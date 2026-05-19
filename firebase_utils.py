# -*- coding: utf-8 -*-
"""
로지판(LogiPan) - Firebase/FCM 모듈
====================================

LogiPan.py에서 분리된 Firebase 관련 기능 모음. (모듈화 3단계)

사용법 (Mixin 패턴):
    from firebase_utils import FirebaseUtilsMixin

    class LogiPanApp(SlackIntegrationMixin, JiraIntegrationMixin, FirebaseUtilsMixin):
        ...

포함된 기능 (총 6개 메서드):
    [FCM 푸시 알림]
      - send_fcm_push  (백그라운드 스레드로 푸시 발송)
    [응답자(대응자) 관리 시스템 - Firestore 트랜잭션]
      - _claim_main_responder  (메인 대응자로 등록, 동시 클릭 방지)
      - _join_extra_responder  (추가 참여자로 등록)
      - _leave_response  (대응에서 빠짐, 메인은 자동 승계)
      - _format_responder_label  (UI 라벨 생성)
    [작업자 목록]
      - get_worker_list  (fcm_tokens에 등록된 작업자)

LogiPan 본체가 제공해야 하는 속성:
    self.db  (firestore.client(), 없으면 None)

LogiPan 본체가 제공해야 하는 메서드: 없음 (완전 자족)

Slack/Jira 모듈과의 의존성: 없음 (완전 독립)

원본 위치: LogiPan.py 라인 689~759 (FCM) + 4309~4437 (응답자 시스템)
분리 일자: 모듈화 3단계
"""

import threading
from firebase_admin import firestore, messaging


class FirebaseUtilsMixin:
    """Firebase/FCM Mixin.

    LogiPanApp이 이 Mixin을 상속하면 self.send_fcm_push() 등을
    그대로 호출 가능. 원본 LogiPan.py와 100% 동작 동일.
    """

    # ========== [FCM 푸시 알림] ==========
    def send_fcm_push(self, target_user, title, body):
        """작업자에게 푸시 알림 발송. (iOS PWA 호환 버전)"""
        if self.db is None:
            return

        import threading
        def _send():
            try:
                token_docs = []
                if target_user == "all":
                    docs = self.db.collection('fcm_tokens').stream()
                    for doc in docs:
                        d = doc.to_dict()
                        if d.get('token'):
                            token_docs.append({'name': doc.id, 'token': d['token']})
                else:
                    doc = self.db.collection('fcm_tokens').document(target_user).get()
                    if doc.exists:
                        d = doc.to_dict()
                        if d.get('token'):
                            token_docs.append({'name': target_user, 'token': d['token']})

                if not token_docs:
                    print(f"📵 FCM 토큰 없음 (대상: {target_user}) - 알림 스킵")
                    return

                success, fail = 0, 0
                for td in token_docs:
                    token = td['token']
                    name = td['name']
                    try:
                        # [수정] 중복 방지: data만 보냄
                        # webpush.notification을 주면 브라우저가 자동으로 띄움 + SW도 띄워서 2번 옴
                        # data만 보내면 SW의 onBackgroundMessage가 한 번만 띄움
                        msg = messaging.Message(
                            data={
                                'title': title,
                                'body': body,
                            },
                            token=token,
                            webpush=messaging.WebpushConfig(
                                headers={
                                    'Urgency': 'high',
                                    'TTL': '86400',
                                },
                                fcm_options=messaging.WebpushFCMOptions(
                                    link='https://ghwkdwjd-debug.github.io/logipan-report/',
                                ),
                            ),
                        )
                        response = messaging.send(msg)
                        success += 1
                        print(f"  ✅ [{name}] 푸시 성공: {response}")
                    except messaging.UnregisteredError:
                        fail += 1
                        print(f"  ❌ [{name}] 토큰 만료 - 자동 정리")
                        try:
                            self.db.collection('fcm_tokens').document(name).delete()
                        except Exception:
                            pass
                    except Exception as e:
                        fail += 1
                        print(f"  ❌ [{name}] 발송 실패: {type(e).__name__}: {e}")

                print(f"📱 FCM 푸시 완료: {success}건 성공 / {fail}건 실패 (대상: {target_user})")
            except Exception as e:
                print(f"⚠️ FCM 발송 전체 오류: {type(e).__name__}: {e}")
                import traceback
                traceback.print_exc()

        threading.Thread(target=_send, daemon=True).start()

    # ========== [추가] 응답자(대응자) 관리 시스템 ==========
    def _claim_main_responder(self, collection_name, doc_id, my_name):
        """[추가] 메인 대응자로 등록 (동시 클릭 방지: 트랜잭션).
        Returns: (success: bool, current_main: str)
        """
        if not (self.db and my_name and doc_id):
            return False, ""
        try:
            doc_ref = self.db.collection(collection_name).document(doc_id)

            @firestore.transactional
            def _try_claim(transaction):
                snap = doc_ref.get(transaction=transaction)
                if not snap.exists:
                    return False, ""
                data = snap.to_dict() or {}
                current = data.get('responder_main', '') or ''
                if current and current != my_name:
                    # 이미 다른 사람이 메인
                    return False, current
                # 비어있거나 나 자신이면 OK (trans-update)
                transaction.update(doc_ref, {
                    'responder_main': my_name,
                    'responder_main_at': firestore.SERVER_TIMESTAMP,
                })
                return True, my_name

            transaction = self.db.transaction()
            return _try_claim(transaction)
        except Exception as e:
            print(f"⚠️ 메인 대응자 등록 실패: {e}")
            return False, ""

    def _join_extra_responder(self, collection_name, doc_id, my_name):
        """[추가] 추가 참여자로 등록 (메인이 아닌 경우)."""
        if not (self.db and my_name and doc_id):
            return False
        try:
            doc_ref = self.db.collection(collection_name).document(doc_id)

            @firestore.transactional
            def _try_join(transaction):
                snap = doc_ref.get(transaction=transaction)
                if not snap.exists:
                    return False
                data = snap.to_dict() or {}
                main = data.get('responder_main', '') or ''
                if main == my_name:
                    return True  # 메인이면 굳이 참여자 추가 X
                extra = data.get('responders_extra', []) or []
                if my_name in extra:
                    return True  # 이미 참여자
                extra.append(my_name)
                transaction.update(doc_ref, {
                    'responders_extra': extra,
                })
                return True

            transaction = self.db.transaction()
            return _try_join(transaction)
        except Exception as e:
            print(f"⚠️ 참여자 등록 실패: {e}")
            return False

    def _leave_response(self, collection_name, doc_id, my_name):
        """[추가] 대응에서 빠지기.
        - 메인이면 → 다음 참여자에게 자동 승계 (없으면 메인 = '')
        - 참여자면 → 그냥 리스트에서 빠짐
        """
        if not (self.db and my_name and doc_id):
            return False
        try:
            doc_ref = self.db.collection(collection_name).document(doc_id)

            @firestore.transactional
            def _try_leave(transaction):
                snap = doc_ref.get(transaction=transaction)
                if not snap.exists:
                    return False
                data = snap.to_dict() or {}
                main = data.get('responder_main', '') or ''
                extra = list(data.get('responders_extra', []) or [])

                update_data = {}
                if main == my_name:
                    # 메인 사퇴 → 자동 승계
                    if extra:
                        new_main = extra.pop(0)
                        update_data['responder_main'] = new_main
                        update_data['responder_main_at'] = firestore.SERVER_TIMESTAMP
                        update_data['responders_extra'] = extra
                    else:
                        update_data['responder_main'] = ''
                        update_data['responder_main_at'] = None
                elif my_name in extra:
                    extra.remove(my_name)
                    update_data['responders_extra'] = extra
                else:
                    return True  # 등록 안 되어 있으면 그냥 무시

                if update_data:
                    transaction.update(doc_ref, update_data)
                return True

            transaction = self.db.transaction()
            return _try_leave(transaction)
        except Exception as e:
            print(f"⚠️ 대응 빠지기 실패: {e}")
            return False

    def _format_responder_label(self, main, extra):
        """[추가] 대응자 라벨 생성: '👤 김관리자' or '👤 김관리자 (외 1명)'"""
        if not main:
            return ""
        extra_count = len(extra) if extra else 0
        if extra_count > 0:
            return f"👤 {main} (외 {extra_count}명) 대응 중"
        return f"👤 {main} 대응 중"

    def get_worker_list(self):
        """fcm_tokens 컬렉션에서 등록된 작업자 이름 목록"""
        try:
            if not hasattr(self, 'db') or self.db is None:
                return []
            docs = self.db.collection('fcm_tokens').stream()
            return sorted([doc.id for doc in docs])
        except Exception as e:
            print(f"⚠️ 작업자 목록 로드 실패: {e}")
            return []
