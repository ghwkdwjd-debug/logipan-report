# -*- coding: utf-8 -*-
"""
로지판(LogiPan) - Slack 통합 모듈
====================================

LogiPan.py에서 분리된 Slack 연동 관련 기능 모음. (모듈화 1단계)

사용법 (Mixin 패턴):
    from slack_integration import SlackIntegrationMixin

    class LogiPanApp(SlackIntegrationMixin):
        ...

포함된 기능 (총 21개 메서드):
    [Slack 설정/메시지]
      - load_slack_settings, save_slack_settings
      - send_slack_message, _test_slack_connection
    [구글시트 목록 (Firestore 공유)]
      - load_sheet_list, save_sheet_list
      - extract_sheet_id_from_url
      - open_sheet_add_dialog
    [Google OAuth + Sheets API]
      - _ensure_google_libs, _get_google_paths
      - is_google_authed, get_google_creds, google_signout
    [시트 검색]
      - _normalize_for_match (@staticmethod)
      - search_sheet_for_brand, search_sheet_smart
    [MD ↔ Slack 매핑 (Firestore 공유)]
      - load_md_mapping, save_md_mapping
      - open_md_mapping_dialog
    [UI 팝업]
      - open_sheet_match_picker
      - open_slack_settings

LogiPan 본체가 제공해야 하는 속성:
    self.root              : tk.Tk
    self.db                : firestore.client() (없으면 None 가능)
    self.config_path       : 로컬 config json 경로
    self.current_user      : 작업자 이름 (Firestore 메타 기록용)

LogiPan 본체가 제공해야 하는 메서드:
    self.position_popup(win, w, h)
    self._bind_esc_close(win)
    self._bind_mousewheel(canvas, container) -> rebind 함수

원본 위치: LogiPan.py 라인 628~2010
분리 일자: 2025-05-08
"""

import os
import json
import re
import urllib.request
import urllib.error
import tkinter as tk
from tkinter import messagebox
from firebase_admin import firestore


class SlackIntegrationMixin:
    """Slack 연동 + 구글시트 검색 + MD 매핑 Mixin.

    LogiPanApp이 이 Mixin을 상속하면 self.open_slack_settings() 등을
    그대로 호출 가능. 원본 LogiPan.py와 100% 동작 동일.
    """

    # ========== [Slack 연동] ==========
    def load_slack_settings(self):
        """로컬 config 파일에서 Slack 설정 로드"""
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, "r", encoding="utf-8") as f:
                    cfg = json.load(f)
                return cfg.get("slack_settings", {})
        except Exception:
            pass
        return {}

    def save_slack_settings(self, settings):
        """로컬 config 파일에 Slack 설정 저장"""
        try:
            cfg = {}
            if os.path.exists(self.config_path):
                try:
                    with open(self.config_path, "r", encoding="utf-8") as f:
                        cfg = json.load(f)
                except Exception:
                    cfg = {}
            cfg["slack_settings"] = settings
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            messagebox.showerror("저장 실패", f"Slack 설정 저장 실패: {e}")
            return False

    def send_slack_message(self, text, blocks=None):
        """Slack Webhook으로 메시지 전송.
        Returns: (성공여부, 메시지)"""
        cfg = self.load_slack_settings()
        if not cfg.get("enabled", False):
            return None, "Slack 비활성화됨"
        webhook_url = cfg.get("webhook_url", "").strip()
        if not webhook_url:
            return False, "Webhook URL 없음"
        try:
            payload = {"text": text}
            if blocks:
                payload["blocks"] = blocks
            data = json.dumps(payload).encode('utf-8')
            req = urllib.request.Request(webhook_url, data=data, headers={
                'Content-Type': 'application/json'
            }, method='POST')
            with urllib.request.urlopen(req, timeout=10) as resp:
                result = resp.read().decode()
            if result == 'ok':
                return True, "전송 완료"
            else:
                return False, f"응답: {result[:100]}"
        except urllib.error.HTTPError as e:
            try:
                err_body = e.read().decode()[:200]
            except:
                err_body = ""
            return False, f"HTTP {e.code}: {err_body}"
        except Exception as e:
            return False, f"오류: {e}"

    def _test_slack_connection(self, webhook_url):
        """Slack 연결 테스트 - 테스트 메시지 보내기"""
        try:
            payload = {"text": "🧪 로지판 Slack 연동 테스트입니다"}
            data = json.dumps(payload).encode('utf-8')
            req = urllib.request.Request(webhook_url.strip(), data=data, headers={
                'Content-Type': 'application/json'
            }, method='POST')
            with urllib.request.urlopen(req, timeout=10) as resp:
                result = resp.read().decode()
            if result == 'ok':
                return True, "✅ 채널에 테스트 메시지 전송됨!"
            else:
                return False, f"응답 이상: {result[:100]}"
        except urllib.error.HTTPError as e:
            try:
                err_body = e.read().decode()[:200]
            except:
                err_body = ""
            return False, f"HTTP {e.code}: {err_body}"
        except Exception as e:
            return False, f"오류: {e}"

    # ========== [구글시트 목록 관리 - Firestore 공유] ==========
    def load_sheet_list(self):
        """Firestore에서 등록된 시트 목록 가져오기.
        캐시 우선, 첫 호출 시에만 Firestore read.
        Returns: [{name, sheet_id, gid}, ...]
        """
        # 캐시 우선
        if hasattr(self, '_sheet_list_cache') and self._sheet_list_cache is not None:
            return self._sheet_list_cache
        try:
            if not hasattr(self, 'db') or self.db is None:
                self._sheet_list_cache = []
                return []
            doc = self.db.collection('config_sheets').document('list').get()
            if doc.exists:
                data = doc.to_dict() or {}
                sheets = data.get('sheets', [])
                if not isinstance(sheets, list):
                    sheets = []
                self._sheet_list_cache = sheets
                return sheets
        except Exception as e:
            print(f"⚠️ 시트 목록 로드 실패: {e}")
        self._sheet_list_cache = []
        return []

    def save_sheet_list(self, sheets):
        """Firestore에 시트 목록 저장.
        Args:
            sheets: [{name, sheet_id, gid}, ...]
        Returns: bool
        """
        try:
            if not hasattr(self, 'db') or self.db is None:
                messagebox.showerror("저장 실패",
                    "Firestore 연결이 없어 저장할 수 없습니다.")
                return False
            self.db.collection('config_sheets').document('list').set({
                'sheets': sheets,
                'updated_at': firestore.SERVER_TIMESTAMP,
                'updated_by': getattr(self, 'current_user', '관리자'),
            })
            # 캐시 갱신
            self._sheet_list_cache = sheets
            return True
        except Exception as e:
            messagebox.showerror("저장 실패", f"시트 목록 저장 실패: {e}")
            return False

    def extract_sheet_id_from_url(self, url_or_id):
        """URL 또는 ID 문자열에서 시트 ID와 gid 추출.
        Returns: (sheet_id, gid)  / 실패 시 (None, None)
        """
        if not url_or_id:
            return None, None
        s = url_or_id.strip()
        # URL이면 정규식으로 추출
        m = re.search(r'/spreadsheets/d/([a-zA-Z0-9_-]+)', s)
        sheet_id = m.group(1) if m else None
        if not sheet_id:
            # URL 형식 아니면 그냥 ID로 간주
            if re.match(r'^[a-zA-Z0-9_-]{20,}$', s):
                sheet_id = s
        # gid 추출
        gid = None
        m2 = re.search(r'[?&#]gid=(\d+)', s)
        if m2:
            gid = m2.group(1)
        return sheet_id, gid

    def open_sheet_add_dialog(self, parent_win, edit_index=None, on_done=None):
        """시트 추가/수정 팝업.
        Args:
            parent_win: 부모 창
            edit_index: 수정 시 기존 인덱스 (None이면 추가)
            on_done: 저장 완료 후 콜백
        """
        win = tk.Toplevel(parent_win)
        is_edit = edit_index is not None
        win.title("✏️ 시트 수정" if is_edit else "➕ 시트 추가")
        win.configure(bg="#F5F6F8")
        try:
            self.position_popup(win, 460, 280)
        except Exception:
            win.geometry("460x280")
        win.transient(parent_win)
        win.grab_set()
        self._bind_esc_close(win)

        sheets = list(self.load_sheet_list())
        existing = sheets[edit_index] if is_edit and 0 <= edit_index < len(sheets) else {}

        # 카드
        card = tk.Frame(win, bg="white",
                          highlightthickness=1, highlightbackground="#E5E7EB")
        card.pack(fill="both", expand=True, padx=18, pady=14)
        tk.Frame(card, bg="#16A34A", width=4).pack(side="left", fill="y")
        inner = tk.Frame(card, bg="white", padx=14, pady=12)
        inner.pack(side="left", fill="both", expand=True)

        # 시트 이름
        tk.Label(inner, text="시트 이름", bg="white",
                 font=("맑은 고딕", 9, "bold"),
                 fg="#374151").pack(anchor="w")
        ent_name = tk.Entry(inner, font=("맑은 고딕", 10),
                              bd=1, relief="solid",
                              highlightthickness=0)
        ent_name.pack(fill="x", ipady=5, pady=(2, 0))
        ent_name.insert(0, existing.get('name', ''))
        tk.Label(inner, text="예: SS26 입출고리스트_매입 VER 1",
                 bg="white", fg="#9CA3AF",
                 font=("맑은 고딕", 8)).pack(anchor="w")

        # 시트 URL/ID
        tk.Label(inner, text="\n시트 URL 또는 ID", bg="white",
                 font=("맑은 고딕", 9, "bold"),
                 fg="#374151").pack(anchor="w")
        ent_url = tk.Entry(inner, font=("맑은 고딕", 9),
                             bd=1, relief="solid",
                             highlightthickness=0)
        ent_url.pack(fill="x", ipady=5, pady=(2, 0))
        # 기존 값 복원 (URL 형태로)
        if existing.get('sheet_id'):
            existing_url = f"https://docs.google.com/spreadsheets/d/{existing['sheet_id']}/edit"
            if existing.get('gid'):
                existing_url += f"?gid={existing['gid']}"
            ent_url.insert(0, existing_url)
        tk.Label(inner, text="구글시트 주소 통째로 붙여넣기 OK (gid도 자동 인식)",
                 bg="white", fg="#9CA3AF",
                 font=("맑은 고딕", 8)).pack(anchor="w")

        # 결과 메시지
        msg_lbl = tk.Label(inner, text="", bg="white",
                             font=("맑은 고딕", 8),
                             wraplength=380, justify="left")
        msg_lbl.pack(fill="x", pady=(6, 0))

        # 버튼
        btn_frame = tk.Frame(win, bg="#F5F6F8")
        btn_frame.pack(fill="x", padx=18, pady=(0, 14))

        def do_save():
            name = ent_name.get().strip()
            url = ent_url.get().strip()
            if not name:
                msg_lbl.config(text="❌ 시트 이름을 입력하세요.", fg="#DC2626")
                return
            sheet_id, gid = self.extract_sheet_id_from_url(url)
            if not sheet_id:
                msg_lbl.config(text="❌ 시트 ID를 인식할 수 없습니다. URL 또는 ID를 확인하세요.",
                                fg="#DC2626")
                return
            new_entry = {'name': name, 'sheet_id': sheet_id, 'gid': gid or ''}
            current = list(self.load_sheet_list())
            if is_edit:
                if 0 <= edit_index < len(current):
                    current[edit_index] = new_entry
            else:
                # 중복 시트 ID 체크
                if any(x.get('sheet_id') == sheet_id for x in current):
                    msg_lbl.config(text="⚠️ 이미 등록된 시트입니다.", fg="#D97706")
                    return
                current.append(new_entry)
            if self.save_sheet_list(current):
                if on_done:
                    on_done()
                win.destroy()

        tk.Button(btn_frame, text="💾 저장",
                   command=do_save,
                   bg="#16A34A", fg="white",
                   font=("맑은 고딕", 9, "bold"),
                   relief="flat", padx=14, pady=6,
                   cursor="hand2").pack(side="right")
        tk.Button(btn_frame, text="취소",
                   command=win.destroy,
                   bg="white", fg="#666",
                   font=("맑은 고딕", 9),
                   relief="flat", padx=14, pady=6,
                   cursor="hand2",
                   highlightthickness=1, highlightbackground="#DDD").pack(side="right", padx=(0, 6))

    # ========== [Google Sheets API 연동] ==========
    def _ensure_google_libs(self):
        """Google API 라이브러리가 없으면 자동 설치.
        Returns: bool (사용 가능 여부)
        """
        try:
            from googleapiclient.discovery import build  # noqa: F401
            from google_auth_oauthlib.flow import InstalledAppFlow  # noqa: F401
            from google.auth.transport.requests import Request  # noqa: F401
            from google.oauth2.credentials import Credentials  # noqa: F401
            return True
        except ImportError:
            pass
        # 자동 설치
        try:
            import subprocess
            import sys
            print("🔄 Google API 라이브러리 설치 중...")
            packages = [
                "google-auth>=2.25.0",
                "google-auth-oauthlib>=1.2.0",
                "google-api-python-client>=2.110.0",
            ]
            result = subprocess.run(
                [sys.executable, "-m", "pip", "install", "--user", "--quiet"] + packages,
                capture_output=True, text=True, timeout=120,
            )
            if result.returncode != 0:
                print(f"⚠️ pip install 실패: {result.stderr}")
                return False
            print("✅ Google API 라이브러리 설치 완료")
            # 다시 import 시도
            from googleapiclient.discovery import build  # noqa: F401
            from google_auth_oauthlib.flow import InstalledAppFlow  # noqa: F401
            from google.auth.transport.requests import Request  # noqa: F401
            from google.oauth2.credentials import Credentials  # noqa: F401
            return True
        except Exception as e:
            print(f"⚠️ Google 라이브러리 설치 실패: {e}")
            return False

    def _get_google_paths(self):
        """OAuth 관련 파일 경로 반환.
        Returns: (client_secret_path, token_path)
        """
        script_dir = os.path.dirname(os.path.abspath(__file__))
        client_path = os.path.join(script_dir, "oauth_client.json")
        token_path = os.path.join(script_dir, "google_token.json")
        return client_path, token_path

    def is_google_authed(self):
        """Google OAuth 인증된 토큰이 있는지 확인 (만료/유효성은 따로)"""
        _, token_path = self._get_google_paths()
        return os.path.exists(token_path)

    def get_google_creds(self, force_reauth=False):
        """Google OAuth credentials 가져오기 (필요 시 자동 갱신).
        Returns: Credentials | None
        """
        if not self._ensure_google_libs():
            return None
        try:
            from google.oauth2.credentials import Credentials
            from google.auth.transport.requests import Request
            from google_auth_oauthlib.flow import InstalledAppFlow
        except ImportError:
            return None

        SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
        client_path, token_path = self._get_google_paths()

        if not os.path.exists(client_path):
            print(f"⚠️ oauth_client.json 없음: {client_path}")
            return None

        creds = None
        if not force_reauth and os.path.exists(token_path):
            try:
                creds = Credentials.from_authorized_user_file(token_path, SCOPES)
            except Exception as e:
                print(f"⚠️ 토큰 로드 실패: {e}")
                creds = None

        # 만료 + refresh 토큰 있으면 자동 갱신
        if creds and not creds.valid:
            if creds.expired and creds.refresh_token:
                try:
                    creds.refresh(Request())
                    # 갱신된 토큰 저장
                    with open(token_path, 'w', encoding='utf-8') as f:
                        f.write(creds.to_json())
                except Exception as e:
                    print(f"⚠️ 토큰 갱신 실패: {e}")
                    creds = None
            else:
                creds = None

        # 토큰 없거나 무효면 브라우저 인증 (사용자 액션 필요!)
        if not creds:
            try:
                flow = InstalledAppFlow.from_client_secrets_file(client_path, SCOPES)
                # run_local_server: 브라우저 자동으로 띄움, 로그인 후 자동 콜백
                creds = flow.run_local_server(port=0, prompt='consent')
                # 토큰 저장
                with open(token_path, 'w', encoding='utf-8') as f:
                    f.write(creds.to_json())
                print("✅ Google OAuth 인증 완료")
            except Exception as e:
                print(f"❌ Google 인증 실패: {e}")
                return None

        return creds

    def google_signout(self):
        """저장된 토큰 삭제 (재인증 강제)"""
        try:
            _, token_path = self._get_google_paths()
            if os.path.exists(token_path):
                os.remove(token_path)
            return True
        except Exception as e:
            print(f"⚠️ 토큰 삭제 실패: {e}")
            return False

    @staticmethod
    def _normalize_for_match(s):
        """매칭용 문자열 정규화 (공백/언더바/대소문자 무시)"""
        if not s:
            return ""
        return re.sub(r'[\s_]+', '', str(s)).lower()

    def search_sheet_for_brand(self, brand_name, md_name=None, progress_callback=None):
        """등록된 모든 시트에서 브랜드+담당MD 매칭 행 찾기.
        Args:
            brand_name: 브랜드명 (예: "ADIDAS26_3월1주차")
            md_name: 담당자명 (선택)
            progress_callback: 진행상황 콜백 (str -> None)
        Returns: [{sheet_name, sheet_id, gid, row_index, brand, md, url}, ...]
            (매칭된 모든 행)
        """
        results = []
        if not brand_name:
            return results
        if not self._ensure_google_libs():
            print("⚠️ Google 라이브러리 없음")
            return results
        creds = self.get_google_creds()
        if not creds:
            print("⚠️ Google 인증 실패")
            return results

        try:
            from googleapiclient.discovery import build
            service = build('sheets', 'v4', credentials=creds, cache_discovery=False)
        except Exception as e:
            print(f"⚠️ Sheets API 초기화 실패: {e}")
            return results

        sheets = self.load_sheet_list()
        target_brand = self._normalize_for_match(brand_name)
        target_md = self._normalize_for_match(md_name) if md_name else None

        for sheet_info in sheets:
            sheet_id = sheet_info.get('sheet_id')
            sheet_display_name = sheet_info.get('name', '(이름없음)')
            target_gid = sheet_info.get('gid')  # 등록된 gid (메인 탭)
            if not sheet_id:
                continue
            if progress_callback:
                progress_callback(f"🔍 {sheet_display_name} 검색 중...")

            try:
                # 1. 시트 메타데이터로 탭 이름 찾기 (gid → 탭이름 변환 필요)
                meta = service.spreadsheets().get(spreadsheetId=sheet_id).execute()
                tab_name = None
                for sh in meta.get('sheets', []):
                    props = sh.get('properties', {})
                    if target_gid and str(props.get('sheetId', '')) == str(target_gid):
                        tab_name = props.get('title')
                        break
                # gid 매칭 실패 시 첫 번째 탭 사용
                if not tab_name and meta.get('sheets'):
                    tab_name = meta['sheets'][0].get('properties', {}).get('title')

                if not tab_name:
                    continue

                # 2. 탭 데이터 모두 가져오기 (헤더 자동 인식 위해 5~6행부터 보고싶지만 일단 전체)
                # 안전하게 1~1000행 X 1~30열 정도 가져옴
                range_str = f"'{tab_name}'!A1:AD1000"
                resp = service.spreadsheets().values().get(
                    spreadsheetId=sheet_id, range=range_str,
                    valueRenderOption='FORMATTED_VALUE',
                ).execute()
                values = resp.get('values', [])
                if not values:
                    continue

                # 3. 헤더 행 찾기 (브랜드/브랜드 담당자/바코드표 키워드 포함된 행)
                header_row_idx = None
                for i, row in enumerate(values[:15]):  # 상위 15행 안에 있다고 가정
                    row_text = "|".join(str(c) for c in row)
                    if ('브랜드' in row_text and
                            ('담당자' in row_text or '바코드표' in row_text or 'URL' in row_text.upper())):
                        header_row_idx = i
                        break
                if header_row_idx is None:
                    if progress_callback:
                        progress_callback(f"⚠️ {sheet_display_name}: 헤더 못 찾음")
                    continue
                header = values[header_row_idx]

                # 4. 컬럼 인덱스 식별
                # - 브랜드 담당자 컬럼: '담당자' 또는 '브랜드 담당자' 정확매칭 우선
                # - 브랜드 컬럼: 헤더 값이 정확히 "브랜드" 인 것 (없으면 부분 매칭)
                # - URL 컬럼: '바코드표' 정확 또는 'URL' 포함
                def find_col(predicate):
                    for idx, h in enumerate(header):
                        if predicate(str(h).strip()):
                            return idx
                    return None

                col_md = find_col(lambda h: h == '브랜드 담당자') \
                          or find_col(lambda h: h == '담당자') \
                          or find_col(lambda h: '담당자' in h)
                # 브랜드 컬럼: '브랜드'로 시작하지만 '담당자' 안 들어간 거
                col_brand = find_col(lambda h: h == '브랜드') \
                             or find_col(lambda h: h.startswith('브랜드') and '담당자' not in h)
                col_url = find_col(lambda h: h == '바코드표') \
                           or find_col(lambda h: 'URL' in h.upper()) \
                           or find_col(lambda h: '링크' in h)

                if col_brand is None or col_url is None:
                    if progress_callback:
                        progress_callback(f"⚠️ {sheet_display_name}: 필수 컬럼 못 찾음")
                    continue

                # 5. 데이터 행 순회 (헤더 다음부터)
                for row_idx in range(header_row_idx + 1, len(values)):
                    row = values[row_idx]
                    if not row:
                        continue
                    # 인덱스 안전하게
                    cell_brand = row[col_brand] if col_brand < len(row) else ""
                    cell_md = row[col_md] if (col_md is not None and col_md < len(row)) else ""
                    cell_url = row[col_url] if col_url < len(row) else ""
                    if not cell_brand or not cell_url:
                        continue
                    # 브랜드 매칭
                    if self._normalize_for_match(cell_brand) != target_brand:
                        continue
                    # 담당자 매칭 (입력했을 때만)
                    md_matched = True  # 담당자 매칭 여부 표시
                    if target_md and cell_md:
                        if self._normalize_for_match(cell_md) != target_md:
                            md_matched = False
                    # URL이 http:// 또는 https://로 시작 안 하면 스킵
                    cell_url_str = str(cell_url).strip()
                    if not (cell_url_str.startswith('http://') or cell_url_str.startswith('https://')):
                        continue
                    results.append({
                        'sheet_name': sheet_display_name,
                        'sheet_id': sheet_id,
                        'gid': target_gid or '',
                        'row_index': row_idx + 1,  # 1-based
                        'brand': cell_brand,
                        'md': cell_md,
                        'url': cell_url_str,
                        'md_matched': md_matched,  # 담당자도 매칭됐는지
                    })
            except Exception as e:
                print(f"⚠️ {sheet_display_name} 검색 오류: {e}")
                if progress_callback:
                    progress_callback(f"⚠️ {sheet_display_name} 오류: {e}")
                continue

        if progress_callback:
            progress_callback(f"✅ 검색 완료 ({len(results)}건)")
        return results

    def search_sheet_smart(self, brand_name, md_name=None, progress_callback=None):
        """스마트 검색: 정확매칭 우선 → 0개면 브랜드명만으로 fallback.
        Returns: (results, mode)
            mode: 'strict' (담당자까지 매칭) | 'brand_only' (브랜드만 매칭) | 'none' (0개)
        """
        # 1차: 브랜드+담당자 모두 매칭된 결과만
        all_results = self.search_sheet_for_brand(
            brand_name, md_name=md_name, progress_callback=progress_callback)
        strict = [r for r in all_results if r.get('md_matched', True)]
        if strict:
            return strict, 'strict'
        # 0개면 브랜드만 매칭된 결과로 fallback
        if all_results:
            if progress_callback:
                progress_callback("🔁 담당자 미매칭 - 브랜드명만으로 재검색")
            return all_results, 'brand_only'
        return [], 'none'

    # ========== [MD ↔ Slack 매핑 - Firestore 공유] ==========
    def load_md_mapping(self):
        """Firestore에서 MD 매핑 가져오기 (캐시 우선).
        Returns: {md_name: slack_member_id, ...}
        """
        if hasattr(self, '_md_mapping_cache') and self._md_mapping_cache is not None:
            return self._md_mapping_cache
        try:
            if not hasattr(self, 'db') or self.db is None:
                self._md_mapping_cache = {}
                return {}
            doc = self.db.collection('config_md_mapping').document('list').get()
            if doc.exists:
                data = doc.to_dict() or {}
                mapping = data.get('mapping', {})
                if not isinstance(mapping, dict):
                    mapping = {}
                self._md_mapping_cache = mapping
                return mapping
        except Exception as e:
            print(f"⚠️ MD 매핑 로드 실패: {e}")
        self._md_mapping_cache = {}
        return {}

    def save_md_mapping(self, mapping):
        """Firestore에 MD 매핑 저장.
        Args:
            mapping: {md_name: slack_member_id, ...}
        """
        try:
            if not hasattr(self, 'db') or self.db is None:
                messagebox.showerror("저장 실패",
                    "Firestore 연결이 없어 저장할 수 없습니다.")
                return False
            self.db.collection('config_md_mapping').document('list').set({
                'mapping': mapping,
                'updated_at': firestore.SERVER_TIMESTAMP,
                'updated_by': getattr(self, 'current_user', '관리자'),
            })
            self._md_mapping_cache = mapping
            return True
        except Exception as e:
            messagebox.showerror("저장 실패", f"MD 매핑 저장 실패: {e}")
            return False

    def open_md_mapping_dialog(self, parent_win, edit_name=None, on_done=None):
        """MD 매핑 추가/수정 팝업"""
        win = tk.Toplevel(parent_win)
        is_edit = edit_name is not None
        win.title("✏️ MD 수정" if is_edit else "➕ MD 추가")
        win.configure(bg="#F5F6F8")
        try:
            self.position_popup(win, 460, 280)
        except Exception:
            win.geometry("460x280")
        win.transient(parent_win)
        win.grab_set()
        self._bind_esc_close(win)

        mapping = dict(self.load_md_mapping())
        existing_id = mapping.get(edit_name, '') if is_edit else ''

        # 카드
        card = tk.Frame(win, bg="white",
                          highlightthickness=1, highlightbackground="#E5E7EB")
        card.pack(fill="both", expand=True, padx=18, pady=14)
        tk.Frame(card, bg="#3B82F6", width=4).pack(side="left", fill="y")
        inner = tk.Frame(card, bg="white", padx=14, pady=12)
        inner.pack(side="left", fill="both", expand=True)

        # MD 이름
        tk.Label(inner, text="MD 이름", bg="white",
                 font=("맑은 고딕", 9, "bold"),
                 fg="#374151").pack(anchor="w")
        ent_name = tk.Entry(inner, font=("맑은 고딕", 10),
                              bd=1, relief="solid",
                              highlightthickness=0)
        ent_name.pack(fill="x", ipady=5, pady=(2, 0))
        ent_name.insert(0, edit_name or '')
        if is_edit:
            ent_name.config(state="readonly")  # 수정 시 이름 변경 불가
        tk.Label(inner,
                 text="시트의 '브랜드 담당자' 컬럼에 적힌 이름과 정확히 일치해야 합니다.",
                 bg="white", fg="#9CA3AF",
                 font=("맑은 고딕", 8), wraplength=400, justify="left").pack(anchor="w")

        # Slack 멤버 ID
        tk.Label(inner, text="\nSlack 멤버 ID", bg="white",
                 font=("맑은 고딕", 9, "bold"),
                 fg="#374151").pack(anchor="w")
        ent_id = tk.Entry(inner, font=("맑은 고딕", 10),
                            bd=1, relief="solid",
                            highlightthickness=0)
        ent_id.pack(fill="x", ipady=5, pady=(2, 0))
        ent_id.insert(0, existing_id)
        tk.Label(inner,
                 text="예: U01ABC123XYZ (Slack 프로필 → ⋮ → '멤버 ID 복사')",
                 bg="white", fg="#9CA3AF",
                 font=("맑은 고딕", 8), wraplength=400, justify="left").pack(anchor="w")

        # 결과 메시지
        msg_lbl = tk.Label(inner, text="", bg="white",
                             font=("맑은 고딕", 8),
                             wraplength=400, justify="left")
        msg_lbl.pack(fill="x", pady=(6, 0))

        # 버튼
        btn_frame = tk.Frame(win, bg="#F5F6F8")
        btn_frame.pack(fill="x", padx=18, pady=(0, 14))

        def do_save():
            name = ent_name.get().strip()
            slack_id = ent_id.get().strip()
            if not name:
                msg_lbl.config(text="❌ MD 이름을 입력하세요.", fg="#DC2626")
                return
            if not slack_id:
                msg_lbl.config(text="❌ Slack 멤버 ID를 입력하세요.", fg="#DC2626")
                return
            # Slack ID 형식 간단 검증 (U나 W로 시작 + 영숫자)
            if not re.match(r'^[UW][A-Z0-9]{5,}$', slack_id):
                msg_lbl.config(
                    text="⚠️ Slack 멤버 ID 형식이 이상합니다. (보통 U로 시작)",
                    fg="#D97706")
                # 경고만 하고 저장은 허용
            current = dict(self.load_md_mapping())
            current[name] = slack_id
            if self.save_md_mapping(current):
                if on_done:
                    on_done()
                win.destroy()

        tk.Button(btn_frame, text="💾 저장",
                   command=do_save,
                   bg="#3B82F6", fg="white",
                   font=("맑은 고딕", 9, "bold"),
                   relief="flat", padx=14, pady=6,
                   cursor="hand2").pack(side="right")
        tk.Button(btn_frame, text="취소",
                   command=win.destroy,
                   bg="white", fg="#666",
                   font=("맑은 고딕", 9),
                   relief="flat", padx=14, pady=6,
                   cursor="hand2",
                   highlightthickness=1, highlightbackground="#DDD").pack(side="right", padx=(0, 6))

    def open_sheet_match_picker(self, brand_name, results, on_pick):
        """다중 매칭 시 선택 팝업.
        Args:
            brand_name: 검색한 브랜드명
            results: search_sheet_for_brand 결과 리스트
            on_pick: 선택 콜백 — on_pick(picked_dict | None)
                None이면 시트 링크 없이 진행
        """
        win = tk.Toplevel(self.root)
        win.title("🔍 시트 행 선택")
        win.configure(bg="#F5F6F8")
        try:
            self.position_popup(win, 600, 480)
        except Exception:
            win.geometry("600x480")
        win.transient(self.root)
        win.grab_set()
        self._bind_esc_close(win)

        # 헤더
        head = tk.Frame(win, bg="#F5F6F8")
        head.pack(fill="x", padx=18, pady=(14, 4))
        tk.Label(head, text="🔍", font=("맑은 고딕", 18), bg="#F5F6F8").pack(side="left", padx=(0, 6))
        tk.Label(head, text=f"'{brand_name}' 매칭 결과 {len(results)}건",
                 font=("맑은 고딕", 13, "bold"),
                 bg="#F5F6F8", fg="#1A1A1A").pack(side="left")
        tk.Label(win, text="Slack 알림에 첨부할 시트 행을 선택하세요.",
                 bg="#F5F6F8", fg="#666",
                 font=("맑은 고딕", 9)).pack(padx=18, anchor="w")

        # 카드 (스크롤 영역)
        card = tk.Frame(win, bg="white",
                          highlightthickness=1, highlightbackground="#E5E7EB")
        card.pack(fill="both", expand=True, padx=18, pady=12)

        canvas = tk.Canvas(card, bg="white", highlightthickness=0)
        canvas.pack(side="left", fill="both", expand=True)
        scroll = tk.Scrollbar(card, command=canvas.yview)
        scroll.pack(side="right", fill="y")
        canvas.configure(yscrollcommand=scroll.set)
        list_inner = tk.Frame(canvas, bg="white")
        canvas.create_window((0, 0), window=list_inner, anchor="nw")

        picked_var = {"value": None}

        def make_pick(item):
            def _pick():
                picked_var["value"] = item
                win.destroy()
            return _pick

        for r in results:
            row = tk.Frame(list_inner, bg="white",
                             highlightthickness=1, highlightbackground="#E5E7EB",
                             cursor="hand2")
            row.pack(fill="x", padx=8, pady=4)

            content = tk.Frame(row, bg="white", padx=10, pady=8)
            content.pack(fill="x")

            # 1줄: 시트명 + 행번호
            line1 = tk.Frame(content, bg="white")
            line1.pack(fill="x")
            tk.Label(line1, text=f"📋 {r['sheet_name']}",
                     bg="white", fg="#1A1A1A",
                     font=("맑은 고딕", 10, "bold")).pack(side="left")
            tk.Label(line1, text=f"  행 {r['row_index']}",
                     bg="white", fg="#9CA3AF",
                     font=("맑은 고딕", 9)).pack(side="left")

            # 2줄: 브랜드 / 담당자
            line2 = tk.Frame(content, bg="white")
            line2.pack(fill="x", pady=(2, 0))
            tk.Label(line2, text=f"🏷️ {r['brand']}",
                     bg="white", fg="#374151",
                     font=("맑은 고딕", 9)).pack(side="left")
            tk.Label(line2, text=f"   👤 {r['md'] or '(담당자 정보 없음)'}",
                     bg="white", fg="#374151",
                     font=("맑은 고딕", 9)).pack(side="left")

            # 3줄: URL 미리보기
            url_short = r['url']
            if len(url_short) > 70:
                url_short = url_short[:65] + "..."
            tk.Label(content, text=f"🔗 {url_short}",
                     bg="white", fg="#3B82F6",
                     font=("맑은 고딕", 8)).pack(anchor="w", pady=(2, 0))

            # 선택 버튼
            tk.Button(content, text="✓ 이 행 선택",
                       command=make_pick(r),
                       bg="#16A34A", fg="white",
                       font=("맑은 고딕", 9, "bold"),
                       relief="flat", padx=10, pady=4,
                       cursor="hand2").pack(anchor="e", pady=(4, 0))

        list_inner.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))

        # 하단 버튼
        bottom = tk.Frame(win, bg="#F5F6F8")
        bottom.pack(fill="x", padx=18, pady=(0, 14))

        tk.Button(bottom, text="시트 링크 없이 진행",
                   command=lambda: (picked_var.update(value=None), win.destroy()),
                   bg="white", fg="#666",
                   font=("맑은 고딕", 9),
                   relief="flat", padx=12, pady=6,
                   cursor="hand2",
                   highlightthickness=1, highlightbackground="#DDD").pack(side="left")

        # 모달 대기
        win.wait_window()
        on_pick(picked_var["value"])

    def open_slack_settings(self):
        """Slack 연동 설정 팝업"""
        win = tk.Toplevel(self.root)
        win.title("💬 Slack 알림 설정")
        win.configure(bg="#F5F6F8")
        try:
            self.position_popup(win, 560, 740)
        except Exception:
            win.geometry("560x740")
        win.transient(self.root)
        win.grab_set()
        self._bind_esc_close(win)

        # ===== 스크롤 컨테이너 =====
        # 외부: 스크롤바 + 캔버스
        scroll_outer = tk.Frame(win, bg="#F5F6F8")
        scroll_outer.pack(fill="both", expand=True)
        scroll_canvas = tk.Canvas(scroll_outer, bg="#F5F6F8",
                                     highlightthickness=0, bd=0)
        scroll_canvas.pack(side="left", fill="both", expand=True)
        scroll_bar = tk.Scrollbar(scroll_outer, command=scroll_canvas.yview)
        scroll_bar.pack(side="right", fill="y")
        scroll_canvas.configure(yscrollcommand=scroll_bar.set)

        # 내부: 실제 위젯이 들어갈 프레임 (이게 곧 'win'을 대체)
        win_body = tk.Frame(scroll_canvas, bg="#F5F6F8")
        body_window = scroll_canvas.create_window((0, 0), window=win_body, anchor="nw")

        def _on_body_resize(event):
            scroll_canvas.configure(scrollregion=scroll_canvas.bbox("all"))
        def _on_canvas_resize(event):
            # 캔버스 폭에 맞춰 내부 프레임 폭 조정
            scroll_canvas.itemconfig(body_window, width=event.width)
        win_body.bind("<Configure>", _on_body_resize)
        scroll_canvas.bind("<Configure>", _on_canvas_resize)

        # 마우스 휠 스크롤 - 캔버스/win_body에만 바인딩 (전역 X)
        # 자손 위젯은 아직 안 그려졌으므로 나중에 다시 바인딩 (deferred)
        rebind_wheel_slack = self._bind_mousewheel(scroll_canvas, win_body)
        # win_body의 자손이 모두 추가된 후 다시 바인딩하기 위한 핸들 저장
        self._slack_wheel_rebind = rebind_wheel_slack

        s = self.load_slack_settings()

        # 헤더
        head = tk.Frame(win_body, bg="#F5F6F8")
        head.pack(fill="x", padx=18, pady=(14, 4))
        tk.Label(head, text="💬", font=("맑은 고딕", 18), bg="#F5F6F8").pack(side="left", padx=(0, 6))
        tk.Label(head, text="Slack 알림 설정",
                 font=("맑은 고딕", 14, "bold"),
                 bg="#F5F6F8", fg="#1A1A1A").pack(side="left")
        tk.Label(win_body, text="Jira 상신 시 슬랙 채널에 자동 알림",
                 bg="#F5F6F8", fg="#666",
                 font=("맑은 고딕", 8)).pack(padx=18, anchor="w")

        # 카드
        card = tk.Frame(win_body, bg="white",
                          highlightthickness=1, highlightbackground="#E5E7EB")
        card.pack(fill="x", padx=18, pady=12)
        tk.Frame(card, bg="#4A154B", width=4).pack(side="left", fill="y")  # Slack 색상
        inner = tk.Frame(card, bg="white", padx=14, pady=12)
        inner.pack(side="left", fill="both", expand=True)

        # Webhook URL
        tk.Label(inner, text="Webhook URL", bg="white",
                 font=("맑은 고딕", 9, "bold"),
                 fg="#374151").pack(anchor="w")
        ent_url = tk.Entry(inner, font=("맑은 고딕", 9),
                            bd=1, relief="solid",
                            highlightthickness=0,
                            show="•")
        ent_url.pack(fill="x", ipady=4, pady=(2, 0))
        ent_url.insert(0, s.get("webhook_url", ""))
        tk.Label(inner, text="https://hooks.slack.com/services/T.../B.../...",
                 bg="white", fg="#9CA3AF",
                 font=("맑은 고딕", 8),
                 anchor="w").pack(fill="x")

        # 채널명 (선택, 표시용)
        tk.Label(inner, text="\n채널명 (참고용)", bg="white",
                 font=("맑은 고딕", 9, "bold"),
                 fg="#374151").pack(anchor="w")
        ent_channel = tk.Entry(inner, font=("맑은 고딕", 10),
                                 bd=1, relief="solid",
                                 highlightthickness=0)
        ent_channel.pack(fill="x", ipady=4, pady=(2, 0))
        ent_channel.insert(0, s.get("channel_name", ""))
        tk.Label(inner, text="예: #입고알림 (메모용, 실제 채널은 Webhook이 결정)",
                 bg="white", fg="#9CA3AF",
                 font=("맑은 고딕", 8),
                 anchor="w").pack(fill="x")

        # CC 멘션 (선택)
        tk.Label(inner, text="\nCC 멘션 (선택)", bg="white",
                 font=("맑은 고딕", 9, "bold"),
                 fg="#374151").pack(anchor="w")
        ent_cc = tk.Entry(inner, font=("맑은 고딕", 10),
                            bd=1, relief="solid",
                            highlightthickness=0)
        ent_cc.pack(fill="x", ipady=4, pady=(2, 0))
        ent_cc.insert(0, s.get("cc_mention", ""))
        tk.Label(inner,
                 text=("모든 메시지 끝에 자동으로 'CC. ...' 줄 추가\n"
                       "사용자그룹: <!subteam^S09B9UPG4UV>  /  여러명: <@U01> <@U02>"),
                 bg="white", fg="#9CA3AF",
                 font=("맑은 고딕", 8),
                 anchor="w", justify="left").pack(fill="x")

        # 활성화
        enabled_var = tk.BooleanVar(value=s.get("enabled", False))
        chk_frame = tk.Frame(inner, bg="white")
        chk_frame.pack(fill="x", pady=(12, 0))
        tk.Checkbutton(chk_frame, text="✅ Jira 상신 시 슬랙 알림 자동 전송",
                        variable=enabled_var, bg="white",
                        font=("맑은 고딕", 9, "bold"),
                        fg="#4A154B").pack(anchor="w")

        # 결과 표시
        result_lbl = tk.Label(inner, text="", bg="white",
                                font=("맑은 고딕", 8),
                                wraplength=380, justify="left")
        result_lbl.pack(fill="x", pady=(6, 0))

        # ========== [구글시트 관리 섹션] ==========
        sheet_card = tk.Frame(win_body, bg="white",
                                highlightthickness=1, highlightbackground="#E5E7EB")
        sheet_card.pack(fill="x", padx=18, pady=(0, 12))
        tk.Frame(sheet_card, bg="#16A34A", width=4).pack(side="left", fill="y")
        sheet_inner = tk.Frame(sheet_card, bg="white", padx=14, pady=10)
        sheet_inner.pack(side="left", fill="both", expand=True)

        # 헤더
        sheet_head = tk.Frame(sheet_inner, bg="white")
        sheet_head.pack(fill="x")
        tk.Label(sheet_head, text="📊 구글시트 자동 검색",
                 bg="white", fg="#1A1A1A",
                 font=("맑은 고딕", 10, "bold")).pack(side="left")
        # 시트 활성화
        sheet_enabled_var = tk.BooleanVar(value=s.get("sheet_search_enabled", False))
        tk.Checkbutton(sheet_head, text="활성화",
                        variable=sheet_enabled_var, bg="white",
                        font=("맑은 고딕", 9),
                        fg="#16A34A").pack(side="right")

        tk.Label(sheet_inner,
                 text="입고 시 브랜드명+담당MD로 시트 검색 → Slack 알림에 시트 링크 자동 첨부",
                 bg="white", fg="#6B7280",
                 font=("맑은 고딕", 8),
                 wraplength=440, justify="left", anchor="w").pack(fill="x", pady=(2, 6))

        # 등록된 시트 리스트 (스크롤 가능 영역)
        list_outer = tk.Frame(sheet_inner, bg="#F9FAFB",
                                 highlightthickness=1, highlightbackground="#E5E7EB")
        list_outer.pack(fill="both", expand=True, pady=(0, 6))

        # 캔버스 + 스크롤바 (Tk 기본은 Listbox에 줄단위 위젯이 어려워서 Canvas 사용)
        list_canvas = tk.Canvas(list_outer, bg="#F9FAFB", highlightthickness=0, height=140)
        list_canvas.pack(side="left", fill="both", expand=True)
        list_scroll = tk.Scrollbar(list_outer, command=list_canvas.yview)
        list_scroll.pack(side="right", fill="y")
        list_canvas.configure(yscrollcommand=list_scroll.set)
        list_inner = tk.Frame(list_canvas, bg="#F9FAFB")
        list_canvas.create_window((0, 0), window=list_inner, anchor="nw")

        def refresh_sheet_list():
            # 기존 위젯 다 지우고
            for w in list_inner.winfo_children():
                w.destroy()
            sheets = self.load_sheet_list()
            if not sheets:
                tk.Label(list_inner, text="등록된 시트가 없습니다.\n[+ 시트 추가] 버튼으로 등록하세요.",
                         bg="#F9FAFB", fg="#9CA3AF",
                         font=("맑은 고딕", 9), justify="center").pack(pady=20)
            else:
                for idx, sh in enumerate(sheets):
                    row = tk.Frame(list_inner, bg="white",
                                     highlightthickness=1, highlightbackground="#E5E7EB")
                    row.pack(fill="x", padx=2, pady=2)

                    info = tk.Frame(row, bg="white")
                    info.pack(side="left", fill="x", expand=True, padx=8, pady=6)
                    tk.Label(info, text="📋 " + sh.get('name', '(이름 없음)'),
                             bg="white", fg="#1A1A1A",
                             font=("맑은 고딕", 9, "bold"),
                             anchor="w").pack(fill="x")
                    sid = sh.get('sheet_id', '')
                    sid_short = (sid[:20] + "...") if len(sid) > 23 else sid
                    tk.Label(info, text=f"ID: {sid_short}",
                             bg="white", fg="#9CA3AF",
                             font=("맑은 고딕", 8),
                             anchor="w").pack(fill="x")

                    btn_box = tk.Frame(row, bg="white")
                    btn_box.pack(side="right", padx=4)

                    def make_edit(i):
                        return lambda: self.open_sheet_add_dialog(
                            win, edit_index=i, on_done=refresh_sheet_list)
                    def make_delete(i):
                        def _del():
                            sheets_now = list(self.load_sheet_list())
                            if 0 <= i < len(sheets_now):
                                if messagebox.askyesno(
                                        "삭제 확인",
                                        f"'{sheets_now[i].get('name', '')}' 시트를 목록에서 제거할까요?",
                                        parent=win):
                                    del sheets_now[i]
                                    if self.save_sheet_list(sheets_now):
                                        refresh_sheet_list()
                        return _del

                    tk.Button(btn_box, text="✏️",
                               command=make_edit(idx),
                               bg="white", fg="#3B82F6",
                               font=("맑은 고딕", 9),
                               relief="flat", cursor="hand2",
                               width=2).pack(side="left")
                    tk.Button(btn_box, text="🗑️",
                               command=make_delete(idx),
                               bg="white", fg="#DC2626",
                               font=("맑은 고딕", 9),
                               relief="flat", cursor="hand2",
                               width=2).pack(side="left")

            # 캔버스 스크롤 영역 갱신
            list_inner.update_idletasks()
            list_canvas.configure(scrollregion=list_canvas.bbox("all"))

        refresh_sheet_list()

        # 시트 추가 + Google 인증 + 검색 테스트 버튼
        sheet_btn_frame = tk.Frame(sheet_inner, bg="white")
        sheet_btn_frame.pack(fill="x", pady=(2, 0))

        tk.Button(sheet_btn_frame, text="➕ 시트 추가",
                   command=lambda: self.open_sheet_add_dialog(
                       win, on_done=refresh_sheet_list),
                   bg="#16A34A", fg="white",
                   font=("맑은 고딕", 9, "bold"),
                   relief="flat", padx=12, pady=5,
                   cursor="hand2").pack(side="left")

        # 인증 상태 라벨
        auth_status_lbl = tk.Label(sheet_btn_frame, text="",
                                      bg="white", fg="#6B7280",
                                      font=("맑은 고딕", 8))
        auth_status_lbl.pack(side="left", padx=(10, 0))

        def update_auth_status():
            if self.is_google_authed():
                auth_status_lbl.config(text="🔐 Google: 인증됨", fg="#16A34A")
            else:
                auth_status_lbl.config(text="🔓 Google: 미인증", fg="#DC2626")
        update_auth_status()

        def do_google_auth():
            """OAuth 인증 (브라우저 띄움) — 끝나면 상태 업데이트"""
            auth_status_lbl.config(text="🔄 인증 중... 브라우저를 확인하세요", fg="#6B7280")
            win.update()
            # 토큰 삭제 후 강제 재인증
            self.google_signout()
            creds = self.get_google_creds(force_reauth=True)
            if creds:
                update_auth_status()
                messagebox.showinfo("인증 완료", "✅ Google 계정 인증이 완료되었습니다.",
                                     parent=win)
            else:
                update_auth_status()
                messagebox.showerror("인증 실패",
                    "❌ Google 인증에 실패했습니다.\n\n"
                    "확인 사항:\n"
                    "- oauth_client.json 파일이 있는지\n"
                    "- 인터넷 연결 상태\n"
                    "- 브라우저에서 동의를 했는지",
                    parent=win)

        tk.Button(sheet_btn_frame, text="🔐 Google 인증",
                   command=do_google_auth,
                   bg="#3B82F6", fg="white",
                   font=("맑은 고딕", 9, "bold"),
                   relief="flat", padx=10, pady=5,
                   cursor="hand2").pack(side="right")

        # 검색 테스트 버튼 (별도 줄)
        test_search_frame = tk.Frame(sheet_inner, bg="white")
        test_search_frame.pack(fill="x", pady=(8, 0))
        tk.Label(test_search_frame, text="🔍 시트 검색 테스트:",
                 bg="white", fg="#6B7280",
                 font=("맑은 고딕", 8)).pack(side="left")
        ent_test_brand = tk.Entry(test_search_frame, font=("맑은 고딕", 9),
                                    bd=1, relief="solid",
                                    highlightthickness=0, width=22)
        ent_test_brand.pack(side="left", ipady=2, padx=(4, 4))
        # placeholder
        _test_ph = "브랜드명 입력"
        ent_test_brand.insert(0, _test_ph)
        ent_test_brand.config(fg="#9CA3AF")
        def _tb_in(e):
            if ent_test_brand.get() == _test_ph:
                ent_test_brand.delete(0, tk.END); ent_test_brand.config(fg="#111827")
        def _tb_out(e):
            if not ent_test_brand.get().strip():
                ent_test_brand.delete(0, tk.END); ent_test_brand.insert(0, _test_ph); ent_test_brand.config(fg="#9CA3AF")
        ent_test_brand.bind("<FocusIn>", _tb_in)
        ent_test_brand.bind("<FocusOut>", _tb_out)

        def do_search_test():
            brand = ent_test_brand.get().strip()
            if not brand or brand == _test_ph:
                messagebox.showwarning("입력 필요", "브랜드명을 입력하세요.", parent=win)
                return
            # 인증 자동 진행
            if not self.is_google_authed():
                if not messagebox.askyesno("Google 인증 필요",
                        "먼저 Google 계정 인증이 필요합니다.\n지금 인증할까요?",
                        parent=win):
                    return
                creds = self.get_google_creds()
                if not creds:
                    messagebox.showerror("인증 실패", "Google 인증 실패", parent=win)
                    return
                update_auth_status()
            # 진행상황 팝업
            progress = tk.Toplevel(win)
            progress.title("검색 중")
            progress.configure(bg="#F5F6F8")
            try:
                self.position_popup(progress, 360, 100)
            except Exception:
                progress.geometry("360x100")
            progress.transient(win)
            progress.grab_set()
            self._bind_esc_close(progress)
            tk.Label(progress, text=f"🔍 '{brand}' 검색 중...",
                     bg="#F5F6F8", font=("맑은 고딕", 10, "bold")).pack(pady=(20, 4))
            progress_lbl = tk.Label(progress, text="시작",
                                      bg="#F5F6F8", fg="#6B7280",
                                      font=("맑은 고딕", 9))
            progress_lbl.pack()
            progress.update()
            def cb(msg):
                try:
                    progress_lbl.config(text=msg)
                    progress.update()
                except Exception:
                    pass
            try:
                results = self.search_sheet_for_brand(brand, md_name=None,
                                                        progress_callback=cb)
            except Exception as e:
                progress.destroy()
                messagebox.showerror("검색 오류", f"검색 중 오류:\n{e}", parent=win)
                return
            progress.destroy()
            # 결과 표시
            if not results:
                messagebox.showinfo("검색 결과", f"❌ '{brand}' 매칭 행 없음", parent=win)
            else:
                lines = [f"✅ {len(results)}건 발견:\n"]
                for r in results[:10]:
                    lines.append(
                        f"• [{r['sheet_name']}] 행{r['row_index']} "
                        f"({r['brand']} / {r['md']})\n  → {r['url'][:50]}..."
                    )
                if len(results) > 10:
                    lines.append(f"... 외 {len(results)-10}건")
                messagebox.showinfo("검색 결과", "\n".join(lines), parent=win)

        tk.Button(test_search_frame, text="검색",
                   command=do_search_test,
                   bg="#F0F2F5", fg="#374151",
                   font=("맑은 고딕", 8, "bold"),
                   relief="flat", padx=8, pady=2,
                   cursor="hand2").pack(side="left")

        # ========== [MD ↔ Slack 매핑 섹션] ==========
        md_card = tk.Frame(win_body, bg="white",
                             highlightthickness=1, highlightbackground="#E5E7EB")
        md_card.pack(fill="x", padx=18, pady=(0, 12))
        tk.Frame(md_card, bg="#3B82F6", width=4).pack(side="left", fill="y")
        md_inner = tk.Frame(md_card, bg="white", padx=14, pady=10)
        md_inner.pack(side="left", fill="both", expand=True)

        # 헤더
        md_head = tk.Frame(md_inner, bg="white")
        md_head.pack(fill="x")
        tk.Label(md_head, text="👥 MD 슬랙 매핑",
                 bg="white", fg="#1A1A1A",
                 font=("맑은 고딕", 10, "bold")).pack(side="left")
        # 활성화
        md_mention_var = tk.BooleanVar(value=s.get("md_mention_enabled", False))
        tk.Checkbutton(md_head, text="멘션 활성화",
                        variable=md_mention_var, bg="white",
                        font=("맑은 고딕", 9),
                        fg="#3B82F6").pack(side="right")

        tk.Label(md_inner,
                 text="MD 이름 → Slack 멤버 ID 매핑. 매핑된 사람은 Slack 메시지에서 진짜 멘션 알림을 받습니다.",
                 bg="white", fg="#6B7280",
                 font=("맑은 고딕", 8),
                 wraplength=460, justify="left", anchor="w").pack(fill="x", pady=(2, 6))

        # 등록된 매핑 리스트
        md_list_outer = tk.Frame(md_inner, bg="#F9FAFB",
                                    highlightthickness=1, highlightbackground="#E5E7EB")
        md_list_outer.pack(fill="both", expand=True, pady=(0, 6))

        md_list_canvas = tk.Canvas(md_list_outer, bg="#F9FAFB",
                                      highlightthickness=0, height=120)
        md_list_canvas.pack(side="left", fill="both", expand=True)
        md_list_scroll = tk.Scrollbar(md_list_outer, command=md_list_canvas.yview)
        md_list_scroll.pack(side="right", fill="y")
        md_list_canvas.configure(yscrollcommand=md_list_scroll.set)
        md_list_inner = tk.Frame(md_list_canvas, bg="#F9FAFB")
        md_list_canvas.create_window((0, 0), window=md_list_inner, anchor="nw")

        def refresh_md_list():
            for w in md_list_inner.winfo_children():
                w.destroy()
            mapping = self.load_md_mapping()
            if not mapping:
                tk.Label(md_list_inner,
                         text="등록된 MD가 없습니다.\n[➕ MD 추가] 버튼으로 등록하세요.",
                         bg="#F9FAFB", fg="#9CA3AF",
                         font=("맑은 고딕", 9), justify="center").pack(pady=20)
            else:
                for md_name in sorted(mapping.keys()):
                    slack_id = mapping[md_name]
                    row = tk.Frame(md_list_inner, bg="white",
                                     highlightthickness=1,
                                     highlightbackground="#E5E7EB")
                    row.pack(fill="x", padx=2, pady=2)
                    info = tk.Frame(row, bg="white")
                    info.pack(side="left", fill="x", expand=True, padx=8, pady=6)
                    tk.Label(info, text=f"👤 {md_name}",
                             bg="white", fg="#1A1A1A",
                             font=("맑은 고딕", 9, "bold"),
                             anchor="w").pack(fill="x")
                    tk.Label(info, text=f"→ {slack_id}",
                             bg="white", fg="#3B82F6",
                             font=("맑은 고딕", 8),
                             anchor="w").pack(fill="x")
                    btn_box = tk.Frame(row, bg="white")
                    btn_box.pack(side="right", padx=4)
                    def make_edit(name):
                        return lambda: self.open_md_mapping_dialog(
                            win, edit_name=name, on_done=refresh_md_list)
                    def make_delete(name):
                        def _del():
                            current = dict(self.load_md_mapping())
                            if name in current:
                                if messagebox.askyesno(
                                        "삭제 확인",
                                        f"'{name}' 매핑을 제거할까요?",
                                        parent=win):
                                    del current[name]
                                    if self.save_md_mapping(current):
                                        refresh_md_list()
                        return _del
                    tk.Button(btn_box, text="✏️",
                               command=make_edit(md_name),
                               bg="white", fg="#3B82F6",
                               font=("맑은 고딕", 9),
                               relief="flat", cursor="hand2",
                               width=2).pack(side="left")
                    tk.Button(btn_box, text="🗑️",
                               command=make_delete(md_name),
                               bg="white", fg="#DC2626",
                               font=("맑은 고딕", 9),
                               relief="flat", cursor="hand2",
                               width=2).pack(side="left")
            md_list_inner.update_idletasks()
            md_list_canvas.configure(scrollregion=md_list_canvas.bbox("all"))

        refresh_md_list()

        # MD 추가 버튼
        tk.Button(md_inner, text="➕ MD 추가",
                   command=lambda: self.open_md_mapping_dialog(
                       win, on_done=refresh_md_list),
                   bg="#3B82F6", fg="white",
                   font=("맑은 고딕", 9, "bold"),
                   relief="flat", padx=12, pady=5,
                   cursor="hand2").pack(anchor="w")

        # 버튼
        btn_frame = tk.Frame(win_body, bg="#F5F6F8")
        btn_frame.pack(fill="x", padx=18, pady=(0, 14))

        def collect():
            return {
                "webhook_url": ent_url.get().strip(),
                "channel_name": ent_channel.get().strip(),
                "cc_mention": ent_cc.get().strip(),
                "enabled": enabled_var.get(),
                "sheet_search_enabled": sheet_enabled_var.get(),
                "md_mention_enabled": md_mention_var.get(),
            }

        def test_connection():
            url = ent_url.get().strip()
            if not url:
                result_lbl.config(text="❌ Webhook URL을 먼저 입력하세요.", fg="#DC2626")
                return
            if not url.startswith("https://hooks.slack.com/"):
                result_lbl.config(text="⚠️ Webhook URL 형식이 이상합니다.", fg="#D97706")
                return
            result_lbl.config(text="🔄 테스트 메시지 전송 중...", fg="#6B7280")
            win.update()
            ok, msg = self._test_slack_connection(url)
            if ok:
                result_lbl.config(text=msg, fg="#16A34A")
            else:
                result_lbl.config(text=f"❌ {msg}", fg="#DC2626")

        def do_save():
            cfg = collect()
            if cfg["enabled"] and not cfg["webhook_url"]:
                messagebox.showwarning("입력 누락",
                    "활성화하려면 Webhook URL이 필요합니다.",
                    parent=win)
                return
            if self.save_slack_settings(cfg):
                messagebox.showinfo("저장 완료",
                    "✅ Slack 설정이 저장되었습니다." +
                    ("\n\nJira 상신 시 자동으로 슬랙 알림이 갑니다." if cfg["enabled"] else ""),
                    parent=win)
                win.destroy()

        tk.Button(btn_frame, text="🧪 테스트 메시지",
                   command=test_connection,
                   bg="#F0F2F5", fg="#374151",
                   font=("맑은 고딕", 9, "bold"),
                   relief="flat", padx=14, pady=6,
                   cursor="hand2").pack(side="left")

        tk.Button(btn_frame, text="💾 저장",
                   command=do_save,
                   bg="#4A154B", fg="white",
                   font=("맑은 고딕", 9, "bold"),
                   relief="flat", padx=14, pady=6,
                   cursor="hand2").pack(side="right")
        tk.Button(btn_frame, text="취소",
                   command=win.destroy,
                   bg="white", fg="#666",
                   font=("맑은 고딕", 9),
                   relief="flat", padx=14, pady=6,
                   cursor="hand2",
                   highlightthickness=1, highlightbackground="#DDD").pack(side="right", padx=(0, 6))

        # 모든 위젯이 추가된 후, 자손에게도 휠 바인딩 적용
        try:
            win.update_idletasks()
            self._slack_wheel_rebind(win_body)
        except Exception:
            pass
