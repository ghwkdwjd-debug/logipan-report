# -*- coding: utf-8 -*-
"""
로지판(LogiPan) - Jira 통합 모듈
====================================

LogiPan.py에서 분리된 Jira 연동 관련 기능 모음. (모듈화 2단계)

사용법 (Mixin 패턴):
    from jira_integration import JiraIntegrationMixin

    class LogiPanApp(SlackIntegrationMixin, JiraIntegrationMixin):
        ...

포함된 기능 (총 7개 메서드):
    [Jira 설정 저장/로드]
      - load_jira_settings, save_jira_settings
    [Jira API]
      - create_jira_ticket
      - _upload_jira_attachment
      - _test_jira_connection
      - _test_jira_connection_OLD_unused (사용 안 함, 참고용 보존)
    [UI 팝업]
      - open_jira_settings

LogiPan 본체가 제공해야 하는 속성:
    self.root, self.config_path

LogiPan 본체가 제공해야 하는 메서드:
    self.position_popup(win, w, h)
    self._bind_esc_close(win)
    self._bind_mousewheel(canvas, container)

Slack 모듈과의 의존성: 없음 (완전 독립)

원본 위치: LogiPan.py 라인 644~1131 (1단계 분리 후 기준)
분리 일자: 다음 세션 (모듈화 2단계)
"""

import os
import json
import urllib.request
import urllib.error
import tkinter as tk
from tkinter import messagebox


class JiraIntegrationMixin:
    """Jira 연동 Mixin.

    LogiPanApp이 이 Mixin을 상속하면 self.open_jira_settings() 등을
    그대로 호출 가능. 원본 LogiPan.py와 100% 동작 동일.
    """

    # ========== [Jira 연동] ==========
    def load_jira_settings(self):
        """로컬 config 파일에서 Jira 설정 로드"""
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, "r", encoding="utf-8") as f:
                    cfg = json.load(f)
                return cfg.get("jira_settings", {})
        except Exception:
            pass
        return {}

    def save_jira_settings(self, settings):
        """로컬 config 파일에 Jira 설정 저장"""
        try:
            cfg = {}
            if os.path.exists(self.config_path):
                try:
                    with open(self.config_path, "r", encoding="utf-8") as f:
                        cfg = json.load(f)
                except Exception:
                    cfg = {}
            cfg["jira_settings"] = settings
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            messagebox.showerror("저장 실패", f"Jira 설정 저장 실패: {e}")
            return False

    def open_jira_settings(self):
        """Jira 연동 설정 팝업"""
        win = tk.Toplevel(self.root)
        win.title("⚙️ Jira 연동 설정")
        win.configure(bg="#F5F6F8")
        try:
            self.position_popup(win, 480, 620)
        except Exception:
            win.geometry("480x620")
        win.transient(self.root)
        win.grab_set()
        self._bind_esc_close(win)

        # ===== 스크롤 컨테이너 =====
        scroll_outer = tk.Frame(win, bg="#F5F6F8")
        scroll_outer.pack(fill="both", expand=True)
        scroll_canvas = tk.Canvas(scroll_outer, bg="#F5F6F8",
                                     highlightthickness=0, bd=0)
        scroll_canvas.pack(side="left", fill="both", expand=True)
        scroll_bar = tk.Scrollbar(scroll_outer, command=scroll_canvas.yview)
        scroll_bar.pack(side="right", fill="y")
        scroll_canvas.configure(yscrollcommand=scroll_bar.set)
        win_body = tk.Frame(scroll_canvas, bg="#F5F6F8")
        body_window = scroll_canvas.create_window((0, 0), window=win_body, anchor="nw")

        def _on_body_resize(event):
            scroll_canvas.configure(scrollregion=scroll_canvas.bbox("all"))
        def _on_canvas_resize(event):
            scroll_canvas.itemconfig(body_window, width=event.width)
        win_body.bind("<Configure>", _on_body_resize)
        scroll_canvas.bind("<Configure>", _on_canvas_resize)

        # 마우스 휠 - 캔버스/win_body에만 (전역 X)
        rebind_wheel_jira = self._bind_mousewheel(scroll_canvas, win_body)
        self._jira_wheel_rebind = rebind_wheel_jira

        s = self.load_jira_settings()

        # 헤더
        head = tk.Frame(win_body, bg="#F5F6F8")
        head.pack(fill="x", padx=18, pady=(14, 4))
        tk.Label(head, text="⚙️", font=("맑은 고딕", 18), bg="#F5F6F8").pack(side="left", padx=(0, 6))
        tk.Label(head, text="Jira 연동 설정",
                 font=("맑은 고딕", 14, "bold"),
                 bg="#F5F6F8", fg="#1A1A1A").pack(side="left")
        tk.Label(win_body, text="입고 CSV 저장 시 자동으로 Jira 티켓 생성 (CSV 첨부)",
                 bg="#F5F6F8", fg="#666",
                 font=("맑은 고딕", 8)).pack(padx=18, anchor="w")

        # 카드
        card = tk.Frame(win_body, bg="white",
                          highlightthickness=1, highlightbackground="#E5E7EB")
        card.pack(fill="x", padx=18, pady=12)
        tk.Frame(card, bg="#3B82F6", width=4).pack(side="left", fill="y")
        inner = tk.Frame(card, bg="white", padx=14, pady=12)
        inner.pack(side="left", fill="both", expand=True)

        # 입력 필드들
        entries = {}

        def add_field(label, key, default="", show=None, hint=""):
            frame = tk.Frame(inner, bg="white")
            frame.pack(fill="x", pady=4)
            tk.Label(frame, text=label, bg="white",
                     font=("맑은 고딕", 9, "bold"),
                     fg="#374151").pack(anchor="w")
            ent = tk.Entry(frame, font=("맑은 고딕", 10),
                            bd=1, relief="solid",
                            highlightthickness=0)
            if show:
                ent.config(show=show)
            ent.pack(fill="x", ipady=4, pady=(2, 0))
            ent.insert(0, s.get(key, default))
            entries[key] = ent
            if hint:
                tk.Label(frame, text=hint,
                         bg="white", fg="#9CA3AF",
                         font=("맑은 고딕", 8),
                         anchor="w").pack(fill="x")
            return ent

        add_field("Atlassian 도메인",
                   "domain",
                   hint="예: yourcompany.atlassian.net (https:// 빼고)")
        add_field("본인 이메일", "email",
                   hint="Atlassian 로그인 이메일")
        add_field("API Token", "api_token", show="•",
                   hint="https://id.atlassian.com/manage-profile/security/api-tokens 에서 생성")
        add_field("프로젝트 키", "project_key",
                   hint="예: LOG, OPS (대문자, Jira에서 확인)")
        add_field("담당자 이름", "assignee_name",
                   hint="본문에 표시될 이름 (예: 김상품)")
        add_field("담당자 Account ID", "assignee_id",
                   hint="비워두면 미지정. Jira Profile에서 Account ID 확인")
        add_field("이슈 타입", "issue_type", default="Task",
                   hint="Task, Bug, Story 등")

        # 활성화 체크박스
        enabled_var = tk.BooleanVar(value=s.get("enabled", False))
        chk_frame = tk.Frame(inner, bg="white")
        chk_frame.pack(fill="x", pady=(10, 0))
        tk.Checkbutton(chk_frame, text="✅ 입고 CSV 저장 시 자동 티켓 생성",
                        variable=enabled_var, bg="white",
                        font=("맑은 고딕", 9, "bold"),
                        fg="#065F46").pack(anchor="w")

        # 결과 표시 라벨
        result_lbl = tk.Label(inner, text="", bg="white",
                                font=("맑은 고딕", 8),
                                wraplength=360, justify="left")
        result_lbl.pack(fill="x", pady=(8, 0))

        # 버튼
        btn_frame = tk.Frame(win_body, bg="#F5F6F8")
        btn_frame.pack(fill="x", padx=18, pady=(0, 14))

        def collect():
            return {
                "domain": entries["domain"].get().strip().rstrip('/'),
                "email": entries["email"].get().strip(),
                "api_token": entries["api_token"].get().strip(),
                "project_key": entries["project_key"].get().strip().upper(),
                "assignee_name": entries["assignee_name"].get().strip(),
                "assignee_id": entries["assignee_id"].get().strip(),
                "issue_type": entries["issue_type"].get().strip() or "Task",
                "enabled": enabled_var.get(),
            }

        def test_connection():
            cfg = collect()
            result_lbl.config(text="🔄 연결 테스트 중...", fg="#6B7280")
            win.update()
            ok, msg = self._test_jira_connection(cfg)
            if ok:
                result_lbl.config(text=f"✅ {msg}", fg="#16A34A")
            else:
                result_lbl.config(text=f"❌ {msg}", fg="#DC2626")

        def do_save():
            cfg = collect()
            if cfg["enabled"]:
                # 활성화 시 필수 항목 체크
                missing = [k for k in ["domain", "email", "api_token", "project_key"]
                           if not cfg[k]]
                if missing:
                    messagebox.showwarning("입력 누락",
                        f"활성화하려면 다음 항목이 필요합니다:\n{', '.join(missing)}",
                        parent=win)
                    return
            if self.save_jira_settings(cfg):
                messagebox.showinfo("저장 완료",
                    "✅ Jira 설정이 저장되었습니다." +
                    ("\n\n다음 입고 CSV 저장 시 자동으로 티켓이 생성됩니다." if cfg["enabled"] else ""),
                    parent=win)
                win.destroy()

        tk.Button(btn_frame, text="🧪 연결 테스트",
                   command=test_connection,
                   bg="#F0F2F5", fg="#374151",
                   font=("맑은 고딕", 9, "bold"),
                   relief="flat", padx=14, pady=6,
                   cursor="hand2").pack(side="left")

        tk.Button(btn_frame, text="💾 저장",
                   command=do_save,
                   bg="#1877F2", fg="white",
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

        # 자손에게 휠 바인딩
        try:
            win.update_idletasks()
            self._jira_wheel_rebind(win_body)
        except Exception:
            pass

    def _test_jira_connection(self, cfg):
        """Jira 연결 테스트 - 단계별로 진단
        1) /myself - 인증 자체가 되는지 (401 = 토큰 문제, 403 = 권한 부족)
        2) /project/{KEY} - 프로젝트 접근 가능한지
        """
        try:
            import base64
            domain = cfg.get("domain", "")
            email = cfg.get("email", "")
            token = cfg.get("api_token", "")
            project = cfg.get("project_key", "")
            if not all([domain, token, project]):
                return False, "도메인/토큰/프로젝트 키 필수입니다."

            # 인증 헤더 후보들
            auth_headers = []
            if email:
                auth_str = f"{email}:{token}"
                auth_b64 = base64.b64encode(auth_str.encode()).decode()
                auth_headers.append(('Basic', f'Basic {auth_b64}'))
            auth_headers.append(('Bearer', f'Bearer {token}'))

            api_versions = ['3', '2']

            # === 1단계: /myself로 인증 테스트 ===
            myself_results = []
            for auth_name, auth_value in auth_headers:
                for api_v in api_versions:
                    url = f"https://{domain}/rest/api/{api_v}/myself"
                    try:
                        req = urllib.request.Request(url, headers={
                            'Authorization': auth_value,
                            'Accept': 'application/json'
                        })
                        with urllib.request.urlopen(req, timeout=10) as resp:
                            data = json.loads(resp.read())
                        # 인증 성공! 본인 정보
                        my_name = data.get('displayName') or data.get('name', '?')
                        my_id = data.get('accountId') or data.get('key') or data.get('name', '?')
                        myself_results.append((auth_name, api_v, my_name, my_id, None))
                        break  # 이 auth로 일단 인증은 됨
                    except urllib.error.HTTPError as e:
                        try:
                            err_body = e.read().decode()[:200]
                        except:
                            err_body = ""
                        myself_results.append((auth_name, api_v, None, None,
                                              f"{e.code}: {err_body or e.reason}"))
                    except Exception as e:
                        myself_results.append((auth_name, api_v, None, None, str(e)[:100]))

            # 성공한 인증 찾기
            auth_success = [r for r in myself_results if r[2] is not None]
            if not auth_success:
                # 인증 자체가 다 실패
                msgs = [f"  {r[0]} api/{r[1]}: {r[4]}" for r in myself_results]
                return False, "인증 실패 (모든 시도):\n" + "\n".join(msgs)

            # 인증 성공한 첫 번째 조합으로 프로젝트 접근 테스트
            best_auth_name, best_api_v, my_name, my_id, _ = auth_success[0]
            best_auth_value = next(av for an, av in auth_headers if an == best_auth_name)

            # === 2단계: /project/{KEY}로 프로젝트 권한 테스트 ===
            url = f"https://{domain}/rest/api/{best_api_v}/project/{project}"
            try:
                req = urllib.request.Request(url, headers={
                    'Authorization': best_auth_value,
                    'Accept': 'application/json'
                })
                with urllib.request.urlopen(req, timeout=10) as resp:
                    data = json.loads(resp.read())
                return True, (f"✅ 연결 성공!\n"
                               f"  로그인: {my_name} (ID: {my_id})\n"
                               f"  프로젝트: {data.get('name', project)}\n"
                               f"  ({best_auth_name}, api/{best_api_v})")
            except urllib.error.HTTPError as e:
                try:
                    err_body = e.read().decode()[:200]
                except:
                    err_body = ""
                if e.code == 403:
                    return False, (f"⚠️ 인증은 OK ({my_name}으로 로그인됨)\n"
                                    f"하지만 '{project}' 프로젝트에 API 접근 권한 없음\n"
                                    f"(403 Forbidden)\n\n"
                                    f"💡 회사 IT/보안팀에 문의 필요:\n"
                                    f"   - REST API 접근 권한\n"
                                    f"   - {project} 프로젝트 권한")
                elif e.code == 404:
                    return False, f"프로젝트 '{project}' 못 찾음. 키 확인하세요."
                else:
                    return False, f"인증 OK인데 프로젝트 접근 HTTP {e.code}: {err_body}"
        except Exception as e:
            return False, f"연결 실패: {e}"

    def _test_jira_connection_OLD_unused(self, cfg):
        """예전 버전 (참고용)"""
        try:
            import base64
            domain = cfg.get("domain", "")
            email = cfg.get("email", "")
            token = cfg.get("api_token", "")
            project = cfg.get("project_key", "")
            if not all([domain, token, project]):
                return False, "도메인/토큰/프로젝트 키 필수입니다."

            # 인증 헤더 후보들
            auth_headers = []
            if email:
                # Basic auth (Cloud + 일부 Server)
                auth_str = f"{email}:{token}"
                auth_b64 = base64.b64encode(auth_str.encode()).decode()
                auth_headers.append(('Basic auth', f'Basic {auth_b64}'))
            # Bearer token (Server PAT)
            auth_headers.append(('Bearer token', f'Bearer {token}'))

            # API 버전 후보들 (Cloud는 3, Server는 2)
            api_versions = ['3', '2']

            last_error = ""
            for auth_name, auth_value in auth_headers:
                for api_v in api_versions:
                    url = f"https://{domain}/rest/api/{api_v}/project/{project}"
                    try:
                        req = urllib.request.Request(url, headers={
                            'Authorization': auth_value,
                            'Accept': 'application/json'
                        })
                        with urllib.request.urlopen(req, timeout=10) as resp:
                            data = json.loads(resp.read())
                        # 성공! 어떤 조합이 통했는지 저장
                        cfg['_api_version'] = api_v
                        cfg['_auth_method'] = auth_name
                        # 임시로 저장 (실제 저장은 사용자가 [저장] 누를 때)
                        return True, (f"연결 성공! 프로젝트: {data.get('name', project)}\n"
                                       f"(API v{api_v}, {auth_name})")
                    except urllib.error.HTTPError as e:
                        try:
                            err_body = e.read().decode()[:300]
                        except:
                            err_body = ""
                        last_error = f"[{auth_name} api/{api_v}] HTTP {e.code}: {err_body or e.reason}"
                        continue
                    except Exception as e:
                        last_error = f"[{auth_name} api/{api_v}] {e}"
                        continue

            return False, f"모든 시도 실패. 마지막: {last_error}"
        except Exception as e:
            return False, f"연결 실패: {e}"

    def create_jira_ticket(self, title, description, attachment_path=None):
        """Jira 티켓 생성 + 첨부파일 업로드.
        Returns: (성공여부, 결과메시지/URL)"""
        cfg = self.load_jira_settings()
        if not cfg.get("enabled", False):
            return None, "Jira 비활성화됨"

        try:
            import base64
            domain = cfg["domain"]
            email = cfg.get("email", "")
            token = cfg["api_token"]
            project = cfg["project_key"]
            assignee = cfg.get("assignee_id", "")
            issue_type = cfg.get("issue_type", "Task")

            # 인증 헤더 후보들
            auth_headers = []
            if email:
                auth_str = f"{email}:{token}"
                auth_b64 = base64.b64encode(auth_str.encode()).decode()
                auth_headers.append(('Basic', f'Basic {auth_b64}'))
            auth_headers.append(('Bearer', f'Bearer {token}'))

            # API 버전 후보들
            api_versions = ['3', '2']

            last_error = ""
            for auth_name, auth_value in auth_headers:
                for api_v in api_versions:
                    issue_url = f"https://{domain}/rest/api/{api_v}/issue"
                    # API v3는 Atlassian Document Format, v2는 plain text
                    if api_v == '3':
                        description_field = {
                            "type": "doc",
                            "version": 1,
                            "content": [{
                                "type": "paragraph",
                                "content": [{
                                    "type": "text",
                                    "text": description
                                }]
                            }]
                        }
                    else:
                        description_field = description  # plain text

                    payload = {
                        "fields": {
                            "project": {"key": project},
                            "summary": title,
                            "issuetype": {"name": issue_type},
                            "description": description_field
                        }
                    }
                    if assignee:
                        # Cloud: accountId / Server: name
                        if api_v == '3':
                            payload["fields"]["assignee"] = {"accountId": assignee}
                        else:
                            payload["fields"]["assignee"] = {"name": assignee}

                    data = json.dumps(payload).encode('utf-8')
                    try:
                        req = urllib.request.Request(issue_url, data=data, headers={
                            'Authorization': auth_value,
                            'Content-Type': 'application/json',
                            'Accept': 'application/json'
                        }, method='POST')
                        with urllib.request.urlopen(req, timeout=15) as resp:
                            result = json.loads(resp.read())
                        issue_key = result.get('key')
                        if not issue_key:
                            last_error = "티켓 키 없음"
                            continue

                        # 성공! 첨부파일 업로드
                        if attachment_path and os.path.exists(attachment_path):
                            try:
                                self._upload_jira_attachment(domain, auth_value, issue_key, attachment_path, api_v)
                            except Exception as e:
                                print(f"⚠️ 첨부파일 업로드 실패: {e}")

                        ticket_url = f"https://{domain}/browse/{issue_key}"
                        return True, ticket_url
                    except urllib.error.HTTPError as e:
                        try:
                            err_body = e.read().decode()
                        except:
                            err_body = ""
                        last_error = f"[{auth_name} api/{api_v}] HTTP {e.code}: {err_body[:200]}"
                        continue
                    except Exception as e:
                        last_error = f"[{auth_name} api/{api_v}] {e}"
                        continue

            return False, f"모든 시도 실패. 마지막: {last_error}"
        except Exception as e:
            return False, f"오류: {e}"

    def _upload_jira_attachment(self, domain, auth_header, issue_key, file_path, api_v='3'):
        """Jira 티켓에 파일 첨부 (multipart/form-data 직접 구성)"""
        url = f"https://{domain}/rest/api/{api_v}/issue/{issue_key}/attachments"
        boundary = '----LogiPanBoundary' + str(int(__import__('time').time()))
        filename = os.path.basename(file_path)

        with open(file_path, 'rb') as f:
            file_content = f.read()

        body = b''
        body += f'--{boundary}\r\n'.encode()
        body += f'Content-Disposition: form-data; name="file"; filename="{filename}"\r\n'.encode('utf-8')
        body += b'Content-Type: text/csv\r\n\r\n'
        body += file_content
        body += f'\r\n--{boundary}--\r\n'.encode()

        req = urllib.request.Request(url, data=body, headers={
            'Authorization': auth_header,
            'X-Atlassian-Token': 'no-check',
            'Content-Type': f'multipart/form-data; boundary={boundary}',
        }, method='POST')
        with urllib.request.urlopen(req, timeout=30) as resp:
            return resp.read()

    # ─────────────────────────────────────────────────────────────
    # [추가] Jira 티켓 검색 - [출고] 같은 prefix로 필터링해서 조회
    # ─────────────────────────────────────────────────────────────
    def search_jira_tickets(self, prefix="[출고]", project_key=None, max_results=50, status_filter=None):
        """JQL로 Jira 티켓 검색.
        - prefix: summary에 포함된 prefix (예: "[출고]")
        - project_key: 프로젝트 키 (None이면 설정값 사용)
        - max_results: 최대 가져올 개수
        - status_filter: ["미해결", "처리중"] 같은 상태 필터 (None이면 전체)

        Returns: (성공여부, 결과리스트 또는 에러메시지)
            결과 항목: {
                'key': 'PROJ-123',
                'summary': '[출고] 매장명 ...',
                'status': '처리중',
                'assignee': '담당자명' or None,
                'reporter': '작성자명',
                'created': '2026-06-04T10:30:00.000+0900',
                'updated': '...',
                'description': '본문 텍스트',
                'url': 'https://domain/browse/PROJ-123',
            }
        """
        cfg = self.load_jira_settings()
        if not cfg.get("enabled", False):
            return False, "Jira 비활성화됨"

        try:
            import base64, json
            import urllib.request, urllib.parse, urllib.error
            domain = cfg["domain"]
            email = cfg.get("email", "")
            token = cfg["api_token"]
            project = project_key or cfg.get("project_key", "")
            if not project:
                return False, "프로젝트 키가 설정되지 않았습니다."

            # JQL 조립
            # 대괄호 "[출고]"는 JQL 텍스트 검색에서 특수문자 취급되어 까다로움.
            # → 대괄호 빼고 "출고"로 검색 (어차피 [출고] 포함하는 티켓이 매칭됨)
            jql_parts = [f'project = "{project}"']
            if prefix:
                # [출고] → 출고 만 추출 (대괄호 떼기)
                safe_text = prefix.replace('[', '').replace(']', '').replace('"', '\\"').strip()
                if safe_text:
                    jql_parts.append(f'summary ~ "{safe_text}"')
            if status_filter:
                statuses = ', '.join(f'"{s}"' for s in status_filter)
                jql_parts.append(f'status in ({statuses})')
            jql_body = ' AND '.join(jql_parts)
            jql = jql_body + ' ORDER BY created DESC'

            # 인증 헤더 - Basic만 사용 (Cloud는 Bearer 거부)
            # email 없으면 token만 (Server 환경)
            if email:
                auth_str = f"{email}:{token}"
                auth_b64 = base64.b64encode(auth_str.encode()).decode()
                auth_header = f'Basic {auth_b64}'
            else:
                auth_header = f'Bearer {token}'

            # API v3는 /search/jql (신규), v2는 /search (구형 호환)
            api_candidates = [('3', 'search/jql'), ('2', 'search')]
            last_error = ""

            for api_v, search_path in api_candidates:
                try:
                    search_url = f"https://{domain}/rest/api/{api_v}/{search_path}"
                    body = {
                        'jql': jql,
                        'maxResults': max_results,
                        'fields': ['summary', 'status', 'assignee', 'reporter',
                                    'created', 'updated', 'description'],
                    }
                    body_bytes = json.dumps(body).encode('utf-8')

                    req = urllib.request.Request(search_url, data=body_bytes, headers={
                        'Authorization': auth_header,
                        'Content-Type': 'application/json',
                        'Accept': 'application/json',
                    }, method='POST')
                    with urllib.request.urlopen(req, timeout=30) as resp:
                        data = json.loads(resp.read().decode())

                    results = []
                    for issue in data.get('issues', []):
                        fields = issue.get('fields', {})
                        assignee = fields.get('assignee')
                        reporter = fields.get('reporter')
                        status = fields.get('status', {})

                        desc_raw = fields.get('description')
                        desc_text = self._extract_description_text(desc_raw)

                        results.append({
                            'key': issue.get('key'),
                            'summary': fields.get('summary', ''),
                            'status': status.get('name', '') if status else '',
                            'assignee': (assignee.get('displayName') if assignee else None),
                            'assignee_id': (assignee.get('accountId') or assignee.get('name')) if assignee else None,
                            'reporter': (reporter.get('displayName') if reporter else None),
                            'created': fields.get('created', ''),
                            'updated': fields.get('updated', ''),
                            'description': desc_text,
                            'url': f"https://{domain}/browse/{issue.get('key')}",
                        })
                    return True, results
                except urllib.error.HTTPError as e:
                    last_error = f"HTTP {e.code}: {e.reason}"
                    try:
                        err_body = e.read().decode()
                        last_error += f" - {err_body[:300]}"
                    except Exception:
                        pass
                    print(f"[Jira 검색] API v{api_v}/{search_path} 실패: {last_error}")
                    continue
                except Exception as e:
                    last_error = str(e)
                    print(f"[Jira 검색] API v{api_v}/{search_path} 예외: {e}")
                    continue

            return False, f"검색 실패: {last_error}"
        except Exception as e:
            import traceback
            print(traceback.format_exc())
            return False, str(e)

    def _extract_description_text(self, desc):
        """Jira description을 plain text로 변환 (API v3 ADF 또는 v2 string 둘 다 지원)"""
        if not desc:
            return ''
        if isinstance(desc, str):
            return desc
        if isinstance(desc, dict):
            # ADF (Atlassian Document Format) - 재귀로 text 추출
            texts = []
            def _walk(node):
                if isinstance(node, dict):
                    if node.get('type') == 'text':
                        texts.append(node.get('text', ''))
                    elif node.get('type') == 'hardBreak':
                        texts.append('\n')
                    for child in node.get('content', []) or []:
                        _walk(child)
                    # paragraph 끝나면 줄바꿈
                    if node.get('type') == 'paragraph':
                        texts.append('\n')
                elif isinstance(node, list):
                    for c in node:
                        _walk(c)
            _walk(desc)
            return ''.join(texts).strip()
        return str(desc)
