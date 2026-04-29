import pandas as pd
from datetime import datetime
import os
import io
import json
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import re
import firebase_admin
from firebase_admin import credentials, firestore, messaging

class LogiPanApp:
    def __init__(self, root):
        self.root = root
        self.root.title("로지판 (Logi-Pan) v11.0 - 통합 물류 파트너")

        # --- 구글 비밀기지 연결 시작 ---
        try:
            # 실행 위치와 상관없이 key.json을 찾도록 절대경로로 접근
            base_path = os.path.dirname(os.path.abspath(__file__))
            key_path = os.path.join(base_path, "key.json")

            # 이미 초기화된 경우 재초기화하지 않음 (중복 init 방지)
            if not firebase_admin._apps:
                cred = credentials.Certificate(key_path)
                firebase_admin.initialize_app(cred)
            self.db = firestore.client()
            print("✅ 구글 비밀기지 연결 성공!")
        except Exception as e:
            print(f"❌ 연결 실패: {e}")
        # --- 구글 비밀기지 연결 끝 ---

        # [창 크기 설정]
        # 노트북 화면 기준으로 컴팩트하게. 화면이 작으면 90%까지만.
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        width = min(820, int(sw * 0.9))
        height = min(720, int(sh * 0.85))
        self.root.geometry(f"{width}x{height}+{(sw-width)//2}+{(sh-height)//2}")
        # 버튼이 잘리지 않을 최소 크기
        self.root.minsize(760, 640)
        self.root.configure(bg="#f5f5f5")        
        self.desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        # [추가] 사용자 설정 파일 경로 (저장 위치를 기억해두는 곳)
        self.config_path = os.path.join(os.path.expanduser("~"), ".logipan_config.json")
        self.save_dir = self.load_save_dir()
        if not os.path.exists(self.save_dir):
            try:
                os.makedirs(self.save_dir)
            except Exception:
                # 만약 저장된 경로가 사라졌거나 접근 불가면 기본값으로 폴백
                self.save_dir = os.path.join(self.desktop, "로지판_작업결과")
                if not os.path.exists(self.save_dir): os.makedirs(self.save_dir)

        self.enter_count = 0
        self.is_matched = False 
        self.selected_store = tk.StringVar(value="성수")
        self.selected_type = tk.StringVar(value="매장출고")
        self.s_btns = {}
        self.t_btns = {}
        self.moms_files = {"send": "", "master": "", "out_list": ""}
        self.chk_files = {"target": "", "master": ""}
        self.filter_var = tk.StringVar(value="전체")
        self.search_var = tk.StringVar()
        # [수정] current_user를 config 파일에서 로드 (없으면 기본값 "장정호")
        self.current_user = self.load_user_name()
        self.start_realtime_listener()

        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure("TNotebook.Tab", padding=[8, 4], font=("맑은 고딕", 9, "bold"))
        self.style.configure("Treeview", font=("맑은 고딕", 10), rowheight=25)
        self.style.configure("Treeview.Heading", font=("맑은 고딕", 10, "bold"))

        # [수정] 하단 상태바를 먼저 pack해서 자리부터 확보 (노트북이 expand=both라 나중에 pack하면 가려짐)
        status_frame = tk.Frame(self.root, bd=1, relief=tk.SUNKEN, bg="#eeeeee", height=24)
        status_frame.pack(side=tk.BOTTOM, fill=tk.X)
        status_frame.pack_propagate(False)  # 높이 고정
        self.status = tk.Label(status_frame, text=f" 📂 저장 위치: {self.save_dir}",
                               anchor=tk.W, bg="#eeeeee", font=("맑은 고딕", 8))
        self.status.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(status_frame, text="📁 경로 변경", command=self.change_save_dir,
                  bg="#ffffff", font=("맑은 고딕", 8, "bold"), bd=1, relief="solid",
                  padx=8, pady=0, cursor="hand2").pack(side=tk.RIGHT, padx=3, pady=1)

        self.nb = ttk.Notebook(self.root)
        self.nb.pack(expand=True, fill="both", padx=5, pady=5)

        self.t_in = ttk.Frame(self.nb); self.nb.add(self.t_in, text="    입고    ")
        self.t_out = ttk.Frame(self.nb); self.nb.add(self.t_out, text="    출고    ")
        self.t_mom = ttk.Frame(self.nb); self.nb.add(self.t_mom, text="    맘스    ")
        self.t_end = ttk.Frame(self.nb); self.nb.add(self.t_end, text="  마감재고  ")
        self.t_chk = ttk.Frame(self.nb); self.nb.add(self.t_chk, text="  재고파악  ")
        self.t_field = tk.Frame(self.nb); self.nb.add(self.t_field, text="   작업보고   ")
        self.t_board = tk.Frame(self.nb); self.nb.add(self.t_board, text="  공지/소통  ")

        self.setup_inbound()
        self.setup_outbound()
        self.setup_moms_v86()
        self.setup_closing_stock()
        self.setup_inventory_check_v95()
        self.setup_field_comm(self.t_field)
        self.setup_board_system(self.t_board)

        # [추가] 탭 알림 시스템 초기화 (모든 탭 만든 다음에)
        self.setup_tab_alert_system()

    # --- [데이터 처리 공통 로직] ---
    def check_double_enter(self, event):
        self.enter_count += 1
        if self.enter_count >= 2:
            self.txt_in_scan.focus_set(); self.enter_count = 0; return "break"

    def count_total_qty(self, text_widget, label_widget, base_text):
        """[수정] parse_logi_data와 같은 규칙으로 합계 계산 (5열 형식 정확히 인식)."""
        raw_text = text_widget.get("1.0", tk.END).strip()
        if not raw_text:
            label_widget.config(text=f"{base_text} (0개)")
            return

        try:
            # parse_logi_data를 사용해서 정확한 수량 계산
            # 단, 이건 사이드이펙트(self.last_scan_detail)를 만들 수 있어서
            # 임시로 보존하고 복원
            saved_detail = getattr(self, 'last_scan_detail', None)
            df = self.parse_logi_data(text_widget)
            # 사이드이펙트 복원 (대조 실행 전이라면 last_scan_detail이 흔들리면 안됨)
            if saved_detail is not None:
                self.last_scan_detail = saved_detail
            elif hasattr(self, 'last_scan_detail') and text_widget is not getattr(self, 'txt_in_scan', None):
                # 스캔칸이 아닌데 채워졌으면 제거
                pass

            total = int(df['수량'].sum()) if df is not None else 0
        except Exception:
            total = 0

        label_widget.config(text=f"{base_text} ({total}개)")

    def normalize_barcode(self, code):
        """바코드 정규화: 공백/하이픈/언더스코어 제거 + 대문자 통일.
        '880-1234-567-890'과 '8801234567890'을 같은 것으로 인식하기 위함."""
        if code is None: return ''
        s = str(code).strip().upper()
        # 하이픈/언더스코어/슬래시/마침표(소수점 .0 끝부분만 제외)/공백 제거
        s = re.sub(r'[-_/\s]', '', s)
        # 끝의 .0 제거 (엑셀에서 숫자로 들어온 경우)
        if s.endswith('.0'):
            s = s[:-2]
        return s

    def _is_qty_token(self, tok):
        """수량 토큰인가? 0 이상의 정수 문자열."""
        return tok.isdigit()

    def _looks_like_location(self, tok):
        """로케이션 형식 문자열인가? 'DD-02-06-03' 또는 '00-00-00-00' 같은 패턴."""
        # 하이픈으로 구분된 토큰이고 영숫자 그룹이 3개 이상이면 로케이션
        return bool(re.match(r'^[A-Za-z0-9]+(-[A-Za-z0-9]+){2,}$', tok))

    def parse_logi_data(self, text_widget):
        """입력을 자동 감지해서 DataFrame으로 반환.

        지원 형식:
        - 바코드만: BARCODE1\nBARCODE2 → 각 1개로
        - 바코드+수량: BARCODE1 5\nBARCODE2 3 → 그대로
        - 5열 (스캔칸 전용): BARCODE 정상수 불량수 정상로케 불량로케
            → 같은 바코드는 (정상+불량) 합쳐서 표시 (대조용)
            → 정상/불량 분리 데이터는 self.last_scan_detail에 보존 (입고파일 생성용)

        반환: DataFrame[바코드, 수량] (정규화된 바코드, 수량 합계)
        """
        raw_text = text_widget.get("1.0", tk.END).strip()
        if not raw_text: return None

        # 한 줄씩 처리. 한 줄 안에 5열 형식이 있으면 그렇게, 아니면 토큰 단위로.
        rows = []  # [(바코드, 정상수, 불량수, 정상로케, 불량로케)] - 5열 또는 (바코드, 수량, 0, '', '')
        token_buffer = []  # 5열이 아닌 단순 형식의 토큰 누적

        for line in raw_text.split('\n'):
            line = line.strip()
            if not line: continue
            # 탭과 공백 모두 구분자로
            tokens = re.split(r'[\t ]+', line)
            tokens = [t for t in tokens if t]
            if not tokens: continue

            # 5열 형식 판별: 토큰 5개 + 2,3번째가 숫자 + 4,5번째가 로케이션처럼 생김
            is_5col = (
                len(tokens) >= 5
                and self._is_qty_token(tokens[1])
                and self._is_qty_token(tokens[2])
                and self._looks_like_location(tokens[3])
                and self._looks_like_location(tokens[4])
            )

            if is_5col:
                # 5열로 인식
                bc = self.normalize_barcode(tokens[0])
                rows.append({
                    '바코드': bc,
                    '정상수': int(tokens[1]),
                    '불량수': int(tokens[2]),
                    '정상로케': tokens[3],
                    '불량로케': tokens[4],
                })
            else:
                # 줄 안의 토큰을 일반 형식으로 누적 (한 줄에 여러 바코드도 가능)
                token_buffer.extend(tokens)

        # 일반 형식 처리 (바코드 / 바코드+수량 패턴)
        i = 0
        while i < len(token_buffer):
            token = token_buffer[i].strip()
            if not token: i += 1; continue
            bc = self.normalize_barcode(token)
            if i + 1 < len(token_buffer) and self._is_qty_token(token_buffer[i+1]):
                rows.append({'바코드': bc, '정상수': int(token_buffer[i+1]),
                             '불량수': 0, '정상로케': '', '불량로케': ''})
                i += 2
            else:
                rows.append({'바코드': bc, '정상수': 1,
                             '불량수': 0, '정상로케': '', '불량로케': ''})
                i += 1

        if not rows: return None

        # 상세 데이터 보관 (입고파일 생성 시 사용) - 스캔칸일 때만 의미 있음
        # 같은 바코드 여러 줄이면 그대로 보존 (입고파일에 각 행으로 들어가야 하니까)
        if text_widget is getattr(self, 'txt_in_scan', None):
            self.last_scan_detail = rows

        # 대조용 DataFrame: 같은 바코드는 (정상+불량) 합쳐서 합계 수량으로
        df = pd.DataFrame(rows)
        df['수량'] = df['정상수'] + df['불량수']
        return df.groupby('바코드', as_index=False)['수량'].sum()

    def get_unique_filename(self, base_name, extension):
        fn = f"{base_name}.{extension}"; cnt = 1
        # [수정] 중복 파일은 '촬영출고1.csv' 처럼 띄어쓰기 없이 숫자 붙이기
        while os.path.exists(os.path.join(self.save_dir, fn)):
            fn = f"{base_name}{cnt}.{extension}"; cnt += 1
        return fn

    # --- [저장 경로 관리] ---
    def load_save_dir(self):
        """설정 파일에서 저장 경로 불러오기. 없으면 기본값(데스크탑/로지판_작업결과) 사용."""
        default_dir = os.path.join(self.desktop, "로지판_작업결과")
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, "r", encoding="utf-8") as f:
                    cfg = json.load(f)
                saved = cfg.get("save_dir", "").strip()
                if saved and os.path.isdir(saved):
                    return saved
        except Exception as e:
            print(f"⚠️ 설정 파일 로드 실패: {e}")
        return default_dir

    def save_save_dir(self, path):
        """저장 경로를 설정 파일에 기록 (다음 실행 때도 유지)."""
        try:
            cfg = {}
            if os.path.exists(self.config_path):
                try:
                    with open(self.config_path, "r", encoding="utf-8") as f:
                        cfg = json.load(f)
                except Exception:
                    cfg = {}
            cfg["save_dir"] = path
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"⚠️ 설정 파일 저장 실패: {e}")

    def load_user_name(self):
        """설정 파일에서 사용자 이름 불러오기. 없으면 기본값 '장정호'."""
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, "r", encoding="utf-8") as f:
                    cfg = json.load(f)
                name = cfg.get("user_name", "").strip()
                if name:
                    return name
        except Exception as e:
            print(f"⚠️ 사용자 이름 로드 실패: {e}")
        return "장정호"

    def save_user_name(self, name):
        """사용자 이름을 설정 파일에 기록 (업데이트 후에도 유지)."""
        try:
            cfg = {}
            if os.path.exists(self.config_path):
                try:
                    with open(self.config_path, "r", encoding="utf-8") as f:
                        cfg = json.load(f)
                except Exception:
                    cfg = {}
            cfg["user_name"] = name
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"⚠️ 사용자 이름 저장 실패: {e}")

    def change_save_dir(self):
        """경로 변경 버튼 - 폴더 선택 다이얼로그 띄우고 변경 후 영구 저장."""
        new_dir = filedialog.askdirectory(
            title="저장할 폴더 선택",
            initialdir=self.save_dir if os.path.isdir(self.save_dir) else self.desktop
        )
        if not new_dir:
            return
        self.save_dir = new_dir
        if not os.path.exists(self.save_dir):
            try:
                os.makedirs(self.save_dir)
            except Exception as e:
                messagebox.showerror("오류", f"폴더 생성 실패: {e}")
                return
        self.save_save_dir(self.save_dir)
        self.status.config(text=f" 📂 저장 위치: {self.save_dir}")
        messagebox.showinfo("경로 변경 완료", f"저장 위치가 변경되었습니다.\n\n{self.save_dir}\n\n(앞으로 모든 파일은 여기에 저장됩니다)")

    # ========== [추가] FCM 푸시 알림 발송 ==========
    def _upload_image_to_imgbb(self, image_path, parent_win=None):
        """이미지 파일을 imgBB에 업로드하고 URL 반환. 실패시 None."""
        try:
            import requests
            IMGBB_KEY = "3db15a11410b0c569fb9c8706f7b8d12"
            with open(image_path, 'rb') as f:
                files = {'image': f}
                response = requests.post(
                    f"https://api.imgbb.com/1/upload?key={IMGBB_KEY}",
                    files=files, timeout=30
                )
            data = response.json()
            if data.get('success'):
                return data['data']['url']
            else:
                from tkinter import messagebox
                messagebox.showerror("업로드 실패", f"이미지 업로드 실패: {data}", parent=parent_win)
                return None
        except Exception as e:
            from tkinter import messagebox
            messagebox.showerror("업로드 오류", f"이미지 업로드 오류:\n{e}", parent=parent_win)
            return None

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

    def position_popup(self, win, width, height):
        """[추가] 팝업창을 메인 창(self.root) 근처에 띄움.

        모니터가 여러 개일 때 다른 모니터로 팝업이 새어나가지 않게 하기 위해
        메인 창의 화면 좌표를 기준으로 위치를 계산한다.

        topmost 속성이 있는 창은 윈도우 매니저가 위치를 무시할 수 있어서
        withdraw → geometry → deiconify 트릭으로 강제로 위치를 적용한다.
        """
        try:
            # 1. 일단 숨김 (윈도우 매니저가 임의 위치에 그리는 걸 방지)
            win.withdraw()

            # 2. 메인 창의 현재 위치/크기 (절대 화면 좌표)
            self.root.update_idletasks()
            rx = self.root.winfo_rootx()
            ry = self.root.winfo_rooty()
            rw = self.root.winfo_width()
            rh = self.root.winfo_height()

            # 3. 메인 창의 정중앙에 팝업 가운데를 맞춤
            x = rx + (rw - width) // 2
            y = ry + (rh - height) // 2

            # 너무 위로 가서 타이틀바가 짤리지 않도록 보정
            if y < 30: y = 30

            # 4. 위치/크기 적용 후 보이게
            win.geometry(f"{width}x{height}+{x}+{y}")
            win.update_idletasks()
            win.deiconify()
            # 5. 부모 창 위로 (모달같은 효과)
            try:
                win.transient(self.root)
            except Exception:
                pass
            win.lift()
            win.focus_force()
        except Exception as e:
            print(f"⚠️ 팝업 위치 설정 실패: {e}")
            try:
                win.deiconify()
                win.geometry(f"{width}x{height}")
            except Exception:
                pass

    def clean_code_strictly(self, series):
        return series.astype(str).str.strip().str.replace(r'\.0$', '', regex=True).str.upper()

    def f_c_helper(self, df, k):
        return next((c for c in df.columns if k in str(c)), None)

    # --- [탭 1: 입고] ---
    def setup_inbound(self):
        container = tk.Frame(self.t_in, bg="white", padx=10, pady=5); container.pack(fill="both", expand=True)
        top = tk.Frame(container, bg="white"); top.pack(fill="x", pady=(5, 5))
        tk.Label(top, text="🔖 브랜드명:", font=("맑은 고딕", 10, "bold"), bg="white").pack(side="left")
        self.ent_brand_in = tk.Entry(top, font=("맑은 고딕", 11), width=45, bd=1, relief="solid")
        self.ent_brand_in.pack(side="left", padx=5, ipady=3)

        # [수정] 버튼 영역을 먼저 BOTTOM에 배치 (창 크기와 무관하게 항상 보이도록)
        btn_f = tk.Frame(container, bg="white")
        btn_f.pack(side="bottom", fill="x", pady=2)
        # [수정] 버튼 순서: 대조 → 입고파일 생성 → CSV 저장&리셋
        tk.Button(btn_f, text="🔍 대조 분석 실행", bg="#FF9800", fg="white", font=("맑은 고딕", 11, "bold"), command=self.run_compare_in).pack(side="left", expand=True, fill="x", padx=1)
        tk.Button(btn_f, text="📋 입고파일 생성", bg="#1565C0", fg="white", font=("맑은 고딕", 11, "bold"), command=self.create_inbound_file).pack(side="left", expand=True, fill="x", padx=1)
        tk.Button(btn_f, text="💾 CSV 저장 & 리셋", bg="#4CAF50", fg="white", font=("맑은 고딕", 11, "bold"), command=self.save_csv_in).pack(side="left", expand=True, fill="x", padx=1)

        # [수정] 리포트 영역도 BOTTOM에 (버튼 위)
        self.txt_in_report = tk.Text(container, height=6, font=("Consolas", 10), bg="#FAFAFA", bd=1, relief="solid")
        self.txt_in_report.pack(side="bottom", fill="x", pady=(2, 5))
        self.txt_in_report.tag_config("match", foreground="#2E7D32", font=("Consolas", 10, "bold"))
        self.txt_in_report.tag_config("error", foreground="#C62828", font=("Consolas", 10, "bold"))
        self.txt_in_report.tag_config("title", foreground="#1565C0", font=("Consolas", 10, "bold"))
        self.txt_in_report.tag_config("info", foreground="#555", font=("Consolas", 10))
        self.txt_in_report.tag_config("warn", foreground="#E65100", font=("Consolas", 10, "bold"))

        report_title_f = tk.Frame(container, bg="white")
        report_title_f.pack(side="bottom", fill="x", pady=(5, 0))
        tk.Label(report_title_f, text="[ 📝 대조 분석 상세 리포트 ]", font=("맑은 고딕", 9, "bold"), bg="white").pack(side="left")
        tk.Label(report_title_f, text="※ 엔터 2번 시 스캔수량 이동", fg="#888", font=("맑은 고딕", 8)).pack(side="right")

        # [수정] mid (텍스트박스 영역)는 마지막에 pack → 남은 공간 다 차지
        mid = tk.Frame(container, bg="white"); mid.pack(fill="both", expand=True)
        mid.columnconfigure(0, weight=1); mid.columnconfigure(1, weight=1)
        l_f = tk.Frame(mid, bg="white"); l_f.grid(row=0, column=0, sticky="nsew", padx=(0, 2))
        self.lbl_in_m = tk.Label(l_f, text="📦 브랜드 수량 (0개)", bg="#E3F2FD", font=("맑은 고딕", 8, "bold")); self.lbl_in_m.pack(fill="x")
        # [수정] 폰트 9 → 11로 확대
        self.txt_in_master = tk.Text(l_f, font=("Consolas", 11), bd=1, relief="solid"); self.txt_in_master.pack(fill="both", expand=True)
        self.txt_in_master.bind("<KeyRelease>", lambda e: (self.count_total_qty(self.txt_in_master, self.lbl_in_m, "📦 브랜드 수량"), self.txt_in_master.tag_remove("mismatch", "1.0", tk.END)))
        self.txt_in_master.bind("<Return>", self.check_double_enter)
        # [추가] 불일치 라인 하이라이트용 태그
        self.txt_in_master.tag_config("mismatch", background="#FFCDD2", foreground="#B71C1C")
        
        r_f = tk.Frame(mid, bg="white"); r_f.grid(row=0, column=1, sticky="nsew", padx=(2, 0))
        self.lbl_in_s = tk.Label(r_f, text="📡 스캔 수량 (0개)", bg="#F1F8E9", font=("맑은 고딕", 8, "bold")); self.lbl_in_s.pack(fill="x")
        # [수정] 폰트 9 → 11로 확대
        self.txt_in_scan = tk.Text(r_f, font=("Consolas", 11), bd=1, relief="solid"); self.txt_in_scan.pack(fill="both", expand=True)
        self.txt_in_scan.bind("<KeyRelease>", lambda e: (self.count_total_qty(self.txt_in_scan, self.lbl_in_s, "📡 스캔 수량"), self.txt_in_scan.tag_remove("mismatch", "1.0", tk.END)))
        # [추가] 불일치 라인 하이라이트용 태그
        self.txt_in_scan.tag_config("mismatch", background="#FFCDD2", foreground="#B71C1C")

    # --- [탭 2: 출고] ---
    def setup_outbound(self):
        container = tk.Frame(self.t_out, bg="white", padx=10, pady=10)
        container.pack(fill="both", expand=True)

        # 🏬 매장 선택 (브랜드반납까지 한 줄 배치)
        sf = tk.LabelFrame(container, text=" 🏬 매장 선택 ", font=("맑은 고딕", 9, "bold"), bg="white", padx=10, pady=12)
        sf.pack(fill="x", pady=(0, 5))
        
        stores = ["성수", "압구정", "갤러리아", "인하우스", "마케팅", "브랜드반납"]
        for idx, s in enumerate(stores):
            # 한 줄에 다 들어가도록 row=0으로 고정
            btn = tk.Button(sf, text=s, width=11, command=lambda x=s: self.select_opt('s', x), bg="#f0f0f0", font=("맑은 고딕", 10))
            btn.grid(row=0, column=idx, padx=4, pady=5) # padx를 살짝 늘려 간격 조정
            self.s_btns[s] = btn

        # 📝 유형 선택 (퀵출고까지 한 줄 배치)
        tf = tk.LabelFrame(container, text=" 📝 유형 선택 ", font=("맑은 고딕", 9, "bold"), bg="white", padx=10, pady=12)
        tf.pack(fill="x", pady=(0, 10))
        
        types = ["매장출고", "신규출고", "긴급(픽업)출고", "보충출고", "택배출고", "촬영출고", "퀵출고"]
        for idx, t in enumerate(types):
            # 한 줄에 다 들어가도록 row=0으로 고정
            btn = tk.Button(tf, text=t, width=10, command=lambda x=t: self.select_opt('t', x), bg="#f0f0f0", font=("맑은 고딕", 9))
            btn.grid(row=0, column=idx, padx=2, pady=5)
            self.t_btns[t] = btn

        # 기본값 설정
        self.select_opt('s', '성수')
        self.select_opt('t', '매장출고')

        # 하단 입력 및 저장 영역 (기존과 동일)
        self.lbl_out_qty = tk.Label(container, text="📡 출고 바코드 & 수량 붙여넣기 (0개)", font=("맑은 고딕", 9, "bold"), bg="white")
        self.lbl_out_qty.pack(anchor="w")

        self.txt_out = tk.Text(container, font=("Consolas", 10), bd=1, relief="solid", height=12)
        self.txt_out.pack(fill="both", expand=True, pady=5)
        self.txt_out.bind("<KeyRelease>", lambda e: self.count_total_qty(self.txt_out, self.lbl_out_qty, "📡 출고 바코드 & 수량 붙여넣기"))

        tk.Button(container, text="🚚 출고 CSV 저장 및 데이터 리셋", bg="#2196F3", fg="white", 
                  font=("맑은 고딕", 13, "bold"), height=2, command=self.run_out).pack(fill="x", pady=(5, 0))
        
    def select_opt(self, mode, val):
        if mode == 's':
            self.selected_store.set(val)
            for k, btn in self.s_btns.items():
                btn.config(bg="#2196F3" if k == val else "#f0f0f0", 
                           fg="white" if k == val else "black")
        else:
            self.selected_type.set(val)
            for k, btn in self.t_btns.items():
                btn.config(bg="#673AB7" if k == val else "#f0f0f0", 
                           fg="white" if k == val else "black")
    # --- [탭 3: 맘스] ---
    def setup_moms_v86(self):
        container = tk.Frame(self.t_mom, bg="#f5f5f5"); container.pack(fill="both", expand=True)
        tk.Label(container, text="📦 맘스 입/출고 등록", font=("맑은 고딕", 18, "bold"), bg="#f5f5f5", pady=15).pack()
        m_in_f = tk.LabelFrame(container, text=" 1. 맘스 마스터 및 입고등록 ", font=("맑은 고딕", 11, "bold"), bg="white", padx=20, pady=20); m_in_f.pack(fill="both", expand=True, padx=15, pady=5)
        r1 = tk.Frame(m_in_f, bg="white"); r1.pack(fill="x", pady=5)
        tk.Button(r1, text="조회 리스트 선택", command=self.sel_mom_s, width=20).pack(side="left")
        self.lbl_mom_s = tk.Label(r1, text="(이동재고) 선택되지 않음", fg="#777", bg="white", font=("맑은 고딕", 10)); self.lbl_mom_s.pack(side="left", padx=10)
        r2 = tk.Frame(m_in_f, bg="white"); r2.pack(fill="x", pady=5)
        tk.Button(r2, text="마스터 재고 선택", command=self.sel_mom_m, width=20).pack(side="left")
        self.lbl_mom_m = tk.Label(r2, text="(맘스 마스터재고) 선택되지 않음", fg="#777", bg="white", font=("맑은 고딕", 10)); self.lbl_mom_m.pack(side="left", padx=10)
        
        btn_r = tk.Frame(m_in_f, bg="white"); btn_r.pack(fill="x", pady=15)
        # [수정] 3개 버튼 크기/폰트 통일 (font, height 동일하게)
        btn_font = ("맑은 고딕", 10, "bold")
        tk.Button(btn_r, text="중복제거 마스터 리스트 생성", bg="#4CAF50", fg="white", font=btn_font, height=2, command=self.run_mom_master_logic).pack(side="left", expand=True, fill="x", padx=5)
        tk.Button(btn_r, text="입고 리스트 생성", bg="#2E7D32", fg="white", font=btn_font, height=2, command=self.run_mom_inbound_logic).pack(side="left", expand=True, fill="x", padx=5)
        # [추가] 마스터재고에 있는 바코드는 제외하고 신규 바코드만 입고 리스트로 생성
        tk.Button(btn_r, text="신규만 입고 리스트 생성", bg="#1B5E20", fg="white", font=btn_font, height=2, command=self.run_mom_inbound_new_only).pack(side="left", expand=True, fill="x", padx=5)
        
        m_out_f = tk.LabelFrame(container, text=" 2. 맘스 출고등록 ", font=("맑은 고딕", 11, "bold"), bg="white", padx=20, pady=20); m_out_f.pack(fill="both", expand=True, padx=15, pady=5)
        inf = tk.Frame(m_out_f, bg="white"); inf.pack(fill="x", pady=5)
        tk.Label(inf, text="👤 주문자/수령자명:", bg="white", font=("bold", 10)).pack(side="left")
        self.ent_mom_user = tk.Entry(inf, width=20, bd=1, relief="solid", font=("맑은 고딕", 11)); self.ent_mom_user.pack(side="left", padx=10)
        tk.Button(m_out_f, text="📁 출고 리스트 파일 선택", command=self.sel_mom_out, width=25).pack(pady=10)
        self.lbl_mom_out = tk.Label(m_out_f, text="파일 미선택", fg="#777", bg="white"); self.lbl_mom_out.pack()
        tk.Button(m_out_f, text="📋 출고 리스트 생성", bg="#9C27B0", fg="white", font=("맑은 고딕", 13, "bold"), height=2, command=self.run_mom_out_logic).pack(side="bottom", fill="x", pady=10)

    # --- [탭 4: 마감재고 (수량 너비 확보)] ---
    def setup_closing_stock(self):
        container = tk.Frame(self.t_end, bg="white", padx=30, pady=30); container.pack(fill="both", expand=True)
        tk.Label(container, text="📊 실시간 마감재고 분석 (EMP)", font=("맑은 고딕", 22, "bold"), bg="white").pack(pady=20)
        tk.Button(container, text="📁 EMP 재고 파일 선택 및 분석 실행", bg="#607D8B", fg="white", font=("맑은 고딕", 14, "bold"), height=2, command=self.run_closing_stock_logic).pack(fill="x", pady=10)
        self.tree_end = ttk.Treeview(container, columns=("zone", "qty"), show="headings", height=8)
        self.tree_end.heading("zone", text="최종 구역"); self.tree_end.column("zone", anchor="center", width=200)
        self.tree_end.heading("qty", text="가용재고 합계"); self.tree_end.column("qty", anchor="e", width=180)
        self.tree_end.pack(fill="both", expand=True)
        self.lbl_end_total = tk.Label(container, text="📊 총 구역 합계: 0개", font=("맑은 고딕", 18, "bold"), bg="#f5f5f5", pady=15); self.lbl_end_total.pack(fill="x", pady=20)

    # --- [탭 5: 재고파악] ---
    def setup_inventory_check_v95(self):
        container = tk.Frame(self.t_chk, bg="#f5f5f5"); container.pack(fill="both", expand=True)
        tk.Label(container, text="🔎 출고가능 재고파악 시스템", font=("맑은 고딕", 20, "bold"), bg="#f5f5f5", pady=15).pack()
        chk_f = tk.LabelFrame(container, text=" 재고 분석 설정 ", font=("맑은 고딕", 12, "bold"), bg="white", padx=30, pady=40)
        chk_f.pack(expand=True, fill="both", padx=30, pady=10)
        r1 = tk.Frame(chk_f, bg="white"); r1.pack(fill="x", pady=15)
        tk.Button(r1, text="1. 조회 리스트 선택", command=self.sel_chk_target, width=22, bg="#E8EAF6", font=("맑은 고딕", 10)).pack(side="left")
        self.lbl_chk_target = tk.Label(r1, text="(출고리스트) 미선택", fg="#777", bg="white", font=("맑은 고딕", 10)); self.lbl_chk_target.pack(side="left", padx=20)
        r2 = tk.Frame(chk_f, bg="white"); r2.pack(fill="x", pady=15)
        tk.Button(r2, text="2. 재고 마스터 선택", command=self.sel_chk_master, width=22, bg="#E8EAF6", font=("맑은 고딕", 10)).pack(side="left")
        self.lbl_chk_master = tk.Label(r2, text="(EMP 재고파일) 미선택", fg="#777", bg="white", font=("맑은 고딕", 10)); self.lbl_chk_master.pack(side="left", padx=20)
        tk.Button(chk_f, text="📋 출고가능 리스트 생성", bg="#673AB7", fg="white", font=("맑은 고딕", 16, "bold"), height=2, command=self.run_inventory_check_logic).pack(side="bottom", fill="x", pady=(30, 0))
    
    # --- [탭 6: 현장소통] ---
    def setup_field_comm(self, container):
        # 1. 배경색 설정
        container.configure(bg="#F0F2F5") 

        # [수정] 제목 + 새로고침 버튼을 한 프레임에 배치
        title_frame = tk.Frame(container, bg="#F0F2F5")
        title_frame.pack(side="top", fill="x", padx=40, pady=(15, 5))
        # 빈 라벨로 좌측 균형 (제목이 가운데로 오도록)
        tk.Frame(title_frame, bg="#F0F2F5", width=120).pack(side="left")
        tk.Label(title_frame, text="📱 실시간 작업 보고", font=("맑은 고딕", 22, "bold"),
                 bg="#F0F2F5", fg="#1C1E21").pack(side="left", expand=True)
        # 제목 오른쪽에 새로고침 버튼
        tk.Button(title_frame, text="🔄 새로고침", command=self.update_table_view,
                  font=("맑은 고딕", 10, "bold"), bg="#34a853", fg="white",
                  relief="flat", padx=14, pady=6, cursor="hand2").pack(side="right")

        # [위치 교정 2] 필터 & 검색 바
        filter_frame = tk.Frame(container, bg="white", padx=15, pady=8, bd=1, relief="flat")
        filter_frame.pack(side="top", fill="x", padx=40, pady=(0, 10)) 
        
        # --- 내부 순서 ---
        tk.Label(filter_frame, text="🎯 상태 보기: ", font=("맑은 고딕", 10, "bold"), bg="white", fg="#65676B").pack(side="left")
        self.filter_var.set("⏳ 처리중") 
        filter_combo = ttk.Combobox(filter_frame, textvariable=self.filter_var, 
                                    values=["전체", "⏳ 처리중", "✅ 완료"], state="readonly", width=12)
        filter_combo.pack(side="left", padx=5)
        filter_combo.bind("<<ComboboxSelected>>", lambda e: self.update_table_view())

        tk.Label(filter_frame, text="🔍 검색: ", font=("맑은 고딕", 10, "bold"), bg="white", fg="#65676B").pack(side="left", padx=(30, 5))
        search_entry = tk.Entry(filter_frame, textvariable=self.search_var, font=("맑은 고딕", 10),
                                 bd=1, relief="solid", highlightthickness=0, width=25)
        search_entry.pack(side="left", padx=5)
        search_entry.bind("<Return>", lambda e: self.update_table_view())

        # [수정] 내 이름 설정 버튼은 필터바 오른쪽에 배치 (새로고침은 제목 옆으로 이동)
        setting_btn = tk.Button(filter_frame, text="⚙️ 내 이름 설정", command=self.change_user_name,
                                font=("맑은 고딕", 9, "bold"), bg="#f0f0f0", fg="#333", 
                                relief="flat", padx=10, cursor="hand2")
        setting_btn.pack(side="right", padx=5)

        # [제거] 하단 버튼부 (공지사항 전송은 공지/소통 탭과 중복이라 제거,
        # 새로고침은 위로 이동했으므로 하단 영역 자체를 제거함)

        # [위치 교정 4] 리스트 상자 (마지막에 배치하여 남은 중간 공간을 다 차지함)
        frame_list = tk.Frame(container, bg="white", bd=1, highlightthickness=1, highlightbackground="#E4E6EB")
        frame_list.pack(side="top", expand=True, fill="both", padx=40, pady=10)

        # --- 스타일 설정 ---
        style = ttk.Style()
        style.theme_use("clam") 
        style.configure("Treeview", rowheight=40, font=("맑은 고딕", 10, "bold"), 
                        background="white", fieldbackground="white", foreground="#333333", borderwidth=0)
        style.configure("Treeview.Heading", font=("맑은 고딕", 11, "bold"), background="#F8F9FA", foreground="#555")
        style.map("Treeview", background=[('selected', '#E7F3FF')], foreground=[('selected', '#1877F2')])

        # 표 생성
        self.chat_tree = ttk.Treeview(frame_list, columns=("상태", "진행", "날짜", "작업자", "제목"), show="headings", height=18)
        self.chat_tree.tag_configure("oddrow", background="white")
        self.chat_tree.tag_configure("evenrow", background="#F9FAFB") 
        
        scrollbar = ttk.Scrollbar(frame_list, orient="vertical", command=self.chat_tree.yview)
        self.chat_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.chat_tree.pack(side="left", fill="both", expand=True)

        # 컬럼 설정
        cols = {"상태": 90, "진행": 90, "날짜": 90, "작업자": 70, "제목": 180}
        for col, width in cols.items():
            self.chat_tree.heading(col, text=col)
            self.chat_tree.column(col, width=width, anchor="center")

        # 우클릭 메뉴 및 이벤트
        self.context_menu = tk.Menu(self.root, tearoff=0, font=("맑은 고딕", 10))
        self.context_menu.add_command(label="🗑️ 이 보고서 완전 삭제", command=self.delete_report)
        self.chat_tree.bind("<Button-3>", self.show_context_menu)
        self.chat_tree.bind("<Double-1>", self.on_message_double_click)

        # 마지막 데이터 로드
        self.update_table_view()

    def setup_board_system(self, parent): # 여기서 parent라고 받았으면
        # 1. 상단 버튼 영역
        btn_frame = tk.Frame(parent) # 여기도 parent!
        btn_frame.pack(fill='x', padx=10, pady=10)

        # 공지 작성 버튼
        tk.Button(btn_frame, text="📢 공지 작성하기", bg="#1a73e8", fg="white", 
                font=("맑은 고딕", 10, "bold"), command=self.send_global_notice, width=20).pack(side='left', padx=5)

        # 새로고침 버튼 (수동 갱신 - 실시간 리스너 있어도 fallback용으로 둠)
        tk.Button(btn_frame, text="🔄 리스트 새로고침", bg="#34a853", fg="white", 
                font=("맑은 고딕", 10, "bold"), command=self.update_board_view, width=20).pack(side='left', padx=5)

        # [추가] 내 이름 설정 버튼 (작업보고 탭과 동일 변수 self.current_user 공유)
        tk.Button(btn_frame, text="⚙️ 내 이름 설정", command=self.change_user_name,
                  font=("맑은 고딕", 9, "bold"), bg="#f0f0f0", fg="#333", 
                  relief="flat", padx=10, cursor="hand2").pack(side="left", padx=10)

        # [추가] 안내 라벨 (우클릭으로 완료 처리 가능하다는 힌트)
        tk.Label(btn_frame, text="💡 우클릭: 완료/삭제", 
                 fg="#888", font=("맑은 고딕", 9)).pack(side="left", padx=10)

        # 2. 작업자 소통 리스트 구역
        list_label_frame = tk.Frame(parent, bg="#F0F2F5")
        list_label_frame.pack(fill="x", padx=35)
        tk.Label(list_label_frame, text="💬 작업자 소통 현황", font=("맑은 고딕", 12, "bold"), 
                 bg="#F0F2F5", fg="#1C1E21").pack(side="left")

        frame_list = tk.Frame(parent, bg="white", bd=1, relief="solid")
        frame_list.pack(expand=True, fill="both", padx=30, pady=(5, 10))

        # [수정] 컬럼 순서: 날짜 / 작업자 / 구분 / 상태 / 내용
        cols = ("날짜", "작업자", "구분", "상태", "내용")
        self.board_tree = ttk.Treeview(frame_list, columns=cols, show="headings", height=12)
        
        widths = {"날짜": 80, "작업자": 100, "구분": 100, "상태": 90, "내용": 430}
        for col in cols:
            self.board_tree.heading(col, text=col)
            self.board_tree.column(col, width=widths[col], anchor="center")
        
        self.board_tree.pack(side="left", expand=True, fill="both")
        
        # 스크롤바 & 더블클릭 이벤트
        sb = ttk.Scrollbar(frame_list, orient="vertical", command=self.board_tree.yview)
        self.board_tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.board_tree.bind("<Double-1>", self.on_board_double_click)

        # [추가] 우클릭 메뉴 - 완료 처리 / 삭제
        self.board_context_menu = tk.Menu(self.root, tearoff=0, font=("맑은 고딕", 10))
        self.board_context_menu.add_command(label="✅ 완료 처리 (리스트에서 숨김)", command=self.complete_board_post)
        self.board_context_menu.add_separator()
        self.board_context_menu.add_command(label="🗑️ 완전 삭제", command=self.delete_board_post)
        self.board_tree.bind("<Button-3>", self.show_board_context_menu)

        # [추가] 실시간 리스너 시작 (자동으로 변경 감지해서 갱신)
        self.start_board_listener()

        # [제거] 하단 게시글 새로고침 버튼 - 위에 이미 있고 실시간 갱신도 됨

    def send_global_notice(self):
            notice_win = tk.Toplevel(self.root)
            notice_win.title("📢 메시지 전송")
            notice_win.configure(bg="#f8f9fa")
            self.position_popup(notice_win, 500, 650)

            # 1. 대상 선택
            target_var = tk.StringVar(value="all")
            
            target_frame = tk.LabelFrame(notice_win, text="수신 대상", font=("맑은 고딕", 10, "bold"), bg="#f8f9fa", pady=10)
            target_frame.pack(fill="x", padx=20, pady=10)

            tk.Radiobutton(target_frame, text="전체 공지사항", variable=target_var, value="all", bg="#f8f9fa").pack(side="left", padx=20)
            tk.Radiobutton(target_frame, text="특정 작업자 지정", variable=target_var, value="individual", bg="#f8f9fa").pack(side="left", padx=20)

            # [추가] 현재 작성자(=내 이름) 표시
            sender_name = getattr(self, 'current_user', '관리자')
            tk.Label(notice_win, text=f"✍️ 작성자: {sender_name}  (변경하려면 '⚙️ 내 이름 설정' 사용)",
                     font=("맑은 고딕", 9, "bold"), bg="#f8f9fa", fg="#1a73e8").pack(anchor="w", padx=25, pady=(5, 0))

            # 2. 작업자명 입력 (타이핑 방식)
            tk.Label(notice_win, text="👤 받는 사람 이름 (개인 전송 시 필수)", font=("맑은 고딕", 9), bg="#f8f9fa", fg="#666").pack(anchor="w", padx=25)
            name_entry = tk.Entry(notice_win, font=("맑은 고딕", 11), bd=1, relief="solid")
            name_entry.pack(fill="x", padx=20, pady=(0, 15))

            # 3. 내용 입력
            tk.Label(notice_win, text="💬 메시지 내용", font=("맑은 고딕", 10, "bold"), bg="#f8f9fa").pack(anchor="w", padx=25)
            notice_text = tk.Text(notice_win, font=("맑은 고딕", 11), height=15, bd=1, relief="solid")
            notice_text.pack(padx=20, pady=5, fill='x')

            def submit_notice():
                target = target_var.get()
                receiver_name = name_entry.get().strip()
                content = notice_text.get("1.0", tk.END).strip()

                if target == "individual" and not receiver_name:
                    messagebox.showwarning("경고", "이름을 입력해주세요.", parent=notice_win)
                    return
                if not content:
                    messagebox.showwarning("경고", "내용을 입력해주세요.", parent=notice_win)
                    return
                
                if messagebox.askyesno("확인", "메시지를 전송하시겠습니까?", parent=notice_win):
                    try:
                        # [수정] 작성자에 self.current_user 사용
                        my_name = getattr(self, 'current_user', '관리자')
                        # 데이터 구성
                        post_data = {
                            'user': receiver_name if target == "individual" else my_name, 
                            'real_sender': my_name, # 실제 보낸 사람 (내 이름)
                            'category': "공지" if target == "all" else "개인지시",
                            'receiver': receiver_name if target == "individual" else "all",
                            'text': content,
                            'status': "📢 공지" if target == "all" else "🆕 지시", # 리스트에 뜰 상태
                            'timestamp': firestore.SERVER_TIMESTAMP
                        }
                        
                        self.db.collection('board_posts').add(post_data)
                        # [추가] 푸시 알림 발송
                        if target == "all":
                            preview = content[:60] + ('...' if len(content) > 60 else '')
                            self.send_fcm_push("all",
                                                f"📢 공지: {preview}",
                                                f"- {my_name}")
                        else:
                            preview = content[:60] + ('...' if len(content) > 60 else '')
                            self.send_fcm_push(receiver_name,
                                                f"🔒 {my_name}님의 지시",
                                                preview)
                        messagebox.showinfo("완료", "전송되었습니다.", parent=notice_win)
                        notice_win.destroy()
                        self.update_board_view()
                    except Exception as e:
                        messagebox.showerror("오류", f"실패: {e}", parent=notice_win)

            tk.Button(notice_win, text="🚀 메시지 보내기", bg="#1a73e8", fg="white", 
                    font=("맑은 고딕", 12, "bold"), command=submit_notice, height=2).pack(fill='x', padx=20, pady=20)

    # --- [기능 1: 소통 글 더블클릭 - 공지는 열람, 개인/문의는 대화 스레드] ---
    def on_board_double_click(self, event):
        selection = self.board_tree.selection()
        if not selection:
            return

        item_id = selection[0]
        row_values = self.board_tree.item(item_id, "values")  # (날짜, 작성자, 구분, 내용, 상태)

        # Firestore에서 원본 문서 로드
        try:
            doc_ref = self.db.collection('board_posts').document(item_id)
            post_data = doc_ref.get().to_dict() or {}
        except Exception as e:
            messagebox.showerror("오류", f"데이터 조회 실패: {e}")
            return

        category = post_data.get('category', '')

        # 공지(전체공지)는 대화 제외 - 간단한 열람 창
        if '공지' in category:
            self._open_notice_view(item_id, post_data)
            return

        # 개인지시/문의는 대화 스레드 창
        self._open_thread_window(item_id, post_data, row_values)

    def _open_notice_view(self, item_id, post_data):
        """전체공지는 단순 열람 + 숨김 버튼"""
        win = tk.Toplevel(self.root)
        win.title("📢 공지 상세")
        win.configure(bg="white")
        self.position_popup(win, 500, 450)

        header = tk.Frame(win, bg="#FFF9C4", pady=15)
        header.pack(fill="x")
        tk.Label(header, text=f"📢 {post_data.get('user', '관리자')} 공지",
                 font=("맑은 고딕", 13, "bold"), bg="#FFF9C4").pack()
        ts = post_data.get('timestamp')
        time_str = ts.strftime('%Y-%m-%d %H:%M') if ts else ""
        tk.Label(header, text=f"📅 {time_str}",
                 font=("맑은 고딕", 9), bg="#FFF9C4", fg="#666").pack()

        body = tk.Text(win, font=("맑은 고딕", 11), bg="#f8f9fa",
                       padx=15, pady=15, relief="flat", wrap="word")
        body.insert("1.0", post_data.get('text', ''))
        body.config(state="disabled")
        body.pack(fill="both", expand=True, padx=20, pady=15)

        btn_frame = tk.Frame(win, bg="white", pady=10)
        btn_frame.pack(fill="x")

        def do_hide():
            if not messagebox.askyesno("확인", "이 공지를 리스트에서 숨기시겠습니까?", parent=win):
                return
            try:
                self.db.collection('board_posts').document(item_id).update({
                    'status': '✅ 확인완료',
                    'closed_time': firestore.SERVER_TIMESTAMP
                })
                win.destroy()
                self.update_board_view()
            except Exception as e:
                messagebox.showerror("오류", f"처리 실패: {e}", parent=win)

        tk.Button(btn_frame, text="✅ 리스트에서 숨김", bg="#34a853", fg="white",
                  font=("맑은 고딕", 10, "bold"), command=do_hide,
                  relief="flat", cursor="hand2", padx=15, pady=5).pack(side="left", padx=20)
        tk.Button(btn_frame, text="닫기", font=("맑은 고딕", 10),
                  command=win.destroy, relief="flat", cursor="hand2",
                  padx=15, pady=5).pack(side="right", padx=20)

    def _open_thread_window(self, item_id, post_data, row_values):
        """개인지시/문의에 대한 대화형 스레드 창"""
        win = tk.Toplevel(self.root)
        win.title(f"💬 {post_data.get('user', '작업자')}님과의 대화")
        win.configure(bg="#F0F2F5")
        self.position_popup(win, 580, 740)

        # 상단 헤더
        header = tk.Frame(win, bg="#1a73e8", pady=12)
        header.pack(fill="x")
        tk.Label(header, text=f"👤 {post_data.get('user', '작업자')}",
                 font=("맑은 고딕", 13, "bold"), bg="#1a73e8", fg="white").pack()
        ts = post_data.get('timestamp')
        time_str = ts.strftime('%Y-%m-%d %H:%M') if ts else ""
        tk.Label(header, text=f"📅 시작: {time_str} | 📁 {post_data.get('category', '')}",
                 font=("맑은 고딕", 9), bg="#1a73e8", fg="#d4e3fc").pack()

        # 하단 입력창 (먼저 배치해서 공간 확보)
        input_frame = tk.Frame(win, bg="white", bd=1, relief="flat",
                               highlightthickness=1, highlightbackground="#E1E4E8")
        input_frame.pack(fill="x", side="bottom")

        reply_text = tk.Text(input_frame, font=("맑은 고딕", 10), height=3,
                             bg="white", padx=10, pady=8, relief="flat", wrap="word")
        reply_text.pack(fill="x", padx=10, pady=(8, 4))
        reply_text.focus_set()

        btn_bar = tk.Frame(input_frame, bg="white")
        btn_bar.pack(fill="x", padx=10, pady=(0, 10))

        # 대화 스크롤 영역
        chat_container = tk.Frame(win, bg="#F0F2F5")
        chat_container.pack(fill="both", expand=True, padx=10, pady=(10, 0))

        canvas = tk.Canvas(chat_container, bg="#F0F2F5", highlightthickness=0)
        sb = tk.Scrollbar(chat_container, orient="vertical", command=canvas.yview)
        chat_frame = tk.Frame(canvas, bg="#F0F2F5")

        chat_frame.bind("<Configure>",
                        lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=chat_frame, anchor="nw", width=540)
        canvas.configure(yscrollcommand=sb.set)

        sb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        def _add_bubble(text, sender_name, side, bg_color, fg_color="#111", time_str=""):
            row = tk.Frame(chat_frame, bg="#F0F2F5")
            row.pack(fill="x", padx=5, pady=4)
            anchor = "w" if side == "left" else "e"
            lbl_name = tk.Label(row, text=sender_name, font=("맑은 고딕", 8),
                                bg="#F0F2F5", fg="#777")
            # [수정] Label → Text 위젯으로 변경 (드래그 복사 가능)
            # 텍스트 줄 수 계산해서 높이 동적 결정
            wraplen_chars = 32  # wraplength 380px 기준 대략적인 글자 수
            line_count = 1
            for paragraph in text.split('\n'):
                line_count += max(1, (len(paragraph) + wraplen_chars - 1) // wraplen_chars) - 1
                line_count += 1
            line_count = max(1, min(20, line_count - 1))  # 1~20줄 제한

            bubble = tk.Text(row, font=("맑은 고딕", 10),
                             bg=bg_color, fg=fg_color, wrap="word",
                             padx=12, pady=8, bd=0, relief="flat",
                             height=line_count, width=38,
                             cursor="xterm",
                             highlightthickness=0)
            bubble.insert("1.0", text)
            bubble.config(state="disabled")  # 편집은 막되 선택/복사는 허용

            # disabled 상태에서도 마우스 드래그 선택 가능하도록
            def _enable_select(e):
                bubble.config(state="normal")
            def _disable_after(e):
                bubble.config(state="disabled")
            bubble.bind("<Button-1>", lambda e: bubble.config(state="normal"))
            bubble.bind("<FocusOut>", lambda e: bubble.config(state="disabled"))

            # Ctrl+C 복사 단축키도 동작하게
            def _copy(e):
                try:
                    sel = bubble.get("sel.first", "sel.last")
                    self.root.clipboard_clear()
                    self.root.clipboard_append(sel)
                except tk.TclError:
                    pass
                return "break"
            bubble.bind("<Control-c>", _copy)
            bubble.bind("<Control-C>", _copy)

            lbl_time = tk.Label(row, text=time_str, font=("맑은 고딕", 7),
                                bg="#F0F2F5", fg="#999")
            lbl_name.pack(anchor=anchor, padx=8)
            bubble.pack(anchor=anchor, padx=8)
            lbl_time.pack(anchor=anchor, padx=8)

        # 1) 첫 메시지 = 원글 (작업자 쪽 / 왼쪽 흰 말풍선)
        orig_ts = post_data.get('timestamp')
        orig_time = orig_ts.strftime('%m/%d %H:%M') if orig_ts else ""
        # 개인지시는 관리자가 보낸 것이므로 오른쪽, 문의는 작업자가 보낸 것이므로 왼쪽
        category = post_data.get('category', '')
        if '지시' in category:
            # [수정] '관리자' 하드코딩 제거, real_sender 우선 사용
            sender_name = post_data.get('real_sender') or '관리자'
            _add_bubble(post_data.get('text', ''), sender_name,
                        "right", "#D8E7FF", "#0d3b78", orig_time)
        else:
            _add_bubble(post_data.get('text', ''), post_data.get('user', '작업자'),
                        "left", "white", "#111", orig_time)

        # 2) 레거시 reply 필드가 있으면 관리자 답변으로 표시 (과거 단일 답장 시스템)
        legacy_reply = (post_data.get('reply') or '').strip()
        if legacy_reply:
            reply_ts = post_data.get('reply_time')
            reply_time = reply_ts.strftime('%m/%d %H:%M') if reply_ts else ""
            _add_bubble(legacy_reply, "관리자 (예전 답장)",
                        "right", "#FEF7CD", "#5c4b00", reply_time)

        # 3) messages 서브컬렉션 로드
        try:
            msgs = self.db.collection('board_posts').document(item_id) \
                       .collection('messages').order_by('timestamp').get()
            for m in msgs:
                md = m.to_dict()
                role = md.get('role', 'worker')
                sender = md.get('sender', '?')
                text = md.get('text', '')
                mts = md.get('timestamp')
                mtime = mts.strftime('%m/%d %H:%M') if mts else ""
                if role == "admin":
                    _add_bubble(text, sender, "right", "#D8E7FF", "#0d3b78", mtime)
                else:
                    _add_bubble(text, sender, "left", "white", "#111", mtime)
        except Exception as e:
            print(f"메시지 로딩 실패: {e}")

        # 하단 버튼 동작
        def do_send_reply():
            content = reply_text.get("1.0", tk.END).strip()
            if not content:
                messagebox.showwarning("입력 필요", "답장 내용을 입력해주세요.", parent=win)
                return
            try:
                # [수정] 작성자 = 내가 설정한 이름 (관리자 하드코딩 제거)
                my_name = getattr(self, 'current_user', '관리자')
                self.db.collection('board_posts').document(item_id) \
                    .collection('messages').add({
                        'sender': my_name,
                        'role': 'admin',
                        'text': content,
                        'timestamp': firestore.SERVER_TIMESTAMP
                    })
                # 상태가 아직 '신규' 등이면 '💬 대화중'으로 업데이트
                current_status = post_data.get('status', '')
                if '확인완료' not in current_status:
                    self.db.collection('board_posts').document(item_id).update({
                        'status': '💬 대화중',
                        'last_reply_time': firestore.SERVER_TIMESTAMP
                    })
                # [추가] 답장 받은 작업자에게 푸시 알림 발송
                target = post_data.get('user', '')
                if target:
                    preview = content[:80] + ('...' if len(content) > 80 else '')
                    self.send_fcm_push(target,
                                        f"💬 {my_name}님의 답장",
                                        preview)
                win.destroy()
                self.update_board_view()
            except Exception as e:
                messagebox.showerror("전송 실패", f"{e}", parent=win)

        def do_close_thread():
            pending = reply_text.get("1.0", tk.END).strip()
            msg = "이 대화를 '확인완료' 처리하시겠습니까?\n(리스트에서 숨겨집니다)"
            if pending:
                msg = "입력창에 남은 내용을 마지막 답장으로 보낸 뒤\n'확인완료' 처리하시겠습니까?\n(리스트에서 숨겨집니다)"
            if not messagebox.askyesno("확인", msg, parent=win):
                return
            try:
                if pending:
                    # [수정] 작성자 = 내 이름
                    my_name = getattr(self, 'current_user', '관리자')
                    self.db.collection('board_posts').document(item_id) \
                        .collection('messages').add({
                            'sender': my_name,
                            'role': 'admin',
                            'text': pending,
                            'timestamp': firestore.SERVER_TIMESTAMP
                        })
                self.db.collection('board_posts').document(item_id).update({
                    'status': '✅ 확인완료',
                    'closed_time': firestore.SERVER_TIMESTAMP
                })
                win.destroy()
                self.update_board_view()
            except Exception as e:
                messagebox.showerror("오류", f"처리 실패: {e}", parent=win)

        tk.Button(btn_bar, text="✅ 확인완료 (숨김)", bg="#34a853", fg="white",
                  font=("맑은 고딕", 10, "bold"), command=do_close_thread,
                  relief="flat", cursor="hand2", padx=12, pady=5).pack(side="left")
        tk.Button(btn_bar, text="🚀 답장 전송", bg="#1a73e8", fg="white",
                  font=("맑은 고딕", 10, "bold"), command=do_send_reply,
                  relief="flat", cursor="hand2", padx=12, pady=5).pack(side="right")

        # 맨 아래로 스크롤
        def _scroll_to_bottom():
            try:
                canvas.update_idletasks()
                canvas.yview_moveto(1.0)
            except Exception:
                pass
        win.after(80, _scroll_to_bottom)

        # 마우스 휠 스크롤
        def _on_mousewheel(event):
            try:
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            except Exception:
                pass
        win.bind("<MouseWheel>", _on_mousewheel)

        # Ctrl+Enter로 빠른 전송
        def _key_send(event):
            do_send_reply()
            return "break"
        reply_text.bind("<Control-Return>", _key_send)

    # --- [기능 2: 소통 글 리스트 업데이트 및 색상 적용] ---
    def update_board_view(self):
        # 기존 리스트 싹 비우기
        for i in self.board_tree.get_children():
            self.board_tree.delete(i)

        try:
            # 최신 순으로 최대 200개 훑어서 숨김 제외하고 100개까지 표시
            posts = self.db.collection('board_posts').order_by('timestamp', direction=firestore.Query.DESCENDING).limit(200).get()

            shown = 0
            for doc in posts:
                d = doc.to_dict()

                current_status = d.get('status', '') # 예: '🆕 신규', '💬 대화중', '✅ 확인완료'
                current_category = d.get('category', '') # 예: '📢 공지', '개인지시', '문의'

                # [숨김 처리] 확인완료된 스레드는 리스트에서 제외
                if '확인완료' in current_status:
                    continue
                # 레거시: 예전 단일 답장 시스템의 '✅ 확인' 상태도 마무리된 것으로 간주하여 숨김
                if current_status.strip() == '✅ 확인':
                    continue
                # [추가] 완료 처리된 게시글도 리스트에서 숨김 (DB에는 남아있음)
                if '완료' in current_status:
                    continue

                ts = d.get('timestamp')
                # [수정] 시간 제거하고 날짜만 표시
                time_str = ts.strftime('%m/%d') if ts else ""

                # [수정] 작업자 칼럼은 가능하면 real_sender(실제 작성자)를 보여줌.
                # 공지/지시는 user 필드에 받는 사람이 들어가있어서 작성자가 안보였음.
                writer = d.get('real_sender') or d.get('user', '')

                # [수정] 컬럼 순서: 날짜 / 작업자 / 구분 / 상태 / 내용
                item_id = self.board_tree.insert("", "end", iid=doc.id,
                                                values=(time_str,
                                                        writer,
                                                        current_category,
                                                        current_status,
                                                        d.get('text', '').replace('\n', ' ')))
                shown += 1

                # --- [색상 태그 설정] ---
                if "공지" in current_category:
                    self.board_tree.tag_configure("notice", background="#FFF9C4") # 연노랑
                    self.board_tree.item(item_id, tags=("notice",))

                elif "대화중" in current_status:
                    self.board_tree.tag_configure("talking", background="#E3F2FD", foreground="#0D47A1") # 연하늘 배경 + 진파랑 글자
                    self.board_tree.item(item_id, tags=("talking",))

                elif "신규" in current_status or "지시" in current_status:
                    self.board_tree.tag_configure("new", background="#FFEBEE", foreground="#C62828") # 연분홍 배경 + 진빨강 글자
                    self.board_tree.item(item_id, tags=("new",))

                if shown >= 100:
                    break

        except Exception as e:
            print(f"새로고침 오류: {e}")

    # --- [공지/소통 우클릭 메뉴 & 완료/삭제 & 실시간 리스너] ---
    def show_board_context_menu(self, event):
        """공지/소통 리스트 우클릭 시 메뉴 표시"""
        item = self.board_tree.identify_row(event.y)
        if item:
            self.board_tree.selection_set(item)
            self.board_context_menu.post(event.x_root, event.y_root)

    def complete_board_post(self):
        """선택한 게시글을 '완료' 처리 (리스트에서는 숨김, DB에는 유지)"""
        selected = self.board_tree.selection()
        if not selected:
            return
        doc_id = selected[0]
        try:
            d = self.board_tree.item(doc_id)
            # [수정] 컬럼 순서가 (날짜, 작업자, 구분, 상태, 내용)이므로 내용은 인덱스 4
            title_preview = d['values'][4][:30] if d.get('values') and len(d['values']) > 4 else ''
            if not messagebox.askyesno("완료 처리",
                f"이 게시글을 완료 처리할까요?\n\n"
                f"내용: {title_preview}...\n\n"
                f"※ 리스트에서는 사라지지만 데이터는 보관됩니다."):
                return
            self.db.collection('board_posts').document(doc_id).update({
                'status': '✅ 완료',
                'completed_at': firestore.SERVER_TIMESTAMP
            })
            self.update_board_view()
        except Exception as e:
            messagebox.showerror("오류", f"완료 처리 실패: {e}")

    def delete_board_post(self):
        """선택한 게시글을 완전 삭제 (DB에서도 제거)"""
        selected = self.board_tree.selection()
        if not selected:
            return
        doc_id = selected[0]
        try:
            if not messagebox.askyesno("완전 삭제",
                "정말로 이 게시글을 완전히 삭제할까요?\n\n"
                "⚠️ 데이터까지 영구 삭제되며 복구할 수 없습니다.\n"
                "(단순히 리스트에서 숨기려면 '완료 처리'를 사용하세요)"):
                return
            self.db.collection('board_posts').document(doc_id).delete()
            messagebox.showinfo("삭제 완료", "게시글이 삭제되었습니다.")
            self.update_board_view()
        except Exception as e:
            messagebox.showerror("오류", f"삭제 실패: {e}")

    def start_board_listener(self):
        """공지/소통 게시판 실시간 리스너.
        Firestore의 on_snapshot은 변경된 문서에 대해서만 read를 사용하므로
        활동 적은 게시판이면 데이터 사용량은 거의 무시 가능."""
        # 부팅 직후 5초 안에 들어오는 ADDED 이벤트는 알림 안 띄움 (이미 있던 글 로딩이라서)
        self._board_is_booting = True
        self.root.after(5000, lambda: setattr(self, '_board_is_booting', False))

        def on_board_snapshot(col_snapshot, changes, read_time):
            # [추가] 신규 글 들어왔는지 확인
            new_post_detected = False
            for change in changes:
                if change.type.name == 'ADDED' and not getattr(self, '_board_is_booting', True):
                    new_post_detected = True
                    break

            # 백그라운드 스레드에서 호출되므로 메인 스레드로 안전하게 넘김
            try:
                self.root.after(10, lambda: self._safe_board_refresh(new_post_detected))
            except Exception:
                pass

        try:
            # 최근 200개만 구독 (옛날 글 변경에는 반응 안함 = 비용 절약)
            query = self.db.collection('board_posts').order_by(
                'timestamp', direction=firestore.Query.DESCENDING
            ).limit(200)
            query.on_snapshot(on_board_snapshot)
            print("📡 공지/소통 실시간 리스너 가동")
        except Exception as e:
            print(f"❌ 공지/소통 리스너 설정 오류: {e}")

    def _safe_board_refresh(self, new_post_detected=False):
        """메인(GUI) 스레드에서 안전하게 board 갱신"""
        try:
            if hasattr(self, 'root') and self.root.winfo_exists():
                self.update_board_view()
                # [추가] 신규 글이 있으면 공지/소통 탭에 빨간 점
                if new_post_detected:
                    self.set_tab_alert(self.t_board, True)
        except Exception as e:
            print(f"❌ board UI 리프레시 오류: {e}")

        # --- [신규 추가] 삭제 관련 함수들 ---
    def show_context_menu(self, event):
        """마우스 우클릭 시 메뉴 팝업"""
        item = self.chat_tree.identify_row(event.y)
        if item:
            self.chat_tree.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)

    def delete_report(self):
        """선택된 보고서와 하위 댓글들 실제 삭제"""
        selected_item = self.chat_tree.selection()
        if not selected_item: return
        
        doc_id = selected_item[0]
        from tkinter import messagebox
        
        if messagebox.askyesno("⚠️ 삭제 확인", "이 보고서와 모든 대화 내역이 영구 삭제됩니다.\n정말 삭제하시겠습니까?"):
            try:
                # 1. 댓글(서브 컬렉션) 삭제
                comments = self.db.collection('field_reports').document(doc_id).collection('comments').get()
                for c in comments:
                    c.reference.delete()
                
                # 2. 본문 삭제
                self.db.collection('field_reports').document(doc_id).delete()
                
                messagebox.showinfo("삭제 완료", "데이터가 깨끗하게 정리되었습니다.")
                self.update_table_view()
            except Exception as e:
                messagebox.showerror("삭제 실패", f"오류가 발생했습니다: {e}")

    # --- [실행 엔진 로직 - 문법 교정 및 파일명 규칙 반영] ---
    def run_compare_in(self):
        self.txt_in_report.delete("1.0", tk.END); self.is_matched = False
        # [추가] 이전 하이라이트 모두 제거
        self.txt_in_master.tag_remove("mismatch", "1.0", tk.END)
        self.txt_in_scan.tag_remove("mismatch", "1.0", tk.END)

        # 5열 상세 데이터 초기화 → parse_logi_data가 다시 채워줌
        self.last_scan_detail = []

        df_m = self.parse_logi_data(self.txt_in_master); df_s = self.parse_logi_data(self.txt_in_scan)
        if df_m is None or df_s is None:
            self.txt_in_report.insert(tk.END, "⚠️ 브랜드/스캔 양쪽 모두 입력해주세요.\n", "warn")
            return

        merged = pd.merge(df_m, df_s, on='바코드', how='outer', suffixes=('_브랜드', '_스캔')).fillna(0)
        # [수정] 차이 = 스캔 - 브랜드 (스캔 기준): 양수면 스캔이 많음, 음수면 부족
        merged['차이'] = merged['수량_스캔'] - merged['수량_브랜드']
        t_b, t_s = int(df_m['수량'].sum()), int(df_s['수량'].sum())

        # [추가] 5열 상세에서 정상/불량 합계 산출
        normal_sum = sum(r.get('정상수', 0) for r in (self.last_scan_detail or []))
        defect_sum = sum(r.get('불량수', 0) for r in (self.last_scan_detail or []))
        has_5col = bool(self.last_scan_detail) and any(r.get('정상로케') for r in self.last_scan_detail)

        # === [헤더] ===
        self.txt_in_report.insert(tk.END, "═" * 50 + "\n", "info")
        self.txt_in_report.insert(tk.END, "📊 [ 대조 결과 요약 ]\n", "title")
        self.txt_in_report.insert(tk.END, "═" * 50 + "\n", "info")
        self.txt_in_report.insert(tk.END, f"  • 브랜드 수량 : {t_b:>4}개\n", "info")
        self.txt_in_report.insert(tk.END, f"  • 스캔 수량   : {t_s:>4}개\n", "info")
        if has_5col:
            self.txt_in_report.insert(tk.END, f"     └ 정상 : {normal_sum}개 / 불량 : {defect_sum}개\n", "info")

        diff = t_s - t_b
        # [수정] 스캔 기준으로 차이 표현
        if diff == 0:
            self.txt_in_report.insert(tk.END, "  • 수량 총계   : ✅ 일치\n", "match")
        elif diff > 0:
            self.txt_in_report.insert(tk.END, f"  • 수량 총계   : 🚨 스캔이 {diff}개 더 많음\n", "error")
        else:
            self.txt_in_report.insert(tk.END, f"  • 수량 총계   : 🚨 스캔이 {abs(diff)}개 부족\n", "error")

        # === [불일치 상세] ===
        mismatch_rows = [r for _, r in merged.iterrows() if r['차이'] != 0]

        if not mismatch_rows and diff == 0:
            self.is_matched = True
            self.txt_in_report.insert(tk.END, "\n✨ 모든 바코드 수량 일치!\n", "match")
            messagebox.showinfo("대조 완료", "✅ 모든 데이터 일치!")
        else:
            self.txt_in_report.insert(tk.END, "\n" + "─" * 50 + "\n", "info")
            self.txt_in_report.insert(tk.END, f"🚨 불일치 항목 ({len(mismatch_rows)}건)\n", "error")
            self.txt_in_report.insert(tk.END, "─" * 50 + "\n", "info")
            for row in mismatch_rows:
                bc = row['바코드']
                qb = int(row['수량_브랜드'])
                qs = int(row['수량_스캔'])
                d = qs - qb
                # 친절한 설명: 어느 쪽이 얼마나 차이나는지
                if qb == 0:
                    detail = f"브랜드에 없는 바코드"
                elif qs == 0:
                    detail = f"스캔 누락"
                elif d > 0:
                    detail = f"스캔이 {d}개 더 많음"
                else:
                    detail = f"스캔이 {abs(d)}개 부족"
                self.txt_in_report.insert(tk.END,
                    f"  ❌ {bc}\n"
                    f"     브랜드 {qb} ↔ 스캔 {qs}  → {detail}\n",
                    "error")

            # 빨간 배경 표시
            mismatch_codes = [str(row['바코드']) for row in mismatch_rows]
            self._highlight_mismatch_lines(self.txt_in_master, mismatch_codes)
            self._highlight_mismatch_lines(self.txt_in_scan, mismatch_codes)
            messagebox.showwarning("대조 결과", f"🚨 불일치 발견 ({len(mismatch_rows)}건)")

        self.last_merged_df = merged

    def _highlight_mismatch_lines(self, text_widget, mismatch_codes):
        """텍스트 위젯에서 mismatch_codes(정규화된 바코드)에 해당하는 줄을 빨간 배경 표시.

        - 5열 형식 줄이면 줄 전체를 색칠
        - 일반 형식이면 (바코드 + 수량 토큰)을 색칠
        - 바코드 비교는 정규화 기준 (하이픈 무관)
        """
        if not mismatch_codes:
            return
        mismatch_set = {self.normalize_barcode(c) for c in mismatch_codes}

        raw = text_widget.get("1.0", tk.END)
        # 줄 단위로 처리 (5열 형식 줄을 통째로 색칠하기 위해)
        offset = 0
        for line in raw.split('\n'):
            line_len = len(line)
            stripped = line.strip()
            if not stripped:
                offset += line_len + 1  # +1 for \n
                continue

            tokens_str = re.split(r'[\t ]+', stripped)
            tokens_str = [t for t in tokens_str if t]

            # 5열 형식 줄?
            is_5col = (
                len(tokens_str) >= 5
                and self._is_qty_token(tokens_str[1])
                and self._is_qty_token(tokens_str[2])
                and self._looks_like_location(tokens_str[3])
                and self._looks_like_location(tokens_str[4])
            )

            if is_5col:
                bc_norm = self.normalize_barcode(tokens_str[0])
                if bc_norm in mismatch_set:
                    # 줄 전체 색칠 (앞뒤 공백 제외)
                    line_start = offset + line.find(stripped)
                    line_end = line_start + len(stripped)
                    start_idx = f"1.0 + {line_start} chars"
                    end_idx = f"1.0 + {line_end} chars"
                    text_widget.tag_add("mismatch", start_idx, end_idx)
            else:
                # 일반 형식: 토큰 위치 직접 추적
                token_iter = re.finditer(r'\S+', line)
                tokens = list(token_iter)
                i = 0
                while i < len(tokens):
                    tok_match = tokens[i]
                    tok = tok_match.group()
                    has_qty = (i + 1 < len(tokens)) and tokens[i + 1].group().isdigit()

                    if self.normalize_barcode(tok) in mismatch_set:
                        start_offset = offset + tok_match.start()
                        end_offset = offset + (tokens[i + 1].end() if has_qty else tok_match.end())
                        start_idx = f"1.0 + {start_offset} chars"
                        end_idx = f"1.0 + {end_offset} chars"
                        text_widget.tag_add("mismatch", start_idx, end_idx)

                    i += 2 if has_qty else 1

            offset += line_len + 1  # +1 for \n

    def save_csv_in(self):
        if not self.is_matched: messagebox.showwarning("주의", "결과 불일치로 저장 불가"); return
        brand = self.ent_brand_in.get().strip(); today = datetime.now().strftime('%y%m%d')
        df = self.last_merged_df[self.last_merged_df['수량_스캔'] > 0][['바코드', '수량_스캔']].copy(); df.columns=['바코드','수량']; df['메모']=brand
        fn = self.get_unique_filename(f"{today} {brand} 입고", "csv")
        df.to_csv(os.path.join(self.save_dir, fn), index=False, encoding="utf-8-sig"); self.reset_inbound(); messagebox.showinfo("완료", f"저장됨: {fn}")

    def create_inbound_file(self):
        """[추가] 스캔칸의 5열 데이터로 엑셀 입고파일 생성.

        컬럼: 박스번호 | 바코드 | 정상수량 | 불량수량 | 정상로케이션 | 불량로케이션
        - 박스번호는 빈칸 (사용자가 나중에 채움)
        - 5열 형식이 아닌 데이터는 정상수량=수량, 불량수량=0, 로케이션=공백으로 처리
        """
        # parse를 먼저 한번 돌려서 last_scan_detail을 채움
        df_check = self.parse_logi_data(self.txt_in_scan)
        if df_check is None or not getattr(self, 'last_scan_detail', None):
            messagebox.showwarning("주의",
                "스캔칸이 비어있거나 인식할 데이터가 없습니다.\n\n"
                "📝 입력 형식 (탭 또는 공백 구분, 한 줄당 1상품):\n"
                "  바코드  정상수량  불량수량  정상로케이션  불량로케이션\n\n"
                "예시:\n"
                "  NIKU261MSSN107BK280  1  0  DD-02-06-03  00-00-00-00")
            return

        rows = self.last_scan_detail
        # 5열 형식이 아예 하나도 없으면 그냥 단순 형식인 거니까 안내
        any_5col = any(r.get('정상로케') for r in rows)
        if not any_5col:
            if not messagebox.askyesno("확인",
                "스캔칸에 5열 형식 데이터가 없는 것 같습니다.\n"
                "(바코드 정상수 불량수 정상로케 불량로케 형식)\n\n"
                "현재 데이터를 정상수량=합계, 불량수량=0으로 입고파일을 만들까요?"):
                return

        try:
            df_out = pd.DataFrame([{
                '박스번호': '',
                '바코드': r['바코드'],
                '정상수량': r.get('정상수', 0),
                '불량수량': r.get('불량수', 0),
                '정상로케이션': r.get('정상로케', ''),
                '불량로케이션': r.get('불량로케', ''),
            } for r in rows])

            today = datetime.now().strftime('%y%m%d')
            brand = self.ent_brand_in.get().strip()
            base_name = f"{today} {brand} 입고파일" if brand else f"{today} 입고파일"
            fn = self.get_unique_filename(base_name, "xlsx")
            full_path = os.path.join(self.save_dir, fn)
            df_out.to_excel(full_path, index=False)

            # 결과 요약
            total_normal = int(df_out['정상수량'].sum())
            total_defect = int(df_out['불량수량'].sum())
            messagebox.showinfo("입고파일 생성 완료",
                f"📄 파일: {fn}\n\n"
                f"📊 요약\n"
                f"  • 행 수: {len(df_out)}건\n"
                f"  • 정상 합계: {total_normal}개\n"
                f"  • 불량 합계: {total_defect}개\n"
                f"  • 박스번호: 빈칸 (수동 입력 필요)\n\n"
                f"📂 위치: {self.save_dir}")
        except Exception as e:
            messagebox.showerror("오류", f"입고파일 생성 실패: {e}")

    def run_out(self):
        df = self.parse_logi_data(self.txt_out); today = datetime.now().strftime('%y%m%d')
        if df is None: return
        s, t = self.selected_store.get(), self.selected_type.get(); df['메모'] = f"{today} {s} {t}"
        fn = self.get_unique_filename(f"{today} {s} {t}", "csv")
        df.to_csv(os.path.join(self.save_dir, fn), index=False, encoding="utf-8-sig"); messagebox.showinfo("완료", f"저장됨: {fn}")
        self.txt_out.delete("1.0", tk.END); self.lbl_out_qty.config(text="📡 출고 바코드 & 수량 붙여넣기 (0개)")

    def run_mom_master_logic(self):
        if not self.moms_files["send"] or not self.moms_files["master"]: return
        try:
            today = datetime.now().strftime('%y%m%d')
            df_s, bc_c = self.smart_load_moms(self.moms_files["send"], '바코드')
            df_m, mc_c = self.smart_load_moms(self.moms_files["master"], '상품코드')
            df_new = df_s[~df_s[bc_c].isin(df_m[mc_c])].copy()
            def f_c_tmp(df_in, k_in):
                return next((c for c in df_in.columns if k_in in str(c)), None)
            m_reg = pd.DataFrame()
            m_reg['브랜드명'] = df_new.get(f_c_tmp(df_new, '브랜드'), '')
            m_reg['상품군'] = df_new.get(f_c_tmp(df_new, '아이템'), '')
            m_reg['스타일번호'] = df_new.get(f_c_tmp(df_new, '상품코드'), '')
            m_reg['바코드명'] = df_new[bc_c]; m_reg['상품코드'] = df_new[bc_c]
            m_reg['상품명'] = df_new.get(f_c_tmp(df_new, '상품명'), ''); m_reg['상품옵션'] = df_new.get(f_c_tmp(df_new, '사이즈'), ''); m_reg['색상'] = '999'
            # [수정] 맘스 마스터 리스트 파일명
            fn = self.get_unique_filename(f"{today} 맘스 마스터 리스트", "xlsx")
            m_reg.to_excel(os.path.join(self.save_dir, fn), index=False)
            messagebox.showinfo("완료", "저장 완료")
            # [수정] 맘스 마스터 생성 시에는 리셋하지 않음 (파일 선택 유지)
        except Exception as e: messagebox.showerror("오류", str(e))

    def run_mom_inbound_logic(self):
        if not self.moms_files["send"]: return
        try:
            today = datetime.now().strftime('%y%m%d'); df_s, bc_col = self.smart_load_moms(self.moms_files["send"], '바코드')
            qty_col = next((c for c in df_s.columns if '가용재고' in str(c)), next((c for c in df_s.columns if '수량' in str(c)), None))
            i_reg = pd.DataFrame(); i_reg['화주코드'] = ['INOOINTEMPTY'] * len(df_s); i_reg['센터명'] = '신여주2'; i_reg['층정보'] = 'A2'; i_reg['마스터코드'] = ""; i_reg['상품바코드(sku)'] = df_s[bc_col]; i_reg['입고'] = df_s[qty_col] if qty_col else '1'
            # [수정] 맘스 입고 리스트 파일명
            fn = self.get_unique_filename(f"{today} 맘스 입고 리스트", "xlsx")
            i_reg.to_excel(os.path.join(self.save_dir, fn), index=False); messagebox.showinfo("완료", "저장 완료")
            # [추가] 입고 리스트 생성 후에는 최종 리셋
            self.moms_files={"send":"","master":"","out_list":""}; self.lbl_mom_s.config(text="미선택", fg="#777"); self.lbl_mom_m.config(text="미선택", fg="#777")
        except Exception as e: messagebox.showerror("오류", str(e))

    def run_mom_inbound_new_only(self):
        """[추가] 조회 리스트에서 마스터 재고에 이미 있는 바코드는 제외하고
        신규 바코드만 입고 리스트로 생성한다."""
        if not self.moms_files["send"]:
            messagebox.showwarning("주의", "조회 리스트를 먼저 선택해주세요.")
            return
        if not self.moms_files["master"]:
            messagebox.showwarning("주의", "마스터 재고 파일을 먼저 선택해주세요.\n(마스터에 있는 바코드를 제외해야 하므로 필수입니다)")
            return
        try:
            today = datetime.now().strftime('%y%m%d')
            df_s, bc_col = self.smart_load_moms(self.moms_files["send"], '바코드')
            df_m, mc_col = self.smart_load_moms(self.moms_files["master"], '상품코드')

            # 비교를 위해 양쪽 코드 정규화 (공백/대소문자/.0 끝 정리)
            send_codes = self.clean_code_strictly(df_s[bc_col])
            master_codes = set(self.clean_code_strictly(df_m[mc_col]).tolist())

            # 마스터에 없는(=신규) 행만 골라낸다
            mask_new = ~send_codes.isin(master_codes)
            df_new = df_s[mask_new].copy()

            if df_new.empty:
                messagebox.showinfo("결과 없음",
                    "조회 리스트의 모든 바코드가 이미 마스터 재고에 존재합니다.\n신규로 추가할 항목이 없습니다.")
                return

            qty_col = next((c for c in df_new.columns if '가용재고' in str(c)),
                           next((c for c in df_new.columns if '수량' in str(c)), None))

            i_reg = pd.DataFrame()
            i_reg['화주코드'] = ['INOOINTEMPTY'] * len(df_new)
            i_reg['센터명'] = '신여주2'
            i_reg['층정보'] = 'A2'
            i_reg['마스터코드'] = ""
            i_reg['상품바코드(sku)'] = df_new[bc_col].values
            i_reg['입고'] = df_new[qty_col].values if qty_col else '1'

            fn = self.get_unique_filename(f"{today} 맘스 입고 리스트(신규만)", "xlsx")
            i_reg.to_excel(os.path.join(self.save_dir, fn), index=False)

            total_cnt = len(df_s)
            new_cnt = len(df_new)
            excluded_cnt = total_cnt - new_cnt
            messagebox.showinfo("완료",
                f"저장 완료!\n\n📊 결과 요약\n"
                f"• 조회 리스트 전체: {total_cnt}건\n"
                f"• 마스터에 이미 존재 (제외): {excluded_cnt}건\n"
                f"• 신규 입고 대상: {new_cnt}건\n\n"
                f"📄 파일: {fn}")

            # 리셋
            self.moms_files = {"send": "", "master": "", "out_list": ""}
            self.lbl_mom_s.config(text="(이동재고) 선택되지 않음", fg="#777")
            self.lbl_mom_m.config(text="(맘스 마스터재고) 선택되지 않음", fg="#777")
        except Exception as e:
            messagebox.showerror("오류", str(e))

    def run_mom_out_logic(self):
        user = self.ent_mom_user.get().strip(); today = datetime.now().strftime('%y%m%d')
        if not user or not self.moms_files["out_list"]: return
        try:
            df_src, bc_c = self.smart_load_moms(self.moms_files["out_list"], '바코드')
            qty_col = next((c for c in df_src.columns if '가용재고' in str(c)), next((c for c in df_src.columns if '수량' in str(c)), None))
            cols = ["바코드", "상품명", "주문번호", "주문일련번호", "주문자명", "수령자명", "우편번호", "주소", "전화번호", "핸드폰", "옵션명", "주문수량", "사용자메모", "결제금액", "상품코드"]
            df_f = pd.DataFrame(columns=cols); df_f["바코드"] = df_src[bc_c]; df_f["상품명"] = df_src.get(self.f_c_helper(df_src, '상품명'), ""); df_f["옵션명"] = df_src.get(self.f_c_helper(df_src, '사이즈'), ""); df_f["주문수량"] = df_src[qty_col] if qty_col else "1"; df_f["상품코드"] = df_src[bc_c]; df_f["주문자명"] = user; df_f["수령자명"] = user; df_f["우편번호"] = "04783"; df_f["주소"] = "서울 성동구 아차산로 126 (더리브 세종타워) 지하 2층"; df_f["전화번호"] = "070-4157-4825"; df_f["핸드폰"] = "070-4157-4825"
            # [수정] 맘스 출고 리스트 파일명
            fn = self.get_unique_filename(f"{today} 맘스 출고등록 {user}", "xlsx")
            df_f.to_excel(os.path.join(self.save_dir, fn), index=False); messagebox.showinfo("완료", "저장 완료")
            self.ent_mom_user.delete(0, tk.END); self.lbl_mom_out.config(text="파일 미선택", fg="#777"); self.moms_files["out_list"]=""
        except Exception as e: messagebox.showerror("오류", str(e))

    def run_inventory_check_logic(self):
        if not self.chk_files["target"] or not self.chk_files["master"]: return
        try:
            df_target = self.smart_load_chk(self.chk_files["target"], False); df_master = self.smart_load_chk(self.chk_files["master"], True); today = datetime.now().strftime('%y%m%d')
            m_map = {k: self.f_c_helper(df_master, k) for k in ['브랜드', '상품명', '사이즈', '바코드', '창고', '다중로케이션', '가용재고', '상품코드']}
            t_key = next((c for c in df_target.columns if any(x in str(c) for x in ['바코드', '상품코드', '코드'])), None)
            df_target[t_key] = self.clean_code_strictly(df_target[t_key]); final_res = pd.DataFrame()
            for match_type in ['바코드', '상품코드']:
                m_key = m_map[match_type]
                if m_key: df_master[m_key] = self.clean_code_strictly(df_master[m_key]); temp = pd.merge(df_target[[t_key]], df_master, left_on=t_key, right_on=m_key, how='inner'); final_res = pd.concat([final_res, temp], ignore_index=True)
            if final_res.empty: messagebox.showwarning("결과", "일치 데이터 없음"); return
            qty_col = m_map['가용재고']
            if qty_col: final_res[qty_col] = pd.to_numeric(final_res[qty_col], errors='coerce').fillna(0); final_res = final_res[final_res[qty_col] >= 1]
            m_bar = m_map['바코드'] if m_map['바코드'] else m_map['상품코드']; u_keys = [m_bar]; [u_keys.append(m_map[k]) for k in ['창고', '다중로케이션'] if m_map[k]]
            final_res = final_res.drop_duplicates(subset=u_keys); output_keys = ['브랜드', '상품명', '사이즈', '바코드', '창고', '다중로케이션', '가용재고']
            actual_cols = [m_map[k] for k in output_keys if m_map[k] in final_res.columns]; final_df = final_res[actual_cols].copy(); final_df.columns = [k for k in output_keys if m_map[k] in final_res.columns]
            fn = self.get_unique_filename(f"{today} 출고가능 리스트", "xlsx")
            final_df.to_excel(os.path.join(self.save_dir, fn), index=False); messagebox.showinfo("성공", f"{len(final_df)}건 추출 완료!")
            self.chk_files={"target":"","master":""}; self.lbl_chk_target.config(text="미선택", fg="#777"); self.lbl_chk_master.config(text="미선택", fg="#777")
        except Exception as e: messagebox.showerror("오류", str(e))

    # --- [유틸리티 헬퍼] ---
    def sel_mom_s(self): p=filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")]); self.moms_files["send"]=p; self.lbl_mom_s.config(text=os.path.basename(p), fg="blue")
    def sel_mom_m(self): p=filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")]); self.moms_files["master"]=p; self.lbl_mom_m.config(text=os.path.basename(p), fg="blue")
    def sel_mom_out(self): p=filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")]); self.moms_files["out_list"]=p; self.lbl_mom_out.config(text=os.path.basename(p), fg="blue")
    def sel_chk_target(self): p=filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")]); self.chk_files["target"]=p; self.lbl_chk_target.config(text=os.path.basename(p), fg="blue")
    def sel_chk_master(self): p=filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")]); self.chk_files["master"]=p; self.lbl_chk_master.config(text=os.path.basename(p), fg="blue")
    def reset_inbound(self): 
        self.txt_in_master.tag_remove("mismatch", "1.0", tk.END)
        self.txt_in_scan.tag_remove("mismatch", "1.0", tk.END)
        self.txt_in_master.delete("1.0", tk.END); self.txt_in_scan.delete("1.0", tk.END); self.ent_brand_in.delete(0, tk.END); self.is_matched = False; self.lbl_in_m.config(text="📦 브랜드 수량 (0개)"); self.lbl_in_s.config(text="📡 스캔 수량 (0개)")
    def smart_load_chk(self, path, is_master=False):
        skips = [1, 0, 2, 3, 4, 5] if is_master else [0, 1, 2, 3, 4, 5]
        for skip in skips:
            try:
                df = pd.read_excel(path, dtype=str, skiprows=skip); df.columns = [str(c).strip().replace(" ", "").replace("\n", "").replace("\r", "") for c in df.columns]
                if any(k in df.columns for k in ['바코드', '상품코드', '가용재고']): return df
            except: continue
        return pd.read_excel(path, dtype=str)
    def smart_load_moms(self, file_path, target_keyword):
        engine = 'xlrd' if file_path.endswith('.xls') else 'openpyxl'
        try: df_raw = pd.read_excel(file_path, header=None, dtype=str, engine=engine)
        except: df_raw = pd.read_excel(file_path, header=None, dtype=str)
        f_row = -1
        for i, row in df_raw.iterrows():
            if any(target_keyword in str(c) for c in row.values): f_row = i; break
        if f_row == -1: raise ValueError(f"'{target_keyword}' 열 없음")
        df = pd.read_excel(file_path, header=f_row, dtype=str, engine=engine); df.columns = [str(c).strip() for c in df.columns]
        bc_col = next((c for c in df.columns if target_keyword in c), df.columns[0])
        df = df[df[bc_col].notna() & (df[bc_col].astype(str).str.strip() != "")]
        return df, bc_col
    def run_closing_stock_logic(self):
        file_path = filedialog.askopenfilename(title="EMP 파일 선택", filetypes=[("Excel", "*.xlsx *.xls")])
        if not file_path: return
        try:
            raw_df = pd.read_excel(file_path, header=None); header_row = -1
            for i, row in raw_df.iterrows():
                if '창고' in row.values: header_row = i; break
            if header_row == -1: raise ValueError("'창고' 행 없음")
            df = raw_df.iloc[header_row+1:].copy(); df.columns = raw_df.iloc[header_row]; df.columns = [str(c).strip() for c in df.columns]
            df = df[df['창고'].astype(str).str.contains('정상창고', na=False)]; df['가용재고'] = pd.to_numeric(df['가용재고'], errors='coerce').fillna(0); df['구역코드'] = df['다중로케이션'].astype(str).str[:2]
            def classify_tmp(code):
                c = str(code).upper()
                if 'AA' <= c <= 'AB': return 'AA~AB구역'
                if c == 'BB': return 'BB구역'
                if 'CC' <= c <= 'DD': return 'CC~DD구역'
                if c in ['EA', 'EE', 'FF']: return 'EA/EE/FF구역'
                return '기타'
            df['최종구역'] = df['구역코드'].apply(classify_tmp); summary = df[df['최종구역'] != '기타'].groupby('최종구역')['가용재고'].sum().reset_index()
            for item in self.tree_end.get_children(): self.tree_end.delete(item)
            total = 0
            for _, row in summary.iterrows(): self.tree_end.insert("", "end", values=(row['최종구역'], f"{int(row['가용재고']):,}개")); total += row['가용재고']
            self.lbl_end_total.config(text=f"📊 합계: {int(total):,}개"); messagebox.showinfo("완료", "분석 완료")
        except Exception as e: messagebox.showerror("오류", str(e))

    def send_notice(self):
        """현장으로 관리자 공지 보내기"""
        from tkinter import simpledialog
        notice = simpledialog.askstring("공지사항", "현장에 보낼 메시지를 입력하세요:")
        if notice:
            try:
                self.db.collection('notifications').add({
                    'admin': "관리자",
                    'message': notice,
                    'timestamp': firestore.SERVER_TIMESTAMP
                })
                messagebox.showinfo("성공", "현장으로 공지가 발송되었습니다!")
            except Exception as e:
                messagebox.showerror("에러", f"발송 실패: {e}")
    def load_messages(self):
        """공지와 현장 보고를 한 표에 다 보여주는 엔진"""
        if not hasattr(self, 'chat_tree'): return
        
        # 1. 표 비우기
        for item in self.chat_tree.get_children():
            self.chat_tree.delete(item)

        try:
            # 2. 관리자 공지사항(notifications) 가져오기
            notices = self.db.collection('notifications').order_by('timestamp', direction=firestore.Query.DESCENDING).limit(5).stream()
            for msg in notices:
                data = msg.to_dict()
                time_str = data.get('timestamp').strftime('%m/%d %H:%M') if data.get('timestamp') else "방금"
                # 공지는 상태칸에 '📢 공지'라고 표시하고 맨 앞에 넣음
                self.chat_tree.insert("", 0, values=("📢 공지", time_str, "관리자", data.get('message', '')))

            # 3. 현장 보고(field_reports) 가져오기
            reports = self.db.collection('field_reports').order_by('timestamp', direction=firestore.Query.DESCENDING).limit(15).stream()
            for msg in reports:
                data = msg.to_dict()
                doc_id = msg.id
                status = data.get('status', '⏳ 처리중')
                time_str = data.get('timestamp').strftime('%m/%d %H:%M') if data.get('timestamp') else "방금"
                
                # 보고 내용은 아래로 차곡차곡 쌓음
                self.chat_tree.insert("", tk.END, iid=doc_id, values=(status, time_str, data.get('user', '현장'), data.get('text', '')))

            print("✅ 공지 포함 전체 업데이트 완료!")
        except Exception as e:
            print(f"❌ 데이터 통합 로딩 실패: {e}")

    def _format_kst_time(self, ts):
        """Firestore timestamp를 한국시간 문자열로 변환. ts가 None/잘못된 값이면 빈 문자열."""
        if not ts:
            return ""
        try:
            from datetime import datetime, timezone, timedelta
            kst = timezone(timedelta(hours=9))
            # Firestore timestamp는 보통 datetime 객체 또는 .seconds/.nanoseconds 가짐
            if hasattr(ts, 'astimezone'):
                dt = ts.astimezone(kst)
            elif hasattr(ts, 'seconds'):
                dt = datetime.fromtimestamp(ts.seconds, tz=kst)
            else:
                return ""
            # 오늘이면 시간만, 다른 날이면 날짜+시간
            now = datetime.now(kst)
            if dt.date() == now.date():
                return dt.strftime("%H:%M")
            elif dt.year == now.year:
                return dt.strftime("%m/%d %H:%M")
            else:
                return dt.strftime("%Y/%m/%d %H:%M")
        except Exception:
            return ""

    def on_message_double_click(self, event):
        import requests
        from io import BytesIO
        from PIL import Image, ImageTk
        import os
        import threading
        from tkinter import messagebox

        selected_items = self.chat_tree.selection()
        if not selected_items: return
        
        selected_id = selected_items[0]
        doc_ref = self.db.collection('field_reports').document(selected_id)
        data = doc_ref.get().to_dict()
        img_urls_raw = data.get('imageUrl', '')
        report_ts = data.get('timestamp')

        # 1. 팝업창 설정 - 카톡 스타일
        detail_win = tk.Toplevel(self.root)
        detail_win.title(f"💬 {data.get('user', '익명')} - {data.get('title', '')[:20]}")
        screen_h = self.root.winfo_screenheight()
        win_h = min(900, int(screen_h * 0.85))
        self.position_popup(detail_win, 580, win_h)
        detail_win.configure(bg="#B2C7DA")  # 카톡 배경 푸르스름

        # ========== [카톡 스타일 상단 헤더 - 고정] ==========
        header = tk.Frame(detail_win, bg="white", height=70,
                          highlightthickness=0, bd=0)
        header.pack(side="top", fill="x")
        header.pack_propagate(False)

        # 작업자 아이콘 (원형 박스)
        avatar_frame = tk.Frame(header, bg="#1877F2", width=44, height=44)
        avatar_frame.place(x=15, y=13)
        avatar_frame.pack_propagate(False)
        tk.Label(avatar_frame, text="👤", bg="#1877F2", fg="white",
                 font=("맑은 고딕", 16)).pack(expand=True)

        # 제목 + 부제
        title_frame = tk.Frame(header, bg="white")
        title_frame.place(x=70, y=10)
        tk.Label(title_frame, text=data.get('title', '제목 없음'),
                 font=("맑은 고딕", 13, "bold"), bg="white", fg="#262626",
                 anchor="w").pack(anchor="w")
        sub_text = f"작업자: {data.get('user', '익명')}  ·  {data.get('status', '처리중')}"
        if report_ts:
            time_str = self._format_kst_time(report_ts)
            if time_str:
                sub_text += f"  ·  📅 {time_str}"
        tk.Label(title_frame, text=sub_text,
                 font=("맑은 고딕", 9), bg="white", fg="#65676B",
                 anchor="w").pack(anchor="w", pady=(3, 0))

        # 헤더 아래 구분선
        tk.Frame(detail_win, bg="#E1E4E8", height=1).pack(side="top", fill="x")

        # ========== [메인 스크롤 영역] ==========
        main_container = tk.Frame(detail_win, bg="#B2C7DA")
        main_container.pack(side="top", fill="both", expand=True)

        canvas = tk.Canvas(main_container, bg="#B2C7DA", highlightthickness=0)
        scrollbar = tk.Scrollbar(main_container, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="#B2C7DA")

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw", width=560)
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        # 하단 고정 영역
        footer = tk.Frame(detail_win, bg="white", pady=10, padx=15, bd=0,
                          highlightthickness=1, highlightbackground="#E1E4E8")
        footer.pack(side="bottom", fill="x")

        # 마우스 휠 스크롤 (Linux/Windows 모두 대응)
        def _safe_scroll(event):
            try:
                if canvas.winfo_exists():
                    canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            except: pass
        detail_win.bind("<MouseWheel>", _safe_scroll)
        canvas.bind("<MouseWheel>", _safe_scroll)
        scrollable_frame.bind("<MouseWheel>", _safe_scroll)

        # 우클릭 메뉴
        comment_context_menu = tk.Menu(detail_win, tearoff=0, font=("맑은 고딕", 10))

        # ========== [본문 카드 - 카톡 받은 메시지 스타일] ==========
        body_outer = tk.Frame(scrollable_frame, bg="#B2C7DA")
        body_outer.pack(fill="x", pady=(15, 5), padx=15)

        # 작업자 표시 (말풍선 위)
        tk.Label(body_outer, text=f"👤 {data.get('user', '익명')}",
                 font=("맑은 고딕", 9), bg="#B2C7DA", fg="#444",
                 anchor="w").pack(anchor="w", padx=(8, 0))

        # 본문 말풍선 (왼쪽 정렬)
        body_bubble_outer = tk.Frame(body_outer, bg="#B2C7DA")
        body_bubble_outer.pack(fill="x", anchor="w")

        body_bubble = tk.Frame(body_bubble_outer, bg="white",
                                highlightthickness=0)
        body_bubble.pack(side="left", anchor="w", padx=(0, 60))

        body_text = data.get('text', '')
        # 본문 길이에 따라 width 조정
        body_label = tk.Label(body_bubble, text=body_text,
                                bg="white", fg="#1A1A1A",
                                font=("맑은 고딕", 11), justify="left",
                                wraplength=400, padx=14, pady=10)
        body_label.pack(anchor="w")

        # 시간 (말풍선 옆 작게)
        if report_ts:
            tt = self._format_kst_time(report_ts)
            if tt:
                tk.Label(body_bubble_outer, text=tt,
                         bg="#B2C7DA", fg="#666",
                         font=("맑은 고딕", 8)).pack(side="left", anchor="s", padx=(4, 0), pady=(0, 4))

        # ========== [본문 사진들 - 본문 아래에 자연스럽게] ==========
        # 사진 PhotoImage 참조 유지 (GC 방지)
        if not hasattr(self, '_detail_img_cache'):
            self._detail_img_cache = []
        self._detail_img_cache.clear()

        # 사진 있을 때만 frame 생성
        photo_frame = None
        if img_urls_raw and img_urls_raw.strip():
            photo_frame = tk.Frame(scrollable_frame, bg="#B2C7DA")
            photo_frame.pack(fill="x", padx=15, pady=(0, 5))

        def load_images_async():
            if not img_urls_raw or not img_urls_raw.strip():
                return
            url_list = [u.strip() for u in img_urls_raw.split(",") if u.strip()]
            for i, url in enumerate(url_list):
                try:
                    if not detail_win.winfo_exists(): return
                    response = requests.get(url, timeout=10)
                    img_raw = Image.open(BytesIO(response.content))
                    img_raw.thumbnail((420, 500))
                    photo = ImageTk.PhotoImage(img_raw)
                    self._detail_img_cache.append(photo)

                    if detail_win.winfo_exists() and photo_frame is not None:
                        # 사진 컨테이너 (본문처럼 왼쪽 정렬)
                        pf = tk.Frame(photo_frame, bg="#B2C7DA")
                        pf.pack(fill="x", anchor="w", pady=4)

                        img_holder = tk.Frame(pf, bg="white", padx=2, pady=2)
                        img_holder.pack(side="left", anchor="w")
                        lbl = tk.Label(img_holder, image=photo, bg="white",
                                       cursor="hand2")
                        lbl.pack()
                        # 클릭하면 원본 보기
                        lbl.bind("<Button-1>", lambda e, u=url: __import__('webbrowser').open(u))

                        # 우클릭으로 저장 메뉴
                        def save_this(u=url, idx=i):
                            try:
                                res = requests.get(u)
                                desktop = os.path.join(os.path.expanduser("~"), "Desktop")
                                path = os.path.join(desktop, f"현장보고_{selected_id[:5]}_{idx+1}.jpg")
                                with open(path, "wb") as f: f.write(res.content)
                                messagebox.showinfo("완료", f"바탕화면에 저장됨:\n{path}", parent=detail_win)
                            except Exception as ex:
                                messagebox.showerror("실패", f"저장 실패: {ex}", parent=detail_win)

                        save_menu = tk.Menu(detail_win, tearoff=0, font=("맑은 고딕", 9))
                        save_menu.add_command(label="🔍 원본 크기로 보기",
                                               command=lambda u=url: __import__('webbrowser').open(u))
                        save_menu.add_command(label="💾 바탕화면에 저장",
                                               command=save_this)
                        lbl.bind("<Button-3>", lambda e, m=save_menu: m.post(e.x_root, e.y_root))
                except Exception as e:
                    if detail_win.winfo_exists() and photo_frame is not None:
                        tk.Label(photo_frame, text=f"❌ 사진 {i+1} 로드 실패: {e}",
                                 fg="#c00", bg="#B2C7DA",
                                 font=("맑은 고딕", 9)).pack(anchor="w", pady=2)

        threading.Thread(target=load_images_async, daemon=True).start()

        # ========== [댓글 영역 - 카톡 스타일 채팅] ==========
        # 구분선
        divider_frame = tk.Frame(scrollable_frame, bg="#B2C7DA")
        divider_frame.pack(fill="x", pady=(15, 5), padx=15)
        tk.Frame(divider_frame, bg="#9DB0C2", height=1).pack(fill="x", pady=8)
        tk.Label(divider_frame, text="💬 대화", bg="#B2C7DA", fg="#444",
                 font=("맑은 고딕", 9, "bold")).pack(anchor="center")

        comment_list_frame = tk.Frame(scrollable_frame, bg="#B2C7DA")
        comment_list_frame.pack(fill="x", padx=10, pady=(5, 15))

        # [댓글 개별 삭제 함수]
        def delete_this_comment(comment_id):
            if messagebox.askyesno("삭제", "이 댓글을 삭제하시겠습니까?"):
                doc_ref.collection('comments').document(comment_id).delete()
                refresh_comments()
                messagebox.showinfo("완료", "댓글이 삭제되었습니다.")

        # [추가] 댓글 수정 함수
        def edit_this_comment(comment_id, current_text):
            from tkinter import simpledialog
            new_text = simpledialog.askstring("댓글 수정", "내용을 수정하세요:",
                                                initialvalue=current_text, parent=detail_win)
            if new_text is not None and new_text.strip():
                doc_ref.collection('comments').document(comment_id).update({
                    'text': new_text.strip(),
                    'edited': True,
                    'editedAt': firestore.SERVER_TIMESTAMP
                })
                refresh_comments()

        # [추가] 인라인 이미지 캐시 (PIL 이미지 참조 유지용)
        if not hasattr(self, '_inline_img_cache'):
            self._inline_img_cache = []

        def refresh_comments():
            if not detail_win.winfo_exists(): return
            for w in comment_list_frame.winfo_children(): w.destroy()
            self._inline_img_cache.clear()

            comments = list(doc_ref.collection('comments').order_by('timestamp').stream())

            if not comments:
                tk.Label(comment_list_frame, text="아직 댓글이 없습니다",
                         bg="#B2C7DA", fg="#666",
                         font=("맑은 고딕", 9, "italic")).pack(pady=15)
                return

            for c in comments:
                cid = c.id
                d = c.to_dict()
                u = d.get('user', '관리자')
                t = d.get('text', '')
                img_url = d.get('imageUrl', '')
                edited = d.get('edited', False)
                ts = d.get('timestamp')

                if not detail_win.winfo_exists(): return

                # 작업자(받은) vs 관리자(보낸) 구분
                if "[답장]" in t:
                    is_admin = False
                    clean_text = t.replace("[답장] ", "").replace("[답장]", "").strip()
                else:
                    is_admin = True
                    clean_text = t

                if edited:
                    clean_text = clean_text + "  (수정됨)"
                # 사진만 있는 경우 (사진) 텍스트 처리
                if clean_text in ("(사진)", ""):
                    clean_text = ""  # 텍스트 없으면 빈칸

                time_str = self._format_kst_time(ts)

                # 말풍선 컨테이너 (카톡처럼 좌우 정렬)
                row = tk.Frame(comment_list_frame, bg="#B2C7DA")
                row.pack(fill="x", pady=4, padx=5)
                row.bind("<Button-3>", lambda e, id=cid, txt=t.replace("[답장] ", ""): show_comment_menu(e, id, txt))

                if is_admin:
                    # 관리자(나): 오른쪽 정렬, 노란 카톡 말풍선
                    bubble_color = "#FFEB33"
                    text_color = "#1A1A1A"
                    name_color = "#444"
                    align_side = "right"
                    text_anchor = "e"
                else:
                    # 작업자(상대): 왼쪽 정렬, 흰 말풍선
                    bubble_color = "white"
                    text_color = "#1A1A1A"
                    name_color = "#444"
                    align_side = "left"
                    text_anchor = "w"

                # [수정] 관리자/작업자 모두 이름 표시 (4명이 같이 쓰니까 누가 썼는지 봐야함)
                name_label = tk.Label(row, text=f"{'📢' if is_admin else '👤'} {u}",
                                       bg="#B2C7DA", fg=name_color,
                                       font=("맑은 고딕", 8, "bold"))
                name_label.pack(anchor=text_anchor, padx=(8 if is_admin else 2, 8 if is_admin else 0))
                name_label.bind("<Button-3>", lambda e, id=cid, txt=t.replace("[답장] ", ""): show_comment_menu(e, id, txt))

                # 말풍선 + 시간 묶음
                bubble_row = tk.Frame(row, bg="#B2C7DA")
                bubble_row.pack(fill="x", anchor=text_anchor)
                bubble_row.bind("<Button-3>", lambda e, id=cid, txt=t.replace("[답장] ", ""): show_comment_menu(e, id, txt))

                # 말풍선 컨테이너
                bubble_container = tk.Frame(bubble_row, bg="#B2C7DA")
                if is_admin:
                    bubble_container.pack(side="right", padx=(60, 0))
                else:
                    bubble_container.pack(side="left", padx=(0, 60))

                # 말풍선 본체
                bubble = tk.Frame(bubble_container, bg=bubble_color)
                if is_admin:
                    bubble.pack(side="right")
                else:
                    bubble.pack(side="left")
                bubble.bind("<Button-3>", lambda e, id=cid, txt=t.replace("[답장] ", ""): show_comment_menu(e, id, txt))

                # 텍스트가 있으면 라벨, 없으면 패스
                if clean_text:
                    text_label = tk.Label(bubble, text=clean_text,
                                           bg=bubble_color, fg=text_color,
                                           font=("맑은 고딕", 11),
                                           justify="left", wraplength=380,
                                           padx=12, pady=8)
                    text_label.pack(anchor="w" if not is_admin else "e")
                    text_label.bind("<Button-3>", lambda e, id=cid, txt=t.replace("[답장] ", ""): show_comment_menu(e, id, txt))

                # 첨부 이미지
                if img_url:
                    try:
                        import requests as _rq
                        from io import BytesIO
                        from PIL import Image, ImageTk
                        resp = _rq.get(img_url, timeout=10)
                        pil_img = Image.open(BytesIO(resp.content))
                        w, h = pil_img.size
                        max_w = 280
                        if w > max_w:
                            ratio = max_w / w
                            pil_img = pil_img.resize((max_w, int(h * ratio)), Image.LANCZOS)
                        tk_img = ImageTk.PhotoImage(pil_img)
                        self._inline_img_cache.append(tk_img)
                        # 사진은 말풍선과 별도 박스로 (카톡처럼)
                        img_holder = tk.Frame(bubble, bg=bubble_color, padx=4, pady=4)
                        img_holder.pack()
                        img_label = tk.Label(img_holder, image=tk_img,
                                              bg=bubble_color, cursor="hand2")
                        img_label.pack()
                        img_label.bind("<Button-1>", lambda e, url=img_url: __import__('webbrowser').open(url))
                        img_label.bind("<Button-3>", lambda e, id=cid, txt=t.replace("[답장] ", ""): show_comment_menu(e, id, txt))
                    except Exception as ex:
                        tk.Label(bubble, text=f"[사진 로드 실패]",
                                 bg=bubble_color, fg="#999",
                                 font=("맑은 고딕", 8), padx=12, pady=4).pack()

                # 시간 (말풍선 옆 작게, 카톡처럼)
                if time_str:
                    time_label = tk.Label(bubble_container, text=time_str,
                                           bg="#B2C7DA", fg="#555",
                                           font=("맑은 고딕", 8))
                    if is_admin:
                        time_label.pack(side="right", anchor="s", padx=(0, 4), pady=(0, 2),
                                        before=bubble)
                    else:
                        time_label.pack(side="left", anchor="s", padx=(4, 0), pady=(0, 2))

        def show_comment_menu(event, cid, current_text):
            comment_context_menu.delete(0, tk.END)
            comment_context_menu.add_command(label="✏️ 이 댓글 수정",
                                              command=lambda: edit_this_comment(cid, current_text))
            comment_context_menu.add_command(label="🗑️ 이 댓글 삭제",
                                              command=lambda: delete_this_comment(cid))
            comment_context_menu.post(event.x_root, event.y_root)

        refresh_comments()

        # 하단 입력창
        # [추가] 안내 문구
        tk.Label(footer, text="💡 Ctrl+V로 사진 붙여넣기 · Enter로 전송 · Shift+Enter로 줄바꿈",
                 bg="white", fg="#888", font=("맑은 고딕", 8), anchor="w").pack(fill="x", pady=(0, 2))

        # [수정] Entry → Text로 변경 (더 크고 여러 줄 입력 가능)
        entry_frame = tk.Frame(footer, bg="white", highlightthickness=1, highlightbackground="#CCC")
        entry_frame.pack(fill="x", pady=(0, 5))
        entry = tk.Text(entry_frame, font=("맑은 고딕", 12), bd=0, relief="flat",
                         height=3, padx=10, pady=8, wrap="word")
        entry.pack(fill="x", expand=True)

        # Text 위젯에서 .get/.delete를 Entry처럼 쓸 수 있도록 헬퍼 메서드 흉내
        # (송신 함수에서 entry.get(), entry.delete(0, tk.END) 같이 호출해도 동작하게)
        _orig_get = entry.get
        _orig_delete = entry.delete
        def _entry_get(*args, **kwargs):
            # Entry처럼 인자 없이 호출하면 전체 텍스트 반환
            if not args:
                return _orig_get("1.0", "end-1c")
            return _orig_get(*args, **kwargs)
        def _entry_delete(*args, **kwargs):
            # Entry.delete(0, tk.END) 호출 시 전체 삭제로 변환
            if len(args) >= 1 and args[0] == 0:
                return _orig_delete("1.0", "end")
            return _orig_delete(*args, **kwargs)
        entry.get = _entry_get
        entry.delete = _entry_delete

        # [추가] 사진 첨부 상태 (Entry 위젯에 attribute로 들고있음)
        self._pending_image_path = None
        # 썸네일 PhotoImage 참조 유지용
        self._pending_thumb_img = None

        # [추가] 첨부 사진 미리보기 영역 (큰 박스)
        photo_status_frame = tk.Frame(footer, bg="#FFF9E6", relief="flat",
                                       highlightthickness=1, highlightbackground="#FFD966")
        photo_status_frame.pack(fill="x", pady=(0, 5))
        # 평소엔 숨겨둠 (사진 첨부 시에만 표시)
        photo_status_frame.pack_forget()

        thumb_label = tk.Label(photo_status_frame, bg="#FFF9E6")
        thumb_label.pack(side="left", padx=8, pady=8)

        info_label = tk.Label(photo_status_frame, text="", bg="#FFF9E6",
                              fg="#444", font=("맑은 고딕", 10, "bold"))
        info_label.pack(side="left", padx=(0, 10))

        def clear_pending_image():
            self._pending_image_path = None
            self._pending_thumb_img = None
            thumb_label.config(image="")
            info_label.config(text="")
            photo_status_frame.pack_forget()

        cancel_btn = tk.Button(photo_status_frame, text="✕ 취소",
                                command=clear_pending_image, bg="#FF6B6B", fg="white",
                                font=("맑은 고딕", 9, "bold"), relief="flat", padx=10, pady=5,
                                cursor="hand2")
        cancel_btn.pack(side="right", padx=8, pady=8)

        def show_thumbnail(image_source):
            """이미지 경로 또는 PIL Image를 받아 썸네일 표시"""
            try:
                from PIL import Image, ImageTk
                if isinstance(image_source, str):
                    pil_img = Image.open(image_source)
                else:
                    pil_img = image_source
                # 썸네일 (60x60)
                pil_thumb = pil_img.copy()
                pil_thumb.thumbnail((60, 60), Image.LANCZOS)
                self._pending_thumb_img = ImageTk.PhotoImage(pil_thumb)
                thumb_label.config(image=self._pending_thumb_img)

                w, h = pil_img.size
                info_label.config(text=f"📎 첨부됨 ({w}×{h})\n전송 시 자동 업로드")
                # 박스 표시
                photo_status_frame.pack(fill="x", pady=(0, 5), before=entry.master.winfo_children()[-1] if False else None)
                # 위치 보정: btn_container 위에 놓이도록 다시 pack
                photo_status_frame.pack_forget()
                photo_status_frame.pack(fill="x", pady=(0, 5))
            except Exception as e:
                info_label.config(text=f"썸네일 생성 실패: {e}")

        def attach_image_from_clipboard():
            """클립보드의 이미지를 임시 파일로 저장 후 첨부"""
            try:
                from PIL import ImageGrab
                import tempfile
                img = ImageGrab.grabclipboard()
                if img is None:
                    messagebox.showinfo("알림", "클립보드에 이미지가 없습니다.\n캡처 도구로 캡처 후 시도해주세요.", parent=detail_win)
                    return False
                if isinstance(img, list):
                    if not img:
                        return False
                    self._pending_image_path = img[0]
                    show_thumbnail(img[0])
                else:
                    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                    img.save(tmp.name, "PNG")
                    self._pending_image_path = tmp.name
                    show_thumbnail(img)
                return True
            except Exception as e:
                messagebox.showerror("오류", f"클립보드 이미지 가져오기 실패:\n{e}", parent=detail_win)
                return False

        def attach_image_from_file():
            """파일 선택 다이얼로그로 이미지 첨부"""
            from tkinter import filedialog
            path = filedialog.askopenfilename(
                title="첨부할 이미지 선택",
                filetypes=[("이미지 파일", "*.png *.jpg *.jpeg *.gif *.bmp"), ("모든 파일", "*.*")],
                parent=detail_win
            )
            if path:
                self._pending_image_path = path
                show_thumbnail(path)

        # [추가] Ctrl+V로 클립보드 이미지 붙여넣기
        def on_paste(event):
            try:
                from PIL import ImageGrab
                img = ImageGrab.grabclipboard()
                if img is not None and not isinstance(img, list):
                    attach_image_from_clipboard()
                    return "break"
            except Exception:
                pass
            return None
        entry.bind("<Control-v>", on_paste)

        btn_container = tk.Frame(footer, bg="white")
        btn_container.pack(fill="x")

        def send_cmd():
            content = entry.get().strip()
            image_url = ""
            if self._pending_image_path:
                # 이미지 업로드
                image_url = self._upload_image_to_imgbb(self._pending_image_path, detail_win)
                if image_url is None:
                    return  # 업로드 실패 시 중단
            if not content and not image_url:
                return
            doc_ref.collection('comments').add({
                'user': self.current_user,
                'role': 'admin',
                'text': content if content else "(사진)",
                'imageUrl': image_url,
                'timestamp': firestore.SERVER_TIMESTAMP
            })
            entry.delete(0, tk.END)
            clear_pending_image()
            refresh_comments()
            # 푸시 알림 발송
            target = data.get('user', '')
            if target:
                preview = (content if content else "(사진)")[:80]
                if len(content) > 80: preview += '...'
                self.send_fcm_push(target,
                                    f"💬 {self.current_user}님의 댓글",
                                    preview)

        # [추가] Enter로 전송 (Shift+Enter는 줄바꿈)
        def _on_enter(event):
            if event.state & 0x0001:  # Shift 키
                return None  # 기본 동작 (줄바꿈)
            send_cmd()
            return "break"  # 기본 줄바꿈 막음
        entry.bind("<Return>", _on_enter)

        # [추가] 사진 첨부 버튼들 (Ctrl+V로 클립보드 이미지 자동 첨부)
        tk.Button(btn_container, text="📎 사진(파일)", command=attach_image_from_file,
                  bg="#f0f2f5", fg="#444", font=("맑은 고딕", 9),
                  relief="flat", pady=8).pack(side="left", padx=(0, 5))
        tk.Button(btn_container, text="🚀 댓글 전송", command=send_cmd, bg="#1877F2", fg="white", 
                  font=("맑은 고딕", 10, "bold"), relief="flat", pady=10).pack(side="left", expand=True, fill="x", padx=(0, 5))
        
        tk.Button(btn_container, text="✅ 완료 처리", bg="#4CAF50", fg="white", font=("맑은 고딕", 10, "bold"),
                  relief="flat", pady=10, command=lambda: self.update_status(selected_id, "✅ 완료", detail_win)).pack(side="left", expand=True, fill="x", padx=(5, 0))

    # --- 여기서부터는 on_message_double_click 함수 밖입니다 ---
     
    # 1. 실시간 감시 (음성 알림 버전 - 필터 및 이모지 대응 수정)
    def start_realtime_listener(self):
        import pyttsx3
        import threading
        from datetime import datetime, timedelta
        from google.cloud.firestore_v1.base_query import FieldFilter

        # [내부 함수] 음성 알림 (별도 쓰레드에서 실행하여 GUI 멈춤 방지)
        def speak_alert(text):
            def _speak():
                try:
                    engine = pyttsx3.init()
                    
                    # 1. 볼륨 설정: 0.0 ~ 1.0 (1.0이 최대)
                    # 혹시 모르니 1.0으로 확실히 박아줍니다.
                    engine.setProperty('volume', 1.0) 
                    
                    # 2. 말하기 속도: 너무 빠르면 작게 들릴 수 있으니 170~180 정도로 조절
                    engine.setProperty('rate', 170) 
                    
                    engine.say(text)
                    engine.runAndWait()
                except: pass
            threading.Thread(target=_speak, daemon=True).start()

        # 부팅 플래그 설정
        self.is_booting = True 

        # [리스너 콜백] ★여기서 self.root.winfo_exists() 등을 절대 호출하지 않음★
        def on_snapshot(col_snapshot, changes, read_time):
            new_report_detected = False
            for change in changes:
                # 새 문서(ADDED)가 들어왔을 때만 감지
                if change.type.name == 'ADDED' and not getattr(self, 'is_booting', True):
                    new_report_detected = True

            # 리스너 쓰레드에서 GUI 쓰레드로 안전하게 신호만 전달
            try:
                # 0.01초 뒤에 메인 쓰레드에서 UI 갱신 함수를 실행하도록 예약
                self.root.after(10, lambda: self._final_safe_refresh(new_report_detected, speak_alert))
            except:
                pass

        # [수정] 인덱스를 요구하던 복합 where+order_by 쿼리 제거.
        # 컬렉션 전체를 그냥 listen하고, 문서가 추가/수정/삭제되면 무조건 update_table_view() 재실행.
        try:
            query = self.db.collection('field_reports').limit(500)

            # 실시간 리스너 가동
            query.on_snapshot(on_snapshot)
            print(f"📡 실시간 리스너 가동 성공 (단순 모드, 컬렉션 전체 구독)")

            # 부팅 로딩 대기 (5초 후부터 알람 허용)
            self.root.after(5000, lambda: setattr(self, 'is_booting', False))

        except Exception as e:
            print(f"❌ 리스너 설정 오류: {e}")

    # [중요] 메인 쓰레드(GUI 전용)에서 안전하게 실행되는 함수
    def _final_safe_refresh(self, detected, alert_func):
        try:
            # 여기서는 메인 쓰레드이므로 winfo_exists() 체크가 안전함
            if hasattr(self, 'root') and self.root.winfo_exists():
                # 1. 화면 목록 새로고침
                self.update_table_view()

                # 2. 신규 확인 시 소리 알림
                if detected:
                    alert_func("신규 확인")
                    # [추가] 탭에 빨간 점 표시 (현재 작업보고 탭이 활성화되지 않았을 때만)
                    self.set_tab_alert(self.t_field, True)
        except Exception as e:
            print(f"❌ UI 리프레시 오류: {e}")

    # --- [탭 알림 시스템] ---
    # 각 탭 옆에 🔴 빨간 점을 붙였다 떼는 기능. 데이터 사용량 영향 0
    # (이미 돌고 있는 실시간 리스너가 신호만 전달, 추가 read 없음)
    def set_tab_alert(self, tab_widget, on):
        """특정 탭에 빨간 점 알림 표시/제거.
        다만 사용자가 그 탭을 보고 있으면 알림을 표시하지 않음 (이미 읽고 있으니까)."""
        try:
            # 현재 보고 있는 탭이면 알림 안 띄움
            current_tab_id = self.nb.select()
            target_tab_id = str(tab_widget)
            if on and current_tab_id == target_tab_id:
                return

            # 기존 텍스트에서 🔴 떼고 새로 붙임
            current_text = self.nb.tab(tab_widget, "text")
            clean_text = current_text.replace("🔴", "").rstrip()
            if on:
                # 오른쪽에 붙임 (공백 포함된 원래 패딩 유지)
                self.nb.tab(tab_widget, text=f"{clean_text} 🔴")
            else:
                # 원래 텍스트로 복원 (앞뒤 공백 패딩 살리기 위해 dict로 관리)
                original = self._original_tab_texts.get(str(tab_widget), clean_text)
                self.nb.tab(tab_widget, text=original)
        except Exception as e:
            print(f"⚠️ 탭 알림 설정 실패: {e}")

    def setup_tab_alert_system(self):
        """프로그램 시작 시 한번 호출. 원본 탭 텍스트 저장 + 탭 클릭 시 알림 자동 제거."""
        self._original_tab_texts = {}
        for tab_widget in [self.t_field, self.t_board]:
            try:
                self._original_tab_texts[str(tab_widget)] = self.nb.tab(tab_widget, "text")
            except Exception:
                pass

        # 탭 변경 이벤트: 클릭한 탭의 알림 자동 제거
        def _on_tab_changed(event):
            try:
                selected_id = self.nb.select()
                for tab_widget in [self.t_field, self.t_board]:
                    if str(tab_widget) == selected_id:
                        # 본 탭이니 알림 끄기
                        self.set_tab_alert(tab_widget, False)
            except Exception:
                pass
        self.nb.bind("<<NotebookTabChanged>>", _on_tab_changed)

    # 2. 상태 업데이트 (완료 버튼 클릭 시 실행)
    # [수정] 이 함수가 start_realtime_listener와 같은 세로 라인이어야 합니다.
    def update_status(self, doc_id, status, window):
        try:
            self.db.collection('field_reports').document(doc_id).update({'status': status})
            messagebox.showinfo("성공", f"상태가 {status}로 변경되었습니다.")
            window.destroy() # 상세 팝업창 닫기
        except Exception as e:
            messagebox.showerror("오류", f"상태 변경 실패: {e}")

# 3. 데이터 목록 갱신 엔진 (오류 수정 및 새 댓글 감지)
    def update_table_view(self, event=None):
        from datetime import datetime, timedelta
        from google.cloud.firestore_v1.base_query import FieldFilter

        # 1. 화면 초기화
        for i in self.chat_tree.get_children():
            self.chat_tree.delete(i)

        try:
            search_keyword = self.search_var.get().strip().lower()
            current_filter = self.filter_var.get()
            ref = self.db.collection('field_reports')

            # 2. 데이터 가져오기
            # [수정] order_by/where가 timestamp 필드 없는 문서를 배제하거나 인덱스를 요구하던 이슈가 있어
            # 일단 전체를 가져와서 파이썬에서 정렬하는 방식으로 변경. 필드가 없는 옛 문서도 안 잘림.
            try:
                raw_docs = list(ref.limit(500).get())
            except Exception as fetch_err:
                print(f"❌ Firestore 조회 실패: {fetch_err}")
                messagebox.showerror("데이터 로딩 실패", f"Firestore 조회 중 오류:\n{fetch_err}")
                return

            def _ts_key(d):
                ts = d.to_dict().get('timestamp')
                try:
                    return ts.timestamp() if ts else 0
                except Exception:
                    return 0
            raw_docs.sort(key=_ts_key, reverse=True)

            # 검색 모드면 더 많이, 일반 모드면 최근 100건만
            reports = raw_docs if search_keyword else raw_docs[:100]
            print(f"[디버그] field_reports 조회: 전체 {len(raw_docs)}건, 표시 대상 {len(reports)}건, 필터={current_filter!r}")
            for _d in raw_docs[:10]:
                _dd = _d.to_dict()
                _ts = _dd.get('timestamp')
                print(f"  └ id={_d.id[:8]} status={_dd.get('status')!r} user={_dd.get('user')!r} date={_dd.get('date')!r} ts={_ts}")

            # 3. 데이터 루프
            for idx, doc in enumerate(reports):
                data = doc.to_dict()
                status = data.get('status', '⏳ 처리중')
                user = data.get('user', '익명')
                title = data.get('title', '(제목 없음)')
                ts = data.get('timestamp')
                time_str = data.get('date') if data.get('date') else (ts.strftime('%y/%m/%d') if ts else "")

                # [검색어 필터링]
                if search_keyword:
                    clean_date = time_str.replace("/", "")
                    combined = f"{time_str} {clean_date} {user} {title}".lower()
                    if search_keyword not in combined:
                        continue

                # [상태 필터링]
                if current_filter != "전체" and status != current_filter:
                    continue

                # 4. 진행 상태 표시
                step_text = "🆕 신규"
                comments_ref = self.db.collection('field_reports').document(doc.id).collection('comments')
                last_c = comments_ref.order_by('timestamp', direction=firestore.Query.DESCENDING).limit(1).get()

                is_worker_reply = False
                if last_c:
                    c_data = last_c[0].to_dict()
                    if "[답장]" in c_data.get('text', ''):
                        step_text = "💬 NEW댓글"; is_worker_reply = True
                    else:
                        step_text = "✅ 확인중"

                # 5. 표에 삽입
                row_tag = "evenrow" if idx % 2 == 0 else "oddrow"
                item_id = self.chat_tree.insert("", "end", iid=doc.id,
                                                values=(status, step_text, time_str, user, title),
                                                tags=(row_tag,))

                # --- [색상 우선순위 수정] 완료를 가장 먼저 체크합니다 ---

                # 1. 완료: 가장 최우선 (이미 끝난 일은 무조건 회색)
                if status == "✅ 완료":
                    self.chat_tree.tag_configure("done_gray", foreground="#888888", background="white")
                    self.chat_tree.item(item_id, tags=("done_gray",))

                # 2. NEW댓글: (완료되지 않은 것 중 답장 온 것)
                elif is_worker_reply:
                    self.chat_tree.tag_configure("sky_blue", foreground="#0078D7", background="#E1F5FE")
                    self.chat_tree.item(item_id, tags=("sky_blue",))

                # 3. 확인중: (완료되지 않은 것 중 내가 확인한 것)
                elif step_text == "✅ 확인중":
                    self.chat_tree.tag_configure("light_green", foreground="#328837", background="#E8F5E9")
                    self.chat_tree.item(item_id, tags=("light_green",))

                # 4. 신규: (완료되지 않은 것 중 아직 안 본 것)
                elif step_text == "🆕 신규":
                    self.chat_tree.tag_configure("light_red", foreground="#C62828", background="#FFEBEE")
                    self.chat_tree.item(item_id, tags=("light_red",))

        except Exception as e:
            import traceback
            print(f"❌ 검색 오류: {e}")
            traceback.print_exc()
            try:
                messagebox.showerror("작업보고 로딩 오류", f"{type(e).__name__}: {e}")
            except Exception:
                pass

    def change_user_name(self):
        from tkinter import simpledialog, messagebox
        current = getattr(self, 'current_user', '관리자')

        new_name = simpledialog.askstring("사용자 설정", "사용하실 이름을 입력하세요:",
                                          initialvalue=current)
        if new_name:
            new_name = new_name.strip()
            if not new_name:
                return
            self.current_user = new_name
            self.save_user_name(new_name)  # [추가] 영구 저장 (업데이트 후에도 유지)
            messagebox.showinfo("완료", f"이제부터 '{new_name}' 이름으로 댓글이 달립니다.\n(다음 실행에도 유지됩니다)")

# --- 여기서부터는 클래스 밖 ---
if __name__ == "__main__":
    root = tk.Tk()
    app = LogiPanApp(root)
    root.mainloop()
