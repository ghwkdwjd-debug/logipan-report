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
        self.root.title("로지판 (Logi-Pan) - 통합 물류 파트너")

        # [추가] 창 아이콘 - 박스 이모지로 동적 생성 (PIL 사용)
        try:
            from PIL import Image, ImageDraw, ImageFont, ImageTk
            # 256x256 박스 아이콘
            img = Image.new('RGBA', (256, 256), (0, 0, 0, 0))
            draw = ImageDraw.Draw(img)
            # 시스템 이모지 폰트 시도 (윈도우)
            try:
                font = ImageFont.truetype("seguiemj.ttf", 200)  # Windows Segoe UI Emoji
            except:
                try:
                    font = ImageFont.truetype("AppleColorEmoji.ttc", 180)  # Mac
                except:
                    font = None
            if font:
                # 박스 이모지 📦
                draw.text((28, 0), "📦", font=font, embedded_color=True)
            self._icon_photo = ImageTk.PhotoImage(img)
            self.root.iconphoto(True, self._icon_photo)
        except Exception as e:
            print(f"⚠️ 아이콘 설정 실패: {e}")

        # [추가] 백그라운드 자가 업데이트
        # 한 번 실행되면 다음부터는 스킵 (마커 파일로 체크)
        # - updater.py 새 버전 다운로드 (옛날 버전 사용자도 자동 업그레이드)
        # - 로지판.ico 다운로드 (없을 때만)
        # - 바탕화면에 박스 아이콘 바로가기 생성 (없을 때만)
        try:
            import threading
            threading.Thread(target=self._self_update_check, daemon=True).start()
        except Exception as e:
            print(f"⚠️ 자가 업데이트 시작 실패: {e}")

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
        # 화면 크기에 비례. 노트북에서도 데스크탑에서도 적당하게.
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        # 화면의 60% 정도 (최대는 980x800)
        width = min(980, int(sw * 0.62))
        height = min(800, int(sh * 0.78))
        self.root.geometry(f"{width}x{height}+{(sw-width)//2}+{(sh-height)//2}")
        # 최소 크기
        self.root.minsize(820, 680)
        self.root.configure(bg="#F5F6F8")        
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
        # [수정] 각 섹션별 파일 따로 관리 (master/inbound/exclude/out 4개 영역)
        self.moms_files = {
            "master_send": "",   # 1. 마스터 생성용 출고리스트
            "master_master": "", # 1. 마스터 생성용 맘스 마스터재고
            "inbound_send": "",  # 2. 입고리스트용 출고리스트
            "excl_send": "",     # 3. 제외 입고리스트용 출고리스트
            "excl_list": "",     # 3. 제외 명단 파일
            "out_list": ""       # 4. 출고등록용
        }
        self.chk_files = {"target": "", "master": ""}
        self.filter_var = tk.StringVar(value="전체")
        self.search_var = tk.StringVar()
        # [수정] current_user를 config 파일에서 로드 (없으면 기본값 "장정호")
        self.current_user = self.load_user_name()
        self.start_realtime_listener()

        self.style = ttk.Style()
        self.style.theme_use('clam')

        # ========== [모던 탭 스타일] ==========
        # 노트북 자체 (탭 영역 배경)
        self.style.configure("TNotebook",
                              background="#F5F6F8",
                              borderwidth=0,
                              tabmargins=[12, 10, 12, 0])
        # 탭 (기본 상태) - 패딩 줄여서 8개 탭 다 보이게
        self.style.configure("TNotebook.Tab",
                              padding=[12, 8],
                              font=("맑은 고딕", 9, "bold"),
                              background="#F5F6F8",
                              foreground="#6B7280",
                              borderwidth=0,
                              focuscolor="#F5F6F8")
        # 탭 (선택/호버 상태)
        self.style.map("TNotebook.Tab",
                        background=[("selected", "white"),
                                    ("active", "#E5E7EB")],
                        foreground=[("selected", "#1877F2"),
                                    ("active", "#374151")],
                        expand=[("selected", [1, 1, 1, 0])])
        # 탭 영역 아래 구분선 제거 + 컨텐츠 영역 배경
        self.style.layout("TNotebook.Tab", [
            ("Notebook.tab", {
                "sticky": "nswe",
                "children": [
                    ("Notebook.padding", {
                        "side": "top",
                        "sticky": "nswe",
                        "children": [
                            ("Notebook.label", {"side": "top", "sticky": ""})
                        ]
                    })
                ]
            })
        ])

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

        self.t_in = ttk.Frame(self.nb); self.nb.add(self.t_in, text="📥 입고")
        self.t_out = ttk.Frame(self.nb); self.nb.add(self.t_out, text="📤 출고")
        self.t_mom = ttk.Frame(self.nb); self.nb.add(self.t_mom, text="📦 맘스")
        self.t_master = ttk.Frame(self.nb); self.nb.add(self.t_master, text="🆕 마스터")
        self.t_rt = ttk.Frame(self.nb); self.nb.add(self.t_rt, text="🔄 RT입고")
        self.t_chk = ttk.Frame(self.nb); self.nb.add(self.t_chk, text="🔍 재고파악")
        self.t_board = tk.Frame(self.nb); self.nb.add(self.t_board, text="📢 공지/소통")
        self.t_field = tk.Frame(self.nb); self.nb.add(self.t_field, text="📋 작업보고")
        self.t_end = ttk.Frame(self.nb); self.nb.add(self.t_end, text="📊 마감재고")

        self.setup_inbound()
        self.setup_outbound()
        self.setup_moms_v86()
        self.setup_master_registration()
        self.setup_rt_inbound()
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

    def _clear_input_box(self, text_widget, label_widget, base_text):
        """입력창 비우기 (확인 후). 내용 없으면 그냥 패스."""
        existing = text_widget.get("1.0", "end-1c").strip()
        if not existing:
            return
        if not messagebox.askyesno("입력 비우기",
                                     f"입력창의 모든 내용을 지우시겠습니까?",
                                     parent=self.root):
            return
        text_widget.delete("1.0", tk.END)
        # 카운트 라벨 리셋
        try:
            label_widget.config(text=f"{base_text} (0개)")
        except Exception:
            pass

    def _is_qty_token(self, tok):
        """수량 토큰인가? 0 이상의 정수 문자열."""
        return tok.isdigit()

    def _looks_like_location(self, tok):
        """로케이션 형식 문자열인가? 'DD-02-06-03' 또는 '00-00-00-00' 같은 패턴."""
        # 하이픈으로 구분된 토큰이고 영숫자 그룹이 3개 이상이면 로케이션
        return bool(re.match(r'^[A-Za-z0-9]+(-[A-Za-z0-9]+){2,}$', tok))

    def _looks_like_barcode(self, tok):
        """바코드처럼 생긴 문자열인가? - 길이 5+ 영숫자 또는 숫자만 6+자리.
        목적: 엑셀 행 복붙 시 첫 토큰이 바코드인지 판단."""
        tok = tok.strip()
        if not tok:
            return False
        # 한글 들어가있으면 바코드 아님
        if any(ord(c) > 127 for c in tok):
            return False
        # 너무 짧으면 바코드 아님
        if len(tok) < 5:
            return False
        # 영숫자만 (하이픈 허용)
        if not re.match(r'^[A-Za-z0-9\-]+$', tok):
            return False
        # 순수 숫자면 6자리 이상
        if tok.isdigit() and len(tok) < 6:
            return False
        return True

    def _find_qty_in_tokens(self, tokens):
        """토큰 리스트에서 '수량'에 해당하는 숫자를 찾기.
        엑셀 행 복붙: 바코드 상품명 사이즈 ... 수량 형태일 때
        - 우선순위 1: 마지막 토큰이 숫자면 그게 수량 (가장 흔한 패턴)
        - 우선순위 2: 1~9999 사이 숫자 중 마지막 거 (수량이라 작은 숫자)
        - 우선순위 3: 못 찾으면 1
        """
        if not tokens:
            return 1

        # 우선순위 1: 마지막 토큰 (가장 흔한 패턴)
        last = tokens[-1].strip()
        if last.isdigit():
            n = int(last)
            if 0 < n < 100000:  # 너무 큰 숫자(바코드 같은)는 제외
                return n

        # 우선순위 2: 끝쪽부터 거꾸로 첫 작은 숫자
        for tok in reversed(tokens):
            tok = tok.strip()
            if tok.isdigit():
                n = int(tok)
                if 0 < n < 10000:
                    return n

        # 못 찾음
        return 1

    def parse_logi_data(self, text_widget):
        """입력을 자동 감지해서 DataFrame으로 반환.

        지원 형식:
        - 바코드만: BARCODE1\nBARCODE2 → 각 1개로
        - 바코드+수량: BARCODE1 5\nBARCODE2 3 → 그대로
        - 5열 (스캔칸 전용): BARCODE 정상수 불량수 정상로케 불량로케
            → 같은 바코드는 (정상+불량) 합쳐서 표시 (대조용)
            → 정상/불량 분리 데이터는 self.last_scan_detail에 보존 (입고파일 생성용)
        - [추가] 엑셀 행 복붙: BARCODE<TAB>상품명<TAB>...<TAB>수량
            → 첫 토큰을 바코드로, 마지막 숫자 토큰을 수량으로 자동 추출

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
            elif len(tokens) >= 3 and self._looks_like_barcode(tokens[0]):
                # [추가] 엑셀 행 복붙 처리: 3개 이상 토큰 + 첫 토큰이 바코드 같음
                # → 첫 토큰을 바코드로, 토큰들 중 "수량스러운 숫자"를 찾아 수량으로 사용
                bc = self.normalize_barcode(tokens[0])
                qty = self._find_qty_in_tokens(tokens[1:])
                rows.append({'바코드': bc, '정상수': qty,
                             '불량수': 0, '정상로케': '', '불량로케': ''})
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

    def _self_update_check(self):
        """[추가] 백그라운드 자가 업데이트 (조용히 실행)
        - updater.py 새 버전 다운로드 (있을 때마다)
        - 로지판.ico 다운로드 (없을 때만)
        - 바탕화면 바로가기 생성 (없을 때만)
        모든 작업은 실패해도 무시 (메인 기능에 영향 X)
        """
        if os.name != 'nt':
            return  # 윈도우만

        try:
            import urllib.request

            base_url = "https://raw.githubusercontent.com/ghwkdwjd-debug/logipan-report/main"
            script_dir = os.path.dirname(os.path.abspath(__file__))

            # 1. 로지판.ico 다운로드 (없을 때만)
            ico_path = os.path.join(script_dir, "로지판.ico")
            if not os.path.exists(ico_path):
                try:
                    ico_url = f"{base_url}/%EB%A1%9C%EC%A7%80%ED%8C%90.ico"
                    urllib.request.urlretrieve(ico_url, ico_path)
                    print("✅ 로지판.ico 다운로드 완료")
                except Exception as e:
                    print(f"⚠️ ico 다운로드 실패 (무시): {e}")

            # 2. updater.py 새 버전 체크 (한 번만)
            # 마커 파일로 한 번 했는지 표시 - 다음 실행부터 스킵
            updater_marker = os.path.join(script_dir, ".updater_v2_done")
            updater_path = os.path.join(script_dir, "updater.py")
            if not os.path.exists(updater_marker) and os.path.exists(updater_path):
                try:
                    updater_url = f"{base_url}/updater.py"
                    tmp_path = updater_path + ".new"
                    urllib.request.urlretrieve(updater_url, tmp_path)

                    # 다운받은 파일이 새 버전인지 확인 (ICON_URL 키워드 있으면 신버전)
                    with open(tmp_path, 'r', encoding='utf-8') as f:
                        new_content = f.read()
                    if 'ICON_URL' in new_content and 'ensure_desktop_shortcut' in new_content:
                        # 안전하게 .bak으로 백업 후 교체
                        bak_path = updater_path + ".bak"
                        try:
                            if os.path.exists(bak_path):
                                os.remove(bak_path)
                            os.rename(updater_path, bak_path)
                            os.rename(tmp_path, updater_path)
                            # 마커 파일 생성
                            with open(updater_marker, 'w', encoding='utf-8') as f:
                                f.write("done")
                            print("✅ updater.py 새 버전으로 교체됨")
                        except Exception as e:
                            print(f"⚠️ updater.py 교체 실패 (무시): {e}")
                            # 교체 실패시 .new 파일 삭제
                            if os.path.exists(tmp_path):
                                try: os.remove(tmp_path)
                                except: pass
                    else:
                        # 새 버전 아니면 그냥 삭제
                        try: os.remove(tmp_path)
                        except: pass
                except Exception as e:
                    print(f"⚠️ updater.py 다운로드 실패 (무시): {e}")

            # 3. 바탕화면 바로가기 생성 (없을 때만)
            try:
                self._ensure_desktop_shortcut(script_dir, ico_path)
            except Exception as e:
                print(f"⚠️ 바로가기 생성 실패 (무시): {e}")

        except Exception as e:
            print(f"⚠️ 자가 업데이트 전체 실패 (무시): {e}")

    def _ensure_desktop_shortcut(self, script_dir, ico_path):
        """바탕화면에 로지판 바로가기가 없으면 자동 생성"""
        if os.name != 'nt':
            return

        # 바탕화면 경로
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        if not os.path.isdir(desktop):
            try:
                import ctypes.wintypes
                CSIDL_DESKTOP = 0
                buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
                ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_DESKTOP, None, 0, buf)
                desktop = buf.value
            except Exception:
                return
        if not os.path.isdir(desktop):
            return

        shortcut_path = os.path.join(desktop, "로지판.lnk")
        bat_path = os.path.join(script_dir, "로지판.bat")

        # 이미 있고 아이콘이 박스로 설정되어 있으면 스킵
        # (간단하게 .lnk 존재 + .ico 존재 모두 만족하면 OK로 간주)
        if os.path.exists(shortcut_path) and os.path.exists(ico_path):
            # 이미 있음 - 스킵
            return

        if not os.path.exists(bat_path):
            return  # .bat 없으면 만들 필요 없음

        # PowerShell로 .lnk 생성
        ps_script = (
            f'$WshShell = New-Object -ComObject WScript.Shell;'
            f'$Shortcut = $WshShell.CreateShortcut("{shortcut_path}");'
            f'$Shortcut.TargetPath = "{bat_path}";'
            f'$Shortcut.WorkingDirectory = "{script_dir}";'
            f'$Shortcut.IconLocation = "{ico_path},0";'
            f'$Shortcut.WindowStyle = 7;'
            f'$Shortcut.Description = "로지판 (Logi-Pan)";'
            f'$Shortcut.Save()'
        )
        import subprocess
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        subprocess.run(
            ["powershell", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
            capture_output=True, text=True,
            startupinfo=startupinfo, timeout=10
        )
        print("✅ 바탕화면에 박스 아이콘 바로가기 생성")

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
        container = tk.Frame(self.t_in, bg="#F5F6F8")
        container.pack(fill="both", expand=True)

        # ========== [상단 헤더] ==========
        header_frame = tk.Frame(container, bg="#F5F6F8")
        header_frame.pack(side="top", fill="x", padx=24, pady=(16, 8))

        title_left = tk.Frame(header_frame, bg="#F5F6F8")
        title_left.pack(side="left")
        tk.Label(title_left, text="📥", font=("맑은 고딕", 22),
                 bg="#F5F6F8").pack(side="left", padx=(0, 6))
        title_text_box = tk.Frame(title_left, bg="#F5F6F8")
        title_text_box.pack(side="left")
        tk.Label(title_text_box, text="입고 등록",
                 font=("맑은 고딕", 15, "bold"),
                 bg="#F5F6F8", fg="#1A1A1A").pack(anchor="w")
        tk.Label(title_text_box, text="브랜드 수량 vs 스캔 수량 대조 후 입고",
                 font=("맑은 고딕", 8),
                 bg="#F5F6F8", fg="#888").pack(anchor="w")

        # [추가] 우측 작은 리셋 버튼 (눈에 안 띄게 회색)
        reset_btn = tk.Button(header_frame, text="🔄 리셋",
                                command=self.reset_inbound_all,
                                bg="#F3F4F6", fg="#9CA3AF",
                                activebackground="#E5E7EB",
                                font=("맑은 고딕", 8),
                                relief="flat", padx=8, pady=4,
                                cursor="hand2")
        reset_btn.pack(side="right", anchor="se")

        # ========== [💾 액션 버튼들 - 하단 고정] ==========
        action_outer = tk.Frame(container, bg="#F5F6F8")
        action_outer.pack(side="bottom", fill="x", padx=18, pady=(0, 14))

        def make_modern_btn(parent, text, bg, hover_bg, command):
            shadow = tk.Frame(parent, bg=hover_bg)
            shadow.pack(side="left", expand=True, fill="x", padx=2)
            btn = tk.Button(shadow, text=text,
                              bg=bg, fg="white",
                              activebackground=hover_bg, activeforeground="white",
                              font=("맑은 고딕", 10, "bold"),
                              relief="flat", bd=0,
                              cursor="hand2", command=command)
            btn.pack(fill="x", ipady=8)
            def on_enter(e): btn.config(bg=hover_bg)
            def on_leave(e): btn.config(bg=bg)
            btn.bind("<Enter>", on_enter)
            btn.bind("<Leave>", on_leave)
            return btn

        make_modern_btn(action_outer, "🔍  대조 분석",
                         bg="#F59E0B", hover_bg="#D97706",
                         command=self.run_compare_in)
        make_modern_btn(action_outer, "📋  입고파일 생성",
                         bg="#1877F2", hover_bg="#1864c8",
                         command=self.create_inbound_file)
        make_modern_btn(action_outer, "📎  클립보드 복사",
                         bg="#8B5CF6", hover_bg="#7C3AED",
                         command=self.copy_inbound_to_clipboard)
        make_modern_btn(action_outer, "💾  CSV 저장",
                         bg="#10B981", hover_bg="#059669",
                         command=self.save_csv_in)

        # ========== [📝 리포트 카드 - 하단 (버튼 위)] ==========
        report_card_outer = tk.Frame(container, bg="#F5F6F8")
        report_card_outer.pack(side="bottom", fill="x", padx=18, pady=(0, 8))

        report_card = tk.Frame(report_card_outer, bg="white",
                                highlightthickness=1, highlightbackground="#E5E7EB")
        report_card.pack(fill="x")

        tk.Frame(report_card, bg="#8B5CF6", width=4).pack(side="left", fill="y")

        report_inner = tk.Frame(report_card, bg="white", padx=14, pady=10)
        report_inner.pack(side="left", fill="both", expand=True)

        # 리포트 헤더
        rh = tk.Frame(report_inner, bg="white")
        rh.pack(fill="x", pady=(0, 6))
        tk.Label(rh, text="📝  대조 분석 상세 리포트",
                 bg="white", fg="#111827",
                 font=("맑은 고딕", 10, "bold")).pack(side="left")
        tk.Label(rh, text="※ 엔터 2번 시 스캔수량으로 이동",
                 bg="white", fg="#9CA3AF",
                 font=("맑은 고딕", 8)).pack(side="right")

        self.txt_in_report = tk.Text(report_inner, height=6,
                                       font=("Consolas", 10),
                                       bg="#FAFAFA", bd=1, relief="solid",
                                       highlightthickness=1, highlightbackground="#E5E7EB",
                                       padx=8, pady=6)
        self.txt_in_report.pack(fill="x")
        self.txt_in_report.tag_config("match", foreground="#16A34A", font=("Consolas", 10, "bold"))
        self.txt_in_report.tag_config("error", foreground="#DC2626", font=("Consolas", 10, "bold"))
        self.txt_in_report.tag_config("title", foreground="#1877F2", font=("Consolas", 10, "bold"))
        self.txt_in_report.tag_config("info", foreground="#555", font=("Consolas", 10))
        self.txt_in_report.tag_config("warn", foreground="#F59E0B", font=("Consolas", 10, "bold"))

        # ========== [브랜드명 입력 카드] ==========
        brand_card_outer = tk.Frame(container, bg="#F5F6F8")
        brand_card_outer.pack(fill="x", padx=18, pady=(0, 8))

        brand_card = tk.Frame(brand_card_outer, bg="white",
                                highlightthickness=1, highlightbackground="#E5E7EB")
        brand_card.pack(fill="x")

        tk.Frame(brand_card, bg="#3B82F6", width=4).pack(side="left", fill="y")

        brand_inner = tk.Frame(brand_card, bg="white", padx=14, pady=10)
        brand_inner.pack(side="left", fill="both", expand=True)

        tk.Label(brand_inner, text="🔖",
                 bg="white", font=("맑은 고딕", 12)).pack(side="left", padx=(0, 6))
        tk.Label(brand_inner, text="브랜드명",
                 bg="white", fg="#374151",
                 font=("맑은 고딕", 10, "bold")).pack(side="left", padx=(0, 8))
        self.ent_brand_in = tk.Entry(brand_inner, font=("맑은 고딕", 11),
                                       bd=1, relief="solid",
                                       highlightthickness=0)
        self.ent_brand_in.pack(side="left", fill="x", expand=True, ipady=5)

        # ========== [📦 브랜드 수량 + 📡 스캔 수량 - 2분할] ==========
        cols_outer = tk.Frame(container, bg="#F5F6F8")
        cols_outer.pack(fill="both", expand=True, padx=18, pady=(0, 8))
        cols_outer.columnconfigure(0, weight=1, uniform="cols")
        cols_outer.columnconfigure(1, weight=1, uniform="cols")

        # === 좌: 브랜드 수량 카드 ===
        l_card_outer = tk.Frame(cols_outer, bg="#F5F6F8")
        l_card_outer.grid(row=0, column=0, sticky="nsew", padx=(0, 4))

        l_card = tk.Frame(l_card_outer, bg="white",
                           highlightthickness=1, highlightbackground="#E5E7EB")
        l_card.pack(fill="both", expand=True)

        tk.Frame(l_card, bg="#0EA5E9", width=4).pack(side="left", fill="y")

        l_inner = tk.Frame(l_card, bg="white", padx=12, pady=10)
        l_inner.pack(side="left", fill="both", expand=True)

        l_head = tk.Frame(l_inner, bg="white")
        l_head.pack(fill="x", pady=(0, 4))
        tk.Label(l_head, text="📦",
                 bg="white", font=("맑은 고딕", 11)).pack(side="left", padx=(0, 4))
        self.lbl_in_m = tk.Label(l_head, text="브랜드 수량 (0개)",
                                   bg="white", fg="#0C4A6E",
                                   font=("맑은 고딕", 10, "bold"))
        self.lbl_in_m.pack(side="left")
        # [추가] 우측 지우기 버튼
        clr_m = tk.Button(l_head, text="🗑️ 지우기",
                           command=lambda: self._clear_input_box(self.txt_in_master, self.lbl_in_m, "브랜드 수량"),
                           bg="#F3F4F6", fg="#6B7280",
                           activebackground="#E5E7EB",
                           font=("맑은 고딕", 8, "bold"),
                           relief="flat", padx=8, pady=2,
                           cursor="hand2")
        clr_m.pack(side="right")

        # [추가] 안내
        tk.Label(l_inner, text="💡 엑셀 행 복붙 OK (바코드+상품명+수량 자동 인식)",
                 bg="white", fg="#9CA3AF",
                 font=("맑은 고딕", 8), anchor="w").pack(fill="x", pady=(0, 4))

        # 입력창 + 스크롤바 컨테이너 (고정 높이)
        master_box = tk.Frame(l_inner, bg="white")
        master_box.pack(fill="both", expand=True)

        self.txt_in_master = tk.Text(master_box, font=("Consolas", 11),
                                       bd=1, relief="solid",
                                       bg="#F0F9FF",
                                       highlightthickness=1, highlightbackground="#E5E7EB",
                                       padx=6, pady=6, wrap="char",
                                       height=15)  # 고정 높이 - 넘으면 스크롤로
        master_sb = ttk.Scrollbar(master_box, orient="vertical",
                                    command=self.txt_in_master.yview)
        self.txt_in_master.configure(yscrollcommand=master_sb.set)
        master_sb.pack(side="right", fill="y")
        self.txt_in_master.pack(side="left", fill="both", expand=True)
        self.txt_in_master.bind("<KeyRelease>",
                                  lambda e: (self.count_total_qty(self.txt_in_master, self.lbl_in_m, "브랜드 수량"),
                                              self.txt_in_master.tag_remove("mismatch", "1.0", tk.END)))
        self.txt_in_master.bind("<Return>", self.check_double_enter)
        self.txt_in_master.tag_config("mismatch", background="#FECACA", foreground="#991B1B")

        # === 우: 스캔 수량 카드 ===
        r_card_outer = tk.Frame(cols_outer, bg="#F5F6F8")
        r_card_outer.grid(row=0, column=1, sticky="nsew", padx=(4, 0))

        r_card = tk.Frame(r_card_outer, bg="white",
                           highlightthickness=1, highlightbackground="#E5E7EB")
        r_card.pack(fill="both", expand=True)

        tk.Frame(r_card, bg="#10B981", width=4).pack(side="left", fill="y")

        r_inner = tk.Frame(r_card, bg="white", padx=12, pady=10)
        r_inner.pack(side="left", fill="both", expand=True)

        r_head = tk.Frame(r_inner, bg="white")
        r_head.pack(fill="x", pady=(0, 4))
        tk.Label(r_head, text="📡",
                 bg="white", font=("맑은 고딕", 11)).pack(side="left", padx=(0, 4))
        self.lbl_in_s = tk.Label(r_head, text="스캔 수량 (0개)",
                                   bg="white", fg="#065F46",
                                   font=("맑은 고딕", 10, "bold"))
        self.lbl_in_s.pack(side="left")
        # [추가] 우측 지우기 버튼
        clr_s = tk.Button(r_head, text="🗑️ 지우기",
                           command=lambda: self._clear_input_box(self.txt_in_scan, self.lbl_in_s, "스캔 수량"),
                           bg="#F3F4F6", fg="#6B7280",
                           activebackground="#E5E7EB",
                           font=("맑은 고딕", 8, "bold"),
                           relief="flat", padx=8, pady=2,
                           cursor="hand2")
        clr_s.pack(side="right")

        # [추가] 안내
        tk.Label(r_inner, text="💡 5열 입력시 정상/불량 구분 가능 (바코드 정상 불량 정상로케 불량로케)",
                 bg="white", fg="#9CA3AF",
                 font=("맑은 고딕", 8), anchor="w").pack(fill="x", pady=(0, 4))

        # 입력창 + 스크롤바 컨테이너
        scan_box = tk.Frame(r_inner, bg="white")
        scan_box.pack(fill="both", expand=True)

        self.txt_in_scan = tk.Text(scan_box, font=("Consolas", 11),
                                     bd=1, relief="solid",
                                     bg="#F0FDF4",
                                     highlightthickness=1, highlightbackground="#E5E7EB",
                                     padx=6, pady=6, wrap="char",
                                     height=15)  # 고정 높이 - 넘으면 스크롤로
        scan_sb = ttk.Scrollbar(scan_box, orient="vertical",
                                  command=self.txt_in_scan.yview)
        self.txt_in_scan.configure(yscrollcommand=scan_sb.set)
        scan_sb.pack(side="right", fill="y")
        self.txt_in_scan.pack(side="left", fill="both", expand=True)
        self.txt_in_scan.bind("<KeyRelease>",
                                lambda e: (self.count_total_qty(self.txt_in_scan, self.lbl_in_s, "스캔 수량"),
                                            self.txt_in_scan.tag_remove("mismatch", "1.0", tk.END)))
        self.txt_in_scan.tag_config("mismatch", background="#FECACA", foreground="#991B1B")

    # --- [탭 2: 출고] ---
    def setup_outbound(self):
        container = tk.Frame(self.t_out, bg="#F5F6F8")
        container.pack(fill="both", expand=True)

        # ========== [상단 헤더] ==========
        header_frame = tk.Frame(container, bg="#F5F6F8")
        header_frame.pack(side="top", fill="x", padx=24, pady=(16, 8))

        title_left = tk.Frame(header_frame, bg="#F5F6F8")
        title_left.pack(side="left")
        tk.Label(title_left, text="📤", font=("맑은 고딕", 22),
                 bg="#F5F6F8").pack(side="left", padx=(0, 6))
        title_text_box = tk.Frame(title_left, bg="#F5F6F8")
        title_text_box.pack(side="left")
        tk.Label(title_text_box, text="출고 등록",
                 font=("맑은 고딕", 15, "bold"),
                 bg="#F5F6F8", fg="#1A1A1A").pack(anchor="w")
        tk.Label(title_text_box, text="매장과 유형 선택 후 바코드 붙여넣기",
                 font=("맑은 고딕", 8),
                 bg="#F5F6F8", fg="#888").pack(anchor="w")

        # ========== [💾 저장 버튼 - 하단 고정 (먼저 pack)] ==========
        save_btn_outer = tk.Frame(container, bg="#F5F6F8")
        save_btn_outer.pack(side="bottom", fill="x", padx=18, pady=(0, 14))

        # 그림자 효과 흉내
        shadow = tk.Frame(save_btn_outer, bg="#1864c8", height=44)
        shadow.pack(fill="x")

        save_btn = tk.Button(shadow, text="🚚  출고 CSV 저장 및 데이터 리셋",
                              bg="#1877F2", fg="white", activebackground="#1864c8",
                              activeforeground="white",
                              font=("맑은 고딕", 11, "bold"),
                              relief="flat", bd=0,
                              cursor="hand2",
                              command=self.run_out)
        save_btn.pack(fill="x", padx=0, pady=0, ipady=10)
        # 호버 효과
        def on_save_enter(e): save_btn.config(bg="#1864c8")
        def on_save_leave(e): save_btn.config(bg="#1877F2")
        save_btn.bind("<Enter>", on_save_enter)
        save_btn.bind("<Leave>", on_save_leave)

        # ========== [선택 카드 영역 - 매장 + 유형 한 카드에] ==========
        select_card_outer = tk.Frame(container, bg="#F5F6F8")
        select_card_outer.pack(fill="x", padx=18, pady=(0, 8))

        select_card = tk.Frame(select_card_outer, bg="white",
                                highlightthickness=1, highlightbackground="#E5E7EB")
        select_card.pack(fill="x")

        # 좌측 액센트
        tk.Frame(select_card, bg="#3B82F6", width=4).pack(side="left", fill="y")

        select_inner = tk.Frame(select_card, bg="white", padx=16, pady=12)
        select_inner.pack(side="left", fill="both", expand=True)

        # === 매장 선택 ===
        store_row = tk.Frame(select_inner, bg="white")
        store_row.pack(fill="x", pady=(0, 8))
        tk.Label(store_row, text="🏬 매장",
                 bg="white", fg="#374151",
                 font=("맑은 고딕", 9, "bold"), width=8, anchor="w").pack(side="left", padx=(0, 8))

        store_btns_frame = tk.Frame(store_row, bg="white")
        store_btns_frame.pack(side="left", fill="x", expand=True)

        stores = ["성수", "압구정", "갤러리아", "인하우스", "마케팅", "브랜드반납"]
        for idx, s in enumerate(stores):
            btn = tk.Button(store_btns_frame, text=s,
                             command=lambda x=s: self.select_opt('s', x),
                             bg="#F0F2F5", fg="#666",
                             font=("맑은 고딕", 9, "bold"),
                             relief="flat", padx=14, pady=6,
                             cursor="hand2")
            btn.pack(side="left", padx=2)
            self.s_btns[s] = btn

        # === 유형 선택 ===
        type_row = tk.Frame(select_inner, bg="white")
        type_row.pack(fill="x")
        tk.Label(type_row, text="📝 유형",
                 bg="white", fg="#374151",
                 font=("맑은 고딕", 9, "bold"), width=8, anchor="w").pack(side="left", padx=(0, 8))

        type_btns_frame = tk.Frame(type_row, bg="white")
        type_btns_frame.pack(side="left", fill="x", expand=True)

        types = ["매장출고", "신규출고", "긴급(픽업)출고", "보충출고", "택배출고", "촬영출고", "퀵출고"]
        for idx, t in enumerate(types):
            btn = tk.Button(type_btns_frame, text=t,
                             command=lambda x=t: self.select_opt('t', x),
                             bg="#F0F2F5", fg="#666",
                             font=("맑은 고딕", 9, "bold"),
                             relief="flat", padx=10, pady=6,
                             cursor="hand2")
            btn.pack(side="left", padx=2)
            self.t_btns[t] = btn

        # 기본값 설정
        self.select_opt('s', '성수')
        self.select_opt('t', '매장출고')

        # ========== [입력 카드 - 남은 공간 채움] ==========
        input_card_outer = tk.Frame(container, bg="#F5F6F8")
        input_card_outer.pack(fill="both", expand=True, padx=18, pady=(0, 8))

        input_card = tk.Frame(input_card_outer, bg="white",
                               highlightthickness=1, highlightbackground="#E5E7EB")
        input_card.pack(fill="both", expand=True)

        # 좌측 액센트
        tk.Frame(input_card, bg="#10B981", width=4).pack(side="left", fill="y")

        input_inner = tk.Frame(input_card, bg="white", padx=16, pady=12)
        input_inner.pack(side="left", fill="both", expand=True)

        # 헤더
        head = tk.Frame(input_inner, bg="white")
        head.pack(fill="x", pady=(0, 8))
        tk.Label(head, text="📡",
                 bg="white", font=("맑은 고딕", 12)).pack(side="left", padx=(0, 6))

        self.lbl_out_qty = tk.Label(head, text="출고 바코드 & 수량 붙여넣기 (0개)",
                                     font=("맑은 고딕", 10, "bold"),
                                     bg="white", fg="#111827")
        self.lbl_out_qty.pack(side="left")

        # [추가] 엑셀 불러오기 버튼 - 모던 스타일
        excel_btn = tk.Button(head, text="📁  엑셀 불러오기",
                               command=self.load_excel_to_outbound,
                               bg="#0EA5E9", fg="white",
                               activebackground="#0284C7", activeforeground="white",
                               font=("맑은 고딕", 9, "bold"),
                               relief="flat", padx=12, pady=5,
                               cursor="hand2")
        excel_btn.pack(side="right")
        def on_excel_enter(e): excel_btn.config(bg="#0284C7")
        def on_excel_leave(e): excel_btn.config(bg="#0EA5E9")
        excel_btn.bind("<Enter>", on_excel_enter)
        excel_btn.bind("<Leave>", on_excel_leave)

        # [추가] 지우기 버튼
        clr_out = tk.Button(head, text="🗑️ 지우기",
                             command=lambda: self._clear_input_box(self.txt_out, self.lbl_out_qty, "📡 출고 바코드 & 수량 붙여넣기"),
                             bg="#F3F4F6", fg="#6B7280",
                             activebackground="#E5E7EB",
                             font=("맑은 고딕", 8, "bold"),
                             relief="flat", padx=8, pady=2,
                             cursor="hand2")
        clr_out.pack(side="right", padx=(0, 6))

        # [추가] 안내 문구
        info_label = tk.Label(input_inner,
                               text="💡 복붙 / 📁 버튼으로 엑셀 불러오기 / 엑셀 파일을 입력창에 드래그",
                               bg="white", fg="#9CA3AF",
                               font=("맑은 고딕", 8), anchor="w")
        info_label.pack(fill="x", pady=(0, 4))

        # 입력창 + 스크롤바 컨테이너
        out_box = tk.Frame(input_inner, bg="white")
        out_box.pack(fill="both", expand=True)

        self.txt_out = tk.Text(out_box, font=("Consolas", 10),
                                bd=1, relief="solid",
                                highlightthickness=1, highlightbackground="#E5E7EB",
                                bg="#FAFAFA", padx=8, pady=6, wrap="char",
                                height=18)  # 고정 높이
        out_sb = ttk.Scrollbar(out_box, orient="vertical",
                                 command=self.txt_out.yview)
        self.txt_out.configure(yscrollcommand=out_sb.set)
        out_sb.pack(side="right", fill="y")
        self.txt_out.pack(side="left", fill="both", expand=True)
        self.txt_out.bind("<KeyRelease>",
                           lambda e: self.count_total_qty(self.txt_out, self.lbl_out_qty,
                                                            "📡 출고 바코드 & 수량 붙여넣기"))

        # [추가] 드래그앤드롭 등록 (tkinterdnd2 있을 때만)
        self._setup_dnd_for_outbound()

    def select_opt(self, mode, val):
        if mode == 's':
            self.selected_store.set(val)
            for k, btn in self.s_btns.items():
                if k == val:
                    btn.config(bg="#1877F2", fg="white")
                else:
                    btn.config(bg="#F0F2F5", fg="#666")
        else:
            self.selected_type.set(val)
            for k, btn in self.t_btns.items():
                if k == val:
                    btn.config(bg="#8B5CF6", fg="white")
                else:
                    btn.config(bg="#F0F2F5", fg="#666")

    # ========== [엑셀 → 출고창 자동 변환] ==========
    def _setup_dnd_for_outbound(self):
        """출고 입력창에 드래그앤드롭 등록 (tkinterdnd2 있을 때만)"""
        try:
            from tkinterdnd2 import DND_FILES
            self.txt_out.drop_target_register(DND_FILES)
            self.txt_out.dnd_bind('<<Drop>>', self._on_excel_dropped)
            print("✅ 드래그앤드롭 활성화됨")
        except ImportError:
            print("ℹ️ tkinterdnd2 미설치 - 드래그앤드롭 비활성화 (파일 선택 버튼 사용)")
        except Exception as e:
            print(f"⚠️ 드래그앤드롭 설정 실패: {e}")

    def _on_excel_dropped(self, event):
        """엑셀 파일이 입력창에 드래그됐을 때"""
        # event.data 는 보통 "{C:/path/to/file.xlsx}" 또는 그냥 경로
        path = event.data.strip()
        if path.startswith('{') and path.endswith('}'):
            path = path[1:-1]
        # 여러 파일이면 첫 번째만
        if ' ' in path and not os.path.exists(path):
            for p in path.split():
                p = p.strip('{}')
                if os.path.exists(p):
                    path = p
                    break
        if not os.path.exists(path):
            messagebox.showerror("오류", f"파일을 찾을 수 없습니다:\n{path}")
            return
        self._process_excel_file(path)
        return event.action

    def load_excel_to_outbound(self):
        """파일 선택 다이얼로그로 엑셀 불러오기"""
        path = filedialog.askopenfilename(
            title="출고용 엑셀 파일 선택",
            filetypes=[("Excel", "*.xlsx *.xls"), ("CSV", "*.csv"), ("모든 파일", "*.*")]
        )
        if not path:
            return
        self._process_excel_file(path)

    def _process_excel_file(self, path):
        """엑셀 파일을 읽어 바코드+수량 추출 후 입력창에 채움"""
        try:
            df = self._smart_parse_outbound_excel(path)
            if df is None or df.empty:
                messagebox.showwarning("결과 없음",
                    "파일에서 바코드와 수량을 인식하지 못했습니다.\n"
                    "헤더에 '바코드'와 '수량'(또는 유사 표현)이 있는지 확인해주세요.")
                return

            # 수량 0인 행 제거
            df = df[df['수량'] > 0].copy()
            if df.empty:
                messagebox.showinfo("결과 없음", "수량이 0보다 큰 행이 없습니다.")
                return

            # 텍스트 형식으로 변환: 바코드\t수량\n...
            new_lines = "\n".join(f"{row['바코드']}\t{int(row['수량'])}"
                                    for _, row in df.iterrows())

            # 기존 내용 확인
            existing = self.txt_out.get("1.0", "end-1c").strip()
            if existing:
                # 선택 다이얼로그
                choice = self._ask_excel_overwrite_choice(len(df), len(existing.split('\n')))
                if choice is None:
                    return
                if choice == 'overwrite':
                    self.txt_out.delete("1.0", tk.END)
                    self.txt_out.insert("1.0", new_lines)
                elif choice == 'append':
                    self.txt_out.insert(tk.END, "\n" + new_lines)
            else:
                self.txt_out.insert("1.0", new_lines)

            # 카운터 업데이트
            self.count_total_qty(self.txt_out, self.lbl_out_qty,
                                  "📡 출고 바코드 & 수량 붙여넣기")
            messagebox.showinfo("불러오기 완료",
                f"✅ {len(df)}건의 바코드+수량이 입력되었습니다.\n"
                f"(수량 0인 행은 자동 제외)")
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("오류", f"엑셀 불러오기 실패:\n{e}")

    def _ask_excel_overwrite_choice(self, new_count, existing_count):
        """기존 내용 있을 때 처리 방식 묻기"""
        win = tk.Toplevel(self.root)
        win.title("입력창에 이미 내용이 있습니다")
        win.configure(bg="white")
        win.transient(self.root)
        win.grab_set()
        try:
            self.position_popup(win, 380, 220)
        except Exception:
            win.geometry("380x220")

        result = {"choice": None}

        tk.Label(win, text="⚠️ 입력창에 이미 데이터가 있습니다",
                 bg="white", fg="#1A1A1A",
                 font=("맑은 고딕", 11, "bold")).pack(pady=(20, 5))
        tk.Label(win, text=f"기존 {existing_count}줄  +  새로 {new_count}줄",
                 bg="white", fg="#666",
                 font=("맑은 고딕", 9)).pack(pady=(0, 15))

        btn_frame = tk.Frame(win, bg="white")
        btn_frame.pack(fill="x", padx=20)

        def pick(choice):
            result["choice"] = choice
            win.destroy()

        tk.Button(btn_frame, text="🔄 덮어쓰기\n(기존 삭제 후 새로)",
                  command=lambda: pick("overwrite"),
                  bg="#FEF2F2", fg="#991B1B",
                  font=("맑은 고딕", 9, "bold"),
                  relief="flat", padx=10, pady=8,
                  cursor="hand2").pack(side="left", expand=True, fill="x", padx=2)
        tk.Button(btn_frame, text="➕ 추가\n(기존 뒤에 붙임)",
                  command=lambda: pick("append"),
                  bg="#ECFDF5", fg="#065F46",
                  font=("맑은 고딕", 9, "bold"),
                  relief="flat", padx=10, pady=8,
                  cursor="hand2").pack(side="left", expand=True, fill="x", padx=2)
        tk.Button(btn_frame, text="❌ 취소",
                  command=lambda: pick(None),
                  bg="#F3F4F6", fg="#374151",
                  font=("맑은 고딕", 9, "bold"),
                  relief="flat", padx=10, pady=8,
                  cursor="hand2").pack(side="left", expand=True, fill="x", padx=2)

        win.wait_window()
        return result["choice"]

    def _smart_parse_outbound_excel(self, path):
        """엑셀 파일에서 바코드+수량 자동 인식.
        - 헤더가 1행이 아닐 수도 있음 (위 5줄까지 헤더 후보 탐색)
        - 컬럼명: 바코드/sku/SKU/상품바코드/코드/Barcode 등
        - 수량 컬럼명: 수량/qty/Qty/QTY/Q'TY/재고/가용재고/출고수량/배송수량 등
        - CSV도 지원"""
        ext = os.path.splitext(path)[1].lower()
        if ext == '.csv':
            df_raw = pd.read_csv(path, header=None, dtype=str, encoding='utf-8-sig', errors='ignore') \
                if False else pd.read_csv(path, header=None, dtype=str, encoding='utf-8-sig')
        else:
            engine = 'xlrd' if ext == '.xls' else 'openpyxl'
            try:
                df_raw = pd.read_excel(path, header=None, dtype=str, engine=engine)
            except Exception:
                df_raw = pd.read_excel(path, header=None, dtype=str)

        # 바코드/수량 키워드 (소문자 비교)
        barcode_keywords = ['바코드', 'sku', '상품바코드', 'barcode', '품번', '제품코드']
        qty_keywords = ['수량', 'qty', "q'ty", 'quantity', '재고', '가용재고',
                        '출고수량', '배송수량', '출고', '주문수량', '발주수량']
        # 우선순위: 정확히 '수량'이 가장 우선, 그 다음 출고/주문, 그 다음 가용재고
        # 우선순위 점수 (낮을수록 우선)
        qty_priority = {
            '수량': 1, 'qty': 1, "q'ty": 1, 'quantity': 1,
            '출고수량': 2, '출고': 2, '주문수량': 2, '발주수량': 2,
            '배송수량': 3,
            '가용재고': 4, '재고': 5
        }

        # 헤더 후보 행 찾기 (위 10행까지 시도)
        header_row = -1
        for i in range(min(10, len(df_raw))):
            row = df_raw.iloc[i].fillna('').astype(str).str.strip().str.lower()
            row_text = ' '.join(row.tolist())
            has_barcode = any(k in row_text for k in barcode_keywords)
            has_qty = any(k in row_text for k in qty_keywords)
            if has_barcode and has_qty:
                header_row = i
                break

        if header_row == -1:
            return None

        # 진짜 헤더로 다시 읽기
        ext = os.path.splitext(path)[1].lower()
        if ext == '.csv':
            df = pd.read_csv(path, header=header_row, dtype=str, encoding='utf-8-sig')
        else:
            engine = 'xlrd' if ext == '.xls' else 'openpyxl'
            try:
                df = pd.read_excel(path, header=header_row, dtype=str, engine=engine)
            except Exception:
                df = pd.read_excel(path, header=header_row, dtype=str)

        df.columns = [str(c).strip() for c in df.columns]

        # 바코드 컬럼 찾기
        bc_col = None
        for col in df.columns:
            col_l = col.lower()
            if any(k in col_l for k in barcode_keywords):
                bc_col = col
                break
        if bc_col is None:
            return None

        # 수량 컬럼 찾기 (우선순위 적용)
        qty_col = None
        best_priority = 999
        for col in df.columns:
            col_l = col.lower()
            for k, prio in qty_priority.items():
                if k in col_l and prio < best_priority:
                    qty_col = col
                    best_priority = prio
                    break
        if qty_col is None:
            # 일반 키워드라도 매치
            for col in df.columns:
                col_l = col.lower()
                if any(k in col_l for k in qty_keywords):
                    qty_col = col
                    break
        if qty_col is None:
            return None

        # 결과 정리
        result = pd.DataFrame()
        result['바코드'] = df[bc_col].fillna('').astype(str).str.strip()
        result['수량'] = pd.to_numeric(df[qty_col], errors='coerce').fillna(0).astype(int)
        # 빈 바코드 제거
        result = result[result['바코드'] != ''].copy()
        # 같은 바코드가 여러 줄에 있으면 합치기 (수량 더하기)
        result = result.groupby('바코드', as_index=False)['수량'].sum()

        return result
    # --- [탭 3: 맘스] ---
    def setup_moms_v86(self):
        """맘스 입/출고 등록 - 모던 카드 스타일"""
        container = tk.Frame(self.t_mom, bg="#F5F6F8")
        container.pack(fill="both", expand=True)

        # ========== [상단 헤더] ==========
        header_frame = tk.Frame(container, bg="#F5F6F8")
        header_frame.pack(side="top", fill="x", padx=24, pady=(16, 8))

        title_left = tk.Frame(header_frame, bg="#F5F6F8")
        title_left.pack(side="left")
        tk.Label(title_left, text="📦", font=("맑은 고딕", 22),
                 bg="#F5F6F8").pack(side="left", padx=(0, 6))
        title_text_box = tk.Frame(title_left, bg="#F5F6F8")
        title_text_box.pack(side="left")
        tk.Label(title_text_box, text="맘스 입/출고 등록",
                 font=("맑은 고딕", 15, "bold"),
                 bg="#F5F6F8", fg="#1A1A1A").pack(anchor="w")
        tk.Label(title_text_box, text="성수↔여주 재고 이동 및 출고",
                 font=("맑은 고딕", 8),
                 bg="#F5F6F8", fg="#888").pack(anchor="w")

        # ========== [카드 컨테이너 - 2열 그리드] ==========
        cards_frame = tk.Frame(container, bg="#F5F6F8")
        cards_frame.pack(fill="both", expand=True, padx=18, pady=(0, 12))
        cards_frame.columnconfigure(0, weight=1, uniform="col")
        cards_frame.columnconfigure(1, weight=1, uniform="col")
        cards_frame.rowconfigure(0, weight=1, uniform="row")
        cards_frame.rowconfigure(1, weight=1, uniform="row")

        # 카드 그리는 공통 함수
        def make_card(parent, row, col, accent_color, badge_text, badge_bg, badge_fg, title, desc):
            """카드 컨테이너 생성. 내부에 위젯 채울 frame 반환."""
            card_outer = tk.Frame(parent, bg="#F5F6F8")
            card_outer.grid(row=row, column=col, sticky="nsew", padx=5, pady=5)

            card = tk.Frame(card_outer, bg="white",
                             highlightthickness=1, highlightbackground="#E5E7EB")
            card.pack(fill="both", expand=True)

            # 좌측 컬러 액센트 바
            tk.Frame(card, bg=accent_color, width=4).pack(side="left", fill="y")

            # 본 컨텐츠
            content = tk.Frame(card, bg="white", padx=14, pady=12)
            content.pack(side="left", fill="both", expand=True)

            # 헤더 (뱃지 + 제목)
            head = tk.Frame(content, bg="white")
            head.pack(fill="x")
            tk.Label(head, text=badge_text,
                     bg=badge_bg, fg=badge_fg,
                     font=("맑은 고딕", 8, "bold"),
                     padx=8, pady=2).pack(side="left")

            tk.Label(content, text=title,
                     bg="white", fg="#111827",
                     font=("맑은 고딕", 11, "bold"),
                     anchor="w").pack(fill="x", anchor="w", pady=(8, 2))
            tk.Label(content, text=desc,
                     bg="white", fg="#6B7280",
                     font=("맑은 고딕", 8),
                     anchor="w", justify="left",
                     wraplength=320).pack(fill="x", anchor="w", pady=(0, 10))

            return content

        # 파일 선택 행 그리는 공통 함수
        def make_file_row(parent, btn_text, btn_color, label_widget_attr, command, default_text="미선택"):
            row = tk.Frame(parent, bg="white")
            row.pack(fill="x", pady=2)
            btn = tk.Button(row, text=btn_text, command=command,
                             bg=btn_color, fg="#374151",
                             font=("맑은 고딕", 8, "bold"),
                             relief="flat", padx=10, pady=5,
                             cursor="hand2")
            btn.pack(side="left")
            lbl = tk.Label(row, text=default_text, fg="#9CA3AF", bg="white",
                            font=("맑은 고딕", 8))
            lbl.pack(side="left", padx=8, fill="x", expand=True)
            setattr(self, label_widget_attr, lbl)
            return lbl

        # 액션 버튼 그리는 공통 함수 (하단 고정)
        def make_action_btn(parent, text, bg, command):
            btn = tk.Button(parent, text=text, command=command,
                              bg=bg, fg="white",
                              font=("맑은 고딕", 9, "bold"),
                              relief="flat", pady=9,
                              cursor="hand2")
            btn.pack(fill="x", side="bottom", pady=(8, 0))
            return btn

        # ========== [카드 1: 신규 마스터 생성] ==========
        c1 = make_card(cards_frame, 0, 0,
                        accent_color="#3B82F6",
                        badge_text="1️⃣ 마스터",
                        badge_bg="#DBEAFE", badge_fg="#1E40AF",
                        title="신규 마스터 리스트 생성",
                        desc="성수 출고 ↔ 여주 마스터 비교 후 신규 상품만 추출")
        make_file_row(c1, "📁 출고리스트 (성수)", "#EFF6FF",
                      "lbl_master_send", self.sel_master_send)
        make_file_row(c1, "📁 마스터재고 (여주)", "#EFF6FF",
                      "lbl_master_master", self.sel_master_master)
        make_action_btn(c1, "✨ 신규 마스터 생성", "#3B82F6", self.run_mom_master_logic)

        # ========== [카드 2: 입고리스트] ==========
        c2 = make_card(cards_frame, 0, 1,
                        accent_color="#10B981",
                        badge_text="2️⃣ 입고",
                        badge_bg="#D1FAE5", badge_fg="#065F46",
                        title="입고 리스트 생성",
                        desc="성수 → 여주 전체 재고 이동용 입고 파일 생성")
        make_file_row(c2, "📁 출고리스트 (성수)", "#ECFDF5",
                      "lbl_inbound_send", self.sel_inbound_send)
        make_action_btn(c2, "📋 입고 리스트 생성", "#10B981", self.run_mom_inbound_logic)

        # ========== [카드 3: 제외 적용 입고] ==========
        c3 = make_card(cards_frame, 1, 0,
                        accent_color="#F59E0B",
                        badge_text="3️⃣ 제외 입고",
                        badge_bg="#FEF3C7", badge_fg="#92400E",
                        title="특정상품 제외 입고",
                        desc="제외 명단에 있는 상품 빼고 입고 (바코드/아이템코드 매치)")
        make_file_row(c3, "📁 출고리스트 (성수)", "#FFFBEB",
                      "lbl_excl_send", self.sel_excl_send)
        make_file_row(c3, "📁 제외 리스트", "#FFFBEB",
                      "lbl_excl_list", self.sel_excl_list,
                      default_text="미선택 (바코드/아이템코드)")
        make_action_btn(c3, "🚫 제외 적용 입고 생성", "#F59E0B", self.run_mom_inbound_new_only)

        # ========== [카드 4: 출고등록] ==========
        c4 = make_card(cards_frame, 1, 1,
                        accent_color="#8B5CF6",
                        badge_text="4️⃣ 출고",
                        badge_bg="#EDE9FE", badge_fg="#5B21B6",
                        title="맘스 출고등록",
                        desc="주문자 입력 후 출고 리스트 파일로 출고 등록")
        # 주문자 입력
        user_row = tk.Frame(c4, bg="white")
        user_row.pack(fill="x", pady=2)
        tk.Label(user_row, text="👤", bg="white",
                 font=("맑은 고딕", 10)).pack(side="left", padx=(0, 4))
        self.ent_mom_user = tk.Entry(user_row, bd=1, relief="solid",
                                       font=("맑은 고딕", 10),
                                       highlightthickness=0)
        self.ent_mom_user.pack(side="left", fill="x", expand=True, ipady=3)
        # 파일 선택
        make_file_row(c4, "📁 출고 리스트", "#F5F3FF",
                      "lbl_mom_out", self.sel_mom_out, default_text="파일 미선택")
        make_action_btn(c4, "📋 출고 리스트 생성", "#8B5CF6", self.run_mom_out_logic)


    # ========================================================================
    # [탭: 마스터 등록] - 바코드표 → 마스터 등록 + 옵션 등록 양식 자동 변환
    # ========================================================================
    # 아이템 코드 매핑 (품번 9번째 글자 → 아이템코드)
    ITEM_CODE_MAP = {
        'R': 'RTW', 'S': 'SHOES', 'A': 'ACC',
        'B': 'BAG', 'E': 'ETC', 'K': 'RAF'
    }

    def setup_master_registration(self):
        """마스터 등록 + 색상코드 + 상품관리 수정 + 상품옵션 수정 (서브탭 구조)"""
        container = tk.Frame(self.t_master, bg="#F5F6F8")
        container.pack(fill="both", expand=True)

        # ========== [서브탭 노트북] ==========
        sub_nb = ttk.Notebook(container, style="SubTab.TNotebook")
        sub_nb.pack(fill="both", expand=True, padx=8, pady=(8, 0))

        # 서브탭 스타일 (메인 탭과 구분되게 더 컴팩트)
        self.style.configure("SubTab.TNotebook",
                              background="#F5F6F8",
                              borderwidth=0,
                              tabmargins=[0, 4, 0, 0])
        self.style.configure("SubTab.TNotebook.Tab",
                              padding=[14, 6],
                              font=("맑은 고딕", 9),
                              background="#F5F6F8",
                              foreground="#6B7280",
                              borderwidth=0,
                              focuscolor="#F5F6F8")
        self.style.map("SubTab.TNotebook.Tab",
                        background=[("selected", "white"),
                                    ("active", "#E5E7EB")],
                        foreground=[("selected", "#1877F2"),
                                    ("active", "#374151")])

        # 2개 서브 탭
        self.t_master_main = ttk.Frame(sub_nb)
        self.t_master_kit = ttk.Frame(sub_nb)

        sub_nb.add(self.t_master_main, text="🆕 마스터 등록")
        sub_nb.add(self.t_master_kit, text="🔧 마스터 수정 키트")

        # 서브탭 1: 마스터 등록
        self._setup_master_main_tab(self.t_master_main)
        # 서브탭 2: 수정 키트 (색상/상품관리/상품옵션 3종)
        self._setup_master_kit_tab(self.t_master_kit)

    def _setup_master_placeholder(self, parent, icon, title, desc):
        """미구현 서브탭 - 안내 화면"""
        wrap = tk.Frame(parent, bg="#F5F6F8")
        wrap.pack(fill="both", expand=True)

        tk.Label(wrap, text=icon, font=("맑은 고딕", 48),
                 bg="#F5F6F8").pack(pady=(80, 12))
        tk.Label(wrap, text=title,
                 bg="#F5F6F8", fg="#1A1A1A",
                 font=("맑은 고딕", 16, "bold")).pack()
        tk.Label(wrap, text="🚧 준비 중",
                 bg="#FEF3C7", fg="#92400E",
                 font=("맑은 고딕", 9, "bold"),
                 padx=12, pady=4).pack(pady=(8, 14))
        tk.Label(wrap, text=desc,
                 bg="#F5F6F8", fg="#6B7280",
                 font=("맑은 고딕", 10),
                 justify="center").pack()

    def _setup_master_main_tab(self, container):
        """마스터 등록 메인 화면 - 바코드표 붙여넣기 → 변환"""
        # ========== [상단 헤더] ==========
        header_frame = tk.Frame(container, bg="#F5F6F8")
        header_frame.pack(side="top", fill="x", padx=18, pady=(12, 6))

        title_left = tk.Frame(header_frame, bg="#F5F6F8")
        title_left.pack(side="left")
        tk.Label(title_left, text="🆕", font=("맑은 고딕", 18),
                 bg="#F5F6F8").pack(side="left", padx=(0, 6))
        title_text_box = tk.Frame(title_left, bg="#F5F6F8")
        title_text_box.pack(side="left")
        tk.Label(title_text_box, text="마스터 등록",
                 font=("맑은 고딕", 13, "bold"),
                 bg="#F5F6F8", fg="#1A1A1A").pack(anchor="w")
        tk.Label(title_text_box, text="바코드표 붙여넣기 → 마스터/옵션 양식 자동 변환",
                 font=("맑은 고딕", 8),
                 bg="#F5F6F8", fg="#888").pack(anchor="w")

        # 우측: 브랜드 관리 버튼
        tk.Button(header_frame, text="🗂️ 브랜드 관리",
                   command=self.open_brand_manager,
                   bg="white", fg="#444",
                   font=("맑은 고딕", 9, "bold"),
                   relief="flat", padx=14, pady=8,
                   cursor="hand2",
                   highlightthickness=1, highlightbackground="#DDD").pack(side="right")

        # ========== [입력 카드] ==========
        input_card_outer = tk.Frame(container, bg="#F5F6F8")
        input_card_outer.pack(fill="both", expand=True, padx=18, pady=(0, 8))

        input_card = tk.Frame(input_card_outer, bg="white",
                               highlightthickness=1, highlightbackground="#E5E7EB")
        input_card.pack(fill="both", expand=True)

        tk.Frame(input_card, bg="#3B82F6", width=4).pack(side="left", fill="y")

        input_inner = tk.Frame(input_card, bg="white", padx=16, pady=12)
        input_inner.pack(side="left", fill="both", expand=True)

        head = tk.Frame(input_inner, bg="white")
        head.pack(fill="x", pady=(0, 4))
        tk.Label(head, text="📥",
                 bg="white", font=("맑은 고딕", 12)).pack(side="left", padx=(0, 6))
        tk.Label(head, text="바코드표 붙여넣기",
                 font=("맑은 고딕", 10, "bold"),
                 bg="white", fg="#111827").pack(side="left")

        clr_btn = tk.Button(head, text="🗑️ 지우기",
                             command=self._clear_master_input,
                             bg="#F3F4F6", fg="#6B7280",
                             font=("맑은 고딕", 8, "bold"),
                             relief="flat", padx=8, pady=2,
                             cursor="hand2")
        clr_btn.pack(side="right")

        tk.Label(input_inner,
                 text="💡 바코드표에서 [상품명, 바코드번호, 상품메모1, 상품메모2, 대표판매가, 옵션내용, 수량] 컬럼 그대로 복붙",
                 bg="white", fg="#9CA3AF",
                 font=("맑은 고딕", 8), anchor="w").pack(fill="x", pady=(0, 4))

        # 입력창 + 스크롤바 컨테이너
        master_in_box = tk.Frame(input_inner, bg="white")
        master_in_box.pack(fill="both", expand=True)

        self.txt_master_in = tk.Text(master_in_box, font=("Consolas", 10),
                                       bd=1, relief="solid",
                                       highlightthickness=1, highlightbackground="#E5E7EB",
                                       bg="#FAFAFA", padx=8, pady=6, wrap="char",
                                       height=18)  # 고정 높이
        master_in_sb = ttk.Scrollbar(master_in_box, orient="vertical",
                                       command=self.txt_master_in.yview)
        self.txt_master_in.configure(yscrollcommand=master_in_sb.set)
        master_in_sb.pack(side="right", fill="y")
        self.txt_master_in.pack(side="left", fill="both", expand=True)

        # ========== [액션 버튼 - 하단 3개 (먼저 pack해서 항상 보이게)] ==========
        action_outer = tk.Frame(container, bg="#F5F6F8")
        action_outer.pack(side="bottom", fill="x", padx=18, pady=(0, 14))

        def make_modern_btn(parent, text, bg, hover_bg, command):
            shadow = tk.Frame(parent, bg=hover_bg)
            shadow.pack(side="left", expand=True, fill="x", padx=2)
            btn = tk.Button(shadow, text=text,
                              bg=bg, fg="white",
                              activebackground=hover_bg, activeforeground="white",
                              font=("맑은 고딕", 10, "bold"),
                              relief="flat", bd=0,
                              cursor="hand2", command=command)
            btn.pack(fill="x", ipady=8)
            def on_enter(e): btn.config(bg=hover_bg)
            def on_leave(e): btn.config(bg=bg)
            btn.bind("<Enter>", on_enter)
            btn.bind("<Leave>", on_leave)
            return btn

        make_modern_btn(action_outer, "🔄  변환 실행",
                         bg="#F59E0B", hover_bg="#D97706",
                         command=self.run_master_conversion)
        make_modern_btn(action_outer, "📎  마스터 등록 복사",
                         bg="#3B82F6", hover_bg="#1D4ED8",
                         command=self.copy_master_table_to_clipboard)
        make_modern_btn(action_outer, "📎  옵션 등록 복사",
                         bg="#10B981", hover_bg="#059669",
                         command=self.copy_option_table_to_clipboard)

        # ========== [결과 리포트 카드 - 하단 (버튼 위)] ==========
        preview_card_outer = tk.Frame(container, bg="#F5F6F8")
        preview_card_outer.pack(side="bottom", fill="x", padx=18, pady=(0, 8))

        preview_card = tk.Frame(preview_card_outer, bg="white",
                                  highlightthickness=1, highlightbackground="#E5E7EB")
        preview_card.pack(fill="x")

        tk.Frame(preview_card, bg="#8B5CF6", width=4).pack(side="left", fill="y")

        preview_inner = tk.Frame(preview_card, bg="white", padx=14, pady=10)
        preview_inner.pack(side="left", fill="both", expand=True)

        tk.Label(preview_inner, text="📝  변환 결과 / 매칭 리포트",
                 bg="white", fg="#111827",
                 font=("맑은 고딕", 10, "bold")).pack(anchor="w", pady=(0, 4))

        self.txt_master_report = tk.Text(preview_inner, height=5,
                                           font=("Consolas", 10),
                                           bg="#FAFAFA", bd=1, relief="solid",
                                           highlightthickness=1, highlightbackground="#E5E7EB",
                                           padx=8, pady=6)
        self.txt_master_report.pack(fill="x")
        self.txt_master_report.tag_config("ok", foreground="#16A34A", font=("Consolas", 10, "bold"))
        self.txt_master_report.tag_config("warn", foreground="#DC2626", font=("Consolas", 10, "bold"))
        self.txt_master_report.tag_config("info", foreground="#555", font=("Consolas", 10))

        # 마지막 결과 보관용
        self._last_master_table_df = None  # 마스터 등록용
        self._last_option_table_df = None  # 옵션 등록용

        # 브랜드 캐시 (Firestore에서 받아옴) - 실시간 리스너로 자동 동기화
        self._brand_cache = {}
        self.refresh_brand_cache()
        self.start_brand_listener()  # [추가] 다른 PC에서 변경 시 즉시 반영

    # ========== [마스터 수정 키트] ==========
    def _setup_master_kit_tab(self, container):
        """3종 수정 양식: 색상코드 / 상품관리 / 상품옵션 - 라디오 버튼으로 전환"""
        wrap = tk.Frame(container, bg="#F5F6F8")
        wrap.pack(fill="both", expand=True)

        # ===== 상단 헤더 + 모드 선택 =====
        header = tk.Frame(wrap, bg="#F5F6F8")
        header.pack(side="top", fill="x", padx=18, pady=(12, 6))

        title_left = tk.Frame(header, bg="#F5F6F8")
        title_left.pack(side="left")
        tk.Label(title_left, text="🔧", font=("맑은 고딕", 18),
                 bg="#F5F6F8").pack(side="left", padx=(0, 6))
        title_box = tk.Frame(title_left, bg="#F5F6F8")
        title_box.pack(side="left")
        tk.Label(title_box, text="마스터 수정 키트",
                 font=("맑은 고딕", 13, "bold"),
                 bg="#F5F6F8", fg="#1A1A1A").pack(anchor="w")
        tk.Label(title_box, text="색상코드 / 상품관리 / 상품옵션 양식 변환",
                 font=("맑은 고딕", 8),
                 bg="#F5F6F8", fg="#888").pack(anchor="w")

        # ===== 모드 선택 카드 =====
        mode_card_outer = tk.Frame(wrap, bg="#F5F6F8")
        mode_card_outer.pack(fill="x", padx=18, pady=(0, 8))

        mode_card = tk.Frame(mode_card_outer, bg="white",
                               highlightthickness=1, highlightbackground="#E5E7EB")
        mode_card.pack(fill="x")
        tk.Frame(mode_card, bg="#3B82F6", width=4).pack(side="left", fill="y")

        mode_inner = tk.Frame(mode_card, bg="white", padx=14, pady=10)
        mode_inner.pack(side="left", fill="both", expand=True)

        tk.Label(mode_inner, text="📋 작업 선택",
                 bg="white", fg="#374151",
                 font=("맑은 고딕", 9, "bold")).pack(anchor="w", pady=(0, 6))

        self._kit_mode = tk.StringVar(value="color")
        self._kit_mode_btns = {}

        btn_row = tk.Frame(mode_inner, bg="white")
        btn_row.pack(fill="x")

        modes = [
            ("color", "🎨 색상코드 등록"),
            ("product", "✏️ 상품 관리 수정"),
            ("option", "🔧 상품 옵션 수정"),
        ]

        def select_mode(mode):
            self._kit_mode.set(mode)
            for k, btn in self._kit_mode_btns.items():
                if k == mode:
                    btn.config(bg="#1877F2", fg="white")
                else:
                    btn.config(bg="#F0F2F5", fg="#666")
            self._update_kit_view()

        for mode, label in modes:
            btn = tk.Button(btn_row, text=label,
                             command=lambda m=mode: select_mode(m),
                             bg="#F0F2F5", fg="#666",
                             font=("맑은 고딕", 9, "bold"),
                             relief="flat", padx=14, pady=6,
                             cursor="hand2")
            btn.pack(side="left", padx=2)
            self._kit_mode_btns[mode] = btn

        # 기본 색상 강조
        self._kit_mode_btns["color"].config(bg="#1877F2", fg="white")

        # ===== 액션 버튼 - 하단 고정 =====
        action_outer = tk.Frame(wrap, bg="#F5F6F8")
        action_outer.pack(side="bottom", fill="x", padx=18, pady=(0, 14))

        def make_btn(parent, text, bg, hover_bg, command):
            shadow = tk.Frame(parent, bg=hover_bg)
            shadow.pack(side="left", expand=True, fill="x", padx=2)
            btn = tk.Button(shadow, text=text, bg=bg, fg="white",
                              activebackground=hover_bg, activeforeground="white",
                              font=("맑은 고딕", 10, "bold"),
                              relief="flat", bd=0, cursor="hand2", command=command)
            btn.pack(fill="x", ipady=8)
            def on_e(e): btn.config(bg=hover_bg)
            def on_l(e): btn.config(bg=bg)
            btn.bind("<Enter>", on_e); btn.bind("<Leave>", on_l)
            return btn

        make_btn(action_outer, "🔄  변환 실행",
                  bg="#F59E0B", hover_bg="#D97706",
                  command=self.run_kit_conversion)
        make_btn(action_outer, "📎  결과 복사",
                  bg="#10B981", hover_bg="#059669",
                  command=self.copy_kit_to_clipboard)

        # ===== 결과 리포트 - 하단 (버튼 위) =====
        report_outer = tk.Frame(wrap, bg="#F5F6F8")
        report_outer.pack(side="bottom", fill="x", padx=18, pady=(0, 8))

        report_card = tk.Frame(report_outer, bg="white",
                                 highlightthickness=1, highlightbackground="#E5E7EB")
        report_card.pack(fill="x")
        tk.Frame(report_card, bg="#8B5CF6", width=4).pack(side="left", fill="y")

        report_inner = tk.Frame(report_card, bg="white", padx=14, pady=10)
        report_inner.pack(side="left", fill="both", expand=True)

        tk.Label(report_inner, text="📝  변환 결과 / 매칭 리포트",
                 bg="white", fg="#111827",
                 font=("맑은 고딕", 10, "bold")).pack(anchor="w", pady=(0, 4))

        self.txt_kit_report = tk.Text(report_inner, height=4,
                                        font=("Consolas", 10),
                                        bg="#FAFAFA", bd=1, relief="solid",
                                        highlightthickness=1, highlightbackground="#E5E7EB",
                                        padx=8, pady=6)
        self.txt_kit_report.pack(fill="x")
        self.txt_kit_report.tag_config("ok", foreground="#16A34A", font=("Consolas", 10, "bold"))
        self.txt_kit_report.tag_config("warn", foreground="#DC2626", font=("Consolas", 10, "bold"))
        self.txt_kit_report.tag_config("info", foreground="#555", font=("Consolas", 10))

        # ===== 입력 카드 (메인) =====
        input_outer = tk.Frame(wrap, bg="#F5F6F8")
        input_outer.pack(fill="both", expand=True, padx=18, pady=(0, 8))

        input_card = tk.Frame(input_outer, bg="white",
                                highlightthickness=1, highlightbackground="#E5E7EB")
        input_card.pack(fill="both", expand=True)
        tk.Frame(input_card, bg="#10B981", width=4).pack(side="left", fill="y")

        input_inner = tk.Frame(input_card, bg="white", padx=14, pady=12)
        input_inner.pack(side="left", fill="both", expand=True)

        head = tk.Frame(input_inner, bg="white")
        head.pack(fill="x", pady=(0, 4))
        tk.Label(head, text="📥",
                 bg="white", font=("맑은 고딕", 12)).pack(side="left", padx=(0, 6))
        self.lbl_kit_title = tk.Label(head, text="입력",
                                         font=("맑은 고딕", 10, "bold"),
                                         bg="white", fg="#111827")
        self.lbl_kit_title.pack(side="left")

        clr_btn = tk.Button(head, text="🗑️ 지우기",
                             command=self._clear_kit_input,
                             bg="#F3F4F6", fg="#6B7280",
                             font=("맑은 고딕", 8, "bold"),
                             relief="flat", padx=8, pady=2,
                             cursor="hand2")
        clr_btn.pack(side="right")

        self.lbl_kit_hint = tk.Label(input_inner,
                                       text="💡 안내",
                                       bg="white", fg="#9CA3AF",
                                       font=("맑은 고딕", 8), anchor="w")
        self.lbl_kit_hint.pack(fill="x", pady=(0, 4))

        # 입력창 + 스크롤
        kit_box = tk.Frame(input_inner, bg="white")
        kit_box.pack(fill="both", expand=True)

        self.txt_kit_in = tk.Text(kit_box, font=("Consolas", 10),
                                    bd=1, relief="solid",
                                    highlightthickness=1, highlightbackground="#E5E7EB",
                                    bg="#FAFAFA", padx=8, pady=6, wrap="char",
                                    height=14)
        kit_sb = ttk.Scrollbar(kit_box, orient="vertical",
                                 command=self.txt_kit_in.yview)
        self.txt_kit_in.configure(yscrollcommand=kit_sb.set)
        kit_sb.pack(side="right", fill="y")
        self.txt_kit_in.pack(side="left", fill="both", expand=True)

        # 마지막 결과
        self._last_kit_df = None

        # 초기 모드 설정
        self._update_kit_view()

    def _update_kit_view(self):
        """선택된 모드에 따라 입력창 안내 + 타이틀 업데이트"""
        mode = self._kit_mode.get()
        if mode == "color":
            self.lbl_kit_title.config(text="브랜드코드 입력 (한 줄에 하나씩)")
            self.lbl_kit_hint.config(
                text="💡 예: R1004 ← 한 줄에 하나씩. 각 브랜드코드별로 색상코드 5개(999) 자동 채움")
        elif mode == "product":
            self.lbl_kit_title.config(text="상품코드 + 수정내용")
            self.lbl_kit_hint.config(
                text="💡 예: EUNW251MBTT010BK<TAB>158000 ← 상품코드와 수정값을 탭/공백으로 구분")
        else:  # option
            self.lbl_kit_title.config(text="상품코드 + 사이즈 + 새 코드")
            self.lbl_kit_hint.config(
                text="💡 예: EUNW251MBTT010BK<TAB>F<TAB>EUNW251MBTT010BK000 ← 현재 코드, 사이즈, 새 코드")

    def _clear_kit_input(self):
        existing = self.txt_kit_in.get("1.0", "end-1c").strip()
        if existing and not messagebox.askyesno("입력 비우기", "내용을 모두 지우시겠습니까?"):
            return
        self.txt_kit_in.delete("1.0", tk.END)
        self.txt_kit_report.delete("1.0", tk.END)
        self._last_kit_df = None

    def run_kit_conversion(self):
        """선택된 모드에 따라 변환 실행"""
        raw = self.txt_kit_in.get("1.0", "end-1c").strip()
        if not raw:
            messagebox.showwarning("주의", "입력값을 먼저 작성해주세요.")
            return

        mode = self._kit_mode.get()
        self.txt_kit_report.delete("1.0", tk.END)

        try:
            if mode == "color":
                self._run_kit_color(raw)
            elif mode == "product":
                self._run_kit_product(raw)
            else:  # option
                self._run_kit_option(raw)
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("오류", f"변환 실패: {e}")

    def _run_kit_color(self, raw):
        """색상코드 등록: 브랜드코드 → 999×5"""
        rows = []
        for line in raw.split('\n'):
            code = line.strip()
            if not code: continue
            # 탭이나 공백 있으면 첫 토큰만
            code = re.split(r'[\t ]+', code)[0]
            rows.append({
                '브랜드코드': code,
                '색상코드': '999',
                '색상명': '999',
                '색상명1': '999',
                '색상명2': '999',
                '색상명3': '999',
            })

        if not rows:
            messagebox.showwarning("결과 없음", "유효한 브랜드코드가 없습니다.")
            return

        df = pd.DataFrame(rows)
        self._last_kit_df = df
        self.txt_kit_report.insert(tk.END, f"✅ 색상코드 등록 양식 생성: {len(df)}건\n", "ok")
        self.txt_kit_report.insert(tk.END,
            f"   각 브랜드코드별로 색상코드/색상명/명1/명2/명3 = 999 5개씩 자동 채움\n", "info")

    def _run_kit_product(self, raw):
        """상품 관리 수정: 상품코드 + 수정내용 → 그대로 출력"""
        rows = []
        for line in raw.split('\n'):
            line = line.strip()
            if not line: continue
            tokens = re.split(r'[\t]+', line) if '\t' in line else re.split(r'\s{2,}', line)
            if len(tokens) < 2:
                tokens = re.split(r'\s+', line, maxsplit=1)
            tokens = [t.strip() for t in tokens if t.strip()]
            if len(tokens) < 2: continue
            rows.append({
                '상품코드': tokens[0],
                '수정내용': tokens[1],
            })

        if not rows:
            messagebox.showwarning("결과 없음", "상품코드+수정내용 형식의 데이터를 인식하지 못했습니다.")
            return

        df = pd.DataFrame(rows)
        self._last_kit_df = df
        self.txt_kit_report.insert(tk.END, f"✅ 상품 관리 수정 양식 생성: {len(df)}건\n", "ok")
        self.txt_kit_report.insert(tk.END,
            f"   복사 후 시트의 적절한 컬럼(예: 판매가, 상품명 등) 위치에 붙여넣기\n", "info")

    def _run_kit_option(self, raw):
        """상품 옵션 수정: 상품코드 + 사이즈 + 새코드 → 색상코드 999 추가"""
        rows = []
        for line in raw.split('\n'):
            line = line.strip()
            if not line: continue
            tokens = re.split(r'[\t]+', line) if '\t' in line else re.split(r'\s+', line)
            tokens = [t.strip() for t in tokens if t.strip()]
            if len(tokens) < 3:
                # 3개 미만이면 스킵
                continue
            rows.append({
                '상품코드': tokens[0],
                '색상코드': '999',
                '사이즈코드': tokens[1],
                '수정내용': tokens[2],
            })

        if not rows:
            messagebox.showwarning("결과 없음",
                "상품코드 + 사이즈 + 새코드 (3개) 형식이 필요합니다.")
            return

        df = pd.DataFrame(rows)
        self._last_kit_df = df
        self.txt_kit_report.insert(tk.END, f"✅ 상품 옵션 수정 양식 생성: {len(df)}건\n", "ok")
        self.txt_kit_report.insert(tk.END,
            f"   색상코드는 999로 자동 채워짐\n", "info")

    def copy_kit_to_clipboard(self):
        """수정 키트 결과를 클립보드에 복사 (TSV 헤더 제외)"""
        df = getattr(self, '_last_kit_df', None)
        if df is None or df.empty:
            messagebox.showwarning("복사 불가",
                "복사할 데이터가 없습니다.\n먼저 [🔄 변환 실행]을 눌러주세요.")
            return
        try:
            tsv = df.to_csv(sep='\t', index=False, header=False)
            self.root.clipboard_clear()
            self.root.clipboard_append(tsv)
            self.root.update()
            messagebox.showinfo("복사 완료",
                f"✅ 클립보드에 복사 완료!\n\n📊 {len(df)}건\n📋 시트에 Ctrl+V")
        except Exception as e:
            messagebox.showerror("복사 실패", f"{e}")

    def _clear_master_input(self):
        """마스터 등록 입력창 + 리포트 비우기"""
        existing = self.txt_master_in.get("1.0", "end-1c").strip()
        if existing and not messagebox.askyesno("입력 비우기", "내용을 모두 지우시겠습니까?"):
            return
        self.txt_master_in.delete("1.0", tk.END)
        self.txt_master_report.delete("1.0", tk.END)
        self._last_master_table_df = None
        self._last_option_table_df = None

    # ========== [브랜드 캐시 + 매니저] ==========
    def refresh_brand_cache(self):
        """Firestore의 brand_master 컬렉션을 메모리에 로드 (수동 갱신용)"""
        try:
            docs = self.db.collection('brand_master').get()
            self._brand_cache = {}
            for d in docs:
                data = d.to_dict()
                self._brand_cache[d.id] = data.get('brand_name', d.id)
            print(f"✅ 브랜드 캐시: {len(self._brand_cache)}건")
        except Exception as e:
            print(f"⚠️ 브랜드 캐시 로드 실패: {e}")
            self._brand_cache = {}

    def start_brand_listener(self):
        """[추가] brand_master 실시간 리스너 - 다른 PC에서 변경하면 즉시 반영"""
        def on_brand_snapshot(col_snapshot, changes, read_time):
            try:
                for change in changes:
                    doc_id = change.document.id
                    data = change.document.to_dict()
                    if change.type.name == 'ADDED' or change.type.name == 'MODIFIED':
                        self._brand_cache[doc_id] = data.get('brand_name', doc_id)
                    elif change.type.name == 'REMOVED':
                        self._brand_cache.pop(doc_id, None)
            except Exception as e:
                print(f"⚠️ 브랜드 리스너 처리 오류: {e}")

        try:
            query = self.db.collection('brand_master')
            self._brand_listener = query.on_snapshot(on_brand_snapshot)
            print("📡 브랜드 마스터 실시간 리스너 가동")
        except Exception as e:
            print(f"❌ 브랜드 리스너 설정 오류: {e}")

    def open_brand_manager(self):
        """브랜드 관리 팝업 - 등록/수정/삭제"""
        win = tk.Toplevel(self.root)
        win.title("🗂️ 브랜드 관리")
        win.configure(bg="#F5F6F8")
        try:
            self.position_popup(win, 480, 600)
        except Exception:
            win.geometry("480x600")

        # 헤더
        head = tk.Frame(win, bg="#F5F6F8")
        head.pack(fill="x", padx=18, pady=(14, 8))
        tk.Label(head, text="🗂️", font=("맑은 고딕", 18), bg="#F5F6F8").pack(side="left", padx=(0, 6))
        tk.Label(head, text="브랜드 관리",
                 font=("맑은 고딕", 14, "bold"),
                 bg="#F5F6F8", fg="#1A1A1A").pack(side="left")

        # ===== 신규 등록 카드 =====
        add_card = tk.Frame(win, bg="white",
                              highlightthickness=1, highlightbackground="#E5E7EB")
        add_card.pack(fill="x", padx=18, pady=(0, 8))
        tk.Frame(add_card, bg="#10B981", width=4).pack(side="left", fill="y")

        add_inner = tk.Frame(add_card, bg="white", padx=14, pady=10)
        add_inner.pack(side="left", fill="both", expand=True)

        tk.Label(add_inner, text="➕ 신규 브랜드 등록",
                 bg="white", fg="#065F46",
                 font=("맑은 고딕", 10, "bold")).pack(anchor="w", pady=(0, 6))

        row1 = tk.Frame(add_inner, bg="white")
        row1.pack(fill="x", pady=2)
        tk.Label(row1, text="코드 (3자리):", bg="white",
                 font=("맑은 고딕", 9), width=12, anchor="w").pack(side="left")
        ent_code = tk.Entry(row1, font=("Consolas", 10),
                             bd=1, relief="solid", width=10)
        ent_code.pack(side="left", padx=4, ipady=2)

        row2 = tk.Frame(add_inner, bg="white")
        row2.pack(fill="x", pady=2)
        tk.Label(row2, text="브랜드명:", bg="white",
                 font=("맑은 고딕", 9), width=12, anchor="w").pack(side="left")
        ent_name = tk.Entry(row2, font=("맑은 고딕", 10),
                             bd=1, relief="solid")
        ent_name.pack(side="left", fill="x", expand=True, padx=4, ipady=2)

        def add_brand():
            code = ent_code.get().strip().upper()
            name = ent_name.get().strip()
            if not code or not name:
                messagebox.showwarning("입력 누락", "코드와 브랜드명을 모두 입력하세요.", parent=win)
                return
            if len(code) != 3:
                messagebox.showwarning("형식 오류", "코드는 정확히 3자리여야 합니다.", parent=win)
                return
            try:
                if code in self._brand_cache:
                    if not messagebox.askyesno("덮어쓰기",
                            f"코드 '{code}'는 이미 '{self._brand_cache[code]}'로 등록되어 있습니다.\n"
                            f"'{name}'으로 덮어쓰시겠습니까?", parent=win):
                        return
                self.db.collection('brand_master').document(code).set({
                    'brand_name': name,
                    'created_at': firestore.SERVER_TIMESTAMP,
                    'updated_at': firestore.SERVER_TIMESTAMP
                })
                self._brand_cache[code] = name
                ent_code.delete(0, tk.END)
                ent_name.delete(0, tk.END)
                refresh_list()
                messagebox.showinfo("등록 완료", f"✅ {code} → {name}", parent=win)
            except Exception as e:
                messagebox.showerror("오류", f"등록 실패: {e}", parent=win)

        # 등록 + 일괄 가져오기 버튼들
        btn_row = tk.Frame(add_inner, bg="white")
        btn_row.pack(fill="x", pady=(6, 0))

        tk.Button(btn_row, text="➕ 등록",
                   command=add_brand,
                   bg="#10B981", fg="white",
                   font=("맑은 고딕", 9, "bold"),
                   relief="flat", padx=14, pady=6,
                   cursor="hand2").pack(side="right")

        def bulk_import():
            """탭으로 구분된 텍스트 파일에서 일괄 가져오기 (코드<TAB>브랜드명)"""
            path = filedialog.askopenfilename(
                title="브랜드 목록 파일 선택 (탭 구분 텍스트)",
                filetypes=[("텍스트 파일", "*.txt"), ("CSV", "*.csv"), ("모든 파일", "*.*")],
                parent=win
            )
            if not path:
                return
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    lines = f.readlines()
            except UnicodeDecodeError:
                with open(path, 'r', encoding='cp949') as f:
                    lines = f.readlines()

            # 파싱
            new_brands = {}  # 첫 번째 값 우선
            for line in lines:
                line = line.strip()
                if not line: continue
                # 탭 또는 콤마 구분
                parts = re.split(r'[\t,]', line)
                if len(parts) < 2: continue
                code = parts[0].strip().upper()
                name = parts[1].strip()
                if not code or not name: continue
                if code == '코드' or code.lower() == 'code': continue
                if len(code) != 3: continue
                if code not in new_brands:
                    new_brands[code] = name

            if not new_brands:
                messagebox.showwarning("결과 없음",
                    "유효한 브랜드 데이터를 찾을 수 없습니다.\n"
                    "형식: 코드<TAB>브랜드명 (한 줄당 하나)", parent=win)
                return

            # 기존 데이터와 비교
            new_count = sum(1 for c in new_brands if c not in self._brand_cache)
            update_count = sum(1 for c, n in new_brands.items()
                                if c in self._brand_cache and self._brand_cache[c] != n)
            same_count = sum(1 for c, n in new_brands.items()
                              if c in self._brand_cache and self._brand_cache[c] == n)

            if not messagebox.askyesno("일괄 가져오기 확인",
                f"📊 발견된 브랜드: {len(new_brands)}건\n\n"
                f"  ✨ 신규 등록: {new_count}건\n"
                f"  🔄 변경 (덮어쓰기): {update_count}건\n"
                f"  ✓ 동일 (스킵): {same_count}건\n\n"
                f"진행하시겠습니까?", parent=win):
                return

            # Firestore에 일괄 등록 - batch 사용
            try:
                # Firestore batch는 500건 제한이라 분할
                items = list(new_brands.items())
                total_done = 0
                CHUNK = 400

                for i in range(0, len(items), CHUNK):
                    chunk = items[i:i+CHUNK]
                    batch = self.db.batch()
                    for code, name in chunk:
                        # 동일하면 스킵
                        if self._brand_cache.get(code) == name:
                            continue
                        ref = self.db.collection('brand_master').document(code)
                        batch.set(ref, {
                            'brand_name': name,
                            'updated_at': firestore.SERVER_TIMESTAMP,
                            'created_at': firestore.SERVER_TIMESTAMP
                        }, merge=True)
                        self._brand_cache[code] = name
                        total_done += 1
                    batch.commit()

                refresh_list()
                messagebox.showinfo("일괄 등록 완료",
                    f"✅ {total_done}건 등록되었습니다!\n\n"
                    f"(동일한 {same_count}건은 스킵됨)", parent=win)
            except Exception as e:
                messagebox.showerror("오류", f"일괄 등록 실패: {e}", parent=win)

        tk.Button(btn_row, text="📥 일괄 가져오기",
                   command=bulk_import,
                   bg="#3B82F6", fg="white",
                   font=("맑은 고딕", 9, "bold"),
                   relief="flat", padx=14, pady=6,
                   cursor="hand2").pack(side="right", padx=(0, 6))

        # ===== 등록 리스트 카드 =====
        list_card = tk.Frame(win, bg="white",
                               highlightthickness=1, highlightbackground="#E5E7EB")
        list_card.pack(fill="both", expand=True, padx=18, pady=(0, 14))
        tk.Frame(list_card, bg="#3B82F6", width=4).pack(side="left", fill="y")

        list_inner = tk.Frame(list_card, bg="white", padx=14, pady=10)
        list_inner.pack(side="left", fill="both", expand=True)

        list_head = tk.Frame(list_inner, bg="white")
        list_head.pack(fill="x", pady=(0, 6))
        list_title = tk.Label(list_head, text="📋 등록된 브랜드",
                                bg="white", fg="#1E40AF",
                                font=("맑은 고딕", 10, "bold"))
        list_title.pack(side="left")

        search_var = tk.StringVar()
        search_entry = tk.Entry(list_head, textvariable=search_var,
                                  font=("맑은 고딕", 9),
                                  bd=1, relief="solid", width=12)
        search_entry.pack(side="right", ipady=2)
        tk.Label(list_head, text="🔍", bg="white",
                 font=("맑은 고딕", 9)).pack(side="right", padx=(0, 4))

        tree_frame = tk.Frame(list_inner, bg="white")
        tree_frame.pack(fill="both", expand=True)

        cols = ("code", "name")
        tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=15)
        tree.heading("code", text="코드")
        tree.heading("name", text="브랜드명")
        tree.column("code", width=80, anchor="center")
        tree.column("name", width=300, anchor="w")
        tree.pack(side="left", fill="both", expand=True)

        sb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")

        def refresh_list():
            self.refresh_brand_cache()
            for it in tree.get_children():
                tree.delete(it)
            keyword = search_var.get().strip().lower()
            sorted_codes = sorted(self._brand_cache.keys())
            for code in sorted_codes:
                name = self._brand_cache[code]
                if keyword and keyword not in code.lower() and keyword not in name.lower():
                    continue
                tree.insert("", "end", iid=code, values=(code, name))
            list_title.config(text=f"📋 등록된 브랜드 ({len(tree.get_children())}건)")

        search_var.trace_add("write", lambda *a: refresh_list())

        ctx_menu = tk.Menu(win, tearoff=0, font=("맑은 고딕", 10))

        def edit_brand():
            sel = tree.selection()
            if not sel: return
            code = sel[0]
            old_name = self._brand_cache.get(code, '')
            new_name = simpledialog.askstring("브랜드명 수정",
                                                 f"코드: {code}\n새 브랜드명:",
                                                 initialvalue=old_name, parent=win)
            if new_name is None or not new_name.strip():
                return
            try:
                self.db.collection('brand_master').document(code).update({
                    'brand_name': new_name.strip(),
                    'updated_at': firestore.SERVER_TIMESTAMP
                })
                self._brand_cache[code] = new_name.strip()
                refresh_list()
            except Exception as e:
                messagebox.showerror("오류", f"수정 실패: {e}", parent=win)

        def delete_brand():
            sel = tree.selection()
            if not sel: return
            code = sel[0]
            name = self._brand_cache.get(code, '')
            if not messagebox.askyesno("삭제 확인",
                                          f"'{code} - {name}' 브랜드를 삭제하시겠습니까?",
                                          parent=win):
                return
            try:
                self.db.collection('brand_master').document(code).delete()
                self._brand_cache.pop(code, None)
                refresh_list()
            except Exception as e:
                messagebox.showerror("오류", f"삭제 실패: {e}", parent=win)

        ctx_menu.add_command(label="✏️ 수정", command=edit_brand)
        ctx_menu.add_separator()
        ctx_menu.add_command(label="🗑️ 삭제", command=delete_brand)

        def show_ctx(e):
            row = tree.identify_row(e.y)
            if row:
                tree.selection_set(row)
                ctx_menu.post(e.x_root, e.y_root)
        tree.bind("<Button-3>", show_ctx)
        tree.bind("<Double-Button-1>", lambda e: edit_brand())

        refresh_list()

    # ========== [바코드표 → 마스터/옵션 양식 변환] ==========
    def _parse_master_input(self):
        """입력창의 바코드표를 파싱.
        헤더 행은 자동 스킵, 컬럼 위치를 자동 인식.
        Returns: list of dict [{상품명, 바코드번호, 상품메모1, 상품메모2, 대표판매가, 옵션내용, 수량}]
        """
        raw = self.txt_master_in.get("1.0", "end-1c").strip()
        if not raw:
            return None

        rows = []
        lines = raw.split('\n')

        # 헤더 행 찾기 (위 5행까지) - "상품명", "바코드", "상품메모" 같은 키워드 있는 행
        header_idx = -1
        col_map = {}  # {필드명: 컬럼 인덱스}

        header_keywords = {
            '상품명': 'product_name',
            '바코드': 'barcode',
            '바코드번호': 'barcode',
            '상품메모1': 'memo1',
            '상품메모2': 'memo2',
            '대표판매가': 'price',
            '판매가': 'price',
            '옵션내용': 'size',
            '옵션': 'size',
            '사이즈': 'size',
            '수량': 'qty',
        }

        for i, line in enumerate(lines[:5]):
            tokens = re.split(r'\t', line)  # 엑셀 복붙은 탭 구분이 정상
            if len(tokens) < 4:
                continue

            tmp_map = {}
            for col_idx, tok in enumerate(tokens):
                tok_clean = tok.strip()
                # 정확 매칭 우선
                if tok_clean in header_keywords:
                    field = header_keywords[tok_clean]
                    if field not in tmp_map:
                        tmp_map[field] = col_idx

            # 헤더로 인식: 적어도 바코드 + memo1이 있어야 함
            if 'barcode' in tmp_map and 'memo1' in tmp_map:
                header_idx = i
                col_map = tmp_map
                break

        if header_idx == -1:
            # 헤더 못 찾음 → 위치 기반으로 추정 (스샷 형식 기준)
            # 상품명 | 바코드번호 | 상품메모1 | 상품메모2 | 대표판매가 | 옵션내용 | 수량
            col_map = {
                'product_name': 0, 'barcode': 1, 'memo1': 2, 'memo2': 3,
                'price': 4, 'size': 5, 'qty': 6
            }
            data_lines = lines
        else:
            data_lines = lines[header_idx + 1:]

        # 데이터 파싱
        for line in data_lines:
            if not line.strip():
                continue
            tokens = re.split(r'\t', line)
            if len(tokens) < 3:
                continue

            def get_col(field, default=''):
                idx = col_map.get(field)
                if idx is None or idx >= len(tokens):
                    return default
                return tokens[idx].strip()

            barcode = get_col('barcode')
            memo1 = get_col('memo1')

            # 바코드와 품번 둘 다 비어있으면 스킵
            if not barcode and not memo1:
                continue

            rows.append({
                '상품명': get_col('product_name'),
                '바코드': barcode,
                '품번': memo1,  # 상품메모1 = 품번
                '상품명_상세': get_col('memo2'),  # 상품메모2 = 상세 상품명
                '대표판매가': get_col('price'),
                '옵션내용': get_col('size'),
                '수량': get_col('qty'),
            })

        return rows if rows else None

    def run_master_conversion(self):
        """입력 바코드표 → 마스터 등록용 + 옵션 등록용 두 양식으로 변환"""
        rows = self._parse_master_input()
        if not rows:
            messagebox.showwarning("주의",
                "바코드표를 인식하지 못했습니다.\n"
                "엑셀에서 [상품명/바코드번호/상품메모1/상품메모2/대표판매가/옵션내용/수량] 컬럼이 있는 행을 복사해주세요.")
            return

        # 리포트 비우기
        self.txt_master_report.delete("1.0", tk.END)

        # 브랜드 매칭
        unknown_brands = set()
        unknown_items = set()

        for r in rows:
            품번 = r['품번']
            if not 품번:
                continue
            # 브랜드 코드 = 품번 앞 3자리
            brand_code = 품번[:3].upper() if len(품번) >= 3 else ''
            r['브랜드코드'] = brand_code
            r['브랜드명'] = self._brand_cache.get(brand_code, '')
            if brand_code and not r['브랜드명']:
                unknown_brands.add(brand_code)

            # 아이템 코드 = 품번 9번째 글자
            if len(품번) >= 9:
                item_letter = 품번[8].upper()
                r['아이템코드'] = self.ITEM_CODE_MAP.get(item_letter, '')
                if not r['아이템코드']:
                    unknown_items.add(item_letter)
            else:
                r['아이템코드'] = ''

        # ========== [출력 1: 마스터 등록 양식] ==========
        # 같은 품번끼리 합치기 (한 행으로)
        seen_품번 = set()
        master_rows = []
        for r in rows:
            품번 = r['품번']
            if not 품번 or 품번 in seen_품번:
                continue
            seen_품번.add(품번)

            # 상품명: 메모2 우선, 없으면 상품명
            상품명 = r['상품명_상세'] or r['상품명']
            # 가격: 콤마 제거
            try:
                가격 = int(str(r['대표판매가']).replace(',', '').replace(' ', '').strip() or 0)
            except:
                가격 = 0

            master_rows.append({
                '품번': 품번,
                '대표코드': 품번,
                '상품명': 상품명,
                '사이즈체계코드': '엠프티',
                '브랜드코드': r['브랜드명'] or f"[미등록:{r['브랜드코드']}]",
                '년도코드': '999',
                '시즌코드': '999',
                '성별코드': '999',
                '아이템코드': r['아이템코드'] or f"[미등록:{품번[8] if len(품번)>=9 else '?'}]",
                '택가': 0,
                '정상가': 가격,
                '판가': 0,
                '원가': 0,
            })

        # ========== [출력 2: 옵션 등록 양식] ==========
        # 행 단위 = 바코드 단위
        option_rows = []
        for r in rows:
            if not r['바코드']:
                continue
            option_rows.append({
                '품번': r['품번'],
                '색상코드': '999',
                '사이즈코드': r['옵션내용'],
                '바코드': r['바코드'],
                '고정로케이션': '00-00-00-00',
            })

        self._last_master_table_df = pd.DataFrame(master_rows)
        self._last_option_table_df = pd.DataFrame(option_rows)

        # ========== [리포트 작성] ==========
        report = self.txt_master_report
        report.insert(tk.END,
            f"✅ 인식: 입력 {len(rows)}건 → 마스터 {len(master_rows)}품번 / 옵션 {len(option_rows)}바코드\n", "ok")

        if unknown_brands:
            report.insert(tk.END,
                f"⚠️ 등록되지 않은 브랜드 코드: {', '.join(sorted(unknown_brands))}\n", "warn")
            report.insert(tk.END,
                "    → 우측 [🗂️ 브랜드 관리]에서 먼저 등록해주세요.\n", "info")

        if unknown_items:
            report.insert(tk.END,
                f"⚠️ 알 수 없는 아이템 코드 글자: {', '.join(sorted(unknown_items))}\n", "warn")
            report.insert(tk.END,
                f"    → 알려진 코드: {', '.join(f'{k}={v}' for k,v in self.ITEM_CODE_MAP.items())}\n", "info")

        # 브랜드별 요약
        if master_rows:
            from collections import Counter
            cnt = Counter(r['브랜드코드'] for r in rows if r['품번'])
            for code, c in sorted(cnt.items()):
                if not code:
                    continue
                name = self._brand_cache.get(code, '?')
                report.insert(tk.END, f"  • {code} ({name}): {c}건\n", "info")

    def copy_master_table_to_clipboard(self):
        """마스터 등록 양식을 클립보드에 복사 (TSV 헤더 제외)"""
        df = getattr(self, '_last_master_table_df', None)
        if df is None or df.empty:
            messagebox.showwarning("복사 불가",
                "복사할 데이터가 없습니다.\n먼저 [🔄 변환 실행]을 눌러주세요.")
            return

        # 미등록 검사
        unknown = df[df['브랜드코드'].astype(str).str.startswith('[미등록:')]
        unknown_item = df[df['아이템코드'].astype(str).str.startswith('[미등록:')]
        if len(unknown) > 0 or len(unknown_item) > 0:
            msg = []
            if len(unknown) > 0:
                msg.append(f"⚠️ 미등록 브랜드 {len(unknown)}건")
            if len(unknown_item) > 0:
                msg.append(f"⚠️ 미등록 아이템 {len(unknown_item)}건")
            if not messagebox.askyesno("미등록 포함",
                f"{', '.join(msg)}이 포함되어 있습니다.\n그래도 복사하시겠습니까?"):
                return

        try:
            tsv = df.to_csv(sep='\t', index=False, header=False)
            self.root.clipboard_clear()
            self.root.clipboard_append(tsv)
            self.root.update()
            messagebox.showinfo("복사 완료",
                f"✅ 마스터 등록 양식 복사 완료!\n\n"
                f"📊 {len(df)}품번\n"
                f"📋 [상품관리 시트]에 Ctrl+V")
        except Exception as e:
            messagebox.showerror("복사 실패", f"{e}")

    def copy_option_table_to_clipboard(self):
        """옵션 등록 양식을 클립보드에 복사 (TSV 헤더 제외)"""
        df = getattr(self, '_last_option_table_df', None)
        if df is None or df.empty:
            messagebox.showwarning("복사 불가",
                "복사할 데이터가 없습니다.\n먼저 [🔄 변환 실행]을 눌러주세요.")
            return
        try:
            tsv = df.to_csv(sep='\t', index=False, header=False)
            self.root.clipboard_clear()
            self.root.clipboard_append(tsv)
            self.root.update()
            messagebox.showinfo("복사 완료",
                f"✅ 옵션 등록 양식 복사 완료!\n\n"
                f"📊 {len(df)}바코드\n"
                f"📋 [상품옵션 시트]에 Ctrl+V")
        except Exception as e:
            messagebox.showerror("복사 실패", f"{e}")

    # --- [탭 4: 마감재고 (수량 너비 확보)] ---
    # 캐파 기본값 (코드 폴백용 - 실제로는 Firestore에서 받아옴)
    DEFAULT_ZONE_CAPACITY = {
        'AA~AB구역': 30000,
        'BB구역': 3200,
        'CC~DD구역': 4400,
        'EA/EE/FF구역': 13696,
    }

    # ========================================================================
    # [탭: RT 바코드 스캔] - RT 바코드 스캔 → 입고 양식 변환
    # ========================================================================
    def setup_rt_inbound(self):
        """RT 바코드 스캔 → 입고 양식 자동 변환"""
        container = tk.Frame(self.t_rt, bg="#F5F6F8")
        container.pack(fill="both", expand=True)

        # ========== [상단 헤더] ==========
        header_frame = tk.Frame(container, bg="#F5F6F8")
        header_frame.pack(side="top", fill="x", padx=24, pady=(16, 8))

        title_left = tk.Frame(header_frame, bg="#F5F6F8")
        title_left.pack(side="left")
        tk.Label(title_left, text="🔄", font=("맑은 고딕", 22),
                 bg="#F5F6F8").pack(side="left", padx=(0, 6))
        title_text_box = tk.Frame(title_left, bg="#F5F6F8")
        title_text_box.pack(side="left")
        tk.Label(title_text_box, text="RT 입고",
                 font=("맑은 고딕", 15, "bold"),
                 bg="#F5F6F8", fg="#1A1A1A").pack(anchor="w")
        tk.Label(title_text_box, text="바코드 스캔 → 입고 양식 자동 변환",
                 font=("맑은 고딕", 8),
                 bg="#F5F6F8", fg="#888").pack(anchor="w")

        # 우측: 리셋 버튼
        tk.Button(header_frame, text="🔄 리셋",
                   command=self.reset_rt_all,
                   bg="#F3F4F6", fg="#9CA3AF",
                   activebackground="#E5E7EB",
                   font=("맑은 고딕", 8),
                   relief="flat", padx=8, pady=4,
                   cursor="hand2").pack(side="right", anchor="se")

        # ========== [액션 버튼 - 하단] ==========
        action_outer = tk.Frame(container, bg="#F5F6F8")
        action_outer.pack(side="bottom", fill="x", padx=18, pady=(0, 14))

        def make_btn(parent, text, bg, hover_bg, command):
            shadow = tk.Frame(parent, bg=hover_bg)
            shadow.pack(side="left", expand=True, fill="x", padx=2)
            btn = tk.Button(shadow, text=text, bg=bg, fg="white",
                              activebackground=hover_bg, activeforeground="white",
                              font=("맑은 고딕", 10, "bold"),
                              relief="flat", bd=0, cursor="hand2", command=command)
            btn.pack(fill="x", ipady=8)
            def on_e(e): btn.config(bg=hover_bg)
            def on_l(e): btn.config(bg=bg)
            btn.bind("<Enter>", on_e); btn.bind("<Leave>", on_l)
            return btn

        make_btn(action_outer, "🔄  변환 실행",
                  bg="#F59E0B", hover_bg="#D97706",
                  command=self.run_rt_conversion)
        make_btn(action_outer, "📎  입고 양식 복사",
                  bg="#10B981", hover_bg="#059669",
                  command=self.copy_rt_to_clipboard)

        # ========== [총 스캔 수량 - 큰 표시] ==========
        total_outer = tk.Frame(container, bg="#F5F6F8")
        total_outer.pack(side="bottom", fill="x", padx=18, pady=(0, 8))

        total_card = tk.Frame(total_outer, bg="#1877F2",
                                highlightthickness=0)
        total_card.pack(fill="x")

        total_inner = tk.Frame(total_card, bg="#1877F2")
        total_inner.pack(pady=12)

        tk.Label(total_inner, text="📡", bg="#1877F2",
                 font=("맑은 고딕", 16)).pack(side="left", padx=(0, 8))
        tk.Label(total_inner, text="총 스캔 수량:",
                 bg="#1877F2", fg="#BFDBFE",
                 font=("맑은 고딕", 11)).pack(side="left", padx=(0, 8))
        self.lbl_rt_total = tk.Label(total_inner, text="0개",
                                        font=("맑은 고딕", 18, "bold"),
                                        bg="#1877F2", fg="white")
        self.lbl_rt_total.pack(side="left")
        tk.Label(total_inner, text="  /  고유 바코드:",
                 bg="#1877F2", fg="#BFDBFE",
                 font=("맑은 고딕", 11)).pack(side="left", padx=(8, 6))
        self.lbl_rt_unique = tk.Label(total_inner, text="0종",
                                        font=("맑은 고딕", 14, "bold"),
                                        bg="#1877F2", fg="white")
        self.lbl_rt_unique.pack(side="left")

        # ========== [로케이션 카드] ==========
        loc_outer = tk.Frame(container, bg="#F5F6F8")
        loc_outer.pack(fill="x", padx=18, pady=(0, 8))

        loc_card = tk.Frame(loc_outer, bg="white",
                              highlightthickness=1, highlightbackground="#E5E7EB")
        loc_card.pack(fill="x")
        tk.Frame(loc_card, bg="#F59E0B", width=4).pack(side="left", fill="y")

        loc_inner = tk.Frame(loc_card, bg="white", padx=14, pady=10)
        loc_inner.pack(side="left", fill="both", expand=True)

        tk.Label(loc_inner, text="📍",
                 bg="white", font=("맑은 고딕", 12)).pack(side="left", padx=(0, 6))
        tk.Label(loc_inner, text="로케이션",
                 bg="white", fg="#374151",
                 font=("맑은 고딕", 10, "bold")).pack(side="left", padx=(0, 8))

        self.ent_rt_location = tk.Entry(loc_inner, font=("Consolas", 11),
                                          bd=1, relief="solid",
                                          highlightthickness=0,
                                          width=20)
        self.ent_rt_location.pack(side="left", padx=(0, 8), ipady=4)
        self.ent_rt_location.insert(0, "RT-00-00-00")

        tk.Label(loc_inner, text="(기본값 RT-00-00-00, 수정 가능)",
                 bg="white", fg="#9CA3AF",
                 font=("맑은 고딕", 8)).pack(side="left")

        # ========== [입력 카드] ==========
        input_outer = tk.Frame(container, bg="#F5F6F8")
        input_outer.pack(fill="both", expand=True, padx=18, pady=(0, 8))

        input_card = tk.Frame(input_outer, bg="white",
                                highlightthickness=1, highlightbackground="#E5E7EB")
        input_card.pack(fill="both", expand=True)
        tk.Frame(input_card, bg="#10B981", width=4).pack(side="left", fill="y")

        input_inner = tk.Frame(input_card, bg="white", padx=14, pady=10)
        input_inner.pack(side="left", fill="both", expand=True)

        head = tk.Frame(input_inner, bg="white")
        head.pack(fill="x", pady=(0, 4))
        tk.Label(head, text="📡",
                 bg="white", font=("맑은 고딕", 12)).pack(side="left", padx=(0, 6))
        tk.Label(head, text="바코드 스캔",
                 font=("맑은 고딕", 10, "bold"),
                 bg="white", fg="#111827").pack(side="left")

        # 지우기 버튼
        tk.Button(head, text="🗑️ 지우기",
                   command=self._clear_rt_input,
                   bg="#F3F4F6", fg="#6B7280",
                   font=("맑은 고딕", 8, "bold"),
                   relief="flat", padx=8, pady=2,
                   cursor="hand2").pack(side="right")

        tk.Label(input_inner,
                 text="💡 바코드 한 줄에 하나씩 스캔 · 같은 바코드 자동 합산 · 모두 정상 입고로 처리됨",
                 bg="white", fg="#9CA3AF",
                 font=("맑은 고딕", 8), anchor="w").pack(fill="x", pady=(0, 4))

        # 입력창 + 스크롤
        rt_box = tk.Frame(input_inner, bg="white")
        rt_box.pack(fill="both", expand=True)

        self.txt_rt_in = tk.Text(rt_box, font=("Consolas", 11),
                                   bd=1, relief="solid",
                                   bg="#F0FDF4",
                                   highlightthickness=1, highlightbackground="#E5E7EB",
                                   padx=8, pady=6, wrap="char",
                                   height=12)
        rt_sb = ttk.Scrollbar(rt_box, orient="vertical",
                                command=self.txt_rt_in.yview)
        self.txt_rt_in.configure(yscrollcommand=rt_sb.set)
        rt_sb.pack(side="right", fill="y")
        self.txt_rt_in.pack(side="left", fill="both", expand=True)

        # 입력 시 실시간 카운트
        self.txt_rt_in.bind("<KeyRelease>", lambda e: self._update_rt_count())

        # 마지막 결과
        self._last_rt_df = None

    def _update_rt_count(self):
        """실시간 총 스캔 수량 업데이트"""
        raw = self.txt_rt_in.get("1.0", "end-1c").strip()
        if not raw:
            self.lbl_rt_total.config(text="0개")
            self.lbl_rt_unique.config(text="0종")
            return

        # 한 줄에 하나씩, 빈 줄 무시
        codes = []
        for line in raw.split('\n'):
            code = line.strip()
            if code:
                # 첫 토큰만 (공백/탭으로 구분된 경우)
                code = re.split(r'[\t ]+', code)[0]
                if code:
                    codes.append(code)

        total = len(codes)
        unique = len(set(codes))
        self.lbl_rt_total.config(text=f"{total:,}개")
        self.lbl_rt_unique.config(text=f"{unique:,}종")

    def _clear_rt_input(self):
        existing = self.txt_rt_in.get("1.0", "end-1c").strip()
        if existing and not messagebox.askyesno("입력 비우기", "바코드 입력을 모두 지우시겠습니까?"):
            return
        self.txt_rt_in.delete("1.0", tk.END)
        self.lbl_rt_total.config(text="0개")
        self.lbl_rt_unique.config(text="0종")
        self._last_rt_df = None

    def reset_rt_all(self):
        """RT 탭 전체 리셋 - 바코드 + 로케이션"""
        if not messagebox.askyesno("RT 탭 리셋",
                                     "바코드와 로케이션을 모두 초기화하시겠습니까?"):
            return
        self.txt_rt_in.delete("1.0", tk.END)
        self.ent_rt_location.delete(0, tk.END)
        self.ent_rt_location.insert(0, "RT-00-00-00")
        self.lbl_rt_total.config(text="0개")
        self.lbl_rt_unique.config(text="0종")
        self._last_rt_df = None

    def run_rt_conversion(self):
        """바코드 스캔 → 입고 양식 변환"""
        raw = self.txt_rt_in.get("1.0", "end-1c").strip()
        if not raw:
            messagebox.showwarning("주의", "바코드를 먼저 스캔해주세요.")
            return

        location = self.ent_rt_location.get().strip()
        if not location:
            location = "RT-00-00-00"

        # 바코드별 카운트
        from collections import Counter
        codes = []
        for line in raw.split('\n'):
            code = line.strip()
            if not code: continue
            code = re.split(r'[\t ]+', code)[0]
            if code:
                codes.append(code)

        if not codes:
            messagebox.showwarning("결과 없음", "유효한 바코드가 없습니다.")
            return

        cnt = Counter(codes)

        # 입고 등록 양식 (입고탭과 동일):
        # 박스번호 | 바코드 | 정상수량 | 불량수량 | 정상로케이션 | 불량로케이션
        # RT 입고는 정상만 처리 → 불량수량 0, 불량로케이션 00-00-00-00 고정
        rows = []
        for bc, qty in cnt.items():
            rows.append({
                '박스번호': '',
                '바코드': bc,
                '정상수량': qty,
                '불량수량': 0,
                '정상로케이션': location,
                '불량로케이션': '00-00-00-00',
            })

        df = pd.DataFrame(rows)
        self._last_rt_df = df

        messagebox.showinfo("변환 완료",
            f"✅ {len(df)}종 / 총 {sum(cnt.values())}개\n\n"
            f"📍 로케이션: {location}\n\n"
            f"[📎 입고 양식 복사] 버튼을 눌러 클립보드에 복사하세요.")

    def copy_rt_to_clipboard(self):
        """RT 바코드 입고 양식을 클립보드에 복사 (TSV 헤더 제외)"""
        df = getattr(self, '_last_rt_df', None)
        if df is None or df.empty:
            messagebox.showwarning("복사 불가",
                "복사할 데이터가 없습니다.\n먼저 [🔄 변환 실행]을 눌러주세요.")
            return
        try:
            tsv = df.to_csv(sep='\t', index=False, header=False)
            self.root.clipboard_clear()
            self.root.clipboard_append(tsv)
            self.root.update()
            messagebox.showinfo("복사 완료",
                f"✅ 클립보드에 복사 완료!\n\n📊 {len(df)}종 / 총 {int(df['정상수량'].sum())}개\n📋 EMP에 Ctrl+V")
        except Exception as e:
            messagebox.showerror("복사 실패", f"{e}")

    def setup_closing_stock(self):
        """마감재고 - 진행바 차트 + 캐파 대비 % + 드래그앤드롭"""
        container = tk.Frame(self.t_end, bg="#F5F6F8")
        container.pack(fill="both", expand=True)

        # 캐파 캐시 + 로드
        self._zone_capacity = dict(self.DEFAULT_ZONE_CAPACITY)
        self.refresh_zone_capacity()

        # ========== [상단 헤더] ==========
        header_frame = tk.Frame(container, bg="#F5F6F8")
        header_frame.pack(side="top", fill="x", padx=24, pady=(16, 8))

        title_left = tk.Frame(header_frame, bg="#F5F6F8")
        title_left.pack(side="left")
        tk.Label(title_left, text="📊", font=("맑은 고딕", 22),
                 bg="#F5F6F8").pack(side="left", padx=(0, 6))
        title_text_box = tk.Frame(title_left, bg="#F5F6F8")
        title_text_box.pack(side="left")
        tk.Label(title_text_box, text="실시간 마감재고",
                 font=("맑은 고딕", 15, "bold"),
                 bg="#F5F6F8", fg="#1A1A1A").pack(anchor="w")
        tk.Label(title_text_box, text="구역별 캐파 대비 가용재고 비율",
                 font=("맑은 고딕", 8),
                 bg="#F5F6F8", fg="#888").pack(anchor="w")

        # 우측: 캐파 설정 + 파일 선택 버튼들
        right_btns = tk.Frame(header_frame, bg="#F5F6F8")
        right_btns.pack(side="right")

        tk.Button(right_btns, text="⚙️ 캐파 설정",
                   command=self.open_capacity_settings,
                   bg="white", fg="#444",
                   font=("맑은 고딕", 9, "bold"),
                   relief="flat", padx=10, pady=6,
                   cursor="hand2",
                   highlightthickness=1, highlightbackground="#DDD").pack(side="right", padx=(8, 0))

        tk.Button(right_btns, text="📁 파일 선택",
                   command=self.run_closing_stock_logic,
                   bg="#1877F2", fg="white",
                   activebackground="#1864c8",
                   font=("맑은 고딕", 9, "bold"),
                   relief="flat", padx=12, pady=6,
                   cursor="hand2").pack(side="right")

        # ========== [총합 + 점유율 - 하단 고정] ==========
        total_outer = tk.Frame(container, bg="#F5F6F8")
        total_outer.pack(side="bottom", fill="x", padx=18, pady=(0, 14))

        total_card = tk.Frame(total_outer, bg="#1877F2",
                                highlightthickness=0)
        total_card.pack(fill="x")

        total_inner = tk.Frame(total_card, bg="#1877F2")
        total_inner.pack(pady=12)

        tk.Label(total_inner, text="📦", bg="#1877F2",
                 font=("맑은 고딕", 16)).pack(side="left", padx=(0, 8))
        self.lbl_end_total = tk.Label(total_inner, text="총 가용재고: 0개  /  전체 캐파: 0개  (0.0%)",
                                         font=("맑은 고딕", 12, "bold"),
                                         bg="#1877F2", fg="white")
        self.lbl_end_total.pack(side="left")

        # ========== [구역별 진행바 카드 - 메인 영역] ==========
        chart_outer = tk.Frame(container, bg="#F5F6F8")
        chart_outer.pack(fill="both", expand=True, padx=18, pady=(0, 8))

        chart_card = tk.Frame(chart_outer, bg="white",
                                highlightthickness=1, highlightbackground="#E5E7EB")
        chart_card.pack(fill="both", expand=True)

        tk.Frame(chart_card, bg="#8B5CF6", width=4).pack(side="left", fill="y")

        chart_inner = tk.Frame(chart_card, bg="white", padx=18, pady=14)
        chart_inner.pack(side="left", fill="both", expand=True)

        ch_head = tk.Frame(chart_inner, bg="white")
        ch_head.pack(fill="x", pady=(0, 12))
        tk.Label(ch_head, text="📍  구역별 점유율",
                 bg="white", fg="#111827",
                 font=("맑은 고딕", 11, "bold")).pack(side="left")
        tk.Label(ch_head, text="🟢 정상 < 80%   🟡 주의 80%+   🔴 경고 90%+",
                 bg="white", fg="#9CA3AF",
                 font=("맑은 고딕", 8)).pack(side="right")

        # 진행바 영역
        self.zone_chart_frame = tk.Frame(chart_inner, bg="white")
        self.zone_chart_frame.pack(fill="both", expand=True)

        # 초기 안내
        self._render_empty_chart()

        # 호환용 (run_closing_stock_logic이 self.tree_end 참조 - 호환 유지)
        self.tree_end = ttk.Treeview(container, columns=("zone", "qty"),
                                       show="headings", height=1)

        # 드래그앤드롭 등록
        self._setup_dnd_for_closing_stock(chart_card, chart_inner, self.zone_chart_frame)

    def _render_empty_chart(self):
        """차트 영역 빈 상태 안내"""
        for w in self.zone_chart_frame.winfo_children():
            w.destroy()
        empty = tk.Label(self.zone_chart_frame,
                          text="💡 EMP 재고 파일을 [📁 파일 선택] 버튼으로 불러오거나\n"
                               "    이 영역에 직접 끌어다 놓으세요",
                          bg="white", fg="#9CA3AF",
                          font=("맑은 고딕", 10),
                          justify="center")
        empty.pack(expand=True, pady=40)

    def _setup_dnd_for_closing_stock(self, *widgets):
        """마감재고 영역에 드래그앤드롭 등록"""
        try:
            from tkinterdnd2 import DND_FILES
            for w in widgets:
                try:
                    w.drop_target_register(DND_FILES)
                    w.dnd_bind('<<Drop>>', self._on_closing_file_dropped)
                except: pass
        except ImportError:
            pass

    def _on_closing_file_dropped(self, event):
        """마감재고 파일 드롭 처리"""
        path = event.data.strip()
        if path.startswith('{') and path.endswith('}'):
            path = path[1:-1]
        if ' ' in path and not os.path.exists(path):
            for p in path.split():
                p = p.strip('{}')
                if os.path.exists(p):
                    path = p
                    break
        if not os.path.exists(path):
            messagebox.showerror("오류", f"파일을 찾을 수 없습니다:\n{path}")
            return
        # 직접 호출 (다이얼로그 안 띄우고 바로)
        self._process_closing_stock_file(path)

    def _process_closing_stock_file(self, file_path):
        """파일 경로 받아서 마감재고 분석 (드래그/버튼 공용)"""
        try:
            raw_df = pd.read_excel(file_path, header=None)
            header_row = -1
            for i, row in raw_df.iterrows():
                if '창고' in row.values:
                    header_row = i; break
            if header_row == -1:
                raise ValueError("'창고' 행 없음")
            df = raw_df.iloc[header_row+1:].copy()
            df.columns = raw_df.iloc[header_row]
            df.columns = [str(c).strip() for c in df.columns]
            df = df[df['창고'].astype(str).str.contains('정상창고', na=False)]
            df['가용재고'] = pd.to_numeric(df['가용재고'], errors='coerce').fillna(0)
            df['구역코드'] = df['다중로케이션'].astype(str).str[:2]

            def classify_tmp(code):
                c = str(code).upper()
                if 'AA' <= c <= 'AB': return 'AA~AB구역'
                if c == 'BB': return 'BB구역'
                if 'CC' <= c <= 'DD': return 'CC~DD구역'
                if c in ['EA', 'EE', 'FF']: return 'EA/EE/FF구역'
                return '기타'

            df['최종구역'] = df['구역코드'].apply(classify_tmp)
            summary = df[df['최종구역'] != '기타'].groupby('최종구역')['가용재고'].sum().reset_index()

            self._render_zone_chart(summary)
        except Exception as e:
            messagebox.showerror("오류", str(e))

    def _render_zone_chart(self, summary_df):
        """구역별 진행바 차트 그리기 (가독성 강화 버전)"""
        # 기존 위젯 제거
        for w in self.zone_chart_frame.winfo_children():
            w.destroy()

        zone_order = ['AA~AB구역', 'BB구역', 'CC~DD구역', 'EA/EE/FF구역']
        data = {row['최종구역']: int(row['가용재고']) for _, row in summary_df.iterrows()}

        zone_colors = {
            'AA~AB구역': '#3B82F6',
            'BB구역': '#10B981',
            'CC~DD구역': '#F59E0B',
            'EA/EE/FF구역': '#8B5CF6',
        }

        total_qty = 0
        total_capa = 0

        for zone in zone_order:
            qty = data.get(zone, 0)
            capa = self._zone_capacity.get(zone, 0)
            total_qty += qty
            total_capa += capa
            pct = (qty / capa * 100) if capa > 0 else 0

            # 색상 결정
            if pct >= 90:
                bar_color = '#DC2626'
                badge_bg = "#FEE2E2"
                badge_fg = "#991B1B"
                badge_text = "⚠️ 경고"
                qty_color = "#DC2626"
            elif pct >= 80:
                bar_color = '#F59E0B'
                badge_bg = "#FEF3C7"
                badge_fg = "#92400E"
                badge_text = "주의"
                qty_color = "#D97706"
            else:
                bar_color = zone_colors.get(zone, '#3B82F6')
                badge_bg = "#DCFCE7"
                badge_fg = "#166534"
                badge_text = "정상"
                qty_color = "#111827"

            # ===== 한 행 카드 형태 =====
            row_card = tk.Frame(self.zone_chart_frame, bg="#FAFBFC",
                                  highlightthickness=1, highlightbackground="#E5E7EB")
            row_card.pack(fill="x", pady=(0, 8))

            row_inner = tk.Frame(row_card, bg="#FAFBFC", padx=14, pady=10)
            row_inner.pack(fill="x")

            # 1행: [구역명] [수량 크게] [%뱃지]
            top_row = tk.Frame(row_inner, bg="#FAFBFC")
            top_row.pack(fill="x", pady=(0, 6))

            # 좌측: 구역명
            tk.Label(top_row, text=zone,
                     bg="#FAFBFC", fg="#374151",
                     font=("맑은 고딕", 11, "bold")).pack(side="left")

            # 우측: 뱃지
            tk.Label(top_row, text=f"{badge_text}  {pct:.1f}%",
                     bg=badge_bg, fg=badge_fg,
                     font=("맑은 고딕", 9, "bold"),
                     padx=10, pady=3).pack(side="right")

            # 2행: 큰 수량 표시 [재고수량 / 캐파]
            qty_row = tk.Frame(row_inner, bg="#FAFBFC")
            qty_row.pack(fill="x", pady=(0, 8))

            # 가용재고 - 진짜 크게
            tk.Label(qty_row, text=f"{qty:,}",
                     bg="#FAFBFC", fg=qty_color,
                     font=("맑은 고딕", 22, "bold")).pack(side="left")
            tk.Label(qty_row, text="개",
                     bg="#FAFBFC", fg=qty_color,
                     font=("맑은 고딕", 12, "bold")).pack(side="left", padx=(2, 8), pady=(8, 0))
            # 구분선
            tk.Label(qty_row, text="/",
                     bg="#FAFBFC", fg="#9CA3AF",
                     font=("맑은 고딕", 14)).pack(side="left", padx=4, pady=(4, 0))
            # 캐파
            tk.Label(qty_row, text=f"{capa:,}",
                     bg="#FAFBFC", fg="#9CA3AF",
                     font=("맑은 고딕", 14)).pack(side="left", padx=(4, 0), pady=(4, 0))
            tk.Label(qty_row, text="개",
                     bg="#FAFBFC", fg="#9CA3AF",
                     font=("맑은 고딕", 10)).pack(side="left", padx=(1, 0), pady=(8, 0))

            # 3행: 진행바 (배경 회색 frame + 채움 frame, 픽셀 단위로 즉시 표시)
            bar_bg = tk.Frame(row_inner, bg="#E5E7EB", height=10,
                                highlightthickness=0)
            bar_bg.pack(fill="x")
            bar_bg.pack_propagate(False)

            # 채움을 place로 % 비율 (resize 시 자동 조정)
            fill_pct = min(pct, 100) / 100  # 0~1
            if fill_pct > 0:
                bar_fill = tk.Frame(bar_bg, bg=bar_color)
                bar_fill.place(relx=0, rely=0, relwidth=fill_pct, relheight=1)

            # 100% 초과 시 추가 표시
            if pct > 100:
                over_label = tk.Label(bar_bg,
                                        text=f"⚠️ {pct:.0f}% 초과 ⚠️",
                                        bg=bar_color, fg="white",
                                        font=("맑은 고딕", 7, "bold"))
                over_label.place(relx=0.5, rely=0.5, anchor="center")

        # 총합 라벨 업데이트
        if total_capa > 0:
            total_pct = total_qty / total_capa * 100
            self.lbl_end_total.config(
                text=f"총 가용재고: {total_qty:,}개  /  전체 캐파: {total_capa:,}개  ({total_pct:.1f}%)")
        else:
            self.lbl_end_total.config(text=f"총 가용재고: {total_qty:,}개")

    # ========== [캐파 설정 - Firestore 저장/로드] ==========
    def refresh_zone_capacity(self):
        """로컬 config 파일에서 구역별 캐파 로드 (개인 PC 전용)"""
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, "r", encoding="utf-8") as f:
                    cfg = json.load(f)
                saved = cfg.get("zone_capacity", {})
                for k, v in saved.items():
                    try:
                        self._zone_capacity[k] = int(v)
                    except: pass
        except Exception as e:
            print(f"⚠️ 캐파 로드 실패: {e}")

    def save_zone_capacity(self, new_capacity):
        """로컬 config 파일에 구역별 캐파 저장 (개인 PC 전용)"""
        try:
            cfg = {}
            if os.path.exists(self.config_path):
                try:
                    with open(self.config_path, "r", encoding="utf-8") as f:
                        cfg = json.load(f)
                except Exception:
                    cfg = {}
            cfg["zone_capacity"] = {k: int(v) for k, v in new_capacity.items()}
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
            self._zone_capacity = dict(new_capacity)
            return True
        except Exception as e:
            messagebox.showerror("저장 실패", str(e))
            return False

    def open_capacity_settings(self):
        """캐파 설정 팝업"""
        win = tk.Toplevel(self.root)
        win.title("⚙️ 구역별 캐파 설정")
        win.configure(bg="#F5F6F8")
        try:
            self.position_popup(win, 380, 360)
        except Exception:
            win.geometry("380x360")

        # 헤더
        head = tk.Frame(win, bg="#F5F6F8")
        head.pack(fill="x", padx=18, pady=(14, 8))
        tk.Label(head, text="⚙️", font=("맑은 고딕", 18), bg="#F5F6F8").pack(side="left", padx=(0, 6))
        tk.Label(head, text="구역별 캐파 설정",
                 font=("맑은 고딕", 13, "bold"),
                 bg="#F5F6F8", fg="#1A1A1A").pack(side="left")

        tk.Label(win, text="저장하면 모든 PC에 동기화됩니다.",
                 bg="#F5F6F8", fg="#6B7280",
                 font=("맑은 고딕", 8)).pack(padx=18, anchor="w")

        # 입력 카드
        card = tk.Frame(win, bg="white",
                          highlightthickness=1, highlightbackground="#E5E7EB")
        card.pack(fill="both", expand=True, padx=18, pady=12)

        inner = tk.Frame(card, bg="white", padx=18, pady=14)
        inner.pack(fill="both", expand=True)

        zone_order = ['AA~AB구역', 'BB구역', 'CC~DD구역', 'EA/EE/FF구역']
        entries = {}

        for zone in zone_order:
            row = tk.Frame(inner, bg="white")
            row.pack(fill="x", pady=6)
            tk.Label(row, text=zone,
                     bg="white", fg="#374151",
                     font=("맑은 고딕", 10, "bold"),
                     width=14, anchor="w").pack(side="left")
            ent = tk.Entry(row, font=("Consolas", 11),
                            bd=1, relief="solid", justify="right")
            current = self._zone_capacity.get(zone, 0)
            ent.insert(0, str(current))
            ent.pack(side="left", fill="x", expand=True, ipady=3)
            tk.Label(row, text="개", bg="white", fg="#9CA3AF",
                     font=("맑은 고딕", 9)).pack(side="left", padx=(4, 0))
            entries[zone] = ent

        # 저장 버튼
        btn_row = tk.Frame(win, bg="#F5F6F8")
        btn_row.pack(fill="x", padx=18, pady=(0, 14))

        def save_and_close():
            new_cap = {}
            for zone, ent in entries.items():
                try:
                    val = int(ent.get().strip().replace(',', ''))
                    if val < 0:
                        raise ValueError("음수 불가")
                    new_cap[zone] = val
                except ValueError:
                    messagebox.showwarning("입력 오류",
                        f"'{zone}'의 값이 올바르지 않습니다. 정수만 입력 가능합니다.",
                        parent=win)
                    return
            if self.save_zone_capacity(new_cap):
                messagebox.showinfo("저장 완료",
                    "✅ 캐파 설정이 저장되었습니다.\n다른 PC에도 동기화됩니다.",
                    parent=win)
                win.destroy()

        tk.Button(btn_row, text="💾 저장",
                   command=save_and_close,
                   bg="#1877F2", fg="white",
                   font=("맑은 고딕", 10, "bold"),
                   relief="flat", padx=20, pady=8,
                   cursor="hand2").pack(side="right")
        tk.Button(btn_row, text="취소",
                   command=win.destroy,
                   bg="#F3F4F6", fg="#6B7280",
                   font=("맑은 고딕", 10),
                   relief="flat", padx=14, pady=8,
                   cursor="hand2").pack(side="right", padx=(0, 6))

    # --- [탭: 재고파악] ---
    def setup_inventory_check_v95(self):
        """재고파악 - 출고리스트 vs 전체재고 → 출고가능 추출"""
        container = tk.Frame(self.t_chk, bg="#F5F6F8")
        container.pack(fill="both", expand=True)

        # ========== [상단 헤더] ==========
        header_frame = tk.Frame(container, bg="#F5F6F8")
        header_frame.pack(side="top", fill="x", padx=24, pady=(16, 8))

        title_left = tk.Frame(header_frame, bg="#F5F6F8")
        title_left.pack(side="left")
        tk.Label(title_left, text="🔍", font=("맑은 고딕", 22),
                 bg="#F5F6F8").pack(side="left", padx=(0, 6))
        title_text_box = tk.Frame(title_left, bg="#F5F6F8")
        title_text_box.pack(side="left")
        tk.Label(title_text_box, text="출고가능 재고파악",
                 font=("맑은 고딕", 15, "bold"),
                 bg="#F5F6F8", fg="#1A1A1A").pack(anchor="w")
        tk.Label(title_text_box, text="출고리스트 ↔ 전체재고 비교 → 출고가능 상품 추출",
                 font=("맑은 고딕", 8),
                 bg="#F5F6F8", fg="#888").pack(anchor="w")

        # ========== [실행 버튼 - 하단 고정] ==========
        action_outer = tk.Frame(container, bg="#F5F6F8")
        action_outer.pack(side="bottom", fill="x", padx=18, pady=(0, 14))

        shadow = tk.Frame(action_outer, bg="#7C3AED")
        shadow.pack(fill="x")

        run_btn = tk.Button(shadow, text="📋  출고가능 리스트 생성",
                              bg="#8B5CF6", fg="white",
                              activebackground="#7C3AED", activeforeground="white",
                              font=("맑은 고딕", 11, "bold"),
                              relief="flat", bd=0,
                              cursor="hand2",
                              command=self.run_inventory_check_logic)
        run_btn.pack(fill="x", ipady=10)
        def on_e2(e): run_btn.config(bg="#7C3AED")
        def on_l2(e): run_btn.config(bg="#8B5CF6")
        run_btn.bind("<Enter>", on_e2)
        run_btn.bind("<Leave>", on_l2)

        # ========== [파일 선택 카드들] ==========
        # 카드 1: 조회 리스트
        c1_outer = tk.Frame(container, bg="#F5F6F8")
        c1_outer.pack(fill="x", padx=18, pady=(0, 8))

        c1 = tk.Frame(c1_outer, bg="white",
                       highlightthickness=1, highlightbackground="#E5E7EB")
        c1.pack(fill="x")
        tk.Frame(c1, bg="#3B82F6", width=4).pack(side="left", fill="y")

        c1_inner = tk.Frame(c1, bg="white", padx=14, pady=14)
        c1_inner.pack(side="left", fill="both", expand=True)

        tk.Label(c1_inner, text="1️⃣  출고리스트 (조회)",
                 bg="white", fg="#1E40AF",
                 font=("맑은 고딕", 10, "bold")).pack(anchor="w", pady=(0, 6))

        c1_row = tk.Frame(c1_inner, bg="white")
        c1_row.pack(fill="x")
        tk.Button(c1_row, text="📁 파일 선택",
                   command=self.sel_chk_target,
                   bg="#EFF6FF", fg="#1E40AF",
                   font=("맑은 고딕", 9, "bold"),
                   relief="flat", padx=14, pady=6,
                   cursor="hand2").pack(side="left")
        self.lbl_chk_target = tk.Label(c1_row, text="미선택",
                                          fg="#9CA3AF", bg="white",
                                          font=("맑은 고딕", 9))
        self.lbl_chk_target.pack(side="left", padx=10)

        # 카드 2: 재고 마스터
        c2_outer = tk.Frame(container, bg="#F5F6F8")
        c2_outer.pack(fill="x", padx=18, pady=(0, 8))

        c2 = tk.Frame(c2_outer, bg="white",
                       highlightthickness=1, highlightbackground="#E5E7EB")
        c2.pack(fill="x")
        tk.Frame(c2, bg="#10B981", width=4).pack(side="left", fill="y")

        c2_inner = tk.Frame(c2, bg="white", padx=14, pady=14)
        c2_inner.pack(side="left", fill="both", expand=True)

        tk.Label(c2_inner, text="2️⃣  EMP 재고 마스터 (전체재고)",
                 bg="white", fg="#065F46",
                 font=("맑은 고딕", 10, "bold")).pack(anchor="w", pady=(0, 6))

        c2_row = tk.Frame(c2_inner, bg="white")
        c2_row.pack(fill="x")
        tk.Button(c2_row, text="📁 파일 선택",
                   command=self.sel_chk_master,
                   bg="#ECFDF5", fg="#065F46",
                   font=("맑은 고딕", 9, "bold"),
                   relief="flat", padx=14, pady=6,
                   cursor="hand2").pack(side="left")
        self.lbl_chk_master = tk.Label(c2_row, text="미선택",
                                          fg="#9CA3AF", bg="white",
                                          font=("맑은 고딕", 9))
        self.lbl_chk_master.pack(side="left", padx=10)

        # 안내
        info_outer = tk.Frame(container, bg="#F5F6F8")
        info_outer.pack(fill="both", expand=True, padx=18, pady=(0, 8))

        info_card = tk.Frame(info_outer, bg="#FFFBEB",
                              highlightthickness=1, highlightbackground="#FEF3C7")
        info_card.pack(fill="both", expand=True)

        info_inner = tk.Frame(info_card, bg="#FFFBEB", padx=20, pady=20)
        info_inner.pack(fill="both", expand=True)

        tk.Label(info_inner, text="💡",
                 bg="#FFFBEB", font=("맑은 고딕", 24)).pack(pady=(20, 8))
        tk.Label(info_inner, text="동작 방식",
                 bg="#FFFBEB", fg="#92400E",
                 font=("맑은 고딕", 11, "bold")).pack()
        tk.Label(info_inner,
                 text="1️⃣ 출고리스트와 2️⃣ 전체재고를 비교하여\n"
                       "현재 출고 가능한 상품만 추출합니다.\n\n"
                       "결과는 엑셀 파일로 저장됩니다.",
                 bg="#FFFBEB", fg="#78350F",
                 font=("맑은 고딕", 9),
                 justify="center").pack(pady=(6, 20))
    
    # --- [탭 6: 현장소통] ---
    def setup_field_comm(self, container):
        """모던 카드 리스트 스타일 작업보고 화면"""
        container.configure(bg="#F5F6F8")

        # ========== [상단 헤더] ==========
        header_frame = tk.Frame(container, bg="#F5F6F8")
        header_frame.pack(side="top", fill="x", padx=30, pady=(20, 10))

        title_left = tk.Frame(header_frame, bg="#F5F6F8")
        title_left.pack(side="left")
        tk.Label(title_left, text="📋", font=("맑은 고딕", 28),
                 bg="#F5F6F8").pack(side="left", padx=(0, 8))
        title_text_box = tk.Frame(title_left, bg="#F5F6F8")
        title_text_box.pack(side="left")
        tk.Label(title_text_box, text="실시간 작업 보고",
                 font=("맑은 고딕", 18, "bold"),
                 bg="#F5F6F8", fg="#1A1A1A").pack(anchor="w")
        # 새 댓글 카운터 라벨
        self.new_comment_counter = tk.Label(title_text_box, text="",
                                              font=("맑은 고딕", 9, "bold"),
                                              bg="#F5F6F8", fg="#FF4757")
        self.new_comment_counter.pack(anchor="w")

        # 우측 액션 버튼들
        action_box = tk.Frame(header_frame, bg="#F5F6F8")
        action_box.pack(side="right")
        tk.Button(action_box, text="⚙️ 내 이름",
                  command=self.change_user_name,
                  font=("맑은 고딕", 9, "bold"),
                  bg="white", fg="#444",
                  relief="flat", padx=14, pady=8,
                  cursor="hand2",
                  highlightthickness=1,
                  highlightbackground="#DDD").pack(side="right", padx=(8, 0))
        tk.Button(action_box, text="🔄 새로고침",
                  command=self.update_table_view,
                  font=("맑은 고딕", 9, "bold"),
                  bg="#1877F2", fg="white",
                  relief="flat", padx=14, pady=8,
                  cursor="hand2").pack(side="right")

        # ========== [필터 + 검색 바 - 깔끔한 둥근 박스] ==========
        filter_card = tk.Frame(container, bg="white",
                                highlightthickness=1,
                                highlightbackground="#E5E7EB")
        filter_card.pack(side="top", fill="x", padx=30, pady=(0, 10), ipady=4)

        # 필터 칩 (탭처럼 보이는 버튼들)
        self.filter_var.set("⏳ 처리중")
        self.filter_chips = {}
        chip_box = tk.Frame(filter_card, bg="white")
        chip_box.pack(side="left", padx=14, pady=10)

        def make_chip(label, value):
            def click():
                self.filter_var.set(value)
                self._update_chip_styles()
                self.update_table_view()
            chip = tk.Label(chip_box, text=label,
                             font=("맑은 고딕", 9, "bold"),
                             bg="#F0F2F5", fg="#666",
                             padx=14, pady=6, cursor="hand2")
            chip.bind("<Button-1>", lambda e: click())
            chip.pack(side="left", padx=2)
            self.filter_chips[value] = chip

        make_chip("⏳ 처리중", "⏳ 처리중")
        make_chip("✅ 완료", "✅ 완료")

        # 구분선
        tk.Frame(filter_card, bg="#E5E7EB", width=1).pack(side="left", fill="y", padx=8, pady=14)

        # 검색
        search_box = tk.Frame(filter_card, bg="white")
        search_box.pack(side="left", padx=10, pady=10, fill="x", expand=True)
        tk.Label(search_box, text="🔍", font=("맑은 고딕", 11),
                 bg="white", fg="#999").pack(side="left", padx=(0, 4))
        search_entry = tk.Entry(search_box, textvariable=self.search_var,
                                 font=("맑은 고딕", 10),
                                 bd=0, relief="flat",
                                 highlightthickness=0,
                                 bg="white", fg="#333")
        search_entry.pack(side="left", fill="x", expand=True, ipady=4)
        search_entry.bind("<Return>", lambda e: self.update_table_view())
        # placeholder 흉내
        def on_focus_in(e):
            if search_entry.get() == "제목·작업자·날짜·내용 검색":
                search_entry.delete(0, tk.END)
                search_entry.config(fg="#333")
        def on_focus_out(e):
            if not search_entry.get():
                search_entry.insert(0, "제목·작업자·날짜·내용 검색")
                search_entry.config(fg="#999")
        if not self.search_var.get():
            search_entry.insert(0, "제목·작업자·날짜·내용 검색")
            search_entry.config(fg="#999")
        search_entry.bind("<FocusIn>", on_focus_in)
        search_entry.bind("<FocusOut>", on_focus_out)

        self._update_chip_styles()

        # ========== [카드 리스트 영역 - 스크롤] ==========
        list_outer = tk.Frame(container, bg="#F5F6F8")
        list_outer.pack(side="top", expand=True, fill="both", padx=30, pady=(0, 15))

        list_canvas = tk.Canvas(list_outer, bg="#F5F6F8",
                                 highlightthickness=0)
        list_scrollbar = ttk.Scrollbar(list_outer, orient="vertical",
                                        command=list_canvas.yview)
        self.cards_frame = tk.Frame(list_canvas, bg="#F5F6F8")

        def _update_field_scrollregion(e):
            list_canvas.configure(scrollregion=list_canvas.bbox("all"))
            # [수정] 카드 새로 그릴 때 항상 맨 위로
            list_canvas.yview_moveto(0)

        self.cards_frame.bind("<Configure>", _update_field_scrollregion)
        canvas_window = list_canvas.create_window((0, 0), window=self.cards_frame, anchor="nw")
        # 캔버스 크기에 맞춰 cards_frame 너비 조절
        def _on_canvas_resize(event):
            list_canvas.itemconfig(canvas_window, width=event.width)
        list_canvas.bind("<Configure>", _on_canvas_resize)
        list_canvas.configure(yscrollcommand=list_scrollbar.set)

        list_scrollbar.pack(side="right", fill="y")
        list_canvas.pack(side="left", fill="both", expand=True)

        def _on_mousewheel(event):
            try:
                list_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            except: pass
        list_canvas.bind_all("<MouseWheel>", _on_mousewheel)
        # 윈도우가 다른 탭으로 전환 시 바인딩 안 풀리는 문제 방지: 마우스가 영역 안에 있을 때만 작동하도록
        def _bind_wheel(e): list_canvas.bind_all("<MouseWheel>", _on_mousewheel)
        def _unbind_wheel(e): list_canvas.unbind_all("<MouseWheel>")
        list_canvas.bind("<Enter>", _bind_wheel)
        list_canvas.bind("<Leave>", _unbind_wheel)

        # 우클릭 메뉴 (삭제용)
        self.context_menu = tk.Menu(self.root, tearoff=0, font=("맑은 고딕", 10))
        self.context_menu.add_command(label="🗑️ 이 보고서 완전 삭제", command=self.delete_report)

        # 첫 로드
        self._selected_card_id = None
        self.update_table_view()

    def _update_chip_styles(self):
        """필터 칩 선택 상태에 맞춰 스타일 업데이트"""
        if not hasattr(self, 'filter_chips'): return
        current = self.filter_var.get()
        for value, chip in self.filter_chips.items():
            if value == current:
                chip.config(bg="#1877F2", fg="white")
            else:
                chip.config(bg="#F0F2F5", fg="#666")

    def setup_board_system(self, parent):
        """모던 카드 리스트 스타일 공지/소통 화면"""
        parent.configure(bg="#F5F6F8")

        # ========== [상단 헤더] ==========
        header_frame = tk.Frame(parent, bg="#F5F6F8")
        header_frame.pack(side="top", fill="x", padx=30, pady=(20, 10))

        title_left = tk.Frame(header_frame, bg="#F5F6F8")
        title_left.pack(side="left")
        tk.Label(title_left, text="📢", font=("맑은 고딕", 28),
                 bg="#F5F6F8").pack(side="left", padx=(0, 8))
        title_text_box = tk.Frame(title_left, bg="#F5F6F8")
        title_text_box.pack(side="left")
        tk.Label(title_text_box, text="공지 및 작업자 소통",
                 font=("맑은 고딕", 18, "bold"),
                 bg="#F5F6F8", fg="#1A1A1A").pack(anchor="w")
        # 새 메시지 카운터
        self.board_new_counter = tk.Label(title_text_box, text="",
                                            font=("맑은 고딕", 9, "bold"),
                                            bg="#F5F6F8", fg="#FF4757")
        self.board_new_counter.pack(anchor="w")

        # 우측 액션 버튼들
        action_box = tk.Frame(header_frame, bg="#F5F6F8")
        action_box.pack(side="right")
        tk.Button(action_box, text="⚙️ 내 이름",
                  command=self.change_user_name,
                  font=("맑은 고딕", 9, "bold"),
                  bg="white", fg="#444",
                  relief="flat", padx=14, pady=8,
                  cursor="hand2",
                  highlightthickness=1,
                  highlightbackground="#DDD").pack(side="right", padx=(8, 0))
        tk.Button(action_box, text="🔄 새로고침",
                  command=self.update_board_view,
                  font=("맑은 고딕", 9, "bold"),
                  bg="#34a853", fg="white",
                  relief="flat", padx=14, pady=8,
                  cursor="hand2").pack(side="right", padx=(8, 0))
        tk.Button(action_box, text="✏️ 공지 작성",
                  command=self.send_global_notice,
                  font=("맑은 고딕", 9, "bold"),
                  bg="#1877F2", fg="white",
                  relief="flat", padx=14, pady=8,
                  cursor="hand2").pack(side="right")

        # ========== [필터 + 검색 바] ==========
        filter_card = tk.Frame(parent, bg="white",
                                highlightthickness=1,
                                highlightbackground="#E5E7EB")
        filter_card.pack(side="top", fill="x", padx=30, pady=(0, 10), ipady=4)

        # 필터 칩
        if not hasattr(self, 'board_filter_var'):
            self.board_filter_var = tk.StringVar(value="📌 미완료")
        else:
            self.board_filter_var.set("📌 미완료")

        self.board_filter_chips = {}
        chip_box = tk.Frame(filter_card, bg="white")
        chip_box.pack(side="left", padx=14, pady=10)

        def make_board_chip(label, value):
            def click():
                self.board_filter_var.set(value)
                self._update_board_chip_styles()
                self.update_board_view()
            chip = tk.Label(chip_box, text=label,
                             font=("맑은 고딕", 9, "bold"),
                             bg="#F0F2F5", fg="#666",
                             padx=14, pady=6, cursor="hand2")
            chip.bind("<Button-1>", lambda e: click())
            chip.pack(side="left", padx=2)
            self.board_filter_chips[value] = chip

        make_board_chip("📌 미완료", "📌 미완료")
        make_board_chip("✅ 완료", "✅ 완료")

        # 구분선
        tk.Frame(filter_card, bg="#E5E7EB", width=1).pack(side="left", fill="y", padx=8, pady=14)

        # 검색
        if not hasattr(self, 'board_search_var'):
            self.board_search_var = tk.StringVar()

        search_box = tk.Frame(filter_card, bg="white")
        search_box.pack(side="left", padx=10, pady=10, fill="x", expand=True)
        tk.Label(search_box, text="🔍", font=("맑은 고딕", 11),
                 bg="white", fg="#999").pack(side="left", padx=(0, 4))
        board_search_entry = tk.Entry(search_box, textvariable=self.board_search_var,
                                       font=("맑은 고딕", 10),
                                       bd=0, relief="flat",
                                       highlightthickness=0,
                                       bg="white", fg="#333")
        board_search_entry.pack(side="left", fill="x", expand=True, ipady=4)
        board_search_entry.bind("<Return>", lambda e: self.update_board_view())

        def board_focus_in(e):
            if board_search_entry.get() == "내용·작업자·날짜·구분 검색":
                board_search_entry.delete(0, tk.END)
                board_search_entry.config(fg="#333")
        def board_focus_out(e):
            if not board_search_entry.get():
                board_search_entry.insert(0, "내용·작업자·날짜·구분 검색")
                board_search_entry.config(fg="#999")
        if not self.board_search_var.get():
            board_search_entry.insert(0, "내용·작업자·날짜·구분 검색")
            board_search_entry.config(fg="#999")
        board_search_entry.bind("<FocusIn>", board_focus_in)
        board_search_entry.bind("<FocusOut>", board_focus_out)

        self._update_board_chip_styles()

        # ========== [카드 리스트 영역] ==========
        list_outer = tk.Frame(parent, bg="#F5F6F8")
        list_outer.pack(side="top", expand=True, fill="both", padx=30, pady=(0, 15))

        list_canvas = tk.Canvas(list_outer, bg="#F5F6F8", highlightthickness=0)
        list_scrollbar = ttk.Scrollbar(list_outer, orient="vertical",
                                        command=list_canvas.yview)
        self.board_cards_frame = tk.Frame(list_canvas, bg="#F5F6F8")

        def _update_scrollregion(e):
            list_canvas.configure(scrollregion=list_canvas.bbox("all"))
            # [수정] 카드 새로 그릴 때 항상 맨 위로
            list_canvas.yview_moveto(0)

        self.board_cards_frame.bind("<Configure>", _update_scrollregion)
        canvas_window = list_canvas.create_window((0, 0), window=self.board_cards_frame, anchor="nw")
        def _on_canvas_resize(event):
            list_canvas.itemconfig(canvas_window, width=event.width)
        list_canvas.bind("<Configure>", _on_canvas_resize)
        list_canvas.configure(yscrollcommand=list_scrollbar.set)

        list_scrollbar.pack(side="right", fill="y")
        list_canvas.pack(side="left", fill="both", expand=True)

        def _on_mousewheel(event):
            try:
                list_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            except: pass
        def _bind_wheel(e): list_canvas.bind_all("<MouseWheel>", _on_mousewheel)
        def _unbind_wheel(e): list_canvas.unbind_all("<MouseWheel>")
        list_canvas.bind("<Enter>", _bind_wheel)
        list_canvas.bind("<Leave>", _unbind_wheel)

        # 우클릭 메뉴
        self.board_context_menu = tk.Menu(self.root, tearoff=0, font=("맑은 고딕", 10))
        self.board_context_menu.add_command(label="✅ 완료 처리", command=self.complete_board_post)
        self.board_context_menu.add_separator()
        self.board_context_menu.add_command(label="🗑️ 완전 삭제", command=self.delete_board_post)

        self._selected_board_id = None

        # 실시간 리스너 + 첫 로드
        self.start_board_listener()
        self.update_board_view()

    def _update_board_chip_styles(self):
        """공지/소통 필터 칩 선택 상태 업데이트"""
        if not hasattr(self, 'board_filter_chips'): return
        current = self.board_filter_var.get()
        for value, chip in self.board_filter_chips.items():
            if value == current:
                chip.config(bg="#1877F2", fg="white")
            else:
                chip.config(bg="#F0F2F5", fg="#666")

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
                            'category': "공지" if target == "all" else "요청",
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
        # [수정] _direct_board_id (카드 클릭) 우선, 없으면 옛 board_tree fallback
        item_id = getattr(self, '_direct_board_id', None)
        self._direct_board_id = None  # 한번 쓰고 비움

        if not item_id:
            # 호환: 옛 board_tree 인터페이스
            if hasattr(self, 'board_tree') and self.board_tree.winfo_exists():
                selection = self.board_tree.selection()
                if not selection:
                    return
                item_id = selection[0]
            else:
                return

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

        # 요청/문의는 대화 스레드 창 (post_data만으로 호출)
        self._open_thread_window(item_id, post_data, None)

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

    def _open_thread_window(self, item_id, post_data, row_values=None):
        """요청/문의 대화창 - 작업보고와 동일한 카톡 스타일"""
        import threading
        from tkinter import messagebox

        doc_ref = self.db.collection('board_posts').document(item_id)
        category = post_data.get('category', '')
        report_ts = post_data.get('timestamp')
        # 원글 작성자 (real_sender 우선, 없으면 user)
        original_writer = post_data.get('real_sender') or post_data.get('user', '익명')
        # '지시'면 관리자가 보낸 거 = is_admin_origin True
        is_admin_origin = ('요청' in category) or ('지시' in category)

        # 팝업 창
        detail_win = tk.Toplevel(self.root)
        title_short = post_data.get('text', '')[:20]
        detail_win.title(f"💬 {original_writer} - {title_short}")
        screen_h = self.root.winfo_screenheight()
        win_h = min(900, int(screen_h * 0.85))
        self.position_popup(detail_win, 580, win_h)
        detail_win.configure(bg="#B2C7DA")

        # ========== [상단 헤더] ==========
        header = tk.Frame(detail_win, bg="white", height=70,
                          highlightthickness=0, bd=0)
        header.pack(side="top", fill="x")
        header.pack_propagate(False)

        # 아이콘
        icon_color = "#8B5CF6" if is_admin_origin else "#3B82F6"
        icon_emoji = "🔒" if is_admin_origin else "💬"
        avatar_frame = tk.Frame(header, bg=icon_color, width=44, height=44)
        avatar_frame.place(x=15, y=13)
        avatar_frame.pack_propagate(False)
        tk.Label(avatar_frame, text=icon_emoji, bg=icon_color, fg="white",
                 font=("맑은 고딕", 16)).pack(expand=True)

        # 제목 + 부제
        title_frame = tk.Frame(header, bg="white")
        title_frame.place(x=70, y=10)

        cat_label = "🔒 요청" if is_admin_origin else "💬 문의"
        if is_admin_origin and post_data.get('user'):
            cat_label += f" → {post_data.get('user')}"

        # 제목 = 카테고리 표시
        title_text_widget = tk.Text(title_frame, height=1, bd=0,
                                     font=("맑은 고딕", 13, "bold"),
                                     bg="white", fg="#262626",
                                     wrap="none", cursor="xterm",
                                     highlightthickness=0, width=40)
        title_text_widget.insert("1.0", cat_label)
        def _readonly_handler(event):
            if event.state & 0x0004:
                return None
            if event.keysym in ('Left', 'Right', 'Up', 'Down', 'Home', 'End', 'Prior', 'Next'):
                return None
            return "break"
        title_text_widget.bind("<Key>", _readonly_handler)
        title_text_widget.pack(anchor="w")

        sub_text = f"작성: {original_writer}  ·  {post_data.get('status', '신규')}"
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

        def _safe_scroll(event):
            try:
                if canvas.winfo_exists():
                    canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            except: pass
        detail_win.bind("<MouseWheel>", _safe_scroll)
        canvas.bind("<MouseWheel>", _safe_scroll)
        scrollable_frame.bind("<MouseWheel>", _safe_scroll)

        # ========== [본문 - 카톡 말풍선] ==========
        body_date = self._get_kst_date(report_ts)
        if body_date:
            sep_frame = tk.Frame(scrollable_frame, bg="#B2C7DA")
            sep_frame.pack(fill="x", pady=(15, 5))
            tk.Label(sep_frame,
                     text=self._format_date_separator(body_date),
                     bg="#9DB0C2", fg="white",
                     font=("맑은 고딕", 8, "bold"),
                     padx=12, pady=3).pack(anchor="center")

        body_outer = tk.Frame(scrollable_frame, bg="#B2C7DA")
        body_outer.pack(fill="x", pady=(5, 5), padx=15)

        # 본문 작성자 표시
        body_align_anchor = "e" if is_admin_origin else "w"
        tk.Label(body_outer, text=f"{'📢' if is_admin_origin else '👤'} {original_writer}",
                 font=("맑은 고딕", 9, "bold"), bg="#B2C7DA", fg="#444",
                 anchor=body_align_anchor).pack(anchor=body_align_anchor, padx=(8, 8))

        # 본문 말풍선
        body_bubble_outer = tk.Frame(body_outer, bg="#B2C7DA")
        body_bubble_outer.pack(fill="x", anchor=body_align_anchor)

        body_color = "#FFEB33" if is_admin_origin else "white"

        body_bubble = tk.Frame(body_bubble_outer, bg=body_color, highlightthickness=0)
        if is_admin_origin:
            body_bubble.pack(side="right", anchor="e", padx=(60, 0))
        else:
            body_bubble.pack(side="left", anchor="w", padx=(0, 60))

        body_text = post_data.get('text', '')
        # [수정] width/height 정확한 계산 (한글 한 글자 = 2칸)
        body_width, text_height = self._calc_text_size(body_text, max_width=42)

        body_label = tk.Text(body_bubble, font=("맑은 고딕", 11),
                              bg=body_color, fg="#1A1A1A",
                              wrap="word", bd=0, padx=14, pady=10,
                              height=text_height, width=body_width,
                              highlightthickness=0, cursor="xterm")
        body_label.insert("1.0", body_text)
        def _body_readonly(event):
            if event.state & 0x0004:
                return None
            if event.keysym in ('Left', 'Right', 'Up', 'Down', 'Home', 'End', 'Prior', 'Next'):
                return None
            return "break"
        body_label.bind("<Key>", _body_readonly)
        body_label.pack(anchor=body_align_anchor)

        # 시간
        if report_ts:
            tt = self._format_kst_time_only(report_ts)
            if tt:
                tk.Label(body_bubble_outer, text=tt,
                         bg="#B2C7DA", fg="#666",
                         font=("맑은 고딕", 8)).pack(side="right" if is_admin_origin else "left",
                                                  anchor="s",
                                                  padx=(4, 0) if not is_admin_origin else (0, 4),
                                                  pady=(0, 4))

        # ========== [댓글(messages) 영역] ==========
        divider_frame = tk.Frame(scrollable_frame, bg="#B2C7DA")
        divider_frame.pack(fill="x", pady=(15, 5), padx=15)
        tk.Frame(divider_frame, bg="#9DB0C2", height=1).pack(fill="x", pady=8)
        tk.Label(divider_frame, text="💬 대화", bg="#B2C7DA", fg="#444",
                 font=("맑은 고딕", 9, "bold")).pack(anchor="center")

        comment_list_frame = tk.Frame(scrollable_frame, bg="#B2C7DA")
        comment_list_frame.pack(fill="x", padx=10, pady=(5, 15))

        # 댓글 영역 우클릭 메뉴 (수정/삭제)
        msg_context_menu = tk.Menu(detail_win, tearoff=0, font=("맑은 고딕", 10))

        if not hasattr(self, '_inline_img_cache_board'):
            self._inline_img_cache_board = []

        def delete_message(msg_id):
            if messagebox.askyesno("삭제", "이 메시지를 삭제하시겠습니까?", parent=detail_win):
                doc_ref.collection('messages').document(msg_id).delete()
                refresh_messages()

        def edit_message(msg_id, current_text):
            from tkinter import simpledialog
            new_text = simpledialog.askstring("메시지 수정", "내용을 수정하세요:",
                                                initialvalue=current_text, parent=detail_win)
            if new_text is not None and new_text.strip():
                doc_ref.collection('messages').document(msg_id).update({
                    'text': new_text.strip(),
                    'edited': True,
                    'editedAt': firestore.SERVER_TIMESTAMP
                })
                refresh_messages()

        def show_msg_menu(event, mid, current_text):
            msg_context_menu.delete(0, tk.END)
            msg_context_menu.add_command(label="✏️ 수정",
                                          command=lambda: edit_message(mid, current_text))
            msg_context_menu.add_command(label="🗑️ 삭제",
                                          command=lambda: delete_message(mid))
            msg_context_menu.post(event.x_root, event.y_root)

        def refresh_messages():
            if not detail_win.winfo_exists(): return
            for w in comment_list_frame.winfo_children(): w.destroy()
            self._inline_img_cache_board.clear()

            try:
                msgs = list(doc_ref.collection('messages').order_by('timestamp').get())
            except Exception as e:
                tk.Label(comment_list_frame, text=f"⚠️ 메시지 로딩 실패: {e}",
                         bg="#B2C7DA", fg="#c00",
                         font=("맑은 고딕", 9)).pack(pady=10)
                return

            if not msgs:
                tk.Label(comment_list_frame, text="아직 답장이 없습니다",
                         bg="#B2C7DA", fg="#666",
                         font=("맑은 고딕", 9, "italic")).pack(pady=15)
                return

            last_date = body_date

            for m in msgs:
                if not detail_win.winfo_exists(): return
                mid = m.id
                md = m.to_dict()
                role = md.get('role', 'worker')
                sender = md.get('sender', '?')
                msg_text = md.get('text', '')
                msg_img = md.get('imageUrl', '')
                edited = md.get('edited', False)
                mts = md.get('timestamp')

                cur_date = self._get_kst_date(mts)
                if cur_date and cur_date != last_date:
                    sep_frame = tk.Frame(comment_list_frame, bg="#B2C7DA")
                    sep_frame.pack(fill="x", pady=(10, 5))
                    tk.Label(sep_frame,
                             text=self._format_date_separator(cur_date),
                             bg="#9DB0C2", fg="white",
                             font=("맑은 고딕", 8, "bold"),
                             padx=12, pady=3).pack(anchor="center")
                    last_date = cur_date

                is_admin = (role == 'admin')
                clean_text = msg_text
                if edited:
                    clean_text += "  (수정됨)"

                time_str = self._format_kst_time_only(mts)

                if is_admin:
                    bubble_color = "#FFEB33"
                    text_anchor = "e"
                else:
                    bubble_color = "white"
                    text_anchor = "w"

                row = tk.Frame(comment_list_frame, bg="#B2C7DA")
                row.pack(fill="x", pady=4, padx=5)
                row.bind("<Button-3>", lambda e, id=mid, txt=msg_text: show_msg_menu(e, id, txt))

                # 이름 표시
                name_label = tk.Label(row, text=f"{'📢' if is_admin else '👤'} {sender}",
                                       bg="#B2C7DA", fg="#444",
                                       font=("맑은 고딕", 8, "bold"))
                name_label.pack(anchor=text_anchor, padx=(8 if is_admin else 2, 8 if is_admin else 0))
                name_label.bind("<Button-3>", lambda e, id=mid, txt=msg_text: show_msg_menu(e, id, txt))

                bubble_row = tk.Frame(row, bg="#B2C7DA")
                bubble_row.pack(fill="x", anchor=text_anchor)
                bubble_row.bind("<Button-3>", lambda e, id=mid, txt=msg_text: show_msg_menu(e, id, txt))

                bubble_container = tk.Frame(bubble_row, bg="#B2C7DA")
                if is_admin:
                    bubble_container.pack(side="right", padx=(60, 0))
                else:
                    bubble_container.pack(side="left", padx=(0, 60))

                bubble = tk.Frame(bubble_container, bg=bubble_color)
                if is_admin:
                    bubble.pack(side="right")
                else:
                    bubble.pack(side="left")
                bubble.bind("<Button-3>", lambda e, id=mid, txt=msg_text: show_msg_menu(e, id, txt))

                if clean_text and clean_text not in ("(사진)", ""):
                    # [수정] 정확한 width/height 계산 사용
                    cmt_width, cmt_height = self._calc_text_size(clean_text, max_width=38)

                    text_label = tk.Text(bubble, font=("맑은 고딕", 11),
                                          bg=bubble_color, fg="#1A1A1A",
                                          wrap="word", bd=0, padx=12, pady=8,
                                          height=cmt_height, width=cmt_width,
                                          highlightthickness=0, cursor="xterm")
                    text_label.insert("1.0", clean_text)
                    def _ro(event):
                        if event.state & 0x0004: return None
                        if event.keysym in ('Left', 'Right', 'Up', 'Down', 'Home', 'End'):
                            return None
                        return "break"
                    text_label.bind("<Key>", _ro)
                    text_label.bind("<Button-3>", lambda e, id=mid, txt=msg_text: show_msg_menu(e, id, txt))
                    text_label.pack(anchor=text_anchor)

                # 첨부 이미지
                if msg_img:
                    try:
                        import requests as _rq
                        from io import BytesIO
                        from PIL import Image, ImageTk
                        resp = _rq.get(msg_img, timeout=10)
                        pil_img = Image.open(BytesIO(resp.content))
                        w, h = pil_img.size
                        max_w = 280
                        if w > max_w:
                            ratio = max_w / w
                            pil_img = pil_img.resize((max_w, int(h * ratio)), Image.LANCZOS)
                        tk_img = ImageTk.PhotoImage(pil_img)
                        self._inline_img_cache_board.append(tk_img)
                        img_holder = tk.Frame(bubble, bg=bubble_color, padx=4, pady=4)
                        img_holder.pack()
                        img_lbl = tk.Label(img_holder, image=tk_img, bg=bubble_color, cursor="hand2")
                        img_lbl.pack()
                        img_lbl.bind("<Button-1>", lambda e, url=msg_img: __import__('webbrowser').open(url))
                    except Exception as ex:
                        tk.Label(bubble, text=f"[사진 로드 실패]",
                                 bg=bubble_color, fg="#999",
                                 font=("맑은 고딕", 8), padx=12, pady=4).pack()

                if time_str:
                    time_label = tk.Label(bubble_container, text=time_str,
                                           bg="#B2C7DA", fg="#555",
                                           font=("맑은 고딕", 8))
                    if is_admin:
                        time_label.pack(side="right", anchor="s", padx=(0, 4), pady=(0, 2),
                                        before=bubble)
                    else:
                        time_label.pack(side="left", anchor="s", padx=(4, 0), pady=(0, 2))

        refresh_messages()

        # ========== [입력창 - 작업보고와 동일] ==========
        tk.Label(footer, text="💡 Ctrl+V로 사진 붙여넣기 · Enter로 전송 · Shift+Enter로 줄바꿈",
                 bg="white", fg="#888", font=("맑은 고딕", 8), anchor="w").pack(fill="x", pady=(0, 2))

        entry_frame = tk.Frame(footer, bg="white", highlightthickness=1, highlightbackground="#CCC")
        entry_frame.pack(fill="x", pady=(0, 5))
        entry = tk.Text(entry_frame, font=("맑은 고딕", 12), bd=0, relief="flat",
                         height=3, padx=10, pady=8, wrap="word")
        entry.pack(fill="x", expand=True)

        _orig_get = entry.get
        _orig_delete = entry.delete
        def _entry_get(*args, **kwargs):
            if not args:
                return _orig_get("1.0", "end-1c")
            return _orig_get(*args, **kwargs)
        def _entry_delete(*args, **kwargs):
            if len(args) >= 1 and args[0] == 0:
                return _orig_delete("1.0", "end")
            return _orig_delete(*args, **kwargs)
        entry.get = _entry_get
        entry.delete = _entry_delete

        # 사진 첨부 시스템 (작업보고와 동일)
        self._board_pending_image_path = None
        self._board_pending_thumb_img = None

        photo_status_frame = tk.Frame(footer, bg="#FFF9E6", relief="flat",
                                       highlightthickness=1, highlightbackground="#FFD966")
        photo_status_frame.pack(fill="x", pady=(0, 5))
        photo_status_frame.pack_forget()

        thumb_label = tk.Label(photo_status_frame, bg="#FFF9E6")
        thumb_label.pack(side="left", padx=8, pady=8)

        info_label = tk.Label(photo_status_frame, text="", bg="#FFF9E6",
                              fg="#444", font=("맑은 고딕", 10, "bold"))
        info_label.pack(side="left", padx=(0, 10))

        def clear_pending_image():
            self._board_pending_image_path = None
            self._board_pending_thumb_img = None
            thumb_label.config(image="")
            info_label.config(text="")
            photo_status_frame.pack_forget()

        cancel_btn = tk.Button(photo_status_frame, text="✕ 취소",
                                command=clear_pending_image, bg="#FF6B6B", fg="white",
                                font=("맑은 고딕", 9, "bold"), relief="flat", padx=10, pady=5,
                                cursor="hand2")
        cancel_btn.pack(side="right", padx=8, pady=8)

        def show_thumbnail(image_source):
            try:
                from PIL import Image, ImageTk
                if isinstance(image_source, str):
                    pil_img = Image.open(image_source)
                else:
                    pil_img = image_source
                pil_thumb = pil_img.copy()
                pil_thumb.thumbnail((60, 60), Image.LANCZOS)
                self._board_pending_thumb_img = ImageTk.PhotoImage(pil_thumb)
                thumb_label.config(image=self._board_pending_thumb_img)
                w, h = pil_img.size
                info_label.config(text=f"📎 첨부됨 ({w}×{h})\n전송 시 자동 업로드")
                photo_status_frame.pack_forget()
                photo_status_frame.pack(fill="x", pady=(0, 5))
            except Exception as e:
                info_label.config(text=f"썸네일 생성 실패: {e}")

        def attach_image_from_clipboard():
            try:
                from PIL import ImageGrab
                import tempfile
                img = ImageGrab.grabclipboard()
                if img is None:
                    messagebox.showinfo("알림", "클립보드에 이미지가 없습니다.", parent=detail_win)
                    return False
                if isinstance(img, list):
                    if not img: return False
                    self._board_pending_image_path = img[0]
                    show_thumbnail(img[0])
                else:
                    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                    img.save(tmp.name, "PNG")
                    self._board_pending_image_path = tmp.name
                    show_thumbnail(img)
                return True
            except Exception as e:
                messagebox.showerror("오류", f"클립보드 이미지 가져오기 실패:\n{e}", parent=detail_win)
                return False

        def attach_image_from_file():
            from tkinter import filedialog
            path = filedialog.askopenfilename(
                title="첨부할 이미지 선택",
                filetypes=[("이미지 파일", "*.png *.jpg *.jpeg *.gif *.bmp"), ("모든 파일", "*.*")],
                parent=detail_win
            )
            if path:
                self._board_pending_image_path = path
                show_thumbnail(path)

        def on_paste(event):
            try:
                from PIL import ImageGrab
                img = ImageGrab.grabclipboard()
                if img is not None and not isinstance(img, list):
                    attach_image_from_clipboard()
                    return "break"
            except: pass
            return None
        entry.bind("<Control-v>", on_paste)

        btn_container = tk.Frame(footer, bg="white")
        btn_container.pack(fill="x")

        def do_send_reply():
            content = entry.get().strip()
            image_url = ""
            if self._board_pending_image_path:
                image_url = self._upload_image_to_imgbb(self._board_pending_image_path, detail_win)
                if image_url is None:
                    return
            if not content and not image_url:
                return
            try:
                my_name = getattr(self, 'current_user', '관리자')
                doc_ref.collection('messages').add({
                    'sender': my_name,
                    'role': 'admin',
                    'text': content if content else "(사진)",
                    'imageUrl': image_url,
                    'timestamp': firestore.SERVER_TIMESTAMP
                })
                # 상태 업데이트
                current_status = post_data.get('status', '')
                if '확인완료' not in current_status and '완료' not in current_status:
                    doc_ref.update({
                        'status': '💬 대화중',
                        'last_reply_time': firestore.SERVER_TIMESTAMP
                    })
                # 푸시 알림
                target = post_data.get('user', '')
                if target and target != 'all':
                    preview = (content if content else "(사진)")[:80]
                    self.send_fcm_push(target,
                                        f"💬 {my_name}님의 답장",
                                        preview)
                entry.delete(0, tk.END)
                clear_pending_image()
                refresh_messages()
            except Exception as e:
                messagebox.showerror("전송 실패", f"{e}", parent=detail_win)

        def do_close_thread():
            pending = entry.get().strip()
            msg = "이 대화를 '확인완료' 처리하시겠습니까?\n(미완료 탭에서 사라집니다)"
            if pending:
                msg = "입력창에 남은 내용을 마지막 답장으로 보낸 뒤\n'확인완료' 처리하시겠습니까?"
            if not messagebox.askyesno("확인", msg, parent=detail_win):
                return
            try:
                if pending:
                    my_name = getattr(self, 'current_user', '관리자')
                    doc_ref.collection('messages').add({
                        'sender': my_name,
                        'role': 'admin',
                        'text': pending,
                        'timestamp': firestore.SERVER_TIMESTAMP
                    })
                doc_ref.update({
                    'status': '✅ 확인완료',
                    'closed_time': firestore.SERVER_TIMESTAMP
                })
                detail_win.destroy()
                self.update_board_view()
            except Exception as e:
                messagebox.showerror("오류", f"처리 실패: {e}", parent=detail_win)

        # Enter로 전송
        def _on_enter(event):
            if event.state & 0x0001:  # Shift
                return None
            do_send_reply()
            return "break"
        entry.bind("<Return>", _on_enter)

        # 버튼들
        tk.Button(btn_container, text="📎 사진(파일)", command=attach_image_from_file,
                  bg="#f0f2f5", fg="#444", font=("맑은 고딕", 9),
                  relief="flat", pady=8).pack(side="left", padx=(0, 5))
        tk.Button(btn_container, text="🚀 답장 전송", command=do_send_reply,
                  bg="#1877F2", fg="white", font=("맑은 고딕", 10, "bold"),
                  relief="flat", pady=10).pack(side="left", expand=True, fill="x", padx=(0, 5))
        tk.Button(btn_container, text="✅ 확인완료", command=do_close_thread,
                  bg="#4CAF50", fg="white", font=("맑은 고딕", 10, "bold"),
                  relief="flat", pady=10).pack(side="left", expand=True, fill="x", padx=(5, 0))

        # 맨 아래 스크롤
        def _scroll_to_bottom():
            try:
                canvas.update_idletasks()
                canvas.yview_moveto(1.0)
            except: pass
        detail_win.after(80, _scroll_to_bottom)

    # --- [기능 2: 소통 글 리스트 업데이트 및 색상 적용] ---
    def update_board_view(self):
        from datetime import datetime, timedelta, timezone

        # 카드 영역 없으면 무시
        if not hasattr(self, 'board_cards_frame') or not self.board_cards_frame.winfo_exists():
            return

        # 1. 기존 카드 제거
        for w in self.board_cards_frame.winfo_children():
            w.destroy()

        try:
            search_keyword = self.board_search_var.get().strip().lower() if hasattr(self, 'board_search_var') else ""
            if search_keyword == "내용·작업자·날짜·구분 검색":
                search_keyword = ""
            current_filter = self.board_filter_var.get() if hasattr(self, 'board_filter_var') else "📌 미완료"

            ref = self.db.collection('board_posts')

            # [최적화] 필터별 다른 쿼리
            try:
                if search_keyword:
                    posts = list(ref.order_by('timestamp', direction=firestore.Query.DESCENDING).limit(500).get())
                elif current_filter == "📌 미완료":
                    # 최근 200건 (대부분 미완료가 적음)
                    posts = list(ref.order_by('timestamp', direction=firestore.Query.DESCENDING).limit(200).get())
                else:
                    # 완료: 최근 7일치만
                    seven_days_ago = datetime.now(timezone.utc) - timedelta(days=7)
                    posts = list(ref.where('timestamp', '>=', seven_days_ago)
                                    .order_by('timestamp', direction=firestore.Query.DESCENDING).limit(200).get())
            except Exception as fetch_err:
                print(f"❌ Firestore 조회 실패: {fetch_err}")
                try:
                    posts = list(ref.order_by('timestamp', direction=firestore.Query.DESCENDING).limit(200).get())
                except Exception:
                    tk.Label(self.board_cards_frame,
                             text=f"⚠️ 데이터 로딩 실패\n{fetch_err}",
                             bg="#F5F6F8", fg="#c00",
                             font=("맑은 고딕", 10), pady=30).pack()
                    return

            new_count = 0  # 미확인/대화중 카운터
            shown = 0

            for idx, doc in enumerate(posts):
                d = doc.to_dict()
                current_status = d.get('status', '신규')
                current_category = d.get('category', '문의')
                writer = d.get('real_sender') or d.get('user', '익명')
                receiver = d.get('receiver', 'all')
                text = d.get('text', '')
                ts = d.get('timestamp')
                time_str = ts.strftime('%y/%m/%d') if ts else ""

                # 완료 여부 판단
                is_done = ('완료' in current_status) or ('확인완료' in current_status) or (current_status.strip() == '✅ 확인')

                # 필터: 미완료 vs 완료
                if current_filter == "📌 미완료" and is_done:
                    continue
                if current_filter == "✅ 완료" and not is_done:
                    continue

                # 검색어 매칭
                if search_keyword:
                    combined = f"{time_str} {writer} {current_category} {current_status} {text} {receiver}".lower()
                    if search_keyword not in combined:
                        continue

                # 새 메시지/대화 카운트
                if not is_done and ("대화중" in current_status or "신규" in current_status):
                    new_count += 1

                self._render_board_card(doc.id, d, current_status, current_category,
                                          writer, receiver, text, time_str, is_done, idx)
                shown += 1
                if shown >= 100:
                    break

            if shown == 0:
                if current_filter == "📌 미완료":
                    msg = "🎉 미완료 항목이 없습니다!"
                elif search_keyword:
                    msg = f"🔍 '{search_keyword}' 검색 결과가 없습니다"
                else:
                    msg = "📭 최근 7일간 완료된 항목이 없습니다\n(더 보려면 검색하세요)"
                tk.Label(self.board_cards_frame, text=msg,
                         bg="#F5F6F8", fg="#999",
                         font=("맑은 고딕", 11), pady=50, justify="center").pack()
            # 새 메시지 카운터
            if hasattr(self, 'board_new_counter'):
                if new_count > 0:
                    self.board_new_counter.config(text=f"🔔 미처리 {new_count}건", fg="#FF4757")
                else:
                    self.board_new_counter.config(text="")

            # 완료 탭 안내
            if not search_keyword and current_filter == "✅ 완료":
                hint = tk.Label(self.board_cards_frame,
                                 text="💡 최근 7일치만 표시됩니다. 더 오래된 데이터는 위 검색창에 키워드를 입력하세요\n(내용·작업자·날짜·구분 모두 검색 가능)",
                                 bg="#F5F6F8", fg="#999",
                                 font=("맑은 고딕", 8), pady=15, justify="center")
                hint.pack(side="bottom")

        except Exception as e:
            import traceback
            print(f"❌ 공지/소통 카드 렌더링 오류: {e}")
            traceback.print_exc()
            tk.Label(self.board_cards_frame, text=f"⚠️ 오류: {e}",
                     bg="#F5F6F8", fg="#c00",
                     font=("맑은 고딕", 10), pady=30).pack()

    def _render_board_card(self, doc_id, data, status, category, writer, receiver,
                            text, time_str, is_done, idx):
        """공지/소통 카드 한 개 그리기"""
        # 카테고리/상태별 색상
        if is_done:
            accent_color = "#9CA3AF"
            badge_bg = "#E5E7EB"
            badge_fg = "#6B7280"
            card_bg = "#FAFAFA"
            cat_label = f"✅ {category}"
        elif "공지" in category:
            accent_color = "#F59E0B"
            badge_bg = "#FEF3C7"
            badge_fg = "#92400E"
            card_bg = "#FFFEF5"
            cat_label = "📢 공지"
        elif ("요청" in category) or ("지시" in category):
            accent_color = "#8B5CF6"
            badge_bg = "#EDE9FE"
            badge_fg = "#5B21B6"
            card_bg = "white"
            cat_label = f"🔒 요청 → {receiver if receiver != 'all' else ''}"
        elif "대화중" in status:
            accent_color = "#3B82F6"
            badge_bg = "#DBEAFE"
            badge_fg = "#1E40AF"
            card_bg = "white"
            cat_label = "💬 대화중"
        else:  # 신규/문의
            accent_color = "#FF4757"
            badge_bg = "#FFE5E8"
            badge_fg = "#DC2626"
            card_bg = "white"
            cat_label = "💬 문의"

        # 카드 컨테이너
        card_outer = tk.Frame(self.board_cards_frame, bg="#F5F6F8")
        card_outer.pack(fill="x", padx=2, pady=4)

        card = tk.Frame(card_outer, bg=card_bg,
                         highlightthickness=1, highlightbackground="#E5E7EB")
        card.pack(fill="x")

        # 좌측 액센트 바
        accent_bar = tk.Frame(card, bg=accent_color, width=4)
        accent_bar.pack(side="left", fill="y")

        content = tk.Frame(card, bg=card_bg, padx=14, pady=12)
        content.pack(side="left", fill="both", expand=True)

        # 1행: 카테고리 뱃지 + 상태 + 우측 작성자
        row1 = tk.Frame(content, bg=card_bg)
        row1.pack(fill="x")

        # 카테고리 뱃지
        tk.Label(row1, text=cat_label,
                 bg=badge_bg, fg=badge_fg,
                 font=("맑은 고딕", 8, "bold"),
                 padx=8, pady=2).pack(side="left")

        # 진행 상태 뱃지 (대화중 등)
        if not is_done and "대화중" in status:
            tk.Label(row1, text="💬 대화중",
                     bg="#DBEAFE", fg="#1E40AF",
                     font=("맑은 고딕", 8, "bold"),
                     padx=6, pady=2).pack(side="left", padx=(4, 0))

        # 시간
        tk.Label(row1, text=f"📅 {time_str}",
                 bg=card_bg, fg="#9CA3AF",
                 font=("맑은 고딕", 9)).pack(side="left", padx=(8, 0))

        # 우측: 작성자
        tk.Label(row1, text=f"👤 {writer}",
                 bg=card_bg, fg="#374151",
                 font=("맑은 고딕", 9, "bold")).pack(side="right")

        # 2행: 내용 (최대 2줄 미리보기)
        preview = text.replace('\n', ' ').strip()
        if len(preview) > 100:
            preview = preview[:100] + "..."
        tk.Label(content, text=preview,
                 bg=card_bg, fg="#111827",
                 font=("맑은 고딕", 11),
                 anchor="w", justify="left",
                 wraplength=600).pack(fill="x", anchor="w", pady=(6, 0))

        # 클릭/우클릭 이벤트
        def on_click(e):
            self._direct_board_id = doc_id
            self.on_board_double_click(None)
        def on_right_click(e):
            self._selected_board_id = doc_id
            self.board_context_menu.post(e.x_root, e.y_root)
        def on_enter(e):
            card.config(highlightbackground="#1877F2", highlightthickness=2)
        def on_leave(e):
            card.config(highlightbackground="#E5E7EB", highlightthickness=1)

        for widget in [card, content, accent_bar, row1] + list(row1.winfo_children()) + list(content.winfo_children()):
            try:
                widget.bind("<Double-Button-1>", on_click)
                widget.bind("<Button-3>", on_right_click)
                widget.config(cursor="hand2")
            except:
                pass
        card.bind("<Enter>", on_enter)
        card.bind("<Leave>", on_leave)

    # --- [공지/소통 우클릭 메뉴 & 완료/삭제 & 실시간 리스너] ---
    def show_board_context_menu(self, event):
        """옛 board_tree 호환 - 카드는 _selected_board_id로 동작"""
        if hasattr(self, 'board_tree') and self.board_tree.winfo_exists():
            item = self.board_tree.identify_row(event.y)
            if item:
                self.board_tree.selection_set(item)
                self.board_context_menu.post(event.x_root, event.y_root)

    def complete_board_post(self):
        """선택한 게시글을 '완료' 처리 (리스트에서는 숨김, DB에는 유지)"""
        # 카드 시스템: _selected_board_id 우선, 옛 board_tree fallback
        doc_id = getattr(self, '_selected_board_id', None)
        if not doc_id and hasattr(self, 'board_tree') and self.board_tree.winfo_exists():
            selected = self.board_tree.selection()
            if selected:
                doc_id = selected[0]
        if not doc_id:
            return
        try:
            if not messagebox.askyesno("완료 처리",
                f"이 게시글을 완료 처리할까요?\n\n"
                f"※ '미완료' 탭에서는 사라지지만 '✅ 완료' 탭에서 확인 가능합니다."):
                return
            self.db.collection('board_posts').document(doc_id).update({
                'status': '✅ 완료',
                'completed_at': firestore.SERVER_TIMESTAMP
            })
            self._selected_board_id = None
            self.update_board_view()
        except Exception as e:
            messagebox.showerror("오류", f"완료 처리 실패: {e}")

    def delete_board_post(self):
        """선택한 게시글을 완전 삭제 (DB에서도 제거)"""
        doc_id = getattr(self, '_selected_board_id', None)
        if not doc_id and hasattr(self, 'board_tree') and self.board_tree.winfo_exists():
            selected = self.board_tree.selection()
            if selected:
                doc_id = selected[0]
        if not doc_id:
            return
        try:
            if not messagebox.askyesno("완전 삭제",
                "정말로 이 게시글을 완전히 삭제할까요?\n\n"
                "⚠️ 데이터까지 영구 삭제되며 복구할 수 없습니다.\n"
                "(단순히 리스트에서 숨기려면 '완료 처리'를 사용하세요)"):
                return
            # 서브컬렉션 messages도 같이 삭제
            try:
                msgs = self.db.collection('board_posts').document(doc_id).collection('messages').get()
                for m in msgs:
                    m.reference.delete()
            except Exception:
                pass
            self.db.collection('board_posts').document(doc_id).delete()
            messagebox.showinfo("삭제 완료", "게시글이 삭제되었습니다.")
            self._selected_board_id = None
            self.update_board_view()
        except Exception as e:
            messagebox.showerror("오류", f"삭제 실패: {e}")

    def start_board_listener(self):
        """공지/소통 게시판 실시간 리스너."""
        import pyttsx3
        import threading

        # 음성 알림 함수
        def speak_alert(text):
            def _speak():
                try:
                    # [추가] 윈도우 알림 사운드 (띵동~)
                    try:
                        import winsound
                        winsound.MessageBeep(winsound.MB_ICONASTERISK)
                    except Exception: pass

                    import time
                    time.sleep(0.3)

                    engine = pyttsx3.init()
                    engine.setProperty('volume', 1.0)
                    engine.setProperty('rate', 160)
                    engine.say(text)
                    engine.runAndWait()
                except: pass
            threading.Thread(target=_speak, daemon=True).start()

        # 부팅 직후 5초 안에 들어오는 ADDED 이벤트는 알림 안 띄움
        self._board_is_booting = True
        self.root.after(5000, lambda: setattr(self, '_board_is_booting', False))

        def on_board_snapshot(col_snapshot, changes, read_time):
            new_post_detected = False
            should_speak = False  # 음성 알림 띄울지
            for change in changes:
                if change.type.name == 'ADDED' and not getattr(self, '_board_is_booting', True):
                    try:
                        d = change.document.to_dict()
                        cat = d.get('category', '')
                        if '공지' in cat or '요청' in cat or '지시' in cat:
                            # 공지/요청은 관리자가 쓴 거니까 음성 X
                            new_post_detected = True
                            continue
                        else:
                            # 문의 (작업자가 보냄) → 음성 ON
                            new_post_detected = True
                            should_speak = True
                    except Exception:
                        new_post_detected = True

            try:
                self.root.after(10, lambda: self._safe_board_refresh(new_post_detected, should_speak, speak_alert))
            except Exception:
                pass

        try:
            query = self.db.collection('board_posts').order_by(
                'timestamp', direction=firestore.Query.DESCENDING
            ).limit(200)
            query.on_snapshot(on_board_snapshot)
            print("📡 공지/소통 실시간 리스너 가동")
        except Exception as e:
            print(f"❌ 공지/소통 리스너 설정 오류: {e}")

    def _safe_board_refresh(self, new_post_detected=False, should_speak=False, alert_func=None):
        """메인(GUI) 스레드에서 안전하게 board 갱신"""
        try:
            if hasattr(self, 'root') and self.root.winfo_exists():
                self.update_board_view()
                if new_post_detected:
                    self.set_tab_alert(self.t_board, True)
                    # [추가] 문의 글일 때만 음성 알림
                    if alert_func and should_speak:
                        alert_func("문의 확인")
        except Exception as e:
            print(f"❌ board UI 리프레시 오류: {e}")

        # --- [신규 추가] 삭제 관련 함수들 ---
    def show_context_menu(self, event):
        """우클릭 메뉴 - 옛날 chat_tree 호환용 (카드 시스템에서는 _selected_card_id 사용)"""
        if hasattr(self, 'chat_tree') and self.chat_tree.winfo_exists():
            item = self.chat_tree.identify_row(event.y)
            if item:
                self.chat_tree.selection_set(item)
                self.context_menu.post(event.x_root, event.y_root)

    def delete_report(self):
        """선택된 보고서와 하위 댓글들 실제 삭제"""
        # 카드 시스템: _selected_card_id 우선, 옛 chat_tree fallback
        doc_id = getattr(self, '_selected_card_id', None)
        if not doc_id and hasattr(self, 'chat_tree') and self.chat_tree.winfo_exists():
            selected_item = self.chat_tree.selection()
            if selected_item:
                doc_id = selected_item[0]
        if not doc_id:
            return

        from tkinter import messagebox

        if messagebox.askyesno("⚠️ 삭제 확인", "이 보고서와 모든 대화 내역이 영구 삭제됩니다.\n정말 삭제하시겠습니까?"):
            try:
                comments = self.db.collection('field_reports').document(doc_id).collection('comments').get()
                for c in comments:
                    c.reference.delete()
                self.db.collection('field_reports').document(doc_id).delete()
                messagebox.showinfo("삭제 완료", "데이터가 깨끗하게 정리되었습니다.")
                self._selected_card_id = None
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
        df.to_csv(os.path.join(self.save_dir, fn), index=False, encoding="utf-8-sig")
        # [수정] 자동 리셋 제거 - 별도 🔄 리셋 버튼으로 분리됨
        messagebox.showinfo("저장 완료",
            f"✅ 저장됨: {fn}\n\n💡 작업 종료 시 우측 상단의 🔄 리셋 버튼을 눌러주세요.")

    def copy_inbound_to_clipboard(self):
        """[추가] 마지막에 생성한 입고파일 데이터를 클립보드에 복사 (헤더 제외, TSV)."""
        df = getattr(self, '_last_inbound_df', None)
        if df is None or df.empty:
            messagebox.showwarning("복사 불가",
                "복사할 입고 데이터가 없습니다.\n먼저 [📋 입고파일 생성]을 눌러주세요.")
            return
        try:
            # TSV 형식 (탭 구분, 헤더 없음) - 엑셀에 그대로 붙여넣기 가능
            tsv = df.to_csv(sep='\t', index=False, header=False)
            self.root.clipboard_clear()
            self.root.clipboard_append(tsv)
            self.root.update()  # 클립보드 확정
            messagebox.showinfo("복사 완료",
                f"✅ 클립보드에 복사되었습니다!\n\n"
                f"📊 {len(df)}건 (박스번호~불량로케이션)\n"
                f"📋 다른 시트나 시스템에 Ctrl+V로 붙여넣기")
        except Exception as e:
            messagebox.showerror("복사 실패", f"클립보드 복사 중 오류:\n{e}")

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

            # [추가] 복사용으로 마지막 데이터 보관
            self._last_inbound_df = df_out.copy()

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

    # ========== [섹션별 파일 선택 함수] ==========
    def _set_file(self, key, label_widget, prefix=""):
        """파일 선택 다이얼로그 띄우고 라벨 업데이트하는 공통 함수"""
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if not p:
            return
        self.moms_files[key] = p
        label_widget.config(text=f"{prefix}{os.path.basename(p)}", fg="blue")

    # 1. 마스터 생성
    def sel_master_send(self): self._set_file("master_send", self.lbl_master_send)
    def sel_master_master(self): self._set_file("master_master", self.lbl_master_master)
    # 2. 입고리스트
    def sel_inbound_send(self): self._set_file("inbound_send", self.lbl_inbound_send)
    # 3. 제외 입고리스트
    def sel_excl_send(self): self._set_file("excl_send", self.lbl_excl_send)
    def sel_excl_list(self): self._set_file("excl_list", self.lbl_excl_list, prefix="📌 ")
    # 4. 맘스 출고등록
    def sel_mom_out(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if not p: return
        self.moms_files["out_list"] = p
        self.lbl_mom_out.config(text=os.path.basename(p), fg="blue")

    # ========== [1. 신규 마스터 리스트 생성] ==========
    def run_mom_master_logic(self):
        send_file = self.moms_files.get("master_send", "")
        master_file = self.moms_files.get("master_master", "")
        if not send_file:
            messagebox.showwarning("주의", "출고리스트(성수)를 먼저 선택해주세요.")
            return
        if not master_file:
            messagebox.showwarning("주의", "맘스 마스터재고(여주) 파일을 먼저 선택해주세요.")
            return
        try:
            today = datetime.now().strftime('%y%m%d')
            df_s, bc_c = self.smart_load_moms(send_file, '바코드')
            df_m, mc_c = self.smart_load_moms(master_file, '상품코드')

            # [개선] 정확한 코드 매칭을 위해 정규화
            send_codes = self.clean_code_strictly(df_s[bc_c])
            master_codes = set(self.clean_code_strictly(df_m[mc_c]).tolist())

            mask_new = ~send_codes.isin(master_codes)
            df_new = df_s[mask_new].copy()

            if df_new.empty:
                messagebox.showinfo("결과 없음",
                    "출고리스트의 모든 바코드가 이미 마스터 재고에 존재합니다.\n신규로 추가할 항목이 없습니다.")
                return

            def f_c_tmp(df_in, k_in):
                return next((c for c in df_in.columns if k_in in str(c)), None)

            m_reg = pd.DataFrame()
            m_reg['브랜드명'] = df_new.get(f_c_tmp(df_new, '브랜드'), '')
            m_reg['상품군'] = df_new.get(f_c_tmp(df_new, '아이템'), '')
            m_reg['스타일번호'] = df_new.get(f_c_tmp(df_new, '상품코드'), '')
            m_reg['바코드명'] = df_new[bc_c]
            m_reg['상품코드'] = df_new[bc_c]
            m_reg['상품명'] = df_new.get(f_c_tmp(df_new, '상품명'), '')
            m_reg['상품옵션'] = df_new.get(f_c_tmp(df_new, '사이즈'), '')
            m_reg['색상'] = '999'

            fn = self.get_unique_filename(f"{today} 맘스 마스터 리스트", "xlsx")
            m_reg.to_excel(os.path.join(self.save_dir, fn), index=False)

            total_cnt = len(df_s)
            new_cnt = len(df_new)
            excluded_cnt = total_cnt - new_cnt
            messagebox.showinfo("완료",
                f"저장 완료!\n\n📊 결과 요약\n"
                f"• 출고리스트 전체: {total_cnt}건\n"
                f"• 마스터에 이미 존재: {excluded_cnt}건\n"
                f"• 신규 마스터 등록 대상: {new_cnt}건\n\n"
                f"📄 파일: {fn}")

            # 1번 섹션 리셋
            self.moms_files["master_send"] = ""
            self.moms_files["master_master"] = ""
            self.lbl_master_send.config(text="미선택", fg="#777")
            self.lbl_master_master.config(text="미선택", fg="#777")
        except Exception as e:
            messagebox.showerror("오류", str(e))

    # ========== [2. 입고 리스트 생성] ==========
    def run_mom_inbound_logic(self):
        send_file = self.moms_files.get("inbound_send", "")
        if not send_file:
            messagebox.showwarning("주의", "출고리스트(성수)를 먼저 선택해주세요.")
            return
        try:
            today = datetime.now().strftime('%y%m%d')
            df_s, bc_col = self.smart_load_moms(send_file, '바코드')

            qty_col = next((c for c in df_s.columns if '가용재고' in str(c)),
                           next((c for c in df_s.columns if '수량' in str(c)), None))

            i_reg = pd.DataFrame()
            i_reg['화주코드'] = ['INOOINTEMPTY'] * len(df_s)
            i_reg['센터명'] = '신여주2'
            i_reg['층정보'] = 'A2'
            i_reg['마스터코드'] = ""
            i_reg['상품바코드(sku)'] = df_s[bc_col].values
            i_reg['입고'] = df_s[qty_col].values if qty_col else '1'

            fn = self.get_unique_filename(f"{today} 맘스 입고 리스트", "xlsx")
            i_reg.to_excel(os.path.join(self.save_dir, fn), index=False)
            messagebox.showinfo("완료",
                f"저장 완료!\n\n📊 입고 대상: {len(df_s)}건\n📄 파일: {fn}")

            # 2번 섹션 리셋
            self.moms_files["inbound_send"] = ""
            self.lbl_inbound_send.config(text="미선택", fg="#777")
        except Exception as e:
            messagebox.showerror("오류", str(e))

    # ========== [3. 특정상품 제외 입고 리스트 생성] ==========
    def run_mom_inbound_new_only(self):
        """출고리스트 + 제외 명단 → 제외 명단에 있는 상품 빼고 입고 리스트 생성.
        제외 명단은 아이템코드 또는 바코드 둘 중 하나라도 매치되면 제외."""
        send_file = self.moms_files.get("excl_send", "")
        excl_file = self.moms_files.get("excl_list", "")
        if not send_file:
            messagebox.showwarning("주의", "출고리스트(성수)를 먼저 선택해주세요.")
            return
        if not excl_file:
            messagebox.showwarning("주의", "제외 리스트 파일을 선택해주세요.")
            return
        try:
            today = datetime.now().strftime('%y%m%d')
            df_s, bc_col = self.smart_load_moms(send_file, '바코드')

            # 제외 명단 로드 - 아이템코드/바코드 둘 다 시도
            excl_item_codes = set()
            excl_barcodes = set()

            # 아이템코드 시도 (헤더 자동 감지)
            try:
                df_e1, item_col = self.smart_load_moms(excl_file, '아이템')
                excl_item_codes = set(self.clean_code_strictly(df_e1[item_col]).tolist())
            except Exception:
                # 아이템코드 헤더 없을 수 있음
                pass

            # 바코드 시도
            try:
                df_e2, bc_e_col = self.smart_load_moms(excl_file, '바코드')
                excl_barcodes = set(self.clean_code_strictly(df_e2[bc_e_col]).tolist())
            except Exception:
                pass

            if not excl_item_codes and not excl_barcodes:
                messagebox.showerror("오류",
                    "제외 리스트에서 '아이템코드' 또는 '바코드' 컬럼을 찾을 수 없습니다.\n"
                    "헤더에 '아이템' 또는 '바코드'가 포함된 컬럼이 있는지 확인해주세요.")
                return

            # 출고리스트의 바코드 + 아이템코드(있으면) 정규화
            send_barcodes = self.clean_code_strictly(df_s[bc_col])
            # 출고리스트에 아이템코드 있는지 찾기
            item_col_in_send = next((c for c in df_s.columns if '아이템' in str(c) or '상품코드' in str(c)), None)
            send_items = self.clean_code_strictly(df_s[item_col_in_send]) if item_col_in_send else None

            # 제외 마스크: 바코드가 제외명단 바코드에 있거나, 아이템코드가 제외명단 아이템코드에 있으면 제외
            mask_excl = pd.Series([False] * len(df_s), index=df_s.index)
            if excl_barcodes:
                mask_excl = mask_excl | send_barcodes.isin(excl_barcodes)
            if excl_item_codes and send_items is not None:
                mask_excl = mask_excl | send_items.isin(excl_item_codes)

            df_new = df_s[~mask_excl].copy()

            if df_new.empty:
                messagebox.showinfo("결과 없음",
                    "출고리스트의 모든 상품이 제외 명단에 포함되어 있습니다.\n생성할 입고 리스트가 없습니다.")
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

            fn = self.get_unique_filename(f"{today} 맘스 입고 리스트(제외적용)", "xlsx")
            i_reg.to_excel(os.path.join(self.save_dir, fn), index=False)

            total_cnt = len(df_s)
            new_cnt = len(df_new)
            excluded_cnt = total_cnt - new_cnt

            excl_info = []
            if excl_barcodes: excl_info.append(f"바코드 {len(excl_barcodes)}건")
            if excl_item_codes: excl_info.append(f"아이템코드 {len(excl_item_codes)}건")
            excl_info_str = ", ".join(excl_info) if excl_info else "없음"

            messagebox.showinfo("완료",
                f"저장 완료!\n\n📊 결과 요약\n"
                f"• 출고리스트 전체: {total_cnt}건\n"
                f"• 제외 명단 ({excl_info_str}): {excluded_cnt}건 제외됨\n"
                f"• 최종 입고 대상: {new_cnt}건\n\n"
                f"📄 파일: {fn}")

            # 3번 섹션 리셋
            self.moms_files["excl_send"] = ""
            self.moms_files["excl_list"] = ""
            self.lbl_excl_send.config(text="미선택", fg="#777")
            self.lbl_excl_list.config(text="미선택 (아이템코드/바코드 포함)", fg="#777")
        except Exception as e:
            import traceback
            traceback.print_exc()
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
    def sel_chk_target(self): p=filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")]); self.chk_files["target"]=p; self.lbl_chk_target.config(text=os.path.basename(p), fg="blue")
    def sel_chk_master(self): p=filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")]); self.chk_files["master"]=p; self.lbl_chk_master.config(text=os.path.basename(p), fg="blue")
    def reset_inbound(self): 
        self.txt_in_master.tag_remove("mismatch", "1.0", tk.END)
        self.txt_in_scan.tag_remove("mismatch", "1.0", tk.END)
        self.txt_in_master.delete("1.0", tk.END); self.txt_in_scan.delete("1.0", tk.END); self.ent_brand_in.delete(0, tk.END); self.is_matched = False; self.lbl_in_m.config(text="브랜드 수량 (0개)"); self.lbl_in_s.config(text="스캔 수량 (0개)")

    def reset_inbound_all(self):
        """[추가] 입고 탭 전체 초기화 - 브랜드명 + 양쪽 입력창 + 리포트.
        확인 다이얼로그 후 실행."""
        # 비어있으면 그냥 패스
        has_data = bool(
            self.txt_in_master.get("1.0", "end-1c").strip() or
            self.txt_in_scan.get("1.0", "end-1c").strip() or
            self.ent_brand_in.get().strip() or
            self.txt_in_report.get("1.0", "end-1c").strip()
        )
        if not has_data:
            return
        if not messagebox.askyesno("입고 탭 리셋",
                                     "브랜드명, 브랜드 수량, 스캔 수량, 리포트를 모두 지우시겠습니까?",
                                     parent=self.root):
            return

        # 입력창 + 라벨 리셋 (기존 reset_inbound 로직)
        self.txt_in_master.tag_remove("mismatch", "1.0", tk.END)
        self.txt_in_scan.tag_remove("mismatch", "1.0", tk.END)
        self.txt_in_master.delete("1.0", tk.END)
        self.txt_in_scan.delete("1.0", tk.END)
        self.ent_brand_in.delete(0, tk.END)
        self.is_matched = False
        self.lbl_in_m.config(text="브랜드 수량 (0개)")
        self.lbl_in_s.config(text="스캔 수량 (0개)")

        # [추가] 리포트 비우기
        self.txt_in_report.delete("1.0", tk.END)

        # [추가] 마지막 입고 데이터도 비움 (혼동 방지)
        self._last_inbound_df = None
        self.last_scan_detail = []
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
        self._process_closing_stock_file(file_path)

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
            if hasattr(ts, 'astimezone'):
                dt = ts.astimezone(kst)
            elif hasattr(ts, 'seconds'):
                dt = datetime.fromtimestamp(ts.seconds, tz=kst)
            else:
                return ""
            now = datetime.now(kst)
            if dt.date() == now.date():
                return dt.strftime("%H:%M")
            elif dt.year == now.year:
                return dt.strftime("%m/%d %H:%M")
            else:
                return dt.strftime("%Y/%m/%d %H:%M")
        except Exception:
            return ""

    def _get_kst_date(self, ts):
        """timestamp의 한국시간 날짜만 반환 (date 객체). None이면 None."""
        if not ts:
            return None
        try:
            from datetime import datetime, timezone, timedelta
            kst = timezone(timedelta(hours=9))
            if hasattr(ts, 'astimezone'):
                return ts.astimezone(kst).date()
            elif hasattr(ts, 'seconds'):
                return datetime.fromtimestamp(ts.seconds, tz=kst).date()
            return None
        except Exception:
            return None

    def _format_date_separator(self, date_obj):
        """카톡 스타일 날짜 구분선 텍스트. 예: '2026년 4월 29일 수요일'"""
        if not date_obj:
            return ""
        try:
            from datetime import datetime, timezone, timedelta
            kst = timezone(timedelta(hours=9))
            today = datetime.now(kst).date()
            yesterday = today - timedelta(days=1)
            weekdays = ['월요일', '화요일', '수요일', '목요일', '금요일', '토요일', '일요일']
            wd = weekdays[date_obj.weekday()]
            if date_obj == today:
                return f"오늘 · {date_obj.month}월 {date_obj.day}일 {wd}"
            elif date_obj == yesterday:
                return f"어제 · {date_obj.month}월 {date_obj.day}일 {wd}"
            elif date_obj.year == today.year:
                return f"{date_obj.month}월 {date_obj.day}일 {wd}"
            else:
                return f"{date_obj.year}년 {date_obj.month}월 {date_obj.day}일 {wd}"
        except Exception:
            return ""

    def _format_kst_time_only(self, ts):
        """시간만 (HH:MM) 반환. 날짜 구분선이 있으니 시간만 짧게."""
        if not ts:
            return ""
        try:
            from datetime import datetime, timezone, timedelta
            kst = timezone(timedelta(hours=9))
            if hasattr(ts, 'astimezone'):
                return ts.astimezone(kst).strftime("%H:%M")
            elif hasattr(ts, 'seconds'):
                return datetime.fromtimestamp(ts.seconds, tz=kst).strftime("%H:%M")
            return ""
        except Exception:
            return ""

    def _calc_text_size(self, text, max_width=42):
        """말풍선 width/height 정확한 계산 (한글 한 글자 = 2칸).
        Returns: (width, height) tuple"""
        def line_visual_width(line):
            w = 0
            for ch in line:
                w += 2 if ord(ch) > 127 else 1
            return w
        paragraphs = text.split('\n')
        longest_visual = max((line_visual_width(p) for p in paragraphs), default=0)
        width = min(max_width, max(5, longest_visual + 2))
        total_lines = 0
        for p in paragraphs:
            pw = line_visual_width(p)
            if pw == 0:
                total_lines += 1
            else:
                total_lines += (pw + width - 1) // width  # 올림 나눗셈
        return width, max(1, min(30, total_lines))

    def on_message_double_click(self, event):
        import requests
        from io import BytesIO
        from PIL import Image, ImageTk
        import os
        import threading
        from tkinter import messagebox

        # [수정] _direct_selected_id (카드 클릭) 우선, 없으면 chat_tree (구 인터페이스)
        selected_id = getattr(self, '_direct_selected_id', None)
        self._direct_selected_id = None  # 한 번 쓰고 비움

        if not selected_id:
            # 호환: 옛날 chat_tree 인터페이스
            if hasattr(self, 'chat_tree'):
                selected_items = self.chat_tree.selection()
                if not selected_items: return
                selected_id = selected_items[0]
            else:
                return
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

        # [수정] 제목을 복사 가능한 Text 위젯으로 (Label은 복사 안됨)
        title_text_widget = tk.Text(title_frame, height=1, bd=0,
                                     font=("맑은 고딕", 13, "bold"),
                                     bg="white", fg="#262626",
                                     wrap="none", cursor="xterm",
                                     highlightthickness=0, width=40)
        title_text_widget.insert("1.0", data.get('title', '제목 없음'))
        # 읽기 전용처럼 동작 (선택/복사는 가능, 수정은 막음)
        def _readonly_handler(event):
            if event.state & 0x0004:  # Ctrl 누른 상태 (복사 등 허용)
                return None
            if event.keysym in ('Left', 'Right', 'Up', 'Down', 'Home', 'End', 'Prior', 'Next', 'shift'):
                return None
            return "break"
        title_text_widget.bind("<Key>", _readonly_handler)
        title_text_widget.pack(anchor="w")

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
        # 본문 날짜 구분선
        body_date = self._get_kst_date(report_ts)
        if body_date:
            sep_frame = tk.Frame(scrollable_frame, bg="#B2C7DA")
            sep_frame.pack(fill="x", pady=(15, 5))
            sep_pill = tk.Label(sep_frame,
                                 text=self._format_date_separator(body_date),
                                 bg="#9DB0C2", fg="white",
                                 font=("맑은 고딕", 8, "bold"),
                                 padx=12, pady=3)
            sep_pill.pack(anchor="center")

        body_outer = tk.Frame(scrollable_frame, bg="#B2C7DA")
        body_outer.pack(fill="x", pady=(5, 5), padx=15)

        # 작업자 표시 (말풍선 위)
        tk.Label(body_outer, text=f"👤 {data.get('user', '익명')}",
                 font=("맑은 고딕", 9, "bold"), bg="#B2C7DA", fg="#444",
                 anchor="w").pack(anchor="w", padx=(8, 0))

        # 본문 말풍선 (왼쪽 정렬)
        body_bubble_outer = tk.Frame(body_outer, bg="#B2C7DA")
        body_bubble_outer.pack(fill="x", anchor="w")

        body_bubble = tk.Frame(body_bubble_outer, bg="white",
                                highlightthickness=0)
        body_bubble.pack(side="left", anchor="w", padx=(0, 60))

        body_text = data.get('text', '')
        # [수정] 정확한 width/height 계산 (한글=2칸)
        body_width, text_height = self._calc_text_size(body_text, max_width=42)

        body_label = tk.Text(body_bubble, font=("맑은 고딕", 11),
                              bg="white", fg="#1A1A1A",
                              wrap="word", bd=0, padx=14, pady=10,
                              height=text_height, width=body_width,
                              highlightthickness=0, cursor="xterm")
        body_label.insert("1.0", body_text)
        # 읽기 전용 (복사는 가능)
        def _body_readonly(event):
            if event.state & 0x0004:
                return None
            if event.keysym in ('Left', 'Right', 'Up', 'Down', 'Home', 'End', 'Prior', 'Next'):
                return None
            return "break"
        body_label.bind("<Key>", _body_readonly)
        body_label.pack(anchor="w")

        # 시간 (말풍선 옆 작게) - 시간만
        if report_ts:
            tt = self._format_kst_time_only(report_ts)
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
                # [추가] 본문 메타 카운트 감소
                try:
                    doc_ref.update({
                        'comment_count': firestore.Increment(-1)
                    })
                except Exception as e:
                    print(f"⚠️ 카운트 감소 실패: {e}")
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

            # 본문 날짜 = 첫 구분선 기준
            body_date_for_compare = self._get_kst_date(report_ts)
            last_date = body_date_for_compare  # 본문에 이미 날짜 구분선 있으니 같은 날이면 또 안 그림

            for c in comments:
                cid = c.id
                d = c.to_dict()
                u = d.get('user', '관리자')
                t = d.get('text', '')
                img_url = d.get('imageUrl', '')
                edited = d.get('edited', False)
                ts = d.get('timestamp')

                if not detail_win.winfo_exists(): return

                # [추가] 날짜 바뀌면 구분선 표시
                cur_date = self._get_kst_date(ts)
                if cur_date and cur_date != last_date:
                    sep_frame = tk.Frame(comment_list_frame, bg="#B2C7DA")
                    sep_frame.pack(fill="x", pady=(10, 5))
                    tk.Label(sep_frame,
                             text=self._format_date_separator(cur_date),
                             bg="#9DB0C2", fg="white",
                             font=("맑은 고딕", 8, "bold"),
                             padx=12, pady=3).pack(anchor="center")
                    last_date = cur_date

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

                time_str = self._format_kst_time_only(ts)

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

                # 텍스트가 있으면 Text 위젯 (복사 가능), 없으면 패스
                if clean_text:
                    # [수정] 정확한 width/height 계산 (한글=2칸)
                    cmt_width, cmt_height = self._calc_text_size(clean_text, max_width=38)

                    text_label = tk.Text(bubble, font=("맑은 고딕", 11),
                                          bg=bubble_color, fg=text_color,
                                          wrap="word", bd=0, padx=12, pady=8,
                                          height=cmt_height, width=cmt_width,
                                          highlightthickness=0, cursor="xterm")
                    text_label.insert("1.0", clean_text)
                    # 읽기 전용
                    def _cmt_readonly(event):
                        if event.state & 0x0004:
                            return None
                        if event.keysym in ('Left', 'Right', 'Up', 'Down', 'Home', 'End', 'Prior', 'Next'):
                            return None
                        return "break"
                    text_label.bind("<Key>", _cmt_readonly)
                    text_label.bind("<Button-3>", lambda e, id=cid, txt=t.replace("[답장] ", ""): show_comment_menu(e, id, txt))
                    text_label.pack(anchor="w" if not is_admin else "e")

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
            # [추가] 본문 문서에 메타 캐싱 (Firestore 읽기 절약)
            try:
                doc_ref.update({
                    'comment_count': firestore.Increment(1),
                    'last_comment_text': content if content else "(사진)",
                    'last_comment_user': self.current_user,
                    'last_comment_role': 'admin',
                    'last_comment_at': firestore.SERVER_TIMESTAMP,
                    'has_attached_image': True if (image_url or data.get('has_attached_image')) else False
                })
            except Exception as e:
                print(f"⚠️ 메타 업데이트 실패: {e}")
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
                    # [추가] 윈도우 알림 사운드 (띵동~)
                    try:
                        import winsound
                        # MB_ICONASTERISK = 윈도우 기본 알림 사운드 (띵동~)
                        winsound.MessageBeep(winsound.MB_ICONASTERISK)
                    except Exception: pass

                    # 사운드 재생 후 살짝 텀 두고 음성
                    import time
                    time.sleep(0.3)

                    engine = pyttsx3.init()
                    engine.setProperty('volume', 1.0)
                    engine.setProperty('rate', 160)
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
        from datetime import datetime, timedelta, timezone

        # 카드 리스트 영역이 없으면 무시
        if not hasattr(self, 'cards_frame') or not self.cards_frame.winfo_exists():
            return

        # 1. 기존 카드 모두 제거
        for w in self.cards_frame.winfo_children():
            w.destroy()

        try:
            search_keyword = self.search_var.get().strip().lower()
            if search_keyword == "제목·작업자·날짜·내용 검색":
                search_keyword = ""
            current_filter = self.filter_var.get()
            ref = self.db.collection('field_reports')

            # [최적화] 필터별 다른 쿼리
            # - 처리중: 처리중인 것만
            # - 완료: 검색어 없으면 최근 7일치만
            try:
                if search_keyword:
                    # 검색 모드는 더 많이 가져옴
                    raw_docs = list(ref.limit(500).get())
                elif current_filter == "⏳ 처리중":
                    raw_docs = list(ref.where('status', '==', '⏳ 처리중').limit(200).get())
                else:
                    # 완료: 최근 7일만
                    seven_days_ago = datetime.now(timezone.utc) - timedelta(days=7)
                    raw_docs = list(ref.where('status', '==', '✅ 완료')
                                       .where('timestamp', '>=', seven_days_ago)
                                       .limit(200).get())
            except Exception as fetch_err:
                print(f"❌ Firestore 조회 실패: {fetch_err}")
                try:
                    raw_docs = list(ref.limit(200).get())
                except Exception:
                    tk.Label(self.cards_frame,
                             text=f"⚠️ 데이터 로딩 실패\n{fetch_err}",
                             bg="#F5F6F8", fg="#c00",
                             font=("맑은 고딕", 10), pady=30).pack()
                    return

            def _ts_key(d):
                ts = d.to_dict().get('timestamp')
                try:
                    return ts.timestamp() if ts else 0
                except Exception:
                    return 0
            raw_docs.sort(key=_ts_key, reverse=True)

            reports = raw_docs if search_keyword else raw_docs[:100]

            new_comment_count = 0

            shown = 0
            for idx, doc in enumerate(reports):
                data = doc.to_dict()
                status = data.get('status', '⏳ 처리중')
                user = data.get('user', '익명')
                title = data.get('title', '(제목 없음)')
                ts = data.get('timestamp')
                time_str = data.get('date') if data.get('date') else (ts.strftime('%y/%m/%d') if ts else "")

                if search_keyword:
                    # [개선] 본문(text) + 댓글 미리보기까지 검색 대상에 포함
                    text_body = data.get('text', '')
                    last_cmt = data.get('last_comment_text', '')
                    clean_date = time_str.replace("/", "")
                    combined = f"{time_str} {clean_date} {user} {title} {text_body} {last_cmt}".lower()
                    if search_keyword not in combined:
                        continue

                if status != current_filter:
                    continue

                # [최적화] 댓글 컬렉션 읽지 않고 본문에 캐싱된 메타 사용
                comment_total = data.get('comment_count', 0)
                last_comment_text = data.get('last_comment_text', '')
                last_comment_user = data.get('last_comment_user', '')
                last_comment_role = data.get('last_comment_role', '')
                has_image = bool(data.get('imageUrl')) or bool(data.get('has_attached_image', False))

                # 진행 상태 결정 (캐싱된 정보 기반)
                step_text = "🆕 신규"
                is_worker_reply = False
                if comment_total > 0:
                    if last_comment_role == 'worker' or (last_comment_text.startswith('[답장]')):
                        step_text = "💬 NEW"
                        is_worker_reply = True
                    else:
                        step_text = "✅ 확인중"
                # 답장 표시 텍스트 정리
                if last_comment_text.startswith('[답장]'):
                    last_comment_text = last_comment_text.replace('[답장]', '').strip()

                if is_worker_reply and status != "✅ 완료":
                    new_comment_count += 1

                # ========== 카드 그리기 ==========
                self._render_report_card(doc.id, data, status, step_text,
                                          time_str, user, title, ts,
                                          is_worker_reply, comment_total,
                                          has_image, last_comment_text,
                                          last_comment_user, idx)
                shown += 1

            if shown == 0:
                if current_filter == "⏳ 처리중":
                    msg = "🎉 처리중인 작업이 없습니다!"
                elif search_keyword:
                    msg = f"🔍 '{search_keyword}' 검색 결과가 없습니다"
                else:
                    msg = "📭 최근 7일간 완료된 작업이 없습니다\n(더 보려면 검색하세요)"
                tk.Label(self.cards_frame,
                         text=msg,
                         bg="#F5F6F8", fg="#999",
                         font=("맑은 고딕", 11), pady=50, justify="center").pack()

            # 새 댓글 카운터
            if hasattr(self, 'new_comment_counter'):
                if new_comment_count > 0:
                    self.new_comment_counter.config(text=f"🔔 새 댓글 {new_comment_count}건",
                                                      fg="#FF4757")
                else:
                    self.new_comment_counter.config(text="")

            # 검색 안내 - 완료 모드일 때만 7일 제한 안내
            if not search_keyword and current_filter == "✅ 완료":
                hint = tk.Label(self.cards_frame,
                                 text="💡 최근 7일치만 표시됩니다. 더 오래된 데이터는 위 검색창에 키워드를 입력하세요\n(제목·작업자·날짜·내용 모두 검색 가능)",
                                 bg="#F5F6F8", fg="#999",
                                 font=("맑은 고딕", 8), pady=15, justify="center")
                hint.pack(side="bottom")

        except Exception as e:
            import traceback
            print(f"❌ 카드 렌더링 오류: {e}")
            traceback.print_exc()
            tk.Label(self.cards_frame,
                     text=f"⚠️ 오류: {e}",
                     bg="#F5F6F8", fg="#c00",
                     font=("맑은 고딕", 10), pady=30).pack()

    def _render_report_card(self, doc_id, data, status, step_text, time_str, user,
                             title, ts, is_worker_reply, comment_total, has_image,
                             last_comment_text, last_comment_user, idx):
        """모던 카드 한 개 그리기"""
        # 상태별 색상 테마
        if status == "✅ 완료":
            accent_color = "#9CA3AF"
            badge_bg = "#E5E7EB"
            badge_fg = "#6B7280"
            card_bg = "#FAFAFA"
        elif is_worker_reply:
            accent_color = "#FF4757"  # 새 댓글 = 빨간색 강조
            badge_bg = "#FFE5E8"
            badge_fg = "#FF4757"
            card_bg = "white"
        elif step_text == "✅ 확인중":
            accent_color = "#10B981"
            badge_bg = "#D1FAE5"
            badge_fg = "#065F46"
            card_bg = "white"
        else:  # 신규
            accent_color = "#F59E0B"
            badge_bg = "#FEF3C7"
            badge_fg = "#92400E"
            card_bg = "white"

        # 카드 컨테이너 (그림자 효과 흉내)
        card_outer = tk.Frame(self.cards_frame, bg="#F5F6F8")
        card_outer.pack(fill="x", padx=2, pady=4)

        card = tk.Frame(card_outer, bg=card_bg,
                         highlightthickness=1, highlightbackground="#E5E7EB")
        card.pack(fill="x")

        # 좌측 컬러 액센트 바
        accent_bar = tk.Frame(card, bg=accent_color, width=4)
        accent_bar.pack(side="left", fill="y")

        # 본 컨텐츠
        content = tk.Frame(card, bg=card_bg, padx=14, pady=12)
        content.pack(side="left", fill="both", expand=True)

        # 1행: 뱃지 + 시간 + (NEW 표시) + 우측 작업자
        row1 = tk.Frame(content, bg=card_bg)
        row1.pack(fill="x")

        # 뱃지 (상태)
        badge = tk.Label(row1, text=step_text,
                          bg=badge_bg, fg=badge_fg,
                          font=("맑은 고딕", 8, "bold"),
                          padx=8, pady=2)
        badge.pack(side="left")

        # NEW 점멸 (새 댓글 있을 때)
        if is_worker_reply and status != "✅ 완료":
            new_dot = tk.Label(row1, text="●",
                                bg=card_bg, fg="#FF4757",
                                font=("맑은 고딕", 12, "bold"))
            new_dot.pack(side="left", padx=(6, 0))
            # 깜빡임 애니메이션 (간단 버전)
            self._blink_dot(new_dot, card_bg)

        # 시간
        tk.Label(row1, text=f"📅 {time_str}",
                 bg=card_bg, fg="#9CA3AF",
                 font=("맑은 고딕", 9)).pack(side="left", padx=(8, 0))

        # 우측: 작업자 + 첨부 표시
        right_row1 = tk.Frame(row1, bg=card_bg)
        right_row1.pack(side="right")

        if has_image:
            tk.Label(right_row1, text="📷",
                     bg=card_bg, fg="#666",
                     font=("맑은 고딕", 10)).pack(side="right", padx=(4, 0))
        if comment_total > 0:
            tk.Label(right_row1, text=f"💬 {comment_total}",
                     bg=card_bg, fg="#666",
                     font=("맑은 고딕", 9)).pack(side="right", padx=(4, 0))

        tk.Label(right_row1, text=f"👤 {user}",
                 bg=card_bg, fg="#374151",
                 font=("맑은 고딕", 9, "bold")).pack(side="right", padx=(0, 6))

        # 2행: 제목 (크게)
        title_label = tk.Label(content, text=title,
                                bg=card_bg, fg="#111827",
                                font=("맑은 고딕", 12, "bold"),
                                anchor="w", justify="left",
                                wraplength=600)
        title_label.pack(fill="x", anchor="w", pady=(6, 2))

        # 3행: 마지막 댓글 미리보기 (있으면)
        if last_comment_text and is_worker_reply:
            preview = last_comment_text[:60] + ("..." if len(last_comment_text) > 60 else "")
            tk.Label(content,
                     text=f"💬 {last_comment_user}: {preview}",
                     bg=card_bg, fg="#FF4757" if is_worker_reply else "#6B7280",
                     font=("맑은 고딕", 9),
                     anchor="w", justify="left").pack(fill="x", anchor="w")

        # 클릭/우클릭 이벤트 (카드 + 모든 자식 위젯에 바인딩)
        def on_click(e):
            self._open_report_detail(doc_id)
        def on_right_click(e):
            self._selected_card_id = doc_id
            self.context_menu.post(e.x_root, e.y_root)
        def on_enter(e):
            card.config(highlightbackground="#1877F2", highlightthickness=2)
        def on_leave(e):
            card.config(highlightbackground="#E5E7EB", highlightthickness=1)

        for widget in [card, content, accent_bar, row1, right_row1, title_label] + list(row1.winfo_children()) + list(content.winfo_children()) + list(right_row1.winfo_children()):
            try:
                widget.bind("<Double-Button-1>", on_click)
                widget.bind("<Button-3>", on_right_click)
                widget.config(cursor="hand2")
            except:
                pass
        card.bind("<Enter>", on_enter)
        card.bind("<Leave>", on_leave)

    def _blink_dot(self, widget, bg_color):
        """빨간 점 깜빡임 (NEW 표시용)"""
        def toggle():
            try:
                if not widget.winfo_exists(): return
                cur = widget.cget("fg")
                widget.config(fg="#FF4757" if cur == bg_color else bg_color)
                widget.after(700, toggle)
            except: pass
        toggle()

    def _open_report_detail(self, doc_id):
        """카드 더블클릭 시 상세 팝업 열기 - 기존 on_message_double_click 호환"""
        # 기존 on_message_double_click이 self.chat_tree.selection() 사용하므로
        # 임시로 selected_id만 따로 저장 + 호환용 함수 호출
        self._direct_selected_id = doc_id
        self._call_detail_directly(doc_id)

    def _call_detail_directly(self, doc_id):
        """상세 팝업 띄우기 (chat_tree 없이도 동작)"""
        # 기존 on_message_double_click 본문을 직접 호출하기 위해
        # selected_items 흉내내는 방법 - 간단하게는 그냥 본문 코드 리팩터해야 하지만,
        # 우선 chat_tree 없으면 동작 안 하니 selected_id만 _direct_selected_id에 저장하고
        # on_message_double_click을 살짝 바꿔서 _direct_selected_id 우선 사용하게 함
        self.on_message_double_click(None)

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
    # [추가] 드래그앤드롭 지원 (tkinterdnd2 있으면 우선 사용, 없으면 기본 Tk)
    try:
        from tkinterdnd2 import TkinterDnD
        root = TkinterDnD.Tk()
        print("✅ tkinterdnd2 활성화 - 드래그앤드롭 사용 가능")
    except ImportError:
        root = tk.Tk()
        print("ℹ️ tkinterdnd2 미설치 - 기본 모드 (파일 선택 버튼만 사용)")
    app = LogiPanApp(root)
    root.mainloop()
