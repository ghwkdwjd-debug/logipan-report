"""
로지판 자동 업데이트 + 실행 스크립트 (멀티파일 + 백업/롤백 지원 버전)
──────────────────────────────────────────────────────────────────────
1. GitHub에서 version.txt 받아옴
2. 내 버전과 비교
3. 새 버전이면 MODULE_FILES 전체 다운로드 (원자적 + 롤백 가능)
4. LogiPan.py 실행

[v2 변경점 - 2025-05-08 - 모듈화 1단계]
- LogiPan.py 단일 파일 → MODULE_FILES 리스트로 확장
- 모든 파일 .tmp로 먼저 다 받기 (하나라도 실패하면 기존 파일 손 안 댐)
- 교체 시 .bak 백업 (실패 시 자동 롤백)
- 4명 PC 동시 영향 위험 최소화

진행 상황을 작은 GUI 창으로 보여줌 (콘솔 X)
"""
import os
import sys
import subprocess
import urllib.request
import urllib.error
import tkinter as tk
from tkinter import messagebox, ttk
import threading
import time
import shutil

# ───────────── 설정 ─────────────
GITHUB_USER = "ghwkdwjd-debug"
GITHUB_REPO = "logipan-report"
GITHUB_BRANCH = "main"

BASE_URL = f"https://raw.githubusercontent.com/{GITHUB_USER}/{GITHUB_REPO}/{GITHUB_BRANCH}"
VERSION_URL = f"{BASE_URL}/version.txt"
REQUIREMENTS_URL = f"{BASE_URL}/requirements.txt"

# [v2] 다운로드할 모듈 파일들 - 추후 모듈화 진행할 때 여기에 계속 추가
MODULE_FILES = [
    "LogiPan.py",
    "slack_integration.py",
    "jira_integration.py",
]

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
LOCAL_VERSION = os.path.join(SCRIPT_DIR, "version.txt")
LOCAL_LOGIPAN = os.path.join(SCRIPT_DIR, "LogiPan.py")
LOCAL_REQUIREMENTS = os.path.join(SCRIPT_DIR, "requirements.txt")


# ────────────── 진행창 클래스 ──────────────
class ProgressWindow:
    """업데이트 진행 상황을 보여주는 작은 GUI 창"""
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("로지판 업데이트")
        self.root.configure(bg="#f5f5f5")
        # 화면 가운데 정렬
        w, h = 380, 140
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.root.geometry(f"{w}x{h}+{x}+{y}")
        # 항상 위에
        self.root.attributes("-topmost", True)
        # 닫기 버튼 비활성 (업데이트 중 끄면 안됨)
        self.root.protocol("WM_DELETE_WINDOW", lambda: None)

        # 제목
        tk.Label(self.root, text="🚀 로지판",
                 font=("맑은 고딕", 14, "bold"),
                 bg="#f5f5f5", fg="#1a73e8").pack(pady=(15, 5))

        # 진행 메시지
        self.message_var = tk.StringVar(value="업데이트 확인 중...")
        tk.Label(self.root, textvariable=self.message_var,
                 font=("맑은 고딕", 10), bg="#f5f5f5", fg="#444").pack(pady=2)

        # 진행 막대 (불확정 - 계속 움직이는 거)
        self.progress = ttk.Progressbar(self.root, mode='indeterminate',
                                          length=300)
        self.progress.pack(pady=10)
        self.progress.start(10)

    def update_message(self, text):
        """메시지 업데이트"""
        try:
            self.message_var.set(text)
            self.root.update_idletasks()
        except Exception:
            pass

    def close(self):
        try:
            self.progress.stop()
            self.root.destroy()
        except Exception:
            pass


# ───────────── 업데이트 로직 ─────────────
def get_local_version():
    if not os.path.exists(LOCAL_VERSION):
        return None
    try:
        with open(LOCAL_VERSION, "r", encoding="utf-8") as f:
            return f.read().strip()
    except Exception:
        return None


def get_remote_version():
    try:
        with urllib.request.urlopen(VERSION_URL, timeout=10) as response:
            return response.read().decode("utf-8").strip()
    except urllib.error.URLError:
        return None
    except Exception:
        return None


def download_file(url, local_path):
    """단일 파일 다운로드 (구버전 호환용 - requirements.txt 같은 보조 파일에 사용)"""
    try:
        tmp_path = local_path + ".tmp"
        urllib.request.urlretrieve(url, tmp_path)
        if os.path.exists(local_path):
            os.remove(local_path)
        os.rename(tmp_path, local_path)
        return True
    except Exception as e:
        print(f"다운로드 실패: {e}")
        return False


# ────────────── [v2] 멀티파일 원자적 다운로드 + 롤백 ──────────────
def download_modules_atomic(module_files, base_url, script_dir, progress_callback=None):
    """모든 모듈 파일을 원자적으로 다운로드.

    동작:
    1. 모든 파일을 .tmp로 받음 (기존 파일 손 안 댐)
    2. 다 받았으면 → 기존 파일들을 .bak으로 백업 → .tmp를 본 이름으로 rename
    3. 중간에 실패하면 → .tmp 다 삭제하고 False 리턴 (기존 파일 유지)
    4. 교체 도중 실패하면 → .bak에서 복구 시도

    Args:
        module_files: ["LogiPan.py", "slack_integration.py", ...]
        base_url: GitHub raw URL prefix
        script_dir: 로컬 디렉토리
        progress_callback: 진행상황 콜백 (str -> None)

    Returns:
        (success: bool, error_msg: str)
    """
    tmp_paths = []   # [(url, tmp_path, final_path), ...]
    bak_paths = []   # [(bak_path, final_path), ...] - 롤백용

    # ── 1단계: 모든 파일을 .tmp로 받기 ──
    try:
        for i, fname in enumerate(module_files, 1):
            url = f"{base_url}/{fname}"
            final_path = os.path.join(script_dir, fname)
            tmp_path = final_path + ".tmp"

            if progress_callback:
                progress_callback(f"📥 다운로드 중... ({i}/{len(module_files)}) {fname}")

            try:
                # 기존 .tmp가 남아있으면 지움
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
                urllib.request.urlretrieve(url, tmp_path)
            except Exception as e:
                # 하나라도 실패하면 → 받아둔 .tmp 다 정리하고 실패 리턴
                _cleanup_tmps(tmp_paths)
                return False, f"{fname} 다운로드 실패: {e}"

            # 빈 파일 검증 (네트워크 끊김 등으로 0바이트 받는 경우 방어)
            if os.path.getsize(tmp_path) < 100:
                _cleanup_tmps(tmp_paths + [(url, tmp_path, final_path)])
                return False, f"{fname} 다운로드 파일이 너무 작음 (손상 의심)"

            tmp_paths.append((url, tmp_path, final_path))

    except Exception as e:
        _cleanup_tmps(tmp_paths)
        return False, f"다운로드 단계 오류: {e}"

    # ── 2단계: 백업 + 교체 (원자적) ──
    # 여기까지 왔으면 .tmp 파일들은 다 안전하게 받아진 상태
    if progress_callback:
        progress_callback("💾 백업 및 교체 중...")

    try:
        for url, tmp_path, final_path in tmp_paths:
            bak_path = final_path + ".bak"
            # 기존 .bak 있으면 삭제
            if os.path.exists(bak_path):
                try:
                    os.remove(bak_path)
                except Exception:
                    pass
            # 기존 파일이 있으면 .bak으로 이름 변경
            if os.path.exists(final_path):
                try:
                    os.rename(final_path, bak_path)
                    bak_paths.append((bak_path, final_path))
                except Exception as e:
                    # 백업 실패 → 롤백
                    _rollback(bak_paths)
                    _cleanup_tmps(tmp_paths)
                    return False, f"{os.path.basename(final_path)} 백업 실패: {e}"
            # .tmp를 본 이름으로 변경
            try:
                os.rename(tmp_path, final_path)
            except Exception as e:
                # 교체 실패 → 롤백
                _rollback(bak_paths)
                _cleanup_tmps(tmp_paths)
                return False, f"{os.path.basename(final_path)} 교체 실패: {e}"

    except Exception as e:
        _rollback(bak_paths)
        _cleanup_tmps(tmp_paths)
        return False, f"교체 단계 오류: {e}"

    # 성공 - .bak은 한 번 더 갱신 안전망으로 남겨둠 (수동 롤백 가능하게)
    return True, ""


def _cleanup_tmps(tmp_paths):
    """다운로드 실패 시 .tmp 파일들 정리"""
    for entry in tmp_paths:
        try:
            tmp_path = entry[1] if isinstance(entry, tuple) else entry
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        except Exception:
            pass


def _rollback(bak_paths):
    """교체 중 실패 시 .bak에서 복구"""
    for bak_path, final_path in bak_paths:
        try:
            if os.path.exists(final_path):
                # 새로 교체된 거 제거
                os.remove(final_path)
            if os.path.exists(bak_path):
                # 백업을 원위치로
                os.rename(bak_path, final_path)
            print(f"⚠️ 롤백 완료: {os.path.basename(final_path)}")
        except Exception as e:
            print(f"⚠️ 롤백 실패 ({final_path}): {e}")


def all_modules_exist():
    """필수 모듈 파일들이 다 있는지 (오프라인 실행 가능 여부 판단용)"""
    for fname in MODULE_FILES:
        if not os.path.exists(os.path.join(SCRIPT_DIR, fname)):
            return False
    return True


def install_requirements():
    if not os.path.exists(LOCAL_REQUIREMENTS):
        return True
    try:
        # Windows에서 콘솔창 안 뜨게
        startupinfo = None
        if os.name == 'nt':
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

        subprocess.run(
            [sys.executable, "-m", "pip", "install", "-r", LOCAL_REQUIREMENTS, "--quiet"],
            capture_output=True, text=True,
            startupinfo=startupinfo
        )
        return True
    except Exception:
        return True


def check_key_json():
    key_path = os.path.join(SCRIPT_DIR, "key.json")
    if not os.path.exists(key_path):
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "key.json 누락",
            "Firebase 키 파일(key.json)이 없습니다.\n\n"
            f"이 폴더에 key.json을 넣어주세요:\n{SCRIPT_DIR}\n\n"
            "관리자에게 문의하세요."
        )
        root.destroy()
        return False
    return True


def run_logipan():
    """LogiPan.py 실행 (이 프로세스를 대체)"""
    try:
        # Windows에서 콘솔창 안 뜨게
        creationflags = 0
        if os.name == 'nt':
            creationflags = subprocess.CREATE_NO_WINDOW

        subprocess.Popen([sys.executable, LOCAL_LOGIPAN],
                         cwd=SCRIPT_DIR,
                         creationflags=creationflags)
    except Exception as e:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("실행 오류", f"로지판 실행 실패:\n{e}")
        root.destroy()


# ───────────── 메인 워크 (백그라운드 스레드용) ─────────────
def update_work(progress_win):
    """진행창 띄운 채로 업데이트 작업"""
    try:
        # 1) key.json 확인
        if not check_key_json():
            progress_win.root.after(0, progress_win.close)
            return

        # 2) 버전 확인
        progress_win.update_message("📡 서버 연결 확인 중...")
        local_v = get_local_version()
        remote_v = get_remote_version()

        # 3) 인터넷 안 되면
        if remote_v is None:
            # [v2] 모든 모듈 파일 다 있어야 오프라인 실행 가능
            if all_modules_exist():
                progress_win.update_message("⚠️ 인터넷 연결 안됨, 기존 버전으로 실행")
                time.sleep(1)
                progress_win.root.after(0, progress_win.close)
                run_logipan()
            else:
                progress_win.root.after(0, progress_win.close)
                missing = [f for f in MODULE_FILES
                           if not os.path.exists(os.path.join(SCRIPT_DIR, f))]
                root = tk.Tk()
                root.withdraw()
                messagebox.showerror(
                    "실행 불가",
                    "인터넷 연결이 안 되고 필수 파일이 없습니다.\n\n"
                    f"누락: {', '.join(missing)}\n\n"
                    "인터넷 연결을 확인해주세요."
                )
                root.destroy()
            return

        # 4) 업데이트 필요 여부 판단
        # [v2] 버전이 다르거나, 모듈 파일 중 하나라도 없으면 업데이트
        need_update = (local_v != remote_v) or (not all_modules_exist())

        if need_update:
            if local_v is None:
                progress_win.update_message(f"📥 로지판 첫 설치 중... (v{remote_v})")
            else:
                progress_win.update_message(f"🆕 새 버전 다운로드 중... ({local_v} → {remote_v})")

            # [v2] 멀티파일 원자적 다운로드
            ok, err_msg = download_modules_atomic(
                MODULE_FILES, BASE_URL, SCRIPT_DIR,
                progress_callback=progress_win.update_message
            )
            if not ok:
                progress_win.root.after(0, progress_win.close)
                # 기존 파일이 있으면 그걸로라도 실행 시도
                if all_modules_exist():
                    root = tk.Tk()
                    root.withdraw()
                    messagebox.showwarning(
                        "업데이트 실패",
                        f"새 버전 다운로드에 실패했습니다.\n\n"
                        f"오류: {err_msg}\n\n"
                        "기존 버전으로 실행합니다."
                    )
                    root.destroy()
                    run_logipan()
                else:
                    root = tk.Tk()
                    root.withdraw()
                    messagebox.showerror(
                        "다운로드 실패",
                        f"필수 파일 다운로드에 실패했습니다.\n\n"
                        f"오류: {err_msg}\n\n"
                        "인터넷 상태를 확인하세요."
                    )
                    root.destroy()
                return

            progress_win.update_message("📦 라이브러리 확인 중...")
            download_file(REQUIREMENTS_URL, LOCAL_REQUIREMENTS)
            install_requirements()

            try:
                with open(LOCAL_VERSION, "w", encoding="utf-8") as f:
                    f.write(remote_v)
            except Exception:
                pass

            progress_win.update_message(f"✅ 업데이트 완료 (v{remote_v})")
            time.sleep(0.8)
        else:
            progress_win.update_message(f"✅ 최신 버전 (v{remote_v})")
            time.sleep(0.5)

        # 5) 로지판 실행 + 진행창 닫기
        progress_win.update_message("🚀 로지판 시작...")
        time.sleep(0.3)
        run_logipan()
        # 메인 스레드에서 닫기
        progress_win.root.after(0, progress_win.close)

    except Exception as e:
        progress_win.root.after(0, progress_win.close)
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("런처 오류", f"예상치 못한 오류가 발생했습니다:\n\n{e}")
        root.destroy()


def main():
    progress_win = ProgressWindow()
    # 메인 mainloop 시작 직후에 워커 스레드 시작
    progress_win.root.after(100, lambda: threading.Thread(
        target=update_work, args=(progress_win,), daemon=True).start())
    progress_win.root.mainloop()


if __name__ == "__main__":
    main()
