# 🚚 로지판 (LogiPan)

무신사 로지스틱스 창고팀 사내 통합 물류 도구

---

## 📦 파일 구조

### 🖥️ 데스크톱 앱 (Python / Tkinter)

| 파일 | 역할 |
|---|---|
| `updater.py` | 런처 - 자동 업데이트 + 로지판 실행 |
| `LogiPan.py` | 메인 앱 (입고/출고/마스터/공지/소통 등) |
| `slack_integration.py` | Slack 통합 모듈 (Mixin 패턴) |
| `version.txt` | 버전 정보 - 이거 갱신 시 자동 배포 트리거 |
| `requirements.txt` | Python 의존성 |

### 📱 작업자 웹앱 (PWA)

| 파일 | 역할 |
|---|---|
| `index.html` | 현장 보고 + 공지/소통 (모바일) |
| `sw.js` | 서비스 워커 (PWA 기본) |
| `firebase-messaging-sw.js` | FCM 푸시 알림 처리 |
| `manifest.json` | PWA 매니페스트 |

호스팅: https://ghwkdwjd-debug.github.io/logipan-report/

### 🔑 설정 / 리소스

| 파일 | 역할 |
|---|---|
| `oauth_client.json` | Google OAuth (시트 검색용) |
| `로지판.ico` | Windows 아이콘 |
| `icon.png`, `favicon.ico`, `android-chrome-512x512.png` | 웹 아이콘 |

---

## 🔄 자동 업데이트 시스템

### 동작 흐름

```
[로지판 실행]
   ↓
[updater.py 실행]
   ↓
[GitHub version.txt 체크]
   ↓
로컬 ≠ 원격? → MODULE_FILES 전체 다운로드 (백업/롤백 가능)
   ↓
[LogiPan.py 실행]
   ↓
[_self_update_check]
   - 부트스트랩: slack_integration.py 등 누락 모듈 자동 다운로드
   - updater.py 해시 비교 후 갱신 (다음 실행부터 적용)
```

### 안전장치

- ✅ 모든 파일을 `.tmp`로 먼저 받음 (실패 시 기존 파일 손 안 댐)
- ✅ 교체 시 `.bak` 백업 (실패 시 자동 롤백)
- ✅ 빈 파일 방어 (100바이트 미만 거부)
- ✅ 부트스트랩: 필수 모듈 누락 시 자가 복구
- ✅ 해시 비교: updater.py 변경 시 자동 갱신

---

## 🛠️ 외부 연동

- **Firebase** (logipan-2026): Firestore + FCM 푸시
- **Slack**: Webhook + 멘션 + CC 그룹
- **Jira**: musinsa-oneteam.atlassian.net
- **Google Sheets API**: 브랜드/MD 자동 검색
- **imgBB**: 이미지 호스팅

---

## ⚠️ 배포 시 주의사항

**`version.txt` 갱신은 신중하게.**

이 값을 올리는 순간 모든 사용자 PC가 다음 실행 때 자동 업데이트 트리거됨. 본인 PC에서 며칠 충분히 검증 후 갱신할 것.

---

## 📊 모듈화 진행 상황

- [x] 1단계: Slack 통합 분리 (`slack_integration.py`)
- [ ] 2단계: Jira 통합 분리 (`jira_integration.py`)
- [ ] 3단계: Firebase/FCM 분리 (`firebase_utils.py`)
- [ ] 4단계 이후: 마스터, 입출고, 공지 등 점진적
