# UpNote ↔ Obsidian 동기화 도구

## 파일 구조

```
UpNote_to_Obsidian/
├── sync_engine.py       # 핵심 동기화 엔진
├── sync_config.py       # .env 로드 및 CONFIG 조립 (수정 불필요)
├── setup_scheduler.py   # Windows 스케줄러 등록/해제
├── .env.example         # 환경변수 템플릿 (Git 포함)
├── .gitignore
└── README.md
```

공유 `.env` 위치 (프로젝트 루트 또는 같은 폴더):
```
python-automation-toolkit/
├── .env                 # ★ 경로 설정 여기서만 (Git 제외)
└── UpNote_to_Obsidian/
    └── sync_engine.py   # 상위 폴더 .env 자동 탐색
```

자동 생성 파일:
```
SYNC_DATA_DIR/           # .env 에서 지정한 경로
├── sync_map.json        # UUID ↔ Obsidian 경로 매핑 DB
├── sync.log             # 동기화 이력 텍스트 로그
└── sync_log.xlsx        # 동기화 이력 엑셀 로그
```

---

## 1. 초기 설정

### 패키지 설치 (최초 1회)

```bash
pip install python-dotenv openpyxl
```

### .env 파일 생성

`.env.example` 을 복사해서 `.env` 로 저장 후 경로 수정:

```env
# 절대 경로 또는 %USERPROFILE% 사용 (~ 사용 불가)
UPNOTE_ROOT=%USERPROFILE%\Desktop\UpNote_Backup\F5p9KpV016SPRSpBH6y3uJ8NcFm2
OBSIDIAN_VAULT=%USERPROFILE%\Documents\Obsidian\MyVault
SYNC_DATA_DIR=%USERPROFILE%\Documents\upnote_obs_sync

# 아래는 기본값 그대로 두면 됨
MTIME_TOLERANCE_SEC=3
INJECT_UUID_FRONTMATTER=true
STRIP_IMAGES=true
```

> `.env`는 `.gitignore`에 등록되어 있어 Git에 올라가지 않습니다.

---

## 2. 수동 실행 (VS Code 터미널)

`sync_engine.py` 가 있는 폴더에서 실행:

```bash
cd C:\Python\python-automation-toolkit\UpNote_to_Obsidian

# 양방향 동기화 (기본)
python sync_engine.py

# UpNote → Obsidian 단방향
python sync_engine.py --up-to-obs

# Obsidian → UpNote 단방향
python sync_engine.py --obs-to-up

# 실제 파일 변경 없이 미리보기
python sync_engine.py --dry-run
```

---

## 3. 자동 실행 (Windows 스케줄러)

```bash
# 30분마다 자동 동기화 등록
python setup_scheduler.py --register

# 등록 상태 확인
python setup_scheduler.py --status

# 해제
python setup_scheduler.py --unregister
```

등록 후 `taskschd.msc` → 작업 스케줄러 라이브러리 → `UpNote_Obsidian_Sync` 에서 확인 가능.

---

## 4. 동기화 규칙

| 상황 | 처리 |
|------|------|
| UpNote만 변경 | Obsidian에 덮어쓰기 |
| Obsidian만 변경 | UpNote UUID.md에 덮어쓰기 |
| 양쪽 모두 변경 (충돌) | 더 최근 수정 파일 우선 적용 |
| 처음 발견된 UpNote 노트 | Obsidian에 폴더 구조 그대로 생성 |
| 처음 발견된 Obsidian 노트 | UpNote Notes/ 폴더에 UUID.md 생성 |

UpNote → Obsidian 변환 시 적용되는 처리:
- 첫 번째 헤딩(제목) 줄 제거 — 파일명이 제목 역할
- 이미지 링크 제거 — 텍스트만 유지
- 빈 줄 / `<br>` 전체 제거

---

## 5. Obsidian frontmatter

동기화된 노트에는 자동으로 frontmatter가 삽입됩니다:

```yaml
---
upnote_uuid: 4b11e655-2207-4c6c-b109-3ed572a6016c
---
```

이 UUID로 역방향(Obsidian → UpNote) 추적이 가능합니다.  
삽입을 원하지 않으면 `.env` 에서:

```env
INJECT_UUID_FRONTMATTER=false
```

---

## 6. 로그 확인

### 텍스트 로그 (`sync.log`)

```
2024-04-06 12:00:01 [INFO] ✅ UP→OBS  [4b11e655] 업무/프로젝트 회고.md
2024-04-06 12:00:02 [INFO] ✅ OBS→UP  [9f3c2a11] → 9f3c2a11-....md
2024-04-06 12:00:03 [WARNING] ⚡ 충돌 감지 [7e8d1b00]: 더 최근 수정 파일 우선 적용
```

### 엑셀 로그 (`sync_log.xlsx`)

| 날짜시간 | 실행모드 | 상태 | UUID | 노트 제목 | 오류 메시지 |
|---|---|---|---|---|---|
| 2024-04-06 12:00:00 | both | 시작 | | 총 스캔 시작 | |
| 2024-04-06 12:00:01 | both | UP→OBS | 4b11e655 | 프로젝트 회고 | |
| 2024-04-06 12:00:02 | both | 충돌 | 9f3c2a11 | 독서 메모 | |
| 2024-04-06 12:00:03 | both | 오류 | 7e8d1b00 | | PermissionError |
| 2024-04-06 12:00:03 | both | 완료 | | UP→OBS 2건 \| 오류 1건 | |

상태별 행 색상이 자동 적용됩니다 (파랑: UP→OBS / 초록: OBS→UP / 노랑: 충돌 / 빨강: 오류).  
스킵된 노트는 텍스트 로그에만 기록되고 엑셀에는 포함되지 않습니다.

---

## 7. 버전 이력

| 버전 | 내용 |
|------|------|
| v1 | `fix_titles.py` — UpNote → Obsidian 단방향 변환 |
| v2 | `sync_engine.py` — 양방향 동기화 + 엑셀 로그 + `.env` 기반 설정 |