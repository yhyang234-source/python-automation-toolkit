"""
sync_config.py — UpNote ↔ Obsidian 동기화 설정
=================================================
가상 환경 내의 패키지 설치 상태와 .env 로드 성공 여부를 
시스템 인터페이스 관점에서 명확히 진단하도록 설계되었습니다.
"""

import os
import sys
from pathlib import Path

# [Logic] 환경변수 로드 및 시스템 진단
try:
    from dotenv import load_dotenv, find_dotenv
    
    # 프로젝트 루트의 .env 위치를 지능적으로 탐색합니다.
    env_path = find_dotenv(usecwd=True)
    
    if env_path:
        load_dotenv(env_path)
    else:
        # 설계 의도: .env가 없을 경우 사용자에게 생성 가이드를 제공합니다.
        print("\n[System Warning] .env 파일을 찾을 수 없습니다.")
        print("프로젝트 루트에 .env 파일을 생성하고 설정을 입력해 주세요.\n")

except ImportError:
    # 설계 의도: 가상 환경 미설치 시 구체적인 해결 명령어를 출력합니다.
    print("\n" + "="*50)
    print("[System Error] 'python-dotenv' 패키지가 현재 가상 환경에 없습니다.")
    print(f"현재 실행 환경: {sys.executable}")
    print("해결 방법: 아래 명령어를 터미널에 복사하여 실행하세요.")
    print(f"{sys.executable} -m pip install python-dotenv")
    print("="*50 + "\n")
    sys.exit(1)

# [Utility] 환경변수 유효성 검사
def _require(key: str) -> str:
    """필수 환경변수 누락 시 명확한 오류 메시지 출력"""
    val = os.getenv(key)
    if not val:
        raise EnvironmentError(
            f"\n[sync_config] 필수 설정값 '{key}' 가 .env 파일에 없습니다.\n"
            f"탐색된 설정 파일 경로: {env_path if 'env_path' in locals() and env_path else '찾을 수 없음'}"
        )
    return val

def _get(key: str, default: str) -> str:
    return os.getenv(key, default)

def _expandvars(path: str) -> str:
    """Windows 환경변수 및 경로 확장 처리"""
    return os.path.expanduser(os.path.expandvars(path))

"""
sync_config.py — UpNote ↔ Obsidian 동기화 설정
=================================================
가상 환경 내의 패키지 설치 상태와 .env 로드 성공 여부를 
시스템 인터페이스 관점에서 명확히 진단하도록 설계되었습니다.
"""

import os
import sys
from pathlib import Path

# [Logic] 환경변수 로드 및 시스템 진단
try:
    from dotenv import load_dotenv, find_dotenv
    
    # 프로젝트 루트의 .env 위치를 지능적으로 탐색합니다.
    env_path = find_dotenv(usecwd=True)
    
    if env_path:
        load_dotenv(env_path)
    else:
        # 설계 의도: .env가 없을 경우 사용자에게 생성 가이드를 제공합니다.
        print("\n[System Warning] .env 파일을 찾을 수 없습니다.")
        print("프로젝트 루트에 .env 파일을 생성하고 설정을 입력해 주세요.\n")

except ImportError:
    # 설계 의도: 가상 환경 미설치 시 구체적인 해결 명령어를 출력합니다.
    print("\n" + "="*50)
    print("[System Error] 'python-dotenv' 패키지가 현재 가상 환경에 없습니다.")
    print(f"현재 실행 환경: {sys.executable}")
    print("해결 방법: 아래 명령어를 터미널에 복사하여 실행하세요.")
    print(f"{sys.executable} -m pip install python-dotenv")
    print("="*50 + "\n")
    sys.exit(1)

# [Utility] 환경변수 유효성 검사
def _require(key: str) -> str:
    """필수 환경변수 누락 시 명확한 오류 메시지 출력"""
    val = os.getenv(key)
    if not val:
        raise EnvironmentError(
            f"\n[sync_config] 필수 설정값 '{key}' 가 .env 파일에 없습니다.\n"
            f"탐색된 설정 파일 경로: {env_path if 'env_path' in locals() and env_path else '찾을 수 없음'}"
        )
    return val

def _get(key: str, default: str) -> str:
    return os.getenv(key, default)

def _expandvars(path: str) -> str:
    """Windows 환경변수 및 경로 확장 처리"""
    return os.path.expanduser(os.path.expandvars(path))

# [Integration] 설정 데이터 조립
# [Integration] 설정 데이터 조립
try:
    _UPNOTE_ROOT    = _expandvars(_require("UPNOTE_ROOT"))
    _OBSIDIAN_VAULT = _expandvars(_require("OBSIDIAN_VAULT"))
    _SYNC_DATA_DIR  = _expandvars(_require("SYNC_DATA_DIR"))

    CONFIG = {
        "UPNOTE_ROOT":          _UPNOTE_ROOT,
        
        # [수정됨] 끝에 있던 "Notes"를 제거하여 원본 파일이 쏟아져 있는 Synapse OS 폴더를 직격으로 가리키게 합니다.
        "UPNOTE_NOTES_DIR":     os.path.join(_UPNOTE_ROOT, "Markdown", "Synapse OS"),
        
        "UPNOTE_NOTEBOOKS_DIR": os.path.join(_UPNOTE_ROOT, "Markdown", "Synapse OS", "notebooks"),
        "OBSIDIAN_VAULT_DIR":   _OBSIDIAN_VAULT,
        "SYNC_MAP_PATH":        os.path.join(_SYNC_DATA_DIR, "sync_map.json"),
        "LOG_PATH":             os.path.join(_SYNC_DATA_DIR, "sync.log"),
        "EXCEL_LOG_PATH":       os.path.join(_SYNC_DATA_DIR, "sync_log.xlsx"),
        "MTIME_TOLERANCE_SEC":      int(_get("MTIME_TOLERANCE_SEC", "3")),
        "INJECT_UUID_FRONTMATTER":  _get("INJECT_UUID_FRONTMATTER", "true").lower() == "true",
        "STRIP_IMAGES":             _get("STRIP_IMAGES", "true").lower() == "true",
    }
    print(f"[System Log] 설정 로드 완료: {_UPNOTE_ROOT}")

except EnvironmentError as e:
    print(e)
    sys.exit(1)