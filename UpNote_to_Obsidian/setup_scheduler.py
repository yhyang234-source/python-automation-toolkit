"""
setup_scheduler.py — Windows 작업 스케줄러 등록/해제
========================================================
관리자 권한 없이도 현재 사용자 세션에 작업을 등록합니다.

사용법:
  python setup_scheduler.py --register    # 스케줄러 등록
  python setup_scheduler.py --unregister  # 스케줄러 해제
  python setup_scheduler.py --status      # 등록 상태 확인
"""

import os
import sys
import subprocess
import argparse
from pathlib import Path

# ──────────────────────────────────────────────
# 설정
# ──────────────────────────────────────────────
TASK_NAME = "UpNote_Obsidian_Sync"
SCRIPT_DIR = Path(__file__).parent.resolve()
SYNC_SCRIPT = SCRIPT_DIR / "sync_engine.py"

# Python 실행 파일 경로 (현재 가상환경/인터프리터 그대로 사용)
PYTHON_EXE = sys.executable

# 동기화 간격 (분)
INTERVAL_MINUTES = 30


def run_ps(command: str) -> tuple[int, str, str]:
    """PowerShell 명령 실행 후 (returncode, stdout, stderr) 반환"""
    result = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", command],
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    return result.returncode, result.stdout.strip(), result.stderr.strip()


def register():
    print(f"[+] 스케줄러 작업 등록: '{TASK_NAME}'")
    print(f"    스크립트: {SYNC_SCRIPT}")
    print(f"    Python:   {PYTHON_EXE}")
    print(f"    주기:     매 {INTERVAL_MINUTES}분\n")

    # 기존 작업 삭제 (재등록 시 충돌 방지)
    run_ps(f'Unregister-ScheduledTask -TaskName "{TASK_NAME}" -Confirm:$false -ErrorAction SilentlyContinue')

    ps_script = f"""
$action  = New-ScheduledTaskAction `
    -Execute "{PYTHON_EXE}" `
    -Argument '"{SYNC_SCRIPT}"' `
    -WorkingDirectory "{SCRIPT_DIR}"

$trigger = New-ScheduledTaskTrigger `
    -RepetitionInterval (New-TimeSpan -Minutes {INTERVAL_MINUTES}) `
    -Once -At (Get-Date).AddMinutes(1) `
    -RepetitionDuration ([TimeSpan]::MaxValue)

$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 5) `
    -MultipleInstances IgnoreNew `
    -StartWhenAvailable

Register-ScheduledTask `
    -TaskName "{TASK_NAME}" `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -RunLevel Limited `
    -Description "UpNote-Obsidian 자동 동기화 (매 {INTERVAL_MINUTES}분)"

Write-Output "OK"
"""
    code, out, err = run_ps(ps_script)
    if code == 0 and "OK" in out:
        print(f"✅ 스케줄러 등록 완료!")
        print(f"   작업 관리자에서 '{TASK_NAME}' 으로 확인 가능합니다.")
        print(f"   taskschd.msc 실행 후 '작업 스케줄러 라이브러리'에서 찾으세요.\n")
    else:
        print(f"❌ 등록 실패 (code={code})")
        if err:
            print(f"   오류: {err}")
        print("\n   수동 등록 방법: taskschd.msc → 기본 작업 만들기")


def unregister():
    print(f"[-] 스케줄러 작업 해제: '{TASK_NAME}'")
    code, out, err = run_ps(
        f'Unregister-ScheduledTask -TaskName "{TASK_NAME}" -Confirm:$false; Write-Output "OK"'
    )
    if code == 0:
        print(f"✅ 스케줄러 작업 '{TASK_NAME}' 해제 완료")
    else:
        print(f"⚠️  해제 실패 또는 등록된 작업 없음")
        if err:
            print(f"   {err}")


def status():
    print(f"[?] 스케줄러 작업 상태 확인: '{TASK_NAME}'\n")
    code, out, err = run_ps(
        f'Get-ScheduledTask -TaskName "{TASK_NAME}" | '
        f'Select-Object TaskName, State, @{{n="NextRun";e={{(Get-ScheduledTaskInfo -TaskName "{TASK_NAME}").NextRunTime}}}} | '
        f'Format-List'
    )
    if code == 0 and out:
        print(out)
    else:
        print(f"  등록된 작업 없음 (python setup_scheduler.py --register 로 등록)")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Windows 작업 스케줄러 관리")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--register",   action="store_true", help="스케줄러 등록")
    group.add_argument("--unregister", action="store_true", help="스케줄러 해제")
    group.add_argument("--status",     action="store_true", help="등록 상태 확인")
    args = parser.parse_args()

    if args.register:
        register()
    elif args.unregister:
        unregister()
    elif args.status:
        status()
