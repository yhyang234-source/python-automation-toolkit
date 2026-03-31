import os
import re
import subprocess
import logging

# 시스템 상태와 오류를 명확히 추적하기 위한 로깅 설정 (방어 로직)
logging.basicConfig(level=logging.INFO, format='%(message)s')

# ==========================================
# 1. CONFIG: 설정값 관리 (GitHub 배포용 동적 경로 적용)
# ==========================================
CONFIG = {
    # '~' 기호가 실행하는 사람의 PC 홈 폴더로 자동 변환됩니다.
    "BASE_DIR": os.path.expanduser(r"~\AppData\Roaming\UpNote\UpNote Backup\F5p9KpV016SPRSpBH6y3uJ8NcFm2\Markdown\Synapse OS"),
    "OUTPUT_DIR": os.path.expanduser("C:\Obsidian\MyVault"),
}

# 파생 경로 자동 생성
CONFIG["NOTES_DIR"] = os.path.join(CONFIG["BASE_DIR"], "Notes")
CONFIG["NOTEBOOKS_DIR"] = os.path.join(CONFIG["BASE_DIR"], "notebooks")


# ==========================================
# 2. UTILITY: 보조 함수 모음
# ==========================================
def clean_filename(filename):
    """파일명으로 쓸 수 없는 특수문자 제거"""
    cleaned = re.sub(r'[\\/*?:"<>|]', "", filename).strip()
    return cleaned or "제목없음"

def get_lnk_target(lnk_path):
    """lnk 파일의 대상 경로 추출"""
    try:
        result = subprocess.run(
            ['powershell', '-command',
             f'$s=(New-Object -ComObject WScript.Shell).CreateShortcut("{lnk_path}"); $s.TargetPath'],
            capture_output=True, text=True, timeout=5
        )
        return result.stdout.strip()
    except Exception as e:
        logging.warning(f"  ⚠️ lnk 읽기 실패: {lnk_path} - {e}")
        return None

def get_title_from_md(md_path):
    """md 파일에서 제목 추출"""
    try:
        with open(md_path, 'r', encoding='utf-8') as f:
            for line in f.readlines()[:10]:
                line = line.strip()
                match = re.match(r'^#{1,6}\s+(.*)', line)
                if match:
                    title = match.group(1)
                    title = re.sub(r'#\w+\s*', '', title).strip()
                    if title:
                        return title
                if line.startswith("title:"):
                    return line.replace("title:", "").strip()
    except Exception as e:
        logging.error(f"  ⚠️ 파일 읽기 실패: {md_path} - {e}")
    return None

def copy_without_title(md_path, out_path):
    """첫 번째 헤딩 줄, 빈 줄 및 이미지 링크 제거 후 저장"""
    with open(md_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    filtered_lines = []
    title_removed = False
    for line in lines:
        if not title_removed and re.match(r'^#{1,6}\s+', line.strip()):
            title_removed = True
            continue

        cleaned_line = re.sub(r'!\[.*?\]\(.*?\)', '', line)

        if cleaned_line.strip() in ('', '<br>'):
            continue
        filtered_lines.append(cleaned_line)

    while filtered_lines and filtered_lines[0].strip() in ('', '<br>'):
        filtered_lines.pop(0)

    with open(out_path, 'w', encoding='utf-8') as f:
        f.writelines(filtered_lines)


# ==========================================
# 3. MAIN LOGIC: 핵심 변환 프로세스
# ==========================================
def process_notebooks():
    count = 0
    fail_count = 0
    skip_count = 0

    if not os.path.exists(CONFIG["NOTEBOOKS_DIR"]):
        logging.error(f"❌ 입력 폴더를 찾을 수 없습니다: {CONFIG['NOTEBOOKS_DIR']}")
        return

    for dirpath, dirnames, filenames in os.walk(CONFIG["NOTEBOOKS_DIR"]):
        lnk_files = [f for f in filenames if f.endswith(".md.lnk")]
        if not lnk_files:
            continue

        rel_path = os.path.relpath(dirpath, CONFIG["NOTEBOOKS_DIR"])
        out_dir = os.path.join(CONFIG["OUTPUT_DIR"], rel_path)
        os.makedirs(out_dir, exist_ok=True)

        for lnk_file in lnk_files:
            lnk_path = os.path.join(dirpath, lnk_file)
            uuid = lnk_file.replace(".md.lnk", "")
            md_path = os.path.join(CONFIG["NOTES_DIR"], f"{uuid}.md")

            if not os.path.exists(md_path):
                target = get_lnk_target(lnk_path)
                if target and os.path.exists(target):
                    md_path = target
                else:
                    logging.warning(f"  ❌ 원본 없음: {lnk_file}")
                    fail_count += 1
                    continue

            title = get_title_from_md(md_path)
            if not title:
                title = uuid  
                skip_count += 1

            safe_title = clean_filename(title)
            out_filename = f"{safe_title}.md"
            out_path = os.path.join(out_dir, out_filename)

            index = 1
            while os.path.exists(out_path):
                out_filename = f"{safe_title}({index}).md"
                out_path = os.path.join(out_dir, out_filename)
                index += 1

            try:
                copy_without_title(md_path, out_path)
                logging.info(f"✅ {rel_path} / {out_filename}")
                count += 1
            except Exception as e:
                logging.error(f"❌ 복사 실패 ({lnk_file}): {e}")
                fail_count += 1

    logging.info(f"\n✅ 성공: {count}개 / ⚠️ 제목없음: {skip_count}개 / ❌ 실패: {fail_count}개")
    logging.info(f"📁 결과물 위치: {CONFIG['OUTPUT_DIR']}")

if __name__ == "__main__":
    try:
        process_notebooks()
    except Exception as e:
        logging.critical(f"\n🚨 시스템 치명적 오류: {e}")