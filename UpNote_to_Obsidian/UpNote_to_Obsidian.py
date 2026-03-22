import os
import re
import subprocess

# ===================== 경로 설정 =====================
base_dir = r"C:\Users\yana_\Desktop\UpNote_Backup_for_Obsidian\F5p9KpV016SPRSpBH6y3uJ8NcFm2\Markdown\Synapse OS"
notes_dir = os.path.join(base_dir, "Notes")
notebooks_dir = os.path.join(base_dir, "notebooks")
output_dir = r"C:\Users\yana_\Desktop\UpNote_Obsidian_Output"  # 결과물 저장 위치
# =====================================================

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
        print(f"  ⚠️ lnk 읽기 실패: {lnk_path} - {e}")
        return None

def get_title_from_md(md_path):
    """md 파일에서 제목 추출 (# 헤딩 or YAML title:)"""
    try:
        with open(md_path, 'r', encoding='utf-8') as f:
            for line in f.readlines()[:10]:
                line = line.strip()
                match = re.match(r'^#{1,6}\s+(.*)', line)
                if match:
                    title = match.group(1)
                    title = re.sub(r'#\w+\s*', '', title).strip()  # #태그 제거
                    if title:
                        return title
                if line.startswith("title:"):
                    return line.replace("title:", "").strip()
    except Exception as e:
        print(f"  ⚠️ 파일 읽기 실패: {md_path} - {e}")
    return None

def copy_without_title(md_path, out_path):
    """첫 번째 헤딩 줄, 빈 줄/br, 이미지 링크 제거 후 저장"""
    with open(md_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    # 첫 번째 헤딩 줄 한 번만 제거
    filtered_lines = []
    title_removed = False
    for line in lines:
        # 첫 헤딩 제거
        if not title_removed and re.match(r'^#{1,6}\s+', line.strip()):
            title_removed = True
            continue

        # 이미지 링크 제거: ![...](...)
        cleaned_line = re.sub(r'!\[.*?\]\(.*?\)', '', line)

        # 이미지 제거 후 줄이 공백/br만 남으면 스킵
        if cleaned_line.strip() in ('', '<br>'):
            continue

        filtered_lines.append(cleaned_line)

    # 앞쪽 빈 줄 / <br> 제거
    while filtered_lines and filtered_lines[0].strip() in ('', '<br>'):
        filtered_lines.pop(0)

    with open(out_path, 'w', encoding='utf-8') as f:
        f.writelines(filtered_lines)

def process_notebooks():
    count = 0
    fail_count = 0
    skip_count = 0

    for dirpath, dirnames, filenames in os.walk(notebooks_dir):
        lnk_files = [f for f in filenames if f.endswith(".md.lnk")]
        if not lnk_files:
            continue

        # notebooks/ 이후 상대 경로 → output 폴더 구조 재현
        rel_path = os.path.relpath(dirpath, notebooks_dir)
        out_dir = os.path.join(output_dir, rel_path)
        os.makedirs(out_dir, exist_ok=True)

        for lnk_file in lnk_files:
            lnk_path = os.path.join(dirpath, lnk_file)

            # UUID 추출 후 Notes/UUID.md 경로 구성
            uuid = lnk_file.replace(".md.lnk", "")
            md_path = os.path.join(notes_dir, f"{uuid}.md")

            # Notes에 없으면 lnk 대상에서 직접 찾기
            if not os.path.exists(md_path):
                target = get_lnk_target(lnk_path)
                if target and os.path.exists(target):
                    md_path = target
                else:
                    print(f"  ❌ 원본 없음: {lnk_file}")
                    fail_count += 1
                    continue

            title = get_title_from_md(md_path)
            if not title:
                title = uuid  # 제목 없으면 UUID 그대로 사용
                skip_count += 1

            safe_title = clean_filename(title)
            out_filename = f"{safe_title}.md"
            out_path = os.path.join(out_dir, out_filename)

            # 중복 파일명 처리
            index = 1
            while os.path.exists(out_path):
                out_filename = f"{safe_title}({index}).md"
                out_path = os.path.join(out_dir, out_filename)
                index += 1

            try:
                copy_without_title(md_path, out_path)
                print(f"✅ {rel_path} / {out_filename}")
                count += 1
            except Exception as e:
                print(f"❌ 복사 실패 ({lnk_file}): {e}")
                fail_count += 1

    print(f"\n✅ 성공: {count}개 / ⚠️ 제목없음(UUID사용): {skip_count}개 / ❌ 실패: {fail_count}개")
    print(f"📁 결과물 위치: {output_dir}")

if __name__ == "__main__":
    process_notebooks()
