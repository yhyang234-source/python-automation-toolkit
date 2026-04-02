"""
csv_to_obsidian.py
독서나이 프로젝트 CSV → Obsidian Markdown 변환기
컬럼: Date, Category, Category_Depth1, Title, Author, Context, Blog_url

실행 방법:
    python csv_to_obsidian.py

출력:
    Input 폴더 내 Output\ 폴더에 md 파일 생성
"""

import pandas as pd
from pathlib import Path
import re
import sys

# ──────────────────────────────────────────────
# 설정
# ──────────────────────────────────────────────
INPUT_DIR  = Path(r"C:\Python\python-automation-toolkit\UpNote_to_Obsidian\Input")
INPUT_FILE = INPUT_DIR / "extracted_books_all_fix.csv"
OUTPUT_DIR = INPUT_DIR / "Output"

# CSV 인코딩 우선순위 (CP949 → UTF-8-sig → UTF-8)
ENCODINGS = ["cp949", "utf-8-sig", "utf-8"]

# 파일명에 사용할 수 없는 문자 제거 패턴
INVALID_CHARS = r'[\\/:*?"<>|]'


# ──────────────────────────────────────────────
# 유틸 함수
# ──────────────────────────────────────────────
def safe_filename(name: str) -> str:
    """파일명으로 사용 불가한 문자 제거"""
    return re.sub(INVALID_CHARS, "", str(name)).strip()


def safe_str(value) -> str:
    """NaN, None → 빈 문자열"""
    if pd.isna(value):
        return ""
    return str(value).strip()


def load_csv(path: Path) -> pd.DataFrame:
    """인코딩 자동 감지하여 CSV 로드"""
    for enc in ENCODINGS:
        try:
            df = pd.read_csv(path, encoding=enc)
            print(f"  [OK] 인코딩: {enc}, 총 {len(df)}행 로드")
            return df
        except (UnicodeDecodeError, Exception):
            continue
    print("[ERROR] CSV 파일을 읽을 수 없습니다. 인코딩을 확인하세요.")
    sys.exit(1)


def build_markdown(row: pd.Series) -> str:
    """row 1개 → md 문자열 생성"""
    title          = safe_str(row.get("Title", ""))
    author         = safe_str(row.get("Author", ""))
    category       = safe_str(row.get("Category", ""))
    category_d1    = safe_str(row.get("Category_Depth1", ""))
    date_read      = safe_str(row.get("Date", ""))
    blog_url       = safe_str(row.get("Blog_url", ""))
    context        = safe_str(row.get("Context", ""))

    # YAML 태그 구성
    tag_lines = ["  - book"]
    if category:
        tag_lines.append(f'  - "{category}"')
    if category_d1 and category_d1 != category:
        tag_lines.append(f'  - "{category_d1}"')
    tags_block = "\n".join(tag_lines)

    # 블로그 링크 처리
    blog_link = f"[리뷰 보기]({blog_url})" if blog_url else "_(없음)_"

    md = f"""---
title: "{title}"
author: "{author}"
category: "{category}"
category_depth1: "{category_d1}"
date_read: {date_read}
blog_url: "{blog_url}"
tags:
{tags_block}
keywords: []
related: []
---

# {title}

> [!info] 기본 정보
> **저자** :: {author}
> **분류** :: {category} / {category_d1}
> **독서일** :: {date_read}
> **블로그** :: {blog_link}

## 📝 리뷰

{context}

## 🔗 연결

### 직접 연결한 책

- 

## 💡 메모

"""
    return md


# ──────────────────────────────────────────────
# 메인
# ──────────────────────────────────────────────
def main():
    print("=" * 50)
    print("독서나이 CSV → Obsidian MD 변환기")
    print("=" * 50)

    # 입력 파일 확인
    if not INPUT_FILE.exists():
        print(f"[ERROR] 파일 없음: {INPUT_FILE}")
        sys.exit(1)

    # 출력 폴더 생성
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    print(f"  [OK] 출력 폴더: {OUTPUT_DIR}")

    # CSV 로드
    df = load_csv(INPUT_FILE)

    # 컬럼 확인
    required = {"Date", "Category", "Category_Depth1", "Title", "Author", "Context", "Blog_url"}
    missing = required - set(df.columns)
    if missing:
        print(f"[WARN] 없는 컬럼: {missing}")
        print(f"       실제 컬럼: {list(df.columns)}")

    # 변환 실행
    success, skipped, duplicate = 0, 0, []
    used_names: set[str] = set()

    for idx, row in df.iterrows():
        title = safe_str(row.get("Title", ""))

        if not title:
            print(f"  [SKIP] {idx+1}행: Title 없음")
            skipped += 1
            continue

        # 중복 파일명 처리 (동명 도서 있을 경우 _2, _3 suffix)
        base_name = safe_filename(title)
        file_name = base_name
        counter = 2
        while file_name in used_names:
            file_name = f"{base_name}_{counter}"
            counter += 1
            duplicate.append(title)
        used_names.add(file_name)

        # md 생성 및 저장
        md_content = build_markdown(row)
        out_path = OUTPUT_DIR / f"{file_name}.md"
        out_path.write_text(md_content, encoding="utf-8")
        success += 1

    # 결과 리포트
    print()
    print("─" * 40)
    print(f"  변환 완료: {success}개")
    print(f"  건너뜀   : {skipped}개 (Title 없는 행)")
    if duplicate:
        print(f"  중복 제목 : {len(duplicate)}개 → suffix 자동 부여")
        for d in duplicate:
            print(f"    - {d}")
    print(f"  출력 경로 : {OUTPUT_DIR}")
    print("─" * 40)
    print("완료! Obsidian Vault에 Output 폴더를 연결하세요.")


if __name__ == "__main__":
    main()