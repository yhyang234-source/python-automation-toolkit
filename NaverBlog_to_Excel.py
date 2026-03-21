import pandas as pd
import logging
import glob
import os
import re
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextBox, LTTextLine, LTChar

# ==========================================
# 1. CONFIG
# ==========================================
CONFIG = {
    "INPUT_DIR": ".",
    "OUTPUT_PATH": "extracted_books_all.csv",
    "COLUMNS": ["블로그 제목", "책 이름", "책 저자", "블로그 날짜", "내용", "url"],
}

TITLE_SIZE  = 18.8              # 블로그 제목 (청크 경계 신호)
COLOR_BLUE  = (0.0, 0.0, 1.0)  # 네이버 도서 모듈
COLOR_BLACK = 0.0               # 본문
COLOR_GRAY1 = 0.2               # 카테고리 라벨 → 제거
COLOR_GRAY2 = 0.6               # 날짜 / URL / 꼬리말 → 메타 후 제거

BLUE_LABEL_RE  = re.compile(r'^(저자|출판|발매){1,2}$')
RELEASE_DATE_RE = re.compile(r'^\d{4}\.\d{2}\.\d{2}')  # 발매일 패턴

logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')


# ==========================================
# 2. Utility
# ==========================================
def get_char_props(line):
    for char in line:
        if isinstance(char, LTChar):
            return round(char.size, 1), getattr(char.graphicstate, 'ncolor', None)
    return None, None

def is_blue(color):
    return color == COLOR_BLUE or color == list(COLOR_BLUE)

def is_title_line(size, color):
    return size == TITLE_SIZE and color == COLOR_BLACK


# ==========================================
# 3. 요소 추출
# ==========================================
def get_text_elements_from_pdf(file_path):
    elements = []
    try:
        for page_layout in extract_pages(file_path):
            for element in page_layout:
                if not isinstance(element, LTTextBox):
                    continue
                block_has_text = False
                for line in element:
                    if not isinstance(line, LTTextLine):
                        continue
                    text = line.get_text().strip()
                    if not text or re.match(r'^\[Image \d+\]$', text):
                        continue
                    size, color = get_char_props(line)
                    elements.append({"text": text, "size": size, "color": color})
                    block_has_text = True
                if block_has_text:
                    elements.append({"text": "<BR>", "size": None, "color": None})
    except Exception as e:
        logging.error(f"[{os.path.basename(file_path)}] 읽기 오류: {e}")
    return elements


# ==========================================
# 4. 청킹: size 18.8 제목 블록 기준
# ==========================================
def split_into_chunks(elements):
    chunks, current = [], []
    for el in elements:
        if is_title_line(el["size"], el["color"]):
            already_has_title = any(is_title_line(e["size"], e["color"]) for e in current)
            if already_has_title and current:
                chunks.append(current)
                current = []
        current.append(el)
    if current:
        chunks.append(current)
    return chunks


# ==========================================
# 5. 파란 블록 → 책 이름 / 저자
#
# 핵심 규칙: 발매일(YYYY.MM.DD.) 이후 파란 블록은 추천링크이므로 전량 무시
#
# 케이스A: [제목x2] [저자명] [출판사] [저자저자] [출판출판] [발매발매] [날짜]
# 케이스B: [제목x2] [저자저자] [출판출판] [발매발매] [저자명] [출판사] [날짜]
# ==========================================
def extract_book_meta(blue_seq):
    # 발매일 이후 블록 제거
    cutoff = len(blue_seq)
    for i, b in enumerate(blue_seq):
        if RELEASE_DATE_RE.match(b):
            cutoff = i + 1  # 날짜 자체는 포함, 이후는 버림
            break
    blue_seq = blue_seq[:cutoff]

    def find_label(labels):
        for i, b in enumerate(blue_seq):
            if b in labels:
                return i
        return None

    def non_label(seq):
        return [b for b in seq if not BLUE_LABEL_RE.match(b)]

    author_li  = find_label(("저자저자", "저자"))
    publish_li = find_label(("출판출판", "출판"))
    release_li = find_label(("발매발매", "발매"))

    if author_li is None:
        return "", ""

    # 책 이름: 저자 레이블 이전, x2 합본 해제 + 중복 제거
    before = non_label(blue_seq[:author_li])
    seen = []
    for b in before:
        half = len(b) // 2
        if len(b) % 2 == 0 and b[:half] == b[half:]:
            b = b[:half]
        if b not in seen:
            seen.append(b)
    book_name = seen[0] if seen else ""

    # 저자: 케이스A(저자라벨~출판라벨 사이) vs 케이스B(발매라벨 이후)
    author = ""
    if publish_li is not None:
        between = non_label(blue_seq[author_li + 1 : publish_li])
        if between:
            author = between[0]  # 케이스A
        elif release_li is not None:
            after_release = non_label(blue_seq[release_li + 1 : cutoff])
            candidates = [b for b in after_release if not RELEASE_DATE_RE.match(b)]
            if candidates:
                author = candidates[0]  # 케이스B
    else:
        after = non_label(blue_seq[author_li + 1:])
        if after:
            author = after[0]

    return book_name, author


# ==========================================
# 6. 청크 파싱
# ==========================================
def parse_chunk(chunk):
    book_info = {col: "" for col in CONFIG["COLUMNS"]}
    content_lines = []

    # 블로그 제목: 첫 번째 size 18.8 텍스트
    for el in chunk:
        if is_title_line(el["size"], el["color"]):
            book_info["블로그 제목"] = el["text"]
            break

    # 책 이름 / 저자
    blue_seq = [el["text"] for el in chunk
                if el["text"] != "<BR>" and is_blue(el["color"])]
    book_info["책 이름"], book_info["책 저자"] = extract_book_meta(blue_seq)

    # URL / 날짜
    for el in chunk:
        if el["text"] == "<BR>" or el["color"] != COLOR_GRAY2:
            continue
        text = el["text"]
        m_url  = re.search(r'https?://blog\.naver\.com/\S+', text)
        m_date = re.search(r'\d{4}/\d{2}/\d{2}\s\d{2}:\d{2}', text)
        if m_url:
            book_info["url"] = m_url.group()
        elif m_date:
            book_info["블로그 날짜"] = m_date.group()

    # 본문: 검정(10.5)만
    for el in chunk:
        text, size, color = el["text"], el["size"], el["color"]
        if text == "<BR>":
            content_lines.append("\n")
            continue
        if is_title_line(size, color):   continue  # 제목 제거
        if is_blue(color):               continue  # 도서 모듈 제거
        if color in (COLOR_GRAY1, COLOR_GRAY2): continue  # 회색 제거
        if "©" in text or "unsplash" in text.lower() or text.strip().startswith("출처"):
            continue

        content_lines.append(text)

    raw = " ".join(content_lines)
    book_info["내용"] = re.sub(r'\s*\n\s*', '\n', raw).strip()
    return book_info


# ==========================================
# 7. 본문 연속 중복 줄 제거
# ==========================================
def deduplicate_lines(content):
    lines = content.split('\n')
    result, i = [], 0
    while i < len(lines):
        if (i + 1 < len(lines)
                and lines[i].strip() == lines[i+1].strip()
                and lines[i].strip()):
            result.append(lines[i])
            i += 2
        else:
            result.append(lines[i])
            i += 1
    return '\n'.join(result)


# ==========================================
# 8. Integration & Main
# ==========================================
def parse_all_books(elements):
    books = []
    for chunk in split_into_chunks(elements):
        info = parse_chunk(chunk)
        if info["url"]:
            info["내용"] = deduplicate_lines(info["내용"])
            books.append(info)
    return books

def save_to_csv(data, output_path):
    try:
        df = pd.DataFrame(data, columns=CONFIG["COLUMNS"])
        df.to_csv(output_path, index=False, encoding='utf-8-sig')
        logging.info(f"✅ 총 {len(data)}건 저장 완료 → {output_path}")
    except Exception as e:
        logging.error(f"저장 오류: {e}")

def main():
    logging.info("🚀 PDFtoCSV v2 시작 (size+색상 구조 기반)")
    pdf_files = glob.glob(os.path.join(CONFIG["INPUT_DIR"], "*.pdf"))
    if not pdf_files:
        logging.warning("PDF 파일 없음")
        return
    all_data = []
    for file_path in pdf_files:
        logging.info(f"처리 중: {os.path.basename(file_path)}")
        elements = get_text_elements_from_pdf(file_path)
        if elements:
            parsed = parse_all_books(elements)
            logging.info(f"  → {len(parsed)}건")
            all_data.extend(parsed)
    if all_data:
        save_to_csv(all_data, CONFIG["OUTPUT_PATH"])

if __name__ == "__main__":
    main()