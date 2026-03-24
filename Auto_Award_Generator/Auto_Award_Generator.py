import os
import shutil
import re
import time
import platform
from datetime import datetime
import win32com.client as win32
import openpyxl

# ==========================================
# 1. CONFIG: 모든 설정값을 한곳에서 관리
# ==========================================
CONFIG = {
    "EXCEL_FILE": "Award_Data.xlsx",
    "SHEET_NAME": "Sheet1",
    "TEMPLATE_FILE": "Award_Template.hwp",
    "OUTPUT_DIR": "결과물",
    "HANGEUL_VISIBLE": True,
    "MAX_PAGE_LIMIT": 1        
}

# ==========================================
# 2. UTILITY: 보조 함수 모음
# ==========================================
def setup_environment():
    print(f"▶ 시스템 정보: Python {platform.python_version()} 환경")
    print("====================================================")
    print(f"🚨 [확인] '{CONFIG['EXCEL_FILE']}'이 열려있다면 반드시 닫아주세요.")
    print("====================================================")

    if not os.path.exists(CONFIG["OUTPUT_DIR"]):
        os.makedirs(CONFIG["OUTPUT_DIR"])
        print(f"📁 '{CONFIG['OUTPUT_DIR']}' 폴더를 생성했습니다.")

def format_date(raw_date):
    if isinstance(raw_date, datetime):
        return raw_date.strftime("%Y年%m月%d日")
    elif isinstance(raw_date, str) and raw_date.strip():
        try:
            dt = datetime.strptime(raw_date.strip(), "%Y-%m-%d")
            return dt.strftime("%Y年%m月%d日")
        except ValueError:
            return raw_date
    return None

def make_filename_safe(text):
    if not text: return "N/A"
    return re.sub(r'[\/:*?"<>|]', '_', str(text))

# ==========================================
# 3. MAIN LOGIC: 핵심 자동화 프로세스
# ==========================================
def main():
    setup_environment()

    current_dir = os.getcwd()
    excel_path = os.path.join(current_dir, CONFIG["EXCEL_FILE"])
    template_path = os.path.join(current_dir, CONFIG["TEMPLATE_FILE"])
    output_dir_path = os.path.join(current_dir, CONFIG["OUTPUT_DIR"])

    if not os.path.exists(excel_path) or not os.path.exists(template_path):
        print("❌ 오류: 엑셀 또는 템플릿 파일을 찾을 수 없습니다.")
        return

    try:
        wb = openpyxl.load_workbook(excel_path)
        ws = wb[CONFIG["SHEET_NAME"]]
    except Exception as e:
        print(f"❌ 엑셀 로드 에러: {e}")
        return

    header = {cell.value: idx + 1 for idx, cell in enumerate(ws[1])}
    
    required_cols = ["문서번호", "부서", "이름", "표창명", "내용", "날짜", "대표이사명", "결과"]
    for col in required_cols:
        if col not in header:
            print(f"❌ 오류: 엑셀 첫 줄에 '{col}' 컬럼이 없습니다.")
            return

    print("⏳ 한글 프로그램을 실행 중입니다...")
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.XHwpWindows.Item(0).Visible = CONFIG["HANGEUL_VISIBLE"]

    counts = {"success": 0, "fail": 0, "skip": 0}

    for row_idx in range(2, ws.max_row + 1):
        status = str(ws.cell(row=row_idx, column=header["결과"]).value or "").strip()
        if status == "성공":
            counts["skip"] += 1
            continue

        doc_no = str(ws.cell(row=row_idx, column=header["문서번호"]).value or "").strip()
        
        # [핵심 변경] 문서번호가 비어있으면 데이터의 끝으로 간주하고 전체 반복 중단
        if not doc_no:
            print(f"\n🛑 [알림] {row_idx}행에서 빈 문서번호를 발견했습니다. 작업을 종료합니다.")
            break 

        dept = str(ws.cell(row=row_idx, column=header["부서"]).value or "").strip()
        name = str(ws.cell(row=row_idx, column=header["이름"]).value or "").strip()
        award = str(ws.cell(row=row_idx, column=header["표창명"]).value or "").strip()
        content = str(ws.cell(row=row_idx, column=header["내용"]).value or "").strip()
        ceo = str(ws.cell(row=row_idx, column=header["대표이사명"]).value or "").strip()
        raw_date = ws.cell(row=row_idx, column=header["날짜"]).value
        formatted_date = format_date(raw_date)

        file_name = f"{make_filename_safe(doc_no)}_{make_filename_safe(dept)}_{make_filename_safe(name)}_표창장.hwp"
        save_path = os.path.join(output_dir_path, file_name)

        print(f"▶ 처리 중: {file_name}", end=" ... ")

        try:
            shutil.copy(template_path, save_path)
            hwp.Open(save_path, "HWP", "forceopen:true")
            
            hwp.PutFieldText("doc_no", doc_no)
            hwp.PutFieldText("name", name)
            hwp.PutFieldText("department", dept)
            hwp.PutFieldText("award_name", award)
            
            if content: hwp.PutFieldText("content", content)
            if formatted_date: hwp.PutFieldText("award_date", formatted_date)
            if ceo: hwp.PutFieldText("ceo", ceo)

            if hwp.PageCount > CONFIG["MAX_PAGE_LIMIT"]:
                hwp.Clear(1)
                os.remove(save_path)
                raise Exception(f"{CONFIG['MAX_PAGE_LIMIT']}페이지 초과")

            hwp.Save()
            hwp.Clear(1)
            ws.cell(row=row_idx, column=header["결과"]).value = "성공"
            counts["success"] += 1
            print("✅")

        except Exception as e:
            ws.cell(row=row_idx, column=header["결과"]).value = f"실패({str(e)[:15]})"
            counts["fail"] += 1
            print(f"❌ ({e})")

    hwp.Quit()
    try:
        wb.save(excel_path)
        print(f"\n🎉 완료! (성공:{counts['success']}, 실패:{counts['fail']}, 스킵:{counts['skip']})")
    except:
        print("\n❌ 엑셀 저장 실패 (파일을 닫고 다시 실행하세요)")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n🚨 치명적 오류: {e}")
    finally:
        input("\n엔터(Enter)를 누르면 창이 닫힙니다...")