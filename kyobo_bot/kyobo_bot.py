"""
kyobo_automation_v2.py
교보문고 장바구니 자동화 - 고도화 버전

[주요 개선사항]
1. 구조 리팩토링  : CONFIG 단일 진입점, 역할별 클래스 분리 (DataManager / Verifier / Reporter / Bot)
2. 재시도 로직    : @retry 데코레이터 + 지수 백오프(exponential backoff) 적용
3. 검증 고도화    : 문자열 유사도(SequenceMatcher) 기반 퍼지 매칭 (임계값 설정 가능)
4. 팝업 방어      : 장바구니 팝업/얼럿 자동 닫기 처리
5. 로깅           : logging 모듈 적용, 파일+콘솔 동시 출력
6. 실패 목록      : failed_report.xlsx 별도 저장 (수동 처리용)
7. 진행률 표시    : tqdm 진행 바
8. 중간 저장      : BATCH_SIZE 단위마다 중간 결과 자동 백업
"""

import time
import random
import logging
import functools
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path

import pandas as pd
from tqdm import tqdm
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException,
    ElementClickInterceptedException, WebDriverException
)

# ─────────────────────────────────────────
# [CONFIG] 통합 설정 — 이 블록만 수정하면 됩니다
# ─────────────────────────────────────────
CONFIG = {
    # 입력 파일
    "EXCEL_PATH"        : "book_list.xlsx",
    "SHEET_NAME"        : "Sheet1",

    # Chrome 디버그 연결
    "CHROME_DEBUG_PORT" : "9222",

    # 교보 검색 URL (keyword → title + publisher 조합)
    "SEARCH_BASE_URL"   : "https://search.kyobobook.co.kr/search?keyword={}&gbCode=TOT&target=total",

    # 컬럼명 매핑
    "COLUMNS": {
        "TITLE"     : "신청 책 제목\n(本タイトル)",
        "PUBLISHER" : "출판사\n(出版社)",
        "AUTHOR"    : "작가\n(作家)",
        "RESULT"    : "CHECK",
        "DETAIL"    : "상세사유",      # 신규: 상세 실패 사유
        "SIMILARITY": "유사도",        # 신규: 매칭 유사도 점수
    },

    # 딜레이 설정 (초)
    "MIN_SLEEP"         : 3,
    "MAX_SLEEP"         : 6,
    "BATCH_SIZE"        : 10,    # 배치 단위 (이 단위마다 장휴식 + 중간저장)
    "BATCH_SLEEP"       : 40,    # 배치 간 대기 시간(초)

    # 재시도 설정
    "MAX_RETRY"         : 3,     # 최대 재시도 횟수
    "RETRY_BASE_WAIT"   : 5,     # 재시도 기본 대기(초), 지수 백오프 적용

    # 유사도 임계값 (0.0 ~ 1.0) — 이 값 이상이면 매칭 성공으로 판정
    "TITLE_THRESHOLD"   : 0.75,
    "AUTHOR_THRESHOLD"  : 0.60,  # 저자는 표기 방식이 다양해 조금 낮게 설정

    # 출력 파일
    "OUTPUT_PATH"       : "result_report_final.xlsx",
    "FAILED_PATH"       : "failed_report.xlsx",
    "BACKUP_PREFIX"     : "backup_",
    "LOG_PATH"          : "automation.log",
}

# ─────────────────────────────────────────
# 로거 설정
# ─────────────────────────────────────────
def setup_logger(log_path: str) -> logging.Logger:
    logger = logging.getLogger("KyoboBot")
    logger.setLevel(logging.DEBUG)
    fmt = logging.Formatter("[%(asctime)s] %(levelname)s — %(message)s", "%H:%M:%S")

    # 콘솔 핸들러
    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    ch.setFormatter(fmt)

    # 파일 핸들러 (DEBUG 레벨까지 전부 기록)
    fh = logging.FileHandler(log_path, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt)

    logger.addHandler(ch)
    logger.addHandler(fh)
    return logger

logger = setup_logger(CONFIG["LOG_PATH"])


# ─────────────────────────────────────────
# 재시도 데코레이터 (지수 백오프)
# ─────────────────────────────────────────
def retry(max_attempts: int = 3, base_wait: float = 5.0, exceptions=(Exception,)):
    """
    지정 횟수만큼 재시도, 매번 대기 시간을 2배씩 늘림 (지수 백오프).
    모든 재시도 실패 시 마지막 예외를 그대로 raise.
    """
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            last_exc = None
            for attempt in range(1, max_attempts + 1):
                try:
                    return func(*args, **kwargs)
                except exceptions as e:
                    last_exc = e
                    wait = base_wait * (2 ** (attempt - 1))
                    logger.warning(
                        f"[재시도 {attempt}/{max_attempts}] {func.__name__} 실패 "
                        f"({type(e).__name__}) — {wait:.0f}초 후 재시도"
                    )
                    time.sleep(wait)
            raise last_exc
        return wrapper
    return decorator


# ─────────────────────────────────────────
# [1] DataManager — 엑셀 입출력 전담
# ─────────────────────────────────────────
class DataManager:
    def __init__(self, excel_path: str, sheet_name: str):
        self.excel_path = excel_path
        self.sheet_name = sheet_name
        self.df: pd.DataFrame = None

    def load(self) -> pd.DataFrame:
        self.df = pd.read_excel(self.excel_path, sheet_name=self.sheet_name)

        # 결과 컬럼 초기화 (없으면 생성)
        for col_key in ("RESULT", "DETAIL", "SIMILARITY"):
            col = CONFIG["COLUMNS"][col_key]
            if col not in self.df.columns:
                self.df[col] = ""

        logger.info(f"데이터 로드 완료 — 총 {len(self.df)}건")
        return self.df

    def save(self, path: str):
        self.df.to_excel(path, index=False)
        logger.info(f"저장 완료 → {path}")

    def save_backup(self, prefix: str, index: int):
        ts = datetime.now().strftime("%H%M%S")
        path = f"{prefix}{index}_{ts}.xlsx"
        self.save(path)
        logger.debug(f"중간 백업 저장 → {path}")

    def save_failed(self, path: str):
        col_result = CONFIG["COLUMNS"]["RESULT"]
        failed = self.df[self.df[col_result].str.startswith("실패", na=False)]
        if failed.empty:
            logger.info("실패 항목 없음 — 실패 리포트 미생성")
            return
        failed.to_excel(path, index=False)
        logger.info(f"실패 목록 저장 완료 → {path} ({len(failed)}건)")


# ─────────────────────────────────────────
# [2] Verifier — 유사도 기반 텍스트 검증 전담
# ─────────────────────────────────────────
class Verifier:
    @staticmethod
    def normalize(text: str) -> str:
        """공백·특수문자 제거 후 소문자화"""
        import re
        return re.sub(r"[\s\-·,.()\[\]《》『』「」【】]", "", text).lower()

    @staticmethod
    def similarity(a: str, b: str) -> float:
        """SequenceMatcher 기반 유사도 (0.0 ~ 1.0)"""
        na, nb = Verifier.normalize(a), Verifier.normalize(b)
        if not na or not nb:
            return 0.0
        # 짧은 쪽이 긴 쪽에 포함되면 1.0 처리 (부분 완전 일치)
        if na in nb or nb in na:
            return 1.0
        return SequenceMatcher(None, na, nb).ratio()

    @staticmethod
    def verify(
        excel_title: str, web_title: str,
        excel_author: str, web_author: str
    ) -> dict:
        """
        제목·저자 유사도를 계산하고 임계값 기준으로 매칭 여부를 반환.
        Returns:
            {
                "is_match": bool,
                "title_score": float,
                "author_score": float,
                "reason": str  # 실패 시 사유
            }
        """
        t_score = Verifier.similarity(excel_title, web_title)
        a_score = Verifier.similarity(excel_author, web_author)

        title_ok  = t_score >= CONFIG["TITLE_THRESHOLD"]
        author_ok = a_score >= CONFIG["AUTHOR_THRESHOLD"]
        is_match  = title_ok and author_ok

        reasons = []
        if not title_ok:
            reasons.append(f"제목 유사도 낮음({t_score:.2f})")
        if not author_ok:
            reasons.append(f"저자 유사도 낮음({a_score:.2f})")

        return {
            "is_match"    : is_match,
            "title_score" : t_score,
            "author_score": a_score,
            "similarity"  : round((t_score + a_score) / 2, 3),
            "reason"      : " / ".join(reasons) if reasons else "매칭 성공",
        }


# ─────────────────────────────────────────
# [3] Reporter — 콘솔 요약 출력 전담
# ─────────────────────────────────────────
class Reporter:
    @staticmethod
    def summary(df: pd.DataFrame):
        col = CONFIG["COLUMNS"]["RESULT"]
        total   = len(df)
        success = df[col].str.startswith("완료", na=False).sum()
        failed  = df[col].str.startswith("실패", na=False).sum()
        pending = total - success - failed

        logger.info("=" * 50)
        logger.info(f"  처리 결과 요약")
        logger.info(f"  전체  : {total}건")
        logger.info(f"  성공  : {success}건")
        logger.info(f"  실패  : {failed}건")
        logger.info(f"  미처리: {pending}건")
        logger.info("=" * 50)


# ─────────────────────────────────────────
# [4] KyoboBot — 핵심 자동화 로직
# ─────────────────────────────────────────
class KyoboBot:
    def __init__(self):
        chrome_options = Options()
        chrome_options.add_experimental_option(
            "debuggerAddress", f"127.0.0.1:{CONFIG['CHROME_DEBUG_PORT']}"
        )
        self.driver = webdriver.Chrome(options=chrome_options)
        self.wait   = WebDriverWait(self.driver, 10)
        self.verifier = Verifier()

    # ── 팝업/얼럿 방어 처리 ──────────────────
    def _dismiss_alert(self):
        try:
            alert = self.driver.switch_to.alert
            alert.accept()
            logger.debug("얼럿 자동 닫기 완료")
        except Exception:
            pass

    def _dismiss_popup(self):
        """레이어 팝업 닫기 버튼 탐색 후 클릭"""
        selectors = [
            ".layer_cart .btn_close",
            ".popup_wrap .btn_close",
            "button[class*='close']",
        ]
        for sel in selectors:
            try:
                btn = self.driver.find_element(By.CSS_SELECTOR, sel)
                if btn.is_displayed():
                    btn.click()
                    logger.debug(f"팝업 닫기 — {sel}")
                    time.sleep(0.3)
                    return
            except NoSuchElementException:
                continue

    # ── 검색 페이지 이동 ──────────────────────
    @retry(
    max_attempts=CONFIG["MAX_RETRY"],
    base_wait=CONFIG["RETRY_BASE_WAIT"],
    exceptions=(TimeoutException, WebDriverException)
)
    def _navigate_to_search(self, keyword: str):
        url = CONFIG["SEARCH_BASE_URL"].format(keyword)
        self.driver.get(url)
        # 기존: presence_of_element_located → 변경: visibility_of_element_located
        self.wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "li.prod_item")))

    # ── 첫 번째 검색 결과 파싱 ────────────────
    def _parse_first_result(self) -> dict:
        all_items = self.driver.find_elements(By.CSS_SELECTOR, "li.prod_item")
        if not all_items:
            raise NoSuchElementException("검색 결과 없음")
        first = all_items[0]
        title  = first.find_element(By.CSS_SELECTOR, "span[id^='cmdtName_']").text.strip()
        author = first.find_element(By.CSS_SELECTOR, ".prod_author_info .author.rep").text.strip()
        return {"element": first, "title": title, "author": author}

    # ── 장바구니 담기 ─────────────────────────
    @retry(
    max_attempts=CONFIG["MAX_RETRY"],
    base_wait=CONFIG["RETRY_BASE_WAIT"],
    exceptions=(ElementClickInterceptedException, WebDriverException)
)
    def _add_to_cart(self, prod_element):
    # button.btn_cart → a.btn_light_gray (장바구니 링크)
        cart_btn = prod_element.find_element(By.CSS_SELECTOR, "a.btn_light_gray")
        self.driver.execute_script(
            "arguments[0].scrollIntoView({block: 'center'});", cart_btn
        )
        time.sleep(0.5)
        cart_btn.click()
        time.sleep(1.0)
        self._dismiss_alert()
        self._dismiss_popup()

    # ── 인간 딜레이 ───────────────────────────
    def _human_delay(self):
        time.sleep(random.uniform(CONFIG["MIN_SLEEP"], CONFIG["MAX_SLEEP"]))

    # ── 단건 처리 ─────────────────────────────
    def process_one(self, index: int, row: pd.Series, df: pd.DataFrame):
        col = CONFIG["COLUMNS"]

        title     = str(row[col["TITLE"]]).strip()
        publisher = str(row[col["PUBLISHER"]]).strip()
        author    = str(row[col["AUTHOR"]]).strip()
        keyword   = f"{title} {publisher}"

        try:
            self._navigate_to_search(keyword)
            result = self._parse_first_result()

            verdict = self.verifier.verify(title, result["title"], author, result["author"])
            logger.debug(
                f"  제목 유사도: {verdict['title_score']:.2f} / "
                f"저자 유사도: {verdict['author_score']:.2f}"
            )

            df.at[index, col["SIMILARITY"]] = str(verdict["similarity"])

            if verdict["is_match"]:
                self._add_to_cart(result["element"])
                df.at[index, col["RESULT"]] = "완료(성공)"
                df.at[index, col["DETAIL"]] = f"제목:{verdict['title_score']:.2f} / 저자:{verdict['author_score']:.2f}"
                logger.info(f"  ✔ 성공 — {title}")
            else:
                df.at[index, col["RESULT"]] = f"실패(검증불일치)"
                df.at[index, col["DETAIL"]] = verdict["reason"]
                logger.warning(f"  ✘ 검증 실패 — {title} | 사유: {verdict['reason']}")

        except TimeoutException:
            df.at[index, col["RESULT"]] = "실패(검색결과없음)"
            df.at[index, col["DETAIL"]] = "페이지 로딩 타임아웃 (재시도 소진)"
            logger.error(f"  ✘ 타임아웃 — {title}")

        except NoSuchElementException as e:
            df.at[index, col["RESULT"]] = "실패(요소없음)"
            df.at[index, col["DETAIL"]] = str(e)[:80]
            logger.error(f"  ✘ 요소 없음 — {title}")

        except Exception as e:
            df.at[index, col["RESULT"]] = "실패(알수없는오류)"
            df.at[index, col["DETAIL"]] = str(e)[:80]
            logger.exception(f"  ✘ 알 수 없는 오류 — {title}")

    # ── 전체 처리 루프 ────────────────────────
    def run(self, df: pd.DataFrame, data_mgr: DataManager):
        total = len(df)
        logger.info(f"자동화 시작 — 총 {total}건")

        with tqdm(total=total, desc="진행", unit="권") as pbar:
            for index, row in df.iterrows():

                title_val = row[CONFIG["COLUMNS"]["TITLE"]]
                if pd.isna(title_val) or str(title_val).strip() == "" or str(title_val).strip() == "nan":
                    logger.info(f"[{index+1}/{total}] 제목 없음 — 스킵")
                    pbar.update(1)
                    continue

                # 배치 단위 장휴식 + 중간 저장
                if index > 0 and index % CONFIG["BATCH_SIZE"] == 0:
                    logger.info(
                        f"--- 배치 휴식 ({index}건 완료) "
                        f"— {CONFIG['BATCH_SLEEP']}초 대기 ---"
                    )
                    data_mgr.save_backup(CONFIG["BACKUP_PREFIX"], index)
                    time.sleep(CONFIG["BATCH_SLEEP"])

                logger.info(f"[{index+1}/{total}] {row[CONFIG['COLUMNS']['TITLE']]}")
                self.process_one(index, row, df)
                self._human_delay()
                pbar.update(1)

        logger.info("모든 처리 완료")


# ─────────────────────────────────────────
# 진입점
# ─────────────────────────────────────────
def main():
    logger.info("=" * 50)
    logger.info(" 교보문고 장바구니 자동화 v2 시작")
    logger.info("=" * 50)

    data_mgr = DataManager(CONFIG["EXCEL_PATH"], CONFIG["SHEET_NAME"])
    df = data_mgr.load()

    bot = KyoboBot()
    bot.run(df, data_mgr)

    # 최종 저장
    data_mgr.save(CONFIG["OUTPUT_PATH"])
    data_mgr.save_failed(CONFIG["FAILED_PATH"])

    # 요약 출력
    Reporter.summary(df)


if __name__ == "__main__":
    main()