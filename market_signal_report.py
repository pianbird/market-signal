import os
import sys
import time
import json
import html
import argparse
import datetime
import subprocess
import requests
import xml.etree.ElementTree as ET
import pandas as pd
from typing import Dict, Any, List, Optional, Tuple

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==============================================================================
# 1. 공통 환경 초기화 및 인코딩/경로 안전장치
# ==============================================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE_DIR)

# 콘솔 UTF-8 인코딩 보장
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
        sys.stderr.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
    except Exception:
        pass

try:
    from dotenv import load_dotenv
    load_dotenv(os.path.join(BASE_DIR, ".env"), override=True)
except ImportError:
    pass

# 환경변수 로드
MY_BOT_TOKEN = os.environ.get("MY_BOT_TOKEN", "")
MY_CHAT_ID = os.environ.get("MY_CHAT_ID", "")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
GEMINI_MODEL = os.environ.get("GEMINI_MODEL", "gemini-2.5-flash")

NOTION_API_KEY = os.environ.get("NOTION_API_KEY", "")
NOTION_DATABASE_ID = os.environ.get("NOTION_DATABASE_ID", "")
NOTION_VERSION = "2022-06-28"

# 구글 드라이브 환경변수
GDRIVE_SERVICE_ACCOUNT_FILE = os.environ.get("GDRIVE_SERVICE_ACCOUNT_FILE", "service_account.json")
GDRIVE_SERVICE_ACCOUNT_JSON = os.environ.get("GDRIVE_SERVICE_ACCOUNT_JSON", "")
GDRIVE_FOLDER_ID = os.environ.get("GDRIVE_FOLDER_ID", "")

# 구글 드라이브 클라이언트 라이브러리 선택적 로드
try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build as gdrive_build
    from googleapiclient.http import MediaFileUpload
    GDRIVE_AVAILABLE = True
except ImportError:
    GDRIVE_AVAILABLE = False

NOTION_HEADERS = {
    "Authorization": f"Bearer {NOTION_API_KEY}",
    "Content-Type": "application/json",
    "Notion-Version": NOTION_VERSION,
}

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
}

# 분석 상수 정의
ANALYSIS_YEARS = 10
FETCH_COUNT_CLOSING = 2600
FETCH_COUNT_INTRADAY = 300

EMA40_SLOPE_THRESHOLD = 0.15      # 5일 EMA40 기울기 임계값 (±0.15%)
EMA40_OVERHEAT_DISPARITY = 104.0  # 이격도 104% 초과는 단기 과열 주의

SCHEDULE_TIMES = ["10:00", "13:00", "15:00"]
STATE_FILE = os.path.join(BASE_DIR, "position_state.json")

# 8대 ETF 유니버스 표준 정의
TARGET_ETFS = [
    {"code": "069500", "name": "KODEX 200", "group": "KOSPI 200", "market": "KOSPI", "type": "정방향"},
    {"code": "102110", "name": "TIGER 200", "group": "KOSPI 200", "market": "KOSPI", "type": "정방향"},
    {"code": "114800", "name": "KODEX 인버스", "group": "KOSPI 200 인버스", "market": "KOSPI", "type": "인버스"},
    {"code": "123310", "name": "TIGER 인버스", "group": "KOSPI 200 인버스", "market": "KOSPI", "type": "인버스"},
    {"code": "229200", "name": "KODEX 코스닥150", "group": "KOSDAQ 150", "market": "KOSDAQ", "type": "정방향"},
    {"code": "232080", "name": "TIGER 코스닥150", "group": "KOSDAQ 150", "market": "KOSDAQ", "type": "정방향"},
    {"code": "251340", "name": "KODEX 코스닥인버스", "group": "KOSDAQ 150 인버스", "market": "KOSDAQ", "type": "인버스"},
    {"code": "250780", "name": "TIGER 코스닥인버스", "group": "KOSDAQ 150 인버스", "market": "KOSDAQ", "type": "인버스"},
]
ETF_MAP: Dict[str, Dict[str, Any]] = {e["code"]: e for e in TARGET_ETFS}



# ==============================================================================
# 2. 공통 유틸리티 (개장일 판별 가드, 거래량 보정, 텔레그램, Git 백업, 노션 API)
# ==============================================================================
def is_market_open_today(mode: str = "closing") -> bool:
    """
    한국 증시 개장일(영업일) 판별 가드
    1. 주말 1차 차단: 토요일/일요일인 경우 즉시 False 반환.
    2. 실시간 네이버 차트 일봉 및 지수 API 검증:
       - 대표 ETF(069500 KODEX 200) 일봉 API 호출
         - 장마감 모드 ('closing'): 15:30 이후 실행되었으나 오늘 날짜 봉 데이터가 없으면 휴장일로 판정 (False)
         - 장 시작 전 / 장중 모드 ('daemon', 'intraday'): 네이버 실시간 국내 지수(KOSPI) API의 marketStatus 및 최신 거래일자를 교차 점검하여 휴장일인 경우 False 반환
    3. 통신 에러 등 예외 발생 시 기본값 True 유지 (경고 로그 출력)
    """
    now = datetime.datetime.now()

    # 1. 주말 1차 차단 (토=5, 일=6)
    if now.weekday() >= 5:
        return False

    today_ymd = now.strftime("%Y%m%d")
    today_dash = now.strftime("%Y-%m-%d")

    try:
        # 대표 ETF (069500) 일봉 API 확인
        fchart_url = "https://fchart.stock.naver.com/sise.nhn?symbol=069500&timeframe=day&count=1&requestType=0"
        res_fchart = requests.get(fchart_url, headers=HEADERS, timeout=5)
        has_today_candle = False
        latest_candle_date = ""

        if res_fchart.status_code == 200:
            root = ET.fromstring(res_fchart.text)
            item = root.find(".//item")
            if item is not None:
                data_attr = item.get("data", "")
                if data_attr:
                    latest_candle_date = data_attr.split("|")[0].strip()
                    if latest_candle_date == today_ymd:
                        has_today_candle = True

        # 오늘자 일봉이 이미 생성되어 있다면 확실히 개장일
        if has_today_candle:
            return True

        market_open_time = now.replace(hour=9, minute=0, second=0, microsecond=0)
        market_close_time = now.replace(hour=15, minute=30, second=0, microsecond=0)

        # 장마감 모드 ('closing'): 15:30 이후 실행되었으나 오늘 날짜의 봉 데이터가 없으면 휴장일
        if mode == "closing":
            if now >= market_close_time and not has_today_candle:
                return False

        # 장 시작 전 / 장중 모드 ('daemon', 'intraday') 또는 15:30 이전 closing
        # 네이버 실시간 국내 지수 API의 marketStatus 및 최신 거래일자 교차 점검
        kospi_url = "https://polling.finance.naver.com/api/realtime/domestic/index/KOSPI"
        res_kospi = requests.get(kospi_url, headers=HEADERS, timeout=5)
        if res_kospi.status_code == 200:
            datas = res_kospi.json().get("datas", [])
            if datas:
                d0 = datas[0]
                market_status = str(d0.get("marketStatus", "")).upper()
                local_traded_at = str(d0.get("localTradedAt", ""))
                trade_date = local_traded_at.split("T")[0] if "T" in local_traded_at else ""

                if market_open_time <= now <= market_close_time:
                    # 장중인데 marketStatus가 OPEN이 아니고 최신 거래일자도 오늘이 아니면 휴장일
                    if market_status != "OPEN" and trade_date != today_dash:
                        return False
                    return True
                elif now > market_close_time:
                    # 장마감 이후인데 봉 데이터도 없고 거래일자도 오늘이 아니면 휴장일
                    if trade_date != today_dash and not has_today_candle:
                        return False
                    return True
                else:
                    # 장 시작 전 (09:00 이전):
                    # 장전 상태에서는 기본 True 반환 (정규장 진입 후 스케줄러/장중 모드에서 2차 검증)
                    return True

        return True

    except Exception as e:
        print(f"⚠️ [증시 개장일 판별 경고] 네이버 시세 API 확인 중 오류 발생 (기본값 True 유지): {e}")
        return True


def get_intraday_volume_multiplier(now_dt: Optional[datetime.datetime] = None) -> float:
    """
    정규장(09:00~15:30, 총 390분) 경과 시간에 따른 거래량 보정 배수 계산
    - 평일 장중 경과 시간에 비례하여 연산 (초기 왜곡 방지를 위해 최대 5.0배 상한)
    - 장마감 또는 주말에는 1.0 반환
    """
    if now_dt is None:
        now_dt = datetime.datetime.now()

    if now_dt.weekday() >= 5:
        return 1.0

    market_open = now_dt.replace(hour=9, minute=0, second=0, microsecond=0)
    market_close = now_dt.replace(hour=15, minute=30, second=0, microsecond=0)

    if now_dt <= market_open or now_dt >= market_close:
        return 1.0

    elapsed_minutes = max(15.0, (now_dt - market_open).total_seconds() / 60.0)
    multiplier = 390.0 / elapsed_minutes
    return min(multiplier, 5.0)


def send_telegram_message(text: str, chat_id: str = MY_CHAT_ID) -> bool:
    """텔레그램 메시지 발송 유틸리티"""
    if not MY_BOT_TOKEN or not chat_id:
        print("ℹ️ 텔레그램 토큰 또는 Chat ID가 설정되지 않아 콘솔 출력만 수행합니다.")
        return False
    url = f"https://api.telegram.org/bot{MY_BOT_TOKEN}/sendMessage"
    payload = {
        "chat_id": chat_id,
        "text": text[:4000],
        "parse_mode": "HTML",
        "disable_web_page_preview": True,
    }
    try:
        res = requests.post(url, data=payload, timeout=10)
        if res.status_code == 200:
            return True
        print(f"⚠️ 텔레그램 발송 실패 ({res.status_code}): {res.text}")
        return False
    except Exception as e:
        print(f"⚠️ 텔레그램 통신 오류: {e}")
        return False


def upload_excel_to_gdrive(file_path: str, custom_filename: Optional[str] = None) -> Optional[str]:
    """
    Google Cloud Service Account를 이용하여 분석 엑셀 보고서를 구글 드라이브에 업로드하고 공유 링크를 반환
    1. GDRIVE_SERVICE_ACCOUNT_JSON (환경변수 문자열) 또는 GDRIVE_SERVICE_ACCOUNT_FILE (JSON 파일 경로) 확인
    2. GDRIVE_FOLDER_ID 폴더가 설정되어 있으면 해당 폴더 내로 업로드 (지정 안될 시 루트)
    3. 링크 열람 권한(anyone reader) 부여
    4. webViewLink 반환 (실패 시 None 반환 및 안전 예외 처리)
    """
    if not GDRIVE_AVAILABLE:
        print("ℹ️ [구글 드라이브] google-api-python-client 라이브러리 미설치로 드라이브 업로드를 건너뜁니다.")
        return None

    if not os.path.exists(file_path):
        print(f"⚠️ [구글 드라이브] 업로드 대상 파일이 존재하지 않습니다: {file_path}")
        return None

    creds = None
    if GDRIVE_SERVICE_ACCOUNT_JSON:
        try:
            info = json.loads(GDRIVE_SERVICE_ACCOUNT_JSON)
            creds = service_account.Credentials.from_service_account_info(
                info, scopes=["https://www.googleapis.com/auth/drive"]
            )
        except Exception as e:
            print(f"⚠️ [구글 드라이브] GDRIVE_SERVICE_ACCOUNT_JSON 파싱 실패: {e}")

    if not creds:
        sa_file = GDRIVE_SERVICE_ACCOUNT_FILE
        if not os.path.isabs(sa_file):
            sa_file = os.path.join(BASE_DIR, sa_file)
        if os.path.exists(sa_file):
            try:
                creds = service_account.Credentials.from_service_account_file(
                    sa_file, scopes=["https://www.googleapis.com/auth/drive"]
                )
            except Exception as e:
                print(f"⚠️ [구글 드라이브] 서비스 계정 파일 로드 실패 ({sa_file}): {e}")

    if not creds:
        print("ℹ️ [구글 드라이브] 서비스 계정 인증 정보가 설정되지 않아 드라이브 업로드를 건너뜁니다.")
        return None

    try:
        service = gdrive_build("drive", "v3", credentials=creds, cache_discovery=False)
        filename = custom_filename or os.path.basename(file_path)

        file_metadata: Dict[str, Any] = {"name": filename}
        if GDRIVE_FOLDER_ID:
            file_metadata["parents"] = [GDRIVE_FOLDER_ID]

        media = MediaFileUpload(
            file_path,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            resumable=True
        )

        file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields="id, webViewLink, webContentLink"
        ).execute()

        file_id = file.get("id")
        web_link = file.get("webViewLink")

        # 링크가 있는 사용자 누구나 열람 가능하도록 권한 추가
        try:
            service.permissions().create(
                fileId=file_id,
                body={"type": "anyone", "role": "reader"}
            ).execute()
        except Exception:
            pass

        print(f"✅ [구글 드라이브] 엑셀 보고서 업로드 완료: {web_link}")
        return web_link
    except Exception as e:
        print(f"⚠️ [구글 드라이브] 업로드 중 예외 발생: {e}")
        return None


def auto_git_backup(excel_path: str, latest_date_str: str) -> bool:
    """분석 완료 후 생성된 엑셀 보고서 및 변경 코드를 GitHub 저장소로 자동 커밋·푸시"""
    if os.environ.get("GITHUB_ACTIONS") == "true":
        return True
    try:
        repo_dir = BASE_DIR
        excel_name = os.path.basename(excel_path)
        add_candidates = [excel_name, ".gitignore", "ALGORITHM_SPEC.md", "README.md"]
        existing_targets = [f for f in add_candidates if os.path.exists(os.path.join(repo_dir, f))]
        
        subprocess.run(["git", "add"] + existing_targets + ["*.py"], cwd=repo_dir, check=True)
        res = subprocess.run(["git", "status", "--porcelain"], cwd=repo_dir, capture_output=True, text=True, encoding="utf-8")
        if not res.stdout.strip():
            print("  └ ℹ️ Git 변경 사항이 없어 커밋을 건너뜁니다.")
            return True

        branch_res = subprocess.run(["git", "branch", "--show-current"], cwd=repo_dir, capture_output=True, text=True, encoding="utf-8")
        curr_branch = branch_res.stdout.strip() or "main"

        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        commit_msg = f"data: {latest_date_str} 마감 분석 엑셀 백업 ({now_str})"
        subprocess.run(["git", "commit", "-m", commit_msg], cwd=repo_dir, check=True)
        subprocess.run(["git", "push", "origin", curr_branch], cwd=repo_dir, check=True)
        print(f"  └ ✅ GitHub 자동 푸시 백업 완료 ({curr_branch}): {commit_msg}")
        return True
    except Exception as e:
        print(f"  └ ⚠️ Git 자동 백업 중 예외 발생: {e}")
        return False


def split_lines_to_notion_paragraphs(header_line: str, lines: List[str], max_chars: int = 1800) -> List[Dict[str, Any]]:
    """노션 2,000자 제한 방어를 위한 문단 분할 유틸리티"""
    blocks = []
    current_text = header_line + "\n"
    for line in lines:
        if len(current_text) + len(line) + 1 > max_chars:
            blocks.append({
                "object": "block",
                "type": "paragraph",
                "paragraph": {"rich_text": [{"type": "text", "text": {"content": current_text.rstrip()}}]}
            })
            current_text = f"{header_line} (계속)\n{line}\n"
        else:
            current_text += line + "\n"

    if current_text.strip():
        blocks.append({
            "object": "block",
            "type": "paragraph",
            "paragraph": {"rich_text": [{"type": "text", "text": {"content": current_text.rstrip()}}]}
        })
    return blocks


def ensure_notion_database_schema() -> Tuple[str, Dict[str, Any]]:
    """노션 데이터베이스 스키마 검사 및 속성 보정"""
    if not NOTION_API_KEY or not NOTION_DATABASE_ID:
        return "제목", {}

    db_url = f"https://api.notion.com/v1/databases/{NOTION_DATABASE_ID.strip()}"
    try:
        res = requests.get(db_url, headers=NOTION_HEADERS, timeout=10)
        title_col_name = "제목"
        existing_props = {}
        if res.status_code == 200:
            props = res.json().get("properties", {})
            existing_props = props
            for prop_name, prop_data in props.items():
                if prop_data.get("type") == "title":
                    title_col_name = prop_name
                    break

            target_props = {
                "날짜": {"date": {}},
                "시장구분": {"select": {}},
                "포착종목수": {"number": {"format": "number"}},
                "KOSPI장세": {"rich_text": {}},
                "KOSDAQ장세": {"rich_text": {}},
                "조기경고수": {"number": {"format": "number"}},
                "등록시간": {"rich_text": {}}
            }
            missing_props = {p: schema for p, schema in target_props.items() if p not in props}
            if missing_props:
                patch_res = requests.patch(db_url, json={"properties": missing_props}, headers=NOTION_HEADERS, timeout=10)
                if patch_res.status_code == 200:
                    existing_props = patch_res.json().get("properties", existing_props)
        return title_col_name, existing_props
    except Exception as e:
        print(f"노션 스키마 검사 오류: {e}")
        return "제목", {}


def append_notion_blocks(page_id: str, blocks: List[Dict[str, Any]]) -> None:
    """노션 페이지에 하위 블록 청크 추가"""
    if not blocks:
        return
    url = f"https://api.notion.com/v1/blocks/{page_id}/children"
    batch_size = 80
    for i in range(0, len(blocks), batch_size):
        chunk = blocks[i:i + batch_size]
        payload = {"children": chunk}
        try:
            res = requests.patch(url, json=payload, headers=NOTION_HEADERS, timeout=15)
            if res.status_code not in [200, 201]:
                print(f"  └ ⚠️ 노션 블록 추가 실패 (Chunk {i // batch_size + 1}): {res.status_code} - {res.text[:200]}")
            time.sleep(0.35)
        except Exception as e:
            print(f"  └ ⚠️ 노션 블록 추가 중 예외: {e}")


# ==============================================================================
# 3. 데이터 파이프라인 (시세 수집 & 40EMA 지표/시그널 연산)
# ==============================================================================
def fetch_etf_ohlcv(code: str, count: int = 300, vol_mult: float = 1.0) -> Optional[pd.DataFrame]:
    """
    네이버 차트 XML API를 통해 ETF 일봉(당일 실시간 포함)을 단일 수집.
    - count: 장중 모드 시 300, 장마감 10년 모드 시 2600
    """
    url = f"https://fchart.stock.naver.com/sise.nhn?symbol={code}&timeframe=day&count={count}&requestType=0"
    try:
        res = requests.get(url, headers=HEADERS, timeout=12)
        if res.status_code != 200:
            return None
        root = ET.fromstring(res.text)
        rows = []
        for item in root.findall(".//item"):
            p = item.attrib["data"].split("|")
            if len(p) >= 6:
                d_str = f"{p[0][:4]}-{p[0][4:6]}-{p[0][6:8]}"
                rows.append({
                    "날짜": d_str,
                    "시가": float(p[1]),
                    "고가": float(p[2]),
                    "저가": float(p[3]),
                    "종가": float(p[4]),
                    "거래량": float(p[5]),
                })
        if not rows:
            return None

        df = pd.DataFrame(rows)
        df["dt"] = pd.to_datetime(df["날짜"])
        df = df.sort_values(by="dt", ascending=True).reset_index(drop=True)

        # 40EMA 및 기술적 지표 기본 산출
        df["MA40"] = df["종가"].ewm(span=40, min_periods=20, adjust=False).mean()
        df["Vol_MA40"] = df["거래량"].ewm(span=40, min_periods=20, adjust=False).mean()
        df["MA40_Slope5"] = ((df["MA40"] - df["MA40"].shift(5)) / df["MA40"].shift(5)) * 100.0
        df["MA40_Disparity"] = (df["종가"] / df["MA40"]) * 100.0
        df["EMA40_2Day_Below"] = (df["종가"].shift(1) < df["MA40"].shift(1)) & (df["종가"] < df["MA40"])

        return calculate_indicators_and_signals(df, vol_mult=vol_mult, code=code)
    except Exception as e:
        print(f"⚠️ [{code}] 시세 수집 오류: {e}")
        return None


def calculate_indicators_and_signals(df: pd.DataFrame, vol_mult: float = 1.0, code: str = "") -> pd.DataFrame:
    """
    6개 보조 시그널, 40EMA 조기경고, 트렌드(+) 시그널 및 EMA40_Filter(10년 통계 기반 등급) 판정 단일화
    - 2-1 지표 연산부: MA40_Slope5(5일 기울기 %), MA40_Disparity(이격도 %), EMA40_2Day_Below(2일 연속 이탈)
    - 2-2 매매신호 필터링: KOSPI 정방향, KOSDAQ 정방향, 인버스 헤지 종목별 10년 검증 등급 부여
    - vol_mult > 1.0일 경우 당일(마지막 행) 장중 예상 거래량 판정 병행
    """
    n = len(df)

    if "MA40" not in df.columns:
        df["MA40"] = df["종가"].ewm(span=40, min_periods=20, adjust=False).mean()
    if "Vol_MA40" not in df.columns:
        df["Vol_MA40"] = df["거래량"].ewm(span=40, min_periods=20, adjust=False).mean()
    df["MA40_Slope5"] = ((df["MA40"] - df["MA40"].shift(5)) / df["MA40"].shift(5)) * 100.0
    df["MA40_Disparity"] = (df["종가"] / df["MA40"]) * 100.0
    df["EMA40_2Day_Below"] = (df["종가"].shift(1) < df["MA40"].shift(1)) & (df["종가"] < df["MA40"])

    etf_meta = ETF_MAP.get(code, {})
    market = etf_meta.get("market", "")
    etf_type = etf_meta.get("type", "")

    sig1_40ma, sig2_gap, sig3_gap_close, sig4_high, sig5_low = [], [], [], [], []
    ema40_early_alerts = []

    for idx in range(n):
        if pd.isna(df.loc[idx, "MA40"]):
            sig1_40ma.append("데이터 부족")
        else:
            cp = df.loc[idx, "종가"]
            m40 = df.loc[idx, "MA40"]
            if cp > m40:
                sig1_40ma.append("상승")
            elif cp < m40:
                sig1_40ma.append("하락")
            else:
                sig1_40ma.append("보합")

        # 40EMA 2거래일 연속 이탈/돌파 조기경고
        if idx < 3 or pd.isna(df.loc[idx, "MA40"]) or pd.isna(df.loc[idx - 2, "MA40"]):
            ema40_early_alerts.append("정상")
        else:
            c0, m0 = df.loc[idx, "종가"], df.loc[idx, "MA40"]
            c1, m1 = df.loc[idx - 1, "종가"], df.loc[idx - 1, "MA40"]
            c2, m2 = df.loc[idx - 2, "종가"], df.loc[idx - 2, "MA40"]
            if (c2 >= m2) and (c1 < m1) and (c0 < m0):
                ema40_early_alerts.append("⚠️ 40EMA 2일연속 하향이탈")
            elif (c2 <= m2) and (c1 > m1) and (c0 > m0):
                ema40_early_alerts.append("🟢 40EMA 2일연속 상향돌파")
            else:
                ema40_early_alerts.append("정상")

        if idx < 1:
            sig2_gap.append("데이터 부족")
            sig3_gap_close.append("데이터 부족")
            sig4_high.append("데이터 부족")
            sig5_low.append("데이터 부족")
            continue

        d0 = df.loc[idx]
        d1 = df.loc[idx - 1]

        # 갭 비교
        if d1["종가"] < d0["시가"]:
            sig2_gap.append("갭상승")
        elif d1["종가"] > d0["시가"]:
            sig2_gap.append("갭하락")
        else:
            sig2_gap.append("보합")

        # 갭과 종가 방향
        if (d1["종가"] < d0["시가"]) and (d0["종가"] < d0["시가"]):
            sig3_gap_close.append("하락마감")
        elif (d1["종가"] > d0["시가"]) and (d0["종가"] > d0["시가"]):
            sig3_gap_close.append("상승마감")
        else:
            sig3_gap_close.append("기타")

        # 고가/저가 비교
        sig4_high.append("고가상승" if d0["고가"] > d1["고가"] else ("고가하락" if d0["고가"] < d1["고가"] else "보합"))
        sig5_low.append("저가상승" if d0["저가"] > d1["저가"] else ("저가하락" if d0["저가"] < d1["저가"] else "보합"))

    sig6_vol = []
    for idx in range(n):
        v = df.loc[idx, "거래량"] if pd.notna(df.loc[idx, "거래량"]) else 0
        v_ma = df.loc[idx, "Vol_MA40"] if pd.notna(df.loc[idx, "Vol_MA40"]) else 0
        if pd.isna(v_ma) or v_ma == 0:
            sig6_vol.append("데이터 부족")
        elif v > v_ma:
            sig6_vol.append("거래량 상승")
        elif v < v_ma:
            sig6_vol.append("거래량 하락")
        else:
            sig6_vol.append("보합")

    df["보조_40MA"] = sig1_40ma
    df["보조_Gap"] = sig2_gap
    df["보조_갭과시가종가"] = sig3_gap_close
    df["보조_고가비교"] = sig4_high
    df["보조_저가비교"] = sig5_low
    df["보조_거래량비교"] = sig6_vol
    df["EMA40_조기경고"] = ema40_early_alerts

    market_trends = []
    ema40_filter_tags = []

    for idx in range(n):
        if idx < 2 or df.loc[idx, "보조_40MA"] == "데이터 부족":
            market_trends.append("데이터 부족")
            ema40_filter_tags.append("-")
            continue

        d0 = df.loc[idx]
        d1 = df.loc[idx - 1]
        d2 = df.loc[idx - 2]

        c_down_d2 = (d2["보조_Gap"] == "갭상승") and (d2["보조_갭과시가종가"] == "하락마감")
        c_down_d1 = (d1["보조_고가비교"] == "고가하락") and (d1["보조_저가비교"] == "저가하락")
        c_down_d0 = (d0["보조_고가비교"] == "고가하락") and (d0["보조_저가비교"] == "저가하락") and (d0["보조_40MA"] == "하락")
        is_downtrend = c_down_d2 and c_down_d1 and c_down_d0

        c_up_d2 = (d2["보조_Gap"] == "갭하락") and (d2["보조_갭과시가종가"] == "상승마감")
        c_up_d1 = (d1["보조_고가비교"] == "고가상승") and (d1["보조_저가비교"] == "저가상승")
        c_up_d0 = (d0["보조_고가비교"] == "고가상승") and (d0["보조_저가비교"] == "저가상승") and (d0["보조_40MA"] == "상승")
        is_uptrend = c_up_d2 and c_up_d1 and c_up_d0

        vol_0 = d0["거래량"] if pd.notna(d0["거래량"]) else 0
        vol_ma40_0 = d0["Vol_MA40"] if pd.notna(d0["Vol_MA40"]) else 0

        # 당일(마지막 행) 장중 거래량 보정 반영
        is_today = (idx == n - 1)
        raw_vol_ok = (vol_0 > vol_ma40_0) and (vol_ma40_0 > 0)
        est_vol_ok = ((vol_0 * vol_mult) > vol_ma40_0) and (vol_ma40_0 > 0)

        slope = d0["MA40_Slope5"] if pd.notna(d0["MA40_Slope5"]) else 0.0
        disp = d0["MA40_Disparity"] if pd.notna(d0["MA40_Disparity"]) else 100.0

        if is_downtrend:
            if is_today and vol_mult > 1.0:
                sig_label = "하락트렌드 +(확정)" if raw_vol_ok else ("하락트렌드 +(장중예상)" if est_vol_ok else "캔들하락(거래량미달)")
            else:
                sig_label = "하락트렌드 +" if raw_vol_ok else "해당없음"

            market_trends.append(sig_label)
            if "하락트렌드 +" in sig_label:
                if etf_type == "정방향":
                    if slope <= -EMA40_SLOPE_THRESHOLD:
                        ema40_filter_tags.append("✅40EMA하향확정")
                    else:
                        ema40_filter_tags.append("⚠️40EMA평탄(휩소주의)")
                elif etf_type == "인버스":
                    ema40_filter_tags.append("✅인버스하향(정방향우호)")
                else:
                    ema40_filter_tags.append("✅하향트렌드확정")
            else:
                ema40_filter_tags.append("-")

        elif is_uptrend:
            if is_today and vol_mult > 1.0:
                sig_label = "상승트렌드 +(확정)" if raw_vol_ok else ("상승트렌드 +(장중예상)" if est_vol_ok else "캔들상승(거래량미달)")
            else:
                sig_label = "상승트렌드 +" if raw_vol_ok else "해당없음"

            market_trends.append(sig_label)
            if "상승트렌드 +" in sig_label:
                # 10년 백테스트 기반 필터 등급 부여
                if etf_type == "정방향":
                    if market == "KOSPI":
                        # KOSPI 정방향: Slope >= +0.15% & Disparity <= 115.0%
                        if slope >= EMA40_SLOPE_THRESHOLD and disp <= 115.0:
                            ema40_filter_tags.append("✅40EMA상향확정 (정방향 매수 승인)")
                        elif disp > 115.0:
                            ema40_filter_tags.append("⚠️단기과열 (추격 매수 금지)")
                        elif abs(slope) < EMA40_SLOPE_THRESHOLD:
                            ema40_filter_tags.append("⚠️40EMA평탄(휩소주의)")
                        else:
                            ema40_filter_tags.append("⚠️40EMA역배열반등")
                    elif market == "KOSDAQ":
                        # KOSDAQ 정방향: Slope >= +0.15% & Disparity <= 104.0% 초입형
                        if slope >= EMA40_SLOPE_THRESHOLD and disp <= 104.0:
                            ema40_filter_tags.append("✅40EMA초입확정 (단기 매수 승인)")
                        elif disp > 108.0:
                            ema40_filter_tags.append("⚠️고위험과열 (추격 승률 12.5% 이하, 매수 금지)")
                        elif disp > 104.0:
                            ema40_filter_tags.append("⚠️초입초과(단기추격주의)")
                        elif abs(slope) < EMA40_SLOPE_THRESHOLD:
                            ema40_filter_tags.append("⚠️40EMA평탄(휩소주의)")
                        else:
                            ema40_filter_tags.append("⚠️40EMA역배열반등")
                    else:
                        if slope >= EMA40_SLOPE_THRESHOLD:
                            ema40_filter_tags.append("✅40EMA상향확정")
                        else:
                            ema40_filter_tags.append("⚠️40EMA평탄(휩소주의)")
                elif etf_type == "인버스":
                    # 인버스 헤지 종목 단독 신호 판정 (4종 동조 검증 전 기본 단독 휩소 또는 패닉 바닥)
                    if disp >= 111.0:
                        ema40_filter_tags.append("🛑패닉바닥도달 (인버스 전량 익절 및 추격 금지)")
                    else:
                        ema40_filter_tags.append("⚠️단독휩소주의 (진입 금지)")
                else:
                    ema40_filter_tags.append("✅상승신호확인")
            else:
                ema40_filter_tags.append("-")
        else:
            market_trends.append("해당없음")
            ema40_filter_tags.append("-")

    df["시장트렌드_시그널"] = market_trends
    df["EMA40_필터판정"] = ema40_filter_tags
    df["EMA40_Filter"] = ema40_filter_tags
    return df





def generate_gemini_advice(report_text: str) -> str:
    """
    구글 Gemini API(gemini-2.5-flash)를 연동하여 10년 통계 기반 실전 장세 대응 조언을 생성합니다.
    - 3-트랙 핵심 규칙(KOSPI 정방향, KOSDAQ 정방향, KOSPI 인버스, 100% 현금 관망) 프롬프트 탑재
    - 응답 텍스트에서 마크다운 볼드(**) 기호 제거 및 html.escape 적용
    """
    if not GEMINI_API_KEY:
        return "⚠️ <b>[Gemini AI 조언 알림]</b> GEMINI_API_KEY가 설정되지 않아 AI 실전 조언을 생성할 수 없습니다."

    try:
        from google import genai
        client = genai.Client(api_key=GEMINI_API_KEY)

        prompt = (
            "당신은 대한민국 ETF 시장의 10년 통계 데이터와 40EMA(지수이동평균) 트렌드 필터링 시스템을 완벽히 마스터한 퀀트 수석 트레이더입니다.\n"
            "아래 제공된 '8대 ETF 실시간 주가 분석 리포트'를 바탕으로, 당사의 [10년 검증 3-트랙 핵심 규칙]에 엄격히 입각하여 "
            "개인 투자자가 지금 당장 취해야 할 가장 냉철하고 명확한 실전 포지션 대응 조언을 작성하세요.\n\n"
            "[10년 검증 3-트랙 핵심 규칙]\n"
            "① KOSPI 정방향 (069500 KODEX 200, 102110 TIGER 200):\n"
            "   - 진입: 40EMA 5일 기울기 ≥+0.15% 및 이격도 ≤115% 진입\n"
            "   - 손절: 40EMA 2일 연속 이탈 시 전량 칼손절 및 현금화\n"
            "   - 과열: 이격도 ≥120% 시 50% 분할 익절\n"
            "② KOSDAQ 정방향 (229200 KODEX 코스닥150, 232080 TIGER 코스닥150):\n"
            "   - 진입: 40EMA 5일 기울기 ≥+0.15% 및 이격도 ≤104% 초입만 20일 단기 스윙 진입\n"
            "   - 추격 금지: 이격도 108% 초과 시 신규 추격 매수 절대 금지\n"
            "   - 익절/손절: 이격도 ≥109% 상투권 전량 익절, 40EMA 2일 연속 이탈 시 전량 손절\n"
            "③ KOSPI 인버스 (114800 KODEX 인버스, 123310 TIGER 인버스):\n"
            "   - 진입: 4대교차 하락 & 이격도 ≤105% 초입만 20일 진입\n"
            "   - 패닉 익절: 지수(KOSPI 200) 이격도 ≤88.5% 또는 인버스 이격도 ≥111.0% 패닉 도달 시 즉시 전량 익절 및 추격 매수 절대 금지\n"
            "④ 정방향 하락세 + 인버스 과열 시:\n"
            "   - 무리한 포지션 진입을 금지하고 100% 현금 관망 유지\n\n"
            "[실시간 리포트 데이터]\n"
            f"{report_text}\n\n"
            "[작성 지침]\n"
            "- 서론/인사말 없이 직관적인 불릿포인트 요약 형태로 작성할 것\n"
            "- 1. 시장 국면 총평 (현재 장세 판정 및 40EMA 방향성 요약)\n"
            "- 2. 트랙별 실전 포지션 대응 지침 (KOSPI 정방향, KOSDAQ 정방향, KOSPI 인버스별 매수/보유/손절/익절/관망 명시)\n"
            "- 3. 오늘의 핵심 주의사항 (추격 매수 금지 구간, 손절 기준가 등)\n"
            "- 분량은 전체 텔레그램 메시지 길이 제한을 고려하여 핵심 위주로 공백 포함 500~800자 내외로 간결하게 작성할 것\n"
            "- 마크다운 굵은 글씨(**)는 일절 사용하지 말 것\n"
        )

        model_name = os.environ.get("GEMINI_MODEL", "gemini-2.5-flash")
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=prompt,
            )
        except Exception as model_err:
            err_msg = str(model_err)
            if "gemini-2.5-flash" in model_name and ("404" in err_msg or "not available" in err_msg.lower()):
                response = client.models.generate_content(
                    model="gemini-3.8-flash",
                    contents=prompt,
                )
            else:
                raise model_err

        raw_advice = response.text or ""
        clean_advice = raw_advice.replace("**", "").strip()
        escaped_advice = html.escape(clean_advice)

        header = "🤖 <b>[Gemini AI 10년 통계 기반 실전 대응 조언]</b>"
        return f"{header}\n\n{escaped_advice}"

    except Exception as e:
        return f"⚠️ <b>[Gemini AI 실전 조언 생성 실패]</b> {html.escape(str(e))}"


# ==============================================================================
# 4. 장중 실시간 브리핑 모드 (run_intraday_mode)
# ==============================================================================
def build_intraday_report(trigger_label: str = "수동 조회", include_ai: bool = False) -> str:
    """현재 주가 분석 및 자동매매 3-트랙 대응 지침 메시지 작성 (include_ai=True 시 Gemini AI 조언 결합)"""
    now_dt = datetime.datetime.now()
    now_str = now_dt.strftime("%Y-%m-%d %H:%M:%S")
    vol_mult = get_intraday_volume_multiplier(now_dt)
    is_market_open = vol_mult > 1.0

    dfs = {}
    for item in TARGET_ETFS:
        df = fetch_etf_ohlcv(item["code"], count=FETCH_COUNT_INTRADAY, vol_mult=vol_mult)
        if df is not None:
            dfs[item["code"]] = df

    if len(dfs) < len(TARGET_ETFS):
        return f"⚠️ ETF 시세 수집 지연으로 리포트를 완성하지 못했습니다. (수집 성공: {len(dfs)}/{len(TARGET_ETFS)})"

    r_k200 = dfs["069500"].iloc[-1]
    r_kinv = dfs["114800"].iloc[-1]
    r_kq150 = dfs["229200"].iloc[-1]
    r_kqinv = dfs["251340"].iloc[-1]

    def eval_mkt(long_row, short_row):
        if long_row["종가"] > long_row["MA40"] and short_row["종가"] < short_row["MA40"]:
            return "🔴 상승 우세 (40EMA 상회)" if long_row["MA40_Slope5"] >= EMA40_SLOPE_THRESHOLD else "🟠 상승 둔화 (기울기 평탄)"
        elif long_row["종가"] < long_row["MA40"] and short_row["종가"] > short_row["MA40"]:
            return "🔵 하락 우세 (40EMA 하회)" if long_row["MA40_Slope5"] <= -EMA40_SLOPE_THRESHOLD else "🟡 하락 둔화 (바닥 탐색)"
        return "⚪ 혼조 / 변곡점 공방"

    kospi_st = eval_mkt(r_k200, r_kinv)
    kosdaq_st = eval_mkt(r_kq150, r_kqinv)

    lines = [
        f"📊 <b>[8대 ETF 실시간 주가 분석 리포트]</b> ({html.escape(trigger_label)})",
        f"⏱️ <b>조회시각:</b> {now_str} " + (f"(장중 거래량보정 ×{vol_mult:.2f})" if is_market_open else "(장마감 확정종가)"),
        "━━━━━━━━━━━━━━━━━━━━━",
        "📌 <b>[1. 양 시장 실시간 장세 진단]</b>",
        f"• <b>KOSPI:</b> {kospi_st}",
        f"• <b>KOSDAQ:</b> {kosdaq_st}",
        "",
        "📈 <b>[2. 주요 ETF 현재가 & 40EMA 지표]</b>",
    ]

    core_codes = ["069500", "114800", "229200", "251340"]
    for item in TARGET_ETFS:
        code = item["code"]
        df = dfs[code]
        r0 = df.iloc[-1]
        r1 = df.iloc[-2]
        r5 = df.iloc[-6] if len(df) >= 6 else df.iloc[0]
        r20 = df.iloc[-21] if len(df) >= 21 else df.iloc[0]

        chg_1d = ((r0["종가"] - r1["종가"]) / r1["종가"]) * 100.0
        chg_5d = ((r0["종가"] - r5["종가"]) / r5["종가"]) * 100.0
        chg_20d = ((r0["종가"] - r20["종가"]) / r20["종가"]) * 100.0
        vol_pct = (r0["거래량"] / r0["Vol_MA40"]) * 100.0 if r0["Vol_MA40"] > 0 else 0.0

        sig = r0["시장트렌드_시그널"]
        f_tag = r0.get("EMA40_Filter") or r0.get("EMA40_필터판정", "-")
        if code in core_codes or sig not in ("해당없음", "데이터 부족"):
            sig_badge = ""
            if sig not in ("해당없음", "데이터 부족"):
                sig_badge = f"\n  └ 🚨 <b>시그널: {html.escape(sig)}</b> | <b>[{html.escape(f_tag)}]</b>"
            elif f_tag and f_tag != "-":
                sig_badge = f"\n  └ 📌 <b>필터: [{html.escape(f_tag)}]</b>"

            lines.append(
                f"• <b>{item['name']}</b>: <b>{int(r0['종가']):,}원</b> ({chg_1d:+.2f}%)\n"
                f"  └ 5일 {chg_5d:+.1f}% | 20일 {chg_20d:+.1f}% | 40MA {int(r0['MA40']):,}원\n"
                f"  └ 이격 <b>{r0['MA40_Disparity']:.1f}%</b> | 기울기 {r0['MA40_Slope5']:+.2f}% | 거래량 {vol_pct:.0f}%{sig_badge}"
            )

    lines.append("\n🧭 <b>[3. 자동매매 3-트랙 실시간 대응 지침]</b>")

    # 4종 동조 검증 (KOSPI 정방향 2종 & 인버스 2종 40EMA 동시 교차 상태)
    r_k200_tiger = dfs["102110"].iloc[-1]
    r_kinv_tiger = dfs["123310"].iloc[-1]
    kospi_long_down = (r_k200["종가"] < r_k200["MA40"]) and (r_k200_tiger["종가"] < r_k200_tiger["MA40"])
    kospi_inv_up = (r_kinv["종가"] > r_kinv["MA40"]) and (r_kinv_tiger["종가"] > r_kinv_tiger["MA40"])
    is_kospi_cross_down = kospi_long_down and kospi_inv_up

    # Track 1: KOSPI 정방향 (069500 / 102110)
    k_two_day_below = (dfs["069500"].iloc[-2]["종가"] < dfs["069500"].iloc[-2]["MA40"]) and (r_k200["종가"] < r_k200["MA40"])
    if k_two_day_below or r_k200["종가"] < r_k200["MA40"]:
        lines.append(f"• <b>KOSPI 정방향:</b> ⚪ <b>100% 현금 관망</b> (40EMA {int(r_k200['MA40']):,}원 하회, 매수 금지)")
    elif r_k200["MA40_Disparity"] > 115.0:
        lines.append(f"• <b>KOSPI 정방향:</b> ⚠️ <b>홀딩 및 분할익절</b> (이격 <b>{r_k200['MA40_Disparity']:.1f}%</b> 과열, 추격 매수 금지)")
    elif r_k200["MA40_Slope5"] >= EMA40_SLOPE_THRESHOLD and r_k200["MA40_Disparity"] <= 115.0:
        lines.append(f"• <b>KOSPI 정방향:</b> 🟢 <b>적극 매수 승인</b> (기울기 {r_k200['MA40_Slope5']:+.2f}% / 이격 {r_k200['MA40_Disparity']:.1f}%)")
    else:
        lines.append(f"• <b>KOSPI 정방향:</b> 🟡 <b>추세 유지 / 홀딩</b> (40EMA {int(r_k200['MA40']):,}원 지지 중)")

    # Track 2: KOSDAQ 정방향 (229200 / 232080)
    q_two_day_below = (dfs["229200"].iloc[-2]["종가"] < dfs["229200"].iloc[-2]["MA40"]) and (r_kq150["종가"] < r_kq150["MA40"])
    if q_two_day_below or r_kq150["종가"] < r_kq150["MA40"]:
        lines.append(f"• <b>KOSDAQ 정방향:</b> ⚪ <b>100% 현금 관망 / 손절</b> (40EMA {int(r_kq150['MA40']):,}원 하회, 진입 절대 금지)")
    elif r_kq150["MA40_Disparity"] > 108.0:
        lines.append(f"• <b>KOSDAQ 정방향:</b> ⚠️ <b>전량 익절 / 신규 금지</b> (고위험 과열 이격 <b>{r_kq150['MA40_Disparity']:.1f}%</b>, 추격 승률 12.5% 이하)")
    elif r_kq150["MA40_Slope5"] >= EMA40_SLOPE_THRESHOLD and r_kq150["MA40_Disparity"] <= 104.0:
        lines.append(f"• <b>KOSDAQ 정방향:</b> 🟢 <b>단기 매수 승인</b> (초입형 이격 {r_kq150['MA40_Disparity']:.1f}% / 20일 스윙 목표)")
    else:
        lines.append(f"• <b>KOSDAQ 정방향:</b> 🟡 <b>신규 보류 / 관망</b> (이격 {r_kq150['MA40_Disparity']:.1f}% 초입 기준 초과)")

    # Track 3: KOSPI 인버스 (114800 / 123310)
    if r_k200["MA40_Disparity"] <= 88.5 or r_kinv["MA40_Disparity"] >= 111.0:
        lines.append(
            f"• <b>KOSPI 인버스:</b> 🛑 <b>패닉 바닥 도달 (인버스 전량 익절 및 추격 금지)</b> "
            f"(지수 이격 <b>{r_k200['MA40_Disparity']:.1f}%</b> / 인버스 이격 <b>{r_kinv['MA40_Disparity']:.1f}%</b>)"
        )
    elif is_kospi_cross_down and r_kinv["MA40_Disparity"] <= 105.0:
        lines.append(
            f"• <b>KOSPI 인버스:</b> 🎯 <b>헤지 매수 승인 (인버스 폭락 초입)</b> "
            f"(KOSPI 4대 교차 하락 동조 & 인버스 이격 <b>{r_kinv['MA40_Disparity']:.1f}%</b>)"
        )
    else:
        concur_msg = "4대교차 미동조" if not is_kospi_cross_down else f"이격 {r_kinv['MA40_Disparity']:.1f}% 과열"
        lines.append(f"• <b>KOSPI 인버스:</b> ⚪ <b>진입 금지 / 단독 휩소 주의</b> ({concur_msg})")

    # 종합 포지션 결론
    is_all_long_down = (r_k200["종가"] < r_k200["MA40"]) and (r_kq150["종가"] < r_kq150["MA40"])
    is_inv_overheat_or_panic = (r_k200["MA40_Disparity"] <= 88.5) or (r_kinv["MA40_Disparity"] >= 105.0)
    if is_all_long_down and (not is_kospi_cross_down or is_inv_overheat_or_panic):
        lines.append("\n🛡️ <b>[종합 포지션 결론]</b> <b>100% 현금 관망 유지</b> (정방향 하락세 + 인버스 과열/미동조 리스크 회피)")
    elif r_k200["MA40_Slope5"] >= EMA40_SLOPE_THRESHOLD and r_k200["MA40_Disparity"] <= 115.0 and not is_kospi_cross_down:
        lines.append("\n🚀 <b>[종합 포지션 결론]</b> <b>KOSPI 정방향 적극 공략</b> (40EMA 상향확정 & 3트랙 매수 승인)")
    elif is_kospi_cross_down and r_kinv["MA40_Disparity"] <= 105.0:
        lines.append("\n📉 <b>[종합 포지션 결론]</b> <b>KOSPI 인버스 헤지 가동</b> (4대 교차 하락 확인 & 단기 20일 헤지)")
    else:
        lines.append("\n⚖️ <b>[종합 포지션 결론]</b> <b>보수적 관망 / 기존 보유분 관리</b> (원칙 외 무리한 추격 매수 금지)")


    base_report = "\n".join(lines)
    if include_ai:
        advice = generate_gemini_advice(base_report)
        full_report = f"{base_report}\n\n━━━━━━━━━━━━━━━━━━━━━\n{advice}"
        if len(full_report) > 3900:
            full_report = full_report[:3850] + "\n\n...(메시지 길이 제한으로 축약됨)"
        return full_report

    return base_report


def run_intraday_mode(include_ai: bool = False) -> None:
    """장중 모드 진입점: 1회 분석 리포트 생성, 콘솔 출력 및 텔레그램 발송 (include_ai 옵션 지원)"""
    print("=" * 65)
    mode_title = "⚡ [장중 실시간 브리핑 모드 실행 (Gemini AI 분석 포함)]" if include_ai else "⚡ [장중 실시간 브리핑 모드 실행]"
    print(mode_title)
    print("=" * 65)
    trigger_title = "장중 실시간 AI 브리핑" if include_ai else "장중 실시간 브리핑"
    report_msg = build_intraday_report(trigger_label=trigger_title, include_ai=include_ai)
    print(report_msg)
    ok = send_telegram_message(report_msg)
    if ok:
        print("\n✅ 텔레그램 발송 완료!")
    else:
        print("\n⚠️ 텔레그램 발송 실패 (콘솔 출력 완료)")



# ==============================================================================
# 5. 장마감 종합 분석 모드 (run_closing_mode)
# ==============================================================================
def evaluate_daily_market(latest_date: pd.Timestamp, cross_signals: List[Dict[str, Any]], analyzed_dfs: Dict[str, pd.DataFrame], target_etfs: List[Dict[str, Any]], days: int = 5) -> List[Dict[str, Any]]:
    """최근 n거래일 기준 KOSPI 및 KOSDAQ 장세와 EMA 40 기울기·이격도·조기경고 종합 분석"""
    sample_df = list(analyzed_dfs.values())[0]
    recent_dates = list(sorted(sample_df[sample_df["dt"] <= latest_date]["dt"].drop_duplicates().tail(days), reverse=True))

    eval_list = []
    for target_date in recent_dates:
        d_str = target_date.strftime("%Y-%m-%d")
        mkt_results = {}

        for mkt in ["KOSPI", "KOSDAQ"]:
            mkt_etfs = [e for e in target_etfs if e["market"] == mkt]
            long_etfs = [e for e in mkt_etfs if e["type"] == "정방향"]
            short_etfs = [e for e in mkt_etfs if e["type"] == "인버스"]

            today_cross = [c for c in cross_signals if c["market"] == mkt and c["date_str"] == d_str]
            long_above = all(
                not analyzed_dfs[e["code"]][analyzed_dfs[e["code"]]["dt"] == target_date].empty and
                analyzed_dfs[e["code"]][analyzed_dfs[e["code"]]["dt"] == target_date].iloc[0]["보조_40MA"] == "상승"
                for e in long_etfs if e["code"] in analyzed_dfs
            )
            short_below = all(
                not analyzed_dfs[e["code"]][analyzed_dfs[e["code"]]["dt"] == target_date].empty and
                analyzed_dfs[e["code"]][analyzed_dfs[e["code"]]["dt"] == target_date].iloc[0]["보조_40MA"] == "하락"
                for e in short_etfs if e["code"] in analyzed_dfs
            )
            long_below = all(
                not analyzed_dfs[e["code"]][analyzed_dfs[e["code"]]["dt"] == target_date].empty and
                analyzed_dfs[e["code"]][analyzed_dfs[e["code"]]["dt"] == target_date].iloc[0]["보조_40MA"] == "하락"
                for e in long_etfs if e["code"] in analyzed_dfs
            )
            short_above = all(
                not analyzed_dfs[e["code"]][analyzed_dfs[e["code"]]["dt"] == target_date].empty and
                analyzed_dfs[e["code"]][analyzed_dfs[e["code"]]["dt"] == target_date].iloc[0]["보조_40MA"] == "상승"
                for e in short_etfs if e["code"] in analyzed_dfs
            )

            rep_df = analyzed_dfs.get(long_etfs[0]["code"])
            rep_row = rep_df[rep_df["dt"] == target_date].iloc[0] if (rep_df is not None and not rep_df[rep_df["dt"] == target_date].empty) else None
            rep_slope = float(rep_row["MA40_Slope5"]) if (rep_row is not None and pd.notna(rep_row["MA40_Slope5"])) else 0.0
            rep_disp = float(rep_row["MA40_Disparity"]) if (rep_row is not None and pd.notna(rep_row["MA40_Disparity"])) else 100.0

            if today_cross:
                direction = today_cross[0]["direction"]
                status = "🔴 [최강 상승장] 4대 ETF 동시 상승 교차 포착" if direction == "🔥 시장 상승 통일" else "🔵 [최강 하락장] 4대 ETF 동시 하락 교차 포착"
            else:
                if long_above and short_below:
                    if rep_slope >= EMA40_SLOPE_THRESHOLD:
                        status = f"🔴 [상승 추세] 40EMA 상회·우상향 (기울기 {rep_slope:+.2f}%, 이격 {rep_disp:.1f}%)"
                    else:
                        status = f"🟠 [상승 둔화/박스권] 40EMA 상회이나 기울기 평탄 ({rep_slope:+.2f}%)"
                elif long_below and short_above:
                    if rep_slope <= -EMA40_SLOPE_THRESHOLD:
                        status = f"🔵 [하락 추세] 40EMA 하회·우하향 (기울기 {rep_slope:+.2f}%, 이격 {rep_disp:.1f}%)"
                    else:
                        status = f"🟡 [하락 둔화/바닥탐색] 40EMA 하회이나 기울기 완화 ({rep_slope:+.2f}%)"
                elif long_below and short_below:
                    status = "🟡 [혼조/음의복리 주의] 지수·인버스 동반 40EMA 하회 (매매 보류)"
                else:
                    status = f"⚪ [중립/변곡점 탐색] 40EMA 공방 구간 (이격 {rep_disp:.1f}%)"

            details = []
            early_warnings = []
            for e in mkt_etfs:
                df_item = analyzed_dfs.get(e["code"])
                if df_item is not None:
                    row_item = df_item[df_item["dt"] == target_date]
                    if not row_item.empty:
                        r0 = row_item.iloc[0]
                        c_val = int(r0["종가"])
                        ma_status = r0["보조_40MA"]
                        disp_val = r0["MA40_Disparity"] if pd.notna(r0["MA40_Disparity"]) else 100.0
                        alert_val = r0["EMA40_조기경고"]
                        details.append(f"{e['name']}: {c_val:,}원 ({'40MA위' if ma_status == '상승' else '40MA아래'}, 이격 {disp_val:.1f}%)")
                        if alert_val != "정상":
                            early_warnings.append(f"{e['name']} {alert_val} ({c_val:,}원)")

            mkt_results[mkt] = {
                "status": status,
                "slope": rep_slope,
                "disparity": rep_disp,
                "details": details,
                "early_warnings": early_warnings
            }

        eval_list.append({
            "date": target_date,
            "date_str": d_str,
            "markets": mkt_results
        })

    return eval_list


def create_notion_etf_report_page(start_3y: pd.Timestamp, latest_date: pd.Timestamp, daily_market_eval: List[Dict[str, Any]], turnaround_signals: List[Dict[str, Any]], cross_signals: List[Dict[str, Any]], dual_signals: List[Dict[str, Any]], df_records: pd.DataFrame, target_etfs: List[Dict[str, Any]]) -> Optional[str]:
    """ETF 트렌드 시그널 분석 리포트를 노션 데이터베이스에 새 페이지로 등록"""
    if not NOTION_API_KEY or NOTION_API_KEY.startswith("your_") or not NOTION_DATABASE_ID:
        print("ℹ️ 노션 API 키 또는 데이터베이스 ID가 없어 등록을 건너뜁니다.")
        return None

    print("📝 [노션 데이터베이스] ETF 트렌드 시그널 리포트 페이지 생성 중...")
    title_col_name, existing_props = ensure_notion_database_schema()

    now_dt = datetime.datetime.now()
    now_str = now_dt.strftime("%Y-%m-%d %H:%M:%S")
    latest_date_str = latest_date.strftime("%Y-%m-%d")

    today_signals_df = df_records[df_records["date_str"] == latest_date_str] if not df_records.empty else pd.DataFrame()
    today_sig_count = len(today_signals_df)

    today_eval = daily_market_eval[0] if daily_market_eval else None
    kospi_status_str = today_eval["markets"]["KOSPI"]["status"] if today_eval else "-"
    kosdaq_status_str = today_eval["markets"]["KOSDAQ"]["status"] if today_eval else "-"
    today_early_warnings = []
    if today_eval:
        today_early_warnings = today_eval["markets"]["KOSPI"]["early_warnings"] + today_eval["markets"]["KOSDAQ"]["early_warnings"]

    properties = {
        title_col_name: {
            "title": [{"text": {"content": f"[{latest_date_str}] 8대 ETF 트렌드(+) 및 40EMA 시그널 리포트 (당일 {today_sig_count}건)"}}]
        },
        "날짜": {"date": {"start": latest_date_str}},
        "시장구분": {"select": {"name": "ETF 트렌드 시그널"}},
        "포착종목수": {"number": today_sig_count},
        "등록시간": {"rich_text": [{"text": {"content": now_str}}]}
    }
    if "KOSPI장세" in existing_props:
        properties["KOSPI장세"] = {"rich_text": [{"text": {"content": kospi_status_str[:100]}}]}
    if "KOSDAQ장세" in existing_props:
        properties["KOSDAQ장세"] = {"rich_text": [{"text": {"content": kosdaq_status_str[:100]}}]}
    if "조기경고수" in existing_props:
        properties["조기경고수"] = {"number": len(today_early_warnings)}

    blocks = []

    # 1. 최상단 요약 Callout 블록
    summary_lines = [
        f"🤖 [8대 ETF 트렌드(+) & 40EMA 필터 분석 보고서]",
        f"📅 기준 거래일: {latest_date_str} (분석 구간: {start_3y.strftime('%Y-%m-%d')} ~ {latest_date_str})",
        f"• KOSPI 장세: {kospi_status_str}",
        f"• KOSDAQ 장세: {kosdaq_status_str}",
        f"• 당일 신규 (+) 시그널: {today_sig_count}건 | 당일 40EMA 돌파/이탈 조기경고: {len(today_early_warnings)}건"
    ]
    if today_early_warnings:
        summary_lines.append("• ⚠️ 당일 40EMA 조기경고 상세: " + " / ".join(today_early_warnings))

    blocks.append({
        "object": "block",
        "type": "callout",
        "callout": {
            "rich_text": [{"type": "text", "text": {"content": "\n".join(summary_lines)}}],
            "icon": {"emoji": "📊"}
        }
    })

    # 2. 최근 5영업일 신규 포착 시그널 타임라인
    blocks.append({
        "object": "block",
        "type": "heading_2",
        "heading_2": {"rich_text": [{"type": "text", "text": {"content": "🚨 최근 5영업일 신규 포착 시그널"}}]}
    })

    recent_5d_cutoff = daily_market_eval[-1]["date"] if daily_market_eval else latest_date
    df_recent_sigs = df_records[df_records["date"] >= recent_5d_cutoff].sort_values(by=["date", "code"], ascending=[False, True]) if not df_records.empty else pd.DataFrame()

    if not df_recent_sigs.empty:
        for _, r in df_recent_sigs.iterrows():
            f_tag = r.get("EMA40_Filter") or r.get("ema40_filter", "-")
            line_txt = f"• {r['date_str']} | {r['name']} | {r['signal']} | 종가 {r['close']:,}원 | [{f_tag}] (이격 {r['disparity']:.1f}%, 기울기 {r['slope']:+.2f}%)"
            blocks.append({
                "object": "block",
                "type": "bulleted_list_item",
                "bulleted_list_item": {"rich_text": [{"type": "text", "text": {"content": line_txt}}]}
            })
    else:
        blocks.append({
            "object": "block",
            "type": "bulleted_list_item",
            "bulleted_list_item": {"rich_text": [{"type": "text", "text": {"content": "• 최근 5영업일 내 발생한 (+) 시그널 없음"}}]}
        })

    # 3. 최근 5영업일 시장 장세 종합 판단
    blocks.append({
        "object": "block",
        "type": "heading_2",
        "heading_2": {"rich_text": [{"type": "text", "text": {"content": "📌 최근 5영업일 시장 장세 종합 판단"}}]}
    })

    for day in daily_market_eval:
        d_str = day["date_str"]
        is_today = (d_str == latest_date_str)
        tag = f"{d_str} (당일)" if is_today else d_str
        for mkt, data in day["markets"].items():
            warn_str = f"\n  - ⚠️ 조기경고: {', '.join(data['early_warnings'])}" if data["early_warnings"] else ""
            txt = f"[{tag}] {mkt}: {data['status']}\n  - 종목별 상태: {', '.join(data['details'])}{warn_str}"
            blocks.append({
                "object": "block",
                "type": "bulleted_list_item",
                "bulleted_list_item": {"rich_text": [{"type": "text", "text": {"content": txt[:1900]}}]}
            })

    # 4. 턴어라운드 및 4대 ETF 동시 교차 시그널
    blocks.append({
        "object": "block",
        "type": "heading_2",
        "heading_2": {"rich_text": [{"type": "text", "text": {"content": "🔄 시장 턴어라운드 & 🔥 4대 ETF 동시 교차 포착 시그널 (최근 10건)"}}]}
    })

    if turnaround_signals:
        for t in turnaround_signals[:5]:
            names_str = ", ".join([f"{item['name']}({item['signal']})" for item in t["items"]])
            txt = f"[턴어라운드] {t['curr_date']} | {t['market']} | {t['turnaround_type']} (직전: {t['prev_date']} {t['prev_direction']}) → {names_str}"
            blocks.append({
                "object": "block",
                "type": "bulleted_list_item",
                "bulleted_list_item": {"rich_text": [{"type": "text", "text": {"content": txt[:1900]}}]}
            })

    if cross_signals:
        for cs in cross_signals[:10]:
            items_summary = ", ".join([f"{item['name']}({item['signal']} {item['close']:,}원)" for item in cs["items"]])
            txt = f"[4대교차] {cs['date_str']} | {cs['market']} | {cs['direction']} → {items_summary}"
            blocks.append({
                "object": "block",
                "type": "bulleted_list_item",
                "bulleted_list_item": {"rich_text": [{"type": "text", "text": {"content": txt[:1900]}}]}
            })

    # 5. 종목별 최근 10년 시그널 상세 이력
    blocks.append({
        "object": "block",
        "type": "heading_2",
        "heading_2": {"rich_text": [{"type": "text", "text": {"content": f"📊 최근 {ANALYSIS_YEARS}년 종목별 (+) 시그널 상세 이력"}}]}
    })

    for item in target_etfs:
        code, name = item["code"], item["name"]
        df_sub = df_records[df_records["code"] == code] if not df_records.empty else pd.DataFrame()
        if not df_sub.empty:
            df_sub = df_sub.sort_values(by="date", ascending=False)
            up_cnt = len(df_sub[df_sub["signal"] == "상승트렌드 +"])
            down_cnt = len(df_sub[df_sub["signal"] == "하락트렌드 +"])
            header_line = f"🔹 {name} (총 {len(df_sub)}회 : 상승+ {up_cnt}회 / 하락+ {down_cnt}회)"
            lines = [
                f"  • {r['date_str']} | {r['name']} | {r['signal']} | 종가 {r['close']:,}원 | [{r.get('EMA40_Filter') or r.get('ema40_filter', '-')}] (이격 {r['disparity']:.1f}%, 기울기 {r['slope']:+.2f}%)"
                for _, r in df_sub.iterrows()
            ]
            blocks.extend(split_lines_to_notion_paragraphs(header_line, lines))
        else:
            blocks.append({
                "object": "block",
                "type": "paragraph",
                "paragraph": {"rich_text": [{"type": "text", "text": {"content": f"🔹 {name}: 최근 {ANALYSIS_YEARS}년 내 시그널 없음"}}]}
            })


    url = "https://api.notion.com/v1/pages"
    first_chunk = blocks[:80]
    remaining_blocks = blocks[80:]

    payload = {
        "parent": {"database_id": NOTION_DATABASE_ID.strip()},
        "properties": properties,
        "children": first_chunk
    }

    try:
        res = requests.post(url, json=payload, headers=NOTION_HEADERS, timeout=15)
        if res.status_code in [200, 201]:
            page_id = res.json().get("id")
            page_url = res.json().get("url")
            if remaining_blocks and page_id:
                append_notion_blocks(page_id, remaining_blocks)
            print(f"  └ ✅ [노션 데이터베이스] ETF 트렌드 시그널 보고서 페이지 등록 완료 (ID: {page_id})")
            return page_url
        print(f"  └ ❌ 노션 등록 실패: Status {res.status_code}, Body: {res.text}")
        return None
    except Exception as e:
        print(f"  └ ❌ 노션 API 요청 중 예외 발생: {e}")
        return None


def build_compact_telegram_briefing(latest_date: pd.Timestamp, daily_market_eval: List[Dict[str, Any]], turnaround_signals: List[Dict[str, Any]], cross_signals: List[Dict[str, Any]], dual_signals: List[Dict[str, Any]], df_records: pd.DataFrame, notion_url: Optional[str], analyzed_dfs: Optional[Dict[str, pd.DataFrame]] = None, gdrive_url: Optional[str] = None) -> str:
    """중요한 메시지를 중심으로 한눈에 들어오도록 요약한 장마감 텔레그램 브리핑 생성 (3-트랙 실전 대응 지침 포함)"""
    latest_str = latest_date.strftime("%Y-%m-%d")
    today_eval = daily_market_eval[0] if daily_market_eval else None

    lines = [
        f"📊 <b>[8대 ETF 마켓 시그널 장마감 종합 요약]</b>",
        f"📅 <b>기준일:</b> {latest_str}",
        "━━━━━━━━━━━━━━━━━━━━━"
    ]

    today_sigs = df_records[df_records["date_str"] == latest_str] if not df_records.empty else pd.DataFrame()
    today_cross = [c for c in cross_signals if c["date_str"] == latest_str]
    today_turn = [t for t in turnaround_signals if t["curr_date"] == latest_str]
    today_dual = [d for d in dual_signals if d["date_str"] == latest_str]

    early_warns = []
    if today_eval:
        early_warns = today_eval["markets"]["KOSPI"]["early_warnings"] + today_eval["markets"]["KOSDAQ"]["early_warnings"]

    lines.append("🚨 <b>[1. 당일 신규 포착 시그널]</b>")
    if today_turn:
        for t in today_turn:
            lines.append(f"• 🔄 <b>{t['market']} 턴어라운드!</b> {html.escape(t['turnaround_type'])}")
    if today_cross:
        for c in today_cross:
            lines.append(f"• 🔥 <b>{c['market']} 4대 ETF 동시 교차!</b> ({html.escape(c['direction'])})")
    elif today_dual:
        for d in today_dual:
            lines.append(f"• ⚡ <b>{html.escape(d['group'])} 동시 포착:</b> {html.escape(d['signal'])}")

    if not today_sigs.empty:
        for _, r in today_sigs.iterrows():
            f_tag = r.get("EMA40_Filter") or r.get("ema40_filter", "-")
            lines.append(
                f"• <b>{html.escape(r['name'])}</b>: {html.escape(r['signal'])} "
                f"({r['close']:,}원 | [{html.escape(f_tag)}])"
            )
    else:
        lines.append("• 당일 신규 트렌드(+) 발생 종목 없음")

    if early_warns:
        for w in early_warns:
            lines.append(f"• {html.escape(w)} <i>(비중 축소/변곡점 주의)</i>")

    lines.append("\n📌 <b>[2. 시장별 40EMA 장세 판단]</b>")
    if today_eval:
        for mkt in ["KOSPI", "KOSDAQ"]:
            m_data = today_eval["markets"][mkt]
            lines.append(f"• <b>{mkt}</b>: {html.escape(m_data['status'])}")

    # 3-트랙 실전 대응 지침 (40EMA 기울기/이격도/4종 동조 반영)
    if analyzed_dfs and "069500" in analyzed_dfs and "114800" in analyzed_dfs:
        r_k200 = analyzed_dfs["069500"].iloc[-1]
        r_k200_tiger = analyzed_dfs["102110"].iloc[-1] if "102110" in analyzed_dfs else r_k200
        r_kinv = analyzed_dfs["114800"].iloc[-1]
        r_kinv_tiger = analyzed_dfs["123310"].iloc[-1] if "123310" in analyzed_dfs else r_kinv
        r_kq150 = analyzed_dfs["229200"].iloc[-1] if "229200" in analyzed_dfs else None

        k_two_day = (len(analyzed_dfs["069500"]) >= 2 and analyzed_dfs["069500"].iloc[-2]["종가"] < analyzed_dfs["069500"].iloc[-2]["MA40"]) and (r_k200["종가"] < r_k200["MA40"])
        k_long_down = (r_k200["종가"] < r_k200["MA40"]) and (r_k200_tiger["종가"] < r_k200_tiger["MA40"])
        k_inv_up = (r_kinv["종가"] > r_kinv["MA40"]) and (r_kinv_tiger["종가"] > r_kinv_tiger["MA40"])
        is_cross_down = k_long_down and k_inv_up

        lines.append("\n🧭 <b>[3. 자동매매 3-트랙 실전 대응 지침]</b>")
        # Track 1
        if k_two_day or r_k200["종가"] < r_k200["MA40"]:
            lines.append(f"• <b>KOSPI 정방향:</b> ⚪ <b>100% 현금 관망</b> (40EMA {int(r_k200['MA40']):,}원 하회, 매수 금지)")
        elif r_k200["MA40_Disparity"] > 115.0:
            lines.append(f"• <b>KOSPI 정방향:</b> ⚠️ <b>홀딩 및 분할익절</b> (이격 {r_k200['MA40_Disparity']:.1f}% 과열, 추격 금지)")
        elif r_k200["MA40_Slope5"] >= EMA40_SLOPE_THRESHOLD and r_k200["MA40_Disparity"] <= 115.0:
            lines.append(f"• <b>KOSPI 정방향:</b> 🟢 <b>적극 매수 승인</b> (기울기 {r_k200['MA40_Slope5']:+.2f}% / 이격 {r_k200['MA40_Disparity']:.1f}%)")
        else:
            lines.append(f"• <b>KOSPI 정방향:</b> 🟡 <b>추세 유지 / 홀딩</b> (40EMA 지지 중)")

        # Track 2
        if r_kq150 is not None:
            q_two_day = (len(analyzed_dfs["229200"]) >= 2 and analyzed_dfs["229200"].iloc[-2]["종가"] < analyzed_dfs["229200"].iloc[-2]["MA40"]) and (r_kq150["종가"] < r_kq150["MA40"])
            if q_two_day or r_kq150["종가"] < r_kq150["MA40"]:
                lines.append(f"• <b>KOSDAQ 정방향:</b> ⚪ <b>100% 현금 관망 / 손절</b> (40EMA {int(r_kq150['MA40']):,}원 하회)")
            elif r_kq150["MA40_Disparity"] > 108.0:
                lines.append(f"• <b>KOSDAQ 정방향:</b> ⚠️ <b>전량 익절 / 신규 금지</b> (고위험 과열 이격 {r_kq150['MA40_Disparity']:.1f}%)")
            elif r_kq150["MA40_Slope5"] >= EMA40_SLOPE_THRESHOLD and r_kq150["MA40_Disparity"] <= 104.0:
                lines.append(f"• <b>KOSDAQ 정방향:</b> 🟢 <b>단기 매수 승인</b> (초입형 이격 {r_kq150['MA40_Disparity']:.1f}% / 20일 스윙)")
            else:
                lines.append(f"• <b>KOSDAQ 정방향:</b> 🟡 <b>신규 보류 / 관망</b> (이격 {r_kq150['MA40_Disparity']:.1f}%)")

        # Track 3
        if r_k200["MA40_Disparity"] <= 88.5 or r_kinv["MA40_Disparity"] >= 111.0:
            lines.append(f"• <b>KOSPI 인버스:</b> 🛑 <b>패닉 바닥 도달 (인버스 전량 익절 & 추격 매수 금지)</b>")
        elif is_cross_down and r_kinv["MA40_Disparity"] <= 105.0:
            lines.append(f"• <b>KOSPI 인버스:</b> 🎯 <b>헤지 매수 승인 (폭락 초입)</b> (4대교차 하락 & 이격 {r_kinv['MA40_Disparity']:.1f}%)")
        else:
            concur_msg = "4대교차 미동조" if not is_cross_down else f"이격 {r_kinv['MA40_Disparity']:.1f}% 과열"
            lines.append(f"• <b>KOSPI 인버스:</b> ⚪ <b>진입 금지 / 단독 휩소 주의</b> ({concur_msg})")

        # 종합 포지션 결론
        is_all_long_down = (r_k200["종가"] < r_k200["MA40"]) and (r_kq150 is None or r_kq150["종가"] < r_kq150["MA40"])
        is_inv_overheat_or_panic = (r_k200["MA40_Disparity"] <= 88.5) or (r_kinv["MA40_Disparity"] >= 105.0)
        if is_all_long_down and (not is_cross_down or is_inv_overheat_or_panic):
            lines.append("🛡️ <b>[종합 결론]</b> <b>100% 현금 관망 유지</b> (정방향 하락세 + 인버스 과열/미동조 리스크 회피)")
        elif r_k200["MA40_Slope5"] >= EMA40_SLOPE_THRESHOLD and r_k200["MA40_Disparity"] <= 115.0 and not is_cross_down:
            lines.append("🚀 <b>[종합 결론]</b> <b>KOSPI 정방향 적극 공략</b> (40EMA 상향확정 & 3트랙 매수 승인)")
        elif is_cross_down and r_kinv["MA40_Disparity"] <= 105.0:
            lines.append("📉 <b>[종합 결론]</b> <b>KOSPI 인버스 헤지 가동</b> (4대 교차 하락 확인 & 단기 20일 헤지)")
        else:
            lines.append("⚖️ <b>[종합 결론]</b> <b>보수적 관망 / 기존 보유분 관리</b> (원칙 외 무리한 추격 매수 금지)")

    lines.append("\n📜 <b>[4. 최근 주요 변곡점 이력]</b>")
    if turnaround_signals:
        rt = turnaround_signals[0]
        lines.append(f"• <b>최근 턴어라운드:</b> {rt['curr_date']} ({rt['market']} {html.escape(rt['curr_direction'])})")
    if cross_signals:
        rc = cross_signals[0]
        lines.append(f"• <b>최근 4대교차:</b> {rc['date_str']} ({rc['market']} {html.escape(rc['direction'])})")
    if dual_signals:
        rd = dual_signals[0]
        lines.append(f"• <b>최근 그룹동시:</b> {rd['date_str']} ({html.escape(rd['group'])} {html.escape(rd['signal'])})")

    if not df_records.empty and len(daily_market_eval) > 1:
        cutoff_5d = daily_market_eval[-1]["date"]
        prev_5d_sigs = df_records[(df_records["date"] >= cutoff_5d) & (df_records["date_str"] != latest_str)]
        if not prev_5d_sigs.empty:
            prev_summary = ", ".join([
                f"{r['date_str'][5:]} {r['name']}({r['signal'].replace('트렌드 ', '')})"
                for _, r in prev_5d_sigs.sort_values(by="date", ascending=False).head(4).iterrows()
            ])
            lines.append(f"• <b>최근 5일내 신호:</b> {html.escape(prev_summary)}")

    if notion_url:
        lines.append(f"\n🔗 <a href='{notion_url}'>노션 상세 리포트 열기</a>")
    if gdrive_url:
        lines.append(f"📁 <a href='{gdrive_url}'>구글 드라이브 엑셀 보고서</a>")


    return "\n".join(lines)


def run_closing_mode() -> None:
    """장마감 종합 분석 모드: 10년 시계열 분석, 엑셀 10년 리포트, 노션 DB 등록, 텔레그램 최종 요약, GitHub 자동 푸시"""
    print("=" * 65)
    print(f"📊 [8대 ETF 최근 {ANALYSIS_YEARS}년 장마감 종합 분석 시작]")
    print("=" * 65)

    analyzed_dfs = {}
    for item in TARGET_ETFS:
        code, name = item["code"], item["name"]
        print(f"[{name} ({code})] 10년 일봉 수집 및 지표/시그널 계산 중...")
        df = fetch_etf_ohlcv(code, count=FETCH_COUNT_CLOSING, vol_mult=1.0)
        if df is not None:
            analyzed_dfs[code] = df
        else:
            print(f"  └ ⚠️ [{name}] 데이터 수집 실패!")

    if not analyzed_dfs:
        raise RuntimeError("분석할 수 있는 시세 데이터가 없습니다.")

    ref_df = list(analyzed_dfs.values())[0]
    latest_date = ref_df["dt"].max()
    start_10y = latest_date - pd.DateOffset(years=ANALYSIS_YEARS)
    latest_date_str = latest_date.strftime("%Y-%m-%d")

    print(f"\n📅 분석 기간: {start_10y.strftime('%Y-%m-%d')} ~ {latest_date_str}")

    records = []
    for item in TARGET_ETFS:
        code, name, group = item["code"], item["name"], item["group"]
        df = analyzed_dfs[code]
        df_10y = df[(df["dt"] >= start_10y) & (df["dt"] <= latest_date)].copy()

        df_plus = df_10y[df_10y["시장트렌드_시그널"].isin(["상승트렌드 +", "하락트렌드 +"])]
        for _, r in df_plus.iterrows():
            records.append({
                "code": code,
                "name": name,
                "group": group,
                "market": item.get("market", ""),
                "type": item.get("type", ""),
                "date": r["dt"],
                "date_str": r["dt"].strftime("%Y-%m-%d"),
                "signal": r["시장트렌드_시그널"],
                "close": int(r["종가"]),
                "vol": int(r["거래량"]),
                "vol_ma": int(r["Vol_MA40"]) if pd.notna(r["Vol_MA40"]) else 0,
                "disparity": float(r["MA40_Disparity"]) if pd.notna(r["MA40_Disparity"]) else 100.0,
                "slope": float(r["MA40_Slope5"]) if pd.notna(r["MA40_Slope5"]) else 0.0,
                "ema40_filter": r["EMA40_필터판정"]
            })

    df_records = pd.DataFrame(records)

    cross_signals = []
    dual_signals = []
    if not df_records.empty:
        for (market, d_str), mkt_date_df in df_records.groupby(["market", "date_str"]):
            long_df = mkt_date_df[mkt_date_df["type"] == "정방향"]
            short_df = mkt_date_df[mkt_date_df["type"] == "인버스"]

            if len(long_df) >= 2 and len(short_df) >= 2:
                long_sigs = set(long_df["signal"])
                short_sigs = set(short_df["signal"])
                is_long_up_short_down = ("상승트렌드 +" in long_sigs) and ("하락트렌드 +" in short_sigs)
                is_long_down_short_up = ("하락트렌드 +" in long_sigs) and ("상승트렌드 +" in short_sigs)

                if is_long_up_short_down or is_long_down_short_up:
                    direction_label = "🔥 시장 상승 통일" if is_long_up_short_down else "❄️ 시장 하락 통일"
                    cross_signals.append({
                        "date_str": d_str,
                        "market": market,
                        "direction": direction_label,
                        "count": len(mkt_date_df),
                        "items": mkt_date_df.to_dict("records")
                    })

        for (d_str, group, sig), group_df in df_records.groupby(["date_str", "group", "signal"]):
            if len(group_df) >= 2:
                dual_signals.append({
                    "date_str": d_str,
                    "group": group,
                    "signal": sig,
                    "items": group_df.to_dict("records")
                })

    cross_signals.sort(key=lambda x: x["date_str"])
    dual_signals.sort(key=lambda x: x["date_str"])

    # 4종 동조 검증 및 인버스 필터 등급 고도화 (KOSPI 4대 교차 하락 동조 & 패닉 바닥 이격도 판정)
    kospi_down_cross_dates = {c["date_str"] for c in cross_signals if c["market"] == "KOSPI" and "하락" in c["direction"]}
    k200_df = analyzed_dfs.get("069500")
    k200_disp_map = {}
    if k200_df is not None:
        for _, row in k200_df.iterrows():
            k200_disp_map[row["dt"].strftime("%Y-%m-%d")] = row["MA40_Disparity"]

    for r in records:
        if r["type"] == "인버스" and "상승트렌드" in r["signal"]:
            d_str = r["date_str"]
            inv_disp = r["disparity"]
            k200_disp = k200_disp_map.get(d_str, 100.0)

            if k200_disp <= 88.5 or inv_disp >= 111.0:
                r["ema40_filter"] = "🛑패닉바닥도달 (인버스 전량 익절 및 추격 금지)"
            elif d_str in kospi_down_cross_dates and inv_disp <= 105.0:
                r["ema40_filter"] = "🎯인버스폭락초입 (헤지 진입 승인)"
            else:
                r["ema40_filter"] = "⚠️단독휩소주의 (진입 금지)"
        r["EMA40_Filter"] = r["ema40_filter"]

    df_records = pd.DataFrame(records)

    turnaround_signals = []
    for mkt in ["KOSPI", "KOSDAQ"]:
        mkt_crosses = [c for c in cross_signals if c["market"] == mkt]
        for i in range(1, len(mkt_crosses)):
            prev = mkt_crosses[i - 1]
            curr = mkt_crosses[i]
            if prev["direction"] != curr["direction"]:
                t_label = "🔄 상승 턴어라운드 (하락→상승 반전)" if curr["direction"] == "🔥 시장 상승 통일" else "🔄 하락 턴어라운드 (상승→하락 반전)"
                turnaround_signals.append({
                    "market": mkt,
                    "turnaround_type": t_label,
                    "prev_date": prev["date_str"],
                    "prev_direction": prev["direction"],
                    "curr_date": curr["date_str"],
                    "curr_direction": curr["direction"],
                    "items": curr["items"]
                })

    turnaround_signals.sort(key=lambda x: x["curr_date"], reverse=True)
    cross_signals.reverse()
    dual_signals.reverse()

    daily_market_eval = evaluate_daily_market(latest_date, cross_signals, analyzed_dfs, TARGET_ETFS)

    # 엑셀 보고서 생성
    output_dir = os.environ.get("OUTPUT_DIR") or BASE_DIR
    os.makedirs(output_dir, exist_ok=True)
    excel_path = os.path.join(output_dir, f"ETF_8종_최근{ANALYSIS_YEARS}년_트렌드시그널_분석.xlsx")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    font_title = Font(name="맑은 고딕", size=15, bold=True, color="1F497D")
    font_sub = Font(name="맑은 고딕", size=10, italic=True, color="595959")
    font_header = Font(name="맑은 고딕", size=11, bold=True, color="FFFFFF")
    font_bold = Font(name="맑은 고딕", size=10, bold=True)
    font_regular = Font(name="맑은 고딕", size=10)

    font_up_plus = Font(name="맑은 고딕", size=10, bold=True, color="C00000")
    font_down_plus = Font(name="맑은 고딕", size=10, bold=True, color="002060")

    fill_header_navy = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    fill_header_slate = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    fill_up_plus = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    fill_down_plus = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    fill_dual = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    fill_turnaround = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")

    border_thin = Side(style="thin", color="D9D9D9")
    border_box = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)

    align_center = Alignment(horizontal="center", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    # 시트 1: 최근 10년 요약
    ws_sum = wb.create_sheet(title=f"최근{ANALYSIS_YEARS}년 요약")
    ws_sum.views.sheetView[0].showGridLines = True
    ws_sum["A1"] = f"8대 ETF 최근 {ANALYSIS_YEARS}년 트렌드(+) 시그널 분석 요약"
    ws_sum["A1"].font = font_title
    ws_sum["A2"] = f"분석 기간: {start_10y.strftime('%Y-%m-%d')} ~ {latest_date_str}"
    ws_sum["A2"].font = font_sub

    ws_sum["A3"] = "📌 [최근 5영업일 시장 장세 종합 판단]"
    ws_sum["A3"].font = Font(name="맑은 고딕", size=11, bold=True, color="1F497D")

    curr_banner_row = 4
    for day in daily_market_eval:
        d_str = day["date_str"]
        is_today = (d_str == latest_date_str)
        label = f"{d_str} (당일)" if is_today else d_str
        k_st = day["markets"]["KOSPI"]["status"]
        q_st = day["markets"]["KOSDAQ"]["status"]
        cell = ws_sum.cell(row=curr_banner_row, column=1, value=f"  • {label} | KOSPI: {k_st} | KOSDAQ: {q_st}")
        cell.font = Font(name="맑은 고딕", size=10, bold=is_today)
        bg_color = "DDEBF7" if is_today else "F2F2F2"
        cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
        curr_banner_row += 1

    if turnaround_signals:
        last_t = turnaround_signals[0]
        t_info = f"🔄 [최근 턴어라운드 발생] {last_t['market']} | {last_t['turnaround_type']} (반전일: {last_t['curr_date']} / 직전: {last_t['prev_date']} [{last_t['prev_direction']}])"
        cell = ws_sum.cell(row=curr_banner_row, column=1, value=t_info)
        cell.font = Font(name="맑은 고딕", size=11, bold=True, color="9C6500")
        cell.fill = fill_turnaround
        curr_banner_row += 1

    header_row = curr_banner_row + 1
    sum_headers = ["종목명", "그룹", "상승트렌드(+) 횟수", "하락트렌드(+) 횟수", "총 (+)시그널 횟수", "시장 4대 ETF 교차포착 참여", "그룹 동시포착 참여"]
    for c_idx, h in enumerate(sum_headers, 1):
        cell = ws_sum.cell(row=header_row, column=c_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header_navy
        cell.alignment = align_center

    for r_idx, item in enumerate(TARGET_ETFS, header_row + 1):
        code, name, group = item["code"], item["name"], item["group"]
        df_sub = df_records[df_records["code"] == code] if not df_records.empty else pd.DataFrame()
        up_cnt = len(df_sub[df_sub["signal"] == "상승트렌드 +"]) if not df_sub.empty else 0
        down_cnt = len(df_sub[df_sub["signal"] == "하락트렌드 +"]) if not df_sub.empty else 0
        cross_cnt = sum(1 for c in cross_signals if any(x["code"] == code for x in c["items"]))
        dual_cnt = sum(1 for d in dual_signals if any(x["code"] == code for x in d["items"]))

        vals = [name, group, up_cnt, down_cnt, up_cnt + down_cnt, cross_cnt, dual_cnt]
        for c_idx, val in enumerate(vals, 1):
            cell = ws_sum.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_regular
            cell.border = border_box
            if c_idx in [1, 2]:
                cell.alignment = align_center
            else:
                cell.alignment = align_right
                cell.number_format = "#,##0"

    # 시트 2: 가장 강력한 동시 시그널
    ws_dual = wb.create_sheet(title="가장 강력한 동시 시그널")
    ws_dual.views.sheetView[0].showGridLines = True
    ws_dual["A1"] = f"🔄 [최신 턴어라운드] & 🔥 KOSPI / KOSDAQ 4대 ETF 교차 동시 발생일 목록 (최근 {ANALYSIS_YEARS}년)"
    ws_dual["A1"].font = font_title

    dual_headers = ["시그널 구분", "발생일자", "시장/방향", "동시 포착 종목", "종가", "당일 거래량", "40EMA 거래량", "비고"]
    for c_idx, h in enumerate(dual_headers, 1):
        cell = ws_dual.cell(row=3, column=c_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header_slate
        cell.alignment = align_center

    curr_row = 4
    font_turnaround = Font(name="맑은 고딕", size=10, bold=True, color="9C6500")

    for t in turnaround_signals:
        for item in t["items"]:
            vals = [
                f"🔄 {t['market']} 턴어라운드", t["curr_date"], t["turnaround_type"], f"{item['name']} ({item['signal']})",
                item["close"], item["vol"], item["vol_ma"], f"🔄 턴어라운드 신호 (직전:{t['prev_date']} {t['prev_direction']})"
            ]
            for c_idx, val in enumerate(vals, 1):
                cell = ws_dual.cell(row=curr_row, column=c_idx, value=val)
                cell.font = font_turnaround if c_idx in [1, 3, 8] else font_regular
                cell.border = border_box
                cell.fill = fill_turnaround
                if c_idx in [1, 2, 3, 4, 8]:
                    cell.alignment = align_center
                else:
                    cell.alignment = align_right
                    cell.number_format = "#,##0"
            curr_row += 1

    for cs in cross_signals:
        for item in cs["items"]:
            vals = [
                f"🔥 {cs['market']} 4대 ETF 교차", cs["date_str"], cs["direction"], f"{item['name']} ({item['signal']})",
                item["close"], item["vol"], item["vol_ma"], "🔥 가장 강력한 신호 (정방향+인버스 교차)"
            ]
            for c_idx, val in enumerate(vals, 1):
                cell = ws_dual.cell(row=curr_row, column=c_idx, value=val)
                cell.font = font_bold if c_idx == 8 else font_regular
                cell.border = border_box
                cell.fill = fill_dual
                if c_idx in [1, 2, 3, 4, 8]:
                    cell.alignment = align_center
                else:
                    cell.alignment = align_right
                    cell.number_format = "#,##0"
                if c_idx in [3, 4]:
                    cell.font = font_up_plus if item["signal"] == "상승트렌드 +" else font_down_plus
            curr_row += 1

    for d in dual_signals:
        for item in d["items"]:
            vals = [
                f"⚡ {d['group']} 동시", d["date_str"], d["signal"], item["name"],
                item["close"], item["vol"], item["vol_ma"], "⚡ 그룹 동시 시그널"
            ]
            for c_idx, val in enumerate(vals, 1):
                cell = ws_dual.cell(row=curr_row, column=c_idx, value=val)
                cell.font = font_regular
                cell.border = border_box
                if c_idx in [1, 2, 3, 4, 8]:
                    cell.alignment = align_center
                else:
                    cell.alignment = align_right
                    cell.number_format = "#,##0"
                if c_idx == 3:
                    cell.font = font_up_plus if d["signal"] == "상승트렌드 +" else font_down_plus
            curr_row += 1

    # 시트 3~10: ETF별 10년 상세 시트
    for item in TARGET_ETFS:
        code, name = item["code"], item["name"]
        df_item = analyzed_dfs[code]
        df_10y_item = df_item[(df_item["dt"] >= start_10y) & (df_item["dt"] <= latest_date)].sort_values(by="dt", ascending=False).reset_index(drop=True)

        ws = wb.create_sheet(title=f"{name} {ANALYSIS_YEARS}년상세")
        ws.views.sheetView[0].showGridLines = True
        ws["A1"] = f"{name} ({code}) 최근 {ANALYSIS_YEARS}년 시그널 분석 상세"
        ws["A1"].font = font_title

        detail_headers = ["날짜", "종가", "시가", "고가", "저가", "거래량", "40EMA 종가", "40EMA 거래량", "트렌드(+) 시그널"]
        for c_idx, h in enumerate(detail_headers, 1):
            cell = ws.cell(row=3, column=c_idx, value=h)
            cell.font = font_header
            cell.fill = fill_header_navy
            cell.alignment = align_center

        for r_idx, r in df_10y_item.iterrows():
            row_num = r_idx + 4
            sig = r["시장트렌드_시그널"]
            vals = [
                r["dt"].strftime("%Y-%m-%d"), int(r["종가"]), int(r["시가"]), int(r["고가"]), int(r["저가"]),
                int(r["거래량"]), int(r["MA40"]) if pd.notna(r["MA40"]) else 0,
                int(r["Vol_MA40"]) if pd.notna(r["Vol_MA40"]) else 0, sig
            ]
            for c_idx, val in enumerate(vals, 1):
                cell = ws.cell(row=row_num, column=c_idx, value=val)
                cell.border = border_box
                cell.font = font_regular
                if c_idx == 1:
                    cell.alignment = align_center
                elif c_idx in [2, 3, 4, 5, 6, 7, 8]:
                    cell.alignment = align_right
                    cell.number_format = "#,##0"
                else:
                    cell.alignment = align_center

                if sig == "상승트렌드 +":
                    cell.fill = fill_up_plus
                    if c_idx == 9:
                        cell.font = font_up_plus
                elif sig == "하락트렌드 +":
                    cell.fill = fill_down_plus
                    if c_idx == 9:
                        cell.font = font_down_plus

    # 열 너비 자동 조정
    for sheet in wb.worksheets:
        for col in sheet.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or "")
                cell_len = sum(2 if ord(c) > 128 else 1 for c in val_str)
                if cell_len > max_len:
                    max_len = cell_len
            sheet.column_dimensions[col_letter].width = max(max_len + 4, 12)

    try:
        wb.save(excel_path)
        print(f"\n✅ 엑셀 파일 저장 완료: {excel_path}")
    except PermissionError:
        alt_excel_path = os.path.join(output_dir, f"ETF_8종_최근{ANALYSIS_YEARS}년_트렌드시그널_분석_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        wb.save(alt_excel_path)
        excel_path = alt_excel_path
        print(f"\n⚠️ 기존 엑셀 파일 사용 중으로 대체 저장 완료: {excel_path}")

    # 구글 드라이브 엑셀 업로드
    gdrive_url = upload_excel_to_gdrive(excel_path)

    # 노션 페이지 등록
    notion_url = create_notion_etf_report_page(start_10y, latest_date, daily_market_eval, turnaround_signals, cross_signals, dual_signals, df_records, TARGET_ETFS)

    # 텔레그램 최종 요약 브리핑 발송 (구글 드라이브 링크 포함)
    briefing_text = build_compact_telegram_briefing(latest_date, daily_market_eval, turnaround_signals, cross_signals, dual_signals, df_records, notion_url, analyzed_dfs=analyzed_dfs, gdrive_url=gdrive_url)
    print("\n[텔레그램 브리핑 메시지]")
    print(briefing_text)
    send_telegram_message(briefing_text)

    # GitHub 자동 커밋·푸시
    auto_git_backup(excel_path, latest_date_str)


# ==============================================================================
# 6. 장중 상주 봇 데몬 모드 (run_daemon_mode)
# ==============================================================================
def build_position_status_report() -> str:
    """position_state.json 기반 현재 보유 슬롯 실시간 평가 손익 브리핑"""
    if not os.path.exists(STATE_FILE):
        return "ℹ️ 아직 생성된 <code>position_state.json</code> 파일이 없습니다. (전 슬롯 현금 관망 중)"

    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            state = json.load(f)
    except Exception as e:
        return f"⚠️ 포지션 상태 파일 로드 오류: {e}"

    lines = [
        "💼 <b>[자동매매 3-트랙 실시간 포지션 현황]</b>",
        f"🔄 <b>상태 파일 갱신:</b> {state.get('updated_at', '-')}",
        "━━━━━━━━━━━━━━━━━━━━━",
    ]

    for slot_key, slot in state.get("slots", {}).items():
        if slot.get("status") == "CASH" or slot.get("qty", 0) <= 0:
            lines.append(f"• <b>{slot_key}</b> ({slot['name']}): ⚪ <b>CASH (현금 관망)</b>")
        else:
            df = fetch_etf_ohlcv(slot["code"], count=30, vol_mult=1.0)
            curr_p = float(df.iloc[-1]["종가"]) if df is not None else float(slot["entry_price"])
            entry_p = float(slot["entry_price"])
            pnl_pct = ((curr_p - entry_p) / entry_p) * 100.0 if entry_p > 0 else 0.0
            pnl_krw = int((curr_p - entry_p) * slot["qty"])
            lines.append(
                f"• <b>{slot_key}</b> ({slot['name']}): 🟢 <b>{slot['status']}</b>\n"
                f"  └ 수량: {slot['qty']:,}주 | 매수가: {int(entry_p):,}원 → 현재가: <b>{int(curr_p):,}원</b>\n"
                f"  └ 평가손익: <b>{pnl_krw:+,}원 ({pnl_pct:+.2f}%)</b> | 보유 {slot.get('hold_days', 0)}일차"
            )

    return "\n".join(lines)


def get_initial_update_offset() -> Optional[int]:
    """봇 시작 시 이전 밀린 메시지를 건너뛰고 최신 offset부터 수신"""
    if not MY_BOT_TOKEN:
        return None
    url = f"https://api.telegram.org/bot{MY_BOT_TOKEN}/getUpdates"
    try:
        res = requests.get(url, params={"offset": -1, "timeout": 3}, timeout=5)
        if res.status_code == 200:
            results = res.json().get("result", [])
            if results:
                return results[-1]["update_id"] + 1
    except Exception:
        pass
    return None


def run_daemon_mode() -> None:
    """장중 백그라운드 상주 데몬: 평일 정기 브리핑(10/13/15시) + 텔레그램 Long Polling 실시간 인터랙션"""
    print("=" * 65)
    print("🤖 [장중 통합 봇 데몬 가동 시작]")
    print(f"   • 평일 자동 브리핑 시각: {', '.join(SCHEDULE_TIMES)}")
    print("   • 실시간 대화형 명령어: /분석, /ai분석, /포지션, /상태")
    print("=" * 65)

    offset = get_initial_update_offset()
    sent_history: Dict[str, List[str]] = {}

    send_telegram_message(
        "🟢 <b>[통합 마켓시그널 봇 데몬 가동]</b>\n"
        f"• <b>정기 자동 브리핑:</b> 평일 {', '.join(SCHEDULE_TIMES)}\n"
        "• <b>명령어 안내:</b> <code>/분석</code>, <code>/ai분석</code>, <code>/포지션</code>, <code>/상태</code>"
    )

    while True:
        try:
            now_dt = datetime.datetime.now()
            today_str = now_dt.strftime("%Y-%m-%d")
            hhmm_str = now_dt.strftime("%H:%M")

            # 1. 평일 정기 브리핑 체크
            if now_dt.weekday() < 5 and hhmm_str in SCHEDULE_TIMES:
                sent_today = sent_history.setdefault(today_str, [])
                if hhmm_str not in sent_today:
                    if not is_market_open_today(mode="intraday"):
                        print(f"\nℹ️ [{today_str} {hhmm_str}] 휴장일 감지: 정기 브리핑을 건너뜁니다.")
                        sent_today.append(hhmm_str)
                    else:
                        print(f"\n⏰ [{today_str} {hhmm_str}] 정기 장중 분석 리포트 자동 발송 중...")
                        report_txt = build_intraday_report(trigger_label=f"{hhmm_str} 정기 브리핑")
                        send_telegram_message(report_txt)
                        sent_today.append(hhmm_str)

            # 2. 텔레그램 실시간 인터랙션 Long Polling
            if MY_BOT_TOKEN and MY_CHAT_ID:
                url = f"https://api.telegram.org/bot{MY_BOT_TOKEN}/getUpdates"
                params: Dict[str, Any] = {"timeout": 15}
                if offset is not None:
                    params["offset"] = offset

                res = requests.get(url, params=params, timeout=20)
                if res.status_code == 200:
                    updates = res.json().get("result", [])
                    for upd in updates:
                        offset = upd["update_id"] + 1
                        msg = upd.get("message") or upd.get("edited_message") or {}
                        chat_id = str(msg.get("chat", {}).get("id", ""))
                        text = (msg.get("text") or "").strip()

                        if chat_id != str(MY_CHAT_ID) or not text:
                            continue

                        cmd = text.split()[0].lower()
                        print(f"📩 [명령어 수신] {now_dt.strftime('%H:%M:%S')} -> {text}")

                        if cmd in ("/ai분석", "ai분석", "/ai", "ai"):
                            send_telegram_message(
                                "🤖 <i>8대 ETF 실시간 지표를 수집하여 Gemini AI 장세 대응 조언을 생성 중입니다...</i>",
                                chat_id
                            )
                            if not is_market_open_today(mode="intraday"):
                                notice = f"ℹ️ <b>[휴장일 안내]</b> 오늘({today_str})은 주식시장 휴장일입니다.\n(최근 거래일 종가 기준으로 AI 분석을 제공합니다.)\n\n"
                                reply = notice + build_intraday_report(trigger_label="실시간 /ai분석 요청 (휴장일)", include_ai=True)
                            else:
                                reply = build_intraday_report(trigger_label="실시간 /ai분석 요청", include_ai=True)
                            send_telegram_message(reply, chat_id)
                        elif cmd in ("/분석", "분석", "/report", "/시황"):
                            if not is_market_open_today(mode="intraday"):
                                notice = f"ℹ️ <b>[휴장일 안내]</b> 오늘({today_str})은 주식시장 휴장일입니다.\n(최근 거래일 종가 기준으로 분석을 제공합니다.)\n\n"
                                reply = notice + build_intraday_report(trigger_label="실시간 /분석 요청 (휴장일)")
                            else:
                                reply = build_intraday_report(trigger_label="실시간 /분석 요청")
                            send_telegram_message(reply, chat_id)
                        elif cmd in ("/포지션", "포지션", "/잔고", "/status"):
                            reply = build_position_status_report()
                            if not is_market_open_today(mode="intraday"):
                                reply = f"ℹ️ <i>[휴장일 안내] 오늘({today_str})은 주식시장 휴장일입니다. (최근 영업일 종가 기준)</i>\n\n" + reply
                            send_telegram_message(reply, chat_id)
                        elif cmd in ("/상태", "/도움말", "/help", "/start"):
                            is_open = is_market_open_today(mode="intraday")
                            market_desc = "🟢 개장일 (정상 운영 중)" if is_open else "🔴 휴장일 (브리핑 스킵 / 명령어 유연 응대)"
                            holiday_notice = "" if is_open else f"\n⚠️ <b>안내:</b> 오늘({today_str})은 주식시장 휴장일입니다."
                            sent_list = ", ".join(sent_history.get(today_str, [])) or "없음"
                            help_txt = (
                                "🤖 <b>[통합 마켓시그널 봇 상태 안내]</b>\n"
                                f"• <b>현재 서버 시각:</b> {now_dt.strftime('%Y-%m-%d %H:%M:%S')}\n"
                                f"• <b>금일 증시 상태:</b> {market_desc}{holiday_notice}\n"
                                f"• <b>평일 자동 브리핑:</b> {', '.join(SCHEDULE_TIMES)}\n"
                                f"• <b>금일 발송 완료:</b> {sent_list}\n\n"
                                "💡 <b>[사용 가능한 명령어]</b>\n"
                                "• <code>/분석</code> : 8대 ETF 실시간 시세·40EMA·3트랙 지침 즉시 조회\n"
                                "• <code>/ai분석</code> : Gemini AI 10년 통계 기반 실전 대응 조언 포함 조회\n"
                                "• <code>/포지션</code> : 보유 슬롯 실시간 평가 손익 조회\n"
                                "• <code>/상태</code> : 봇 정상 동작 여부 확인"
                            )
                            send_telegram_message(help_txt, chat_id)
            else:
                time.sleep(10)

        except KeyboardInterrupt:
            print("\n🛑 사용자 요청으로 데몬을 종료합니다.")
            break
        except Exception as e:
            print(f"⚠️ 폴링 루프 일시 오류 (5초 후 재시도): {e}")
            time.sleep(5)


# ==============================================================================
# 7. 메인 진입점 및 CLI 스마트 분기
# ==============================================================================
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="통합 마켓 시그널 리포트 엔진 (장중/장마감/데몬 모드)")
    parser.add_argument("--intraday", action="store_true", help="장중 실시간 브리핑 1회 발송 후 종료")
    parser.add_argument("--closing", action="store_true", help="장마감 10년 종합 분석 실행 (엑셀/노션/텔레그램/Git 백업)")
    parser.add_argument("--daemon", action="store_true", help="장중 정기 브리핑 및 텔레그램 명령어 응답 백그라운드 데몬 가동")
    parser.add_argument("--ai", action="store_true", help="Gemini AI 10년 통계 기반 실전 대응 조언 포함")
    parser.add_argument("--force", action="store_true", help="휴장일 검사를 건너뛰고 강제 실행")

    args = parser.parse_args()

    # 단발성 실행 모드 (--closing, --intraday, 인자 없는 기본 실행): 휴장일 가드 점검
    if not args.daemon and not args.force:
        target_mode = "closing" if args.closing else ("intraday" if args.intraday else "auto")
        if target_mode == "auto":
            now_check = datetime.datetime.now()
            if datetime.time(9, 0) <= now_check.time() < datetime.time(15, 30):
                target_mode = "intraday"
            else:
                target_mode = "closing"

        # CLI에서 명시적으로 --intraday 또는 --ai 옵션을 준 경우에는 휴장일이라도 최근 종가 기준으로 발송 및 테스트 진행
        if not (args.intraday or args.ai):
            if not is_market_open_today(mode=target_mode):
                today_str = datetime.datetime.now().strftime("%Y-%m-%d")
                print(f"ℹ️ [휴장일 감지] 오늘({today_str})은 주식시장 휴장일이므로 작업을 건너뜁니다. (테스트 실행은 --force 추가)")
                sys.exit(0)


    if args.intraday or (args.ai and not args.closing and not args.daemon):
        run_intraday_mode(include_ai=args.ai)
    elif args.closing:
        run_closing_mode()
    elif args.daemon:
        run_daemon_mode()
    else:
        # 스마트 자동 분기: 평일 09:00 ~ 15:30 사이는 장중 모드, 그 외는 장마감 종합 분석 모드
        now = datetime.datetime.now()
        is_weekday = (now.weekday() < 5)
        is_market_hours = (datetime.time(9, 0) <= now.time() < datetime.time(15, 30))

        if is_weekday and is_market_hours:
            print(f"ℹ️ [스마트 분기] 평일 정규장 시간대({now.strftime('%H:%M')}) 감지 -> 장중 실시간 모드(--intraday) 자동 실행")
            run_intraday_mode(include_ai=args.ai)
        else:
            print(f"ℹ️ [스마트 분기] 정규장 마감 시간대({now.strftime('%H:%M')}) 감지 -> 장마감 종합 분석 모드(--closing) 자동 실행")
            run_closing_mode()

