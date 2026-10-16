import sys
import os
import time
import html
import datetime
import requests
import xml.etree.ElementTree as ET
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# 콘솔 인코딩 및 버퍼링 설정 (UTF-8 출력 보장)
if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace', line_buffering=True)
        sys.stderr.reconfigure(encoding='utf-8', errors='replace', line_buffering=True)
    except Exception:
        pass

# .env 파일에서 환경변수 로드
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# 텔레그램 설정
MY_BOT_TOKEN = os.environ.get("MY_BOT_TOKEN", "")
MY_CHAT_ID = os.environ.get("MY_CHAT_ID", "")

# 노션(Notion) API 설정
NOTION_API_KEY = os.environ.get("NOTION_API_KEY", "")
NOTION_DATABASE_ID = os.environ.get("NOTION_DATABASE_ID", "")
NOTION_VERSION = "2022-06-28"

NOTION_HEADERS = {
    "Authorization": f"Bearer {NOTION_API_KEY}",
    "Content-Type": "application/json",
    "Notion-Version": NOTION_VERSION
}

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
}

# EMA 40 필터 파라미터 상수
EMA40_SLOPE_THRESHOLD = 0.15      # 5일 EMA40 기울기 임계값 (±0.15% 미만은 횡보/평탄화로 간주)
EMA40_OVERHEAT_DISPARITY = 104.0  # 이격도 104% 초과는 단기 과열 주의

# 분석 대상 8대 ETF 정의
TARGET_ETFS = [
    {'code': '069500', 'name': 'KODEX 200', 'group': 'KOSPI 200', 'market': 'KOSPI', 'type': '정방향'},
    {'code': '102110', 'name': 'TIGER 200', 'group': 'KOSPI 200', 'market': 'KOSPI', 'type': '정방향'},
    {'code': '114800', 'name': 'KODEX 인버스', 'group': 'KOSPI 200 인버스', 'market': 'KOSPI', 'type': '인버스'},
    {'code': '123310', 'name': 'TIGER 인버스', 'group': 'KOSPI 200 인버스', 'market': 'KOSPI', 'type': '인버스'},
    {'code': '229200', 'name': 'KODEX 코스닥150', 'group': 'KOSDAQ 150', 'market': 'KOSDAQ', 'type': '정방향'},
    {'code': '232080', 'name': 'TIGER 코스닥150', 'group': 'KOSDAQ 150', 'market': 'KOSDAQ', 'type': '정방향'},
    {'code': '251340', 'name': 'KODEX 코스닥150선물인버스', 'group': 'KOSDAQ 150 인버스', 'market': 'KOSDAQ', 'type': '인버스'},
    {'code': '250780', 'name': 'TIGER 코스닥150선물인버스', 'group': 'KOSDAQ 150 인버스', 'market': 'KOSDAQ', 'type': '인버스'}
]


def ensure_notion_database_schema():
    """노션 데이터베이스 스키마 검사 및 확장 컬럼 자동 생성"""
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
                else:
                    print(f"⚠️ 노션 스키마 업데이트 경고: {patch_res.status_code} - {patch_res.text}")
        return title_col_name, existing_props
    except Exception as e:
        print(f"노션 스키마 검사 오류: {e}")
        return "제목", {}


def append_notion_blocks(page_id, blocks):
    """노션 페이지 하위 블록 배치 추가 (에러 로깅 및 Rate Limit 방어 포함)"""
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
            print(f"  └ ⚠️ 노션 블록 추가 중 예외 발생 (Chunk {i // batch_size + 1}): {e}")


def split_lines_to_notion_paragraphs(header_line, lines, max_chars=1800):
    """노션 블록 2,000자 제한으로 인한 중간 잘림을 방지하기 위해 줄 단위로 안전하게 분할"""
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


def fetch_historical_ohlcv(code, count=1200):
    """네이버 차트 XML API를 이용하여 일별 시세 수집 및 40EMA·기울기·이격도 연산"""
    url_fchart = f"https://fchart.stock.naver.com/sise.nhn?symbol={code}&timeframe=day&count={count}&requestType=0"
    res = requests.get(url_fchart, headers=HEADERS, timeout=10)
    if res.status_code == 200:
        root = ET.fromstring(res.text)
        items = root.findall('.//item')
        rows = []
        for item in items:
            p = item.attrib['data'].split('|')
            if len(p) >= 6:
                d_str = f"{p[0][:4]}-{p[0][4:6]}-{p[0][6:8]}"
                op, hp, lp, cp, vol = float(p[1]), float(p[2]), float(p[3]), float(p[4]), float(p[5])
                rows.append({'날짜': d_str, '시가': op, '고가': hp, '저가': lp, '종가': cp, '거래량': vol})
        if rows:
            df = pd.DataFrame(rows)
            df['날짜'] = pd.to_datetime(df['날짜'])
            df = df.sort_values(by='날짜', ascending=True).reset_index(drop=True)
            df['MA40'] = df['종가'].ewm(span=40, min_periods=20, adjust=False).mean()
            df['Vol_MA40'] = df['거래량'].ewm(span=40, min_periods=20, adjust=False).mean()
            # EMA 40의 5거래일 변화율(기울기, %) 및 이격도(%) 추가
            df['MA40_Slope5'] = ((df['MA40'] - df['MA40'].shift(5)) / df['MA40'].shift(5)) * 100.0
            df['MA40_Disparity'] = (df['종가'] / df['MA40']) * 100.0
            return df
    return None


def calculate_all_signals(df):
    """6개 보조 시그널, 통합 트렌드(+) 시그널, EMA 40 조기경고 및 필터 상태 판정"""
    sig1_40ma, sig2_gap, sig3_gap_close, sig4_high, sig5_low = [], [], [], [], []
    ema40_early_alerts = []

    for idx in range(len(df)):
        if pd.isna(df.loc[idx, 'MA40']):
            sig1_40ma.append("데이터 부족")
        else:
            close = df.loc[idx, '종가']
            ma40 = df.loc[idx, 'MA40']
            if close > ma40:
                sig1_40ma.append("상승")
            elif close < ma40:
                sig1_40ma.append("하락")
            else:
                sig1_40ma.append("보합")

        # EMA 40 2거래일 연속 하향이탈 / 상향돌파 조기 포착 (변곡점 후행성 방어)
        if idx < 3 or pd.isna(df.loc[idx, 'MA40']) or pd.isna(df.loc[idx - 2, 'MA40']):
            ema40_early_alerts.append("정상")
        else:
            c0, m0 = df.loc[idx, '종가'], df.loc[idx, 'MA40']
            c1, m1 = df.loc[idx - 1, '종가'], df.loc[idx - 1, 'MA40']
            c2, m2 = df.loc[idx - 2, '종가'], df.loc[idx - 2, 'MA40']
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

        if d1['종가'] < d0['시가']:
            sig2_gap.append("갭상승")
        elif d1['종가'] > d0['시가']:
            sig2_gap.append("갭하락")
        else:
            sig2_gap.append("보합")

        if (d1['종가'] < d0['시가']) and (d0['종가'] < d0['시가']):
            sig3_gap_close.append("하락마감")
        elif (d1['종가'] > d0['시가']) and (d0['종가'] > d0['시가']):
            sig3_gap_close.append("상승마감")
        else:
            sig3_gap_close.append("기타")

        if d0['고가'] > d1['고가']:
            sig4_high.append("고가상승")
        elif d0['고가'] < d1['고가']:
            sig4_high.append("고가하락")
        else:
            sig4_high.append("보합")

        if d0['저가'] > d1['저가']:
            sig5_low.append("저가상승")
        elif d0['저가'] < d1['저가']:
            sig5_low.append("저가하락")
        else:
            sig5_low.append("보합")

    sig6_vol = []
    for idx in range(len(df)):
        vol = df.loc[idx, '거래량'] if ('거래량' in df.columns and pd.notna(df.loc[idx, '거래량'])) else 0
        vol_ma = df.loc[idx, 'Vol_MA40'] if ('Vol_MA40' in df.columns and pd.notna(df.loc[idx, 'Vol_MA40'])) else 0
        if pd.isna(vol_ma) or vol_ma == 0:
            sig6_vol.append("데이터 부족")
        elif vol > vol_ma:
            sig6_vol.append("거래량 상승")
        elif vol < vol_ma:
            sig6_vol.append("거래량 하락")
        else:
            sig6_vol.append("보합")

    df['보조_40MA'] = sig1_40ma
    df['보조_Gap'] = sig2_gap
    df['보조_갭과시가종가'] = sig3_gap_close
    df['보조_고가비교'] = sig4_high
    df['보조_저가비교'] = sig5_low
    df['보조_거래량비교'] = sig6_vol
    df['EMA40_조기경고'] = ema40_early_alerts

    market_trends = []
    ema40_filter_tags = []

    for idx in range(len(df)):
        if idx < 2 or df.loc[idx, '보조_40MA'] == "데이터 부족":
            market_trends.append("데이터 부족")
            ema40_filter_tags.append("-")
            continue

        d0 = df.loc[idx]
        d1 = df.loc[idx - 1]
        d2 = df.loc[idx - 2]

        c_down_d2 = (d2['보조_Gap'] == "갭상승") and (d2['보조_갭과시가종가'] == "하락마감")
        c_down_d1 = (d1['보조_고가비교'] == "고가하락") and (d1['보조_저가비교'] == "저가하락")
        c_down_d0 = (d0['보조_고가비교'] == "고가하락") and (d0['보조_저가비교'] == "저가하락") and (d0['보조_40MA'] == "하락")
        is_downtrend = c_down_d2 and c_down_d1 and c_down_d0

        c_up_d2 = (d2['보조_Gap'] == "갭하락") and (d2['보조_갭과시가종가'] == "상승마감")
        c_up_d1 = (d1['보조_고가비교'] == "고가상승") and (d1['보조_저가비교'] == "저가상승")
        c_up_d0 = (d0['보조_고가비교'] == "고가상승") and (d0['보조_저가비교'] == "저가상승") and (d0['보조_40MA'] == "상승")
        is_uptrend = c_up_d2 and c_up_d1 and c_up_d0

        vol_0 = d0['거래량'] if ('거래량' in d0 and pd.notna(d0['거래량'])) else 0
        vol_ma40_0 = d0['Vol_MA40'] if ('Vol_MA40' in d0 and pd.notna(d0['Vol_MA40'])) else 0
        is_vol_higher = (vol_0 > vol_ma40_0) and (vol_ma40_0 > 0)

        slope = d0['MA40_Slope5'] if pd.notna(d0['MA40_Slope5']) else 0.0
        disp = d0['MA40_Disparity'] if pd.notna(d0['MA40_Disparity']) else 100.0

        if is_downtrend and is_vol_higher:
            market_trends.append("하락트렌드 +")
            if slope <= -EMA40_SLOPE_THRESHOLD:
                ema40_filter_tags.append("✅40EMA하향확정")
            else:
                ema40_filter_tags.append("⚠️40EMA평탄(휩소주의)")
        elif is_uptrend and is_vol_higher:
            market_trends.append("상승트렌드 +")
            if abs(slope) < EMA40_SLOPE_THRESHOLD:
                ema40_filter_tags.append("⚠️40EMA평탄(휩소주의)")
            elif disp > EMA40_OVERHEAT_DISPARITY:
                ema40_filter_tags.append(f"⚠️단기과열(이격{disp:.1f}%)")
            elif slope >= EMA40_SLOPE_THRESHOLD:
                ema40_filter_tags.append("✅40EMA상향확정")
            else:
                ema40_filter_tags.append("⚠️40EMA역배열반등")
        else:
            market_trends.append("해당없음")
            ema40_filter_tags.append("-")

    df['시장트렌드_시그널'] = market_trends
    df['EMA40_필터판정'] = ema40_filter_tags
    return df


def evaluate_daily_market(latest_date, cross_signals, analyzed_dfs, target_etfs, days=5):
    """최근 n거래일 기준 KOSPI 및 KOSDAQ 장세와 EMA 40 기울기·이격도·조기경고 종합 분석"""
    sample_df = list(analyzed_dfs.values())[0]
    recent_dates = list(sorted(sample_df[sample_df['날짜'] <= latest_date]['날짜'].drop_duplicates().tail(days), reverse=True))

    eval_list = []

    for target_date in recent_dates:
        d_str = target_date.strftime('%Y-%m-%d')
        mkt_results = {}

        for mkt in ['KOSPI', 'KOSDAQ']:
            mkt_etfs = [e for e in target_etfs if e['market'] == mkt]
            long_etfs = [e for e in mkt_etfs if e['type'] == '정방향']
            short_etfs = [e for e in mkt_etfs if e['type'] == '인버스']

            today_cross = [c for c in cross_signals if c['market'] == mkt and c['date_str'] == d_str]
            long_above = all(
                not analyzed_dfs[e['code']][analyzed_dfs[e['code']]['날짜'] == target_date].empty and
                analyzed_dfs[e['code']][analyzed_dfs[e['code']]['날짜'] == target_date].iloc[0]['보조_40MA'] == '상승'
                for e in long_etfs if e['code'] in analyzed_dfs
            )
            short_below = all(
                not analyzed_dfs[e['code']][analyzed_dfs[e['code']]['날짜'] == target_date].empty and
                analyzed_dfs[e['code']][analyzed_dfs[e['code']]['날짜'] == target_date].iloc[0]['보조_40MA'] == '하락'
                for e in short_etfs if e['code'] in analyzed_dfs
            )
            long_below = all(
                not analyzed_dfs[e['code']][analyzed_dfs[e['code']]['날짜'] == target_date].empty and
                analyzed_dfs[e['code']][analyzed_dfs[e['code']]['날짜'] == target_date].iloc[0]['보조_40MA'] == '하락'
                for e in long_etfs if e['code'] in analyzed_dfs
            )
            short_above = all(
                not analyzed_dfs[e['code']][analyzed_dfs[e['code']]['날짜'] == target_date].empty and
                analyzed_dfs[e['code']][analyzed_dfs[e['code']]['날짜'] == target_date].iloc[0]['보조_40MA'] == '상승'
                for e in short_etfs if e['code'] in analyzed_dfs
            )

            # 대표 정방향 ETF(KODEX 200 / KODEX 코스닥150)의 EMA40 기울기 및 이격도 추출
            rep_df = analyzed_dfs.get(long_etfs[0]['code'])
            rep_row = rep_df[rep_df['날짜'] == target_date].iloc[0] if (rep_df is not None and not rep_df[rep_df['날짜'] == target_date].empty) else None
            rep_slope = float(rep_row['MA40_Slope5']) if (rep_row is not None and pd.notna(rep_row['MA40_Slope5'])) else 0.0
            rep_disp = float(rep_row['MA40_Disparity']) if (rep_row is not None and pd.notna(rep_row['MA40_Disparity'])) else 100.0

            if today_cross:
                direction = today_cross[0]['direction']
                if direction == '🔥 시장 상승 통일':
                    status = "🔴 [최강 상승장] 4대 ETF 동시 상승 교차 포착"
                else:
                    status = "🔵 [최강 하락장] 4대 ETF 동시 하락 교차 포착"
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
                df = analyzed_dfs.get(e['code'])
                if df is not None:
                    row = df[df['날짜'] == target_date]
                    if not row.empty:
                        r0 = row.iloc[0]
                        c_val = int(r0['종가'])
                        ma_status = r0['보조_40MA']
                        disp_val = r0['MA40_Disparity'] if pd.notna(r0['MA40_Disparity']) else 100.0
                        alert_val = r0['EMA40_조기경고']
                        details.append(f"{e['name']}: {c_val:,}원 ({'40MA위' if ma_status == '상승' else '40MA아래'}, 이격 {disp_val:.1f}%)")
                        if alert_val != "정상":
                            early_warnings.append(f"{e['name']} {alert_val} ({c_val:,}원)")

            mkt_results[mkt] = {
                'status': status,
                'slope': rep_slope,
                'disparity': rep_disp,
                'details': details,
                'early_warnings': early_warnings
            }

        eval_list.append({
            'date': target_date,
            'date_str': d_str,
            'markets': mkt_results
        })

    return eval_list


def create_notion_etf_report_page(start_3y, latest_date, daily_market_eval, turnaround_signals, cross_signals, dual_signals, df_records, target_etfs):
    """ETF 트렌드 시그널 분석 리포트를 노션 데이터베이스에 새 페이지로 등록 (속성 정합성 및 줄바꿈 완벽 교정)"""
    if not NOTION_API_KEY or NOTION_API_KEY.startswith("your_") or not NOTION_DATABASE_ID:
        print("ℹ️ 노션 API 키 또는 데이터베이스 ID가 없어 등록을 건너뜁니다.")
        return None

    print("📝 [노션 데이터베이스] ETF 트렌드 시그널 리포트 페이지 생성 중...")
    title_col_name, existing_props = ensure_notion_database_schema()

    now_dt = datetime.datetime.now()
    now_str = now_dt.strftime("%Y-%m-%d %H:%M:%S")
    latest_date_str = latest_date.strftime("%Y-%m-%d")

    # 당일 실제 발생한 시그널 및 조기경고 집계
    today_signals_df = df_records[df_records['date_str'] == latest_date_str] if not df_records.empty else pd.DataFrame()
    today_sig_count = len(today_signals_df)

    today_eval = daily_market_eval[0] if daily_market_eval else None
    kospi_status_str = today_eval['markets']['KOSPI']['status'] if today_eval else "-"
    kosdaq_status_str = today_eval['markets']['KOSDAQ']['status'] if today_eval else "-"
    today_early_warnings = []
    if today_eval:
        today_early_warnings = today_eval['markets']['KOSPI']['early_warnings'] + today_eval['markets']['KOSDAQ']['early_warnings']

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

    # 1. 최상단 핵심 요약 Callout 블록
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

    # 2. 최근 5영업일 내 발생 시그널 타임라인 (종목명 1행 밀림 방지 포맷)
    blocks.append({
        "object": "block",
        "type": "heading_2",
        "heading_2": {"rich_text": [{"type": "text", "text": {"content": "🚨 최근 5영업일 신규 포착 시그널 (날짜 | 종목명 | 시그널 | 종가 | 40EMA 판정)"}}]}
    })

    recent_5d_cutoff = daily_market_eval[-1]['date'] if daily_market_eval else latest_date
    df_recent_sigs = df_records[df_records['date'] >= recent_5d_cutoff].sort_values(by=['date', 'code'], ascending=[False, True]) if not df_records.empty else pd.DataFrame()

    if not df_recent_sigs.empty:
        for _, r in df_recent_sigs.iterrows():
            line_txt = f"• {r['date_str']} | {r['name']} | {r['signal']} | 종가 {r['close']:,}원 | {r['ema40_filter']} (이격 {r['disparity']:.1f}%, 기울기 {r['slope']:+.2f}%)"
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

    # 3. 최근 5영업일 시장 장세 종합 판단 블록
    blocks.append({
        "object": "block",
        "type": "heading_2",
        "heading_2": {"rich_text": [{"type": "text", "text": {"content": "📌 최근 5영업일 시장 장세 종합 판단"}}]}
    })

    for day in daily_market_eval:
        d_str = day['date_str']
        is_today = (d_str == latest_date_str)
        tag = f"{d_str} (당일)" if is_today else d_str
        for mkt, data in day['markets'].items():
            warn_str = f"\n  - ⚠️ 조기경고: {', '.join(data['early_warnings'])}" if data['early_warnings'] else ""
            txt = f"[{tag}] {mkt}: {data['status']}\n  - 종목별 상태: {', '.join(data['details'])}{warn_str}"
            blocks.append({
                "object": "block",
                "type": "bulleted_list_item",
                "bulleted_list_item": {"rich_text": [{"type": "text", "text": {"content": txt[:1900]}}]}
            })

    # 4. 시장 턴어라운드 (Turnaround) 및 4대 ETF 교차 시그널 블록
    blocks.append({
        "object": "block",
        "type": "heading_2",
        "heading_2": {"rich_text": [{"type": "text", "text": {"content": "🔄 시장 턴어라운드 & 🔥 4대 ETF 동시 교차 포착 시그널 (최근 10건)"}}]}
    })

    if turnaround_signals:
        for t in turnaround_signals[:5]:
            names_str = ", ".join([f"{item['name']}({item['signal']})" for item in t['items']])
            txt = f"[턴어라운드] {t['curr_date']} | {t['market']} | {t['turnaround_type']} (직전: {t['prev_date']} {t['prev_direction']}) → {names_str}"
            blocks.append({
                "object": "block",
                "type": "bulleted_list_item",
                "bulleted_list_item": {"rich_text": [{"type": "text", "text": {"content": txt[:1900]}}]}
            })

    if cross_signals:
        for cs in cross_signals[:10]:
            items_summary = ", ".join([f"{item['name']}({item['signal']} {item['close']:,}원)" for item in cs['items']])
            txt = f"[4대교차] {cs['date_str']} | {cs['market']} | {cs['direction']} → {items_summary}"
            blocks.append({
                "object": "block",
                "type": "bulleted_list_item",
                "bulleted_list_item": {"rich_text": [{"type": "text", "text": {"content": txt[:1900]}}]}
            })

    # 5. 종목별 최근 3년 시그널 현황 (줄바꿈 밀림 방지: 각 줄마다 날짜 | 종목명 | 시그널 | 종가 | 필터 명시)
    blocks.append({
        "object": "block",
        "type": "heading_2",
        "heading_2": {"rich_text": [{"type": "text", "text": {"content": "📊 최근 3년 종목별 (+) 시그널 상세 이력"}}]}
    })

    for item in target_etfs:
        code, name = item['code'], item['name']
        df_sub = df_records[df_records['code'] == code] if not df_records.empty else pd.DataFrame()
        if not df_sub.empty:
            df_sub = df_sub.sort_values(by='date', ascending=False)
            up_cnt = len(df_sub[df_sub['signal'] == '상승트렌드 +'])
            down_cnt = len(df_sub[df_sub['signal'] == '하락트렌드 +'])
            header_line = f"🔹 {name} (총 {len(df_sub)}회 : 상승+ {up_cnt}회 / 하락+ {down_cnt}회)"
            lines = [
                f"  • {r['date_str']} | {r['name']} | {r['signal']} | 종가 {r['close']:,}원 | {r['ema40_filter']}"
                for _, r in df_sub.iterrows()
            ]
            blocks.extend(split_lines_to_notion_paragraphs(header_line, lines))
        else:
            blocks.append({
                "object": "block",
                "type": "paragraph",
                "paragraph": {"rich_text": [{"type": "text", "text": {"content": f"🔹 {name}: 최근 3년 내 시그널 없음"}}]}
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
        else:
            print(f"  └ ❌ 노션 등록 실패: Status {res.status_code}, Body: {res.text}")
            return None
    except Exception as e:
        print(f"  └ ❌ 노션 API 요청 중 예외 발생: {e}")
        return None


def build_compact_telegram_briefing(latest_date, daily_market_eval, turnaround_signals, cross_signals, dual_signals, df_records, notion_url):
    """중요한 메시지를 중심으로 한눈에 들어오도록 요약한 단일 텔레그램 브리핑 생성"""
    latest_str = latest_date.strftime('%Y-%m-%d')
    today_eval = daily_market_eval[0] if daily_market_eval else None

    lines = [
        f"📊 <b>[8대 ETF 마켓 시그널 핵심 요약]</b>",
        f"📅 <b>기준일:</b> {latest_str}",
        "━━━━━━━━━━━━━━━━━━━━━"
    ]

    # 1. 오늘(당일) 신규 발생한 액션 시그널 & EMA 40 조기경고 (가장 중요한 정보 최상단 배치)
    today_sigs = df_records[df_records['date_str'] == latest_str] if not df_records.empty else pd.DataFrame()
    today_cross = [c for c in cross_signals if c['date_str'] == latest_str]
    today_turn = [t for t in turnaround_signals if t['curr_date'] == latest_str]
    today_dual = [d for d in dual_signals if d['date_str'] == latest_str]

    early_warns = []
    if today_eval:
        early_warns = today_eval['markets']['KOSPI']['early_warnings'] + today_eval['markets']['KOSDAQ']['early_warnings']

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
            lines.append(
                f"• <b>{html.escape(r['name'])}</b>: {html.escape(r['signal'])} "
                f"({r['close']:,}원 | {html.escape(r['ema40_filter'])})"
            )
    else:
        lines.append("• 당일 신규 트렌드(+) 발생 종목 없음")

    if early_warns:
        for w in early_warns:
            lines.append(f"• {html.escape(w)} <i>(비중 축소/변곡점 주의)</i>")

    # 2. KOSPI / KOSDAQ 40EMA 장세 요약
    lines.append("\n📌 <b>[2. 시장별 40EMA 장세 판단]</b>")
    if today_eval:
        for mkt in ['KOSPI', 'KOSDAQ']:
            m_data = today_eval['markets'][mkt]
            lines.append(f"• <b>{mkt}</b>: {html.escape(m_data['status'])}")

    # 3. 직전 최근 핵심 시그널 이력 (발생일자 명확히 표기)
    lines.append("\n🧭 <b>[3. 최근 주요 변곡점 이력]</b>")
    if turnaround_signals:
        rt = turnaround_signals[0]
        lines.append(f"• <b>최근 턴어라운드:</b> {rt['curr_date']} ({rt['market']} {html.escape(rt['curr_direction'])})")
    if cross_signals:
        rc = cross_signals[0]
        lines.append(f"• <b>최근 4대교차:</b> {rc['date_str']} ({rc['market']} {html.escape(rc['direction'])})")
    if dual_signals:
        rd = dual_signals[0]
        lines.append(f"• <b>최근 그룹동시:</b> {rd['date_str']} ({html.escape(rd['group'])} {html.escape(rd['signal'])})")

    # 최근 5영업일 내 당일 외 시그널이 있었다면 1줄 요약
    if not df_records.empty and len(daily_market_eval) > 1:
        cutoff_5d = daily_market_eval[-1]['date']
        prev_5d_sigs = df_records[(df_records['date'] >= cutoff_5d) & (df_records['date_str'] != latest_str)]
        if not prev_5d_sigs.empty:
            prev_summary = ", ".join([
                f"{r['date_str'][5:]} {r['name']}({r['signal'].replace('트렌드 ', '')})"
                for _, r in prev_5d_sigs.sort_values(by='date', ascending=False).head(4).iterrows()
            ])
            lines.append(f"• <b>최근 5일내 신호:</b> {html.escape(prev_summary)}")

    if notion_url:
        lines.append(f"\n🔗 <a href='{notion_url}'>노션 상세 리포트 열기</a>")

    return "\n".join(lines)


def send_telegram_message(bot_token, chat_id, text):
    """텔레그램 메시지 전송 유틸리티 (오류 상세 출력 포함)"""
    if not bot_token or not chat_id:
        print("텔레그램 인증정보가 없어 전송을 건너뜁니다.")
        return False
    url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
    payload = {
        'chat_id': chat_id,
        'text': text[:4000],
        'parse_mode': 'HTML',
        'disable_web_page_preview': True
    }
    try:
        res = requests.post(url, data=payload, timeout=10)
        if res.status_code == 200:
            return True
        else:
            print(f"⚠️ 텔레그램 API 응답 오류 ({res.status_code}): {res.text}")
            return False
    except Exception as e:
        print(f"텔레그램 전송 오류: {e}")
        return False