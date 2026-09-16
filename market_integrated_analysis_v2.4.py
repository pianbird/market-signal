import sys
import io
import os
import time
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
MY_CHAT_ID = os.environ.get("MY_CHAT_ID", "7719903334")

# 노션(Notion) API 설정
NOTION_API_KEY = os.environ.get("NOTION_API_KEY", "")
NOTION_DATABASE_ID = os.environ.get("NOTION_DATABASE_ID", "3bfed8456d788084a904f35566400acf")
NOTION_VERSION = "2022-06-28"

NOTION_HEADERS = {
    "Authorization": f"Bearer {NOTION_API_KEY}",
    "Content-Type": "application/json",
    "Notion-Version": NOTION_VERSION
}

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
}

def ensure_notion_database_schema():
    """노션 데이터베이스 스키마 검사 및 제목 컬럼명 추출"""
    if not NOTION_API_KEY or not NOTION_DATABASE_ID:
        return "제목"

    db_url = f"https://api.notion.com/v1/databases/{NOTION_DATABASE_ID.strip()}"
    try:
        res = requests.get(db_url, headers=NOTION_HEADERS, timeout=10)
        title_col_name = "제목"
        if res.status_code == 200:
            props = res.json().get("properties", {})
            for prop_name, prop_data in props.items():
                if prop_data.get("type") == "title":
                    title_col_name = prop_name
                    break
            
            target_props = {
                "날짜": {"date": {}},
                "시장구분": {"select": {}},
                "포착종목수": {"number": {"format": "number"}},
                "등록시간": {"rich_text": {}}
            }
            missing_props = {p: schema for p, schema in target_props.items() if p not in props}
            if missing_props:
                requests.patch(db_url, json={"properties": missing_props}, headers=NOTION_HEADERS, timeout=10)
        return title_col_name
    except Exception as e:
        print(f"노션 스키마 검사 오류: {e}")
        return "제목"

def append_notion_blocks(page_id, blocks):
    """노션 페이지 하위 블록 배치 추가"""
    if not blocks:
        return
    url = f"https://api.notion.com/v1/blocks/{page_id}/children"
    batch_size = 90
    for i in range(0, len(blocks), batch_size):
        chunk = blocks[i:i+batch_size]
        payload = {"children": chunk}
        try:
            requests.patch(url, json=payload, headers=NOTION_HEADERS, timeout=15)
        except Exception:
            pass

def evaluate_daily_market(latest_date, cross_signals, analyzed_dfs, target_etfs, days=5):
    """최근 n거래일(기본 5영업일) 기준 KOSPI 및 KOSDAQ 장세를 각각 종합 분석하여 리스트 형태로 반환"""
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

            # 1) 당일 4대 ETF 교차 포착 확인
            today_cross = [c for c in cross_signals if c['market'] == mkt and c['date_str'] == d_str]
            if today_cross:
                direction = today_cross[0]['direction']
                if direction == '🔥 시장 상승 통일':
                    status = "🔴 [최강 상승장] 4대 ETF 동시 상승 교차 포착"
                else:
                    status = "🔵 [최강 하락장] 4대 ETF 동시 하락 교차 포착"
            else:
                # 2) 40EMA 위치 관계 종합 판단
                long_above = True
                for e in long_etfs:
                    df = analyzed_dfs.get(e['code'])
                    if df is not None:
                        row = df[df['날짜'] == target_date]
                        if row.empty or row.iloc[0]['보조_40MA'] != '상승':
                            long_above = False
                            break

                short_below = True
                for e in short_etfs:
                    df = analyzed_dfs.get(e['code'])
                    if df is not None:
                        row = df[df['날짜'] == target_date]
                        if row.empty or row.iloc[0]['보조_40MA'] != '하락':
                            short_below = False
                            break

                long_below = True
                for e in long_etfs:
                    df = analyzed_dfs.get(e['code'])
                    if df is not None:
                        row = df[df['날짜'] == target_date]
                        if row.empty or row.iloc[0]['보조_40MA'] != '하락':
                            long_below = False
                            break

                short_above = True
                for e in short_etfs:
                    df = analyzed_dfs.get(e['code'])
                    if df is not None:
                        row = df[df['날짜'] == target_date]
                        if row.empty or row.iloc[0]['보조_40MA'] != '상승':
                            short_above = False
                            break

                if long_above and short_below:
                    status = "🔴 [상승 우세] 40EMA 상회 정방향 우세장"
                elif long_below and short_above:
                    status = "🔵 [하락 우세] 40EMA 하회 인버스 우세장"
                elif long_below and short_below:
                    status = "🟡 [혼조/관망장] 지수-인버스 40EMA 혼조세"
                else:
                    status = "⚪ [중립/보합장] 트렌드 탐색 구간"

            # 상세 종목 종가 요약
            details = []
            for e in mkt_etfs:
                df = analyzed_dfs.get(e['code'])
                if df is not None:
                    row = df[df['날짜'] == target_date]
                    if not row.empty:
                        c_val = int(row.iloc[0]['종가'])
                        ma_status = row.iloc[0]['보조_40MA']
                        details.append(f"{e['name']}: {c_val:,}원 ({'40MA 위' if ma_status == '상승' else '40MA 아래'})")

            mkt_results[mkt] = {
                'status': status,
                'details': details
            }

        eval_list.append({
            'date': target_date,
            'date_str': d_str,
            'markets': mkt_results
        })

    return eval_list

def create_notion_etf_report_page(start_3y, latest_date, daily_market_eval, turnaround_signals, cross_signals, dual_signals, df_records, target_etfs):
    """ETF 트렌드 시그널 분석 리포트를 노션 데이터베이스에 새 페이지로 등록"""
    if not NOTION_API_KEY or NOTION_API_KEY.startswith("your_"):
        print("ℹ️ 노션 API 키가 설정되지 않아 노션 데이터베이스 등록을 건너뜁니다.")
        return False

    print("📝 [노션 데이터베이스] ETF 트렌드 시그널 리포트 페이지 생성 중...")
    title_col_name = ensure_notion_database_schema()

    now_dt = datetime.datetime.now()
    now_str = now_dt.strftime("%Y-%m-%d %H:%M:%S")
    now_iso = now_dt.strftime("%Y-%m-%dT%H:%M:%S+09:00")

    properties = {
        title_col_name: {
            "title": [{"text": {"content": f"8대 ETF 최근 3년 트렌드(+) 시그널 분석 보고서 ({now_str})"}}]
        },
        "날짜": {"date": {"start": now_iso}},
        "시장구분": {"select": {"name": "ETF 트렌드 시그널"}},
        "포착종목수": {"number": len(target_etfs)},
        "등록시간": {"rich_text": [{"text": {"content": now_str}}]}
    }

    blocks = []
    
    # 1. 요약 Callout 블록
    summary_txt = f"🤖 [8대 ETF 최근 3년 트렌드(+) 시그널 분석 보고서]\n📅 분석 기간: {start_3y.strftime('%Y-%m-%d')} ~ {latest_date.strftime('%Y-%m-%d')}\n등록시간: {now_str}"
    blocks.append({
        "object": "block",
        "type": "callout",
        "callout": {
            "rich_text": [{"type": "text", "text": {"content": summary_txt}}],
            "icon": {"emoji": "📊"}
        }
    })

    # 2. 최근 5영업일 시장 장세 종합 판단 블록
    blocks.append({
        "object": "block",
        "type": "heading_2",
        "heading_2": {"rich_text": [{"type": "text", "text": {"content": "📌 최근 5영업일 시장 장세 종합 판단"}}]}
    })

    for day in daily_market_eval:
        d_str = day['date_str']
        is_today = (d_str == latest_date.strftime('%Y-%m-%d'))
        tag = f"{d_str} (당일)" if is_today else d_str
        for mkt, data in day['markets'].items():
            txt = f"• [{tag}] {mkt} 시장: {data['status']}\n  - 종목별 상태: " + ", ".join(data['details']) + "\n"
            blocks.append({
                "object": "block",
                "type": "bulleted_list_item",
                "bulleted_list_item": {"rich_text": [{"type": "text", "text": {"content": txt[:1900]}}]}
            })

    # 3. 최상단: 턴어라운드 (Turnaround) 최신 포착 시그널 블록
    blocks.append({
        "object": "block",
        "type": "heading_2",
        "heading_2": {"rich_text": [{"type": "text", "text": {"content": "🔄 시장 턴어라운드 (Turnaround) 최신 포착 신호"}}]}
    })

    if turnaround_signals:
        for t in turnaround_signals:
            names_str = ", ".join([f"{item['name']} ({item['signal']})" for item in t['items']])
            txt = f"• [{t['market']}] {t['turnaround_type']}\n  - 반전일자: {t['curr_date']} (직전: {t['prev_date']} [{t['prev_direction']}])\n  - 포착 종목: {names_str}\n"
            blocks.append({
                "object": "block",
                "type": "bulleted_list_item",
                "bulleted_list_item": {"rich_text": [{"type": "text", "text": {"content": txt[:1900]}}]}
            })
    else:
        blocks.append({
            "object": "block",
            "type": "bulleted_list_item",
            "bulleted_list_item": {"rich_text": [{"type": "text", "text": {"content": "• 최근 3년 내 시장 턴어라운드 신호 없음"}}]}
        })

    # 3. KOSPI / KOSDAQ 시장별 4대 ETF 교차 포착 시그널 (가장 강력한 신호)
    blocks.append({
        "object": "block",
        "type": "heading_2",
        "heading_2": {"rich_text": [{"type": "text", "text": {"content": "🔥 KOSPI / KOSDAQ 4대 ETF 교차 포착 시그널 (가장 강력)"}}]}
    })

    if cross_signals:
        for cs in cross_signals:
            names_str = ", ".join([f"{item['name']} ({item['signal']})" for item in cs['items']])
            txt = f"• {cs['market']} | 동시 발생일: {cs['date_str']} [{cs['direction']}]\n  - 포착 종목: {names_str}\n"
            for item in cs['items']:
                txt += f"    └ {item['name']}: {item['signal']} | 종가 {item['close']:,}원 | 거래량 {item['vol']:,}\n"
            blocks.append({
                "object": "block",
                "type": "bulleted_list_item",
                "bulleted_list_item": {"rich_text": [{"type": "text", "text": {"content": txt[:1900]}}]}
            })
    else:
        blocks.append({
            "object": "block",
            "type": "bulleted_list_item",
            "bulleted_list_item": {"rich_text": [{"type": "text", "text": {"content": "• 최근 3년 내 시장 4대 ETF 교차 포착 시그널 없음"}}]}
        })

    # 3. 그룹별 동시 포착 시그널 블록
    blocks.append({
        "object": "block",
        "type": "heading_2",
        "heading_2": {"rich_text": [{"type": "text", "text": {"content": "⚡ 그룹별 최근 동시 포착 시그널 (Dual Signal)"}}]}
    })

    group_list = [
        ('KOSPI 200', 'KODEX 200 & TIGER 200'),
        ('KOSPI 200 인버스', 'KODEX 인버스 & TIGER 인버스'),
        ('KOSDAQ 150', 'KODEX 코스닥150 & TIGER 코스닥150'),
        ('KOSDAQ 150 인버스', 'KODEX 코스닥150선물인버스 & TIGER 코스닥150선물인버스')
    ]

    for g_key, g_label in group_list:
        group_duals = [d for d in dual_signals if d['group'] == g_key]
        if group_duals:
            last_d = group_duals[0]
            txt = f"• {g_key} ({g_label})\n  - 최근 동시 시그널일: {last_d['date_str']} ({last_d['signal']})\n"
            for item in last_d['items']:
                txt += f"    └ {item['name']}: 종가 {item['close']:,}원 | 거래량 {item['vol']:,}\n"
        else:
            txt = f"• {g_key} ({g_label}): 최근 3년 내 동시 포착 시그널 없음"
        
        blocks.append({
            "object": "block",
            "type": "bulleted_list_item",
            "bulleted_list_item": {"rich_text": [{"type": "text", "text": {"content": txt[:1900]}}]}
        })

    # 4. 종목별 3년 시그널 현황
    blocks.append({
        "object": "block",
        "type": "heading_2",
        "heading_2": {"rich_text": [{"type": "text", "text": {"content": "📊 최근 3년 종목별 (+) 시그널 현황"}}]}
    })

    for item in target_etfs:
        code, name = item['code'], item['name']
        df_sub = df_records[df_records['code'] == code] if not df_records.empty else pd.DataFrame()
        if not df_sub.empty:
            df_sub = df_sub.sort_values(by='date', ascending=False)
            up_cnt = len(df_sub[df_sub['signal'] == '상승트렌드 +'])
            down_cnt = len(df_sub[df_sub['signal'] == '하락트렌드 +'])
            txt = f"🔹 {name} (총 {len(df_sub)}회 : 상승+ {up_cnt}회 / 하락+ {down_cnt}회)\n"
            for _, r in df_sub.iterrows():
                txt += f"  • {r['date_str']} | {r['signal']} | 종가 {r['close']:,}원\n"
        else:
            txt = f"🔹 {name}: 시그널 없음"
            
        blocks.append({
            "object": "block",
            "type": "paragraph",
            "paragraph": {"rich_text": [{"type": "text", "text": {"content": txt[:1900]}}]}
        })

    url = "https://api.notion.com/v1/pages"
    first_chunk = blocks[:90]
    remaining_blocks = blocks[90:]

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

def fetch_historical_ohlcv(code, count=1200):
    """네이버 차트 XML API를 이용하여 일별 시세 수집 및 40EMA 연산"""
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
            return df
    return None

def calculate_all_signals(df):
    """6개 보조 시그널 및 통합 트렌드(+) 시그널 판정"""
    sig1_40ma, sig2_gap, sig3_gap_close, sig4_high, sig5_low = [], [], [], [], []

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

    market_trends = []
    for idx in range(len(df)):
        if idx < 2 or df.loc[idx, '보조_40MA'] == "데이터 부족":
            market_trends.append("데이터 부족")
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

        if is_downtrend and is_vol_higher:
            market_trends.append("하락트렌드 +")
        elif is_uptrend and is_vol_higher:
            market_trends.append("상승트렌드 +")
        else:
            market_trends.append("해당없음")

    df['시장트렌드_시그널'] = market_trends
    return df

def send_telegram_message(bot_token, chat_id, text):
    """텔레그램 메시지 전송 유틸리티"""
    url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
    payload = {'chat_id': chat_id, 'text': text, 'parse_mode': 'HTML'}
    try:
        res = requests.post(url, data=payload, timeout=10)
        return res.status_code == 200
    except Exception as e:
        print(f"텔레그램 전송 오류: {e}")
        return False

def run_analysis():
    print("=" * 60)
    print("8대 ETF 최근 3년 상승트랜드(+), 하락트랜드(+) 시그널 분석 시작")
    print("=" * 60)

    # 1. 8개 ETF 시세 데이터 수집 및 시그널 계산
    analyzed_dfs = {}
    for item in TARGET_ETFS:
        code, name, group = item['code'], item['name'], item['group']
        print(f"[{name} ({code})] 시세 수집 및 시그널 계산 중...")
        df = fetch_historical_ohlcv(code, count=1200)
        if df is not None:
            df_sig = calculate_all_signals(df)
            analyzed_dfs[code] = df_sig
        else:
            print(f"  - [{name}] 데이터 수집 실패!")

    if not analyzed_dfs:
        print("분석할 수 있는 데이터가 없습니다.")
        return

    # 2. 최근 3년 기간 설정 (가장 최근 일자 기준 1년)
    ref_df = list(analyzed_dfs.values())[0]
    latest_date = ref_df['날짜'].max()
    start_3y = latest_date - pd.DateOffset(years=3)

    print(f"\n📅 최근 3년 분석 기간: {start_3y.strftime('%Y-%m-%d')} ~ {latest_date.strftime('%Y-%m-%d')}")

    # 3. 최근 3년 내 (+) 시그널 레코드 수집
    records = []
    for item in TARGET_ETFS:
        code, name, group = item['code'], item['name'], item['group']
        df = analyzed_dfs[code]
        df_3y = df[(df['날짜'] >= start_3y) & (df['날짜'] <= latest_date)].copy()
        
        # (+) 시그널만 필터링
        df_plus = df_3y[df_3y['시장트렌드_시그널'].isin(['상승트렌드 +', '하락트렌드 +'])]
        for _, r in df_plus.iterrows():
            records.append({
                'code': code,
                'name': name,
                'group': group,
                'market': item.get('market', ''),
                'type': item.get('type', ''),
                'date': r['날짜'],
                'date_str': r['날짜'].strftime('%Y-%m-%d'),
                'signal': r['시장트렌드_시그널'],
                'close': int(r['종가']),
                'vol': int(r['거래량']),
                'vol_ma': int(r['Vol_MA40']) if pd.notna(r['Vol_MA40']) else 0
            })

    df_records = pd.DataFrame(records)

    # 4. 동시 시그널 검출 (가장 강력한 신호: KOSPI/KOSDAQ 시장 4대 ETF 정방향-인버스 교차 포착 + 턴어라운드 반전 + 그룹별 Dual Signal)
    cross_signals = []
    dual_signals = []
    if not df_records.empty:
        # (1) KOSPI / KOSDAQ 시장별 4대 ETF 정방향-인버스 교차 포착 시그널 (가장 강력한 신호)
        for (market, d_str), mkt_date_df in df_records.groupby(['market', 'date_str']):
            long_df = mkt_date_df[mkt_date_df['type'] == '정방향']
            short_df = mkt_date_df[mkt_date_df['type'] == '인버스']

            # 정방향 2종 및 인버스 2종 모두 (+) 시그널 발생한 경우
            if len(long_df) >= 2 and len(short_df) >= 2:
                long_sigs = set(long_df['signal'])
                short_sigs = set(short_df['signal'])

                is_long_up_short_down = ('상승트렌드 +' in long_sigs) and ('하락트렌드 +' in short_sigs)
                is_long_down_short_up = ('하락트렌드 +' in long_sigs) and ('상승트렌드 +' in short_sigs)

                if is_long_up_short_down or is_long_down_short_up:
                    direction_label = "🔥 시장 상승 통일" if is_long_up_short_down else "❄️ 시장 하락 통일"
                    cross_signals.append({
                        'date_str': d_str,
                        'market': market,
                        'direction': direction_label,
                        'count': len(mkt_date_df),
                        'items': mkt_date_df.to_dict('records')
                    })

        # (2) 그룹별 2개 ETF 동시 포착 시그널 (Dual Signal)
        for (d_str, group, sig), group_df in df_records.groupby(['date_str', 'group', 'signal']):
            if len(group_df) >= 2:  # 동일 그룹(KODEX/TIGER) 동시 발생
                dual_signals.append({
                    'date_str': d_str,
                    'group': group,
                    'signal': sig,
                    'items': group_df.to_dict('records')
                })

    cross_signals.sort(key=lambda x: x['date_str'])
    dual_signals.sort(key=lambda x: x['date_str'])

    # (3) 턴어라운드 (Turnaround) 시그널 검출 (직전 4대 ETF 동시 시그널 대비 반대 방향 반전)
    turnaround_signals = []
    for mkt in ['KOSPI', 'KOSDAQ']:
        mkt_crosses = [c for c in cross_signals if c['market'] == mkt]
        for i in range(1, len(mkt_crosses)):
            prev = mkt_crosses[i - 1]
            curr = mkt_crosses[i]
            if prev['direction'] != curr['direction']:
                t_label = "🔄 상승 턴어라운드 (하락→상승 반전)" if curr['direction'] == "🔥 시장 상승 통일" else "🔄 하락 턴어라운드 (상승→하락 반전)"
                turnaround_signals.append({
                    'market': mkt,
                    'turnaround_type': t_label,
                    'prev_date': prev['date_str'],
                    'prev_direction': prev['direction'],
                    'curr_date': curr['date_str'],
                    'curr_direction': curr['direction'],
                    'items': curr['items']
                })

    turnaround_signals.sort(key=lambda x: x['curr_date'], reverse=True)
    cross_signals.reverse()
    dual_signals.reverse()

    # 당일 시장 장세 종합 판단
    daily_market_eval = evaluate_daily_market(latest_date, cross_signals, analyzed_dfs, TARGET_ETFS)

    # 5. 엑셀 보고서 생성
    excel_path = os.path.join(os.path.dirname(__file__), "ETF_8종_최근3년_트렌드시그널_분석.xlsx")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    font_title = Font(name='맑은 고딕', size=15, bold=True, color='1F497D')
    font_sub = Font(name='맑은 고딕', size=10, italic=True, color='595959')
    font_header = Font(name='맑은 고딕', size=11, bold=True, color='FFFFFF')
    font_bold = Font(name='맑은 고딕', size=10, bold=True)
    font_regular = Font(name='맑은 고딕', size=10)

    font_up_plus = Font(name='맑은 고딕', size=10, bold=True, color='C00000')
    font_down_plus = Font(name='맑은 고딕', size=10, bold=True, color='002060')

    fill_header_navy = PatternFill(start_color='1F497D', end_color='1F497D', fill_type='solid')
    fill_header_slate = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
    fill_up_plus = PatternFill(start_color='FCE4D6', end_color='FCE4D6', fill_type='solid')
    fill_down_plus = PatternFill(start_color='DDEBF7', end_color='DDEBF7', fill_type='solid')
    fill_dual = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')

    border_thin = Side(style='thin', color='D9D9D9')
    border_box = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)

    align_center = Alignment(horizontal='center', vertical='center')
    align_right = Alignment(horizontal='right', vertical='center')
    align_left = Alignment(horizontal='left', vertical='center')

    # [시트 1: 요약]
    ws_sum = wb.create_sheet(title="최근3년 요약")
    ws_sum.views.sheetView[0].showGridLines = True
    ws_sum['A1'] = "8대 ETF 최근 3년 트렌드(+) 시그널 분석 요약"
    ws_sum['A1'].font = font_title
    ws_sum['A2'] = f"분석 기간: {start_3y.strftime('%Y-%m-%d')} ~ {latest_date.strftime('%Y-%m-%d')}"
    ws_sum['A2'].font = font_sub

    # 최근 5영업일 시장 장세 종합 판단 배너 (A3)
    ws_sum['A3'] = "📌 [최근 5영업일 시장 장세 종합 판단]"
    ws_sum['A3'].font = Font(name='맑은 고딕', size=11, bold=True, color='1F497D')

    curr_banner_row = 4
    for day in daily_market_eval:
        d_str = day['date_str']
        is_today = (d_str == latest_date.strftime('%Y-%m-%d'))
        label = f"{d_str} (당일)" if is_today else d_str
        k_st = day['markets']['KOSPI']['status']
        q_st = day['markets']['KOSDAQ']['status']
        
        mkt_summary_txt = f"  • {label} | KOSPI: {k_st} | KOSDAQ: {q_st}"
        cell = ws_sum.cell(row=curr_banner_row, column=1, value=mkt_summary_txt)
        cell.font = Font(name='맑은 고딕', size=10, bold=is_today)
        bg_color = 'DDEBF7' if is_today else 'F2F2F2'
        cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type='solid')
        curr_banner_row += 1

    # 최상단 턴어라운드 배너
    if turnaround_signals:
        last_t = turnaround_signals[0]
        t_info = f"🔄 [최근 턴어라운드 발생] {last_t['market']} | {last_t['turnaround_type']} (반전일: {last_t['curr_date']} / 직전: {last_t['prev_date']} [{last_t['prev_direction']}])"
        cell = ws_sum.cell(row=curr_banner_row, column=1, value=t_info)
        cell.font = Font(name='맑은 고딕', size=11, bold=True, color='9C6500')
        cell.fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
        curr_banner_row += 1

    header_row = curr_banner_row + 1

    sum_headers = ["종목명", "그룹", "상승트렌드(+) 횟수", "하락트렌드(+) 횟수", "총 (+)시그널 횟수", "시장 4대 ETF 교차포착 참여", "그룹 동시포착 참여"]
    for c_idx, h in enumerate(sum_headers, 1):
        cell = ws_sum.cell(row=header_row, column=c_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header_navy
        cell.alignment = align_center

    for r_idx, item in enumerate(TARGET_ETFS, header_row + 1):
        code, name, group = item['code'], item['name'], item['group']
        df_sub = df_records[df_records['code'] == code] if not df_records.empty else pd.DataFrame()
        
        up_cnt = len(df_sub[df_sub['signal'] == '상승트렌드 +']) if not df_sub.empty else 0
        down_cnt = len(df_sub[df_sub['signal'] == '하락트렌드 +']) if not df_sub.empty else 0
        
        # 4대 ETF 교차 포착 참여 횟수 및 그룹 동시 포착 참여 횟수
        cross_cnt = sum(1 for c in cross_signals if any(x['code'] == code for x in c['items']))
        dual_cnt = sum(1 for d in dual_signals if any(x['code'] == code for x in d['items']))

        vals = [name, group, up_cnt, down_cnt, up_cnt + down_cnt, cross_cnt, dual_cnt]
        for c_idx, val in enumerate(vals, 1):
            cell = ws_sum.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_regular
            cell.border = border_box
            if c_idx in [1, 2]:
                cell.alignment = align_center
            else:
                cell.alignment = align_right
                cell.number_format = '#,##0'

    # [시트 2: 가장 강력한 동시 시그널 모음]
    ws_dual = wb.create_sheet(title="가장 강력한 동시 시그널")
    ws_dual.views.sheetView[0].showGridLines = True
    ws_dual['A1'] = "🔄 [최신 턴어라운드] & 🔥 KOSPI / KOSDAQ 4대 ETF 교차 동시 발생일 목록 (최근 3년)"
    ws_dual['A1'].font = font_title

    dual_headers = ["시그널 구분", "발생일자", "시장/방향", "동시 포착 종목", "종가", "당일 거래량", "40EMA 거래량", "비고"]
    for c_idx, h in enumerate(dual_headers, 1):
        cell = ws_dual.cell(row=3, column=c_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header_slate
        cell.alignment = align_center

    curr_row = 4

    # 1) 최상단: 턴어라운드 (Turnaround) 발생일 및 포착 종목 우선 배치
    fill_turnaround = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
    font_turnaround = Font(name='맑은 고딕', size=10, bold=True, color='9C6500')

    for t in turnaround_signals:
        for item in t['items']:
            vals = [
                f"🔄 {t['market']} 턴어라운드", t['curr_date'], t['turnaround_type'], f"{item['name']} ({item['signal']})",
                item['close'], item['vol'], item['vol_ma'], f"🔄 턴어라운드 신호 (직전:{t['prev_date']} {t['prev_direction']})"
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
                    cell.number_format = '#,##0'
            curr_row += 1

    # 2) 시장 4대 ETF 정방향-인버스 교차 발생 (가장 강력한 신호)
    for cs in cross_signals:
        for item in cs['items']:
            vals = [
                f"🔥 {cs['market']} 4대 ETF 교차", cs['date_str'], cs['direction'], f"{item['name']} ({item['signal']})",
                item['close'], item['vol'], item['vol_ma'], "🔥 가장 강력한 신호 (정방향+인버스 교차)"
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
                    cell.number_format = '#,##0'
                if c_idx in [3, 4]:
                    cell.font = font_up_plus if item['signal'] == '상승트렌드 +' else font_down_plus
            curr_row += 1

    # 3) 그룹별 동시 포착 시그널 (Dual Signal)
    for d in dual_signals:
        for item in d['items']:
            vals = [
                f"⚡ {d['group']} 동시", d['date_str'], d['signal'], item['name'],
                item['close'], item['vol'], item['vol_ma'], "⚡ 그룹 동시 시그널"
            ]
            for c_idx, val in enumerate(vals, 1):
                cell = ws_dual.cell(row=curr_row, column=c_idx, value=val)
                cell.font = font_regular
                cell.border = border_box
                if c_idx in [1, 2, 3, 4, 8]:
                    cell.alignment = align_center
                else:
                    cell.alignment = align_right
                    cell.number_format = '#,##0'
                if c_idx == 3:
                    cell.font = font_up_plus if d['signal'] == '상승트렌드 +' else font_down_plus
            curr_row += 1

    # [시트 3~10: ETF별 3년 상세 시트]
    for item in TARGET_ETFS:
        code, name, group = item['code'], item['name'], item['group']
        df = analyzed_dfs[code]
        df_3y = df[(df['날짜'] >= start_3y) & (df['날짜'] <= latest_date)].sort_values(by='날짜', ascending=False).reset_index(drop=True)

        ws = wb.create_sheet(title=f"{name} 3년상세")
        ws.views.sheetView[0].showGridLines = True
        ws['A1'] = f"{name} ({code}) 최근 3년 시그널 분석 상세"
        ws['A1'].font = font_title

        detail_headers = ["날짜", "종가", "시가", "고가", "저가", "거래량", "40EMA 종가", "40EMA 거래량", "트렌드(+) 시그널"]
        for c_idx, h in enumerate(detail_headers, 1):
            cell = ws.cell(row=3, column=c_idx, value=h)
            cell.font = font_header
            cell.fill = fill_header_navy
            cell.alignment = align_center

        for r_idx, r in df_3y.iterrows():
            row_num = r_idx + 4
            sig = r['시장트렌드_시그널']
            vals = [
                r['날짜'].strftime('%Y-%m-%d'), int(r['종가']), int(r['시가']), int(r['고가']), int(r['저가']),
                int(r['거래량']), int(r['MA40']) if pd.notna(r['MA40']) else 0,
                int(r['Vol_MA40']) if pd.notna(r['Vol_MA40']) else 0, sig
            ]

            for c_idx, val in enumerate(vals, 1):
                cell = ws.cell(row=row_num, column=c_idx, value=val)
                cell.border = border_box
                cell.font = font_regular
                if c_idx == 1:
                    cell.alignment = align_center
                elif c_idx in [2, 3, 4, 5, 6, 7, 8]:
                    cell.alignment = align_right
                    cell.number_format = '#,##0'
                else:
                    cell.alignment = align_center

                if sig == '상승트렌드 +':
                    cell.fill = fill_up_plus
                    if c_idx == 9:
                        cell.font = font_up_plus
                elif sig == '하락트렌드 +':
                    cell.fill = fill_down_plus
                    if c_idx == 9:
                        cell.font = font_down_plus

    # 너비 자동 조정
    for sheet in wb.worksheets:
        for col in sheet.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or '')
                cell_len = sum(2 if ord(c) > 128 else 1 for c in val_str)
                if cell_len > max_len:
                    max_len = cell_len
            sheet.column_dimensions[col_letter].width = max(max_len + 4, 12)

    try:
        wb.save(excel_path)
        print(f"\n✅ 엑셀 파일 저장 완료: {excel_path}")
    except PermissionError:
        alt_excel_path = os.path.join(os.path.dirname(__file__), f"ETF_8종_최근3년_트렌드시그널_분석_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        wb.save(alt_excel_path)
        print(f"\n⚠️ 기존 엑셀 파일이 사용 중이어서 새 파일명으로 저장 완료: {alt_excel_path}")
    except Exception as e:
        print(f"\n❌ 엑셀 파일 저장 중 오류 발생: {e}")

    # 6. 노션 데이터베이스 리포트 페이지 등록 (텔레그램에 링크 포함을 위해 먼저 실행)
    notion_url = create_notion_etf_report_page(start_3y, latest_date, daily_market_eval, turnaround_signals, cross_signals, dual_signals, df_records, TARGET_ETFS)

    # 7. 텔레그램 브리핑 메시지 작성 및 전송 (요점만 간단히)
    current_time_str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
    
    # (6-1) 시장 장세 종합 요약 메시지
    msg_market = f"📌 <b>[오늘의 시장 장세 요약]</b> ({current_time_str})\n"
    for day in daily_market_eval:
        d_str = day['date_str']
        if d_str == latest_date.strftime('%Y-%m-%d'):
            msg_market += f"• <b>{d_str} (최근 영업일)</b>\n"
            for mkt, data in day['markets'].items():
                msg_market += f"  - <b>{mkt}</b>: {data['status']}\n"

    # (6-2) 8대 ETF 트렌드 시그널 요약 메시지
    msg_etf = f"🤖 <b>[ETF 트렌드 포착 요약]</b> ({current_time_str})\n"
    msg_etf += f"📅 <b>기준일</b>: {latest_date.strftime('%Y-%m-%d')}\n"
    msg_etf += "━━━━━━━━━━━━━━━━━━━━━\n"
    
    msg_etf += "🔄 <b>[시장 턴어라운드 신호]</b>\n"
    if turnaround_signals:
        recent_turn = turnaround_signals[0] if turnaround_signals else None
        msg_etf += f"• <b>{len(turnaround_signals)}건</b> 포착 (최근: {recent_turn['market']} {recent_turn['turnaround_type']})\n"
    else:
        msg_etf += "• 신호 없음\n"

    msg_etf += "\n🔥 <b>[KOSPI/KOSDAQ 교차 신호]</b>\n"
    if cross_signals:
        recent_cross = cross_signals[0] if cross_signals else None
        msg_etf += f"• <b>{len(cross_signals)}건</b> 포착 (최근: {recent_cross['market']} {recent_cross['direction']})\n"
    else:
        msg_etf += "• 신호 없음\n"

    msg_etf += "\n⚡ <b>[그룹별 동시 신호 (Dual Signal)]</b>\n"
    if dual_signals:
        recent_dual = dual_signals[0] if dual_signals else None
        msg_etf += f"• 총 <b>{len(dual_signals)}건</b> 포착 (최근: {recent_dual['group']} {recent_dual['signal']})\n"
    else:
        msg_etf += "• 신호 없음\n"

    if notion_url:
        msg_etf += f"\n💡 <i>※ 상세 내역(3년 분석, 종목별 시그널)은 <a href='{notion_url}'>노션 리포트</a> 및 엑셀을 확인하세요.</i>"
    else:
        msg_etf += "\n💡 <i>※ 상세 내역(3년 분석, 종목별 시그널)은 노션 리포트 및 엑셀을 확인하세요.</i>"

    print("\n[텔레그램 브리핑 메시지 1: 최근 5영업일 시장 장세 종합 판단]")
    print(msg_market)
    print("\n[텔레그램 브리핑 메시지 2: 8대 ETF 최근 3년 트렌드(+) 시그널 분석 보고서]")
    print(msg_etf)

    # 텔레그램 전송 (2개 메시지 분리 전송)
    success1 = send_telegram_message(MY_BOT_TOKEN, MY_CHAT_ID, msg_market)
    time.sleep(1)
    success2 = send_telegram_message(MY_BOT_TOKEN, MY_CHAT_ID, msg_etf)

    if success1 and success2:
        print("\n✅ 텔레그램 메시지 2건이 성공적으로 전송되었습니다.")
    else:
        print(f"\n⚠️ 텔레그램 전송 결과 - 메시지1(장세판단): {'성공' if success1 else '실패'}, 메시지2(ETF분석): {'성공' if success2 else '실패'}")

if __name__ == '__main__':
    run_analysis()

# ==============================================================================

# ==============================================================================
# 📜 개발 히스토리 (Development History)
# ==============================================================================
#
# [v1.5 - v1.6] 초작 및 기본 시장 통합 수집 모듈 개발
#   - 네이버 증권 기반 주요 시장 지수(KOSPI, KOSDAQ, 환율, 원자재 등) 데이터 수집.
#   - Notion API 연동 (일간/주간 시장 분석 보고서 페이지 및 블록 구조 생성).
#   - Telegram Bot API 연동 (수집/분석 결과 텔레그램 요약 브리핑 전송).
#
# [v1.7 - v1.9] 시장 분석 및 포맷 확장
#   - 네이버 증권 업종/테마 상세 데이터, 거래대금 상위, 투자자 동향(개인/외국인/기관) 수집 추가.
#   - 상승/하락 종목 수 및 시장 매수/매도 Breadth 시각화/분석 기능 확장.
#   - 노션 페이지 포맷 다중 블록(상세 종목 리스트, 주요 지수 차트 구조화)으로 대폭 확장.
#
# [v1.91 - v1.92] API 안정화 및 오류 처리 강화
#   - 노션 블록 생성 시 API 전송 제한(100개 chunk 단위 분할) 및 타임아웃 예외 처리 강화.
#   - 데이터 수집 시 실패/재시도(Retry) 및 네트워크 예외 처리 모듈화.
#
# [v2.0] ETF 트렌드 시그널 전용 모듈로 모듈화 및 대폭 경량화 (~2,700줄 → ~640줄)
#   - 분석 대상을 대표 8대 ETF(KODEX 200, TIGER 200, KODEX 인버스, TIGER 인버스,
#     KODEX 코스닥150, TIGER 코스닥150, KODEX 코스닥150선물인버스, TIGER 코스닥150선물인버스)로 집중.
#   - 50일 이동평균선 거래량 지표 기반 상승/하락 (+) 시그널 및
#     동일 그룹 내 2개 ETF 동시 포착 시그널(Dual Signal) 분석 기능 도입.
#   - openpyxl 기반 Excel 자동 리포트 생성 (자동 컬럼 폭 맞춤, 셀 조건부 색상 적용).
#   - 텔레그램 브리핑 전송 + 노션 ETF 전용 리포트 DB 자동 등록.
#
# [v2.1] 경량 오프라인 테스트 버전
#   - 노션 API를 제외하고 엑셀 보고서 + 텔레그램 메시지 브리핑 중심 오프라인 분석 테스트 (465줄).
#
# [v2.2] 노션 API 자동화 및 예외 대응 강화
#   - 노션 API DB 스키마 자동 감지(ensure_notion_database_schema)로 컬럼명 자동 매핑.
#   - 텔레그램 메시지 길이 제한 방지 (최근 3개 시그널 요약 표시).
#   - 엑셀 파일 열림(PermissionError) 발생 시 타임스탬프 파일명 자동 대체 저장 안전장치 추가.
#
# 
# ==============================================================================
