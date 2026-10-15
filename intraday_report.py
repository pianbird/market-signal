import os
import sys
import html
import datetime
import requests
import xml.etree.ElementTree as ET
import pandas as pd

# 콘솔 UTF-8 출력 설정
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
    except Exception:
        pass

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

MY_BOT_TOKEN = os.environ.get("MY_BOT_TOKEN", "")
MY_CHAT_ID = os.environ.get("MY_CHAT_ID", "")

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
}

TARGET_ETFS = [
    {"code": "069500", "name": "KODEX 200", "market": "KOSPI", "type": "정방향"},
    {"code": "102110", "name": "TIGER 200", "market": "KOSPI", "type": "정방향"},
    {"code": "114800", "name": "KODEX 인버스", "market": "KOSPI", "type": "인버스"},
    {"code": "123310", "name": "TIGER 인버스", "market": "KOSPI", "type": "인버스"},
    {"code": "229200", "name": "KODEX 코스닥150", "market": "KOSDAQ", "type": "정방향"},
    {"code": "232080", "name": "TIGER 코스닥150", "market": "KOSDAQ", "type": "정방향"},
    {"code": "251340", "name": "KODEX 코스닥인버스", "market": "KOSDAQ", "type": "인버스"},
    {"code": "250780", "name": "TIGER 코스닥인버스", "market": "KOSDAQ", "type": "인버스"},
]


def get_intraday_volume_multiplier(now_dt: datetime.datetime) -> float:
    """
    장중(09:00~15:30, 총 390분) 경과 시간에 따른 거래량 보정 계수 산출.
    예: 10:00(60분 경과) -> 390 / 60 = 6.5배 보정 (단, 초기 왜곡 방지를 위해 최대 5.0배 상한)
    """
    market_open = now_dt.replace(hour=9, minute=0, second=0, microsecond=0)
    market_close = now_dt.replace(hour=15, minute=30, second=0, microsecond=0)

    if now_dt <= market_open or now_dt >= market_close:
        return 1.0

    elapsed_minutes = max(15.0, (now_dt - market_open).total_seconds() / 60.0)
    multiplier = 390.0 / elapsed_minutes
    return min(multiplier, 5.0)


def fetch_intraday_etf_data(code: str, vol_mult: float, count: int = 300) -> pd.DataFrame:
    """네이버 차트 XML API에서 일봉(당일 실시간 현재가 포함) 수집 및 장중 지표 계산"""
    url = f"https://fchart.stock.naver.com/sise.nhn?symbol={code}&timeframe=day&count={count}&requestType=0"
    res = requests.get(url, headers=HEADERS, timeout=10)
    res.raise_for_status()

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

    df = pd.DataFrame(rows)
    df["dt"] = pd.to_datetime(df["날짜"])
    df = df.sort_values(by="dt", ascending=True).reset_index(drop=True)

    df["MA40"] = df["종가"].ewm(span=40, min_periods=20, adjust=False).mean()
    df["Vol_MA40"] = df["거래량"].ewm(span=40, min_periods=20, adjust=False).mean()
    df["Slope5"] = ((df["MA40"] - df["MA40"].shift(5)) / df["MA40"].shift(5)) * 100.0
    df["Disparity"] = (df["종가"] / df["MA40"]) * 100.0

    # 당일(마지막 행) 장중 시그널 판정 (실제 거래량 기준 vs 시간비례 예상 거래량 기준)
    d0, d1, d2, d3 = df.iloc[-1], df.iloc[-2], df.iloc[-3], df.iloc[-4]
    c_up = (
        (d2["시가"] < d3["종가"] and d2["종가"] > d2["시가"]) and
        (d1["고가"] > d2["고가"] and d1["저가"] > d2["저가"]) and
        (d0["고가"] > d1["고가"] and d0["저가"] > d1["저가"] and d0["종가"] > d0["MA40"])
    )
    c_down = (
        (d2["시가"] > d3["종가"] and d2["종가"] < d2["시가"]) and
        (d1["고가"] < d2["고가"] and d1["저가"] < d2["저가"]) and
        (d0["고가"] < d1["고가"] and d0["저가"] < d1["저가"] and d0["종가"] < d0["MA40"])
    )

    raw_vol_ok = d0["거래량"] > d0["Vol_MA40"]
    est_vol_ok = (d0["거래량"] * vol_mult) > d0["Vol_MA40"]

    sig_str = "해당없음"
    if c_up:
        sig_str = "상승트렌드 +(확정)" if raw_vol_ok else ("상승트렌드 +(장중예상)" if est_vol_ok else "캔들상승(거래량미달)")
    elif c_down:
        sig_str = "하락트렌드 +(확정)" if raw_vol_ok else ("하락트렌드 +(장중예상)" if est_vol_ok else "캔들하락(거래량미달)")

    df.loc[df.index[-1], "장중시그널"] = sig_str
    return df


def build_intraday_report() -> str:
    now_dt = datetime.datetime.now()
    now_str = now_dt.strftime("%Y-%m-%d %H:%M")
    vol_mult = get_intraday_volume_multiplier(now_dt)
    is_market_open = vol_mult > 1.0

    dfs = {}
    for item in TARGET_ETFS:
        dfs[item["code"]] = fetch_intraday_etf_data(item["code"], vol_mult)

    r_k200 = dfs["069500"].iloc[-1]
    r_kinv = dfs["114800"].iloc[-1]
    r_kq150 = dfs["229200"].iloc[-1]
    r_kqinv = dfs["251340"].iloc[-1]

    # 1. 시장 장세 판정
    def eval_mkt(long_row, short_row):
        if long_row["종가"] > long_row["MA40"] and short_row["종가"] < short_row["MA40"]:
            return "🔴 상승 우세 (40EMA 위)" if long_row["Slope5"] >= 0.15 else "🟠 상승 둔화 (기울기 평탄)"
        elif long_row["종가"] < long_row["MA40"] and short_row["종가"] > short_row["MA40"]:
            return "🔵 하락 우세 (40EMA 아래)" if long_row["Slope5"] <= -0.15 else "🟡 하락 둔화 (바닥 탐색)"
        return "⚪ 혼조/변곡점 공방"

    kospi_st = eval_mkt(r_k200, r_kinv)
    kosdaq_st = eval_mkt(r_kq150, r_kqinv)

    lines = [
        f"⏱️ <b>[{'장중 실시간' if is_market_open else '장마감'} 주가 분석 리포트]</b>",
        f"📅 <b>시각:</b> {now_str} " + (f"(거래량 보정 x{vol_mult:.1f}배)" if is_market_open else "(확정 종가)"),
        "━━━━━━━━━━━━━━━━━━━━━",
        "📌 <b>[1. 양 시장 실시간 장세 진단]</b>",
        f"• <b>KOSPI:</b> {kospi_st}",
        f"• <b>KOSDAQ:</b> {kosdaq_st}",
        "",
        "📊 <b>[2. 주요 ETF 현재가 & 40EMA 지표]</b>",
    ]

    # 대표 4종 + 신호 발생 종목 상세 요약
    core_codes = ["069500", "114800", "229200", "251340"]
    for item in TARGET_ETFS:
        code = item["code"]
        df = dfs[code]
        r0, r1, r5, r20 = df.iloc[-1], df.iloc[-2], df.iloc[-6], df.iloc[-21]
        chg_1d = ((r0["종가"] - r1["종가"]) / r1["종가"]) * 100.0
        chg_5d = ((r0["종가"] - r5["종가"]) / r5["종가"]) * 100.0
        chg_20d = ((r0["종가"] - r20["종가"]) / r20["종가"]) * 100.0
        vol_pct = (r0["거래량"] / r0["Vol_MA40"]) * 100.0

        if code in core_codes or r0["장중시그널"] != "해당없음":
            sig_badge = f" | 🚨<b>{r0['장중시그널']}</b>" if r0["장중시그널"] != "해당없음" else ""
            lines.append(
                f"• <b>{item['name']}</b>: <b>{int(r0['종가']):,}원</b> ({chg_1d:+.2f}%)\n"
                f"  └ 5일 {chg_5d:+.1f}% | 20일 {chg_20d:+.1f}% | 이격 <b>{r0['Disparity']:.1f}%</b> | 기울기 {r0['Slope5']:+.2f}% | 거래량 {vol_pct:.0f}%{sig_badge}"
            )

    # 3. 3-트랙 알고리즘 실시간 행동 지침 생성
    lines.append("\n🧭 <b>[3. 자동매매 3-트랙 실시간 대응 지침]</b>")

    # Track 1: KOSPI 정방향
    k_two_day_below = (dfs["069500"].iloc[-2]["종가"] < dfs["069500"].iloc[-2]["MA40"]) and (r_k200["종가"] < r_k200["MA40"])
    if k_two_day_below:
        lines.append(f"• <b>KOSPI 정방향:</b> ⚪ 관망/청산유지 (40EMA {int(r_k200['MA40']):,}원 하회 지속)")
    elif r_k200["Disparity"] >= 120.0:
        lines.append(f"• <b>KOSPI 정방향:</b> 💰 과열 구간(이격 {r_k200['Disparity']:.1f}%) → 보유분 50% 분할익절 권장")
    else:
        lines.append(f"• <b>KOSPI 정방향:</b> 🟢 추세 유지 (40EMA {int(r_k200['MA40']):,}원 지지 확인)")

    # Track 2: KOSDAQ 정방향
    q_two_day_below = (dfs["229200"].iloc[-2]["종가"] < dfs["229200"].iloc[-2]["MA40"]) and (r_kq150["종가"] < r_kq150["MA40"])
    if q_two_day_below:
        lines.append(f"• <b>KOSDAQ 정방향:</b> ⚠️ 40EMA 2일연속 이탈({int(r_kq150['MA40']):,}원) → 전량 손절/진입금지")
    elif r_kq150["Disparity"] >= 109.0:
        lines.append(f"• <b>KOSDAQ 정방향:</b> 💰 단기 과열(이격 {r_kq150['Disparity']:.1f}%) → 전량 차익실현 권장")
    else:
        lines.append(f"• <b>KOSDAQ 정방향:</b> 이격 {r_kq150['Disparity']:.1f}% (초입 매수 기준 ≤104.0%)")

    # Track 3: KOSPI 인버스
    if r_k200["Disparity"] <= 88.5 or r_kinv["Disparity"] >= 111.0:
        lines.append(
            f"• <b>KOSPI 인버스:</b> 🎯 패닉 과매도 도달(K200 이격 {r_k200['Disparity']:.1f}% / 인버스 {r_kinv['Disparity']:.1f}%) "
            f"→ <b>보유분 조기 익절 및 신규 추격매수 금지</b>"
        )
    else:
        lines.append(f"• <b>KOSPI 인버스:</b> 인버스 이격 {r_kinv['Disparity']:.1f}% (4대교차 하락 초입만 진입)")

    return "\n".join(lines)


def send_telegram(text: str) -> None:
    if not MY_BOT_TOKEN or not MY_CHAT_ID:
        print("ℹ️ 텔레그램 설정이 없어 콘솔 출력만 수행합니다.")
        return
    url = f"https://api.telegram.org/bot{MY_BOT_TOKEN}/sendMessage"
    try:
        res = requests.post(url, data={"chat_id": MY_CHAT_ID, "text": text[:4000], "parse_mode": "HTML"}, timeout=10)
        if res.status_code == 200:
            print("✅ 텔레그램 발송 완료!")
        else:
            print(f"⚠️ 텔레그램 발송 실패 ({res.status_code}): {res.text}")
    except Exception as e:
        print(f"⚠️ 텔레그램 통신 오류: {e}")



if __name__ == "__main__":
    report_msg = build_intraday_report()
    print(report_msg)
    send_telegram(report_msg)