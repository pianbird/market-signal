import os
import sys
import time
import json
import html
import argparse
import datetime
import requests
import xml.etree.ElementTree as ET
import pandas as pd
from typing import Dict, Any, List, Optional

# intraday_bot.py 상단에 추가 (안전장치)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE_DIR)
load_dotenv(os.path.join(BASE_DIR, ".env"))

# 콘솔 UTF-8 인코딩 보장
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
        sys.stderr.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
    except Exception:
        pass

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

MY_BOT_TOKEN = os.environ.get("MY_BOT_TOKEN", "")
MY_CHAT_ID = os.environ.get("MY_CHAT_ID", "")

# 평일 장중 자동 발송 시각 설정 (원하는 시각으로 추가/변경 가능)
SCHEDULE_TIMES = ["10:00", "13:00", "15:00"]

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATE_FILE = os.path.join(BASE_DIR, "position_state.json")

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
    정규장(09:00~15:30, 390분) 경과 시간에 따른 거래량 보정 배수 계산
    - 10:00 (60분 경과): 390 ÷ 60 = 6.50 -> 상한 5.00배 적용
    - 13:00 (240분 경과): 390 ÷ 240 = 1.62배 적용
    - 15:00 (360분 경과): 390 ÷ 360 = 1.08배 적용
    """
    if now_dt.weekday() >= 5:  # 주말(토, 일)은 보정 없음
        return 1.0
    market_open = now_dt.replace(hour=9, minute=0, second=0, microsecond=0)
    market_close = now_dt.replace(hour=15, minute=30, second=0, microsecond=0)

    if now_dt <= market_open or now_dt >= market_close:
        return 1.0

    elapsed_minutes = max(15.0, (now_dt - market_open).total_seconds() / 60.0)
    return min(390.0 / elapsed_minutes, 5.0)


def fetch_intraday_etf_data(code: str, vol_mult: float, count: int = 300) -> pd.DataFrame:
    """네이버 차트 XML API를 통해 실시간 일봉 수집 및 40EMA·이격도·장중 시그널 계산"""
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


def build_intraday_report(trigger_label: str = "수동 조회") -> str:
    """현재 주가 분석 및 3-트랙 대응 지침 텔레그램 메시지 생성"""
    now_dt = datetime.datetime.now()
    now_str = now_dt.strftime("%Y-%m-%d %H:%M:%S")
    vol_mult = get_intraday_volume_multiplier(now_dt)
    is_market_open = vol_mult > 1.0

    dfs = {}
    for item in TARGET_ETFS:
        dfs[item["code"]] = fetch_intraday_etf_data(item["code"], vol_mult)

    r_k200 = dfs["069500"].iloc[-1]
    r_kinv = dfs["114800"].iloc[-1]
    r_kq150 = dfs["229200"].iloc[-1]
    r_kqinv = dfs["251340"].iloc[-1]

    def eval_mkt(long_row, short_row):
        if long_row["종가"] > long_row["MA40"] and short_row["종가"] < short_row["MA40"]:
            return "🔴 상승 우세 (40EMA 상회)" if long_row["Slope5"] >= 0.15 else "🟠 상승 둔화 (기울기 평탄)"
        elif long_row["종가"] < long_row["MA40"] and short_row["종가"] > short_row["MA40"]:
            return "🔵 하락 우세 (40EMA 하회)" if long_row["Slope5"] <= -0.15 else "🟡 하락 둔화 (바닥 탐색)"
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
        r0, r1, r5, r20 = df.iloc[-1], df.iloc[-2], df.iloc[-6], df.iloc[-21]
        chg_1d = ((r0["종가"] - r1["종가"]) / r1["종가"]) * 100.0
        chg_5d = ((r0["종가"] - r5["종가"]) / r5["종가"]) * 100.0
        chg_20d = ((r0["종가"] - r20["종가"]) / r20["종가"]) * 100.0
        vol_pct = (r0["거래량"] / r0["Vol_MA40"]) * 100.0

        if code in core_codes or r0["장중시그널"] != "해당없음":
            sig_badge = f"\n  └ 🚨 <b>시그널: {html.escape(r0['장중시그널'])}</b>" if r0["장중시그널"] != "해당없음" else ""
            lines.append(
                f"• <b>{item['name']}</b>: <b>{int(r0['종가']):,}원</b> ({chg_1d:+.2f}%)\n"
                f"  └ 5일 {chg_5d:+.1f}% | 20일 {chg_20d:+.1f}% | 40MA {int(r0['MA40']):,}원\n"
                f"  └ 이격 <b>{r0['Disparity']:.1f}%</b> | 기울기 {r0['Slope5']:+.2f}% | 거래량 {vol_pct:.0f}%{sig_badge}"
            )

    lines.append("\n🧭 <b>[3. 자동매매 3-트랙 실시간 대응 지침]</b>")

    # Track 1: KOSPI 정방향
    k_two_day_below = (dfs["069500"].iloc[-2]["종가"] < dfs["069500"].iloc[-2]["MA40"]) and (r_k200["종가"] < r_k200["MA40"])
    if k_two_day_below:
        lines.append(f"• <b>KOSPI 정방향:</b> ⚪ 현금 관망 (40EMA {int(r_k200['MA40']):,}원 회복 대기)")
    elif r_k200["Disparity"] >= 120.0:
        lines.append(f"• <b>KOSPI 정방향:</b> 💰 과열권(이격 {r_k200['Disparity']:.1f}%) → 50% 분할익절 구간")
    else:
        lines.append(f"• <b>KOSPI 정방향:</b> 🟢 추세 유지 (40EMA {int(r_k200['MA40']):,}원 지지 중)")

    # Track 2: KOSDAQ 정방향
    q_two_day_below = (dfs["229200"].iloc[-2]["종가"] < dfs["229200"].iloc[-2]["MA40"]) and (r_kq150["종가"] < r_kq150["MA40"])
    if q_two_day_below:
        lines.append(f"• <b>KOSDAQ 정방향:</b> ⚠️ 40EMA 2일연속 이탈({int(r_kq150['MA40']):,}원) → 전량 청산/진입금지")
    elif r_kq150["Disparity"] >= 109.0:
        lines.append(f"• <b>KOSDAQ 정방향:</b> 💰 과열 상투권(이격 {r_kq150['Disparity']:.1f}%) → 전량 익절 구간")
    else:
        lines.append(f"• <b>KOSDAQ 정방향:</b> 이격 {r_kq150['Disparity']:.1f}% (초입 매수 기준 ≤104.0%)")

    # Track 3: KOSPI 인버스
    if r_k200["Disparity"] <= 88.5 or r_kinv["Disparity"] >= 111.0:
        lines.append(
            f"• <b>KOSPI 인버스:</b> 🎯 패닉 과매도 도달(K200 이격 {r_k200['Disparity']:.1f}% / 인버스 {r_kinv['Disparity']:.1f}%) "
            f"→ <b>보유분 익절 및 신규 추격 금지</b>"
        )
    else:
        lines.append(f"• <b>KOSPI 인버스:</b> 인버스 이격 {r_kinv['Disparity']:.1f}% (4대교차 하락 초입만 진입)")

    return "\n".join(lines)


def build_position_status_report() -> str:
    """position_state.json 기반 현재 보유 슬롯 실시간 평가 손익 브리핑"""
    if not os.path.exists(STATE_FILE):
        return "ℹ️ 아직 생성된 <code>position_state.json</code> 파일이 없습니다. (전 슬롯 현금 관망 중)"

    with open(STATE_FILE, "r", encoding="utf-8") as f:
        state = json.load(f)

    lines = [
        "💼 <b>[자동매매 3-트랙 실시간 포지션 현황]</b>",
        f"🔄 <b>상태 파일 갱신:</b> {state.get('updated_at', '-')}",
        "━━━━━━━━━━━━━━━━━━━━━",
    ]

    for slot_key, slot in state.get("slots", {}).items():
        if slot.get("status") == "CASH" or slot.get("qty", 0) <= 0:
            lines.append(f"• <b>{slot_key}</b> ({slot['name']}): ⚪ <b>CASH (현금 관망)</b>")
        else:
            df = fetch_intraday_etf_data(slot["code"], 1.0, count=30)
            curr_p = float(df.iloc[-1]["종가"])
            entry_p = float(slot["entry_price"])
            pnl_pct = ((curr_p - entry_p) / entry_p) * 100.0
            pnl_krw = int((curr_p - entry_p) * slot["qty"])
            lines.append(
                f"• <b>{slot_key}</b> ({slot['name']}): 🟢 <b>{slot['status']}</b>\n"
                f"  └ 수량: {slot['qty']:,}주 | 매수가: {int(entry_p):,}원 → 현재가: <b>{int(curr_p):,}원</b>\n"
                f"  └ 평가손익: <b>{pnl_krw:+,}원 ({pnl_pct:+.2f}%)</b> | 보유 {slot.get('hold_days', 0)}일차"
            )

    return "\n".join(lines)


def send_telegram_message(text: str, chat_id: str = MY_CHAT_ID) -> bool:
    if not MY_BOT_TOKEN or not chat_id:
        print("ℹ️ 텔레그램 토큰/채팅방 ID 미설정으로 콘솔 출력만 수행합니다.")
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
        return res.status_code == 200
    except Exception as e:
        print(f"⚠️ 텔레그램 전송 오류: {e}")
        return False


def get_initial_update_offset() -> Optional[int]:
    """봇 시작 시 이전 밀린 메시지를 무시하고 최신 offset부터 수신"""
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


def run_daemon() -> None:
    """
    백그라운드 데몬 모드:
    1) 평일 10:00, 13:00, 15:00 정각 자동 리포트 발송
    2) 텔레그램 채팅방 '/분석', '/포지션', '/상태' 명령어 실시간 감지 및 즉시 답장
    """
    print("=" * 65)
    print(f"🤖 [장중 하이브리드 텔레그램 봇 가동 시작]")
    print(f"   • 평일 자동 발송 시각: {', '.join(SCHEDULE_TIMES)}")
    print("   • 실시간 대화형 명령어: /분석 (또는 '분석'), /포지션, /상태")
    print("=" * 65)

    offset = get_initial_update_offset()
    sent_history: Dict[str, List[str]] = {}  # {'YYYY-MM-DD': ['10:00', ...]}

    send_telegram_message(
        "🟢 <b>[장중 마켓시그널 봇 대기 시작]</b>\n"
        f"• <b>정기 자동 브리핑:</b> 평일 {', '.join(SCHEDULE_TIMES)}\n"
        "• <b>즉시 조회 명령어:</b> 채팅창에 <code>/분석</code> 또는 <code>/포지션</code> 입력"
    )

    while True:
        try:
            now_dt = datetime.datetime.now()
            today_str = now_dt.strftime("%Y-%m-%d")
            hhmm_str = now_dt.strftime("%H:%M")

            # 1. 평일(월~금) 지정 시각(10:00, 13:00, 15:00) 자동 발송 체크
            if now_dt.weekday() < 5 and hhmm_str in SCHEDULE_TIMES:
                sent_today = sent_history.setdefault(today_str, [])
                if hhmm_str not in sent_today:
                    print(f"\n⏰ [{today_str} {hhmm_str}] 정기 장중 분석 리포트 자동 발송 중...")
                    report_txt = build_intraday_report(trigger_label=f"{hhmm_str} 정기 브리핑")
                    send_telegram_message(report_txt)
                    sent_today.append(hhmm_str)

            # 2. 텔레그램 실시간 명령어 Long Polling (15초 대기)
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

                        # 본인 채팅방(MY_CHAT_ID)에서 온 명령어만 허용
                        if chat_id != str(MY_CHAT_ID) or not text:
                            continue

                        cmd = text.split()[0].lower()
                        print(f"📩 [명령어 수신] {now_dt.strftime('%H:%M:%S')} -> {text}")

                        if cmd in ("/분석", "분석", "/report", "/시황"):
                            reply = build_intraday_report(trigger_label="실시간 /분석 요청")
                            send_telegram_message(reply, chat_id)
                        elif cmd in ("/포지션", "포지션", "/잔고", "/status"):
                            reply = build_position_status_report()
                            send_telegram_message(reply, chat_id)
                        elif cmd in ("/상태", "/도움말", "/help", "/start"):
                            sent_list = ", ".join(sent_history.get(today_str, [])) or "없음"
                            help_txt = (
                                "🤖 <b>[장중 마켓시그널 봇 상태 및 명령어 안내]</b>\n"
                                f"• <b>현재 서버 시각:</b> {now_dt.strftime('%Y-%m-%d %H:%M:%S')}\n"
                                f"• <b>평일 자동 브리핑:</b> {', '.join(SCHEDULE_TIMES)}\n"
                                f"• <b>금일 발송 완료:</b> {sent_list}\n\n"
                                "💡 <b>[사용 가능한 명령어]</b>\n"
                                "• <code>/분석</code> (또는 <code>분석</code>) : 실시간 8대 ETF 주가·40EMA·3트랙 지침 즉시 조회\n"
                                "• <code>/포지션</code> (또는 <code>포지션</code>) : 보유 슬롯 실시간 평가 손익 조회\n"
                                "• <code>/상태</code> : 봇 정상 동작 여부 확인"
                            )
                            send_telegram_message(help_txt, chat_id)
            else:
                time.sleep(10)

        except KeyboardInterrupt:
            print("\n🛑 사용자 요청으로 장중 봇을 종료합니다.")
            break
        except Exception as e:
            print(f"⚠️ 폴링 루프 일시 오류 (5초 후 재시도): {e}")
            time.sleep(5)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="8대 ETF 장중 정기 브리핑 및 텔레그램 명령어 응답 봇")
    parser.add_argument("--once", action="store_true", help="데몬 대기 없이 즉시 1회 분석 리포트만 발송 후 종료")
    args = parser.parse_args()

    if args.once:
        msg = build_intraday_report(trigger_label="즉시 1회 실행")
        print(msg)
        ok = send_telegram_message(msg)
        if ok:
            print("✅ 텔레그램 발송 완료!")
        else:
            print("⚠️ 텔레그램 발송 실패")
    else:
        run_daemon()