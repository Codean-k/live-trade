"""
LIVE TRADE - 점수 공식 v3.2 스캐너
코스피200 + 코스닥150 (350종목) → scores.json

v3.2 변경:
  - 수급 함수 수정 (investor 파라미터)
  - 이격도 구간 확대 (-15%까지 점수 인정)
  - 작전주 페널티 단계화 (-30/-60/-100)
  - 신고가 종목 미세 보정 (-5)
"""

import json
import time
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
from pykrx import stock


# ============================================================
# 유틸
# ============================================================

def log(msg):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


def get_business_days(end_date_str: str, n: int = 70):
    end = datetime.strptime(end_date_str, "%Y%m%d")
    start = end - timedelta(days=int(n * 1.6))
    df = stock.get_index_ohlcv(
        start.strftime("%Y%m%d"),
        end.strftime("%Y%m%d"),
        "1028",
    )
    days = [d.strftime("%Y%m%d") for d in df.index]
    return days[-n:]


def get_last_business_day():
    d = datetime.now()
    for _ in range(10):
        ymd = d.strftime("%Y%m%d")
        df = stock.get_index_ohlcv(ymd, ymd, "1028")
        if not df.empty:
            return ymd
        d -= timedelta(days=1)
    raise RuntimeError("영업일 못 찾음")


# ============================================================
# 데이터 수집
# ============================================================

def fetch_ohlcv(code: str, start: str, end: str) -> pd.DataFrame:
    df = stock.get_market_ohlcv(start, end, code)
    if df.empty:
        return df
    df = df.rename(columns={
        "시가": "open", "고가": "high", "저가": "low",
        "종가": "close", "거래량": "volume", "거래대금": "value",
        "등락률": "change_pct",
    })
    if "value" not in df.columns:
        df["value"] = df["volume"] * df["close"]
    return df


def fetch_supply_bulk(start: str, end: str):
    """
    KOSPI/KOSDAQ × 외국인/기관 4개 한 번에 받아서 dict로
    return: {code: {"foreign": int, "institution": int}}
    """
    supply = {}

    for market in ["KOSPI", "KOSDAQ"]:
        for investor, key in [("외국인", "foreign"), ("기관합계", "institution")]:
            try:
                df = stock.get_market_net_purchases_of_equities_by_ticker(
                    start, end, market, investor,
                )
                if df.empty:
                    continue
                for code, row in df.iterrows():
                    code_str = str(code).zfill(6)
                    if code_str not in supply:
                        supply[code_str] = {"foreign": 0, "institution": 0}
                    supply[code_str][key] = int(row.get("순매수거래대금", 0))
            except Exception as e:
                log(f"⚠️ 수급 {market}/{investor} 실패: {e}")

    return supply


# ============================================================
# 점수 계산
# ============================================================

def calc_score(df: pd.DataFrame, supply_row: dict):
    if len(df) < 60:
        return None, None, None

    close = df["close"]
    high = df["high"]
    low = df["low"]
    vol = df["volume"]
    val = df["value"]

    today = df.iloc[-1]

    breakdown = {
        "낙폭": 0, "바닥다지기": 0, "반등시작": 0,
        "이격도": 0, "수급": 0, "페널티": 0,
    }
    details = {}

    # ──────────────────────────────────────────
    # ① 낙폭 (30점)
    # ──────────────────────────────────────────
    high_60d = high.iloc[-60:].max()
    drop = (today["close"] - high_60d) / high_60d * 100
    details["drop_60d"] = round(drop, 2)

    if drop <= -25:
        breakdown["낙폭"] = 30
    elif drop <= -20:
        breakdown["낙폭"] = 24
    elif drop <= -15:
        breakdown["낙폭"] = 18
    elif drop <= -10:
        breakdown["낙폭"] = 12
    else:
        breakdown["낙폭"] = 6

    # ──────────────────────────────────────────
    # ② 바닥다지기 (40점)
    # ──────────────────────────────────────────
    vol_5d = vol.iloc[-5:].mean()
    vol_20d = vol.iloc[-20:].mean()
    vol_ratio = vol_5d / vol_20d if vol_20d > 0 else 0
    details["volume_ratio_5d_vs_20d"] = round(vol_ratio, 2)
    if vol_ratio >= 1.2:
        breakdown["바닥다지기"] += 15

    std_5d = close.iloc[-5:].std()
    std_10d = close.iloc[-10:].std()
    details["std_ratio_5d_vs_10d"] = round(std_5d / std_10d, 2) if std_10d > 0 else 0
    if std_10d > 0 and std_5d < std_10d:
        breakdown["바닥다지기"] += 15

    is_bear = today["close"] < today["open"]
    if is_bear:
        candle_range = today["high"] - today["low"]
        if candle_range > 0:
            tail_recovery = (today["close"] - today["low"]) / candle_range
            details["tail_recovery"] = round(tail_recovery, 2)
            if tail_recovery >= 0.5:
                breakdown["바닥다지기"] += 10

    # ──────────────────────────────────────────
    # ③ 반등시작 (20점)
    # ──────────────────────────────────────────
    is_bull = today["close"] > today["open"]
    today_vol_ratio = today["volume"] / vol_20d if vol_20d > 0 else 0
    details["today_volume_ratio"] = round(today_vol_ratio, 2)

    if is_bull and today_vol_ratio >= 1.5:
        breakdown["반등시작"] += 7
        if today_vol_ratio >= 2.0:
            breakdown["반등시작"] += 3

    ma5 = close.rolling(5).mean()
    if len(ma5) >= 2 and pd.notna(ma5.iloc[-1]) and pd.notna(ma5.iloc[-2]):
        if ma5.iloc[-1] > ma5.iloc[-2]:
            breakdown["반등시작"] += 5

    ema12 = close.ewm(span=12, adjust=False).mean()
    ema26 = close.ewm(span=26, adjust=False).mean()
    macd = ema12 - ema26
    signal = macd.ewm(span=9, adjust=False).mean()
    hist = macd - signal
    if len(hist) >= 2 and hist.iloc[-2] < 0 and hist.iloc[-1] > 0:
        breakdown["반등시작"] += 5

    # ──────────────────────────────────────────
    # ④ 이격도 (15점) - v3.2 완화
    # ──────────────────────────────────────────
    ma20 = close.rolling(20).mean().iloc[-1]
    ma60 = close.rolling(60).mean().iloc[-1]
    disp20 = (today["close"] - ma20) / ma20 * 100 if ma20 > 0 else 0
    disp60 = (today["close"] - ma60) / ma60 * 100 if ma60 > 0 else 0
    details["disparity_20d"] = round(disp20, 2)
    details["disparity_60d"] = round(disp60, 2)

    # 20일 이격도 (8점) - v3.2: -15% 까지 점수
    score_d20 = 0
    if -15 <= disp20 <= -10:
        score_d20 = 4
    elif -10 < disp20 <= -7:
        score_d20 = 4
    elif -7 < disp20 <= -3:
        score_d20 = 3
    if len(close) >= 25:
        ma20_5d_ago = close.iloc[-25:-5].rolling(20).mean().iloc[-1]
        if pd.notna(ma20_5d_ago) and ma20_5d_ago > 0:
            disp20_5d_ago = (close.iloc[-6] - ma20_5d_ago) / ma20_5d_ago * 100
            if disp20 - disp20_5d_ago >= 3:
                score_d20 += 2
            if disp20_5d_ago <= -10:
                score_d20 += 2
    breakdown["이격도"] += min(score_d20, 8)

    # 60일 이격도 (7점) - v3.2: -25%까지 점수
    score_d60 = 0
    if -25 <= disp60 <= -20:
        score_d60 = 3
    elif -20 < disp60 <= -15:
        score_d60 = 3
    elif -15 < disp60 <= -10:
        score_d60 = 4
    if len(close) >= 65:
        ma60_5d_ago = close.iloc[-65:-5].rolling(60).mean().iloc[-1]
        if pd.notna(ma60_5d_ago) and ma60_5d_ago > 0:
            disp60_5d_ago = (close.iloc[-6] - ma60_5d_ago) / ma60_5d_ago * 100
            if disp60 - disp60_5d_ago >= 2:
                score_d60 += 2
            if disp60_5d_ago <= -20:
                score_d60 += 1
    breakdown["이격도"] += min(score_d60, 7)

    # ──────────────────────────────────────────
    # ⑤ 수급 (15점) - v3.2: bulk dict 룩업
    # ──────────────────────────────────────────
    if supply_row:
        foreign = supply_row.get("foreign", 0)
        inst = supply_row.get("institution", 0)
        value_5d = val.iloc[-5:].sum()

        if value_5d > 0:
            f_ratio = foreign / value_5d * 100
            i_ratio = inst / value_5d * 100
            details["foreign_ratio_pct"] = round(f_ratio, 2)
            details["institution_ratio_pct"] = round(i_ratio, 2)

            if foreign > 0:
                breakdown["수급"] += 4
                if f_ratio >= 3:
                    breakdown["수급"] += 2
                if f_ratio >= 5:
                    breakdown["수급"] += 2

            if inst > 0:
                breakdown["수급"] += 3
                if i_ratio >= 3:
                    breakdown["수급"] += 2
                if i_ratio >= 5:
                    breakdown["수급"] += 2

    # ──────────────────────────────────────────
    # 페널티 - v3.2: 작전주 단계화
    # ──────────────────────────────────────────
    avg_value_60d = val.iloc[-60:].mean()
    details["avg_value_60d_won"] = int(avg_value_60d)
    if avg_value_60d < 500_000_000:
        breakdown["페널티"] -= 15

    # 작전주: 일일 변동률 ±15% 횟수
    daily_change = (close.pct_change().abs() * 100).iloc[-60:]
    big_moves = (daily_change >= 15).sum()
    details["big_moves_15pct"] = int(big_moves)
    if big_moves >= 7:
        breakdown["페널티"] -= 100
    elif big_moves >= 5:
        breakdown["페널티"] -= 60
    elif big_moves >= 3:
        breakdown["페널티"] -= 30

    # 작전주: 거래량 10배 + ±10% (강력한 신호)
    vol_spike = (vol.iloc[-60:] > vol.iloc[-60:].mean() * 10)
    big_change = (daily_change >= 10)
    if (vol_spike & big_change).any():
        breakdown["페널티"] -= 100

    # 작전주: 20일 100%+ 급등 (확실한 작전)
    rolling_return_20d = (close / close.shift(20) - 1) * 100
    if (rolling_return_20d.iloc[-60:] >= 100).any():
        breakdown["페널티"] -= 100

    # v3.2 신규: 신고가 종목 미세 보정 (역발상 전략에 안 맞음)
    if drop > -5 and disp60 > 20:
        breakdown["페널티"] -= 5
        details["near_high_warning"] = True

    raw = sum(breakdown.values())
    return raw, breakdown, details


def to_grade(score_100: int):
    if score_100 < 0:
        return "경고"
    if score_100 >= 85:
        return "S"
    if score_100 >= 70:
        return "A"
    if score_100 >= 50:
        return "B"
    if score_100 >= 30:
        return "C"
    return "D"


# ============================================================
# 메인
# ============================================================

def main():
    log("LIVE TRADE 스캐너 v3.2 시작")

    with open("tickers.json", "r", encoding="utf-8") as f:
        ticker_data = json.load(f)
    tickers = ticker_data["tickers"]
    log(f"종목 {len(tickers)}개 로드")

    end = get_last_business_day()
    days = get_business_days(end, 65)
    start = days[0]
    supply_start = days[-5]
    log(f"기간: {start} ~ {end} (수급은 {supply_start} ~ {end})")

    # 수급 데이터 한 번에 받기 (4번 호출, 빠름)
    log("수급 데이터 받는 중...")
    supply_data = fetch_supply_bulk(supply_start, end)
    log(f"수급 데이터 {len(supply_data)}개 로드")

    results = []
    fail_count = 0

    for i, t in enumerate(tickers, 1):
        code = t["code"]
        name = t["name"]
        market = t["market"]

        try:
            df = fetch_ohlcv(code, start, end)
            if df.empty or len(df) < 60:
                fail_count += 1
                continue

            supply_row = supply_data.get(code, None)

            raw, breakdown, details = calc_score(df, supply_row)
            if raw is None:
                fail_count += 1
                continue

            score_100 = round(raw * 100 / 120)
            grade = to_grade(score_100)
            today_row = df.iloc[-1]
            prev_row = df.iloc[-2]

            if "change_pct" in df.columns:
                change_pct = float(today_row["change_pct"])
            else:
                change_pct = (today_row["close"] - prev_row["close"]) / prev_row["close"] * 100

            results.append({
                "code": code,
                "name": name,
                "market": market,
                "score": score_100,
                "grade": grade,
                "price": int(today_row["close"]),
                "change_pct": round(change_pct, 2),
                "breakdown": breakdown,
                "details": details,
            })

            if i % 50 == 0:
                log(f"진행 {i}/{len(tickers)} (실패 {fail_count}건)")

            time.sleep(0.15)
        except Exception as e:
            fail_count += 1
            if fail_count <= 5:
                log(f"⚠️ {code} {name} 실패: {e}")

    grade_order = {"S": 0, "A": 1, "B": 2, "C": 3, "D": 4, "경고": 5}
    results.sort(key=lambda r: (grade_order[r["grade"]], -r["score"]))

    output = {
        "updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "trading_date": end,
        "scan_count": len(results),
        "fail_count": fail_count,
        "scoring_version": "v3.2",
        "results": results,
    }

    Path("scores.json").write_text(
        json.dumps(output, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    grade_counts = {}
    for r in results:
        grade_counts[r["grade"]] = grade_counts.get(r["grade"], 0) + 1

    log("=" * 50)
    log(f"✅ 완료: 성공 {len(results)} / 실패 {fail_count}")
    for g in ["S", "A", "B", "C", "D", "경고"]:
        if g in grade_counts:
            log(f"  {g}급: {grade_counts[g]}개")
    log("scores.json 저장됨")


if __name__ == "__main__":
    main()