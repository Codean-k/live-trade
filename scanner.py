"""
LIVE TRADE Scanner v3.5
- 코스피200 + 코스닥150 (350종목)
- 점수 만점 100점 (환산 없음)
- 가중치: 낙폭25 + 바닥다지기20 + 반등시작25 + 이격도15 + 수급15 = 100
- 페널티 별도 차감
- v3.5: details.chart_60d 추가 (60일 close + MA5/20/60 시계열, PWA 차트용)
"""

import os
import json
import time
from datetime import datetime, timedelta
import pandas as pd
from pykrx import stock

# ========================================
# 설정
# ========================================

TICKERS_FILE = "tickers.json"
OUTPUT_FILE = "scores.json"

# ========================================
# 유틸리티
# ========================================

def get_trading_dates(n_days=70):
    """최근 n일 영업일 리스트 (가장 최근이 마지막)"""
    today = datetime.now().strftime("%Y%m%d")
    start = (datetime.now() - timedelta(days=n_days + 20)).strftime("%Y%m%d")
    dates = stock.get_previous_business_days(fromdate=start, todate=today)
    return [d.strftime("%Y%m%d") for d in dates][-n_days:]


def safe_div(a, b, default=0):
    return a / b if b != 0 else default


# ========================================
# 캔들 패턴 인식
# ========================================

def detect_candle_pattern(df):
    """
    최근 캔들 패턴 인식
    Returns: (today, two_day, three_day)
    """
    if len(df) < 3:
        return "양봉", None, None

    last = df.iloc[-1]
    prev = df.iloc[-2]
    prev2 = df.iloc[-3]

    open_p, close, high, low = last["시가"], last["종가"], last["고가"], last["저가"]
    body = abs(close - open_p)
    candle_range = high - low

    today_pattern = "양봉" if close >= open_p else "음봉"

    if candle_range > 0:
        body_ratio = body / candle_range
        upper_tail = high - max(open_p, close)
        lower_tail = min(open_p, close) - low

        if body_ratio < 0.1:
            today_pattern = "도지"
        elif lower_tail > body * 2 and upper_tail < body * 0.5 and close > open_p:
            today_pattern = "망치형"
        elif upper_tail > body * 2 and lower_tail < body * 0.5 and close > open_p:
            today_pattern = "역망치"
        elif close > open_p and body_ratio > 0.7 and body > df["종가"].iloc[-20:].std() * 1.5:
            today_pattern = "장대양봉"
        elif close < open_p and body_ratio > 0.7 and body > df["종가"].iloc[-20:].std() * 1.5:
            today_pattern = "장대음봉"

    two_day_pattern = None
    if (
        prev["종가"] < prev["시가"]
        and close > open_p
        and open_p < prev["종가"]
        and close > prev["시가"]
    ):
        two_day_pattern = "상승장악"
    elif (
        prev["종가"] > prev["시가"]
        and close < open_p
        and open_p > prev["종가"]
        and close < prev["시가"]
    ):
        two_day_pattern = "하락장악"

    three_day_pattern = None
    if (
        prev2["종가"] < prev2["시가"]
        and abs(prev["종가"] - prev["시가"]) / max(prev["고가"] - prev["저가"], 1) < 0.3
        and close > open_p
        and close > (prev2["시가"] + prev2["종가"]) / 2
    ):
        three_day_pattern = "샛별"

    return today_pattern, two_day_pattern, three_day_pattern


# ========================================
# 일별 수급 (B 방식)
# ========================================

def fetch_daily_supply(dates, markets):
    """5일치 일별 수급. Returns: {ticker: [{date, foreign, inst}, ...]}"""
    supply_map = {}

    for date in dates:
        for market in markets:
            for investor in ["외국인", "기관합계"]:
                try:
                    df = stock.get_market_net_purchases_of_equities(
                        date, date, market, investor
                    )
                    if df is None or df.empty:
                        continue

                    value_col = None
                    for col in df.columns:
                        if "순매수" in col and ("대금" in col or "금액" in col):
                            value_col = col
                            break

                    if value_col is None:
                        continue

                    for ticker, row in df.iterrows():
                        if ticker not in supply_map:
                            supply_map[ticker] = {}
                        if date not in supply_map[ticker]:
                            supply_map[ticker][date] = {"foreign": 0, "inst": 0}

                        key = "foreign" if investor == "외국인" else "inst"
                        supply_map[ticker][date][key] = int(row[value_col])

                    time.sleep(0.3)
                except Exception as e:
                    print(f"  수급 수집 오류 ({date} {market} {investor}): {e}")
                    continue

    result = {}
    for ticker, date_data in supply_map.items():
        sorted_dates = sorted(date_data.keys())
        result[ticker] = [
            {
                "date": d[4:],
                "foreign": date_data[d]["foreign"],
                "inst": date_data[d]["inst"],
            }
            for d in sorted_dates
        ]
    return result


# ========================================
# 종목별 분석
# ========================================

def analyze_stock(ticker, name, market, daily_supply):
    """단일 종목 분석"""
    try:
        today = datetime.now().strftime("%Y%m%d")
        start = (datetime.now() - timedelta(days=100)).strftime("%Y%m%d")
        df = stock.get_market_ohlcv(start, today, ticker)

        if df is None or df.empty or len(df) < 30:
            return None

        if "거래대금" not in df.columns:
            df["거래대금"] = df["종가"] * df["거래량"]

        if "등락률" not in df.columns:
            df["등락률"] = df["종가"].pct_change() * 100

        last = df.iloc[-1]
        price = int(last["종가"])
        change_pct = float(last["등락률"])
        high_60d = float(df["고가"].iloc[-60:].max())
        low_60d = float(df["저가"].iloc[-60:].min())

        breakdown = {}

        # 1. 낙폭 (max 25)
        drop_60d = (price - high_60d) / high_60d * 100
        if drop_60d <= -25:
            drop_score = 25
        elif drop_60d <= -20:
            drop_score = 18
        elif drop_60d <= -15:
            drop_score = 10
        elif drop_60d <= -10:
            drop_score = 5
        else:
            drop_score = 0

        disp_60d = (price - df["종가"].iloc[-60:].mean()) / df["종가"].iloc[-60:].mean() * 100
        if drop_60d > -5 and disp_60d > 20:
            drop_score = max(0, drop_score - 5)
        breakdown["낙폭"] = drop_score

        # 2. 바닥다지기 (max 20)
        vol_5d = df["거래량"].iloc[-5:].mean()
        vol_20d = df["거래량"].iloc[-20:].mean()
        vol_60d = df["거래량"].iloc[-60:].mean()
        vol_ratio_5_20 = safe_div(vol_5d, vol_20d, 1)

        std_5d = df["종가"].iloc[-5:].std()
        std_10d = df["종가"].iloc[-10:].std()
        std_ratio = safe_div(std_5d, std_10d, 1)

        base_score = 0
        if vol_ratio_5_20 >= 1.3:
            base_score += 8
        elif vol_ratio_5_20 >= 1.0:
            base_score += 5
        if std_ratio < 0.7:
            base_score += 8
        elif std_ratio < 0.85:
            base_score += 5
        if std_ratio < 0.7 and vol_ratio_5_20 >= 1.2:
            base_score += 4
        breakdown["바닥다지기"] = min(base_score, 20)

        # 3. 반등시작 (max 25, 캔들 가산 포함)
        rebound_score = 0
        if change_pct > 0:
            rebound_score += 8
        vol_today_ratio = safe_div(last["거래량"], vol_20d, 1)
        if vol_today_ratio >= 1.5:
            rebound_score += 5

        ema12 = df["종가"].ewm(span=12).mean()
        ema26 = df["종가"].ewm(span=26).mean()
        macd = ema12 - ema26
        if len(macd) >= 2 and macd.iloc[-1] > 0 and macd.iloc[-2] <= 0:
            rebound_score += 5

        candle_today, candle_2d, candle_3d = detect_candle_pattern(df)

        if candle_today == "망치형":
            rebound_score += 4
        elif candle_today == "장대양봉":
            rebound_score += 6
        if candle_2d == "상승장악":
            rebound_score += 7
        if candle_3d == "샛별":
            rebound_score += 10

        breakdown["반등시작"] = min(rebound_score, 25)

        # 4. 이격도 (max 15)
        disp_20d = (price - df["종가"].iloc[-20:].mean()) / df["종가"].iloc[-20:].mean() * 100
        disp_60d_val = disp_60d

        disp_score = 0
        if -15 <= disp_20d <= -3:
            disp_score += 8
        elif -3 < disp_20d <= 0:
            disp_score += 4
        if -15 <= disp_60d_val <= -5:
            disp_score += 5

        try:
            ma20_5d_ago = df["종가"].iloc[-25:-5].mean()
            price_5d_ago = float(df["종가"].iloc[-6])
            disp_20d_5d_ago = (price_5d_ago - ma20_5d_ago) / ma20_5d_ago * 100
            disp_recovery = disp_20d - disp_20d_5d_ago
            if disp_recovery > 3:
                disp_score += 2
        except Exception:
            disp_recovery = 0

        breakdown["이격도"] = min(disp_score, 15)

        # 5. 수급 (max 15)
        supply_data = daily_supply.get(ticker, [])
        foreign_5d_total = sum(d["foreign"] for d in supply_data) if supply_data else 0
        inst_5d_total = sum(d["inst"] for d in supply_data) if supply_data else 0
        combined_5d = foreign_5d_total + inst_5d_total

        foreign_buy_days = sum(1 for d in supply_data if d["foreign"] > 0)
        inst_buy_days = sum(1 for d in supply_data if d["inst"] > 0)

        supply_score = 0
        if combined_5d > 0:
            supply_score += 8
        if foreign_5d_total > 0:
            supply_score += 4
        if inst_5d_total > 0:
            supply_score += 3
        breakdown["수급"] = min(supply_score, 15)

        # 매수/매도 우세
        if combined_5d > 0 and (foreign_buy_days + inst_buy_days) >= 5:
            dominance = "매수"
        elif combined_5d < 0 and (foreign_buy_days + inst_buy_days) <= 3:
            dominance = "매도"
        else:
            dominance = "중립"

        # 6. 페널티
        penalty = 0
        big_moves = sum(1 for r in df["등락률"].iloc[-60:].abs() if r >= 15)
        if big_moves >= 5:
            penalty -= 100
        elif big_moves >= 3:
            penalty -= 60
        elif big_moves >= 2:
            penalty -= 30

        if vol_today_ratio >= 10 and abs(change_pct) >= 10:
            penalty -= 100

        try:
            price_20d_ago = float(df["종가"].iloc[-21])
            if (price - price_20d_ago) / price_20d_ago >= 1.0:
                penalty -= 100
        except Exception:
            pass

        breakdown["페널티"] = penalty

        # 총점/등급
        raw_total = sum(breakdown.values())
        total = max(0, min(100, raw_total))

        if penalty <= -100:
            grade = "경고"
        elif total >= 80:
            grade = "S"
        elif total >= 60:
            grade = "A"
        elif total >= 40:
            grade = "B"
        elif total >= 20:
            grade = "C"
        else:
            grade = "D"

        # 이동평균
        ma5 = int(df["종가"].iloc[-5:].mean())
        ma20 = int(df["종가"].iloc[-20:].mean())
        ma60 = int(df["종가"].iloc[-60:].mean())

        if ma5 > ma20 > ma60:
            alignment = "정배열"
        elif ma5 < ma20 < ma60:
            alignment = "역배열"
        else:
            alignment = "혼조"

        # ATR
        df["tr"] = df.apply(
            lambda r: max(r["고가"] - r["저가"],
                          abs(r["고가"] - r["종가"]),
                          abs(r["저가"] - r["종가"])),
            axis=1,
        )
        atr = int(df["tr"].iloc[-14:].mean())

        # 시가총액
        try:
            cap_df = stock.get_market_cap(today, today, ticker)
            market_cap = int(cap_df["시가총액"].iloc[0]) if not cap_df.empty else 0
        except Exception:
            market_cap = 0

        return {
            "code": ticker,
            "name": name,
            "market": market,
            "price": price,
            "change_pct": round(change_pct, 2),
            "score": total,
            "grade": grade,
            "dominance": dominance,
            "breakdown": breakdown,
            "details": {
                "drop_60d": round(drop_60d, 1),
                "high_60d": int(high_60d),
                "low_60d": int(low_60d),
                "disparity_20d": round(disp_20d, 1),
                "disparity_60d": round(disp_60d_val, 1),
                "disp_recovery_5d": round(disp_recovery, 1),
                "volume_ratio_5d_vs_20d": round(vol_ratio_5_20, 2),
                "volume_today_vs_5d": round(safe_div(last["거래량"], vol_5d, 0), 2),
                "volume_today_vs_20d": round(vol_today_ratio, 2),
                "volume_today_vs_60d": round(safe_div(last["거래량"], vol_60d, 0), 2),
                "std_ratio_5d_vs_10d": round(std_ratio, 2),
                "std_shrink_pct": round((1 - std_ratio) * 100, 1),
                "avg_value_60d_won": int(df["거래대금"].iloc[-60:].mean()),
                "big_moves_15pct": int(big_moves),
                "atr_14": atr,
                "market_cap": market_cap,
                "ma5": ma5,
                "ma20": ma20,
                "ma60": ma60,
                "ma_alignment": alignment,
                "price_vs_ma5_pct": round((price - ma5) / ma5 * 100, 2),
                "price_vs_ma20_pct": round((price - ma20) / ma20 * 100, 2),
                "price_vs_ma60_pct": round((price - ma60) / ma60 * 100, 2),
                "candle_today": candle_today,
                "candle_2d": candle_2d,
                "candle_3d": candle_3d,
                "foreign_5d_total": foreign_5d_total,
                "inst_5d_total": inst_5d_total,
                "combined_5d": combined_5d,
                "foreign_buy_days": foreign_buy_days,
                "inst_buy_days": inst_buy_days,
                "supply_daily": supply_data,
                "chart_60d": [
                    {
                        "d": d.strftime("%Y%m%d"),
                        "c": int(c),
                        "m5": int(m5) if pd.notna(m5) else None,
                        "m20": int(m20) if pd.notna(m20) else None,
                        "m60": int(m60) if pd.notna(m60) else None,
                    }
                    for d, c, m5, m20, m60 in zip(
                        df.index[-60:],
                        df["종가"].iloc[-60:],
                        df["종가"].rolling(5).mean().iloc[-60:],
                        df["종가"].rolling(20).mean().iloc[-60:],
                        df["종가"].rolling(60).mean().iloc[-60:],
                    )
                ],
            },
        }

    except Exception as e:
        print(f"  분석 오류 ({ticker} {name}): {e}")
        return None


# ========================================
# 메인
# ========================================

def main():
    print("=" * 60)
    print("LIVE TRADE Scanner v3.4 (100점 만점)")
    print("=" * 60)

    with open(TICKERS_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)
    tickers = data["tickers"] if isinstance(data, dict) and "tickers" in data else data

    print(f"\n총 {len(tickers)}개 종목 스캔 시작")
    print(f"시작 시간: {datetime.now().strftime('%H:%M:%S')}\n")

    print("[1/2] 일별 수급 데이터 수집 중...")
    dates = get_trading_dates(n_days=5)
    print(f"  대상 영업일 (최근 5일): {dates}")

    daily_supply = fetch_daily_supply(dates, ["KOSPI", "KOSDAQ"])
    print(f"  수급 데이터 수집 완료: {len(daily_supply)}개 종목\n")

    print("[2/2] 종목별 분석 중...")
    results = []
    for i, t in enumerate(tickers):
        result = analyze_stock(t["code"], t["name"], t["market"], daily_supply)
        if result:
            results.append(result)
        if (i + 1) % 50 == 0:
            print(f"  진행: {i + 1}/{len(tickers)} ({len(results)}개 분석 성공)")

    results.sort(key=lambda x: -x["score"])

    output = {
        "updated": datetime.now().isoformat(),
        "version": "v3.4",
        "count": len(results),
        "results": results,
    }
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    grade_counts = {}
    for r in results:
        grade_counts[r["grade"]] = grade_counts.get(r["grade"], 0) + 1

    print(f"\n{'=' * 60}")
    print(f"완료: {len(results)}개 분석")
    print(f"등급 분포: {grade_counts}")
    print(f"저장: {OUTPUT_FILE}")
    print(f"종료 시간: {datetime.now().strftime('%H:%M:%S')}")


if __name__ == "__main__":
    main()
