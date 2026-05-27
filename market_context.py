"""
LIVE TRADE — Market Context Module (v3.7)
================================
scanner.py가 호출하는 시장 지표 계산 모듈.

검증된 사항 (2026-05-28):
- pykrx get_market_trading_value_by_investor 컬럼: ['매도', '매수', '순매수'] (단위: 원)
- pykrx index: ['금융투자','보험','투신','사모','은행','기타금융','연기금 등',
                '기관합계','기타법인','개인','외국인','기타외국인','전체']
- scanner.py v3.7이 details.today_value_won 제공

영향:
- 기존 점수 계산 로직 변경 없음
- scores.json 끝에 새 필드 추가만
- PWA가 이 필드 안 읽어도 기존처럼 동작 (하위 호환)
"""

from pykrx import stock
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

KST = ZoneInfo("Asia/Seoul")


# ========================================
# 영업일 fallback 헬퍼
# ========================================
def _find_recent_business_day(max_lookback_days=10):
    """
    오늘부터 거꾸로 거슬러 올라가며 첫 영업일을 반환.
    pykrx get_previous_business_days로 정확히 영업일만 가져옴.
    """
    today = datetime.now(KST).strftime("%Y%m%d")
    start = (datetime.now(KST) - timedelta(days=max_lookback_days)).strftime("%Y%m%d")
    try:
        dates = stock.get_previous_business_days(fromdate=start, todate=today)
        if dates:
            return dates[-1].strftime("%Y%m%d")
    except Exception as e:
        print(f"  영업일 조회 실패, today 사용: {e}")
    return today


# ========================================
# 1. 코스피/코스닥 지수
# ========================================
def get_market_indices():
    """
    코스피 / 코스닥 지수 OHLCV.

    pykrx 지수 코드:
    - 1001: KOSPI
    - 2001: KOSDAQ
    - 1028: KOSPI200
    - 2203: KOSDAQ150
    """
    today = datetime.now(KST).strftime("%Y%m%d")
    # 14일 잡는 이유: 연휴(설/추석) 끼면 7일도 부족할 수 있음
    start = (datetime.now(KST) - timedelta(days=14)).strftime("%Y%m%d")

    indices = {}
    targets = {
        "KOSPI": "1001",
        "KOSDAQ": "2001",
        "KOSPI200": "1028",
        "KOSDAQ150": "2203",
    }

    for name, code in targets.items():
        try:
            df = stock.get_index_ohlcv(start, today, code)
            if df is None or df.empty or len(df) < 2:
                indices[name] = None
                continue
            today_row = df.iloc[-1]
            prev_row = df.iloc[-2]
            prev_close = float(prev_row["종가"])
            today_close = float(today_row["종가"])
            change = today_close - prev_close
            change_pct = (change / prev_close * 100) if prev_close > 0 else 0

            indices[name] = {
                "value": round(today_close, 2),
                "change": round(change, 2),
                "change_pct": round(change_pct, 2),
                "volume": int(today_row["거래량"]),
                "value_amount": int(today_row["거래대금"]),
                "as_of": df.index[-1].strftime("%Y-%m-%d"),
            }
        except Exception as e:
            print(f"  지수 {name} 가져오기 실패: {e}")
            indices[name] = None

    return indices


# ========================================
# 2. 시장 폭 (Market Breadth)
# ========================================
def calculate_breadth(results):
    """
    우리 350종목 기준 시장 폭.

    시장 폭 넓음 = 광범위한 상승 (좋은 시장)
    시장 폭 좁음 = 소수 종목만 상승 (위험)
    """
    total = len(results)
    if total == 0:
        return None

    advance = sum(1 for r in results if r["change_pct"] > 0.05)
    decline = sum(1 for r in results if r["change_pct"] < -0.05)
    unchanged = total - advance - decline

    strong_up = sum(1 for r in results if r["change_pct"] > 3)
    strong_down = sum(1 for r in results if r["change_pct"] < -3)

    kospi = [r for r in results if r["market"] == "KOSPI200"]
    kosdaq = [r for r in results if r["market"] == "KOSDAQ150"]

    return {
        "total": total,
        "advance": advance,
        "decline": decline,
        "unchanged": unchanged,
        "advance_pct": round(advance / total * 100, 1),
        "strong_up": strong_up,
        "strong_down": strong_down,
        "kospi200": {
            "total": len(kospi),
            "advance": sum(1 for r in kospi if r["change_pct"] > 0.05),
            "avg_change": round(sum(r["change_pct"] for r in kospi) / len(kospi), 2) if kospi else 0,
        },
        "kosdaq150": {
            "total": len(kosdaq),
            "advance": sum(1 for r in kosdaq if r["change_pct"] > 0.05),
            "avg_change": round(sum(r["change_pct"] for r in kosdaq) / len(kosdaq), 2) if kosdaq else 0,
        },
    }


# ========================================
# 3. 거래대금 상위 + 집중도
# ========================================
def get_volume_leaders(results, top_n=10):
    """
    오늘 거래대금 상위 종목 + 집중도.

    scanner.py v3.7이 details.today_value_won을 제공한다는 전제.
    (없으면 avg_value_60d_won으로 fallback — 옛 데이터 호환)

    집중도 높음 = 소수 종목에 거래 쏠림 (위험 신호)
    집중도 낮음 = 거래 분산 (건강)
    """
    volume_data = []
    fallback_used = 0
    for r in results:
        dt = r.get("details", {})
        today_value = dt.get("today_value_won")
        if today_value is None:
            today_value = dt.get("avg_value_60d_won", 0)
            fallback_used += 1

        volume_data.append({
            "code": r["code"],
            "name": r["name"],
            "market": r["market"],
            "price": r["price"],
            "change_pct": r["change_pct"],
            "score": r["score"],
            "grade": r["grade"],
            "dominance": r.get("dominance", "중립"),
            "today_value": int(today_value),
        })

    if fallback_used > 0:
        print(f"  ⚠ today_value_won 없는 종목 {fallback_used}개: 60d 평균으로 fallback")

    # 거래대금 내림차순
    volume_data.sort(key=lambda x: -x["today_value"])
    top = volume_data[:top_n]

    total_value = sum(v["today_value"] for v in volume_data)
    top_value = sum(v["today_value"] for v in top)
    concentration = (top_value / total_value * 100) if total_value > 0 else 0

    return {
        "top_n": top_n,
        "concentration_pct": round(concentration, 1),
        "total_value_won": total_value,
        "leaders": [
            {
                "code": v["code"],
                "name": v["name"],
                "market": v["market"],
                "price": v["price"],
                "change_pct": v["change_pct"],
                "score": v["score"],
                "grade": v["grade"],
                "dominance": v["dominance"],
                "today_value_won": v["today_value"],
            }
            for v in top
        ],
    }


# ========================================
# 4. 점수 분포
# ========================================
def calculate_score_distribution(results):
    """우리 점수 시스템의 등급 분포."""
    dist = {"S": 0, "A": 0, "B": 0, "C": 0, "D": 0, "경고": 0}
    for r in results:
        g = r.get("grade", "D")
        if g in dist:
            dist[g] += 1

    total = len(results)
    return {
        "counts": dist,
        "pct": {g: round(c / total * 100, 1) if total > 0 else 0 for g, c in dist.items()},
        "buy_candidates": dist["S"] + dist["A"],
        "buy_candidates_pct": round((dist["S"] + dist["A"]) / total * 100, 1) if total > 0 else 0,
    }


# ========================================
# 5. 외국인/기관 시장 전체 자금 흐름
# ========================================
# 검증된 pykrx 응답 (2026-05-28):
#   df.columns = ['매도', '매수', '순매수']  (단위: 원)
#   df.index   = ['금융투자','보험','투신','사모','은행','기타금융','연기금 등',
#                 '기관합계','기타법인','개인','외국인','기타외국인','전체']

# 디스플레이용으로 노출할 투자자만 선별
KEY_INVESTORS = ["외국인", "기관합계", "개인", "연기금 등", "금융투자", "투신"]


def get_market_flow():
    """
    시장 전체 외국인/기관/개인 등 순매수 (가장 최근 영업일 기준).
    """
    target_date = _find_recent_business_day()

    flow = {}
    for market in ["KOSPI", "KOSDAQ"]:
        try:
            df = stock.get_market_trading_value_by_investor(target_date, target_date, market)

            if df is None or df.empty:
                # 한 번 더 거슬러서 시도
                start = (datetime.now(KST) - timedelta(days=10)).strftime("%Y%m%d")
                try:
                    dates = stock.get_previous_business_days(fromdate=start, todate=target_date)
                    if len(dates) >= 2:
                        retry_date = dates[-2].strftime("%Y%m%d")
                        df = stock.get_market_trading_value_by_investor(retry_date, retry_date, market)
                        if not df.empty:
                            target_date = retry_date
                except Exception:
                    pass

            if df is None or df.empty:
                print(f"  시장 {market} 자금 흐름: 데이터 없음")
                flow[market] = None
                continue

            result = {}
            for investor in KEY_INVESTORS:
                if investor not in df.index:
                    continue
                row = df.loc[investor]
                buy = int(row["매수"])
                sell = int(row["매도"])
                net = int(row["순매수"])
                result[investor] = {
                    "buy": buy,
                    "sell": sell,
                    "net": net,
                }
            # 전체 거래대금도 참고용으로
            if "전체" in df.index:
                result["_total_value"] = int(df.loc["전체", "매수"])

            flow[market] = {
                "as_of": f"{target_date[:4]}-{target_date[4:6]}-{target_date[6:8]}",
                "investors": result,
            }
        except Exception as e:
            print(f"  시장 {market} 자금 흐름 실패: {e}")
            flow[market] = None

    return flow


# ========================================
# 6. 매수 후보 계산 (Today's Candidates)
# ========================================
# 두 관점 + 교집합:
#  관점 1 (시장 흐름): 거래대금 Top 30 ∩ 5일 수급+매수우세 ∩ 경고 제외
#  관점 2 (낙폭 분석): 점수 50+ ∩ 경고 제외
#  ★ SWEET SPOT     : 위 둘의 교집합

def _slim_candidate(r, today_value_won=None, rank=None):
    """후보 카드용 슬림 dict. results 원본보다 가벼움."""
    dt = r.get("details", {}) or {}
    out = {
        "code": r["code"],
        "name": r["name"],
        "market": r["market"],
        "price": r["price"],
        "change_pct": r["change_pct"],
        "score": r["score"],
        "grade": r["grade"],
        "dominance": r.get("dominance", "중립"),
        "drop_60d": dt.get("drop_60d"),
        "combined_5d": dt.get("combined_5d", 0),
        "today_value_won": dt.get("today_value_won", 0),
    }
    if today_value_won is not None:
        out["today_value_won"] = int(today_value_won)
    if rank is not None:
        out["volume_rank"] = rank
    return out


def calculate_candidates(results):
    """
    Today's Candidates 산출.
    기준 (Dean 확정, 2026-05-28):
      관점 1: 거래대금 Top 30 + combined_5d > 0 + dominance == "매수" + grade != "경고"
      관점 2: score >= 50 + grade != "경고"
      SWEET : 위 둘의 교집합
    노출: 전부 (상한 없음)
    """
    # ===== 관점 1: 시장 흐름 (주도주) =====
    # 오늘 거래대금 내림차순, Top 30
    valid_for_volume = [
        (r, (r.get("details") or {}).get("today_value_won", 0))
        for r in results
    ]
    valid_for_volume.sort(key=lambda x: -x[1])
    top30 = valid_for_volume[:30]

    trend_codes = set()
    trend_candidates = []
    for rank, (r, tv) in enumerate(top30, start=1):
        if r.get("grade") == "경고":
            continue
        dt = r.get("details") or {}
        if dt.get("combined_5d", 0) <= 0:
            continue
        if r.get("dominance") != "매수":
            continue
        trend_candidates.append(_slim_candidate(r, today_value_won=tv, rank=rank))
        trend_codes.add(r["code"])

    # ===== 관점 2: 낙폭 분석 (역발상) =====
    drop_codes = set()
    drop_candidates = []
    # 점수 내림차순 (이미 results가 점수순으로 정렬돼 있지만 안전하게)
    sorted_by_score = sorted(results, key=lambda r: -r.get("score", 0))
    for r in sorted_by_score:
        if r.get("grade") == "경고":
            continue
        if r.get("score", 0) < 50:
            break  # 정렬돼 있으니 50 미만 만나면 종료
        drop_candidates.append(_slim_candidate(r))
        drop_codes.add(r["code"])

    # ===== SWEET SPOT: 교집합 =====
    sweet_codes = trend_codes & drop_codes
    sweet_candidates = []
    # 점수 높은 순으로 정렬
    for r in sorted_by_score:
        if r["code"] in sweet_codes:
            # today_value_won도 함께
            tv = (r.get("details") or {}).get("today_value_won", 0)
            sweet_candidates.append(_slim_candidate(r, today_value_won=tv))

    return {
        "criteria": {
            "trend": "거래대금 Top 30 + 외인/기관 5일 매수 + 매수 우세 + 경고 제외",
            "drop": "점수 50+ + 경고 제외",
            "sweet": "두 관점 교집합",
        },
        "trend": {
            "count": len(trend_candidates),
            "items": trend_candidates,
        },
        "drop": {
            "count": len(drop_candidates),
            "items": drop_candidates,
        },
        "sweet_spot": {
            "count": len(sweet_candidates),
            "items": sweet_candidates,
        },
    }


# ========================================
# 7. 메인: 모든 시장 지표 통합
# ========================================
def build_market_context(results):
    """
    모든 시장 지표를 하나의 dict로 통합.
    scanner.py main()에서 호출.
    """
    print("\n[시장 지표 계산]")

    print("  - 코스피/코스닥 지수...")
    indices = get_market_indices()

    print("  - 시장 폭 계산...")
    breadth = calculate_breadth(results)

    print("  - 거래대금 상위...")
    volume = get_volume_leaders(results, top_n=10)

    print("  - 점수 분포...")
    score_dist = calculate_score_distribution(results)

    print("  - 시장 전체 자금 흐름...")
    flow = get_market_flow()

    print("  - Today's Candidates 산출...")
    candidates = calculate_candidates(results)
    print(f"    관점 1 (시장 흐름) {candidates['trend']['count']}개")
    print(f"    관점 2 (낙폭 분석) {candidates['drop']['count']}개")
    print(f"    ★ SWEET SPOT       {candidates['sweet_spot']['count']}개")

    return {
        "indices": indices,
        "breadth": breadth,
        "volume_leaders": volume,
        "score_distribution": score_dist,
        "investor_flow": flow,
        "candidates": candidates,
        "calculated_at": datetime.now(KST).isoformat(),
    }


# ========================================
# 모듈 단독 실행 (디버깅용)
# ========================================
if __name__ == "__main__":
    import json
    print("=== get_market_indices ===")
    print(json.dumps(get_market_indices(), ensure_ascii=False, indent=2))
    print()
    print("=== get_market_flow ===")
    print(json.dumps(get_market_flow(), ensure_ascii=False, indent=2))
