"""
코스피200 + 코스닥150 종목 리스트 생성 (pykrx 버전, KRX 로그인 사용)
출력: tickers.json
"""
from pykrx import stock
import json
from datetime import datetime, timedelta

def get_last_business_day():
    d = datetime.now()
    for _ in range(10):
        ymd = d.strftime("%Y%m%d")
        df = stock.get_index_ohlcv(ymd, ymd, "1028")
        if not df.empty:
            return ymd
        d -= timedelta(days=1)
    raise RuntimeError("영업일 못 찾음")

date = get_last_business_day()
print(f"기준일: {date}")

# 코스피200
kospi200 = stock.get_index_portfolio_deposit_file("1028", date)
print(f"코스피200: {len(kospi200)}종목")

# 코스닥150
kosdaq150 = stock.get_index_portfolio_deposit_file("2203", date)
print(f"코스닥150: {len(kosdaq150)}종목")

tickers = []
for code in kospi200:
    name = stock.get_market_ticker_name(code)
    tickers.append({"code": code, "name": name, "market": "KOSPI200"})

for code in kosdaq150:
    name = stock.get_market_ticker_name(code)
    tickers.append({"code": code, "name": name, "market": "KOSDAQ150"})

result = {
    "updated": date,
    "count": len(tickers),
    "tickers": tickers
}

with open("tickers.json", "w", encoding="utf-8") as f:
    json.dump(result, f, ensure_ascii=False, indent=2)

print(f"\n✅ 총 {len(tickers)}종목 → tickers.json 저장 완료")