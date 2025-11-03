uv run python rebalance_k_allweather_now.py --holdings "holding/2025년 구매수량.xlsx"

# 수수료 0.015% 반영 + 잔여현금 그리디(기본 on)

uv run python rebalance_k_allweather_now.py --holdings "holding/2025년 구매수량.xlsx" --fee 0.00015

# 잔여현금 추가매수 비활성화

uv run python rebalance_k_allweather_now.py --holdings "holding/2025년 구매수량.xlsx" --no-greedy
