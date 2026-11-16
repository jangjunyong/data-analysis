# -*- coding: utf-8 -*-
"""W1-6 일별 하루 매출 지수 추이(5/1~8/31), W1-7 요일×시간대 주문 분포(홀 전표+배달, 6/9~8/31). 금액은 지수로만 출력."""
import importlib.util
from pathlib import Path
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

HERE = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location("w1", HERE / "w1_분석.py")
w1 = importlib.util.module_from_spec(spec); spec.loader.exec_module(w1)

# ── W1-6 일별 추이
pos = w1.load_pos_daily().set_index("날짜")
o = w1.load_orders()
deliv = o.groupby("날짜")["음식판매"].sum()
d = pos.loc["2026-05-01":"2026-08-31", ["POS"]].copy()
d["배달"] = deliv.reindex(d.index).fillna(0)
d["합산"] = d["POS"] + d["배달"]
base = d.loc[d.index.month == 5, "합산"].mean()
idx = d["합산"] / base * 100
roll = idx.rolling(7, min_periods=4, center=True).mean()
fig, ax = plt.subplots(figsize=(10, 4.2))
ax.bar(idx.index, idx, width=0.8, color="#9AA3AE", alpha=0.45, label="하루 매출 지수(홀·직접 포장+배달앱)")
ax.plot(roll.index, roll, color=w1.NAVY, lw=2.4, label="앞뒤 7일 평균")
for m in [5, 6, 7, 8]:
    s = idx[idx.index.month == m]
    ax.hlines(s.mean(), s.index.min(), s.index.max(), color=w1.GOLD, lw=2, ls="--")
    ax.text(s.index.min(), s.mean() + 5, f"{m}월 평균 {s.mean():.0f}", fontsize=9, color=w1.TEXT, fontweight="bold", bbox=dict(fc="white", ec="none", pad=1, alpha=0.85))
ax.axhline(100, color=w1.GRAY, lw=0.8, ls=":")
ax.set_ylabel("하루 매출 지수(5월 평균=100)")
ax.xaxis.set_major_locator(mdates.MonthLocator()); ax.xaxis.set_major_formatter(mdates.DateFormatter("%m월"))
for s in ["top", "right"]: ax.spines[s].set_visible(False)
ax.legend(loc="upper right", fontsize=8.5)
ax.set_title("하루 매출은 날마다 기복이 크고, 6월부터 내려가 8월 평균은 5월보다 약 30% 낮다", fontsize=11.5, loc="left", pad=10)
w1.footnote(fig, "출처: 가게 POS 일자별 매출(홀·직접 포장), 배달앱 3곳 주문(음식 판매금액). 2026-05-01~08-31 영업일. 지수는 5월 영업일 평균=100. 점선=월 평균.", y=-0.04)
fig.savefig(HERE / "W1-6_일별_매출지수_추이.png")
print("W1-6 월평균", {m: round(idx[idx.index.month == m].mean(), 1) for m in [5, 6, 7, 8]}, "일별 최소·최대", round(idx.min()), round(idx.max()))

# ── W1-7 요일×시간대 주문 분포 (6/9~8/31)
t, hall = w1.load_receipts()
h = hall[(hall["영업일자"] >= "2026-06-09") & (hall["영업일자"] <= "2026-08-31")].copy()
h["요일"] = h["영업일자"].dt.dayofweek
od = o[(o["날짜"] >= "2026-06-09") & (o["날짜"] <= "2026-08-31")].copy()
od["요일"] = od["날짜"].dt.dayofweek
days = pd.Series(sorted(t[(t["영업일자"] >= "2026-06-09") & (t["영업일자"] <= "2026-08-31")]["영업일자"].unique()))
nday = days.dt.dayofweek.value_counts()
HOURS = list(range(14, 24)) + [0, 1]
def grid(df, hcol="시"):
    g = df.groupby(["요일", hcol]).size().unstack(fill_value=0).reindex(index=range(6), columns=HOURS, fill_value=0)
    return g.div(nday.reindex(range(6)), axis=0)
G = grid(h) + grid(od)
fig, ax = plt.subplots(figsize=(8, 3.4))
im = ax.imshow(G.values, aspect="auto", cmap="Blues"); ax.grid(False)
ax.set_xticks(range(len(HOURS))); ax.set_xticklabels([f"{x}시" for x in HOURS], fontsize=8.5)
ax.set_yticks(range(6)); ax.set_yticklabels(["월", "화", "수", "목", "금", "토"])
for i in range(6):
    for j in range(len(HOURS)):
        v = G.values[i, j]
        if v >= 0.3: ax.text(j, i, f"{v:.1f}", ha="center", va="center", fontsize=8.5, color="white" if v > G.values.max() * 0.6 else w1.TEXT)
cb = fig.colorbar(im, ax=ax, fraction=0.025, pad=0.01); cb.set_label("영업일당 주문 수(건)", fontsize=8.5)
ax.set_title("주문은 요일과 관계없이 저녁 18~20시에 몰림", fontsize=12, fontweight="bold", loc="center", pad=10)
fig.savefig(HERE / "W1-7_요일x시간대_주문분포.png")
print("W1-7 요일 합(영업일당)", G.sum(axis=1).round(1).to_dict(), "14~16시 비중", round(G[[14, 15]].values.sum() / G.values.sum() * 100, 1), "%")
