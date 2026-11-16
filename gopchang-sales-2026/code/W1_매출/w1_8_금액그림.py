# -*- coding: utf-8 -*-
"""v6(2026-10-17): 그림 안 출처 삭제, 제목 가운데. v4(사장님 매출 공개 동의, 2026-10-16): 매출 그림을 금액(만 원)으로.
W1-8 월별 하루 매출(홀·직접 포장 + 배달앱, 3~8월), W1-6b 날짜별 하루 매출(5~8월). w1_분석.py 로더를 그대로 쓴다."""
import importlib.util
from pathlib import Path
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import numpy as np

HERE = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location("w1", HERE / "w1_분석.py")
w1 = importlib.util.module_from_spec(spec); spec.loader.exec_module(w1)

pos = w1.load_pos_daily().set_index("날짜")
o = w1.load_orders()
deliv = o.groupby("날짜")["음식판매"].sum()
d = pos[["POS"]].copy()
d["배달"] = deliv.reindex(d.index).fillna(0)
MAN = 1e4

# ── W1-8 월별 하루 매출(만 원)
months = [3, 4, 5, 6, 7, 8]
hall = [d.loc[d.index.month == m, "POS"].mean() / MAN for m in months]
app = [d.loc[d.index.month == m, "배달"].mean() / MAN for m in months]
# v8(10-17): 팀장 "디자인이 짜쳐 보임" → y축·테두리 없앤 미니멀 막대, 옅은 배달색, 5→8월 -30% 강조
LIGHT_DELIV = "#9DB8DE"
fig, ax = plt.subplots(figsize=(7.4, 3.7), dpi=300)
x = np.arange(len(months))
ax.bar(x, hall, 0.5, color=w1.NAVY, label="홀·직접 포장", zorder=2)
ax.bar(x, app, 0.5, bottom=hall, color=LIGHT_DELIV, label="배달앱", zorder=2)
for i, (h, a) in enumerate(zip(hall, app)):
    ax.text(i, h + a + 1.0, f"{h + a:.1f}", ha="center", va="bottom", fontsize=11.5, fontweight="bold", color=w1.TEXT)
    if a > 0:  # 배달이 있는 달만 칸별 값
        ax.text(i, h - 1.6, f"{h:.1f}", ha="center", va="top", fontsize=8.5, color="white")
        ax.text(i, h + a / 2, f"{a:.1f}", ha="center", va="center", fontsize=8, color=w1.NAVY)
# 5월 → 8월 감소 강조
ax.annotate("", xy=(5, hall[5] + app[5] + 5.2), xytext=(2.25, hall[2] + app[2] + 5.2),
            arrowprops=dict(arrowstyle="-|>", color="#0068B7", lw=1.3, connectionstyle="arc3,rad=-0.15"))
ax.text(3.6, hall[2] + app[2] + 7.6, "5월 대비 8월 하루 매출 -30%", ha="center", fontsize=9.5, color="#0068B7", fontweight="bold")
ax.set_xticks(x, [f"{m}월" for m in months], fontsize=10.5)
ax.set_ylim(0, 56); ax.set_yticks([]); ax.grid(False)
for s in ["top", "right", "left"]: ax.spines[s].set_visible(False)
ax.spines["bottom"].set_color("#C9D3E3"); ax.tick_params(axis="x", length=0, pad=6)
ax.legend(loc="upper left", ncol=2, fontsize=9, frameon=False, handlelength=1.0, handleheight=1.0, borderaxespad=0)
ax.text(1.0, 1.0, "단위: 만 원(영업일 하루 평균)", transform=ax.transAxes, ha="right", va="bottom", fontsize=8, color=w1.GRAY)
ax.set_title("6월부터 홀 매출이 줄고 배달이 일부를 메웠음", fontsize=12.5, fontweight="bold", loc="center", pad=22)
fig.tight_layout(); fig.savefig(HERE / "W1-8_월별_하루매출_금액.png", facecolor="white"); plt.close(fig)
print("W1-8", [round(h, 1) for h in hall], [round(a, 1) for a in app])

# ── W1-6b 날짜별 하루 매출(만 원), 5~8월
s = d.loc["2026-05-01":"2026-08-31"]
tot = (s["POS"] + s["배달"]) / MAN
roll = tot.rolling(7, min_periods=4, center=True).mean()
roll[tot.index.to_series().diff().dt.days.gt(3).values] = np.nan  # 사흘 넘게 쉰 구간은 평균선을 끊는다(v11)
fig, ax = plt.subplots(figsize=(8.4, 3.9))
ax.bar(tot.index, tot, width=0.8, color="#9AA3AE", alpha=0.45, label="하루 매출(홀·직접 포장+배달앱)")
ax.plot(roll.index, roll, color=w1.NAVY, lw=2.4, label="앞뒤 7일 평균")
for m in [5, 6, 7, 8]:
    mm = tot[tot.index.month == m]
    ax.hlines(mm.mean(), mm.index.min(), mm.index.max(), color=w1.GOLD, lw=2, ls="--")
    ax.text(mm.index.min(), mm.mean() + 2, f"{m}월 평균 {mm.mean():.1f}만 원", fontsize=9, color=w1.TEXT, fontweight="bold",
            bbox=dict(fc="white", ec="none", pad=1, alpha=0.85))
ax.set_ylabel("하루 매출(만 원)")
ax.xaxis.set_major_locator(mdates.MonthLocator()); ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda v, p: f"{mdates.num2date(v).month}월"))
for sp in ["top", "right"]: ax.spines[sp].set_visible(False)
ax.legend(loc="upper right", fontsize=8.5)
ax.set_title("6월부터 하루 매출 평균이 내려갔음", fontsize=12, fontweight="bold", loc="center", pad=10)
fig.savefig(HERE / "W1-6b_일별_매출_금액.png"); plt.close(fig)
print("W1-6b 최소·최대", round(tot.min(), 1), round(tot.max(), 1))
