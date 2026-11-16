# -*- coding: utf-8 -*-
"""좌석 구분 없이 홀 전체 전표 × 배달 주문 관계(6/9~8/31): 요일·월 통제 잔차 상관, 18~23시 6월→8월 변화."""
import importlib.util, sys
import pandas as pd, statsmodels.formula.api as smf
from scipy import stats
sys.stdout.reconfigure(encoding="utf-8")
from pathlib import Path
P = Path(__file__).resolve().parent / "w1_분석.py"
spec = importlib.util.spec_from_file_location("w1", P); w1 = importlib.util.module_from_spec(spec); spec.loader.exec_module(w1)
t, hall = w1.load_receipts(); o = w1.load_orders()
s, e = "2026-06-09", "2026-08-31"
days = pd.DatetimeIndex(sorted(t[(t["영업일자"] >= s) & (t["영업일자"] <= e)]["영업일자"].unique()))
hi = hall[(hall["영업일자"] >= s) & (hall["영업일자"] <= e)].groupby("영업일자").size().reindex(days, fill_value=0)
dv = o[(o["날짜"] >= s) & (o["날짜"] <= e)].groupby("날짜").size().reindex(days, fill_value=0)
dd = pd.DataFrame({"홀": hi.values, "배달": dv.values, "요일": days.dayofweek, "월": days.month})
r1 = smf.ols("홀 ~ C(요일) + C(월)", dd).fit().resid; r2 = smf.ols("배달 ~ C(요일) + C(월)", dd).fit().resid
print("n", len(dd), "원값", stats.pearsonr(dd["홀"], dd["배달"]), "잔차", stats.pearsonr(r1, r2))
# 18~23시 첫 주문 기준 6월(9~30) vs 8월 영업일당
for lab, a, b in [("6월", "2026-06-09", "2026-06-30"), ("8월", "2026-08-01", "2026-08-31")]:
    dset = days[(days >= a) & (days <= b)]
    h = hall[(hall["영업일자"] >= a) & (hall["영업일자"] <= b) & hall["시"].between(18, 22)]
    oo = o[(o["날짜"] >= a) & (o["날짜"] <= b)]
    print(lab, "일수", len(dset), "홀 18~23시 전표/일", round(len(h) / len(dset), 2), "홀 전체/일", round(len(hall[(hall['영업일자'] >= a) & (hall['영업일자'] <= b)]) / len(dset), 2), "배달/일", round(len(oo) / len(dset), 2))
