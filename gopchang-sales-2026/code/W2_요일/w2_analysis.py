# -*- coding: utf-8 -*-
"""W2 요일·사장님 말씀 검증 — 일요일(H3)·2주 주기(H4)·가족(H5)·폐업(H6)·단골(H7) + 기타 말씀.

입력(읽기만): 볼트 02 POS 일자별, 예준 레포 19 배달·20 전표·리뷰 2종·업종4개_요일가설·8월까지_재계산.
출력(이 폴더): csv 여러 개, 그림 W2-1~3(PNG 300dpi). 금액은 지수·비율·건수만 저장한다.
실행: python w2_analysis.py
"""
import re, sys, io, textwrap, warnings
from pathlib import Path
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch, Rectangle

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
warnings.filterwarnings("ignore")

OUT = Path(__file__).resolve().parent
R = OUT.parents[2]  # 레포 최상위
POS_X = OUT.parent / "data" / "02_매출데이터_20260906.xlsx"
DLV_X = R / "매출 데이터" / "19_배달데이터_통합.xlsx"
RCP_X = R / "매출 데이터" / "20_테이블별매출_통합_0609-0918.xlsx"
REV_X = R / "리뷰" / "리뷰_전체.xlsx"
DRV_X = R / "리뷰" / "배달리뷰_전체.xlsx"
DOW_X = R / "분석" / "상권분석" / "output" / "업종4개_요일가설.xlsx"
AUG_X = R / "분석" / "최종보고서" / "output" / "8월까지_재계산.xlsx"

# ---- 팔레트 (발표자료 디자인 기준 2026-10-16). 여기만 바꾸면 그림 3장에 모두 반영된다.
PAL = {"navy": "#1B2A4A",   # 주색: 가게
       "blue": "#0068B7",   # 보조: 명일역 상권
       "gold": "#D4A843",   # 강조: 일요일·판단 기준
       "gray": "#8A8A8A",   # 서울 전체·되돌림
       "light": "#C9D3E3",  # 연한 남회색: 배경 띠·테두리
       "text": "#1F2937",   # 글자
       "grid": "#EEF1F5",   # 격자선(아주 연하게)
       "gold_bg": "#FBF3DC", "navy_bg": "#E9EEF5", "gray_bg": "#F1F1F1"}
NAVY, BLUE, GOLD, GRAY, LIGHT, TEXT, GRID = (PAL[k] for k in ["navy", "blue", "gold", "gray", "light", "text", "grid"])
from matplotlib import font_manager
# Noto Sans KR은 가변 글꼴(NotoSansKR-VF.ttf) 하나만 설치돼 있으면 matplotlib이 굵기 100(Thin)으로만 그린다 →
# 정적 Regular/Bold 파일이 있을 때만 쓰고, 없으면 Malgun Gothic.
_noto_static = [f for f in font_manager.fontManager.ttflist if f.name == "Noto Sans KR" and "VF" not in f.fname.upper() and f.weight >= 400]
FONT = "Noto Sans KR" if _noto_static else "Malgun Gothic"
plt.rcParams.update({"font.family": FONT, "axes.unicode_minus": False, "font.size": 10,
                     "figure.facecolor": "white", "axes.facecolor": "white",
                     "axes.edgecolor": LIGHT, "axes.labelcolor": TEXT, "xtick.color": TEXT, "ytick.color": TEXT, "text.color": TEXT})
print("font:", FONT)
WD = ["월", "화", "수", "목", "금", "토", "일"]
HOLIDAYS = {  # 2026 공휴일·대체공휴일 (분석 기간 안)
    "2026-03-01": "삼일절", "2026-03-02": "대체공휴일", "2026-05-01": "근로자의 날(참고)",
    "2026-05-05": "어린이날", "2026-05-24": "부처님오신날", "2026-05-25": "대체공휴일",
    "2026-06-03": "지방선거", "2026-06-06": "현충일", "2026-08-15": "광복절", "2026-08-17": "대체공휴일",
}
HOL = {pd.Timestamp(k): v for k, v in HOLIDAYS.items()}
def save(df, name):
    df.to_csv(OUT / name, index=False, encoding="utf-8-sig"); print("saved", name)
def idx(s): return s / s.mean() * 100

# ---------------------------------------------------------------- 0. 자료 적재
pos = pd.read_excel(POS_X, sheet_name="일자별매출", header=3)
pos["날짜"] = pd.to_datetime(pos["날짜"], errors="coerce"); pos = pos[(pos["영업여부"] == "영업") & pos["날짜"].notna()].copy()
pos["dow"] = pos["날짜"].dt.dayofweek; pos["월"] = pos["날짜"].dt.month
pos["공휴일"] = pos["날짜"].map(HOL).fillna("")
assert len(pos) == 165, len(pos)

daily = pd.read_excel(DLV_X, sheet_name="일별_채널"); daily["날짜"] = pd.to_datetime(daily["날짜"])
orders = pd.read_excel(DLV_X, sheet_name="주문"); orders["날짜"] = pd.to_datetime(orders["날짜"])
orders = orders[orders["상태"] != "취소"].copy()
for c in ["음식금액", "가게즉시할인", "고객결제금액"]:
    orders[c] = pd.to_numeric(orders[c], errors="coerce")
# 예준 최종보고서와 같은 정의: 음식금액−가게할인, 미확보는 결제금액−가게배달팁 3,000
known = orders["음식금액"] - orders["가게즉시할인"].fillna(0)
est = orders["고객결제금액"] - np.where(orders["서비스유형"] == "가게배달", 3000, 0)
orders["앱음식"] = known.where(orders["음식금액"].notna(), est)
app = orders.groupby("날짜").agg(앱음식=("앱음식", "sum"), 앱건수=("채널", "size"))
d = daily.set_index("날짜").join(app).fillna({"앱음식": 0, "앱건수": 0})
assert (d["앱건수"] == d["배달앱_건수"]).all()
d = d[(d.index <= "2026-08-31") & (d["POS영업여부"] == "영업")].copy()
d["합산"] = d["POS총매출"] + d["앱음식"]; d["dow"] = d.index.dayofweek; d["월"] = d.index.month
assert len(d) == 100

rc = pd.read_excel(RCP_X, sheet_name="전표요약"); rc["영업일자"] = pd.to_datetime(rc["영업일자"])
hall = rc[(rc["채널"] == "매장") & (rc["영업일자"] <= "2026-08-31")].copy()
def hr(t):
    h = int(str(t)[:2]); return h + 24 if h < 6 else h
hall["시"] = hall["최초주문"].map(hr); hall["결제시"] = hall["결제시간"].map(lambda t: int(str(t)[:2]) + (24 if int(str(t)[:2]) < 6 else 0))
hall["dow"] = hall["영업일자"].dt.dayofweek
hall["tno"] = hall["테이블번호"].astype(str).str.extract(r"(\d+)")[0].astype(float)  # '테이블 6' 같은 표기 섞임
assert hall["tno"].notna().all() and hall["tno"].between(1, 9).all(), hall["테이블번호"].unique()

# ---------------------------------------------------------------- 1. H3 일요일 — 가게 자체 근거
# 1-1 3월 요일 지수 (개업 첫 달, 정기휴무 전이라 일요일 5회 영업)
m3 = pos[pos["월"] == 3].copy()
base_mon_sat = m3[m3.dow < 6]["총매출액"].mean()
rows = []
for k, w in enumerate(WD):
    g = m3[m3.dow == k]
    g_nh = g[g["공휴일"] == ""]
    rows.append({"요일": w, "영업일수": len(g), "공휴일 포함": ", ".join(f"{x.strftime('%m/%d')} {v}" for x, v in zip(g["날짜"], g["공휴일"]) if v),
                 "지수(3월 월~토 평균=100)": round(g["총매출액"].mean() / base_mon_sat * 100, 1),
                 "지수_공휴일 제외": round(g_nh["총매출액"].mean() / base_mon_sat * 100, 1) if len(g_nh) else np.nan,
                 "최고일 지수": round(g["총매출액"].max() / base_mon_sat * 100, 1), "최저일 지수": round(g["총매출액"].min() / base_mon_sat * 100, 1),
                 "순위(평균)": None})
t_m3 = pd.DataFrame(rows); t_m3["순위(평균)"] = t_m3["지수(3월 월~토 평균=100)"].rank(ascending=False).astype(int)
# 일요일 3/8 돌출일 제외 민감도
sun = m3[m3.dow == 6]; sun_ex38 = sun[sun["날짜"] != "2026-03-08"]
t_m3.loc[t_m3["요일"] == "일", "지수_3/8 제외"] = round(sun_ex38["총매출액"].mean() / base_mon_sat * 100, 1)
t_m3.loc[t_m3["요일"] == "일", "지수_3/8 제외"] = t_m3.loc[t_m3["요일"] == "일", "지수_3/8 제외"]
t_m3["기준"] = "POS 총매출(홀+직접포장, 배달앱 제외), 2026-03-01~03-31 영업일 30일"
save(t_m3, "W2_3월_요일지수.csv")
sun_idx = float(t_m3.loc[t_m3["요일"] == "일", "지수(3월 월~토 평균=100)"].iloc[0])
tue_idx3 = float(t_m3.loc[t_m3["요일"] == "화", "지수(3월 월~토 평균=100)"].iloc[0])
print("3월 일요일 지수", sun_idx, "화요일", tue_idx3, "일/화", round(sun_idx / tue_idx3, 2))

# 1-2 정기휴무 뒤 일요일 영업 2회 (5/24·8/16, 모두 연휴) — 홀+배달 합산, 같은 달 화요일 대비
rows = []
for sd in [pd.Timestamp("2026-05-24"), pd.Timestamp("2026-08-16")]:
    m = d[d["월"] == sd.month]; tue = m[m.dow == 1]; tue_nh = tue[~tue.index.isin(HOL)]
    ms = m[m.dow < 6]
    for lab, ref in [("같은 달 화요일 전체", tue), ("같은 달 화요일(공휴일 제외)", tue_nh), ("같은 달 월~토 평균", ms)]:
        rows.append({"일요일": sd.strftime("%Y-%m-%d"), "비고": HOL.get(sd, ""), "비교 기준": lab, "기준 영업일수": len(ref),
                     "홀(POS) 지수": round(d.loc[sd, "POS총매출"] / ref["POS총매출"].mean() * 100, 1),
                     "배달 지수": round(d.loc[sd, "앱음식"] / ref["앱음식"].mean() * 100, 1) if ref["앱음식"].mean() > 0 else np.nan,
                     "합산 지수": round(d.loc[sd, "합산"] / ref["합산"].mean() * 100, 1),
                     "일요일 배달 건수": int(d.loc[sd, "앱건수"]), "기준일 배달 건수 평균": round(ref["앱건수"].mean(), 2)})
t_sun2 = pd.DataFrame(rows); save(t_sun2, "W2_일요일영업2회_비교.csv")
print(t_sun2[["일요일", "비교 기준", "합산 지수", "홀(POS) 지수"]].to_string())

# ---------------------------------------------------------------- 2. 상권 근거 재확인 (원자료 xlsx)
raw = pd.read_excel(DOW_X, sheet_name="원자료_요일")
raw["일÷화_재계산"] = (raw["일_비중%"] / raw["화_비중%"]).round(3)
raw["차이(시트값−재계산)"] = (raw["일÷화_금액"] - raw["일÷화_재계산"]).round(3)
chk = raw[["묶음", "업종", "점포수_2026_2Q", "화_비중%", "일_비중%", "일÷화_금액", "일÷화_재계산", "차이(시트값−재계산)"]]
qtr = pd.read_excel(DOW_X, sheet_name="원자료_분기별")
q_my = qtr[(qtr["묶음"] == "명일역 상권") & (qtr["업종"].isin(["한식", "호프-간이주점", "한식+호프"]))]
q_sum = q_my.groupby("업종")["일÷화_금액"].agg(**{"분기수": "size", "최소": "min", "최대": "max", "중앙": "median", "일÷화 1초과 분기수": lambda s: int((s > 1).sum())}).round(3).reset_index()
pop = pd.read_excel(DOW_X, sheet_name="생활인구_요일저녁")
pop_sel = pop[pop["시간"].isin(["17~21시", "17~24시", "21~24시"])][["영역", "시간", "인구", "화", "일", "일÷화"]]
chk.to_csv(OUT / "W2_상권_일÷화_재확인.csv", index=False, encoding="utf-8-sig")
with open(OUT / "W2_상권_일÷화_재확인.csv", "a", encoding="utf-8-sig") as f:
    f.write("\n[명일역 상권 10개 분기(2024 1Q~2026 2Q) 일÷화 분포]\n"); q_sum.to_csv(f, index=False)
    f.write("\n[도보권 저녁 생활인구 일÷화 (2026-09-01~21, 요일당 3일)]\n"); pop_sel.to_csv(f, index=False)
print("saved W2_상권_일÷화_재확인.csv"); print(chk[chk["업종"] == "한식+호프"].to_string()); print(q_sum.to_string())

# 2-2 요일 지수 비교표: 가게 홀(전표 6/9~8/31) vs 명일역 한식+호프 vs 서울 한식+호프 (월~토 평균=100)
hd = hall.groupby("영업일자")["총금액"].sum().to_frame("홀매출"); hd["dow"] = hd.index.dayofweek
hd["공휴일"] = hd.index.map(HOL).fillna("")
wk = hd.groupby("dow")["홀매출"].agg(["mean", "size"])
store = wk["mean"] / wk.loc[:5, "mean"].mean() * 100
aug = pd.read_excel(AUG_X, sheet_name="요일", header=1).dropna(subset=["요일"])
aug = aug[aug["요일"].isin(WD)].set_index("요일")
aug = aug[~aug.index.duplicated(keep="first")]  # 시트 아래쪽 참고 블록(요일 중복) 제외
aug_idx_col = [c for c in aug.columns if "지수" in str(c)][0]
def mkt(group):
    r = raw[(raw["묶음"] == group) & (raw["업종"] == "한식+호프")].iloc[0]
    v = np.array([r[f"{w}_비중%"] for w in WD], float); return v / v[:6].mean() * 100
my, se = mkt("명일역 상권"), mkt("서울 전체 상권")
t_dow = pd.DataFrame({"요일": WD,
    "가게 홀 지수(전표 6/9~8/31, 월~토 평균=100)": [round(store.get(k, np.nan), 1) if k < 6 else np.nan for k in range(7)],
    "가게 영업일수": [int(wk.loc[k, "size"]) if k in wk.index else 0 for k in range(7)],
    "예준 8월까지_재계산 홀 지수(평균=100)": [float(aug.loc[w, aug_idx_col]) if w in aug.index else np.nan for w in WD],
    "명일역 한식+호프 지수": my.round(1), "서울 한식+호프 지수": se.round(1)})
t_dow["명일역 비중 출처"] = "서울시 상권분석서비스 추정매출-상권(OA-15572) 2025 3Q~2026 2Q 합산, 업종4개_요일가설.xlsx 원자료_요일"
save(t_dow, "W2_요일지수_가게vs상권.csv"); print(t_dow.iloc[:, :6].to_string())
store_tue = float(store[1]); store_wed = float(store[2])

# ---------------------------------------------------------------- 3. H4 2주 주기 재검정 (홀·배달·합산)
def weekly(df, cols):
    df = df.copy(); df["주"] = df.index - pd.to_timedelta(df["dow"], unit="D")
    g = df.groupby("주").agg(영업일=("dow", "size"), **{c: (c, "mean") for c in cols})
    return g[g["영업일"] >= 3]
def acf(x, k): x = np.asarray(x) - np.mean(x); return float(np.sum(x[k:] * x[:-k]) / np.sum(x * x))
def detrend(y, deg):
    t = np.arange(len(y)); return y - np.polyval(np.polyfit(t, y, deg), t)
def alt_tests(y, label, deg=2):
    y = np.asarray(y, float); r = detrend(idx(pd.Series(y)).values, deg); n = len(r)
    s = np.sign(r); alt = float(np.mean(s[1:] != s[:-1]))
    # 격주 더미 (-1)^t 회귀: 교대가 있으면 계수가 뚜렷하고 p가 작다
    z = np.where(np.arange(n) % 2 == 0, 1.0, -1.0); b = float(np.sum(z * r) / np.sum(z * z))
    resid = r - b * z; se_b = float(np.sqrt(np.sum(resid ** 2) / (n - 1) / np.sum(z * z))); tstat = b / se_b
    from scipy import stats
    p = float(2 * stats.t.sf(abs(tstat), n - 1))
    return {"계열": label, "주 수": n, "추세 제거": f"{deg}차 다항", "lag1 자기상관": round(acf(r, 1), 3), "lag2 자기상관": round(acf(r, 2), 3),
            "부호 교대 비율": round(alt, 2), "격주 더미 계수(지수점)": round(b, 1), "격주 더미 p값": round(p, 3),
            "무작위 기준(±2/√n)": round(2 / np.sqrt(n), 2), "교대 패턴이면": "lag1 뚜렷한 음(−), lag2 양(+), 교대 비율 1에 가까움, p<0.05"}
# 3-1 POS 홀+직접포장 전기간 28주 (17 노트 재현)
pos_w = weekly(pos.set_index("날짜"), ["총매출액"])
# 3-2 5/1~8/31 홀·배달·합산·앱비중
d["앱비중"] = d["앱음식"] / d["합산"] * 100
dw = weekly(d, ["POS총매출", "앱음식", "합산", "앱비중"])
res = [alt_tests(pos_w["총매출액"], "홀(POS) 2/24~9/5"), alt_tests(dw["POS총매출"], "홀(POS) 5/1~8/31"),
       alt_tests(dw["앱음식"], "배달 5/1~8/31"), alt_tests(dw["합산"], "홀+배달 합산 5/1~8/31"), alt_tests(dw["앱비중"], "배달 비중(%) 5/1~8/31", deg=1)]
for deg in [1]:
    res.append(alt_tests(dw["POS총매출"], "홀(POS) 5/1~8/31 (1차 추세)", deg=1)); res.append(alt_tests(dw["앱음식"], "배달 5/1~8/31 (1차 추세)", deg=1))
t_cyc = pd.DataFrame(res)
rh, ra = detrend(idx(dw["POS총매출"]).values, 2), detrend(idx(dw["앱음식"]).values, 2)
xcorr = float(np.corrcoef(rh, ra)[0, 1])
xcorr_diff = float(np.corrcoef(np.diff(idx(dw["POS총매출"]).values), np.diff(idx(dw["앱음식"]).values))[0, 1])
extra = pd.DataFrame([{"계열": "홀 잔차 × 배달 잔차 동시 상관(5/1~8/31)", "주 수": len(dw), "lag1 자기상관": round(xcorr, 3), "교대 패턴이면": "홀이 많은 주에 배달이 적다면 뚜렷한 음(−)"},
                      {"계열": "홀 주간 변화 × 배달 주간 변화 상관", "주 수": len(dw) - 1, "lag1 자기상관": round(xcorr_diff, 3), "교대 패턴이면": "뚜렷한 음(−)"}])
t_cyc = pd.concat([t_cyc, extra], ignore_index=True); save(t_cyc, "W2_2주주기_검정.csv"); print(t_cyc.iloc[:, :8].to_string())
wk_out = dw.copy(); wk_out.index = wk_out.index.strftime("%m-%d(월)")
for c in ["POS총매출", "앱음식", "합산"]: wk_out[c + "_지수"] = idx(wk_out[c]).round(1)
wk_out = wk_out.drop(columns=["POS총매출", "앱음식", "합산"]).rename(columns={"앱비중": "배달비중_%"}).round(1)
wk_out.index.name = "주(월요일 시작)"; wk_out.reset_index().to_csv(OUT / "W2_주별시계열_5~8월.csv", index=False, encoding="utf-8-sig")

# ---------------------------------------------------------------- 4. H5 가족 손님 — 전표·리뷰
n = len(hall)
qty = hall["안주수량"].clip(upper=4)
dist = qty.value_counts().sort_index()
fam = hall[(hall["안주수량"] >= 3) & (hall["주류부착"] == "N")]
fam_eve = fam[(fam["시"] >= 17) & (fam["시"] < 21)]
noalc = hall[hall["주류부착"] == "N"]
t_fam = pd.DataFrame([
    {"지표": "홀 전표 수(6/9~8/31)", "값": n},
    {"지표": "안주 1인분 비율(%)", "값": round(dist.get(1, 0) / n * 100, 1)}, {"지표": "안주 2인분 비율(%)", "값": round(dist.get(2, 0) / n * 100, 1)},
    {"지표": "안주 3인분 비율(%)", "값": round(dist.get(3, 0) / n * 100, 1)}, {"지표": "안주 4인분 이상 비율(%)", "값": round(dist.get(4, 0) / n * 100, 1)},
    {"지표": "안주 2인분 이하 비율(%)", "값": round((qty <= 2).mean() * 100, 1)},
    {"지표": "술 동반 전표 비율(%)", "값": round((hall["주류부착"] == "Y").mean() * 100, 1)},
    {"지표": "첫 주문 17~20시 비율(%)", "값": round(((hall["시"] >= 17) & (hall["시"] < 21)).mean() * 100, 1)},
    {"지표": "첫 주문 21시 이후 비율(%)", "값": round((hall["시"] >= 21).mean() * 100, 1)},
    {"지표": "가족형 후보(안주 3인분 이상·술 없음) 전표 수", "값": len(fam)}, {"지표": "가족형 후보 비율(%)", "값": round(len(fam) / n * 100, 1)},
    {"지표": "가족형 후보 중 17~20시 전표 수", "값": len(fam_eve)},
    {"지표": "술 없는 전표의 안주 수량 평균", "값": round(noalc["안주수량"].mean(), 2)}, {"지표": "술 있는 전표의 안주 수량 평균", "값": round(hall[hall["주류부착"] == "Y"]["안주수량"].mean(), 2)},
    {"지표": "토요일 홀 전표 중 가족형 후보 비율(%)", "값": round(len(fam[fam.dow == 5]) / max(len(hall[hall.dow == 5]), 1) * 100, 1)},
    {"지표": "토요일 홀 전표 중 17~20시 술 없는 비율(%)", "값": round(len(noalc[(noalc.dow == 5) & (noalc["시"] >= 17) & (noalc["시"] < 21)]) / max(len(hall[hall.dow == 5]), 1) * 100, 1)},
])
# 시간대별 술 동반·안주 수량
tb = hall.assign(시간대=pd.cut(hall["시"], [0, 17, 19, 21, 23, 30], labels=["~16시", "17~18시", "19~20시", "21~22시", "23시~"], right=False)) \
    .groupby("시간대").agg(전표수=("시", "size"), 술동반비율=("주류부착", lambda s: round((s == "Y").mean() * 100, 1)), 안주수량평균=("안주수량", "mean"),
                        가족형후보=("안주수량", lambda s: int(((s >= 3) & (hall.loc[s.index, "주류부착"] == "N")).sum()))).round(2).reset_index()
rev = pd.read_excel(REV_X, sheet_name="네이버_리뷰"); rev["작성일"] = pd.to_datetime(rev["작성일"])
comp = rev["동행"].dropna().str.split(" / ").explode().str.strip().value_counts()
purpose = rev["방문 목적"].dropna().str.split(" / ").explode().str.strip().value_counts()
drv = {s: pd.read_excel(DRV_X, sheet_name=s) for s in ["땡겨요", "배민", "쿠팡이츠"]}
fam_re = re.compile(r"아이|애들|애기|아기|아들|딸|가족|식구|남편|아내|와이프|신랑|엄마|아빠|부모님|어머니|아버지")
fam_mention = sum(int(df["본문"].fillna("").str.contains(fam_re).sum()) for df in drv.values())
n_drv = sum(len(df) for df in drv.values())
t_rev = pd.DataFrame([{"자료": "네이버 리뷰 동행", "항목": k, "건수": int(v)} for k, v in comp.items()] +
                     [{"자료": "네이버 리뷰 방문 목적", "항목": k, "건수": int(v)} for k, v in purpose.items()] +
                     [{"자료": "배달 리뷰 본문", "항목": "가족·아이·배우자 언급(키워드)", "건수": fam_mention}, {"자료": "배달 리뷰 본문", "항목": "리뷰 전체", "건수": n_drv}])
with open(OUT / "W2_가족손님_전표리뷰.csv", "w", encoding="utf-8-sig") as f:
    f.write("[홀 전표 6/9~8/31]\n"); t_fam.to_csv(f, index=False); f.write("\n[시간대별 홀 전표]\n"); tb.to_csv(f, index=False)
    f.write("\n[리뷰 동행·목적 (네이버 48건, 중복 선택 가능) · 배달 리뷰 124건]\n"); t_rev.to_csv(f, index=False)
print("saved W2_가족손님_전표리뷰.csv"); print(t_fam.to_string()); print(tb.to_string()); print(comp.to_dict(), purpose.to_dict(), "배달 가족 언급", fam_mention)

# ---------------------------------------------------------------- 5. H7 단골·새 손님 — 리뷰 추이·재방문·재주문
rev["월"] = rev["작성일"].dt.to_period("M")
mon = rev.groupby("월").size()
first = int((rev["방문 차수"] == 1).sum()); rev_n = len(rev)
coup = drv["쿠팡이츠"]; bm = drv["배민"]; dd = drv["땡겨요"]
rean = pd.read_excel(DRV_X, sheet_name="리뷰별_분석")
re_col = [c for c in rean.columns if "재주문" in str(c)][0]
re_cnt = int(rean[re_col].notna().sum())
dd_re = int(dd["본문"].fillna("").str.contains(r"재주문|또 시|또시|단골|번째|자주|매번|늘 |항상|다시 시").sum())
t_loyal = pd.DataFrame([
    *[{"지표": f"네이버 리뷰 작성 수 {m}", "값": int(v)} for m, v in mon.items()],
    {"지표": "네이버 리뷰 3~5월 합", "값": int(mon[[p for p in mon.index if p.month in (3, 4, 5)]].sum())},
    {"지표": "네이버 리뷰 6~8월 합", "값": int(mon[[p for p in mon.index if p.month in (6, 7, 8)]].sum())},
    {"지표": "네이버 리뷰 첫 방문 비율(%)", "값": round(first / rev_n * 100, 1)}, {"지표": "네이버 리뷰 2회 이상 방문 건수", "값": rev_n - first},
    {"지표": "네이버 리뷰 작성자 수 / 리뷰 수", "값": f"{rev['작성자 닉네임'].nunique()} / {rev_n}"},
    {"지표": "쿠팡이츠 리뷰 중 'N번째 재주문' 표시 비율(%)", "값": round(coup["재주문"].notna().mean() * 100, 1)},
    {"지표": "쿠팡이츠 리뷰 수", "값": len(coup)},
    {"지표": "배민 리뷰 중 '최근 N번 주문' 표시 건수", "값": int(bm["최근 주문"].notna().sum())}, {"지표": "배민 리뷰 수", "값": len(bm)},
    {"지표": "땡겨요 본문에 재주문·단골 표현 건수(키워드)", "값": dd_re}, {"지표": "땡겨요 리뷰 수", "값": len(dd)},
    {"지표": f"배달 리뷰 124건 중 재주문·단골 언급('{re_col}' 분류)", "값": re_cnt}, {"지표": "배달 리뷰 재주문·단골 비율(%)", "값": round(re_cnt / len(rean) * 100, 1)},
    {"지표": "땡겨요 월별 리뷰 5/6/7/8/9월", "값": "/".join(str(int(v)) for v in pd.to_datetime(dd["작성일"]).dt.month.value_counts().reindex([5, 6, 7, 8, 9]).fillna(0))},
])
save(t_loyal, "W2_단골_새손님_리뷰.csv"); print(t_loyal.to_string())

# ---------------------------------------------------------------- 6. 기타 말씀 — 회전·웨이팅·배달 비중·객단가
daily_hall = hall.groupby("영업일자").size()
def max_concurrent(g):
    ev = sorted([(r["시"] * 60 + int(str(r["최초주문"])[3:5]), 1) for _, r in g.iterrows()] + [(r["결제시"] * 60 + int(str(r["결제시간"])[3:5]), -1) for _, r in g.iterrows()], key=lambda x: (x[0], x[1]))
    cur = best = 0
    for _, e in ev: cur += e; best = max(best, cur)
    return best
conc = hall.groupby("영업일자").apply(max_concurrent)
c1 = pd.read_csv(R / "분석" / "최종보고서" / "output" / "tables" / "C1_월별_채널구성_음식판매기준.csv").iloc[:4]
t_misc = pd.DataFrame([
    {"말씀": "하루 2회전", "지표": "홀 전표/영업일 평균(6/9~8/31)", "값": round(daily_hall.mean(), 2)},
    {"말씀": "하루 2회전", "지표": "테이블 9개 기준 회전(평균 / 최대)", "값": f"{daily_hall.mean() / 9:.2f} / {daily_hall.max() / 9:.2f}"},
    {"말씀": "하루 2회전", "지표": "실내 5개 기준 회전(평균 / 최대, 실내 전표만)", "값": f"{hall[hall['tno'] <= 5].groupby('영업일자').size().mean() / 5:.2f} / {hall[hall['tno'] <= 5].groupby('영업일자').size().max() / 5:.2f}"},
    {"말씀": "하루 2회전", "지표": "2회전(전표 18장 이상) 넘긴 날 수 / 영업일", "값": f"{int((daily_hall >= 18).sum())} / {len(daily_hall)}"},
    {"말씀": "웨이팅 없음", "지표": "하루 최대 동시 착석 테이블(중앙 / 최대, 전체 9개)", "값": f"{int(conc.median())} / {int(conc.max())}"},
    {"말씀": "웨이팅 없음", "지표": "동시 착석 7개 이상 날 수", "값": int((conc >= 7).sum())},
    {"말씀": "배달 30%", "지표": "배달앱 매출 비중(%) 5/6/7/8월 (음식 판매액 기준)", "값": "/".join(str(x) for x in c1["앱비중_기준(추정 포함)(%)"])},
    {"말씀": "배달 30%", "지표": "5~8월 합산 배달 비중(%)", "값": round(d["앱음식"].sum() / d["합산"].sum() * 100, 1)},
    {"말씀": "객단가 2만원대", "지표": "홀 안주 1인분당 매출 ÷ 메뉴판 야채곱창 가격(12,000원) 배수", "값": round(hall["총금액"].sum() / hall["안주수량"].sum() / 12000, 2)},
    {"말씀": "객단가 2만원대", "지표": "홀 안주 1인분당 매출(만원 단위 반올림)", "값": f"약 {hall['총금액'].sum() / hall['안주수량'].sum() / 10000:.1f}만원"},
    {"말씀": "객단가 2만원대", "지표": "홀 전표당 안주 수량 평균", "값": round(hall["안주수량"].mean(), 2)},
])
save(t_misc, "W2_기타말씀_회전웨이팅배달객단가.csv"); print(t_misc.to_string())

# ---------------------------------------------------------------- 7. 그림 W2-1 요일 지수 비교
fig, ax = plt.subplots(figsize=(7.6, 4.5), dpi=300)  # v9: 인쇄 폭에 맞춘 캔버스(글씨가 줄지 않게)
x = np.arange(7); w = 0.26
series = [("꿀꿀곱창 홀 (전표 6/9~8/31)", t_dow.iloc[:, 1].values, NAVY), ("명일역 상권 한식+호프", my, BLUE), ("서울 전체 한식+호프", se, GRAY)]
# 일요일·화요일 묶음 뒤에 연한 띠(강조 요일)
for k in (1, 6): ax.add_patch(Rectangle((x[k] - 0.47, 0), 0.94, 165, fc="#EEF1F5", ec="none", zorder=0))
for i, (lab, v, c) in enumerate(series):
    xs = x + (i - 1) * w
    vals = np.nan_to_num(v, nan=0)
    ax.bar(xs, vals, w * 0.92, color=c, label=lab, zorder=3, linewidth=0)
    for k in (1, 6):
        if not np.isnan(v[k]): ax.text(xs[k], v[k] + 2, f"{v[k]:.0f}", ha="center", va="bottom", fontsize=8.5, color=TEXT)
# 가게 일요일: 자료 없음
xs0 = x[6] - w
ax.add_patch(Rectangle((xs0 - w * 0.46, 0), w * 0.92, 100, fill=False, edgecolor=NAVY, linestyle=(0, (3, 2)), linewidth=1.2, zorder=3))
ax.text(xs0, 50, "자료\n없음", ha="center", va="center", fontsize=7.5, color=NAVY, linespacing=1.3)
ax.axhline(100, color=GRAY, linewidth=0.7, zorder=2); ax.text(2.0, 101.5, "월~토 평균 = 100", fontsize=8, color=GRAY, ha="center", va="bottom")
ax.set_xticks(x); ax.set_xticklabels([f"{w_}요일" for w_ in WD]); ax.set_ylabel("요일 지수 (월~토 평균 = 100)")
ax.set_ylim(0, 165); ax.grid(axis="y", color=GRID, linewidth=0.6, zorder=0); ax.set_axisbelow(True)
for s in ["top", "right"]: ax.spines[s].set_visible(False)
ax.legend(loc="lower left", bbox_to_anchor=(0, 1.0), frameon=False, fontsize=9, ncol=3, borderaxespad=0.2)
ax.set_title("명일역 상권은 일요일 매출이 화요일보다 큼", fontsize=12.5, color=TEXT, loc="center", pad=28, fontweight="bold")
ax.annotate(f"화요일: 가게 {t_dow.iloc[1, 1]:.0f} · 명일역 {my[1]:.0f}\n일요일: 명일역 {my[6]:.0f} (일÷화 1.44), 서울 {se[6]:.0f} (일÷화 {se[6] / se[1]:.2f})",
            xy=(0.02, 0.97), xycoords="axes fraction", ha="left", va="top", fontsize=8.5, color=TEXT,
            bbox=dict(boxstyle="round,pad=0.4", fc="white", ec=LIGHT))
fig.tight_layout(); fig.savefig(OUT / "W2-1_요일지수_가게vs명일역vs서울.png"); plt.close(fig); print("saved W2-1")

# ---------------------------------------------------------------- 8. 그림 W2-2 판단 흐름도
fig, ax = plt.subplots(figsize=(11, 6.8), dpi=300); ax.set_xlim(0, 100); ax.set_ylim(0, 100); ax.axis("off")
def box(xc, yc, wd, ht, text, fc="white", ec=TEXT, lw=1.2, fs=9.5, tc=TEXT, bold=False):
    ax.add_patch(FancyBboxPatch((xc - wd / 2, yc - ht / 2), wd, ht, boxstyle="round,pad=0.4,rounding_size=1.2", fc=fc, ec=ec, lw=lw, zorder=3))
    ax.text(xc, yc, text, ha="center", va="center", fontsize=fs, color=tc, zorder=4, fontweight="bold" if bold else "normal", linespacing=1.4)
def arrow(p, q, text=None, tx=0, ty=0, color=TEXT):
    ax.add_patch(FancyArrowPatch(p, q, arrowstyle="-|>", mutation_scale=14, color=color, lw=1.1, zorder=2, shrinkA=2, shrinkB=2))
    if text: ax.text((p[0] + q[0]) / 2 + tx, (p[1] + q[1]) / 2 + ty, text, fontsize=8.3, color=TEXT, ha="center", va="center", bbox=dict(fc="white", ec="none", pad=1))
box(50, 91, 70, 9, "시작: 평일 하루(화 또는 수) 휴무·일요일 영업으로 4주(일요일 4번) 시범, 근무일 수는 그대로\n시작 전에 네이버·카카오·배달앱 영업일을 같은 날 바꾸고, 직전 4주 쉬는 평일 기록을 적어 둔다", fc=PAL["navy_bg"], ec=NAVY, fs=9)
box(50, 73, 60, 9, "매주 기록: 일요일 홀·배달 매출, 그 주 문 연 평일 평균, 날씨·공휴일,\n피로도(1 여유 ~ 5 다음 날이 걱정될 만큼 지침)", ec=LIGHT, fs=9)
box(50, 55, 60, 9, "4주 뒤 계산: 일요일 4일 평균 ÷ 직전 4주 쉬는 평일 평균 (홀+배달 합산 매출)\n= '본전 비율'. 공휴일·연휴가 낀 주는 빼고 한 주를 더한다", fc=PAL["gold_bg"], ec=GOLD, fs=9, bold=True)
arrow((50, 86.5), (50, 77.5)); arrow((50, 68.5), (50, 59.5))
box(17, 28, 28, 15, "계속 (정식 전환)\n일요일 영업·평일 휴무 확정\n지도·앱·안내문 표기 유지\n다음 점검: 겨울 첫 달", fc=PAL["navy_bg"], ec=NAVY, fs=9, bold=False)
box(50, 28, 28, 15, "4주 연장\n단골이 일요일 영업을 아는 데\n시간이 걸린다 → 소식·당근 1회 더\n8주째에 같은 기준으로 다시 판정", fc=PAL["gold_bg"], ec=GOLD, fs=9)
box(83, 28, 28, 15, "되돌림 (일요일 휴무)\n쉬던 평일 영업 복귀, 표기 원상 복구\n'일요일은 손님이 적다'는 말씀이\n이 가게에선 맞았다고 기록", fc=PAL["gray_bg"], ec=GRAY, fs=9)
arrow((35, 50.5), (17, 36), "본전 비율 1.0 이상\n그리고 4주 내내 피로도 4 이하", tx=-10, ty=3, color=NAVY)
arrow((50, 50.5), (50, 36), "0.8 ~ 1.0", tx=0, ty=1, color=GOLD)
arrow((65, 50.5), (83, 36), "0.8 미만 또는\n피로도 5가 2주 연속", tx=10, ty=3, color=GRAY)
ax.text(50, 9, "본전 조건: 일요일 하루 매출(홀+배달)이 쉬는 평일의 직전 4주 평균 이상이면 손해가 아니다.\n"
        "상권 자료(명일역 한식+호프 일÷화 1.44)대로라면 1.0을 넘고, 가게 3월 경험(일요일 81.6 ÷ 화요일 97.3)대로라면 못 넘는다.\n"
        "0.8 미만은 상권 자료가 이 가게에는 맞지 않는다고 보는 선이다. 4주는 방향 확인용이라 경계값 근처면 연장한다. 계절이 바뀌어 평일도 20% 넘게 줄면 기준값도 그만큼 낮춘다.",
        ha="center", va="center", fontsize=8.3, color=TEXT, linespacing=1.5)
ax.set_title("일요일 영업 4주 시범은 '일요일 평균이 직전 4주 쉬는 평일 평균을 넘는가'로 계속·연장·되돌림을 정한다", fontsize=12.5, color=TEXT, loc="center", pad=8, fontweight="bold")
fig.tight_layout(); fig.savefig(OUT / "W2-2_일요일시범_판단흐름도.png"); plt.close(fig); print("saved W2-2")

# ---------------------------------------------------------------- 9. 진술 확인 카드 표 + 그림 W2-3
sun2 = t_sun2.set_index(["일요일", "비교 기준"])["합산 지수"]
cards = [
    ("H1", "여름이라 손님이 줄었다", "좌석 구역 × 저녁 기온, 상권 계절지수", "W1 결과 반영", "W1(매출·좌석·날씨) 판정을 그대로 옮긴다"),
    ("H2", "줄어든 만큼 배달로 메웠다", "홀+배달 합산 지수(C1)", "W1 결과 반영", "W1 판정을 그대로 옮긴다"),
    ("H3", "일요일은 장사가 안 된다", "3월 일요일 5일 POS, 연휴 일요일 2일(홀+배달), 상권 요일·생활인구·경쟁점", "부분적으로 맞았다",
     f"개업 첫 달엔 일요일이 요일 꼴찌(지수 {sun_idx:.0f}, 화요일 {tue_idx3:.0f}). 그러나 상권 자료는 일요일이 화요일의 1.44배, 도보 10분 안 곱창집은 일요일에 모두 쉰다 → 4주 시범으로 확인"),
    ("H4", "홀과 배달이 2주 주기로 번갈아 많아진다", f"주별 홀·배달·합산 시계열(5~8월 {len(dw)}주, 홀은 2~9월 {len(pos_w)}주)", "다르게 나타났다",
     f"추세를 걷어낸 주별 값에 격주 교대 신호가 없다(홀 lag1 {t_cyc.iloc[1]['lag1 자기상관']:+.2f}, 배달 {t_cyc.iloc[2]['lag1 자기상관']:+.2f}, 홀·배달 잔차 상관 {xcorr:+.2f}, 교대 비율 0.5 안팎)"),
    ("H5", "주 손님은 가족이다", "홀 전표 421장(안주 수량·술·시간), 네이버 리뷰 동행·목적, 배달 리뷰 본문", "다르게 나타났다",
     f"안주 2인분 이하 {(qty <= 2).mean() * 100:.0f}%, 술 동반 {(hall['주류부착'] == 'Y').mean() * 100:.0f}%. 가족형(3인분 이상·술 없음) 전표는 {len(fam) / n * 100:.1f}%. 리뷰 동행은 연인·친구·혼자가 대부분, 아이 동반 1건"),
    ("H6", "근처 곱창집이 문을 닫았다", "지도 3회 조회(9/7·10/7·10/15), 상가정보", "부분적으로 맞았다",
     "돌다리곱창(17m)이 10월 지도 두 곳에서 사라지고 같은 건물에 태국음식점이 들어섬. 다만 도보 5분 안 곱창집은 여전히 4곳, 줄어든 쪽은 호프 등 술자리 업종"),
    ("H7", "단골은 많고 새 손님이 적다", "네이버 리뷰 48건 월별·방문 차수, 배달 리뷰 124건 재주문 표시", "부분적으로 맞았다",
     f"배달 리뷰 {re_cnt / len(rean) * 100:.0f}%가 재주문·단골 언급(쿠팡 'N번째 재주문' {coup['재주문'].notna().mean() * 100:.0f}%). 네이버 새 리뷰는 3~5월 {int(mon[[p for p in mon.index if p.month in (3, 4, 5)]].sum())}건 → 6~8월 {int(mon[[p for p in mon.index if p.month in (6, 7, 8)]].sum())}건. 단, 리뷰 {first / rev_n * 100:.0f}%는 첫 방문자가 썼다 → 새 손님이 끊긴 건 아니다"),
    ("기타", "하루 2회전은 된다", "홀 전표/영업일 ÷ 테이블 수(9개, 실내 5개)", "부분적으로 맞았다",
     f"평균은 테이블 9개 기준 {daily_hall.mean() / 9:.1f}회전, 실내 5개 기준 {hall[hall['tno'] <= 5].groupby('영업일자').size().mean() / 5:.1f}회전. 2회전을 넘긴 날은 {int((daily_hall >= 18).sum())}일/{len(daily_hall)}일"),
    ("기타", "웨이팅은 거의 없다", "전표 주문~결제 시각으로 동시 착석 테이블 수", "맞았다", f"하루 최대 동시 착석 중앙 {int(conc.median())}개, 최대 {int(conc.max())}개(테이블 9개). 자리가 차서 돌아간 날은 자료상 보이지 않는다"),
    ("기타", "배달이 30%쯤 된다", "배달앱 3사 주문 vs POS(음식 판매액 기준)", "부분적으로 맞았다", f"5월 {c1.iloc[0]['앱비중_기준(추정 포함)(%)']:.0f}% → 8월 {c1.iloc[3]['앱비중_기준(추정 포함)(%)']:.0f}%. 5~8월 합산 {d['앱음식'].sum() / d['합산'].sum() * 100:.0f}%. 8월에는 말씀에 가까워졌다"),
    ("기타", "객단가는 2만원대", "홀 매출 ÷ 안주 수량(안주 1인분 = 1명으로 봄)", "맞았다", f"1인당 지출은 야채곱창 가격(12,000원)의 약 {hall['총금액'].sum() / hall['안주수량'].sum() / 12000:.1f}배로 말씀하신 '2만원대' 바로 아래. 사장님은 1인 기준으로 말씀하셨다"),
]
t_cards = pd.DataFrame(cards, columns=["번호", "사장님 말씀", "확인한 자료", "결과", "한 줄 근거"]); save(t_cards, "W2_진술확인카드.csv")

CHIP = {"맞았다": (BLUE, "white"), "부분적으로 맞았다": (GOLD, TEXT), "다르게 나타났다": (NAVY, "white"), "아직 판단할 수 없다": (GRAY, "white"), "W1 결과 반영": ("white", TEXT)}
fig, ax = plt.subplots(figsize=(13, 10.2), dpi=300); ax.set_xlim(0, 100); ax.set_ylim(0, 100); ax.axis("off")
cols = [(1.0, 4.5, "번호"), (5.8, 19, "사장님 말씀"), (25.5, 23, "확인한 자료"), (49.5, 13.5, "결과"), (63.5, 36, "한 줄 근거")]
y = 93.5
for x0, wd_, name in cols: ax.text(x0, y, name, fontsize=10, fontweight="bold", color=TEXT, va="center")
ax.plot([0.5, 99.5], [y - 2.2, y - 2.2], color=NAVY, lw=1)
y -= 3.2; rh_ = (y - 3) / len(cards)
for i, (no, st, src, res_, why) in enumerate(cards):
    yc = y - rh_ * (i + 0.5)
    if i % 2 == 1: ax.add_patch(Rectangle((0.5, yc - rh_ / 2), 99, rh_, fc="#F4F6FA", ec="none", zorder=0))
    ax.text(cols[0][0], yc, no, fontsize=9.5, color=TEXT, va="center", fontweight="bold")
    ax.text(cols[1][0], yc, "\n".join(textwrap.wrap("“" + st + "”", 15)), fontsize=9.5, color=TEXT, va="center", linespacing=1.35)
    ax.text(cols[2][0], yc, "\n".join(textwrap.wrap(src, 20)), fontsize=8.3, color=GRAY, va="center", linespacing=1.35)
    fc, tc = CHIP[res_]
    ax.add_patch(FancyBboxPatch((cols[3][0], yc - 1.6), cols[3][1] - 1.2, 3.2, boxstyle="round,pad=0.25,rounding_size=0.8", fc=fc, ec=TEXT if fc == "white" else fc, lw=0.8, zorder=2))
    ax.text(cols[3][0] + (cols[3][1] - 1.2) / 2, yc, res_, fontsize=8.6, color=tc, ha="center", va="center", zorder=3, fontweight="bold")
    ax.text(cols[4][0], yc, "\n".join(textwrap.wrap(why, 44)), fontsize=8.1, color=TEXT, va="center", linespacing=1.35)
ax.plot([0.5, 99.5], [3, 3], color=LIGHT, lw=0.8)
ax.set_title("사장님 말씀 일곱 가지 가운데 다섯 가지는 자료와 맞거나 일부 맞았다 (H1·H2는 W1 판정 반영)", fontsize=12.5, color=TEXT, loc="left", pad=6, fontweight="bold")
fig.text(0.01, 0.006, "출처: POS 일자별(2026-02-24~09-05), 테이블별 전표 홀 421장(06-09~08-31), 배달앱 3사 주문(05-01~08-31), 네이버 리뷰 48건·배달 리뷰 124건(10-07 수집), 서울시 상권분석서비스, 네이버·카카오 지도(10-15).\n"
         "결과 구분: 맞았다 / 부분적으로 맞았다 / 다르게 나타났다 / 아직 판단할 수 없다. 단위: 비율·건수·지수. 통계 세부는 결과.md 부록.", fontsize=7, color=GRAY, linespacing=1.5)
fig.tight_layout(rect=(0, 0.03, 1, 1)); fig.savefig(OUT / "W2-3_진술확인카드.png"); plt.close(fig); print("saved W2-3")

# ---------------------------------------------------------------- 10. 기록표 초안
rec = []
for wk_ in range(1, 5):
    rec.append({"주차": wk_, "구분": "기준(직전 4주 화요일)", "날짜": "", "요일": "화", "영업": "O", "날씨(맑음/비/더움/추움)": "", "공휴일·연휴": "", "홀 전표 수": "", "홀 매출(사장님 기록)": "", "배달 건수": "", "배달 매출": "", "홀+배달 합계": "", "새 손님 체감(많음/보통/적음)": "", "피로도(1~5)": "", "메모": "시범 시작 전 4주의 화요일 값을 먼저 적는다"})
for wk_ in range(1, 5):
    rec.append({"주차": wk_, "구분": "시범(일요일 영업)", "날짜": "", "요일": "일", "영업": "O", "날씨(맑음/비/더움/추움)": "", "공휴일·연휴": "", "홀 전표 수": "", "홀 매출(사장님 기록)": "", "배달 건수": "", "배달 매출": "", "홀+배달 합계": "", "새 손님 체감(많음/보통/적음)": "", "피로도(1~5)": "", "메모": ""})
    rec.append({"주차": wk_, "구분": "시범(화요일 휴무)", "날짜": "", "요일": "화", "영업": "X", "날씨(맑음/비/더움/추움)": "", "공휴일·연휴": "", "홀 전표 수": "-", "홀 매출(사장님 기록)": "-", "배달 건수": "-", "배달 매출": "-", "홀+배달 합계": "-", "새 손님 체감(많음/보통/적음)": "-", "피로도(1~5)": "", "메모": "쉬는 날 피로도만 적는다"})
save(pd.DataFrame(rec), "W2_일요일시범_기록표.csv")
print("done")
