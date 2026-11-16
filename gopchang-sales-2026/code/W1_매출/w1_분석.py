"""W1 매출 흐름·좌석 구역·날씨 심층 분석 (꿀꿀곱창).

입력(읽기만): 볼트 02_매출데이터(POS 일자별), 레포 20_테이블별매출(전표요약·원본판독), 19_배달데이터(주문·일별_채널),
같은 폴더 weather_openmeteo_myeongil.json.
출력(이 폴더에만): 표 csv, 그림 png(300dpi), 콘솔 요약. 원 금액은 계산에만 쓰고 출력에는 지수·비율·건수만 남긴다.
실행: python -X utf8 w1_분석.py
"""
import json
import warnings
from pathlib import Path

import matplotlib
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import statsmodels.api as sm
import statsmodels.formula.api as smf
from scipy import stats

warnings.filterwarnings("ignore")
matplotlib.use("Agg")

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[2]  # 레포 최상위(최종제출_북적북적팀/분석코드/W1_매출 기준)
POS_X = HERE.parent / "data" / "02_매출데이터_20260906.xlsx"
TABLE_X = REPO / "매출 데이터" / "20_테이블별매출_통합_0609-0918.xlsx"
DELIV_X = REPO / "매출 데이터" / "19_배달데이터_통합.xlsx"
C1_CSV = REPO / "분석" / "최종보고서" / "output" / "tables" / "C1_월별_채널구성_음식판매기준.csv"
WEATHER = HERE / "weather_openmeteo_myeongil.json"

# ---- 팔레트(발표자료 디자인 기준, 2026-10-16 변경). 여기만 바꾸면 모든 그림이 따라간다
NAVY = "#1B2A4A"    # 주색 · 실내 좌석
BLUE = "#0068B7"    # 보조 파랑 · 천막 좌석
GOLD = "#D4A843"    # 강조 금색 · 야외 좌석
GRAY = "#8A8A8A"    # 회색 · 비교 기준(6월 등)
LIGHT = "#C9D3E3"   # 연한 남회색 · 합계 막대·배경 강조
TEXT = "#1F2937"    # 글자
DELIV = "#5B7FB5"   # 배달앱
GRID = "#EEEEEE"    # 격자선(아주 연하게)
ZONE_COLOR = {"실내": NAVY, "천막": BLUE, "야외": GOLD}
import matplotlib.font_manager as _fm  # noqa: E402
# Noto Sans KR은 고정 두께(Regular/Bold) 파일이 있을 때만 쓴다. 가변 글꼴(…-VF.ttf) 하나뿐이면 matplotlib이 Thin(100)으로 그려
# 글자가 너무 가늘어지므로 Malgun Gothic으로 내려간다. 고정 두께 Noto Sans KR을 설치하면 자동으로 바뀐다.
_NOTO_STATIC = [f for f in _fm.fontManager.ttflist if f.name == "Noto Sans KR" and "VF" not in Path(f.fname).name.upper()]
_FONT = "Noto Sans KR" if _NOTO_STATIC else "Malgun Gothic"
plt.rcParams.update({"font.family": _FONT, "axes.unicode_minus": False, "font.size": 10, "text.color": TEXT,
                     "axes.labelcolor": TEXT, "axes.titlecolor": TEXT, "figure.facecolor": "white", "axes.facecolor": "white",
                     "axes.spines.top": False, "axes.spines.right": False, "axes.edgecolor": "#CFCFCF",
                     "axes.grid": True, "axes.grid.axis": "y", "grid.color": GRID, "grid.linewidth": 0.8,
                     "axes.axisbelow": True, "xtick.color": TEXT, "ytick.color": TEXT,
                     "legend.frameon": False, "savefig.dpi": 300, "savefig.bbox": "tight", "savefig.facecolor": "white"})
WD = ["월", "화", "수", "목", "금", "토", "일"]
# 2026년 공휴일(대체공휴일 포함). 분석 구간(2/24~9/5)에 걸리는 것만 써도 되지만 전체를 둔다
HOLIDAYS = pd.to_datetime(["2026-01-01", "2026-02-16", "2026-02-17", "2026-02-18", "2026-03-01", "2026-03-02",
                           "2026-05-05", "2026-05-24", "2026-05-25", "2026-06-03", "2026-06-06", "2026-08-15",
                           "2026-08-17", "2026-09-24", "2026-09-25", "2026-09-26", "2026-09-28", "2026-10-03",
                           "2026-10-05", "2026-10-09", "2026-12-25"])
TIP_STORE = 3000  # f06과 같은 가정(배민 가게배달 미확보분 배달팁 중앙값)


def save_csv(df, name, index=False):
    df.to_csv(HERE / name, index=index, encoding="utf-8-sig")


def footnote(fig, text, y=0.005):
    fig.text(0.01, y, text, fontsize=7.5, color=GRAY, ha="left", va="bottom", wrap=True)


def zone_of(table):
    n = pd.to_numeric(table.astype(str).str.extract(r"(\d+)")[0], errors="coerce")
    return pd.cut(n, [0, 5, 7, 9], labels=["실내", "천막", "야외"]).astype(str)


# ------------------------------------------------------------------ 읽기
def load_pos_daily():
    """POS 일자별 매출(홀+직접 포장). 02-24~09-05 165영업일."""
    df = pd.read_excel(POS_X, sheet_name="일자별매출", header=3)
    df = df[df["영업여부"] == "영업"].copy()
    df["날짜"] = pd.to_datetime(df["날짜"])
    df = df.rename(columns={"총매출액": "POS"})[["날짜", "POS"]].sort_values("날짜").reset_index(drop=True)
    df["월"] = df["날짜"].dt.month
    df["요일"] = df["날짜"].dt.dayofweek.map(dict(enumerate(WD)))
    df["경과일"] = (df["날짜"] - pd.Timestamp("2026-02-24")).dt.days + 1
    df["공휴일전날"] = (df["날짜"] + pd.Timedelta(days=1)).isin(HOLIDAYS).astype(int)
    return df


def load_orders():
    o = pd.read_excel(DELIV_X, sheet_name="주문")
    o["주문일시"] = pd.to_datetime(o["주문일시"])
    o["날짜"] = pd.to_datetime(o["날짜"])
    for c in ["음식금액", "고객배달팁", "가게즉시할인", "고객결제금액"]:
        o[c] = pd.to_numeric(o[c], errors="coerce")
    o = o[o["상태"] != "취소"].copy()
    o["월"] = o["날짜"].dt.month
    o["시"] = o["주문일시"].dt.hour
    # f06.food_sales 와 같은 정의: 메뉴 가격 합계 - 가게 부담 즉시할인. 미확보 배민은 결제금액 - 배달팁(가게배달 3,000 추정)
    f = o["음식금액"] - o["가게즉시할인"].fillna(0)
    est = o["고객결제금액"] - np.where(o["서비스유형"] == "가게배달", TIP_STORE, 0)
    o["음식판매"] = f.where(o["음식금액"].notna(), est)
    return o


def load_receipts():
    t = pd.read_excel(TABLE_X, sheet_name="전표요약")
    t["영업일자"] = pd.to_datetime(t["영업일자"])
    t["월"] = t["영업일자"].dt.month
    t["시"] = pd.to_datetime(t["최초주문"].astype(str), format="%H:%M:%S", errors="coerce").dt.hour
    hall = t[t["채널"] == "매장"].copy()  # '포장 1'(배달) · 배달-포장 제외
    hall["구역"] = zone_of(hall["테이블번호"])
    assert hall["구역"].isin(["실내", "천막", "야외"]).all()
    return t, hall


def load_weather():
    w = json.load(open(WEATHER, encoding="utf-8"))
    h = pd.DataFrame(w["hourly"])
    h["time"] = pd.to_datetime(h["time"])
    h["날짜"] = h["time"].dt.normalize()
    h["시"] = h["time"].dt.hour
    ev = h[h["시"].between(18, 23)].groupby("날짜").agg(저녁기온=("temperature_2m", "mean"), 저녁강수=("precipitation", "sum"),
                                                  저녁습도=("relative_humidity_2m", "mean"))
    d = pd.DataFrame(w["daily"]).rename(columns={"time": "날짜", "temperature_2m_max": "일최고기온", "temperature_2m_min": "일최저기온",
                                                 "temperature_2m_mean": "일평균기온", "precipitation_sum": "일강수",
                                                 "apparent_temperature_max": "체감최고"})
    d["날짜"] = pd.to_datetime(d["날짜"])
    wx = d.set_index("날짜").join(ev)
    # 저녁 체감온도(간이): Steadman 식 근사 — 기온·습도 기반. 보고서에는 '저녁 열지수(기온·습도)'로만 쓴다
    e = wx["저녁습도"] / 100 * 6.105 * np.exp(17.27 * wx["저녁기온"] / (237.7 + wx["저녁기온"]))
    wx["저녁체감"] = wx["저녁기온"] + 0.33 * e - 4.0
    return wx.reset_index()


# ------------------------------------------------------------------ 1. 월별 합산 지수(C1 재현)
def sec1_monthly(o, pos_daily, deliv_daily):
    """5~8월. POS 영업일 기준(일별_채널 POS영업여부=영업, f01.daily_pos와 같음)."""
    rows = []
    for m in [5, 6, 7, 8]:
        p = pos_daily[pos_daily["월"] == m]
        om = o[o["월"] == m]
        rows.append({"월": f"{m}월", "영업일": len(p), "_P": p["POS"].sum() / len(p), "_A": om["음식판매"].sum() / len(p),
                     "앱주문_영업일당": round(len(om) / len(p), 2)})
    df = pd.DataFrame(rows)
    b = df.iloc[0]
    tot0 = b["_P"] + b["_A"]
    df["홀포장_지수(5월=100)"] = (df["_P"] / b["_P"] * 100).round(1)
    df["배달앱_지수(5월=100)"] = (df["_A"] / b["_A"] * 100).round(1)
    df["합산_지수(5월=100)"] = ((df["_P"] + df["_A"]) / tot0 * 100).round(1)
    df["합산중_홀포장(5월합산=100)"] = (df["_P"] / tot0 * 100).round(1)
    df["합산중_배달앱(5월합산=100)"] = (df["_A"] / tot0 * 100).round(1)
    df["배달비중(%)"] = (df["_A"] / (df["_P"] + df["_A"]) * 100).round(1)
    r5, r8 = df.iloc[0], df.iloc[3]
    keep = (r8["_A"] - r5["_A"]) / (r5["_P"] - r8["_P"])
    keep6 = (df.iloc[1]["_A"] - r5["_A"]) / (r5["_P"] - df.iloc[1]["_P"])
    pub = df.drop(columns=["_P", "_A"])
    pub.loc[len(pub)] = {"월": "보전몫(배달 증가분÷홀 감소분)", "영업일": "", "앱주문_영업일당": "",
                         "홀포장_지수(5월=100)": "", "배달앱_지수(5월=100)": "", "합산_지수(5월=100)": "",
                         "합산중_홀포장(5월합산=100)": f"5→6월 {keep6:.3f}", "합산중_배달앱(5월합산=100)": f"5→8월 {keep:.3f}", "배달비중(%)": ""}
    save_csv(pub, "표1_월별_합산지수.csv")
    # C1 대조
    c1 = pd.read_csv(C1_CSV, encoding="utf-8-sig").set_index("월")
    chk = pd.DataFrame({"재현_합산": df["합산_지수(5월=100)"].values, "C1_합산": c1.loc[["5월", "6월", "7월", "8월"], "합산_지수(5월=100)"].values,
                        "재현_POS": df["홀포장_지수(5월=100)"].values, "C1_POS": c1.loc[["5월", "6월", "7월", "8월"], "POS_지수(5월=100)"].values,
                        "재현_앱": df["배달앱_지수(5월=100)"].values, "C1_앱": c1.loc[["5월", "6월", "7월", "8월"], "앱_지수(5월=100)"].values,
                        "재현_배달비중": df["배달비중(%)"].values, "C1_배달비중": c1.loc[["5월", "6월", "7월", "8월"], "앱비중_기준(추정 포함)(%)"].values},
                       index=["5월", "6월", "7월", "8월"])
    print("== 1. C1 대조\n", chk.to_string(), f"\n보전몫 5→8월 {keep:.3f}, 5→6월 {keep6:.3f}")
    assert (chk["재현_합산"] - chk["C1_합산"]).abs().max() < 0.06 and abs(keep - 0.188) < 0.002
    return df, keep, keep6


def fig1(df, keep):
    fig, ax = plt.subplots(figsize=(7.2, 4.2))
    x = np.arange(len(df))
    hall, app = df["합산중_홀포장(5월합산=100)"], df["합산중_배달앱(5월합산=100)"]
    ax.bar(x, hall, 0.5, color=NAVY, label="홀·직접 포장(POS)", zorder=2)
    ax.bar(x, app, 0.5, bottom=hall + 0.8, color=DELIV, label="배달앱(음식 판매금액)", zorder=2)
    for i, (h, a, tot) in enumerate(zip(hall, app, df["합산_지수(5월=100)"])):
        ax.text(i, h + a + 2.5, f"{tot:.0f}", ha="center", va="bottom", fontsize=11, fontweight="bold", color=TEXT)
        ax.text(i, h / 2, f"{h:.0f}", ha="center", va="center", fontsize=9, color="white")
        if a > 9:
            ax.text(i, h + 0.8 + a / 2, f"{a:.0f}", ha="center", va="center", fontsize=9, color="white")
        else:
            ax.text(i + 0.28, h + 0.8 + a / 2, f"{a:.0f}", ha="left", va="center", fontsize=9, color=TEXT)
    ax.set_xticks(x, df["월"])
    ax.set_ylim(0, 118)
    ax.set_ylabel("하루 매출 지수 (5월 합산 = 100)")
    ax.set_title("5월→8월 하루 매출은 30% 줄었고, 늘어난 배달이 메운 몫은 홀 감소분의 약 1/5이다", fontsize=11.5, loc="left", pad=12)
    ax.legend(loc="upper right", ncol=2, fontsize=8.5, bbox_to_anchor=(1, 1.02))
    ax.annotate(f"배달 증가분 ÷ 홀 감소분 = {keep:.2f}", xy=(3, df.iloc[3]["합산_지수(5월=100)"] + 9), fontsize=9, ha="center", color=DELIV)
    footnote(fig, "출처: POS 일자별 매출(홀+직접 포장), 배달앱 3사 주문(땡겨요·쿠팡이츠·배민, 음식 판매금액 = 메뉴 가격 - 가게 부담 할인). "
                  "기간 2026-05-01~08-31, 영업일당 평균. 막대 위 숫자는 합산 지수, 막대 안 숫자는 채널 몫.", y=-0.06)
    fig.savefig(HERE / "W1-1_월별_합산지수.png")
    plt.close(fig)


# ------------------------------------------------------------------ 2. 6월→8월 홀 매출 구역 분해
PERIODS = [("6월(9~30일)", "2026-06-09", "2026-06-30"), ("7월", "2026-07-01", "2026-07-31"), ("8월", "2026-08-01", "2026-08-31"),
           ("9월(1~18일) 참고", "2026-09-01", "2026-09-18")]


def sec2_zone(t, hall):
    rows = []
    for lab, s, e in PERIODS:
        days = t[(t["영업일자"] >= s) & (t["영업일자"] <= e)]["영업일자"].nunique()
        h = hall[(hall["영업일자"] >= s) & (hall["영업일자"] <= e)]
        for z in ["실내", "천막", "야외"]:
            hz = h[h["구역"] == z]
            rows.append({"기간": lab, "영업일": days, "구역": z, "전표": len(hz), "전표_일": len(hz) / days,
                         "_전표당": hz["총금액"].mean() if len(hz) else 0.0, "_일매출": hz["총금액"].sum() / days,
                         "술동반(%)": round((hz["주류부착"] == "Y").mean() * 100, 1) if len(hz) else np.nan})
        rows.append({"기간": lab, "영업일": days, "구역": "홀 전체", "전표": len(h), "전표_일": len(h) / days,
                     "_전표당": h["총금액"].mean(), "_일매출": h["총금액"].sum() / days,
                     "술동반(%)": round((h["주류부착"] == "Y").mean() * 100, 1)})
    df = pd.DataFrame(rows)
    base_total = df[(df["기간"] == PERIODS[0][0]) & (df["구역"] == "홀 전체")].iloc[0]
    base_zone = df[df["기간"] == PERIODS[0][0]].set_index("구역")
    df["일매출_지수(6월홀=100)"] = (df["_일매출"] / base_total["_일매출"] * 100).round(1)
    df["전표당금액_지수(6월같은구역=100)"] = df.apply(lambda r: r["_전표당"] / base_zone.loc[r["구역"], "_전표당"] * 100 if base_zone.loc[r["구역"], "_전표당"] else np.nan, axis=1).round(1)
    df["6월대비_기여(%p)"] = df.apply(lambda r: (r["_일매출"] - base_zone.loc[r["구역"], "_일매출"]) / base_total["_일매출"] * 100, axis=1).round(1)
    df["6월_홀매출_비중(%)"] = df["구역"].map((base_zone["_일매출"] / base_total["_일매출"] * 100).round(1))
    df["전표_일"] = df["전표_일"].round(2)
    pub = df.drop(columns=["_전표당", "_일매출"])
    save_csv(pub, "표2_구역별_홀매출_분해.csv")

    # 실내: 전표 수 몫 vs 전표당 금액 몫 (6월→8월, 로그 분해)
    i6 = df[(df["기간"] == PERIODS[0][0]) & (df["구역"] == "실내")].iloc[0]
    i8 = df[(df["기간"] == "8월") & (df["구역"] == "실내")].iloc[0]
    dl = np.log(i8["_일매출"] / i6["_일매출"])
    dn = np.log(i8["전표_일"] / i6["전표_일"])
    dp = np.log(i8["_전표당"] / i6["_전표당"])
    inner = pd.DataFrame([{"구역": "실내", "6→8월 일매출 변화(%)": round((np.exp(dl) - 1) * 100, 1),
                           "전표/일 변화(%)": round((np.exp(dn) - 1) * 100, 1), "전표당 금액 변화(%)": round((np.exp(dp) - 1) * 100, 1),
                           "기여 %p(6월 홀=100)": i8["6월대비_기여(%p)"],
                           "기여 중 전표수 몫(%p)": round(i8["6월대비_기여(%p)"] * dn / dl, 1), "기여 중 전표당금액 몫(%p)": round(i8["6월대비_기여(%p)"] * dp / dl, 1)}])
    # 홀 전체도 같은 방식
    h6 = df[(df["기간"] == PERIODS[0][0]) & (df["구역"] == "홀 전체")].iloc[0]
    h8 = df[(df["기간"] == "8월") & (df["구역"] == "홀 전체")].iloc[0]
    dlh, dnh = np.log(h8["_일매출"] / h6["_일매출"]), np.log(h8["전표_일"] / h6["전표_일"])
    inner.loc[1] = {"구역": "홀 전체", "6→8월 일매출 변화(%)": round((np.exp(dlh) - 1) * 100, 1), "전표/일 변화(%)": round((np.exp(dnh) - 1) * 100, 1),
                    "전표당 금액 변화(%)": round((h8["_전표당"] / h6["_전표당"] - 1) * 100, 1), "기여 %p(6월 홀=100)": round(h8["일매출_지수(6월홀=100)"] - 100, 1),
                    "기여 중 전표수 몫(%p)": round((h8["일매출_지수(6월홀=100)"] - 100) * dnh / dlh, 1),
                    "기여 중 전표당금액 몫(%p)": round((h8["일매출_지수(6월홀=100)"] - 100) * (1 - dnh / dlh), 1)}
    save_csv(inner, "표2b_실내_전표수vs전표당금액_분해.csv")
    print("== 2. 구역 분해\n", pub.to_string(index=False), "\n", inner.to_string(index=False))
    return df, inner


def fig2(df):
    p0, p8, p9 = PERIODS[0][0], "8월", PERIODS[3][0]
    g = df.set_index(["기간", "구역"])
    contrib = {z: g.loc[(p8, z), "6월대비_기여(%p)"] for z in ["실내", "천막", "야외"]}
    tot8 = g.loc[(p8, "홀 전체"), "일매출_지수(6월홀=100)"]
    tot9 = g.loc[(p9, "홀 전체"), "일매출_지수(6월홀=100)"]
    fig, (ax, ax2) = plt.subplots(1, 2, figsize=(8.6, 4.4), gridspec_kw={"width_ratios": [3.2, 1.3], "wspace": 0.25})
    labels = ["6월 홀\n(9~30일)", "실내\n1~5번", "천막\n6·7번", "야외\n8·9번", "8월 홀"]
    colors = [LIGHT, NAVY, BLUE, GOLD, LIGHT]
    zones = ["실내", "천막", "야외"]
    bottoms, heights, run = [0.0], [100.0], 100.0
    levels = [100.0]  # 각 단계가 끝난 뒤의 누적 수준(연결선용)
    for z in zones:
        c = contrib[z]
        bottoms.append(run if c >= 0 else run + c); heights.append(abs(c)); run += c; levels.append(run)
    bottoms.append(0.0); heights.append(tot8)
    x = np.arange(5)
    ax.bar(x, heights, 0.55, bottom=bottoms, color=colors, zorder=2)
    for i in range(1, 5):  # 연결선: 이전 단계 누적 수준에서 다음 막대까지
        ax.plot([x[i - 1] + 0.275, x[i] - 0.275], [levels[i - 1]] * 2, color="#BDBDBD", lw=0.8, zorder=1)
    ax.text(0, 100 + 2, "100", ha="center", va="bottom", fontsize=10, fontweight="bold", color=TEXT)
    for i, z in enumerate(zones, start=1):
        c = contrib[z]
        ax.text(x[i], bottoms[i] + heights[i] + 2, f"{c:+.1f}", ha="center", va="bottom", fontsize=10, fontweight="bold", color=TEXT)
    ax.text(4, tot8 + 2, f"{tot8:.1f}", ha="center", va="bottom", fontsize=10, fontweight="bold", color=TEXT)
    ax.set_xticks(x, labels)
    ax.set_ylim(0, 125)
    ax.set_ylabel("홀 하루 매출 지수 (6월 홀 = 100)")
    fig.suptitle("6월→8월 홀 매출 감소는 모두 천막·야외 좌석에서 나왔고, 실내는 늘었다", fontsize=11.5, x=0.01, ha="left", y=1.0)
    # 오른쪽: 구역별 전표/일 6·7·8·9월
    per = ["6월(9~30일)", "7월", "8월", PERIODS[3][0]]
    xx = np.arange(4)
    for j, z in enumerate(["실내", "천막", "야외"]):
        y = [g.loc[(p, z), "전표_일"] for p in per]
        ax2.plot(xx, y, color=ZONE_COLOR[z], lw=2, marker="o", ms=5, mec="white", mew=1.2, label=z, zorder=3)
        ax2.text(3.12, y[3], f"{y[3]:.2f}", va="center", fontsize=8, color=TEXT)
    ax2.axvspan(2.5, 3.5, color="#F3F3F3", zorder=0)
    ax2.text(3, ax2.get_ylim()[1] * 0.97, "참고", ha="center", va="top", fontsize=8, color=GRAY)
    ax2.set_xticks(xx, ["6월", "7월", "8월", "9월\n~18일"], fontsize=8.5)
    ax2.set_xlim(-0.3, 3.6)
    ax2.set_ylabel("구역별 전표/영업일(건)", fontsize=9)
    ax2.legend(fontsize=8, loc="center left", bbox_to_anchor=(0, 0.55))
    ax2.set_title("구역별 하루 전표 수", fontsize=9.5, loc="left", color=TEXT, pad=6)
    footnote(fig, "출처: 테이블별 전표(채널=매장) 2026-06-09~08-31, 9월은 1~18일 참고. 좌석 구역: 실내 1~5번, 천막 6·7번(비닐 테라스), 야외 8·9번(길가). "
                  f"기여 %p = (8월 구역 일매출 - 6월 구역 일매출) ÷ 6월 홀 일매출. 9월 홀 지수 {tot9:.1f}.", y=-0.07)
    fig.savefig(HERE / "W1-2_구역분해_워터폴.png")
    plt.close(fig)


# ------------------------------------------------------------------ 3. 저녁 기온·강수 × 구역별 전표 수
def zone_daily(t, hall, wx, s, e):
    days = pd.DatetimeIndex(sorted(t[(t["영업일자"] >= s) & (t["영업일자"] <= e)]["영업일자"].unique()))
    h = hall[(hall["영업일자"] >= s) & (hall["영업일자"] <= e)]
    cnt = h.groupby(["영업일자", "구역"]).size().unstack(fill_value=0).reindex(days, fill_value=0)
    for z in ["실내", "천막", "야외"]:
        if z not in cnt:
            cnt[z] = 0
    cnt["천막야외"] = cnt["천막"] + cnt["야외"]
    cnt["홀"] = cnt[["실내", "천막", "야외"]].sum(axis=1)
    d = cnt.join(wx.set_index("날짜"), how="left")
    d["요일"] = d.index.dayofweek.map(dict(enumerate(WD)))
    d["월"] = d.index.month
    d["주말"] = d.index.dayofweek.isin([4, 5]).astype(int)  # 금·토
    d["공휴일전날"] = (d.index + pd.Timedelta(days=1)).isin(HOLIDAYS).astype(int)
    return d.reset_index().rename(columns={"index": "날짜"})


def glm(formula, data, family="poisson"):
    fam = sm.families.Poisson() if family == "poisson" else sm.families.NegativeBinomial(alpha=1.0)
    if family == "nb":
        # alpha 추정: statsmodels NegativeBinomial(이산 모형)으로 alpha를 구해 GLM에 넣는다
        try:
            nbm = smf.negativebinomial(formula, data).fit(disp=0, maxiter=200)
            alpha = float(nbm.params.get("alpha", 1.0))
            alpha = min(max(alpha, 1e-3), 10)
        except Exception:
            alpha = 1.0
        fam = sm.families.NegativeBinomial(alpha=alpha)
    return smf.glm(formula, data, family=fam).fit()


def coef_row(res, var, label, model, zone, n, extra=""):
    b, se, p = res.params[var], res.bse[var], res.pvalues[var]
    ci = res.conf_int().loc[var]
    return {"구역": zone, "모형": model, "변수": label, "계수(log)": round(b, 4), "표준오차": round(se, 4),
            "1단위당 변화(%)": round((np.exp(b) - 1) * 100, 1), "95%CI 하한(%)": round((np.exp(ci[0]) - 1) * 100, 1),
            "95%CI 상한(%)": round((np.exp(ci[1]) - 1) * 100, 1), "p": round(p, 4), "n(일)": n, "비고": extra}


def sec3_weather(t, hall, wx):
    d = zone_daily(t, hall, wx, "2026-06-09", "2026-08-31")
    d9 = zone_daily(t, hall, wx, "2026-06-09", "2026-09-18")
    save_csv(d[["날짜", "요일", "실내", "천막", "야외", "천막야외", "홀", "저녁기온", "저녁강수", "저녁습도", "저녁체감", "일최고기온", "체감최고", "일강수"]].round(2), "표3a_일별_구역전표_날씨.csv")
    # 기온 구간표
    bins = [-np.inf, 24, 27, 29, np.inf]
    labs = ["24℃ 미만", "24~27℃", "27~29℃", "29℃ 이상"]
    for name, dd in [("6/9~8/31", d), ("6/9~9/18(참고)", d9)]:
        dd["기온구간"] = pd.cut(dd["저녁기온"], bins, labels=labs, right=False)
        g = dd.groupby("기온구간", observed=False).agg(일수=("홀", "size"), 실내=("실내", "mean"), 천막=("천막", "mean"), 야외=("야외", "mean"),
                                                   천막야외=("천막야외", "mean"), 홀=("홀", "mean"), 저녁기온평균=("저녁기온", "mean")).round(2)
        g.insert(0, "기간", name)
        if name == "6/9~8/31":
            bin_tab = g.copy()
        else:
            bin_tab = pd.concat([bin_tab, g])
    save_csv(bin_tab.reset_index(), "표3b_기온구간별_구역전표.csv")
    print("== 3. 기온 구간표\n", bin_tab.to_string())

    # 회귀
    rows = []
    specs = [
        ("포아송(요일 통제)", "poisson", "{y} ~ 저녁기온 + 저녁강수 + C(요일)", d, "저녁기온", ""),
        ("음이항(요일 통제)", "nb", "{y} ~ 저녁기온 + 저녁강수 + C(요일)", d, "저녁기온", ""),
        ("포아송 + 월 고정효과", "poisson", "{y} ~ 저녁기온 + 저녁강수 + C(요일) + C(월)", d, "저녁기온", "월 안의 기온 변동만 사용"),
        ("포아송, 일 최고기온", "poisson", "{y} ~ 일최고기온 + 일강수 + C(요일)", d, "일최고기온", ""),
        ("포아송, 저녁 체감(기온·습도)", "poisson", "{y} ~ 저녁체감 + 저녁강수 + C(요일)", d, "저녁체감", ""),
        ("포아송, 금·토만", "poisson", "{y} ~ 저녁기온 + 저녁강수", d[d["주말"] == 1], "저녁기온", ""),
        ("포아송, 월~목만", "poisson", "{y} ~ 저녁기온 + 저녁강수 + C(요일)", d[d["주말"] == 0], "저녁기온", ""),
        ("포아송, 9/18까지(참고)", "poisson", "{y} ~ 저녁기온 + 저녁강수 + C(요일)", d9, "저녁기온", "82일"),
        ("포아송 + 공휴일 전날", "poisson", "{y} ~ 저녁기온 + 저녁강수 + 공휴일전날 + C(요일)", d, "저녁기온", ""),
    ]
    for zone, y in [("천막·야외", "천막야외"), ("실내", "실내"), ("천막", "천막"), ("야외", "야외"), ("홀 전체", "홀")]:
        for model, fam, f, dd, tv, extra in specs:
            if zone in ("천막", "야외") and model not in ("포아송(요일 통제)", "음이항(요일 통제)"):
                continue
            try:
                res = glm(f.format(y=y), dd, fam)
                rows.append(coef_row(res, tv, "기온 +1℃", model, zone, len(dd), extra))
                rv = "일강수" if "일강수" in f else "저녁강수"
                rows.append(coef_row(res, rv, "강수 +1mm", model, zone, len(dd), extra))
            except Exception as ex:  # noqa: BLE001
                rows.append({"구역": zone, "모형": model, "변수": "실패", "비고": str(ex)[:60]})
    reg = pd.DataFrame(rows)
    save_csv(reg, "표3c_구역별_기온회귀.csv")
    main = reg[(reg["모형"] == "포아송(요일 통제)")]
    print(main.to_string(index=False))

    # 실내 vs 천막·야외 계수 차이: 두 구역을 쌓아 교차항으로 검정
    long = pd.concat([d.assign(구역="실내", y=d["실내"]), d.assign(구역="천막야외", y=d["천막야외"])], ignore_index=True)
    inter = smf.glm("y ~ C(구역, Treatment('실내')) * (저녁기온 + 저녁강수) + C(요일)", long, family=sm.families.Poisson()).fit(
        cov_type="cluster", cov_kwds={"groups": long["날짜"].astype("category").cat.codes})
    k_t = [k for k in inter.params.index if "천막야외" in k and "저녁기온" in k][0]
    k_r = [k for k in inter.params.index if "천막야외" in k and "저녁강수" in k][0]
    diff = pd.DataFrame([
        {"항목": "기온 계수 차이(천막·야외 − 실내), log", "값": round(inter.params[k_t], 4), "p": round(inter.pvalues[k_t], 4)},
        {"항목": "강수 계수 차이(천막·야외 − 실내), log", "값": round(inter.params[k_r], 4), "p": round(inter.pvalues[k_r], 4)},
        {"항목": "날짜 군집 표준오차, 66일×2구역", "값": len(long), "p": ""},
    ])
    save_csv(diff, "표3d_계수차이_교차항검정.csv")
    print(diff.to_string(index=False))
    # 월별 저녁 평균기온(5~9월, 전체 달력일)
    wm = wx[(wx["날짜"] >= "2026-05-01") & (wx["날짜"] <= "2026-09-30")].groupby(wx["날짜"].dt.month)[["저녁기온", "일최고기온", "저녁강수", "일강수"]].mean().round(1)
    wm.index = [f"{m}월" for m in wm.index]
    save_csv(wm.reset_index().rename(columns={"index": "월"}), "표3e_월별_저녁기온.csv")
    print(wm.to_string())
    return d, bin_tab, reg, diff, wm


def fig3(bin_tab, reg):
    b = bin_tab[bin_tab["기간"] == "6/9~8/31"]
    labs = b.index.astype(str).tolist()
    x = np.arange(len(labs))
    w = 0.25
    fig, ax = plt.subplots(figsize=(7.2, 4.3))
    for j, z in enumerate(["실내", "천막", "야외"]):
        ax.bar(x + (j - 1) * (w + 0.015), b[z], w, color=ZONE_COLOR[z], label=f"{z} {'1~5' if z == '실내' else ('6·7' if z == '천막' else '8·9')}번", zorder=2)
        for i, v in enumerate(b[z]):
            ax.text(x[i] + (j - 1) * (w + 0.015), v + 0.08, f"{v:.2f}" if z != "실내" else f"{v:.1f}", ha="center", va="bottom", fontsize=8, color=TEXT)
    ax.set_xticks(x, [f"{l}\n({n}일)" for l, n in zip(labs, b["일수"])])
    ax.set_ylabel("하루 전표 수(건/영업일)")
    ax.set_xlabel("저녁(18~24시) 평균기온")
    ax.set_ylim(0, max(b["실내"]) * 1.25)
    m = reg[(reg["모형"] == "포아송(요일 통제)") & (reg["변수"] == "기온 +1℃")].set_index("구역")
    ax.set_title("저녁 기온이 오를수록 천막·야외 전표는 줄었고, 실내 전표는 변하지 않았다", fontsize=11.5, loc="left", pad=12)
    ax.text(0.99, 0.97, f"포아송 회귀(요일·강수 통제), 기온 +1℃당\n천막·야외 {m.loc['천막·야외', '1단위당 변화(%)']:+.0f}% (p={m.loc['천막·야외', 'p']:.3f})  ·  "
            f"실내 {m.loc['실내', '1단위당 변화(%)']:+.1f}% (p={m.loc['실내', 'p']:.2f})", transform=ax.transAxes, ha="right", va="top", fontsize=8.5, color=TEXT,
            bbox=dict(boxstyle="round,pad=0.4", fc="#F5F7FA", ec="none"))
    ax.legend(loc="upper left", fontsize=8.5)
    footnote(fig, "출처: 테이블별 전표(채널=매장) 2026-06-09~08-31 66영업일 × Open-Meteo 재분석(명일동 좌표) 시간별 기온. "
                  "막대 = 기온 구간별 영업일당 평균 전표 수. 괄호는 구간 일수.", y=-0.12)
    fig.savefig(HERE / "W1-3_기온구간별_구역전표.png")
    plt.close(fig)


# ------------------------------------------------------------------ 4. POS 일매출 × 저녁기온 (165일)
def sec4_pos_temp(pos, wx):
    d = pos.merge(wx, on="날짜", how="left")
    d["logPOS"] = np.log(d["POS"])
    d["log경과일"] = np.log(d["경과일"])
    d["강수유무"] = (d["저녁강수"] > 0.5).astype(int)
    save_csv(d[["날짜", "요일", "월", "경과일", "공휴일전날", "저녁기온", "저녁강수", "일최고기온", "저녁체감"]].round(2).assign(POS지수_기간평균100=(d["POS"] / d["POS"].mean() * 100).round(1)),
             "표4a_POS일별_날씨(지수).csv")
    corr = d[["경과일", "log경과일", "저녁기온", "일최고기온"]].corr().round(3)
    print("== 4. 상관\n", corr.to_string())
    specs = [
        ("A. 추세 없음(상한)", "logPOS ~ 저녁기온 + 저녁강수 + 공휴일전날 + C(요일)"),
        ("B. 경과일 선형", "logPOS ~ 경과일 + 저녁기온 + 저녁강수 + 공휴일전날 + C(요일)"),
        ("C. 경과일 로그", "logPOS ~ log경과일 + 저녁기온 + 저녁강수 + 공휴일전날 + C(요일)"),
        ("D. 경과일 선형+로그", "logPOS ~ 경과일 + log경과일 + 저녁기온 + 저녁강수 + 공휴일전날 + C(요일)"),
        ("E. 경과일 2차", "logPOS ~ 경과일 + I(경과일**2) + 저녁기온 + 저녁강수 + 공휴일전날 + C(요일)"),
        ("F. 월 고정효과(월 안 변동만)", "logPOS ~ C(월) + 저녁기온 + 저녁강수 + 공휴일전날 + C(요일)"),
        ("G. 경과일 선형, 일 최고기온", "logPOS ~ 경과일 + 일최고기온 + 일강수 + 공휴일전날 + C(요일)"),
        ("H. 경과일 선형, 5~9월만", "logPOS ~ 경과일 + 저녁기온 + 저녁강수 + 공휴일전날 + C(요일)"),
    ]
    rows, fits = [], {}
    for name, f in specs:
        dd = d[d["월"] >= 5] if name.startswith("H") else d
        res = smf.ols(f, dd).fit(cov_type="HC1")
        tv = "일최고기온" if "일최고기온" in f else "저녁기온"
        b, p = res.params[tv], res.pvalues[tv]
        ci = res.conf_int().loc[tv]
        row = {"모형": name, "기온 계수(log/℃)": round(b, 4), "기온 +1℃ 효과(%)": round((np.exp(b) - 1) * 100, 2), "95%CI(%)": f"{(np.exp(ci[0])-1)*100:+.1f}~{(np.exp(ci[1])-1)*100:+.1f}",
               "p": round(p, 4), "R²adj": round(res.rsquared_adj, 3), "n": int(res.nobs)}
        for tr in ["경과일", "log경과일"]:
            if tr in res.params:
                row[f"{tr} 계수"] = round(res.params[tr], 5)
                row[f"{tr} p"] = round(res.pvalues[tr], 4)
        vif = None
        if "경과일" in res.params and tv == "저녁기온" and "log경과일" not in res.params and "I(경과일 ** 2)" not in res.params:
            X = res.model.exog
            names = res.model.exog_names
            j = names.index("저녁기온")
            r2 = sm.OLS(X[:, j], np.delete(X, j, axis=1)).fit().rsquared
            vif = 1 / (1 - r2)
        row["기온 VIF"] = round(vif, 1) if vif else ""
        rows.append(row)
        fits[name] = res
    reg = pd.DataFrame(rows)
    save_csv(reg, "표4b_POS기온회귀_사양별.csv")
    print(reg.to_string(index=False))

    # 5월→6월, 5월→8월 POS 감소 중 기온으로 설명되는 몫
    mm = d.groupby("월").agg(POS=("POS", "mean"), 저녁기온=("저녁기온", "mean"), 일최고기온=("일최고기온", "mean"), 일수=("POS", "size"))
    rows = []
    for name, res in fits.items():
        tv = "일최고기온" if "일최고기온" in res.params else "저녁기온"
        b = res.params[tv]
        for m in [6, 7, 8]:
            dlog = np.log(mm.loc[m, "POS"] / mm.loc[5, "POS"])
            dtemp = mm.loc[m, tv] - mm.loc[5, tv]
            expl = b * dtemp
            rows.append({"모형": name, "비교": f"5월→{m}월", "실제 POS 변화(%)": round((np.exp(dlog) - 1) * 100, 1), "기온차(℃)": round(dtemp, 1),
                         "기온으로 설명되는 변화(%)": round((np.exp(expl) - 1) * 100, 1), "설명 몫(%)": round(expl / dlog * 100, 0) if dlog < 0 else np.nan})
    share = pd.DataFrame(rows)
    save_csv(share, "표4c_기온설명몫_월별.csv")
    print(share[share["비교"].isin(["5월→6월", "5월→8월"])].to_string(index=False))
    print(mm.round(2).to_string())
    return d, corr, reg, share, mm


def sec4b_seat_scenario(pos, hall, d3):
    """[추정] 5월→6월 POS 감소 중 '천막·야외 좌석 경로'로 설명 가능한 몫의 범위.
    5월 전표가 없으므로 5월 천막·야외 전표/일을 시나리오로 두고, 6월(9~30일) 천막·야외 전표당 금액을 곱해 POS 감소분과 비교한다."""
    p5 = pos[pos["월"] == 5]["POS"].mean()
    p6 = pos[pos["월"] == 6]["POS"].mean()
    h6 = hall[(hall["영업일자"] >= "2026-06-09") & (hall["영업일자"] <= "2026-06-30")]
    to6 = h6[h6["구역"].isin(["천막", "야외"])]
    per = to6["총금액"].mean()
    n6 = len(to6) / h6["영업일자"].nunique()
    cool = d3[d3["저녁기온"] < 23]
    june = d3[d3["날짜"].dt.month == 6]
    scen = {"6월 수준 그대로(좌석 몫 0)": n6,
            "6월 중 24℃ 미만 날 수준": june[june["저녁기온"] < 24]["천막야외"].mean(),
            "23℃ 미만 날 수준(6/9~8/31, 5일)": cool["천막야외"].mean(),
            "서늘한 날 상위(23℃ 미만 날 75% 분위)": cool["천막야외"].quantile(0.75)}
    rows = []
    for k, n5 in scen.items():
        extra = (n5 - n6) * per  # 5월에 더 있었을 천막·야외 하루 매출
        rows.append({"5월 천막·야외 전표/일 가정": k, "전표/일": round(n5, 2), "6월 대비 차이(건/일)": round(n5 - n6, 2),
                     "5→6월 POS 감소 중 좌석 몫(%)": round(extra / (p5 - p6) * 100, 0)})
    out = pd.DataFrame(rows)
    out["비고"] = f"5→6월 POS 일평균 {((p6 / p5) - 1) * 100:.1f}%, 6월 천막·야외 전표당 금액은 6월 홀 평균의 {per / h6['총금액'].mean() * 100:.0f}%"
    save_csv(out, "표4d_5→6월_좌석경로_시나리오.csv")
    print("== 4b. 5→6월 좌석 경로 시나리오\n", out.to_string(index=False))
    return out


def fig4_supp(d):
    """부록용: 경과일·기온과 POS(지수) — 공선성을 보여주는 그림."""
    fig, ax = plt.subplots(figsize=(7.2, 3.8))
    idx = d["POS"] / d["POS"].mean() * 100
    cmap = matplotlib.colors.LinearSegmentedColormap.from_list("navy_seq", [LIGHT, BLUE, NAVY])  # 한 색조 연→진
    sc = ax.scatter(d["경과일"], idx, c=d["저녁기온"], cmap=cmap, s=18, edgecolor="white", linewidth=0.5, zorder=3)
    roll = idx.rolling(7, center=True).mean()
    ax.plot(d["경과일"], roll, color=GOLD, lw=1.8, zorder=4, label="7영업일 이동평균")
    cb = fig.colorbar(sc, ax=ax, pad=0.01)
    cb.set_label("저녁 평균기온(℃)", fontsize=8.5)
    cb.ax.tick_params(labelsize=8)
    ymax = idx.max() * 1.08
    for m, lab in [(5, "5월"), (6, "6월"), (7, "7월"), (8, "8월")]:
        x0 = d[d["월"] == m]["경과일"].min()
        ax.axvline(x0, color="#DDDDDD", lw=0.8, zorder=1)
        ax.text(x0 + 1, ymax * 0.985, lab, fontsize=8, color=GRAY, va="top")
    ax.set_ylim(0, ymax)
    ax.set_xlabel("개업 경과일(2026-02-24 = 1)")
    ax.set_ylabel("POS 하루 매출 지수(165일 평균=100)")
    ax.set_title("개업 경과일과 저녁 기온은 같은 방향으로 움직여 두 효과를 따로 떼기 어렵다", fontsize=11, loc="left", pad=10)
    ax.legend(loc="lower left", fontsize=8.5)
    footnote(fig, "출처: POS 일자별 매출(홀+직접 포장) 2026-02-24~09-05 165영업일, Open-Meteo 시간별 기온(18~24시 평균). 점 색이 짙을수록 더운 저녁.", y=-0.07)
    fig.savefig(HERE / "W1-4b_부록_경과일_기온_POS.png")
    plt.close(fig)


# ------------------------------------------------------------------ 5. 시간대별 홀(구역) vs 배달, 6월 vs 8월
HOURS = [16, 17, 18, 19, 20, 21, 22, 23, 0, 1]
BUCKETS = [("16~17시", [16, 17]), ("18~19시", [18, 19]), ("20~21시", [20, 21]), ("22~23시", [22, 23]), ("0~1시", [0, 1])]


def sec5_hours(t, hall, o, pos_daily):
    out = []
    spans = [("6월(9~30일)", "2026-06-09", "2026-06-30"), ("8월", "2026-08-01", "2026-08-31")]
    for lab, s, e in spans:
        days_t = t[(t["영업일자"] >= s) & (t["영업일자"] <= e)]["영업일자"].nunique()
        days_p = len(pos_daily[(pos_daily["날짜"] >= s) & (pos_daily["날짜"] <= e)])
        h = hall[(hall["영업일자"] >= s) & (hall["영업일자"] <= e)]
        od = o[(o["날짜"] >= s) & (o["날짜"] <= e)]
        for name, hrs in BUCKETS:
            r = {"기간": lab, "시간대": name, "영업일(전표)": days_t, "영업일(POS)": days_p}
            for z in ["실내", "천막", "야외"]:
                r[f"{z}_전표/일"] = round((h[(h["구역"] == z) & h["시"].isin(hrs)].shape[0]) / days_t, 3)
            r["천막야외_전표/일"] = round(r["천막_전표/일"] + r["야외_전표/일"], 3)
            r["배달앱_주문/일"] = round(od[od["시"].isin(hrs)].shape[0] / days_p, 3)
            r["배달앱_배달만/일"] = round(od[od["시"].isin(hrs) & (od["수령"] == "배달")].shape[0] / days_p, 3)
            out.append(r)
        # 영업시간 밖(11~15시) 배달
        r = {"기간": lab, "시간대": "11~15시(홀 영업 전)", "영업일(전표)": days_t, "영업일(POS)": days_p, "실내_전표/일": round(h[(h["구역"] == "실내") & h["시"].between(11, 15)].shape[0] / days_t, 3),
             "천막_전표/일": 0.0, "야외_전표/일": 0.0, "천막야외_전표/일": 0.0, "배달앱_주문/일": round(od[od["시"].between(11, 15)].shape[0] / days_p, 3),
             "배달앱_배달만/일": round(od[od["시"].between(11, 15) & (od["수령"] == "배달")].shape[0] / days_p, 3)}
        out.append(r)
    df = pd.DataFrame(out)
    piv = df.pivot(index="시간대", columns="기간", values=["실내_전표/일", "천막야외_전표/일", "배달앱_주문/일"])
    piv = piv.reindex([b[0] for b in BUCKETS] + ["11~15시(홀 영업 전)"])
    chg = pd.DataFrame({"시간대": piv.index})
    for v in ["실내_전표/일", "천막야외_전표/일", "배달앱_주문/일"]:
        chg[f"{v} 6월"] = piv[(v, "6월(9~30일)")].values
        chg[f"{v} 8월"] = piv[(v, "8월")].values
        chg[f"{v} 변화"] = (piv[(v, "8월")] - piv[(v, "6월(9~30일)")]).round(3).values
    save_csv(df, "표5a_시간대별_구역_배달_6월vs8월.csv")
    save_csv(chg, "표5b_시간대별_변화.csv")
    # 상관: 날짜별 실내 전표 vs 배달 주문(6/9~8/31), 요일 통제 잔차 상관
    s, e = "2026-06-09", "2026-08-31"
    days = pd.DatetimeIndex(sorted(t[(t["영업일자"] >= s) & (t["영업일자"] <= e)]["영업일자"].unique()))
    hi = hall[(hall["영업일자"] >= s) & (hall["영업일자"] <= e) & (hall["구역"] == "실내")].groupby("영업일자").size().reindex(days, fill_value=0)
    dv = o[(o["날짜"] >= s) & (o["날짜"] <= e)].groupby("날짜").size().reindex(days, fill_value=0)
    dd = pd.DataFrame({"실내": hi.values, "배달": dv.values, "요일": days.dayofweek.map(dict(enumerate(WD))), "월": days.month})
    r1 = smf.ols("실내 ~ C(요일) + C(월)", dd).fit().resid
    r2 = smf.ols("배달 ~ C(요일) + C(월)", dd).fit().resid
    rho, p = stats.pearsonr(r1, r2)
    raw_rho, raw_p = stats.pearsonr(dd["실내"], dd["배달"])
    corr = pd.DataFrame([{"항목": "날짜별 실내 전표 × 배달 주문 상관(원값)", "r": round(raw_rho, 3), "p": round(raw_p, 3), "n(일)": len(dd)},
                         {"항목": "요일·월 통제 잔차 상관", "r": round(rho, 3), "p": round(p, 3), "n(일)": len(dd)}])
    save_csv(corr, "표5c_실내vs배달_일별상관.csv")
    print("== 5. 시간대\n", chg.to_string(index=False), "\n", corr.to_string(index=False))
    return df, chg, corr


def fig5(df):
    labs = [b[0] for b in BUCKETS]
    sub = df[df["시간대"].isin(labs)]
    g6 = sub[sub["기간"] == "6월(9~30일)"].set_index("시간대").reindex(labs)
    g8 = sub[sub["기간"] == "8월"].set_index("시간대").reindex(labs)
    fig, axes = plt.subplots(1, 3, figsize=(9.6, 3.9), sharey=False)
    panels = [("실내 1~5번 홀 전표", "실내_전표/일", NAVY), ("천막·야외 6~9번 홀 전표", "천막야외_전표/일", BLUE), ("배달앱 주문", "배달앱_주문/일", DELIV)]
    x = np.arange(len(labs))
    w = 0.36
    for ax, (title, col, color) in zip(axes, panels):
        ax.bar(x - w / 2 - 0.01, g6[col], w, color=LIGHT, label="6월(9~30일)", zorder=2)
        ax.bar(x + w / 2 + 0.01, g8[col], w, color=color, label="8월", zorder=2)
        for i in range(len(labs)):
            d = g8[col].iloc[i] - g6[col].iloc[i]
            top = max(g6[col].iloc[i], g8[col].iloc[i])
            ax.text(x[i], top + 0.05, f"{d:+.2f}", ha="center", va="bottom", fontsize=8, color=TEXT)
        ax.set_xticks(x, [l.replace("시", "") for l in labs], fontsize=8)
        ax.set_xlabel("첫 주문 시각(시)", fontsize=8.5)
        ax.set_title(title, fontsize=9.5, loc="left", color=TEXT)
        ax.set_ylim(0, max(g6[col].max(), g8[col].max()) * 1.35 + 0.1)
        ax.tick_params(axis="y", labelsize=8)
    axes[0].set_ylabel("영업일당 건수")
    axes[0].legend(fontsize=8, loc="upper left")
    fig.suptitle("저녁 피크(18~21시)에 배달이 늘었지만 같은 시간 실내 홀 전표는 줄지 않았다", fontsize=11.5, x=0.01, ha="left", y=1.0)
    footnote(fig, "출처: 테이블별 전표(채널=매장, 첫 주문 시각) · 배달앱 3사 주문(주문 시각). 6월 9~30일(19영업일) vs 8월(23영업일), 영업일당 건수. 막대 위 숫자 = 8월 - 6월.", y=-0.10)
    fig.savefig(HERE / "W1-4_시간대_홀구역_배달_6월vs8월.png")
    plt.close(fig)


# ------------------------------------------------------------------ 6. 술 동반·계란찜
def sec6_alcohol_egg(hall, t):
    rows = []
    for z in ["실내", "천막", "야외", "홀 전체"]:
        h = hall if z == "홀 전체" else hall[hall["구역"] == z]
        a = h[(h["영업일자"] >= "2026-06-09") & (h["영업일자"] <= "2026-06-30")]["주류부착"].eq("Y")
        b = h[(h["영업일자"] >= "2026-08-01") & (h["영업일자"] <= "2026-08-31")]["주류부착"].eq("Y")
        c = h[(h["영업일자"] >= "2026-07-01") & (h["영업일자"] <= "2026-07-31")]["주류부착"].eq("Y")
        p = stats.fisher_exact([[a.sum(), len(a) - a.sum()], [b.sum(), len(b) - b.sum()]])[1] if len(a) and len(b) else np.nan
        rows.append({"구역": z, "6월 전표": len(a), "6월 술동반(%)": round(a.mean() * 100, 1) if len(a) else np.nan, "7월 전표": len(c),
                     "7월 술동반(%)": round(c.mean() * 100, 1) if len(c) else np.nan, "8월 전표": len(b), "8월 술동반(%)": round(b.mean() * 100, 1) if len(b) else np.nan,
                     "6→8월 피셔 p": round(p, 3) if p == p else ""})
    alc = pd.DataFrame(rows)
    save_csv(alc, "표6a_구역별_술동반.csv")
    # 계란찜: 원본판독에서 전표별 포함 여부
    r = pd.read_excel(TABLE_X, sheet_name="원본판독")
    r["영업일자"] = pd.to_datetime(r["영업일자"])
    r = r[r["채널"] == "매장"] if "채널" in r else r
    egg = r[r["상품명"].astype(str).str.contains("계란찜")]
    last = egg["영업일자"].max()
    key = ["영업일자", "전표번호"]
    rec = hall[key + ["월"]].drop_duplicates()
    rec = rec.merge(egg.groupby(key).size().rename("계란찜").reset_index(), on=key, how="left").fillna({"계란찜": 0})
    side = r[r["구분"].astype(str) == "사이드"] if "구분" in r else r.iloc[0:0]
    rec = rec.merge(side[~side["상품명"].astype(str).str.contains("계란찜")].groupby(key).size().rename("다른사이드").reset_index(), on=key, how="left").fillna({"다른사이드": 0})
    def span(s, e):
        m = (rec["영업일자"] >= s) & (rec["영업일자"] <= e)
        x = rec[m]
        return {"홀 전표": len(x), "계란찜 포함 전표(%)": round((x["계란찜"] > 0).mean() * 100, 1), "계란찜 제외 사이드 포함(%)": round((x["다른사이드"] > 0).mean() * 100, 1)}
    eggtab = pd.DataFrame([{"기간": "6월(9~30일)", **span("2026-06-09", "2026-06-30")}, {"기간": "7월(1~28일, 판매 중)", **span("2026-07-01", "2026-07-28")},
                           {"기간": "7월(29~31일)+8월, 중단 후", **span("2026-07-29", "2026-08-31")}, {"기간": "8월", **span("2026-08-01", "2026-08-31")}])
    eggtab["마지막 판매일"] = last.strftime("%Y-%m-%d") if pd.notna(last) else ""
    # 계란찜 전표당 금액 몫: 6월 홀 매출 중 계란찜 금액 비중
    if "금액" in egg.columns:
        e6 = egg[(egg["영업일자"] >= "2026-06-09") & (egg["영업일자"] <= "2026-06-30")]["금액"].sum()
        h6 = hall[(hall["영업일자"] >= "2026-06-09") & (hall["영업일자"] <= "2026-06-30")]["총금액"].sum()
        eggtab["6월 홀 매출 중 계란찜 비중(%)"] = round(e6 / h6 * 100, 2)
    save_csv(eggtab, "표6b_계란찜_판매중단.csv")
    print("== 6. 술 동반·계란찜\n", alc.to_string(index=False), "\n", eggtab.to_string(index=False))
    return alc, eggtab


def main():
    pos = load_pos_daily()
    o = load_orders()
    t, hall = load_receipts()
    wx = load_weather()
    # 일별_채널의 POS와 02 파일 대조(5/1~9/5)
    dc = pd.read_excel(DELIV_X, sheet_name="일별_채널")
    dc["날짜"] = pd.to_datetime(dc["날짜"])
    dcp = dc[dc["POS영업여부"] == "영업"][["날짜", "POS총매출"]]
    chk = pos.merge(dcp, on="날짜", how="inner")
    print(f"POS 대조: 겹치는 영업일 {len(chk)}일, 불일치 {int((chk['POS'] != chk['POS총매출']).sum())}일; 02파일 영업일 {len(pos)}일 {pos['날짜'].min().date()}~{pos['날짜'].max().date()}")
    pos58 = pos[(pos["날짜"] >= "2026-05-01") & (pos["날짜"] <= "2026-08-31")]
    print("5~8월 POS 영업일", len(pos58), "앱 주문(5~8월, 취소 제외)", int((o["월"] <= 8).sum()), "홀 전표(6/9~8/31)", int((hall["영업일자"] <= "2026-08-31").sum()))

    m, keep, keep6 = sec1_monthly(o, pos58, None)
    fig1(m, keep)
    z, inner = sec2_zone(t, hall)
    fig2(z)
    d3, bins, reg3, diff, wm = sec3_weather(t, hall, wx)
    fig3(bins, reg3)
    d4, corr4, reg4, share4, mm = sec4_pos_temp(pos, wx)
    sec4b_seat_scenario(pos, hall, d3)
    fig4_supp(d4)
    df5, chg5, corr5 = sec5_hours(t, hall, o, pos)
    fig5(df5)
    sec6_alcohol_egg(hall, t)
    print("완료")


if __name__ == "__main__":
    main()
