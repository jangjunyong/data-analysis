"""W3 배달·포장 주문 구조와 세트 설계 — 집계·장바구니·세트 시뮬레이션.

읽기·정의는 예준 레포 분석/최종보고서/src/f01_집계.py·f06_검토반영_재계산.py를 그대로 쓴다.
기간: 배달앱 5/1~8/31, 전표(홀·직접 포장) 6/9~8/31. 원 금액 합계는 쓰지 않는다(건수·비율·지수·메뉴판 가격만).
출력: 이 폴더의 csv·png·w3_수치.json
"""
import json
import sys
from itertools import combinations
from pathlib import Path

import numpy as np
import pandas as pd

REPO = Path(__file__).resolve().parents[3]  # 레포 최상위
sys.path.insert(0, str(REPO / "분석" / "최종보고서" / "src"))
import f01_집계 as f01  # noqa: E402
import f06_검토반영_재계산 as f06  # noqa: E402

OUT = Path(__file__).resolve().parent
END = pd.Timestamp("2026-08-31")

# ---- 그림 팔레트(발표자료 디자인 기준, 2026-10-16 변경): 배달=파랑, 홀=남색, 포장=금색, 강조값=금색
NAVY, BLUE, GOLD, GRAY, PALE, INK = "#1B2A4A", "#0068B7", "#D4A843", "#8A8A8A", "#C9D3E3", "#1F2937"
GRID = "#EDF0F5"
CH_COLOR = {"배달": BLUE, "앱 포장": GOLD, "직접 포장": GOLD, "홀": NAVY}
FONT = "Noto Sans KR"  # 없으면 Malgun Gothic
PRICE = {"야채곱창": 12000, "알곱창": 13000, "순대곱창": 14000, "양념막창": 14000, "곱창전골": 16000, "마라곱창": 12000,
         "볶음밥": 2000, "주먹밥": 2000, "순대사리": 2000, "당면": 1000, "공기밥": 1000, "주류": 4000, "음료": 2000}
# 쿠팡이츠 배달 정산식(C3, 179건, R² 0.998). 거의 모든 주문에 가게 즉시할인 1,000원이 붙은 상태의 식
SLOPE, CONST = 0.8857, -4301
# 할인 분리식(상권보강/쿠팡이츠_정산식_재추정.txt)
SLOPE2, DISC2, CONST2 = 0.8854, -1.265, -3058
PICKUP_APP = 0.892   # 앱 포장 정산률 중앙값(A4: 쿠팡이츠 포장 89.2, 배민 픽업 89.2)
MONTHS = {5: "5월", 6: "6월", 7: "7월", 8: "8월"}
NUM = {}


def load():
    o, it, t, r = f01.load_orders(), f01.load_items(), f01.load_receipts(), f01.load_lines()
    o = o[o["날짜"] <= END].copy()
    t = t[t["영업일자"] <= END].copy()
    r = r[r["영업일자"] <= END].copy()
    it = it.merge(o[["채널", "주문번호", "월", "수령", "시각", "음식금액"]], on=["채널", "주문번호"], how="inner")
    return o, it, t, r


# ------------------------------------------------------------------ 품목 표준화 → 주문별 장바구니
def norm_app(row):
    name = str(row["메뉴_표준"])
    cls = row["메뉴분류"]
    if cls == "음료":
        return "음료"
    if cls == "주류":
        return "주류"
    if cls == "점심특선":
        return "점심특선"
    return {"당면추가": "당면", "중화면 추가": "중화면"}.get(name, name)


def baskets_app(it):
    rows = []
    for (ch, no), g in it.groupby(["채널", "주문번호"]):
        items = []
        for _, x in g.iterrows():
            q = int(x["수량"]) if pd.notna(x["수량"]) else 1
            items += [norm_app(x)] * max(q, 1)
            opt = str(x["옵션"])
            if opt != "nan":
                for p in opt.split("+"):
                    items.append({"당면추가": "당면"}.get(p.strip(), p.strip()))
        rows.append({"채널": ch, "주문번호": no, "수령": g["수령"].iloc[0], "월": g["월"].iloc[0], "시각": g["시각"].iloc[0],
                     "음식금액": g["음식금액"].iloc[0], "items": items,
                     "메인수": int(g.loc[g["메뉴분류"] == "메인", "수량"].sum())})  # f01.a5_menu 정의
    b = pd.DataFrame(rows)
    b["구분"] = np.where(b["수령"] == "포장", "앱 포장", "배달")
    return b


def baskets_receipt(r):
    r = r.copy()
    r["수량"] = pd.to_numeric(r["수량"], errors="coerce").fillna(1).astype(int)
    r["월"] = r["영업일자"].dt.month

    def nm(x):
        if x["구분"] == "음료":
            return "음료"
        if x["구분"] == "주류":
            return "주류"
        n = str(x["상품명"])
        return "점심특선" if n.startswith("[점심특선]") else n

    rows = []
    for (d, no), g in r.groupby(["영업일자", "전표번호"]):
        items = []
        for _, x in g.iterrows():
            items += [nm(x)] * max(x["수량"], 1)
        food = int(g.loc[g["구분"] != "주류", "금액"].sum())
        rows.append({"영업일자": d, "전표번호": no, "월": g["월"].iloc[0], "채널": g["채널"].iloc[0], "items": items,
                     "음식금액": food, "메인수": int(g.loc[g["구분"] == "안주", "수량"].sum())})  # 전표요약 '안주수량'(점심특선 포함)과 같은 정의
    b = pd.DataFrame(rows)
    b["구분"] = np.where(b["채널"] == "매장", "홀", "직접 포장")
    return b


MAINS = ["야채곱창", "알곱창", "순대곱창", "양념막창", "곱창전골", "마라곱창", "곱창짬뽕", "껍데기", "소창볶음", "비빔냉면"]
SIDES = ["볶음밥", "주먹밥", "순대사리", "당면", "공기밥", "계란찜", "곱창볶음밥", "중화면"]


def features(b):
    b = b.copy()
    b["사이드"] = b["items"].apply(lambda s: any(i in SIDES for i in s))
    b["사이드2천"] = b["items"].apply(lambda s: any(i in ["볶음밥", "주먹밥", "순대사리"] for i in s))
    b["볶음밥"] = b["items"].apply(lambda s: "볶음밥" in s)
    b["당면"] = b["items"].apply(lambda s: "당면" in s)
    b["주류"] = b["items"].apply(lambda s: "주류" in s)
    b["음료"] = b["items"].apply(lambda s: "음료" in s)
    b["메인구간"] = pd.cut(b["메인수"], [-1, 0, 1, 2, 99], labels=["0개", "1개", "2개", "3개+"])
    b["금액구간"] = pd.cut(b["음식금액"], [0, 14999, 19999, 29999, 1e9], labels=["1.5만 원 미만", "1.5~2만 원", "2~3만 원", "3만 원 이상"])
    b["단독1인"] = (b["메인수"] == 1) & (~b["주류"]) & (b["음식금액"] < 15000)
    return b


# ------------------------------------------------------------------ 1. 주문 구조
def order_structure(B):
    rows = []
    for ch in ["배달", "앱 포장", "직접 포장", "홀"]:
        x = B[B["구분"] == ch]
        xp = x[x["음식금액"].notna() & (x["음식금액"] > 0)]
        rows.append({
            "채널": ch, "건수": len(x), "메뉴금액확인_건수": len(xp),
            "메인0개(%)": round((x["메인수"] == 0).mean() * 100, 1),
            "메인1개(%)": round((x["메인수"] == 1).mean() * 100, 1),
            "메인2개(%)": round((x["메인수"] == 2).mean() * 100, 1),
            "메인3개이상(%)": round((x["메인수"] >= 3).mean() * 100, 1),
            "1인주문추정(메인1·술없음·1.5만미만)(%)": round(x["단독1인"].mean() * 100, 1),
            "메뉴금액_중앙값(원, 주류제외)": xp["음식금액"].median(),
            "1.5만원미만(%)": round((xp["음식금액"] < 15000).mean() * 100, 1),
            "사이드부착(%)": round(x["사이드"].mean() * 100, 1),
            "2천원사이드부착(%)": round(x["사이드2천"].mean() * 100, 1),
            "볶음밥부착(%)": round(x["볶음밥"].mean() * 100, 1),
            "당면부착(%)": round(x["당면"].mean() * 100, 1),
            "음료부착(%)": round(x["음료"].mean() * 100, 1),
            "주류부착(%)": round(x["주류"].mean() * 100, 1),
        })
    df = pd.DataFrame(rows)
    base = df.loc[df["채널"] == "홀", "메뉴금액_중앙값(원, 주류제외)"].iloc[0]
    df["메뉴금액_중앙값_지수(홀=100)"] = (df["메뉴금액_중앙값(원, 주류제외)"] / base * 100).round(1)
    df.to_csv(OUT / "W3_표1_채널별_주문구조.csv", index=False, encoding="utf-8-sig")
    return df


def monthly_small(B):
    """1.5만 원 미만 비중·중앙값 월별(배달; C4 정의 = 음식금액 확인 주문)."""
    x = B[(B["구분"] == "배달") & B["음식금액"].notna()]
    g = x.groupby("월").agg(건수=("음식금액", "size"),
                           **{"1.5만원미만(%)": ("음식금액", lambda s: round((s < 15000).mean() * 100, 1)),
                              "메뉴금액_중앙값(원)": ("음식금액", "median"),
                              "메인1개(%)": ("메인수", lambda s: round((s == 1).mean() * 100, 1)),
                              "볶음밥부착(%)": ("볶음밥", lambda s: round(s.mean() * 100, 1)),
                              "당면부착(%)": ("당면", lambda s: round(s.mean() * 100, 1))})
    g.index = [MONTHS[m] for m in g.index]
    tot = x
    g.loc["5~8월"] = [len(tot), round((tot["음식금액"] < 15000).mean() * 100, 1), tot["음식금액"].median(),
                     round((tot["메인수"] == 1).mean() * 100, 1), round(tot["볶음밥"].mean() * 100, 1), round(tot["당면"].mean() * 100, 1)]
    g.to_csv(OUT / "W3_표2_배달_월별_주문크기.csv", encoding="utf-8-sig")
    return g


def main_mix(B):
    rows = {}
    for ch in ["배달", "앱 포장", "직접 포장", "홀"]:
        cnt = pd.Series([i for s in B.loc[B["구분"] == ch, "items"] for i in s if i in MAINS]).value_counts()
        rows[ch] = (cnt / cnt.sum() * 100).round(1)
    df = pd.DataFrame(rows).fillna(0).sort_values("홀", ascending=False)
    df.to_csv(OUT / "W3_표3_메인메뉴_구성비.csv", encoding="utf-8-sig")
    return df


def hours(B, o, t):
    hrs = list(range(11, 24)) + [0, 1]
    d = {}
    d["배달"] = o[o["수령"] == "배달"].groupby("시각").size()
    d["앱 포장"] = o[o["수령"] == "포장"].groupby("시각").size()
    d["직접 포장"] = t[t["포장"]].groupby("최초시").size()
    d["홀"] = t[~t["포장"]].groupby("최초시").size()
    df = pd.DataFrame({k: v.reindex(hrs, fill_value=0) for k, v in d.items()})
    pct = (df / df.sum() * 100).round(1)
    pct.index = [f"{h:02d}시" for h in hrs]
    pct.to_csv(OUT / "W3_표4_시간대_비중.csv", encoding="utf-8-sig")
    return pct


# ------------------------------------------------------------------ 2. 장바구니(품목 쌍)
def pairs(B, ch, min_cnt=3):
    x = B[B["구분"] == ch]
    n = len(x)
    single = {}
    pair = {}
    for s in x["items"]:
        u = sorted(set(s))
        for i in u:
            single[i] = single.get(i, 0) + 1
        for a, b in combinations(u, 2):
            pair[(a, b)] = pair.get((a, b), 0) + 1
    rows = []
    for (a, b), c in pair.items():
        if c < min_cnt:
            continue
        rows.append({"채널": ch, "품목A": a, "품목B": b, "건수": c, "지지도(%)": round(c / n * 100, 1),
                     "신뢰도 A→B(%)": round(c / single[a] * 100, 1), "신뢰도 B→A(%)": round(c / single[b] * 100, 1),
                     "향상도": round((c / n) / ((single[a] / n) * (single[b] / n)), 2), "주문수": n})
    df = pd.DataFrame(rows).sort_values(["건수", "향상도"], ascending=False)
    return df, single, n


def basket_patterns(B, ch, top=10):
    x = B[B["구분"] == ch]
    pat = x["items"].apply(lambda s: " + ".join(f"{k}×{v}" if v > 1 else k for k, v in sorted(pd.Series(s).value_counts().items(), key=lambda kv: (kv[0] not in MAINS, kv[0]))))
    vc = pat.value_counts().head(top)
    return pd.DataFrame({"채널": ch, "장바구니": vc.index, "건수": vc.values, "비중(%)": (vc.values / len(x) * 100).round(1)})


def two_main_combos(B, ch):
    x = B[(B["구분"] == ch) & (B["메인수"] == 2)]
    pat = x["items"].apply(lambda s: " + ".join(sorted(i for i in s if i in MAINS)))
    vc = pat.value_counts()
    return pd.DataFrame({"채널": ch, "메인2개_조합": vc.index, "건수": vc.values, "메인2개주문중(%)": (vc.values / len(x) * 100).round(1)})


# ------------------------------------------------------------------ 3. 세트 시뮬레이션
def settle(p, disc=1000):
    return SLOPE * p + CONST, SLOPE2 * p + DISC2 * disc + CONST2


def set_sim():
    sets = [
        # 이름, 구성(가격 합산용), 세트 가격, 비교 기준 단품 가격
        ("기준: 야채곱창 단품", ["야채곱창"], 12000, 12000),
        ("기준: 곱창전골 단품", ["곱창전골"], 16000, 16000),
        ("A 혼곱 세트(1인): 야채곱창+당면+볶음밥(주먹밥 선택 가능)", ["야채곱창", "당면", "볶음밥"], 14000, 12000),
        ("B 둘이 세트(2인): 야채곱창+알곱창+당면+볶음밥", ["야채곱창", "알곱창", "당면", "볶음밥"], 27000, 12000),
        ("B' 둘이 세트(야채곱창 2개 선택 시)", ["야채곱창", "야채곱창", "당면", "볶음밥"], 26000, 12000),
        ("C 전골 집밥 세트(포장·배달): 곱창전골+당면+공기밥 2", ["곱창전골", "당면", "공기밥", "공기밥"], 18000, 16000),
    ]
    rows = []
    for name, comp, sp, basep in sets:
        sep = sum(PRICE[c] for c in comp)
        d1, d2 = settle(sp)
        sd1, _ = settle(sep)
        base_d, _ = settle(basep)
        rows.append({
            "세트": name, "따로 시킬 때(원)": sep, "세트 가격(원)": sp, "세트 할인(원)": sep - sp,
            "배달 점포몫(원, C3식)": round(d1), "배달 정산비율(%)": round(d1 / sp * 100, 1),
            "배달 점포몫(원, 할인분리식)": round(d2),
            "비교 기준 단품(원)": basep, "단품 대비 증가(원)": round(d1 - base_d), "단품 대비 배수": round(d1 / base_d, 2),
            "따로 시킬 때 배달 점포몫(원)": round(sd1), "세트화로 덜 받는 몫(원, 따로 대비)": round(sd1 - d1),
            "앱 포장 점포몫(원, 89.2%)": round(sp * PICKUP_APP), "직접 포장 점포몫(원, 100%)": sp,
        })
    df = pd.DataFrame(rows)
    df.to_csv(OUT / "W3_표6_세트_정산시뮬레이션.csv", index=False, encoding="utf-8-sig")
    return df


# ------------------------------------------------------------------ 5. 홀 볶음밥
def hall_fried_rice(B, t):
    h = B[B["구분"] == "홀"].copy()
    days = t.groupby("월")["영업일자"].nunique()
    qty = h.assign(q=h["items"].apply(lambda s: s.count("볶음밥"))).groupby("월")["q"].sum()
    perday = (qty / days).round(2)
    att = h.groupby("월")["볶음밥"].mean().mul(100).round(1)
    m = pd.DataFrame({"영업일": days, "볶음밥_수량": qty, "영업일당_볶음밥": perday, "볶음밥_부착전표(%)": att})
    m.index = ["6월(9~30일)", "7월", "8월"]
    m.to_csv(OUT / "W3_표7_홀_볶음밥_월별.csv", encoding="utf-8-sig")
    h["메인구간"] = pd.cut(h["메인수"], [-1, 0, 1, 2, 3, 99], labels=["0개", "1개", "2개", "3개", "4개+"])
    by = h.groupby("메인구간", observed=True).agg(전표=("볶음밥", "size"), 볶음밥부착=("볶음밥", "mean"), 사이드부착=("사이드", "mean"),
                                               주류부착=("주류", "mean"))
    by[["볶음밥부착", "사이드부착", "주류부착"]] = (by[["볶음밥부착", "사이드부착", "주류부착"]] * 100).round(1)
    by.to_csv(OUT / "W3_표8_홀_메인수별_볶음밥부착.csv", encoding="utf-8-sig")
    alc = h.groupby("주류")["볶음밥"].agg(["size", "mean"])
    alc["mean"] = (alc["mean"] * 100).round(1)
    return m, by, alc


# ------------------------------------------------------------------ 그림
def setup_font():
    """Noto Sans KR이 가변 글꼴(VF)뿐이면 matplotlib이 굵기 100으로만 그린다 → 임시 폴더에 Regular·Bold 정적 인스턴스를 만들어 등록."""
    import tempfile
    from matplotlib import font_manager as fm
    files = [f.fname for f in fm.fontManager.ttflist if f.name == FONT]
    if not files:
        return "Malgun Gothic"
    if any(f.weight in (400, 700) for f in fm.fontManager.ttflist if f.name == FONT):
        return FONT
    try:
        from fontTools.ttLib import TTFont
        from fontTools.varLib.instancer import instantiateVariableFont
        d = Path(tempfile.gettempdir()) / "w3_fonts"
        d.mkdir(exist_ok=True)
        for w, nm in [(400, "Regular"), (700, "Bold")]:
            p = d / f"NotoSansKR-{nm}.ttf"
            if not p.exists():
                instantiateVariableFont(TTFont(files[0]), {"wght": w}, updateFontNames=True).save(p)
            fm.fontManager.addfont(str(p))
        return FONT
    except Exception as e:  # noqa: BLE001
        print("글꼴 인스턴스 실패, Malgun Gothic 사용:", e)
        return "Malgun Gothic"


def figures(S, mix, pr_app, pr_hall, Bmonthly, fr_m, fr_by):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    font = setup_font()
    plt.rcParams.update({"font.family": font, "axes.unicode_minus": False, "font.size": 10,
                         "figure.facecolor": "white", "axes.facecolor": "white",
                         "axes.spines.top": False, "axes.spines.right": False, "axes.edgecolor": "#D5DAE3",
                         "axes.linewidth": 0.8, "xtick.color": INK, "ytick.color": INK, "text.color": INK,
                         "axes.labelcolor": INK})
    SUB, FOOT = "#4B5563", "#6B7280"
    chans = ["배달", "앱 포장", "직접 포장", "홀"]
    labels = [f"{c}\n(n={int(S.set_index('채널').loc[c, '건수'])})" for c in chans]

    def label_color(c):
        return "white" if c in (NAVY, BLUE, GRAY) else INK

    # ---- W3-1
    fig, axes = plt.subplots(1, 2, figsize=(8.4, 4.4), dpi=300, gridspec_kw={"wspace": 0.42})
    ax = axes[0]
    cols = [("메인1개(%)", GOLD, "메인 1개"), ("메인2개(%)", PALE, "메인 2개"), ("메인3개이상(%)", GRAY, "메인 3개 이상"), ("메인0개(%)", "#EEF1F6", "메인 없음")]
    left = np.zeros(len(chans))
    y = np.arange(len(chans))[::-1]
    for key, c, lab in cols:
        v = S.set_index("채널").loc[chans, key].values
        ax.barh(y, v, left=left, color=c, height=0.55, label=lab, edgecolor="white", linewidth=1.5)
        for yi, vi, li in zip(y, v, left):
            if vi >= 8:
                ax.text(li + vi / 2, yi, f"{vi:.0f}%", ha="center", va="center", fontsize=9, color=label_color(c))
        left += v
    ax.set_yticks(y)
    ax.set_yticklabels(labels)
    ax.set_xlim(0, 100)
    ax.set_xticks([0, 25, 50, 75, 100])
    ax.set_xlabel("주문(전표) 비중, %")
    ax.set_title("주문 1건의 메인 안주 개수", loc="left", fontsize=11, color=INK)
    ax.legend(loc="lower center", bbox_to_anchor=(0.5, -0.36), ncol=2, frameon=False, fontsize=8.5)
    ax.grid(axis="x", color=GRID, linewidth=0.8)
    ax.set_axisbelow(True)

    ax = axes[1]
    bins = ["1.5만 원 미만", "1.5~2만 원", "2~3만 원", "3만 원 이상"]
    bcol = [GOLD, PALE, GRAY, NAVY]
    left = np.zeros(len(chans))
    for b, c in zip(bins, bcol):
        v = np.array([Bmonthly[ch][b] for ch in chans])
        ax.barh(y, v, left=left, color=c, height=0.55, label=b, edgecolor="white", linewidth=1.5)
        for yi, vi, li in zip(y, v, left):
            if vi >= 8:
                ax.text(li + vi / 2, yi, f"{vi:.0f}%", ha="center", va="center", fontsize=9, color=label_color(c))
        left += v
    ax.set_yticks(y)
    ax.set_yticklabels(labels)
    ax.set_xlim(0, 100)
    ax.set_xticks([0, 25, 50, 75, 100])
    ax.set_xlabel("주문(전표) 비중, %")
    ax.set_title("주문 1건의 메뉴 금액(주류 제외)", loc="left", fontsize=11, color=INK)
    ax.legend(loc="lower center", bbox_to_anchor=(0.5, -0.36), ncol=2, frameon=False, fontsize=8.5)
    ax.grid(axis="x", color=GRID, linewidth=0.8)
    ax.set_axisbelow(True)

    s = S.set_index("채널")
    fig.suptitle("배달 주문 10건 중 7건은 메인이 하나뿐임",
                 x=0.5, ha="center", fontsize=13, fontweight="bold", color=INK)
    fig.subplots_adjust(top=0.87, bottom=0.24, left=0.09, right=0.98)
    fig.savefig(OUT / "그림_W3-1_주문크기_메인개수.png", facecolor="white")
    plt.close(fig)

    # ---- W3-2
    fig, axes = plt.subplots(1, 2, figsize=(8.6, 4.6), dpi=300, gridspec_kw={"wspace": 0.62, "width_ratios": [1, 1.25]})
    ax = axes[0]
    top = ["야채곱창", "알곱창", "순대곱창", "곱창전골", "마라곱창"]
    m = mix.loc[top, ["배달", "직접 포장", "홀"]]
    yy = np.arange(len(top))[::-1]
    hgt = 0.24
    for i, ch in enumerate(["배달", "직접 포장", "홀"]):
        v = m[ch].values
        ax.barh(yy + (1 - i) * hgt, v, height=hgt - 0.03, color=CH_COLOR[ch], label=ch)
        for yi, vi in zip(yy + (1 - i) * hgt, v):
            ax.text(vi + 0.8, yi, f"{vi:.0f}", va="center", fontsize=8, color=INK)
    ax.set_yticks(yy)
    ax.set_yticklabels(top)
    ax.set_xlim(0, 70)
    ax.set_xlabel("메인 수량 중 비중, %")
    ax.set_title("메인 메뉴 구성(수량 기준)", loc="left", fontsize=11, color=INK)
    ax.legend(frameon=False, fontsize=8.5, loc="lower right")
    ax.grid(axis="x", color=GRID, linewidth=0.8)
    ax.set_axisbelow(True)

    ax = axes[1]
    pa = pr_app.head(8).copy()
    ph = pr_hall.set_index(pr_hall["품목A"] + " + " + pr_hall["품목B"])
    pa["조합"] = pa["품목A"] + " + " + pa["품목B"]
    yy = np.arange(len(pa))[::-1]
    ax.barh(yy + 0.18, pa["지지도(%)"].values, height=0.34, color=BLUE, label="배달 주문")
    hv = [ph["지지도(%)"].get(k, np.nan) for k in pa["조합"]]
    ax.barh(yy - 0.18, [0 if np.isnan(v) else v for v in hv], height=0.34, color=NAVY, label="홀 전표")
    for yi, v in zip(yy + 0.18, pa["지지도(%)"].values):
        ax.text(v + 0.8, yi, f"{v:.0f}%", va="center", fontsize=8, color=INK)
    for yi, v in zip(yy - 0.18, hv):
        ax.text((0 if np.isnan(v) else v) + 0.8, yi, "-" if np.isnan(v) else f"{v:.0f}%", va="center", fontsize=8, color=GRAY)
    ax.set_yticks(yy)
    ax.set_yticklabels(pa["조합"])
    ax.set_xlim(0, max(pa["지지도(%)"].max(), np.nanmax(hv)) * 1.25)
    ax.set_xlabel("두 품목을 함께 담은 주문 비중, %")
    ax.set_title("배달에서 자주 함께 담은 품목 쌍 상위 8 (홀과 비교)", loc="left", fontsize=11, color=INK)
    ax.legend(frameon=False, fontsize=8.5, loc="lower right")
    ax.grid(axis="x", color=GRID, linewidth=0.8)
    ax.set_axisbelow(True)

    fig.suptitle("배달에서 가장 많이 함께 담는 것은 야채곱창과 당면임",
                 x=0.5, ha="center", fontsize=13, fontweight="bold", color=INK)
    fig.subplots_adjust(top=0.87, bottom=0.11, left=0.08, right=0.98)
    fig.savefig(OUT / "그림_W3-2_메인구성_조합.png", facecolor="white")
    plt.close(fig)

    # ---- W3-3 홀 볶음밥
    fig, axes = plt.subplots(1, 2, figsize=(10, 4), dpi=300, gridspec_kw={"wspace": 0.4})
    ax = axes[0]
    x = np.arange(len(fr_m))
    ax.bar(x, fr_m["영업일당_볶음밥"].values, width=0.5, color=[PALE, PALE, GOLD])
    for xi, v in zip(x, fr_m["영업일당_볶음밥"].values):
        ax.text(xi, v + 0.03, f"{v:.2f}", ha="center", fontsize=9, color=INK)
    ax.set_xticks(x)
    ax.set_xticklabels(fr_m.index)
    ax.set_ylim(0, fr_m["영업일당_볶음밥"].max() * 1.3)
    ax.set_ylabel("영업일당 볶음밥 수량(개)")
    ax.set_title("홀 볶음밥, 영업일당 수량", loc="left", fontsize=11, color=INK)
    ax.grid(axis="y", color=GRID, linewidth=0.8)
    ax.set_axisbelow(True)

    ax = axes[1]
    b = fr_by.loc[["1개", "2개", "3개", "4개+"]]
    x = np.arange(len(b))
    ax.bar(x, b["볶음밥부착"].values, width=0.5, color=NAVY)
    for xi, v in zip(x, b["볶음밥부착"].values):
        ax.text(xi, v + 0.8, f"{v:.0f}%", ha="center", fontsize=9, color=INK)
    ax.set_xticks(x)
    ax.set_xticklabels([f"메인 {k}\n(n={n})" for k, n in zip(b.index, b["전표"].values)])
    ax.set_ylim(0, max(b["볶음밥부착"].max() * 1.3, 10))
    ax.set_ylabel("볶음밥을 주문한 전표 비중, %")
    ax.set_title("메인 안주 개수별 볶음밥 주문 비중(홀)", loc="left", fontsize=11, color=INK)
    ax.grid(axis="y", color=GRID, linewidth=0.8)
    ax.set_axisbelow(True)

    fig.suptitle(f"홀 볶음밥은 영업일당 {fr_m['영업일당_볶음밥'].iloc[0]:.2f}개에서 {fr_m['영업일당_볶음밥'].iloc[-1]:.2f}개로 줄었고, 메인이 많은 테이블일수록 더 주문한다",
                 x=0.02, ha="left", fontsize=12.5, fontweight="bold", color=INK)
    fig.subplots_adjust(top=0.83, bottom=0.2, left=0.08, right=0.98)
    fig.savefig(OUT / "그림_W3-3_홀_볶음밥.png", facecolor="white")
    plt.close(fig)


def main():
    o, it, t, r = load()
    ba = features(baskets_app(it))
    br = features(baskets_receipt(r))
    B = pd.concat([ba[["구분", "월", "items", "음식금액", "메인수", "사이드", "사이드2천", "볶음밥", "당면", "주류", "음료", "메인구간", "금액구간", "단독1인"]],
                   br[["구분", "월", "items", "음식금액", "메인수", "사이드", "사이드2천", "볶음밥", "당면", "주류", "음료", "메인구간", "금액구간", "단독1인"]]],
                  ignore_index=True)

    S = order_structure(B)
    print("== 표1 채널별 주문구조\n", S.T.to_string())
    M = monthly_small(B)
    print("== 표2 배달 월별\n", M.to_string())
    mix = main_mix(B)
    print("== 표3 메인 구성비\n", mix.to_string())
    H = hours(B, o, t)
    print("== 표4 시간대\n", H.to_string())

    # 금액 구간 분포(채널별)
    amt = {}
    for ch in ["배달", "앱 포장", "직접 포장", "홀"]:
        x = B[(B["구분"] == ch) & B["음식금액"].notna() & (B["음식금액"] > 0)]
        amt[ch] = (x["금액구간"].value_counts(normalize=True).sort_index() * 100).round(1).to_dict()
    pd.DataFrame(amt).to_csv(OUT / "W3_표1b_채널별_금액구간.csv", encoding="utf-8-sig")

    # 장바구니
    P, singles, n = {}, {}, {}
    for ch in ["배달", "앱 포장", "직접 포장", "홀"]:
        P[ch], singles[ch], n[ch] = pairs(B, ch, min_cnt=3)
    pd.concat([p.head(10) for p in P.values()]).to_csv(OUT / "W3_표5_품목쌍_상위10.csv", index=False, encoding="utf-8-sig")
    print("== 표5 품목 쌍(배달)\n", P["배달"].head(10).to_string())
    print("== 표5 품목 쌍(홀)\n", P["홀"].head(10).to_string())
    print("== 표5 품목 쌍(직접 포장)\n", P["직접 포장"].head(10).to_string())
    pat = pd.concat([basket_patterns(B, ch) for ch in ["배달", "앱 포장", "직접 포장", "홀"]])
    pat.to_csv(OUT / "W3_표5b_장바구니_패턴_상위10.csv", index=False, encoding="utf-8-sig")
    print("== 표5b 장바구니 패턴\n", pat.to_string())
    tm = pd.concat([two_main_combos(B, ch) for ch in ["배달", "홀", "직접 포장"]])
    tm.to_csv(OUT / "W3_표5c_메인2개_조합.csv", index=False, encoding="utf-8-sig")
    print("== 표5c 메인 2개 조합\n", tm.head(25).to_string())

    # 단품 야채곱창 하나만(사이드 없음) 비중
    d = B[B["구분"] == "배달"]
    only_veg = d["items"].apply(lambda s: sorted(s) == ["야채곱창"]).mean() * 100
    veg_single = d["items"].apply(lambda s: s.count("야채곱창") == 1 and sum(1 for i in s if i in MAINS) == 1).mean() * 100
    veg_rice = d["items"].apply(lambda s: "야채곱창" in s and "볶음밥" in s).mean() * 100
    veg_dm = d["items"].apply(lambda s: "야채곱창" in s and "당면" in s).mean() * 100
    main1_noside = ((d["메인수"] == 1) & (~d["사이드"])).mean() * 100
    print(f"배달: 야채곱창만 단독 {only_veg:.1f}%, 야채곱창 1개 메인 {veg_single:.1f}%, 야채+볶음밥 {veg_rice:.1f}%, 야채+당면 {veg_dm:.1f}%, 메인1·사이드없음 {main1_noside:.1f}%")

    # 세트
    SIM = set_sim()
    print("== 표6 세트 시뮬레이션\n", SIM.T.to_string())

    # 홀 볶음밥
    fr_m, fr_by, alc = hall_fried_rice(B, t)
    print("== 표7 홀 볶음밥 월별\n", fr_m.to_string())
    print("== 표8 메인수별\n", fr_by.to_string())
    print("== 주류 동반별 볶음밥 부착\n", alc.to_string())

    figures(S, mix, P["배달"], P["홀"], amt, fr_m, fr_by)

    NUM.update({
        "표1": S.set_index("채널").to_dict(orient="index"),
        "배달_야채곱창만단독%": round(only_veg, 1), "배달_야채곱창1개메인%": round(veg_single, 1),
        "배달_야채+볶음밥%": round(veg_rice, 1), "배달_야채+당면%": round(veg_dm, 1), "배달_메인1사이드없음%": round(main1_noside, 1),
        "표2": M.reset_index().to_dict(orient="records"),
        "표7": fr_m.reset_index().to_dict(orient="records"), "표8": fr_by.reset_index().to_dict(orient="records"),
        "주류동반별_볶음밥": alc.reset_index().to_dict(orient="records"),
        "시간대": H.to_dict(),
    })
    with open(OUT / "w3_수치.json", "w", encoding="utf-8") as f:
        json.dump(NUM, f, ensure_ascii=False, indent=1, default=str)


if __name__ == "__main__":
    main()
