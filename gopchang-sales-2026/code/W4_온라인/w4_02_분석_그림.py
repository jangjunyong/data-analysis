# -*- coding: utf-8 -*-
"""W4 찾기 쉬움: 리뷰·온라인 정보·검색 노출·SNS. 그림 W4-1~W4-5 + csv.
v6(2026-10-17): 그림 안 출처 삭제(본문 출처 줄로 대신), 제목 가운데, W4-4 정비 전·뒤 비교.
입력(읽기만): 예준 레포 리뷰 xlsx 2개, 인스타 xlsx, 볼트 02_매출데이터, 이 폴더의 온라인정보_조회.json
실행: python -X utf8 -I w4_02_분석_그림.py
"""
import json, csv, textwrap
from pathlib import Path
import numpy as np, pandas as pd, openpyxl
import matplotlib; matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.patches import Patch

# ---------- 팔레트·글꼴 (발표자료 디자인 기준 2026-10-16) ----------
NAVY, BLUE, GOLD, GRAY, LIGHT, TEXT = "#1B2A4A", "#0068B7", "#D4A843", "#8A8A8A", "#C9D3E3", "#1F2937"
GRID = "#E5E9F0"
# Noto Sans KR은 가변 글꼴(VF)만 설치돼 있어 matplotlib이 굵기 100(얇게)으로만 그린다 → 정적 글꼴이 있을 때만 쓰고, 없으면 Malgun Gothic
_static_noto = [f for f in font_manager.fontManager.ttflist if f.name == "Noto Sans KR" and "VF" not in f.fname]
FONT = "Noto Sans KR" if _static_noto else "Malgun Gothic"
print("font:", FONT)
plt.rcParams.update({"font.family": FONT, "axes.unicode_minus": False, "figure.facecolor": "white", "axes.facecolor": "white",
                     "axes.edgecolor": GRID, "axes.labelcolor": TEXT, "xtick.color": TEXT, "ytick.color": TEXT, "text.color": TEXT,
                     "axes.grid": True, "grid.color": GRID, "grid.linewidth": 0.8, "axes.spines.top": False, "axes.spines.right": False,
                     "axes.titlesize": 13, "axes.titleweight": "bold", "font.size": 10})
DPI = 300

OUT = Path(__file__).resolve().parent
REPO = OUT.parents[2]  # 레포 최상위
online = json.loads((OUT / "온라인정보_조회.json").read_text(encoding="utf-8"))


def save(fig, name):
    fig.savefig(OUT / name, dpi=DPI, bbox_inches="tight"); plt.close(fig); print("saved", name)


# ================= 1. 리뷰 =================
wb = openpyxl.load_workbook(REPO / "리뷰/리뷰_전체.xlsx", read_only=True, data_only=True)
rows = list(wb["네이버_리뷰"].iter_rows(values_only=True)); hdr = rows[0]
nv = pd.DataFrame(rows[1:], columns=hdr)
nv["작성일"] = pd.to_datetime(nv["작성일"]); nv["월"] = nv["작성일"].dt.month
monthly = nv.groupby("월").size().reindex(range(3, 11), fill_value=0)
monthly[10] += 1  # 10-16 재조회: 10/14 작성 리뷰 1건 추가(총 49건, 수집본 48건은 10/7까지)
assert monthly.sum() == 49, monthly.sum()
reply_cnt = nv["사장님 답글"].notna().sum() if "사장님 답글" in nv else 0

# POS 월별 일평균(홀+직접 포장) 지수, 5월=100. 02_매출데이터 월별요약
wb2 = openpyxl.load_workbook(OUT.parent / "data" / "02_매출데이터_20260906.xlsx", read_only=True, data_only=True)
pos = {}
for r in wb2["월별요약"].iter_rows(values_only=True):
    if r and isinstance(r[0], str) and r[0].endswith("월") and r[0][:-1].isdigit() and r[5]:
        pos[int(r[0][:-1])] = float(r[5])
pos_idx = {m: round(pos[m] / pos[5] * 100, 1) for m in range(3, 9)}  # 9월은 5영업일뿐이라 제외
pd.DataFrame({"월": list(range(3, 11)), "네이버_리뷰수": monthly.values,
              "홀POS_일평균_지수(5월=100)": [pos_idx.get(m) for m in range(3, 11)]}).to_csv(OUT / "W4-1_월별리뷰_홀지수.csv", index=False, encoding="utf-8-sig")

# 그림 W4-1
fig, (a1, a2) = plt.subplots(2, 1, figsize=(8, 6), sharex=True, gridspec_kw=dict(hspace=0.35))
months = list(range(3, 11)); x = np.arange(len(months))
cols = [NAVY] * 7 + [LIGHT]
a1.bar(x, monthly.values, color=cols, width=0.55, zorder=3)
for i, v in enumerate(monthly.values):
    a1.text(i, v + 0.25, str(v), ha="center", va="bottom", fontsize=9, color=TEXT)
a1.set_title("6월 이후 새 리뷰는 줄어든 채로 있음", loc="center")
a1.set_ylabel("네이버 방문자 리뷰 수(건)"); a1.set_ylim(0, 14); a1.grid(axis="x", visible=False)
a1.text(7, 3.2, "~16일", ha="center", fontsize=8, color=GRAY)
a2.plot(x[:6], [pos[m] / 1e4 for m in months[:6]], color=BLUE, lw=2, marker="o", ms=6, mec="white", mew=1.5, zorder=3)
for i, m in enumerate(months[:6]):
    a2.text(i, pos[m] / 1e4 + 1.2, f"{pos[m] / 1e4:.1f}", ha="center", fontsize=9, color=TEXT)
a2.set_ylim(15, 42); a2.grid(axis="x", visible=False)
a2.set_ylabel("홀·직접 포장 하루 매출(만 원)")
a2.set_xticks(x); a2.set_xticklabels([f"{m}월" for m in months])
a2.text(6.5, 24.5, "9월(18일까지) 홀 매출은\n전표 기준 6월의 90.7%\n(9·10월 POS 자료 없음)", ha="center", fontsize=8.5, color=TEXT, linespacing=1.4)
save(fig, "W4-1_월별리뷰_홀지수.png")

# ---- 배달 리뷰 주제 (배달리뷰_전체.xlsx 주제_분석) ----
wb3 = openpyxl.load_workbook(REPO / "리뷰/배달리뷰_전체.xlsx", read_only=True, data_only=True)
tr = [r for r in wb3["주제_분석"].iter_rows(values_only=True)]
sec, topics = None, {"긍정": {}, "부정": {}, "언급": {}}
for r in tr:
    if not r or r[0] is None: continue
    s = str(r[0])
    if s.startswith("긍정으로"): sec = "긍정"
    elif s.startswith("부정·아쉬움으로"): sec = "부정"
    elif s.startswith("평가 없이"): sec = "언급"
    elif s.startswith("부정·아쉬움·요청이"): sec = None
    elif sec and s not in ("주제", "리뷰 수") and isinstance(r[4], (int, float)):
        topics[sec][s] = int(r[4])
N_DELIV = 124
assert topics["긍정"]["양"] == 101 and topics["부정"]["맵기"] == 7
pd.DataFrame([(k, s, v, round(v / N_DELIV * 100, 1)) for k, d in topics.items() for s, v in d.items()],
             columns=["구분", "주제", "건수", "비율%"]).to_csv(OUT / "W4-2_배달리뷰_주제.csv", index=False, encoding="utf-8-sig")

kw = online["네이버"]["키워드투표"]; kw_n = online["네이버"]["방문자리뷰수"]  # 투표는 리뷰 단위(참여자 46명이 리뷰 49건에 투표)
fig, axes = plt.subplots(1, 3, figsize=(9.4, 4.0), gridspec_kw=dict(wspace=1.05))
# (a) 네이버 키워드 투표
labs, vals = zip(*kw["상위"][:6]); pct = [v / kw_n * 100 for v in vals]
ax = axes[0]; ax.barh(range(len(labs))[::-1], pct, color=NAVY, height=0.55, zorder=3)
ax.set_yticks(range(len(labs))[::-1]); ax.set_yticklabels(labs); ax.set_xlim(0, 115); ax.grid(axis="y", visible=False)
for i, (p, v) in enumerate(zip(pct, vals)): ax.text(p + 2, len(labs) - 1 - i, f"{p:.0f}% ({v})", va="center", fontsize=8.5)
ax.set_title(f"네이버 키워드 투표(리뷰 {kw_n}건 중 %)", loc="left", fontsize=11)
# (b) 배달 리뷰 긍정
pos_t = [("양이 많다", "양"), ("맛있다", "맛"), ("재주문·단골", "재주문·단골"), ("잡내 없다", "잡내·냄새"), ("친절·응대", "친절·응대"), ("가성비", "가성비"), ("포장 깔끔", "포장·소분")]
ax = axes[1]; v = [topics["긍정"][k] / N_DELIV * 100 for _, k in pos_t]
ax.barh(range(len(pos_t))[::-1], v, color=BLUE, height=0.55, zorder=3)
ax.set_yticks(range(len(pos_t))[::-1]); ax.set_yticklabels([l for l, _ in pos_t]); ax.set_xlim(0, 100); ax.grid(axis="y", visible=False)
for i, (p, (_, k)) in enumerate(zip(v, pos_t)): ax.text(p + 2, len(pos_t) - 1 - i, f"{p:.0f}% ({topics['긍정'][k]})", va="center", fontsize=8.5)
ax.set_title("배달앱 리뷰 124건: 칭찬한 것", loc="left", fontsize=11)
# (c) 배달 리뷰 아쉬운 점·요청 (건수)
neg_t = [("맵다", 7), ("당면 없음·위치", 3), ("맵기 조절 요청", 2), ("달다", 2), ("야채·쌈 더", 2), ("소스 넘침(포장)", 1), ("용기·두부 등 기타", 3)]
ax = axes[2]; ax.barh(range(len(neg_t))[::-1], [c for _, c in neg_t], color=GOLD, height=0.55, zorder=3)
ax.set_yticks(range(len(neg_t))[::-1]); ax.set_yticklabels([l for l, _ in neg_t]); ax.set_xlim(0, 10); ax.grid(axis="y", visible=False)
for i, (_, c) in enumerate(neg_t): ax.text(c + 0.2, len(neg_t) - 1 - i, f"{c}건 ({c / N_DELIV * 100:.1f}%)", va="center", fontsize=8.5)
ax.set_title("배달앱 리뷰 124건: 아쉬운 점·요청(건)", loc="left", fontsize=11)
fig.suptitle("손님은 양과 맛을 칭찬하고 맵기를 아쉬워함", x=0.5, ha="center", fontsize=13, fontweight="bold", y=1.02)
save(fig, "W4-2_칭찬과불편.png")

# ================= 2. 채널 정보 감사 매트릭스 =================
COLS = ["영업 시작", "마감", "휴무일", "메뉴 최신\n(짬뽕·계란찜 삭제)", "주류 표시", "1인분 가능", "포장 가능", "위치 안내", "대표 사진"]
M, X, N, U = "일치", "불일치", "정보 없음", "확인 불가"
matrix = {  # (상태, 셀 표기)
    "네이버 플레이스": [(M, "16:00"), (M, "02:00"), (M, "일요일"), (M, "9개 등록\n옛 메뉴 없음"), (X, "2/9종\n(켈리·무알콜)"), (N, "설명 없음"), (M, "포장·주문\n버튼 있음"), (M, "길 안내 있음"), (M, "사장님 사진\n+메뉴 사진")],
    "카카오맵":       [(M, "16:00\n(10-13 정정)"), (M, "02:00"), (M, "일요일"), (M, "메인 6개\n(사이드 없음)"), (N, "미등록"), (N, "없음"), (M, "포장가능"), (N, "지하철 자동\n길 안내 없음"), (X, "사장님 사진 0\n블로그 사진 자동")],
    "배달의민족":     [(U, "앱 전용"), (U, "접수 23시"), (U, "앱 전용"), (U, "앱 전용"), (U, "앱 전용"), (M, "1인분 단위\n(리뷰 근거)"), (M, "픽업 5건\n(리뷰)"), (U, "앱 전용"), (U, "앱 전용")],
    "쿠팡이츠":       [(U, "앱 전용"), (U, "접수 23시"), (U, "앱 전용"), (U, "앱 전용"), (U, "앱 전용"), (M, "340g 1인분\n(메뉴 칩)"), (U, "앱 전용"), (U, "앱 전용"), (U, "앱 전용")],
    "땡겨요":         [(U, "앱 전용"), (U, "접수 23시"), (U, "앱 전용"), (U, "앱 전용"), (U, "앱 전용"), (M, "1인분 단위\n(리뷰 근거)"), (M, "포장 11건\n(리뷰)"), (U, "앱 전용"), (U, "앱 전용")],
    "인스타 프로필":  [(N, "소개 없음"), (N, "소개 없음"), (N, "소개 없음"), (X, "마지막 게시 7월\n(3개월 휴면)"), (N, "없음"), (N, "없음"), (N, "없음"), (N, "없음"), (X, "메뉴·가게\n게시 12개뿐")],
    "매장 외관(사진)": [(N, "표기 없음"), (N, "표기 없음"), (M, "종이 안내\n'일요일 휴무'"), (X, "벽 메뉴판에\n곱창짬뽕 남음(9/7)"), (M, "소주·맥주\n·무알콜 표기"), (N, "표기 없음"), (N, "스티커 없음\n(사장님 방침)"), (X, "간판이 큰길에서\n잘 안 보임"), (M, "노란 천막·돼지 간판")],
}
CELL = {M: LIGHT, X: GOLD, N: GRAY, U: "white"}
rowsM = list(matrix.keys())
fig, ax = plt.subplots(figsize=(9.8, 5.6)); ax.grid(False)
for i, r in enumerate(rowsM):
    for j, (st, txt) in enumerate(matrix[r]):
        ax.add_patch(plt.Rectangle((j, len(rowsM) - 1 - i), 1, 1, facecolor=CELL[st], edgecolor="white", lw=2, alpha=0.35 if st == N else 1))
        if st == U:  # 빗금은 edgecolor를 따르므로 따로 얇은 회색 테두리로 그린다
            ax.add_patch(plt.Rectangle((j + 0.04, len(rowsM) - 1 - i + 0.04), 0.92, 0.92, facecolor="none", edgecolor=LIGHT, lw=0.6, hatch="///"))
        txt = "\n".join(textwrap.fill(ln, 8) for ln in txt.split("\n"))  # 칸 밖으로 넘치지 않게 7자마다 줄바꿈
        ax.text(j + 0.5, len(rowsM) - 1 - i + 0.5, txt, ha="center", va="center", fontsize=7.8,
                color=("#4A4A4A" if st == N else TEXT))
ax.set_xlim(0, len(COLS)); ax.set_ylim(0, len(rowsM))
ax.set_xticks(np.arange(len(COLS)) + 0.5); ax.set_xticklabels(COLS, fontsize=8.5); ax.xaxis.tick_top()
ax.set_yticks(np.arange(len(rowsM)) + 0.5); ax.set_yticklabels(rowsM[::-1], fontsize=10)
for s in ax.spines.values(): s.set_visible(False)
ax.tick_params(length=0)
cnt = {k: sum(1 for r in rowsM for st, _ in matrix[r] if st == k) for k in (M, X, N, U)}
ax.legend(handles=[Patch(facecolor=LIGHT, label=f"일치 {cnt[M]}"), Patch(facecolor=GOLD, label=f"불일치 {cnt[X]}"),
                   Patch(facecolor=GRAY, alpha=0.35, label=f"정보 없음 {cnt[N]}"), Patch(facecolor="white", edgecolor=GRAY, hatch="///", label=f"확인 불가(앱 전용·로그인) {cnt[U]}")],
          loc="upper center", bbox_to_anchor=(0.5, -0.02), ncol=4, frameon=False, fontsize=9)
fig.suptitle("영업시간은 맞춰졌고 메뉴 설명이 남았음", x=0.5, ha="center", fontsize=12.5, fontweight="bold", y=1.04)
save(fig, "W4-3_정보감사_매트릭스.png")
with open(OUT / "W4-3_정보감사_매트릭스.csv", "w", newline="", encoding="utf-8-sig") as f:
    w = csv.writer(f); w.writerow(["채널"] + [c.replace("\n", " ") for c in COLS])
    for r in rowsM: w.writerow([r] + [f"{st}: {txt.replace(chr(10), ' ')}" for st, txt in matrix[r]])
print("매트릭스 집계", cnt)

# ================= 3. 검색 노출 =================
rk = online["검색순위"]  # 정비 전(10-16 00:30)
# 정비 뒤(10-17): 사장님 네이버 플레이스 정비·플레이스 플러스 적용 뒤 같은 방법으로 다시 조회한 목록
after = pd.read_csv(OUT / "검색노출_순위_전체_정비후_1017.csv", encoding="utf-8-sig")
def rank_after(q):
    r = after[(after["키워드"] == q) & (after["꿀꿀"])]
    return f"{int(r['순위'].iloc[0])}위" if len(r) else "50위 밖"
def top3_after(q):
    r = after[after["키워드"] == q].head(3)
    return "  ·  ".join(f"{n}({c})" for n, c in zip(r["상호"], r["방문자리뷰"]))
fig, ax = plt.subplots(figsize=(10.6, 2.7)); ax.axis("off"); ax.grid(False)
cells = [[q, (f"{v['꿀꿀순위']}위" if isinstance(v["꿀꿀순위"], int) else v["꿀꿀순위"]), rank_after(q), top3_after(q)] for q, v in rk.items()]
tb = ax.table(cellText=cells, colLabels=["검색어", "정비 전", "정비 뒤", "정비 뒤 상위 3곳 (방문자 리뷰 수)"], loc="center", cellLoc="center",
              colWidths=[0.14, 0.09, 0.09, 0.68])
tb.auto_set_font_size(False); tb.set_fontsize(8.6); tb.scale(1, 2.2)
for (r, c), cell in tb.get_celld().items():
    cell.set_edgecolor(GRID); cell.set_linewidth(0.8)
    if r == 0: cell.set_facecolor(NAVY); cell.set_text_props(color="white", fontweight="bold")
    elif c in (1, 2):
        v = cells[r - 1][c]
        cell.set_facecolor(NAVY if v == "1위" else (GOLD if v == "50위 밖" else (LIGHT if v == "2위" else "white")))
        if v == "1위": cell.set_text_props(color="white", fontweight="bold")
fig.suptitle("정비 뒤 '명일역 곱창' 검색 1위가 됐음", x=0.5, ha="center", fontsize=12.5, fontweight="bold", y=0.98)
save(fig, "W4-4_검색노출_순위.png")
with open(OUT / "W4-4_검색노출_순위.csv", "w", newline="", encoding="utf-8-sig") as f:
    w = csv.writer(f); w.writerow(["검색어", "정비 전(10-16)", "정비 뒤(10-17)", "정비 뒤 상위3(리뷰 수)"]); [w.writerow(c) for c in cells]

# ================= 4. 인스타 =================
wb4 = openpyxl.load_workbook(REPO / "reports/instagram_post_analysis/곱창_인스타_게시물별_분석_최종.xlsx", read_only=True, data_only=True)
rows = list(wb4["전체 게시물 목록"].iter_rows(values_only=True))
ig = pd.DataFrame(rows[1:], columns=rows[0])
ig["t"] = pd.to_datetime(ig["게시일 KST"].str[:19]); ig["likes"] = pd.to_numeric(ig["좋아요"], errors="coerce")
ig = ig[ig["상단 고정"].isna() & ig["likes"].notna()].copy()
W = 7; parts = []
for acc, g in ig.groupby("계정"):
    g = g.sort_values("t").copy(); y = np.log1p(g["likes"].to_numpy(float))
    base = [np.median(np.r_[y[max(0, i - W):i], y[i + 1:i + W + 1]]) for i in range(len(y))]
    g["rel"] = y - np.array(base); parts.append(g)
ig = pd.concat(parts); assert len(ig) == 1054, len(ig)
def hbin(h): return "23~05시" if (h == 23 or h <= 5) else next(l for lo, hi, l in [(6, 10, "06~10시"), (11, 13, "11~13시"), (14, 16, "14~16시"), (17, 19, "17~19시"), (20, 22, "20~22시")] if lo <= h <= hi)
ig["hbin"] = ig["t"].dt.hour.map(hbin); ig["wd"] = ig["t"].dt.dayofweek
rng = np.random.default_rng(7)
def boot(v):
    v = np.asarray(v); ms = [np.exp(rng.choice(v, len(v)).mean()) for _ in range(2000)]
    return np.exp(v.mean()), np.percentile(ms, 5), np.percentile(ms, 95)
topic_tab = []
for tp, g in ig.groupby("큰 주제"):
    m, lo, hi = boot(g["rel"]); topic_tab.append(dict(주제=tp, 게시물=len(g), 배수=round(m, 2), 하한90=round(lo, 2), 상한90=round(hi, 2), 계정수=g["계정"].nunique()))
topic_tab = pd.DataFrame(topic_tab).sort_values("배수", ascending=False)
acc_tab = []
for (acc, tp), g in ig.groupby(["계정", "큰 주제"]):
    if len(g) >= 5:
        m, lo, hi = boot(g["rel"]); acc_tab.append(dict(계정=acc, 주제=tp, 게시물=len(g), 배수=round(m, 2), 하한90=round(lo, 2), 상한90=round(hi, 2)))
pd.concat([topic_tab.assign(계정="전체(4계정 합산)"), pd.DataFrame(acc_tab)]).to_csv(OUT / "W4-5_인스타_주제별_반응배수.csv", index=False, encoding="utf-8-sig")
# 회귀: 계정·시간대·요일·형식·주제 → F검정(강건 표준오차)
try:
    import statsmodels.formula.api as smf
    d = ig.rename(columns={"계정": "acc", "큰 주제": "topic", "형식": "fmt"})
    mod = smf.ols("rel ~ C(acc) + C(hbin) + C(wd) + C(fmt) + C(topic)", data=d).fit(cov_type="HC1")
    wt = mod.wald_test_terms(scalar=True); pv = {k: float(v) for k, v in wt.table["P>F"].items()} if "P>F" in wt.table else {k: float(v) for k, v in wt.table.iloc[:, 1].items()}
    p_h = pv.get("C(hbin)"); p_t = pv.get("C(topic)"); p_f = pv.get("C(fmt)"); p_w = pv.get("C(wd)")
except Exception as e:
    print("회귀 실패", e); p_h = p_t = p_f = p_w = None
print("F검정 p: 시간대", p_h, "주제", p_t, "형식", p_f, "요일", p_w)
# 운영·이용 안내 세부: 사정이 있는 공지 vs 날짜만
sub = ig[ig["큰 주제"] == "운영·이용 안내"].groupby("세부 주제")["rel"].agg(["size", lambda v: np.exp(v.mean())]).sort_values("size", ascending=False)
print(sub.head(12))

show = topic_tab[topic_tab["게시물"] >= 5].copy()
rel_for = {"메뉴·조리·고객 후기", "단체 주문 사례", "운영·이용 안내", "홀 이용·감사", "매장 일상·관리"}
fig, ax = plt.subplots(figsize=(7.8, 5.0))
yy = np.arange(len(show))[::-1]
colors = [GOLD if t in rel_for else LIGHT for t in show["주제"]]
ax.barh(yy, show["배수"] - 1, left=1, color=colors, height=0.55, zorder=3)
ax.errorbar(show["배수"], yy, xerr=[show["배수"] - show["하한90"], show["상한90"] - show["배수"]], fmt="none", ecolor=NAVY, elinewidth=1.2, capsize=3, zorder=4)
ax.axvline(1, color=NAVY, lw=1.2); ax.set_yticks(yy); ax.set_yticklabels([f"{t} (n={n})" for t, n in zip(show["주제"], show["게시물"])])
ax.set_xlabel("같은 계정 '보통 게시물' 대비 좋아요 배수(앞뒤 14개 중앙값 기준, 90% 구간)"); ax.grid(axis="y", visible=False)
ax.set_xlim(0.5, max(2.2, show["상한90"].max() + 0.15))
for y_, m, lo, hi in zip(yy, show["배수"], show["하한90"], show["상한90"]):
    ax.text(hi + 0.03 if m >= 1 else lo - 0.03, y_, f"{m:.2f}배", va="center", ha="left" if m >= 1 else "right", fontsize=8.5)
fig.suptitle("인스타 반응은 올린 시간보다 주제에 따라 갈렸음", x=0.5, ha="center", fontsize=13, fontweight="bold")
sub_note = (f"공지도 내용에 따라 갈림: 날짜만 적은 휴무 공지 {sub.loc['휴무 공지'].iloc[1]:.2f}배(n={int(sub.loc['휴무 공지'].iloc[0])}) vs "
            f"사정을 적은 명절·특별 휴무 공지 {sub.loc['명절·특별 휴무 공지'].iloc[1]:.2f}배(n={int(sub.loc['명절·특별 휴무 공지'].iloc[0])}), 영업시간 변경 {sub.loc['영업시간 변경'].iloc[1]:.2f}배(n={int(sub.loc['영업시간 변경'].iloc[0])})")
ax.text(0.99, 0.03, f"회귀(계정·시간대·요일·형식·주제 동시 통제, n=1,054) F검정\n게시 시간대 p={p_h:.2f} · 요일 p={p_w:.2f} · 형식 p={p_f:.3f} · 주제 p{'<0.001' if p_t is not None and p_t < 0.001 else f'={p_t:.3f}'}",
        transform=ax.transAxes, ha="right", va="bottom", fontsize=8.5, color=TEXT, bbox=dict(boxstyle="round", fc="white", ec=GRID))
ax.text(0.99, 0.22, sub_note.replace(" vs ", "\nvs ").replace(", 영업시간", "\n영업시간"), transform=ax.transAxes, ha="right", va="bottom", fontsize=8, color=TEXT)
ax.legend(handles=[Patch(facecolor=GOLD, label="꿀꿀곱창이 바로 쓸 수 있는 주제"), Patch(facecolor=LIGHT, label="참고(가맹·지점·외부 식당 등)")], loc="lower right", bbox_to_anchor=(1, 0.36), frameon=False, fontsize=8.5)
save(fig, "W4-5_인스타_주제별반응.png")
print(topic_tab.to_string(index=False))
