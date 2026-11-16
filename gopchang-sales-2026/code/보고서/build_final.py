# -*- coding: utf-8 -*-
"""양식 복사본 안에 최종 분석보고서 본문을 채운다. 원본 양식은 건드리지 않는다. (중간보고서 build.py 확장)"""
import os, sys, io, shutil, re, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.path.insert(0, os.path.dirname(__file__))
from content_final import BODY, APPENDIX
from PIL import Image
from pyhwpx import Hwp

HERE = os.path.dirname(os.path.abspath(__file__))
TPL = os.path.join(HERE, "양식_표준_분석보고서.hwp")
os.makedirs(os.path.join(HERE, "출력"), exist_ok=True)   # 제출본(사람이 다듬은 최종본)을 덮어쓰지 않게 따로 저장
OUT = os.path.join(HERE, "출력", "최종 분석보고서_북적북적팀.hwp")
PDF_OUT = os.path.join(HERE, "출력", "최종 분석보고서_북적북적팀.pdf")
PDF = os.path.join(os.environ["TEMP"], "최종보고서_check.pdf")

SEC_TITLES = {1: ["점포 기본 현황", "현장방문·인터뷰 주요 내용", "핵심 문제 선정 및 정의", "분석대상 및 범위 설정"],
              2: ["분석 목적 및 핵심 질문", "활용자료 및 분석방법", "주요 분석결과", "현황·원인 종합 및 핵심 시사점"],
              3: ["분석결과와 개선방향 도출", "개선전략 수립", "점포주·멘토 의견 및 반영내용", "전략 보완·구체화 내용"],
              4: ["전략별 실행계획", "점포 적용방안 및 역할분담", "실행 시 고려사항 및 대응방안", "전략 적용·실행 현황"],
              5: ["효과 확인 항목 및 방법", "현재 확인 가능한 변화 및 진행상황", "향후 효과 확인 및 보완계획", "산출물 및 점포 활용·관리방안"]}
CH_TITLES = {1: "점포 현황 및 주요 문제 확인", 2: "근거 기반 현황·원인 분석", 3: "개선전략 수립 및 구체화", 4: "실행계획 및 점포 적용", 5: "효과 확인방안 및 후속 활용"}

import psutil
OLD_HWP = {q.pid for q in psutil.process_iter(["name"]) if (q.info["name"] or "").lower() == "hwp.exe"}
shutil.copyfile(TPL, OUT)
hwp = Hwp(new=True, visible=False)   # 창을 숨겨 빌드 중 클릭으로 커서가 옮겨지는 일을 막는다
hwp.set_message_box_mode(0x10011)   # 확인·예 자동 (창이 숨어 있어 대화상자가 뜨면 멈추므로)
POS = {}

def style_name():
    return hwp.get_style()["Name"]

def find_style(text, style):
    """현재 위치에서 아래로 text 를 찾되, 문단 스타일이 style 인 것만. 문단 시작 pos 반환"""
    while hwp.find(text):
        if style_name() == style:
            hwp.MoveParaBegin(); return hwp.get_pos()
        hwp.MoveParaEnd()
    return None

def paste_para_from(pos):
    # 클립보드를 쓰므로 빌드 중 다른 곳에서 복사하면 붙여넣기가 비어 버린다 → 그러면 뒤의 역방향 찾기가
    # 양식 원본 제목(1장)을 바꿔 버리므로, 커서가 움직였는지 보고 다시 시도한다
    cur = hwp.get_pos()
    for _ in range(5):
        hwp.set_pos(*pos); hwp.MoveParaBegin(); hwp.MoveSelParaEnd(); hwp.Copy(); hwp.Cancel()
        hwp.set_pos(*cur); hwp.Paste()
        if hwp.get_pos() != cur: return
        time.sleep(0.5)
    raise RuntimeError(f"붙여넣기 실패(클립보드 사용 중?) 원본 {pos} 현재 {cur}")

def rb(old, new):
    ok = hwp.find(old, direction="Backward")
    if ok: hwp.insert_text(new)
    return ok

def keep_with_next():
    try:
        ps = hwp.HParameterSet.HParaShape
        hwp.HAction.GetDefault("ParagraphShape", ps.HSet)
        ps.KeepWithNext = 1
        hwp.HAction.Execute("ParagraphShape", ps.HSet)
    except Exception as e: print("  keep_with_next fail", e)

def para(style, text, keep=False):
    hwp.set_style(style); hwp.insert_text(text)
    if keep: keep_with_next()
    hwp.BreakPara()

def chapter(num, title):
    if num == 1: return
    hwp.BreakPage()
    paste_para_from(POS["ch"])
    keep_with_next()
    rb(CH_TITLES[1], title); hwp.MoveParaEnd(); rb("제1장", f"제{num}장")
    hwp.MoveParaEnd(); hwp.BreakPara()

def section(num, title):
    if POS.get("first_section_pending"):
        POS["first_section_pending"] = False; return   # 1장 1절 헤딩은 양식에 있음
    paste_para_from(POS["sec"])
    rb(SEC_TITLES[1][0], title); hwp.MoveParaEnd(); rb("1", str(num))
    hwp.HAction.Run("CloseEx")          # 번호 도형 밖으로
    hwp.MoveParaEnd(); keep_with_next(); hwp.BreakPara()

def table(header, rows, widths=None):
    rows_n, cols_n = len(rows) + 1, len(header)
    hwp.set_style("바탕글")
    pset = hwp.HParameterSet.HTableCreation
    hwp.HAction.GetDefault("TableCreate", pset.HSet)
    pset.Rows = rows_n; pset.Cols = cols_n
    TOTAL = 49324  # 본문 폭 (HWPUNIT, 양식 lineseg horzsize)
    pset.WidthType = 2; pset.HeightType = 0; pset.WidthValue = TOTAL
    pset.CreateItemArray("ColWidth", cols_n); pset.CreateItemArray("RowHeight", rows_n)
    ws = list(widths) if widths else [1] * cols_n
    eff = TOTAL - 120 - cols_n * 1020   # 셀 안쪽 여백(좌우 510씩)이 열마다 더해짐
    for i, w in enumerate(ws):
        pset.ColWidth.SetItem(i, int(eff * w / sum(ws)))
    pset.TableProperties.TreatAsChar = 1
    hwp.HAction.Execute("TableCreate", pset.HSet)
    data = [header] + [[str(c) for c in r] for r in rows]
    TABLES.append((hwp.ParentCtrl, data))   # 정렬은 다 쓴 뒤 align_tables()에서
    for r in range(rows_n):
        for c in range(cols_n):
            if not hwp.is_cell():   # 커서가 표 밖으로 나가면 본문이 뒤섞이므로 바로 멈춘다
                raise RuntimeError(f"표 채우기 중 셀 밖으로 나감: {header[0]} ({r},{c}) pos={hwp.get_pos()}")
            hwp.set_style("표내용-강조" if r == 0 else "표내용-일반")
            if r == 0:
                try:
                    hwp.HAction.Run("TableCellBlock"); hwp.cell_fill((226, 232, 240)); hwp.Cancel()
                except Exception as e: print("  fill fail", e)
            hwp.insert_text(data[r][c])
            if not (r == rows_n - 1 and c == cols_n - 1): hwp.TableRightCell()
    hwp.HAction.Run("CloseEx"); hwp.MoveParaEnd(); hwp.BreakPara()

TABLES = []

def align_tables():
    """양식 표 스타일은 오른쪽 정렬 → 머리행·짧은 칸은 가운데, 긴 글은 왼쪽. 글을 다 쓴 뒤 표마다 들어가 정렬만 바꾼다."""
    bad = 0
    for ctrl, data in TABLES:
        # pyhwpx get_into_nth_table(ctrl)는 버전 버그가 있어 같은 동작을 직접 한다
        hwp.set_pos_by_set(ctrl.GetAnchorPos(0)); hwp.hwp.FindCtrl(); hwp.ShapeObjTableSelCell(); hwp.Cancel()
        # 열 단위로 정한다: 본문 칸이 모두 짧은 열만 가운데, 긴 칸이 하나라도 있으면 그 열 전체 왼쪽(칸마다 섞이지 않게)
        short = [all(len(rw[c]) <= 12 for rw in data[1:]) for c in range(len(data[0]))]
        for r, row in enumerate(data):
            for c, txt in enumerate(row):
                if not hwp.is_cell():
                    bad += 1; break
                hwp.HAction.Run("ParagraphShapeAlignCenter" if r == 0 or short[c] else "ParagraphShapeAlignLeft")
                if not (r == len(data) - 1 and c == len(row) - 1): hwp.TableRightCell()
            else:
                continue
            break
        hwp.HAction.Run("CloseEx")
    print("표 정렬", len(TABLES), "개, 실패", bad)

def figure(path, caption, width_mm):
    w, h = Image.open(path).size
    width_mm = min(width_mm, 120 * w / h)   # 세로로 긴 그림이 한 쪽을 다 먹지 않게 높이 120mm 상한
    hwp.set_style("바탕글")
    hwp.insert_picture(path, treat_as_char=True, sizeoption=3, width=width_mm, height=round(width_mm * h / w))
    hwp.HAction.Run("ParagraphShapeAlignCenter")
    hwp.BreakPara()
    para("표/그림제목", caption)

def write_items(items):
    last = None
    for n, it in enumerate(items):
        k = it[0]
        pos = hwp.get_pos()
        # 커서가 본문(list 0)에서 뒤로 튀면 그 뒤 글이 앞쪽에 섞이므로 바로 멈춘다
        if last is not None and (pos[0] != 0 or pos[1] < last[1]):
            raise RuntimeError(f"커서 이상 #{n} {k} {str(it[1])[:30]!r}: 이전 {last} → 지금 {pos} (직전 항목 {items[n-1][0]} {str(items[n-1][1])[:30]!r})")
        last = pos
        if k == "pb": hwp.BreakPage(); continue
        if k == "ch": chapter(it[1], it[2]); print("ch", it[1])
        elif k == "sec": section(it[1], it[2])
        elif k == "p1" and it[1][:2] in ("B.", "C."):   # 부록 B·C 제목을 부록 A와 같은 장 제목 스타일로(v9)
            para("헤드-장", "부록 " + it[1][0].translate(str.maketrans("BC", "ＢＣ")) + it[1][1:], keep=True)
        elif k == "p1": para("본문-1단계", it[1], keep=len(it[1]) < 45)
        elif k == "p2": para("본문-2단계", it[1])
        elif k == "p3": para("본문-3단계(-)", "- " + it[1])
        elif k == "cap": para("표/그림제목", it[1], keep=True)
        elif k == "tbl": table(it[1], it[2], it[3] if len(it) > 3 else None)
        elif k == "fig": figure(it[1], it[2], it[3] if len(it) > 3 else 150)

FW = str.maketrans("0123456789", "０１２３４５６７８９")
def fullwidth(n): return f"{n:02d}".translate(FW)

try:
    hwp.open(OUT)
    # ── 앵커 위치 (본문 헤딩 원본) ──
    hwp.MoveDocBegin(); POS["ch"] = find_style(CH_TITLES[1], "헤드-장")
    hwp.MoveDocBegin(); POS["sec"] = find_style(SEC_TITLES[1][0], "본문-0단계")
    hwp.MoveDocBegin(); s = find_style("본문 문장작성 레벨1", "본문-1단계")
    hwp.MoveDocBegin(); e = find_style("부록 Ａ. 세부 분석자료", "헤드-장")
    print("anchors", POS, s, e)
    assert all([POS["ch"], POS["sec"], s, e])

    # ── 양식 안내문(※) 삭제: 텍스트만 지운다 (문단·구역 설정은 유지) ──
    NOTES = ["※ 보고서 작성 전 아래 지정 폰트를 웹에서 검색하여 설치해 주세요.",
             "(KoPub돋움체 Light, Medium, Bold / 에스코어 드림 1 Thin ~ 에스코어 드림 9 Black)",
             "※ 팀이름과 구성원 이름을 기재해주세요.",
             "※ 중간 분석보고서와 최종 분석보고서는 동일한 양식을 사용합니다.",
             "(최종 분석보고서 작성 시 표지 제목을 변경하고, 중간보고 이후 추가·변경된 내용을 각 항목에 반영하여 보완해 주세요.)",
             "※ 목차의 페이지 수는 각 팀의 보고서 분량에 맞게 수정하여 기재해 주세요.",
             "(목차의 장·절 명칭은 변경하지 않습니다.)",
             "※ 표지·목차에는 쪽번호가 표시되지 않도록 하고, 「제1장 1. 점포 기본 현황」부터 1페이지로 시작되도록 페이지 번호를 설정해 주세요.",
             "※ 부록은 팀별 분석내용, 참고자료 및 산출물 등에 따라 필요한 항목을 선택하여 작성",
             "(아래 부록 제목은 예시이며, 팀별 수행내용에 맞게 항목을 추가·삭제하거나 명칭 변경 가능)"]
    for t in NOTES:
        hwp.MoveDocBegin()
        if hwp.find(t): hwp.Delete()
        else: print("  note not found:", t[:30])
    # ── 작성가이드 페이지 삭제 (가이드 헤딩 ~ 제1장 헤딩 직전 문단) ──
    hwp.MoveDocBegin(); assert hwp.find("<보고서 작성가이드>"); hwp.MoveParaBegin(); g = hwp.get_pos()
    hwp.MoveDocBegin(); c1 = find_style(CH_TITLES[1], "헤드-장")
    print("guide range", g, c1)
    hwp.select_text(g[1], 0, c1[1] - 1, -1); hwp.Cut()
    # 남은 빈 문단(가이드 헤딩 자리) 제거 → 제1장 헤딩과 합치지 않도록 DeleteBack 만
    hwp.set_pos(g[0], g[1], 0); hwp.DeleteBack()

    # ── 표지 ──
    hwp.MoveDocBegin()
    if hwp.find("중간 분석보고서"): hwp.insert_text("최종 분석보고서")
    hwp.MoveDocBegin()
    hwp.find("________팀"); hwp.insert_text("북적북적팀")
    hwp.find("홍길동"); hwp.insert_text("장준용")
    hwp.find("홍길동"); hwp.insert_text("허예준")
    for _ in range(2):
        if hwp.find("홍길동"):
            hwp.MoveParaBegin(); hwp.MoveSelParaEnd(); hwp.Delete(); hwp.DeleteBack()

    # ── 앵커 재계산 (가이드 삭제로 문단 번호가 바뀜) ──
    hwp.MoveDocBegin(); POS["ch"] = find_style(CH_TITLES[1], "헤드-장")
    hwp.MoveDocBegin(); POS["sec"] = find_style(SEC_TITLES[1][0], "본문-0단계")
    print("anchors after guide cut", POS)
    # ── 샘플 본문 삭제 후 본문 작성 ──
    hwp.MoveDocBegin(); s = find_style("본문 문장작성 레벨1", "본문-1단계")
    hwp.MoveDocBegin(); e = find_style("부록 Ａ. 세부 분석자료", "헤드-장")
    hwp.select_text(s[1], 0, e[1] - 1, -1); hwp.Cut(); hwp.set_pos(*s)
    print('after cut pos', hwp.get_pos(), 'style', style_name())
    POS["first_section_pending"] = True
    write_items(BODY); hwp.DeleteBack()   # 본문 끝 빈 글머리 문단 제거(부록과 같은 방식)
    # v11: 본문이 쪽 끝까지 차면 부록 제목 앞 빈 문단이 홀로 한 쪽을 차지해 빈 쪽이 생긴다.
    # 부록 제목은 건드리지 않고, 앞 본문 문단 끝에서 Delete로 빈 문단을 흡수한다(본문 문단 모양 유지).
    def _ptext(k):
        hwp.set_pos(0, k, 0); hwp.MoveSelParaEnd(); t = hwp.get_selected_text() or ""; hwp.Cancel(); return t
    hwp.MoveDocBegin(); e = find_style("부록 Ａ. 세부 분석자료", "헤드-장")
    while e and e[1] >= 2 and not _ptext(e[1] - 1).strip() and _ptext(e[1] - 2).strip():
        hwp.set_pos(0, e[1] - 2, 0); hwp.MoveParaEnd(); hwp.Delete()
        print("부록 앞 빈 문단 흡수", e[1] - 1)
        hwp.MoveDocBegin(); e = find_style("부록 Ａ. 세부 분석자료", "헤드-장")

    # ── 부록 ──
    hwp.MoveDocBegin(); e = find_style("부록 Ａ. 세부 분석자료", "헤드-장")
    hwp.set_pos(*e); hwp.MoveParaEnd(); hwp.BreakPara()
    write_items(APPENDIX[1:]); hwp.DeleteBack()

    align_tables()
    nw = [c for c in hwp.ctrl_list if c.CtrlID == "nwno"]
    print("newNum ctrls:", len(nw))
    for c in nw[1:]: hwp.delete_ctrl(c)
    hwp.save_as(OUT); hwp.save_as(PDF, format="PDF")
    print("first save ok, pages", hwp.PageCount)

    # ── 목차 쪽번호 계산 ──
    import pymupdf
    doc = pymupdf.open(PDF)
    texts = [p.get_text().replace("\x01", " ") for p in doc]   # 한글 PDF 띄어쓰기는 \x01
    lines = [[l.strip() for l in t.splitlines()] for t in texts]
    def is_ch(ls, ch): return any(l.startswith(f"제{ch}장") and CH_TITLES[ch] in l for l in ls)
    body0 = next(i for i, ls in enumerate(lines) if is_ch(ls, 1) and "목   차" not in texts[i])
    def page_of(needle, start, chapter=None):
        for i in range(start, len(texts)):
            if chapter is not None:
                if is_ch(lines[i], chapter): return i - body0 + 1
            elif needle in lines[i] or any(l.startswith("부록") and needle in l for l in lines[i]):
                return i - body0 + 1
        return None
    nums = []   # 순서대로 (token, number)
    cur = body0
    for ch in range(1, 6):
        p = page_of(None, cur, chapter=ch); nums.append(("＃#" if ch > 1 else "01", p)); cur_ch = body0 + p - 1
        for si, st in enumerate(SEC_TITLES[ch], 1):
            ps = page_of(st, cur_ch); nums.append(("＃＃" if not (ch == 1 and si == 1) else "０１", ps))
        cur = cur_ch
    pa = page_of("부록 Ａ", cur)
    def page_sub(needle, start):   # 본문 부록 소제목은 앞에 '■'가 붙어 줄 전체 비교로는 안 잡힌다
        return next((i - body0 + 1 for i in range(start, len(texts)) if any(needle in l for l in lines[i])), None)
    pb = page_sub("Ｂ. 제작·적용", body0 + pa - 1); pc = page_sub("Ｃ. 참고자료", body0 + pa - 1)
    assert pb and pc, ("부록 쪽 번호 못 찾음", pb, pc)
    nums += [("##", pa), ("##", pa), ("##", pb), ("##", pc)]   # 부록 제목, A, B, C 순
    print("toc pages", nums)
    doc.close()   # PDF 잠금 해제 (다음 save_as 덮어쓰기용)
    hwp.MoveDocBegin()
    for tok, n in nums:
        if n is None: continue
        if tok in ("01", "０１"):   # 이미 01. 값이 다르면 교체
            if hwp.find(tok): hwp.insert_text(f"{n:02d}")   # v9: 목차 쪽 번호를 모두 반각으로(1장만 반각이던 것 통일)
            continue
        if hwp.find(tok): hwp.insert_text(f"{n:02d}")

    # ── 목차를 한 쪽에 맞춤: 양식 안내문(※)을 지운 빈 문단 정리 + 줄 간격 15% 축소 ──
    def para_text(i):
        hwp.set_pos(0, i, 0); hwp.MoveSelParaEnd(); t = hwp.get_selected_text() or ""; hwp.Cancel(); return t
    hwp.MoveDocBegin(); assert hwp.find("목   차"); top = hwp.get_pos()[1]
    idx = []
    for i in range(top + 1, top + 80):
        t = para_text(i); idx.append(i)
        if "참고자료 및 출처 목록" in t: break
    else:
        raise RuntimeError("목차 끝(참고자료 및 출처 목록)을 못 찾음")
    j5 = max(i for i in idx if "산출물 및 점포 활용" in para_text(i))
    jb = min(i for i in idx if para_text(i).startswith("부록"))
    for k in range(jb - 2, j5, -1):   # 부록 바로 앞 빈 문단(jb-1) 하나만 남긴다. 뒤에서부터 지워 번호가 안 밀리게
        if not para_text(k).strip():
            hwp.set_pos(0, k, 0); hwp.MoveSelParaEnd(); hwp.Delete(); hwp.DeleteBack()
    end = max(i for i in range(top + 1, top + 80) if "참고자료 및 출처 목록" in para_text(i))
    # 목차 뒤 빈 문단이 넘쳐 쪽 번호만 있는 빈 쪽이 생기므로, 1장 제목 바로 앞 하나만 남기고 지운다
    c1 = next(i for i in range(end + 1, end + 20) if para_text(i).startswith("제1장"))
    for k in range(c1 - 2, end, -1):
        if not para_text(k).strip():
            hwp.set_pos(0, k, 0); hwp.MoveSelParaEnd(); hwp.Delete(); hwp.DeleteBack()
    for k in range(top + 1, end + 1):
        hwp.set_pos(0, k, 0)
        ps = hwp.HParameterSet.HParaShape; hwp.HAction.GetDefault("ParagraphShape", ps.HSet)
        if ps.LineSpacing >= 170:
            ps.LineSpacing = int(ps.LineSpacing * 0.85); hwp.HAction.Execute("ParagraphShape", ps.HSet)
    hwp.save_as(OUT); hwp.save_as(PDF, format="PDF"); hwp.save_as(PDF_OUT, format="PDF")
    print("final save ok, pages", hwp.PageCount)
finally:
    try: hwp.quit()
    except Exception as ex: print("quit warn", ex)
    # 숨긴 창이 남으면 파일이 잠기므로 이 빌드가 띄운 한글만 정리
    for q in psutil.process_iter(["name"]):
        if (q.info["name"] or "").lower() == "hwp.exe" and q.pid not in OLD_HWP:
            q.kill()

import pymupdf
doc = pymupdf.open(PDF)
outdir = os.path.join(os.environ["TEMP"], "final_pages"); os.makedirs(outdir, exist_ok=True)
for f in os.listdir(outdir): os.remove(os.path.join(outdir, f))
for i, p in enumerate(doc):
    p.get_pixmap(dpi=60).save(os.path.join(outdir, f"p{i+1:02d}.png"))
print("rendered", len(doc), "pages to", outdir)
