# -*- coding: utf-8 -*-
"""최종보고서 원고(마크다운) → content_final.py (build_final.py 가 읽는 BODY·APPENDIX 목록)

원고 규칙
  # 제N장 장제목          → ("ch", N, 장제목)
  ## N. 절제목            → ("sec", N, 절제목)
  ### 소제목              → ("p1", 소제목)        본문-1단계(■)
  일반 문단               → ("p2", 문단)          본문-2단계
  - 항목                  → ("p3", 항목)          본문-3단계(-)
  [표 n-n] 캡션            → ("cap", 캡션)          다음에 오는 마크다운 표의 제목
  <!-- widths: 1,3,1.5 --> 표 바로 위에 두면 열 너비 비율
  | 머리 | 머리 |           → ("tbl", 머리행, 본문행들, 너비)
  ![[그림 n-n] 캡션](절대경로){너비mm}  → ("fig", 경로, 캡션, 너비)
  # 부록                  → 이후는 APPENDIX. 첫 항목은 ("p1","A. 세부 분석자료 및 원자료")
  <!-- ... -->, 빈 줄      → 무시
사용: python md2content.py 원고1.md 원고2.md ... → content_final.py 생성
"""
import re, sys, os, io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))


def parse(lines):
    body, appx = [], []
    cur = body
    widths = None
    i = 0
    while i < len(lines):
        raw = lines[i].rstrip("\n")
        s = raw.strip()
        i += 1
        if not s:
            continue
        m = re.match(r"<!--\s*widths:\s*([\d.,\s]+)-->", s)
        if m:
            widths = [float(x) for x in m.group(1).split(",") if x.strip()]
            continue
        if s == "<!-- pagebreak -->":   # 이 자리에서 쪽을 넘긴다(팀이 HWP에서 넣은 쪽 나눔을 원고에 기록)
            cur.append(("pb",)); continue
        if s.startswith("<!--"):
            while "-->" not in s and i < len(lines):
                s = lines[i].strip(); i += 1
            continue
        if s == "# 부록":
            cur = appx
            continue
        m = re.match(r"#\s*제(\d)장\s+(.+)$", s)
        if m:
            cur.append(("ch", int(m.group(1)), m.group(2).strip())); continue
        m = re.match(r"##\s*(\d)\.\s+(.+)$", s)
        if m:
            cur.append(("sec", int(m.group(1)), m.group(2).strip())); continue
        if s.startswith("### "):
            cur.append(("p1", s[4:].strip())); continue
        m = re.match(r"!\[(.+?)\]\((.+?)\)(?:\{(\d+)\})?$", s)
        if m:
            path = m.group(2).strip()
            if not os.path.isabs(path):   # 원고의 상대 경로는 이 폴더 기준
                path = os.path.normpath(os.path.join(HERE, path))
            if not os.path.exists(path):
                print("  [경고] 그림 파일 없음:", path)
            cur.append(("fig", path, m.group(1).strip(), int(m.group(3) or 150))); continue
        if s.startswith("[표 "):
            cur.append(("cap", s)); continue
        if s.startswith("|"):
            rows = [s]
            while i < len(lines) and lines[i].strip().startswith("|"):
                rows.append(lines[i].strip()); i += 1
            cells = [[c.strip() for c in r.strip("|").split("|")] for r in rows]
            cells = [r for r in cells if not all(re.fullmatch(r":?-{2,}:?", c) for c in r)]
            head, data = cells[0], cells[1:]
            n = len(head)
            for r in data:
                if len(r) != n:
                    print("  [경고] 열 수 불일치:", r[:2], len(r), "≠", n)
            cur.append(("tbl", head, data, widths) if widths else ("tbl", head, data)); widths = None
            continue
        if s.startswith("- "):
            cur.append(("p3", s[2:].strip())); continue
        cur.append(("p2", s))
    return body, appx


def main(files):
    lines = []
    for f in files:
        with open(f, encoding="utf-8") as fh:
            lines += fh.readlines() + ["\n"]
    body, appx = parse(lines)
    out = os.path.join(HERE, "content_final.py")
    with open(out, "w", encoding="utf-8") as fh:
        fh.write("# -*- coding: utf-8 -*-\n# 자동 생성: md2content.py — 직접 고치지 말고 원고 md를 고친 뒤 다시 만든다\n")
        fh.write("BODY = " + repr(body) + "\n")
        fh.write("APPENDIX = " + repr(appx) + "\n")
    kinds = {}
    for it in body + appx:
        kinds[it[0]] = kinds.get(it[0], 0) + 1
    chars = sum(len(it[1]) for it in body if it[0] in ("p1", "p2", "p3"))
    print("저장:", out, "| 항목", kinds, "| 본문 글자 수", chars)


if __name__ == "__main__":
    main(sys.argv[1:])
