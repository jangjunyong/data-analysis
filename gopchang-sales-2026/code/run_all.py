# -*- coding: utf-8 -*-
"""꿀꿀곱창 최종 분석보고서: 분석·그림 → 원고 변환 → HWP·PDF 빌드를 한 번에 실행한다.

사용(레포 어디서든):
    python -X utf8 최종제출_북적북적팀/분석코드/run_all.py            # 분석 + 보고서 빌드(Windows + 한글 필요)
    python -X utf8 최종제출_북적북적팀/분석코드/run_all.py --no-build # 분석·그림만
    python -X utf8 최종제출_북적북적팀/분석코드/run_all.py --online   # 네이버·카카오 공개 페이지를 다시 조회(W4-01)부터

빌드 결과는 보고서/출력/ 에 저장된다. 제출본(최종제출_북적북적팀/*.hwp, *.pdf)은 빌드 뒤 HWP에서 마지막으로
고친 판이라 빌드로 덮어쓰지 않는다.
"""
import subprocess, sys, time
from pathlib import Path

HERE = Path(__file__).resolve().parent
ONLINE = [("W4_온라인", ["w4_01_온라인정보_조회.py", "--refresh"])]
ANALYSIS = [
    ("W1_매출", ["w1_8_금액그림.py"]),               # 그림 1-2, 2-1
    ("W1_매출", ["w1_6_일별추이_요일시간대.py"]),      # 그림 2-5
    ("W1_매출", ["w1_9_홀배달_상관.py"]),             # 홀×배달 잔차 상관(2장 3절 (5))
    ("W2_요일", ["w2_analysis.py"]),                 # 그림 2-6, 요일·말씀 확인 표
    ("W3_배달세트", ["w3_분석.py"]),                  # 그림 2-8, 3-1, 세트 계산
    ("W4_온라인", ["w4_02_분석_그림.py"]),            # 그림 2-7, 2-9~2-12
]
BUILD = [
    ("보고서", ["md2content.py", "원고_1_2장.md", "원고_3_5장_부록.md"]),
    ("보고서", ["build_final.py"]),                  # 한글(HWP) 자동화, 약 5분
]


def run(folder, args):
    t = time.time()
    print(f"\n▶ {folder}/{args[0]}", flush=True)
    subprocess.run([sys.executable, "-X", "utf8", *args], cwd=HERE / folder, check=True)
    print(f"  완료 {time.time() - t:.0f}초", flush=True)


if __name__ == "__main__":
    steps = (ONLINE if "--online" in sys.argv else []) + ANALYSIS + ([] if "--no-build" in sys.argv else BUILD)
    for folder, args in steps:
        run(folder, args)
    print("\n전체 완료" + ("" if "--no-build" in sys.argv else f": {HERE / '보고서' / '출력'}"))
