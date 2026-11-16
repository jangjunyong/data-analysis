# -*- coding: utf-8 -*-
"""W4-① 온라인 정보 재조회(읽기 전용 공개 페이지): 네이버 플레이스·카카오맵·네이버 지도 검색 순위.
로그인·글쓰기 없음. 요청 사이 2.5초. 결과는 이 폴더의 json/csv로 저장.
실행: python -X utf8 -I w4_01_온라인정보_조회.py [--refresh]
"""
import json, re, sys, time, csv, urllib.request, urllib.parse, datetime
from pathlib import Path

OUT = Path(__file__).resolve().parent
CACHE = OUT / "_cache_html"; CACHE.mkdir(exist_ok=True)
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
NAVER_ID, KAKAO_ID = "2059104960", "358993799"
KEYWORDS = ["명일역 곱창", "명일동 곱창", "명일역 맛집", "명일역 술집", "강동 야채곱창"]
REFRESH = "--refresh" in sys.argv


def get(url, path, headers=None):
    if path.exists() and not REFRESH:
        return path.read_bytes()
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept-Language": "ko-KR,ko;q=0.9", **(headers or {})})
    data = urllib.request.urlopen(req, timeout=30).read()
    path.write_bytes(data); time.sleep(2.5)
    return data


def apollo(html):
    m = re.search(r'window\.__APOLLO_STATE__\s*=\s*(\{.*?\});\s*\n', html, re.S) or \
        re.search(r'window\.__APOLLO_STATE__\s*=\s*(\{.*?\})\s*;?\s*</script>', html, re.S)
    return json.loads(m.group(1)) if m else {}


def naver_place():
    html = get(f"https://pcmap.place.naver.com/restaurant/{NAVER_ID}/home", CACHE / "naver_home.html").decode("utf-8", "replace")
    js = apollo(html)
    base = js[f"PlaceDetailBase:{NAVER_ID}"]
    stats = js[f"VisitorReviewStatsResult:{NAVER_ID}"]
    menus = [dict(name=v["name"], price=v["price"]["displayText"], repr="repr" in (v.get("badges") or []), desc=v.get("description"))
             for k, v in js.items() if k.startswith("PlaceMenuItem:")]
    # 영업시간은 APOLLO가 아니라 페이지 상태 JSON에 있다 → 정규식으로 WorkingHoursInfo를 꺼낸다
    hours = re.findall(r'"day":"([월화수목금토일])","businessHours":(null|\{[^}]*\}),"breakHours":[^,]*,"description":(null|"[^"]*")', html)
    hours = [(d, (json.loads(b) if b != "null" else None), (json.loads(desc) if desc != "null" else None)) for d, b, desc in hours]
    ig = re.search(r'https://www\.instagram\.com/[A-Za-z0-9_.]+', html)
    corkage = re.search(r'"alcoholTypeName":"([^"]*)","policyDescription":"([^"]*)"', html)
    order = [dict(type=v["type"], available=v["isOrderAvailable"]) for k, v in js.items() if k.startswith("NaverOrderItem:")]
    vk = stats["analysis"]["votedKeyword"]
    rv_html = get(f"https://pcmap.place.naver.com/restaurant/{NAVER_ID}/review/visitor", CACHE / "naver_review.html").decode("utf-8", "replace")
    rjs = apollo(rv_html)
    latest = [dict(created=v.get("created"), visited=v.get("visited"), rating=v.get("rating"),
                   reply=bool((v.get("reply") or {}).get("body"))) for k, v in rjs.items() if k.startswith("VisitorReview:") and "created" in v]
    return dict(
        조회일시=datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), 이름=base["name"], 카테고리=base["category"],
        방문자리뷰수=base["visitorReviewsTotal"], 별점=base["visitorReviewsScore"], 블로그카페리뷰수=base.get("cafeBlogReviewsTotal"),
        편의=base["conveniences"], 결제=base["paymentInfo"], 한줄소개=base["microReviews"],
        영업시간=[dict(요일=d, 시간=(f"{b['start']}~{b['end']}" if b else None), 비고=desc) for d, b, desc in hours],
        메뉴=menus, 인스타링크=ig.group(0) if ig else None,
        콜키지=(corkage.group(1), corkage.group(2)) if corkage else None, 네이버주문=order,
        키워드투표=dict(총표=vk.get("totalCount"), 참여자=vk.get("userCount"),
                   상위=[(d["displayName"], d["count"]) for d in vk.get("details", [])[:10]]),
        리뷰테마=[(t["label"], t["count"]) for t in stats["analysis"]["themes"]],
        최근리뷰10건=latest,
    )


def kakao_place():
    raw = get(f"https://place-api.map.kakao.com/places/panel3/{KAKAO_ID}", CACHE / "kakao_panel.json",
              headers={"pf": "web", "Accept": "application/json", "Referer": "https://place.map.kakao.com/", "Origin": "https://place.map.kakao.com"})
    d = json.loads(raw)
    wk = []
    for p in d["open_hours"]["week_from_today"]["week_periods"]:
        for day in p["days"]:
            wk.append((day["day_of_the_week_desc"], day.get("on_days", {}).get("start_end_time_desc") or day.get("off_days_desc")))
    return dict(
        조회일시=datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), 이름=d["summary"]["name"], 카테고리=d["summary"]["category"]["name"],
        정보갱신일=d["summary"]["meta"].get("updated_at"), 영업시간=wk, 영업상태=d["open_hours"]["headline"],
        메뉴=[(m["name"], m["price"], m.get("mod_at")) for m in d["menu"]["menus"]["items"]],
        홈페이지=d["summary"].get("homepages"), 결제=[p["image_keyword"] for p in d["summary"]["payments"]],
        후기수=d["kakaomap_review"]["score_set"]["review_count"], 후기평점=d["kakaomap_review"]["score_set"]["average_score"],
        블로그리뷰수=d["blog_review"]["review_count"], 사진수=d["photos"]["counts"],
        AI요약=d["ai_mate"]["bottom_sheet"].get("summary"), AI추천메뉴=d["ai_mate"]["bottom_sheet"]["sheets"][0].get("keywords"),
        편의=[i["text"] for i in d["place_add_info"]["ai_mate"]["store_facility_icons"]],
        좌석=[i["text"] for i in d["place_add_info"]["simple_detail_infos"]],
        지하철=d["find_way"]["subway"], 대표사진출처=("블로그 사진" if "pstatic" in (d["summary"].get("main_photo_url") or "") else "기타"),
    )


def naver_rank():
    rows = []
    for i, q in enumerate(KEYWORDS, 1):
        url = f"https://pcmap.place.naver.com/restaurant/list?query={urllib.parse.quote(q)}&x=127.144&y=37.551"
        html = get(url, CACHE / f"nlist_{i}.html").decode("utf-8", "replace")
        js = apollo(html); rq = js.get("ROOT_QUERY", {})
        key = [k for k in rq if k.startswith("placeList(") and '"display":50' in k]
        if not key:
            rows.append(dict(키워드=q, 순위="조회 실패", 상호="", 방문자리뷰="", 블로그="", 거리="")); continue
        items = rq[key[0]]["businesses"]["items"]
        for n, ref in enumerate(items, 1):
            v = js[ref["__ref"]]
            rows.append(dict(키워드=q, 순위=n, 상호=v["name"], 방문자리뷰=v.get("visitorReviewCount"), 블로그=v.get("blogCafeReviewCount"),
                             거리=v.get("distance"), 카테고리=v.get("category"), 꿀꿀=(v["id"] == NAVER_ID)))
    with open(OUT / "검색노출_순위_전체.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys())); w.writeheader(); w.writerows(rows)
    return rows


if __name__ == "__main__":
    res = dict(네이버=naver_place(), 카카오=kakao_place())
    rows = naver_rank()
    res["검색순위"] = {q: dict(꿀꿀순위=next((r["순위"] for r in rows if r["키워드"] == q and r.get("꿀꿀")), "50위 밖"),
                          상위5=[(r["상호"], r["방문자리뷰"]) for r in rows if r["키워드"] == q and isinstance(r["순위"], int) and r["순위"] <= 5],
                          표시수=sum(1 for r in rows if r["키워드"] == q and isinstance(r["순위"], int))) for q in KEYWORDS}
    (OUT / "온라인정보_조회.json").write_text(json.dumps(res, ensure_ascii=False, indent=1, default=str), encoding="utf-8")
    print(json.dumps(res["검색순위"], ensure_ascii=False, indent=1))
    print("네이버 영업시간", res["네이버"]["영업시간"]); print("카카오 영업시간", res["카카오"]["영업시간"], res["카카오"]["정보갱신일"])
    print("네이버 최근리뷰", res["네이버"]["최근리뷰10건"][:3])
