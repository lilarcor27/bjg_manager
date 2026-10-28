"""Update last year's budget survey report with this year's survey results.

    python3 tools/build_survey_report.py

Input 붙임1 is the county's tally workbook (sheet '설문조사 결과' with one
'응답항목/응답수/비율' block per question, sheet '설문조사 DB' with one row
per respondent, column AD = where the form came from). 붙임2 is the 2026
report; its numbers, sentences and tables are rewritten from 붙임1. The
charts in 붙임2 are pasted Excel pictures (WMF) that cannot be redrawn
from the workbook here, so each becomes a same-size red box telling the
county which chart in 붙임1 to paste there (담당자 요청, 2026-10-28).
"""
import os
import re
import sys
import tempfile
import warnings
import zipfile
from decimal import ROUND_HALF_UP, Decimal

import openpyxl
from PIL import Image, ImageDraw, ImageFont

sys.path.insert(0, os.path.dirname(__file__))
from hwpx_fill import HC, Hwpx, _q, stamped  # noqa: E402

warnings.filterwarnings("ignore")
SRC = "docs/자료/2027 예산편성 설문조사/"
XLSX = SRC + "2027 예산편성 설문조사 결과(취합).xlsx"
LAST = SRC + "2026년도 예산편성 관련 주민설문조사 결과(작년 보고서).hwpx"
BASE = "docs/작성본/2027년도 예산편성 관련 주민설문조사 결과"
FONT = "/usr/local/lib/python3.13/dist-packages/koreanize_matplotlib/fonts/NanumGothicBold.ttf"

# Body paragraph of each chart in 붙임2, in the workbook's question order (block 4..24).
CHART_PARAS = [28, 31, 33, 36, 38, 41, 43, 46, 48, 51, 53, 56, 58, 61, 63, 66, 68, 71, 73, 76, 78]


def pct(n, total):
    return "%.1f%%" % (Decimal(n) * 100 / Decimal(total)).quantize(Decimal("0.1"), ROUND_HALF_UP)


def load_blocks(path):
    ws = openpyxl.load_workbook(path, data_only=True)["설문조사 결과"]
    out, cur = [], None
    for r in range(1, ws.max_row + 1):
        c, d = ws.cell(r, 3).value, ws.cell(r, 4).value
        if c == "응답항목":
            cur = {"row": r, "total": None, "items": []}
            out.append(cur)
        elif cur is not None and c is not None and d is not None:
            if str(c).replace(" ", "") == "합계":
                cur["total"] = d
            else:
                cur["items"].append((" ".join(str(c).split()), d))
        elif cur is not None and cur["items"]:
            cur = None
    for b in out:
        assert sum(n for _, n in b["items"]) == b["total"], b
    return out


def chart_rows(path):
    """(first, last) sheet row of every chart on '설문조사 결과', 1-based."""
    with zipfile.ZipFile(path) as z:
        xml = z.read("xl/drawings/drawing1.xml").decode("utf-8")
    rows = []
    for a in re.findall(r"<xdr:twoCellAnchor.*?</xdr:twoCellAnchor>", xml, re.S):
        if "graphicFrame" in a:
            f = int(re.search(r"<xdr:from>.*?<xdr:row>(\d+)</xdr:row>", a, re.S).group(1))
            t = int(re.search(r"<xdr:to>.*?<xdr:row>(\d+)</xdr:row>", a, re.S).group(1))
            rows.append((f + 1, t + 1))
    return sorted(rows)


def respondents(path):
    ws = openpyxl.load_workbook(path, data_only=True)["설문조사 DB"]
    web = paper = 0
    suggestions = 0
    for r in ws.iter_rows(min_row=17, values_only=True):
        if not any(v not in (None, "") for v in r[1:29]):
            continue  # blank form
        if r[29] == "홈페이지":
            web += 1
        else:
            paper += 1
        if r[28] not in (None, ""):
            suggestions += 1
    return web, paper, suggestions


B = load_blocks(XLSX)
WEB, PAPER, SUGGEST = respondents(XLSX)
TOTAL = WEB + PAPER


def find(b, key):
    hits = [(k, n) for k, n in B[b]["items"] if key in k.replace("∙", "·")]
    assert len(hits) == 1, (b, key, hits)
    return hits[0][1]


def fmt(text):
    """'{8:시설물:p}' -> share of that answer in question 8, ':n' -> count."""
    def sub(m):
        b, key, kind = int(m.group(1)), m.group(2), m.group(3)
        n = find(b, key)
        return pct(n, B[b]["total"]) if kind == "p" else str(n)
    return re.sub(r"\{(\d+):([^:}]+):([pn])\}", sub, text)


def ranked(b):
    return [k for k, _ in sorted(B[b]["items"], key=lambda x: -x[1])]


def expect_order(b, *keys):
    """Guard the hand-written sentences: they name answers in rank order."""
    top = ranked(b)[: len(keys)]
    assert all(k in t.replace("∙", "·") for k, t in zip(keys, top)), (b, keys, top)


def runs_text(p, texts):
    """Set the text of each run that holds text, in order; None keeps a run."""
    runs = [r for r in p.findall(_q("run")) if r.find(_q("t")) is not None]
    assert len(runs) >= len(texts), (len(runs), texts)
    for r, text in zip(runs, texts):
        if text is not None:
            ts = r.findall(_q("t"))
            ts[0].text = text
            for t in ts[1:]:
                r.remove(t)
    for ls in p.findall(_q("linesegarray")):
        p.remove(ls)


def placeholder(path, w, h, lines):
    """Red dashed box, same shape as the chart it stands in for."""
    W = 1400
    H = round(W * h / w)
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)
    red = (204, 0, 0)
    for x in range(10, W - 10, 28):
        d.line([(x, 10), (min(x + 16, W - 10), 10)], fill=red, width=5)
        d.line([(x, H - 11), (min(x + 16, W - 10), H - 11)], fill=red, width=5)
    for y in range(10, H - 10, 28):
        d.line([(10, y), (10, min(y + 16, H - 10))], fill=red, width=5)
        d.line([(W - 11, y), (W - 11, min(y + 16, H - 10))], fill=red, width=5)
    big, small = ImageFont.truetype(FONT, 64), ImageFont.truetype(FONT, 44)
    sizes = [big] + [small] * (len(lines) - 1)
    heights = [f.getbbox(t)[3] + 30 for f, t in zip(sizes, lines)]
    y = (H - sum(heights)) // 2
    for f, t, lh in zip(sizes, lines, heights):
        d.text(((W - d.textlength(t, font=f)) // 2, y), t, font=f, fill=red if f is big else (60, 60, 60))
        y += lh
    img.save(path)


d = Hwpx(LAST)
P = list(d.root)
HP_T = _q("t")

# ---- 1. 총괄 -----------------------------------------------------------------
runs_text(P[0], ["2027년도 예산편성 관련 주민설문조사 결과"])
runs_text(P[3], [None, "총 %d명" % TOTAL, "의 주민이 참여(온라인 %d명, 서면조사 %d명)" % (WEB, PAPER)])
runs_text(P[4], [None, None, " 2027년도 재정운용에 관한 설문조사로"])
expect_order(7, "국토", "산업", "문화", "공공질서", "사회복지")
runs_text(P[6], [None, None, None, None, fmt("국토및지역개발 {7:국토:p}"),
                 fmt(", 산업·중소기업 및 에너지 {7:산업:p}, 문화및관광 {7:문화:p}, "
                     "공공질서및안전 {7:공공질서:p}, 사회복지 {7:사회복지:p} 순으로 나타남")])

# 분야별 1순위 사업 (block 8~17)
tops = d.table_with("- 공공질서및안전분야")
for row, b in enumerate(range(8, 18)):
    k, n = sorted(B[b]["items"], key=lambda x: -x[1])[0]
    name = re.sub(r"^[①-⑩]\s*", "", k)
    d.set_cell(tops, 2, row, "%s(%s)" % (name, pct(n, B[b]["total"])))

expect_order(18, "로컬푸드", "상하수도", "장계")
runs_text(P[10], [None, None, fmt("로컬푸드 등 향토산업 육성({18:로컬푸드:p}),")])
runs_text(P[11], [fmt("       상하수도 확충 및 현대화사업({18:상하수도:p}), 장계관광지 정비 사업({18:장계:p})")])
expect_order(23, "행사", "사회단체")
runs_text(P[13], [None, None, fmt("행사·축제성 경비({23:행사:p})"), None,
                  fmt("사회단체보조 등 민간지원경비({23:사회단체:p}) ")])
expect_order(24, "일자리", "문화관광", "인재", "신재생", "대규모")
runs_text(P[15], [None, fmt("중소기업 유치지원 등 일자리 창출({24:일자리:p})"), None,
                  fmt(" 문화관광 인프라 확충({24:문화관광:p}), 차별화된 인재양성({24:인재:p}), "
                      "신재생에너지 등 녹색성장 및 신성장 동력({24:신재생:p}), "
                      "대규모 산업단지 및 농업단지 조성({24:대규모:p}) 순으로 나타남.")])
runs_text(P[16], [None, None, None, "건의 및 제안사항에는 %d명" % SUGGEST,
                  "의 응답자가 의견을 제시하였으며, 예산편성과 관련된 사항은 2028년도 예산편성 시 참고하여 반영토록 추진."])

# ---- 2. 기본사항 -------------------------------------------------------------
runs_text(P[19], [None, "총 응답자 %d명 중," % TOTAL])
expect_order(0, "여", "남")
runs_text(P[20], [None, fmt("여성이 {0:여:n}명({0:여:p})"), fmt(", 남성이 {0:남:n}명({0:남:p})이며,")])
expect_order(1, "60대")
runs_text(P[21], [None, fmt("연령별로 19세 이하 {1:19세:n}명({1:19세:p}), 20대 {1:20대:n}명({1:20대:p}), "
                            "30대 {1:30대:n}명({1:30대:p}), 40대 {1:40대:n}명({1:40대:p}), "
                            "50대 {1:50대:n}명({1:50대:p}), "),
                  fmt("60대 이상 {1:60대:n}명({1:60대:p})"), None])
expect_order(2, "옥천읍", "안내면", "군북면")
runs_text(P[22], [None, None, fmt("옥천읍이 {2:옥천읍:n}명({2:옥천읍:p})"),
                  fmt("으로 가장 많았으며, 안내면 {2:안내면:n}명({2:안내면:p}), 군북면 {2:군북면:n}명({2:군북면:p}) 순임.")])
expect_order(3, "농", "자영업", "가정주부")
runs_text(P[23], [None, None, fmt("농·축·어업 {3:농:n}명({3:농:p})"),
                  fmt(", 자영업 {3:자영업:n}명({3:자영업:p}), 가정주부 {3:가정주부:n}명({3:가정주부:p}) 순으로 응답함.")])

basic = d.table_with("성 별")
for col_label, col_n, col_p, b, keys in [
        (1, 2, 3, 0, {"남": "남", "여": "여"}),
        (5, 6, 7, 1, {"19세 이하": "19세", "20대": "20대", "30대": "30대", "40대": "40대",
                      "50대": "50대", "60대 이상": "60대"}),
        (1, 2, 3, 2, {e: e for e in ["옥천읍", "동이면", "안남면", "안내면", "청성면", "청산면",
                                      "이원면", "군서면", "군북면"]} | {"관  외": "관"}),
        (5, 6, 7, 3, {"농축어업": "농", "가정주부": "가정주부", "자영업": "자영업", "회사원": "회사원",
                      "전문직": "전문직", "학  생": "학생", "공무원": "공무원", "기타": "기"})]:
    cells = {t: (c, r) for c, r, t in d.dump(basic)}
    for label, key in keys.items():
        _, row = cells[label]
        n = find(b, key) if key != "기" else dict(B[b]["items"])["⑧ 기 타"]
        d.set_cell(basic, col_n, row, str(n))
        d.set_cell(basic, col_p, row, pct(n, B[b]["total"]))
# 소계 rows: 성별 row 1 (and its repeat at row 7), 연령 row 1, 직업 row 14, 거주지 has none.
for col_n, col_p, row, b in [(2, 3, 1, 0), (2, 3, 7, 0), (6, 7, 1, 1), (6, 7, 14, 3)]:
    d.set_cell(basic, col_n, row, str(B[b]["total"]))
    d.set_cell(basic, col_p, row, "100.0%")

# ---- 3. 문항별 표, 설명, 차트 자리 ---------------------------------------------
COMMENT = {
    28: [None, "☞ 이 항목에서는 군민들이 우선적으로 생각하는 재원배분 항목이 무엇인지 볼 수 있는데, 응답자의 "
               "{4:경기:p}인 {4:경기:n}명이 지역경기활성화를 현시점 우리 군민들이 상대적으로 중시하는 것으로 나타남."],
    31: ["☞ 우리군의 예산 운영에 관해서는, 응답자의 {5:보통:p}인 {5:보통:n}명이 ‘보통이다’, "
         "{5:② 잘하고:p}인 {5:② 잘하고:n}명이 ‘잘하고 있다’고 응답해 군민들은 양호한 평가를 함.",
         "☞ 응답자의 {5:모르:p}가 잘 모르겠다고 응답해 예산 운영 상황에 대한 홍보가 필요한 것으로 보임. ",
         "☞ 잘못하고 있다는 {5:잘못:p}임. "],
    33: [None, "☞ 응답자의 {6:국·도비:p}인 {6:국·도비:n}명이 국·도비 등 중앙 및 광역지자체 지원예산 확보로 "
               "군비 부담 경감이 필요하다고 응답함.",
         "☞ 경상경비 절감 등 예산절감 노력이 {6:경상경비:p}",
         "☞ 탈루세원 발굴, 체납액 징수 등 자주재원 확충이 {6:탈루:p} 순으로 나타남."],
    36: [None, "☞ 국토및지역개발이 {7:국토:p}로 가장 높았으며, ",
         "☞ 산업·중소기업 및 에너지 {7:산업:p}, 문화및관광 {7:문화:p}, 공공질서및안전 {7:공공질서:p}, "
         "사회복지 {7:사회복지:p} 순으로 나타남."],
    38: ["☞ 각종 재난재해 발생으로 관심이 증대된 공공질서 및 안전분야에서는 우리 군민들의 {8:시설물:p}인 "
         "{8:시설물:n}명이 시설물 안전관리점검 강화가 필요하다고 응답함.",
         "☞ 재해위험지구 정비 {8:재해위험:p}, ", "☞ 재난 예·경보시설 확충 {8:예·경보:p} 순으로 나타남."],
    41: ["☞ 응답자의 {9:평생학습:p}인 {9:평생학습:n}명이 평생학습 프로그램 지원 강화가 필요하다고 응답함.",
         "☞ 인재육성을 위한 고등교육 재정지원 {9:인재:p}", "☞ 학교무상급식지원 확대 {9:무상급식:p} 순으로 나타남."],
    43: ["☞ 응답자의 {10:읍면소재지:p}인 {10:읍면소재지:n}명이 기초생활인프라 확충을 위한 읍면소재지 정비가 "
         "필요하다고 응답함.",
         "☞ 도시계획도로 개설을 통한 도시환경정비 {10:도시계획:p}", "☞ 주거복지 실현을 위한 주택건설 {10:주거:p} 순으로 나타남."],
    46: ["☞ 응답자의 {11:관광지:p}인 {11:관광지:n}명이 관광지를 개발해야 한다고 응답함.",
         "☞ 차별화된 축제 및 행사 개발 {11:축제:p}", "☞ 체육시설확충 및 운영 활성화 {11:체육:p} 순으로 나타남.", None],
    48: ["☞ 응답자의 {12:폐기물:p}인 {12:폐기물:n}명이 폐기물의 안정적 처리시설을 확충해야 한다고 응답함.",
         "☞ 자연생태계 보전 및 자연공원 관리 {12:자연생태계:p},", "☞ 안심하고 마실 수 있는 상수도 시설 확충 {12:상수도:p},",
         "☞ 완벽한 하수처리를 통한 수질환경 개선, 악취저감 등 대기환경 개선 순으로 관심이 높은 것으로 나타남."],
    51: ["☞ 응답자의 {13:노인:p}인 {13:노인:n}명이 노인생활안정을 위한 노인일자리 지원을 위한 사업을 추진해야 한다고 응답함.",
         "☞ 저소득, 차상위계층 등 취약계층 기초생활안정 {13:저소득:p}",
         "☞ 여성인력 양성 등 여성의 권익증진 {13:여성:p} 순으로 높게 나타남."],
    53: ["☞ 응답자의 {14:공공의료:p}인 {14:공공의료:n}명이 공공의료기관 의료서비스를 강화해야 한다고 응답함.",
         "☞ 건강증진 프로그램 다양화 {14:건강증진:p}, 공중보건 및 위생에 관한 관리감독 강화 {14:공중보건:p} 순으로,", None],
    56: ["☞ 응답자의 {15:농촌마을:p}인 {15:농촌마을:n}명이 농촌마을종합개발(권역단위종합정비)가 필요하다고 응답함. ",
         "☞ 과수, 채소 등 경쟁우위품목 육성지원 {15:과수:p}, 농산물 유통구조 개선 {15:유통:p}, "
         "친환경농산물 생산 및 판매 지원 {15:친환경:p} 순으로 나타남."],
    58: ["☞ 응답자의 {16:④ 일자리:p}인 {16:④ 일자리:n}명이 일자리 창출이 필요하다고 응답함.",
         "☞ 다음으로 기업유치 및 산업단지 조성이 {16:기업유치:p}로 나타나,", None],
    61: ["☞ 응답자의 {17:위험도로:p}인 {17:위험도로:n}명이 위험도로 및 노후도로 구조개선이 필요하다고 응답했으며, ",
         "☞ 대중교통망 및 교통편의시설 확충 {17:대중교통:p}, 주차장 확보 및 시내 주정차 단속 강화 {17:주차장:p}, "
         "군도, 농어촌도로 등 도로망 확충 {17:군도:p} 순으로 나타남."],
    63: ["☞ 응답자의 {18:로컬푸드:p}와 {18:상하수도:p}가 각각 로컬푸드 등 향토산업 육성과 상하수도 확충 및 "
         "현대화사업에 우선적으로 투자해야 한다고 응답함.",
         "☞ 장계관광지 정비사업({18:장계:p}), 소상공인·자영업 육성({18:소상공인:p}) 순으로 나타남."],
    66: ["☞ 응답자의 {19:현수준:p}가 현 수준을 유지해야 한다고 응답함.",
         "☞ 차츰 감소 지원은 {19:차츰:p}, 지속 증가 지원은 {19:지속:p}로 나타남"],
    68: ["☞ 응답자의 {20:50%이상:p}인 {20:50%이상:n}명이 보조율 50%이상~70%이하가 적정하다고 응답함.",
         "☞ 보조율 70%이상 {20:70% 이상:p}, ", "☞ 보조율 30%이상~50%이하 {20:30%이상:p} 순으로 적정하다고 답함."],
    71: ["☞ 응답자의 {21:50%이상:p}인 {21:50%이상:n}명이 보조율 50%이상~70% 이하가 적정하다고 응답함.",
         "☞ 보조율 70% 이상이 {21:70% 이상:p},", "☞ 보조율 30%이상~50% 이하가 적정하다는 응답 {21:30%이상:p} 순으로 나타남."],
    73: ["☞ 응답자의 {22:공공목적:p}인 {22:공공목적:n}명이 공공목적(농로 등)이 우선 고려되어야 한다고 응답함.",
         "☞ 재해예방(공공,개인) {22:재해예방:p}, ", "☞ 편의시설(정자,운동기구 등) {22:편의시설:p} 순으로 나타남."],
    76: ["☞ 응답자의 {23:행사:p}인 {23:행사:n}명이 행사·축제성 경비 축소가 필요하다고 응답함.",
         "☞ 다음으로 사회단체보조 등 민간지원경비({23:사회단체:p}),",
         "☞ 농로 확·포장 등 소규모 민원해결성 사업({23:농로:p}) 순으로 축소가 필요하다고 응답함."],
    78: ["☞ 응답자의 {24:일자리:p}인 {24:일자리:n}명이 중소기업 유치지원 등 일자리 창출 확대가 필요하다고 응답함.",
         "☞ 다음으로 문화관광 인프라 확충 {24:문화관광:p}, ",
         "☞ 차별화된 인재양성 {24:인재:p}, 신재생에너지 등 녹색성장 및 신성장 동력 {24:신재생:p} 순으로 확대가 필요하다고 답함."],
}
ORDER = {4: ["경기"], 5: ["보통", "잘하고", "모르"], 6: ["국", "경상", "탈루"], 8: ["시설물", "재해위험", "예·경보"],
         9: ["평생", "인재", "무상"], 10: ["읍면", "도시계획", "주거"], 11: ["관광지", "축제", "체육"],
         12: ["폐기물", "자연", "상수도"], 13: ["노인", "저소득", "여성"], 14: ["공공의료", "건강", "공중"],
         15: ["농촌", "과수", "유통", "친환경"], 16: ["일자리", "기업"], 17: ["위험", "대중", "주차", "군도"],
         18: ["로컬", "상하수도", "장계", "소상공인"], 19: ["현수준", "차츰", "지속"], 20: ["50%", "70%", "30%이상"],
         21: ["50%", "70%", "30%이상"], 22: ["공공목적", "재해", "편의"], 23: ["행사", "사회단체", "농로"],
         24: ["일자리", "문화관광", "인재", "신재생"]}
for b, keys in ORDER.items():
    expect_order(b, *keys)

charts = chart_rows(XLSX)
tmp = tempfile.mkdtemp()
for b, pi in zip(range(4, 25), CHART_PARAS):
    outer = P[pi].find(".//" + _q("tbl"))
    inner = outer.find(".//" + _q("tc")).find(".//" + _q("tbl"))
    blk = B[b]
    # 표: 3열(응답항목/응답수/비율) 또는 투자분야처럼 6열(①~⑤ | ⑥~⑩)
    cells = d.dump(inner)
    ncol = 1 + max(c for c, _, _ in cells)
    slots = [(0, r) for r in range(2, 1 + max(r for _, r, _ in cells))] if ncol == 3 else \
        [(0, r) for r in range(2, 7)] + [(3, r) for r in range(1, 6)]
    assert len(slots) == len(blk["items"]), (b, len(slots), blk["items"])
    d.set_cell(inner, 1, 1, str(blk["total"]))
    d.set_cell(inner, 2, 1, "100.0%")
    for (c, r), (k, n) in zip(slots, blk["items"]):
        old = d.cell(inner, c, r).xpath("string(.)").strip()
        assert old[:1] == k[:1], (b, old, k)  # same circled number
        d.set_cell(inner, c + 1, r, str(n))
        d.set_cell(inner, c + 2, r, pct(n, blk["total"]))
    # 설명
    note = d.cell(outer, 1, 0)
    paras = note.find(_q("subList")).findall(_q("p"))
    lines = COMMENT[pi]
    assert len(paras) == len(lines), (pi, len(paras), len(lines))
    for p, line in zip(paras, lines):
        if line is not None:
            runs_text(p, [fmt(line)] + [""] * (len([r for r in p.findall(_q("run")) if r.find(HP_T) is not None]) - 1))
    # 차트 자리
    first, last = next((f, t) for f, t in charts if f > blk["row"])
    pic = next(outer.iter(_q("pic")))
    ref = pic.find(".//{%s}img" % HC).get("binaryItemIDRef")
    cur = pic.find(_q("curSz"))
    png = os.path.join(tmp, ref + ".png")
    placeholder(png, int(cur.get("width")), int(cur.get("height")),
                ["【차트 붙여넣기】", "붙임1 '설문조사 결과' 시트 %d~%d행 차트를" % (first, last),
                 "복사해서 이 박스 대신 붙여넣어 주세요"])
    d.swap_image(ref, png)

# ---- 4. 건의사항 ---------------------------------------------------------------
runs_text(P[82], [None, "설문조사에서는 %d명의 응답자가 건의사항 및 제안사항을 작성하였음." % SUGGEST])
runs_text(P[83], ["   ▷ 내용을 보면, 도로·교통·주차, 농업·농민 지원, 문화·행사 관련 의견과 출렁다리·박물관 등 "
                  "대형사업 재검토 의견이 주를 이루고 있음."])

os.makedirs(os.path.dirname(BASE), exist_ok=True)
out = stamped(BASE)
d.save(out)
print("saved", out, "| 응답자 %d명(온라인 %d, 서면 %d), 건의 %d건" % (TOTAL, WEB, PAPER, SUGGEST))
