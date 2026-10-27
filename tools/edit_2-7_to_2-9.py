"""Fill indicator 2-7 and raise 2-8 to general residents + body (content only).

    python3 tools/edit_2-7_to_2-9.py [input.hwpx]

Answers from the county (2026-10-27): projects rated '매우미흡' in the
major fiscal project post-evaluation were reflected in the budget; the
priority-decision shares follow the 2026 budget (resident-council type
decided by residents' general assemblies 450,000 of 2,740,000 thousand
won = 16%, regional-meeting type 84%). 2-9 was filled in the first draft
and is left as it is.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(__file__))
from hwpx_fill import Hwpx, latest, stamped  # noqa: E402

BASE = "docs/작성본/1. 종합분야 실적서(옥천군)"
CHECK = "【확인 필요】"
d = Hwpx(sys.argv[1] if len(sys.argv) > 1 else latest(BASE))
T = list(d.tables())
put = d.set_cell

# ---- 2-7 결산·평가 환류: 주요재정사업 사후평가 결과 반영 ----------------------
put(T[114], 1, 1, "○")
put(T[114], 2, 1, "")
put(T[115], 0, 1, CHECK + "\n사후평가 '매우미흡' 사업명")
put(T[115], 1, 1,
    "○ 주요재정사업 사후평가(지방보조금관리위원회, 참여예산주민위원 참여 " + CHECK + ")에서 "
    "사업 성과 '매우미흡' 판정 및 개선 의견")
put(T[115], 2, 1, "수용")
put(T[115], 3, 1,
    "○ ’27년 예산편성 시 평가결과 반영(감액·폐지·사업방식 변경 등)\n"
    "  - 사업별 조치 내용 " + CHECK)

# ---- 2-8 우선순위 결정권: 일반주민(주민총회) + 기구(지역회의·분과) --------------
put(T[121], 1, 2, "○")
put(T[121], 2, 2, "○")
put(T[121], 3, 2, "")
put(T[121], 1, 3, "16%")
put(T[121], 2, 3, "84%")
put(T[121], 3, 3, "0%")
put(T[122], 1, 1,
    "○ 읍면 주민자치회가 발굴한 사업을 주민총회에서 주민 투표로 선정·우선순위 결정 → "
    "읍면별 한도(5천만원) 내 예산 편성")
put(T[122], 2, 1, "주민자치회형 사업 선정·우선순위 결정\n(’26년 450,000천원, 16%)")
put(T[122], 2, 2,
    "중점 투자사업에 대한 우선순위 결정 및 의견 제출(운영조례 제9조)\n"
    "(지역회의형 ’26년 2,290,000천원, 84%)")

out = stamped(BASE)
d.save(out)
print("saved", out)
