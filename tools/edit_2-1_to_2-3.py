"""Fill indicators 2-1, 2-2 and 2-3 in the working report (content only).

    python3 tools/edit_2-1_to_2-3.py [input.hwpx]

Answers from the county (2026-10-27): no resident input asked on the
mid-term fiscal plan or the 2026 operating plan; both the '2027 budget
wishes' collection and the budget-direction survey were run. The
subcommittees' budget-direction opinions come from their 2026 result
reports. Dates and counts the county has not given yet are marked.
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

# ---- 2-1 중기지방재정계획: 미참여 ------------------------------------------
put(T[72], 1, 1, "")
put(T[72], 2, 1, "○")
put(T[73], 0, 1, "해당 없음")
for c in (1, 2):
    put(T[73], c, 1, "")
put(T[73], 3, 1, "-")
for c in range(4):
    put(T[73], c, 2, "")

# ---- 2-2 예산편성 기본방향: 주민 참여 ---------------------------------------
put(T[79], 1, 1, "○")
put(T[79], 2, 1, "")
put(T[80], 0, 1, "’26.7.~8월\n" + CHECK + " 기간")
put(T[80], 1, 1, "일반 주민")
put(T[80], 2, 1, "‘2027년도 예산편성에 바란다’ 주민의견 수렴, ‘2027년 예산편성 방향 설문조사’")
put(T[80], 3, 1,
    "○ 군 홈페이지·카카오톡채널·읍면 서면으로 ’27년 예산편성에 대한 주민의견 수렴\n"
    "  - 접수 " + CHECK + "건\n"
    "○ 홈페이지·서면 설문조사로 ’27년 예산편성 방향에 대한 주민 의견 조사\n"
    "  - 응답 " + CHECK + "명")
put(T[80], 0, 2, "’26.10.12.~10.16.")
put(T[80], 1, 2, "참여예산주민위원회\n(3개 분과, 23명)")
put(T[80], 2, 2, "분과위원회 회의에서 ’27년 예산편성 방향 의견 제시")
put(T[80], 3, 2,
    "○ (자치행정) 보조사업 집행 단가 현실화, 체육진흥기금 계획적 적립 확대\n"
    "○ (농산개발) 축종 간 지원 형평성 확대, 성과 부진 사업 추진상황 점검\n"
    "○ (복지문화) 고령화 여건을 고려한 주민 체감형 편성, 경로당 등 이용 실태에 맞는 지원")

# ---- 2-3 운영·활성화 계획: 미참여 ------------------------------------------
put(T[86], 1, 1, "")
put(T[86], 2, 1, "○")
put(T[87], 0, 1, "해당 없음")
for c in (1, 2):
    put(T[87], c, 1, "")
put(T[87], 3, 1, "-")
for c in range(4):
    put(T[87], c, 2, "")

out = stamped(BASE)
d.save(out)
print("saved", out)
