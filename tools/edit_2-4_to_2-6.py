"""Fill indicators 2-4, 2-5 and 2-6 in the working report (content only).

    python3 tools/edit_2-4_to_2-6.py [input.hwpx]

Answers from the county (2026-10-27): resident-council projects were
chosen by township residents' general assemblies; the "내 세금감시단"
(budget watch group) ran again in 2026; residents took part in the
performance analysis. Wording follows the 2025 operating plan and the
2025 budget-waste reporting plan; dates and counts the county has not
given yet are marked 【확인 필요】.
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


def participated(mark):
    put(mark, 1, 1, "○")
    put(mark, 2, 1, "")


# ---- 2-4 주민제안사업 선정: 읍면 주민자치회 주민총회 -----------------------------
participated(T[93])
put(T[94], 0, 1, CHECK + "\n주민총회 개최 시기")
put(T[94], 1, 1, "읍면 주민\n(주민총회 참석 주민)")
put(T[94], 2, 1, "읍면 주민자치회 주민총회 주민 투표")
put(T[94], 3, 1,
    "○ 읍면 주민자치회가 발굴한 주민자치회형(주민참여예산) 사업을 주민총회에서 주민 투표로 선정\n"
    "  - 개최 읍면 " + CHECK + "개, 참여 주민 " + CHECK + "명\n"
    "○ 선정 사업은 읍면별 한도(5천만원) 내에서 예산 편성(’26년 본예산 9개 읍면 450,000천원)")
d.delete_row(T[94], 2)

# ---- 2-5 집행 과정: 내 세금감시단 ---------------------------------------------
participated(T[100])
put(T[101], 0, 1, CHECK + "\n활동 기간")
put(T[101], 1, 1, "내 세금감시단\n(참여예산주민위원 등)")
put(T[101], 2, 1, "예산 집행 감시, 자체 감사 참여 현장점검")
put(T[101], 3, 1,
    "○ 「예산낭비신고제도 운영계획」에 따라 ’내 세금감시단’이 예산 집행 과정의 예산낭비 감시·홍보 활동\n"
    "○ 자체 감사(직속기관·사업소·읍면 행정지도 감사)에 참여하여 예산 집행 실태 점검\n"
    "  - 참여 인원 " + CHECK + "명, 점검 " + CHECK + "회")
d.delete_row(T[101], 2)

# ---- 2-6 결산·평가 과정: 성과분석 ---------------------------------------------
participated(T[107])
put(T[108], 0, 1, CHECK + "\n성과분석 시기")
put(T[108], 1, 1, "주민, 참여예산주민위원회 " + CHECK)
put(T[108], 2, 1, "주민참여예산 성과분석 의견 수렴 " + CHECK + " (성과보고회·의견조사 등)")
put(T[108], 3, 1,
    "○ ’25년 주민참여예산사업 추진성과 분석 및 ’26년 개선방향 모색 과정에 주민·주민참여예산기구 의견 수렴\n"
    "  - 참여 " + CHECK + "명, 주요 의견 " + CHECK)
d.delete_row(T[108], 2)

out = stamped(BASE)
d.save(out)
print("saved", out)
