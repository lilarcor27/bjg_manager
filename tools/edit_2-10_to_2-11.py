"""Complete 2-10 (execution and settlement steps) and fill 2-11 (content only).

    python3 tools/edit_2-10_to_2-11.py [input.hwpx]

2-10: the execution and settlement procedures follow the 2025 operating
plan (budget-waste reporting centre, performance analysis) and the
county's answers for 2-5 to 2-7. 2-11: the county says this year's
opinion letter has the same items as 2025's; the 2025 evidence shows 7
of the 9 items (공모 금액 and 의견 수렴방법 not found).
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

# ---- 2-10 집행·결산(평가) 단계 절차 --------------------------------------------
put(T[136], 1, 2, "주민\n내 세금감시단")
put(T[136], 2, 2,
    "○ 예산낭비 신고·예산절감 제안 접수(홈페이지·카카오톡채널·국민신문고) → 소관부서 검토·조사 → "
    "결과 통보 및 예산성과금 지급(예산낭비신고센터, 연중)")
put(T[136], 3, 2,
    "○ 내 세금감시단 감시활동 및 주민 신고·제안 검토·반영\n  - 접수·반영 " + CHECK + "건")
put(T[136], 1, 3, "주민\n참여예산주민위원회\n지방보조금관리위원회")
put(T[136], 2, 3,
    "○ 주민참여예산 성과분석(주민·기구 의견 수렴) → 개선방향 마련 → 다음 연도 운영계획·예산편성 반영\n"
    "○ 주요재정사업 사후평가 → 평가결과 예산편성 반영")
put(T[136], 3, 3, "○ 사후평가 '매우미흡' 사업 ’27년 예산편성 반영(지표 2-7)")

# ---- 2-11 주민의견서 기재 항목: 9개 중 7개 ------------------------------------
for c in (1, 3, 4):
    put(T[142], c, 1, "")
put(T[142], 2, 1, "○")
items = ["-", "○", "○", "○", "○", "-", "○", "○", "○"]  # ①~⑨
put(T[143], 0, 2, "총 7개")
for c, v in enumerate(items, start=1):
    put(T[143], c, 2, v)

out = stamped(BASE)
d.save(out)
print("saved", out)
