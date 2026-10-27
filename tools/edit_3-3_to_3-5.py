"""Fill 3-3 and 3-4, and strengthen 3-5 (content only).

    python3 tools/edit_3-3_to_3-5.py [input.hwpx]

Answers from the county (2026-10-27): 3-3 at the 2025 level (the
departments' review of proposals with proposer contact, and site
checks); 3-4 no facilitator or outside expert. 3-5 adds the residents'
general assemblies that choose resident-council projects (see 2-4).
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

# ---- 3-3 제안서 보완 지원: 부서검토·제안자 소통, 현장 확인 ---------------------
put(T[162], 1, 1, "○")
put(T[162], 2, 1, "")
put(T[163], 0, 1, "주민제안자\n(홈페이지·카카오톡채널·서면 접수)")
put(T[163], 1, 1, "담당자 상담")
put(T[163], 2, 1,
    "○ 접수된 주민제안사업을 소관부서에 보내 검토의견 제출 요청(’26." + CHECK + ")\n"
    "  - 검토 시 제안자와 소통하여 제안사업의 취지·내용을 명확히 파악\n"
    "  - 사업 구체화 시 제안자와 숙의과정을 거쳐 사업 만족도 제고")
put(T[163], 0, 2, "주민제안사업")
put(T[163], 1, 2, "현장 확인")
put(T[163], 2, 2, "○ 소관부서 현장 확인으로 제안사업의 추진 가능성 검토 및 사업내용 보완")

# ---- 3-4 기구 숙의 지원: 미지원 ---------------------------------------------
put(T[169], 1, 1, "")
put(T[169], 2, 1, "○")
put(T[170], 0, 1, "해당 없음")

# ---- 3-5 주민자치회 연계: 주민총회 추가 -------------------------------------
put(T[179], 2, 1,
    "○ 주민참여예산 유형에 '주민자치회형' 운영: 읍면 주민자치회가 사업을 발굴하고, 주민총회에서 "
    "주민 투표로 선정한 사업을 읍면별 한도 내에서 예산 편성\n"
    "  - ’26년 본예산 9개 읍면 각 50,000천원, 총 450,000천원"
    "(주민주도형 5,000천원 + 민관협력형 45,000천원)\n"
    "○ 지역회의 구성 시 주민자치회·이장협의회 등 읍면 단위 단체가 대행 가능(제8기 구성 계획)")
put(T[179], 2, 2,
    "○ 읍면 주민자치회 대상 주민참여예산 홍보 및 교육 10회(’26.7.~8월)\n"
    "○ 주민자치회 주민총회를 통한 주민자치회형 사업 선정(지표 2-4)")

out = stamped(BASE)
d.save(out)
print("saved", out)
