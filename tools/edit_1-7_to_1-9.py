"""Fill indicators 1-7, 1-8 and 1-9 in the working report (content only).

    python3 tools/edit_1-7_to_1-9.py [input.hwpx]

Answers from the county (2026-10-27): proposals were taken on the county
homepage (0 received) and the KakaoTalk channel (10 received), no online
vote; one budget officer with under 80% of duties on the system; no
external support organisation. Evidence is added later in one pass.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(__file__))
from hwpx_fill import Hwpx, latest, stamped  # noqa: E402

BASE = "docs/작성본/1. 종합분야 실적서(옥천군)"
d = Hwpx(sys.argv[1] if len(sys.argv) > 1 else latest(BASE))
T = list(d.tables())
put = d.set_cell

# ---- 1-7 온라인플랫폼: 홈페이지·카카오톡채널 접수, 투표 미실시 ----------------
put(T[48], 2, 1, "○")
put(T[49], 3, 1, "○\n(군 홈페이지, 카카오톡채널)")
put(T[49], 3, 2, "군 홈페이지 www.oc.go.kr\n(정보공개 > 예산정보 > 예산편성에바란다)\n카카오톡채널 '옥천군 주민참여예산'")
put(T[49], 3, 3, "10건\n(카카오톡채널 10, 홈페이지 0)")
put(T[49], 3, 4, "미실시")
for c in (1, 2):
    for r in (1, 2, 3, 4):
        put(T[49], c, r, "-")

# ---- 1-8 내부 전담조직: 해당 없음 --------------------------------------------
put(T[55], 1, 2, "X")
put(T[55], 2, 2, "0명")
put(T[55], 3, 2, "0명")
put(T[56], 0, 1, "해당 없음")
put(T[57], 0, 1, "해당 없음")
put(T[57], 1, 1, "주민참여예산 담당자 1명(기획예산담당관 예산팀), 업무비중 80% 미만")

# ---- 1-9 외부 전담조직: 해당 없음 --------------------------------------------
put(T[63], 1, 2, "X")
put(T[63], 2, 2, "-")
put(T[63], 3, 2, "0명")
put(T[64], 0, 1, "해당 없음")
put(T[65], 0, 1, "해당 없음")

out = stamped(BASE)
d.save(out)
print("saved", out)
