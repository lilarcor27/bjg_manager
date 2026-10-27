"""Fill indicators 3-6, 3-7 and 3-8 in the working report (content only).

    python3 tools/edit_3-6_to_3-8.py [input.hwpx]

Answers from the county (2026-10-27): the homepage discloses the same
items as in 2025 (ordinance, operating plan, manual/guide, opinion
letter); promotion used the same media as 2025 (homepage, KakaoTalk
channel, posters, banners, placards, 옥천소식지) with no own content; no
award or incentive was chosen. Counts and dates are marked to confirm.
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

# ---- 3-6 정보공개: 11개 중 4개 -----------------------------------------------
put(T[185], 3, 1, "○")
put(T[186], 2, 1, "4개")
disclosed = {(4, 1): "○", (6, 1): "○", (2, 2): "○", (4, 2): "-", (6, 2): "-",
             (2, 3): "-", (4, 3): "-", (6, 3): "-", (2, 4): "○", (4, 4): "-", (6, 4): "-"}
for (c, r), v in disclosed.items():
    put(T[186], c, r, v)

# ---- 3-7 홍보: 매체 6종, 고유 콘텐츠 없음 ------------------------------------
put(T[192], 1, 1, "○")
put(T[193], 1, 1, "6종")
put(T[193], 2, 1, CHECK)
put(T[193], 1, 2, "군 홈페이지")
put(T[193], 2, 2, CHECK)
put(T[193], 3, 2, "’26.7.~8월")
put(T[193], 4, 2, "‘2027년도 예산편성에 바란다’ 주민의견 수렴 안내")
put(T[193], 3, 3, "’26.7.~8월")
put(T[193], 4, 3, "‘2027년 예산편성 방향 설문조사’ 안내")
put(T[193], 3, 4, CHECK)
put(T[193], 4, 4, "주민참여예산제도 안내 및 주민참여예산 게시판 운영")
put(T[193], 1, 5, "카카오톡채널\n(옥천군 주민참여예산)")
put(T[193], 2, 5, CHECK)
put(T[193], 3, 5, "연중")
put(T[193], 4, 5, "주민참여예산제도 안내 및 주민제안 접수")
put(T[193], 4, 6, "")
put(T[193], 1, 7, "포스터·배너·현수막")
put(T[193], 2, 7, CHECK)
put(T[193], 3, 7, CHECK)
put(T[193], 4, 7, "주민참여예산 제안사업 공모 안내(장애인복지관·도서관 등 다중이용시설 게시)")
r = d.append_row(T[193], 7)
put(T[193], 0, r, "4")
put(T[193], 1, r, "옥천소식지")
put(T[193], 2, r, CHECK)
put(T[193], 3, r, CHECK)
put(T[193], 4, r, "주민참여예산 제안사업 공모 안내")
put(T[194], 1, 1, "해당 없음")

# ---- 3-8 포상·인센티브: 미제공 ----------------------------------------------
put(T[199], 1, 1, "")
put(T[199], 2, 1, "○")
put(T[200], 1, 1, "해당 없음")

out = stamped(BASE)
d.save(out)
print("saved", out)
