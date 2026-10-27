"""Complete indicators 1-4, 1-5 and 1-6 in the working report.

    python3 tools/edit_1-4_to_1-6.py [input.hwpx]

Reads the given file, or the newest `docs/작성본/1. 종합분야 실적서(옥천군)
yymmdd hhmm.hwpx`, adds the evidence images and detailed-evidence lines,
and saves a new file stamped with the current Korea time.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(__file__))
from hwpx_fill import Hwpx, latest, stamped  # noqa: E402

BASE = "docs/작성본/1. 종합분야 실적서(옥천군)"
IMG = "docs/작성본/이미지/"
NEED = "【자료 필요】"
CELL_H = 25500  # image rows are about 26,000 HWPUNIT tall

d = Hwpx(sys.argv[1] if len(sys.argv) > 1 else latest(BASE))
T = list(d.tables())
put = d.set_cell
plan = IMG + "1-4·1-5 제8기 위원회 구성 계획.png"

# ---- 1-4 대표성: 각종 위원회 조례 제7조⑥, 제8기 구성 계획 --------------------
put(T[31], 0, 0, "옥천군 각종 위원회 설치 및 운영 조례 제7조")
put(T[31], 1, 0, "제8기 참여예산주민위원회 구성 계획(’25.8.28.)")
d.set_cell_image(T[31], 0, 1, IMG + "1-4 각종 위원회 조례 제7조.png",
                 "옥천군 각종 위원회 설치 및 운영 조례 제7조", CELL_H)
d.set_cell_image(T[31], 1, 1, plan, "제8기 참여예산주민위원회 및 주민참여예산 지역회의 구성 계획", CELL_H)
d.delete_row(T[31], 3)
d.delete_row(T[31], 2)
d.set_lines("1-4-1 (평가지표-일련번호)", [
    "1-4-1 옥천군 각종 위원회 설치 및 운영 조례(제7조제6항)",
    "1-4-2 제8기 참여예산주민위원회 및 주민참여예산 지역회의 구성 계획(’25.8.28.)",
])

# ---- 1-5 전문성: 운영조례 제16조, 구성 계획 -----------------------------------
put(T[38], 0, 0, "운영조례 제16조(분과위원회)")
put(T[38], 1, 0, "제8기 구성 계획(분과위원회·지역회의 구성)")
d.set_cell_image(T[38], 0, 1, IMG + "1-5 운영조례 제16조.png", "옥천군 주민참여예산제 운영조례 제16조", CELL_H)
d.set_cell_image(T[38], 1, 1, plan, "제8기 참여예산주민위원회 및 주민참여예산 지역회의 구성 계획", CELL_H)
d.delete_row(T[38], 3)
d.delete_row(T[38], 2)
d.set_lines("1-5-1 (평가지표-일련번호)", [
    "1-5-1 옥천군 주민참여예산제 운영조례(제16조 분과위원회, 제24조 지역회의)",
    "1-5-2 제8기 참여예산주민위원회 및 주민참여예산 지역회의 구성 계획(’25.8.28.)",
])

# ---- 1-6 운영 활성화: 분과회의 결과 보고, 지역회의 회의록 ------------------------
put(T[45], 0, 0, "2026 참여예산주민위원회 분과별회의 결과 보고")
put(T[45], 1, 0, "주민참여예산 지역회의 회의록(9개 읍면)")
d.set_cell_image(T[45], 0, 1, IMG + "1-6 분과별회의 결과 보고.png",
                 "2026년 참여예산주민위원회 분과별회의 결과 보고(기획예산담당관-22066)", CELL_H)
put(T[45], 1, 1, NEED + " 지역회의 회의록 1쪽(대표 1개 읍면)")
put(T[45], 0, 2, "분과위원회 회의 사진")
put(T[45], 1, 2, "지역회의 회의 사진")
put(T[45], 0, 3, NEED + " 분과위원회 회의 사진")
put(T[45], 1, 3, NEED + " 지역회의 회의 사진")
d.set_lines("1-6-1 (평가지표-일련번호)", [
    "1-6-1 2026년 참여예산주민위원회 분과별회의 결과 보고(기획예산담당관-22066, ’26.10.21.) 및 분과별 회의결과·회의록",
    "1-6-2 2026년 주민참여예산 지역회의 회의록(9개 읍면)",
])

out = stamped(BASE)
d.save(out)
print("saved", out)
