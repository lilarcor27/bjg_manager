"""Complete indicators 1-1 and 1-3 (and tidy 1-2, 3-2) in the working report.

    python3 tools/edit_1-1_to_1-3.py

Edits `docs/작성본/1. 종합분야 실적서(옥천군).hwpx` in place, on top of the
version the county saved from Hangul. Table numbers (T[n]) follow the form.
"""
import copy
import os
import sys

sys.path.insert(0, os.path.dirname(__file__))
from hwpx_fill import Hwpx, _q  # noqa: E402

DOC = "docs/작성본/1. 종합분야 실적서(옥천군).hwpx"
IMG = "docs/작성본/이미지/"
NEED = "【자료 필요】"

d = Hwpx(DOC)
T = list(d.tables())
put = d.set_cell


def serial_lines(old, lines):
    """Replace a '▣ 세부 증빙자료' line with one paragraph per evidence item."""
    for t in d.root.iter(_q("t")):
        if t.text and t.text.strip() == old:
            p = t.getparent().getparent()
            prev = p
            for i, line in enumerate(lines):
                q = p if i == 0 else copy.deepcopy(p)
                for tt in q.iter(_q("t")):
                    tt.text = ""
                next(q.iter(_q("t"))).text = "  " + line
                for ls in q.findall(_q("linesegarray")):
                    q.remove(ls)
                if i:
                    q.set("id", "0")
                    prev.addnext(q)
                prev = q
            return
    raise KeyError(old)


# ---- 1-1 증빙: '25·'26 주민제안사업 예산액 ---------------------------------
d.set_cell_image(T[10], 0, 3, IMG + "1-1 25년 주민제안사업 예산액.png", "’25년 주민제안사업 예산액")
d.set_cell_image(T[10], 1, 3, IMG + "1-1 26년 주민제안사업 예산액.png", "’26년 주민제안사업 예산액")
serial_lines("1-1-1 (평가지표-일련번호)", [
    "1-1-1 ’24년 주민세 징수상황총괄(표준지방세정보시스템 7202)",
    "1-1-2 ’26년 주민제안사업 예산 내역(2026 본예산, 111건)",
    "1-1-3 ’25년 주민제안사업 예산 내역(2025 당초예산, 103건)",
])

# ---- 1-2: 지표 원문 복구, 증빙 칸 정리 ----------------------------------------
for t in list(d.root.iter(_q("t"))):
    if t.text == "2억이상 투자 + 행사성":  # the county's memo, typed into the indicator text
        run = t.getparent()
        para = run.getparent()
        para.remove(run)
        last = list(para.iter(_q("t")))[-1]
        last.text = last.text.rstrip()
        for ls in para.findall(_q("linesegarray")):
            para.remove(ls)
put(T[17], 0, 0, "2026년 주요재정사업 평가 계획")
put(T[17], 1, 0, "지방보조금관리위원회 심의 결과(’26.8.26.)")
put(T[17], 0, 1, NEED + " 평가 대상(2억원 이상 투자·행사성 사업)과 평가 절차가 나온 쪽")
put(T[17], 1, 1, NEED + " 결과보고 또는 회의록 중 신규 행사성사업 의견·평가결과 환류 부분")
put(T[17], 0, 2, "지방보조금관리위원회 위원 명단")
put(T[17], 1, 2, "참여예산주민위원회 위원 명단")
put(T[17], 0, 3, NEED + " 참여예산주민위원 겸직 위원 표시")
put(T[17], 1, 3, NEED + " 겸직 위원 표시")
serial_lines("1-2-1 (평가지표-일련번호)", [
    "1-2-1 2026년 주요재정사업 평가 계획",
    "1-2-2 지방보조금관리위원회 심의 결과보고 및 회의록(’26.8.26.)",
    "1-2-3 지방보조금관리위원회·참여예산주민위원회 위원 명단",
])

# ---- 1-3 증빙: 운영조례 조문 ----------------------------------------------
put(T[24], 0, 0, "운영조례 제8조(설치)·제9조(기능)")
put(T[24], 1, 0, "운영조례 제15조(회의)·제16조(분과위원회)")
d.set_cell_image(T[24], 0, 1, IMG + "1-3 운영조례 제8조·제9조.png", "옥천군 주민참여예산제 운영조례 제8조·제9조")
d.set_cell_image(T[24], 1, 1, IMG + "1-3 운영조례 제15조·제16조.png", "옥천군 주민참여예산제 운영조례 제15조·제16조")
d.delete_row(T[24], 3)
d.delete_row(T[24], 2)
serial_lines("1-3-1 (평가지표-일련번호)", [
    "1-3-1 옥천군 주민참여예산제 운영조례(제8조, 제9조, 제15조, 제16조)",
])

# ---- 3-2: 읍면 교육 두 행을 하나로 (담당자 요청) ------------------------------
for c, v in enumerate([
        "주민참여예산 홍보 및 교육", "’26.7.8.~8.28.(35회)",
        "9개 읍면 주민자치회·이장협의회·새마을단체·기관단체협의회 등",
        "35시간\n(회당 1시간)", "주민참여예산 운영개요, 지역회의 역할 및 참여방법, 주요 재정 현황"], start=1):
    put(T[156], c, 3, v)
d.delete_row(T[156], 5)
d.delete_row(T[156], 4)

d.save(DOC + ".tmp")
os.replace(DOC + ".tmp", DOC)
print("saved", DOC)
