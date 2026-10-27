"""Resident-proposed projects from a budget workbook marked in yellow.

The county marks resident-proposed projects by filling the 세부사업명
cell (column I) yellow. Rows the file already highlighted across the
whole row (column A yellow too) are not counted. Headers (◎) whose
children are marked are grouped; generic headers (일반수용비 etc.) are
split into their children.
"""
import re

import openpyxl
GENERIC = {'일반수용비', '운영수당', '임차료', '일반용역비', '주민참여숙원사업'}
EUPMYEON = ['옥천읍','동이면','안남면','안내면','청성면','청산면','이원면','군서면','군북면']
def yellow(c):
    f = c.fill
    return f and f.fill_type == 'solid' and f.fgColor.type == 'rgb' and f.fgColor.rgb == 'FFFFFF00'
def load(path, year):
    ws = openpyxl.load_workbook(path).worksheets[0]
    rows = []
    for row in ws.iter_rows(min_row=3):
        v = [c.value for c in row]
        if v[2] is None: continue
        rows.append(dict(r=row[0].row, y=yellow(row[8]) and not yellow(row[0]), dept=v[2], biz=v[8], mok=v[12],
                         name=(v[15] or ''), calc=v[16], lvl=v[17] or 0, amt=v[18] or 0))
    # attach children to ◎ headers (any highlight), check consistency
    issues, entries = [], []
    i = 0
    while i < len(rows):
        h = rows[i]
        if not h['name'].strip().startswith('◎'):
            i += 1; continue
        kids = []
        j = i + 1
        while j < len(rows) and not rows[j]['name'].strip().startswith('◎') and rows[j]['biz'] == h['biz'] and rows[j]['dept'] == h['dept']:
            kids.append(rows[j]); j += 1
        if h['y'] or any(k['y'] for k in kids):
            ky = [k for k in kids if k['y']]
            if not h['y'] or len(ky) != len(kids):
                issues.append(('부분 표시', h['r'], h['name'], [k['r'] for k in kids if not k['y']]))
            leaf = [k for k in ky if k['amt']] if ky else []
            if kids and h['lvl'] and sum(k['amt'] for k in kids) != h['lvl']:
                issues.append(('레벨합 불일치', h['r'], h['name'], h['lvl'], sum(k['amt'] for k in kids)))
            hname = h['name'].strip().lstrip('◎').strip()
            eup = next((e for e in EUPMYEON if e in h['dept'] or e in hname or e[:2] + '면' in hname), '')
            if not eup:
                m = re.search(r'\((\S+?[읍면])\)', hname); eup = m.group(1) if m else ''
            typ = '지역회의형' if '숙원' in h['biz'] else ('주민자치회형' if '주민자치회' in h['biz'] else '기타')
            base = dict(year=year, typ=typ, eup=eup, dept=h['dept'], biz=h['biz'], mok=h['mok'])
            if leaf and hname in GENERIC:
                for k in leaf:
                    entries.append(dict(base, name=k['name'].strip().lstrip('○').strip(), amt=k['amt'], rows=str(k['r'])))
            elif leaf:
                entries.append(dict(base, name=hname, amt=sum(k['amt'] for k in leaf), rows=f"{h['r']}~{leaf[-1]['r']}"))
            else:
                entries.append(dict(base, name=hname, amt=h['amt'], rows=str(h['r'])))
        i = j
    total_y = sum(r['amt'] for r in rows if r['y'])
    return entries, issues, total_y
