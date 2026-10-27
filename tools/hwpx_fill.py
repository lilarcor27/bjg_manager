"""Fill text into an existing HWPX form without touching its layout.

Usage from Python:

    from hwpx_fill import Hwpx
    doc = Hwpx("docs/1. 종합분야 실적서.hwpx")
    doc.replace_text("OO시‧도‧군‧구 추진실적 보고서", "옥천군 추진실적 보고서")
    tbl = doc.table_with("담당자/직급")
    doc.set_cell(tbl, col=3, row=0, text="군")
    doc.save("out.hwpx")

Cells are addressed by the form's own cellAddr (col, row), which already
accounts for merged cells. Multi-line text is split on "\n" into separate
paragraphs that copy the cell's first paragraph style.
"""
import copy
import zipfile

from lxml import etree

HP = "http://www.hancom.co.kr/hwpml/2011/paragraph"
NS = {"hp": HP}
SECTION = "Contents/section0.xml"


def _q(tag):
    return "{%s}%s" % (HP, tag)


def _text(el):
    return "".join(el.itertext())


class Hwpx:
    def __init__(self, path):
        self.path = path
        with zipfile.ZipFile(path) as z:
            raw = z.read(SECTION)
        self.decl = raw[: raw.index(b"?>") + 2] if raw.startswith(b"<?xml") else b""
        self.root = etree.fromstring(raw)

    # ---- lookup -------------------------------------------------------
    def tables(self):
        return self.root.iter(_q("tbl"))

    def table_with(self, anchor, nth=0):
        """Innermost table that has a cell whose text equals anchor."""
        hits = []
        for tbl in self.tables():
            for tc in self._cells(tbl):
                if _text(tc).strip() == anchor:
                    hits.append(tbl)
                    break
        if len(hits) <= nth:
            raise KeyError("table with cell %r (#%d) not found" % (anchor, nth))
        return hits[nth]

    def _cells(self, tbl):
        return tbl.xpath("./hp:tr/hp:tc", namespaces=NS)

    def cell(self, tbl, col, row):
        for tc in self._cells(tbl):
            addr = tc.find(_q("cellAddr"))
            if int(addr.get("colAddr")) == col and int(addr.get("rowAddr")) == row:
                return tc
        raise KeyError("cell (%d,%d) not found" % (col, row))

    def dump(self, tbl):
        """List (col, row, text) for every cell, for working out addresses."""
        out = []
        for tc in self._cells(tbl):
            addr = tc.find(_q("cellAddr"))
            out.append((int(addr.get("colAddr")), int(addr.get("rowAddr")), _text(tc).strip()))
        return out

    # ---- editing ------------------------------------------------------
    def set_cell(self, tbl, col, row, text):
        tc = self.cell(tbl, col, row)
        sub = tc.find(_q("subList"))
        paras = sub.findall(_q("p"))
        template = paras[0]
        for p in paras:
            sub.remove(p)
        for i, line in enumerate(text.split("\n")):
            p = copy.deepcopy(template)
            self._set_para_text(p, line)
            if i:
                p.set("id", "0")
            sub.append(p)
        return tc

    def append_row(self, tbl, src_row):
        """Copy a plain row (no merged cells) to the end of the table."""
        rows = tbl.findall(_q("tr"))
        src = None
        for tr in rows:
            addrs = [int(tc.find(_q("cellAddr")).get("rowAddr")) for tc in tr.findall(_q("tc"))]
            if addrs and addrs[0] == src_row:
                src = tr
        for tc in self._cells(tbl):
            span = tc.find(_q("cellSpan"))
            top = int(tc.find(_q("cellAddr")).get("rowAddr"))
            if top <= src_row < top + int(span.get("rowSpan")) and span.get("rowSpan") != "1":
                raise ValueError("row %d is part of a merged cell" % src_row)
        new_row = int(tbl.get("rowCnt"))
        tr = copy.deepcopy(src)
        height = 0
        for tc in tr.findall(_q("tc")):
            tc.find(_q("cellAddr")).set("rowAddr", str(new_row))
            height = max(height, int(tc.find(_q("cellSz")).get("height")))
        rows[-1].addnext(tr)
        tbl.set("rowCnt", str(new_row + 1))
        sz = tbl.find(_q("sz"))
        sz.set("height", str(int(sz.get("height")) + height))
        return new_row

    def replace_text(self, old, new, count=1):
        done = 0
        for t in self.root.iter(_q("t")):
            if t.text and old in t.text:
                t.text = t.text.replace(old, new)
                self._drop_lineseg(t)
                done += 1
                if done == count:
                    break
        if done == 0:
            raise KeyError("text %r not found" % old)
        return done

    def _set_para_text(self, p, line):
        runs = p.findall(_q("run"))
        first = runs[0]
        for r in runs[1:]:
            p.remove(r)
        for child in list(first):
            first.remove(child)
        if line:
            t = etree.SubElement(first, _q("t"))
            t.text = line
        self._drop_lineseg_para(p)

    def _drop_lineseg(self, el):
        p = el
        while p is not None and p.tag != _q("p"):
            p = p.getparent()
        if p is not None:
            self._drop_lineseg_para(p)

    @staticmethod
    def _drop_lineseg_para(p):
        # Cached line layout; Hangul recomputes it when the element is absent.
        for ls in p.findall(_q("linesegarray")):
            p.remove(ls)

    # ---- output -------------------------------------------------------
    def save(self, out):
        body = etree.tostring(self.root, encoding="UTF-8", xml_declaration=False)
        data = (self.decl or b'<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>') + body
        with zipfile.ZipFile(self.path) as src, zipfile.ZipFile(out, "w") as dst:
            for info in src.infolist():
                payload = data if info.filename == SECTION else src.read(info.filename)
                dst.writestr(info, payload, compress_type=info.compress_type)
