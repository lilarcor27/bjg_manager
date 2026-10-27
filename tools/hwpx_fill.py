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
import random
import re
import struct
import zipfile

from lxml import etree

HP = "http://www.hancom.co.kr/hwpml/2011/paragraph"
HC = "http://www.hancom.co.kr/hwpml/2011/core"
NS = {"hp": HP}
SECTION = "Contents/section0.xml"
MANIFEST = "Contents/content.hpf"
HWPUNIT_PER_PX = 75  # 7200 per inch at 96 dpi

# Picture element as Hangul writes it for an inline (treat-as-char) image.
PIC = """<hp:pic xmlns:hp="{hp}" xmlns:hc="{hc}" id="{id}" zOrder="{z}" numberingType="PICTURE"
 textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" href="" groupLevel="0"
 instid="{inst}" reverse="0"><hp:offset x="0" y="0"/><hp:orgSz width="{ow}" height="{oh}"/>
<hp:curSz width="{w}" height="{h}"/><hp:flip horizontal="0" vertical="0"/>
<hp:rotationInfo angle="0" centerX="{cx}" centerY="{cy}" rotateimage="1"/><hp:renderingInfo>
<hc:transMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>
<hc:scaMatrix e1="{sx}" e2="0" e3="0" e4="0" e5="{sy}" e6="0"/>
<hc:rotMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/></hp:renderingInfo>
<hc:img binaryItemIDRef="{ref}" bright="0" contrast="0" effect="REAL_PIC" alpha="0"/>
<hp:imgRect><hc:pt0 x="0" y="0"/><hc:pt1 x="{ow}" y="0"/><hc:pt2 x="{ow}" y="{oh}"/><hc:pt3 x="0" y="{oh}"/></hp:imgRect>
<hp:imgClip left="0" right="{ow}" top="0" bottom="{oh}"/><hp:inMargin left="0" right="0" top="0" bottom="0"/>
<hp:imgDim dimwidth="{ow}" dimheight="{oh}"/><hp:effects/>
<hp:sz width="{w}" widthRelTo="ABSOLUTE" height="{h}" heightRelTo="ABSOLUTE" protect="0"/>
<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0"
 vertRelTo="PARA" horzRelTo="COLUMN" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/>
<hp:outMargin left="0" right="0" top="0" bottom="0"/><hp:shapeComment>그림입니다.</hp:shapeComment></hp:pic>"""


def _png_size(data):
    if data[:8] != b"\x89PNG\r\n\x1a\n":
        raise ValueError("only PNG images are supported")
    return struct.unpack(">II", data[16:24])


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
        self.new_bins = []  # (item id, zip path, bytes)

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

    def delete_row(self, tbl, row):
        """Remove a plain row (no merged cells) and renumber the rows below."""
        for tc in self._cells(tbl):
            span = tc.find(_q("cellSpan"))
            top = int(tc.find(_q("cellAddr")).get("rowAddr"))
            if span.get("rowSpan") != "1" and top <= row < top + int(span.get("rowSpan")):
                raise ValueError("row %d is part of a merged cell" % row)
        height = 0
        for tr in tbl.findall(_q("tr")):
            tcs = tr.findall(_q("tc"))
            r = int(tcs[0].find(_q("cellAddr")).get("rowAddr"))
            if r == row:
                height = max(int(tc.find(_q("cellSz")).get("height")) for tc in tcs)
                tbl.remove(tr)
        for tc in self._cells(tbl):
            addr = tc.find(_q("cellAddr"))
            if int(addr.get("rowAddr")) > row:
                addr.set("rowAddr", str(int(addr.get("rowAddr")) - 1))
        tbl.set("rowCnt", str(int(tbl.get("rowCnt")) - 1))
        sz = tbl.find(_q("sz"))
        sz.set("height", str(max(0, int(sz.get("height")) - height)))

    def set_cell_image(self, tbl, col, row, png_path, comment=None):
        """Replace a cell's content with a PNG scaled to the cell's inner width."""
        data = open(png_path, "rb").read()
        px_w, px_h = _png_size(data)
        tc = self.cell(tbl, col, row)
        margin = tc.find(_q("cellMargin"))
        inner = int(tc.find(_q("cellSz")).get("width")) - int(margin.get("left")) - int(margin.get("right"))
        ow, oh = px_w * HWPUNIT_PER_PX, px_h * HWPUNIT_PER_PX
        w = min(inner, ow)
        h = round(oh * w / ow)
        used_ids = {int(x) for x in re.findall(r'\bid="(\d+)"', etree.tostring(self.root).decode()) if len(x) < 12}
        pid = random.randint(10**9, 2 * 10**9)
        while pid in used_ids:
            pid += 1
        zs = [int(z) for z in self.root.xpath("//@zOrder")]
        n = len(self._existing_bins()) + len(self.new_bins) + 1
        ref = "image%d" % n
        pic = etree.fromstring(PIC.replace("\n", "").format(
            hp=HP, hc=HC, id=pid, z=max(zs, default=0) + 1, inst=pid - 1000, ow=ow, oh=oh, w=w, h=h,
            cx=w // 2, cy=h // 2, sx="%.6f" % (w / ow), sy="%.6f" % (h / oh), ref=ref))
        if comment:
            pic.find(_q("shapeComment")).text = comment
        self.set_cell(tbl, col, row, "")
        run = tc.find(_q("subList")).find(_q("p")).find(_q("run"))
        for child in list(run):
            run.remove(child)
        run.append(pic)
        etree.SubElement(run, _q("t"))
        self.new_bins.append((ref, "BinData/%s.png" % ref, data))
        return ref

    def _existing_bins(self):
        with zipfile.ZipFile(self.path) as z:
            return [n for n in z.namelist() if n.startswith("BinData/")]

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
                if info.filename == MANIFEST and self.new_bins:
                    payload = self._manifest_with_bins(payload)
                dst.writestr(info, payload, compress_type=info.compress_type)
            for _, name, blob in self.new_bins:
                dst.writestr(zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0)), blob, compress_type=zipfile.ZIP_STORED)

    def _manifest_with_bins(self, payload):
        text = payload.decode("utf-8")
        items = "".join('<opf:item id="%s" href="%s" media-type="image/png" isEmbeded="1"/>' % (ref, name)
                        for ref, name, _ in self.new_bins)
        anchor = '<opf:item id="section0"'
        if anchor not in text:
            raise ValueError("manifest has no section0 item")
        return text.replace(anchor, items + anchor, 1).encode("utf-8")
