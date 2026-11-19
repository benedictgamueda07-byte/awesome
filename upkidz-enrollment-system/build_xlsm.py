"""
Builds the ready-to-use macro-enabled workbook:

    UpKidz_Enrollment_System.xlsx  +  UpKidzEnrollment.bas  ->  UpKidz_Enrollment_System.xlsm

It does three things, so nobody has to open the VBA editor:
  1. Generates xl/vbaProject.bin from the .bas source, following Microsoft's
     [MS-OVBA] specification (dir stream, compressed module source, PROJECT
     stream). The project has no compiled p-code ("_VBA_PROJECT" version
     0xFFFF), so Excel compiles the source itself the first time it opens it.
  2. Adds ThisWorkbook / Sheet document modules. ThisWorkbook.Workbook_Open
     calls EnsureButtons, which re-links every button to its macro.
  3. Draws the buttons (rounded rectangles already assigned to their macros)
     into the Dashboard and Student Search sheets.

Usage:  python3 build_xlsm.py [in.xlsx] [module.bas] [out.xlsm]
Needs:  Python 3.9+ standard library only
"""
import os
import re
import struct
import sys
import uuid
import random
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "UpKidz_Enrollment_System.xlsx")
BAS = sys.argv[2] if len(sys.argv) > 2 else os.path.join(HERE, "UpKidzEnrollment.bas")
OUT = sys.argv[3] if len(sys.argv) > 3 else os.path.join(HERE, "UpKidz_Enrollment_System.xlsm")

CODEPAGE = 1252
ENC = "cp1252"


# ---------------------------------------------------------------------------
# [MS-OVBA] 2.4.1 Compression (verified against oletools' decompressor)
# ---------------------------------------------------------------------------
def _compress_chunk(chunk: bytes) -> bytes:
    out = bytearray()
    pos, n = 0, len(chunk)
    heads = {}                                   # 3-byte prefix -> positions
    while pos < n:
        flag_index = len(out)
        out.append(0)
        flag = 0
        for bit in range(8):
            if pos >= n:
                break
            best_len, best_off = 0, 0
            if pos > 0:
                bit_count = max((pos - 1).bit_length(), 4)
                max_len = (0xFFFF >> bit_count) + 3
                limit = min(max_len, n - pos)
                if limit >= 3:
                    for cand in reversed(heads.get(chunk[pos:pos + 3], [])[-64:]):
                        ln = 0
                        while ln < limit and chunk[cand + ln] == chunk[pos + ln]:
                            ln += 1
                        if ln > best_len:
                            best_len, best_off = ln, pos - cand
                            if ln == limit:
                                break
            if best_len >= 3:
                bit_count = max((pos - 1).bit_length(), 4)
                token = ((best_off - 1) << (16 - bit_count)) | (best_len - 3)
                out += struct.pack("<H", token)
                flag |= 1 << bit
                step = best_len
            else:
                out.append(chunk[pos])
                step = 1
            for i in range(pos, pos + step):
                if i + 3 <= n:
                    heads.setdefault(chunk[i:i + 3], []).append(i)
            pos += step
        out[flag_index] = flag
    return bytes(out)


def compress(data: bytes) -> bytes:
    out = bytearray(b"\x01")
    for start in range(0, len(data), 4096):
        chunk = data[start:start + 4096]
        comp = _compress_chunk(chunk)
        if len(comp) <= 4096:
            out += struct.pack("<H", 0xB000 | (len(comp) + 2 - 3)) + comp
        else:                                     # store raw (only valid for a full chunk)
            assert len(chunk) == 4096, "incompressible short chunk"
            out += struct.pack("<H", 0x3000 | (4096 + 2 - 3)) + chunk
    return bytes(out)


# ---------------------------------------------------------------------------
# [MS-OVBA] 2.4.3 Data Encryption (used for CMG / DPB / GC in PROJECT stream)
# ---------------------------------------------------------------------------
def ovba_encrypt(project_id: str, data: bytes) -> str:
    seed = random.randint(0, 255)
    version_enc = 2 ^ seed
    proj_key = sum(ord(c) for c in project_id) & 0xFF
    proj_key_enc = proj_key ^ seed
    unenc1, enc1, enc2 = proj_key, proj_key_enc, version_enc
    out = bytearray([seed, version_enc, proj_key_enc])
    for _ in range((seed & 6) // 2):
        temp = random.randint(0, 255)
        b = temp ^ ((enc2 + unenc1) & 0xFF)
        out.append(b)
        enc2, enc1, unenc1 = enc1, b, temp
    for byte in struct.pack("<I", len(data)) + data:
        b = byte ^ ((enc2 + unenc1) & 0xFF)
        out.append(b)
        enc2, enc1, unenc1 = enc1, b, byte
    return out.hex().upper()


# ---------------------------------------------------------------------------
# module sources
# ---------------------------------------------------------------------------
DOC_ATTRS = ('Attribute VB_Base = "0{{{base}}}"\r\n'
             'Attribute VB_GlobalNameSpace = False\r\n'
             'Attribute VB_Creatable = False\r\n'
             'Attribute VB_PredeclaredId = True\r\n'
             'Attribute VB_Exposed = True\r\n'
             'Attribute VB_TemplateDerived = False\r\n'
             'Attribute VB_Customizable = True\r\n')
WORKBOOK_BASE = "00020819-0000-0000-C000-000000000046"
SHEET_BASE = "00020820-0000-0000-C000-000000000046"

THISWORKBOOK_CODE = (
    "Option Explicit\r\n"
    "\r\n"
    "Private Sub Workbook_Open()\r\n"
    "    ' Make sure every button exists and is linked to its macro\r\n"
    "    On Error Resume Next\r\n"
    "    Application.Run \"'\" & Me.Name & \"'!EnsureButtons\"\r\n"
    "End Sub\r\n")


def crlf(s: str) -> str:
    return s.replace("\r\n", "\n").replace("\n", "\r\n")


def std_module_source(path):
    text = open(path, encoding="utf-8").read()
    text.encode(ENC)                       # fails loudly on non-cp1252 characters
    m = re.search(r'Attribute VB_Name = "([^"]+)"', text)
    name = m.group(1) if m else "Module1"
    if not m:
        text = f'Attribute VB_Name = "{name}"\n' + text
    return name, crlf(text)


def doc_module_source(name, base, code=""):
    return (f'Attribute VB_Name = "{name}"\r\n' + DOC_ATTRS.format(base=base) + code)


# ---------------------------------------------------------------------------
# dir stream  ([MS-OVBA] 2.3.4.2)
# ---------------------------------------------------------------------------
def rec(rid, payload: bytes) -> bytes:
    return struct.pack("<HI", rid, len(payload)) + payload


def u16(s):
    return s.encode("utf-16-le")


def build_dir(project_name, modules):
    d = b""
    d += rec(0x0001, struct.pack("<I", 1))            # SYSKIND = Win32
    d += rec(0x0002, struct.pack("<I", 0x0409))       # LCID
    d += rec(0x0014, struct.pack("<I", 0x0409))       # LCIDINVOKE
    d += rec(0x0003, struct.pack("<H", CODEPAGE))     # CODEPAGE
    d += rec(0x0004, project_name.encode(ENC))        # NAME
    d += rec(0x0005, b"") + rec(0x0040, b"")          # DOCSTRING (+unicode)
    d += rec(0x0006, b"") + rec(0x003D, b"")          # HELPFILEPATH 1 / 2
    d += rec(0x0007, struct.pack("<I", 0))            # HELPCONTEXT
    d += rec(0x0008, struct.pack("<I", 0))            # LIBFLAGS
    d += struct.pack("<HIIH", 0x0009, 4, 1, 0)        # VERSION (reserved=4, major, minor)
    d += rec(0x000C, b"") + rec(0x003C, b"")          # CONSTANTS (+unicode)
    # REFERENCES: OLE Automation (stdole). VBA and Excel libraries are implicit.
    ref_name = "stdole"
    libid = (r"*\G{00020430-0000-0000-C000-000000000046}#2.0#0#"
             r"C:\Windows\System32\stdole2.tlb#OLE Automation").encode(ENC)
    d += rec(0x0016, ref_name.encode(ENC)) + rec(0x003E, u16(ref_name))
    d += rec(0x000D, struct.pack("<I", len(libid)) + libid + struct.pack("<IH", 0, 0))
    # MODULES
    d += rec(0x000F, struct.pack("<H", len(modules)))
    d += rec(0x0013, struct.pack("<H", 0xFFFF))       # PROJECTCOOKIE
    for m in modules:
        n = m["name"]
        d += rec(0x0019, n.encode(ENC))               # MODULENAME
        d += rec(0x0047, u16(n))                      # MODULENAMEUNICODE
        d += rec(0x001A, n.encode(ENC)) + rec(0x0032, u16(n))   # STREAMNAME
        d += rec(0x001C, b"") + rec(0x0048, b"")      # DOCSTRING
        d += rec(0x0031, struct.pack("<I", 0))        # OFFSET (no p-code cache)
        d += rec(0x001E, struct.pack("<I", 0))        # HELPCONTEXT
        d += rec(0x002C, struct.pack("<H", 0xFFFF))   # COOKIE
        d += struct.pack("<HI", 0x0021 if m["type"] == "std" else 0x0022, 0)  # TYPE
        d += struct.pack("<HI", 0x002B, 0)            # module terminator
    d += struct.pack("<HI", 0x0010, 0)                # dir terminator
    return d


def build_project_stream(project_id, modules):
    lines = [f'ID="{project_id}"']
    for m in modules:
        if m["type"] == "doc":
            lines.append(f"Document={m['name']}/&H00000000")
        else:
            lines.append(f"Module={m['name']}")
    lines += ['Name="VBAProject"', 'HelpContextID="0"', 'VersionCompatible32="393222000"',
              f'CMG="{ovba_encrypt(project_id, struct.pack("<I", 0))}"',   # not protected
              f'DPB="{ovba_encrypt(project_id, b"\x00")}"',               # no password
              f'GC="{ovba_encrypt(project_id, b"\xff")}"',                # visible
              "", "[Host Extender Info]",
              "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000",
              "", "[Workspace]"]
    for m in modules:
        lines.append(f"{m['name']}=0, 0, 0, 0, C")
    return ("\r\n".join(lines) + "\r\n").encode(ENC)


def build_projectwm(modules):
    b = b""
    for m in modules:
        b += m["name"].encode(ENC) + b"\x00" + u16(m["name"]) + b"\x00\x00"
    return b + b"\x00\x00"


def build_vba_project_bin(sheet_codenames, bas_path):
    std_name, std_src = std_module_source(bas_path)
    modules = [dict(name="ThisWorkbook", type="doc",
                    src=doc_module_source("ThisWorkbook", WORKBOOK_BASE, THISWORKBOOK_CODE))]
    for cn in sheet_codenames:
        modules.append(dict(name=cn, type="doc", src=doc_module_source(cn, SHEET_BASE)))
    modules.append(dict(name=std_name, type="std", src=std_src))

    project_id = "{" + str(uuid.uuid4()).upper() + "}"
    streams = {
        "PROJECT": build_project_stream(project_id, modules),
        "PROJECTwm": build_projectwm(modules),
        "VBA/_VBA_PROJECT": bytes([0xCC, 0x61, 0xFF, 0xFF, 0x00, 0x03, 0x00]),
        "VBA/dir": compress(build_dir("VBAProject", modules)),
    }
    for m in modules:
        streams["VBA/" + m["name"]] = compress(m["src"].encode(ENC))

    return write_cfb(streams)


# ---------------------------------------------------------------------------
# Minimal OLE compound file writer ([MS-CFB] version 3, 512-byte sectors)
# ---------------------------------------------------------------------------
FREESECT, ENDOFCHAIN, FATSECT, NOSTREAM = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD, 0xFFFFFFFF
SECTOR, MINI, CUTOFF = 512, 64, 4096


def _cfb_key(name):
    return (len(name), name.upper())


def write_cfb(streams):
    """streams: {"A/B/stream": bytes}. Returns compound file bytes."""
    # ---- build the directory tree
    root = {"name": "Root Entry", "type": 5, "kids": {}}
    for path, data in streams.items():
        node = root
        parts = path.split("/")
        for part in parts[:-1]:
            node = node["kids"].setdefault(part, {"name": part, "type": 1, "kids": {}})
        node["kids"][parts[-1]] = {"name": parts[-1], "type": 2, "data": data, "kids": {}}
    entries = []

    def collect(node):
        node["id"] = len(entries)
        entries.append(node)
        for k in node["kids"].values():
            collect(k)
    collect(root)
    for e in entries:
        e.update(left=NOSTREAM, right=NOSTREAM, child=NOSTREAM, color=1)

    # Siblings form a balanced binary tree (middle element at the root). All
    # null links then sit at depth h or h+1 (h = floor(log2 n)); colouring the
    # nodes on level h red (unless the tree is perfect) gives a valid red-black tree.
    def build_tree(sib, depth, h, perfect):
        if not sib:
            return NOSTREAM
        mid = len(sib) // 2
        n = sib[mid]
        n["left"] = build_tree(sib[:mid], depth + 1, h, perfect)
        n["right"] = build_tree(sib[mid + 1:], depth + 1, h, perfect)
        n["color"] = 0 if (depth == h and h > 0 and not perfect) else 1   # 0 = red, 1 = black
        return n["id"]

    for e in entries:
        kids = sorted(e["kids"].values(), key=lambda k: _cfb_key(k["name"]))
        if kids:
            n = len(kids)
            e["child"] = build_tree(kids, 0, n.bit_length() - 1, (n + 1) & n == 0)

    # ---- place stream data
    fat, sectors = [], []          # sectors: list of 512-byte blocks
    minifat, ministream = [], bytearray()

    def alloc(data):
        n = (len(data) + SECTOR - 1) // SECTOR
        start = len(sectors)
        for i in range(n):
            sectors.append(data[i * SECTOR:(i + 1) * SECTOR].ljust(SECTOR, b"\x00"))
            fat.append(start + i + 1 if i < n - 1 else ENDOFCHAIN)
        return start if n else ENDOFCHAIN

    for e in entries:
        if e["type"] != 2:
            continue
        d = e["data"]
        e["size"] = len(d)
        if len(d) == 0:
            e["start"] = ENDOFCHAIN
        elif len(d) < CUTOFF:
            n = (len(d) + MINI - 1) // MINI
            start = len(minifat)
            for i in range(n):
                minifat.append(start + i + 1 if i < n - 1 else ENDOFCHAIN)
            ministream += d.ljust(n * MINI, b"\x00")
            e["start"] = start
        else:
            e["start"] = alloc(d)
    root["size"] = len(ministream)
    root["start"] = alloc(bytes(ministream)) if ministream else ENDOFCHAIN
    mf = b"".join(struct.pack("<I", x) for x in minifat)
    mf += struct.pack("<I", FREESECT) * ((-len(minifat)) % 128)
    first_minifat = alloc(mf) if minifat else ENDOFCHAIN
    n_minifat = (len(mf) // SECTOR) if minifat else 0

    # ---- directory
    def dirent(e):
        if e is None:
            return (b"\x00" * 64 + struct.pack("<HBB", 0, 0, 0) + struct.pack("<III", NOSTREAM, NOSTREAM, NOSTREAM)
                    + b"\x00" * 16 + b"\x00" * 4 + b"\x00" * 16 + struct.pack("<IQ", 0, 0))
        nm = e["name"].encode("utf-16-le") + b"\x00\x00"
        assert len(nm) <= 64
        return (nm.ljust(64, b"\x00") + struct.pack("<HBB", len(nm), e["type"], e["color"])
                + struct.pack("<III", e["left"], e["right"], e["child"]) + b"\x00" * 16 + b"\x00" * 4
                + b"\x00" * 16 + struct.pack("<IQ", e.get("start", 0) if e["type"] != 1 else 0,
                                             e.get("size", 0) if e["type"] != 1 else 0))
    dir_bytes = b"".join(dirent(e) for e in entries)
    dir_bytes += dirent(None) * ((-len(entries)) % 4)
    first_dir = alloc(dir_bytes)

    # ---- FAT (sized so it can describe itself)
    n_fat = 1
    while (len(sectors) + n_fat) > n_fat * 128:
        n_fat += 1
    assert n_fat <= 109, "file too large for header DIFAT"
    fat_start = len(sectors)
    fat += [FATSECT] * n_fat
    fat += [FREESECT] * (n_fat * 128 - len(fat))
    fb = b"".join(struct.pack("<I", x) for x in fat)
    for i in range(n_fat):
        sectors.append(fb[i * SECTOR:(i + 1) * SECTOR])

    difat = [fat_start + i for i in range(n_fat)] + [FREESECT] * (109 - n_fat)
    header = (bytes.fromhex("D0CF11E0A1B11AE1") + b"\x00" * 16
              + struct.pack("<HHHHH", 0x003E, 0x0003, 0xFFFE, 9, 6) + b"\x00" * 6
              + struct.pack("<IIIIIIIII", 0, n_fat, first_dir, 0, CUTOFF, first_minifat, n_minifat,
                            ENDOFCHAIN, 0)
              + b"".join(struct.pack("<I", x) for x in difat))
    assert len(header) == 512
    return header + b"".join(sectors)


# ---------------------------------------------------------------------------
# button shapes (DrawingML) already linked to their macros
# ---------------------------------------------------------------------------
EMU_PX, EMU_PT = 9525, 12700


def col_px(width):
    return int(((256 * width + int(128 / 7)) / 256) * 7)


def anchor_xml(idx, name, text, macro, color, size_pt, c1, c2, rows, col_widths):
    """Button spanning columns c1..c2 and rows r1..r2 (0-based, inclusive), inset by 3px."""
    r1, r2 = rows
    inset = 3 * EMU_PX
    right_off = col_px(col_widths[c2]) * EMU_PX - inset
    return f"""<xdr:twoCellAnchor editAs="absolute">
<xdr:from><xdr:col>{c1}</xdr:col><xdr:colOff>{inset}</xdr:colOff><xdr:row>{r1}</xdr:row><xdr:rowOff>{2 * EMU_PX}</xdr:rowOff></xdr:from>
<xdr:to><xdr:col>{c2}</xdr:col><xdr:colOff>{right_off}</xdr:colOff><xdr:row>{r2 + 1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
<xdr:sp macro="[0]!{macro}" textlink="">
<xdr:nvSpPr><xdr:cNvPr id="{idx + 1}" name="{name}"/><xdr:cNvSpPr/></xdr:nvSpPr>
<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>
<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:ln><a:noFill/></a:ln></xdr:spPr>
<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" lIns="36000" rIns="36000" tIns="0" bIns="0" rtlCol="0" anchor="ctr"/><a:lstStyle/>
<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="{int(size_pt * 100)}" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="Arial"/><a:cs typeface="Arial"/></a:rPr><a:t>{text}</a:t></a:r></a:p>
</xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>"""


def drawing_xml(anchors):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
            'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' + "".join(anchors) + "</xdr:wsDr>")


# column widths (index 0 = A) - must match build_workbook.py
DASH_W = [2, 22, 22, 22, 34, 3, 18, 18, 18, 18, 2]
SEARCH_W = [13, 20, 6, 26, 30, 11, 12, 13, 14, 11]
DASH_BUTTONS = [  # name, text, macro, color, size, c1, c2, (row1, row2) 0-based
    ("btnSubmit", "SUBMIT ENROLLMENT", "SubmitEnrollment", "2E7D32", 14, 1, 2, (17, 18)),
    ("btnUpdate", "UPDATE ENROLLMENT", "UpdateEnrollment", "1E88E5", 10, 3, 3, (17, 18)),
    ("btnClear", "CLEAR FORM", "ClearForm", "78909C", 10, 4, 4, (17, 18)),
    ("btnLoad", "LOAD FOR EDIT", "LoadEnrollmentForEdit", "1E88E5", 10, 6, 7, (7, 7)),
    ("btnCancel", "CANCEL ENROLLMENT", "CancelEnrollment", "C62828", 10, 8, 9, (7, 7)),
    ("btnSearch", "SEARCH STUDENT", "SearchStudent", "00897B", 10, 6, 9, (11, 11)),
]
SEARCH_BUTTONS = [("btnBack", "BACK TO DASHBOARD", "GoToDashboard", "37474F", 9, 7, 9, (2, 2))]

NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DRAWING_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
VBA_REL = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"


def add_drawing(files, sheet_part, drawing_name, anchors_xml):
    """Attach a new drawing part to a worksheet part (sheet must have no drawing yet)."""
    sheet_xml = files[sheet_part].decode("utf-8")
    assert "<drawing " not in sheet_xml, f"{sheet_part} already has a drawing"
    rels_part = sheet_part.replace("worksheets/", "worksheets/_rels/") + ".rels"
    if rels_part in files:
        rels = files[rels_part].decode("utf-8")
    else:
        rels = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
    rid = "rIdBtn1"
    rels = rels.replace("</Relationships>",
                        f'<Relationship Id="{rid}" Type="{DRAWING_REL}" Target="../drawings/{drawing_name}"/></Relationships>')
    files[rels_part] = rels.encode("utf-8")
    files[f"xl/drawings/{drawing_name}"] = drawing_xml(anchors_xml).encode("utf-8")
    # <drawing> goes after pageSetup/headerFooter/... and before legacyDrawing/tableParts/extLst
    tag = f'<drawing xmlns:r="{NS_REL}" r:id="{rid}"/>'
    m = re.search(r"<(legacyDrawing|legacyDrawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b", sheet_xml)
    if m:
        sheet_xml = sheet_xml[:m.start()] + tag + sheet_xml[m.start():]
    else:
        sheet_xml = sheet_xml.replace("</worksheet>", tag + "</worksheet>")
    files[sheet_part] = sheet_xml.encode("utf-8")
    ct = files["[Content_Types].xml"].decode("utf-8")
    ct = ct.replace("</Types>", f'<Override PartName="/xl/drawings/{drawing_name}" '
                    'ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>')
    files["[Content_Types].xml"] = ct.encode("utf-8")


def sheet_parts(files):
    """map sheet name -> (part path, codeName) using workbook.xml + rels."""
    wb = files["xl/workbook.xml"].decode("utf-8")
    rels = files["xl/_rels/workbook.xml.rels"].decode("utf-8")
    rmap = {}
    for m in re.finditer(r"<Relationship\b[^>]*>", rels):
        t = m.group(0)
        rid = re.search(r'Id="([^"]+)"', t).group(1)
        tgt = re.search(r'Target="([^"]+)"', t).group(1)
        rmap[rid] = tgt.lstrip("/") if tgt.startswith("/") else "xl/" + tgt
    out = []
    for m in re.finditer(r"<sheet\b[^>]*>", wb):
        t = m.group(0)
        name = re.search(r'name="([^"]+)"', t).group(1)
        rid = re.search(r'r:id="([^"]+)"', t).group(1)
        part = rmap[rid]
        cn = re.search(r'<sheetPr[^>]*codeName="([^"]+)"', files[part].decode("utf-8"))
        out.append((name, part, cn.group(1) if cn else None))
    return out


def main():
    with zipfile.ZipFile(SRC) as z:
        files = {n: z.read(n) for n in z.namelist()}

    sheets = sheet_parts(files)
    missing = [n for n, _, cn in sheets if not cn]
    assert not missing, f"sheets without codeName (rebuild with build_workbook.py): {missing}"
    assert b'codeName="ThisWorkbook"' in files["xl/workbook.xml"], "workbook has no codeName"
    part_of = {n: p for n, p, _ in sheets}

    vba = build_vba_project_bin([cn for _, _, cn in sheets], BAS)
    files["xl/vbaProject.bin"] = vba

    rels = files["xl/_rels/workbook.xml.rels"].decode("utf-8")
    rels = rels.replace("</Relationships>",
                        f'<Relationship Id="rIdVBA1" Type="{VBA_REL}" Target="vbaProject.bin"/></Relationships>')
    files["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")

    ct = files["[Content_Types].xml"].decode("utf-8")
    ct = ct.replace("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                    "application/vnd.ms-excel.sheet.macroEnabled.main+xml")
    ct = ct.replace("<Default Extension=\"xml\"",
                    '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default Extension="xml"', 1)
    files["[Content_Types].xml"] = ct.encode("utf-8")
    assert b"macroEnabled.main+xml" in files["[Content_Types].xml"]

    existing = [n for n in files if n.startswith("xl/drawings/drawing")]
    nxt = len(existing) + 1
    add_drawing(files, part_of["Dashboard"], f"drawing{nxt}.xml",
                [anchor_xml(i + 1, *b[:5], b[5], b[6], b[7], DASH_W) for i, b in enumerate(DASH_BUTTONS)])
    add_drawing(files, part_of["Student Search"], f"drawing{nxt + 1}.xml",
                [anchor_xml(i + 1, *b[:5], b[5], b[6], b[7], SEARCH_W) for i, b in enumerate(SEARCH_BUTTONS)])

    # remove the "run InstallUpKidzSystem" placeholder text in Dashboard!B18
    dash = files[part_of["Dashboard"]].decode("utf-8")
    dash = re.sub(r'(<c r="B18"[^>]*?)\s+t="inlineStr"([^>]*)><is><t>[^<]*</t></is></c>', r'\1\2/>', dash)
    files[part_of["Dashboard"]] = dash.encode("utf-8")

    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
        order = ["[Content_Types].xml"] + [n for n in files if n != "[Content_Types].xml"]
        for n in order:
            z.writestr(n, files[n])
    print("saved", OUT, f"(vbaProject.bin {len(vba)} bytes)")


if __name__ == "__main__":
    main()
