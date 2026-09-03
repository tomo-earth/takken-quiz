# -*- coding: utf-8 -*-
"""谷直美様邸_現場日報台帳_集計.xlsx に VBA（印刷メニュー）を組み込み .xlsm を生成する。

Excel COM が使えない Linux 環境向けに、MS-OVBA / MS-CFB 仕様どおり
vbaProject.bin をゼロから生成して zip に注入する。
生成後は oletools(olevba) でソースが完全一致で取り出せることを検証する。

構造は EPPlus（Excel で実績のある C# ライブラリ）が生成する VBA プロジェクトと
同じレイアウト（PROJECT / PROJECTwm / VBA/_VBA_PROJECT / VBA/dir / 各モジュール、
参照は stdole と Office、文書モジュール ThisWorkbook + SheetN を空で用意）。

使い方: python3 build_xlsm.py <入力xlsx> <出力xlsm>
"""
import re
import struct
import sys
import uuid
import zipfile

from vba_source import (VBA_CODE, MODULE_NAME, DOC_MODULE_TEMPLATE,
                        GUID_WORKBOOK, GUID_WORKSHEET, to_vba_bytes)

CODEPAGE = "cp932"

# ---------------------------------------------------------------- MS-OVBA 圧縮

def _copytoken_params(pos_in_chunk):
    diff = pos_in_chunk
    bit_count = 4
    while (1 << bit_count) < diff:
        bit_count += 1
    bit_count = max(bit_count, 4)
    length_mask = 0xFFFF >> bit_count
    max_length = length_mask + 3
    return bit_count, length_mask, max_length


def compress_chunk(data):
    """1チャンク（decompressed <=4096 bytes）を圧縮トークン列にする"""
    out = bytearray()
    pos = 0
    n = len(data)
    while pos < n:
        flag = 0
        tokens = bytearray()
        for i in range(8):
            if pos >= n:
                break
            bit_count, length_mask, max_length = _copytoken_params(pos)
            best_len = 0
            best_off = 0
            if pos >= 1:
                max_off = min(pos, 1 << bit_count)
                # 後方一致の最長を探す（単純探索で十分小さい）
                for off in range(1, max_off + 1):
                    src = pos - off
                    l = 0
                    while (pos + l < n and l < max_length
                           and data[src + l] == data[pos + l]):
                        l += 1
                    if l > best_len:
                        best_len = l
                        best_off = off
                        if l >= max_length:
                            break
            if best_len >= 3:
                token = ((best_off - 1) << (16 - bit_count)) | (best_len - 3)
                tokens += struct.pack("<H", token)
                flag |= 1 << i
                pos += best_len
            else:
                tokens.append(data[pos])
                pos += 1
        out.append(flag)
        out += tokens
    return bytes(out)


def compress_container(data):
    """MS-OVBA 2.4.1 CompressedContainer"""
    out = bytearray(b"\x01")
    for start in range(0, len(data), 4096):
        chunk = data[start:start + 4096]
        comp = compress_chunk(chunk)
        if len(comp) < 4096:
            size_field = (len(comp) + 2) - 3     # ヘッダ2byte込みサイズ-3
            header = 0x8000 | 0x3000 | size_field  # compressed, sig=0b011
            out += struct.pack("<H", header) + comp
        else:  # 圧縮で膨らむ場合は raw チャンク（4096byteちょうどに詰める）
            raw = chunk + b"\x00" * (4096 - len(chunk))
            header = 0x3000 | (4098 - 3)
            out += struct.pack("<H", header) + raw
    return bytes(out)

# ------------------------------------------------------------ dir ストリーム

def _rec(rec_id, payload):
    return struct.pack("<HI", rec_id, len(payload)) + payload


def build_dir_stream(project_name, modules, references):
    """modules: [(name, stream_name, is_document)] / references: [(name, libid)]"""
    cp = CODEPAGE
    out = bytearray()
    # PROJECTINFORMATION
    out += _rec(0x0001, struct.pack("<I", 1))               # SYSKIND: Win32
    out += _rec(0x0002, struct.pack("<I", 0x409))           # LCID
    out += _rec(0x0014, struct.pack("<I", 0x409))           # LCIDINVOKE
    out += _rec(0x0003, struct.pack("<H", 932))             # CODEPAGE
    out += _rec(0x0004, project_name.encode(cp))            # NAME
    out += _rec(0x0005, b"") + _rec(0x0040, b"")            # DOCSTRING
    out += _rec(0x0006, b"") + _rec(0x003D, b"")            # HELPFILEPATH
    out += _rec(0x0007, struct.pack("<I", 0))               # HELPCONTEXT
    out += _rec(0x0008, struct.pack("<I", 0))               # LIBFLAGS
    # VERSION: Id + Reserved(=4) + VersionMajor(4) + VersionMinor(2)
    out += struct.pack("<HIIH", 0x0009, 4, 1, 0)
    out += _rec(0x000C, b"") + _rec(0x003C, b"")            # CONSTANTS
    # PROJECTREFERENCES
    for ref_name, libid in references:
        out += _rec(0x0016, ref_name.encode(cp))            # REFERENCENAME
        out += _rec(0x003E, ref_name.encode("utf-16-le"))
        libid_b = libid.encode(cp)
        payload = struct.pack("<I", len(libid_b)) + libid_b + b"\x00\x00\x00\x00\x00\x00"
        out += _rec(0x000D, payload)                        # REFERENCEREGISTERED
    # PROJECTMODULES
    out += _rec(0x000F, struct.pack("<H", len(modules)))
    out += _rec(0x0013, struct.pack("<H", 0xFFFF))          # COOKIE
    for name, stream_name, is_document in modules:
        out += _rec(0x0019, name.encode(cp))                # MODULENAME
        out += _rec(0x0047, name.encode("utf-16-le"))       # MODULENAMEUNICODE
        out += _rec(0x001A, stream_name.encode(cp))         # MODULESTREAMNAME
        out += _rec(0x0032, stream_name.encode("utf-16-le"))
        out += _rec(0x001C, b"") + _rec(0x0048, b"")        # MODULEDOCSTRING
        out += _rec(0x0031, struct.pack("<I", 0))           # MODULEOFFSET
        out += _rec(0x001E, struct.pack("<I", 0))           # MODULEHELPCONTEXT
        out += _rec(0x002C, struct.pack("<H", 0xFFFF))      # MODULECOOKIE
        type_id = 0x0022 if is_document else 0x0021         # MODULETYPE
        out += struct.pack("<HI", type_id, 0)
        out += struct.pack("<HI", 0x002B, 0)                # モジュール終端
    out += struct.pack("<HI", 0x0010, 0)                    # dir 終端
    return bytes(out)

# ------------------------------------------- PROJECT ストリームの保護フィールド

def _project_key(project_id):
    return sum(project_id.encode("ascii")) & 0xFF


def encrypt_field(project_id, data, seed=0x2F):
    """MS-OVBA 2.4.3.2 Data Encryption（保護なしの既定値を暗号化するだけ）"""
    proj_key = _project_key(project_id)
    version = 2
    version_enc = seed ^ version
    proj_key_enc = seed ^ proj_key
    ignored_length = (seed & 6) >> 1
    plain = bytes(ignored_length) + struct.pack("<I", len(data)) + data
    out = [seed, version_enc, proj_key_enc]
    unenc_prev = proj_key
    enc_prev1 = proj_key_enc
    enc_prev2 = version_enc
    for b in plain:
        be = b ^ ((enc_prev2 + unenc_prev) & 0xFF)
        out.append(be)
        enc_prev2 = enc_prev1
        enc_prev1 = be
        unenc_prev = b
    return "".join(f"{x:02X}" for x in out)


def decrypt_field(project_id, hexstr):
    """検証用の復号（2.4.3.3）"""
    raw = bytes.fromhex(hexstr)
    seed, version_enc, proj_key_enc = raw[0], raw[1], raw[2]
    proj_key = _project_key(project_id)
    assert seed ^ version_enc == 2 and seed ^ proj_key_enc == proj_key
    ignored_length = (seed & 6) >> 1
    unenc_prev = proj_key
    enc_prev1 = proj_key_enc
    enc_prev2 = version_enc
    plain = bytearray()
    for be in raw[3:]:
        b = be ^ ((enc_prev2 + unenc_prev) & 0xFF)
        plain.append(b)
        enc_prev2 = enc_prev1
        enc_prev1 = be
        unenc_prev = b
    body = plain[ignored_length:]
    (length,) = struct.unpack("<I", body[:4])
    return bytes(body[4:4 + length])


def build_project_stream(project_id, doc_modules, std_modules):
    visibility = b"\xff"
    lines = [f'ID="{project_id}"']
    for m in doc_modules:
        lines.append(f"Document={m}/&H00000000")
    for m in std_modules:
        lines.append(f"Module={m}")
    lines += [
        'HelpFile=""',
        'Name="VBAProject"',
        'HelpContextID="0"',
        'VersionCompatible32="393222000"',
        f'CMG="{encrypt_field(project_id, bytes(4), seed=0x11)}"',
        f'DPB="{encrypt_field(project_id, bytes(1), seed=0x25)}"',
        f'GC="{encrypt_field(project_id, visibility, seed=0x33)}"',
        "",
        "[Host Extender Info]",
        "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000",
        "",
    ]
    return ("\r\n".join(lines)).encode(CODEPAGE)


def build_projectwm_stream(module_names):
    out = bytearray()
    for name in module_names:
        out += name.encode(CODEPAGE) + b"\x00"
        out += name.encode("utf-16-le") + b"\x00\x00"
    out += b"\x00\x00"
    return bytes(out)

# ---------------------------------------------------------------- CFB ライタ

FREESECT, ENDOFCHAIN, FATSECT = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD
NOSTREAM = 0xFFFFFFFF


class _Entry:
    def __init__(self, name, etype, data=None):
        self.name = name
        self.etype = etype          # 5=root, 1=storage, 2=stream
        self.data = data or b""
        self.children = []          # storage のみ
        self.left = self.right = self.child = NOSTREAM
        self.start = ENDOFCHAIN
        self.size = 0
        self.did = None


def _cfb_name_key(name):
    return (len(name), name.upper())


def _build_sibling_tree(entries):
    """ソート済みエントリ群から平衡二分木を作り、部分木の根の did を返す"""
    if not entries:
        return NOSTREAM
    mid = len(entries) // 2
    node = entries[mid]
    node.left = _build_sibling_tree(entries[:mid])
    node.right = _build_sibling_tree(entries[mid + 1:])
    return node.did


def write_cfb(root_children):
    """root_children: [_Entry] （root 直下）。CFB v3 バイト列を返す"""
    SECT = 512
    MINI = 64

    root = _Entry("Root Entry", 5)
    root.children = root_children

    # ディレクトリエントリを列挙（root, その後 幅優先）
    entries = [root]
    queue = [root]
    while queue:
        st = queue.pop(0)
        for ch in sorted(st.children, key=lambda e: _cfb_name_key(e.name)):
            entries.append(ch)
            if ch.etype == 1:
                queue.append(ch)
    for i, e in enumerate(entries):
        e.did = i
    for e in entries:
        if e.etype in (1, 5):
            kids = sorted(e.children, key=lambda k: _cfb_name_key(k.name))
            e.child = _build_sibling_tree(kids)

    # ミニストリーム構築（4096 未満のストリーム）
    mini_data = bytearray()
    minifat = []
    for e in entries:
        if e.etype != 2:
            continue
        e.size = len(e.data)
        if 0 < e.size < 4096:
            first = len(mini_data) // MINI
            nsect = (e.size + MINI - 1) // MINI
            mini_data += e.data + b"\x00" * (nsect * MINI - e.size)
            for k in range(nsect - 1):
                minifat.append(first + k + 1)
            minifat.append(ENDOFCHAIN)
            e.start = first
        elif e.size == 0:
            e.start = ENDOFCHAIN

    # 大ストリーム（>=4096） + ミニストリーム本体 + directory + miniFAT を
    # セクタ列に並べ、FAT を後ろに置く
    big_streams = [e for e in entries if e.etype == 2 and e.size >= 4096]

    def nsect(nbytes):
        return (nbytes + SECT - 1) // SECT

    dir_count = ((len(entries) + 3) // 4) * 4          # 4エントリ/セクタ
    dir_bytes = dir_count * 128
    minifat_bytes = len(minifat) * 4
    payload = []                                        # (kind, obj, nbytes)
    payload.append(("dir", None, dir_bytes))
    if minifat:
        payload.append(("minifat", None, minifat_bytes))
    if mini_data:
        payload.append(("ministream", root, len(mini_data)))
    for e in big_streams:
        payload.append(("stream", e, e.size))

    payload_sectors = sum(nsect(nb) for _, _, nb in payload)
    fat_sectors = 1
    while (payload_sectors + fat_sectors) > fat_sectors * (SECT // 4):
        fat_sectors += 1

    # セクタ割り当て（payload 先、FAT 最後）
    fat = [FREESECT] * (fat_sectors * (SECT // 4))
    sect_cursor = 0
    locations = {}
    for kind, obj, nb in payload:
        first = sect_cursor
        cnt = nsect(nb)
        for k in range(cnt - 1):
            fat[first + k] = first + k + 1
        fat[first + cnt - 1] = ENDOFCHAIN
        sect_cursor += cnt
        locations[(kind, id(obj))] = (first, cnt)
        if kind == "stream":
            obj.start = first
        elif kind == "ministream":
            root.start = first
            root.size = len(mini_data)
    fat_first = sect_cursor
    for k in range(fat_sectors):
        fat[fat_first + k] = FATSECT
    total_sectors = fat_first + fat_sectors

    dir_first = locations[("dir", id(None))][0]
    minifat_first = locations.get(("minifat", id(None)), (ENDOFCHAIN, 0))[0]
    minifat_count = locations.get(("minifat", id(None)), (0, 0))[1]

    # ヘッダ
    difat = [FREESECT] * 109
    for k in range(fat_sectors):
        difat[k] = fat_first + k
    header = struct.pack(
        "<8s16sHHHHH6sIIIIIIII",
        b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1", b"\x00" * 16,
        0x003E, 0x0003, 0xFFFE, 9, 6, b"\x00" * 6,
        0,                    # v3: dir sector 数は 0
        fat_sectors,
        dir_first,
        0,                    # transaction
        4096,                 # mini cutoff
        minifat_first if minifat else ENDOFCHAIN,
        minifat_count,
        ENDOFCHAIN,           # DIFAT 開始
    )
    header += struct.pack("<I", 0)          # DIFAT セクタ数
    header += b"".join(struct.pack("<I", v) for v in difat)
    assert len(header) == 512

    # ディレクトリストリーム
    def entry_bytes(e):
        if e is None:
            return (b"\x00" * 64 + struct.pack("<H", 0) + b"\x00\x00"
                    + struct.pack("<III", NOSTREAM, NOSTREAM, NOSTREAM)
                    + b"\x00" * 16 + b"\x00" * 4 + b"\x00" * 16
                    + struct.pack("<II", 0, 0))
        nm = e.name.encode("utf-16-le") + b"\x00\x00"
        nm_padded = nm + b"\x00" * (64 - len(nm))
        start = e.start if e.start != ENDOFCHAIN or e.etype == 2 else 0
        if e.etype == 5:
            start = root.start if mini_data else ENDOFCHAIN
            size = root.size
        else:
            size = e.size
        if e.etype == 2 and e.size == 0:
            start = ENDOFCHAIN
        return (nm_padded + struct.pack("<H", len(nm)) +
                bytes([e.etype, 1]) +                     # color=black
                struct.pack("<III", e.left, e.right, e.child) +
                b"\x00" * 16 +                            # CLSID
                b"\x00" * 4 +                             # state bits
                b"\x00" * 16 +                            # timestamps
                struct.pack("<IQ", start & 0xFFFFFFFF, size))

    dir_stream = bytearray()
    for i in range(dir_count):
        dir_stream += entry_bytes(entries[i] if i < len(entries) else None)

    # 本体を並べる
    body = bytearray(b"\x00" * (total_sectors * SECT))

    def put(first_sector, data):
        body[first_sector * SECT: first_sector * SECT + len(data)] = data

    put(dir_first, dir_stream)
    if minifat:
        mf = b"".join(struct.pack("<I", v) for v in minifat)
        put(minifat_first, mf)
    if mini_data:
        put(root.start, bytes(mini_data))
    for e in big_streams:
        put(e.start, e.data)
    fat_bytes = b"".join(struct.pack("<I", v) for v in fat)
    put(fat_first, fat_bytes[: fat_sectors * SECT])

    return bytes(header) + bytes(body)

# ------------------------------------------------------------- vbaProject.bin

REFERENCES = [
    ("stdole",
     r"*\G{00020430-0000-0000-C000-000000000046}#2.0#0"
     r"#C:\Windows\system32\stdole2.tlb#OLE Automation"),
    ("Office",
     r"*\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.0#0"
     r"#C:\Program Files\Common Files\Microsoft Shared\OFFICE14\MSO.DLL"
     r"#Microsoft Office 14.0 Object Library"),
]


def build_vba_project(sheet_count):
    project_id = "{" + str(uuid.uuid4()).upper() + "}"
    doc_modules = ["ThisWorkbook"] + [f"Sheet{i}" for i in range(1, sheet_count + 1)]
    module_sources = {}
    module_sources["ThisWorkbook"] = to_vba_bytes(
        DOC_MODULE_TEMPLATE.format(name="ThisWorkbook", guid=GUID_WORKBOOK))
    for i in range(1, sheet_count + 1):
        module_sources[f"Sheet{i}"] = to_vba_bytes(
            DOC_MODULE_TEMPLATE.format(name=f"Sheet{i}", guid=GUID_WORKSHEET))
    module_sources[MODULE_NAME] = to_vba_bytes(VBA_CODE)

    all_modules = doc_modules + [MODULE_NAME]
    dir_modules = [(m, m, m in doc_modules) for m in all_modules]
    dir_stream = compress_container(
        build_dir_stream("VBAProject", dir_modules, REFERENCES))

    vba_children = [
        _Entry("dir", 2, dir_stream),
        _Entry("_VBA_PROJECT", 2, b"\xcc\x61\xff\xff\x00\x00\x00"),
    ]
    for m in all_modules:
        vba_children.append(_Entry(m, 2, compress_container(module_sources[m])))

    vba_storage = _Entry("VBA", 1)
    vba_storage.children = vba_children
    root_children = [
        vba_storage,
        _Entry("PROJECT", 2, build_project_stream(project_id, doc_modules, [MODULE_NAME])),
        _Entry("PROJECTwm", 2, build_projectwm_stream(all_modules)),
    ]
    return write_cfb(root_children), project_id, module_sources

# ------------------------------------------------------------------ zip 注入

def build_xlsm(src_xlsx, dst_xlsm):
    zin = zipfile.ZipFile(src_xlsx)
    sheet_files = sorted(
        n for n in zin.namelist() if re.match(r"xl/worksheets/sheet\d+\.xml$", n))
    sheet_count = len(sheet_files)
    vba_bin, project_id, module_sources = build_vba_project(sheet_count)

    zout = zipfile.ZipFile(dst_xlsm, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        name = item.filename
        if name == "[Content_Types].xml":
            text = data.decode("utf-8")
            text = text.replace(
                'PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"',
                'PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"')
            text = text.replace(
                "<Default Extension=\"xml\"",
                "<Default Extension=\"bin\" ContentType=\"application/vnd.ms-office.vbaProject\"/><Default Extension=\"xml\"",
                1)
            data = text.encode("utf-8")
        elif name == "xl/_rels/workbook.xml.rels":
            text = data.decode("utf-8")
            rids = [int(m) for m in re.findall(r'Id="rId(\d+)"', text)]
            new_rid = max(rids) + 1
            text = text.replace(
                "</Relationships>",
                f'<Relationship Id="rId{new_rid}" '
                'Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" '
                'Target="vbaProject.bin"/></Relationships>')
            data = text.encode("utf-8")
        elif name == "xl/workbook.xml":
            text = data.decode("utf-8")
            if "<workbookPr/>" in text:
                text = text.replace("<workbookPr/>",
                                    '<workbookPr codeName="ThisWorkbook"/>')
            elif "<workbookPr " in text:
                text = text.replace("<workbookPr ",
                                    '<workbookPr codeName="ThisWorkbook" ', 1)
            data = text.encode("utf-8")
        elif re.match(r"xl/worksheets/sheet(\d+)\.xml$", name):
            num = re.match(r"xl/worksheets/sheet(\d+)\.xml$", name).group(1)
            text = data.decode("utf-8")
            if "<sheetPr>" in text:
                text = text.replace("<sheetPr>", f'<sheetPr codeName="Sheet{num}">', 1)
            elif "<sheetPr/>" in text:
                text = text.replace("<sheetPr/>", f'<sheetPr codeName="Sheet{num}"/>', 1)
            elif "<sheetPr " in text:
                text = text.replace("<sheetPr ", f'<sheetPr codeName="Sheet{num}" ', 1)
            data = text.encode("utf-8")
        zi = zipfile.ZipInfo(name, date_time=item.date_time)
        zi.compress_type = zipfile.ZIP_DEFLATED
        zi.external_attr = item.external_attr
        zout.writestr(zi, data)
    zout.writestr("xl/vbaProject.bin", vba_bin)
    zout.close()
    zin.close()
    return project_id, module_sources


# ------------------------------------------------------------------- 自己検証

def self_test():
    from oletools.olevba import decompress_stream
    import random
    rnd = random.Random(0)
    samples = [
        b"", b"a", b"abc" * 5,
        to_vba_bytes(VBA_CODE),
        bytes(rnd.randrange(256) for _ in range(10000)),      # 非圧縮チャンク経路
        b"x" * 5000,
        to_vba_bytes(VBA_CODE) * 3,
    ]
    for s in samples:
        assert decompress_stream(bytearray(compress_container(s))) == s, "圧縮検証NG"
    pid = "{12345678-1234-1234-1234-123456789012}"
    for data, seed in [(bytes(4), 0x11), (bytes(1), 0x25), (b"\xff", 0x33)]:
        assert decrypt_field(pid, encrypt_field(pid, data, seed)) == data, "暗号化検証NG"
    print("self_test OK（MS-OVBA圧縮 roundtrip / 保護フィールド暗号化）")


if __name__ == "__main__":
    self_test()
    src = sys.argv[1] if len(sys.argv) > 1 else "谷直美様邸_現場日報台帳_集計.xlsx"
    dst = sys.argv[2] if len(sys.argv) > 2 else "谷直美様邸_現場日報台帳_集計.xlsm"
    project_id, module_sources = build_xlsm(src, dst)
    print(f"wrote {dst} (project ID {project_id}, modules: {list(module_sources)})")
