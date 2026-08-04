#!/usr/bin/env python3
"""Generate a vbaProject.bin (MS-OVBA inside a CFB container) from VBA source.

Implements:
  - [MS-CFB] minimal writer (v3, 512-byte sectors, mini-stream for small streams)
  - [MS-OVBA] dir stream, module streams, PROJECT / PROJECTwm / _VBA_PROJECT
Validated with olefile (container) and oletools.olevba (VBA extraction).
"""
import struct
import uuid
from ms_ovba_compression.ms_ovba import MsOvba


# --- Fix upstream bug in ms_ovba_compression 1.0.1: copytoken_help computed
# max_length as (0xFFFF << bit_count + 3) instead of (0xFFFF >> bit_count) + 3,
# letting copy-token lengths overflow into the offset bits on long runs.
def _copytoken_help_fixed(difference):
    bit_count = MsOvba.ceil_log2(difference)
    length_mask = 0xFFFF >> bit_count
    return {
        "lengthMask": length_mask,
        "offsetMask": ~length_mask & 0xFFFF,
        "bitCount": bit_count,
        "maxLength": length_mask + 3,
    }


MsOvba.copytoken_help = staticmethod(_copytoken_help_fixed)

CODEPAGE = 1252  # keep all VBA source ASCII-only

# ---------------------------------------------------------------- CFB writer

FREESECT = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD
NOSTREAM = 0xFFFFFFFF


class Entry:
    def __init__(self, name, typ, data=b"", children=None):
        self.name = name          # str
        self.typ = typ            # 1 storage, 2 stream, 5 root
        self.data = data          # bytes (streams only)
        self.children = children or []
        # filled during layout
        self.index = None
        self.left = NOSTREAM
        self.right = NOSTREAM
        self.child = NOSTREAM
        self.start = ENDOFCHAIN
        self.size = 0


def _cfb_name_key(e):
    return (len(e.name), e.name.upper())


def _build_bst(entries):
    """Return index of root of a balanced BST over CFB-sorted entries."""
    entries = sorted(entries, key=_cfb_name_key)

    def build(lo, hi):
        if lo > hi:
            return NOSTREAM
        mid = (lo + hi) // 2
        e = entries[mid]
        e.left = build(lo, mid - 1)
        e.right = build(mid + 1, hi)
        return e.index

    return build(0, len(entries) - 1)


def write_cfb(root):
    # flatten: root first, then DFS storages/streams
    flat = []

    def walk(e):
        e.index = len(flat)
        flat.append(e)
        for c in e.children:
            walk(c)

    walk(root)
    # sibling trees
    for e in flat:
        if e.children:
            e.child = _build_bst(e.children)

    SEC = 512
    MINISEC = 64
    CUTOFF = 4096

    # ---- mini stream layout (all streams < 4096 bytes)
    mini_data = bytearray()
    minifat = []
    big_streams = []
    for e in flat:
        if e.typ != 2:
            continue
        e.size = len(e.data)
        if e.size == 0:
            e.start = ENDOFCHAIN
            continue
        if e.size < CUTOFF:
            nsec = (e.size + MINISEC - 1) // MINISEC
            e.start = len(minifat)
            for i in range(nsec - 1):
                minifat.append(len(minifat) + 1)
            minifat.append(ENDOFCHAIN)
            mini_data += e.data.ljust(nsec * MINISEC, b"\x00")
        else:
            big_streams.append(e)

    # ---- sector layout
    ndirsec = (len(flat) * 128 + SEC - 1) // SEC
    nminifatsec = (len(minifat) * 4 + SEC - 1) // SEC if minifat else 0
    nministreamsec = (len(mini_data) + SEC - 1) // SEC
    nbigsec = sum((e.size + SEC - 1) // SEC for e in big_streams)

    # iterate: FAT sector count depends on total sectors
    nfat = 1
    while True:
        total = nfat + ndirsec + nminifatsec + nministreamsec + nbigsec
        need = (total * 4 + SEC - 1) // SEC
        if need <= nfat:
            break
        nfat = need

    fat = [FREESECT] * (nfat * SEC // 4)
    sec = 0
    fat_secs = list(range(sec, sec + nfat))
    for s in fat_secs:
        fat[s] = FATSECT
    sec += nfat

    dir_start = sec
    for i in range(ndirsec):
        fat[sec] = sec + 1 if i < ndirsec - 1 else ENDOFCHAIN
        sec += 1

    minifat_start = ENDOFCHAIN
    if nminifatsec:
        minifat_start = sec
        for i in range(nminifatsec):
            fat[sec] = sec + 1 if i < nminifatsec - 1 else ENDOFCHAIN
            sec += 1

    ministream_start = ENDOFCHAIN
    if nministreamsec:
        ministream_start = sec
        for i in range(nministreamsec):
            fat[sec] = sec + 1 if i < nministreamsec - 1 else ENDOFCHAIN
            sec += 1

    for e in big_streams:
        n = (e.size + SEC - 1) // SEC
        e.start = sec
        for i in range(n):
            fat[sec] = sec + 1 if i < n - 1 else ENDOFCHAIN
            sec += 1

    root.start = ministream_start
    root.size = len(mini_data)

    # ---- directory sectors
    dirbytes = bytearray()
    for e in flat:
        name16 = e.name.encode("utf-16-le") + b"\x00\x00"
        assert len(name16) <= 64
        ent = bytearray(128)
        ent[0:len(name16)] = name16
        struct.pack_into("<H", ent, 64, len(name16))
        ent[66] = e.typ
        ent[67] = 1  # black
        struct.pack_into("<LLL", ent, 68, e.left, e.right, e.child)
        # clsid 16 zero, state 4 zero, times 16 zero
        struct.pack_into("<L", ent, 116, e.start if (e.typ != 1) else 0)
        struct.pack_into("<Q", ent, 120, e.size if e.typ in (2, 5) else 0)
        if e.typ == 1:
            struct.pack_into("<L", ent, 116, 0)
        dirbytes += ent
    dirbytes = dirbytes.ljust(ndirsec * SEC, b"\x00")
    # unused dir entries must read as empty: set left/right/child NOSTREAM
    for i in range(len(flat), ndirsec * SEC // 128):
        off = i * 128
        struct.pack_into("<LLL", dirbytes, off + 68, NOSTREAM, NOSTREAM, NOSTREAM)

    # ---- header
    hdr = bytearray(SEC)
    hdr[0:8] = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
    struct.pack_into("<H", hdr, 24, 0x003E)      # minor
    struct.pack_into("<H", hdr, 26, 0x0003)      # major v3
    struct.pack_into("<H", hdr, 28, 0xFFFE)      # little endian
    struct.pack_into("<H", hdr, 30, 9)           # sector shift
    struct.pack_into("<H", hdr, 32, 6)           # mini shift
    struct.pack_into("<L", hdr, 44, nfat)
    struct.pack_into("<L", hdr, 48, dir_start)
    struct.pack_into("<L", hdr, 56, CUTOFF)
    struct.pack_into("<L", hdr, 60, minifat_start if minifat else ENDOFCHAIN)
    struct.pack_into("<L", hdr, 64, nminifatsec)
    struct.pack_into("<L", hdr, 68, ENDOFCHAIN)  # first DIFAT sector
    struct.pack_into("<L", hdr, 72, 0)           # n DIFAT sectors
    for i in range(109):
        struct.pack_into("<L", hdr, 76 + i * 4, fat_secs[i] if i < len(fat_secs) else FREESECT)

    # ---- assemble
    out = bytearray(hdr)
    out += b"".join(struct.pack("<L", v) for v in fat)
    out += dirbytes
    if nminifatsec:
        mf = b"".join(struct.pack("<L", v) for v in minifat)
        out += mf.ljust(nminifatsec * SEC, b"\xff")  # pad with FREESECT
    out += mini_data.ljust(nministreamsec * SEC, b"\x00")
    for e in big_streams:
        n = (e.size + SEC - 1) // SEC
        out += e.data.ljust(n * SEC, b"\x00")
    return bytes(out)


# ------------------------------------------------------------- OVBA streams

def _rec(rid, payload):
    return struct.pack("<HL", rid, len(payload)) + payload


def build_dir_stream(project_name, modules):
    """modules: list of (name, is_document)."""
    cp = "cp%d" % CODEPAGE
    b = b""
    # PROJECTINFORMATION
    b += _rec(0x0001, struct.pack("<L", 1))                     # SysKind Win32
    b += _rec(0x0002, struct.pack("<L", 0x0409))                # LCID
    b += _rec(0x0014, struct.pack("<L", 0x0409))                # LCIDINVOKE
    b += _rec(0x0003, struct.pack("<H", CODEPAGE))              # CODEPAGE
    b += _rec(0x0004, project_name.encode(cp))                  # NAME
    b += _rec(0x0005, b"") + _rec(0x0040, b"")                  # DOCSTRING
    b += _rec(0x0006, b"") + _rec(0x003D, b"")                  # HELPFILEPATH
    b += _rec(0x0007, struct.pack("<L", 0))                     # HELPCONTEXT
    b += _rec(0x0008, struct.pack("<L", 0))                     # LIBFLAGS
    b += struct.pack("<HLLH", 0x0009, 4, 0x0197B3C9, 0x0009)    # VERSION
    b += _rec(0x000C, b"") + _rec(0x003C, b"")                  # CONSTANTS
    # PROJECTREFERENCES: stdole + Office
    refs = [
        ("stdole",
         "*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#"
         "C:\\Windows\\System32\\stdole2.tlb#OLE Automation"),
        ("Office",
         "*\\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.0#0#"
         "C:\\Program Files\\Common Files\\Microsoft Shared\\OFFICE16\\MSO.DLL"
         "#Microsoft Office 16.0 Object Library"),
    ]
    for name, libid in refs:
        b += _rec(0x0016, name.encode(cp))                      # REFERENCENAME
        b += _rec(0x003E, name.encode("utf-16-le"))             # ...unicode
        libid_b = libid.encode(cp)
        payload = struct.pack("<L", len(libid_b)) + libid_b + b"\x00\x00\x00\x00\x00\x00"
        b += _rec(0x000D, payload)                              # REFERENCEREGISTERED
    # PROJECTMODULES
    b += _rec(0x000F, struct.pack("<H", len(modules)))
    b += _rec(0x0013, struct.pack("<H", 0xFFFF))                # COOKIE
    for name, is_doc in modules:
        nm = name.encode(cp)
        b += _rec(0x0019, nm)                                   # MODULENAME
        b += _rec(0x0047, name.encode("utf-16-le"))             # ...unicode
        b += _rec(0x001A, nm)                                   # STREAMNAME
        b += _rec(0x0032, name.encode("utf-16-le"))
        b += _rec(0x001C, b"") + _rec(0x0048, b"")              # DOCSTRING
        b += _rec(0x0031, struct.pack("<L", 0))                 # OFFSET
        b += _rec(0x001E, struct.pack("<L", 0))                 # HELPCONTEXT
        b += _rec(0x002C, struct.pack("<H", 0xFFFF))            # COOKIE
        b += struct.pack("<HLH", 0x0022 if is_doc else 0x0021, 0, 0x002B)
        b += struct.pack("<L", 0)                               # terminator size
    b += struct.pack("<HL", 0x0010, 0)                          # dir terminator
    return b


def build_project_stream(project_id, doc_modules, std_modules):
    lines = ['ID="%s"' % project_id]
    for m in doc_modules:
        lines.append("Document=%s/&H00000000" % m)
    for m in std_modules:
        lines.append("Module=%s" % m)
    lines += [
        'Name="VBAProject"',
        'HelpContextID="0"',
        'VersionCompatible32="393222000"',
        "",
        "[Host Extender Info]",
        "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000",
        "",
        "[Workspace]",
    ]
    for m in doc_modules + std_modules:
        lines.append("%s=0, 0, 0, 0, C" % m)
    return ("\r\n".join(lines) + "\r\n").encode("cp%d" % CODEPAGE)


def build_projectwm(module_names):
    b = b""
    for m in module_names:
        b += m.encode("cp%d" % CODEPAGE) + b"\x00" + m.encode("utf-16-le") + b"\x00\x00"
    return b + b"\x00\x00"


def _compress_checked(ovba, data):
    """Compress and assert the result decompresses back to the input.

    Verified with oletools' independent decompressor (ms_ovba_compression's
    own decompress() mis-advances the buffer on multi-chunk streams).
    """
    from oletools.olevba import decompress_stream
    c = ovba.compress(data)
    assert decompress_stream(bytearray(c)) == data, \
        "OVBA compression round-trip failed"
    return c


def make_vba_project(modules_src, doc_modules):
    """modules_src: dict name -> VBA source (str, ASCII). doc_modules: set of names."""
    ovba = MsOvba()
    names = list(modules_src)
    dir_stream = build_dir_stream(
        "VBAProject", [(n, n in doc_modules) for n in names])
    vba_children = [
        Entry("dir", 2, _compress_checked(ovba, dir_stream)),
        Entry("_VBA_PROJECT", 2, b"\xcc\x61\xff\xff\x00\x00\x00"),
    ]
    for n in names:
        src = modules_src[n].replace("\n", "\r\n").replace("\r\r\n", "\r\n")
        vba_children.append(Entry(n, 2, _compress_checked(ovba, src.encode("cp%d" % CODEPAGE))))
    project_id = "{%s}" % str(uuid.uuid4()).upper()
    root = Entry("Root Entry", 5, children=[
        Entry("VBA", 1, children=vba_children),
        Entry("PROJECT", 2, build_project_stream(
            project_id,
            [n for n in names if n in doc_modules],
            [n for n in names if n not in doc_modules])),
        Entry("PROJECTwm", 2, build_projectwm(names)),
    ])
    return write_cfb(root)


if __name__ == "__main__":
    import sys
    sys.exit("import and call make_vba_project()")
