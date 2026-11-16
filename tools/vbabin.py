"""Minimal vbaProject.bin writer (MS-OVBA + MS-CFB), source-only (no p-code)."""
import struct, uuid, random

CP = 'cp932'

# ---------------- MS-OVBA compression ----------------
def _compress_chunk(chunk):
    out = bytearray(); i = 0; n = len(chunk); table = {}
    def add(p):
        if p + 3 <= n:
            table.setdefault(chunk[p:p+3], []).append(p)
    while i < n:
        fpos = len(out); out.append(0); flags = 0
        for bit in range(8):
            if i >= n: break
            bc = max((i - 1).bit_length(), 4) if i > 0 else 4
            maxlen = (0xFFFF >> bc) + 3
            maxoff = 1 << bc
            best_len = 0; best_off = 0
            if i + 3 <= n:
                for p in reversed(table.get(chunk[i:i+3], [])):
                    off = i - p
                    if off > maxoff: break
                    l = 0; lim = min(maxlen, n - i)
                    while l < lim and chunk[p + l] == chunk[i + l]: l += 1
                    if l > best_len:
                        best_len, best_off = l, off
                        if l == lim: break
            if best_len >= 3:
                out += struct.pack('<H', ((best_off - 1) << (16 - bc)) | (best_len - 3))
                flags |= 1 << bit
                for p in range(i, i + best_len): add(p)
                i += best_len
            else:
                out.append(chunk[i]); add(i); i += 1
        out[fpos] = flags
    return bytes(out)

def compress(data):
    out = bytearray([1])
    for pos in range(0, len(data), 4096):
        chunk = data[pos:pos+4096]
        comp = _compress_chunk(chunk)
        if len(comp) > 4096:
            assert len(chunk) == 4096
            out += struct.pack('<H', 0x3000 | 4095) + chunk
        else:
            out += struct.pack('<H', 0xB000 | (len(comp) + 2 - 3)) + comp
    return bytes(out)

def decompress(d):
    assert d[0] == 1; out = bytearray(); i = 1
    while i < len(d):
        hdr = struct.unpack('<H', d[i:i+2])[0]; size = (hdr & 0xfff) + 3
        chunk = d[i+2:i+size]; i += size; start = len(out)
        if not hdr >> 15: out += chunk; continue
        j = 0
        while j < len(chunk):
            fb = chunk[j]; j += 1
            for b in range(8):
                if j >= len(chunk): break
                if not (fb >> b) & 1: out.append(chunk[j]); j += 1
                else:
                    tok = struct.unpack('<H', chunk[j:j+2])[0]; j += 2
                    dpos = len(out) - start; bc = max((dpos - 1).bit_length(), 4)
                    ln = (tok & (0xffff >> bc)) + 3; off = (tok >> (16 - bc)) + 1
                    for _ in range(ln): out.append(out[-off])
    return bytes(out)

# ---------------- PROJECT stream encryption (CMG/DPB/GC) ----------------
def _encrypt(data, proj_id):
    seed = random.randint(0, 255); version = 2
    projkey = sum(proj_id.encode('ascii')) & 0xFF
    out = [seed, seed ^ version, seed ^ projkey]
    unenc1 = projkey; enc1 = out[2]; enc2 = out[1]
    def emit(b):
        nonlocal unenc1, enc1, enc2
        e = b ^ ((enc2 + unenc1) & 0xFF)
        out.append(e); enc2 = enc1; enc1 = e; unenc1 = b
    for _ in range((seed & 6) // 2): emit(random.randint(0, 255))
    for b in struct.pack('<I', len(data)): emit(b)
    for b in data: emit(b)
    return ''.join('%02X' % b for b in out)

# ---------------- dir stream ----------------
def _rec(rid, payload):
    return struct.pack('<HI', rid, len(payload)) + payload

def _dir_stream(project, modules):
    b = bytearray()
    b += _rec(0x01, struct.pack('<I', 1))          # SYSKIND win32
    b += _rec(0x02, struct.pack('<I', 0x409))      # LCID
    b += _rec(0x14, struct.pack('<I', 0x409))      # LCIDINVOKE
    b += _rec(0x03, struct.pack('<H', 932))        # CODEPAGE
    b += _rec(0x04, project.encode(CP))            # NAME
    b += _rec(0x05, b'') + _rec(0x40, b'')         # DOCSTRING
    b += _rec(0x06, b'') + _rec(0x3D, b'')         # HELPFILEPATH
    b += _rec(0x07, struct.pack('<I', 0))          # HELPCONTEXT
    b += _rec(0x08, struct.pack('<I', 0))          # LIBFLAGS
    b += struct.pack('<HIIH', 0x09, 4, 1, 0)       # VERSION
    b += _rec(0x0C, b'') + _rec(0x3C, b'')         # CONSTANTS
    b += _rec(0x0F, struct.pack('<H', len(modules)))
    b += _rec(0x13, struct.pack('<H', 0xFFFF))
    for name, _src, is_doc in modules:
        b += _rec(0x19, name.encode(CP))
        b += _rec(0x47, name.encode('utf-16-le'))
        b += _rec(0x1A, name.encode(CP)) + _rec(0x32, name.encode('utf-16-le'))
        b += _rec(0x1C, b'') + _rec(0x48, b'')
        b += _rec(0x31, struct.pack('<I', 0))      # text offset 0 (no p-code)
        b += _rec(0x1E, struct.pack('<I', 0))
        b += _rec(0x2C, struct.pack('<H', 0xFFFF))
        b += _rec(0x22 if is_doc else 0x21, b'')
        b += _rec(0x2B, b'')
    b += _rec(0x10, b'')
    return bytes(b)

# ---------------- MS-CFB writer ----------------
ENDOFCHAIN, FREESECT, FATSECT, NOSTREAM = 0xFFFFFFFE, 0xFFFFFFFF, 0xFFFFFFFD, 0xFFFFFFFF

class _E:
    def __init__(s, name, typ, data=b'', children=None):
        s.name, s.typ, s.data, s.children = name, typ, data, children or []
        s.left = s.right = s.child = NOSTREAM; s.start = ENDOFCHAIN; s.size = 0

def _key(e): return (len(e.name), e.name.upper())

def cfb(tree):
    """tree: dict name -> bytes | dict (storage)."""
    def build(name, node, typ):
        if isinstance(node, dict):
            return _E(name, typ, children=[build(k, v, 2 if isinstance(v, bytes) else 1) for k, v in node.items()])
        return _E(name, 2, node)
    root = build('Root Entry', tree, 5)
    entries = []
    def flatten(e):
        entries.append(e)
        for c in e.children: flatten(c)
    flatten(root)
    idx = {id(e): i for i, e in enumerate(entries)}
    def bal(lst):
        if not lst: return NOSTREAM
        m = len(lst) // 2; e = lst[m]
        e.left = bal(lst[:m]); e.right = bal(lst[m+1:])
        return idx[id(e)]
    for e in entries:
        if e.children: e.child = bal(sorted(e.children, key=_key))

    SS = 512
    big = [e for e in entries if e.typ == 2 and len(e.data) >= 4096]
    small = [e for e in entries if e.typ == 2 and len(e.data) < 4096]
    mini = bytearray(); minifat = []
    for e in small:
        e.size = len(e.data)
        if not e.data: e.start = ENDOFCHAIN; continue
        n = (len(e.data) + 63) // 64; first = len(minifat); e.start = first
        minifat += [first + k + 1 for k in range(n - 1)] + [ENDOFCHAIN]
        mini += e.data + b'\0' * (n * 64 - len(e.data))
    sectors = []; fat = []
    def place(data):
        n = (len(data) + SS - 1) // SS; first = len(sectors)
        for k in range(n): sectors.append(data[k*SS:(k+1)*SS].ljust(SS, b'\0'))
        fat.extend([first + k + 1 for k in range(n - 1)] + [ENDOFCHAIN])
        return first, n
    for e in big:
        e.size = len(e.data); e.start, _ = place(e.data)
    root.size = len(mini)
    root.start = place(bytes(mini))[0] if mini else ENDOFCHAIN
    mf_bytes = b''.join(struct.pack('<I', v) for v in minifat)
    mf_bytes += struct.pack('<I', FREESECT) * ((-len(minifat)) % 128)
    mf_start, mf_n = place(mf_bytes) if minifat else (ENDOFCHAIN, 0)
    d = bytearray()
    for e in entries:
        nm = e.name.encode('utf-16-le') + b'\0\0'
        d += nm.ljust(64, b'\0') + struct.pack('<HBB', len(nm), e.typ, 1)
        d += struct.pack('<III', e.left, e.right, e.child) + b'\0' * 16 + b'\0' * 4 + b'\0' * 16
        d += struct.pack('<IQ', e.start, e.size)
    while len(d) % SS:
        d += b'\0' * 64 + struct.pack('<HBB', 0, 0, 0) + struct.pack('<III', NOSTREAM, NOSTREAM, NOSTREAM) + b'\0' * 36 + struct.pack('<IQ', 0, 0)
    dir_start, _ = place(bytes(d))
    nfat = 1
    while (len(sectors) + nfat) * 4 > nfat * SS: nfat += 1
    assert nfat <= 109
    fat_first = len(sectors)
    fat += [FATSECT] * nfat
    fat += [FREESECT] * (nfat * 128 - len(fat))
    fb = b''.join(struct.pack('<I', v) for v in fat)
    for k in range(nfat): sectors.append(fb[k*SS:(k+1)*SS])
    difat = [fat_first + k for k in range(nfat)] + [FREESECT] * (109 - nfat)
    hdr = bytes.fromhex('D0CF11E0A1B11AE1') + b'\0' * 16
    hdr += struct.pack('<HHHHH', 0x3E, 3, 0xFFFE, 9, 6) + b'\0' * 6
    hdr += struct.pack('<IIIIIIIII', 0, nfat, dir_start, 0, 4096, mf_start, mf_n, ENDOFCHAIN, 0)
    hdr += b''.join(struct.pack('<I', v) for v in difat)
    assert len(hdr) == 512
    return hdr + b''.join(sectors)

# ---------------- project ----------------
WB_BASE = '0{00020819-0000-0000-C000-000000000046}'
WS_BASE = '0{00020820-0000-0000-C000-000000000046}'

def _doc_src(name, base):
    return ('Attribute VB_Name = "%s"\r\nAttribute VB_Base = "%s"\r\n'
            'Attribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\n'
            'Attribute VB_PredeclaredId = True\r\nAttribute VB_Exposed = True\r\n'
            'Attribute VB_TemplateDerived = False\r\nAttribute VB_Customizable = True\r\n') % (name, base)

def build_vba_project(std_modules, sheet_codenames, workbook_codename='ThisWorkbook', project='VBAProject'):
    """std_modules: list of (name, source_text). Returns vbaProject.bin bytes."""
    mods = [(workbook_codename, _doc_src(workbook_codename, WB_BASE), True)]
    mods += [(n, _doc_src(n, WS_BASE), True) for n in sheet_codenames]
    mods += [(n, src.replace('\r\n', '\n').replace('\n', '\r\n'), False) for n, src in std_modules]
    pid = '{%s}' % str(uuid.uuid4()).upper()
    lines = ['ID="%s"' % pid]
    for n, _, doc in mods:
        lines.append(('Document=%s/&H00000000' % n) if doc else ('Module=%s' % n))
    lines += ['Name="%s"' % project, 'HelpContextID="0"', 'VersionCompatible32="393222000"',
              'CMG="%s"' % _encrypt(b'\0\0\0\0', pid), 'DPB="%s"' % _encrypt(b'\0', pid),
              'GC="%s"' % _encrypt(b'\xff', pid), '',
              '[Host Extender Info]', '&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000', '',
              '[Workspace]']
    lines += ['%s=0, 0, 0, 0, C' % n for n, _, _ in mods]
    project_stream = ('\r\n'.join(lines) + '\r\n').encode(CP)
    wm = b''.join(n.encode(CP) + b'\0' + n.encode('utf-16-le') + b'\0\0' for n, _, _ in mods) + b'\0\0'
    vba = {'_VBA_PROJECT': bytes([0xCC, 0x61, 0xFF, 0xFF, 0x00, 0x00, 0x00]),
           'dir': compress(_dir_stream(project, mods))}
    for n, src, _ in mods:
        vba[n] = compress(src.encode(CP))
    return cfb({'VBA': vba, 'PROJECT': project_stream, 'PROJECTwm': wm})
