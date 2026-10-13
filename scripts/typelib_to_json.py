#!/usr/bin/env python3
"""Describe the type libraries embedded in a file as the engine's JSON.

The session takes external libraries as plain data (`Library`, see
`crates/vb6-engine/src/session/library.rs`). This script produces that data
from a compiled type library (`.tlb`, `.olb`) or from a binary that embeds
one (`.dll`, `.ocx`): it finds every embedded type library, reads its types
and members, and writes a JSON array of libraries.

    typelib_to_json.py FILE [--library NAME]... [--out OUT.json]

With no `--library`, every library found is written. The output is
deterministic: types and members keep the order of the type library.

Only what name resolution needs is kept: type kinds, the interface and
coclass links, member names and kinds, result type names, parameter counts
and the default member. Values of constants, help strings and GUIDs are not.
"""

import argparse
import json
import struct
import sys

# ── type library reader ─────────────────────────────────────────────────────

MAGIC = b"MSFT\x02\x00\x01\x00"

VT_NAMES = {
    0: "void", 2: "Integer", 3: "Long", 4: "Single", 5: "Double", 6: "Currency",
    7: "Date", 8: "String", 9: "Object", 10: "Long", 11: "Boolean", 12: "Variant",
    13: "IUnknown", 14: "Decimal", 16: "Byte", 17: "Byte", 18: "Integer",
    19: "Long", 20: "Long", 21: "Long", 22: "Long", 23: "Long", 24: "void",
    25: "HRESULT", 30: "String", 31: "String",
}
VT_PTR, VT_SAFEARRAY, VT_USERDEFINED = 26, 27, 29

WELL_KNOWN = {
    "00000000-0000-0000-c000-000000000046": "IUnknown",
    "00020400-0000-0000-c000-000000000046": "IDispatch",
    # Types of the OLE automation library, qualified by its name.
    "0be35203-8f91-11ce-9de3-00aa004bb851": "stdole.StdFont",
    "0be35204-8f91-11ce-9de3-00aa004bb851": "stdole.StdPicture",
    "bef6e003-a874-101a-8bba-00aa00300cab": "stdole.Font",
    "7bf80981-bf32-101a-8bba-00aa00300cab": "stdole.Picture",
}

TKIND_ENUM, TKIND_RECORD, TKIND_MODULE, TKIND_INTERFACE, TKIND_DISPATCH, TKIND_COCLASS = range(6)
TYPEFLAG_APPOBJECT = 0x1
TYPEFLAG_DUAL, TYPEFLAG_NONEXTENSIBLE, TYPEFLAG_DISPATCHABLE = 0x40, 0x80, 0x1000
FUNCFLAG_RESTRICTED = 0x1
PARAMFLAG_LCID, PARAMFLAG_RETVAL = 0x4, 0x8
IMPLTYPE_DEFAULT, IMPLTYPE_SOURCE = 0x1, 0x2
INVOKE_FUNC, INVOKE_PROPGET, INVOKE_PROPPUT, INVOKE_PROPPUTREF = 1, 2, 4, 8
SEGMENTS = [
    "typeinfo", "impinfo", "impfiles", "reftab", "libtab", "guid", "res6", "nametab",
    "strtab", "typdesc", "arraydesc", "custdata", "cdguids", "r13", "r14",
]


class TypeLib:
    def __init__(self, data, base):
        self.d, self.b = data, base
        self.varflags = self.u32(20)
        self.count = self.u32(32)
        pos = 84 + (4 if self.varflags & 0x100 else 0) + 4 * self.count
        self.seg = {}
        for k, name in enumerate(SEGMENTS):
            off, length, _, _ = struct.unpack_from("<iiii", data, base + pos + 16 * k)
            self.seg[name] = (off, length)
        self.name_offset = self.u32(56)

    def u32(self, o):
        return struct.unpack_from("<I", self.d, self.b + o)[0]

    def i32(self, o):
        return struct.unpack_from("<i", self.d, self.b + o)[0]

    def u16(self, o):
        return struct.unpack_from("<H", self.d, self.b + o)[0]

    def i16(self, o):
        return struct.unpack_from("<h", self.d, self.b + o)[0]

    def name(self, off):
        if off < 0:
            return None
        o = self.seg["nametab"][0] + off
        n = self.d[self.b + o + 8] & 0xFF
        return self.d[self.b + o + 12:self.b + o + 12 + n].decode("latin1")

    def info(self, idx):
        o = self.seg["typeinfo"][0] + idx * 100
        return [self.i32(o + 4 * k) for k in range(25)]

    def guid(self, off):
        o = self.b + self.seg["guid"][0] + off
        a, b, c = struct.unpack_from("<IHH", self.d, o)
        tail = self.d[o + 8:o + 16]
        return "%08x-%04x-%04x-%s-%s" % (a, b, c, tail[:2].hex(), tail[2:].hex())

    def href_name(self, href):
        if href & 1:
            o = self.seg["impinfo"][0] + (href & ~3)
            g = self.guid(self.i32(o + 8))
            return WELL_KNOWN.get(g, "{%s}" % g.upper())
        idx, rem = divmod(href, 100)
        if rem == 0 and 0 <= idx < self.count:
            return self.name(self.info(idx)[13])
        return "{unresolved}"

    def type_name(self, v):
        if v < 0:
            return VT_NAMES.get(v & 0xFFF, "Variant")
        o = self.seg["typdesc"][0] + v
        vt, data = self.u32(o) & 0xFFF, self.i32(o + 4)
        if vt == VT_PTR:
            return self.type_name(data)
        if vt == VT_SAFEARRAY:
            return self.type_name(data) + "()"
        if vt == VT_USERDEFINED:
            return self.href_name(data)
        return VT_NAMES.get(vt, "Variant")


def read_type(tl, idx):
    b = tl.info(idx)
    kind, n_funcs, n_vars = b[0] & 15, b[6] & 0xFFFF, (b[6] >> 16) & 0xFFFF
    t = {"name": tl.name(b[13]), "kind": kind, "flags": b[12] & 0xFFFFFFFF,
         "funcs": [], "vars": [], "impl": []}
    n = n_funcs + n_vars
    if n:
        recs = b[1] + 4
        after = recs + tl.i32(b[1])
        names = [tl.i32(after + 4 * n + 4 * k) for k in range(n)]
        offs = [tl.i32(after + 8 * n + 4 * k) for k in range(n)]
        for k in range(n):
            r = recs + offs[k]
            size = tl.u16(r)
            if k < n_funcs:
                fkccic = tl.i32(r + 16)
                nargs = tl.i16(r + 20)
                params = []
                for a in range(nargs):
                    p = r + size - nargs * 12 + 12 * a
                    params.append({"name": tl.name(tl.i32(p + 4)), "type": tl.type_name(tl.i32(p)),
                                   "flags": tl.i32(p + 8)})
                t["funcs"].append({
                    "name": tl.name(names[k]), "memid": None,
                    "invkind": (fkccic >> 3) & 0xF, "flags": tl.i32(r + 8) & 0xFFFF,
                    "ret": tl.type_name(tl.i32(r + 4)), "params": params,
                })
            else:
                t["vars"].append({"name": tl.name(names[k])})
    # Implemented interfaces of a coclass, or the base of an interface.
    ref = b[21]
    if kind == TKIND_COCLASS and ref != -1:
        o = ref
        while o != -1:
            ro = tl.seg["reftab"][0] + o
            t["impl"].append((tl.href_name(tl.i32(ro)), tl.i32(ro + 4)))
            o = tl.i32(ro + 12)
    elif kind in (TKIND_INTERFACE, TKIND_DISPATCH) and ref != -1:
        t["impl"].append((tl.href_name(ref), 0))
    return t


def read_memids(tl, idx):
    """Member ids of type `idx`, in function-then-variable order."""
    b = tl.info(idx)
    n = (b[6] & 0xFFFF) + ((b[6] >> 16) & 0xFFFF)
    if not n:
        return []
    after = b[1] + 4 + tl.i32(b[1])
    return [tl.i32(after + 4 * k) for k in range(n)]


def read_library(data, base):
    tl = TypeLib(data, base)
    types = []
    for i in range(tl.count):
        t = read_type(tl, i)
        ids = read_memids(tl, i)
        for f, mid in zip(t["funcs"], ids):
            f["memid"] = mid
        types.append(t)
    return tl.name(tl.name_offset), types


# ── engine schema ───────────────────────────────────────────────────────────

IGNORED_BASES = {"IUnknown", "IDispatch"}


def engine_type(t):
    kind = t["kind"]
    out = {"name": t["name"]}
    members = []
    if kind == TKIND_ENUM:
        out["kind"] = "enum"
        members = [{"name": v["name"], "kind": "constant"} for v in t["vars"]]
    elif kind == TKIND_MODULE:
        out["kind"] = "module"
        members = [{"name": v["name"], "kind": "constant"} for v in t["vars"]]
    elif kind in (TKIND_INTERFACE, TKIND_DISPATCH):
        out["kind"] = "interface"
        extends = [n for n, _ in t["impl"] if n not in IGNORED_BASES]
        if extends:
            out["extends"] = extends
        # An interface that is dispatchable, neither dual nor marked
        # non-extensible, gains members at run time: an unknown member is
        # looked up when the program runs, not rejected.
        flags = t["flags"]
        if flags & TYPEFLAG_DISPATCHABLE and not flags & (TYPEFLAG_NONEXTENSIBLE | TYPEFLAG_DUAL):
            out["extensible"] = True
        members = [{"name": v["name"], "kind": "field"} for v in t["vars"]]
    elif kind == TKIND_COCLASS:
        out["kind"] = "coclass"
        listed = [(n, fl) for n, fl in t["impl"] if not fl & IMPLTYPE_SOURCE]
        defaults = [n for n, fl in listed if fl & IMPLTYPE_DEFAULT]
        if defaults:
            out["default_interface"] = defaults[0]
        others = [n for n, fl in listed if not fl & IMPLTYPE_DEFAULT]
        if others:
            out["interfaces"] = others
    else:
        return None
    if t["flags"] & TYPEFLAG_APPOBJECT:
        out["is_global"] = True
    funcs = []
    for f in t["funcs"]:
        if f["flags"] & FUNCFLAG_RESTRICTED:
            continue
        params = [p for p in f["params"] if not p["flags"] & PARAMFLAG_LCID]
        retval = [p for p in params if p["flags"] & PARAMFLAG_RETVAL]
        args = [p for p in params if not p["flags"] & PARAMFLAG_RETVAL]
        ret = retval[0]["type"] if retval else f["ret"]
        if ret in ("void", "HRESULT"):
            ret = None
        kind_name = {INVOKE_PROPGET: "propertyget", INVOKE_PROPPUT: "propertylet",
                     INVOKE_PROPPUTREF: "propertyset"}.get(f["invkind"], "method")
        m = {"name": f["name"], "kind": kind_name}
        if ret and kind_name in ("method", "propertyget"):
            m["return_type"] = ret
        m["param_count"] = len(args)
        if f["memid"] == 0:
            m["is_default"] = True
        funcs.append(m)
    out["members"] = funcs + members
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("file")
    ap.add_argument("--library", action="append", default=[])
    ap.add_argument("--out")
    args = ap.parse_args()
    data = open(args.file, "rb").read()
    offsets = []
    at = data.find(MAGIC)
    while at != -1:
        offsets.append(at)
        at = data.find(MAGIC, at + 1)
    wanted = {n.lower() for n in args.library}
    libs = []
    for base in offsets:
        name, types = read_library(data, base)
        if wanted and name.lower() not in wanted:
            continue
        libs.append({"name": name, "types": [e for e in map(engine_type, types) if e]})
    text = json.dumps(libs, indent=1, ensure_ascii=True) + "\n"
    if args.out:
        open(args.out, "w").write(text)
    else:
        sys.stdout.write(text)


if __name__ == "__main__":
    main()
