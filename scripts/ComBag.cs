// COM bridge: decode a proprietary VB6 OCX property bag by hosting the
// control via COM.
//
// The .frm `Object=` GUID is the control's *type library* id, NOT the coclass
// CLSID that CoCreateInstance needs. So we resolve the coclass from the OCX type
// library (LoadTypeLibEx REGKIND_NONE — no registration required): find the
// coclass whose name matches the control class, take its CLSID, and remember its
// default interface so we can drive property reads from the real schema (the live
// RCW's IDispatch GetTypeInfo(0) walk comes back empty for some controls, e.g.
// MSChart's _DMSChart).
//
// Created LICENSE-AWARE (IClassFactory2::RequestLicKey -> CreateInstanceLic,
// falling back to CoCreateInstance). When the control exists but its license is
// absent, returns "NOTLICENSED" so the caller can HARD-ERROR. Must run in a
// 32-bit, STA host (VB6 OCXs are x86).
//
// Output (stdout, one line of JSON):
//   {"ok":true,"clsid":"{...}","properties":[["RowCount","5"],["ColumnCount","4"],...]}
//   {"ok":false,"error":"NOTLICENSED"}      control present, no license (hard error)
//   {"ok":false,"error":"NOTREG"}           control not registered
//   {"ok":false,"error":"<message>"}        other failure
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using CT = System.Runtime.InteropServices.ComTypes;

public static class ComBag
{
    [DllImport("ole32.dll")] static extern int CoCreateInstance(ref Guid clsid, IntPtr outer, uint ctx, ref Guid iid, [MarshalAs(UnmanagedType.IUnknown)] out object ppv);
    [DllImport("ole32.dll")] static extern int CoGetClassObject(ref Guid clsid, uint ctx, IntPtr res, ref Guid iid, [MarshalAs(UnmanagedType.Interface)] out object ppv);
    [DllImport("ole32.dll")] static extern int CreateStreamOnHGlobal(IntPtr hg, bool del, out IStream s);
    [DllImport("oleaut32.dll", CharSet = CharSet.Unicode)] static extern int LoadTypeLibEx(string file, int regkind, out CT.ITypeLib tlb);
    [DllImport("oleaut32.dll")] static extern int LoadRegTypeLib(ref Guid libid, ushort major, ushort minor, int lcid, out CT.ITypeLib tlb);

    const uint CLSCTX_INPROC_SERVER = 1;
    const int CLASS_E_NOTLICENSED = unchecked((int)0x80040112);
    const int REGDB_E_CLASSNOTREG = unchecked((int)0x80040154);

    static Guid IID_IUnknown = new Guid("00000000-0000-0000-C000-000000000046");
    static Guid IID_IDispatch = new Guid("00020400-0000-0000-C000-000000000046");
    static Guid IID_IClassFactory2 = new Guid("B196B28F-BAB4-101A-B69C-00AA00341D07");

    [ComImport, Guid("B196B28F-BAB4-101A-B69C-00AA00341D07"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IClassFactory2
    {
        void CreateInstance(IntPtr outer, ref Guid iid, [MarshalAs(UnmanagedType.IUnknown)] out object obj);
        void LockServer(bool fLock);
        void GetLicInfo(IntPtr licInfo);
        [PreserveSig] int RequestLicKey(uint reserved, [MarshalAs(UnmanagedType.BStr)] out string key);
        [PreserveSig] int CreateInstanceLic(IntPtr outer, IntPtr reserved, ref Guid iid, [MarshalAs(UnmanagedType.BStr)] string key, [MarshalAs(UnmanagedType.IUnknown)] out object obj);
    }

    [ComImport, Guid("7FD52380-4E07-101B-AE2D-08002B2EC713"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IPersistStreamInit
    {
        void GetClassID(out Guid id); [PreserveSig] int IsDirty();
        void Load(IStream s); void Save(IStream s, [MarshalAs(UnmanagedType.Bool)] bool clear);
        void GetSizeMax(out long size); void InitNew();
    }
    [ComImport, Guid("00000109-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IPersistStream
    {
        void GetClassID(out Guid id); [PreserveSig] int IsDirty();
        void Load(IStream s); void Save(IStream s, [MarshalAs(UnmanagedType.Bool)] bool clear); void GetSizeMax(out long size);
    }

    static IStream StreamFromBytes(byte[] d)
    {
        IStream s; Marshal.ThrowExceptionForHR(CreateStreamOnHGlobal(IntPtr.Zero, true, out s));
        s.Write(d, d.Length, IntPtr.Zero);
        IntPtr po = Marshal.AllocHGlobal(8); s.Seek(0, 0, po); Marshal.FreeHGlobal(po); return s;
    }

    // Current seek position of a stream = bytes the control consumed during Load.
    static long StreamPos(IStream s)
    {
        IntPtr p = Marshal.AllocHGlobal(8);
        try { s.Seek(0, 1 /*STREAM_SEEK_CUR*/, p); return Marshal.ReadInt64(p); }
        catch { return 0; }
        finally { Marshal.FreeHGlobal(p); }
    }

    static bool TryGuid(string s, out Guid g)
    {
        g = Guid.Empty;
        if (string.IsNullOrEmpty(s)) return false;
        try { g = new Guid(s); return true; } catch { return false; }
    }

    // Load a candidate's type library: from the OCX file when present (reg-free),
    // else from the registry by type-library GUID + version. Returns null on failure.
    static CT.ITypeLib GetTypeLib(string ocxPath, string libid, string version)
    {
        CT.ITypeLib tlb;
        if (!string.IsNullOrEmpty(ocxPath) && LoadTypeLibEx(ocxPath, 2 /*REGKIND_NONE*/, out tlb) == 0)
            return tlb;
        Guid g;
        if (TryGuid(libid, out g))
        {
            ushort maj = 1, min = 0;
            var parts = (version ?? "").Split('.');
            if (parts.Length > 0) ushort.TryParse(parts[0], out maj);
            if (parts.Length > 1) ushort.TryParse(parts[1], out min);
            if (LoadRegTypeLib(ref g, maj, min, 0, out tlb) == 0) return tlb;
        }
        return null;
    }

    // Resolve the coclass CLSID (and its default interface, for the property schema)
    // from a type library by matching the control class name. When `libName` is
    // given, the typelib's own library name must match it — this disambiguates a
    // coclass-name collision across the candidate typelibs.
    static bool ResolveCoclass(CT.ITypeLib tlb, string className, string libName, out Guid clsid, out CT.ITypeInfo defIface)
    {
        clsid = Guid.Empty; defIface = null;
        if (tlb == null || string.IsNullOrEmpty(className)) return false;
        if (!string.IsNullOrEmpty(libName))
        {
            string ln, ld; int lh; string lf; tlb.GetDocumentation(-1, out ln, out ld, out lh, out lf);
            if (ln != null && !string.Equals(ln, libName, StringComparison.OrdinalIgnoreCase)) return false;
        }
        int n = tlb.GetTypeInfoCount();
        for (int i = 0; i < n; i++)
        {
            CT.TYPEKIND kind; tlb.GetTypeInfoType(i, out kind);
            if (kind != CT.TYPEKIND.TKIND_COCLASS) continue;
            CT.ITypeInfo ti; tlb.GetTypeInfo(i, out ti);
            string name, doc; int hc; string hf; ti.GetDocumentation(-1, out name, out doc, out hc, out hf);
            if (!string.Equals(name, className, StringComparison.OrdinalIgnoreCase)) continue;
            IntPtr pAttr; ti.GetTypeAttr(out pAttr);
            var attr = (CT.TYPEATTR)Marshal.PtrToStructure(pAttr, typeof(CT.TYPEATTR));
            clsid = attr.guid; short cImpl = attr.cImplTypes; ti.ReleaseTypeAttr(pAttr);
            for (int t = 0; t < cImpl; t++)
            {
                CT.IMPLTYPEFLAGS flags; ti.GetImplTypeFlags(t, out flags);
                if ((flags & CT.IMPLTYPEFLAGS.IMPLTYPEFLAG_FSOURCE) != 0) continue; // skip event sources
                int href; ti.GetRefTypeOfImplType(t, out href);
                CT.ITypeInfo iface; ti.GetRefTypeInfo(href, out iface);
                if (defIface == null) defIface = iface;
                if ((flags & CT.IMPLTYPEFLAGS.IMPLTYPEFLAG_FDEFAULT) != 0) { defIface = iface; break; }
            }
            return clsid != Guid.Empty;
        }
        return false;
    }

    // The no-parameter, browsable property getters that make up the bag schema.
    // Skips hidden/restricted/non-browsable members (runtime-only noise such as a
    // window handle), and for vtable interfaces walks the inheritance chain so
    // properties declared on a base interface aren't missed.
    static List<string> SchemaGetters(CT.ITypeInfo ti)
    {
        var list = new List<string>();
        CollectGetters(ti, new HashSet<string>(), new HashSet<Guid>(), list);
        return list;
    }

    const CT.FUNCFLAGS NONBROWSABLE =
        CT.FUNCFLAGS.FUNCFLAG_FRESTRICTED | CT.FUNCFLAGS.FUNCFLAG_FHIDDEN | CT.FUNCFLAGS.FUNCFLAG_FNONBROWSABLE;

    static void CollectGetters(CT.ITypeInfo ti, HashSet<string> seen, HashSet<Guid> visited, List<string> list)
    {
        if (ti == null) return;
        IntPtr pAttr; ti.GetTypeAttr(out pAttr);
        var attr = (CT.TYPEATTR)Marshal.PtrToStructure(pAttr, typeof(CT.TYPEATTR));
        Guid g = attr.guid; short cFuncs = attr.cFuncs; short cImpl = attr.cImplTypes;
        CT.TYPEKIND kind = attr.typekind;
        ti.ReleaseTypeAttr(pAttr);
        // Stop at the universal bases and guard against inheritance cycles.
        if (g == IID_IUnknown || g == IID_IDispatch || !visited.Add(g)) return;

        for (int f = 0; f < cFuncs; f++)
        {
            IntPtr pf; ti.GetFuncDesc(f, out pf);
            var fd = (CT.FUNCDESC)Marshal.PtrToStructure(pf, typeof(CT.FUNCDESC));
            bool getter = fd.invkind == CT.INVOKEKIND.INVOKE_PROPERTYGET;
            short cParams = fd.cParams; int memid = fd.memid;
            bool browsable = ((CT.FUNCFLAGS)fd.wFuncFlags & NONBROWSABLE) == 0;
            ti.ReleaseFuncDesc(pf);
            if (getter && cParams == 0 && browsable)
            {
                string[] names = new string[1]; int cn;
                ti.GetNames(memid, names, 1, out cn);
                if (cn > 0 && names[0] != null && seen.Add(names[0])) list.Add(names[0]);
            }
        }
        // Dispinterfaces list every member flat; vtable interfaces inherit, so walk
        // the base interface (impl type 0) too.
        if (kind == CT.TYPEKIND.TKIND_INTERFACE && cImpl > 0)
        {
            int href; ti.GetRefTypeOfImplType(0, out href);
            CT.ITypeInfo baseTi; ti.GetRefTypeInfo(href, out baseTi);
            CollectGetters(baseTi, seen, visited, list);
        }
    }

    // License-aware creation. Throws "NOTLICENSED"/"NOTREG" on those specific failures.
    static object Create(Guid clsid)
    {
        Guid iid = IID_IUnknown; object obj;
        int hr = CoCreateInstance(ref clsid, IntPtr.Zero, CLSCTX_INPROC_SERVER, ref iid, out obj);
        if (hr == 0) return obj;
        if (hr == REGDB_E_CLASSNOTREG) throw new Exception("NOTREG");
        if (hr == CLASS_E_NOTLICENSED)
        {
            Guid icf2 = IID_IClassFactory2; object factObj;
            int fhr = CoGetClassObject(ref clsid, CLSCTX_INPROC_SERVER, IntPtr.Zero, ref icf2, out factObj);
            if (fhr != 0) throw new Exception("NOTLICENSED");
            var fac = (IClassFactory2)factObj;
            string key;
            if (fac.RequestLicKey(0, out key) != 0 || key == null) throw new Exception("NOTLICENSED");
            Guid iid2 = IID_IUnknown; object lic;
            if (fac.CreateInstanceLic(IntPtr.Zero, IntPtr.Zero, ref iid2, key, out lic) != 0) throw new Exception("NOTLICENSED");
            return lic;
        }
        throw new Exception("create hr=0x" + hr.ToString("X8"));
    }

    static void LoadBag(object obj, byte[] bag)
    {
        // A bag may begin with a class id / VB framing (e.g. OleObjectBlob's "LB"
        // header) before the persisted stream; try a few leading-offset skips with
        // both persistence interfaces. A Load that returns S_OK but consumes zero
        // bytes hasn't really loaded (a lenient control accepting a wrong offset),
        // so we require the control to have advanced the stream before accepting.
        int[] skips = { 0, 16, 20, 24 };
        Exception last = null;
        foreach (int skip in skips)
        {
            if (skip >= bag.Length) continue;
            byte[] d = bag;
            if (skip > 0) { d = new byte[bag.Length - skip]; Array.Copy(bag, skip, d, 0, d.Length); }
            IStream s1 = StreamFromBytes(d);
            try { ((IPersistStreamInit)obj).Load(s1); if (StreamPos(s1) > 0) return; } catch (Exception e) { last = e; }
            IStream s2 = StreamFromBytes(d);
            try { ((IPersistStream)obj).Load(s2); if (StreamPos(s2) > 0) return; } catch (Exception e) { last = e; }
        }
        throw last ?? new Exception("load failed or consumed no bytes");
    }

    // Read a known list of property names off the live control (schema-driven).
    static List<string[]> ReadNamed(object obj, List<string> names)
    {
        var result = new List<string[]>();
        foreach (string n in names)
        {
            if (result.Count > 2000) break;
            object v;
            try { v = obj.GetType().InvokeMember(n, BindingFlags.GetProperty, null, obj, null); }
            catch (Exception e) { result.Add(new string[] { n, "<err:" + (e.InnerException != null ? e.InnerException.Message : e.Message) + ">" }); continue; }
            if (v == null) { result.Add(new string[] { n, "" }); continue; }
            if (Marshal.IsComObject(v)) ReadCollection(v, n, 1, result);
            else result.Add(new string[] { n, v.ToString() });
        }
        return result;
    }

    // Fallback: read gettable scalar properties off the live RCW's own type info,
    // recursing one or two levels into object-valued (collection) properties.
    static List<string[]> ReadProps(object obj)
    {
        var result = new List<string[]>();
        ReadPropsInto(obj, "", 2, result);
        return result;
    }

    static void ReadPropsInto(object obj, string prefix, int depth, List<string[]> result)
    {
        if (obj == null || result.Count > 2000) return;
        CT.ITypeInfo ti = null;
        try
        {
            var disp = (IDispatchInfo)obj;
            int n; disp.GetTypeInfoCount(out n);
            if (n > 0) disp.GetTypeInfo(0, 0, out ti);
        }
        catch { }
        if (ti == null) return;

        IntPtr pAttr; ti.GetTypeAttr(out pAttr);
        var attr = (CT.TYPEATTR)Marshal.PtrToStructure(pAttr, typeof(CT.TYPEATTR));
        short cFuncs = attr.cFuncs; short cVars = attr.cVars; ti.ReleaseTypeAttr(pAttr);

        var seen = new HashSet<string>();
        for (int f = 0; f < cFuncs; f++)
        {
            IntPtr pf; ti.GetFuncDesc(f, out pf);
            var fd = (CT.FUNCDESC)Marshal.PtrToStructure(pf, typeof(CT.FUNCDESC));
            bool getter = (fd.invkind == CT.INVOKEKIND.INVOKE_PROPERTYGET);
            short cParams = fd.cParams;
            int memid = fd.memid;
            ti.ReleaseFuncDesc(pf);
            if (getter && cParams == 0) ReadMember(obj, ti, memid, prefix, depth, seen, result);
        }
        for (int vi = 0; vi < cVars; vi++)
        {
            IntPtr pv; ti.GetVarDesc(vi, out pv);
            var vd = (CT.VARDESC)Marshal.PtrToStructure(pv, typeof(CT.VARDESC));
            int memid = vd.memid;
            ti.ReleaseVarDesc(pv);
            ReadMember(obj, ti, memid, prefix, depth, seen, result);
        }
    }

    static void ReadMember(object obj, CT.ITypeInfo ti, int memid, string prefix, int depth, HashSet<string> seen, List<string[]> result)
    {
        string[] names = new string[1]; int cn;
        ti.GetNames(memid, names, 1, out cn);
        if (cn == 0 || names[0] == null || !seen.Add(names[0])) return;
        string full = prefix.Length == 0 ? names[0] : prefix + "." + names[0];
        object v;
        try { v = obj.GetType().InvokeMember(names[0], BindingFlags.GetProperty, null, obj, null); }
        catch (Exception e) { result.Add(new string[] { full, "<err:" + (e.InnerException != null ? e.InnerException.Message : e.Message) + ">" }); return; }
        if (v == null) { result.Add(new string[] { full, "" }); return; }
        if (depth > 0 && Marshal.IsComObject(v)) ReadCollection(v, full, depth - 1, result);
        else result.Add(new string[] { full, v.ToString() });
    }

    static void ReadCollection(object col, string prefix, int depth, List<string[]> result)
    {
        int count = -1;
        try { object c = col.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, col, null); count = Convert.ToInt32(c); }
        catch { }
        if (count < 0) { ReadPropsInto(col, prefix, depth, result); return; }

        result.Add(new string[] { prefix + ".Count", count.ToString() });
        int cap = Math.Min(count, 100);
        for (int i = 0; i < cap; i++)
        {
            object item = null;
            try { item = col.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, col, new object[] { i + 1 }); }
            catch { try { item = col.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, col, new object[] { i }); } catch { } }
            if (item == null) continue;
            string ip = prefix + "(" + i + ")";
            if (Marshal.IsComObject(item)) ReadPropsInto(item, ip, depth, result);
            else result.Add(new string[] { ip, item.ToString() });
        }
    }

    [ComImport, Guid("00020400-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IDispatchInfo
    {
        void GetTypeInfoCount(out int count);
        void GetTypeInfo(int index, int lcid, out CT.ITypeInfo ti);
    }

    static string JsonStr(string s)
    {
        var sb = new StringBuilder("\"");
        foreach (char c in s)
        {
            if (c == '"' || c == '\\') sb.Append('\\').Append(c);
            else if (c == '\n') sb.Append("\\n");
            else if (c == '\r') sb.Append("\\r");
            else if (c == '\t') sb.Append("\\t");
            else if (c < 0x20) sb.Append("\\u").Append(((int)c).ToString("x4"));
            else sb.Append(c);
        }
        return sb.Append('"').ToString();
    }

    // `;`-separated list -> entries. Empty slots are KEPT (the OCX-path / GUID /
    // version lists are index-aligned, so dropping empties would misalign them).
    static string[] SplitList(string s)
    {
        if (string.IsNullOrEmpty(s)) return new string[0];
        return s.Split(';');
    }

    public static string Decode(string ocxPaths, string className, string libName, string typelibClsids, string versions, string embeddedClsid, byte[] bag)
    {
        // Declared out here so the catch can report which coclass we identified
        // (e.g. on NOTLICENSED the control IS known — resolution preceded Create).
        Guid clsid = Guid.Empty;
        try
        {
            // The lists are index-aligned: entry i = one Object= line. Try each
            // candidate typelib (OCX path, else registry by GUID+version) and use
            // the one whose library/coclass name matches.
            string[] paths = SplitList(ocxPaths);
            string[] libids = SplitList(typelibClsids);
            string[] vers = SplitList(versions);
            int count = Math.Max(paths.Length, Math.Max(libids.Length, vers.Length));
            CT.ITypeInfo defIface = null; bool resolved = false;
            for (int i = 0; i < count; i++)
            {
                string p = i < paths.Length ? paths[i] : "";
                string lid = i < libids.Length ? libids[i] : "";
                string ver = i < vers.Length ? vers[i] : "";
                CT.ITypeLib tlb = GetTypeLib(p, lid, ver);
                if (tlb == null) continue;
                if (ResolveCoclass(tlb, className, libName, out clsid, out defIface)) { resolved = true; break; }
            }
            if (!resolved)
            {
                defIface = null;
                bool fb = TryGuid(embeddedClsid, out clsid);
                if (!fb) foreach (string g in libids) { if (TryGuid(g, out clsid)) { fb = true; break; } }
                if (!fb)
                    return "{\"ok\":false,\"error\":" + JsonStr("could not resolve coclass for class '" + className + "'") + "}";
            }

            object obj = Create(clsid);
            LoadBag(obj, bag);

            var names = SchemaGetters(defIface);
            List<string[]> props = names.Count > 0 ? ReadNamed(obj, names) : ReadProps(obj);

            var sb = new StringBuilder();
            sb.Append("{\"ok\":true,\"clsid\":").Append(JsonStr(clsid.ToString("B"))).Append(",\"properties\":[");
            for (int i = 0; i < props.Count; i++)
            {
                if (i > 0) sb.Append(',');
                sb.Append('[').Append(JsonStr(props[i][0])).Append(',').Append(JsonStr(props[i][1])).Append(']');
            }
            sb.Append("]}");
            return sb.ToString();
        }
        catch (Exception e)
        {
            string cl = clsid != Guid.Empty ? ",\"clsid\":" + JsonStr(clsid.ToString("B")) : "";
            return "{\"ok\":false" + cl + ",\"error\":" + JsonStr(e.Message) + "}";
        }
    }
}
