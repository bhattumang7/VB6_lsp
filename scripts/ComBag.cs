// Tier-3 COM bridge: decode a proprietary VB6 OCX property bag by hosting the
// control via COM. Created LICENSE-AWARE (IClassFactory2::RequestLicKey ->
// CreateInstanceLic, falling back to CoCreateInstance). When the control exists
// but its license is absent, returns the marker "NOTLICENSED" so the caller can
// HARD-ERROR rather than degrade. Must run in a 32-bit, STA host (VB6 OCXs are x86).
//
// Output (stdout, one line of JSON):
//   {"ok":true,"clsid":"{...}","properties":[["Rows","5"],["Cols","3"],...]}
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

    const uint CLSCTX_INPROC_SERVER = 1;
    const int CLASS_E_NOTLICENSED = unchecked((int)0x80040112);
    const int REGDB_E_CLASSNOTREG = unchecked((int)0x80040154);

    static Guid IID_IUnknown = new Guid("00000000-0000-0000-C000-000000000046");
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
        // A bag may begin with a class id / VB framing before the persisted stream;
        // try a few leading-offset skips with both persistence interfaces.
        int[] skips = { 0, 16, 20, 24 };
        Exception last = null;
        foreach (int skip in skips)
        {
            if (skip >= bag.Length) continue;
            byte[] d = bag;
            if (skip > 0) { d = new byte[bag.Length - skip]; Array.Copy(bag, skip, d, 0, d.Length); }
            try { ((IPersistStreamInit)obj).Load(StreamFromBytes(d)); return; } catch (Exception e) { last = e; }
            try { ((IPersistStream)obj).Load(StreamFromBytes(d)); return; } catch (Exception e) { last = e; }
        }
        throw last ?? new Exception("load failed");
    }

    // Read gettable scalar properties, recursing one or two levels into object-valued
    // (collection) properties such as a grid's Columns or an ActiveBar's Bands.
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
        // PROPERTYGET methods with no parameters (dual / vtable interfaces).
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
        // VARDESC properties (dispinterfaces, e.g. StdFont's IFontDisp, expose
        // their properties as variables rather than get-methods).
        for (int vi = 0; vi < cVars; vi++)
        {
            IntPtr pv; ti.GetVarDesc(vi, out pv);
            var vd = (CT.VARDESC)Marshal.PtrToStructure(pv, typeof(CT.VARDESC));
            int memid = vd.memid;
            ti.ReleaseVarDesc(pv);
            ReadMember(obj, ti, memid, prefix, depth, seen, result);
        }
    }

    // Resolve a member's name, fetch its value, and either record it or (for a
    // collection/object) recurse.
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

    // A property that returned a COM object: if it exposes Count + Item, enumerate its
    // items (capped); otherwise recurse its scalar properties.
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

    // Minimal IDispatch surface needed to fetch ITypeInfo.
    [ComImport, Guid("00020400-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IDispatchInfo
    {
        void GetTypeInfoCount(out int count);
        void GetTypeInfo(int index, int lcid, out CT.ITypeInfo ti);
        // (GetIDsOfNames / Invoke omitted — not needed; we use reflection InvokeMember)
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

    public static string Decode(string clsidStr, byte[] bag)
    {
        try
        {
            Guid clsid = new Guid(clsidStr);
            object obj = Create(clsid);
            LoadBag(obj, bag);
            var props = ReadProps(obj);
            var sb = new StringBuilder();
            sb.Append("{\"ok\":true,\"clsid\":").Append(JsonStr(clsidStr)).Append(",\"properties\":[");
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
            string msg = e.Message == "NOTLICENSED" || e.Message == "NOTREG" ? e.Message : e.Message;
            return "{\"ok\":false,\"error\":" + JsonStr(msg) + "}";
        }
    }
}
