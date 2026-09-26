// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Threading;

namespace OfficeCli.Core;

[SupportedOSPlatform("windows")]
internal static class WordPdfBackend
{
    [DllImport("combase.dll", PreserveSig = false)]
    static extern void WindowsCreateString([MarshalAs(UnmanagedType.LPWStr)] string s, uint len, out IntPtr h);

    [DllImport("combase.dll", PreserveSig = false)]
    static extern void WindowsDeleteString(IntPtr h);

    [DllImport("combase.dll", PreserveSig = false)]
    static extern void RoGetActivationFactory(IntPtr classId, ref Guid iid, out IntPtr factory);

    [DllImport("combase.dll", PreserveSig = false)]
    static extern void RoActivateInstance(IntPtr classId, out IntPtr instance);

    [DllImport("ole32.dll", PreserveSig = false)]
    internal static extern void CoCreateInstance(ref Guid clsid, IntPtr unkOuter, int ctx, ref Guid iid, out IntPtr ppv);

    [DllImport("ole32.dll", PreserveSig = false)]
    static extern void CreateStreamOnHGlobal(IntPtr hGlobal, [MarshalAs(UnmanagedType.Bool)] bool fDeleteOnRelease, out IntPtr ppstm);

    [DllImport("ole32.dll", PreserveSig = false)]
    static extern void GetHGlobalFromStream(IntPtr pstm, out IntPtr phglobal);

    [DllImport("oleaut32.dll", CharSet = CharSet.Unicode)]
    static extern IntPtr SysAllocString(string s);

    [DllImport("oleaut32.dll")]
    static extern void SysFreeString(IntPtr bstr);

    [DllImport("kernel32.dll")] static extern IntPtr GlobalLock(IntPtr h);
    [DllImport("kernel32.dll")] static extern bool GlobalUnlock(IntPtr h);
    [DllImport("kernel32.dll")] static extern uint GlobalSize(IntPtr h);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int F_QI(IntPtr self, ref Guid iid, out IntPtr p);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int F_OneIn(IntPtr self, IntPtr a, out IntPtr r);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int F_Out(IntPtr self, out IntPtr r);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int F_OutVoid(IntPtr self);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int F_OutU32(IntPtr self, out uint r);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int F_OutU64(IntPtr self, out ulong r);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int F_GetPage(IntPtr self, uint i, out IntPtr p);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int F_GetInAt(IntPtr self, ulong pos, out IntPtr s);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int F_LoadAsync(IntPtr self, uint count, out IntPtr op);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int F_ReadBytes(IntPtr self, uint len, IntPtr buf);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int F_GetIDsOfNames(IntPtr self, ref Guid riid, IntPtr rgszNames, uint cNames, uint lcid, IntPtr rgDispId);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int F_Invoke(IntPtr self, int dispIdMember, ref Guid riid, uint lcid, ushort wFlags, IntPtr pDispParams, IntPtr pVarResult, IntPtr pExcepInfo, IntPtr puArgErr);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int F_CopyPixels(IntPtr self, IntPtr rect, uint stride, uint bufSize, IntPtr buf);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int F_InitMem(IntPtr self, IntPtr buf, uint cb);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int F_CreateDecoder(IntPtr self, IntPtr stream, IntPtr vendor, int cache, out IntPtr decoder);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int F_GetFrame(IntPtr self, uint i, out IntPtr frame);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int F_GetSize(IntPtr self, out uint w, out uint h);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int F_CreateConverter(IntPtr self, out IntPtr converter);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int F_ConvInit(IntPtr self, IntPtr src, ref Guid dst, int dither, IntPtr palette, double alpha, int paletteType);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int F_CreateEncoder(IntPtr self, ref Guid containerFmt, IntPtr vendor, out IntPtr encoder);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int F_EncInit(IntPtr self, IntPtr stream, int cache);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int F_CreateNewFrame(IntPtr self, out IntPtr frame, out IntPtr propBag);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int F_FrameInit(IntPtr self, IntPtr propBag);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int F_SetSize(IntPtr self, uint w, uint h);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int F_SetPixelFormat(IntPtr self, ref Guid pf);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int F_FrameWritePixels(IntPtr self, uint lineCount, uint stride, uint bufSize, IntPtr pixels);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int F_Commit(IntPtr self);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int F_CreateStream(IntPtr self, out IntPtr stream);

    static readonly Guid G_AsyncInfo      = new("00000036-0000-0000-c000-000000000046");
    static readonly Guid G_PdfDocStatics  = new("433a0b5f-c007-4788-90f2-08143d922599");
    static readonly Guid G_FileStatics    = new("5984c710-daf2-43c8-8bb4-a4d3eacfd03f");
    static readonly Guid G_DataReaderFact = new("d7527847-57da-4e15-914c-06806699a098");
    static readonly Guid G_RAS            = new("905a0fe1-bc53-11df-8c49-001e4fc686da");
    static readonly Guid G_Word           = new("000209FF-0000-0000-C000-000000000046");
    internal static readonly Guid G_IDispatch      = new("00020400-0000-0000-C000-000000000046");
    static readonly Guid G_WICFactory_C   = new("CACAF262-9370-4615-A13B-9F5539DA4C0A");
    static readonly Guid G_WICFactory_I   = new("EC5EC8A9-C395-4314-9C77-54D7A935FF70");
    static readonly Guid G_PngContainer   = new("1B7CFAF4-713F-473C-BBCD-6137425FAEAF");
    static readonly Guid G_BGRA32         = new("6FDDC324-4E03-4BFE-B185-3D77768DC90F");

    const int VAR_SZ = 24;
    const ushort VT_EMPTY = 0;
    const ushort VT_I4 = 3;
    const ushort VT_BSTR = 8;
    const ushort VT_DISPATCH = 9;
    const ushort VT_BOOL = 11;
    const ushort VT_ERROR = 10;
    const int DISPID_PROPERTYPUT = -3;
    const uint DISP_E_PARAMNOTFOUND = 0x80020004;
    static readonly object MISSING = new();

    static T VT<T>(IntPtr p, int slot) where T : Delegate
        => (T)Marshal.GetDelegateForFunctionPointer(Marshal.ReadIntPtr(Marshal.ReadIntPtr(p), slot * IntPtr.Size), typeof(T));

    static IntPtr Hs(string s) { WindowsCreateString(s, (uint)s.Length, out var h); return h; }

    static IntPtr Factory(string cls, Guid iid)
    {
        var n = Hs(cls);
        try { RoGetActivationFactory(n, ref iid, out var f); return f; }
        finally { WindowsDeleteString(n); }
    }

    static IntPtr Activate(string cls)
    {
        var n = Hs(cls);
        try { RoActivateInstance(n, out var p); return p; }
        finally { WindowsDeleteString(n); }
    }

    static IntPtr QI(IntPtr p, Guid iid)
    {
        var fn = VT<F_QI>(p, 0);
        var iidCopy = iid;
        if (fn(p, ref iidCopy, out var r) != 0) throw new InvalidOperationException();
        return r;
    }

    static int Rel(IntPtr p) => p == IntPtr.Zero ? 0 : Marshal.Release(p);

    static void Wait(IntPtr op, int timeoutMs)
    {
        var info = QI(op, G_AsyncInfo);
        try
        {
            var getStatus = VT<F_OutU32>(info, 7);
            var deadline = Environment.TickCount64 + timeoutMs;
            while (Environment.TickCount64 < deadline)
            {
                getStatus(info, out var s);
                if (s == 0) { Thread.Sleep(20); continue; }
                if (s == 1) return;
                throw new InvalidOperationException();
            }
            throw new TimeoutException();
        }
        finally { Rel(info); }
    }

    static int DispId(IntPtr d, string name)
    {
        var np = Marshal.StringToHGlobalUni(name);
        var arr = Marshal.AllocHGlobal(IntPtr.Size);
        var did = Marshal.AllocHGlobal(4);
        try
        {
            Marshal.WriteIntPtr(arr, np);
            var iid = Guid.Empty;
            int hr = VT<F_GetIDsOfNames>(d, 5)(d, ref iid, arr, 1, 0, did);
            if (hr != 0) throw new InvalidOperationException($"DispId({name}) hr=0x{hr:X8}");
            return Marshal.ReadInt32(did);
        }
        finally { Marshal.FreeHGlobal(did); Marshal.FreeHGlobal(arr); Marshal.FreeHGlobal(np); }
    }

    static void Wv(IntPtr v, object? o)
    {
        Marshal.WriteInt64(v, 0); Marshal.WriteInt64(v, 8, 0); Marshal.WriteInt64(v, 16, 0);
        switch (o)
        {
            case null: Marshal.WriteInt16(v, (short)VT_EMPTY); break;
            case int i: Marshal.WriteInt16(v, (short)VT_I4); Marshal.WriteInt32(v, 8, i); break;
            case bool b: Marshal.WriteInt16(v, (short)VT_BOOL); Marshal.WriteInt16(v, 8, (short)(b ? -1 : 0)); break;
            case string s: Marshal.WriteInt16(v, (short)VT_BSTR); Marshal.WriteIntPtr(v, 8, SysAllocString(s)); break;
            case IntPtr p:
                Marshal.WriteInt16(v, (short)VT_DISPATCH); Marshal.WriteIntPtr(v, 8, p);
                if (p != IntPtr.Zero) Marshal.AddRef(p);
                break;
            default:
                if (ReferenceEquals(o, MISSING)) { Marshal.WriteInt16(v, (short)VT_ERROR); Marshal.WriteInt32(v, 8, unchecked((int)DISP_E_PARAMNOTFOUND)); }
                else throw new ArgumentException($"unsupported variant type: {o.GetType()}");
                break;
        }
    }

    static object? Rv(IntPtr v, bool addRefDispatch = true)
    {
        var vt = Marshal.ReadInt16(v);
        switch (vt)
        {
            case (short)VT_EMPTY: return null;
            case (short)VT_I4: return Marshal.ReadInt32(v, 8);
            case (short)VT_BOOL: return Marshal.ReadInt16(v, 8) != 0;
            case (short)VT_BSTR: return Marshal.PtrToStringBSTR(Marshal.ReadIntPtr(v, 8));
            case (short)VT_DISPATCH:
                var p = Marshal.ReadIntPtr(v, 8);
                if (addRefDispatch && p != IntPtr.Zero) Marshal.AddRef(p);
                return p;
            default: return null;
        }
    }

    static void Cv(IntPtr v)
    {
        var vt = Marshal.ReadInt16(v);
        if (vt == (short)VT_BSTR) { var b = Marshal.ReadIntPtr(v, 8); if (b != IntPtr.Zero) SysFreeString(b); }
        else if (vt == (short)VT_DISPATCH) { var p = Marshal.ReadIntPtr(v, 8); if (p != IntPtr.Zero) Marshal.Release(p); }
        Marshal.WriteInt64(v, 0); Marshal.WriteInt64(v, 8, 0); Marshal.WriteInt64(v, 16, 0);
    }

    static object? DispCall(IntPtr d, string name, ushort flags, object?[] args, bool isPut = false)
    {
        int dispId = DispId(d, name);
        IntPtr argArr = args.Length > 0 ? Marshal.AllocHGlobal(VAR_SZ * args.Length) : IntPtr.Zero;
        IntPtr namedArr = isPut ? Marshal.AllocHGlobal(4) : IntPtr.Zero;
        IntPtr dp = Marshal.AllocHGlobal(IntPtr.Size * 2 + 8);
        IntPtr result = Marshal.AllocHGlobal(VAR_SZ);
        Marshal.WriteInt64(result, 0); Marshal.WriteInt64(result, 8, 0); Marshal.WriteInt64(result, 16, 0);
        try
        {
            for (int i = 0; i < args.Length; i++) Wv(argArr + (args.Length - 1 - i) * VAR_SZ, args[i]);
            if (isPut) Marshal.WriteInt32(namedArr, DISPID_PROPERTYPUT);
            Marshal.WriteIntPtr(dp, argArr);
            Marshal.WriteIntPtr(dp, IntPtr.Size, namedArr);
            Marshal.WriteInt32(dp, IntPtr.Size * 2, args.Length);
            Marshal.WriteInt32(dp, IntPtr.Size * 2 + 4, isPut ? 1 : 0);

            var iid = Guid.Empty;
            int hr = VT<F_Invoke>(d, 6)(d, dispId, ref iid, 0, flags, dp, result, IntPtr.Zero, IntPtr.Zero);
            if (hr != 0) throw new InvalidOperationException($"Invoke({name}) hr=0x{hr:X8}");
            return Rv(result);
        }
        finally
        {
            Cv(result); Marshal.FreeHGlobal(result);
            Marshal.FreeHGlobal(dp);
            for (int i = 0; i < args.Length; i++) Cv(argArr + i * VAR_SZ);
            if (argArr != IntPtr.Zero) Marshal.FreeHGlobal(argArr);
            if (namedArr != IntPtr.Zero) Marshal.FreeHGlobal(namedArr);
        }
    }

    internal static void DispSet(IntPtr d, string name, object? v) => DispCall(d, name, 4, [v], true);
    internal static object? DispGet(IntPtr d, string name) => DispCall(d, name, 2, []);
    internal static object? DispMethod(IntPtr d, string name, params object?[] args) => DispCall(d, name, 1, args);

    static byte[] RenderOne(IntPtr doc, uint i, IntPtr drFactory, int timeoutMs)
    {
        var getPage = VT<F_GetPage>(doc, 6);
        if (getPage(doc, i, out var page) != 0) throw new InvalidOperationException();
        try
        {
            var streamObj = Activate("Windows.Storage.Streams.InMemoryRandomAccessStream");
            var stream = QI(streamObj, G_RAS);
            Rel(streamObj);
            try
            {
                var render = VT<F_OneIn>(page, 6);
                if (render(page, stream, out var op) != 0) throw new InvalidOperationException();
                Wait(op, timeoutMs);
                VT<F_OutVoid>(op, 8)(op);
                Rel(op);

                VT<F_OutU64>(stream, 6)(stream, out var size);
                VT<F_GetInAt>(stream, 8)(stream, 0, out var inStream);
                try
                {
                    VT<F_OneIn>(drFactory, 6)(drFactory, inStream, out var reader);
                    try
                    {
                        VT<F_LoadAsync>(reader, 29)(reader, (uint)size, out var lop);
                        Wait(lop, timeoutMs);
                        VT<F_OutU32>(lop, 8)(lop, out _);
                        Rel(lop);

                        var bytes = new byte[size];
                        var buf = Marshal.AllocHGlobal((int)size);
                        try
                        {
                            VT<F_ReadBytes>(reader, 14)(reader, (uint)size, buf);
                            Marshal.Copy(buf, bytes, 0, (int)size);
                        }
                        finally { Marshal.FreeHGlobal(buf); }
                        return bytes;
                    }
                    finally { Rel(reader); }
                }
                finally { Rel(inStream); }
            }
            finally { Rel(stream); }
        }
        finally { Rel(page); }
    }

    static int[] ParsePages(string filter, int total)
    {
        var set = new SortedSet<int>();
        if (string.IsNullOrWhiteSpace(filter)) return [1];
        foreach (var part in filter.Split(','))
        {
            var t = part.Trim();
            if (t.Contains('-'))
            {
                var r = t.Split('-', 2);
                if (int.TryParse(r[0].Trim(), out var from) && int.TryParse(r[1].Trim(), out var to))
                    for (int p = from; p <= to; p++) if (p >= 1 && p <= total) set.Add(p);
            }
            else if (int.TryParse(t, out var n) && n >= 1 && n <= total) set.Add(n);
        }
        if (set.Count == 0) set.Add(1);
        return set.ToArray();
    }

    static (byte[] pixels, int w, int h) DecodePngBgra(IntPtr factory, byte[] pngBytes)
    {
        var memBuf = Marshal.AllocHGlobal(pngBytes.Length);
        Marshal.Copy(pngBytes, 0, memBuf, pngBytes.Length);
        var stream = IntPtr.Zero; var decoder = IntPtr.Zero; var frame = IntPtr.Zero; var converter = IntPtr.Zero;
        try
        {
            VT<F_CreateStream>(factory, 14)(factory, out stream);
            VT<F_InitMem>(stream, 16)(stream, memBuf, (uint)pngBytes.Length);
            VT<F_CreateDecoder>(factory, 4)(factory, stream, IntPtr.Zero, 0, out decoder);
            VT<F_GetFrame>(decoder, 13)(decoder, 0, out frame);
            VT<F_GetSize>(frame, 3)(frame, out var w, out var h);
            VT<F_CreateConverter>(factory, 10)(factory, out converter);
            var dst = G_BGRA32;
            VT<F_ConvInit>(converter, 8)(converter, frame, ref dst, 0, IntPtr.Zero, 0.0, 0);
            int stride = (int)w * 4;
            int byteCount = stride * (int)h;
            var pixels = new byte[byteCount];
            var pinHandle = GCHandle.Alloc(pixels, GCHandleType.Pinned);
            try { VT<F_CopyPixels>(converter, 7)(converter, IntPtr.Zero, (uint)stride, (uint)byteCount, pinHandle.AddrOfPinnedObject()); }
            finally { pinHandle.Free(); }
            return (pixels, (int)w, (int)h);
        }
        finally
        {
            if (converter != IntPtr.Zero) Marshal.Release(converter);
            if (frame != IntPtr.Zero) Marshal.Release(frame);
            if (decoder != IntPtr.Zero) Marshal.Release(decoder);
            if (stream != IntPtr.Zero) Marshal.Release(stream);
            Marshal.FreeHGlobal(memBuf);
        }
    }

    static byte[] EncodeBgraToPng(IntPtr factory, byte[] pixels, int w, int h)
    {
        CreateStreamOnHGlobal(IntPtr.Zero, true, out var outStream);
        var encoder = IntPtr.Zero; var frame = IntPtr.Zero; var propBag = IntPtr.Zero;
        try
        {
            var c = G_PngContainer;
            VT<F_CreateEncoder>(factory, 8)(factory, ref c, IntPtr.Zero, out encoder);
            VT<F_EncInit>(encoder, 3)(encoder, outStream, 2);
            VT<F_CreateNewFrame>(encoder, 10)(encoder, out frame, out propBag);
            VT<F_FrameInit>(frame, 3)(frame, propBag);
            VT<F_SetSize>(frame, 4)(frame, (uint)w, (uint)h);
            var pf = G_BGRA32;
            VT<F_SetPixelFormat>(frame, 6)(frame, ref pf);
            int stride = w * 4;
            var pinHandle = GCHandle.Alloc(pixels, GCHandleType.Pinned);
            try { VT<F_FrameWritePixels>(frame, 10)(frame, (uint)h, (uint)stride, (uint)pixels.Length, pinHandle.AddrOfPinnedObject()); }
            finally { pinHandle.Free(); }
            VT<F_Commit>(frame, 12)(frame);
            VT<F_Commit>(encoder, 11)(encoder);

            GetHGlobalFromStream(outStream, out var hg);
            uint sz = GlobalSize(hg);
            var p = GlobalLock(hg);
            try
            {
                var result = new byte[sz];
                Marshal.Copy(p, result, 0, (int)sz);
                return result;
            }
            finally { GlobalUnlock(hg); }
        }
        finally
        {
            if (propBag != IntPtr.Zero) Marshal.Release(propBag);
            if (frame != IntPtr.Zero) Marshal.Release(frame);
            if (encoder != IntPtr.Zero) Marshal.Release(encoder);
            Marshal.Release(outStream);
        }
    }

    internal static byte[] Stitch(List<byte[]> pngs)
    {
        if (pngs.Count == 1) return pngs[0];
        var clsid = G_WICFactory_C; var iid = G_WICFactory_I;
        CoCreateInstance(ref clsid, IntPtr.Zero, 1, ref iid, out var factory);
        try
        {
            var pages = new List<(byte[] pixels, int w, int h)>();
            foreach (var b in pngs) pages.Add(DecodePngBgra(factory, b));

            int W = pages.Max(p => p.w);
            int H = pages.Sum(p => p.h);
            int targetStride = W * 4;
            var target = new byte[targetStride * H];
            for (int i = 0; i < target.Length; i++) target[i] = 0xFF;
            int yOff = 0;
            foreach (var p in pages)
            {
                int srcStride = p.w * 4;
                for (int row = 0; row < p.h; row++)
                    Array.Copy(p.pixels, row * srcStride, target, (yOff + row) * targetStride, srcStride);
                yOff += p.h;
            }
            return EncodeBgraToPng(factory, target, W, H);
        }
        finally { Marshal.Release(factory); }
    }

    // Tile equal-size cell PNGs into an N-column grid on a white background
    // with uniform gap and outer padding (all in pixels). Reuses the same
    // BGRA decode/encode primitives as Stitch. Cells are placed top-left in
    // their cell; the last row may be partially filled.
    internal static byte[] StitchGrid(List<byte[]> cells, int cols, int gap, int pad)
    {
        if (cells.Count == 0) throw new InvalidOperationException("no cells to tile");
        if (cols < 1) cols = 1;
        var clsid = G_WICFactory_C; var iid = G_WICFactory_I;
        CoCreateInstance(ref clsid, IntPtr.Zero, 1, ref iid, out var factory);
        try
        {
            var imgs = new List<(byte[] pixels, int w, int h)>();
            foreach (var b in cells) imgs.Add(DecodePngBgra(factory, b));

            int cellW = imgs.Max(i => i.w);
            int cellH = imgs.Max(i => i.h);
            int rows = (imgs.Count + cols - 1) / cols;
            int W = pad * 2 + cols * cellW + (cols - 1) * gap;
            int H = pad * 2 + rows * cellH + (rows - 1) * gap;
            int targetStride = W * 4;
            var target = new byte[targetStride * H];
            for (int i = 0; i < target.Length; i++) target[i] = 0xFF;
            for (int idx = 0; idx < imgs.Count; idx++)
            {
                int r = idx / cols, c = idx % cols;
                int x0 = pad + c * (cellW + gap);
                int y0 = pad + r * (cellH + gap);
                var p = imgs[idx];
                int srcStride = p.w * 4;
                for (int row = 0; row < p.h; row++)
                    Array.Copy(p.pixels, row * srcStride, target, (y0 + row) * targetStride + x0 * 4, srcStride);
            }
            return EncodeBgraToPng(factory, target, W, H);
        }
        finally { Marshal.Release(factory); }
    }

    static string DocxToPdf(string docx)
    {
        var pdf = Path.Combine(Path.GetTempPath(), $"_w_{Guid.NewGuid():N}.pdf");
        var clsid = G_Word; var iid = G_IDispatch;
        CoCreateInstance(ref clsid, IntPtr.Zero, 4, ref iid, out var word);
        try
        {
            var name = (string?)DispGet(word, "Name") ?? "";
            if (!name.Contains("Microsoft Word", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("word_not_authentic: " + name);

            DispSet(word, "Visible", false);
            DispSet(word, "DisplayAlerts", 0);
            try { DispSet(word, "AutomationSecurity", 3); } catch { }

            var docs = (IntPtr)DispGet(word, "Documents")!;
            try
            {
                var doc = (IntPtr)DispMethod(docs, "Open", docx, MISSING, true, false)!;
                try { DispMethod(doc, "SaveAs2", pdf, 17); }
                finally { try { DispMethod(doc, "Close", false); } catch { } Marshal.Release(doc); }
            }
            finally { Marshal.Release(docs); }
        }
        finally
        {
            try { DispMethod(word, "Quit"); } catch { }
            Marshal.Release(word);
        }
        return pdf;
    }

    // Render the requested PDF pages to one PNG each (native page pixels), in
    // page order. Shared by Render (vertical stitch) and RenderGrid (downscale +
    // tile). Returns an empty list if the filter selects nothing.
    static List<byte[]> PdfToPngList(string pdf, string pageFilter, int timeoutMs)
    {
        var fileFact = Factory("Windows.Storage.StorageFile", G_FileStatics);
        var pdfFact = Factory("Windows.Data.Pdf.PdfDocument", G_PdfDocStatics);
        var drFact = Factory("Windows.Storage.Streams.DataReader", G_DataReaderFact);
        try
        {
            var pathHs = Hs(pdf);
            IntPtr getOp;
            try { if (VT<F_OneIn>(fileFact, 6)(fileFact, pathHs, out getOp) != 0) throw new InvalidOperationException(); }
            finally { WindowsDeleteString(pathHs); }
            Wait(getOp, timeoutMs);
            VT<F_Out>(getOp, 8)(getOp, out var sf);
            Rel(getOp);

            VT<F_OneIn>(pdfFact, 6)(pdfFact, sf, out var loadOp);
            Wait(loadOp, timeoutMs);
            VT<F_Out>(loadOp, 8)(loadOp, out var doc);
            Rel(loadOp); Rel(sf);

            try
            {
                VT<F_OutU32>(doc, 7)(doc, out var pageCount);
                var pages = ParsePages(pageFilter, (int)pageCount);
                var pngs = new List<byte[]>();
                foreach (var p in pages) pngs.Add(RenderOne(doc, (uint)(p - 1), drFact, timeoutMs));
                return pngs;
            }
            finally { Rel(doc); }
        }
        finally { Rel(fileFact); Rel(pdfFact); Rel(drFact); }
    }

    // Box-average downscale of a BGRA buffer to dw×dh. Pure managed (no COM), so
    // it runs and unit-tests anywhere; only invoked on the Windows render path.
    // Area averaging (not bilinear) keeps downscaled page text legible in the
    // thumbnail. Upscaling degenerates to nearest-neighbour (n == 1), which is
    // fine — cells are only ever smaller than a native page.
    internal static byte[] ScaleBgra(byte[] src, int sw, int sh, int dw, int dh)
    {
        if (sw == dw && sh == dh) return src;
        var dst = new byte[dw * dh * 4];
        for (int y = 0; y < dh; y++)
        {
            int sy0 = (int)((long)y * sh / dh);
            int sy1 = Math.Min(sh, Math.Max(sy0 + 1, (int)((long)(y + 1) * sh / dh)));
            for (int x = 0; x < dw; x++)
            {
                int sx0 = (int)((long)x * sw / dw);
                int sx1 = Math.Min(sw, Math.Max(sx0 + 1, (int)((long)(x + 1) * sw / dw)));
                long b = 0, g = 0, r = 0, a = 0; int n = 0;
                for (int yy = sy0; yy < sy1; yy++)
                {
                    int rowOff = yy * sw * 4;
                    for (int xx = sx0; xx < sx1; xx++)
                    {
                        int si = rowOff + xx * 4;
                        b += src[si]; g += src[si + 1]; r += src[si + 2]; a += src[si + 3]; n++;
                    }
                }
                int di = (y * dw + x) * 4;
                dst[di] = (byte)(b / n); dst[di + 1] = (byte)(g / n);
                dst[di + 2] = (byte)(r / n); dst[di + 3] = (byte)(a / n);
            }
        }
        return dst;
    }

    public static bool RefreshFields(string docx, int timeoutMs = 180000)
    {
        bool ok = false;
        var th = new Thread(() =>
        {
            try
            {
                var clsid = G_Word; var iid = G_IDispatch;
                CoCreateInstance(ref clsid, IntPtr.Zero, 4, ref iid, out var word);
                try
                {
                    var name = (string?)DispGet(word, "Name") ?? "";
                    if (!name.Contains("Microsoft Word", StringComparison.OrdinalIgnoreCase)) return;
                    DispSet(word, "Visible", false);
                    DispSet(word, "DisplayAlerts", 0);
                    try { DispSet(word, "AutomationSecurity", 3); } catch { }
                    var docs = (IntPtr)DispGet(word, "Documents")!;
                    try
                    {
                        var doc = (IntPtr)DispMethod(docs, "Open", docx, MISSING, false, false)!;
                        try
                        {
                            var fields = (IntPtr)DispGet(doc, "Fields")!;
                            try { DispMethod(fields, "Update"); }
                            finally { Marshal.Release(fields); }
                            DispMethod(doc, "Save");
                            ok = true;
                        }
                        finally { try { DispMethod(doc, "Close", false); } catch { } Marshal.Release(doc); }
                    }
                    finally { Marshal.Release(docs); }
                }
                finally { try { DispMethod(word, "Quit"); } catch { } Marshal.Release(word); }
            }
            catch { }
        });
        th.SetApartmentState(ApartmentState.STA);
        th.IsBackground = true;
        th.Start();
        if (!th.Join(timeoutMs + 30000)) return false;
        return ok;
    }

    public static int? GetPageCount(string docx, int timeoutMs = 120000)
    {
        int? result = null;
        var th = new Thread(() =>
        {
            try
            {
                var clsid = G_Word; var iid = G_IDispatch;
                CoCreateInstance(ref clsid, IntPtr.Zero, 4, ref iid, out var word);
                try
                {
                    var name = (string?)DispGet(word, "Name") ?? "";
                    if (!name.Contains("Microsoft Word", StringComparison.OrdinalIgnoreCase)) return;
                    DispSet(word, "Visible", false);
                    DispSet(word, "DisplayAlerts", 0);
                    try { DispSet(word, "AutomationSecurity", 3); } catch { }
                    var docs = (IntPtr)DispGet(word, "Documents")!;
                    try
                    {
                        var doc = (IntPtr)DispMethod(docs, "Open", docx, MISSING, true, false)!;
                        try
                        {
                            var pages = DispMethod(doc, "ComputeStatistics", 2);
                            if (pages is int p) result = p;
                        }
                        finally { try { DispMethod(doc, "Close", false); } catch { } Marshal.Release(doc); }
                    }
                    finally { Marshal.Release(docs); }
                }
                finally { try { DispMethod(word, "Quit"); } catch { } Marshal.Release(word); }
            }
            catch { }
        });
        th.SetApartmentState(ApartmentState.STA);
        th.IsBackground = true;
        th.Start();
        if (!th.Join(timeoutMs + 30000)) return null;
        return result;
    }

    public static byte[]? Render(string docx, string pageFilter, int timeoutMs = 60000)
    {
        byte[]? result = null;
        Exception? error = null;
        var th = new Thread(() =>
        {
            string? pdf = null;
            try
            {
                pdf = DocxToPdf(docx);
                var pngs = PdfToPngList(pdf, pageFilter, timeoutMs);
                result = pngs.Count == 0 ? null : Stitch(pngs);
            }
            catch (Exception e)
            {
                error = e;
            }
            finally
            {
                if (pdf != null) try { File.Delete(pdf); } catch { }
            }
        });
        th.SetApartmentState(ApartmentState.STA);
        th.IsBackground = true;
        th.Start();
        if (!th.Join(timeoutMs + 30000))
            throw new TimeoutException("Word native rendering timed out.");
        if (error != null) ExceptionDispatchInfo.Capture(error).Throw();
        return result;
    }

    /// <summary>
    /// Render every page via real Word (Word → PDF → per-page raster), downscale
    /// each to <paramref name="cellW"/>×<paramref name="cellH"/>, and tile them
    /// into a <paramref name="cols"/>-column contact sheet. The docx analogue of
    /// PowerPointPngBackend.RenderGrid (which exports each slide at cell size via
    /// PowerPoint). Failures are surfaced to the caller; auto mode catches them
    /// and falls back to the HTML grid. cellW/cellH are the FINAL (already
    /// 1920-capped) cell size, so the stitched image needs no further cap.
    /// </summary>
    public static byte[]? RenderGrid(string docx, string pageFilter, int cellW, int cellH, int cols, int gap, int pad, int timeoutMs = 60000)
    {
        byte[]? result = null;
        Exception? error = null;
        var th = new Thread(() =>
        {
            string? pdf = null;
            try
            {
                pdf = DocxToPdf(docx);
                var nativePngs = PdfToPngList(pdf, pageFilter, timeoutMs);
                if (nativePngs.Count == 0) return;

                var clsid = G_WICFactory_C; var iid = G_WICFactory_I;
                CoCreateInstance(ref clsid, IntPtr.Zero, 1, ref iid, out var factory);
                try
                {
                    var cells = new List<byte[]>(nativePngs.Count);
                    foreach (var png in nativePngs)
                    {
                        var (px, w, h) = DecodePngBgra(factory, png);
                        var scaled = ScaleBgra(px, w, h, cellW, cellH);
                        cells.Add(EncodeBgraToPng(factory, scaled, cellW, cellH));
                    }
                    result = StitchGrid(cells, cols, gap, pad);
                }
                finally { Marshal.Release(factory); }
            }
            catch (Exception e) { error = e; }
            finally { if (pdf != null) try { File.Delete(pdf); } catch { } }
        });
        th.SetApartmentState(ApartmentState.STA);
        th.IsBackground = true;
        th.Start();
        if (!th.Join(timeoutMs + 30000))
            throw new TimeoutException("Word native grid rendering timed out.");
        if (error != null) ExceptionDispatchInfo.Capture(error).Throw();
        return result;
    }
}
