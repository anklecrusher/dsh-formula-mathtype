// MathTypeInjector.exe — insert a real MathType equation (Equation.DSMT4) at the
// caret of the active Microsoft Word document.
//
// Why this exists
// ---------------
// MathType's SDK call MTXFormEqn() is designed to *export* an equation object to
// text (MathType OLE on the clipboard -> LaTeX/MathML text). Trying to run it in
// the other direction (text -> MathType OLE on the clipboard) does not feed the
// equation parser: Word then embeds an object whose content is the literal text,
// which renders as plain LaTeX source or, after a big garbled input, as a black
// block.
//
// The quiet, reliable route is the one MathType's own Word integration uses:
//   1. let Word create an empty Equation.DSMT4 OLE object at the caret,
//   2. ask that OLE server for its IDataObject,
//   3. hand it the formula in a clipboard format MathType understands
//      ("TeX Input Language" parses LaTeX, "MathML" parses MathML),
//   4. close the object so it saves the parsed equation.
//
// This file intentionally uses late binding (`dynamic`) instead of the Office
// PIA, so it compiles with the .NET Framework compiler that ships with Windows
// and has no extra runtime dependency. The COM interface declarations below are
// the standard objidl ones.
//
// The C# source is compiled on first use by helper.py:
//   csc -nologo -target:exe -out:MathTypeInjector.exe MathTypeInjector.cs
//
// Exit codes: 0 success, 1 failure (a JSON object with "ok":false is printed).

using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using ComIDataObject = System.Runtime.InteropServices.ComTypes.IDataObject;

namespace MathTypeInjector
{
    [ComImport]
    [Guid("00000112-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IOleObject
    {
        [PreserveSig] int SetClientSite(IntPtr pClientSite);
        [PreserveSig] int GetClientSite(out IntPtr ppClientSite);
        [PreserveSig] int SetHostNames([MarshalAs(UnmanagedType.LPWStr)] string szContainerApp, [MarshalAs(UnmanagedType.LPWStr)] string szContainerObj);
        [PreserveSig] int Close(uint dwSaveOption);
        [PreserveSig] int SetMoniker(uint dwWhichMoniker, IntPtr pmk);
        [PreserveSig] int GetMoniker(uint dwAssign, uint dwWhichMoniker, out IntPtr ppmk);
        [PreserveSig] int InitFromData(ComIDataObject pDataObject, int fCreation, uint dwReserved);
        [PreserveSig] int GetClipboardData(uint dwReserved, out ComIDataObject ppDataObject);
        [PreserveSig] int DoVerb(int iVerb, IntPtr lpmsg, IntPtr pActiveSite, int lindex, IntPtr hwndParent, IntPtr lprcPosRect);
        [PreserveSig] int EnumVerbs(out IntPtr ppEnumOleVerb);
        [PreserveSig] int Update();
        [PreserveSig] int IsUpToDate();
        [PreserveSig] int GetUserClassID(out Guid pClsid);
        [PreserveSig] int GetUserType(uint dwFormOfType, [MarshalAs(UnmanagedType.LPWStr)] out string pszUserType);
        [PreserveSig] int SetExtent(uint dwDrawAspect, IntPtr psizel);
        [PreserveSig] int GetExtent(uint dwDrawAspect, IntPtr psizel);
        [PreserveSig] int Advise(IntPtr pAdvSink, out uint pdwConnection);
        [PreserveSig] int Unadvise(uint dwConnection);
        [PreserveSig] int EnumAdvise(out IntPtr ppenumAdvise);
        [PreserveSig] int GetMiscStatus(uint dwAspect, out uint pdwStatus);
        [PreserveSig] int SetColorScheme(IntPtr pLogpal);
    }

    internal static class Program
    {
        private const uint GMEM_MOVEABLE = 0x0002;
        private const uint GMEM_ZEROINIT = 0x0040;
        private const uint OLECLOSE_SAVEIFDIRTY = 0;
        private const int WORD_XML_DOCUMENT = 12;

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern uint RegisterClipboardFormat(string lpszFormat);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr GlobalAlloc(uint uFlags, UIntPtr dwBytes);
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr GlobalLock(IntPtr hMem);
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool GlobalUnlock(IntPtr hMem);
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr GlobalFree(IntPtr hMem);
        [DllImport("ole32.dll", PreserveSig = false)]
        private static extern void CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string progId, out Guid clsid);

        private static readonly Dictionary<string, uint> FormatIds = new Dictionary<string, uint>();

        private static uint FormatId(string name)
        {
            uint id;
            if (!FormatIds.TryGetValue(name, out id))
            {
                id = RegisterClipboardFormat(name);
                FormatIds[name] = id;
            }
            return id;
        }

        private static string NormalizeFormat(string format)
        {
            if (string.Equals(format, "tex", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(format, "latex", StringComparison.OrdinalIgnoreCase))
                return "TeX Input Language";
            if (string.Equals(format, "amslatex", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(format, "ams", StringComparison.OrdinalIgnoreCase))
                return "AMS-LaTeX";
            if (string.Equals(format, "mathml", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(format, "mml", StringComparison.OrdinalIgnoreCase))
                return "MathML";
            return format;
        }

        private static object ActiveWord()
        {
            try
            {
                return Marshal.GetActiveObject("Word.Application");
            }
            catch (COMException)
            {
                throw new InvalidOperationException(
                    "没有找到正在运行的 Word。请先打开要写入的报告文档，把光标放到插入位置，再重新触发。");
            }
        }

        private static object NewWord()
        {
            Guid clsid;
            CLSIDFromProgID("Word.Application", out clsid);
            Type type = Type.GetTypeFromCLSID(clsid);
            return Activator.CreateInstance(type);
        }

        private static string Escape(string value)
        {
            if (value == null) return "";
            var builder = new StringBuilder(value.Length + 16);
            foreach (char ch in value)
            {
                switch (ch)
                {
                    case '\\': builder.Append("\\\\"); break;
                    case '"': builder.Append("\\\""); break;
                    case '\r': builder.Append(' '); break;
                    case '\n': builder.Append(' '); break;
                    case '\t': builder.Append(' '); break;
                    default:
                        // Keep the JSON pure ASCII so console code pages can never corrupt it.
                        if (ch < 0x20 || ch > 0x7E)
                            builder.Append("\\u").Append(((int)ch).ToString("x4"));
                        else
                            builder.Append(ch);
                        break;
                }
            }
            return builder.ToString();
        }

        private static void InjectAtRange(dynamic doc, dynamic range, string formula, string formatName)
        {
            object missing = Type.Missing;
            dynamic shape = doc.InlineShapes.AddOLEObject(
                "Equation.DSMT4", "", false, false, missing, missing, missing, range);
            dynamic oleFormat = shape.OLEFormat;
            oleFormat.DoVerb(2);
            System.Threading.Thread.Sleep(120);
            object raw = oleFormat.Object;
            ComIDataObject dataObject = raw as ComIDataObject;
            IOleObject oleObject = raw as IOleObject;
            if (dataObject == null)
                throw new InvalidOperationException("Equation.DSMT4 没有暴露 IDataObject，无法写入公式内容。");
            SetUnicodeData(dataObject, formatName, formula);
            if (oleObject != null)
            {
                int hr = oleObject.Close(OLECLOSE_SAVEIFDIRTY);
                if (hr < 0) Marshal.ThrowExceptionForHR(hr);
            }
            System.Threading.Thread.Sleep(80);
            Console.WriteLine(
                "{\"ok\":true,\"width\":\"" +
                shape.Width.ToString(System.Globalization.CultureInfo.InvariantCulture) +
                "\",\"height\":\"" +
                shape.Height.ToString(System.Globalization.CultureInfo.InvariantCulture) + "\"}");
        }

        private static int Main(string[] args)
        {
            try { Console.OutputEncoding = Encoding.UTF8; } catch { }
            try
            {
                string mode = "insert";
                string file = null, output = null, format = "tex";
                for (int i = 0; i < args.Length; i++)
                {
                    switch (args[i])
                    {
                        case "--mode": mode = args[++i]; break;
                        case "--file": file = args[++i]; break;
                        case "--output": output = args[++i]; break;
                        case "--format": format = args[++i]; break;
                        default:
                            throw new ArgumentException("未知参数：" + args[i]);
                    }
                }
                if (file == null) throw new ArgumentException("缺少 --file 参数");
                string formula = File.ReadAllText(file, Encoding.UTF8);
                if (formula.Trim().Length == 0) throw new ArgumentException("公式文件为空");
                string formatName = NormalizeFormat(format);

                if (mode == "insert")
                {
                    dynamic app = ActiveWord();
                    dynamic docs = app.Documents;
                    if (docs.Count < 1)
                        throw new InvalidOperationException("Word 里没有打开的文档，请先打开报告文档并定位光标。");
                    dynamic doc = app.ActiveDocument;
                    dynamic range = app.Selection.Range;
                    InjectAtRange(doc, range, formula, formatName);
                    return 0;
                }

                if (mode == "create-docx")
                {
                    if (output == null) throw new ArgumentException("create-docx 需要 --output 参数");
                    dynamic word = NewWord();
                    dynamic wdoc = null;
                    try
                    {
                        word.Visible = false;
                        wdoc = word.Documents.Add();
                        InjectAtRange(wdoc, wdoc.Range(0, 0), formula, formatName);
                        wdoc.SaveAs2(Path.GetFullPath(output), WORD_XML_DOCUMENT);
                    }
                    finally
                    {
                        try { if (wdoc != null) wdoc.Close(false); } catch { }
                        try { word.Quit(false); } catch { }
                    }
                    Console.WriteLine("{\"ok\":true,\"mode\":\"create-docx\",\"output\":\"" + Escape(output) + "\"}");
                    return 0;
                }

                throw new ArgumentException("未知 mode：" + mode);
            }
            catch (Exception ex)
            {
                Console.WriteLine("{\"ok\":false,\"error\":\"" + Escape(ex.Message) + "\"}");
                Console.Error.WriteLine(ex.ToString());
                return 1;
            }
        }

        private static void SetUnicodeData(ComIDataObject dataObject, string formatName, string value)
        {
            byte[] bytes = Encoding.Unicode.GetBytes(value + "\0");
            IntPtr hGlobal = GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, (UIntPtr)bytes.Length);
            if (hGlobal == IntPtr.Zero)
                throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error(), "GlobalAlloc 失败");
            bool ownershipTransferred = false;
            try
            {
                IntPtr ptr = GlobalLock(hGlobal);
                if (ptr == IntPtr.Zero)
                    throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error(), "GlobalLock 失败");
                try { Marshal.Copy(bytes, 0, ptr, bytes.Length); }
                finally { GlobalUnlock(hGlobal); }

                FORMATETC formatetc = new FORMATETC
                {
                    cfFormat = (short)FormatId(formatName),
                    dwAspect = DVASPECT.DVASPECT_CONTENT,
                    lindex = -1,
                    ptd = IntPtr.Zero,
                    tymed = TYMED.TYMED_HGLOBAL
                };
                STGMEDIUM medium = new STGMEDIUM
                {
                    tymed = TYMED.TYMED_HGLOBAL,
                    unionmember = hGlobal,
                    pUnkForRelease = null
                };
                dataObject.SetData(ref formatetc, ref medium, true);
                ownershipTransferred = true;
            }
            finally
            {
                if (!ownershipTransferred) GlobalFree(hGlobal);
            }
        }
    }
}