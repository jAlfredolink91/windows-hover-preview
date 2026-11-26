using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Windows.Automation;
using Windows.Foundation;
using Windows.Data.Pdf;
using Windows.Storage;
using Windows.Storage.Streams;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;

// Vista Previa (al pasar el cursor en el Explorador de Windows).
// Soporta: PDF (multipagina + scroll), imagenes, Word y Excel.
// Motor PDF: Windows.Data.Pdf (nativo de Windows). Office: conversion a PDF via COM.
class PdfHoverPreview
{
    // ---------- Win32 ----------
    [StructLayout(LayoutKind.Sequential)]
    struct POINT { public int X; public int Y; public POINT(int x, int y) { X = x; Y = y; } }
    [DllImport("user32.dll")] static extern IntPtr WindowFromPoint(POINT p);
    [DllImport("user32.dll")] static extern IntPtr GetAncestor(IntPtr hwnd, uint flags);
    [DllImport("user32.dll")] static extern short GetKeyState(int nVirtKey);
    const uint GA_ROOT = 2;
    const int VK_SHIFT = 0x10;
    const int VK_CONTROL = 0x11;

    [StructLayout(LayoutKind.Sequential)]
    struct MSLLHOOKSTRUCT { public POINT pt; public uint mouseData; public uint flags; public uint time; public IntPtr dwExtraInfo; }
    delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll", SetLastError = true)] static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);
    [DllImport("user32.dll", SetLastError = true)] static extern bool UnhookWindowsHookEx(IntPtr hhk);
    [DllImport("user32.dll")] static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);
    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)] static extern IntPtr GetModuleHandle(string lpModuleName);
    const int WH_MOUSE_LL = 14;
    const int WM_MOUSEWHEEL = 0x020A;

    // ---------- Tipos de archivo ----------
    const int KIND_NONE = 0, KIND_PDF = 1, KIND_IMAGE = 2, KIND_OFFICE = 3, KIND_TEXT = 4, KIND_ZIP = 5, KIND_EMAIL = 6;
    static readonly string[] ImageExts = { ".jpg", ".jpeg", ".jfif", ".png", ".gif", ".bmp", ".tif", ".tiff", ".ico", ".webp", ".heic", ".heif" };
    static readonly string[] OfficeExts = { ".doc", ".docx", ".docm", ".rtf", ".xls", ".xlsx", ".xlsm", ".xlsb", ".ppt", ".pptx", ".pptm" };
    static readonly string[] TextExts = { ".txt", ".csv", ".tsv", ".log", ".xml", ".json", ".ini", ".cfg", ".conf", ".config", ".md", ".yml", ".yaml",
                                          ".cs", ".js", ".ts", ".html", ".htm", ".css", ".bat", ".cmd", ".ps1", ".py", ".sql", ".java", ".c", ".cpp", ".h", ".php", ".vb", ".reg" };
    static readonly string[] ZipExts = { ".zip" };
    static readonly string[] EmailExts = { ".eml", ".msg" };
    static readonly string[] AllExts;
    static PdfHoverPreview()
    {
        List<string> all = new List<string>();
        all.Add(".pdf");
        all.AddRange(ImageExts);
        all.AddRange(OfficeExts);
        all.AddRange(TextExts);
        all.AddRange(ZipExts);
        all.AddRange(EmailExts);
        AllExts = all.ToArray();
    }
    static int KindOf(string path)
    {
        string e = Path.GetExtension(path).ToLowerInvariant();
        if (e == ".pdf") return KIND_PDF;
        foreach (string x in ImageExts) if (x == e) return KIND_IMAGE;
        foreach (string x in OfficeExts) if (x == e) return KIND_OFFICE;
        foreach (string x in TextExts) if (x == e) return KIND_TEXT;
        foreach (string x in ZipExts) if (x == e) return KIND_ZIP;
        foreach (string x in EmailExts) if (x == e) return KIND_EMAIL;
        return KIND_NONE;
    }

    class Hit { public string Path; public int Kind; }

    // ---------- Espera manual de async WinRT ----------
    static T AwaitOp<T>(IAsyncOperation<T> op)
    {
        ManualResetEventSlim done = new ManualResetEventSlim(false);
        op.Completed = delegate(IAsyncOperation<T> o, AsyncStatus s) { done.Set(); };
        done.Wait();
        if (op.Status != AsyncStatus.Completed) throw new Exception("async op status=" + op.Status);
        return op.GetResults();
    }
    static void AwaitAct(IAsyncAction act)
    {
        ManualResetEventSlim done = new ManualResetEventSlim(false);
        act.Completed = delegate(IAsyncAction a, AsyncStatus s) { done.Set(); };
        done.Wait();
        if (act.Status != AsyncStatus.Completed) throw new Exception("async act status=" + act.Status);
        act.GetResults();
    }

    static byte[] RenderPageBytes(PdfPage page, uint width)
    {
        InMemoryRandomAccessStream stream = new InMemoryRandomAccessStream();
        PdfPageRenderOptions opts = new PdfPageRenderOptions();
        opts.DestinationWidth = width;
        AwaitAct(page.RenderToStreamAsync(stream, opts));
        uint size = (uint)stream.Size;
        DataReader reader = new DataReader(stream.GetInputStreamAt(0));
        AwaitOp(reader.LoadAsync(size));
        byte[] b = new byte[size];
        reader.ReadBytes(b);
        return b;
    }

    // ---------- Conversion Office -> PDF (COM, cacheado) ----------
    static string Hash(string s)
    {
        using (MD5 md5 = MD5.Create())
        {
            byte[] h = md5.ComputeHash(Encoding.UTF8.GetBytes(s));
            StringBuilder sb = new StringBuilder();
            foreach (byte x in h) sb.Append(x.ToString("x2"));
            return sb.ToString();
        }
    }

    static string CacheDir()
    {
        string d = Path.Combine(Path.GetTempPath(), "VistaPreviaPDF_cache");
        Directory.CreateDirectory(d);
        return d;
    }

    static string OfficeToPdf(string path)
    {
        FileInfo fi = new FileInfo(path);
        string key = path.ToLowerInvariant() + "|" + fi.LastWriteTimeUtc.Ticks;
        string outPdf = Path.Combine(CacheDir(), Hash(key) + ".pdf");
        if (File.Exists(outPdf) && new FileInfo(outPdf).Length > 0) return outPdf;

        string ext = Path.GetExtension(path).ToLowerInvariant();
        if (ext.StartsWith(".doc") || ext == ".rtf") WordExport(path, outPdf);
        else if (ext.StartsWith(".xls")) ExcelExport(path, outPdf);
        else if (ext.StartsWith(".ppt")) PowerPointExport(path, outPdf);
        else throw new Exception("formato no soportado: " + ext);

        if (!File.Exists(outPdf)) throw new Exception("no se genero el PDF");
        return outPdf;
    }

    static System.Collections.Generic.HashSet<int> Pids(string procName)
    {
        System.Collections.Generic.HashSet<int> set = new System.Collections.Generic.HashSet<int>();
        try { foreach (System.Diagnostics.Process p in System.Diagnostics.Process.GetProcessesByName(procName)) set.Add(p.Id); }
        catch { }
        return set;
    }
    static int NewPid(string procName, System.Collections.Generic.HashSet<int> before)
    {
        try { foreach (System.Diagnostics.Process p in System.Diagnostics.Process.GetProcessesByName(procName)) if (!before.Contains(p.Id)) return p.Id; }
        catch { }
        return 0;
    }
    static void KillPid(int pid)
    {
        if (pid <= 0) return;
        try { System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(pid); if (!p.HasExited) p.Kill(); }
        catch { }
    }

    static void WordExport(string path, string outPdf)
    {
        Type t = Type.GetTypeFromProgID("Word.Application");
        if (t == null) throw new Exception("Word no esta instalado");
        dynamic app = null; dynamic doc = null; dynamic docs = null;
        System.Collections.Generic.HashSet<int> before = Pids("WINWORD");
        int myPid = 0;
        try
        {
            app = Activator.CreateInstance(t);
            myPid = NewPid("WINWORD", before);
            try { app.Visible = false; } catch { }
            try { app.DisplayAlerts = 0; } catch { }
            docs = app.Documents;
            doc = docs.Open(path, false, true); // FileName, ConfirmConversions, ReadOnly
            doc.ExportAsFixedFormat(outPdf, 17); // wdExportFormatPDF
        }
        finally
        {
            try { if (doc != null) doc.Close(false); } catch { }
            try { if (app != null) app.Quit(); } catch { }
            try { if (doc != null) Marshal.FinalReleaseComObject(doc); } catch { }
            try { if (docs != null) Marshal.FinalReleaseComObject(docs); } catch { }
            try { if (app != null) Marshal.FinalReleaseComObject(app); } catch { }
            doc = null; docs = null; app = null;
            GC.Collect(); GC.WaitForPendingFinalizers();
            GC.Collect(); GC.WaitForPendingFinalizers();
            KillPid(myPid); // garantizar que no quede un WINWORD zombie bloqueando archivos
        }
    }

    static void ExcelExport(string path, string outPdf)
    {
        Type t = Type.GetTypeFromProgID("Excel.Application");
        if (t == null) throw new Exception("Excel no esta instalado");
        dynamic app = null; dynamic wb = null; dynamic wbs = null;
        System.Collections.Generic.HashSet<int> before = Pids("EXCEL");
        int myPid = 0;
        try
        {
            app = Activator.CreateInstance(t);
            myPid = NewPid("EXCEL", before);
            try { app.Visible = false; } catch { }
            try { app.DisplayAlerts = false; } catch { }
            try { app.ScreenUpdating = false; } catch { }
            wbs = app.Workbooks;
            wb = wbs.Open(path, 0, true); // FileName, UpdateLinks, ReadOnly
            wb.ExportAsFixedFormat(0, outPdf); // xlTypePDF
        }
        finally
        {
            try { if (wb != null) wb.Close(false); } catch { }
            try { if (app != null) app.Quit(); } catch { }
            try { if (wb != null) Marshal.FinalReleaseComObject(wb); } catch { }
            try { if (wbs != null) Marshal.FinalReleaseComObject(wbs); } catch { }
            try { if (app != null) Marshal.FinalReleaseComObject(app); } catch { }
            wb = null; wbs = null; app = null;
            GC.Collect(); GC.WaitForPendingFinalizers();
            GC.Collect(); GC.WaitForPendingFinalizers();
            KillPid(myPid); // garantizar que no quede un EXCEL zombie bloqueando archivos
        }
    }

    static void PowerPointExport(string path, string outPdf)
    {
        Type t = Type.GetTypeFromProgID("PowerPoint.Application");
        if (t == null) throw new Exception("PowerPoint no esta instalado");
        dynamic app = null; dynamic pres = null; dynamic press = null;
        System.Collections.Generic.HashSet<int> before = Pids("POWERPNT");
        int myPid = 0;
        try
        {
            app = Activator.CreateInstance(t);
            myPid = NewPid("POWERPNT", before);
            try { app.DisplayAlerts = 1; } catch { } // ppAlertsNone
            press = app.Presentations;
            // Open(FileName, ReadOnly=msoTrue(-1), Untitled=msoFalse(0), WithWindow=msoFalse(0))
            pres = press.Open(path, -1, 0, 0);
            pres.SaveAs(outPdf, 32); // 32 = ppSaveAsPDF (mas confiable que ExportAsFixedFormat)
        }
        finally
        {
            try { if (pres != null) pres.Close(); } catch { }
            try { if (app != null) app.Quit(); } catch { }
            try { if (pres != null) Marshal.FinalReleaseComObject(pres); } catch { }
            try { if (press != null) Marshal.FinalReleaseComObject(press); } catch { }
            try { if (app != null) Marshal.FinalReleaseComObject(app); } catch { }
            pres = null; press = null; app = null;
            GC.Collect(); GC.WaitForPendingFinalizers();
            GC.Collect(); GC.WaitForPendingFinalizers();
            KillPid(myPid); // garantizar que no quede un POWERPNT zombie
        }
    }

    // ---------- Configuracion persistente (%LOCALAPPDATA%\VistaPreviaPDF\config.txt) ----------
    static Dictionary<string, string> _cfg;
    static string ConfigFile()
    {
        string d = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "VistaPreviaPDF");
        try { Directory.CreateDirectory(d); } catch { }
        return Path.Combine(d, "config.txt");
    }
    static void LoadCfg()
    {
        _cfg = new Dictionary<string, string>();
        try
        {
            string f = ConfigFile();
            if (File.Exists(f))
                foreach (string ln in File.ReadAllLines(f))
                {
                    int i = ln.IndexOf('=');
                    if (i > 0) _cfg[ln.Substring(0, i).Trim()] = ln.Substring(i + 1).Trim();
                }
        }
        catch { }
    }
    static string CfgGet(string k, string def)
    {
        if (_cfg == null) LoadCfg();
        string v;
        if (_cfg.TryGetValue(k, out v)) return v;
        return def;
    }
    static void CfgSet(string k, string v)
    {
        if (_cfg == null) LoadCfg();
        _cfg[k] = v;
        try
        {
            StringBuilder sb = new StringBuilder();
            foreach (KeyValuePair<string, string> kv in _cfg) sb.AppendLine(kv.Key + "=" + kv.Value);
            File.WriteAllText(ConfigFile(), sb.ToString());
        }
        catch { }
    }

    static bool IsDarkTheme()
    {
        try
        {
            object v = Microsoft.Win32.Registry.GetValue(
                "HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize",
                "AppsUseLightTheme", 1);
            if (v is int) return ((int)v) == 0;
        }
        catch { }
        return false;
    }

    // ---------- Lectores de contenido (texto, ZIP, correo) ----------
    static string ReadTextPreview(string path)
    {
        const int maxBytes = 400 * 1024;
        byte[] data = File.ReadAllBytes(path); // lee y suelta: no bloquea el archivo
        int n = data.Length; bool trunc = false;
        if (n > maxBytes) { n = maxBytes; trunc = true; }
        Encoding enc = Encoding.UTF8; int start = 0;
        if (n >= 3 && data[0] == 0xEF && data[1] == 0xBB && data[2] == 0xBF) { enc = Encoding.UTF8; start = 3; }
        else if (n >= 2 && data[0] == 0xFF && data[1] == 0xFE) { enc = Encoding.Unicode; start = 2; }
        else if (n >= 2 && data[0] == 0xFE && data[1] == 0xFF) { enc = Encoding.BigEndianUnicode; start = 2; }
        string text;
        try { text = enc.GetString(data, start, n - start); }
        catch { text = Encoding.Default.GetString(data, 0, n); }
        text = text.Replace("\r\n", "\n").Replace("\r", "\n").Replace("\n", "\r\n");
        if (trunc) text += "\r\n\r\n--- (archivo grande: mostrando el inicio) ---";
        return text;
    }

    static string FormatSize(long b)
    {
        if (b < 1024) return b + " B";
        if (b < 1024 * 1024) return (b / 1024.0).ToString("0.0") + " KB";
        return (b / (1024.0 * 1024)).ToString("0.0") + " MB";
    }

    static string ReadZipListing(string path)
    {
        StringBuilder sb = new StringBuilder();
        sb.AppendLine("Contenido de: " + Path.GetFileName(path));
        sb.AppendLine(new string('-', 44));
        byte[] data = File.ReadAllBytes(path);
        using (MemoryStream ms = new MemoryStream(data))
        using (System.IO.Compression.ZipArchive zip = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Read))
        {
            long total = 0; int count = 0;
            foreach (System.IO.Compression.ZipArchiveEntry e in zip.Entries)
            {
                if (string.IsNullOrEmpty(e.Name)) { sb.AppendLine("  [carpeta]  " + e.FullName); continue; }
                sb.AppendLine(FormatSize(e.Length).PadLeft(10) + "   " + e.FullName);
                total += e.Length; count++;
            }
            sb.AppendLine(new string('-', 44));
            sb.AppendLine(count + " archivo(s),  " + FormatSize(total) + " sin comprimir");
        }
        return sb.ToString().Replace("\r\n", "\n").Replace("\r", "\n").Replace("\n", "\r\n");
    }

    static string HeaderValue(string headers, string name)
    {
        foreach (string ln in headers.Replace("\r\n", "\n").Split('\n'))
        {
            if (ln.Length > name.Length && ln.Substring(0, name.Length + 1).ToLowerInvariant() == name.ToLowerInvariant() + ":")
                return ln.Substring(name.Length + 1).Trim();
        }
        return "";
    }

    static string ReadEmailPreview(string path)
    {
        string ext = Path.GetExtension(path).ToLowerInvariant();
        if (ext == ".msg") return ReadMsgViaOutlook(path);
        // .eml (texto MIME): mostramos cabeceras + cuerpo (vista rapida)
        byte[] data = File.ReadAllBytes(path);
        string raw = Encoding.UTF8.GetString(data);
        int sep = raw.IndexOf("\r\n\r\n"); int sl = 4;
        if (sep < 0) { sep = raw.IndexOf("\n\n"); sl = 2; }
        string head = (sep >= 0) ? raw.Substring(0, sep) : raw;
        string body = (sep >= 0) ? raw.Substring(sep + sl) : "";
        StringBuilder sb = new StringBuilder();
        sb.AppendLine("De:     " + HeaderValue(head, "From"));
        sb.AppendLine("Para:   " + HeaderValue(head, "To"));
        sb.AppendLine("Asunto: " + HeaderValue(head, "Subject"));
        sb.AppendLine("Fecha:  " + HeaderValue(head, "Date"));
        sb.AppendLine(new string('-', 50));
        sb.AppendLine();
        if (body.Length > 20000) body = body.Substring(0, 20000) + "\r\n\r\n--- (recortado) ---";
        sb.Append(body);
        return sb.ToString().Replace("\r\n", "\n").Replace("\r", "\n").Replace("\n", "\r\n");
    }

    static string ReadMsgViaOutlook(string path)
    {
        Type t = Type.GetTypeFromProgID("Outlook.Application");
        if (t == null) return "No se puede mostrar .msg aqui: requiere Microsoft Outlook instalado.";
        dynamic app = null; dynamic ns = null; dynamic item = null;
        try
        {
            app = Activator.CreateInstance(t);
            ns = app.GetNamespace("MAPI");
            item = ns.OpenSharedItem(path);
            string de = "", para = "", asunto = "", fecha = "", body = "";
            try { de = (string)item.SenderName; } catch { }
            try { para = (string)item.To; } catch { }
            try { asunto = (string)item.Subject; } catch { }
            try { fecha = Convert.ToString(item.ReceivedTime); } catch { }
            try { body = (string)item.Body; } catch { }
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("De:     " + de);
            sb.AppendLine("Para:   " + para);
            sb.AppendLine("Asunto: " + asunto);
            sb.AppendLine("Fecha:  " + fecha);
            sb.AppendLine(new string('-', 50));
            sb.AppendLine();
            if (body != null)
            {
                if (body.Length > 30000) body = body.Substring(0, 30000) + "\r\n\r\n--- (recortado) ---";
                sb.Append(body);
            }
            return sb.ToString().Replace("\r\n", "\n").Replace("\r", "\n").Replace("\n", "\r\n");
        }
        catch (Exception e) { return "No se pudo leer el correo (.msg): " + e.Message; }
        finally
        {
            // No cerramos Outlook: si estaba abierto es la instancia del usuario.
            try { if (item != null) Marshal.FinalReleaseComObject(item); } catch { }
            try { if (ns != null) Marshal.FinalReleaseComObject(ns); } catch { }
            try { if (app != null) Marshal.FinalReleaseComObject(app); } catch { }
            item = null; ns = null; app = null;
            GC.Collect(); GC.WaitForPendingFinalizers();
        }
    }

    // ---------- Deteccion del archivo bajo el cursor ----------
    static Hit DetectUnderCursor(Point pt)
    {
        AutomationElement el = AutomationElement.FromPoint(new System.Windows.Point(pt.X, pt.Y));
        if (el == null) return null;

        AutomationElement item = el;
        TreeWalker walker = TreeWalker.ControlViewWalker;
        int guard = 0;
        while (item != null && guard++ < 14)
        {
            ControlType ct = null;
            try { ct = item.Current.ControlType; } catch { }
            if (ct == ControlType.ListItem || ct == ControlType.DataItem) break;
            item = walker.GetParent(item);
        }
        if (item == null) return null;
        ControlType fct = null;
        try { fct = item.Current.ControlType; } catch { return null; }
        if (fct != ControlType.ListItem && fct != ControlType.DataItem) return null;

        string name;
        try { name = item.Current.Name; } catch { return null; }
        if (string.IsNullOrEmpty(name)) return null;

        IntPtr hwnd = GetAncestor(WindowFromPoint(new POINT(pt.X, pt.Y)), GA_ROOT);
        string folder = GetExplorerFolderPath(hwnd);
        if (folder == null) folder = GetDialogFolderPath(hwnd); // dialogos Abrir/Guardar
        if (folder == null) return null;

        string full = Path.Combine(folder, name);
        string found = null;
        if (File.Exists(full)) found = full;
        else
        {
            foreach (string ext in AllExts)
            {
                if (File.Exists(full + ext)) { found = full + ext; break; }
            }
        }
        if (found == null) return null;
        int kind = KindOf(found);
        if (kind == KIND_NONE) return null;
        Hit h = new Hit(); h.Path = found; h.Kind = kind;
        return h;
    }

    static string GetExplorerFolderPath(IntPtr hwnd)
    {
        try
        {
            Type shellType = Type.GetTypeFromProgID("Shell.Application");
            if (shellType == null) return null;
            dynamic shell = Activator.CreateInstance(shellType);
            dynamic windows = shell.Windows();
            int count = (int)windows.Count;
            for (int i = 0; i < count; i++)
            {
                dynamic w = windows.Item(i);
                if (w == null) continue;
                IntPtr h;
                try { h = new IntPtr((long)w.HWND); } catch { continue; }
                if (h != hwnd) continue;
                try
                {
                    dynamic self = w.Document.Folder.Self;
                    string path = (string)self.Path;
                    if (!string.IsNullOrEmpty(path) && Directory.Exists(path)) return path;
                }
                catch { }
                return null;
            }
        }
        catch { }
        return null;
    }

    // Carpeta actual en un dialogo "Abrir/Guardar" (no es una ventana del Explorador).
    // Se obtiene de la barra de direcciones, cuyo Name es "Direccion: C:\ruta...".
    // Buscamos el patron de unidad (C:\) para no depender del idioma.
    static string GetDialogFolderPath(IntPtr hwnd)
    {
        try
        {
            AutomationElement win = AutomationElement.FromHandle(hwnd);
            if (win == null) return null;
            // La barra de direcciones puede ser ToolBar o Pane segun la carpeta.
            OrCondition cond = new OrCondition(
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ToolBar),
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane));
            AutomationElementCollection els = win.FindAll(TreeScope.Descendants, cond);
            foreach (AutomationElement e in els)
            {
                string nm = null;
                try { nm = e.Current.Name; } catch { }
                if (string.IsNullOrEmpty(nm)) continue;
                if (nm.IndexOf("Direc", StringComparison.OrdinalIgnoreCase) < 0
                    && IndexOfPath(nm) < 0) continue; // acelera: solo barras de direcciones
                string f = ExtractFolder(nm);
                if (f != null) return f;
            }
        }
        catch { }
        return null;
    }

    // De un texto como "Direccion: C:\ruta" o "Direccion: Documentos" saca la carpeta real.
    static string ExtractFolder(string nm)
    {
        // 1) Ruta directa tipo C:\ o UNC \\
        int idx = IndexOfPath(nm);
        if (idx >= 0)
        {
            string p = nm.Substring(idx).Trim();
            if (p.Length > 0 && Directory.Exists(p)) return p;
        }
        // 2) Nombre de carpeta especial tras "...: " (Documentos, Descargas, etc.)
        int c = nm.IndexOf(": ");
        if (c >= 0)
        {
            string disp = nm.Substring(c + 2).Trim();
            if (!string.IsNullOrEmpty(disp))
            {
                string mapped = MapKnownFolder(disp);
                if (mapped != null) return mapped;
            }
        }
        return null;
    }

    // Traduce el nombre mostrado de una carpeta especial a su ruta real (segun el idioma del sistema).
    static string MapKnownFolder(string disp)
    {
        try
        {
            string profile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string[] candidates = new string[]
            {
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                Environment.GetFolderPath(Environment.SpecialFolder.MyPictures),
                Environment.GetFolderPath(Environment.SpecialFolder.MyMusic),
                Environment.GetFolderPath(Environment.SpecialFolder.MyVideos),
                Path.Combine(profile, "Downloads"),
                profile
            };
            foreach (string path in candidates)
            {
                if (string.IsNullOrEmpty(path) || !Directory.Exists(path)) continue;
                string dn = ShellDisplayName(path);
                if (dn != null && string.Equals(dn, disp, StringComparison.OrdinalIgnoreCase))
                    return path;
            }
        }
        catch { }
        return null;
    }

    static string ShellDisplayName(string path)
    {
        try
        {
            Type t = Type.GetTypeFromProgID("Shell.Application");
            if (t == null) return null;
            dynamic shell = Activator.CreateInstance(t);
            dynamic ns = shell.NameSpace(path);
            if (ns != null) { dynamic self = ns.Self; return (string)self.Name; }
        }
        catch { }
        return null;
    }

    // Devuelve el indice donde empieza una ruta tipo "C:\" o UNC "\\", o -1.
    static int IndexOfPath(string s)
    {
        for (int i = 0; i + 2 < s.Length; i++)
        {
            char c = s[i];
            if (((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z')) && s[i + 1] == ':' && s[i + 2] == '\\')
                return i;
            if (c == '\\' && s[i + 1] == '\\') return i;
        }
        return -1;
    }

    // ---------- Ventana de vista previa (sin robar foco, con scroll) ----------
    class PreviewForm : Form
    {
        public Panel Host;
        Label loading;
        Panel zoomBar;
        Label zoomLbl;
        TextBox textBox;          // modo texto (txt, csv, zip, correos)
        Panel actionBar;          // botones Abrir / Fijar
        Button pinBtn;
        List<double> ratios = new List<double>();
        float zoom = 1f;
        const float ZoomMin = 0.6f, ZoomMax = 2.6f;
        const int Gap = 6;
        public const int FormWidth = 560;
        int BaseContentW;

        const int MODE_PAGES = 0, MODE_TEXT = 1;
        int mode = MODE_PAGES;
        public bool Pinned;
        public string CurrentFile;   // archivo original (para el boton Abrir)
        bool dragging; Point dragStart;

        [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);
        const int EM_LINESCROLL = 0x00B6;

        public PreviewForm()
        {
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.Manual;
            ShowInTaskbar = false;
            TopMost = true;
            BackColor = Color.FromArgb(90, 90, 90);
            Padding = new Padding(1);

            Host = new Panel();
            Host.Dock = DockStyle.Fill;
            Host.AutoScroll = true;
            Host.BackColor = Color.FromArgb(140, 140, 140);
            Controls.Add(Host);

            loading = new Label();
            loading.AutoSize = false;
            loading.Dock = DockStyle.Fill;
            loading.TextAlign = ContentAlignment.MiddleCenter;
            loading.BackColor = Color.White;
            loading.ForeColor = Color.FromArgb(70, 70, 70);
            loading.Font = new Font("Segoe UI", 10f);
            loading.Visible = false;
            Controls.Add(loading);

            // Barra de zoom (botones - / % / +) fija arriba a la derecha
            zoomBar = new Panel();
            zoomBar.Size = new Size(108, 28);
            zoomBar.BackColor = Color.FromArgb(245, 245, 245);
            zoomBar.BorderStyle = BorderStyle.FixedSingle;
            zoomBar.Visible = false;
            Button bMinus = MakeZoomButton("−", 0);
            zoomLbl = new Label();
            zoomLbl.Text = "100%";
            zoomLbl.TextAlign = ContentAlignment.MiddleCenter;
            zoomLbl.Font = new Font("Segoe UI", 8f);
            zoomLbl.ForeColor = Color.FromArgb(60, 60, 60);
            zoomLbl.SetBounds(28, 0, 48, 26);
            Button bPlus = MakeZoomButton("+", 76);
            bMinus.Click += delegate { ZoomBy(-1); };
            bPlus.Click += delegate { ZoomBy(1); };
            zoomBar.Controls.Add(bMinus);
            zoomBar.Controls.Add(zoomLbl);
            zoomBar.Controls.Add(bPlus);
            Controls.Add(zoomBar);
            zoomBar.BringToFront();

            // Cuadro de texto (para .txt, .csv, .zip, correos)
            textBox = new TextBox();
            textBox.Multiline = true;
            textBox.ReadOnly = true;
            textBox.ScrollBars = ScrollBars.Both;
            textBox.WordWrap = false;
            textBox.Dock = DockStyle.Fill;
            textBox.BackColor = Color.White;
            textBox.ForeColor = Color.FromArgb(30, 30, 30);
            textBox.Font = new Font("Consolas", 9.5f);
            textBox.BorderStyle = BorderStyle.None;
            textBox.Visible = false;
            Controls.Add(textBox);

            // Barra de acciones: Abrir / Fijar (arriba a la izquierda)
            actionBar = new Panel();
            actionBar.Size = new Size(150, 28);
            actionBar.BackColor = Color.FromArgb(245, 245, 245);
            actionBar.BorderStyle = BorderStyle.FixedSingle;
            actionBar.Visible = false;
            Button openBtn = new Button();
            openBtn.Text = "Abrir";
            openBtn.SetBounds(0, 0, 72, 26);
            openBtn.FlatStyle = FlatStyle.Flat;
            openBtn.FlatAppearance.BorderSize = 0;
            openBtn.BackColor = Color.White;
            openBtn.Font = new Font("Segoe UI", 8.5f);
            openBtn.TabStop = false;
            openBtn.Click += delegate { OpenCurrentFile(); };
            pinBtn = new Button();
            pinBtn.Text = "Fijar";
            pinBtn.SetBounds(74, 0, 74, 26);
            pinBtn.FlatStyle = FlatStyle.Flat;
            pinBtn.FlatAppearance.BorderSize = 0;
            pinBtn.BackColor = Color.White;
            pinBtn.Font = new Font("Segoe UI", 8.5f);
            pinBtn.TabStop = false;
            pinBtn.Click += delegate { TogglePin(); };
            actionBar.Controls.Add(openBtn);
            actionBar.Controls.Add(pinBtn);
            // arrastrar la ventana tomando la barra de acciones (util al fijar)
            actionBar.MouseDown += delegate(object s, MouseEventArgs e) { dragging = true; dragStart = e.Location; };
            actionBar.MouseMove += delegate(object s, MouseEventArgs e) { if (dragging) { Left += e.X - dragStart.X; Top += e.Y - dragStart.Y; } };
            actionBar.MouseUp += delegate(object s, MouseEventArgs e) { dragging = false; };
            Controls.Add(actionBar);
            actionBar.BringToFront();

            BaseContentW = FormWidth - 2 - SystemInformation.VerticalScrollBarWidth - 2;
        }

        void OpenCurrentFile()
        {
            try
            {
                if (!string.IsNullOrEmpty(CurrentFile) && File.Exists(CurrentFile))
                {
                    System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo(CurrentFile);
                    psi.UseShellExecute = true;
                    System.Diagnostics.Process.Start(psi);
                    Pinned = false;
                    HideIt();
                }
            }
            catch { }
        }

        void TogglePin()
        {
            Pinned = !Pinned;
            if (pinBtn != null) pinBtn.Text = Pinned ? "Fijado" : "Fijar";
            BackColor = Pinned ? Color.FromArgb(29, 78, 216) : Color.FromArgb(90, 90, 90);
        }

        void ShowActionBar()
        {
            if (pinBtn != null) pinBtn.Text = Pinned ? "Fijado" : "Fijar";
            // En modo texto el contenido va a la izquierda: barra a la derecha.
            // En modo paginas la barra de zoom esta a la derecha: barra a la izquierda.
            if (mode == MODE_TEXT)
                actionBar.Location = new Point(ClientSize.Width - actionBar.Width - 3 - SystemInformation.VerticalScrollBarWidth, 3);
            else
                actionBar.Location = new Point(3, 3);
            actionBar.Visible = true;
            actionBar.BringToFront();
        }

        void ApplyTheme()
        {
            bool dark = IsDarkTheme();
            if (dark)
            {
                if (!Pinned) BackColor = Color.FromArgb(50, 50, 50);
                Host.BackColor = Color.FromArgb(60, 60, 60);
                textBox.BackColor = Color.FromArgb(32, 32, 32);
                textBox.ForeColor = Color.FromArgb(230, 230, 230);
                loading.BackColor = Color.FromArgb(40, 40, 40);
                loading.ForeColor = Color.FromArgb(220, 220, 220);
            }
            else
            {
                if (!Pinned) BackColor = Color.FromArgb(90, 90, 90);
                Host.BackColor = Color.FromArgb(140, 140, 140);
                textBox.BackColor = Color.White;
                textBox.ForeColor = Color.FromArgb(30, 30, 30);
                loading.BackColor = Color.White;
                loading.ForeColor = Color.FromArgb(70, 70, 70);
            }
        }

        Button MakeZoomButton(string text, int left)
        {
            Button b = new Button();
            b.Text = text;
            b.SetBounds(left, 0, 32, 26);
            b.FlatStyle = FlatStyle.Flat;
            b.FlatAppearance.BorderSize = 0;
            b.BackColor = Color.White;
            b.Font = new Font("Segoe UI", 11f, FontStyle.Bold);
            b.TabStop = false;
            return b;
        }

        protected override bool ShowWithoutActivation { get { return true; } }
        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= 0x08000000; // WS_EX_NOACTIVATE
                cp.ExStyle |= 0x00000080; // WS_EX_TOOLWINDOW
                return cp;
            }
        }

        void PositionZoomBar()
        {
            zoomBar.Location = new Point(ClientSize.Width - zoomBar.Width - 3, 3);
            zoomBar.BringToFront();
        }

        void ClearHost()
        {
            Host.SuspendLayout();
            List<Control> old = new List<Control>();
            foreach (Control c in Host.Controls) old.Add(c);
            Host.Controls.Clear();
            foreach (Control c in old)
            {
                PictureBox pb = c as PictureBox;
                if (pb != null && pb.Image != null) { pb.Image.Dispose(); pb.Image = null; }
                c.Dispose();
            }
            Host.AutoScrollPosition = new Point(0, 0);
            Host.ResumeLayout();
        }

        float SavedZoom()
        {
            float z;
            if (float.TryParse(CfgGet("zoom", "1"), System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out z))
            {
                if (z < ZoomMin) z = ZoomMin;
                if (z > ZoomMax) z = ZoomMax;
                return z;
            }
            return 1f;
        }

        public void ShowLoading(string text, Point near)
        {
            mode = MODE_PAGES;
            textBox.Visible = false;
            Host.Visible = true;
            ApplyTheme();
            ClearHost();
            ratios.Clear();
            zoomBar.Visible = false;
            actionBar.Visible = false;
            loading.Text = text;
            loading.Visible = true;
            loading.BringToFront();
            SizeAndPlace(360, near, true);
            if (!Visible) Show();
        }

        public void ShowText(string text, Point near)
        {
            mode = MODE_TEXT;
            loading.Visible = false;
            ClearHost();
            ratios.Clear();
            zoomBar.Visible = false;
            Host.Visible = false;
            ApplyTheme();
            textBox.Text = text;
            textBox.Visible = true;
            textBox.BringToFront();
            try { textBox.Select(0, 0); } catch { }
            SizeAndPlaceText(near);
            ShowActionBar();
            if (!Visible) Show();
        }

        void SizeAndPlaceText(Point near)
        {
            Screen scr = Screen.FromPoint(near);
            int winW = Math.Min(640, scr.WorkingArea.Width - 8);
            int winH = (int)(scr.WorkingArea.Height * 0.70);
            if (winH < 220) winH = 220;
            int x = near.X + 22, y = near.Y + 18;
            if (x + winW + 6 > scr.WorkingArea.Right) x = scr.WorkingArea.Right - winW - 6;
            if (x < scr.WorkingArea.Left) x = scr.WorkingArea.Left + 4;
            if (y + winH + 6 > scr.WorkingArea.Bottom) y = scr.WorkingArea.Bottom - winH - 6;
            if (y < scr.WorkingArea.Top) y = scr.WorkingArea.Top + 4;
            Bounds = new Rectangle(x, y, winW, winH);
            PositionZoomBar();
        }

        // PDF: crea los "huecos" (uno por pagina) con su altura, de una vez.
        public void BuildPdfSlots(double[] rs, Point near)
        {
            mode = MODE_PAGES;
            textBox.Visible = false;
            Host.Visible = true;
            ApplyTheme();
            loading.Visible = false;
            ClearHost();
            ratios.Clear();
            zoom = SavedZoom();
            Host.SuspendLayout();
            for (int i = 0; i < rs.Length; i++)
            {
                ratios.Add(rs[i]);
                PictureBox pb = new PictureBox();
                pb.BackColor = Color.White;
                pb.SizeMode = PictureBoxSizeMode.Zoom;
                Host.Controls.Add(pb);
            }
            Host.ResumeLayout();
            Relayout(near, true);
            zoomBar.Visible = true;
            ShowActionBar();
            UpdateZoomLabel();
            if (!Visible) Show();
        }

        public void SetPdfPage(int index, Bitmap bmp)
        {
            if (index < 0 || index >= Host.Controls.Count) { if (bmp != null) bmp.Dispose(); return; }
            PictureBox pb = Host.Controls[index] as PictureBox;
            if (pb == null) { if (bmp != null) bmp.Dispose(); return; }
            Image old = pb.Image;
            pb.Image = bmp;
            if (old != null) old.Dispose();
        }

        public void ShowSingleImage(Bitmap bmp, Point near)
        {
            mode = MODE_PAGES;
            textBox.Visible = false;
            Host.Visible = true;
            ApplyTheme();
            loading.Visible = false;
            ClearHost();
            ratios.Clear();
            zoom = SavedZoom();
            ratios.Add((double)bmp.Height / bmp.Width);
            Host.SuspendLayout();
            PictureBox pb = new PictureBox();
            pb.BackColor = Color.White;
            pb.SizeMode = PictureBoxSizeMode.Zoom;
            pb.Image = bmp;
            Host.Controls.Add(pb);
            Host.ResumeLayout();
            Relayout(near, true);
            zoomBar.Visible = true;
            ShowActionBar();
            UpdateZoomLabel();
            if (!Visible) Show();
        }

        // Recoloca las paginas segun el zoom y ajusta/posiciona la ventana.
        void Relayout(Point near, bool reposition)
        {
            if (Host.Controls.Count == 0) return;
            int w = (int)Math.Round(BaseContentW * zoom);
            if (w < 60) w = 60;

            int oldDisp = Host.DisplayRectangle.Height;
            double frac = (oldDisp > 0) ? ((double)(-Host.AutoScrollPosition.Y) / oldDisp) : 0.0;

            Host.SuspendLayout();
            int y = 0, idx = 0;
            foreach (Control c in Host.Controls)
            {
                PictureBox pb = c as PictureBox;
                if (pb == null) continue;
                double r = (idx < ratios.Count) ? ratios[idx] : 1.3;
                int h = (int)Math.Round(w * r);
                if (h < 10) h = 10;
                pb.SetBounds(1, y, w, h);
                y += h + Gap;
                idx++;
            }
            int contentH = y;
            Host.ResumeLayout();

            SizeAndPlaceForContent(w, contentH, near, reposition);

            int newDisp = Host.DisplayRectangle.Height;
            Host.AutoScrollPosition = new Point(0, (int)(frac * newDisp));
        }

        void SizeAndPlaceForContent(int contentW, int contentHeight, Point near, bool reposition)
        {
            Screen scr = reposition ? Screen.FromPoint(near) : Screen.FromRectangle(Bounds);
            int sbW = SystemInformation.VerticalScrollBarWidth;
            int maxW = scr.WorkingArea.Width - 8;
            int maxH = (int)(scr.WorkingArea.Height * 0.95);
            int winW = Math.Min(contentW + sbW + 4, maxW);
            int winH = Math.Min(contentHeight + 2, maxH);
            if (winW < 120) winW = 120;
            if (winH < 80) winH = 80;

            int x, y;
            if (reposition) { x = near.X + 22; y = near.Y + 18; }
            else { x = Left; y = Top; }
            if (x + winW + 6 > scr.WorkingArea.Right) x = scr.WorkingArea.Right - winW - 6;
            if (x < scr.WorkingArea.Left) x = scr.WorkingArea.Left + 4;
            if (y + winH + 6 > scr.WorkingArea.Bottom) y = scr.WorkingArea.Bottom - winH - 6;
            if (y < scr.WorkingArea.Top) y = scr.WorkingArea.Top + 4;
            Bounds = new Rectangle(x, y, winW, winH);
            PositionZoomBar();
        }

        void SizeAndPlace(int contentHeight, Point near, bool reposition)
        {
            SizeAndPlaceForContent(BaseContentW, contentHeight, near, reposition);
        }

        void UpdateZoomLabel()
        {
            if (zoomLbl != null) zoomLbl.Text = ((int)Math.Round(zoom * 100)) + "%";
        }

        public void ZoomBy(int steps)
        {
            if (ratios.Count == 0) return;
            float nz = zoom * (float)Math.Pow(1.15, steps);
            if (nz < ZoomMin) nz = ZoomMin;
            if (nz > ZoomMax) nz = ZoomMax;
            if (Math.Abs(nz - zoom) < 0.001f) return;
            zoom = nz;
            Relayout(Point.Empty, false);
            UpdateZoomLabel();
            CfgSet("zoom", zoom.ToString(System.Globalization.CultureInfo.InvariantCulture)); // recordar
        }

        public void ZoomStep(int wheelDelta)
        {
            int steps = wheelDelta / 120;
            if (steps == 0) steps = (wheelDelta > 0) ? 1 : -1;
            ZoomBy(steps);
        }

        public void HideIt()
        {
            if (Visible) Hide();
            ClearHost();
            ratios.Clear();
            loading.Visible = false;
            zoomBar.Visible = false;
            if (textBox != null) { textBox.Visible = false; textBox.Clear(); }
            if (actionBar != null) actionBar.Visible = false;
            mode = MODE_PAGES;
            if (Pinned) { Pinned = false; if (pinBtn != null) pinBtn.Text = "Fijar"; BackColor = Color.FromArgb(90, 90, 90); }
        }

        public void ScrollBy(int dy)
        {
            if (!Visible) return;
            if (mode == MODE_TEXT)
            {
                int lines = dy / 30;
                if (lines == 0) lines = (dy > 0) ? 1 : -1;
                try { SendMessage(textBox.Handle, EM_LINESCROLL, IntPtr.Zero, (IntPtr)lines); } catch { }
                return;
            }
            int x = -Host.AutoScrollPosition.X;
            int y = -Host.AutoScrollPosition.Y;
            Host.AutoScrollPosition = new Point(x, y + dy);
        }

        public void ScrollByH(int dx)
        {
            if (!Visible) return;
            if (mode == MODE_TEXT)
            {
                int chars = dx / 16;
                if (chars == 0) chars = (dx > 0) ? 1 : -1;
                try { SendMessage(textBox.Handle, EM_LINESCROLL, (IntPtr)chars, IntPtr.Zero); } catch { }
                return;
            }
            int x = -Host.AutoScrollPosition.X;
            int y = -Host.AutoScrollPosition.Y;
            Host.AutoScrollPosition = new Point(x + dx, y);
        }
    }

    // ---------- Contexto de bandeja + logica ----------
    class TrayContext : ApplicationContext
    {
        NotifyIcon tray;
        System.Windows.Forms.Timer timer;
        PreviewForm preview;
        Point lastPos;
        DateTime restSince;
        bool handledThisRest;
        string shownPath;
        bool paused;
        volatile int gen;
        IntPtr mouseHook = IntPtr.Zero;
        HookProc hookProc; // mantener viva la referencia

        int DwellMs = 350;
        const uint RenderWidth = 1000;
        const int MaxPages = 40;

        public TrayContext()
        {
            preview = new PreviewForm();
            IntPtr force = preview.Handle; // crea el handle para poder Invoke

            tray = new NotifyIcon();
            try { tray.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); }
            catch { tray.Icon = SystemIcons.Application; }
            if (tray.Icon == null) tray.Icon = SystemIcons.Application;
            tray.Visible = true;
            tray.Text = "Vista Previa (hover): PDF, imagenes, Word, Excel";
            ContextMenuStrip menu = new ContextMenuStrip();
            ToolStripMenuItem pauseItem = new ToolStripMenuItem("Pausar");
            pauseItem.Click += delegate
            {
                paused = !paused;
                pauseItem.Text = paused ? "Reanudar" : "Pausar";
                if (paused) { gen++; preview.HideIt(); shownPath = null; }
            };
            menu.Items.Add(pauseItem);

            ToolStripMenuItem autoItem = new ToolStripMenuItem("Iniciar con Windows");
            autoItem.Checked = IsAutoStartEnabled();
            autoItem.Click += delegate
            {
                bool enable = !IsAutoStartEnabled();
                bool ok = SetAutoStart(enable);
                autoItem.Checked = IsAutoStartEnabled();
                if (ok)
                    tray.ShowBalloonTip(3000, "Vista Previa",
                        enable
                          ? "Se iniciara automaticamente cada vez que enciendas la PC."
                          : "Ya NO se iniciara solo con Windows.",
                        ToolTipIcon.Info);
                else
                    tray.ShowBalloonTip(3000, "Vista Previa",
                        "No se pudo cambiar el inicio automatico.", ToolTipIcon.Warning);
            };
            menu.Items.Add(autoItem);

            // Tiempo de aparicion (dwell) configurable
            int dw; if (!int.TryParse(CfgGet("dwell", "350"), out dw)) dw = 350; DwellMs = dw;
            ToolStripMenuItem speedItem = new ToolStripMenuItem("Tiempo de aparicion");
            ToolStripMenuItem sFast = new ToolStripMenuItem("Rapido (0.15 s)");
            ToolStripMenuItem sNormal = new ToolStripMenuItem("Normal (0.35 s)");
            ToolStripMenuItem sSlow = new ToolStripMenuItem("Lento (0.6 s)");
            sFast.Checked = (DwellMs <= 150);
            sNormal.Checked = (DwellMs > 150 && DwellMs <= 350);
            sSlow.Checked = (DwellMs > 350);
            sFast.Click += delegate { DwellMs = 150; CfgSet("dwell", "150"); sFast.Checked = true; sNormal.Checked = false; sSlow.Checked = false; };
            sNormal.Click += delegate { DwellMs = 350; CfgSet("dwell", "350"); sFast.Checked = false; sNormal.Checked = true; sSlow.Checked = false; };
            sSlow.Click += delegate { DwellMs = 600; CfgSet("dwell", "600"); sFast.Checked = false; sNormal.Checked = false; sSlow.Checked = true; };
            speedItem.DropDownItems.Add(sFast);
            speedItem.DropDownItems.Add(sNormal);
            speedItem.DropDownItems.Add(sSlow);
            menu.Items.Add(speedItem);

            menu.Items.Add(new ToolStripSeparator());
            ToolStripMenuItem exitItem = new ToolStripMenuItem("Salir");
            exitItem.Click += delegate { ExitApp(); };
            menu.Items.Add(exitItem);
            tray.ContextMenuStrip = menu;
            tray.ShowBalloonTip(5000, "Vista Previa de documentos activada",
                "El servicio ya esta corriendo en segundo plano. Pasa el cursor sobre un PDF, imagen, Word o Excel para ver la vista previa.\r\nPara cerrarlo cuando gustes: clic derecho en este icono (lupa azul) > Salir.",
                ToolTipIcon.Info);

            // Hook de la rueda del mouse (para scroll sin robar foco)
            hookProc = new HookProc(WheelHook);
            try
            {
                using (System.Diagnostics.Process p = System.Diagnostics.Process.GetCurrentProcess())
                using (System.Diagnostics.ProcessModule m = p.MainModule)
                    mouseHook = SetWindowsHookEx(WH_MOUSE_LL, hookProc, GetModuleHandle(m.ModuleName), 0);
            }
            catch { mouseHook = IntPtr.Zero; }

            lastPos = Cursor.Position;
            restSince = DateTime.Now;

            timer = new System.Windows.Forms.Timer();
            timer.Interval = 180;
            timer.Tick += Tick;
            timer.Start();
        }

        IntPtr WheelHook(int nCode, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                if (nCode >= 0 && (int)wParam == WM_MOUSEWHEEL)
                {
                    MSLLHOOKSTRUCT data = (MSLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(MSLLHOOKSTRUCT));
                    if (preview.Visible && preview.Bounds.Contains(data.pt.X, data.pt.Y))
                    {
                        short delta = (short)((data.mouseData >> 16) & 0xffff);
                        bool ctrl = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
                        bool shift = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
                        if (ctrl)
                        {
                            preview.ZoomStep(delta); // Ctrl + rueda = zoom
                        }
                        else if (shift)
                        {
                            preview.ScrollByH(-(delta / 120) * 90); // Shift + rueda = horizontal
                        }
                        else
                        {
                            preview.ScrollBy(-(delta / 120) * 90); // rueda = vertical
                        }
                        return (IntPtr)1; // consumir: que el Explorador no haga scroll
                    }
                }
            }
            catch { }
            return CallNextHookEx(IntPtr.Zero, nCode, wParam, lParam);
        }

        void Tick(object sender, EventArgs e)
        {
            if (paused) return;
            if (preview.Visible && preview.Pinned) return; // fijada: no cambiarla
            try
            {
                Point p = Cursor.Position;

                // Si el cursor esta sobre la vista previa, mantenerla (para hacer scroll)
                if (preview.Visible && preview.Bounds.Contains(p)) return;

                int dx = p.X - lastPos.X, dy = p.Y - lastPos.Y;
                if (dx * dx + dy * dy > 49) // se movio > 7px
                {
                    lastPos = p;
                    restSince = DateTime.Now;
                    handledThisRest = false;
                    if (shownPath != null) { gen++; preview.HideIt(); shownPath = null; }
                    return;
                }
                if (handledThisRest) return;
                if ((DateTime.Now - restSince).TotalMilliseconds < DwellMs) return;
                handledThisRest = true;

                Hit hit = DetectUnderCursor(p);
                if (hit == null) { if (shownPath != null) { gen++; preview.HideIt(); shownPath = null; } return; }
                if (hit.Path == shownPath) return;

                StartPreview(hit, p);
            }
            catch { /* nunca crashear por un tick */ }
        }

        void StartPreview(Hit hit, Point at)
        {
            int myGen = ++gen;
            shownPath = hit.Path;
            preview.CurrentFile = hit.Path; // para el boton "Abrir"

            if (hit.Kind == KIND_IMAGE)
            {
                try
                {
                    Bitmap bmp = LoadImage(hit.Path);
                    if (gen == myGen) preview.ShowSingleImage(bmp, at);
                    else bmp.Dispose();
                }
                catch { if (gen == myGen) preview.ShowLoading("No se pudo abrir la imagen", at); }
            }
            else if (hit.Kind == KIND_PDF)
            {
                StartPdf(hit.Path, myGen, at);
            }
            else if (hit.Kind == KIND_OFFICE)
            {
                preview.ShowLoading("Generando vista previa...", at);
                Thread t = new Thread(delegate()
                {
                    try
                    {
                        string pdf = OfficeToPdf(hit.Path);
                        if (gen != myGen) return;
                        preview.BeginInvoke((MethodInvoker)delegate
                        {
                            if (gen == myGen) StartPdf(pdf, myGen, at);
                        });
                    }
                    catch
                    {
                        if (gen != myGen) return;
                        try { preview.BeginInvoke((MethodInvoker)delegate { if (gen == myGen) preview.ShowLoading("No se pudo generar la vista previa", at); }); }
                        catch { }
                    }
                });
                t.IsBackground = true;
                t.SetApartmentState(ApartmentState.STA); // Office COM prefiere STA
                t.Start();
            }
            else if (hit.Kind == KIND_TEXT || hit.Kind == KIND_ZIP || hit.Kind == KIND_EMAIL)
            {
                int kind = hit.Kind; string path = hit.Path;
                preview.ShowLoading("Cargando vista previa...", at);
                Thread t = new Thread(delegate()
                {
                    string text;
                    try
                    {
                        if (kind == KIND_TEXT) text = ReadTextPreview(path);
                        else if (kind == KIND_ZIP) text = ReadZipListing(path);
                        else text = ReadEmailPreview(path);
                    }
                    catch (Exception ex) { text = "No se pudo mostrar la vista previa:\r\n" + ex.Message; }
                    if (gen != myGen) return;
                    try { preview.BeginInvoke((MethodInvoker)delegate { if (gen == myGen) preview.ShowText(text, at); }); }
                    catch { }
                });
                t.IsBackground = true;
                t.SetApartmentState(ApartmentState.STA); // .msg usa Outlook COM (STA)
                t.Start();
            }
        }

        static Bitmap LoadImage(string path)
        {
            byte[] data = File.ReadAllBytes(path);
            try
            {
                using (MemoryStream ms = new MemoryStream(data))
                    return new Bitmap(ms);
            }
            catch
            {
                // GDI no soporta HEIC/WebP: usar el decodificador WIC de Windows
                return LoadImageWic(data);
            }
        }

        static Bitmap LoadImageWic(byte[] data)
        {
            using (MemoryStream ms = new MemoryStream(data))
            {
                System.Windows.Media.Imaging.BitmapDecoder dec = System.Windows.Media.Imaging.BitmapDecoder.Create(
                    ms,
                    System.Windows.Media.Imaging.BitmapCreateOptions.PreservePixelFormat,
                    System.Windows.Media.Imaging.BitmapCacheOption.OnLoad);
                System.Windows.Media.Imaging.BitmapFrame frame = dec.Frames[0];
                System.Windows.Media.Imaging.FormatConvertedBitmap conv =
                    new System.Windows.Media.Imaging.FormatConvertedBitmap(frame, System.Windows.Media.PixelFormats.Bgra32, null, 0);
                int w = conv.PixelWidth, h = conv.PixelHeight;
                int stride = w * 4;
                byte[] px = new byte[h * stride];
                conv.CopyPixels(px, stride, 0);
                Bitmap bmp = new Bitmap(w, h, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                System.Drawing.Imaging.BitmapData bd = bmp.LockBits(
                    new Rectangle(0, 0, w, h),
                    System.Drawing.Imaging.ImageLockMode.WriteOnly,
                    System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                try { Marshal.Copy(px, 0, bd.Scan0, px.Length); }
                finally { bmp.UnlockBits(bd); }
                return bmp;
            }
        }

        void StartPdf(string pdfPath, int myGen, Point at)
        {
            Thread t = new Thread(delegate() { RenderPdfJob(pdfPath, myGen, at); });
            t.IsBackground = true;
            t.SetApartmentState(ApartmentState.MTA);
            t.Start();
        }

        void RenderPdfJob(string pdfPath, int myGen, Point at)
        {
            try
            {
                // Leer el PDF a memoria y soltar el archivo de inmediato,
                // para NO bloquearlo (asi el usuario puede borrarlo/moverlo).
                byte[] pdfBytes = File.ReadAllBytes(pdfPath);
                InMemoryRandomAccessStream mem = new InMemoryRandomAccessStream();
                DataWriter writer = new DataWriter(mem);
                writer.WriteBytes(pdfBytes);
                AwaitOp(writer.StoreAsync());
                writer.DetachStream();
                mem.Seek(0);
                PdfDocument doc = AwaitOp(PdfDocument.LoadFromStreamAsync(mem));
                int count = (int)doc.PageCount;
                if (count > MaxPages) count = MaxPages;
                if (count < 1) count = 1;

                // Pass 1: medir proporciones de cada pagina
                double[] ratios = new double[count];
                for (int i = 0; i < count; i++)
                {
                    using (PdfPage pg = doc.GetPage((uint)i))
                    {
                        Windows.Foundation.Size sz = pg.Size;
                        double r = (sz.Width > 0) ? (sz.Height / sz.Width) : 1.3;
                        ratios[i] = r;
                    }
                }
                if (gen != myGen) return;
                preview.BeginInvoke((MethodInvoker)delegate { if (gen == myGen) preview.BuildPdfSlots(ratios, at); });

                // Pass 2: renderizar cada pagina y colocarla
                for (int i = 0; i < count; i++)
                {
                    if (gen != myGen) return;
                    byte[] bytes;
                    using (PdfPage pg = doc.GetPage((uint)i))
                        bytes = RenderPageBytes(pg, RenderWidth);
                    if (gen != myGen) return;
                    Bitmap bmp = new Bitmap(new MemoryStream(bytes));
                    int idx = i;
                    preview.BeginInvoke((MethodInvoker)delegate
                    {
                        if (gen == myGen) preview.SetPdfPage(idx, bmp);
                        else bmp.Dispose();
                    });
                }
            }
            catch
            {
                if (gen != myGen) return;
                try { preview.BeginInvoke((MethodInvoker)delegate { if (gen == myGen) preview.ShowLoading("No se pudo mostrar el PDF", at); }); }
                catch { }
            }
        }

        void ExitApp()
        {
            try { timer.Stop(); } catch { }
            try { if (mouseHook != IntPtr.Zero) UnhookWindowsHookEx(mouseHook); } catch { }
            try { tray.Visible = false; tray.Dispose(); } catch { }
            try { preview.Dispose(); } catch { }
            ExitThread();
        }

        // ---------- Inicio con Windows (acceso directo en carpeta Inicio) ----------
        static string StartupLnkPath()
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Startup),
                "Vista Previa.lnk");
        }

        static bool IsAutoStartEnabled()
        {
            try { return File.Exists(StartupLnkPath()); }
            catch { return false; }
        }

        static bool SetAutoStart(bool enable)
        {
            try
            {
                string lnk = StartupLnkPath();
                if (enable)
                {
                    string exe = Application.ExecutablePath;
                    Type t = Type.GetTypeFromProgID("WScript.Shell");
                    dynamic shell = Activator.CreateInstance(t);
                    dynamic sc = shell.CreateShortcut(lnk);
                    sc.TargetPath = exe;
                    sc.WorkingDirectory = Path.GetDirectoryName(exe);
                    sc.IconLocation = exe + ",0";
                    sc.Description = "Vista Previa de archivos";
                    sc.Save();
                }
                else
                {
                    if (File.Exists(lnk)) File.Delete(lnk);
                }
                return true;
            }
            catch { return false; }
        }
    }

    [STAThread]
    static void Main()
    {
        // Version simple: solo corre la vista previa. NO se auto-copia ni toca
        // el registro (eso activaba el antivirus). La instalacion y el arranque
        // con Windows se hacen por fuera (acceso directo en la carpeta Inicio).
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new TrayContext());
    }
}
