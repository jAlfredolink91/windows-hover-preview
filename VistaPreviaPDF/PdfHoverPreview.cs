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
    const int KIND_NONE = 0, KIND_PDF = 1, KIND_IMAGE = 2, KIND_OFFICE = 3;
    static readonly string[] ImageExts = { ".jpg", ".jpeg", ".jfif", ".png", ".gif", ".bmp", ".tif", ".tiff", ".ico" };
    static readonly string[] OfficeExts = { ".doc", ".docx", ".docm", ".rtf", ".xls", ".xlsx", ".xlsm", ".xlsb" };
    static readonly string[] AllExts;
    static PdfHoverPreview()
    {
        List<string> all = new List<string>();
        all.Add(".pdf");
        all.AddRange(ImageExts);
        all.AddRange(OfficeExts);
        AllExts = all.ToArray();
    }
    static int KindOf(string path)
    {
        string e = Path.GetExtension(path).ToLowerInvariant();
        if (e == ".pdf") return KIND_PDF;
        foreach (string x in ImageExts) if (x == e) return KIND_IMAGE;
        foreach (string x in OfficeExts) if (x == e) return KIND_OFFICE;
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
        List<double> ratios = new List<double>();
        float zoom = 1f;
        const float ZoomMin = 0.6f, ZoomMax = 2.6f;
        const int Gap = 6;
        public const int FormWidth = 560;
        int BaseContentW;

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

            BaseContentW = FormWidth - 2 - SystemInformation.VerticalScrollBarWidth - 2;
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

        public void ShowLoading(string text, Point near)
        {
            ClearHost();
            ratios.Clear();
            zoomBar.Visible = false;
            loading.Text = text;
            loading.Visible = true;
            loading.BringToFront();
            SizeAndPlace(360, near, true);
            if (!Visible) Show();
        }

        // PDF: crea los "huecos" (uno por pagina) con su altura, de una vez.
        public void BuildPdfSlots(double[] rs, Point near)
        {
            loading.Visible = false;
            ClearHost();
            ratios.Clear();
            zoom = 1f;
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
            loading.Visible = false;
            ClearHost();
            ratios.Clear();
            zoom = 1f;
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
        }

        public void ZoomStep(int wheelDelta)
        {
            int steps = wheelDelta / 120;
            if (steps == 0) steps = (wheelDelta > 0) ? 1 : -1;
            ZoomBy(steps);
        }

        public void HideIt() { if (Visible) { Hide(); ClearHost(); ratios.Clear(); loading.Visible = false; zoomBar.Visible = false; } }

        public void ScrollBy(int dy)
        {
            if (!Visible) return;
            int x = -Host.AutoScrollPosition.X;
            int y = -Host.AutoScrollPosition.Y;
            Host.AutoScrollPosition = new Point(x, y + dy);
        }

        public void ScrollByH(int dx)
        {
            if (!Visible) return;
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

        const int DwellMs = 350;
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
        }

        static Bitmap LoadImage(string path)
        {
            byte[] data = File.ReadAllBytes(path);
            using (MemoryStream ms = new MemoryStream(data))
                return new Bitmap(ms);
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
