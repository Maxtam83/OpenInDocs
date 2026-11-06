// OpenInDocs - open Office files in Google Docs / Sheets / Slides through Google Drive for desktop.
// Ouvre les fichiers Word / Excel / PowerPoint dans Google Docs / Sheets / Slides via Google Drive pour ordinateur.
//
// Not affiliated with Google. "Google Docs", "Sheets", "Slides" and "Drive" are trademarks of Google LLC.
//
// Built by installer\Install.ps1 with the C# compiler shipped with Windows (C# 5 / .NET Framework 4), in 3 variants:
//   /define:DOCS -> OpenInDocs.exe     /define:SHEETS -> OpenInSheets.exe     /define:SLIDES -> OpenInSlides.exe
// The three programs are identical; only the name and the icon shown by Windows differ.
// Started without argument: shows a diagnostic.
//
// How it works: the file is copied into "My Drive" (Drive for desktop uploads it), its Google Drive ID is read from
// Drive for desktop's local database (no Google API, no sign-in), then the matching editor URL is opened.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Windows.Forms;

#if SHEETS
[assembly: AssemblyTitle("Google Sheets")]
#elif SLIDES
[assembly: AssemblyTitle("Google Slides")]
#else
[assembly: AssemblyTitle("Google Docs")]
#endif
[assembly: AssemblyProduct("OpenInDocs")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]

// ---------------------------------------------------------------------------------------------
// Tiny FR/EN localisation: French if the Windows display language is French, English otherwise.
// ---------------------------------------------------------------------------------------------
static class L
{
    static readonly bool Fr = CultureInfo.CurrentUICulture.TwoLetterISOLanguageName == "fr";
    public static string T(string fr, string en) { return Fr ? fr : en; }
}

// ---------------------------------------------------------------------------------------------
// Read-only SQLite access through winsqlite3.dll (ships with Windows 10/11, nothing to install)
// ---------------------------------------------------------------------------------------------
static class WinSqlite
{
    const string D = "winsqlite3.dll";
    [DllImport(D, CallingConvention = CallingConvention.Winapi)] static extern int sqlite3_open_v2(byte[] f, out IntPtr db, int flags, IntPtr vfs);
    [DllImport(D, CallingConvention = CallingConvention.Winapi)] static extern int sqlite3_close(IntPtr db);
    [DllImport(D, CallingConvention = CallingConvention.Winapi, CharSet = CharSet.Unicode)] static extern int sqlite3_prepare16_v2(IntPtr db, string sql, int n, out IntPtr stmt, IntPtr tail);
    [DllImport(D, CallingConvention = CallingConvention.Winapi, CharSet = CharSet.Unicode)] static extern int sqlite3_bind_text16(IntPtr stmt, int i, string v, int n, IntPtr destructor);
    [DllImport(D, CallingConvention = CallingConvention.Winapi)] static extern int sqlite3_step(IntPtr stmt);
    [DllImport(D, CallingConvention = CallingConvention.Winapi)] static extern int sqlite3_column_count(IntPtr stmt);
    [DllImport(D, CallingConvention = CallingConvention.Winapi)] static extern IntPtr sqlite3_column_text16(IntPtr stmt, int i);
    [DllImport(D, CallingConvention = CallingConvention.Winapi)] static extern int sqlite3_finalize(IntPtr stmt);
    [DllImport(D, CallingConvention = CallingConvention.Winapi)] static extern IntPtr sqlite3_errmsg16(IntPtr db);

    public static List<string[]> Query(string path, string sql, string[] args)
    {
        IntPtr db;
        int rc = sqlite3_open_v2(Encoding.UTF8.GetBytes(path + "\0"), out db, 1 /* SQLITE_OPEN_READONLY */, IntPtr.Zero);
        if (rc != 0)
        {
            if (db != IntPtr.Zero) sqlite3_close(db);
            throw new Exception("Cannot open database (code " + rc + ")");
        }
        try
        {
            IntPtr st;
            rc = sqlite3_prepare16_v2(db, sql, -1, out st, IntPtr.Zero);
            if (rc != 0) throw new Exception(Marshal.PtrToStringUni(sqlite3_errmsg16(db)));
            try
            {
                for (int i = 0; i < args.Length; i++) sqlite3_bind_text16(st, i + 1, args[i], -1, new IntPtr(-1));
                List<string[]> rows = new List<string[]>();
                int n = sqlite3_column_count(st);
                while (true)
                {
                    rc = sqlite3_step(st);
                    if (rc == 100)
                    {
                        string[] r = new string[n];
                        for (int i = 0; i < n; i++)
                        {
                            IntPtr t = sqlite3_column_text16(st, i);
                            r[i] = (t == IntPtr.Zero) ? null : Marshal.PtrToStringUni(t);
                        }
                        rows.Add(r);
                    }
                    else if (rc == 101) break;
                    else throw new Exception(Marshal.PtrToStringUni(sqlite3_errmsg16(db)));
                }
                return rows;
            }
            finally { sqlite3_finalize(st); }
        }
        finally { sqlite3_close(db); }
    }
}

// ---------------------------------------------------------------------------------------------
// Settings: OpenInDocs.ini next to the program (old French key names are still accepted)
// ---------------------------------------------------------------------------------------------
class Config
{
    public string TargetFolder = "Opened from PC";
    public string GoogleAccount = "";
    public int MaxWaitSeconds = 90;
    public string DriveRoot = "";

    public static Config Load(string path)
    {
        Config c = new Config();
        if (!File.Exists(path)) return c;
        foreach (string raw in File.ReadAllLines(path, Encoding.UTF8))
        {
            string l = raw.Trim();
            if (l.Length == 0 || l[0] == '#' || l[0] == ';') continue;
            int i = l.IndexOf('=');
            if (i < 0) continue;
            string k = l.Substring(0, i).Trim().ToLowerInvariant();
            string v = l.Substring(i + 1).Trim();
            if (k == "targetfolder" || k == "dossiercible") { if (v.Trim('\\').Length > 0) c.TargetFolder = v.Trim('\\'); }
            else if (k == "googleaccount" || k == "comptegoogle") c.GoogleAccount = v;
            else if (k == "maxwaitseconds" || k == "delaimaxsecondes") { int n; if (int.TryParse(v, out n) && n > 0) c.MaxWaitSeconds = n; }
            else if (k == "driveroot" || k == "racinemondrive") c.DriveRoot = v.TrimEnd('\\');
        }
        return c;
    }
}

// ---------------------------------------------------------------------------------------------
// Google Drive for desktop: local metadata database
// ---------------------------------------------------------------------------------------------
static class Drive
{
    // Default name of the "My Drive" folder in some Drive languages. The real name is also read from the database,
    // so this list is only a fallback.
    static readonly string[] KnownRootNames = new string[] {
        "my drive", "mon drive", "meine ablage", "mi unidad", "il mio drive", "o meu drive", "mijn drive", "min disk", "mitt drev"
    };

    public static List<string> Databases()
    {
        List<string> res = new List<string>();
        string root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Google\\DriveFS");
        if (!Directory.Exists(root)) return res;
        foreach (string d in Directory.GetDirectories(root))
        {
            string f = Path.Combine(d, "metadata_sqlite_db");
            if (File.Exists(f)) res.Add(f);
        }
        return res;
    }

    // Drive uses a non-breaking space in some names ("Mon Drive"): normalise before comparing.
    public static string Norm(string s) { return s.Replace(' ', ' '); }

    // Finds the local "My Drive" folder, whatever the language: the root is the folder without parent in the database.
    public static string FindRoot(Config cfg, string file)
    {
        if (!string.IsNullOrEmpty(cfg.DriveRoot)) return cfg.DriveRoot;

        HashSet<string> names = new HashSet<string>(KnownRootNames);
        foreach (string db in Databases())
        {
            try
            {
                foreach (string[] r in Q(db, SqlRootFolders))
                    if (r[0] != null) names.Add(Norm(r[0]).ToLowerInvariant());
            }
            catch (Exception) { }
        }

        List<string> parents = new List<string>();
        if (file != null) parents.Add(Path.GetPathRoot(file));   // the drive of the file first
        foreach (DriveInfo d in DriveInfo.GetDrives())
        {
            try { if (d.DriveType != DriveType.CDRom && d.IsReady) parents.Add(d.RootDirectory.FullName); } catch (Exception) { }
        }
        parents.Add(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile));   // "mirror files" mode
        foreach (string p in parents)
        {
            string[] dirs;
            try { dirs = Directory.GetDirectories(p); } catch (Exception) { continue; }
            foreach (string d in dirs)
                if (names.Contains(Norm(Path.GetFileName(d)).ToLowerInvariant())) return d;
        }
        return null;
    }

    // Copy that tolerates files kept open by Drive.
    static void SharedCopy(string src, string dst)
    {
        using (FileStream i = new FileStream(src, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
        using (FileStream o = new FileStream(dst, FileMode.CreateNew, FileAccess.Write))
            i.CopyTo(o);
    }

    public static List<string[]> Q(string db, string sql, params string[] a)
    {
        try { return WinSqlite.Query(db, sql, a); }
        catch (Exception)
        {
            // Database locked by Drive: work on a temporary copy (database + WAL + SHM).
            string tmp = Path.Combine(Path.GetTempPath(), "openindocs_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tmp);
            try
            {
                foreach (string s in new string[] { "", "-wal", "-shm" })
                    if (File.Exists(db + s)) SharedCopy(db + s, Path.Combine(tmp, "metadata_sqlite_db" + s));
                return WinSqlite.Query(Path.Combine(tmp, "metadata_sqlite_db"), sql, a);
            }
            finally { try { Directory.Delete(tmp, true); } catch (Exception) { } }
        }
    }

    const string SqlRootFolders = "SELECT local_title FROM items WHERE is_folder = 1 AND local_title IS NOT NULL AND stable_id NOT IN (SELECT item_stable_id FROM stable_parents)";
    const string SqlCandidates = "SELECT stable_id, id FROM items WHERE local_title = ? COLLATE NOCASE AND is_folder = 0 AND trashed = 0 ORDER BY stable_id DESC";
    const string SqlNamedParent = "SELECT p.parent_stable_id FROM stable_parents p JOIN items i ON i.stable_id = p.parent_stable_id WHERE p.item_stable_id = CAST(? AS INTEGER) AND i.local_title = ? COLLATE NOCASE";
    const string SqlParent = "SELECT parent_stable_id FROM stable_parents WHERE item_stable_id = CAST(? AS INTEGER)";

    // seg = path relative to "My Drive". Returns the Drive ID ("local-..." while not uploaded yet) or null.
    public static string FindId(string[] seg)
    {
        foreach (string db in Databases())
            foreach (string[] c in Q(db, SqlCandidates, seg[seg.Length - 1]))
                if (CheckAncestry(db, c[0], seg)) return c[1];
        return null;
    }

    static bool CheckAncestry(string db, string stableId, string[] seg)
    {
        string cur = stableId;
        for (int i = seg.Length - 2; i >= 0; i--)
        {
            List<string[]> p = Q(db, SqlNamedParent, cur, seg[i]);
            if (p.Count == 0) return false;
            cur = p[0][0];
        }
        // The last parent must be the root of "My Drive" (the item without parent).
        List<string[]> r = Q(db, SqlParent, cur);
        if (r.Count == 0) return false;
        return Q(db, SqlParent, r[0][0]).Count == 0;
    }
}

// ---------------------------------------------------------------------------------------------
// Small "Sending to Google Drive..." window, bottom-right of the screen
// ---------------------------------------------------------------------------------------------
class ProgressForm : Form
{
    public volatile bool Cancelled;
    readonly Label lblStatus;

    public ProgressForm(string title, Icon icon, string status)
    {
        SuspendLayout();
        AutoScaleDimensions = new SizeF(96F, 96F);
        AutoScaleMode = AutoScaleMode.Dpi;
        Text = title;
        Icon = icon;
        FormBorderStyle = FormBorderStyle.None;
        StartPosition = FormStartPosition.Manual;
        TopMost = true;
        ShowInTaskbar = true;
        BackColor = Color.White;
        ClientSize = new Size(380, 92);
        Font = new Font("Segoe UI", 9F);

        PictureBox pic = new PictureBox();
        pic.Location = new Point(16, 18);
        pic.Size = new Size(36, 36);
        pic.SizeMode = PictureBoxSizeMode.Zoom;
        try { pic.Image = new Icon(icon, 48, 48).ToBitmap(); } catch (Exception) { }
        Controls.Add(pic);

        Label lblTitle = new Label();
        lblTitle.Text = title;
        lblTitle.Font = new Font("Segoe UI Semibold", 10.5F);
        lblTitle.Location = new Point(64, 14);
        lblTitle.Size = new Size(300, 22);
        Controls.Add(lblTitle);

        lblStatus = new Label();
        lblStatus.Text = status;
        lblStatus.ForeColor = Color.FromArgb(80, 80, 80);
        lblStatus.Location = new Point(64, 38);
        lblStatus.Size = new Size(300, 20);
        lblStatus.AutoEllipsis = true;
        Controls.Add(lblStatus);

        ProgressBar bar = new ProgressBar();
        bar.Style = ProgressBarStyle.Marquee;
        bar.MarqueeAnimationSpeed = 25;
        bar.Location = new Point(64, 66);
        bar.Size = new Size(230, 8);
        Controls.Add(bar);

        LinkLabel cancel = new LinkLabel();
        cancel.Text = L.T("Annuler", "Cancel");
        cancel.AutoSize = true;
        cancel.Location = new Point(308, 61);
        cancel.LinkColor = Color.FromArgb(26, 115, 232);
        cancel.LinkBehavior = LinkBehavior.HoverUnderline;
        cancel.LinkClicked += delegate { Cancelled = true; Close(); };
        Controls.Add(cancel);

        ResumeLayout(true);
    }

    protected override void OnLoad(EventArgs e)
    {
        base.OnLoad(e);
        Rectangle wa = Screen.PrimaryScreen.WorkingArea;
        Location = new Point(wa.Right - Width - 16, wa.Bottom - Height - 16);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);
        using (Pen p = new Pen(Color.FromArgb(210, 210, 210)))
            e.Graphics.DrawRectangle(p, 0, 0, ClientSize.Width - 1, ClientSize.Height - 1);
    }

    public void SetStatus(string s)
    {
        if (IsHandleCreated) BeginInvoke((MethodInvoker)delegate { lblStatus.Text = s; });
    }

    public void CloseSafe()
    {
        if (IsHandleCreated) BeginInvoke((MethodInvoker)delegate { Close(); });
    }
}

// ---------------------------------------------------------------------------------------------
// Main program
// ---------------------------------------------------------------------------------------------
static class Program
{
#if SHEETS
    const string Key = "sheets";
#elif SLIDES
    const string Key = "slides";
#else
    const string Key = "docs";
#endif
    static Config Cfg;
    static string FilePath;
    static string AppName;
    static Icon AppIcon;

    [DllImport("user32.dll")] static extern bool SetProcessDPIAware();
    [DllImport("user32.dll")] static extern bool AllowSetForegroundWindow(int processId);

    [STAThread]
    static int Main(string[] args)
    {
        try { SetProcessDPIAware(); } catch (Exception) { }
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        string exeDir = Path.GetDirectoryName(Application.ExecutablePath);
        Cfg = Config.Load(Path.Combine(exeDir, "OpenInDocs.ini"));
        AssemblyTitleAttribute t = (AssemblyTitleAttribute)Attribute.GetCustomAttribute(Assembly.GetExecutingAssembly(), typeof(AssemblyTitleAttribute));
        AppName = (t != null) ? t.Title : "Google Docs";
        try { AppIcon = new Icon(Path.Combine(exeDir, "icons\\app-" + Key + ".ico"), 48, 48); }
        catch (Exception)
        {
            try { AppIcon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch (Exception) { AppIcon = SystemIcons.Application; }
        }

        if (args.Length == 0) { Diagnostic(); return 0; }
        try
        {
            FilePath = Path.GetFullPath(args[0]);
            Run();
        }
        catch (Exception e) { Fallback(L.T("Erreur : ", "Error: ") + e.Message); }
        return 0;
    }

    static void Run()
    {
        if (!File.Exists(FilePath)) { Show(L.T("Fichier introuvable :\n", "File not found:\n") + FilePath, MessageBoxIcon.Error); return; }
        string root = Drive.FindRoot(Cfg, FilePath);
        if (root == null) { Fallback(L.T("Le dossier « Mon Drive » est introuvable. Google Drive pour ordinateur est-il lancé ?", "Your \"My Drive\" folder was not found. Is Google Drive for desktop running?")); return; }
        if (Drive.Databases().Count == 0) { Fallback(L.T("La base locale de Google Drive est introuvable.", "Google Drive for desktop's local database was not found.")); return; }
        string ext = Path.GetExtension(FilePath);

        string[] seg = null;
        if (Drive.Norm(FilePath).StartsWith(Drive.Norm(root) + "\\", StringComparison.OrdinalIgnoreCase))
        {
            // Already on Drive: if its ID is known, open right away, no window.
            seg = FilePath.Substring(root.Length).TrimStart('\\').Split('\\');
            string id0 = Drive.FindId(seg);
            if (id0 != null && !id0.StartsWith("local-")) { OpenEditor(id0, ext); return; }
        }

        ProgressForm f = new ProgressForm(AppName, AppIcon, L.T("Préparation de « ", "Preparing \"") + Path.GetFileName(FilePath) + L.T(" »…", "\"..."));
        string id = null;
        string error = null;
        string[] segments = seg;
        Thread worker = new Thread(delegate ()
        {
            try
            {
                if (segments == null) segments = CopyToDrive(root, f);
                else f.SetStatus(L.T("Envoi de « ", "Uploading \"") + segments[segments.Length - 1] + L.T(" » vers Google Drive…", "\" to Google Drive..."));
                DateTime end = DateTime.Now.AddSeconds(Cfg.MaxWaitSeconds);
                while (DateTime.Now < end && !f.Cancelled)
                {
                    string x = Drive.FindId(segments);
                    if (x != null && !x.StartsWith("local-")) { id = x; break; }
                    Thread.Sleep(700);
                }
            }
            catch (Exception e) { error = e.Message; }
            f.CloseSafe();
        });
        worker.IsBackground = true;
        f.Shown += delegate { worker.Start(); };
        Application.Run(f);

        if (f.Cancelled) return;
        worker.Join();
        if (id != null) OpenEditor(id, ext);
        else if (error != null) Fallback(L.T("Erreur : ", "Error: ") + error);
        else Fallback(L.T("Le fichier n'a pas reçu d'identifiant Google Drive au bout de ", "The file did not get a Google Drive ID after ") + Cfg.MaxWaitSeconds
                      + L.T(" s (pas de connexion, ou Google Drive en pause).", " s (no connection, or Google Drive is paused)."));
    }

    static string[] CopyToDrive(string root, ProgressForm f)
    {
        string folder = Path.Combine(root, Cfg.TargetFolder);
        Directory.CreateDirectory(folder);
        string b = Path.GetFileNameWithoutExtension(FilePath);
        string e = Path.GetExtension(FilePath);
        string name = b + e;
        int n = 2;
        bool already = false;
        long size = new FileInfo(FilePath).Length;
        string hash = null;
        while (File.Exists(Path.Combine(folder, name)))
        {
            // Same file already copied before? Reuse it instead of creating a duplicate.
            string target = Path.Combine(folder, name);
            if (new FileInfo(target).Length == size)
            {
                if (hash == null) hash = Fingerprint(FilePath);
                if (Fingerprint(target) == hash) { already = true; break; }
            }
            name = b + " (" + n + ")" + e;
            n++;
        }
        if (!already)
        {
            f.SetStatus(L.T("Copie de « ", "Copying \"") + name + L.T(" » dans Google Drive…", "\" into Google Drive..."));
            File.Copy(FilePath, Path.Combine(folder, name));
        }
        f.SetStatus(L.T("Envoi de « ", "Uploading \"") + name + L.T(" » vers Google Drive…", "\" to Google Drive..."));

        List<string> seg = new List<string>(Cfg.TargetFolder.Split(new char[] { '\\' }, StringSplitOptions.RemoveEmptyEntries));
        seg.Add(name);
        return seg.ToArray();
    }

    static string Fingerprint(string path)
    {
        using (FileStream s = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
        using (SHA256 h = SHA256.Create())
            return Convert.ToBase64String(h.ComputeHash(s));
    }

    static string BuildUrl(string id, string ext)
    {
        string e = ext.ToLowerInvariant();
        string type = null;
        if (e == ".docx" || e == ".doc" || e == ".docm" || e == ".dotx" || e == ".odt" || e == ".rtf") type = "document";
        else if (e == ".xlsx" || e == ".xls" || e == ".xlsm" || e == ".ods" || e == ".csv") type = "spreadsheets";
        else if (e == ".pptx" || e == ".ppt" || e == ".pptm" || e == ".ppsx" || e == ".odp") type = "presentation";

        string url = (type == null)
            ? "https://drive.google.com/file/d/" + id + "/view"
            : "https://docs.google.com/" + type + "/d/" + id + "/edit?rtpof=true&sd=true";
        if (!string.IsNullOrEmpty(Cfg.GoogleAccount))
            url += (url.Contains("?") ? "&" : "?") + "authuser=" + Uri.EscapeDataString(Cfg.GoogleAccount);
        return url;
    }

    static void OpenEditor(string id, string ext)
    {
        try { AllowSetForegroundWindow(-1); } catch (Exception) { }   // let the browser come to the front
        Process.Start(BuildUrl(id, ext));
    }

    static DialogResult Show(string text, MessageBoxIcon icon)
    {
        return MessageBox.Show(text, AppName, MessageBoxButtons.OK, icon);
    }

    // When something goes wrong, offer to open the file in the regular Office application instead.
    static void Fallback(string reason)
    {
        string ext = (FilePath == null) ? "" : Path.GetExtension(FilePath).ToLowerInvariant();
        string exe = "winword.exe", app = "Word";
        if (ext.StartsWith(".xl") || ext == ".ods" || ext == ".csv") { exe = "excel.exe"; app = "Excel"; }
        else if (ext.StartsWith(".pp") || ext == ".odp") { exe = "powerpnt.exe"; app = "PowerPoint"; }

        if (FilePath == null || !File.Exists(FilePath)) { Show(reason, MessageBoxIcon.Warning); return; }
        DialogResult r = MessageBox.Show(reason + L.T("\n\nOuvrir le fichier dans ", "\n\nOpen the file in ") + app + L.T(" à la place ?", " instead?"), AppName,
                                         MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
        if (r != DialogResult.Yes) return;
        try { Process.Start(exe, "\"" + FilePath + "\""); }
        catch (Exception) { Show(app + L.T(" n'a pas été trouvé sur ce PC.", " was not found on this PC."), MessageBoxIcon.Error); }
    }

    static void Diagnostic()
    {
        StringBuilder sb = new StringBuilder();
        sb.AppendLine(AppName + " (OpenInDocs) - " + L.T("diagnostic", "diagnostic"));
        sb.AppendLine();
        string root = Drive.FindRoot(Cfg, null);
        sb.AppendLine(L.T("Mon Drive : ", "My Drive: ") + (root ?? L.T("INTROUVABLE", "NOT FOUND")));
        sb.AppendLine(L.T("Dossier de dépôt : ", "Target folder: ") + Cfg.TargetFolder);
        sb.AppendLine(L.T("Compte Google : ", "Google account: ") + (Cfg.GoogleAccount.Length > 0 ? Cfg.GoogleAccount : L.T("(navigateur par défaut)", "(browser default)")));
        List<string> dbs = Drive.Databases();
        sb.AppendLine(L.T("Bases Google Drive : ", "Google Drive databases: ") + dbs.Count);
        bool ok = root != null && dbs.Count > 0;
        foreach (string db in dbs)
        {
            try
            {
                List<string[]> r = Drive.Q(db, "SELECT COUNT(*), SUM(id LIKE 'local-%') FROM items");
                sb.AppendLine("  " + r[0][0] + L.T(" éléments, ", " items, ") + (r[0][1] ?? "0") + L.T(" en cours d'envoi", " being uploaded"));
            }
            catch (Exception e) { ok = false; sb.AppendLine("  " + L.T("ERREUR : ", "ERROR: ") + e.Message); }
        }
        sb.AppendLine();
        sb.AppendLine(ok ? L.T("Tout est prêt.", "Everything is ready.")
                         : L.T("Un problème empêche le fonctionnement (voir ci-dessus).", "Something prevents it from working (see above)."));
        Show(sb.ToString(), ok ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
    }
}
