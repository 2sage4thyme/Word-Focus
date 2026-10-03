# Word Focus - highlights the word under your mouse plus the next few words.
# Required Notice: Copyright 2sage4thyme (https://github.com/2sage4thyme)
# License: Big Time Public License 2.0.2 - https://bigtimelicense.com/versions/2.0.2 (see LICENSE.md)
#   Ctrl+Alt+H  = turn highlight on/off
#   Ctrl+Alt+K  = open settings (word-count slider, ombre on/off)
#   Tray icon (bottom-right) -> Exit to quit.
# What this script does: compiles the C# code below in memory (using Windows' built-in
# compiler), then runs it. It installs nothing, writes no files, changes no settings,
# and does not start with Windows. It only reads the text under your mouse through
# Windows' accessibility interface (UI Automation) and draws a see-through overlay.
# While a highlight is showing, it temporarily swaps the arrow, text I-beam and link-hand
# pointers for see-through copies (in memory only), and puts your normal pointers back
# when you stop hovering, turn it off, or exit. Nothing about your pointers is saved.
# A diagnostic log exists but is OFF (Log.Enabled = false in the code below).

$ErrorActionPreference = 'Stop'
try {
    Add-Type -AssemblyName System.Windows.Forms, System.Drawing, UIAutomationClient, UIAutomationTypes, WindowsBase
    $refs = @(
        [System.Windows.Forms.Form].Assembly.Location,
        [System.Drawing.Bitmap].Assembly.Location,
        [System.Windows.Automation.AutomationElement].Assembly.Location,
        [System.Windows.Automation.Text.TextUnit].Assembly.Location,
        [System.Windows.Rect].Assembly.Location
    )
    $code = @'
// Word Focus - highlights the word under the mouse plus the next few words.
// Written for C# 5 so Windows 10's built-in compiler (used by PowerShell Add-Type) accepts it.
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using System.Windows.Automation;
using System.Windows.Automation.Text;

namespace WordFocus
{
    // ---- Windows functions we call directly ----
    public static class Native
    {
        [StructLayout(LayoutKind.Sequential)]
        public struct POINT { public int X; public int Y; public POINT(int x, int y) { X = x; Y = y; } }
        [StructLayout(LayoutKind.Sequential)]
        public struct SIZE { public int cx; public int cy; public SIZE(int w, int h) { cx = w; cy = h; } }
        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        public struct BLENDFUNCTION { public byte BlendOp; public byte BlendFlags; public byte SourceConstantAlpha; public byte AlphaFormat; }

        [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
        [DllImport("user32.dll")] public static extern bool GetCursorPos(out POINT p);
        [DllImport("user32.dll")] public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint mods, uint vk);
        [DllImport("user32.dll")] public static extern bool UnregisterHotKey(IntPtr hWnd, int id);
        [DllImport("user32.dll")] public static extern IntPtr GetDC(IntPtr hWnd);
        [DllImport("user32.dll")] public static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);
        [DllImport("user32.dll")]
        public static extern bool UpdateLayeredWindow(IntPtr hwnd, IntPtr hdcDst, ref POINT pptDst, ref SIZE psize,
            IntPtr hdcSrc, ref POINT pptSrc, int crKey, ref BLENDFUNCTION pblend, int dwFlags);
        [DllImport("gdi32.dll")] public static extern IntPtr CreateCompatibleDC(IntPtr hDC);
        [DllImport("gdi32.dll")] public static extern bool DeleteDC(IntPtr hdc);
        [DllImport("gdi32.dll")] public static extern IntPtr SelectObject(IntPtr hDC, IntPtr obj);
        [DllImport("gdi32.dll")] public static extern bool DeleteObject(IntPtr obj);

        // --- cursor functions (for the see-through pointer) ---
        [StructLayout(LayoutKind.Sequential)]
        public struct ICONINFO { public bool fIcon; public int xHotspot; public int yHotspot; public IntPtr hbmMask; public IntPtr hbmColor; }
        [StructLayout(LayoutKind.Sequential)]
        public struct BITMAP { public int bmType; public int bmWidth; public int bmHeight; public int bmWidthBytes; public ushort bmPlanes; public ushort bmBitsPixel; public IntPtr bmBits; }
        [StructLayout(LayoutKind.Sequential)]
        public struct BITMAPINFOHEADER { public int biSize; public int biWidth; public int biHeight; public short biPlanes; public short biBitCount; public int biCompression; public int biSizeImage; public int biXPelsPerMeter; public int biYPelsPerMeter; public int biClrUsed; public int biClrImportant; }

        [DllImport("user32.dll")] public static extern IntPtr LoadCursor(IntPtr hInstance, IntPtr id);
        [DllImport("user32.dll")] public static extern bool GetIconInfo(IntPtr hIcon, out ICONINFO info);
        [DllImport("user32.dll")] public static extern IntPtr CreateIconIndirect(ref ICONINFO info);
        [DllImport("user32.dll")] public static extern bool DrawIconEx(IntPtr hdc, int x, int y, IntPtr hIcon, int w, int h, int step, IntPtr brush, int flags);
        [DllImport("user32.dll", SetLastError = true)] public static extern bool SetSystemCursor(IntPtr hcur, int id);
        [DllImport("user32.dll")] public static extern bool SystemParametersInfo(int action, int param, IntPtr pv, int winIni);
        [DllImport("gdi32.dll")] public static extern int GetObject(IntPtr h, int size, out BITMAP bm);
        [DllImport("gdi32.dll")] public static extern IntPtr CreateDIBSection(IntPtr hdc, ref BITMAPINFOHEADER bmi, int usage, out IntPtr bits, IntPtr section, int offset);
        [DllImport("gdi32.dll")] public static extern IntPtr CreateBitmap(int w, int h, int planes, int bpp, byte[] bits);
    }

    // ---- Makes the mouse pointer see-through while a highlight is showing ----
    // Uses SetSystemCursor (temporary, in memory, nothing saved). Restoring calls
    // SPI_SETCURSORS, which reloads your normal pointer scheme exactly as Windows does at login.
    public class CursorFader
    {
        private class CurData { public int Id; public int W; public int H; public int HotX; public int HotY; public int[] Pixels; }

        // 32512 = normal arrow, 32513 = text I-beam, 32649 = link hand
        private static readonly int[] Ids = new int[] { 32512, 32513, 32649 };
        private readonly List<CurData> data = new List<CurData>();
        private bool faded = false;
        public bool IsFaded { get { return faded; } }

        public CursorFader(float opacity)
        {
            RestoreAlways(); // in case an earlier run was force-closed while faded
            foreach (int id in Ids)
            {
                try
                {
                    IntPtr h = Native.LoadCursor(IntPtr.Zero, new IntPtr(id));
                    CurData d = Capture(h, opacity);
                    if (d != null)
                    {
                        d.Id = id; data.Add(d);
                        int vis = 0; foreach (int px in d.Pixels) if (((px >> 24) & 255) > 0) vis++;
                        Log.Write("Cursor " + id + " captured " + d.W + "x" + d.H + " hotspot " + d.HotX + "," + d.HotY + " visible pixels " + vis);
                    }
                    else Log.Write("Cursor " + id + " capture returned nothing (handle " + h + ")");
                }
                catch (Exception ex) { Log.Write("Cursor " + id + " capture ERROR " + ex.Message); }
            }
        }

        public void Fade()
        {
            if (faded || data.Count == 0) return;
            foreach (CurData d in data)
            {
                IntPtr c = Create(d);
                bool ok = c != IntPtr.Zero && Native.SetSystemCursor(c, d.Id); // Windows takes ownership of c
                Log.Write("Fade cursor " + d.Id + " created=" + (c != IntPtr.Zero) + " set=" + ok + (ok ? "" : " err=" + Marshal.GetLastWin32Error()));
            }
            faded = true;
        }

        public void Restore() { if (faded) RestoreAlways(); }

        public void RestoreAlways()
        {
            bool ok = Native.SystemParametersInfo(0x57, 0, IntPtr.Zero, 0); // SPI_SETCURSORS: reload normal pointers
            Log.Write("Restore cursors ok=" + ok);
            faded = false;
        }

        private static int[] DrawOn(IntPtr cur, int w, int h, Color bg)
        {
            using (Bitmap b = new Bitmap(w, h, PixelFormat.Format32bppRgb))
            {
                using (Graphics g = Graphics.FromImage(b))
                {
                    g.Clear(bg);
                    IntPtr hdc = g.GetHdc();
                    Native.DrawIconEx(hdc, 0, 0, cur, w, h, 0, IntPtr.Zero, 3);
                    g.ReleaseHdc(hdc);
                }
                BitmapData bd = b.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.ReadOnly, PixelFormat.Format32bppRgb);
                int[] px = new int[w * h];
                for (int y = 0; y < h; y++)
                    Marshal.Copy(new IntPtr(bd.Scan0.ToInt64() + (long)y * bd.Stride), px, y * w, w);
                b.UnlockBits(bd);
                return px;
            }
        }

        private static CurData Capture(IntPtr cur, float opacity)
        {
            Native.ICONINFO ii;
            if (cur == IntPtr.Zero || !Native.GetIconInfo(cur, out ii)) return null;
            Native.BITMAP bm;
            Native.GetObject(ii.hbmMask, Marshal.SizeOf(typeof(Native.BITMAP)), out bm);
            int w = bm.bmWidth;
            int h = ii.hbmColor == IntPtr.Zero ? bm.bmHeight / 2 : bm.bmHeight;
            if (ii.hbmMask != IntPtr.Zero) Native.DeleteObject(ii.hbmMask);
            if (ii.hbmColor != IntPtr.Zero) Native.DeleteObject(ii.hbmColor);
            if (w <= 0 || h <= 0) return null;

            // Draw the pointer on black and on white; the difference tells us its true transparency.
            int[] k = DrawOn(cur, w, h, Color.Black);
            int[] wt = DrawOn(cur, w, h, Color.White);
            int n = w * h;
            int[] outp = new int[n];
            bool[] inv = new bool[n];
            for (int i = 0; i < n; i++)
            {
                int kr = (k[i] >> 16) & 255, kg = (k[i] >> 8) & 255, kb = k[i] & 255;
                int wr = (wt[i] >> 16) & 255, wg = (wt[i] >> 8) & 255, wb = wt[i] & 255;
                if (wr + wg + wb < kr + kg + kb - 30) { inv[i] = true; outp[i] = unchecked((int)0xFF000000); continue; } // "inverting" pixels -> black
                int a = 255 - ((wr - kr) + (wg - kg) + (wb - kb)) / 3;
                if (a <= 0) { outp[i] = 0; continue; }
                if (a > 255) a = 255;
                int r = Math.Min(255, kr * 255 / a), gg = Math.Min(255, kg * 255 / a), bb = Math.Min(255, kb * 255 / a);
                outp[i] = (a << 24) | (r << 16) | (gg << 8) | bb;
            }
            // White outline around formerly-inverting pixels so the I-beam stays visible on dark backgrounds.
            for (int y = 0; y < h; y++)
                for (int x = 0; x < w; x++)
                {
                    int i = y * w + x;
                    if (inv[i] || ((outp[i] >> 24) & 255) != 0) continue;
                    bool near = (x > 0 && inv[i - 1]) || (x < w - 1 && inv[i + 1]) || (y > 0 && inv[i - w]) || (y < h - 1 && inv[i + w]);
                    if (near) outp[i] = unchecked((int)0xFFFFFFFF);
                }
            // Apply the see-through amount.
            for (int i = 0; i < n; i++)
            {
                int a = (outp[i] >> 24) & 255;
                int a2 = (int)(a * opacity);
                outp[i] = (a2 << 24) | (outp[i] & 0xFFFFFF);
            }

            CurData d = new CurData();
            d.W = w; d.H = h; d.HotX = ii.xHotspot; d.HotY = ii.yHotspot; d.Pixels = outp;
            return d;
        }

        private static IntPtr Create(CurData d)
        {
            Native.BITMAPINFOHEADER bih = new Native.BITMAPINFOHEADER();
            bih.biSize = Marshal.SizeOf(typeof(Native.BITMAPINFOHEADER));
            bih.biWidth = d.W; bih.biHeight = -d.H; bih.biPlanes = 1; bih.biBitCount = 32; bih.biCompression = 0;
            IntPtr bits;
            IntPtr color = Native.CreateDIBSection(IntPtr.Zero, ref bih, 0, out bits, IntPtr.Zero, 0);
            if (color == IntPtr.Zero) return IntPtr.Zero;
            Marshal.Copy(d.Pixels, 0, bits, d.Pixels.Length);
            byte[] maskBits = new byte[((d.W + 15) / 16 * 2) * d.H]; // all zero
            IntPtr mask = Native.CreateBitmap(d.W, d.H, 1, 1, maskBits);
            Native.ICONINFO ni = new Native.ICONINFO();
            ni.fIcon = false; ni.xHotspot = d.HotX; ni.yHotspot = d.HotY; ni.hbmMask = mask; ni.hbmColor = color;
            IntPtr cur = Native.CreateIconIndirect(ref ni);
            Native.DeleteObject(color);
            Native.DeleteObject(mask);
            return cur;
        }
    }

    // ---- See-through, click-through window that draws the highlight ----
    public class Overlay : Form
    {
        public Overlay()
        {
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            TopMost = true;
        }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                // layered | click-through | no taskbar/alt-tab | topmost | never takes focus
                cp.ExStyle |= 0x80000 | 0x20 | 0x80 | 0x8 | 0x08000000;
                return cp;
            }
        }

        protected override bool ShowWithoutActivation { get { return true; } }

        public void SetBitmap(Bitmap bmp, int x, int y)
        {
            IntPtr screenDc = Native.GetDC(IntPtr.Zero);
            IntPtr memDc = Native.CreateCompatibleDC(screenDc);
            IntPtr hBmp = IntPtr.Zero, oldBmp = IntPtr.Zero;
            try
            {
                hBmp = bmp.GetHbitmap(Color.FromArgb(0));
                oldBmp = Native.SelectObject(memDc, hBmp);
                Native.SIZE size = new Native.SIZE(bmp.Width, bmp.Height);
                Native.POINT src = new Native.POINT(0, 0);
                Native.POINT dst = new Native.POINT(x, y);
                Native.BLENDFUNCTION blend = new Native.BLENDFUNCTION();
                blend.BlendOp = 0; blend.BlendFlags = 0; blend.SourceConstantAlpha = 255; blend.AlphaFormat = 1;
                Native.UpdateLayeredWindow(Handle, screenDc, ref dst, ref size, memDc, ref src, 0, ref blend, 2);
            }
            finally
            {
                Native.ReleaseDC(IntPtr.Zero, screenDc);
                if (hBmp != IntPtr.Zero) { Native.SelectObject(memDc, oldBmp); Native.DeleteObject(hBmp); }
                Native.DeleteDC(memDc);
            }
        }
    }

    // ---- Small settings window: slider + ombre checkbox ----
    public class SettingsForm : Form
    {
        private readonly Controller ctl;
        private readonly TrackBar slider;
        private readonly Label sliderLabel;
        private readonly CheckBox ombreBox;
        private readonly CheckBox onBox;

        public SettingsForm(Controller c)
        {
            ctl = c;
            Text = "Word Focus";
            FormBorderStyle = FormBorderStyle.FixedToolWindow;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(320, 210);
            TopMost = true;

            onBox = new CheckBox();
            onBox.Text = "Highlight on   (Ctrl+Alt+H)";
            onBox.SetBounds(14, 10, 290, 24);
            onBox.CheckedChanged += delegate { if (onBox.Checked != ctl.Enabled) ctl.Toggle(); };

            sliderLabel = new Label();
            sliderLabel.SetBounds(14, 40, 290, 20);

            slider = new TrackBar();
            slider.Minimum = 1; slider.Maximum = 11; slider.TickFrequency = 1; slider.LargeChange = 1;
            slider.SetBounds(10, 60, 300, 45);
            slider.ValueChanged += delegate { ctl.SetCount(slider.Value); UpdateLabel(); };

            ombreBox = new CheckBox();
            ombreBox.Text = "Ombre fade   (off = one solid color)";
            ombreBox.SetBounds(14, 110, 290, 24);
            ombreBox.CheckedChanged += delegate { ctl.SetOmbre(ombreBox.Checked); };

            Label hint = new Label();
            hint.Text = "Ctrl+Alt+K opens this window. Closing it keeps Word Focus running.";
            hint.ForeColor = SystemColors.GrayText;
            hint.SetBounds(14, 182, 300, 20);

            Controls.Add(onBox); Controls.Add(sliderLabel); Controls.Add(slider);
            Controls.Add(ombreBox); Controls.Add(hint);

            // Color row: preset swatches + a "Custom..." button that opens the Windows color picker.
            Label colorLabel = new Label();
            colorLabel.Text = "Color:";
            colorLabel.SetBounds(14, 146, 42, 20);
            Controls.Add(colorLabel);
            Color[] presets = new Color[] {
                Color.FromArgb(255, 214, 0),   // yellow
                Color.FromArgb(120, 220, 80),  // green
                Color.FromArgb(80, 170, 255),  // blue
                Color.FromArgb(255, 120, 190), // pink
                Color.FromArgb(255, 150, 40),  // orange
                Color.FromArgb(170, 120, 255)  // purple
            };
            string[] names = new string[] { "Yellow", "Green", "Blue", "Pink", "Orange", "Purple" };
            ToolTip tips = new ToolTip();
            for (int i = 0; i < presets.Length; i++)
            {
                Button sw = new Button();
                sw.FlatStyle = FlatStyle.Flat;
                sw.BackColor = presets[i];
                sw.SetBounds(58 + i * 28, 142, 24, 24);
                Color picked = presets[i];
                sw.Click += delegate { ctl.SetColor(picked); };
                tips.SetToolTip(sw, names[i]);
                Controls.Add(sw);
            }
            Button custom = new Button();
            custom.Text = "Custom...";
            custom.SetBounds(228, 141, 80, 26);
            custom.Click += delegate
            {
                using (ColorDialog dlg = new ColorDialog())
                {
                    dlg.Color = Controller.HighlightColor;
                    dlg.FullOpen = true;
                    if (dlg.ShowDialog(this) == DialogResult.OK) ctl.SetColor(dlg.Color);
                }
            };
            Controls.Add(custom);
            SyncFromController();
        }

        public void SyncFromController()
        {
            onBox.Checked = ctl.Enabled;
            slider.Value = ctl.Count;
            ombreBox.Checked = ctl.Ombre;
            UpdateLabel();
        }

        private void UpdateLabel()
        {
            int v = slider.Value;
            if (v >= 11) sliderLabel.Text = "Words highlighted: rest of the line";
            else if (v == 1) sliderLabel.Text = "Words highlighted: 1 (just the hovered word)";
            else sliderLabel.Text = "Words highlighted: " + v + " (hovered word + next " + (v - 1) + ")";
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing) { e.Cancel = true; Hide(); }
            base.OnFormClosing(e);
        }
    }

    // ---- Invisible window that receives the global hotkeys ----
    public class HotkeyWindow : NativeWindow, IDisposable
    {
        public event Action<int> Pressed;
        public HotkeyWindow() { CreateHandle(new CreateParams()); }
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x0312 && Pressed != null) Pressed(m.WParam.ToInt32()); // WM_HOTKEY
            base.WndProc(ref m);
        }
        public void Dispose() { DestroyHandle(); }
    }

    public class WordBox
    {
        public System.Windows.Rect[] Rects;
        public WordBox(System.Windows.Rect[] r) { Rects = r; }
    }

    // ---- Ties it all together ----
    public class Controller
    {
        // ===== Easy-to-change defaults =====
        public static Color HighlightColor = Color.FromArgb(255, 214, 0); // yellow
        public const int MaxAlpha = 150;     // strength of the first word (0-255)
        public const float FadeTo = 0.15f;   // last word fades to this fraction of MaxAlpha
        public const int LineModeValue = 11; // slider value that means "rest of the line"
        public const float CursorOpacity = 0.10f; // pointer see-through amount while highlighting (0 = invisible, 1 = normal)
        public const int CursorRestoreDelayMs = 700; // keeps the pointer faded while moving across gaps between words
        // ====================================

        private volatile bool enabled = false;
        private volatile bool ombre = true;
        private volatile int count = 3;
        private volatile bool dirty = true;
        private volatile bool running = true;

        private readonly Overlay overlay;
        private readonly HotkeyWindow hotkeys;
        private readonly NotifyIcon tray;
        private SettingsForm settings;
        private CursorFader fader;
        private int lastHitTick;
        private System.Windows.Forms.Timer restoreTimer;

        public bool Enabled { get { return enabled; } }
        public bool Ombre { get { return ombre; } }
        public int Count { get { return count; } }

        public Controller()
        {
            try { fader = new CursorFader(CursorOpacity); } catch { fader = null; }
            AppDomain.CurrentDomain.ProcessExit += delegate { RestoreCursor(); };
            AppDomain.CurrentDomain.UnhandledException += delegate { RestoreCursor(); };
            Application.ThreadException += delegate { RestoreCursor(); };

            overlay = new Overlay();
            overlay.Show();
            Clear();

            // Puts the normal pointer back shortly after you stop hovering over text.
            restoreTimer = new System.Windows.Forms.Timer();
            restoreTimer.Interval = 150;
            restoreTimer.Tick += delegate
            {
                if (fader != null && fader.IsFaded &&
                    (!enabled || unchecked(Environment.TickCount - lastHitTick) > CursorRestoreDelayMs))
                    RestoreCursor();
            };
            restoreTimer.Start();

            hotkeys = new HotkeyWindow();
            hotkeys.Pressed += OnHotkey;
            bool ok1 = Native.RegisterHotKey(hotkeys.Handle, 1, 0x2 | 0x1 | 0x4000, 0x48); // Ctrl+Alt+H
            bool ok2 = Native.RegisterHotKey(hotkeys.Handle, 2, 0x2 | 0x1 | 0x4000, 0x4B); // Ctrl+Alt+K

            tray = new NotifyIcon();
            tray.Icon = SystemIcons.Application;
            tray.Visible = true;
            ContextMenuStrip menu = new ContextMenuStrip();
            menu.Items.Add("Highlight on/off  (Ctrl+Alt+H)", null, delegate { Toggle(); });
            menu.Items.Add("Settings...  (Ctrl+Alt+K)", null, delegate { ShowSettings(); });
            menu.Items.Add("Exit", null, delegate { Exit(); });
            tray.ContextMenuStrip = menu;
            tray.DoubleClick += delegate { ShowSettings(); };
            UpdateTrayText();

            if (!ok1 || !ok2)
                MessageBox.Show("Another program is already using Ctrl+Alt+H or Ctrl+Alt+K, so that hotkey won't work.\n" +
                                "You can still use the tray icon (bottom-right of the taskbar).", "Word Focus");

            Thread worker = new Thread(Loop);
            worker.IsBackground = true;
            worker.SetApartmentState(ApartmentState.MTA);
            worker.Start();
        }

        private void OnHotkey(int id)
        {
            if (id == 1) Toggle();
            else if (id == 2) ShowSettings();
        }

        public void Toggle()
        {
            enabled = !enabled;
            dirty = true;
            if (!enabled) { Clear(); RestoreCursor(); }
            UpdateTrayText();
            if (settings != null) settings.SyncFromController();
        }

        public void SetCount(int v) { count = v; dirty = true; }
        public void SetOmbre(bool v) { ombre = v; dirty = true; }
        public void SetColor(Color c) { HighlightColor = Color.FromArgb(255, c.R, c.G, c.B); dirty = true; }

        private void UpdateTrayText() { tray.Text = "Word Focus - " + (enabled ? "ON" : "off"); }

        private void ShowSettings()
        {
            if (settings == null) settings = new SettingsForm(this);
            settings.SyncFromController();
            settings.Show();
            settings.Activate();
        }

        private void Exit()
        {
            running = false;
            RestoreCursor();
            Native.UnregisterHotKey(hotkeys.Handle, 1);
            Native.UnregisterHotKey(hotkeys.Handle, 2);
            hotkeys.Dispose();
            tray.Visible = false;
            tray.Dispose();
            Application.Exit();
        }

        // Background loop: when the mouse moves, ask Windows which words are under it.
        private void Loop()
        {
            Native.POINT last = new Native.POINT(int.MinValue, int.MinValue);
            while (running)
            {
                Thread.Sleep(70);
                if (!enabled) continue;
                Native.POINT p;
                if (!Native.GetCursorPos(out p)) continue;
                if (p.X == last.X && p.Y == last.Y && !dirty) continue;
                last = p;
                dirty = false;

                List<WordBox> boxes = null;
                try { boxes = Query(p); } catch (Exception ex) { Log.Write("Q " + p.X + "," + p.Y + " ERROR " + ex.GetType().Name + ": " + ex.Message); boxes = null; }

                List<WordBox> result = boxes;
                try { overlay.BeginInvoke(new Action(delegate { Render(result); })); } catch { }
            }
        }

        private static bool HasWordChar(string s)
        {
            if (s == null) return false;
            foreach (char c in s) if (char.IsLetterOrDigit(c)) return true;
            return false;
        }

        private static bool Hit(System.Windows.Rect[] rects, System.Windows.Point pt)
        {
            foreach (System.Windows.Rect r in rects)
            {
                System.Windows.Rect g = r; g.Inflate(2, Math.Max(4, r.Height * 0.35)); // covers the gap between lines
                if (g.Contains(pt)) return true;
            }
            return false;
        }

        private static string Describe(AutomationElement e)
        {
            if (e == null) return "none";
            try { return e.Current.ControlType.ProgrammaticName.Replace("ControlType.", "") + "(" + e.Current.ClassName + ")"; }
            catch { return "?"; }
        }

        private static TextPattern GetTP(AutomationElement e)
        {
            object o;
            try { if (e != null && e.TryGetCurrentPattern(TextPattern.Pattern, out o)) return (TextPattern)o; } catch { }
            return null;
        }

        private static bool IsDocument(AutomationElement e)
        {
            try { return e.Current.ControlType == ControlType.Document; } catch { return false; }
        }

        private static bool Contains(AutomationElement e, System.Windows.Point pt)
        {
            try
            {
                System.Windows.Rect r = e.Current.BoundingRectangle;
                if (r.IsEmpty) return false;
                r.Inflate(1, 1);
                return r.Contains(pt);
            }
            catch { return false; }
        }

        // Finds the smallest piece of text under the mouse that can report its words.
        // Browser pages (Claude, Chrome, Edge) often report the paragraph box instead of the text inside it,
        // and asking the whole page is unreliable, so: 1) the element itself, 2) look inside it,
        // 3) look at its parents, 4) last resort, the whole page.
        private static TextPattern FindTextPattern(AutomationElement el, System.Windows.Point pt, out AutomationElement src, out string how)
        {
            src = null; how = "none";
            if (el == null) return null;
            TreeWalker walker = TreeWalker.ControlViewWalker;

            TextPattern tp = GetTP(el);
            if (tp != null && !IsDocument(el)) { src = el; how = "self"; return tp; }

            // Look inside: step down into whichever child is under the mouse.
            AutomationElement node = el;
            for (int depth = 0; depth < 40 && node != null; depth++)
            {
                AutomationElement next = null;
                AutomationElement child = null;
                try { child = walker.GetFirstChild(node); } catch { child = null; }
                for (int k = 0; child != null && k < 400; k++)
                {
                    if (Contains(child, pt)) { next = child; break; }
                    try { child = walker.GetNextSibling(child); } catch { child = null; }
                }
                if (next == null) break;
                node = next;
                tp = GetTP(node);
                if (tp != null && !IsDocument(node)) { src = node; how = "inside"; return tp; }
            }

            // Look at parents.
            TextPattern docTp = IsDocument(el) ? GetTP(el) : null;
            AutomationElement docEl = docTp != null ? el : null;
            AutomationElement up = el;
            for (int hops = 0; hops < 40; hops++)
            {
                try { up = walker.GetParent(up); } catch { up = null; }
                if (up == null) break;
                tp = GetTP(up);
                if (tp == null) continue;
                if (!IsDocument(up)) { src = up; how = "parent"; return tp; }
                if (docTp == null) { docTp = tp; docEl = up; }
            }
            if (docTp != null) { src = docEl; how = "page"; }
            return docTp;
        }

        private List<WordBox> Query(Native.POINT p)
        {
            System.Windows.Point pt = new System.Windows.Point(p.X, p.Y);
            AutomationElement el = AutomationElement.FromPoint(pt);

            // Find the nearest element (or parent) that exposes its text.
            string how;
            AutomationElement src;
            TextPattern tp = FindTextPattern(el, pt, out src, out how);
            string elInfo = Describe(el) + " -> " + how + " " + Describe(src);
            if (tp == null) { Log.Write("Q " + p.X + "," + p.Y + " no TextPattern under " + elInfo); return null; }

            // The position Windows gives can land on the word next door (e.g. right edge of the last letter),
            // so check the word it picked, then the previous and next words.
            TextPatternRange start = tp.RangeFromPoint(pt);
            TextPatternRange cur = null;
            System.Windows.Rect[] firstRects = null;
            string tried = "";
            int[] offsets = new int[] { 0, -1, 1 };
            foreach (int off in offsets)
            {
                TextPatternRange c = start.Clone();
                if (off != 0 && c.Move(TextUnit.Word, off) == 0) continue;
                c.ExpandToEnclosingUnit(TextUnit.Word);
                System.Windows.Rect[] rs = c.GetBoundingRectangles();
                string tx = c.GetText(40);
                bool hit = rs.Length > 0 && Hit(rs, pt);
                tried += " [" + off + ":'" + (tx ?? "").Replace("\n", " ") + "' rects=" + rs.Length +
                         (rs.Length > 0 ? " " + (int)rs[0].Left + "," + (int)rs[0].Top + " " + (int)rs[0].Width + "x" + (int)rs[0].Height : "") +
                         " hit=" + hit + "]";
                if (hit && HasWordChar(tx)) { cur = c; firstRects = rs; break; }
            }
            Log.Write("Q " + p.X + "," + p.Y + " el=" + elInfo + (cur == null ? " MISS" : " OK") + tried);
            if (cur == null) return null;

            bool lineMode = count >= LineModeValue;
            int want = lineMode ? 300 : count;
            double lineMid = firstRects[0].Top + firstRects[0].Height / 2;
            double lineH = firstRects[0].Height;

            List<WordBox> list = new List<WordBox>();
            for (int guard = 0; list.Count < want && guard < 400; guard++)
            {
                string text = cur.GetText(100);
                if (HasWordChar(text))
                {
                    System.Windows.Rect[] rects = cur.GetBoundingRectangles();
                    if (rects.Length == 0) break; // scrolled off-screen
                    if (lineMode)
                    {
                        double mid = rects[0].Top + rects[0].Height / 2;
                        if (Math.Abs(mid - lineMid) > lineH * 0.5) break; // reached the next line
                    }
                    list.Add(new WordBox(rects));
                }
                if (cur.Move(TextUnit.Word, 1) == 0) break; // end of text
                cur.ExpandToEnclosingUnit(TextUnit.Word);
            }
            return list;
        }

        private int AlphaAt(int i, int n)
        {
            if (!ombre || n <= 1) return MaxAlpha;
            float t = (float)i / n; // 0 at first word, 1 at end of last word
            return (int)(MaxAlpha * (1f - (1f - FadeTo) * t));
        }

        private void RestoreCursor()
        {
            try { if (fader != null) fader.Restore(); } catch { }
        }

        private void Clear()
        {
            using (Bitmap b = new Bitmap(1, 1, PixelFormat.Format32bppArgb))
            {
                b.SetPixel(0, 0, Color.Transparent);
                overlay.SetBitmap(b, 0, 0);
            }
        }

        private void Render(List<WordBox> boxes)
        {
            if (!enabled || boxes == null || boxes.Count == 0) { Clear(); return; }
            lastHitTick = Environment.TickCount;
            if (fader != null) { try { fader.Fade(); } catch { } }

            double l = double.MaxValue, t = double.MaxValue, r = double.MinValue, b = double.MinValue;
            foreach (WordBox w in boxes)
                foreach (System.Windows.Rect rc in w.Rects)
                {
                    l = Math.Min(l, rc.Left); t = Math.Min(t, rc.Top);
                    r = Math.Max(r, rc.Right); b = Math.Max(b, rc.Bottom);
                }
            int pad = 2;
            int x0 = (int)Math.Floor(l) - pad, y0 = (int)Math.Floor(t) - pad;
            int w0 = (int)Math.Ceiling(r) - x0 + pad, h0 = (int)Math.Ceiling(b) - y0 + pad;
            if (w0 < 1 || h0 < 1 || w0 > 10000 || h0 > 10000) { Clear(); return; }

            using (Bitmap bmp = new Bitmap(w0, h0, PixelFormat.Format32bppArgb))
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.Transparent);
                int n = boxes.Count;
                for (int i = 0; i < n; i++)
                {
                    int a0 = AlphaAt(i, n), a1 = AlphaAt(i + 1, n);
                    foreach (System.Windows.Rect rc in boxes[i].Rects)
                    {
                        RectangleF rf = new RectangleF((float)(rc.Left - x0) - 1, (float)(rc.Top - y0),
                                                       (float)rc.Width + 2, (float)rc.Height);
                        if (rf.Width < 1 || rf.Height < 1) continue;
                        RectangleF brushRect = new RectangleF(rf.X - 1, rf.Y, rf.Width + 2, rf.Height);
                        using (LinearGradientBrush br = new LinearGradientBrush(brushRect,
                                   Color.FromArgb(a0, HighlightColor), Color.FromArgb(a1, HighlightColor),
                                   LinearGradientMode.Horizontal))
                        {
                            g.FillRectangle(br, rf);
                        }
                    }
                }
                overlay.SetBitmap(bmp, x0, y0);
            }
        }
    }

    // ---- Diagnostic log (off by default; WordFocus-log.txt beside the script, max 3000 lines) ----
    public static class Log
    {
        private static readonly object gate = new object();
        private static System.IO.StreamWriter w;
        private static int lines = 0;
        public static bool Enabled = false; // diagnostic log off; set to true to write WordFocus-log.txt again
        public static void Start(string folder)
        {
            if (!Enabled) return;
            try
            {
                w = new System.IO.StreamWriter(System.IO.Path.Combine(folder, "WordFocus-log.txt"), false);
                w.AutoFlush = true;
                Write("Word Focus started. Windows " + Environment.OSVersion.Version);
            }
            catch { w = null; }
        }
        public static void Write(string msg)
        {
            lock (gate)
            {
                if (w == null || lines >= 3000) return;
                lines++;
                try { w.WriteLine(DateTime.Now.ToString("HH:mm:ss.fff") + "  " + msg); } catch { }
            }
        }
    }

    public static class App
    {
        public static void Run(string folder)
        {
            Log.Start(folder);
            Native.SetProcessDPIAware(); // so word positions match the screen exactly
            Application.EnableVisualStyles();
            new Controller();
            Application.Run();
        }
    }
}
'@
    Add-Type -TypeDefinition $code -ReferencedAssemblies $refs
    [WordFocus.App]::Run($PSScriptRoot)
}
catch {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show("Word Focus couldn't start:`n`n" + $_.Exception.Message, "Word Focus") | Out-Null
}
