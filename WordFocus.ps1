# Word Focus - highlights the word under your mouse plus the next few words.
# Required Notice: Copyright 2sage4thyme (https://github.com/2sage4thyme)
# License: Big Time Public License 2.0.2 - https://bigtimelicense.com/versions/2.0.2 (see LICENSE.md)
#   Ctrl+Alt+H  = turn highlight on/off
#   Ctrl+Alt+K  = open settings (word count, ombre, strength, pointer, color, and more)
#   Ctrl+Alt+A  = Paragraph Focus (point at text; press again to turn off)
#   Ctrl+Alt+B  = Page Focus (point at text; press again to turn off)
#   Every action can have its own hotkey: Settings > Advanced > Hotkeys...
#   Tray icon (bottom-right) -> Exit to quit.
# What this script does: compiles the C# code below in memory (using Windows' built-in
# compiler), then runs it. It installs nothing and changes no system settings. The only file
# it writes is your Word Focus settings (%APPDATA%\WordFocus\settings.txt). It
# does not start with Windows unless you tick "Start with Windows" in settings
# (that adds a shortcut to your personal Startup folder; unticking removes it). It only reads the text under your mouse through
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
        [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr hWnd, IntPtr after, int x, int y, int cx, int cy, uint flags);
        [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
        [DllImport("user32.dll")] public static extern IntPtr WindowFromPoint(POINT p);
        [DllImport("user32.dll")] public static extern IntPtr GetAncestor(IntPtr hwnd, uint flags);
        [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
        [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hwnd, out RECT r);
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

        private float opacity;
        public float Opacity { get { return opacity; } }

        // Changes how see-through the pointer is; re-applies immediately if currently faded.
        public void SetOpacity(float o)
        {
            opacity = Math.Max(0f, Math.Min(1f, o));
            if (!faded) return;
            if (opacity >= 0.999f) { RestoreAlways(); return; }
            faded = false;
            Fade();
        }

        public CursorFader(float opacity)
        {
            this.opacity = opacity;
            RestoreAlways(); // in case an earlier run was force-closed while faded
            foreach (int id in Ids)
            {
                try
                {
                    IntPtr h = Native.LoadCursor(IntPtr.Zero, new IntPtr(id));
                    CurData d = Capture(h);
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
            if (faded || data.Count == 0 || opacity >= 0.999f) return; // 100% = leave the normal pointer alone
            foreach (CurData d in data)
            {
                IntPtr c = Create(d, opacity);
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

        private static CurData Capture(IntPtr cur)
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

            CurData d = new CurData();
            d.W = w; d.H = h; d.HotX = ii.xHotspot; d.HotY = ii.yHotspot; d.Pixels = outp;
            return d;
        }

        private static IntPtr Create(CurData d, float opacity)
        {
            Native.BITMAPINFOHEADER bih = new Native.BITMAPINFOHEADER();
            bih.biSize = Marshal.SizeOf(typeof(Native.BITMAPINFOHEADER));
            bih.biWidth = d.W; bih.biHeight = -d.H; bih.biPlanes = 1; bih.biBitCount = 32; bih.biCompression = 0;
            IntPtr bits;
            IntPtr color = Native.CreateDIBSection(IntPtr.Zero, ref bih, 0, out bits, IntPtr.Zero, 0);
            if (color == IntPtr.Zero) return IntPtr.Zero;
            int[] px = new int[d.Pixels.Length];
            for (int i = 0; i < px.Length; i++) // apply the see-through amount
            {
                int a = (int)(((d.Pixels[i] >> 24) & 255) * opacity);
                px[i] = (a << 24) | (d.Pixels[i] & 0xFFFFFF);
            }
            Marshal.Copy(px, 0, bits, px.Length);
            // Mask of all 1s: if every pixel is fully see-through (0%), Windows falls back to this mask,
            // and all 1s means "leave the screen as is", i.e. an invisible pointer (an all-0 mask showed a black box).
            byte[] maskBits = new byte[((d.W + 15) / 16 * 2) * d.H];
            for (int i = 0; i < maskBits.Length; i++) maskBits[i] = 0xFF;
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

        // Other "always on top" windows (e.g. an app in a special on-top mode) can end up above us.
        // Re-claim the top spot each time we draw, without taking focus.
        public void BringToTop()
        {
            // HWND_TOPMOST = -1; NOSIZE | NOMOVE | NOACTIVATE | NOOWNERZORDER
            Native.SetWindowPos(Handle, new IntPtr(-1), 0, 0, 0, 0, 0x0001 | 0x0002 | 0x0010 | 0x0200);
        }

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
                BringToTop();
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
        private readonly CheckBox overlapBox;
        private readonly CheckBox onBox;
        private readonly CheckBox startupBox;
        private readonly TrackBar strength;
        private readonly Label strengthLabel;
        private readonly TrackBar pointer;
        private readonly Label pointerLabel;

        private const int W = 380;   // window width
        private const int CW = 350;  // control width

        private static TrackBar MakeSlider(int min, int max, int tick, int small, int large, int y)
        {
            TrackBar t = new TrackBar();
            t.Minimum = min; t.Maximum = max; t.TickFrequency = tick; t.SmallChange = small; t.LargeChange = large;
            t.SetBounds(10, y, CW + 6, 45);
            return t;
        }

        private Label MakeLabel(int y)
        {
            Label l = new Label();
            l.SetBounds(14, y, CW, 20);
            Controls.Add(l);
            return l;
        }

        public SettingsForm(Controller c)
        {
            ctl = c;
            Text = "Word Focus";
            FormBorderStyle = FormBorderStyle.FixedToolWindow;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(W, 412);
            TopMost = true;

            onBox = new CheckBox();
            onBox.SetBounds(14, 10, CW, 24);
            onBox.CheckedChanged += delegate { if (onBox.Checked != ctl.Enabled) ctl.Toggle(); };
            Controls.Add(onBox);

            // Words: 1-10, then end of sentence, end of line, end of line + next word.
            sliderLabel = MakeLabel(134);
            slider = MakeSlider(1, Controller.MaxCountValue, 1, 1, 1, 152);
            slider.ValueChanged += delegate { ctl.SetCount(slider.Value); UpdateLabel(); };
            Controls.Add(slider);

            ombreBox = new CheckBox();
            ombreBox.Text = "Ombre fade   (off = one solid color)";
            ombreBox.SetBounds(14, 78, CW, 24);
            ombreBox.CheckedChanged += delegate { ctl.SetOmbre(ombreBox.Checked); };
            Controls.Add(ombreBox);

            overlapBox = new CheckBox();
            overlapBox.Text = "Overlap lines   (thin darker lines where words meet)";
            overlapBox.SetBounds(14, 102, CW, 24);
            overlapBox.CheckedChanged += delegate { ctl.SetOverlapLines(overlapBox.Checked); };
            Controls.Add(overlapBox);

            // Highlight strength: 10% = faint, 100% = solid.
            strengthLabel = MakeLabel(198);
            strength = MakeSlider(1, 10, 1, 1, 1, 216); // tens: 10%..100%
            strength.ValueChanged += delegate { ctl.SetStrength(strength.Value * 10); UpdateStrengthLabel(); };
            Controls.Add(strength);

            // Pointer visibility while highlighting: 0% = invisible, 100% = normal.
            pointerLabel = MakeLabel(262);
            pointer = MakeSlider(0, 10, 1, 1, 1, 280); // tens: 0%..100%
            pointer.ValueChanged += delegate { ctl.SetCursorPercent(pointer.Value * 10); UpdatePointerLabel(); };
            Controls.Add(pointer);

            // Color row: preset swatches + a "Custom..." button that opens the Windows color picker.
            Label colorLabel = new Label();
            colorLabel.Text = "Color:";
            colorLabel.SetBounds(14, 46, 42, 20);
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
                sw.SetBounds(60 + i * 30, 42, 26, 26);
                Color picked = presets[i];
                sw.Click += delegate { ctl.SetColor(picked); };
                tips.SetToolTip(sw, names[i]);
                Controls.Add(sw);
            }
            Button custom = new Button();
            custom.Text = "Custom...";
            custom.SetBounds(W - 104, 41, 90, 28);
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

            // Start with Windows: adds/removes a shortcut in your personal Startup folder.
            startupBox = new CheckBox();
            startupBox.Text = "Start with Windows   (uses your last settings)";
            startupBox.SetBounds(14, 332, CW, 24);
            startupBox.Checked = Startup.IsEnabled();
            startupBox.CheckedChanged += delegate
            {
                if (startupBox.Checked == Startup.IsEnabled()) return; // just syncing the box, nothing to change
                string err = Startup.Set(startupBox.Checked);
                if (err != null)
                {
                    MessageBox.Show("Couldn't change the Startup shortcut:\n" + err, "Word Focus");
                    startupBox.Checked = Startup.IsEnabled();
                }
            };
            Controls.Add(startupBox);

            Label hint = new Label();
            hint.Text = "Closing this window keeps Word Focus running.";
            hint.ForeColor = SystemColors.GrayText;
            hint.SetBounds(14, 378, W - 130, 20);
            Controls.Add(hint);

            Button adv = new Button();
            adv.Text = "Advanced...";
            adv.SetBounds(W - 104, 372, 90, 28);
            adv.Click += delegate { ctl.ShowAdvanced(); };
            Controls.Add(adv);

            if (Controller.AppIcon != null) Icon = Controller.AppIcon;
            SyncFromController();
        }

        public void SyncFromController()
        {
            onBox.Checked = ctl.Enabled;
            startupBox.Checked = Startup.IsEnabled();
            onBox.Text = "Highlight on   (" + ctl.KeyText("toggle") + ")";
            slider.Value = ctl.Count;
            ombreBox.Checked = ctl.Ombre;
            overlapBox.Checked = ctl.OverlapLines;
            strength.Value = Math.Max(1, Math.Min(10, (ctl.StrengthPercent + 5) / 10));
            pointer.Value = Math.Max(0, Math.Min(10, (ctl.CursorPercent + 5) / 10));
            UpdateStrengthLabel();
            UpdatePointerLabel();
            UpdateLabel();
        }

        private void UpdateStrengthLabel() { strengthLabel.Text = "Highlight strength: " + (strength.Value * 10) + "%"; }

        private void UpdatePointerLabel()
        {
            int v = pointer.Value * 10;
            pointerLabel.Text = "Mouse pointer while over text: " + v + "% visible" +
                (v == 0 ? " (invisible)" : v == 100 ? " (normal)" : "");
        }

        private void UpdateLabel() { sliderLabel.Text = "Words highlighted: " + Controller.CountText(slider.Value); }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing) { e.Cancel = true; Hide(); }
            base.OnFormClosing(e);
        }
    }

    public interface IKeyHost { void ReleaseFocus(); void KeysChanged(); }

    // ---- A box that records a key combination when clicked ----
    public class HotkeyBox : TextBox
    {
        private readonly Controller ctl;
        private readonly int index;
        private readonly Label status;

        public HotkeyBox(Controller c, int i, Label statusLabel)
        {
            ctl = c; index = i; status = statusLabel;
            ReadOnly = true;
            BackColor = SystemColors.Window;
            ShortcutsEnabled = false;
            Cursor = Cursors.Hand;
        }

        public void ShowCurrent() { Text = ctl.KeyText(index); }

        protected override void OnEnter(EventArgs e)
        {
            ctl.SuspendHotkeys();
            Text = "Press new keys...";
            status.ForeColor = SystemColors.GrayText;
            status.Text = "Hold Ctrl and/or Alt, then press a key. Backspace = none, Esc = cancel.";
            base.OnEnter(e);
        }

        protected override void OnLeave(EventArgs e)
        {
            ctl.ResumeHotkeys();
            ShowCurrent();
            base.OnLeave(e);
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            Keys key = keyData & Keys.KeyCode;
            Keys mods = keyData & Keys.Modifiers;
            if (key == Keys.Tab && mods == Keys.None) return base.ProcessCmdKey(ref msg, keyData);
            if (key == Keys.ShiftKey || key == Keys.ControlKey || key == Keys.Menu || key == Keys.LWin || key == Keys.RWin)
                return true; // still waiting for the main key
            if (key == Keys.Escape && mods == Keys.None) { ShowCurrent(); Done(null, "Cancelled."); return true; }

            uint m = 0;
            if ((mods & Keys.Control) != 0) m |= 2;
            if ((mods & Keys.Alt) != 0) m |= 1;
            if ((mods & Keys.Shift) != 0) m |= 4;

            string err;
            if ((key == Keys.Back || key == Keys.Delete) && mods == Keys.None) err = ctl.SetHotkey(index, 0, 0);
            else if ((m & 3) == 0) err = "Include Ctrl or Alt, so the hotkey doesn't block normal typing.";
            else err = ctl.SetHotkey(index, m, (uint)key);
            ShowCurrent();
            Done(err, null);
            return true;
        }

        private void Done(string err, string note)
        {
            status.ForeColor = err == null ? SystemColors.GrayText : Color.Firebrick;
            string warn = err == null && note == null ? ctl.KeyWarning(index) : null;
            status.Text = err ?? note ?? ("Saved: " + Controller.KeyNames[index] + " = " + ctl.KeyText(index) +
                                          (warn != null ? "  (see the red note)" : ""));
            IKeyHost f = FindForm() as IKeyHost;
            if (f != null) { f.ReleaseFocus(); f.KeysChanged(); } // stop recording (turns our hotkeys back on), re-check warnings
        }
    }

    // ---- Advanced settings: style + hotkeys ----
    public class AdvancedForm : Form
    {
        private readonly Controller ctl;
        private readonly RadioButton hlRadio, ulRadio, bothRadio;
        private readonly RadioButton fHlRadio, fUlRadio, fBothRadio;
        private readonly TrackBar fStrength;
        private readonly Label fStrengthLabel;

        private const int W = 360;

        private Label Heading(string text, int y)
        {
            Label l = new Label();
            l.Text = text;
            l.Font = new Font(Font, FontStyle.Bold);
            l.SetBounds(14, y, W - 28, 20);
            Controls.Add(l);
            return l;
        }

        // Three radio buttons in their own panel, so each group works independently.
        private Panel StyleRow(int y, out RadioButton a, out RadioButton b, out RadioButton c)
        {
            Panel p = new Panel();
            p.SetBounds(14, y, W - 28, 26);
            a = new RadioButton(); a.Text = "Highlight"; a.SetBounds(6, 0, 90, 24);
            b = new RadioButton(); b.Text = "Underline"; b.SetBounds(101, 0, 90, 24);
            c = new RadioButton(); c.Text = "Both"; c.SetBounds(196, 0, 70, 24);
            p.Controls.Add(a); p.Controls.Add(b); p.Controls.Add(c);
            Controls.Add(p);
            return p;
        }

        public AdvancedForm(Controller c)
        {
            ctl = c;
            Text = "Word Focus - Advanced";
            FormBorderStyle = FormBorderStyle.FixedToolWindow;
            StartPosition = FormStartPosition.CenterScreen;
            TopMost = true;
            if (Controller.AppIcon != null) Icon = Controller.AppIcon;
            ToolTip tips = new ToolTip();

            // --- Hover style ---
            Heading("Hover style", 10);
            StyleRow(32, out hlRadio, out ulRadio, out bothRadio);
            hlRadio.CheckedChanged += delegate { if (hlRadio.Checked) ctl.SetStyle(Controller.StyleHighlight); };
            ulRadio.CheckedChanged += delegate { if (ulRadio.Checked) ctl.SetStyle(Controller.StyleUnderline); };
            bothRadio.CheckedChanged += delegate { if (bothRadio.Checked) ctl.SetStyle(Controller.StyleBoth); };
            tips.SetToolTip(bothRadio, "Highlight in your color, plus an underline that is automatically white on dark pages and black on light pages.");

            // --- Paragraph & Page Focus ---
            Heading("Paragraph / Page Focus", 68);
            Label fHint = new Label();
            fHint.Text = "Point at text and press the hotkey to light up the whole paragraph (or the visible page). Press it again to turn it off.";
            fHint.ForeColor = SystemColors.GrayText;
            fHint.SetBounds(14, 90, W - 28, 32);
            Controls.Add(fHint);
            StyleRow(124, out fHlRadio, out fUlRadio, out fBothRadio);
            fHlRadio.CheckedChanged += delegate { if (fHlRadio.Checked) ctl.SetFocusStyle(Controller.StyleHighlight); };
            fUlRadio.CheckedChanged += delegate { if (fUlRadio.Checked) ctl.SetFocusStyle(Controller.StyleUnderline); };
            fBothRadio.CheckedChanged += delegate { if (fBothRadio.Checked) ctl.SetFocusStyle(Controller.StyleBoth); };
            fStrengthLabel = new Label();
            fStrengthLabel.SetBounds(14, 154, W - 28, 20);
            Controls.Add(fStrengthLabel);
            fStrength = new TrackBar();
            fStrength.Minimum = 1; fStrength.Maximum = 10; fStrength.TickFrequency = 1;
            fStrength.SmallChange = 1; fStrength.LargeChange = 1;
            fStrength.SetBounds(10, 172, W - 20, 45);
            fStrength.ValueChanged += delegate { ctl.SetFocusStrength(fStrength.Value * 10); UpdateFocusLabel(); };
            Controls.Add(fStrength);

            // --- Hotkeys: their own window ---
            Heading("Hotkeys", 222);
            Label kHint = new Label();
            kHint.Text = "Set a hotkey for anything Word Focus can do.";
            kHint.ForeColor = SystemColors.GrayText;
            kHint.SetBounds(14, 246, W - 130, 32);
            Controls.Add(kHint);
            Button keysBtn = new Button();
            keysBtn.Text = "Hotkeys...";
            keysBtn.SetBounds(W - 104, 242, 90, 28);
            keysBtn.Click += delegate { ctl.ShowHotkeys(); };
            Controls.Add(keysBtn);

            ClientSize = new Size(W, 286);
        }

        private void UpdateFocusLabel() { fStrengthLabel.Text = "Focus area strength: " + (fStrength.Value * 10) + "%"; }

        public void SyncFromController()
        {
            hlRadio.Checked = ctl.Style == Controller.StyleHighlight;
            ulRadio.Checked = ctl.Style == Controller.StyleUnderline;
            bothRadio.Checked = ctl.Style == Controller.StyleBoth;
            fHlRadio.Checked = ctl.FocusStyle == Controller.StyleHighlight;
            fUlRadio.Checked = ctl.FocusStyle == Controller.StyleUnderline;
            fBothRadio.Checked = ctl.FocusStyle == Controller.StyleBoth;
            fStrength.Value = Math.Max(1, Math.Min(10, (ctl.FocusStrength + 5) / 10));
            UpdateFocusLabel();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing) { e.Cancel = true; Hide(); }
            base.OnFormClosing(e);
        }
    }

    // ---- Hotkeys menu: one row per action, scrollable. Problem keys turn red with a note above them. ----
    public class HotkeysForm : Form, IKeyHost
    {
        private readonly Controller ctl;
        private readonly HotkeyBox[] boxes;
        private readonly Label[] names;
        private readonly Label[] notes;
        private readonly Panel list;
        private readonly Label status;
        private readonly Button resetBtn;
        private const int W = 440;
        private static readonly Color WarnBack = Color.FromArgb(255, 205, 205);
        private static readonly Color WarnText = Color.FromArgb(160, 20, 20);

        public HotkeysForm(Controller c)
        {
            ctl = c;
            Text = "Word Focus - Hotkeys";
            FormBorderStyle = FormBorderStyle.FixedToolWindow;
            StartPosition = FormStartPosition.CenterScreen;
            TopMost = true;
            if (Controller.AppIcon != null) Icon = Controller.AppIcon;

            Label hint = new Label();
            hint.Text = "Click a box, then press the new keys (hold Ctrl and/or Alt). Backspace = none, Esc = cancel.";
            hint.ForeColor = SystemColors.GrayText;
            hint.SetBounds(14, 10, W - 28, 32);
            Controls.Add(hint);

            status = new Label();
            status.ForeColor = SystemColors.GrayText;

            list = new Panel();
            list.AutoScroll = true;
            list.SetBounds(8, 46, W - 16, 420);
            list.BorderStyle = BorderStyle.FixedSingle;
            Controls.Add(list);

            int n = Controller.KeyIds.Length;
            boxes = new HotkeyBox[n];
            names = new Label[n];
            notes = new Label[n];
            for (int i = 0; i < n; i++)
            {
                notes[i] = new Label();
                notes[i].ForeColor = WarnText;
                notes[i].Visible = false;
                notes[i].Cursor = Cursors.Hand;
                int which = i;
                notes[i].Click += delegate { ctl.DismissNote(which); Relayout(); };
                list.Controls.Add(notes[i]);
                names[i] = new Label();
                names[i].Text = Controller.KeyNames[i];
                list.Controls.Add(names[i]);
                boxes[i] = new HotkeyBox(ctl, i, status);
                list.Controls.Add(boxes[i]);
            }

            status.SetBounds(14, 472, W - 28, 32);
            Controls.Add(status);

            resetBtn = new Button();
            resetBtn.Text = "Reset hotkeys to defaults";
            resetBtn.SetBounds(14, 508, 180, 28);
            resetBtn.Click += delegate
            {
                ctl.ResetHotkeys();
                SyncFromController();
                status.ForeColor = SystemColors.GrayText;
                status.Text = "Hotkeys reset: Ctrl+Alt+H, Ctrl+Alt+K, Ctrl+Alt+A, Ctrl+Alt+B; everything else none.";
            };
            Controls.Add(resetBtn);
            ClientSize = new Size(W, 546);
            Relayout();

            // Make sure our hotkeys come back if this window loses focus while a box is recording.
            Deactivate += delegate { ActiveControl = resetBtn; ctl.ResumeHotkeys(); };
        }

        // Places each row, with its red note (if any) just above it.
        private void Relayout()
        {
            int scroll = list.VerticalScroll.Value;
            list.SuspendLayout();
            list.AutoScrollPosition = new Point(0, 0);
            int innerW = W - 16 - 26;
            int y = 6;
            for (int i = 0; i < boxes.Length; i++)
            {
                string warn = ctl.KeyWarning(i);
                bool showNote = warn != null && !ctl.NoteDismissed(i);
                if (showNote)
                {
                    warn += "  (click to hide)";
                    Size sz = TextRenderer.MeasureText(warn, notes[i].Font, new Size(innerW - 8, 1000), TextFormatFlags.WordBreak);
                    notes[i].Text = warn;
                    notes[i].SetBounds(8, y, innerW - 8, sz.Height + 2);
                    notes[i].Visible = true;
                    y += sz.Height + 4;
                    boxes[i].BackColor = WarnBack;
                }
                else
                {
                    notes[i].Visible = false;
                    boxes[i].BackColor = warn != null ? WarnBack : SystemColors.Window; // still red as a reminder
                }
                names[i].SetBounds(8, y + 3, 230, 20);
                boxes[i].SetBounds(244, y, innerW - 244, 22);
                y += 30;
            }
            list.ResumeLayout();
            list.AutoScrollPosition = new Point(0, scroll);
        }

        public void ReleaseFocus() { ActiveControl = resetBtn; }
        public void KeysChanged() { Relayout(); }

        public void SyncKeysOnly()
        {
            for (int i = 0; i < boxes.Length; i++) if (!boxes[i].Focused) boxes[i].ShowCurrent();
            Relayout();
        }

        public void SyncFromController()
        {
            for (int i = 0; i < boxes.Length; i++) boxes[i].ShowCurrent();
            Relayout();
            ActiveControl = resetBtn;
            status.ForeColor = SystemColors.GrayText;
            status.Text = "Keys showing \"none\" are off until you set them. Red = a conflict or a risky choice.";
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing) { e.Cancel = true; ctl.ResumeHotkeys(); Hide(); }
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
        public const int DefaultStrengthPercent = 60; // highlight strength at start-up (10-100)
        public const float FadeTo = 0.15f;   // last word fades to this fraction of the first word's strength
        // Slider positions after 1-10 words:
        public const int SentenceValue = 11;  // to the end of the sentence
        public const int LineValue = 12;      // to the end of the line
        public const int LinePlusValue = 13;  // end of the line + first word of the next line
        public const int MaxCountValue = 13;
        public const int DefaultCursorPercent = 10; // pointer visibility while highlighting (0 = invisible, 100 = normal)
        public const int CursorRestoreDelayMs = 700; // keeps the pointer faded while moving across gaps between words
        // ====================================

        private volatile bool enabled = false;
        private volatile bool ombre = true;
        private volatile int strengthPercent = DefaultStrengthPercent;
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

        // ---- Style ----
        // 0 = highlight, 1 = underline (in your color), 2 = both (highlight in your color + automatic black/white underline)
        public const int StyleHighlight = 0, StyleUnderline = 1, StyleBoth = 2;
        private volatile int style = StyleHighlight;
        public int Style { get { return style; } }
        public void SetStyle(int v) { style = Math.Max(0, Math.Min(2, v)); dirty = true; settingsChanged = true; }

        // ---- Paragraph / Page Focus: a steady layer over a whole paragraph or the visible page ----
        public const int FocusNone = 0, FocusParagraph = 1, FocusPage = 2, FocusClear = 3;
        private volatile int focusStyle = StyleHighlight;
        private volatile int focusStrength = 10;
        public int FocusStyle { get { return focusStyle; } }
        public int FocusStrength { get { return focusStrength; } }
        public void SetFocusStyle(int v) { focusStyle = Math.Max(0, Math.Min(2, v)); focusDirty = true; settingsChanged = true; }
        public void SetFocusStrength(int v) { focusStrength = Math.Max(10, Math.Min(100, v)); focusDirty = true; settingsChanged = true; }

        private static readonly uint MyPid = (uint)System.Diagnostics.Process.GetCurrentProcess().Id;
        private Overlay focusOverlay;
        private volatile int focusRequest = FocusNone;    // set by hotkeys (UI thread), handled by the worker
        private Native.POINT focusRequestPt;
        private volatile int focusKind = FocusNone;       // what is showing now
        private volatile bool focusDirty = false;         // redraw even if the text hasn't moved
        // Worker-thread-only:
        private List<TextPatternRange> focusRanges; // one range for a paragraph; one per run of body lines for a page
        private IntPtr focusRoot = IntPtr.Zero;
        private string focusSig = "";
        private int focusRefreshTick = 0;
        private int focusSlowPenalty = 0;
        private string lastPagePos = "";

        // Same hotkey again (or the other focus hotkey on the same kind) turns it off; otherwise start/switch.
        private void RequestFocus(int kind)
        {
            Native.POINT p;
            Native.GetCursorPos(out p);
            focusRequestPt = p;
            focusRequest = (focusKind == kind) ? FocusClear : kind;
        }

        // ---- Custom icon: WordFocus.ico beside the script, if present ----
        public static Icon AppIcon;

        // ---- Hotkeys (index 0..3 -> hotkey id 1..4) ----
        // Every action Word Focus can do, in the order shown in the Hotkeys menu.
        // Ids are stable (they're what gets saved); names are what you see.
        public static readonly string[] KeyIds = new string[] {
            "toggle", "settings", "more", "fewer", "nextstyle", "ombre", "overlap",
            "stronger", "weaker", "nextcolor", "prevcolor", "pointermore", "pointerless",
            "paragraph", "page", "fstronger", "fweaker",
            "advanced", "hotkeys", "startup", "exit" };
        public static readonly string[] KeyNames = new string[] {
            "Highlight on/off", "Open settings", "More words", "Fewer words", "Next style (highlight / underline / both)",
            "Ombre on/off", "Overlap lines on/off", "Highlight stronger", "Highlight weaker", "Next color", "Previous color",
            "Mouse pointer more visible", "Mouse pointer less visible",
            "Paragraph Focus on/off", "Page Focus on/off", "Focus layer stronger", "Focus layer weaker",
            "Open advanced settings", "Open hotkeys menu", "Start with Windows on/off", "Exit Word Focus" };
        public static int KeyIndex(string id) { return Array.IndexOf(KeyIds, id); }
        // Defaults: Ctrl+Alt+H and Ctrl+Alt+K. Word-count keys are off until set in Settings > Advanced.
        // Paragraph Focus defaults to Ctrl+Alt+A; Page Focus is off until set.
        // Defaults: Ctrl+Alt+H (on/off), Ctrl+Alt+K (settings), Ctrl+Alt+A (Paragraph Focus), Ctrl+Alt+B (Page Focus). Everything else starts as "none".
        private static readonly uint[] DefMods = MakeDefaults(true);
        private static readonly uint[] DefVk = MakeDefaults(false);
        private readonly uint[] keyMods = MakeDefaults(true);
        private readonly uint[] keyVk = MakeDefaults(false);
        private static uint[] MakeDefaults(bool mods)
        {
            uint[] a = new uint[KeyIds.Length];
            string[] ids = new string[] { "toggle", "settings", "paragraph", "page" };
            uint[] vks = new uint[] { 0x48, 0x4B, 0x41, 0x42 }; // H, K, A, B
            for (int i = 0; i < ids.Length; i++) a[Array.IndexOf(KeyIds, ids[i])] = mods ? 3u : vks[i];
            return a;
        }
        public int KeyCount { get { return keyVk.Length; } }
        private bool hotkeysSuspended = false;
        private ToolStripMenuItem toggleItem, settingsItem;
        private AdvancedForm advanced;

        public static string FormatKey(uint mods, uint vk)
        {
            if (vk == 0) return "none";
            string t = "";
            if ((mods & 2) != 0) t += "Ctrl+";
            if ((mods & 1) != 0) t += "Alt+";
            if ((mods & 4) != 0) t += "Shift+";
            Keys k = (Keys)vk;
            string name = k.ToString();
            if (k >= Keys.D0 && k <= Keys.D9) name = ((int)(k - Keys.D0)).ToString();
            else if (k == Keys.Up) name = "Up arrow";
            else if (k == Keys.Down) name = "Down arrow";
            else if (k == Keys.Left) name = "Left arrow";
            else if (k == Keys.Right) name = "Right arrow";
            return t + name;
        }

        public string KeyText(int i) { return FormatKey(keyMods[i], keyVk[i]); }
        public string KeyText(string id) { return KeyText(KeyIndex(id)); }

        // Registers every hotkey; returns the names of any that another program already uses.
        private List<string> RegisterAll()
        {
            List<string> failed = new List<string>();
            for (int i = 0; i < keyVk.Length; i++) Native.UnregisterHotKey(hotkeys.Handle, i + 1);
            for (int i = 0; i < keyVk.Length; i++)
            {
                keyTaken[i] = false;
                if (keyVk[i] == 0 || DuplicateOf(i) >= 0 && DuplicateOf(i) < i) continue; // a duplicate: the first one wins
                if (!Native.RegisterHotKey(hotkeys.Handle, i + 1, keyMods[i] | 0x4000, keyVk[i]))
                {
                    keyTaken[i] = true;
                    failed.Add(KeyText(i) + " (" + KeyNames[i] + ")");
                }
            }
            return failed;
        }

        private void UnregisterAll()
        {
            for (int i = 0; i < keyVk.Length; i++) Native.UnregisterHotKey(hotkeys.Handle, i + 1);
        }

        // While a hotkey box is recording, our own hotkeys are paused so the keys reach the box.
        public void SuspendHotkeys() { if (!hotkeysSuspended) { UnregisterAll(); hotkeysSuspended = true; } }
        public void ResumeHotkeys() { if (hotkeysSuspended) { hotkeysSuspended = false; RegisterAll(); } }

        // Returns null if OK, or a message explaining why the key can't be used.
        public string SetHotkey(int i, uint mods, uint vk)
        {
            keyTaken[i] = false;
            if (vk != 0)
            {
                // Test whether another program owns it (our own keys are paused while recording).
                if (Native.RegisterHotKey(hotkeys.Handle, 99, mods | 0x4000, vk)) Native.UnregisterHotKey(hotkeys.Handle, 99);
                else keyTaken[i] = true;
            }
            keyMods[i] = mods; keyVk[i] = vk;
            settingsChanged = true;
            if (!hotkeysSuspended) RegisterAll();
            UpdateKeyLabels();
            return null;
        }

        private readonly bool[] keyTaken = new bool[KeyIds.Length]; // true = another program owns it

        // Notes you've clicked away. Remembered with the exact key, so a new key brings its note back.
        private readonly Dictionary<string, string> dismissedNotes = new Dictionary<string, string>();
        private string KeySig(int i) { return keyMods[i] + "," + keyVk[i]; }
        public bool NoteDismissed(int i)
        {
            string v;
            return dismissedNotes.TryGetValue(KeyIds[i], out v) && v == KeySig(i);
        }
        public void DismissNote(int i) { dismissedNotes[KeyIds[i]] = KeySig(i); settingsChanged = true; }

        // Index of another action using the same key, or -1.
        public int DuplicateOf(int i)
        {
            if (keyVk[i] == 0) return -1;
            for (int j = 0; j < keyVk.Length; j++)
                if (j != i && keyVk[j] == keyVk[i] && keyMods[j] == keyMods[i]) return j;
            return -1;
        }

        // A short explanation if this hotkey is a conflict or a bad idea, otherwise null.
        public string KeyWarning(int i)
        {
            uint m = keyMods[i], vk = keyVk[i];
            if (vk == 0) return null;
            string k = FormatKey(m, vk);
            int dup = DuplicateOf(i);
            if (dup >= 0)
                return k + " is also set for \"" + KeyNames[dup] + "\". Only one of them can work - pick a different key for one.";
            if (keyTaken[i])
                return "Another program already uses " + k + ", so it won't work here. Try a different letter with Ctrl+Alt.";
            return Advice(m, vk);
        }

        // Combinations that technically work but cause trouble elsewhere.
        private static string Advice(uint m, uint vk)
        {
            bool ctrl = (m & 2) != 0, alt = (m & 1) != 0, shift = (m & 4) != 0;
            Keys key = (Keys)vk;
            string k = FormatKey(m, vk);
            string tip = " Try Ctrl+Alt + a letter (or Ctrl+Alt+Shift + a letter) instead.";
            if (ctrl && alt && key == Keys.Delete) return "Windows reserves Ctrl+Alt+Delete." + tip;
            if (ctrl && shift && !alt && key == Keys.Escape) return "Windows uses Ctrl+Shift+Esc to open Task Manager." + tip;
            if (alt && !ctrl && (key == Keys.F4 || key == Keys.Tab || key == Keys.Escape || key == Keys.Space))
                return "Windows uses " + k + " (closing or switching windows)." + tip;
            if (ctrl && !alt && key == Keys.Escape) return "Windows uses Ctrl+Esc to open the Start menu." + tip;
            if (ctrl && alt && (key == Keys.Up || key == Keys.Down || key == Keys.Left || key == Keys.Right))
                return "On some PCs, Ctrl+Alt + arrow keys rotate the screen. It may be fine on yours, but be careful sharing this setup.";
            if (ctrl && !alt && !shift)
                return "Apps use Ctrl + one key for everyday things (copy, paste, save, undo...). " + k + " would stop doing that everywhere while Word Focus runs." + tip;
            if (alt && !ctrl && !shift)
                return "Alt + one key opens menus in many apps, so " + k + " would stop doing that while Word Focus runs." + tip;
            if (ctrl && shift && !alt)
                return "Browsers and apps use many Ctrl+Shift shortcuts (for example Ctrl+Shift+T reopens a tab)." + tip;
            if (alt && shift && !ctrl)
                return "Windows uses Alt+Shift to switch keyboard languages on some PCs, and apps use some Alt+Shift shortcuts." + tip;
            if (ctrl && alt && !shift && (key == Keys.E || key == Keys.Q || (key >= Keys.D0 && key <= Keys.D9)))
                return "On many non-US keyboards, " + k + " types a character (like the euro sign or @), so it would stop typing that." + tip.Replace("a letter", "another letter");
            return null;
        }

        public void ResetHotkeys()
        {
            for (int i = 0; i < keyVk.Length; i++) { keyMods[i] = DefMods[i]; keyVk[i] = DefVk[i]; }
            settingsChanged = true;
            if (!hotkeysSuspended) RegisterAll();
            UpdateKeyLabels();
        }

        private void UpdateKeyLabels()
        {
            if (toggleItem != null) toggleItem.Text = "Highlight on/off  (" + KeyText("toggle") + ")";
            if (settingsItem != null) settingsItem.Text = "Settings...  (" + KeyText("settings") + ")";
            if (settings != null) settings.SyncFromController();
            if (hotkeysForm != null) hotkeysForm.SyncKeysOnly();
        }

        private HotkeysForm hotkeysForm;
        public void ShowHotkeys()
        {
            if (hotkeysForm == null) hotkeysForm = new HotkeysForm(this);
            hotkeysForm.SyncFromController();
            hotkeysForm.Show();
            hotkeysForm.Activate();
        }

        public void ShowAdvanced()
        {
            if (advanced == null) advanced = new AdvancedForm(this);
            advanced.SyncFromController();
            advanced.Show();
            advanced.Activate();
        }
        public int Count { get { return count; } }

        // Overlap lines: where neighbouring words overlap, the highlight doubles up into a thin line.
        private volatile bool overlapLines = true;
        public bool OverlapLines { get { return overlapLines; } }
        public void SetOverlapLines(bool v) { overlapLines = v; dirty = true; settingsChanged = true; }

        // Pointer visibility while highlighting.
        private volatile int cursorPercent = DefaultCursorPercent;
        public int CursorPercent { get { return cursorPercent; } }
        public void SetCursorPercent(int v)
        {
            cursorPercent = Math.Max(0, Math.Min(100, v));
            settingsChanged = true;
            try { if (fader != null) fader.SetOpacity(cursorPercent / 100f); } catch { }
        }

        public static string CountText(int v)
        {
            if (v == SentenceValue) return "to the end of the sentence";
            if (v == LineValue) return "to the end of the line";
            if (v == LinePlusValue) return "end of the line + first word of the next";
            if (v == 1) return "1 (just the hovered word)";
            return v + " (hovered word + next " + (v - 1) + ")";
        }

        public Controller()
        {
            try { fader = new CursorFader(cursorPercent / 100f); } catch { fader = null; }
            AppDomain.CurrentDomain.ProcessExit += delegate { RestoreCursor(); };
            AppDomain.CurrentDomain.UnhandledException += delegate { RestoreCursor(); };
            Application.ThreadException += delegate { RestoreCursor(); };

            LoadSettings();

            focusOverlay = new Overlay();   // created first, so the hover highlight sits on top of it
            focusOverlay.Show();
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

            // Saves changed settings about once a second (so dragging a slider doesn't write constantly).
            System.Windows.Forms.Timer saveTimer = new System.Windows.Forms.Timer();
            saveTimer.Interval = 1000;
            saveTimer.Tick += delegate { if (settingsChanged) SaveSettings(); };
            saveTimer.Start();

            hotkeys = new HotkeyWindow();
            hotkeys.Pressed += OnHotkey;
            List<string> failed = RegisterAll();

            try
            {
                string ico = System.IO.Path.Combine(Startup.Folder ?? "", "WordFocus.ico");
                if (System.IO.File.Exists(ico)) AppIcon = new Icon(ico);
            }
            catch { AppIcon = null; }

            tray = new NotifyIcon();
            tray.Icon = AppIcon ?? SystemIcons.Application;
            tray.Visible = true;
            ContextMenuStrip menu = new ContextMenuStrip();
            toggleItem = (ToolStripMenuItem)menu.Items.Add("Highlight on/off", null, delegate { Toggle(); });
            settingsItem = (ToolStripMenuItem)menu.Items.Add("Settings...", null, delegate { ShowSettings(); });
            menu.Items.Add("Exit", null, delegate { Exit(); });
            tray.ContextMenuStrip = menu;
            tray.DoubleClick += delegate { ShowSettings(); };
            UpdateKeyLabels();
            UpdateTrayText();

            if (failed.Count > 0)
                MessageBox.Show("Another program is already using these hotkeys, so they won't work:\n\n" +
                                string.Join("\n", failed.ToArray()) +
                                "\n\nYou can pick different keys in Settings > Advanced, or use the tray icon by the clock.", "Word Focus");

            Thread worker = new Thread(Loop);
            worker.IsBackground = true;
            worker.SetApartmentState(ApartmentState.MTA);
            worker.Start();

            // Separate worker for Paragraph/Page Focus, so a slow app can only delay the Focus layer.
            Thread focusWorker = new Thread(FocusLoop);
            focusWorker.IsBackground = true;
            focusWorker.SetApartmentState(ApartmentState.MTA);
            focusWorker.Start();
        }

        // Preset colors (same as the swatches in Settings), used by Next / Previous color.
        public static readonly Color[] Presets = new Color[] {
            Color.FromArgb(255, 214, 0), Color.FromArgb(120, 220, 80), Color.FromArgb(80, 170, 255),
            Color.FromArgb(255, 120, 190), Color.FromArgb(255, 150, 40), Color.FromArgb(170, 120, 255) };

        private void StepColor(int dir)
        {
            Color c = HighlightColor;
            int at = -1;
            for (int i = 0; i < Presets.Length; i++)
                if (Presets[i].R == c.R && Presets[i].G == c.G && Presets[i].B == c.B) { at = i; break; }
            int next = at < 0 ? (dir > 0 ? 0 : Presets.Length - 1) : (at + dir + Presets.Length) % Presets.Length;
            SetColor(Presets[next]);
        }

        public void ToggleStartup()
        {
            string err = Startup.Set(!Startup.IsEnabled());
            if (err != null) MessageBox.Show("Couldn't change the Startup shortcut:\n" + err, "Word Focus");
        }

        private void SyncWindows()
        {
            if (settings != null) settings.SyncFromController();
            if (advanced != null) advanced.SyncFromController();
        }

        private void OnHotkey(int id)
        {
            if (id < 1 || id > KeyIds.Length) return;
            switch (KeyIds[id - 1])
            {
                case "toggle": Toggle(); return;
                case "settings": ShowSettings(); return;
                case "more": SetCount(count + 1); break;
                case "fewer": SetCount(count - 1); break;
                case "nextstyle": SetStyle((style + 1) % 3); break;
                case "ombre": SetOmbre(!ombre); break;
                case "overlap": SetOverlapLines(!overlapLines); break;
                case "stronger": SetStrength(strengthPercent + 10); break;
                case "weaker": SetStrength(strengthPercent - 10); break;
                case "nextcolor": StepColor(1); break;
                case "prevcolor": StepColor(-1); break;
                case "pointermore": SetCursorPercent(cursorPercent + 10); break;
                case "pointerless": SetCursorPercent(cursorPercent - 10); break;
                case "paragraph": RequestFocus(FocusParagraph); return;
                case "page": RequestFocus(FocusPage); return;
                case "fstronger": SetFocusStrength(focusStrength + 10); break;
                case "fweaker": SetFocusStrength(focusStrength - 10); break;
                case "advanced": ShowAdvanced(); return;
                case "hotkeys": ShowHotkeys(); return;
                case "startup": ToggleStartup(); break;
                case "exit": Exit(); return;
            }
            SyncWindows();
        }

        public void Toggle()
        {
            enabled = !enabled;
            dirty = true;
            settingsChanged = true;
            if (!enabled) { Clear(); RestoreCursor(); }
            UpdateTrayText();
            if (settings != null) settings.SyncFromController();
        }

        public void SetCount(int v) { count = Math.Max(1, Math.Min(MaxCountValue, v)); dirty = true; settingsChanged = true; }
        public void SetOmbre(bool v) { ombre = v; dirty = true; settingsChanged = true; }
        public void SetStrength(int percent) { strengthPercent = Math.Max(10, Math.Min(100, percent)); dirty = true; settingsChanged = true; }
        public int StrengthPercent { get { return strengthPercent; } }
        public void SetColor(Color c) { HighlightColor = Color.FromArgb(255, c.R, c.G, c.B); dirty = true; focusDirty = true; settingsChanged = true; }

        // ---- Remembered settings: %APPDATA%\WordFocus\settings.txt (plain text, one setting per line) ----
        private volatile bool settingsChanged = false;

        private static string SettingsPath()
        {
            return System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "WordFocus", "settings.txt");
        }

        private void LoadSettings()
        {
            try
            {
                string path = SettingsPath();
                if (!System.IO.File.Exists(path)) return;
                foreach (string line in System.IO.File.ReadAllLines(path))
                {
                    int eq = line.IndexOf('=');
                    if (eq <= 0) continue;
                    string k = line.Substring(0, eq).Trim().ToLowerInvariant();
                    string v = line.Substring(eq + 1).Trim();
                    int n;
                    if (k == "on") enabled = v == "1";
                    else if (k == "words")
                    {
                        if (v == "sentence") count = SentenceValue;
                        else if (v == "line") count = LineValue;
                        else if (v == "lineplus") count = LinePlusValue;
                        else if (int.TryParse(v, out n)) count = n >= 11 ? LineValue : Math.Max(1, n); // 11 meant "line" in older versions
                    }
                    else if (k == "overlap") overlapLines = v == "1";
                    else if (k == "pointer" && int.TryParse(v, out n)) cursorPercent = Math.Max(0, Math.Min(100, n));
                    else if (k == "strength" && int.TryParse(v, out n)) strengthPercent = Math.Max(10, Math.Min(100, n));
                    else if (k == "ombre") ombre = v == "1";
                    else if (k == "fstyle" && int.TryParse(v, out n)) focusStyle = Math.Max(0, Math.Min(2, n));
                    else if (k == "fstrength" && int.TryParse(v, out n)) focusStrength = Math.Max(10, Math.Min(100, n));
                    else if (k == "style") style = v == "underline" ? StyleUnderline : v == "both" ? StyleBoth : StyleHighlight;
                    else if (k.StartsWith("noteok.")) dismissedNotes[k.Substring(7)] = v;
                    else if (k.StartsWith("key"))
                    {
                        // New form: key.<id>=mods,vk. Older versions saved key1..key6 by position.
                        string id = null;
                        if (k.StartsWith("key.")) id = k.Substring(4);
                        else
                        {
                            string[] legacy = new string[] { "toggle", "settings", "more", "fewer", "paragraph", "page" };
                            int pos;
                            if (int.TryParse(k.Substring(3), out pos) && pos >= 1 && pos <= legacy.Length) id = legacy[pos - 1];
                        }
                        int idx = id == null ? -1 : KeyIndex(id);
                        string[] parts = v.Split(',');
                        uint m, vk;
                        if (idx >= 0 && parts.Length == 2 && uint.TryParse(parts[0], out m) && uint.TryParse(parts[1], out vk))
                        { keyMods[idx] = m & 7; keyVk[idx] = vk; }
                    }
                    else if (k == "color" && v.Length == 6 && int.TryParse(v, System.Globalization.NumberStyles.HexNumber, null, out n))
                        HighlightColor = Color.FromArgb(255, (n >> 16) & 255, (n >> 8) & 255, n & 255);
                }
            }
            catch { } // a missing or damaged file just means defaults
        }

        private void SaveSettings()
        {
            try
            {
                string path = SettingsPath();
                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(path));
                Color c = HighlightColor;
                string[] lines = new string[] {
                    "# Word Focus settings (saved automatically). Delete this file to go back to defaults.",
                    "on=" + (enabled ? "1" : "0"),
                    "words=" + (count == SentenceValue ? "sentence" : count == LineValue ? "line" : count == LinePlusValue ? "lineplus" : count.ToString()),
                    "overlap=" + (overlapLines ? "1" : "0"),
                    "pointer=" + cursorPercent,
                    "strength=" + strengthPercent,
                    "ombre=" + (ombre ? "1" : "0"),
                    "color=" + c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2"),
                    "style=" + (style == StyleUnderline ? "underline" : style == StyleBoth ? "both" : "highlight"),
                    "fstyle=" + focusStyle,
                    "fstrength=" + focusStrength
                };
                List<string> all = new List<string>(lines);
                for (int i = 0; i < KeyIds.Length; i++) all.Add("key." + KeyIds[i] + "=" + keyMods[i] + "," + keyVk[i]);
                for (int i = 0; i < KeyIds.Length; i++) if (NoteDismissed(i)) all.Add("noteok." + KeyIds[i] + "=" + KeySig(i));
                lines = all.ToArray();
                System.IO.File.WriteAllLines(path, lines);
                settingsChanged = false;
            }
            catch { }
        }

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
            if (settingsChanged) SaveSettings();
            RestoreCursor();
            UnregisterAll();
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

        private void FocusLoop()
        {
            while (running)
            {
                Thread.Sleep(70);
                try { FocusStep(); } catch (Exception ex) { Log.Write("Focus ERROR " + ex.GetType().Name + ": " + ex.Message); ClearFocusWorker(); }
            }
        }

        // ---- Focus layer (focus worker thread) ----
        private void FocusStep()
        {
            int req = focusRequest;
            if (req != FocusNone)
            {
                focusRequest = FocusNone;
                if (req == FocusClear) ClearFocusWorker();
                else StartFocusWorker(req, focusRequestPt);
            }
            if (focusRanges == null) return;
            // Check ~7x a second (the work per check is small and capped). If the app is slow to answer, wait longer.
            int interval = 140 + focusSlowPenalty;
            if (unchecked(Environment.TickCount - focusRefreshTick) < interval && !focusDirty && !pageRebuildPending) return;
            focusRefreshTick = Environment.TickCount;

            // Hide while another window is in front (e.g. you switched apps); show again when you come back.
            IntPtr fg = Native.GetForegroundWindow();
            uint fgPid;
            Native.GetWindowThreadProcessId(fg, out fgPid);
            bool ours = fgPid == MyPid; // our settings windows count as "still here"
            bool visible = focusRoot == IntPtr.Zero || fg == focusRoot || ours;
            int t0 = Environment.TickCount;
            List<System.Windows.Rect> rawList = new List<System.Windows.Rect>();
            if (visible) foreach (TextPatternRange fr in focusRanges) rawList.AddRange(fr.GetBoundingRectangles());
            System.Windows.Rect[] raw = rawList.ToArray();
            int took = unchecked(Environment.TickCount - t0);
            // Be gentle with slow apps: back off, and give up on a Focus area that is too heavy to track.
            focusSlowPenalty = took > 150 ? Math.Min(5000, took * 8) : 0;
            if (took > 2000 || raw.Length > 1500)
            {
                Log.Write("Focus: too heavy for this app (" + took + " ms, " + raw.Length + " shapes) - turned off");
                ClearFocusWorker();
                return;
            }
            System.Windows.Rect[] rects = TextLinesOnly(raw);

            // Page Focus follows scrolling: when the text moved, re-find the visible body lines
            // (at most ~3x a second while you scroll, and once more after you stop).
            if (focusKind == FocusPage && visible)
            {
                string moved = rects.Length > 0 ? (int)rects[0].Y + "," + (int)rects[rects.Length - 1].Y + "," + rects.Length : "none";
                if (moved != lastPagePos) { lastPagePos = moved; pageRebuildPending = true; }
                if (pageRebuildPending && unchecked(Environment.TickCount - lastPageBuild) >= 300)
                {
                    pageRebuildPending = false;
                    RebuildPage(rects);
                    raw = new System.Windows.Rect[0];
                    List<System.Windows.Rect> again = new List<System.Windows.Rect>();
                    foreach (TextPatternRange fr in focusRanges) again.AddRange(fr.GetBoundingRectangles());
                    rects = TextLinesOnly(again.ToArray());
                    lastPagePos = rects.Length > 0 ? (int)rects[0].Y + "," + (int)rects[rects.Length - 1].Y + "," + rects.Length : "none";
                }
            }

            string sig = visible + "|" + rects.Length;
            foreach (System.Windows.Rect r in rects) sig += "|" + (int)r.X + "," + (int)r.Y + "," + (int)r.Width + "," + (int)r.Height;
            if (sig == focusSig && !focusDirty) return; // nothing moved
            focusSig = sig;
            focusDirty = false;
            System.Windows.Rect[] result = rects;
            try { focusOverlay.BeginInvoke(new Action(delegate { RenderFocus(result); })); } catch { }
        }

        // Some apps include the boxes that hold the text (the page, sections, images) along with the
        // lines themselves. Keep only shapes about the height of a line of text.
        private static System.Windows.Rect[] TextLinesOnly(System.Windows.Rect[] rects)
        {
            if (rects == null || rects.Length == 0) return new System.Windows.Rect[0];
            List<double> heights = new List<double>();
            foreach (System.Windows.Rect r in rects) if (r.Height >= 4 && r.Width >= 1) heights.Add(r.Height);
            if (heights.Count == 0) return new System.Windows.Rect[0];
            heights.Sort();
            // The most common line height: the median of the shorter half, so a few huge boxes can't skew it.
            double typical = heights[(heights.Count - 1) / 4];
            double maxH = Math.Max(typical * 2.2, 8);
            List<System.Windows.Rect> keep = new List<System.Windows.Rect>();
            foreach (System.Windows.Rect r in rects)
                if (r.Height >= 4 && r.Width >= 1 && r.Height <= maxH) keep.Add(r);
            return keep.ToArray();
        }

        private void ClearFocusWorker()
        {
            focusRanges = null;
            focusRoot = IntPtr.Zero;
            focusSig = "";
            focusKind = FocusNone;
            try { focusOverlay.BeginInvoke(new Action(delegate { RenderFocus(null); })); } catch { }
        }

        private class LineInfo { public TextPatternRange Range; public System.Windows.Rect Box; }

        private static System.Windows.Rect Union(System.Windows.Rect[] rs)
        {
            System.Windows.Rect u = rs[0];
            for (int i = 1; i < rs.Length; i++) u.Union(rs[i]);
            return u;
        }

        // Steps line by line (dir = -1 up, +1 down) from 'from' while lines are inside the window,
        // collecting each visible line. At most 120 steps, so it never gets expensive.
        private static List<LineInfo> WalkLines(TextPatternRange from, int dir, double winTop, double winBottom)
        {
            List<LineInfo> found = new List<LineInfo>();
            TextPatternRange probe = from.Clone();
            int blanks = 0;
            for (int i = 0; i < 120; i++)
            {
                if (probe.Move(TextUnit.Line, dir) == 0) break;   // start or end of the document
                probe.ExpandToEnclosingUnit(TextUnit.Line);
                System.Windows.Rect[] rs = probe.GetBoundingRectangles();
                if (rs.Length == 0) { if (++blanks > 3) break; continue; } // blank or hidden line; allow a few
                blanks = 0;
                System.Windows.Rect box = Union(rs);
                if (dir < 0 && box.Bottom < winTop) break;     // scrolled above the window
                if (dir > 0 && box.Top > winBottom) break;     // below the window
                LineInfo li = new LineInfo(); li.Range = probe.Clone(); li.Box = box;
                found.Add(li);
            }
            return found;
        }

        // Controls whose text isn't reading material (links inside body text are NOT in this list).
        private static readonly ControlType[] NotBodyTypes = new ControlType[] {
            ControlType.Button, ControlType.SplitButton, ControlType.Edit, ControlType.ComboBox,
            ControlType.Menu, ControlType.MenuBar, ControlType.MenuItem, ControlType.ToolBar,
            ControlType.Tab, ControlType.TabItem, ControlType.CheckBox, ControlType.RadioButton,
            ControlType.Spinner, ControlType.Slider, ControlType.ScrollBar, ControlType.Header,
            ControlType.HeaderItem, ControlType.TitleBar, ControlType.StatusBar, ControlType.ToolTip
        };

        // Skip a line if the item it sits in (or one just above it) is a heading or a control like a button,
        // text box, drop-down, menu or tab. Works with the built-in Windows accessibility component.
        private static bool IsHeadingLine(TextPatternRange r)
        {
            try
            {
                AutomationElement e = r.GetEnclosingElement();
                TreeWalker walker = TreeWalker.ControlViewWalker;
                for (int hops = 0; e != null && hops < 3; hops++)
                {
                    if (IsDocument(e)) break;
                    string t = e.Current.LocalizedControlType;
                    if (t != null && t.ToLowerInvariant().Contains("heading")) return true;
                    ControlType ct = e.Current.ControlType;
                    foreach (ControlType nb in NotBodyTypes) if (ct == nb) return true;
                    e = walker.GetParent(e);
                }
            }
            catch { }
            return false;
        }

        // Body text = lines in the same column as the line you pointed at, not much taller than it, and not headings.
        private static bool IsBodyLine(LineInfo li, System.Windows.Rect startBox)
        {
            if (li.Box.Height > startBox.Height * 1.3) return false;                       // big title text
            if (li.Box.Right <= startBox.Left || li.Box.Left >= startBox.Right) return false; // other column (menus, sidebars)
            if (IsHeadingLine(li.Range)) return false;
            return true;
        }

        // Page Focus state (focus worker only): used to re-find the visible body lines after scrolling.
        private TextPattern pageTp;
        private System.Windows.Rect pageColumn;
        private int lastPageBuild;
        private bool pageRebuildPending;

        // Walks up and down from 'start' to the window edges and joins the body lines into runs.
        private static List<TextPatternRange> BuildPageRuns(TextPatternRange start, System.Windows.Rect column,
                                                            double winTop, double winBottom, bool startIsPointed)
        {
            int tw = Environment.TickCount;
            List<TextPatternRange> runs = new List<TextPatternRange>();
            System.Windows.Rect[] startRects = start.GetBoundingRectangles();
            if (startRects.Length == 0) return runs;
            LineInfo startLine = new LineInfo(); startLine.Range = start; startLine.Box = Union(startRects);
            List<LineInfo> up = WalkLines(start, -1, winTop, winBottom);
            List<LineInfo> down = WalkLines(start, 1, winTop, winBottom);
            up.Reverse();
            List<LineInfo> all = new List<LineInfo>(up);
            all.Add(startLine);
            all.AddRange(down);

            // Join neighbouring body lines into runs (fewer, bigger ranges = fewer questions to the app later).
            TextPatternRange run = null;
            int kept = 0;
            foreach (LineInfo li in all)
            {
                bool body = (li == startLine && startIsPointed) ? !IsHeadingLine(li.Range) : IsBodyLine(li, column);
                if (body)
                {
                    kept++;
                    if (run == null) run = li.Range.Clone();
                    else run.MoveEndpointByRange(TextPatternRangeEndpoint.End, li.Range, TextPatternRangeEndpoint.End);
                }
                else if (run != null) { runs.Add(run); run = null; }
            }
            if (run != null) runs.Add(run);
            Log.Write("Page focus: " + all.Count + " visible lines, " + kept + " body lines in " + runs.Count + " runs; window " +
                      winTop + ".." + winBottom + "; took " + unchecked(Environment.TickCount - tw) + " ms");
            return runs;
        }

        // After a scroll: start again from the body line nearest the middle of the window and re-walk,
        // so lines that scrolled into view light up and ones that left drop off.
        private void RebuildPage(System.Windows.Rect[] currentRects)
        {
            Native.RECT wr;
            if (pageTp == null || focusRoot == IntPtr.Zero || !Native.GetWindowRect(focusRoot, out wr)) return;
            double mid = (wr.Top + wr.Bottom) / 2.0;
            System.Windows.Point anchor;
            System.Windows.Rect best = System.Windows.Rect.Empty;
            double bestDist = double.MaxValue;
            foreach (System.Windows.Rect r in currentRects)
            {
                if (r.Bottom < wr.Top || r.Top > wr.Bottom) continue;
                double d = Math.Abs((r.Top + r.Bottom) / 2 - mid);
                if (d < bestDist) { bestDist = d; best = r; }
            }
            if (!best.IsEmpty) anchor = new System.Windows.Point(best.Left + Math.Min(best.Width / 2, 20), (best.Top + best.Bottom) / 2);
            else anchor = new System.Windows.Point(pageColumn.Left + Math.Min(pageColumn.Width / 2, 20), mid); // everything scrolled away

            TextPatternRange start;
            try { start = pageTp.RangeFromPoint(anchor); start.ExpandToEnclosingUnit(TextUnit.Line); }
            catch { return; }
            List<TextPatternRange> runs = BuildPageRuns(start, pageColumn, wr.Top, wr.Bottom, false);
            if (runs.Count > 0) focusRanges = runs;
            lastPageBuild = Environment.TickCount;
        }

        private static AutomationElement FindDocument(AutomationElement e)
        {
            TreeWalker walker = TreeWalker.ControlViewWalker;
            for (int hops = 0; e != null && hops < 40; hops++)
            {
                if (IsDocument(e) && GetTP(e) != null) return e;
                try { e = walker.GetParent(e); } catch { return null; }
            }
            return null;
        }

        private void StartFocusWorker(int kind, Native.POINT p)
        {
            System.Windows.Point pt = new System.Windows.Point(p.X, p.Y);
            AutomationElement el = AutomationElement.FromPoint(pt);
            AutomationElement src; string how;
            TextPattern tp = FindTextPattern(el, pt, out src, out how);
            if (tp == null) { Log.Write("Focus: no text under mouse"); ClearFocusWorker(); return; }
            AutomationElement doc = FindDocument(src ?? el);
            TextPattern docTp = doc != null ? GetTP(doc) : null;

            TextPatternRange range = null;
            if (kind == FocusPage)
            {
                // "Page" = the lines visible in that window. Built line by line outward from the mouse,
                // with a hard cap, so a huge document (e.g. a 1000-page PDF) costs no more than a short one.
                TextPattern pageTp0 = docTp ?? tp;
                TextPatternRange start = null;
                if (docTp != null && src != null && !IsDocument(src)) { try { start = docTp.RangeFromChild(src); } catch { start = null; } }
                if (start == null) start = pageTp0.RangeFromPoint(pt);
                start.ExpandToEnclosingUnit(TextUnit.Line);

                IntPtr root = Native.GetAncestor(Native.WindowFromPoint(p), 2);
                Native.RECT wr;
                double winTop = double.MinValue, winBottom = double.MaxValue;
                if (root != IntPtr.Zero && Native.GetWindowRect(root, out wr)) { winTop = wr.Top; winBottom = wr.Bottom; }

                System.Windows.Rect[] startRects = start.GetBoundingRectangles();
                if (startRects.Length == 0) { Log.Write("Focus: line under mouse has no shape"); ClearFocusWorker(); return; }
                System.Windows.Rect column = Union(startRects);   // the column + line height to treat as body text
                List<TextPatternRange> runs = BuildPageRuns(start, column, winTop, winBottom, true);
                if (runs.Count == 0) { Log.Write("Focus: no body text found"); ClearFocusWorker(); return; }
                pageTp = pageTp0;
                pageColumn = column;
                lastPageBuild = Environment.TickCount;
                pageRebuildPending = false;
                StartFocusRanges(kind, p, runs, how);
                return;
            }
            else
            {
                // Paragraph: ask the whole page for the paragraph containing this text, so links/bold
                // in the middle of a paragraph don't cut it short. Fall back to the local text.
                if (docTp != null && src != null && !IsDocument(src))
                {
                    try { range = docTp.RangeFromChild(src); } catch { range = null; }
                }
                if (range == null) range = tp.RangeFromPoint(pt);
                range.ExpandToEnclosingUnit(TextUnit.Paragraph);
            }

            System.Windows.Rect[] firstRects = range.GetBoundingRectangles();
            Log.Write("Focus range: " + firstRects.Length + " shapes, " + TextLinesOnly(firstRects).Length + " kept as text lines; via " + how + " " + Describe(src) + " doc=" + (docTp != null));
            if (TextLinesOnly(firstRects).Length == 0) { Log.Write("Focus: range has no visible text lines"); ClearFocusWorker(); return; }
            List<TextPatternRange> one = new List<TextPatternRange>();
            one.Add(range);
            StartFocusRanges(kind, p, one, how);
        }

        private void StartFocusRanges(int kind, Native.POINT p, List<TextPatternRange> ranges, string how)
        {
            focusRanges = ranges;
            focusRoot = Native.GetAncestor(Native.WindowFromPoint(p), 2); // GA_ROOT: the app's main window
            focusKind = kind;
            focusSig = "";
            focusDirty = true;
            focusRefreshTick = 0;
            focusSlowPenalty = 0;
            Log.Write("Focus " + (kind == FocusPage ? "page" : "paragraph") + " via " + how);
        }

        // ---- Focus layer drawing (UI thread) ----
        private void RenderFocus(System.Windows.Rect[] rects)
        {
            if (rects == null || rects.Length == 0) { ClearOverlay(focusOverlay); return; }
            double l = double.MaxValue, t = double.MaxValue, r = double.MinValue, b = double.MinValue;
            foreach (System.Windows.Rect rc in rects)
            {
                l = Math.Min(l, rc.Left); t = Math.Min(t, rc.Top);
                r = Math.Max(r, rc.Right); b = Math.Max(b, rc.Bottom);
            }
            int x0 = (int)Math.Floor(l) - 2, y0 = (int)Math.Floor(t) - 2;
            int w0 = (int)Math.Ceiling(r) - x0 + 2, h0 = (int)Math.Ceiling(b) - y0 + 2;
            if (w0 < 1 || h0 < 1 || w0 > 10000 || h0 > 10000) { ClearOverlay(focusOverlay); return; }

            int st = focusStyle;
            int alpha = focusStrength * 255 / 100;
            using (Bitmap bmp = new Bitmap(w0, h0, PixelFormat.Format32bppArgb))
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.Transparent);
                Color lineColor = Color.Black;
                if (st == StyleBoth) lineColor = IsDarkBehind(x0, y0, w0, h0) ? Color.White : Color.Black;
                using (SolidBrush hb = new SolidBrush(Color.FromArgb(alpha, HighlightColor)))
                using (SolidBrush lb = new SolidBrush(lineColor))
                {
                    foreach (System.Windows.Rect rc in rects)
                    {
                        RectangleF rf = new RectangleF((float)(rc.Left - x0), (float)(rc.Top - y0), (float)rc.Width, (float)rc.Height);
                        if (rf.Width < 1 || rf.Height < 1) continue;
                        float thick = Math.Max(2f, rf.Height * 0.12f);
                        RectangleF line = new RectangleF(rf.X, rf.Bottom - thick, rf.Width, thick);
                        if (st == StyleHighlight || st == StyleBoth) g.FillRectangle(hb, rf);
                        if (st == StyleUnderline) g.FillRectangle(hb, line);
                        if (st == StyleBoth) g.FillRectangle(lb, line);
                    }
                }
                focusOverlay.SetBitmap(bmp, x0, y0);
                overlay.BringToTop(); // hover highlight stays above the Focus layer
            }
        }

        private static void ClearOverlay(Overlay ov)
        {
            using (Bitmap b = new Bitmap(1, 1, PixelFormat.Format32bppArgb))
            {
                b.SetPixel(0, 0, Color.Transparent);
                ov.SetBitmap(b, 0, 0);
            }
        }

        // True if a word ends a sentence: . ! or ? possibly followed by closing quotes/brackets.
        private static bool EndsSentence(string s)
        {
            if (s == null) return false;
            string t = s.TrimEnd();
            while (t.Length > 0 && "\"')]}\u201D\u2019".IndexOf(t[t.Length - 1]) >= 0) t = t.Substring(0, t.Length - 1);
            if (t.Length == 0) return false;
            char c = t[t.Length - 1];
            return c == '.' || c == '!' || c == '?' || c == '\u2026';
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

            int mode = count;
            bool lineMode = mode == LineValue || mode == LinePlusValue;
            bool sentenceMode = mode == SentenceValue;
            int want = lineMode ? 300 : sentenceMode ? 60 : count;
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
                        if (Math.Abs(mid - lineMid) > lineH * 0.5) // reached the next line
                        {
                            if (mode == LinePlusValue) list.Add(new WordBox(rects)); // ...include its first word
                            break;
                        }
                    }
                    list.Add(new WordBox(rects));
                    if (sentenceMode && EndsSentence(text)) break;
                }
                if (cur.Move(TextUnit.Word, 1) == 0) break; // end of text
                cur.ExpandToEnclosingUnit(TextUnit.Word);
            }
            return list;
        }

        private int AlphaAt(int i, int n)
        {
            int MaxAlpha = strengthPercent * 255 / 100;
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

        // Looks at the screen behind the words (our own see-through overlay isn't captured) and
        // returns true if the background is dark. Uses the median brightness, so the text itself barely counts.
        private static bool IsDarkBehind(int x, int y, int w, int h)
        {
            try
            {
                using (Bitmap shot = new Bitmap(w, h, PixelFormat.Format32bppRgb))
                {
                    using (Graphics sg = Graphics.FromImage(shot))
                        sg.CopyFromScreen(x, y, 0, 0, new Size(w, h), CopyPixelOperation.SourceCopy);
                    List<int> lum = new List<int>();
                    int stepX = Math.Max(1, w / 40), stepY = Math.Max(1, h / 10);
                    for (int yy = 0; yy < h; yy += stepY)
                        for (int xx = 0; xx < w; xx += stepX)
                        {
                            Color c = shot.GetPixel(xx, yy);
                            lum.Add((c.R * 299 + c.G * 587 + c.B * 114) / 1000);
                        }
                    if (lum.Count == 0) return false;
                    lum.Sort();
                    return lum[lum.Count / 2] < 128;
                }
            }
            catch { return false; }
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
                int st = style;
                Region painted = overlapLines ? null : new Region(RectangleF.Empty);
                List<RectangleF> lines = st == StyleBoth ? new List<RectangleF>() : null;
                for (int i = 0; i < n; i++)
                {
                    int a0 = AlphaAt(i, n), a1 = AlphaAt(i + 1, n);
                    foreach (System.Windows.Rect rc in boxes[i].Rects)
                    {
                        RectangleF rf = new RectangleF((float)(rc.Left - x0) - 1, (float)(rc.Top - y0),
                                                       (float)rc.Width + 2, (float)rc.Height);
                        if (rf.Width < 1 || rf.Height < 1) continue;
                        float thick = Math.Max(2f, rf.Height * 0.12f);
                        if (lines != null) lines.Add(new RectangleF(rf.X + 1, rf.Bottom - thick, rf.Width - 2, thick));
                        if (st == StyleUnderline)
                            rf = new RectangleF(rf.X, rf.Bottom - thick, rf.Width, thick);
                        RectangleF brushRect = new RectangleF(rf.X - 1, rf.Y, rf.Width + 2, rf.Height);
                        using (LinearGradientBrush br = new LinearGradientBrush(brushRect,
                                   Color.FromArgb(a0, HighlightColor), Color.FromArgb(a1, HighlightColor),
                                   LinearGradientMode.Horizontal))
                        {
                            if (painted == null) g.FillRectangle(br, rf);
                            else
                            {
                                // Paint only where nothing has been painted yet, so neighbours don't double up.
                                g.SetClip(painted, CombineMode.Exclude);
                                g.FillRectangle(br, rf);
                                g.ResetClip();
                                painted.Union(rf);
                            }
                        }
                    }
                }
                if (painted != null) painted.Dispose();

                // "Both": solid underline, white on dark pages and black on light pages.
                if (lines != null && lines.Count > 0)
                {
                    Color lineColor = IsDarkBehind(x0, y0, w0, h0) ? Color.White : Color.Black;
                    using (SolidBrush lb = new SolidBrush(lineColor))
                        foreach (RectangleF lr in lines) g.FillRectangle(lb, lr);
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
        public static bool Hover = false;   // routine hover lines off
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
                if (!Hover && msg.StartsWith("Q ") && msg.IndexOf("ERROR") < 0) return; // skip routine hover lines this round
                lines++;
                try { w.WriteLine(DateTime.Now.ToString("HH:mm:ss.fff") + "  " + msg); } catch { }
            }
        }
    }

    // ---- "Start with Windows": a shortcut in the user's Startup folder (no admin, no registry) ----
    public static class Startup
    {
        public static string Folder; // where WordFocus.ps1 and Start Word Focus.cmd live

        private static string LinkPath()
        {
            return System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Startup), "Word Focus.lnk");
        }

        public static bool IsEnabled() { return System.IO.File.Exists(LinkPath()); }

        // Returns null on success, or an error message.
        public static string Set(bool on)
        {
            try
            {
                string link = LinkPath();
                if (!on) { if (System.IO.File.Exists(link)) System.IO.File.Delete(link); return null; }
                string target = System.IO.Path.Combine(Folder, "Start Word Focus.cmd");
                if (!System.IO.File.Exists(target)) return "Can't find " + target;
                Type t = Type.GetTypeFromProgID("WScript.Shell");
                object shell = Activator.CreateInstance(t);
                object sc = t.InvokeMember("CreateShortcut", System.Reflection.BindingFlags.InvokeMethod, null, shell, new object[] { link });
                Type st = sc.GetType();
                st.InvokeMember("TargetPath", System.Reflection.BindingFlags.SetProperty, null, sc, new object[] { target });
                st.InvokeMember("WorkingDirectory", System.Reflection.BindingFlags.SetProperty, null, sc, new object[] { Folder });
                st.InvokeMember("WindowStyle", System.Reflection.BindingFlags.SetProperty, null, sc, new object[] { 7 }); // minimized
                st.InvokeMember("Description", System.Reflection.BindingFlags.SetProperty, null, sc, new object[] { "Word Focus reading highlighter" });
                st.InvokeMember("Save", System.Reflection.BindingFlags.InvokeMethod, null, sc, null);
                Marshal.FinalReleaseComObject(sc);
                Marshal.FinalReleaseComObject(shell);
                return null;
            }
            catch (Exception ex) { return ex.Message; }
        }
    }

    public static class App
    {
        public static void Run(string folder)
        {
            // Only one copy at a time (e.g. if it already started with Windows and you double-click it again).
            bool firstCopy;
            Mutex single = new Mutex(true, "WordFocus_SingleInstance_2sage4thyme", out firstCopy);
            if (!firstCopy)
            {
                MessageBox.Show("Word Focus is already running. Look for its icon in the system tray, by the clock.\nCtrl+Alt+H turns the highlight on or off.", "Word Focus");
                return;
            }
            Startup.Folder = folder;
            Log.Start(folder);
            Native.SetProcessDPIAware(); // so word positions match the screen exactly
            Application.EnableVisualStyles();
            new Controller();
            Application.Run();
            GC.KeepAlive(single); // holds the one-copy lock until exit
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
