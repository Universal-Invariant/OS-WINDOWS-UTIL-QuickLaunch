using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Windows.Forms;
using System.Xml.Serialization;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;

// https://archive.softwareheritage.org/save/


public class QuickLauncher : Form, IMessageFilter
{

    static string windowName = "Quick Launcher"; // Should match the titleLabel.Text or Form.Text    
    private NotifyIcon trayIcon;
    private ContextMenuStrip trayMenu;
    private ToolStripMenuItem idleToggleItem;

    private Panel mainPanel;
    private FlowLayoutPanel itemsPanel;
    private QuickLauncherSettings appSettings;
    private readonly string configPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), windowName.Replace(" ", ""), windowName.Replace(" ", "") + "Settings.xml");
    private QuickItemCollection currentItems { get { return (navigationStack.Count() == 0) ? rootItems : navigationStack.Peek(); } }
    private QuickItemCollection rootItems { get { return appSettings.Items; } }
    private Keys registeredHotkey { get { return appSettings.GlobalHotkey; } }
    private Stack<QuickItemCollection> navigationStack = new Stack<QuickItemCollection>();

    // Title bar elements
    private Panel titleBar;
    private FlowLayoutPanel titleBarRHS;
    private Label titleLabel;
    private Button closeButton;
    private Button editButton;
    private bool isDragging = false;
    private Point lastCursor;
    private Point lastForm;

    // Add a unique message ID for communication between instances
    private const int WM_SHOW_OR_HIDE_LAUNCHER = 0x0401; // Choose a unique value

    // Single-instance IPC: the second instance posts this message to the first
    // instance's window. The wParam encodes the requested action so the same
    // hidden window can serve multiple entry points (hotkey, tray, CLI).
    private const int SHOWMSG_TOGGLE = 0;   // show if hidden / close-or-navigate if visible
    private const int SHOWMSG_SHOW = 1;     // always show
    private const int SHOWMSG_EDIT = 2;     // open the editor

    /// <summary>
    /// Command-line interface (handled by the primary instance):
    ///   QuickLaunch.exe               -> toggle launcher (same as hotkey)
    ///   QuickLaunch.exe /show         -> show launcher
    ///   QuickLaunch.exe /hide         -> hide/close launcher
    ///   QuickLaunch.exe /edit         -> open the editor
    /// Unknown args are treated as toggle (legacy behaviour).
    /// </summary>
    private static int ShowMessageWParamForArgs(string[] args)
    {
        foreach (var raw in args ?? Array.Empty<string>())
        {
            var a = raw.TrimStart('/', '-').ToLowerInvariant();
            switch (a)
            {
                case "show": return SHOWMSG_SHOW;
                case "hide": return SHOWMSG_TOGGLE; // toggles off if visible, shows if hidden
                case "edit": return SHOWMSG_EDIT;
            }
        }
        return SHOWMSG_TOGGLE;
    }

    // P/Invoke declarations for finding and manipulating the window
    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    private static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    private const int SW_RESTORE = 9; // Restore window if minimized
    private const int SW_SHOW = 5;    // Show window
    private const int SW_HIDE = 0;    // Hide window

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern IntPtr SetActiveWindow(IntPtr hWnd);

    private const int WM_HOTKEY = 0x0312;
    private const int HOTKEY_ID = 9000; // Choose a unique ID for your hotkey

    // ---------------------------------------------------------------------
    // Idle process-exiter: quits the app after a configurable period without
    // any keyboard/mouse interaction (see ActivityMonitor below). The timer
    // is only ticking while the launcher window is hidden; showing the
    // launcher always counts as activity and resets it.
    // ---------------------------------------------------------------------
    private System.Windows.Forms.Timer idleExitTimer;
    private bool exitingDueToIdle = false;
    private bool activityFlagPending = false;   // coalescing flag for the message-filter hook

    public static QuickLauncher Instance { get; private set; }

    /// <summary>
    /// Win32 LASTINPUTINFO: seconds since the system-wide last input event.
    /// Used as a safety net so an external process injecting keystrokes into a
    /// child of this process (e.g. AutoHotkey re-sending SendKeys shortcuts)
    /// can never trigger the idle auto-exit while the user is at the keyboard.
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    private struct LASTINPUTINFO
    {
        public uint cbSize;
        public uint dwTime;
    }

    [DllImport("user32.dll")]
    private static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);

    private static double SystemIdleAgeMs()
    {
        try
        {
            var lii = new LASTINPUTINFO { cbSize = (uint)Marshal.SizeOf(typeof(LASTINPUTINFO)) };
            if (!GetLastInputInfo(ref lii)) return -1;
            return (Environment.TickCount64 - lii.dwTime);
        }
        catch { return -1; }
    }

    private void InitializeIdleExitTimer()
    {
        idleExitTimer = new System.Windows.Forms.Timer();
        // Poll at a fixed cadence; the actual timeout is compared against the
        // configurable interval inside the Tick handler (see ApplyIdleSettings).
        idleExitTimer.Interval = IdlePollMs;
        idleExitTimer.Tick += (s, e) =>
        {
            if (!appSettings.AutoExitWhenIdle) return;
            // Never exit while the window is visible or an item is running.
            if (this.Visible || stopAutoClose) return;

            double age = ActivityMonitor.LastActivityAgeMs;
            // Safety net: LASTINPUTINFO reports input that happened *after* our
            // recorded activity (e.g. keystrokes injected into one of our child
            // processes by AutoHotkey, which low-level hooks do not observe).
            // Trust whichever timestamp is more recent.
            double sysAge = SystemIdleAgeMs();
            if (sysAge >= 0 && sysAge < age) age = sysAge;

            int timeout = (idleExitTimer.Tag is int tmo) ? tmo : IdleTimeoutMs;
            if (age < timeout) return;

            exitingDueToIdle = true;
            idleExitTimer.Stop();
            Application.Exit();
        };
        idleExitTimer.Start();
    }

    // How often the idle watchdog evaluates the timestamps (independent of the
    // user-configurable timeout so polling stays cheap and predictable).
    private const int IdlePollMs = 2000;

    // Recompute the timeout whenever settings change (e.g. editor save / tray toggle).
    private void ApplyIdleSettings()
    {
        if (idleExitTimer == null) return;
        idleExitTimer.Tag = IdleTimeoutMs;   // effective timeout, checked on each tick
        idleExitTimer.Enabled = appSettings.AutoExitWhenIdle;
        UpdateIdleToggleText();
    }

    // Keep the tray menu item label in sync with the configured timeout.
    private void UpdateIdleToggleText()
    {
        if (idleToggleItem == null) return;
        idleToggleItem.Text = appSettings.AutoExitWhenIdle
            ? $"Auto-exit when idle ({appSettings.IdleTimeoutSeconds}s)"
            : "Auto-exit when idle (off)";
    }

    private int IdleTimeoutMs
    {
        get
        {
            var secs = appSettings.IdleTimeoutSeconds;
            if (secs <= 0) secs = QuickLauncherSettings.DefaultIdleTimeoutSeconds;
            return secs * 1000;
        }
    }

    /// <summary>
    /// Call this from every keyboard/mouse event that indicates the user is
    /// interacting with the launcher. It resets the idle-exit countdown and,
    /// as a safety net, restarts the monitor if a previous Application.Exit
    /// was cancelled by another open form (e.g. the editor dialog).
    /// </summary>
    public static void RecordActivity()
    {
        ActivityMonitor.RecordActivity();
        var inst = Instance;
        if (inst != null && inst.exitingDueToIdle)
        {
            inst.exitingDueToIdle = false;
            inst.ApplyIdleSettings();
        }
    }

    // ------------------------------------------------------------------
    // IMessageFilter: coalesced activity tracking for ALL messages pumped
    // by this app's message loop (launcher window, editor dialog, tray).
    // Input messages are frequent, so instead of writing the timestamp on
    // every single one we set a flag and flush it at most once per 250 ms.
    // ------------------------------------------------------------------
    private const int ActivityFlushIntervalMs = 250;
    private long lastActivityFlushTicks = 0;

    public bool PreFilterMessage(ref Message m)
    {
        // WM_KEYDOWN..WM_MOUSELAST covers all keyboard & mouse input messages.
        if (m.Msg >= 0x0100 && m.Msg <= 0x02FF)
        {
            activityFlagPending = true;
        }
        else if (activityFlagPending)
        {
            activityFlagPending = false;
            var now = Environment.TickCount64;
            if (now - lastActivityFlushTicks >= ActivityFlushIntervalMs)
            {
                lastActivityFlushTicks = now;
                RecordActivity();
            }
        }
        return false; // never consume messages
    }


    [DllImport("user32.dll")]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, Keys vk);

    [DllImport("user32.dll")]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);


    [DllImport("user32")]
    private static extern bool SetWindowPos(IntPtr hwnd, IntPtr hwnd2, int x, int y, int cx, int cy, int flags);
    // https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowpos
    //instead of calling SetForegroundWindow
    //SWP_NOSIZE = 1, SWP_NOMOVE = 2  -> keep the current pos and size (ignore x,y,cx,cy).
    //the second param = -1   -> set window as Topmost.


    protected override void OnLoad(EventArgs e)
    {
        base.OnLoad(e);
        this.Hide(); // Hide initially

        // Register the global hotkey when the form loads. The hotkey is what
        // summons the launcher, so it must work even without the tray icon;
        // (previously it was gated on UseTrayIcon, which left no way to show
        // the window when running tray-less).
        RegisterGlobalHotkey();
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        // If the user clicks the 'X', just hide the form and keep the tray icon
        if (appSettings.UseTrayIcon && e.CloseReason == CloseReason.UserClosing)
        {
            e.Cancel = true; // Cancel the closing event
            CloseOrHide();
        }
        else
        {
            // If closing for other reasons (e.g., Application.Exit), let it proceed
            // Unregister hotkey (already handled in Dispose)
            base.OnFormClosing(e);
        }
    }

    private void InitializeTrayIcon()
    {
        Icon appIcon = SystemIcons.Application;
        var iconResourceName = "QuickLaunch.Plant - Leafs.ico";
        using (var stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(iconResourceName))
        {
            if (stream != null)
            {
                try
                {
                    appIcon = new Icon(stream);
                }
                catch (ArgumentException ex) // Catch potential errors from Icon constructor (e.g., invalid icon format)
                {
                }
            }
            else
            {

            }
        }

        trayMenu = new ContextMenuStrip();
        var showItem = new ToolStripMenuItem("Show");
        showItem.Click += (s, e) => ShowLauncher();
        var editItem = new ToolStripMenuItem("Edit");
        editItem.Click += (s, e) => OpenEditor();
        var reloadItem = new ToolStripMenuItem("Reload Settings");
        reloadItem.Click += (s, e) => ReloadSettingsFromDisk();
        var openFolderItem = new ToolStripMenuItem("Open Settings Folder");
        openFolderItem.Click += (s, e) => { try { Process.Start("explorer.exe", Path.GetDirectoryName(configPath)); } catch { } };

        // Tray icon on/off toggle (hotkey registration is only active with the tray).
        var trayToggleItem = new ToolStripMenuItem("Show Tray Icon")
        {
            Checked = appSettings.UseTrayIcon,
            CheckOnClick = true
        };
        trayToggleItem.CheckedChanged += (s, e) =>
        {
            appSettings.UseTrayIcon = ((ToolStripMenuItem)s).Checked;
            SaveConfiguration();
            UpdateTrayVisibility();
            RefreshGlobalHotkey();
        };

        // Idle-exit on/off toggle (kept in sync via SettingChanged).
        idleToggleItem = new ToolStripMenuItem()
        {
            Checked = appSettings.AutoExitWhenIdle,
            CheckOnClick = true
        };
        UpdateIdleToggleText();
        idleToggleItem.CheckedChanged += (s, e) =>
        {
            appSettings.AutoExitWhenIdle = ((ToolStripMenuItem)s).Checked;
            SaveConfiguration();
            ApplyIdleSettings();
            RecordActivity(); // toggling counts as usage -> fresh countdown
        };

        var exitItem = new ToolStripMenuItem("Exit");
        exitItem.Click += (s, e) => Application.Exit(); // Use Application.Exit() to close the app properly

        trayMenu.Items.Add(showItem);
        trayMenu.Items.Add(editItem);
        trayMenu.Items.Add(new ToolStripSeparator());
        trayMenu.Items.Add(trayToggleItem);
        trayMenu.Items.Add(idleToggleItem);
        trayMenu.Items.Add(reloadItem);
        trayMenu.Items.Add(openFolderItem);
        trayMenu.Items.Add(new ToolStripSeparator()); // Optional separator
        trayMenu.Items.Add(exitItem);

        // Refresh dynamic labels (idle timeout text) each time the menu opens.
        trayMenu.Opening += (s, e) => UpdateIdleToggleText();

        trayIcon = new NotifyIcon()
        {
            Icon = appIcon,
            // Icon = new Icon("path/to/your/icon.ico"), // Load a custom icon
            Text = windowName, // Tooltip text
            ContextMenuStrip = trayMenu, // Assign the context menu
            Visible = false // actual visibility is driven by UpdateTrayVisibility()
        };

        // Optional: Double-clicking the tray icon can show the launcher
        trayIcon.DoubleClick += (s, e) => ShowLauncher();
    }
    private int titleBarHeight = 26;
    public QuickLauncher()
    {
        Instance = this;

        // Load configuration
        LoadConfiguration();


        var panelWidth = 400;



        var panelColor = Color.FromArgb(30, 30, 30);
        var titleBarColor = Color.FromArgb(40, 40, 40);



        // Initialize form
        this.FormBorderStyle = FormBorderStyle.None;
        this.TopMost = true;
        this.ShowInTaskbar = false;
        this.StartPosition = FormStartPosition.Manual;
        this.BackColor = panelColor;
        this.Size = new Size(panelWidth, 600);
        this.WindowState = FormWindowState.Normal;
        this.MinimizeBox = false;
        this.MaximizeBox = false;




        titleLabel = new Label
        {
            Text = windowName,
            ForeColor = Color.White,
            //BackColor = Color.Red,
            TextAlign = ContentAlignment.MiddleCenter,
            //Location = new Point(5, 5),
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
            Dock = DockStyle.Fill,
            //Width = panelWidth,
            //AutoSize = true,
        };

        editButton = new Button
        {
            FlatStyle = FlatStyle.Flat,
            ForeColor = Color.White,
            Text = "Σ",
            Size = new Size(titleBarHeight, titleBarHeight - 4),
            Anchor = AnchorStyles.Right,
        };
        editButton.FlatAppearance.BorderSize = 0;
        editButton.Click += (s, e) => OpenEditor();

        closeButton = new Button
        {
            FlatStyle = FlatStyle.Flat,
            ForeColor = Color.White,
            Text = "✖",
            Size = new Size(titleBarHeight, titleBarHeight - 4),
            Anchor = AnchorStyles.Right // Anchor to the top-right
        };
        closeButton.FlatAppearance.BorderSize = 0;
        closeButton.Click += (s, e) => { CloseOrHide(); };

        // Title bar elements
        titleBarRHS = new FlowLayoutPanel
        {
            Anchor = AnchorStyles.Right,
            FlowDirection = FlowDirection.RightToLeft,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
        };

        titleBarRHS.Controls.Add(closeButton);
        titleBarRHS.Controls.Add(editButton);

        // Create title bar
        titleBar = new Panel
        {
            Height = titleBarHeight,
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
        };


        titleBar.Controls.Add(titleBarRHS);
        titleBar.Controls.Add(titleLabel);

        // Make title bar draggable
        titleLabel.MouseDown += (s, e) => { isDragging = true; lastCursor = Cursor.Position; lastForm = this.Location; };
        titleLabel.MouseMove += (s, e) => { if (isDragging) { this.Location = new Point(lastForm.X + (Cursor.Position.X - lastCursor.X), lastForm.Y + (Cursor.Position.Y - lastCursor.Y)); } };
        titleLabel.MouseUp += (s, e) => { isDragging = false; };


        // Create main content panel
        itemsPanel = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = true,
            AutoScroll = true,
            //Width = this.Width,
            //Height = this.Height,
            Dock = DockStyle.Fill,
            Padding = new Padding { Top = titleBarHeight },
            //BackColor = Color.Red            
        };

        mainPanel = new Panel
        {
            BackColor = panelColor,
            Dock = DockStyle.Fill,
        };

        mainPanel.Controls.Add(titleBar);
        mainPanel.Controls.Add(itemsPanel);
        this.Controls.Add(mainPanel);

        BuildUI();

        // Handle form events. Any keyboard/mouse interaction with the launcher
        // counts as activity and resets the idle-exit countdown (RecordActivity).
        this.Deactivate += (s, e) => this.Hide();
        this.KeyPreview = true;
        this.KeyDown += (s, e) => {
            RecordActivity();
            if (e.KeyCode == Keys.Escape)
            {
                if (navigationStack.Count > 0)
                {
                    // Go back to parent
                    navigationStack.Pop();
                    BuildUI();
                }
                else
                {
                    CloseOrHide();
                }
            }
        };

        // Center on screen initially
        CenterOnScreen();
        SetWindowPos(this.Handle, new IntPtr(-1), 0, 0, 0, 0, 0x1 | 0x2); // required to prevent app from closing automaticlaly when started if not focused and topmost.
        SetActiveWindow(this.Handle);

        // Install the global mouse/keyboard activity monitor that feeds the
        // idle-exit timer (low-level hooks + message filter fallback).
        ActivityMonitor.Install();

        // Coalesced in-app activity tracking for every form/dialog we pump
        // messages for (also covers the case where the low-level hooks fail).
        Application.AddMessageFilter(this);

        // Start the idle process-exiter countdown.
        InitializeIdleExitTimer();

        // DEBUG ONLY
        //OpenEditor();
        //System.Environment.Exit(0);

        // The tray icon is always created so the app remains reachable even
        // when UseTrayIcon is later toggled off/on without a restart; its
        // visibility is driven by the setting via UpdateTrayVisibility().
        InitializeTrayIcon();
        UpdateTrayVisibility();
    }

    /// <summary>
    /// Show or hide the tray icon according to <c>appSettings.UseTrayIcon</c>.
    /// Safe to call before the tray icon exists.
    /// </summary>
    private void UpdateTrayVisibility()
    {
        if (trayIcon != null)
            trayIcon.Visible = appSettings.UseTrayIcon;
    }

    /// <summary>
    /// Re-register (or unregister) the global hotkey from the current settings.
    /// </summary>
    private void RefreshGlobalHotkey()
    {
        try { UnregisterHotKey(this.Handle, HOTKEY_ID); } catch { }
        try { RegisterGlobalHotkey(); } catch { }
    }


    private void CloseOrHide()
    {
        RecordActivity(); // hiding via user action is still "usage"
        this.Hide();
        if (!appSettings.UseTrayIcon)
            Application.Exit(); // no tray to live in -> same shutdown as idle-exit
    }



    protected override void OnSizeChanged(EventArgs e)
    {
        base.OnSizeChanged(e);
    }

    private void LoadConfiguration()
    {
        if (!File.Exists(configPath))
        {
            CreateDefaultConfiguration();
            SaveConfiguration();
            return;
        }

        try
        {
            var serializer = new XmlSerializer(typeof(QuickLauncherSettings));
            using var reader = new StreamReader(configPath);
            appSettings = (QuickLauncherSettings)serializer.Deserialize(reader);
        }
        catch
        {
            appSettings = new QuickLauncherSettings();
        }

        if (appSettings.Items == null)
        {
            CreateDefaultConfiguration();
        }
    }

    /// <summary>
    /// Re-read the settings file from disk and rebuild the UI. Handy if you
    /// edit QuickLauncherSettings.xml by hand (or with another tool) without
    /// going through the editor.
    /// </summary>
    private void ReloadSettingsFromDisk()
    {
        navigationStack.Clear();
        LoadConfiguration();
        BuildUI();
        ApplyIdleSettings();
        UpdateTrayVisibility();
        RefreshGlobalHotkey();
    }

    private void SaveConfiguration()
    {
        var dir = Path.GetDirectoryName(configPath);
        if (!Directory.Exists(dir))
            Directory.CreateDirectory(dir);

        var serializer = new XmlSerializer(typeof(QuickLauncherSettings));
        using var writer = new StreamWriter(configPath);
        serializer.Serialize(writer, appSettings);
    }

    private void CreateDefaultConfiguration()
    {
        if (appSettings == null) appSettings = new QuickLauncherSettings();
        appSettings.Items = new QuickItemCollection()
        {
            Items = {
                new QuickItem { Name = "Notepad", Type = QuickItemType.App, Path = "notepad.exe", ShortcutKey = Keys.N },
                new QuickItem { Name = "Calculator", Type = QuickItemType.App, Path = "calc.exe", ShortcutKey = Keys.C },
                new QuickItem { Name = "Command Prompt", Type = QuickItemType.App, Path = "cmd.exe", ShortcutKey = Keys.D },
                new QuickItem { Name = "Utilities", Type = QuickItemType.Folder, ButtonColor = GetDefaultItemColor(QuickItemType.Folder), Items = new List<QuickItem>
                    {
                        new QuickItem { Name = "Paint", Type = QuickItemType.App, Path = "mspaint.exe", ShortcutKey = Keys.P },
                        new QuickItem { Name = "WordPad", Type = QuickItemType.App, Path = "write.exe", ShortcutKey = Keys.W },
                        new QuickItem { Name = "System Info", Type = QuickItemType.Command, ButtonColor = GetDefaultItemColor(QuickItemType.Command), Path = "systeminfo", ShortcutKey = Keys.I }
                    }
                }
            }
        };
    }



    private void CenterOnScreen()
    {
        var screen = Screen.FromPoint(Cursor.Position);
        var workingArea = screen.WorkingArea;
        this.Location = new Point(
            workingArea.Left + (workingArea.Width - this.Width) / 2,
            workingArea.Top + (workingArea.Height - this.Height) / 2
        );
    }

    public void ShowLauncher()
    {        
        RecordActivity(); // summoning the launcher counts as usage
        FixStack();
        CenterOnScreen();
        this.Show();
        this.BringToFront();
        this.Focus();
        FocusFirstButton();
    }

    private void FocusFirstButton()
    {
        if (itemsPanel.Controls.Count > 0)
        {
            ((Button)itemsPanel.Controls[0]).Focus();
        }
    }

    private void BuildUI()
    {
        // Button dimensions
        const int bw = 180;
        const int bh = 50;
        const int p = 3;

        itemsPanel.Controls.Clear();
        if (currentItems == null)
        {
            this.Size = new Size(bw + 2 * p, bh + 2 * p + titleBarHeight);
            return;
        }



        // Get monitor info and calculate aspect ratios
        var screen = Screen.FromControl(this); // Use 'this' or the mainPanel
        var workingArea = screen.WorkingArea;
        double a = (double)workingArea.Width / workingArea.Height;

        // Determine size
        var N = currentItems.Items.Count();

        double n = Math.Sqrt(N * a * bh / bw);
        int n_cols = Math.Max(1, (int)Math.Ceiling(n));
        int n_rows = (int)Math.Ceiling((double)N / n_cols);

        int calculated_width = n_cols * (bw + p) + p;
        int calculated_height = n_rows * (bh + p) + p;
        this.Size = new Size(calculated_width, calculated_height + titleBarHeight);

        foreach (var item in currentItems.Items)
        {
            var btn = new Button
            {
                Text = $"{item.Name}".Replace("\\n", "\n"),
                Width = bw,
                Height = bh,
                FlatStyle = FlatStyle.Flat,
                BackColor = (item.ButtonColor == Color.Transparent) ? GetDefaultItemColor(item.Type) : item.ButtonColor,
                ForeColor = item.FontColor,
                Font = item.Font,
                UseVisualStyleBackColor = false,
                Tag = item,
                Padding = new Padding(0, 0, 0, 0),
                Margin = new Padding(p, p, 0, 0),
                AutoSize = true,
                MaximumSize = new Size(bw, bh),
                MinimumSize = new Size(bw, bh),

            };

            QuickLauncherEditor.CTT(btn, item.Desc);


            // Shortcut Label
            var sLbl = new Label
            {
                Text = $"{sKey.GetKeyDisplay(item.ShortcutKey, true)}",
                // Create a new font based on the button's font family, but with a smaller size
                Font = new Font(btn.Font.FontFamily, 8F, FontStyle.Regular),
                Location = new Point(btn.Width - 50, btn.Height - 18), // Now X = btn.Width - label.Width, Y = btn.Height - label.Height
                                                                       // Set a fixed size for the label
                Size = new Size(50, 15), // Width = 20, Height = 15
                                         // Optional: Style the label
                ForeColor = item.FontColor, // Example: Different color for the shortcut
                BackColor = Color.Transparent, // So it doesn't obscure the button's background
                TextAlign = ContentAlignment.BottomRight // Align text within the label
            };
            btn.Controls.Add(sLbl);

            btn.FlatAppearance.BorderSize = 0;

            btn.Click += (s, e) => { RecordActivity(); ExecuteItem((QuickItem)((Button)s).Tag, (ModifierKeys == Keys.Control) ? QuickItemType.App : ((ModifierKeys == Keys.Shift) ? QuickItemType.Command : null)); };
            btn.KeyDown += ButtonKeyDown;
            btn.MouseEnter += (s, e) => {
                var btn = (Button)s;
                btn.ForeColor = (item.ButtonColor == Color.Transparent) ? GetDefaultItemColor(item.Type) : item.ButtonColor;
                btn.BackColor = item.FontColor;
                ((Button)s).Focus();
            };
            btn.MouseDown += (s, e) => RecordActivity(); // mouse interaction resets idle-exit timer
            btn.MouseLeave += (s, e) =>
            {
                var btn = (Button)s;
                btn.BackColor = (item.ButtonColor == Color.Transparent) ? GetDefaultItemColor(item.Type) : item.ButtonColor;
                btn.ForeColor = item.FontColor;
                ((Button)s).Focus();
            };
            itemsPanel.Controls.Add(btn);
        }

        CenterOnScreen();
    }

    public static Color GetDefaultItemColor(QuickItemType type)
    {
        switch (type)
        {
            case QuickItemType.App: return Color.FromArgb(50, 100, 50);
            case QuickItemType.Command: return Color.FromArgb(100, 50, 50);
            case QuickItemType.Shortcut: return Color.FromArgb(50, 50, 100);
            case QuickItemType.Folder: return Color.FromArgb(100, 100, 50);
            default: return Color.FromArgb(50, 50, 50);
        }
    }


    /*
     * additional processinfo flags that might be useful or should be optional
    UseShellExecute = false, // Can be true, but false might be preferred for hiding cmd window
    CreateNoWindow = true, // Hide the schtasks window
    WindowStyle = ProcessWindowStyle.Hidden // Alternative to CreateNoWindow
    */

    private void ExecuteItem(QuickItem item, QuickItemType? forceType = null)
    {
        RecordActivity();
        var type = (forceType.HasValue) ? forceType.Value : item.Type;
        var path = item.Path;
        var args = item.Args;
        var workingDir = item.WorkingDir;
        var se = item.useShellExecute;
        var cnw = item.noWindow;

    top:
        switch (type)
        {
            case QuickItemType.App:
                if (path != "")
                {
                    stopAutoClose = true;
                    try
                    {
                        var si = new ProcessStartInfo() {
                            WorkingDirectory = (workingDir == "") ? Path.GetDirectoryName(item.Path) : workingDir,
                            FileName = path,
                            Arguments = args,
                            Verb = (item.RunAsAdmin) ? "runas" : "",
                            UseShellExecute = se,
                            CreateNoWindow = cnw,
                            WindowStyle = (cnw) ? ProcessWindowStyle.Hidden : ProcessWindowStyle.Normal,
                        };
                        var p = Process.Start(si);

                        this.Hide();
                        if (item.Close)
                            CloseOrHide();
                        else
                        {
                            p.WaitForExit();
                            SetWindowPos(this.Handle, new IntPtr(-1), 0, 0, 0, 0, 0x1 | 0x2);
                            this.Show();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error executing app: {ex.Message}");
                        SetWindowPos(this.Handle, new IntPtr(0), 0, 0, 0, 0, 0x1 | 0x2);
                        this.Show();
                        stopAutoClose = false;
                    }
                    stopAutoClose = false;
                }
                break;
            case QuickItemType.Command:
                if (path != "")
                {
                    ExecuteCommand(item);
                    if (item.Close)
                        CloseOrHide();
                }
                break;
            case QuickItemType.Shortcut:
                // SendKeys targets the foreground window, so the launcher must
                // be hidden first or it would receive its own keystrokes.
                this.Hide();
                System.Threading.Thread.Sleep(50); // let activation settle
                try { SendKeys.SendWait(item.Path); }
                catch (Exception ex) { MessageBox.Show($"Error sending keys: {ex.Message}"); }
                if (item.Close)
                    CloseOrHide();
                else
                    this.Show();
                break;
            case QuickItemType.GlobalShortcut:
                // Synthesize a system-wide key combo via keybd_event. The Path
                // holds the combo with modifiers (e.g. "Ctrl+Alt+Delete",
                // "Win+D"); Args is unused.
                this.Hide(); // give the foreground to whatever receives it
                System.Threading.Thread.Sleep(50);
                try { SendGlobalShortcut(item.Path); }
                catch (Exception ex) { MessageBox.Show($"Error sending global shortcut: {ex.Message}"); }
                if (item.Close)
                    CloseOrHide();
                else
                    this.Show();
                break;
            case QuickItemType.Folder:
                navigationStack.Push(item);
                BuildUI();
                break;
            case QuickItemType.Task:
                type = QuickItemType.App;
                path = "C:\\Windows\\System32\\schtasks.exe";
                // Quote the task name too so names with spaces/special chars are safe.
                args = "/Run /TN \"" + (item.TaskName ?? "").Replace("\"", "\\\"") + "\"";
                workingDir = "";
                goto top;
                break;
        }
    }

    // ------------------------------------------------------------------
    // GlobalShortcut support: parse a combo string like "Ctrl+Alt+F1" or
    // "Win+D" into VK codes and inject it system-wide with keybd_event.
    // ------------------------------------------------------------------
    [DllImport("user32.dll")]
    private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

    private const uint KEYEVENTF_KEYUP = 0x0002;
    private const uint KEYEVENTF_EXTENDEDKEY = 0x0001;

    private static Keys? ParseShortcutToken(string token)
    {
        token = token.Trim();
        switch (token.ToLowerInvariant())
        {
            case "ctrl": case "control": return Keys.Control;
            case "alt": case "menu": return Keys.Alt;
            case "shift": return Keys.Shift;
            case "win": case "windows": return Keys.LWin;
        }
        if (token.Length == 1)
        {
            char c = char.ToUpperInvariant(token[0]);
            if ((c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9')) return (Keys)c;
        }
        if (Enum.TryParse<Keys>(token, true, out var parsed)) return parsed;
        return null;
    }

    private static void SendGlobalShortcut(string combo)
    {
        if (string.IsNullOrWhiteSpace(combo)) return;
        var parts = combo.Split(new[] { '+' }, StringSplitOptions.RemoveEmptyEntries);
        var mods = new List<Keys>();
        Keys main = Keys.None;
        foreach (var p in parts)
        {
            var k = ParseShortcutToken(p);
            if (k == null) continue;
            var kk = k.Value;
            if (kk == Keys.Control || kk == Keys.Alt || kk == Keys.Shift || kk == Keys.LWin || kk == Keys.RWin)
                mods.Add(kk);
            else
                main = kk;
        }
        if (main == Keys.None) return;

        void Tap(Keys vk, bool down) =>
            keybd_event((byte)vk, 0, (down ? 0u : KEYEVENTF_KEYUP) | (((uint)vk & 0xF0) == 0x70 ? KEYEVENTF_EXTENDEDKEY : 0), UIntPtr.Zero);

        foreach (var m in mods) Tap(m, true);
        Tap(main, true);
        Tap(main, false);
        for (int i = mods.Count - 1; i >= 0; i--) Tap(mods[i], false);
    }

    private void ExecuteCommand(QuickItem item)
    {
        var processInfo = new ProcessStartInfo
        {
            FileName = "cmd.exe",
            Arguments = $"/C {item.Path} {item.Args}",
            UseShellExecute = item.useShellExecute,
            CreateNoWindow = item.noWindow,
            WindowStyle = (item.noWindow) ? ProcessWindowStyle.Hidden : ProcessWindowStyle.Normal,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            WorkingDirectory = (item.WorkingDir == "") ? Path.GetDirectoryName(item.Path) : item.WorkingDir,
            Verb = (item.RunAsAdmin) ? "runas" : "",
        };

        //processInfo.ArgumentList.Add(item.Args);

        try
        {
            using (var process = Process.Start(processInfo))
            {

                stopAutoClose = true;
                this.Hide();
                SetWindowPos(this.Handle, new IntPtr(-1), 0, 0, 0, 0, 0x1 | 0x2);
                process.WaitForExit();
                SetWindowPos(this.Handle, new IntPtr(0), 0, 0, 0, 0, 0x1 | 0x2);
                this.Show();
                stopAutoClose = false;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error executing command: {ex.Message}");
            SetWindowPos(this.Handle, new IntPtr(0), 0, 0, 0, 0, 0x1 | 0x2);
            this.Show();
            stopAutoClose = false;
        }
    }

    private void ButtonKeyDown(object sender, KeyEventArgs e)
    {
        RecordActivity(); // keyboard interaction resets idle-exit timer
        var btn = (Button)sender;
        var item = (QuickItem)btn.Tag;
        var idx = itemsPanel.Controls.IndexOf(btn);

        switch (e.KeyCode)
        {
            case Keys.Enter:
                ExecuteItem(item);
                break;
            case Keys.Tab:
                var next = (idx + 1) % itemsPanel.Controls.Count;
                ((Button)itemsPanel.Controls[next]).Focus();
                break;
            case Keys.Up:
                if (idx > 0) ((Button)itemsPanel.Controls[idx - 1]).Focus();
                break;
            case Keys.Down:
                if (idx < itemsPanel.Controls.Count - 1) ((Button)itemsPanel.Controls[idx + 1]).Focus();
                break;
        }

        // Handle key bindings
        foreach (var i in currentItems.Items)
        {
            if (i.ShortcutKey == e.KeyCode)
            {
                ExecuteItem(i);
                return;
            }
        }
    }


    private IEnumerable<QuickItemCollection> GetAllCollections(QuickItemCollection collection)
    {
        yield return collection;
        foreach (var item in collection.Items.Where(i => i.Type == QuickItemType.Folder))
        {
            // Pass the folder's Items list as a new QuickItemCollection
            foreach (var nested in GetAllCollections(item))
            {
                yield return nested;
            }
        }
    }


    // Inside the QuickLauncher class
    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        RecordActivity(); // any key reaching the launcher counts as usage

        // Check for Ctrl+Shift+L to open editor
        if (keyData == (Keys.Control | Keys.Shift | Keys.L))
        {
            OpenEditor();
            return true;
        }

        // Handle shortcut keys in current view. Strip modifier-only noise so
        // plain-key shortcuts still match when e.g. Shift is held for letters.
        foreach (var item in currentItems.Items)
        {
            if (item.ShortcutKey == Keys.None) continue;
            if (item.ShortcutKey == keyData ||
                ((item.ShortcutKey & Keys.KeyCode) == (keyData & Keys.KeyCode) &&
                 (item.ShortcutKey & Keys.Modifiers) == Keys.None))
            {
                ExecuteItem(item);
                return true;
            }
        }
        return base.ProcessCmdKey(ref msg, keyData);
    }


    protected override void WndProc(ref Message m)
    {
        // Any input message that reaches this window (WM_KEY*, WM_MOUSE*, wheel, etc.)
        // counts as usage and resets the idle-exit countdown. Cheap timestamp write.
        if (m.Msg >= 0x0100 && m.Msg <= 0x02FF) RecordActivity();

        const int WM_ACTIVATEAPP = 0x1C;
        if (m.Msg == WM_ACTIVATEAPP)
        {
            // Close add when lost focus
            if (m.WParam == IntPtr.Zero)
            {
                if (!stopAutoClose)
                    CloseOrHide();
            }

        }

        if (m.Msg == WM_HOTKEY)
        {
            int id = m.WParam.ToInt32();
            if (id == HOTKEY_ID)
            {
                RecordActivity(); // pressing the hotkey is user interaction
                if (this.Visible)
                {
                    CloseOrHide();
                }
                else
                {
                    ShowLauncher();
                }
                return;
            }
        }

        // Handle the custom message from another instance (wParam = requested verb)
        if (m.Msg == WM_SHOW_OR_HIDE_LAUNCHER)
        {
            int verb = m.WParam.ToInt32();

            if (verb == SHOWMSG_EDIT)
            {
                if (!this.Visible) ShowLauncher();
                OpenEditor();
                return;
            }

            if (verb == SHOWMSG_SHOW && !this.Visible)
            {
                ShowLauncher();
                return;
            }

            // Toggle visibility based on current state (SHOWMSG_TOGGLE, or legacy behaviour)
            if (this.Visible)
            {
                if (appSettings.closeInsteadOfNavigate)
                    CloseOrHide();
                else
                {
                    if (navigationStack.Count > 0)
                    {
                        // Go back to parent
                        navigationStack.Pop();
                        BuildUI();
                    }
                    else
                    {
                        CloseOrHide();
                    }
                }

            }
            else
            {
                ShowLauncher();
            }
            return; // Don't pass the message to base class
        }

        base.WndProc(ref m);
    }

    private void RegisterGlobalHotkey()
    {
        // Define modifiers (Win = 0x0008, Ctrl = 0x0002, Alt = 0x0001, Shift = 0x0004)
        // Combine them using bitwise OR. Example: Ctrl+Alt = 0x0002 | 0x0001 = 0x0003
        uint modifiers = 0;
        if ((registeredHotkey & Keys.Control) == Keys.Control) modifiers |= 0x0002;
        if ((registeredHotkey & Keys.Alt) == Keys.Alt) modifiers |= 0x0001;
        if ((registeredHotkey & Keys.Shift) == Keys.Shift) modifiers |= 0x0004;
        // MOD_NOREPEAT (0x4000): don't let auto-repeat re-trigger the toggle.
        modifiers |= 0x4000;
        // Note: Windows key is 0x0008, but requires special handling and might conflict

        Keys key = registeredHotkey & ~Keys.Control & ~Keys.Alt & ~Keys.Shift; // Get the main key

        // Register the hotkey
        bool result = RegisterHotKey(this.Handle, HOTKEY_ID, modifiers, key);
        if (!result)
        {
            // Handle registration failure (e.g., hotkey already registered by another app)
            MessageBox.Show($"Failed to register hotkey {registeredHotkey}. It might already be in use.", "Hotkey Registration Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }


    // Add a method to find the target collection based on a sequence of Keys
    private QuickItemCollection FindCollectionByKeys(QuickItemCollection startCollection, List<Keys> keySequence)
    {
        QuickItemCollection currentCollection = startCollection;

        foreach (var targetKey in keySequence)
        {
            // Find the first item in the current collection that matches the target key
            var targetItem = currentCollection.Items.FirstOrDefault(item => item.ShortcutKey == targetKey);

            if (targetItem != null && targetItem.Type == QuickItemType.Folder)
            {
                // If found and it's a folder, move into its Items collection
                currentCollection = targetItem;
            }
            else if (targetItem != null)
            {
                if (targetItem.Type != QuickItemType.Folder)
                {
                    // Path led to a non-folder item, cannot navigate further into it as a collection.
                    // You might want to execute the item here instead, but for "changing root", it fails.
                    return null; // Indicate failure to find a valid folder root
                }

                if (targetItem.Type != QuickItemType.Folder)
                {
                    // The key resolved to a non-folder item, cannot navigate further as a root.
                    return null; // Path is invalid for folder navigation
                }
                // If it was a folder, currentCollection was updated above.
            }
            else
            {
                // Key not found in the current collection
                return null; // Path is invalid
            }
        }

        // If we successfully navigated through all keys and ended at a folder collection
        return currentCollection;
    }

    public void FixStack()
    {
        // try to find closest path to original
        var n = navigationStack.Reverse().ToArray();
        navigationStack.Clear();

        var cis = rootItems.Items;
        nest:
        if (n.Length > 0)
        {
            foreach (var item in cis)
            {
                if (n[0].Name == item.Name)
                {
                    n = n[1..];
                    navigationStack.Push(item);
                    cis = item.Items;
                    goto nest;
                }
            }
        }
    }

    private bool stopAutoClose = false;
    // Public wrapper used by the Shown event handler (CLI "/edit" verb).
    public void OpenEditorViaShown() => OpenEditor();

    private void OpenEditor()
    {
        RecordActivity(); // opening the editor counts as usage
        stopAutoClose = true;
        this.Hide(); // Hide the launcher while editing                        
        var editor = new QuickLauncherEditor(appSettings, navigationStack);
        if (editor.ShowDialog() == DialogResult.OK)
        {
            appSettings.Items = editor.rootCollection;
            // The editor modifies the rootItems directly
            SaveConfiguration(); // Save changes after editor closes with OK
            
            if (navigationStack.Count == 0)
            {
                // Refresh the current view if it's the root
                navigationStack.Clear();
            } else
            {
                FixStack();
            }

            try
            {
                UnregisterHotKey(this.Handle, HOTKEY_ID);
            }
            catch { }

            try
            {
                RegisterGlobalHotkey(); // hotkey may have changed in the editor
            }
            catch { }

            // Tray visibility / idle-exit settings changed in the editor take effect immediately.
            UpdateTrayVisibility();
            ApplyIdleSettings();
        }

        BuildUI();
        this.Show();
        stopAutoClose = false;
    }




    [STAThread]
    public static void Main(string[] args)
    {
        const string appName = "QuickLauncher"; // Unique name for your application
                                                // Use a unique window name/class for FindWindow. We'll use the form's Text property.



        using (var mutex = new Mutex(true, appName, out bool createdNew))
        {
            if (!createdNew)
            {
                // Another instance is already running
                // Find the existing instance's window
                IntPtr existingWindowHandle = FindWindow(null, windowName); // lpClassName is null, use window name

                if (existingWindowHandle != IntPtr.Zero)
                {
                    // Send the custom message to the existing instance
                    int verb = ShowMessageWParamForArgs(args);
                    bool messageSent = PostMessage(existingWindowHandle, WM_SHOW_OR_HIDE_LAUNCHER, (IntPtr)verb, IntPtr.Zero);

                    if (messageSent)
                    {
                        // Optionally, try to ensure the existing window is brought to the foreground if it's now visible
                        // This might be necessary depending on focus rules
                        ShowWindow(existingWindowHandle, SW_RESTORE);
                        SetForegroundWindow(existingWindowHandle);
                        SetWindowPos(existingWindowHandle, new IntPtr(-1), 0, 0, 0, 0, 0x1 | 0x2);
                        return;
                    }
                    else
                    {
                        // If sending the message failed, fallback might be to just show a message
                        // Console.WriteLine("Failed to communicate with existing instance.");
                        return;
                    }
                }
                else
                {
                    // Could not find the window handle, maybe it's starting up or has a different title?
                    // Console.WriteLine("Existing instance window not found.");
                }
                return; // Exit this new instance
            }

            // We are the primary instance
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            var form = new QuickLauncher(); // Pass true for primary instance
            form.Text = windowName;
            // If launched with a verb (e.g. "QuickLaunch.exe /edit"), honour it right away.
            int verb = ShowMessageWParamForArgs(args);
            if (verb == SHOWMSG_EDIT)
                form.Shown += (s, e) => form.OpenEditorViaShown();
            else if (verb == SHOWMSG_SHOW)
                form.Shown += (s, e) => form.ShowLauncher();
            Application.Run(form);

            // The 'using' statement ensures the mutex is released when the application exits
            // or when the 'using' block is exited
        }
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            // Unregister the hotkey and activity hooks when the application closes
            try
            {
                ActivityMonitor.Uninstall();
                UnregisterHotKey(this.Handle, HOTKEY_ID);

                trayIcon?.Dispose();
                trayMenu?.Dispose();
            }
            catch { }

        }
        base.Dispose(disposing);
    }


}


[Serializable]
public class QuickLauncherSettings
{

    public const int DefaultIdleTimeoutSeconds = 60;

    public Keys GlobalHotkey { get; set; } = Keys.Control | Keys.Alt | Keys.L;
    public bool UseTrayIcon { get; set; } = false;
    public bool closeInsteadOfNavigate { get; set; } = false; // Specifies that the app will close rather than navigate up the stack to top and then closed when called. 

    // --- Idle process-exiter ---
    // When enabled, the app quits itself after it has been hidden AND there has
    // been no keyboard/mouse interaction with the launcher for IdleTimeoutSeconds.
    public bool AutoExitWhenIdle { get; set; } = true;
    public int IdleTimeoutSeconds { get; set; } = DefaultIdleTimeoutSeconds;

    public QuickItemCollection Items { get; set; } = new QuickItemCollection();
}


// Serializable classes for configuration
[Serializable]
public class QuickItemCollection
{
    public string Name { get; set; }
    public List<QuickItem> Items { get; set; } = new List<QuickItem>();
}



public enum QuickItemType
{
    App,        // Runs an app
    Command,    // Executes a command
    Shortcut,   // Triggers a shortcut key that is send to the active app (uses PostMessage)
    GlobalShortcut, // Triggers a global shortcut key
    Folder,     // Represents a folder that contains more items    
    Task,   // Run a windows task
}


[Serializable]
public class QuickItem : QuickItemCollection
{
    // Don't serialize the Font object directly
    [XmlIgnore]
    private Font _font = null; // Cache the reconstructed font

    // Store font properties for serialization
    public string FontFamilyName { get; set; } = FontFamily.GenericSansSerif.Name;
    public float FontSize { get; set; } = 12.0f;
    public FontStyle FontStyle { get; set; } = FontStyle.Regular;

    public string Desc { get; set; }
    public bool Close { get; set; } = true;

    public bool RunAsAdmin { get; set; } = false;
    public bool useShellExecute { get; set; } = false;
    public bool noWindow { get; set; } = false;



    public QuickItemType Type { get; set; }
    public string Path { get; set; }
    public string Args { get; set; }
    public string WorkingDir { get; set; }

    public string TaskName { get; set; }

    public Keys ShortcutKey { get; set; } = Keys.None;
    public float SortIndex { get; set; } = 0.0f;




    // Store colors as ARGB integers for reliable serialization
    private int _fontColorArgb = Color.White.ToArgb();
    private int _buttonColorArgb = QuickLauncher.GetDefaultItemColor(QuickItemType.App).ToArgb();

    // Public properties that convert between Color and ARGB
    [XmlIgnore] // Don't serialize this directly
    public Color FontColor
    {
        get { return Color.FromArgb(_fontColorArgb); }
        set { _fontColorArgb = value.ToArgb(); }
    }

    [XmlIgnore] // Don't serialize this directly
    public Color ButtonColor
    {
        get { return Color.FromArgb(_buttonColorArgb); }
        set { _buttonColorArgb = value.ToArgb(); }
    }

    // XML-serializable properties for the ARGB values
    [XmlElement("FontColorArgb")] // Give it a specific name for the XML
    public int FontColorArgb
    {
        get { return _fontColorArgb; }
        set { _fontColorArgb = value; }
    }

    [XmlElement("ButtonColorArgb")] // Give it a specific name for the XML
    public int ButtonColorArgb
    {
        get { return _buttonColorArgb; }
        set { _buttonColorArgb = value; }
    }



    // Public property to get/set the font, handling reconstruction
    [XmlIgnore] // Don't serialize this property
    public Font Font
    {
        get
        {
            if (_font == null || _font.FontFamily.Name != FontFamilyName || _font.Size != FontSize || _font.Style != FontStyle)
            {
                try
                {
                    _font = new Font(new FontFamily(FontFamilyName), FontSize, FontStyle);
                }
                catch
                {
                    // Fallback if the specific font family is not available
                    _font = new Font(FontFamily.GenericSansSerif, FontSize, FontStyle);
                }
            }
            return _font;
        }
        set
        {
            if (value != null)
            {
                FontFamilyName = value.FontFamily.Name;
                FontSize = value.Size;
                FontStyle = value.Style;
                _font = value; // Cache the new font object
            }
        }
    }

    // Optional: A helper to reset the cached font if properties change individually
    private void InvalidateFontCache()
    {
        _font = null;
    }

    // Ensure the cache is invalidated when deserialized
    public void OnDeserialized()
    {
        InvalidateFontCache();
    }
}


// ---------------------------------------------------------------------------
// ActivityMonitor: tracks the timestamp of the last user keyboard/mouse
// interaction using low-level global hooks (WH_KEYBOARD_LL / WH_MOUSE_LL).
// The QuickLauncher idle-exit timer consults LastActivityAgeMs to decide
// whether the app has been unused long enough to quit.
//
// Design notes:
//  * Hooks are installed once at startup and uninstalled on form disposal.
//  * Callbacks only write a tick count (Interlocked.Exchange) - they never
//    allocate, block, or touch UI, which keeps hook latency minimal.
//  * If hook installation fails for any reason, we degrade gracefully:
//    QuickLauncher still records activity from its own WndProc/events, so the
//    exit timer works based on in-app usage alone.
// ---------------------------------------------------------------------------
public static class ActivityMonitor
{
    private const int WH_KEYBOARD_LL = 13;
    private const int WH_MOUSE_LL = 14;
    private const int WM_KEYDOWN = 0x0100;
    private const int WM_SYSKEYDOWN = 0x0104;
    private const int WM_LBUTTONDOWN = 0x0201;
    private const int WM_RBUTTONDOWN = 0x0204;
    private const int WM_MBUTTONDOWN = 0x0207;
    private const int WM_MOUSEWHEEL = 0x020A;

    private delegate IntPtr LowLevelHookProc(int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelHookProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr GetModuleHandle(string lpModuleName);

    // Keep references alive so the delegates are not garbage collected.
    private static LowLevelHookProc kbProc;
    private static LowLevelHookProc msProc;
    private static IntPtr kbHook = IntPtr.Zero;
    private static IntPtr msHook = IntPtr.Zero;

    private static long lastActivityTicks = DateTime.UtcNow.Ticks;

    /// <summary>Milliseconds elapsed since the last recorded activity.</summary>
    public static double LastActivityAgeMs =>
        (DateTime.UtcNow.Ticks - Interlocked.Read(ref lastActivityTicks)) * 1000.0 / TimeSpan.TicksPerMillisecond;

    public static void RecordActivity()
    {
        Interlocked.Exchange(ref lastActivityTicks, DateTime.UtcNow.Ticks);
    }

    public static void Install()
    {
        try
        {
            if (kbHook == IntPtr.Zero)
            {
                kbProc = KeyboardHookCallback;
                kbHook = SetWindowsHookEx(WH_KEYBOARD_LL, kbProc, GetModuleHandle(null), 0);
            }
            if (msHook == IntPtr.Zero)
            {
                msProc = MouseHookCallback;
                msHook = SetWindowsHookEx(WH_MOUSE_LL, msProc, GetModuleHandle(null), 0);
            }
        }
        catch
        {
            // Hooking failed - fall back to in-app event tracking only.
        }
    }

    public static void Uninstall()
    {
        try
        {
            if (kbHook != IntPtr.Zero) { UnhookWindowsHookEx(kbHook); kbHook = IntPtr.Zero; }
            if (msHook != IntPtr.Zero) { UnhookWindowsHookEx(msHook); msHook = IntPtr.Zero; }
        }
        catch { }
    }

    private static IntPtr KeyboardHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0)
        {
            int msg = (int)(long)wParam;
            if (msg == WM_KEYDOWN || msg == WM_SYSKEYDOWN)
                RecordActivity();
        }
        return CallNextHookEx(kbHook, nCode, wParam, lParam);
    }

    private static IntPtr MouseHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0)
        {
            int msg = (int)(long)wParam;
            // Button clicks & wheel only (not raw movement): moving the mouse
            // across the screen without clicking should NOT keep the app alive.
            if (msg == WM_LBUTTONDOWN || msg == WM_RBUTTONDOWN ||
                msg == WM_MBUTTONDOWN || msg == WM_MOUSEWHEEL)
                RecordActivity();
        }
        return CallNextHookEx(msHook, nCode, wParam, lParam);
    }
}
