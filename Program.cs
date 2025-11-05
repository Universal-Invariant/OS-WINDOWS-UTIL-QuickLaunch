using System;
using System.Collections.Generic;
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


public class QuickLauncher : Form
{

    static string windowName = "Quick Launcher"; // Should match the titleLabel.Text or Form.Text    
    private NotifyIcon trayIcon;
    private ContextMenuStrip trayMenu;

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

        // Register the global hotkey when the form loads
        if (appSettings.UseTrayIcon)
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
        var exitItem = new ToolStripMenuItem("Exit");
        exitItem.Click += (s, e) => Application.Exit(); // Use Application.Exit() to close the app properly

        trayMenu.Items.Add(showItem);
        trayMenu.Items.Add(editItem);
        trayMenu.Items.Add(new ToolStripSeparator()); // Optional separator
        trayMenu.Items.Add(exitItem);

        trayIcon = new NotifyIcon()
        {
            Icon = appIcon,
            // Icon = new Icon("path/to/your/icon.ico"), // Load a custom icon
            Text = windowName, // Tooltip text
            ContextMenuStrip = trayMenu, // Assign the context menu
            Visible = true // Make the icon visible immediately when the form starts
        };

        // Optional: Double-clicking the tray icon can show the launcher
        trayIcon.DoubleClick += (s, e) => ShowLauncher();
    }
    private int titleBarHeight = 26;
    public QuickLauncher()
    {

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

        // Handle form events
        this.Deactivate += (s, e) => this.Hide();
        this.KeyPreview = true;
        this.KeyDown += (s, e) => {
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
        // DEBUG ONLY
        //OpenEditor();
        //System.Environment.Exit(0);

        if (appSettings.UseTrayIcon)
            InitializeTrayIcon();
    }


    private void CloseOrHide()
    {

        this.Hide();
        if (!appSettings.UseTrayIcon)
            this.Close();
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

            btn.Click += (s, e) => ExecuteItem((QuickItem)((Button)s).Tag, (ModifierKeys == Keys.Control) ? QuickItemType.App : ((ModifierKeys == Keys.Shift) ? QuickItemType.Command : null));
            btn.KeyDown += ButtonKeyDown;
            btn.MouseEnter += (s, e) => {
                var btn = (Button)s;
                btn.ForeColor = (item.ButtonColor == Color.Transparent) ? GetDefaultItemColor(item.Type) : item.ButtonColor;
                btn.BackColor = item.FontColor;
                ((Button)s).Focus();
            };
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
                SendKeys.SendWait(item.Path);
                if (item.Close)
                    CloseOrHide();
                break;
            case QuickItemType.Folder:
                navigationStack.Push(item);
                BuildUI();
                break;
            case QuickItemType.Task:
                type = QuickItemType.App;
                path = "C:\\Windows\\System32\\schtasks.exe";
                args = "/Run /TN \"" + item.TaskName + "\"";
                workingDir = "";
                goto top;
                break;
        }
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
        // Check for Ctrl+Shift+L to open editor
        if (keyData == (Keys.Control | Keys.Shift | Keys.L))
        {
            OpenEditor();
            return true;
        }

        // Handle shortcut keys in current view
        foreach (var item in currentItems.Items)
        {
            if (item.ShortcutKey == keyData)
            {
                ExecuteItem(item);
                return true;
            }
        }
        return base.ProcessCmdKey(ref msg, keyData);
    }


    protected override void WndProc(ref Message m)
    {
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

        // Handle the custom message from another instance
        if (m.Msg == WM_SHOW_OR_HIDE_LAUNCHER)
        {
            // Toggle visibility based on current state
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
    private void OpenEditor()
    {
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
                if (appSettings.UseTrayIcon)
                    RegisterGlobalHotkey();
            }
            catch { }
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
                    bool messageSent = PostMessage(existingWindowHandle, WM_SHOW_OR_HIDE_LAUNCHER, IntPtr.Zero, IntPtr.Zero);

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
            Application.Run(form);

            // The 'using' statement ensures the mutex is released when the application exits
            // or when the 'using' block is exited
        }
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            // Unregister the hotkey when the application closes
            try
            {
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

    public Keys GlobalHotkey { get; set; } = Keys.Control | Keys.Alt | Keys.L;
    public bool UseTrayIcon { get; set; } = false;
    public bool closeInsteadOfNavigate { get; set; } = false; // Specifies that the app will close rather than navigate up the stack to top and then closed when called. 

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

