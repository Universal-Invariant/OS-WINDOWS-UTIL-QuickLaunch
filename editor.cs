using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing.Text;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Xml.Serialization;


using TagT = System.Tuple<QuickItem, System.Windows.Forms.TreeNode>;





public partial class QuickLauncherEditor : Form
{
    private QuickLauncherSettings appSettings;
    public QuickItemCollection rootCollection;
    private QuickItemCollection orgCollection;
    private TreeView treeView;
    private TextBox nameTextBox;
    private TextBox descTextBox;
    private TextBox sortIndexTextBox;
    private ComboBox typeComboBox;
    private TextBox pathTextBox;
    private TextBox argsTextBox;
    private TextBox workingDirTextBox;
    private TextBox shortcutKeyTextBox; // New control for shortcut key
    private TextBox globalShortCutTextBox;
    private TextBox taskNameTextBox;
    private Button addButton;
    private Button deleteButton;
    private Button saveButton;
    private Button cancelButton;
    private CheckBox useTrayCheckBox;
    private CheckBox closeCheckBox;
    private CheckBox runAsAdminCheckBox;
    private CheckBox useShellExecuteCheckBox;
    private CheckBox noWindowCheckBox;
    private TreeNode dragNode; // For drag-and-drop

    private NumericUpDown fontSizeUpDown; // For font size
    private Button fontColorButton;       // For font color
    private Button buttonColorButton;     // For button color
    private Button selectFontButton;      // For font family/style
    private Font currentSelectedFont = SystemFonts.DefaultFont; // Track selected font family/style
    private List<QuickItemCollection> navigationStack;
    // We need this for the auto-scroll functionality in DragOver.
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wp, IntPtr lp);

    public T DeepCopy<T>(ref T object2Copy)
    {
        T objectCopy;
        using (var stream = new MemoryStream())
        {
            var serializer = new XmlSerializer(typeof(T));

            serializer.Serialize(stream, object2Copy);
            stream.Position = 0;
            objectCopy = (T)serializer.Deserialize(stream);
        }

        return objectCopy;
    }

    public Control prevSibling(Control c)
    {
        var i = c.Parent.Controls.IndexOf(c);
        if (i <= 0) return null;
        return c.Parent.Controls[i - 1];
    }


    public QuickLauncherEditor(QuickLauncherSettings appSettings, Stack<QuickItemCollection> navigationStack)
    {
        this.appSettings = appSettings;
        orgCollection = appSettings.Items;
        rootCollection = DeepCopy(ref orgCollection);
        this.navigationStack = navigationStack.Reverse().ToList();
        InitializeComponent();
        BuildTreeView(null, null, this.navigationStack);

        useTrayCheckBox.Checked = appSettings.UseTrayIcon;
    }

    public static ToolTip CTT(Control ctrl, string Desc)
    {
        ToolTip toolTip = new ToolTip();

        // Set up the delays for the ToolTip.
        toolTip.AutoPopDelay = 5000;
        toolTip.InitialDelay = 1000;
        toolTip.ReshowDelay = 500;
        // Force the ToolTip text to be displayed whether or not the form is active.
        toolTip.ShowAlways = true;

        // Set up the ToolTip text for the Button and Checkbox.
        toolTip.SetToolTip(ctrl, Desc);
        return toolTip;
    }

    // Helper function - Focus on FlowLayoutPanel with Dock.Fill
    private TableLayoutPanel nLTB(string label, out TextBox tb, string toolTip = "", int Width = -1)
    {
        var f = new TableLayoutPanel
        {
            // Crucially, fill the cell in the outer TableLayoutPanel
            Dock = DockStyle.Fill,
            // Ensure horizontal flow
            Margin = new Padding(0),
            RowCount = 1,
            ColumnCount = 2,
            Padding = new Padding(0),
            Height = 28,
        };
        if (Width != -1) f.Width = Width;

        // Label: AutoSize=true is critical to size to content
        var l = new Label
        {
            Text = label,
            Dock = DockStyle.Left,
            Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top | AnchorStyles.Bottom,
            TextAlign = ContentAlignment.MiddleLeft,
            AutoSize = true,
        };
        f.Controls.Add(l, 0, 0);

        // TextBox: Dock=Fill is critical to take remaining space
        tb = new TextBox
        {
            Text = "",
            Dock = DockStyle.Fill,
        };        
        f.Controls.Add(tb, 1, 0);
        if (toolTip != "")
            CTT(tb, toolTip);
        return f;
    }

    private void InitializeComponent()
    {
        this.Text = "Quick Launcher Editor";
        this.Size = new System.Drawing.Size(900, 700); // Larger initial size
        this.StartPosition = FormStartPosition.CenterParent;

        // Create controls
        var splitContainer = new SplitContainer { BackColor = SystemColors.Control, SplitterWidth = 5, Dock = DockStyle.Fill };
        var LHSContainer = new FlowLayoutPanel();

        // Left panel for tree
        treeView = new TreeView
        {
            Anchor = AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Left,
            Dock = DockStyle.Fill,
            HideSelection = false,
            LabelEdit = true, // Allows in-place renaming
            AllowDrop = true // Enable drag-and-drop
        };
        treeView.AfterLabelEdit += TreeView_AfterLabelEdit;
        treeView.AfterSelect += TreeView_UpdateFormAfterSelect;

        // Drag-and-drop handlers
        treeView.ItemDrag += TreeView_ItemDrag;
        treeView.DragEnter += TreeView_DragEnter;
        treeView.DragDrop += TreeView_DragDrop;
        treeView.DragOver += TreeView_DragOver;

        splitContainer.Panel1.Controls.Add(treeView);



        // Right panel for properties
        var rightPanel = new Panel { Dock = DockStyle.Fill, BackColor = SystemColors.Control };
        var editPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.TopDown,
            Padding = new Padding(10),
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
        };

        // Name        
        editPanel.Controls.Add(nLTB("Name:", out nameTextBox, "Label for action"));
        nameTextBox.TextChanged += (o, e) => UpdateItemFromForm();


        // Desc
        editPanel.Controls.Add(nLTB("  Desc:", out descTextBox, "Description of Item"));
        descTextBox.TextChanged += (o, e) => UpdateItemFromForm();

        // Common controls
        var layoutF = new FlowLayoutPanel() { AutoSize = true };
        // Sort
        layoutF.Controls.Add(nLTB("Sort Index:", out sortIndexTextBox, "The Sort Order", 120));
        sortIndexTextBox.TextChanged += (o, e) => UpdateItemFromForm();

        // Type
        layoutF.Controls.Add(new Label { Text = "Type:", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true });
        typeComboBox = new ComboBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, DropDownStyle = ComboBoxStyle.DropDownList, AutoSize = true };
        typeComboBox.Items.AddRange(Enum.GetNames(typeof(QuickItemType)));
        typeComboBox.SelectedIndexChanged += (o, e) =>
        {
            if (Enum.TryParse<QuickItemType>((o as Control).Text, out var type) && type == QuickItemType.Task)
            {
                taskNameTextBox.Show();
                prevSibling(taskNameTextBox).Show();

                workingDirTextBox.Hide();
                prevSibling(workingDirTextBox).Hide();

                argsTextBox.Hide();
                prevSibling(argsTextBox).Hide();

                pathTextBox.Hide();
                prevSibling(pathTextBox).Hide();


            }
            else
            {
                taskNameTextBox.Hide();
                prevSibling(taskNameTextBox).Hide();

                workingDirTextBox.Show();
                prevSibling(workingDirTextBox).Show();
                
                argsTextBox.Show();
                prevSibling(argsTextBox).Show();

                pathTextBox.Show();
                prevSibling(pathTextBox).Show();

            }
            UpdateItemFromForm();
        };
        layoutF.Controls.Add(typeComboBox);

        // Shortcut Key      
        layoutF.Controls.Add(nLTB("Shortcut:", out shortcutKeyTextBox, "The key combination to trigger item", 250));
        shortcutKeyTextBox.ReadOnly = true; shortcutKeyTextBox.Width = 220;
        shortcutKeyTextBox.Click += (s, e) =>
        {
            if (treeView.SelectedNode == null) return;
            var tag = (treeView.SelectedNode.Tag as TagT);
            if (tag != null)
                tag.Item1.ShortcutKey = PromptForShortcutKey();
        };

        // close/hide checkbox
        closeCheckBox = new CheckBox() { Text = ":Close on Run ", RightToLeft = RightToLeft.Yes, Margin = new Padding(10, 4, 0, 0) };
        layoutF.Controls.Add(closeCheckBox);
        closeCheckBox.Click += (s, e) => UpdateItemFromForm();


        editPanel.Controls.Add(layoutF);


        // Common font controls
        layoutF = new FlowLayoutPanel() { AutoSize = true };

        // Font Size
        layoutF.Controls.Add(new Label { Text = "Font Size:", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true });
        fontSizeUpDown = new NumericUpDown { Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true, Minimum = 6, Maximum = 72, DecimalPlaces = 1, Increment = 0.5m };
        fontSizeUpDown.ValueChanged += (s, e) => UpdateItemFromForm();
        layoutF.Controls.Add(fontSizeUpDown);

        // Font Color
        layoutF.Controls.Add(new Label { Text = "Font Color:", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true });
        fontColorButton = new Button { Width = 25, Height = 25, Anchor = AnchorStyles.Left | AnchorStyles.Right, Text = " ", BackColor = Color.Red };
        fontColorButton.Click += (s, e) => PromptForColor(fontColorButton, e);
        layoutF.Controls.Add(fontColorButton);
        CTT(fontColorButton, "Control Click for Default Color");

        // Button Color
        layoutF.Controls.Add(new Label { Text = "Button Color:", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true });
        buttonColorButton = new Button { Width = 25, Height = 25, Text = " ", BackColor = Color.Red };
        buttonColorButton.Click += (s, e) => PromptForColor(buttonColorButton, e);
        layoutF.Controls.Add(buttonColorButton);
        CTT(buttonColorButton, "Control Click for Default Color");

        // Font Family/Style Selection Button (Optional, advanced)
        layoutF.Controls.Add(new Label { Text = "Font Family/Style:", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true });
        selectFontButton = new Button { Width = 200, AutoSize = true, Anchor = AnchorStyles.Left | AnchorStyles.Right, Text = "Select Font" };
        selectFontButton.Click += (s, e) => PromptForFont(ref currentSelectedFont);
        layoutF.Controls.Add(selectFontButton);

        editPanel.Controls.Add(layoutF);

        // Path
        editPanel.Controls.Add(nLTB("App/Command:", out pathTextBox, "The path of the executable to run when item is activated.\n (Only if type is App or Command or Shift-Click Launch Button)"));
        pathTextBox.TextChanged += (o, e) => UpdateItemFromForm();

        // Args
        editPanel.Controls.Add(nLTB("Args:", out argsTextBox, "Arguments passed to App/Command"));
        argsTextBox.TextChanged += (o, e) => UpdateItemFromForm();


        // WorkingDir
        editPanel.Controls.Add(nLTB("Working Dir:", out workingDirTextBox, "The working director for the executable to run in."));
        workingDirTextBox.TextChanged += (o, e) => UpdateItemFromForm();

        layoutF = new FlowLayoutPanel() { AutoSize = true };

        // admin run checkbox
        runAsAdminCheckBox = new CheckBox() { Text = ":Admin ", RightToLeft = RightToLeft.Yes, Margin = new Padding(10, 4, 0, 0) };
        layoutF.Controls.Add(runAsAdminCheckBox);
        runAsAdminCheckBox.Click += (s, e) => UpdateItemFromForm();

        // useShellExecute
        useShellExecuteCheckBox = new CheckBox() { Text = ":Shell ", RightToLeft = RightToLeft.Yes, Margin = new Padding(10, 4, 0, 0) };
        layoutF.Controls.Add(useShellExecuteCheckBox);
        useShellExecuteCheckBox.Click += (s, e) => UpdateItemFromForm();

        // no window
        noWindowCheckBox = new CheckBox() { Text = ":No Window ", RightToLeft = RightToLeft.Yes, Margin = new Padding(10, 4, 0, 0) };
        layoutF.Controls.Add(noWindowCheckBox);
        noWindowCheckBox.Click += (s, e) => UpdateItemFromForm();


        editPanel.Controls.Add(layoutF);

        // Task
        editPanel.Controls.Add(nLTB("Task:", out taskNameTextBox, "A Windows task to run."));
        taskNameTextBox.TextChanged += (o, e) => UpdateItemFromForm();
        //taskNameTextBox.Click += 
        taskNameTextBox.Hide();
        prevSibling(taskNameTextBox).Hide();


        // Buttons - Use a FlowLayoutPanel for vertical stacking
        var buttonFlowPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            FlowDirection = FlowDirection.LeftToRight,
            Padding = new Padding(10),
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink
        };

        // Add, Delete button
        addButton = new Button { Text = "Add", Size = new System.Drawing.Size(70, 25), Margin = new Padding(0, 0, 15, 5) };
        addButton.Click += AddButton_Click;
        deleteButton = new Button { Text = "Delete", Size = new System.Drawing.Size(70, 25), Margin = new Padding(0, 0, 0, 5) };
        deleteButton.Click += DeleteButton_Click;

        buttonFlowPanel.Controls.AddRange(addButton, deleteButton);
        splitContainer.Panel1.Controls.Add(buttonFlowPanel);

        // Save Cancel
        var okCancelPanel = new FlowLayoutPanel { Height = 35, Dock = DockStyle.Bottom, FlowDirection = FlowDirection.RightToLeft, BackColor = SystemColors.Control, Width = 500 };
        saveButton = new Button { Text = "Save", Size = new System.Drawing.Size(70, 25), DialogResult = DialogResult.OK, Margin = new Padding(0, 0, 15, 5) };
        cancelButton = new Button { Text = "Cancel", Size = new System.Drawing.Size(70, 25), DialogResult = DialogResult.Cancel, Margin = new Padding(0, 0, 15, 5) };
        useTrayCheckBox = new CheckBox { Text = "Use Tray" };
        useTrayCheckBox.CheckedChanged += (s, e) => { var cb = (CheckBox)s; appSettings.UseTrayIcon = cb.Checked; };

        // global Shortcut Key              
        var gsl = new Label { Text = "Global Shortcut: ", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true };
        globalShortCutTextBox = new TextBox() { };
        globalShortCutTextBox.ReadOnly = true; globalShortCutTextBox.Width = 220;
        globalShortCutTextBox.Click += (s, e) => { appSettings.GlobalHotkey = PromptForShortcutKey(true); };
        okCancelPanel.Controls.AddRange(saveButton, cancelButton, useTrayCheckBox, globalShortCutTextBox, gsl);

        rightPanel.Controls.Add(editPanel);
        rightPanel.Controls.Add(okCancelPanel);

        splitContainer.Panel2.Controls.Add(rightPanel);
        splitContainer.Dock = DockStyle.Fill;

        this.Controls.Add(splitContainer);

        splitContainer.SplitterDistance = 200;

        this.KeyPreview = true;
        this.KeyDown += (s, e) => {
            // Exit app on escape. May not be desirable but at least for debugging
            if (e.KeyCode == Keys.Escape)
            {
                this.Hide();
                this.Close();
            }
        };


    }

    private bool disableItemUpdate = false;

    // Update QuickItem from Form
    private void UpdateItemFromForm()
    {
        if (disableItemUpdate) return;

        var node = treeView.SelectedNode;
        if (node == null) return;
        var tag = (node.Tag as TagT);
        if (tag == null) return;
        var item = tag.Item1;


        if (item != null)
        {
            // Update item properties from form controls
            item.Name = nameTextBox.Text;
            item.Desc = descTextBox.Text;
            if (Enum.TryParse<QuickItemType>(typeComboBox.SelectedItem?.ToString(), out var type))
            {
                item.Type = type;
            }
            item.Path = pathTextBox.Text;
            item.Args = argsTextBox.Text;
            item.WorkingDir = workingDirTextBox.Text;
            item.ShortcutKey = (Keys)(shortcutKeyTextBox.Tag ?? Keys.None);

            item.FontSize = (float)fontSizeUpDown.Value;
            item.FontColor = fontColorButton.BackColor;
            item.ButtonColor = buttonColorButton.BackColor;
            // Update font properties based on currentSelectedFont
            item.FontFamilyName = currentSelectedFont.FontFamily.Name;
            item.FontStyle = currentSelectedFont.Style;
            item.Close = closeCheckBox.Checked;
            item.RunAsAdmin = runAsAdminCheckBox.Checked;
            item.useShellExecute = useShellExecuteCheckBox.Checked;
            item.noWindow = noWindowCheckBox.Checked;
            item.TaskName = taskNameTextBox.Text;
            // The Font property setter will automatically update FontFamilyName, Size, Style

            // Update the tree node text if name changed
            if (node.Text != item.Name)
            {
                node.Text = item.Name;
            }
        }
    }

    // Update Form from TreeView change
    private void TreeView_UpdateFormAfterSelect(object sender, TreeViewEventArgs e)
    {
        disableItemUpdate = true;
        if (e.Node == null) return;
        var tag = e.Node?.Tag;
        if (tag == null) return;
        var item = (e.Node?.Tag as TagT).Item1;
        if (item != null)
        {

            nameTextBox.Text = item.Name;
            descTextBox.Text = item.Desc;
            sortIndexTextBox.Text = item.SortIndex.ToString();
            typeComboBox.SelectedItem = item.Type.ToString();
            pathTextBox.Text = item.Path;
            argsTextBox.Text = item.Args;
            workingDirTextBox.Text = item.WorkingDir;
            shortcutKeyTextBox.Text = sKey.GetKeyDisplay(item.ShortcutKey);
            shortcutKeyTextBox.Tag = item.ShortcutKey;
            selectFontButton.Text = $"{item.FontFamilyName}, {item.FontStyle}";
            fontColorButton.BackColor = item.FontColor;
            buttonColorButton.BackColor = item.ButtonColor;
            fontSizeUpDown.Value = (decimal)item.FontSize;
            closeCheckBox.Checked = item.Close;
            runAsAdminCheckBox.Checked = item.RunAsAdmin;
            useShellExecuteCheckBox.Checked = item.useShellExecute;
            noWindowCheckBox.Checked = item.noWindow;
            taskNameTextBox.Text = item.TaskName;
        }
        else
        {
            ClearForm();
        }
        disableItemUpdate = false;
    }


    private void PromptForColor(Button button, EventArgs e)
    {
        if (ModifierKeys == Keys.Control)
        {
            if (button == buttonColorButton)
                button.BackColor = QuickLauncher.GetDefaultItemColor((treeView.SelectedNode.Tag as TagT).Item1.Type);
            else if (button == fontColorButton)
                button.BackColor = Color.White;

            UpdateItemFromForm();
            return;
        }

        var colorDialog = new ColorDialog();
        if (colorDialog.ShowDialog() == DialogResult.OK)
        {
            var color = colorDialog.Color;
            button.BackColor = color;

            UpdateItemFromForm(); // Apply changes immediately
        }

    }

    private void PromptForFont(ref Font fontRef)
    {
        var fontDialog = new FontDialog { Font = fontRef };
        if (fontDialog.ShowDialog() == DialogResult.OK)
        {
            fontRef = fontDialog.Font;
            selectFontButton.Text = $"{fontRef.FontFamily.Name}, {fontRef.Style}"; // Update button text
            UpdateItemFromForm(); // Apply changes immediately
        }
    }



    private void BuildTreeView(QuickItemCollection collection = null, TreeNode parentNode = null, List<QuickItemCollection> navigationStack = null)
    {
        if (collection == null)
        {
            collection = rootCollection;
            if (collection == null)
                collection = new QuickItemCollection();

            treeView.Nodes.Clear(); // Start fresh for the root
        }


        foreach (var item in collection.Items)
        {
            var node = new TreeNode(item.Name);
            node.Tag = Tuple.Create(item, node);                        
            if (navigationStack != null && navigationStack.Count() > 0 && navigationStack[0].Name == item.Name)
            {
                node.Expand();
                if (navigationStack.Count() == 1)
                {
                    treeView.SelectedNode = node;
                }
                BuildTreeView(item, node, navigationStack[1..]);
            }
            else                
                BuildTreeView(item, node, null);
            
            

            if (parentNode != null)
            {
                parentNode.Nodes.Add(node);
            }
            else
            {
                treeView.Nodes.Add(node);
            }
        }
    }

    private void ClearForm()
    {
        nameTextBox.Clear();
        nameTextBox.Tag = null;
        typeComboBox.SelectedValue = QuickItemType.Folder;
        typeComboBox.Tag = null;
        pathTextBox.Clear();
        descTextBox.Clear();
        argsTextBox.Clear();
        workingDirTextBox.Clear();
        closeCheckBox.Checked = false;
        // TODO: Set other values to default
        shortcutKeyTextBox.Clear();
    }


    private void TreeView_AfterLabelEdit(object sender, NodeLabelEditEventArgs e)
    {
        if (e.Label != null && !string.IsNullOrWhiteSpace(e.Label))
        {
            var item = (e.Node.Tag as TagT).Item1;
            if (item != null)
            {
                item.Name = e.Label;
                nameTextBox.Text = e.Label;
            }
        }
        else
        {
            e.CancelEdit = true; // Revert if label is empty
        }
    }

    private Keys PromptForShortcutKey(bool global = false)
    {
        // Simple prompt for key input
        var promptForm = new Form()
        {
            Width = 300,
            Height = 150,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            Text = "Enter Shortcut Key",
            StartPosition = FormStartPosition.CenterParent,
            MaximizeBox = false,
            MinimizeBox = false
        };

        var label = new Label() { Left = 10, Top = 20, Text = "Press a key for the shortcut:" };
        var keyLabel = new Label() { Left = 10, Top = 45, Width = 250, Text = "Waiting for key..." };
        var okButton = new Button() { Text = "OK", Left = 10, Width = 75, Top = 70, DialogResult = DialogResult.OK };
        var cancelButton = new Button() { Text = "Cancel", Left = 90, Width = 75, Top = 70, DialogResult = DialogResult.Cancel };

        promptForm.Controls.AddRange(new Control[] { label, keyLabel, okButton, cancelButton });
        promptForm.AcceptButton = okButton;
        promptForm.CancelButton = cancelButton;

        KeyEventArgs? capturedKey = null;
        promptForm.KeyPreview = true;

        promptForm.KeyDown += (s, e) => {
            e.Handled = true; // Prevent closing on Enter/Escape            
            if (sKey.IsAssignableKey(e.KeyData))
            {
                capturedKey = e;
                KeysConverter kc = new KeysConverter();
                keyLabel.Text = $"Key captured: {(sKey.GetKeyDisplay(e.KeyData))}";
            }
        };

        if (promptForm.ShowDialog() == DialogResult.OK && capturedKey != null)
        {
            var e = capturedKey;
            if (global)
                globalShortCutTextBox.Text = $"{(sKey.GetKeyDisplay(e.KeyData))}";
            else
                shortcutKeyTextBox.Text = $"{(sKey.GetKeyDisplay(e.KeyData))}"; ;
            return e.KeyData;
        }

        if (global)
            globalShortCutTextBox.Text = "";
        else
            shortcutKeyTextBox.Text = "";
        return Keys.None;
    }

    protected override void OnClosed(EventArgs e)
    {
        base.OnClosed(e);
    }

    private void AddButton_Click(object sender, EventArgs e)
    {
        var selectedNode = treeView.SelectedNode;
        var newItem = new QuickItem { Name = "New Item", Type = QuickItemType.App, Path = "" };
        // Create the tuple tag first
        var newTag = Tuple.Create(newItem, (TreeNode)null); // TreeNode will be set after creating the node

        // Create the node and assign the tag (which contains the item and will reference the node itself)
        var newNode = new TreeNode(newItem.Name);
        // Update the tuple to include the node itself before assigning it as the tag
        var finalTag = Tuple.Create(newItem, newNode);
        newNode.Tag = finalTag;

        // Check if the selected node is a folder.
        if (selectedNode?.Tag is TagT selectedTag && selectedTag.Item1.Type == QuickItemType.Folder)
        {
            // Add as a child to the selected folder.
            var selectedFolderItem = selectedTag.Item1;
            if (selectedFolderItem.Items == null) selectedFolderItem.Items = new List<QuickItem>();
            selectedFolderItem.Items.Add(newItem);
            selectedNode.Nodes.Add(newNode);
            selectedNode.Expand();
        }
        else
        {
            // Otherwise, add to the root of the tree.
            if (rootCollection == null) rootCollection = new QuickItemCollection();
            rootCollection.Items.Add(newItem);
            treeView.Nodes.Add(newNode);
        }

        treeView.SelectedNode = newNode;
        newNode.BeginEdit(); // Allow immediate renaming
    }

    private void DeleteButton_Click(object sender, EventArgs e)
    {
        var node = treeView.SelectedNode;
        //if (node != null && MessageBox.Show("Delete this item and all its children?", "Confirm Delete", MessageBoxButtons.YesNo) == DialogResult.Yes)
        {
            // Use the new, correct helper method
            var parentItemsList = GetParentItemsList(node);
            var item = node.Tag as TagT;

            if (parentItemsList != null && item != null)
            {
                parentItemsList.Remove(item.Item1); // Update data model
                node.Remove();                // Update tree view
                if (parentItemsList.Count == 0)
                {
                    ClearForm();
                }
            }
        }
    }



    // Drag-and-Drop Methods

    private void TreeView_ItemDrag(object sender, ItemDragEventArgs e)
    {
        // Start the drag operation.
        dragNode = (TreeNode)e.Item;
        DoDragDrop(e.Item, DragDropEffects.Move);
    }

    private void TreeView_DragEnter(object sender, DragEventArgs e)
    {
        // Show the move cursor.
        e.Effect = e.AllowedEffect;
    }

    private void TreeView_DragOver(object sender, DragEventArgs e)
    {
        // Get the node that the mouse is currently hovering over.
        Point targetPoint = treeView.PointToClient(new Point(e.X, e.Y));
        TreeNode targetNode = treeView.GetNodeAt(targetPoint);


        // Allow drop if we have a valid target that isn't the dragged node or one of its children.
        if (dragNode != null && targetNode != null && targetNode != dragNode && !IsNodeDescendant(dragNode, targetNode))
        {
            // Auto-scroll the tree view if dragging near the top or bottom.
            int scrollThreshold = 20;
            if (targetPoint.Y < scrollThreshold)
                SendMessage(treeView.Handle, 277, (IntPtr)0, IntPtr.Zero); // WM_VSCROLL, SB_LINEUP
            else if (targetPoint.Y > treeView.ClientSize.Height - scrollThreshold)
                SendMessage(treeView.Handle, 277, (IntPtr)1, IntPtr.Zero); // WM_VSCROLL, SB_LINEDOWN
        }

        // Allow moving to anywhere in the treeView
        e.Effect = DragDropEffects.Move;
    }



    private void TreeView_DragDrop(object sender, DragEventArgs e)
    {
        if (dragNode == null) return;

        // --- 1. Get Destination and Preliminary Checks ---
        Point targetPoint = treeView.PointToClient(new Point(e.X, e.Y));
        TreeNode destinationNode = treeView.GetNodeAt(targetPoint);
        var nodeToMove = dragNode; // Keep a reference
        dragNode = null; // Reset for next operation

        // If dropped onto empty space, move to the root level.
        if (destinationNode == null)
        {
            MoveItemToRoot(nodeToMove);
            return;
        }

        // Prevent dropping a node onto itself or one of its own children.
        if (nodeToMove == destinationNode || IsNodeDescendant(nodeToMove, destinationNode))
        {
            return;
        }

        // --- 2. Get Data Model Items ---
        var itemToMove = (nodeToMove.Tag as TagT).Item1;
        var sourceItemsList = GetParentItemsList(nodeToMove);
        if (itemToMove == null || sourceItemsList == null) return;

        // --- 3. Determine Drop Action (Before, After, or Into) ---
        const int dropZoneSlack = 5; // Pixels from top/bottom to be considered "between"
        var destBounds = destinationNode.Bounds;
        var dropTargetItem = (destinationNode.Tag as TagT).Item1;

        // Update data and view in one go.
        sourceItemsList.Remove(itemToMove);
        nodeToMove.Remove();

        if (targetPoint.Y < destBounds.Top + dropZoneSlack)
        {
            // ACTION: Drop BEFORE destinationNode
            var destItemsList = GetParentItemsList(destinationNode);
            int destIndex = destinationNode.Parent == null ? treeView.Nodes.IndexOf(destinationNode) : destinationNode.Parent.Nodes.IndexOf(destinationNode);

            destItemsList.Insert(destIndex, itemToMove);
            if (destinationNode.Parent == null)
                treeView.Nodes.Insert(destIndex, nodeToMove);
            else
                destinationNode.Parent.Nodes.Insert(destIndex, nodeToMove);
        }
        else if (targetPoint.Y > destBounds.Bottom - dropZoneSlack)
        {
            // ACTION: Drop AFTER destinationNode
            var destItemsList = GetParentItemsList(destinationNode);
            int destIndex = destinationNode.Parent == null ? treeView.Nodes.IndexOf(destinationNode) : destinationNode.Parent.Nodes.IndexOf(destinationNode);

            destItemsList.Insert(destIndex + 1, itemToMove);
            if (destinationNode.Parent == null)
                treeView.Nodes.Insert(destIndex + 1, nodeToMove);
            else
                destinationNode.Parent.Nodes.Insert(destIndex + 1, nodeToMove);
        }
        //else if (dropTargetItem?.Type == QuickItemType.Folder)
        else
        {
            // ACTION: Drop INTO destinationNode (as a child)
            if (dropTargetItem.Items == null) dropTargetItem.Items = new List<QuickItem>();

            dropTargetItem.Items.Add(itemToMove);
            destinationNode.Nodes.Add(nodeToMove);
            destinationNode.Expand(); // Show the newly added child
        }


        treeView.SelectedNode = nodeToMove; // Reselect the moved node
    }


    /// <summary>
    /// Checks if a node is a descendant of another node, to prevent circular references.
    /// </summary>
    private bool IsNodeDescendant(TreeNode parent, TreeNode potentialDescendant)
    {
        TreeNode current = potentialDescendant.Parent;
        while (current != null)
        {
            if (current == parent) return true;
            current = current.Parent;
        }
        return false;
    }

    /// <summary>
    /// Moves a node to the root level of the tree.
    /// </summary>
    private void MoveItemToRoot(TreeNode nodeToMove, bool Bottom = false)
    {
        var itemToMove = (nodeToMove.Tag as TagT).Item1;
        var sourceItemsList = GetParentItemsList(nodeToMove);

        if (itemToMove != null && sourceItemsList != null)// && sourceItemsList != rootCollection.Items)
        {

            // Update Data Model
            sourceItemsList.Remove(itemToMove);
            rootCollection.Items.Add(itemToMove);

            // Update TreeView
            nodeToMove.Remove();
            treeView.Nodes.Add(nodeToMove);
            treeView.SelectedNode = nodeToMove;
        }
    }

    /// <summary>
    /// Gets the actual List<QuickItem> that contains the item associated with a given TreeNode.
    /// </summary>
    /// <summary>
    /// Gets the actual List<QuickItem> that contains the item associated with a given TreeNode.
    /// </summary>
    private List<QuickItem> GetParentItemsList(TreeNode node)
    {
        // Check if the node has a parent
        if (node?.Parent != null)
        {
            // Get the TagT from the parent node
            var parentTagT = node.Parent.Tag as TagT;
            if (parentTagT != null)
            {
                // Get the QuickItem from the TagT
                var parentItem = parentTagT.Item1; // Item1 is the QuickItem
                                                   // Check if the parent item is a folder
                if (parentItem.Type == QuickItemType.Folder)
                {
                    // Ensure the parent's Items list is not null
                    if (parentItem.Items == null) parentItem.Items = new List<QuickItem>();
                    return parentItem.Items; // Return the parent folder's Items list
                }
            }
        }
        // If node has no parent or parent is not a folder, return the root collection's Items list.
        return rootCollection.Items;
    }

    private TreeNode FindNodeByItem(TreeNodeCollection nodes, QuickItem item)
    {
        foreach (TreeNode node in nodes)
        {
            if (node.Tag == item) return node;
            var found = FindNodeByItem(node.Nodes, item);
            if (found != null) return found;
        }
        return null;
    }

    private TreeNode FindNodeByText(TreeView treeView, string text)
    {
        foreach (TreeNode node in treeView.Nodes)
        {
            if (node.Text == text) return node;
            var found = FindNodeByTextRecursive(node, text);
            if (found != null) return found;
        }
        return null;
    }

    private TreeNode FindNodeByTextRecursive(TreeNode node, string text)
    {
        foreach (TreeNode child in node.Nodes)
        {
            if (child.Text == text) return child;
            var found = FindNodeByTextRecursive(child, text);
            if (found != null) return found;
        }
        return null;
    }


    private IEnumerable<QuickItemCollection> GetAllCollections(QuickItemCollection collection)
    {
        yield return collection;
        foreach (var item in collection.Items.Where(i => i.Type == QuickItemType.Folder))
        {
            foreach (var nested in GetAllCollections(item))
            {
                yield return nested;
            }
        }
    }

}