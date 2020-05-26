<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class mainForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(mainForm))
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.NewToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OpenToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FileToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.FolderToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.toolStripSeparator = New System.Windows.Forms.ToolStripSeparator()
        Me.SaveToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SaveAsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.toolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.PrintToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PrintPreviewToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.toolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.EditToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.UndoToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RedoToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.toolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.CutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CopyToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PasteToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.toolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator()
        Me.SelectAllToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CustomizeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OptionsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ContentsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.IndexToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SearchToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.toolStripSeparator5 = New System.Windows.Forms.ToolStripSeparator()
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuItem_department = New System.Windows.Forms.ToolStripMenuItem()
        Me.item_lighting = New System.Windows.Forms.ToolStripMenuItem()
        Me.item_movHeads = New System.Windows.Forms.ToolStripMenuItem()
        Me.item_strobes = New System.Windows.Forms.ToolStripMenuItem()
        Me.item_blinders = New System.Windows.Forms.ToolStripMenuItem()
        Me.item_arch = New System.Windows.Forms.ToolStripMenuItem()
        Me.item_LED = New System.Windows.Forms.ToolStripMenuItem()
        Me.item_smoke = New System.Windows.Forms.ToolStripMenuItem()
        Me.item_consoles = New System.Windows.Forms.ToolStripMenuItem()
        Me.item_intercom = New System.Windows.Forms.ToolStripMenuItem()
        Me.item_screen = New System.Windows.Forms.ToolStripMenuItem()
        Me.item_modules = New System.Windows.Forms.ToolStripMenuItem()
        Me.item_servers = New System.Windows.Forms.ToolStripMenuItem()
        Me.item_controllers1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.item_controllers2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.item_projectors = New System.Windows.Forms.ToolStripMenuItem()
        Me.item_scr_construction = New System.Windows.Forms.ToolStripMenuItem()
        Me.item_lightDesks = New System.Windows.Forms.ToolStripMenuItem()
        Me.item_cameras = New System.Windows.Forms.ToolStripMenuItem()
        Me.item_commutation = New System.Windows.Forms.ToolStripMenuItem()
        Me.item_truss = New System.Windows.Forms.ToolStripMenuItem()
        Me.item_construction = New System.Windows.Forms.ToolStripMenuItem()
        Me.item_sound = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuItem_company = New System.Windows.Forms.ToolStripMenuItem()
        Me.item_belimlight = New System.Windows.Forms.ToolStripMenuItem()
        Me.item_PRLighting = New System.Windows.Forms.ToolStripMenuItem()
        Me.item_blackout = New System.Windows.Forms.ToolStripMenuItem()
        Me.item_vision = New System.Windows.Forms.ToolStripMenuItem()
        Me.item_stage = New System.Windows.Forms.ToolStripMenuItem()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.lbl_companyValue = New System.Windows.Forms.Label()
        Me.lbl_company = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.lbl_subsectionValue = New System.Windows.Forms.Label()
        Me.lbl_subsection = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.lbl_dpartmentValue = New System.Windows.Forms.Label()
        Me.lbl_department = New System.Windows.Forms.Label()
        Me.dgv = New System.Windows.Forms.DataGridView()
        Me.FBD = New System.Windows.Forms.FolderBrowserDialog()
        Me.btn_test = New System.Windows.Forms.Button()
        Me.MenuStrip1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.dgv, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.EditToolStripMenuItem, Me.ToolsToolStripMenuItem, Me.HelpToolStripMenuItem, Me.menuItem_department, Me.menuItem_company})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(1068, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.NewToolStripMenuItem, Me.OpenToolStripMenuItem, Me.toolStripSeparator, Me.SaveToolStripMenuItem, Me.SaveAsToolStripMenuItem, Me.toolStripSeparator1, Me.PrintToolStripMenuItem, Me.PrintPreviewToolStripMenuItem, Me.toolStripSeparator2, Me.ExitToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.FileToolStripMenuItem.Text = "&File"
        '
        'NewToolStripMenuItem
        '
        Me.NewToolStripMenuItem.Image = CType(resources.GetObject("NewToolStripMenuItem.Image"), System.Drawing.Image)
        Me.NewToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.NewToolStripMenuItem.Name = "NewToolStripMenuItem"
        Me.NewToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.N), System.Windows.Forms.Keys)
        Me.NewToolStripMenuItem.Size = New System.Drawing.Size(146, 22)
        Me.NewToolStripMenuItem.Text = "&New"
        '
        'OpenToolStripMenuItem
        '
        Me.OpenToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem1, Me.FolderToolStripMenuItem})
        Me.OpenToolStripMenuItem.Image = CType(resources.GetObject("OpenToolStripMenuItem.Image"), System.Drawing.Image)
        Me.OpenToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.OpenToolStripMenuItem.Name = "OpenToolStripMenuItem"
        Me.OpenToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.O), System.Windows.Forms.Keys)
        Me.OpenToolStripMenuItem.Size = New System.Drawing.Size(146, 22)
        Me.OpenToolStripMenuItem.Text = "&Open"
        '
        'FileToolStripMenuItem1
        '
        Me.FileToolStripMenuItem1.Name = "FileToolStripMenuItem1"
        Me.FileToolStripMenuItem1.Size = New System.Drawing.Size(107, 22)
        Me.FileToolStripMenuItem1.Text = "&File"
        '
        'FolderToolStripMenuItem
        '
        Me.FolderToolStripMenuItem.Name = "FolderToolStripMenuItem"
        Me.FolderToolStripMenuItem.Size = New System.Drawing.Size(107, 22)
        Me.FolderToolStripMenuItem.Text = "Folder"
        '
        'toolStripSeparator
        '
        Me.toolStripSeparator.Name = "toolStripSeparator"
        Me.toolStripSeparator.Size = New System.Drawing.Size(143, 6)
        '
        'SaveToolStripMenuItem
        '
        Me.SaveToolStripMenuItem.Image = CType(resources.GetObject("SaveToolStripMenuItem.Image"), System.Drawing.Image)
        Me.SaveToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.SaveToolStripMenuItem.Name = "SaveToolStripMenuItem"
        Me.SaveToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.S), System.Windows.Forms.Keys)
        Me.SaveToolStripMenuItem.Size = New System.Drawing.Size(146, 22)
        Me.SaveToolStripMenuItem.Text = "&Save"
        '
        'SaveAsToolStripMenuItem
        '
        Me.SaveAsToolStripMenuItem.Name = "SaveAsToolStripMenuItem"
        Me.SaveAsToolStripMenuItem.Size = New System.Drawing.Size(146, 22)
        Me.SaveAsToolStripMenuItem.Text = "Save &As"
        '
        'toolStripSeparator1
        '
        Me.toolStripSeparator1.Name = "toolStripSeparator1"
        Me.toolStripSeparator1.Size = New System.Drawing.Size(143, 6)
        '
        'PrintToolStripMenuItem
        '
        Me.PrintToolStripMenuItem.Image = CType(resources.GetObject("PrintToolStripMenuItem.Image"), System.Drawing.Image)
        Me.PrintToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.PrintToolStripMenuItem.Name = "PrintToolStripMenuItem"
        Me.PrintToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.P), System.Windows.Forms.Keys)
        Me.PrintToolStripMenuItem.Size = New System.Drawing.Size(146, 22)
        Me.PrintToolStripMenuItem.Text = "&Print"
        '
        'PrintPreviewToolStripMenuItem
        '
        Me.PrintPreviewToolStripMenuItem.Image = CType(resources.GetObject("PrintPreviewToolStripMenuItem.Image"), System.Drawing.Image)
        Me.PrintPreviewToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.PrintPreviewToolStripMenuItem.Name = "PrintPreviewToolStripMenuItem"
        Me.PrintPreviewToolStripMenuItem.Size = New System.Drawing.Size(146, 22)
        Me.PrintPreviewToolStripMenuItem.Text = "Print Pre&view"
        '
        'toolStripSeparator2
        '
        Me.toolStripSeparator2.Name = "toolStripSeparator2"
        Me.toolStripSeparator2.Size = New System.Drawing.Size(143, 6)
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(146, 22)
        Me.ExitToolStripMenuItem.Text = "E&xit"
        '
        'EditToolStripMenuItem
        '
        Me.EditToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.UndoToolStripMenuItem, Me.RedoToolStripMenuItem, Me.toolStripSeparator3, Me.CutToolStripMenuItem, Me.CopyToolStripMenuItem, Me.PasteToolStripMenuItem, Me.toolStripSeparator4, Me.SelectAllToolStripMenuItem})
        Me.EditToolStripMenuItem.Name = "EditToolStripMenuItem"
        Me.EditToolStripMenuItem.Size = New System.Drawing.Size(39, 20)
        Me.EditToolStripMenuItem.Text = "&Edit"
        '
        'UndoToolStripMenuItem
        '
        Me.UndoToolStripMenuItem.Name = "UndoToolStripMenuItem"
        Me.UndoToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Z), System.Windows.Forms.Keys)
        Me.UndoToolStripMenuItem.Size = New System.Drawing.Size(144, 22)
        Me.UndoToolStripMenuItem.Text = "&Undo"
        '
        'RedoToolStripMenuItem
        '
        Me.RedoToolStripMenuItem.Name = "RedoToolStripMenuItem"
        Me.RedoToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Y), System.Windows.Forms.Keys)
        Me.RedoToolStripMenuItem.Size = New System.Drawing.Size(144, 22)
        Me.RedoToolStripMenuItem.Text = "&Redo"
        '
        'toolStripSeparator3
        '
        Me.toolStripSeparator3.Name = "toolStripSeparator3"
        Me.toolStripSeparator3.Size = New System.Drawing.Size(141, 6)
        '
        'CutToolStripMenuItem
        '
        Me.CutToolStripMenuItem.Image = CType(resources.GetObject("CutToolStripMenuItem.Image"), System.Drawing.Image)
        Me.CutToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.CutToolStripMenuItem.Name = "CutToolStripMenuItem"
        Me.CutToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.X), System.Windows.Forms.Keys)
        Me.CutToolStripMenuItem.Size = New System.Drawing.Size(144, 22)
        Me.CutToolStripMenuItem.Text = "Cu&t"
        '
        'CopyToolStripMenuItem
        '
        Me.CopyToolStripMenuItem.Image = CType(resources.GetObject("CopyToolStripMenuItem.Image"), System.Drawing.Image)
        Me.CopyToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.CopyToolStripMenuItem.Name = "CopyToolStripMenuItem"
        Me.CopyToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.C), System.Windows.Forms.Keys)
        Me.CopyToolStripMenuItem.Size = New System.Drawing.Size(144, 22)
        Me.CopyToolStripMenuItem.Text = "&Copy"
        '
        'PasteToolStripMenuItem
        '
        Me.PasteToolStripMenuItem.Image = CType(resources.GetObject("PasteToolStripMenuItem.Image"), System.Drawing.Image)
        Me.PasteToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.PasteToolStripMenuItem.Name = "PasteToolStripMenuItem"
        Me.PasteToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.V), System.Windows.Forms.Keys)
        Me.PasteToolStripMenuItem.Size = New System.Drawing.Size(144, 22)
        Me.PasteToolStripMenuItem.Text = "&Paste"
        '
        'toolStripSeparator4
        '
        Me.toolStripSeparator4.Name = "toolStripSeparator4"
        Me.toolStripSeparator4.Size = New System.Drawing.Size(141, 6)
        '
        'SelectAllToolStripMenuItem
        '
        Me.SelectAllToolStripMenuItem.Name = "SelectAllToolStripMenuItem"
        Me.SelectAllToolStripMenuItem.Size = New System.Drawing.Size(144, 22)
        Me.SelectAllToolStripMenuItem.Text = "Select &All"
        '
        'ToolsToolStripMenuItem
        '
        Me.ToolsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CustomizeToolStripMenuItem, Me.OptionsToolStripMenuItem})
        Me.ToolsToolStripMenuItem.Name = "ToolsToolStripMenuItem"
        Me.ToolsToolStripMenuItem.Size = New System.Drawing.Size(46, 20)
        Me.ToolsToolStripMenuItem.Text = "&Tools"
        '
        'CustomizeToolStripMenuItem
        '
        Me.CustomizeToolStripMenuItem.Name = "CustomizeToolStripMenuItem"
        Me.CustomizeToolStripMenuItem.Size = New System.Drawing.Size(130, 22)
        Me.CustomizeToolStripMenuItem.Text = "&Customize"
        '
        'OptionsToolStripMenuItem
        '
        Me.OptionsToolStripMenuItem.Name = "OptionsToolStripMenuItem"
        Me.OptionsToolStripMenuItem.Size = New System.Drawing.Size(130, 22)
        Me.OptionsToolStripMenuItem.Text = "&Options"
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ContentsToolStripMenuItem, Me.IndexToolStripMenuItem, Me.SearchToolStripMenuItem, Me.toolStripSeparator5, Me.AboutToolStripMenuItem})
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
        Me.HelpToolStripMenuItem.Text = "&Help"
        '
        'ContentsToolStripMenuItem
        '
        Me.ContentsToolStripMenuItem.Name = "ContentsToolStripMenuItem"
        Me.ContentsToolStripMenuItem.Size = New System.Drawing.Size(122, 22)
        Me.ContentsToolStripMenuItem.Text = "&Contents"
        '
        'IndexToolStripMenuItem
        '
        Me.IndexToolStripMenuItem.Name = "IndexToolStripMenuItem"
        Me.IndexToolStripMenuItem.Size = New System.Drawing.Size(122, 22)
        Me.IndexToolStripMenuItem.Text = "&Index"
        '
        'SearchToolStripMenuItem
        '
        Me.SearchToolStripMenuItem.Name = "SearchToolStripMenuItem"
        Me.SearchToolStripMenuItem.Size = New System.Drawing.Size(122, 22)
        Me.SearchToolStripMenuItem.Text = "&Search"
        '
        'toolStripSeparator5
        '
        Me.toolStripSeparator5.Name = "toolStripSeparator5"
        Me.toolStripSeparator5.Size = New System.Drawing.Size(119, 6)
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(122, 22)
        Me.AboutToolStripMenuItem.Text = "&About..."
        '
        'menuItem_department
        '
        Me.menuItem_department.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.item_lighting, Me.item_screen, Me.item_commutation, Me.item_truss, Me.item_construction, Me.item_sound})
        Me.menuItem_department.Name = "menuItem_department"
        Me.menuItem_department.Size = New System.Drawing.Size(82, 20)
        Me.menuItem_department.Text = "Department"
        '
        'item_lighting
        '
        Me.item_lighting.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.item_movHeads, Me.item_strobes, Me.item_blinders, Me.item_arch, Me.item_LED, Me.item_smoke, Me.item_consoles, Me.item_intercom})
        Me.item_lighting.Name = "item_lighting"
        Me.item_lighting.Size = New System.Drawing.Size(180, 22)
        Me.item_lighting.Text = "Lighting"
        '
        'item_movHeads
        '
        Me.item_movHeads.Name = "item_movHeads"
        Me.item_movHeads.Size = New System.Drawing.Size(151, 22)
        Me.item_movHeads.Text = "&Moving Heads"
        '
        'item_strobes
        '
        Me.item_strobes.Name = "item_strobes"
        Me.item_strobes.Size = New System.Drawing.Size(151, 22)
        Me.item_strobes.Text = "&Strobes"
        '
        'item_blinders
        '
        Me.item_blinders.Name = "item_blinders"
        Me.item_blinders.Size = New System.Drawing.Size(151, 22)
        Me.item_blinders.Text = "&Blinders"
        '
        'item_arch
        '
        Me.item_arch.Name = "item_arch"
        Me.item_arch.Size = New System.Drawing.Size(151, 22)
        Me.item_arch.Text = "&Arch"
        '
        'item_LED
        '
        Me.item_LED.Name = "item_LED"
        Me.item_LED.Size = New System.Drawing.Size(151, 22)
        Me.item_LED.Text = "&LED"
        '
        'item_smoke
        '
        Me.item_smoke.Name = "item_smoke"
        Me.item_smoke.Size = New System.Drawing.Size(151, 22)
        Me.item_smoke.Text = "&Smoke"
        '
        'item_consoles
        '
        Me.item_consoles.Name = "item_consoles"
        Me.item_consoles.Size = New System.Drawing.Size(151, 22)
        Me.item_consoles.Text = "&Lighting desks"
        '
        'item_intercom
        '
        Me.item_intercom.Name = "item_intercom"
        Me.item_intercom.Size = New System.Drawing.Size(151, 22)
        Me.item_intercom.Text = "&Intercoms"
        '
        'item_screen
        '
        Me.item_screen.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.item_modules, Me.item_servers, Me.item_controllers1, Me.item_controllers2, Me.item_projectors, Me.item_scr_construction, Me.item_lightDesks, Me.item_cameras})
        Me.item_screen.Name = "item_screen"
        Me.item_screen.Size = New System.Drawing.Size(180, 22)
        Me.item_screen.Text = "Screen"
        '
        'item_modules
        '
        Me.item_modules.Name = "item_modules"
        Me.item_modules.Size = New System.Drawing.Size(203, 22)
        Me.item_modules.Text = "&Screen modules"
        '
        'item_servers
        '
        Me.item_servers.Name = "item_servers"
        Me.item_servers.Size = New System.Drawing.Size(203, 22)
        Me.item_servers.Text = "&Media servers"
        '
        'item_controllers1
        '
        Me.item_controllers1.Name = "item_controllers1"
        Me.item_controllers1.Size = New System.Drawing.Size(203, 22)
        Me.item_controllers1.Text = "&Controllers, convertors 1"
        '
        'item_controllers2
        '
        Me.item_controllers2.Name = "item_controllers2"
        Me.item_controllers2.Size = New System.Drawing.Size(203, 22)
        Me.item_controllers2.Text = "&Controllers, convertors 2"
        '
        'item_projectors
        '
        Me.item_projectors.Name = "item_projectors"
        Me.item_projectors.Size = New System.Drawing.Size(203, 22)
        Me.item_projectors.Text = "&Projectors and screens"
        '
        'item_scr_construction
        '
        Me.item_scr_construction.Name = "item_scr_construction"
        Me.item_scr_construction.Size = New System.Drawing.Size(203, 22)
        Me.item_scr_construction.Text = "&Screen construction"
        '
        'item_lightDesks
        '
        Me.item_lightDesks.Name = "item_lightDesks"
        Me.item_lightDesks.Size = New System.Drawing.Size(203, 22)
        Me.item_lightDesks.Text = "&Lighting desks"
        '
        'item_cameras
        '
        Me.item_cameras.Name = "item_cameras"
        Me.item_cameras.Size = New System.Drawing.Size(203, 22)
        Me.item_cameras.Text = "&Cameras, mixing desks"
        '
        'item_commutation
        '
        Me.item_commutation.Name = "item_commutation"
        Me.item_commutation.Size = New System.Drawing.Size(180, 22)
        Me.item_commutation.Text = "Commutation"
        '
        'item_truss
        '
        Me.item_truss.Name = "item_truss"
        Me.item_truss.Size = New System.Drawing.Size(180, 22)
        Me.item_truss.Text = "Trusses and motors"
        '
        'item_construction
        '
        Me.item_construction.Name = "item_construction"
        Me.item_construction.Size = New System.Drawing.Size(180, 22)
        Me.item_construction.Text = "Construction"
        '
        'item_sound
        '
        Me.item_sound.Name = "item_sound"
        Me.item_sound.Size = New System.Drawing.Size(180, 22)
        Me.item_sound.Text = "Sound"
        '
        'menuItem_company
        '
        Me.menuItem_company.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.item_belimlight, Me.item_PRLighting, Me.item_blackout, Me.item_vision, Me.item_stage})
        Me.menuItem_company.Name = "menuItem_company"
        Me.menuItem_company.Size = New System.Drawing.Size(71, 20)
        Me.menuItem_company.Text = "Company"
        '
        'item_belimlight
        '
        Me.item_belimlight.Name = "item_belimlight"
        Me.item_belimlight.Size = New System.Drawing.Size(169, 22)
        Me.item_belimlight.Text = "&Belimlight"
        '
        'item_PRLighting
        '
        Me.item_PRLighting.Name = "item_PRLighting"
        Me.item_PRLighting.Size = New System.Drawing.Size(169, 22)
        Me.item_PRLighting.Text = "&PRLighting"
        '
        'item_blackout
        '
        Me.item_blackout.Name = "item_blackout"
        Me.item_blackout.Size = New System.Drawing.Size(169, 22)
        Me.item_blackout.Text = "&Blackout"
        '
        'item_vision
        '
        Me.item_vision.Name = "item_vision"
        Me.item_vision.Size = New System.Drawing.Size(169, 22)
        Me.item_vision.Text = "&Multivision"
        '
        'item_stage
        '
        Me.item_stage.Name = "item_stage"
        Me.item_stage.Size = New System.Drawing.Size(169, 22)
        Me.item_stage.Text = "&Stage engineering"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Silver
        Me.Panel1.Controls.Add(Me.GroupBox3)
        Me.Panel1.Controls.Add(Me.GroupBox2)
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Location = New System.Drawing.Point(0, 38)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1068, 46)
        Me.Panel1.TabIndex = 1
        '
        'GroupBox3
        '
        Me.GroupBox3.BackColor = System.Drawing.Color.WhiteSmoke
        Me.GroupBox3.Controls.Add(Me.lbl_companyValue)
        Me.GroupBox3.Controls.Add(Me.lbl_company)
        Me.GroupBox3.Location = New System.Drawing.Point(842, 5)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(223, 33)
        Me.GroupBox3.TabIndex = 3
        Me.GroupBox3.TabStop = False
        '
        'lbl_companyValue
        '
        Me.lbl_companyValue.AutoSize = True
        Me.lbl_companyValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.lbl_companyValue.Location = New System.Drawing.Point(83, 12)
        Me.lbl_companyValue.Name = "lbl_companyValue"
        Me.lbl_companyValue.Size = New System.Drawing.Size(76, 16)
        Me.lbl_companyValue.TabIndex = 0
        Me.lbl_companyValue.Text = "Belimlight"
        '
        'lbl_company
        '
        Me.lbl_company.AutoSize = True
        Me.lbl_company.Location = New System.Drawing.Point(15, 12)
        Me.lbl_company.Name = "lbl_company"
        Me.lbl_company.Size = New System.Drawing.Size(54, 13)
        Me.lbl_company.TabIndex = 0
        Me.lbl_company.Text = "Company:"
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.WhiteSmoke
        Me.GroupBox2.Controls.Add(Me.lbl_subsectionValue)
        Me.GroupBox2.Controls.Add(Me.lbl_subsection)
        Me.GroupBox2.Location = New System.Drawing.Point(391, 5)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(261, 33)
        Me.GroupBox2.TabIndex = 2
        Me.GroupBox2.TabStop = False
        '
        'lbl_subsectionValue
        '
        Me.lbl_subsectionValue.AutoSize = True
        Me.lbl_subsectionValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.lbl_subsectionValue.Location = New System.Drawing.Point(73, 12)
        Me.lbl_subsectionValue.Name = "lbl_subsectionValue"
        Me.lbl_subsectionValue.Size = New System.Drawing.Size(108, 16)
        Me.lbl_subsectionValue.TabIndex = 0
        Me.lbl_subsectionValue.Text = "Moving Heads"
        '
        'lbl_subsection
        '
        Me.lbl_subsection.AutoSize = True
        Me.lbl_subsection.Location = New System.Drawing.Point(5, 12)
        Me.lbl_subsection.Name = "lbl_subsection"
        Me.lbl_subsection.Size = New System.Drawing.Size(63, 13)
        Me.lbl_subsection.TabIndex = 0
        Me.lbl_subsection.Text = "Subsection:"
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.WhiteSmoke
        Me.GroupBox1.Controls.Add(Me.lbl_dpartmentValue)
        Me.GroupBox1.Controls.Add(Me.lbl_department)
        Me.GroupBox1.Location = New System.Drawing.Point(1, 5)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(182, 33)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        '
        'lbl_dpartmentValue
        '
        Me.lbl_dpartmentValue.AutoSize = True
        Me.lbl_dpartmentValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.lbl_dpartmentValue.Location = New System.Drawing.Point(81, 12)
        Me.lbl_dpartmentValue.Name = "lbl_dpartmentValue"
        Me.lbl_dpartmentValue.Size = New System.Drawing.Size(62, 16)
        Me.lbl_dpartmentValue.TabIndex = 0
        Me.lbl_dpartmentValue.Text = "Lighting"
        '
        'lbl_department
        '
        Me.lbl_department.AutoSize = True
        Me.lbl_department.Location = New System.Drawing.Point(13, 12)
        Me.lbl_department.Name = "lbl_department"
        Me.lbl_department.Size = New System.Drawing.Size(65, 13)
        Me.lbl_department.TabIndex = 0
        Me.lbl_department.Text = "Department:"
        '
        'dgv
        '
        Me.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgv.Location = New System.Drawing.Point(6, 104)
        Me.dgv.Name = "dgv"
        Me.dgv.Size = New System.Drawing.Size(1056, 543)
        Me.dgv.TabIndex = 2
        '
        'btn_test
        '
        Me.btn_test.Location = New System.Drawing.Point(72, 680)
        Me.btn_test.Name = "btn_test"
        Me.btn_test.Size = New System.Drawing.Size(75, 23)
        Me.btn_test.TabIndex = 3
        Me.btn_test.Text = "Test"
        Me.btn_test.UseVisualStyleBackColor = True
        '
        'mainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1068, 736)
        Me.Controls.Add(Me.btn_test)
        Me.Controls.Add(Me.dgv)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "mainForm"
        Me.Text = "datasetForm"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.dgv, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents FileToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents NewToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents OpenToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents toolStripSeparator As ToolStripSeparator
    Friend WithEvents SaveToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SaveAsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents toolStripSeparator1 As ToolStripSeparator
    Friend WithEvents PrintToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents PrintPreviewToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents toolStripSeparator2 As ToolStripSeparator
    Friend WithEvents ExitToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents EditToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents UndoToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents RedoToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents toolStripSeparator3 As ToolStripSeparator
    Friend WithEvents CutToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents CopyToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents PasteToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents toolStripSeparator4 As ToolStripSeparator
    Friend WithEvents SelectAllToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents CustomizeToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents OptionsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents HelpToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ContentsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents IndexToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SearchToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents toolStripSeparator5 As ToolStripSeparator
    Friend WithEvents AboutToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents menuItem_department As ToolStripMenuItem
    Friend WithEvents item_lighting As ToolStripMenuItem
    Friend WithEvents item_screen As ToolStripMenuItem
    Friend WithEvents item_commutation As ToolStripMenuItem
    Friend WithEvents item_truss As ToolStripMenuItem
    Friend WithEvents item_construction As ToolStripMenuItem
    Friend WithEvents item_sound As ToolStripMenuItem
    Friend WithEvents item_modules As ToolStripMenuItem
    Friend WithEvents item_servers As ToolStripMenuItem
    Friend WithEvents item_controllers1 As ToolStripMenuItem
    Friend WithEvents item_controllers2 As ToolStripMenuItem
    Friend WithEvents item_projectors As ToolStripMenuItem
    Friend WithEvents item_scr_construction As ToolStripMenuItem
    Friend WithEvents item_lightDesks As ToolStripMenuItem
    Friend WithEvents item_cameras As ToolStripMenuItem
    Friend WithEvents item_movHeads As ToolStripMenuItem
    Friend WithEvents item_strobes As ToolStripMenuItem
    Friend WithEvents item_blinders As ToolStripMenuItem
    Friend WithEvents item_arch As ToolStripMenuItem
    Friend WithEvents item_LED As ToolStripMenuItem
    Friend WithEvents item_smoke As ToolStripMenuItem
    Friend WithEvents item_consoles As ToolStripMenuItem
    Friend WithEvents item_intercom As ToolStripMenuItem
    Friend WithEvents menuItem_company As ToolStripMenuItem
    Friend WithEvents item_belimlight As ToolStripMenuItem
    Friend WithEvents item_PRLighting As ToolStripMenuItem
    Friend WithEvents item_blackout As ToolStripMenuItem
    Friend WithEvents item_vision As ToolStripMenuItem
    Friend WithEvents item_stage As ToolStripMenuItem
    Friend WithEvents Panel1 As Panel
    Friend WithEvents lbl_dpartmentValue As Label
    Friend WithEvents lbl_department As Label
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents lbl_companyValue As Label
    Friend WithEvents lbl_company As Label
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents lbl_subsectionValue As Label
    Friend WithEvents lbl_subsection As Label
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents dgv As DataGridView
    Friend WithEvents FileToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents FolderToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents FBD As FolderBrowserDialog
    Friend WithEvents btn_test As Button
End Class
