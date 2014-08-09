Imports Microsoft.Win32
Imports System.IO

Public Class Main_Screen
    Inherits System.Windows.Forms.Form

    Dim WithEvents Worker1 As Worker
    Public Delegate Sub WorkerhHandler(ByVal Result As String)
    Public Delegate Sub WorkerProgresshHandler()
    Public Delegate Sub WorkerFileCreatedhHandler()
    Public Delegate Sub WorkerErrorEncountered(ByVal ex As Exception)

    Private splash_loader As Splash_Screen
    Public dataloaded As Boolean = False

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Worker1 = New Worker
        AddHandler Worker1.WorkerComplete, AddressOf WorkerHandler
        AddHandler Worker1.WorkerProgress, AddressOf WorkerProgressHandler
        AddHandler Worker1.WorkerFileCreated, AddressOf WorkerFileCreatedHandler
        AddHandler Worker1.WorkerErrorEncountered, AddressOf WorkerErrorEncounteredHandler
    End Sub

    Public Sub New(ByVal splash As Splash_Screen)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        splash_loader = splash
        Worker1 = New Worker
        AddHandler Worker1.WorkerComplete, AddressOf WorkerHandler
        AddHandler Worker1.WorkerProgress, AddressOf WorkerProgressHandler
        AddHandler Worker1.WorkerFileCreated, AddressOf WorkerFileCreatedHandler
        AddHandler Worker1.WorkerErrorEncountered, AddressOf WorkerErrorEncounteredHandler
    End Sub
    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents Timer2 As System.Windows.Forms.Timer
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents NotifyIcon1 As System.Windows.Forms.NotifyIcon
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents PictureBox5 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox4 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents folderscreated As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents filescreated As System.Windows.Forms.Label
    Friend WithEvents term1 As System.Windows.Forms.TextBox
    Friend WithEvents term2 As System.Windows.Forms.TextBox
    Friend WithEvents term3 As System.Windows.Forms.TextBox
    Friend WithEvents term4 As System.Windows.Forms.TextBox
    Friend WithEvents term5 As System.Windows.Forms.TextBox
    Friend WithEvents extension5 As System.Windows.Forms.TextBox
    Friend WithEvents extension4 As System.Windows.Forms.TextBox
    Friend WithEvents extension3 As System.Windows.Forms.TextBox
    Friend WithEvents extension2 As System.Windows.Forms.TextBox
    Friend WithEvents extension1 As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents basefolder As System.Windows.Forms.TextBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents foldername As System.Windows.Forms.TextBox
    Friend WithEvents depth As System.Windows.Forms.NumericUpDown
    Friend WithEvents amount As System.Windows.Forms.NumericUpDown
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Main_Screen))
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
        Me.Label8 = New System.Windows.Forms.Label
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog
        Me.Label9 = New System.Windows.Forms.Label
        Me.Button3 = New System.Windows.Forms.Button
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.PictureBox5 = New System.Windows.Forms.PictureBox
        Me.PictureBox4 = New System.Windows.Forms.PictureBox
        Me.PictureBox3 = New System.Windows.Forms.PictureBox
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.term1 = New System.Windows.Forms.TextBox
        Me.folderscreated = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.filescreated = New System.Windows.Forms.Label
        Me.term2 = New System.Windows.Forms.TextBox
        Me.term3 = New System.Windows.Forms.TextBox
        Me.term4 = New System.Windows.Forms.TextBox
        Me.term5 = New System.Windows.Forms.TextBox
        Me.extension5 = New System.Windows.Forms.TextBox
        Me.extension4 = New System.Windows.Forms.TextBox
        Me.extension3 = New System.Windows.Forms.TextBox
        Me.extension2 = New System.Windows.Forms.TextBox
        Me.extension1 = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.basefolder = New System.Windows.Forms.TextBox
        Me.Button2 = New System.Windows.Forms.Button
        Me.Label6 = New System.Windows.Forms.Label
        Me.foldername = New System.Windows.Forms.TextBox
        Me.depth = New System.Windows.Forms.NumericUpDown
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.amount = New System.Windows.Forms.NumericUpDown
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.Panel1 = New System.Windows.Forms.Panel
        CType(Me.depth, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.amount, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ImageList1
        '
        Me.ImageList1.ImageSize = New System.Drawing.Size(16, 16)
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        '
        'Timer2
        '
        Me.Timer2.Interval = 1000
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.Transparent
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.Gray
        Me.Label8.Location = New System.Drawing.Point(208, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(152, 16)
        Me.Label8.TabIndex = 33
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.ToolTip1.SetToolTip(Me.Label8, "Current System Time")
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.ContextMenu = Me.ContextMenu1
        Me.NotifyIcon1.Icon = CType(resources.GetObject("NotifyIcon1.Icon"), System.Drawing.Icon)
        Me.NotifyIcon1.Text = "Resting..."
        Me.NotifyIcon1.Visible = True
        '
        'ContextMenu1
        '
        Me.ContextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem2, Me.MenuItem3})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.Text = "Display Main Screen"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 1
        Me.MenuItem2.Text = "-"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 2
        Me.MenuItem3.Text = "Exit Application"
        '
        'FolderBrowserDialog1
        '
        Me.FolderBrowserDialog1.Description = "Select the base folder in which to create the Test Bed folder in"
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.Transparent
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.Gray
        Me.Label9.Location = New System.Drawing.Point(336, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(109, 16)
        Me.Label9.TabIndex = 54
        Me.Label9.Text = "BUILD 20060118.3"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.Color.DeepSkyBlue
        Me.Button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button3.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.ForeColor = System.Drawing.Color.Black
        Me.Button3.Location = New System.Drawing.Point(24, 224)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(104, 23)
        Me.Button3.TabIndex = 70
        Me.Button3.Text = "Generate Test Bed"
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.White
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Gray
        Me.Label7.Location = New System.Drawing.Point(152, 257)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(112, 26)
        Me.Label7.TabIndex = 69
        Me.Label7.Text = "Resting..."
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.White
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.LimeGreen
        Me.Label1.Location = New System.Drawing.Point(360, 264)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(66, 16)
        Me.Label1.TabIndex = 67
        Me.Label1.Text = "Waiting"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'PictureBox5
        '
        Me.PictureBox5.BackgroundImage = CType(resources.GetObject("PictureBox5.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox5.Location = New System.Drawing.Point(336, 264)
        Me.PictureBox5.Name = "PictureBox5"
        Me.PictureBox5.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox5.TabIndex = 66
        Me.PictureBox5.TabStop = False
        '
        'PictureBox4
        '
        Me.PictureBox4.BackgroundImage = CType(resources.GetObject("PictureBox4.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox4.Location = New System.Drawing.Point(320, 264)
        Me.PictureBox4.Name = "PictureBox4"
        Me.PictureBox4.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox4.TabIndex = 65
        Me.PictureBox4.TabStop = False
        '
        'PictureBox3
        '
        Me.PictureBox3.BackgroundImage = CType(resources.GetObject("PictureBox3.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox3.Location = New System.Drawing.Point(304, 264)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox3.TabIndex = 64
        Me.PictureBox3.TabStop = False
        '
        'PictureBox2
        '
        Me.PictureBox2.BackgroundImage = CType(resources.GetObject("PictureBox2.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(288, 264)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox2.TabIndex = 63
        Me.PictureBox2.TabStop = False
        '
        'PictureBox1
        '
        Me.PictureBox1.BackgroundImage = CType(resources.GetObject("PictureBox1.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(272, 264)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox1.TabIndex = 62
        Me.PictureBox1.TabStop = False
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.White
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label2.Location = New System.Drawing.Point(144, 256)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(288, 32)
        Me.Label2.TabIndex = 68
        Me.Label2.Text = "Label2"
        '
        'term1
        '
        Me.term1.BackColor = System.Drawing.Color.White
        Me.term1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.term1.ForeColor = System.Drawing.Color.Black
        Me.term1.Location = New System.Drawing.Point(208, 80)
        Me.term1.Name = "term1"
        Me.term1.Size = New System.Drawing.Size(72, 20)
        Me.term1.TabIndex = 71
        Me.term1.Text = "Term 1"
        '
        'folderscreated
        '
        Me.folderscreated.BackColor = System.Drawing.Color.Transparent
        Me.folderscreated.ForeColor = System.Drawing.Color.Black
        Me.folderscreated.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.folderscreated.Location = New System.Drawing.Point(208, 216)
        Me.folderscreated.Name = "folderscreated"
        Me.folderscreated.Size = New System.Drawing.Size(80, 16)
        Me.folderscreated.TabIndex = 73
        Me.folderscreated.Text = "0"
        Me.folderscreated.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.DeepSkyBlue
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.Black
        Me.Button1.Location = New System.Drawing.Point(24, 256)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(104, 23)
        Me.Button1.TabIndex = 74
        Me.Button1.Text = "Remove Test Bed"
        '
        'filescreated
        '
        Me.filescreated.BackColor = System.Drawing.Color.Transparent
        Me.filescreated.ForeColor = System.Drawing.Color.Black
        Me.filescreated.Location = New System.Drawing.Point(208, 232)
        Me.filescreated.Name = "filescreated"
        Me.filescreated.Size = New System.Drawing.Size(80, 16)
        Me.filescreated.TabIndex = 75
        Me.filescreated.Text = "0"
        Me.filescreated.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'term2
        '
        Me.term2.BackColor = System.Drawing.Color.White
        Me.term2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.term2.ForeColor = System.Drawing.Color.Black
        Me.term2.Location = New System.Drawing.Point(208, 104)
        Me.term2.Name = "term2"
        Me.term2.Size = New System.Drawing.Size(72, 20)
        Me.term2.TabIndex = 76
        Me.term2.Text = "Term 2"
        '
        'term3
        '
        Me.term3.BackColor = System.Drawing.Color.White
        Me.term3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.term3.ForeColor = System.Drawing.Color.Black
        Me.term3.Location = New System.Drawing.Point(208, 128)
        Me.term3.Name = "term3"
        Me.term3.Size = New System.Drawing.Size(72, 20)
        Me.term3.TabIndex = 77
        Me.term3.Text = "Term 3"
        '
        'term4
        '
        Me.term4.BackColor = System.Drawing.Color.White
        Me.term4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.term4.ForeColor = System.Drawing.Color.Black
        Me.term4.Location = New System.Drawing.Point(208, 152)
        Me.term4.Name = "term4"
        Me.term4.Size = New System.Drawing.Size(72, 20)
        Me.term4.TabIndex = 78
        Me.term4.Text = "Term 4"
        '
        'term5
        '
        Me.term5.BackColor = System.Drawing.Color.White
        Me.term5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.term5.ForeColor = System.Drawing.Color.Black
        Me.term5.Location = New System.Drawing.Point(208, 176)
        Me.term5.Name = "term5"
        Me.term5.Size = New System.Drawing.Size(72, 20)
        Me.term5.TabIndex = 79
        Me.term5.Text = "Term 5"
        '
        'extension5
        '
        Me.extension5.BackColor = System.Drawing.Color.White
        Me.extension5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.extension5.ForeColor = System.Drawing.Color.Black
        Me.extension5.Location = New System.Drawing.Point(304, 176)
        Me.extension5.Name = "extension5"
        Me.extension5.Size = New System.Drawing.Size(72, 20)
        Me.extension5.TabIndex = 84
        Me.extension5.Text = "ex5"
        '
        'extension4
        '
        Me.extension4.BackColor = System.Drawing.Color.White
        Me.extension4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.extension4.ForeColor = System.Drawing.Color.Black
        Me.extension4.Location = New System.Drawing.Point(304, 152)
        Me.extension4.Name = "extension4"
        Me.extension4.Size = New System.Drawing.Size(72, 20)
        Me.extension4.TabIndex = 83
        Me.extension4.Text = "ex4"
        '
        'extension3
        '
        Me.extension3.BackColor = System.Drawing.Color.White
        Me.extension3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.extension3.ForeColor = System.Drawing.Color.Black
        Me.extension3.Location = New System.Drawing.Point(304, 128)
        Me.extension3.Name = "extension3"
        Me.extension3.Size = New System.Drawing.Size(72, 20)
        Me.extension3.TabIndex = 82
        Me.extension3.Text = "ex3"
        '
        'extension2
        '
        Me.extension2.BackColor = System.Drawing.Color.White
        Me.extension2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.extension2.ForeColor = System.Drawing.Color.Black
        Me.extension2.Location = New System.Drawing.Point(304, 104)
        Me.extension2.Name = "extension2"
        Me.extension2.Size = New System.Drawing.Size(72, 20)
        Me.extension2.TabIndex = 81
        Me.extension2.Text = "ex2"
        '
        'extension1
        '
        Me.extension1.BackColor = System.Drawing.Color.White
        Me.extension1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.extension1.ForeColor = System.Drawing.Color.Black
        Me.extension1.Location = New System.Drawing.Point(304, 80)
        Me.extension1.Name = "extension1"
        Me.extension1.Size = New System.Drawing.Size(72, 20)
        Me.extension1.TabIndex = 80
        Me.extension1.Text = "ex1"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(200, 64)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(72, 16)
        Me.Label3.TabIndex = 85
        Me.Label3.Text = "Input Terms:"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.ForeColor = System.Drawing.Color.Black
        Me.Label4.Location = New System.Drawing.Point(296, 64)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(88, 16)
        Me.Label4.TabIndex = 86
        Me.Label4.Text = "File Extensions:"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.ForeColor = System.Drawing.Color.Black
        Me.Label5.Location = New System.Drawing.Point(16, 16)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(144, 16)
        Me.Label5.TabIndex = 88
        Me.Label5.Text = "Test Bed Base Folder:"
        '
        'basefolder
        '
        Me.basefolder.BackColor = System.Drawing.Color.White
        Me.basefolder.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.basefolder.ForeColor = System.Drawing.Color.Black
        Me.basefolder.Location = New System.Drawing.Point(24, 32)
        Me.basefolder.Name = "basefolder"
        Me.basefolder.ReadOnly = True
        Me.basefolder.Size = New System.Drawing.Size(344, 20)
        Me.basefolder.TabIndex = 87
        Me.basefolder.Text = ""
        '
        'Button2
        '
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.ForeColor = System.Drawing.Color.DimGray
        Me.Button2.Location = New System.Drawing.Point(374, 32)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(56, 20)
        Me.Button2.TabIndex = 89
        Me.Button2.Text = "BROWSE"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.Transparent
        Me.Label6.ForeColor = System.Drawing.Color.Black
        Me.Label6.Location = New System.Drawing.Point(16, 64)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(152, 16)
        Me.Label6.TabIndex = 91
        Me.Label6.Text = "Test Bed Folder Name:"
        '
        'foldername
        '
        Me.foldername.BackColor = System.Drawing.Color.White
        Me.foldername.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.foldername.ForeColor = System.Drawing.Color.Black
        Me.foldername.Location = New System.Drawing.Point(24, 80)
        Me.foldername.Name = "foldername"
        Me.foldername.Size = New System.Drawing.Size(152, 20)
        Me.foldername.TabIndex = 90
        Me.foldername.Text = "Test Bed"
        '
        'depth
        '
        Me.depth.BackColor = System.Drawing.Color.White
        Me.depth.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.depth.ForeColor = System.Drawing.Color.Black
        Me.depth.Location = New System.Drawing.Point(24, 176)
        Me.depth.Maximum = New Decimal(New Integer() {5, 0, 0, 0})
        Me.depth.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.depth.Name = "depth"
        Me.depth.Size = New System.Drawing.Size(56, 20)
        Me.depth.TabIndex = 92
        Me.depth.Value = New Decimal(New Integer() {5, 0, 0, 0})
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.Color.Transparent
        Me.Label10.ForeColor = System.Drawing.Color.Black
        Me.Label10.Location = New System.Drawing.Point(16, 160)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(80, 16)
        Me.Label10.TabIndex = 93
        Me.Label10.Text = "Folder Depth:"
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.Color.Transparent
        Me.Label11.ForeColor = System.Drawing.Color.Black
        Me.Label11.Location = New System.Drawing.Point(16, 112)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(128, 16)
        Me.Label11.TabIndex = 95
        Me.Label11.Text = "Max File/Folder Amount:"
        '
        'amount
        '
        Me.amount.BackColor = System.Drawing.Color.White
        Me.amount.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.amount.ForeColor = System.Drawing.Color.Black
        Me.amount.Location = New System.Drawing.Point(24, 128)
        Me.amount.Maximum = New Decimal(New Integer() {10, 0, 0, 0})
        Me.amount.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.amount.Name = "amount"
        Me.amount.Size = New System.Drawing.Size(56, 20)
        Me.amount.TabIndex = 94
        Me.amount.Value = New Decimal(New Integer() {9, 0, 0, 0})
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.Color.Transparent
        Me.Label12.ForeColor = System.Drawing.Color.Black
        Me.Label12.Location = New System.Drawing.Point(296, 216)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(88, 16)
        Me.Label12.TabIndex = 96
        Me.Label12.Text = "Folders Effected"
        '
        'Label13
        '
        Me.Label13.BackColor = System.Drawing.Color.Transparent
        Me.Label13.ForeColor = System.Drawing.Color.Black
        Me.Label13.Location = New System.Drawing.Point(296, 232)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(80, 16)
        Me.Label13.TabIndex = 97
        Me.Label13.Text = "Files Effected"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Transparent
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Location = New System.Drawing.Point(16, 216)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(120, 72)
        Me.Panel1.TabIndex = 98
        '
        'Main_Screen
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.LightSkyBlue
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.ClientSize = New System.Drawing.Size(450, 304)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.amount)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.depth)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.foldername)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.basefolder)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.extension5)
        Me.Controls.Add(Me.extension4)
        Me.Controls.Add(Me.extension3)
        Me.Controls.Add(Me.extension2)
        Me.Controls.Add(Me.extension1)
        Me.Controls.Add(Me.term5)
        Me.Controls.Add(Me.term4)
        Me.Controls.Add(Me.term3)
        Me.Controls.Add(Me.term2)
        Me.Controls.Add(Me.filescreated)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.folderscreated)
        Me.Controls.Add(Me.term1)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.PictureBox5)
        Me.Controls.Add(Me.PictureBox4)
        Me.Controls.Add(Me.PictureBox3)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Panel1)
        Me.ForeColor = System.Drawing.Color.White
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(456, 336)
        Me.MinimumSize = New System.Drawing.Size(456, 336)
        Me.Name = "Main_Screen"
        Me.ShowInTaskbar = False
        Me.Text = "Testbed Folder Contents Generator"
        CType(Me.depth, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.amount, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private current_light As Integer = 0
    Private current_colour As Integer = 0
    Private currently_working As Boolean = False




    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            If ex.Message.IndexOf("Thread was being aborted") < 0 Then
                Dim Display_Message1 As New Display_Message("The Application encountered the following problem: " & vbCrLf & ex.ToString)
                Display_Message1.ShowDialog()
                Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
                If dir.Exists = False Then
                    dir.Create()
                End If
                Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
                filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & " - " & ex.ToString)
                filewriter.Flush()
                filewriter.Close()
            End If
        Catch exc As Exception

            MsgBox("An error occurred in Testbed Folder Contents Generator's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Private Sub run_green_lights()
        Try
            Label1.ForeColor = Color.LimeGreen
            Label1.Text = "Waiting"
            Label7.Text = "Resting..."

            current_light = current_light - 1
            If current_light < 1 Then
                current_light = 5
            End If
            current_colour = 0
            PictureBox1.Image = ImageList1.Images(1)
            PictureBox2.Image = ImageList1.Images(1)
            PictureBox3.Image = ImageList1.Images(1)
            PictureBox4.Image = ImageList1.Images(1)
            PictureBox5.Image = ImageList1.Images(1)

            Select Case current_light
                Case 0

                    PictureBox1.Image = ImageList1.Images(0)
                Case 1

                    PictureBox2.Image = ImageList1.Images(0)
                Case 2

                    PictureBox3.Image = ImageList1.Images(0)
                Case 3

                    PictureBox4.Image = ImageList1.Images(0)
                Case 4

                    PictureBox5.Image = ImageList1.Images(0)
                Case 5

                    PictureBox1.Image = ImageList1.Images(0)
            End Select

            current_light = current_light + 1
            If current_light > 5 Then
                current_light = 1
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub run_orange_lights()
        Try
            Label1.ForeColor = Color.DarkOrange
            Label1.Text = "Working"

            current_light = current_light - 1
            If current_light < 1 Then
                current_light = 5
            End If
            current_colour = 1
            PictureBox1.Image = ImageList1.Images(3)
            PictureBox2.Image = ImageList1.Images(3)
            PictureBox3.Image = ImageList1.Images(3)
            PictureBox4.Image = ImageList1.Images(3)
            PictureBox5.Image = ImageList1.Images(3)
            Select Case current_light
                Case 0
                    PictureBox1.Image = ImageList1.Images(2)
                Case 1
                    PictureBox2.Image = ImageList1.Images(2)
                Case 2
                    PictureBox3.Image = ImageList1.Images(2)
                Case 3
                    PictureBox4.Image = ImageList1.Images(2)
                Case 4
                    PictureBox5.Image = ImageList1.Images(2)
                Case 5
                    PictureBox1.Image = ImageList1.Images(2)
            End Select

            current_light = current_light + 1
            If current_light > 5 Then
                current_light = 1
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub run_lights()
        Try
            If current_colour = 1 Then
                Select Case current_light
                    Case 0
                        PictureBox5.Image = ImageList1.Images(3)
                        PictureBox1.Image = ImageList1.Images(2)
                    Case 1
                        PictureBox1.Image = ImageList1.Images(3)
                        PictureBox2.Image = ImageList1.Images(2)
                    Case 2
                        PictureBox2.Image = ImageList1.Images(3)
                        PictureBox3.Image = ImageList1.Images(2)
                    Case 3
                        PictureBox3.Image = ImageList1.Images(3)
                        PictureBox4.Image = ImageList1.Images(2)
                    Case 4
                        PictureBox4.Image = ImageList1.Images(3)
                        PictureBox5.Image = ImageList1.Images(2)
                    Case 5
                        PictureBox5.Image = ImageList1.Images(3)
                        PictureBox1.Image = ImageList1.Images(2)
                End Select
            Else
                Select Case current_light
                    Case 0
                        PictureBox5.Image = ImageList1.Images(1)
                        PictureBox1.Image = ImageList1.Images(0)
                    Case 1
                        PictureBox1.Image = ImageList1.Images(1)
                        PictureBox2.Image = ImageList1.Images(0)
                    Case 2
                        PictureBox2.Image = ImageList1.Images(1)
                        PictureBox3.Image = ImageList1.Images(0)
                    Case 3
                        PictureBox3.Image = ImageList1.Images(1)
                        PictureBox4.Image = ImageList1.Images(0)
                    Case 4
                        PictureBox4.Image = ImageList1.Images(1)
                        PictureBox5.Image = ImageList1.Images(0)
                    Case 5
                        PictureBox5.Image = ImageList1.Images(1)
                        PictureBox1.Image = ImageList1.Images(0)
                End Select

            End If

            current_light = current_light + 1
            If current_light > 5 Then
                current_light = 1
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Try
            run_lights()
            Label8.Text = Format(Now(), "dd/MM/yyyy hh:mm:ss tt")
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Main_Screen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'basefolder.Text = Application.StartupPath
            Load_Registry_Values()
            Label8.Text = Format(Now(), "dd/MM/yyyy hh:mm:ss tt")
            Timer2.Start()
            Label5.Select()
            dataloaded = True
            splash_loader.Visible = False
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub exit_application()
        Try
            Save_Registry_Values()
            If Worker1.WorkerThread Is Nothing = False Then
                Worker1.WorkerThread.Abort()
                Worker1.Dispose()
            End If
            NotifyIcon1.Dispose()
            Application.Exit()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Main_Screen_closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        Try
            exit_application()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub


    Public Sub WorkerHandler(ByVal Result As String)
        Try
            currently_working = False

            NotifyIcon1.Text = "Resting... "
            run_green_lights()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Public Sub WorkerFileCreatedHandler()
        Try
            filescreated.Text = CLng(filescreated.Text) + 1
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Public Sub WorkerProgressHandler()
        Try
            folderscreated.Text = CLng(folderscreated.Text) + 1
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub
    Public Sub WorkerErrorEncounteredHandler(ByVal ex As Exception)
        Try
            Error_Handler(ex, "Worker Error Encountered")
        Catch exc As Exception
            Error_Handler(exc)
        End Try
    End Sub

    Private Sub run_worker(Optional ByVal threadselect As Integer = 1)
        run_orange_lights()
        folderscreated.Text = 0
        filescreated.Text = 0
        Worker1.foldername = (basefolder.Text & "\" & foldername.Text).Replace("\\", "\")
        Worker1.depth = depth.Value
        Worker1.amount = amount.Value
        Worker1.term1 = term1.Text
        Worker1.term2 = term2.Text
        Worker1.term3 = term3.Text
        Worker1.term4 = term4.Text
        Worker1.term5 = term5.Text
        Worker1.extension1 = extension1.Text
        Worker1.extension2 = extension2.Text
        Worker1.extension3 = extension3.Text
        Worker1.extension4 = extension4.Text
        Worker1.extension5 = extension5.Text

        Select Case threadselect
            Case 1
                Worker1.ChooseThreads(1)
            Case 2
                Worker1.ChooseThreads(2)
        End Select

        currently_working = True
    End Sub



    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click
        Try
            Me.Close()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub show_application()
        Try
            Me.Opacity = 1

            Me.BringToFront()
            Me.Refresh()
            Me.WindowState = FormWindowState.Normal

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub NotifyIcon1_dblclick(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon1.DoubleClick
        show_application()
    End Sub
    Private Sub NotifyIcon1_snglclick(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon1.Click
        show_application()
    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click
        show_application()
    End Sub

    Private Sub Main_Screen_resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        Try

            If Me.WindowState = FormWindowState.Minimized Then
                NotifyIcon1.Visible = True
                Me.Opacity = 0
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub force_check(Optional ByVal threadselect As Integer = 1)
        Try
            Label7.Text = "Busy Working..."
            NotifyIcon1.Text = "Generating Test Bed..."
            If currently_working = False Then
                run_worker(threadselect)
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub


    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        force_check()
    End Sub




    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        force_check(2)
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            Dim tinfo As DirectoryInfo = New DirectoryInfo(basefolder.Text)
            If tinfo.Exists = True Then
                FolderBrowserDialog1.SelectedPath = basefolder.Text
            End If
            Dim result As DialogResult = FolderBrowserDialog1.ShowDialog
            If result = DialogResult.OK Then
                basefolder.Text = FolderBrowserDialog1.SelectedPath
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Public Sub Load_Registry_Values()
        Try
            Dim configflag As Boolean
            configflag = False
            Dim str As String
            Dim keyflag1 As Boolean = False
            Dim oReg As RegistryKey = Registry.LocalMachine
            Dim keys() As String = oReg.GetSubKeyNames()
            System.Array.Sort(keys)

            For Each str In keys
                If str.Equals("Software\Testbed Folder Contents Generator") = True Then
                    keyflag1 = True
                    Exit For
                End If
            Next str

            If keyflag1 = False Then
                oReg.CreateSubKey("Software\Testbed Folder Contents Generator")
            End If

            keyflag1 = False

            Dim oKey As RegistryKey = oReg.OpenSubKey("Software\Testbed Folder Contents Generator", True)

            str = oKey.GetValue("term1")
            If Not IsNothing(str) And Not (str = "") Then
                term1.Text = str
            Else
                oKey.SetValue("term1", "Term 1")
                term1.Text = "Term 1"
            End If

            str = oKey.GetValue("term2")
            If Not IsNothing(str) And Not (str = "") Then
                term2.Text = str
            Else
                oKey.SetValue("term2", "Term 2")
                term2.Text = "Term 2"
            End If

            str = oKey.GetValue("term3")
            If Not IsNothing(str) And Not (str = "") Then
                term3.Text = str
            Else
                oKey.SetValue("term3", "Term 3")
                term3.Text = "Term 3"
            End If

            str = oKey.GetValue("term4")
            If Not IsNothing(str) And Not (str = "") Then
                term4.Text = str
            Else
                oKey.SetValue("term4", "Term 4")
                term1.Text = "Term 4"
            End If

            str = oKey.GetValue("term5")
            If Not IsNothing(str) And Not (str = "") Then
                term5.Text = str
            Else
                oKey.SetValue("term5", "Term 5")
                term5.Text = "Term 5"
            End If

            str = oKey.GetValue("extension1")
            If Not IsNothing(str) And Not (str = "") Then
                extension1.Text = str
            Else
                oKey.SetValue("extension1", "ex1")
                extension1.Text = "ex1"
            End If

            str = oKey.GetValue("extension2")
            If Not IsNothing(str) And Not (str = "") Then
                extension2.Text = str
            Else
                oKey.SetValue("extension2", "ex2")
                extension2.Text = "ex2"
            End If

            str = oKey.GetValue("extension3")
            If Not IsNothing(str) And Not (str = "") Then
                extension3.Text = str
            Else
                oKey.SetValue("extension3", "ex3")
                extension3.Text = "ex3"
            End If

            str = oKey.GetValue("extension4")
            If Not IsNothing(str) And Not (str = "") Then
                extension4.Text = str
            Else
                oKey.SetValue("extension4", "ex4")
                extension4.Text = "ex4"
            End If

            str = oKey.GetValue("extension5")
            If Not IsNothing(str) And Not (str = "") Then
                extension5.Text = str
            Else
                oKey.SetValue("extension5", "ex5")
                extension5.Text = "ex5"
            End If

            str = oKey.GetValue("foldername")
            If Not IsNothing(str) And Not (str = "") Then
                foldername.Text = str
            Else
                oKey.SetValue("foldername", "Test Bed")
                foldername.Text = "Test Bed"
            End If

            str = oKey.GetValue("basefolder")
            If Not IsNothing(str) And Not (str = "") Then
                basefolder.Text = str
            Else
                oKey.SetValue("basefolder", Application.StartupPath)
                basefolder.Text = Application.StartupPath
            End If

            str = oKey.GetValue("amount")
            If Not IsNothing(str) And Not (str = "") Then
                amount.Value = CInt(str)
            Else
                oKey.SetValue("amount", 9)
                amount.Value = 9
            End If

            str = oKey.GetValue("depth")
            If Not IsNothing(str) And Not (str = "") Then
                depth.Value = CInt(str)
            Else
                oKey.SetValue("depth", 5)
                depth.Value = 5
            End If

            oKey.Close()
            oReg.Close()

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Save_Registry_Values()
        Try
            Dim oReg As RegistryKey = Registry.LocalMachine
            Dim oKey As RegistryKey = oReg.OpenSubKey("Software\Testbed Folder Contents Generator", True)

            oKey.SetValue("term1", term1.Text)
            oKey.SetValue("term2", term2.Text)
            oKey.SetValue("term3", term3.Text)
            oKey.SetValue("term4", term4.Text)
            oKey.SetValue("term5", term5.Text)
            oKey.SetValue("extension1", extension1.Text)
            oKey.SetValue("extension2", extension2.Text)
            oKey.SetValue("extension3", extension3.Text)
            oKey.SetValue("extension4", extension4.Text)
            oKey.SetValue("extension5", extension5.Text)
            oKey.SetValue("depth", depth.Value)
            oKey.SetValue("amount", amount.Value)
            oKey.SetValue("foldername", foldername.Text)
            oKey.SetValue("basefolder", basefolder.Text)
            

            oKey.Close()
            oReg.Close()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

End Class
