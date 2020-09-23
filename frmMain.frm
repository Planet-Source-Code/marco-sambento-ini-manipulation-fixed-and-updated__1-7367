VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "INI Manipulation"
   ClientHeight    =   8655
   ClientLeft      =   285
   ClientTop       =   1800
   ClientWidth     =   10965
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MouseIcon       =   "frmMain.frx":1272
   MousePointer    =   99  'Custom
   ScaleHeight     =   8655
   ScaleWidth      =   10965
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ilDisabled 
      Left            =   2031
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":157C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3D30
            Key             =   "Value"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":460C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4EE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":57C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":60A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6EF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7D48
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":85DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8E70
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":918C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B940
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BC5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C538
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CE14
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D6F0
            Key             =   "Settings"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FEA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10780
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1105C
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11378
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11C54
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilSmall 
      Left            =   2037
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11F70
            Key             =   "Section"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14724
            Key             =   "Key"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15578
            Key             =   "Value"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15E54
            Key             =   "selKey"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16CA8
            Key             =   "selSection"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16FC4
            Key             =   "Move"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":178A0
            Key             =   "Copy"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   8400
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2
            MinWidth        =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14977
            MinWidth        =   2
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   979
            MinWidth        =   970
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1931
            MinWidth        =   1940
            TextSave        =   "25-04-2000"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ilTollbar 
      Left            =   2031
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1817C
            Key             =   "Section"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A930
            Key             =   "Value"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B20C
            Key             =   "WINini"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BAE8
            Key             =   "SystemIni"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C3C4
            Key             =   "Gridlines"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CCA0
            Key             =   "Key"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D57C
            Key             =   "selKey"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1DE58
            Key             =   "Left"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E6EC
            Key             =   "Right"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1EF80
            Key             =   "selSection"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F29C
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":21A50
            Key             =   "New"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":21D6C
            Key             =   "lOpen"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22648
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22F24
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23800
            Key             =   "Settings"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":25FB4
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":26890
            Key             =   "About"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27464
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27780
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2805C
            Key             =   "Notepad"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   7815
      Left            =   10671
      ScaleHeight     =   3402.987
      ScaleMode       =   0  'User
      ScaleWidth      =   780
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   72
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2031
      Top             =   6960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "ini"
      DialogTitle     =   "Open INI"
      Filter          =   "INI Files|*.ini|All Files|*.*"
   End
   Begin MSComctlLib.ListView lvListView 
      Height          =   7815
      Left            =   2637
      TabIndex        =   2
      Top             =   600
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   13785
      View            =   3
      Arrange         =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HoverSelection  =   -1  'True
      PictureAlignment=   3
      _Version        =   393217
      Icons           =   "ilTollbar"
      SmallIcons      =   "ilSmall"
      ColHdrIcons     =   "ilSmall"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Key"
         Text            =   "Key"
         Object.Width           =   0
         ImageIndex      =   2
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Value"
         Text            =   "Value"
         Object.Width           =   0
         ImageIndex      =   3
      EndProperty
      Picture         =   "frmMain.frx":28938
   End
   Begin MSComctlLib.TreeView tvTreeView 
      Height          =   7815
      Left            =   117
      TabIndex        =   3
      Top             =   600
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   13785
      _Version        =   393217
      Indentation     =   0
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   1
      HotTracking     =   -1  'True
      ImageList       =   "ilSmall"
      Appearance      =   0
   End
   Begin MSComctlLib.Toolbar tbToolbar 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   1005
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "ilTollbar"
      DisabledImageList=   "ilDisabled"
      HotImageList    =   "ilDisabled"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   22
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbStart"
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Left"
            Object.ToolTipText     =   "Back"
            ImageKey        =   "Left"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Right"
            Object.ToolTipText     =   "Forward"
            ImageKey        =   "Right"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Description     =   "New Ini"
            Object.ToolTipText     =   "New Ini"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open INI"
            Description     =   "Open INI"
            Object.ToolTipText     =   "Open INI"
            ImageKey        =   "lOpen"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Close"
            Description     =   "Close"
            Object.ToolTipText     =   "Close"
            ImageKey        =   "Close"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Rename"
            Description     =   "Rename Selected Item"
            Object.ToolTipText     =   "Rename Selected Item"
            ImageKey        =   "Value"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Description     =   "Delete Item"
            Object.ToolTipText     =   "Delete"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "WordPad"
            Description     =   "Open in WordPad"
            Object.ToolTipText     =   "Open Ini in WordPad"
            ImageKey        =   "Notepad"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Description     =   "Refresh"
            Object.ToolTipText     =   "Refresh"
            ImageKey        =   "Refresh"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Win.ini"
            Description     =   "Open Win.ini"
            Object.ToolTipText     =   "Open Win.ini"
            ImageKey        =   "WINini"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "System.ini"
            Description     =   "Open system.ini"
            Object.ToolTipText     =   "Open system.ini"
            ImageKey        =   "SystemIni"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "View Gridlines"
            Description     =   "View Gridlines"
            Object.ToolTipText     =   "View Gridlines"
            ImageKey        =   "Gridlines"
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Description     =   "Find Item"
            Object.ToolTipText     =   "Find"
            ImageKey        =   "Find"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "about"
            Description     =   "about"
            Object.ToolTipText     =   "About"
            ImageKey        =   "About"
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Description     =   "Exit"
            Object.ToolTipText     =   "Exit"
            ImageKey        =   "Exit"
         EndProperty
      EndProperty
   End
   Begin VB.Image imgSplitter 
      Height          =   7815
      Left            =   2517
      MousePointer    =   9  'Size W E
      Top             =   600
      Width           =   120
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New..."
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnublank2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditEditWordPad 
         Caption         =   "&Edit in WordPad"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuEditRename 
         Caption         =   "&Rename"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnublank10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "&Find"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnublank3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuViewStatusbar 
         Caption         =   "&Status Bar"
         Checked         =   -1  'True
         Shortcut        =   ^S
      End
      Begin VB.Menu mnublank5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewGridlines 
         Caption         =   "&View GridLines"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnublank6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewIcons 
         Caption         =   "&Icons"
      End
      Begin VB.Menu mnuViewSmallIcons 
         Caption         =   "S&mall Icons"
      End
      Begin VB.Menu mnuViewList 
         Caption         =   "&List"
      End
      Begin VB.Menu mnuViewDetails 
         Caption         =   "&Details"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnublank7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
   Begin VB.Menu mnupop 
      Caption         =   "&pop"
      Visible         =   0   'False
      Begin VB.Menu mnuCreate 
         Caption         =   "&Create"
         Begin VB.Menu mnuCreateSection 
            Caption         =   "Create &Section"
         End
         Begin VB.Menu mnuCreateKeyValue 
            Caption         =   "Create &Key/Value"
         End
      End
      Begin VB.Menu mnuRename 
         Caption         =   "&Rename"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
         Begin VB.Menu mnuDeleteSection 
            Caption         =   "Delete &Section"
         End
         Begin VB.Menu mnuDeleteKey 
            Caption         =   "Delete &Key"
         End
         Begin VB.Menu mnuDeleteValue 
            Caption         =   "Delete &Value"
         End
      End
      Begin VB.Menu mnublank8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExpande 
         Caption         =   "&Expande"
      End
      Begin VB.Menu mnuExpandeAll 
         Caption         =   "Expande &All"
      End
      Begin VB.Menu mnublank9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCollapse 
         Caption         =   "Co&llapse"
      End
      Begin VB.Menu mnuCollapseAll 
         Caption         =   "Colla&pse All"
      End
      Begin VB.Menu mnublank11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Re&fresh"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'|'==================================================================================
'|'**********************************************************************************
'|'
'|'* Author: Marco Sambento
'|'* Date: April/2000
'|'* Notes: This project was made to help understanding, mainly the routines of
'|'   manipulating INI Files, such as reading/writing, renaming and deleting sections
'|'   and keys/values, but it's also a good exercise on TreeView operations and some
'|'   other usefull routines.
'|'
'|'   I take no responsability in the misuse of this program or for any damage
'|'   caused!
'|'
'|'   Be specially careful when editing 'win.ini' and 'system.ini' files!
'|'
'|'FEATURES:
'|'   - Dragdrop keys, option for move or copy Key, by using Ctrl Key
'|'   - Context Menus
'|'   - Open/Close/Create INI File
'|'   - Resize, etc...
'|'
'|'
'|'  Please send all feedback and report bugs/suggestions to:
'|'* Email: marco.sambento@netc.pt
'|'
'|'
'|' ---------  Please Keep This Text When Using/Distributing This Code  ------------
'|'**********************************************************************************
'|'==================================================================================

Dim indrag As Boolean           ' Flag that signals a Drag Drop operation.
Public nodDrag As Node          ' Item that is being dragged.
Public CopyNode As Boolean      ' in a drag 'n drop operation, sets move or copy node
Dim mbMoving As Boolean         ' determines if treeview/listview are being sized
Public FileOpen As Boolean      ' determines if a file is open
Public FileIsEmpty As Boolean   ' if file is empty
Dim numSections As Integer      ' number of sections to display in statusbar
Const sglSplitLimit = 2250      ' limit for resizing treeview and listview

Private Sub Form_Load()

If Command$ <> vbNullString Then
    INIPath = Command$
    
    If Dir(INIPath) <> vbNullString Then
        FilltvTreeView
    End If

Else
'replace with new file or simply erase it
   INIPath = App.Path & "\marco.ini"
    If Dir(INIPath) = vbNullString Then
        mnuFileOpen_Click
    Else
        FilltvTreeView
    End If
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Button As VbMsgBoxResult

    Button = MsgBox("Are You Sure?", vbInformation + vbYesNo, "Exit...")
    If Button = vbNo Then Cancel = True: Exit Sub
    Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
 Dim i As Integer
    'close all sub forms
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
End Sub

'**********************************************************************************

'==================================================================================
'            READ INI FILE AND FILL TREEVIEW    marco.sambento@netc.pt
'==================================================================================

Public Function FilltvTreeView()
On Error Resume Next
    If Not FileOpen Then Exit Function
    
    Dim Button As Integer
    Dim firstChar As String
    Dim lastChar As String
    Dim FileData As String
    Dim Section As String
    Dim posKey As Integer
    Dim Key As String
    Dim nodeKey As Node
    
    Filename = INIPath
    Me.Caption = App.Title & " - '" & StrConv(Filename, vbProperCase) & "'"
    tvTreeView.Nodes.clear
    numSections = 0

    If Len(Filename) Then
        Open Filename For Input As #1
        Do While Not EOF(1)     ' while it isn't the end of file, get another line
            Line Input #1, FileData 'line text
            
            firstChar = Left(FileData, 1)
            lastChar = Right(FileData, 1)
            
            If firstChar = "[" And lastChar = "]" Then 'if it's a section
                Section = Mid(FileData, 2, Len(FileData) - 2)
                'then add section
                tvTreeView.Nodes.Add , , Section, Section, "Section", "selSection"
                numSections = numSections + 1
            Else ' it's a key
                posKey = InStr(FileData, "=")
                If posKey <> 0 And Section <> vbNullString Then
                'add key to current section
                    Key = Left(FileData, posKey - 1)
                    tvTreeView.Nodes.Add Section, tvwChild, , Key, "Key", "selKey"
                End If
            End If
        Loop
        Close #1
    End If

Call CheckFile
If frmFind.FindIsLoaded Then frmFind.txtFind_KeyUp 0, 0 'update findform to new file
If FileIsEmpty Then lvListView.ListItems.clear _
: MsgBox "Invalid or Empty INI File!", vbCritical, "Invalid" _
: Exit Function
    
    tvTreeView.SelectedItem = tvTreeView.Nodes.Item(1)
    tvTreeView.SelectedItem.Expanded = True
    tvTreeView_NodeClick tvTreeView.SelectedItem    'update listview
End Function

'==================================================================================
'  READ SECTIONS AND KEYS IN TREEVIEW AND FILL LISTVIEW     marco.sambento@netc.pt
'==================================================================================

Public Sub tvTreeView_NodeClick(ByVal Node As MSComctlLib.Node)
If Not FileOpen Or FileIsEmpty Then Exit Sub

    Dim Value As String
    Dim numKey As Integer
    Dim Childrens As Integer
    Dim first As Integer
    Dim last As Integer
    Dim Item As ListItem

    lvListView.ListItems.clear      'clear listview
    
    If Node.Parent Is Nothing Then 'if node is a section
    
    mnuViewDetails.Enabled = True 'enable menu 'Details'
    mnuViewDetails_Click          'set listview mode as details
        
        If Node.Children > 0 Then    'if it has at least a key
        numKey = Node.Child.FirstSibling.Index    'first key index
        
        Do
             Set Item = lvListView.ListItems.Add(, , _
             tvTreeView.Nodes.Item(numKey).Text, "Key", "Key")
             'add the key to listview
             With Item
                Section = Node.Text
                Key = Item.Text
                Value = GetVal(Section, Key)    'get value for current key
                .SubItems(1) = Value    'add the value to second column on listview
                .Tag = numKey     'holds the node index because the key can't be a number
             End With

            If numKey <> Node.Child.LastSibling.Index Then _
            numKey = tvTreeView.Nodes(numKey).Next.Index Else Exit Do
        Loop
        End If
    Else    'if selected node is a key in tree
        mnuViewDetails.Enabled = False 'disable menu 'Details'
        mnuViewList_Click
        Section = Node.Parent.Text
        Key = Node.Text
        Value = GetVal(Section, Key)
    
        lvListView.ListItems.Add , , Value, "Value", "Value"
    End If
'checks if there are items and enable/disable menus
    If lvListView.SelectedItem Is Nothing Then
        mnuDeleteValue.Enabled = False
        mnuDeleteKey.Enabled = False
    Else
        mnuDeleteValue.Enabled = True
        mnuDeleteKey.Enabled = True
    End If

UpdateStatusBar
End Sub

'**********************************************************************************

'==================================================================================
'          CREATE      SECTIONS/KEYS/VALUES          marco.sambento@netc.pt
'==================================================================================

Private Sub mnuCreateSection_Click()
   Dim i As Integer
   Dim newSection As String
   Dim newSectionNode As Node
   Dim newKeyNode As Node

        If FileIsEmpty Then  'if it's 1st section create "New Section"
                Create "New Section", "New Key", "New Value"
                'add section to tree, and new key/value
                Set newSectionNode = tvTreeView.Nodes.Add(, , "New Section", "New Section", "Section", "selSection")
                tvTreeView.Nodes.Add "New Section", tvwChild, , "New Key", "Key", "selKey"
                Call CheckFile
        Else    ' check if new section already exists, if so increase section number
ReStartCheck:
            numSection = tvTreeView.SelectedItem.Root.FirstSibling.Index
        Do    'check if matches section
            If tvTreeView.Nodes.Item(numSection).Text = "New Section" & i _
            Then i = i + 1: GoTo ReStartCheck ' if so increases and restart
            'if not check next section
            If numSection <> tvTreeView.SelectedItem.Root.LastSibling.Index _
            Then numSection = tvTreeView.Nodes.Item(numSection).Next.Index Else Exit Do
        Loop
                'create section
            newSection = "New Section" & i
            Create newSection, "New Key", "New Value"
                ' and add it to tree
            Set newSectionNode = tvTreeView.Nodes.Add(, , newSection, newSection, "Section", "selSection")
            tvTreeView.Nodes.Add newSection, tvwChild, , "New Key", "Key", "selKey"
            
        End If
    numSections = numSections + 1
    tvTreeView.SelectedItem = newSectionNode
    tvTreeView_NodeClick newSectionNode
    newSectionNode.Expanded = True
    tvTreeView.StartLabelEdit
End Sub

Private Sub mnuCreateKeyValue_Click()
' does the same as create section, but for keys
   Dim i As Integer
   Dim newKey As Node
   Dim Section As Node
   
   If FileIsEmpty Then mnuCreateSection_Click: Exit Sub
   
    If tvTreeView.SelectedItem.Parent Is Nothing Then
        Set Section = tvTreeView.SelectedItem
    Else
        Set Section = tvTreeView.SelectedItem.Parent
    End If

      If Section.Children = 0 Then '1st key
           Create Section, "New Key", "New Value"
           tvTreeView.Nodes.Add Section, tvwChild, , "New Key", "Key", "selKey"
      Else   'check if key exists
Repeat:    numKey = Section.Child.FirstSibling.Index
         Do
            
         If tvTreeView.Nodes.Item(numKey) = "New Key" & i Then i = i + 1: GoTo Repeat
            
         If numKey <> Section.Child.LastSibling.Index Then _
           numKey = tvTreeView.Nodes.Item(numKey).Next.Index Else Exit Do
         Loop
           
         Create Section, "New Key" & i, "New Value"
         Set newKey = tvTreeView.Nodes.Add(Section, tvwChild, , "New Key" & i, "Key", "selKey")
       End If

   tvTreeView.SelectedItem = newKey
   tvTreeView.StartLabelEdit
End Sub

'==================================================================================
'             DELETE     SECTIONS/KEYS/VALUES         marco.sambento@netc.pt
'==================================================================================

Public Sub mnuDeleteSection_Click()
Dim Button As VbMsgBoxResult

    If tvTreeView.SelectedItem.Parent Is Nothing Then 'menu was poped up in a section
        Section = tvTreeView.SelectedItem
    Else    'menu was poped up in a key
        Section = tvTreeView.SelectedItem.Parent
    End If
        
        Button = MsgBox("Delete Section '" & Section & "'?", _
        vbOKCancel + vbExclamation, "Delete Section...")
        
        If Button = vbCancel Then Exit Sub
    
        DeleteSection Section               'delete section in file
        tvTreeView.Nodes.Remove Section     'delete section in treee
        numSections = numSections - 1   'decrease numsections
        
        If numSections = 0 Then lvListView.ListItems.clear: Call CheckFile: Exit Sub

        tvTreeView_NodeClick tvTreeView.SelectedItem
End Sub

Public Sub mnuDeleteKey_Click()

Dim Button As VbMsgBoxResult

    If tvTreeView.SelectedItem.Parent Is Nothing Then 'menu poped up in section
        Section = tvTreeView.SelectedItem
        Key = lvListView.SelectedItem
        KeyNode = lvListView.SelectedItem.Tag 'if you remember keys in listview
            'hold the key index from tree (see tvtreeview_nodeclick event)
    Else 'menu poped up in key
        Section = tvTreeView.SelectedItem.Parent
        Key = tvTreeView.SelectedItem
        KeyNode = tvTreeView.SelectedItem.Index
    End If
        
        Button = MsgBox("Delete Key '" & Key & "'?", vbOKCancel + vbExclamation, "Delete Key...")
        If Button = vbCancel Then Exit Sub
        
        DeleteKey Section, Key 'delete key in file
        tvTreeView.Nodes.Remove (KeyNode) 'delete key in section

    tvTreeView_NodeClick tvTreeView.SelectedItem
End Sub

Public Sub mnuDeleteValue_Click()
'the same as above, except for subitem
Dim Button As VbMsgBoxResult

    If lvListView.SelectedItem.Tag = vbNullString Then 'only keys in listview have tag
        Section = tvTreeView.SelectedItem.Parent       'see tvtree_nodeclick, key add part
        Key = tvTreeView.SelectedItem
        Value = lvListView.SelectedItem

    Else 'in listview it's only a value and not keys and subitems
        Section = tvTreeView.SelectedItem
        Key = lvListView.SelectedItem
        Value = lvListView.SelectedItem.SubItems(1)
        subItem = True 'to know wich value to delete: a subitem
    End If
    
        Button = MsgBox("Delete value '" & Value & "'?", vbOKCancel + vbExclamation, "Delete value...")
        If Button = vbCancel Then Exit Sub
        
        DeleteValue Section, Key ' delete in file
        'delete in listview, if subitem is true then delete the subitem, else delete value
        If subItem Then lvListView.SelectedItem.SubItems(1) = vbNullString _
        Else lvListView.SelectedItem.Text = vbNullString: _
        tvTreeView_NodeClick tvTreeView.SelectedItem.Parent
End Sub

'**********************************************************************************

'==================================================================================
'                 RENAME    SECTIONS/KEYS        marco.sambento@netc.pt
'==================================================================================

Private Sub tvTreeView_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Error GoTo error
    Dim nullCharPos As Integer
    Dim StringStart As Integer
    Dim Section As String
    Dim SectionStrings As String
    Dim KeyValueStr As String
    Dim Key As String
    Dim Value As String
    Dim Button As VbMsgBoxResult

    If tvTreeView.SelectedItem.Parent Is Nothing Then 'if is a section, rename it
        numSection = tvTreeView.SelectedItem.FirstSibling.Index    'first section index

    Do 'check if already exists section and prompt if so
        If LCase(tvTreeView.Nodes.Item(numSection).Text) = LCase(NewString) Then
            MsgBox "Section '" & NewString & "' already exists!", vbExclamation, "Rename Section..."
            Cancel = True
            tvTreeView.StartLabelEdit
            SendKeys NewString & "{home}+({end})"
            Exit Sub
        End If
        'if not the last section, get next
        If numSection <> tvTreeView.SelectedItem.LastSibling.Index Then _
        numSection = tvTreeView.Nodes.Item(numSection).Next.Index Else Exit Do
        
    Loop ' next key

        Button = MsgBox("Change section '" & tvTreeView.SelectedItem.Text & _
        "' with '" & NewString & "'?", vbExclamation + vbOKCancel, "Rename Section...")
        
        If Button = vbCancel Then Cancel = True: Exit Sub
        
        tvTreeView.SelectedItem.Key = NewString ' change key to newstring(sections hve keys with their names, see filltvtreeview)
        Section = tvTreeView.SelectedItem
        SectionStrings = GetSection(Section) 'retrieves keys and values from section
        Size = Len(SectionStrings)
        
        DeleteSection Section

StringStart = 1
NextString:
            nullCharPos = InStr(StringStart, SectionStrings, vbNullChar)
            'find 1st 'key=value' string and extract them
            KeyValueStr = Mid(SectionStrings, StringStart, nullCharPos - 1)
            Key = Left(KeyValueStr, InStr(KeyValueStr, "=") - 1)
            Value = Right(KeyValueStr, Len(KeyValueStr) - InStr(KeyValueStr, "="))
            'create new key
            Create NewString, Key, Value
            'adds key and value to file
            StringStart = nullCharPos + 1 ' new 'key=value' string start
            
            If nullCharPos <> Size Then GoTo NextString ' if it's not the end of
                                                'strings then get another string
    
    Else    'its a key, so rename it
            Section = tvTreeView.SelectedItem.Parent
            Key = tvTreeView.SelectedItem
            Value = GetVal(Section, Key)
        
        If tvTreeView.SelectedItem.Parent.Children > 1 Then ' if exists more than a key in section
            numKey = tvTreeView.SelectedItem.FirstSibling.Index
            
            Do    'for each key check if already exists
            If LCase(tvTreeView.Nodes.Item(numKey).Text) = LCase(NewString) Then
                    MsgBox "Key '" & NewString & "' already exists!", vbExclamation, "Rename Key..."
                    Cancel = True
                    tvTreeView.StartLabelEdit
                    SendKeys NewString & "{home}+({end})"
                    Exit Sub
            End If
                If numKey <> tvTreeView.SelectedItem.LastSibling.Index Then _
                numKey = tvTreeView.Nodes.Item(numKey).Next.Index Else Exit Do
            Loop
        End If

            Button = MsgBox("Change Key '" & Key & _
            "' with '" & NewString & "'?", vbExclamation + vbOKCancel, "Rename Key...")

            If Button = vbCancel Then Cancel = True: Exit Sub

            Create Section, NewString, Value
            DeleteKey Section, Key
    End If
Exit Sub
error:

If Err = 35603 Then MsgBox "Please choose a name with at least a letter!", _
vbCritical, "Invalid Section...": Cancel = True:: tvTreeView.StartLabelEdit: _
SendKeys NewString & "{home}+({end})": Exit Sub
'the key property doesn't suport numbers
'but it's unlikely to use only numbers to define a section
'however that is easily correct, by for example adding "S" to the key,
'it make the code confuse.

MsgBox "Error '" & Err & " - " & Err.Description & "' occured.", vbCritical, "Error"

End Sub

'==================================================================================
'          RENAME/CHANGE        KEYS/VALUES        marco.sambento@netc.pt
'==================================================================================

Private Sub lvlistView_AfterLabelEdit(Cancel As Integer, NewString As String)
Dim Section As String
Dim Key As String
Dim Value As String
Dim Button As VbMsgBoxResult
    ' just has above, but now to rename keys/values
    If lvListView.SelectedItem.Tag <> vbNullString Then 'only keys have tag, see nodeclick
        Value = lvListView.SelectedItem.SubItems(1)
        Key = lvListView.SelectedItem
        Section = tvTreeView.SelectedItem

        For numKey = 1 To lvListView.ListItems.Count
            If LCase(lvListView.ListItems.Item(numKey).Text) = LCase(NewString) Then
                MsgBox "Key '" & NewString & "' already exists!", vbExclamation, "Rename Key..."
                Cancel = True
                lvListView.StartLabelEdit
                SendKeys NewString & "{home}+({end})"
                Exit Sub
            End If
        Next numKey

        Button = MsgBox("Change Key '" & Key & _
        "' with '" & NewString & "'?", vbExclamation + vbOKCancel, "Rename Key...")
        
        If Button = vbCancel Then Cancel = True: Exit Sub
        
        Create Section, NewString, Value
        DeleteKey Section, Key
    
        tvTreeView.Nodes.Item(lvListView.SelectedItem.Tag).Text = NewString
    Else 'it's a value
        Section = tvTreeView.SelectedItem.Parent
        Key = tvTreeView.SelectedItem.Text
        Value = lvListView.SelectedItem.Text
        
        Button = MsgBox("Change value '" & Value & _
        "' with '" & NewString & "'?", vbExclamation + vbOKCancel, "Change Value...")
        
        If Button = vbCancel Then Cancel = True: Exit Sub
        
        Value = NewString
        Create Section, Key, Value
    End If
End Sub

'**********************************************************************************

Private Sub tbtoolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "WordPad"
        mnuEditEditWordPad_Click
        
        Case "Close"
        mnuFileClose_Click
        
        Case "New"
        mnuFileNew_Click
        
        Case "Exit"
        mnuFileExit_Click
        
        Case "Win.ini"
        
        INIPath = WinDir & "\win.ini"
        EnableDisable
        FilltvTreeView
        
        Case "System.ini"

        INIPath = WinDir & "\system.ini"
        EnableDisable
        FilltvTreeView

        Case "View Gridlines"
        mnuViewGridlines_Click
            
        Case "Find"
        mnuEditFind_Click
        
        Case "Open INI"
        mnuFileOpen_Click
        
        Case "Left"
        tvTreeView.SetFocus
        SendKeys "{up}"
                        
        Case "Right"
        tvTreeView.SetFocus
        If tvTreeView.SelectedItem.Parent Is Nothing Then
            If tvTreeView.SelectedItem.Children = 0 Then SendKeys "{down}"
            SendKeys "{right}"
        Else
            SendKeys "{down}"
        End If
        
        Case "Delete"
        mnuEditDelete_Click
        
        Case "Refresh"
        mnuViewRefresh_Click
        
        Case "Rename"
        mnuEditRename_Click
                
        Case "about"
        mnuAbout_Click
        
    End Select
End Sub

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Not FileOpen Or FileIsEmpty Then Exit Sub

    If indrag Then If Shift = 2 Then CopyNode = True: _
    tvTreeView.DragIcon = ilSmall.ListImages.Item("Copy").Picture
End Sub

Public Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If Not FileOpen Or FileIsEmpty Then Exit Sub

    Select Case KeyCode
        Case Is = 93
            PopupMenu mnupop
        Case vbKeyF1
            mnuAbout_Click
    End Select

        CopyNode = False
        tvTreeView.DragIcon = ilSmall.ListImages.Item("Move").Picture
End Sub

Private Sub lvListView_KeyUp(KeyCode As Integer, Shift As Integer)
If Not FileOpen Or FileIsEmpty Then Exit Sub
        
    Select Case KeyCode
        Case vbKeyBack
            If Not (tvTreeView.SelectedItem.Parent Is Nothing) Then
                tvTreeView.SelectedItem = tvTreeView.SelectedItem.Parent
                tvTreeView_NodeClick tvTreeView.SelectedItem
            End If
        Case vbKeyReturn
            lvlistView_DblClick
    End Select
End Sub

Private Sub tvTreeView_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Not FileOpen Then Exit Sub
    If Button = vbRightButton Then
        If FileIsEmpty Then
            PopupMenu mnuCreate
            Exit Sub
        Else
            If Not (tvTreeView.SelectedItem Is Nothing) Then _
            mnuRename.Enabled = True: PopupMenu mnupop
        End If
    End If
End Sub

Private Sub lvlistView_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Not FileOpen Then Exit Sub
    If Button = vbRightButton Then
        If FileIsEmpty Then
            PopupMenu mnuCreate
        Else
            If lvListView.SelectedItem Is Nothing Then _
            mnuDeleteValue.Enabled = False: mnuDeleteKey.Enabled = False: _
            mnuRename.Enabled = False
            PopupMenu mnupop
        End If
    End If
End Sub

Private Sub lvlistView_DblClick()
    If Not FileOpen Or FileIsEmpty Then Exit Sub
    If lvListView.SelectedItem Is Nothing Then Exit Sub
    If lvListView.SelectedItem.Tag <> vbNullString Then
        tvTreeView.SelectedItem = tvTreeView.Nodes.Item(lvListView.SelectedItem.Tag)

        tvTreeView_NodeClick tvTreeView.SelectedItem
    End If
End Sub

'**********************************************************************************
'
'==================================================================================
'                 DragDrop OPERATION               marco.sambento@netc.pt
'==================================================================================
'
Private Sub tvtreeview_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Not FileOpen Then Exit Sub

If FileIsEmpty Then
    If Button = vbRightButton Then PopupMenu mnuCreate
    Exit Sub
End If

    Set nodDrag = tvTreeView.SelectedItem ' Set the item being dragged.
End Sub

Public Sub lvListView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Not FileOpen Then Exit Sub

If FileIsEmpty Then
    If Button = vbRightButton Then PopupMenu mnuCreate
    Exit Sub
End If

If lvListView.ListItems.Count = 0 Then Exit Sub

    If lvListView.SelectedItem.Tag <> vbNullString Then _
    Set nodDrag = tvTreeView.Nodes.Item(lvListView.SelectedItem.Tag)

End Sub

Private Sub tvtreeview_MouseMove _
(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next

    If Button = vbLeftButton Then ' Signal a Drag operation.
        If nodDrag.Parent Is Nothing Then
        Else
            indrag = True   ' Set the flag to true.
            tvTreeView.DragIcon = ilSmall.ListImages.Item("Move").Picture 'dragicon
            tvTreeView.Drag vbBeginDrag ' Drag operation.
        End If
    End If
End Sub

Private Sub tvtreeview_DragDrop(Source As Control, x As Single, y As Single)

        If tvTreeView.DropHighlight Is Nothing Then GoTo clear
        If tvTreeView.DropHighlight = nodDrag.Parent Or _
        tvTreeView.DropHighlight = nodDrag Then GoTo clear
        
        If tvTreeView.DropHighlight.Parent Is Nothing Then _
        Set newSection = tvTreeView.DropHighlight Else _
        Set newSection = tvTreeView.DropHighlight.Parent

        'check for existing key
        If newSection.Children > 0 Then ' if exists keys in section
            numKey = newSection.Child.FirstSibling.Index
nextkey:
        'for each key child in dropped section, checks is already exists dropped key
                If LCase(tvTreeView.Nodes.Item(numKey).Text) = LCase(nodDrag.Text) Then
                Button = MsgBox("Key '" & nodDrag.Text & "' already exists in '" _
                & newSection & "'!" & vbLf & "Replace?", vbExclamation + vbYesNo, "Replace Key...")
                
                If Button = vbNo Then GoTo clear
                If Button = vbYes Then tvTreeView.Nodes.Remove numKey: GoTo Replace
                End If
                
                If numKey <> newSection.Child.LastSibling.Index Then _
                numKey = tvTreeView.Nodes.Item(numKey).Next.Index: GoTo nextkey
        End If
Replace:
        Section = nodDrag.Parent.Text
        Key = nodDrag.Text
        Value = GetVal(Section, Key)
        
        Create newSection, Key, Value
        With nodDrag
        tvTreeView.Nodes.Add newSection, tvwChild, , .Text, .Image, .SelectedImage
        End With

        If Not CopyNode Then 'if copyKey flag is set to false then delete dragged key
        DeleteKey Section, Key
        tvTreeView.Nodes.Remove nodDrag.Index
        End If
clear:
        Set tvTreeView.DropHighlight = Nothing
        indrag = False
        CopyNode = False
        tvTreeView.DragIcon = ilSmall.ListImages.Item("Move").Picture
        tvTreeView_NodeClick tvTreeView.SelectedItem
End Sub

Private Sub tvtreeview_DragOver(Source As Control, x As Single, y As Single, State As Integer)
On Error Resume Next

    If indrag = True Then
        If tvTreeView.HitTest(x, y).Parent Is Nothing Then
            ' Set DropHighlight to the mouse's coordinates.
            Set tvTreeView.DropHighlight = tvTreeView.HitTest(x, y)
        Else
            Set tvTreeView.DropHighlight = tvTreeView.HitTest(x, y).Parent
        End If
    End If
End Sub

'**********************************************************************************'

Private Sub mnuCollapse_Click()
    
    If tvTreeView.SelectedItem.Parent Is Nothing Then
        tvTreeView.SelectedItem.Expanded = False
    Else
        tvTreeView.SelectedItem.Parent.Expanded = False
    End If
End Sub

Private Sub mnuCollapseAll_Click()
    For i = 1 To tvTreeView.Nodes.Count
        tvTreeView.Nodes(i).Expanded = False
    Next i
End Sub

Private Sub mnuExpande_Click()
    tvTreeView.SelectedItem.Expanded = True
End Sub

Private Sub mnuExpandeAll_Click()
    For i = 1 To tvTreeView.Nodes.Count
        tvTreeView.Nodes(i).Expanded = True
    Next i
End Sub

Private Sub mnuRefresh_Click()
    mnuViewRefresh_Click
End Sub
Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileNew_Click()
    On Error GoTo error
    Dim Section As Node
    Dim Key As Node
    
    With CommonDialog1
        .CancelError = True
        .DialogTitle = "Create new INI File..."
        .InitDir = WinDir
        .ShowSave
        
        If .Filename <> vbNullString Then INIPath = .Filename
    End With
    
    Create "New Section", "New Key", "New Value"
    tvTreeView.Nodes.clear
    lvListView.ListItems.clear
    
    Set Section = tvTreeView.Nodes.Add(, , "New Section", "New Section", "Section", "selSection")
    tvTreeView.Nodes.Add Section, tvwChild, , "New Key", "Key", "selKey"

Me.Caption = App.Title & " - '" & StrConv(INIPath, vbProperCase) & "'"
Call CheckFile

numSections = 1
tvTreeView.SelectedItem = Section
tvTreeView_NodeClick tvTreeView.SelectedItem

Unload frmFind
error:
If Err = 32755 Then Exit Sub 'cancel was pressed

End Sub

Private Sub mnuFileOpen_Click()
    On Error GoTo error
    With CommonDialog1
        .CancelError = True
        .Flags = cdlOFNFileMustExist
        .InitDir = WinDir
        .ShowOpen
    End With

        If CommonDialog1.Filename <> vbNullString Then
        INIPath = CommonDialog1.Filename
        EnableDisable

        Call FilltvTreeView  ' fill tree and listview
        End If

error:
If Err = 32755 Then Exit Sub 'cancel was pressed

End Sub

Private Sub mnuFileClose_Click()
    INIPath = vbNullString
    EnableDisable
    Unload frmFind
End Sub

Private Sub mnuViewDetails_Click()
    mnuViewDetails.Checked = True
    lvListView.View = lvwReport
    
    mnuViewIcons.Checked = False
    mnuViewList.Checked = False
    mnuViewSmallIcons.Checked = False
End Sub

Private Sub mnuViewGridlines_Click()
    If mnuViewGridlines.Checked Then
        lvListView.GridLines = False
        tbToolbar.Buttons.Item("View Gridlines").Value = tbrUnpressed
        mnuViewGridlines.Checked = False
    Else
        lvListView.GridLines = True
        tbToolbar.Buttons.Item("View Gridlines").Value = tbrPressed
        mnuViewGridlines.Checked = True
    End If
End Sub

Private Sub mnuViewIcons_Click()
    mnuViewIcons.Checked = True
    lvListView.View = lvwIcon

    mnuViewDetails.Checked = False
    mnuViewList.Checked = False
    mnuViewSmallIcons.Checked = False
End Sub

Private Sub mnuViewList_Click()
    mnuViewList.Checked = True
    lvListView.View = lvwList
    
    mnuViewIcons.Checked = False
    mnuViewDetails.Checked = False
    mnuViewSmallIcons.Checked = False
End Sub

Private Sub mnuViewSmallIcons_Click()
    mnuViewSmallIcons.Checked = True
    lvListView.View = lvwSmallIcon
    
    mnuViewIcons.Checked = False
    mnuViewDetails.Checked = False
    mnuViewList.Checked = False
End Sub

Private Sub mnuViewStatusbar_Click()
    If mnuViewStatusbar.Checked Then
        sbStatusBar.Visible = False
        mnuViewStatusbar.Checked = False
    Else
        sbStatusBar.Visible = True
        mnuViewStatusbar.Checked = True
    End If
    
    SizeControls imgSplitter.Left
End Sub

Private Sub mnuViewToolbar_Click()
    If mnuViewToolbar.Checked = True Then
        tbToolbar.Visible = False
        mnuViewToolbar.Checked = False
    Else
        tbToolbar.Visible = True
        mnuViewToolbar.Checked = True
    End If
    
    SizeControls imgSplitter.Left
End Sub

Public Sub mnuEditRename_Click()
    ActiveControl.StartLabelEdit
End Sub

Private Sub mnuRename_Click()
    mnuEditRename_Click
End Sub

Private Sub mnuEditEditWordPad_Click()
    OpenWordPad
End Sub

Private Sub mnuEditDelete_Click()
    PopupMenu mnuDelete
End Sub

Private Sub mnuEditFind_Click()
    frmFind.Show
End Sub

Private Sub mnuViewRefresh_Click()
    FilltvTreeView
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Function EnableDisable()

If Not FileOpen Then

        tvTreeView.Nodes.clear
        lvListView.ListItems.clear
        
        For Button = 1 To tbToolbar.Buttons.Count
            tbToolbar.Buttons.Item(Button).Enabled = False
        Next Button
        
            tbToolbar.Buttons.Item("New").Enabled = True
            tbToolbar.Buttons.Item("Open INI").Enabled = True
            tbToolbar.Buttons.Item("View Gridlines").Enabled = True
            tbToolbar.Buttons.Item("about").Enabled = True
            tbToolbar.Buttons.Item("Win.ini").Enabled = True
            tbToolbar.Buttons.Item("System.ini").Enabled = True
            tbToolbar.Buttons.Item("Exit").Enabled = True

            UpdateStatusBar
        
            mnuEdit.Enabled = False
            mnuViewRefresh.Enabled = False
            mnuFileClose.Enabled = False
                       
            Me.Caption = StrConv(App.Title, vbProperCase)
Else
            For Button = 1 To tbToolbar.Buttons.Count
                tbToolbar.Buttons.Item(Button).Enabled = True
            Next Button
            
            mnuEdit.Enabled = True
            mnuViewRefresh.Enabled = True
            mnuFileClose.Enabled = True
End If
End Function

Function CheckFile()

    If tvTreeView.Nodes.Count = 0 Then  ' if tree hasn't nodes
        FileIsEmpty = True
        mnuEditDelete.Enabled = False
        mnuEditRename.Enabled = False
        mnuEditFind.Enabled = False
        
        For Button = 1 To tbToolbar.Buttons.Count
            tbToolbar.Buttons.Item(Button).Enabled = False
        Next Button
        
            tbToolbar.Buttons.Item("New").Enabled = True
            tbToolbar.Buttons.Item("Open INI").Enabled = True
            tbToolbar.Buttons.Item("Close").Enabled = True
            tbToolbar.Buttons.Item("View Gridlines").Enabled = True
            tbToolbar.Buttons.Item("about").Enabled = True
            tbToolbar.Buttons.Item("Win.ini").Enabled = True
            tbToolbar.Buttons.Item("System.ini").Enabled = True
            tbToolbar.Buttons.Item("Exit").Enabled = True
            tbToolbar.Buttons.Item("WordPad").Enabled = True
            tbToolbar.Buttons.Item("Refresh").Enabled = True
        
        sbStatusBar.Panels.Item(1).Text = vbNullString
        sbStatusBar.Panels.Item(2).Text = vbNullString
    Else
        FileIsEmpty = False
        
        For Button = 1 To tbToolbar.Buttons.Count
            tbToolbar.Buttons.Item(Button).Enabled = True
        Next Button
        
        mnuEditDelete.Enabled = True
        mnuEditRename.Enabled = True
        mnuEditFind.Enabled = True
    End If

End Function

Public Function UpdateStatusBar()
    If Not FileIsEmpty And FileOpen Then
    
    Set Node = tvTreeView.SelectedItem
        If Node.Parent Is Nothing Then
            sbStatusBar.Panels.Item(1).Text = numSections & " Sections - " & tvTreeView.Nodes.Count - numSections & " Keys" 'for sections
            sbStatusBar.Panels.Item(2).Text = lvListView.ListItems.Count & " keys/values in section '" & Node & "'"
        Else
            sbStatusBar.Panels.Item(1).Text = Node.Parent.Children & " Keys in '" & Node.Parent & "'" 'keys
            sbStatusBar.Panels.Item(2).Text = "Value '" & lvListView.SelectedItem & "' in key '" & Node & "'"
        End If
    Else
        sbStatusBar.Panels.Item(1).Text = vbNullString
        sbStatusBar.Panels.Item(2).Text = vbNullString
    End If
End Function

'**********************************************************************************

'==================================================================================
'              RESIZE CONTROLS OPERATION           marco.sambento@netc.pt
'==================================================================================

Private Sub Form_Resize()

    If Me.WindowState = vbNormal Then
        If Me.Width < 10000 Then Me.Width = 10000
        Me.Move (Screen.Width - Me.Width) / 2, _
        (Screen.Height - Me.Height) / 2, Me.Width, Me.Height
    End If
    SizeControls imgSplitter.Left
End Sub

Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 4, .Height - 20
    End With
    picSplitter.Visible = True
    mbMoving = True
End Sub

Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sglPos As Single
    

    If mbMoving Then
        sglPos = x + imgSplitter.Left
        If sglPos < sglSplitLimit Then
            picSplitter.Left = sglSplitLimit
        ElseIf sglPos > Me.Width - sglSplitLimit * 3 Then
            picSplitter.Left = Me.Width - sglSplitLimit * 3
        Else
            picSplitter.Left = sglPos
        End If
    End If
End Sub

Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    SizeControls picSplitter.Left
    picSplitter.Visible = False
    mbMoving = False
End Sub

Sub SizeControls(x As Single)
    
    Dim numColumns As Integer
    Dim Column As Integer
    Dim sizeColumn As Integer

    'set the width
    If x < 1500 Then x = 1500
    If x > (Me.Width - 1500) Then x = Me.Width - 1500
    tvTreeView.Width = x - imgSplitter.Width
    sbStatusBar.Panels.Item(1).Width = x
    
    imgSplitter.Left = x
    lvListView.Left = x + imgSplitter.Width / 2
    lvListView.Width = Me.ScaleWidth - lvListView.Left - tvTreeView.Left
    
    numColumns = lvListView.ColumnHeaders.Count
    sizeColumn = (lvListView.Width / numColumns) - 180

    For Column = 1 To numColumns
        lvListView.ColumnHeaders.Item(Column).Width = sizeColumn
    Next Column

    'set the top

    If tbToolbar.Visible Then
        tvTreeView.Top = tbToolbar.Height + 120
    Else
        tvTreeView.Top = 120
    End If
    
  lvListView.Top = tvTreeView.Top
    

    'set the height
    If sbStatusBar.Visible Then
        tvTreeView.Height = Me.ScaleHeight - tvTreeView.Top - sbStatusBar.Height
    Else
        tvTreeView.Height = Me.ScaleHeight - tvTreeView.Top - 80
    End If
    
    lvListView.Height = tvTreeView.Height
    imgSplitter.Top = tvTreeView.Top
    imgSplitter.Height = tvTreeView.Height
End Sub

'**********************************************************************************
