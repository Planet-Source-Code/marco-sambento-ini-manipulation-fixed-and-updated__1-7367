VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9015
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   1215
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChKCase 
      Caption         =   "&Match Case"
      Height          =   255
      Left            =   184
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox txtFind 
      Height          =   405
      Left            =   183
      TabIndex        =   0
      Top             =   240
      Width           =   8520
   End
   Begin MSComctlLib.ListView lvFound 
      Height          =   4695
      Left            =   191
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   8281
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Section"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Key"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Value"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label lblFound 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   8658
      TabIndex        =   3
      Top             =   720
      Width           =   45
   End
   Begin VB.Menu mnuPop 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuDeleteSection 
         Caption         =   "Delete &Section"
      End
      Begin VB.Menu mnuDeleteKey 
         Caption         =   "Delete &key"
      End
      Begin VB.Menu mnuDeleteValue 
         Caption         =   "Delete &Value"
      End
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FindIsLoaded As Boolean ' i don't know how to check if a form is loaded, so this is my solution
Dim Resized As Boolean
Dim noTextFocus As Boolean

Private Sub ChKCase_Click()
txtFind_KeyUp 0, 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    frmMain.Form_KeyDown 0, Shift
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
    If noTextFocus Then 'used to know if you are typing, so that when
    'you press one of this keys during typing, you won't goto the find button or check...
        Select Case KeyCode
        Case vbKeyM 'if key is 'M'
            ChKCase.SetFocus 'sets focus to check mark and changes it
            If ChKCase.Value = 1 Then ChKCase.Value = 0 Else ChKCase.Value = 1
        End Select
    End If

frmMain.CopyNode = False
frmMain.tvTreeView.DragIcon = frmMain.ilSmall.ListImages.Item("Move").Picture
End Sub

Private Sub Form_Load()
    FindIsLoaded = True

    Me.Icon = frmMain.Icon
    Me.MouseIcon = frmMain.MouseIcon
    colWidth = lvFound.Width / lvFound.ColumnHeaders.Count - 100
    For col = 1 To lvFound.ColumnHeaders.Count
        lvFound.ColumnHeaderIcons = frmMain.ilSmall
        With lvFound.ColumnHeaders.Item(col)
        .Width = colWidth
        .Icon = col
        End With
    Next col
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Resized = False
    FindIsLoaded = False
End Sub

Function MatchCase(ByVal Text As String)
    If ChKCase.Value = 1 Then ' if matchcase is checked then the returned
    'string to the sub above is the text as it was written
        MatchCase = Text
    Else 'else the text is passed as lower case to compare
        MatchCase = LCase(Text)
    End If
End Function

Private Sub lvFound_ItemClick(ByVal Item As MSComctlLib.ListItem)
        'find section in tvtreeview and select it
        With frmMain.tvTreeView
            If .SelectedItem <> .Nodes.Item(Item.Text) Then
                .SelectedItem = .Nodes.Item(Item.Text)
                frmMain.tvTreeView_NodeClick .SelectedItem
            End If
        End With
    'if 3rd column has a value, if it's not empty
    If Item.SubItems(2) <> vbNullString Then 'finds in lvlistview the correspondent
        With frmMain.lvListView              'value and selects its key
            For Key = 1 To .ListItems.Count
                If .ListItems.Item(Key).Text = Item.SubItems(1) And _
                .ListItems.Item(Key).SubItems(1) = Item.SubItems(2) Then
                    .SelectedItem = .ListItems.Item(Key)
                    Exit Sub
                End If
            Next Key
        End With
        
    ElseIf Item.SubItems(1) <> vbNullString Then 'the same as above but now for keys
        With frmMain.lvListView
            For Key = 1 To .ListItems.Count
                If .ListItems.Item(Key).Text = Item.SubItems(1) Then
                    .SelectedItem = .ListItems.Item(Key)
                    Exit Sub
                End If
            Next Key
        End With
    End If
End Sub

Private Sub lvFound_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If lvFound.ListItems.Count = 0 Then Exit Sub
    If lvFound.SelectedItem.SubItems(1) <> vbNullString Then _
    frmMain.lvListView_MouseDown Button, Shift, x, y 'call dragdrop event
End Sub

Private Sub lvFound_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If lvFound.ListItems.Count = 0 Then Exit Sub
    If lvFound.SelectedItem.SubItems(2) <> vbNullString Then
        mnuDeleteKey.Enabled = True
        mnuDeleteValue.Enabled = True
    ElseIf lvFound.SelectedItem.SubItems(1) <> vbNullString Then
        mnuDeleteValue.Enabled = False
    Else
        mnuDeleteKey.Enabled = False
        mnuDeleteValue.Enabled = False
    End If
If Button = vbRightButton Then PopupMenu mnupop

End Sub

Private Sub mnuDeleteKey_Click()
    With lvFound.SelectedItem
    frmFindSection = .Text
    frmFindKey = .SubItems(1)
    End With

    frmMainSection = frmMain.tvTreeView.SelectedItem.Text
    frmMainKey = frmMain.lvListView.SelectedItem
    frmMainValue = frmMain.lvListView.SelectedItem.SubItems(1)

If frmFindSection = frmMainSection _
And frmFindKey = frmMainKey _
And frmFindValue = frmMainValue _
Then frmMain.mnuDeleteValue_Click

txtFind_KeyUp 0, 0
End Sub

Private Sub mnuDeleteSection_Click()
    frmFindSection = lvFound.SelectedItem.Text
    frmMainSection = frmMain.tvTreeView.SelectedItem
    
    If frmFindSection = frmMainSection Then frmMain.mnuDeleteSection_Click

txtFind_KeyUp 0, 0
End Sub

Private Sub mnuDeleteValue_Click()

    With lvFound.SelectedItem
    frmFindSection = .Text
    frmFindKey = .SubItems(1)
    frmFindValue = .SubItems(2)
    End With

    frmMainSection = frmMain.tvTreeView.SelectedItem.Text
    frmMainKey = frmMain.lvListView.SelectedItem.Text
    frmMainValue = frmMain.lvListView.SelectedItem.SubItems(1)

If frmFindSection = frmMainSection _
And frmFindKey = frmMainKey _
Then frmMain.mnuDeleteKey_Click

txtFind_KeyUp 0, 0
End Sub

Private Sub txtFind_GotFocus()
    txtFind.SelStart = 0
    txtFind.SelLength = Len(txtFind)
    noTextFocus = False
End Sub

Public Sub txtFind_KeyUp(KeyCode As Integer, Shift As Integer)

Dim FileData As String
Dim Section As String
Dim Key As String
Dim Value As String
Dim Position As Integer
Dim posKey As Integer
Dim numFile As Integer

If txtFind.Text = vbNullString Then lvFound.ListItems.clear: Exit Sub    'if string to search is empty
'resize form and show listview if it hasn't been done already
If Not Resized Then
frmFind.Height = frmFind.Height + lvFound.Height
lvFound.Visible = True
Resized = True
End If

lvFound.ListItems.clear
Filename = INIPath
If Len(Filename) Then
    numFile = FreeFile
    Open Filename For Input As numFile
    Do While Not EOF(numFile)     ' while it isn't the end of file, get another line
        Line Input #numFile, FileData 'line text
        
        firstChar = Left(FileData, 1)
        lastChar = Right(FileData, 1)
        
        If firstChar = "[" And lastChar = "]" Then 'if it's a section
            Section = Mid(FileData, 2, Len(FileData) - 2)
            'finds string in section
            Position = InStr(MatchCase(Section), MatchCase(txtFind))
            If Position <> 0 Then 'if found add to listview
                lvFound.ListItems.Add , , Section
            End If
        Else ' it's a key
            If Section <> vbNullString Then
                ' see also function MatchCase
                posKey = InStr(FileData, "=") 'if it's a 'key=value' string
                If posKey <> 0 Then '

            Key = Left(FileData, posKey - 1)
            Position = InStr(MatchCase(Key), MatchCase(txtFind))
        
                If Position <> 0 Then 'if string was found on key
                    Set Item = lvFound.ListItems.Add(, , Section) 'add section to 1st column
                    Item.SubItems(1) = Key 'add key to 2nd column
                End If
            
            Value = GetVal(Section, Key)
            Position = InStr(MatchCase(Value), MatchCase(txtFind))

                If Position <> 0 Then ' if string was found in value
                    Set Item = lvFound.ListItems.Add(, , Section) 'add section
                    Item.SubItems(1) = Key 'add key
                    Item.SubItems(2) = Value 'add value
                End If
                End If
            End If
        End If
    Loop
    Close numFile 'close file
End If
If lvFound.ListItems.Count = 0 Then
    lblFound.Caption = "No Items"
Else
    lblFound.Caption = lvFound.ListItems.Count & " Items found"
End If
End Sub

Private Sub txtFind_LostFocus()
    noTextFocus = True
End Sub
