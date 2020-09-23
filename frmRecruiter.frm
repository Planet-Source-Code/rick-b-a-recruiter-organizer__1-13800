VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmRecruiter 
   Caption         =   "Recruiter Organizer v1.0  copyright@2002 by Indra"
   ClientHeight    =   7950
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10965
   Icon            =   "frmRecruiter.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   10965
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imlTree 
      Left            =   2880
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecruiter.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecruiter.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecruiter.frx":0CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecruiter.frx":1138
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecruiter.frx":158A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecruiter.frx":19DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecruiter.frx":1E2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRecruiter.frx":2148
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picTitles 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   75
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   10965
      TabIndex        =   5
      Top             =   630
      Width           =   10965
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   6360
      Left            =   10800
      ScaleHeight     =   2769.418
      ScaleMode       =   0  'User
      ScaleWidth      =   780
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   75
   End
   Begin MSComctlLib.ListView lvListView 
      Height          =   6375
      Left            =   1320
      TabIndex        =   1
      Top             =   720
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   11245
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView tvTreeView 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   11245
      _Version        =   393217
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imlTree"
      Appearance      =   1
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   1111
      ButtonWidth     =   1270
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "imlTree"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Back"
            Key             =   "Back"
            Object.ToolTipText     =   "Previous contact"
            Object.Tag             =   "Back"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Forward"
            Key             =   "Forward"
            Object.ToolTipText     =   "Next contact"
            Object.Tag             =   "Forward"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete contact"
            Object.Tag             =   "Delete"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Insert"
            Key             =   "Insert"
            Object.ToolTipText     =   "Insert new contact"
            Object.Tag             =   "Insert"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            Key             =   "Search"
            Object.ToolTipText     =   "Search contact"
            Object.Tag             =   "Search"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Info"
            Key             =   "Info"
            Object.ToolTipText     =   "Info about selected entry"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Email Me"
            Key             =   "Email"
            Object.ToolTipText     =   "Email"
            Object.Tag             =   "Email"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit"
            Object.Tag             =   "Exit"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   7695
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13679
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "1/10/2002"
            Object.ToolTipText     =   "Date"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "6:45 PM"
            Object.ToolTipText     =   "Time"
         EndProperty
      EndProperty
   End
   Begin VB.Image imgSplitter 
      Height          =   6225
      Left            =   1200
      MousePointer    =   9  'Size W E
      Top             =   720
      Width           =   135
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuPopUpTree 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuSearch 
         Caption         =   "Search Contact List"
      End
      Begin VB.Menu mnuInsert 
         Caption         =   "Insert New Contact"
      End
   End
   Begin VB.Menu mnuPopUpList 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Contact"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit Contact"
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "Properties"
      End
   End
End
Attribute VB_Name = "frmRecruiter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim mbMoving As Boolean
Const sglSplitLimit = 500

Private Sub Form_Load()
    'set the database file path
    dbPath = App.Path & "\RecruiterOrganizer.mdb"
    If (Not openDatabase) Then
        MsgBox "ERROR, cant open/create database file", vbCritical, "Error Open/Create DB"
    End If
    
    'open the database
    Set dbRecruiter = DBEngine.Workspaces(0).openDatabase(dbPath, , False)
    'open the recordset
    Set rsRecruiter = dbRecruiter.OpenRecordset("tblRecruiter", dbOpenDynaset)

    'to make column in list view
    Call makeColumns1
    
    'put form's icon
    Me.Icon = imlTree.ListImages(8).Picture
    
    Dim lIndex As Long
    Dim nodeX As Node
    'create root for number
    Set nodeX = tvTreeView.Nodes.Add(, 1, "0-9", "0 - 9", 1, 2)
    nodeX.key = "0-9"
    nodeX.Tag = "Root"
    nodeX.Sorted = True
    'create all root nodes from A to Z
    For lIndex = Asc("A") To Asc("Z")
        Set nodeX = tvTreeView.Nodes.Add(, 1, Chr$(lIndex), Chr$(lIndex), 1, 2)
        nodeX.key = Chr$(lIndex)
        nodeX.Tag = "Root"
        nodeX.Sorted = True
    Next lIndex
    
    tvTreeView.Nodes(1).Selected = True
    Call populateListView(tvTreeView.SelectedItem.key) 'display first company list
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set dbRecruiter = Nothing
    Set rsRecruiter = Nothing
End Sub


' for sorting purposes in listview
Private Sub lvListView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If lvListView.SortKey <> ColumnHeader.Index - 1 Then
        lvListView.SortKey = ColumnHeader.Index - 1
    Else
        If lvListView.SortOrder <> lvwAscending Then
            lvListView.SortOrder = lvwAscending
        Else
            lvListView.SortOrder = lvwDescending
        End If
    End If
End Sub

Private Sub lvListView_DblClick()
    Dim listitemX As ListItem
    If (lvListView.ListItems.count > 0) Then
        Set listitemX = lvListView.SelectedItem
        companyID = listitemX.key
        companyID = Left(companyID, InStr(companyID, "_") - 1)
        Call frmInfo.info
        frmInfo.Show vbModal
    End If
End Sub


Private Sub tvTreeView_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Tag = "Root" Then
        Call populateListView(Node.key)
    Else
        lvListView.ListItems.Clear
    End If
End Sub


'------------------------------------ Got Focus ------------------------------------'
Private Sub lvListView_GotFocus()
    If lvListView.ListItems.count > 0 Then
        ctlGotFocus = lvGotFocus
    End If
End Sub

Private Sub tvTreeView_GotFocus()
    If tvTreeView.Nodes.count > 1 Then
        ctlGotFocus = tvGotFocus
    End If
End Sub
'------------------------------------ Got Focus ------------------------------------'



'------------------------------------ Menu Toolbar -----------------------------------'
Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.key
        'Back Button
        Case "Back"
            Call back
        'Forward Button
        Case "Forward"
            Call forward
        'Delete Button
        Case "Delete"
            'listview got focus
            If ctlGotFocus = lvGotFocus Then
                If (lvListView.ListItems.count > 0) Then
                    Dim listitemX As ListItem
                    Set listitemX = lvListView.SelectedItem
                    Call delete(listitemX)
                End If
            Else
                MsgBox "Select an entry to be deleted in the right box"
            End If
        'Add Address Book Button
        Case "Insert"
            isEditMode = False
            frmNewEditContact.Show vbModal
        Case "Search"
            frmSearch.Show vbModal
        Case "Email"
            Call email("ihonggoseput@yahoo.com")
        Case "Info"
            'listview got focus
            If ctlGotFocus = lvGotFocus Then
                If (lvListView.ListItems.count > 0) Then
                    Dim listitemX2 As ListItem
                    Set listitemX2 = lvListView.SelectedItem
                    companyID = listitemX2.key
                    companyID = Left(companyID, InStr(companyID, "_") - 1)
                    Call frmInfo.info
                    frmInfo.Show vbModal
                End If
            Else
                MsgBox "Select an entry to be displayed in the right box"
            End If
        Case "Exit"
            Dim response As String
            response = MsgBox("Are you sure?", vbYesNo, "Exit")
            If response = vbYes Then
                Unload Me
                Exit Sub
            End If
    End Select
    'refresh the listview
    Call populateListView(tvTreeView.SelectedItem.key)
End Sub
'------------------------------------ Menu Toolbar -------------------------------------'

'--------------------------------- RightClick Toolbar -------------------------------------'
Private Sub mnuInfo_Click()
    If (lvListView.ListItems.count > 0) Then
        Dim listitemX2 As ListItem
        Set listitemX2 = lvListView.SelectedItem
        companyID = listitemX2.key
        companyID = Left(companyID, InStr(companyID, "_") - 1)
        Call frmInfo.info
        frmInfo.Show vbModal
    End If
End Sub

Private Sub mnuedit_Click()
    If (lvListView.ListItems.count > 0) Then
        Dim listitemX2 As ListItem
        Set listitemX2 = lvListView.SelectedItem
        companyID = listitemX2.key
        companyID = Left(companyID, InStr(companyID, "_") - 1)
        Call frmInfo.edit
    End If
End Sub

Private Sub mnuSearch_Click()
    frmSearch.Show vbModal
End Sub

Private Sub mnuDelete_Click()
    If (lvListView.ListItems.count > 0) Then
        Dim listitemX As ListItem
        Set listitemX = lvListView.SelectedItem
        Call delete(listitemX)
    End If
End Sub

Private Sub mnuInsert_Click()
    isEditMode = False
    frmNewEditContact.Show vbModal
End Sub
'--------------------------------- RightClick Toolbar -------------------------------------'



'about menu bar click
Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

'exit menu click
Private Sub mnuExit_Click()
    Unload Me
End Sub


'++++++++++++++++++++++++++++++++++ Resize Functions +++++++++++++++++++++++++++++++++++
Private Sub Form_Resize()
    On Error Resume Next
    If Me.Width < 3000 Then Me.Width = 3000
    SizeControls imgSplitter.Left
End Sub
Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
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
        ElseIf sglPos > Me.Width - sglSplitLimit Then
            picSplitter.Left = Me.Width - sglSplitLimit
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
Private Sub TreeView1_DragDrop(Source As Control, x As Single, y As Single)
    If Source = imgSplitter Then
        SizeControls x
    End If
End Sub
Sub SizeControls(x As Single)
    On Error Resume Next
    'set the width
    If x < 1500 Then x = 1500
    If x > (Me.Width - 1500) Then x = Me.Width - 1500
    tvTreeView.Width = x
    imgSplitter.Left = x
    lvListView.Left = x + 40
    lvListView.Width = Me.Width - (tvTreeView.Width + 140)
    'set the top
    If tbToolBar.Visible Then
        tvTreeView.Top = tbToolBar.Height + picTitles.Height
    Else
        tvTreeView.Top = picTitles.Height
    End If
    lvListView.Top = tvTreeView.Top
    'set the height
    If sbStatusBar.Visible Then
        tvTreeView.Height = Me.ScaleHeight - (picTitles.Top + picTitles.Height + sbStatusBar.Height)
    Else
        tvTreeView.Height = Me.ScaleHeight - (picTitles.Top + picTitles.Height)
    End If
    lvListView.Height = tvTreeView.Height
    imgSplitter.Top = tvTreeView.Top
    imgSplitter.Height = tvTreeView.Height
End Sub
'++++++++++++++++++++++++++++++++++ Resize Functions +++++++++++++++++++++++++++++++++++






'---------------------------------- Right Click Event ------------------------------'
Private Sub tvTreeView_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        frmRecruiter.PopupMenu mnuPopUpTree
    End If
End Sub

Private Sub lvListView_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        If lvListView.ListItems.count > 0 Then
            frmRecruiter.mnuDelete.Enabled = True
            frmRecruiter.mnuEdit.Enabled = True
            frmRecruiter.PopupMenu mnuPopUpList
        End If
    End If
End Sub
'---------------------------------- Right Click Event ------------------------------'
