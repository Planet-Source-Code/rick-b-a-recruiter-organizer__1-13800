VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSearchResult 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Search Result (double click on entry to view the complete info)   copyright@2002 by Indra"
   ClientHeight    =   6030
   ClientLeft      =   165
   ClientTop       =   435
   ClientWidth     =   13530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   13530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvListView 
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   10398
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit Contact"
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "Properties"
      End
   End
End
Attribute VB_Name = "frmSearchResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Icon = frmRecruiter.imlTree.ListImages(8).Picture
End Sub


Public Sub displayResult(typeSearch As String, stateSearch As String, firstNameSearch As String, lastNameSearch As String, followUpSearch As String)
    'create a new recordset
    Dim rs As Recordset
    Dim sSQL As String
    Dim whereClause As String
    Dim flag As Boolean
    
    flag = True 'indicate that the where clause has been modified
    
    sSQL = "SELECT CompanyName, URL, Type, State, FirstName, LastName, FollowUp, Date, Time, cID " & _
            "FROM tblRecruiter "
           
    whereClause = "WHERE "
    
    'check for type search
    If typeSearch <> "" Then
        whereClause = whereClause & "Type like '" & typeSearch & "' "
        flag = False
    End If
    'check for zip search
    If stateSearch <> "" Then
        If Not flag Then
            whereClause = whereClause & "AND "
        End If
        whereClause = whereClause & "State like '" & stateSearch & "' "
        flag = False
    End If
    'check for name search
    If firstNameSearch <> "" Then
        If Not flag Then
            whereClause = whereClause & "AND "
        End If
        whereClause = whereClause & "FirstName like '" & firstNameSearch & "' "
        flag = False
    End If
    If lastNameSearch <> "" Then
        If Not flag Then
            whereClause = whereClause & "AND "
        End If
        whereClause = whereClause & "LastName like '" & lastNameSearch & "' "
        flag = False
    End If
    'check for follow-up search
    If followUpSearch <> "" Then
        If Not flag Then
            whereClause = whereClause & "AND "
        End If
        If followUpSearch = "Yes" Then
            whereClause = whereClause & "FollowUp <> 0 "
        Else
            whereClause = whereClause & "FollowUp = 0 "
        End If
    End If
        
    'concatenate the string sql
    sSQL = sSQL & whereClause
    sSQL = Trim(sSQL)
    Set rs = dbRecruiter.OpenRecordset(sSQL, dbOpenDynaset)
     
    Call makeColumns2 'create columns in listview

    lvListView.ListItems.Clear
    With rs
        If .RecordCount > 0 Then
            .MoveFirst
            'create a listitem object to be inserted in listview
            Dim itemX As ListItem
            Do While Not .EOF
                Set itemX = lvListView.ListItems.Add(, CStr(.Fields("cID")) & "_ " & _
                            CStr(.AbsolutePosition), .Fields("CompanyName"))
                'display URL
                If .Fields("URL") <> "" Then
                    itemX.SubItems(1) = .Fields("URL")
                Else
                    itemX.SubItems(1) = "No URL listed"
                End If
                'display Type
                If .Fields("Type") <> "" Then
                    itemX.SubItems(2) = .Fields("Type")
                End If
                'display Zip Code
                If .Fields("State") <> "" Then
                    itemX.SubItems(3) = .Fields("State")
                End If
                'display Recruiter Name
                If .Fields("FirstName") <> "" Then
                    itemX.SubItems(4) = .Fields("FirstName")
                End If
                If .Fields("LastName") <> "" Then
                    itemX.SubItems(4) = itemX.SubItems(4) & " " & .Fields("LastName")
                End If
                'display Follow-Up Info
                If (.Fields("Followup")) Then
                    itemX.SubItems(5) = "Yes"
                    itemX.SubItems(6) = .Fields("Date")
                    itemX.SubItems(7) = .Fields("Time")
                Else
                    itemX.SubItems(5) = "No"
                    itemX.SubItems(6) = ""
                    itemX.SubItems(7) = ""
                End If
                .MoveNext
            Loop
            'display the search result form
            frmSearchResult.Show vbModal
        Else
            MsgBox "Found nothing"
        End If
    End With
    
    'delete recordset
    Set rs = Nothing
End Sub


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


Private Sub lvListView_ItemClick(ByVal Item As MSComctlLib.ListItem)
    lvListView.ToolTipText = "Double Click on an entry to view the complete info"
End Sub



'---------------------------------- Right Click Event ------------------------------'
Private Sub lvListView_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        If lvListView.ListItems.count > 0 Then
            frmSearchResult.mnuEdit.Enabled = True
            frmSearchResult.mnuInfo.Enabled = True
            frmRecruiter.PopupMenu mnuPopUp
        End If
    End If
End Sub

Private Sub mnuInfo_Click()
    If (lvListView.ListItems.count > 0) Then
        Dim listitemX As ListItem
        Set listitemX = lvListView.SelectedItem
        companyID = listitemX.key
        companyID = Left(companyID, InStr(companyID, "_") - 1)
        Call frmInfo.info
        frmInfo.Show vbModal
    End If
End Sub

Private Sub mnuedit_Click()
    If (lvListView.ListItems.count > 0) Then
        Dim listitemX As ListItem
        Set listitemX = lvListView.SelectedItem
        companyID = listitemX.key
        companyID = Left(companyID, InStr(companyID, "_") - 1)
        Call frmInfo.edit
    End If
End Sub
'---------------------------------- Right Click Event ------------------------------'


