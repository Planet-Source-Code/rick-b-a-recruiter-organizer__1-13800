VERSION 5.00
Begin VB.Form frmSearch 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Search  copyright@2002 by Indra"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   11
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search Now"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Frame frOption 
      Caption         =   "Select what to search"
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   5775
      Begin VB.ComboBox cboFollowUp 
         Height          =   315
         Left            =   1800
         TabIndex        =   8
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtNameSearch 
         Height          =   375
         Left            =   1800
         TabIndex        =   7
         Top             =   1800
         Width           =   3735
      End
      Begin VB.CheckBox chkName 
         Caption         =   "Recruiter Name"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox txtStateSearch 
         Height          =   405
         Left            =   1800
         TabIndex        =   5
         Top             =   1080
         Width           =   3735
      End
      Begin VB.TextBox txtTypeSearch 
         Height          =   405
         Left            =   1800
         TabIndex        =   3
         Top             =   360
         Width           =   3735
      End
      Begin VB.CheckBox chkState 
         Caption         =   "State"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   855
      End
      Begin VB.CheckBox chkType 
         Caption         =   "Type of Company"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Follow-Up Only"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2520
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Choose the field(s) you want to search for:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Icon = frmRecruiter.imlTree.ListImages(8).Picture
    cboFollowUp.AddItem ""
    cboFollowUp.AddItem "Yes"
    cboFollowUp.AddItem "No"
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdReset_Click()
    txtTypeSearch.Text = ""
    txtStateSearch.Text = ""
    txtNameSearch.Text = ""
    cboFollowUp.Text = ""
    chkState.Value = vbUnchecked
    chkType.Value = vbUnchecked
    chkName.Value = vbUnchecked
End Sub


Private Sub cmdSearch_Click()
    Dim typeString As String
    Dim stateString As String
    Dim firstNameString As String
    Dim lastNameString As String
    Dim followUpString As String
    Dim flag As Boolean
    
    flag = True 'true if nothing to search
    typeString = ""
    stateString = ""
    firstNameString = ""
    lastNameString = ""
    followUpString = ""
    
    'check for type search
    If chkType.Value = vbChecked Then
        If txtTypeSearch.Text = "" Then
            MsgBox "You can check 'TYPE' checkbox while the textbox is empty"
        Else
            typeString = txtTypeSearch.Text
            flag = False
        End If
    End If
    ' check for state search
    If chkState.Value = vbChecked Then
        If txtStateSearch.Text = "" Then
            MsgBox "You cant check 'STATE' checkbox while the textbox is empty"
        Else
            stateString = txtStateSearch.Text
            flag = False
        End If
    End If
    ' check for recruiter name search
    If chkName.Value = vbChecked Then
        If txtNameSearch.Text = "" Then
            MsgBox "You cant check 'RECRUITER NAME' checkbox while the textbox is empty"
        Else
            Dim temp As String
            temp = txtNameSearch.Text
            If InStr(temp, " ") <> 0 Then
                firstNameString = Left(temp, InStr(temp, " ") - 1)
                lastNameString = Mid(temp, InStr(temp, " ") + 1)
            Else
                firstNameString = temp
                lastNameString = temp
            End If
            flag = False
        End If
    End If
    'check for follow-up only
    If cboFollowUp.Text <> "" Then
        followUpString = cboFollowUp.Text
        flag = False
    End If
    
    'check for error, ie: empty all search fields
    If flag Then
        MsgBox "Error, action aborted! Nothing to search", vbCritical
        Exit Sub
    End If
    
    Call frmSearchResult.displayResult(Trim(typeString), Trim(stateString), Trim(firstNameString), Trim(lastNameString), Trim(followUpString))
End Sub

