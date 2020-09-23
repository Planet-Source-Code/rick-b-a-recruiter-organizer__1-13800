VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmNewEditContact 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "New/Edit Form  copyright@2002 by Indra"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   9975
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "General Info"
      TabPicture(0)   =   "frmNewEditContact.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frCompany"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frRecruiter"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdNext1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdReset1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Follow Up Info"
      TabPicture(1)   =   "frmNewEditContact.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblFollowUp"
      Tab(1).Control(1)=   "Line1"
      Tab(1).Control(2)=   "lblDate"
      Tab(1).Control(3)=   "lblTime"
      Tab(1).Control(4)=   "optYes"
      Tab(1).Control(5)=   "optNo"
      Tab(1).Control(6)=   "mvDate"
      Tab(1).Control(7)=   "DTPicker1"
      Tab(1).Control(8)=   "cmdReset2"
      Tab(1).Control(9)=   "cmdNext2"
      Tab(1).Control(10)=   "cmdBack2"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Notes"
      TabPicture(2)   =   "frmNewEditContact.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtNotes"
      Tab(2).Control(1)=   "cmdReset3"
      Tab(2).Control(2)=   "cmdSave"
      Tab(2).Control(3)=   "cmdCancel"
      Tab(2).Control(4)=   "cmdBack3"
      Tab(2).ControlCount=   5
      Begin VB.CommandButton cmdBack2 
         Caption         =   "Previous Page"
         Height          =   495
         Left            =   -74640
         Picture         =   "frmNewEditContact.frx":0054
         TabIndex        =   22
         Top             =   4920
         Width           =   1215
      End
      Begin VB.CommandButton cmdBack3 
         Caption         =   "Previous Page"
         Height          =   495
         Left            =   -74640
         Picture         =   "frmNewEditContact.frx":0496
         TabIndex        =   25
         Top             =   4920
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   -68040
         TabIndex        =   28
         Top             =   4920
         Width           =   1335
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   495
         Left            =   -70320
         TabIndex        =   27
         Top             =   4920
         Width           =   1575
      End
      Begin VB.CommandButton cmdReset3 
         Caption         =   "Reset Notes"
         Height          =   495
         Left            =   -72720
         TabIndex        =   26
         Top             =   4920
         Width           =   1575
      End
      Begin VB.CommandButton cmdReset1 
         Caption         =   "Reset Textbox"
         Height          =   495
         Left            =   1680
         TabIndex        =   15
         Top             =   4920
         Width           =   1575
      End
      Begin VB.CommandButton cmdNext1 
         Caption         =   "Next Page"
         Height          =   495
         Left            =   7080
         Picture         =   "frmNewEditContact.frx":08D8
         TabIndex        =   16
         Top             =   4920
         Width           =   1215
      End
      Begin VB.CommandButton cmdNext2 
         Caption         =   "Next Page"
         Height          =   495
         Left            =   -67920
         Picture         =   "frmNewEditContact.frx":0D1A
         TabIndex        =   23
         Top             =   4920
         Width           =   1215
      End
      Begin VB.CommandButton cmdReset2 
         Caption         =   "Reset"
         Height          =   495
         Left            =   -67920
         TabIndex        =   21
         Top             =   2280
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "h:nn AM/PM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   4
         EndProperty
         Height          =   495
         Left            =   -71040
         TabIndex        =   20
         ToolTipText     =   "Select your follow-up time"
         Top             =   2280
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   19660802
         UpDown          =   -1  'True
         CurrentDate     =   37265
         MinDate         =   36526
      End
      Begin MSComCtl2.MonthView mvDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "M/d/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   2370
         Left            =   -74640
         TabIndex        =   19
         ToolTipText     =   "Select your follow-up date"
         Top             =   2280
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         Enabled         =   0   'False
         ShowToday       =   0   'False
         StartOfWeek     =   19660801
         CurrentDate     =   37265
         MinDate         =   36892
      End
      Begin VB.OptionButton optNo 
         Caption         =   "No"
         Height          =   255
         Left            =   -70920
         TabIndex        =   18
         Top             =   840
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton optYes 
         Caption         =   "Yes"
         Height          =   255
         Left            =   -72000
         TabIndex        =   17
         Top             =   840
         Width           =   735
      End
      Begin VB.Frame frRecruiter 
         Caption         =   "Recruiter Info"
         Height          =   1815
         Left            =   120
         TabIndex        =   38
         Top             =   3000
         Width           =   8535
         Begin VB.TextBox txtExtension 
            Height          =   375
            Left            =   6360
            TabIndex        =   14
            Top             =   1200
            Width           =   2055
         End
         Begin VB.TextBox txtPhone 
            Height          =   375
            Left            =   6360
            TabIndex        =   13
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox txtEmail 
            Height          =   375
            Left            =   6360
            TabIndex        =   12
            Top             =   240
            Width           =   2055
         End
         Begin VB.ComboBox cboTitle 
            DataSource      =   """Mr."", ""Mrs."", ""Miss"""
            Height          =   315
            ItemData        =   "frmNewEditContact.frx":115C
            Left            =   1200
            List            =   "frmNewEditContact.frx":115E
            TabIndex        =   9
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtFirst 
            Height          =   405
            Left            =   1200
            TabIndex        =   10
            Top             =   720
            Width           =   3735
         End
         Begin VB.TextBox txtLast 
            Height          =   405
            Left            =   1200
            TabIndex        =   11
            Top             =   1200
            Width           =   3735
         End
         Begin VB.Label Label2 
            Caption         =   "Extension"
            Height          =   375
            Left            =   5040
            TabIndex        =   44
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Phone Number"
            Height          =   255
            Left            =   5040
            TabIndex        =   43
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label lblEmail 
            Caption         =   "Email"
            Height          =   255
            Left            =   5040
            TabIndex        =   42
            Top             =   360
            Width           =   615
         End
         Begin VB.Label lblTitle 
            Caption         =   "Title"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   360
            Width           =   495
         End
         Begin VB.Label lblFirst 
            Caption         =   "First Name:"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   840
            Width           =   975
         End
         Begin VB.Label lblLast 
            Caption         =   "Last Name:"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   1320
            Width           =   975
         End
      End
      Begin VB.Frame frCompany 
         Caption         =   "Company Info"
         Height          =   2295
         Left            =   120
         TabIndex        =   29
         Top             =   480
         Width           =   8535
         Begin VB.TextBox txtCompany 
            Height          =   375
            Left            =   1080
            TabIndex        =   1
            Top             =   240
            Width           =   3855
         End
         Begin VB.TextBox txtStreet2 
            Height          =   375
            Left            =   6360
            TabIndex        =   5
            Top             =   240
            Width           =   2055
         End
         Begin VB.TextBox txtURL 
            Height          =   405
            Left            =   1080
            TabIndex        =   2
            Top             =   720
            Width           =   3855
         End
         Begin VB.TextBox txtType 
            Height          =   405
            Left            =   1080
            TabIndex        =   3
            Top             =   1200
            Width           =   3855
         End
         Begin VB.TextBox txtStreet1 
            Height          =   375
            Left            =   1080
            TabIndex        =   4
            Top             =   1680
            Width           =   3855
         End
         Begin VB.TextBox txtCity 
            Height          =   375
            Left            =   6360
            TabIndex        =   6
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox txtState 
            Height          =   375
            Left            =   6360
            TabIndex        =   7
            Top             =   1200
            Width           =   2055
         End
         Begin VB.TextBox txtZip 
            Height          =   405
            Left            =   6360
            TabIndex        =   8
            Top             =   1680
            Width           =   2055
         End
         Begin VB.Label lblStreet2 
            Caption         =   "Street 2:"
            Height          =   255
            Left            =   5040
            TabIndex        =   37
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblCompany 
            Caption         =   "Company:"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblWebsite 
            Caption         =   "Website:"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   840
            Width           =   855
         End
         Begin VB.Label lbtype 
            Caption         =   "Type:"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label lblStreet1 
            Caption         =   "Street 1:"
            Height          =   375
            Left            =   120
            TabIndex        =   33
            Top             =   1800
            Width           =   1095
         End
         Begin VB.Label lblCity 
            Caption         =   "City:"
            Height          =   255
            Left            =   5040
            TabIndex        =   32
            Top             =   840
            Width           =   735
         End
         Begin VB.Label lblState 
            Caption         =   "State /Province:"
            Height          =   255
            Left            =   5040
            TabIndex        =   31
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label lblZipcode 
            Caption         =   "Zip/ Postal Code:"
            Height          =   255
            Left            =   5040
            TabIndex        =   30
            Top             =   1800
            Width           =   1455
         End
      End
      Begin VB.TextBox txtNotes 
         Height          =   4335
         Left            =   -74880
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   480
         Width           =   8415
      End
      Begin VB.Label lblTime 
         Caption         =   "Pick your follow-up time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71040
         TabIndex        =   47
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label lblDate 
         Caption         =   "Pick your follow-up date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74640
         TabIndex        =   46
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Line Line1 
         X1              =   -74760
         X2              =   -66480
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label lblFollowUp 
         Caption         =   "Need to follow up?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74640
         TabIndex        =   45
         Top             =   840
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmNewEditContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
    Me.Icon = frmRecruiter.imlTree.ListImages(8).Picture
    optNo.Value = True
    cboTitle.AddItem "Mr"
    cboTitle.AddItem "Mrs"
    cboTitle.AddItem "Miss"
    DTPicker1.Format = dtpCustom
    DTPicker1.CustomFormat = "hh:mm tt"
    SSTab1.Tab = 0
End Sub

Private Sub optNo_Click()
    mvDate.Enabled = False
    DTPicker1.Enabled = False
End Sub

Private Sub optYes_Click()
    mvDate.Enabled = True
    DTPicker1.Enabled = True
End Sub


'---------------------------- save -----------------------------'
Private Sub cmdSave_Click()
    Call save
    If isEditMode Then
        Call frmInfo.info
    End If
End Sub

Private Sub save()
    If txtCompany = "" Then
        MsgBox "Please insert the company name in order to save the info or quit otherwise!", vbCritical, "Error insertion"
        SSTab1.Tab = 0
        txtCompany.SetFocus
        Exit Sub
    End If
    'check for correct action to recordset, either add new record or update it
    If isEditMode Then
        rsRecruiter.edit
    Else
        'check if the company is in database
        If rsRecruiter.RecordCount > 0 Then
            rsRecruiter.MoveFirst
            rsRecruiter.FindFirst ("[CompanyName]='" & txtCompany & "'")
            If Not rsRecruiter.NoMatch Then
                MsgBox "The company name you inserted is alreaady in the database!", vbCritical, "Duplicate"
                SSTab1.Tab = 0
                txtCompany.SetFocus
                Exit Sub
            End If
        End If
        rsRecruiter.AddNew
    End If
    Call insert1 'save 1st form
    Call insert2 'save 2nd form
    Call insert3 'save 3rd form
    rsRecruiter.Update
    
    'refresh treeview
    Dim newKey As String
    newKey = UCase(Left(txtCompany, 1))
    Dim lIndex As Long
    'trace all nodes in treeview
    For lIndex = 1 To 26
        If frmRecruiter.tvTreeView.Nodes.Item(lIndex).key Like newKey Then
            frmRecruiter.tvTreeView.Nodes.Item(lIndex).Selected = True
        End If
    Next lIndex

    Unload Me 'unload the form
End Sub

Private Sub insert1()
    With rsRecruiter
        !CompanyName = Trim(UCase(Left(txtCompany, 1))) & Trim(Mid(txtCompany, 2))
        !URL = Trim(LCase(txtURL))
        !Type = Trim(UCase(Left(txtType, 1))) & Trim(Mid(txtType, 2))
        !Street = Trim(txtStreet1) & " " & Trim(txtStreet2)
        !City = Trim(UCase(Left(txtCity, 1))) & Trim(Mid(txtCity, 2))
        !State = Trim(UCase(txtState))
        !Zip = Trim(txtZip)
        !Title = Trim(cboTitle.Text)
        !LastName = Trim(UCase(Left(txtLast, 1))) & Trim(Mid(txtLast, 2))
        !FirstName = Trim(UCase(Left(txtFirst, 1))) & Trim(Mid(txtFirst, 2))
        !email = Trim(txtEmail)
        !Phone = Trim(txtPhone)
        !Extension = Trim(txtExtension)
    End With
End Sub

Private Sub insert2()
    If optYes.Value = True Then
        rsRecruiter!followUp = True
        Dim newdate As String
        Dim newtime As String
        newdate = mvDate.Value
        newtime = DTPicker1.Value
        
        Dim temp1 As String
        Dim temp2 As String
        'remove the date
        temp1 = InStrRev(newtime, "/")
        If temp1 <> 0 Then
            newtime = Mid(newtime, temp1 + 5)
        End If
        'remove the seconds
        temp1 = InStrRev(newtime, ":")
        If temp1 <> 0 Then
            temp2 = Mid(newtime, 1, temp1 - 1)
            temp2 = temp2 & Mid(newtime, temp1 + 3)
            newtime = temp2
        End If
        rsRecruiter!Date = Trim(newdate)
        rsRecruiter!Time = Trim(newtime)
    Else
        rsRecruiter!followUp = False
    End If
End Sub

Private Sub insert3()
    rsRecruiter!Notes = Trim(txtNotes)
End Sub
'---------------------------- save -----------------------------'


'---------------------------- reset -----------------------------'
Private Sub cmdreset1_Click()
    txtCompany.Text = ""
    txtURL.Text = ""
    txtType.Text = ""
    txtStreet1.Text = ""
    txtStreet2.Text = ""
    txtCity.Text = ""
    txtState.Text = ""
    txtZip.Text = ""
    cboTitle.Text = ""
    txtFirst.Text = ""
    txtLast.Text = ""
    txtEmail.Text = ""
    txtPhone.Text = ""
    txtExtension.Text = ""
End Sub

Private Sub cmdreset2_click()
    optNo.Value = True
End Sub

Private Sub cmdreset3_Click()
    txtNotes.Text = ""
End Sub
'---------------------------- reset -----------------------------'


'---------------------------- cancel -----------------------------'
Private Sub cmdCancel_Click()
    Unload Me
    If isEditMode Then
        Call frmInfo.info
    End If
End Sub
'---------------------------- cancel -----------------------------'


'---------------------------- next -----------------------------'
Private Sub cmdNext1_Click()
    SSTab1.Tab = 1
End Sub

Private Sub cmdNext2_Click()
    SSTab1.Tab = 2
End Sub

Private Sub cmdBack2_Click()
    SSTab1.Tab = 0
End Sub

Private Sub cmdBack3_Click()
    SSTab1.Tab = 1
End Sub
'---------------------------- next -----------------------------'

