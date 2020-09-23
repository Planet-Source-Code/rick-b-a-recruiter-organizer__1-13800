VERSION 5.00
Begin VB.Form frmInfo 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   7920
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   4200
      TabIndex        =   25
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton cmdEditInfo 
      Caption         =   "Edit Info"
      Height          =   375
      Left            =   1560
      TabIndex        =   24
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Frame frNotes 
      Caption         =   "Notes"
      Height          =   1095
      Left            =   120
      TabIndex        =   23
      Top             =   6240
      Width           =   6855
      Begin VB.TextBox txtDisplayNotes 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   6615
      End
   End
   Begin VB.Frame frFollowUp 
      Caption         =   "Follow-Up Info"
      Height          =   855
      Left            =   120
      TabIndex        =   20
      Top             =   5280
      Width           =   6855
      Begin VB.TextBox txtDisplayTime 
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
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtDisplayDate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   270
         Width           =   1215
      End
      Begin VB.TextBox txtDisplayFollowUp 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Time:"
         Height          =   255
         Left            =   4680
         TabIndex        =   36
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Date:"
         Height          =   255
         Index           =   7
         Left            =   2760
         TabIndex        =   22
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Need to follow-up?"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame frRecruiter 
      Caption         =   "Recruter Info"
      Height          =   2055
      Left            =   0
      TabIndex        =   13
      Top             =   3120
      Width           =   6855
      Begin VB.TextBox txtDisplayLastName 
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
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox txtDisplayExtension 
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
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton cmdEmailRecruiter 
         Caption         =   "Email Now"
         Height          =   615
         Left            =   4080
         Picture         =   "frmInfo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtDisplayPhone 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1575
         Width           =   1815
      End
      Begin VB.TextBox txtDisplayEmail 
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txtDisplayFirstName 
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "Last Name:"
         Height          =   255
         Left            =   3240
         TabIndex        =   32
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Extension:"
         Height          =   255
         Left            =   3120
         TabIndex        =   30
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Phone:"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   19
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Email:"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "First Name:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame frGeneral 
      Caption         =   "General Info"
      Height          =   2655
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   360
      Width           =   6855
      Begin VB.TextBox txtDisplayState 
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
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtDisplayZip 
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
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtDisplayCity 
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   1680
         Width           =   2895
      End
      Begin VB.CommandButton cmdOpenWebsite 
         Caption         =   "Open Browser?"
         Height          =   375
         Left            =   5160
         Picture         =   "frmInfo.frx":0442
         TabIndex        =   3
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtDisplayURL 
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   720
         Width           =   3855
      End
      Begin VB.TextBox txtDisplayType 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   5655
      End
      Begin VB.TextBox txtDisplayStreet 
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
         Left            =   1080
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   1200
         Width           =   5655
      End
      Begin VB.Label Label6 
         Caption         =   "State:"
         Height          =   255
         Left            =   4080
         TabIndex        =   34
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Zip:"
         Height          =   255
         Left            =   5400
         TabIndex        =   28
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "City:"
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Website:"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Type:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Street:"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   1320
         Width           =   735
      End
   End
   Begin VB.Label lblCompany 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6855
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdEditInfo_Click()
    Call edit
End Sub

Public Sub edit()
    isEditMode = True
    With rsRecruiter 'find the correct record
        .MoveFirst
        Call .FindFirst("[cID]= " & companyID)
        If Not .NoMatch Then
            frmNewEditContact.txtCompany = !CompanyName
            If .Fields("URL") <> "" Then
                frmNewEditContact.txtURL = !URL
            End If
            If .Fields("Type") <> "" Then
                frmNewEditContact.txtType = !Type
            End If
            If .Fields("Street") <> "" Then
                frmNewEditContact.txtStreet1 = !Street
            End If
            If .Fields("City") <> "" Then
                frmNewEditContact.txtCity = !City
            End If
            If .Fields("State") <> "" Then
                frmNewEditContact.txtState = !State
            End If
            If .Fields("Zip") <> "" Then
                frmNewEditContact.txtZip = !Zip
            End If
            If .Fields("Title") <> "" Then
                frmNewEditContact.cboTitle.Text = !Title
            End If
            If .Fields("LastName") <> "" Then
                frmNewEditContact.txtLast = !LastName
            End If
            If .Fields("FirstName") <> "" Then
                frmNewEditContact.txtFirst = !FirstName
            End If
            If .Fields("Email") <> "" Then
                frmNewEditContact.txtEmail = !email
            End If
            If .Fields("Phone") <> "" Then
                frmNewEditContact.txtPhone = !Phone
            End If
            If .Fields("Extension") <> "" Then
                frmNewEditContact.txtExtension = !Extension
            End If
            If !followUp <> 0 Then
                frmNewEditContact.optYes.Value = True
            Else
                frmNewEditContact.optNo.Value = True
            End If
            If .Fields("Notes") <> "" Then
                frmNewEditContact.txtNotes = !Notes
            End If
        End If
    End With
    frmNewEditContact.Show vbModal
End Sub


Public Sub info()
    Dim tempID As Long
    tempID = CInt(companyID)
    With rsRecruiter 'find the correct record and delete it
        .MoveFirst
        Call .FindFirst("[cID]= " & companyID)
        If Not .NoMatch Then
            'display company name
            frmInfo.lblCompany.Caption = .Fields("CompanyName")
            'display company type
            If .Fields("Type") <> "" Then
                frmInfo.txtDisplayType = .Fields("Type")
            End If
            'display company URL
            If .Fields("URL") <> "" Then
                frmInfo.txtDisplayURL = .Fields("URL")
            End If
            'display company address
            If .Fields("Street") <> "" Then
                frmInfo.txtDisplayStreet = .Fields("Street")
            End If
            If .Fields("City") <> "" Then
                frmInfo.txtDisplayCity = .Fields("City")
            End If
            If .Fields("State") <> "" Then
                frmInfo.txtDisplayState = .Fields("State")
            End If
            If .Fields("Zip") <> "" Then
                frmInfo.txtDisplayZip = .Fields("Zip")
            End If
            'display recruiter name
            If .Fields("FirstName") <> "" Then
                frmInfo.txtDisplayFirstName = .Fields("FirstName")
            End If
            If .Fields("LastName") <> "" Then
                frmInfo.txtDisplayLastName = .Fields("LastName")
            End If
            'display recruiter email
            If .Fields("Email") <> "" Then
                frmInfo.txtDisplayEmail = .Fields("Email")
            End If
            'display recruiter phone
            If .Fields("Phone") <> "" Then
                frmInfo.txtDisplayPhone = .Fields("Phone")
            End If
            If .Fields("Extension") <> "" Then
                frmInfo.txtDisplayExtension = .Fields("Extension")
            End If
            'display follow-up info
            If .Fields("FollowUp") Then
                frmInfo.txtDisplayFollowUp = "Yes"
                frmInfo.txtDisplayDate = .Fields("Date")
                frmInfo.txtDisplayTime = .Fields("Time")
            Else
                frmInfo.txtDisplayFollowUp = "No"
            End If
            'display notes
            If .Fields("Notes") <> "" Then
                frmInfo.txtDisplayNotes = .Fields("Notes")
            End If
            'displaying the form
            frmInfo.Caption = "Info about " & frmInfo.lblCompany.Caption & _
                         " copyright@2002 by Indra"
        Else
            MsgBox "Should never go here, no record found!!!"
        End If
    End With
End Sub


Private Sub cmdEmailRecruiter_Click()
    If txtDisplayEmail <> "" Then
        Call email(Trim(txtDisplayEmail))
    Else
        MsgBox "No email address provided, action aborted!", vbCritical, "ERROR"
    End If
End Sub

Private Sub cmdOpenWebsite_Click()
    If txtDisplayURL <> "" Then
        Call webBrowse(Trim(txtDisplayURL))
    Else
        MsgBox "No URL provided, action aborted!", vbCritical, "ERROR"
    End If
End Sub


Private Sub Form_Load()
    Me.Icon = frmRecruiter.imlTree.ListImages(8).Picture
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub cmdDone_Click()
    Unload Me
End Sub

