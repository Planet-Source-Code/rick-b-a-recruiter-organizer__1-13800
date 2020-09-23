VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "About RecruiterOrganizer"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   225
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   2160
      Width           =   1095
   End
   Begin VB.PictureBox picIcon 
      Height          =   495
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Any comment or suggestions please send email to      kurt_kx@yahoo.com"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Width           =   3855
   End
   Begin VB.Label lblWarning 
      Caption         =   "Warning - this application is provided as is, without any explicit or implied warranty."
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   3855
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   5280
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label lblMessage2 
      Caption         =   "Written as an extended course project for IS-371 at"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   1440
      Width           =   4695
   End
   Begin VB.Label lblMessage1 
      Caption         =   "RecruiterOrganizer is a general purpose contact list program"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   1080
      Width           =   5055
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version 1.0"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblTitle 
      Caption         =   "Recruiter Organizer"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmRecruiter.imlTree.ListImages(8).Picture
End Sub
