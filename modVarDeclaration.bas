Attribute VB_Name = "modVarDeclaration"
Option Explicit

Public ctlGotFocus As Integer
Public Const tvGotFocus As Integer = 1
Public Const lvGotFocus As Integer = 2
Public dbRecruiter As Database
Public rsRecruiter As Recordset
Public dbPath As String
Public gFindString As String
Public companyID As String
Public isEditMode As Boolean
