Attribute VB_Name = "modExternalApp"
Option Explicit


Public Sub email(emailAddress As String)
    Dim IE As Object
    Set IE = CreateObject("internetexplorer.application")
    IE.Visible = False
    IE.Navigate "Mailto:" & emailAddress
    IE.Quit   ' When you finish, use the Quit method to close
    Set IE = Nothing   ' the application, then release the reference.
End Sub

Public Sub webBrowse(urlAddress As String)
    Dim IE As Object
    Set IE = CreateObject("internetexplorer.application")
    IE.Visible = True
    IE.Navigate urlAddress
End Sub
