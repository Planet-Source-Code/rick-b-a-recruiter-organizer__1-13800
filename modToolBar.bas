Attribute VB_Name = "modToolBar"
Option Explicit

'----------------------------- Correspond Menu Functions ---------------------------'
Public Sub back()
    Dim temp As Node
    Dim Index As Integer
    Dim count As Integer
    Set temp = frmRecruiter.tvTreeView.SelectedItem
    Index = temp.Index - 1
    count = frmRecruiter.tvTreeView.Nodes.count
    If Index = 0 Then
        Index = count
    End If
    frmRecruiter.tvTreeView.Nodes(Index).Selected = True
End Sub

Public Sub forward()
    Dim temp As Node
    Dim Index As Integer
    Dim count As Integer
    Set temp = frmRecruiter.tvTreeView.SelectedItem
    Index = temp.Index + 1
    count = frmRecruiter.tvTreeView.Nodes.count
    If Index = (count + 1) Then
        Index = 1
    End If
    frmRecruiter.tvTreeView.Nodes(Index).Selected = True
End Sub

Public Sub delete(itemX As ListItem)
    Dim answer As Variant
    Dim prompt As String
    prompt = "Are you sure want to delete this record?"
    answer = MsgBox(prompt, vbQuestion + vbYesNo, "Confirm Delete?")
    If answer = vbYes Then
        With rsRecruiter 'find the correct record and delete it
            .MoveFirst
            Call .FindFirst("[CompanyName]='" & itemX.Text & "'")
            If Not .NoMatch Then
                .delete
            Else
                MsgBox "Should never go here, no record found!!!"
            End If
        End With
        frmRecruiter.lvListView.ListItems.Remove (itemX.Index)
    Else
        MsgBox "User cancel deletion"
        Exit Sub
    End If
End Sub

Public Sub insert()
    frmNewEditContact.Show vbModal
End Sub


Public Sub search()
    frmSearch.Show vbModal
End Sub
