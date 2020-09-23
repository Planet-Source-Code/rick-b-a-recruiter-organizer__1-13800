Attribute VB_Name = "modPopulateListView"
Public Sub populateListView(key As String)
    Dim rs As Recordset
    Dim sSQL As String
    Dim flag As Boolean
    
    flag = False 'if false then letter company name, true then digit company name
    
    sSQL = "SELECT CompanyName, URL, FollowUp, Date, Time, cID " & _
           "FROM tblRecruiter "
    Set rs = dbRecruiter.OpenRecordset(sSQL, dbOpenDynaset)
    
    frmRecruiter.lvListView.ListItems.Clear
    
    'check for digit start letter
    If InStr(key, "0") <> 0 Then
        flag = True
    End If
    
    With rs
        If .RecordCount > 0 Then
            'create a listitem object to be inserted in listview
            Dim itemX As ListItem
            Dim displayRecord As Boolean 'to filter the record
            .MoveFirst
            Do While Not .EOF
                If Not flag And (UCase(Left(.Fields("CompanyName"), 1)) Like key) Then
                    displayRecord = True
                ElseIf flag And InStr("0123456789", Left(.Fields("CompanyName"), 1)) <> 0 Then
                    displayRecord = True
                Else
                    displayRecord = False
                End If
                'display filtered record
                If displayRecord Then
                    Set itemX = frmRecruiter.lvListView.ListItems.Add(, CStr(.Fields("cID")) & _
                                "_ " & CStr(.AbsolutePosition), .Fields("CompanyName"))
                    If (.Fields("URL") <> "") Then
                        itemX.SubItems(1) = .Fields("URL")
                    Else
                        itemX.SubItems(1) = "No URL listed"
                    End If
                    If (.Fields("Followup")) Then
                        itemX.SubItems(2) = "Yes"
                        itemX.SubItems(3) = .Fields("Date")
                        itemX.SubItems(4) = .Fields("Time")
                    Else
                        itemX.SubItems(2) = "No"
                        itemX.SubItems(3) = ""
                        itemX.SubItems(4) = ""
                    End If
                End If
                .MoveNext
            Loop
        End If
    End With
    Set rs = Nothing
End Sub

