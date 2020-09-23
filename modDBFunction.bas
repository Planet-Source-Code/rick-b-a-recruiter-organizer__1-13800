Attribute VB_Name = "modDBFunction"
Option Explicit


Public Function openDatabase() As Boolean
    Dim result As String
On Error GoTo dbErrors
    result = Dir(dbPath)
    'MsgBox "filename: " & result
    If IsNull(result) Or result = "" Then
        Call createDatabase
    End If
    openDatabase = True
    Exit Function
dbErrors:
    openDatabase = False
    MsgBox (Err.Description)
    
End Function


Public Sub makeColumns1()
    frmRecruiter.lvListView.ColumnHeaders.Clear
    frmRecruiter.lvListView.ColumnHeaders.Add 1, , "Company Name", 2500
    frmRecruiter.lvListView.ColumnHeaders.Add 2, , "URL", 2500
    frmRecruiter.lvListView.ColumnHeaders.Add 3, , "FollowUp", 900
    frmRecruiter.lvListView.ColumnHeaders.Add 4, , "Date", 1100
    frmRecruiter.lvListView.ColumnHeaders.Add 5, , "Time", 1100
End Sub

Public Sub makeColumns2()
    frmSearchResult.lvListView.ColumnHeaders.Clear
    frmSearchResult.lvListView.ColumnHeaders.Add 1, , "Company Name", 2400
    frmSearchResult.lvListView.ColumnHeaders.Add 2, , "URL", 2500
    frmSearchResult.lvListView.ColumnHeaders.Add 3, , "Type", 2000
    frmSearchResult.lvListView.ColumnHeaders.Add 4, , "State", 700
    frmSearchResult.lvListView.ColumnHeaders.Add 5, , "Recruiter Name", 2500
    frmSearchResult.lvListView.ColumnHeaders.Add 6, , "FollowUp", 900
    frmSearchResult.lvListView.ColumnHeaders.Add 7, , "Date", 1100
    frmSearchResult.lvListView.ColumnHeaders.Add 8, , "Time", 1100
End Sub


Public Sub createDatabase()
    Dim db As Database
    Dim tbl As TableDef
    Dim fld As Field
    'create workspaces
    Dim wrkIS As Workspace
    Set wrkIS = CreateWorkspace("RecruiterOrganizer.mdb", "admin", "", dbUseJet)
    'open the database
    Set db = wrkIS.createDatabase("RecruiterOrganizer.mdb", dbLangGeneral)
    
    '---------------Create Contacts Table-------------
    'create a tabledef object
    Set tbl = db.CreateTableDef("tblRecruiter")
    'create the fields with its type and ordinal position, then add it to table
    
    Set fld = tbl.CreateField("CompanyName", dbText, 50)
    fld.OrdinalPosition = 1
    fld.Required = True
    fld.AllowZeroLength = False
    tbl.Fields.Append fld
    
    Set fld = tbl.CreateField("URL", dbText, 50)
    fld.OrdinalPosition = 2
    fld.AllowZeroLength = True
    tbl.Fields.Append fld
    
    Set fld = tbl.CreateField("Type", dbText, 50)
    fld.OrdinalPosition = 3
    fld.AllowZeroLength = True
    tbl.Fields.Append fld
    
    Set fld = tbl.CreateField("Street", dbText, 255)
    fld.OrdinalPosition = 4
    fld.AllowZeroLength = True
    tbl.Fields.Append fld
    
    Set fld = tbl.CreateField("City", dbText, 25)
    fld.OrdinalPosition = 5
    fld.AllowZeroLength = True
    tbl.Fields.Append fld
    
    Set fld = tbl.CreateField("State", dbText, 2)
    fld.OrdinalPosition = 6
    fld.AllowZeroLength = True
    tbl.Fields.Append fld
    
    Set fld = tbl.CreateField("Zip", dbText, 10)
    fld.OrdinalPosition = 7
    fld.AllowZeroLength = True
    tbl.Fields.Append fld
    
    Set fld = tbl.CreateField("Title", dbText, 4)
    fld.OrdinalPosition = 8
    fld.AllowZeroLength = True
    tbl.Fields.Append fld
    
    Set fld = tbl.CreateField("LastName", dbText, 50)
    fld.OrdinalPosition = 9
    fld.AllowZeroLength = True
    tbl.Fields.Append fld
    
    Set fld = tbl.CreateField("FirstName", dbText, 50)
    fld.OrdinalPosition = 10
    fld.AllowZeroLength = True
    tbl.Fields.Append fld
    
    Set fld = tbl.CreateField("Email", dbText, 30)
    fld.OrdinalPosition = 11
    fld.AllowZeroLength = True
    tbl.Fields.Append fld
    
    Set fld = tbl.CreateField("Phone", dbText, 20)
    fld.OrdinalPosition = 12
    fld.AllowZeroLength = True
    tbl.Fields.Append fld
    
    Set fld = tbl.CreateField("Extension", dbText, 5)
    fld.OrdinalPosition = 13
    fld.AllowZeroLength = True
    tbl.Fields.Append fld
        
    Set fld = tbl.CreateField("FollowUp", dbBoolean)
    fld.OrdinalPosition = 14
    fld.DefaultValue = False
    tbl.Fields.Append fld
    
    Set fld = tbl.CreateField("Date", dbText, 15)
    fld.OrdinalPosition = 15
    fld.AllowZeroLength = True
    tbl.Fields.Append fld
    
    Set fld = tbl.CreateField("Time", dbText, 15)
    fld.OrdinalPosition = 16
    fld.AllowZeroLength = True
    tbl.Fields.Append fld
    
    Set fld = tbl.CreateField("Notes", dbMemo)
    fld.OrdinalPosition = 17
    fld.AllowZeroLength = True
    tbl.Fields.Append fld
   
    Set fld = tbl.CreateField("cID", dbLong)
    fld.OrdinalPosition = 18
    fld.Required = True
    fld.Attributes = dbAutoIncrField
    tbl.Fields.Append fld
    
    'add the table to database
    db.TableDefs.Append tbl
    '---------------Create Contacts Table-------------
    Set wrkIS = Nothing
    Set db = Nothing
    Set tbl = Nothing
    Set fld = Nothing
End Sub

