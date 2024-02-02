VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} QueryForm 
   Caption         =   "INSERT QUERY (Created by Seokhoon Joo)"
   ClientHeight    =   9583.001
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10425
   OleObjectBlob   =   "QueryForm.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "QueryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ==============================
' 2013 Created by Seokhoon Joo
' SQL with Excel
' ==============================

Option Explicit

Private Sub UserForm_Initialize()
    Dim i As Long
    Dim lastRowQuery As Long
    Dim lastRowDB As Long
    Dim strConn As String
    Dim adoConn As ADODB.Connection
    Dim adoRS As ADODB.Recordset
    Dim wbName As String
    Dim fileExt As String
    Dim tblCollection As New collection
    
    lastRowQuery = Sheets("snippet").Cells(Rows.Count, 2).End(3).Row
    For i = 2 To lastRowQuery
        With Sheets("snippet")
            lbxQueryCollection.AddItem .Cells(i, 2)
        End With
    Next i
    
    lastRowDB = Sheets("list").Cells(Rows.Count, 2).End(3).Row
    For i = 2 To lastRowDB
        With Sheets("list")
            cmbDB.AddItem .Cells(i, 2)
        End With
    Next i
    
    cmbDB = Range("db")
    fileExt = GetFileExtension(cmbDB)
    Select Case fileExt
        Case "xlsb", "xlsm", "xlsx", "xls"
       
            wbName = Range("db")
            Set tblCollection = GetSheetsNames(wbName)
        
            For i = 1 To tblCollection.Count
               Sheets("list").Cells(i + 1, 3) = tblCollection(i)
            Next i
                
            With Sheets("query")
                cmbDB = .Range("db")
                cmbTable = .Range("table")
                txtQuery = .Range("query")
            End With
            
        Case "accdb", "mdb"
        
            wbName = Range("db")
            Set adoConn = New ADODB.Connection
            
            ' Open the connection.
            With adoConn
                .Provider = "Microsoft.ACE.OLEDB.12.0"
                .Open wbName
            End With
        
            ' Open the tables schema rowset.
            Set adoRS = adoConn.OpenSchema(adSchemaTables)
        
            ' Loop through the results and print the
            ' names and types in the Immediate pane.
            i = 2
            With adoRS
                Do While Not .EOF
                    If .Fields("TABLE_TYPE") <> "VIEW" And Left(.Fields("TABLE_NAME"), 4) <> "MSys" Then
                       Sheets("list").Cells(i, 3) = .Fields("TABLE_NAME") '& vbTab & .Fields("TABLE_TYPE")
                    End If
                    .MoveNext
                    i = i + 1
                Loop
            End With
            adoRS.Close
            adoConn.Close
            Set adoRS = Nothing
            Set adoConn = Nothing
           
            With Sheets("query")
                cmbDB = .Range("db")
                cmbTable = .Range("table")
                txtQuery = .Range("query")
            End With
       
        Case Else
        
            Set tblCollection = GetSheetsNames(ThisWorkbook.FullName)
            For i = 1 To tblCollection.Count
               Sheets("list").Cells(i + 1, 3) = tblCollection(i)
            Next i
    
            With Sheets("query")
                cmbDB = ThisWorkbook.FullName
                cmbTable = .Range("table")
                txtQuery = .Range("query")
            End With
    
    End Select
End Sub
Private Sub cmbDB_Change()
    Dim i As Long
    Dim lastRowTbl As Long
    Dim strConn As String
    Dim adoConn As ADODB.Connection
    Dim adoRS As ADODB.Recordset
    Dim fileExt As String
    Dim wbName As String
    Dim tblCollection As New collection
    
    Sheets("list").Range("C2:C1048576").ClearContents
    
    fileExt = GetFileExtension(cmbDB)
    Select Case fileExt
        Case "xlsb", "xlsm", "xlsx", "xls"
            Set tblCollection = GetSheetsNames(cmbDB)
            For i = 1 To tblCollection.Count
               Sheets("list").Cells(i + 1, 3) = tblCollection(i)
            Next i
        Case "accdb", "mdb"
            ' Open the connection.
            Set adoConn = New ADODB.Connection
            With adoConn
               .Provider = "Microsoft.ACE.OLEDB.12.0"
               .Open cmbDB
            End With
            ' Open the tables schema rowset.
            Set adoRS = adoConn.OpenSchema(adSchemaTables)
            ' Loop through the results and print the
            ' names and types in the Immediate pane.
            i = 2
            With adoRS
                Do While Not .EOF
                    lastRowTbl = Sheets("list").Cells(Rows.Count, 3).End(3).Row
                    If .Fields("TABLE_TYPE") <> "VIEW" And Left(.Fields("TABLE_NAME"), 4) <> "MSys" Then
                          Sheets("list").Cells(lastRowTbl + 1, 3) = .Fields("TABLE_NAME") '& vbTab & .Fields("TABLE_TYPE")
                    End If
                   .MoveNext
               Loop
            End With
            adoRS.Close
            adoConn.Close
            Set adoRS = Nothing
            Set adoConn = Nothing
    End Select

    lastRowTbl = Sheets("list").Cells(Rows.Count, 3).End(3).Row
    With Sheets("list")
        cmbTable.Clear
        For i = 2 To lastRowTbl
            cmbTable.AddItem .Cells(i, 3)
        Next i
    End With
        
    With Sheets("query")
        .Range("db") = cmbDB
    End With
            
End Sub
Private Sub cmbTable_Change()
    Dim adoConn As New ADODB.Connection
    Dim adoRS As ADODB.Recordset
    Dim strConn As String
    Dim query As String
    Dim fileExt As String
    Dim i As Long, lastrowCol As Long
    
    Dim objConn As ADODB.Connection
    Dim objCat As ADOX.Catalog
    Dim tbl As ADOX.Table
    Dim tblCollection As New collection
    
    Sheets("list").Range("D2:D1048576").ClearContents

    Set adoConn = New ADODB.Connection
    Set adoRS = New ADODB.Recordset
             
    fileExt = GetFileExtension(cmbDB)
    Select Case fileExt
        Case "xlsb", "xlsm", "xlsx", "xls"
      
            strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                           "Data Source=" & cmbDB & ";" & _
                           "Extended Properties=Excel 12.0;"
                                
            adoConn.Open strConn
                
            query = "SELECT TOP 1 * FROM [Excel 12.0;HDR=YES;DATABASE=" & cmbDB & "]." & cmbTable
            If adoConn.State = adStateOpen And cmbTable <> "" Then
                adoRS.Open query, strConn, adOpenForwardOnly, adLockReadOnly, adCmdText
                If Not adoRS.EOF Then
                     With Sheets("list")
                         For i = 0 To adoRS.Fields.Count - 1
                            .Cells(i + 2, 4).Value = adoRS.Fields(i).Name
                         Next i
                    End With
                Else
                    MsgBox "No data", 64, "Error"
                End If
                adoRS.Close
            Else
            End If
            adoConn.Close
            
            Set adoConn = Nothing
            Set adoRS = Nothing
            
        Case "accdb", "mdb"
             
            strConn = "Provider=Microsoft.ACE.OLEDB.12.0; " & _
                           "Data Source=" & cmbDB
                                
            adoConn.Open strConn
                
            query = "SELECT TOP 1 * FROM " & cmbTable
            If adoConn.State = adStateOpen And cmbTable <> "" Then
                adoRS.Open query, strConn, adOpenForwardOnly, adLockReadOnly, adCmdText
                If Not adoRS.EOF Then
                    With Sheets("list")
                        For i = 0 To adoRS.Fields.Count - 1
                           .Cells(i + 2, 4).Value = adoRS.Fields(i).Name
                        Next i
                    End With
                Else
                    MsgBox "No data", 64, "Error"
                End If
                adoRS.Close
            Else
            End If
            adoConn.Close
            
            Set adoConn = Nothing
            Set adoRS = Nothing
        
        Case Else
        
            MsgBox "Unsupported file format"
        
    End Select

    lastrowCol = Sheets("list").Cells(Rows.Count, 4).End(3).Row
    With Sheets("list")
        lbxCol.Clear
        For i = 2 To lastrowCol
            lbxCol.AddItem .Cells(i, 4)
        Next i
    End With
        
    With Sheets("Query")
        .Range("table") = cmbTable
    End With
             
End Sub

Private Sub cmdInsert_Click()
    Dim adoConn As ADODB.Connection
    Dim adoRS As ADODB.Recordset
    Dim strConn As String
    Dim query As String
    Dim i As Long
    
    Range("query") = txtQuery
    query = txtQuery

    If InStr(1, txtQuery, "INTO") <> 0 Or InStr(1, txtQuery, "INSERT") <> 0 Or InStr(1, txtQuery, "DELETE") <> 0 Or InStr(1, txtQuery, "DROP") <> 0 Then
        Set adoConn = New ADODB.Connection
        Set adoRS = New ADODB.Recordset
        
        strConn = "Provider=Microsoft.ACE.OLEDB.12.0; " & _
                       "Data Source=" & cmbDB
                       
        adoConn.Open strConn
        adoConn.Execute query
        adoRS.Close
        adoConn.Close
        Set adoConn = Nothing
        Set adoRS = Nothing
        MsgBox "        Done"
        Exit Sub
    End If
   
    Range("query") = txtQuery
    query = txtQuery
    If InStr(1, query, "$") = 0 Then
        Set adoConn = New ADODB.Connection
        Set adoRS = New ADODB.Recordset
             
        strConn = "Provider=Microsoft.ACE.OLEDB.12.0; " & _
                       "Data Source=" & cmbDB
                       
        On Error Resume Next

        adoConn.Open strConn
        
        If Err.Number <> 0 Then
            MsgBox "Error: " & Err.Source, vbExclamation, "Error"
            Err.Clear
        End If
            
        If adoConn.State = adStateOpen Then
            adoRS.Open query, strConn, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not adoRS.EOF Then
                With Sheets("query")
                    Rows(.Range("header").Row).ClearContents
                    .Range("data").CurrentRegion.Clear
                    For i = 0 To adoRS.Fields.Count - 1
                       .Cells(.Range("header").Row, i + 1).Value = adoRS.Fields(i).Name
                    Next i
                    .Range("data").CopyFromRecordset adoRS
                    .Range("B:ZZ").Columns.AutoFit
                    .Activate
                End With
            Else
                MsgBox "No data", 64, "Error"
            End If
            adoRS.Close
        Else
        End If
        adoConn.Close
        
        Set adoRS = Nothing
        Set adoConn = Nothing
            
    Else
    
        Set adoConn = New ADODB.Connection
        Set adoRS = New ADODB.Recordset
        
        strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
            "Data Source=" & cmbDB & ";" & _
            "Extended Properties=Excel 12.0;"
            
        On Error Resume Next
        
        adoConn.Open strConn
        
        If Err.Number <> 0 Then
            MsgBox "Error: " & Err.Source, vbExclamation, "Error"
            Err.Clear
        End If
        
        If adoConn.State = adStateOpen Then
            adoRS.Open query, strConn, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not adoRS.EOF Then
                With Sheets("query")
                    Rows(.Range("header").Row).ClearContents
                    .Range("data").CurrentRegion.Clear
                    For i = 0 To adoRS.Fields.Count - 1
                       .Cells(.Range("header").Row, i + 1).Value = adoRS.Fields(i).Name
                    Next i
                   .Range("data").CopyFromRecordset adoRS
                   .Range("B:ZZ").Columns.AutoFit
                   .Activate
                End With
            Else
                MsgBox "No data", 64, "Error"
            End If
            adoRS.Close
        End If
        adoConn.Close
        
        Set adoRS = Nothing
        Set adoConn = Nothing
        
    End If
    
    Range("data").Select
End Sub

Private Sub cmdInsert2_Click()
    Dim adoConn As ADODB.Connection
    Dim adoRS As ADODB.Recordset
    Dim strConn As String
    Dim query As String
    Dim i As Long
    
    Range("query") = txtQueryCollection
    txtQuery = txtQueryCollection
    query = txtQuery
   
    If InStr(1, query, "INTO") <> 0 Or InStr(1, query, "INSERT") <> 0 Or InStr(1, query, "DELETE") <> 0 Or InStr(1, query, "DROP") <> 0 Then
   
        Set adoConn = New ADODB.Connection
        Set adoRS = New ADODB.Recordset
        
        strConn = "Provider=Microsoft.ACE.OLEDB.12.0; " & _
            "Data Source=" & cmbDB
                       
        adoConn.Open strConn
        adoConn.Execute query
        adoRS.Close
        adoConn.Close
        Set adoConn = Nothing
        Set adoRS = Nothing
        MsgBox "        Done"
        Exit Sub
    Else
    End If
   
    Range("query") = txtQueryCollection
    txtQuery = txtQueryCollection
    query = txtQuery
   
    If InStr(1, txtQuery, "$") = 0 Then
       
        Set adoConn = New ADODB.Connection
        Set adoRS = New ADODB.Recordset
             
        strConn = "Provider=Microsoft.ACE.OLEDB.12.0; " & _
             "Data Source=" & cmbDB
                            
        adoConn.Open strConn
            
        If adoConn.State = adStateOpen Then
            adoRS.Open query, strConn, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not adoRS.EOF Then
                With Sheets("query")
                    Rows(.Range("header").Row).ClearContents
                    .Range("data").CurrentRegion.Clear
                    For i = 0 To adoRS.Fields.Count - 1
                       .Cells(.Range("header").Row, i + 1).Value = adoRS.Fields(i).Name
                    Next i
                    .Range("data").CopyFromRecordset adoRS
                    .Range("B:ZZ").Columns.AutoFit
                    .Activate
                End With
            Else
                MsgBox "No data", 64, "Error"
            End If
            adoRS.Close
        End If
        adoConn.Close
        
        Set adoRS = Nothing
        Set adoConn = Nothing
            
    Else
        
        Set adoConn = New ADODB.Connection
        Set adoRS = New ADODB.Recordset
        
        strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
            "Data Source=" & cmbDB & ";" & _
            "Extended Properties=Excel 12.0;"
                
        adoConn.Open strConn
            
        If adoConn.State = adStateOpen Then
            adoRS.Open query, strConn, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not adoRS.EOF Then
                With Sheets("query")
                   Rows(.Range("header").Row).ClearContents
                    .Range("data").CurrentRegion.Clear
                    For i = 0 To adoRS.Fields.Count - 1
                       .Cells(.Range("header").Row, i + 1).Value = adoRS.Fields(i).Name
                    Next i
                   .Range("data").CopyFromRecordset adoRS
                   .Range("B:ZZ").Columns.AutoFit
                   .Activate
                End With
            Else
                MsgBox "No data", 64, "Error"
            End If
            adoRS.Close
        End If
        adoConn.Close
        
        Set adoRS = Nothing
        Set adoConn = Nothing
        
    End If
    
    Range("data").Select
End Sub

Private Sub cmdDesc_Click()
    txtQuery.SelText = "DESC "
    txtQuery.SetFocus
End Sub

Private Sub cmdDrop_Click()
    txtQuery.SelText = "DROP TABLE "
    txtQuery.SetFocus
End Sub

Private Sub cmdHaving_Click()
    txtQuery.SelText = "HAVING "
    txtQuery.SetFocus
End Sub

Private Sub cmdInto_Click()
    txtQuery.SelText = "SELECT * INTO " & vbCrLf & "FROM [Excel 12.0;HDR=YES;DATABASE=" & ThisWorkbook.FullName & "].[table$]"
    txtQuery.SetFocus
End Sub

Private Sub cmdJoin_Click()
    txtQuery.SelText = "LEFT OUTER JOIN "
    txtQuery.SetFocus
End Sub

Private Sub cmdOn_Click()
    txtQuery.SelText = "ON "
    txtQuery.SetFocus
End Sub

Private Sub cmdOrder_Click()
    txtQuery.SelText = "ORDER BY "
    txtQuery.SetFocus
End Sub

Private Sub cmdPer_Click()
    txtQuery.SelText = "%"
    txtQuery.SetFocus
End Sub

Private Sub cmdA_Click()
    txtQuery.SelText = "A"
    txtQuery.SetFocus
End Sub

Private Sub cmdAnd_Click()
    txtQuery.SelText = "AND "
    txtQuery.SetFocus
End Sub

Private Sub cmdAs_Click()
    txtQuery.SelText = "AS "
    txtQuery.SetFocus
End Sub

Private Sub cmdB_Click()
    txtQuery.SelText = "B"
    txtQuery.SetFocus
End Sub

Private Sub cmdBase_Click()
    txtQuery.SelText = "SELECT "
    txtQuery.SetFocus
End Sub

Private Sub cmdCol_Click()
    txtQuery.SelText = cmbCol
    txtQuery.SetFocus
End Sub

Private Sub cmdComma_Click()
    txtQuery.SelText = ", "
    txtQuery.SetFocus
End Sub

Private Sub cmdCount_Click()
    txtQuery.SelText = "COUNT"
    txtQuery.SetFocus
End Sub

Private Sub cmdDistinct_Click()
    txtQuery.SelText = "DISTINCT "
    txtQuery.SetFocus
End Sub

Private Sub cmdEqual_Click()
    txtQuery.SelText = "="
    txtQuery.SetFocus
End Sub

Private Sub cmdFrom_Click()
    txtQuery.SelText = "FROM "
    txtQuery.SetFocus
End Sub

Private Sub cmdGroup_Click()
    txtQuery.SelText = "GROUP BY "
    txtQuery.SetFocus
End Sub

Private Sub cmdL_Click()
    txtQuery.SelText = ">"
    txtQuery.SetFocus
End Sub

Private Sub cmdOr_Click()
    txtQuery.SelText = "OR "
    txtQuery.SetFocus
End Sub

Private Sub cmdP_Click()
    txtQuery.SelText = "."
    txtQuery.SetFocus
End Sub

Private Sub cmdR_Click()
    txtQuery.SelText = "<"
    txtQuery.SetFocus
End Sub

Private Sub cmdStar_Click()
    txtQuery.SelText = "* "
    txtQuery.SetFocus
End Sub

Private Sub cmdSub_Click()
    Dim i As Long
    Dim k As Long
    Dim lastRow As Long
    
    lastRow = Sheets("snippet").Cells(Rows.Count, 3).Row
    With Sheets("snippet")
        .Range(.Cells(2, 3), .Cells(lastRow, 3)).Replace What:=txtFindWord, Replacement:=txtReplaceWord, LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        txtQueryCollection = Replace(txtQueryCollection, txtFindWord, txtReplaceWord, , , vbTextCompare)
    End With
End Sub

Private Sub cmdSum_Click()
    txtQuery.SelText = "SUM"
    txtQuery.SetFocus
End Sub

Private Sub cmdTable_Click()
    If Left(cmbTable, 1) <> "[" Then
        txtQuery.SelText = cmbTable
        txtQuery.SetFocus
    ElseIf cmbDB <> ThisWorkbook.FullName Then
        txtQuery.SelText = "[Excel 12.0;HDR=YES;DATABASE=" & cmbDB & "]." & cmbTable
        txtQuery.SetFocus
    Else
        txtQuery.SelText = cmbTable
        txtQuery.SetFocus
    End If
End Sub

Private Sub cmdTop_Click()
    txtQuery.SelText = "SELECT TOP 30 * " & vbCrLf & "FROM "
    txtQuery.SetFocus
End Sub

Private Sub cmdWhere_Click()
    txtQuery.SelText = "WHERE "
    txtQuery.SetFocus
End Sub

Private Sub cmdRightBracket_Click()
    txtQuery.SelText = ") "
    txtQuery.SetFocus
End Sub

Private Sub cmdLeftBracket_Click()
    txtQuery.SelText = "("
    txtQuery.SetFocus
End Sub

Private Sub lbxCol_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim i As Long
    Dim c As Range
    Dim lastrowCol  As Long
    
    lastrowCol = Sheets("list").Cells(Rows.Count, 3).End(3).Row
    For i = 0 To lbxCol.ListCount - 1
        If lbxCol.Selected(i) = True Then
            txtQuery.SelText = lbxCol & ", "
            txtQuery.SetFocus
        End If
    Next i
End Sub

Private Sub lbxQueryCollection_Change()
    Dim i As Long
    Dim c As Range
    Dim lastRow As Long
    
    lastRow = Sheets("snippet").Cells(Rows.Count, 2).End(3).Row
    For i = 0 To lbxQueryCollection.ListCount - 1
        If lbxQueryCollection.Selected(i) = True Then
            txtQueryCollection = lbxQueryCollection
            With Sheets("snippet")
                Set c = .Range(.Cells(1, 1), .Cells(lastRow, 3)).Find(lbxQueryCollection, LookAt:=xlWhole)
                If Not c Is Nothing Then
                    txtQueryCollection = c.Offset(0, 1)
                End If
            End With
        End If
    Next i
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdExit2_Click()
    Unload Me
End Sub

