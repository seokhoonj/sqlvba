VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "INSERT strSQL (Created by Seokhoon Joo)"
   ClientHeight    =   9585
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10425
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================
' 2013 Created by Seokhoon Joo (seokhoonj.github.io)
' SQL with Excel
'==============================

Option Explicit

Private Sub UserForm_Initialize()

    Dim i As Long
    Dim lastrow_strSQL As Long
    Dim lastrow_DBList As Long
    Dim strConn As String
    Dim adoConn As ADODB.Connection
    Dim adoRS As ADODB.Recordset
    Dim WBname As String
    Dim collection As New collection
    
        lastrow_strSQL = Sheets("strSQL").Cells(Rows.Count, 2).End(3).Row
        lastrow_DBList = Sheets("List").Cells(Rows.Count, 2).End(3).Row

    
    For i = 2 To lastrow_strSQL
        With Sheets("strSQL")
            listboxstrSQLCollection.AddItem .Cells(i, 2)
        End With
    Next i


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        With Sheets("List")
                    For i = 2 To lastrow_DBList
                            cmbDB.AddItem .Cells(i, 2)
                    Next i
        End With
        
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
   
If Mid(Cells(8, 1), InStrRev(Cells(8, 1), ".") + 1, Len(Cells(8, 1)) - InStrRev(Cells(8, 1), ".")) = "xls" Or _
   Mid(Cells(8, 1), InStrRev(Cells(8, 1), ".") + 1, Len(Cells(8, 1)) - InStrRev(Cells(8, 1), ".")) = "xlsx" Or _
   Mid(Cells(8, 1), InStrRev(Cells(8, 1), ".") + 1, Len(Cells(8, 1)) - InStrRev(Cells(8, 1), ".")) = "xlsm" Or _
   Mid(Cells(8, 1), InStrRev(Cells(8, 1), ".") + 1, Len(Cells(8, 1)) - InStrRev(Cells(8, 1), ".")) = "xlsb" Then
        
        WBname = Cells(8, 1)
        Set collection = GetSheetsNames(WBname)

        For i = 1 To collection.Count
           Sheets("List").Cells(i + 1, 3) = collection(i)
        Next i
        
    With Sheets("Query")
        cmbDB = .Cells(8, 1)
        cmbTable = .Cells(11, 1)
        txtstrSQL = .Cells(1, 1)
    End With
        
ElseIf Mid(Cells(8, 1), InStrRev(Cells(8, 1), ".") + 1, Len(Cells(8, 1)) - InStrRev(Cells(8, 1), ".")) = "accdb" Or _
         Mid(Cells(8, 1), InStrRev(Cells(8, 1), ".") + 1, Len(Cells(8, 1)) - InStrRev(Cells(8, 1), ".")) = "mdb" Then
                     
   Set adoConn = New ADODB.Connection

   ' Open the connection.
   With adoConn
      .Provider = "Microsoft.ACE.OLEDB.12.0"
      .Open Cells(8, 1)
   End With

   ' Open the tables schema rowset.
   Set adoRS = adoConn.OpenSchema(adSchemaTables)

   ' Loop through the results and print the
   ' names and types in the Immediate pane.
   
   i = 2
   With adoRS
      Do While Not .EOF
         If .Fields("TABLE_TYPE") <> "VIEW" And Left(.Fields("TABLE_NAME"), 4) <> "MSys" Then
               Sheets("List").Cells(i, 3) = .Fields("TABLE_NAME") '& vbTab & .Fields("TABLE_TYPE")
         End If
         .MoveNext
         i = i + 1
      Loop
   End With
   adoRS.Close
   adoConn.Close
   
   Set adoRS = Nothing
   Set adoConn = Nothing
   
    With Sheets("Query")
        cmbDB = .Cells(8, 1)
        cmbTable = .Cells(11, 1)
        txtstrSQL = .Cells(1, 1)
    End With
   
Else

        WBname = ThisWorkbook.FullName
        Set collection = GetSheetsNames(WBname)

        For i = 1 To collection.Count
           Sheets("List").Cells(i + 1, 3) = collection(i)
        Next i

    With Sheets("Query")
        cmbDB = ThisWorkbook.FullName
        cmbTable = .Cells(11, 1)
        txtstrSQL = .Cells(1, 1)
    End With

End If
    
End Sub
Private Sub cmbDB_Change()

    Dim i As Long
    Dim lastrow_tbl As Long
    Dim strConn As String
    Dim adoConn As ADODB.Connection
    Dim adoRS As ADODB.Recordset
    Dim WBname As String
    Dim tblCollection As New collection
    
    lastrow_tbl = Sheets("List").Cells(Rows.Count, 3).End(3).Row
    
    With Sheets("List")
        .Range(.Cells(2, 3), .Cells(lastrow_tbl, 3)).ClearContents
    End With
    
    
If Mid(cmbDB, InStrRev(cmbDB, ".") + 1, Len(cmbDB) - InStrRev(cmbDB, ".")) = "xls" Or _
   Mid(cmbDB, InStrRev(cmbDB, ".") + 1, Len(cmbDB) - InStrRev(cmbDB, ".")) = "xlsx" Or _
   Mid(cmbDB, InStrRev(cmbDB, ".") + 1, Len(cmbDB) - InStrRev(cmbDB, ".")) = "xlsm" Or _
   Mid(cmbDB, InStrRev(cmbDB, ".") + 1, Len(cmbDB) - InStrRev(cmbDB, ".")) = "xlsb" Then
           
        WBname = cmbDB
        Set tblCollection = GetSheetsNames(WBname)

        For i = 1 To tblCollection.Count
           Sheets("List").Cells(i + 1, 3) = tblCollection(i)
        Next i
        
Else

       Set adoConn = New ADODB.Connection

   ' Open the connection.
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
      lastrow_tbl = Sheets("List").Cells(Rows.Count, 3).End(3).Row
         If .Fields("TABLE_TYPE") <> "VIEW" And Left(.Fields("TABLE_NAME"), 4) <> "MSys" Then
               Sheets("List").Cells(lastrow_tbl + 1, 3) = .Fields("TABLE_NAME") '& vbTab & .Fields("TABLE_TYPE")
         End If
         .MoveNext
      Loop
   End With
   adoRS.Close
   adoConn.Close
   
   Set adoRS = Nothing
   Set adoConn = Nothing
   
End If

        lastrow_tbl = Sheets("List").Cells(Rows.Count, 3).End(3).Row
        With Sheets("List")
                cmbTable.Clear
                    For i = 2 To lastrow_tbl
                        cmbTable.AddItem .Cells(i, 3)
                    Next i
        End With
        
    With Sheets("Query")
        .Cells(8, 1) = cmbDB
    End With
            
End Sub
Private Sub cmbTable_Change()
        
  Dim adoConn As New ADODB.Connection
  Dim adoRS As ADODB.Recordset
  Dim strConn As String
  Dim strSQL As String
  Dim i As Long, lastrow_col As Long
  
  Dim objConn As ADODB.Connection
  Dim objCat As ADOX.Catalog
  Dim tbl As ADOX.Table
  Dim tblCollection As New collection
  
      lastrow_col = Sheets("List").Cells(Rows.Count, 4).End(3).Row
    
    With Sheets("List")
        .Range(.Cells(2, 4), .Cells(lastrow_col, 4)).ClearContents
    End With

        Set adoConn = New ADODB.Connection
        Set adoRS = New ADODB.Recordset
             
             
If Mid(cmbDB, InStrRev(cmbDB, ".") + 1, Len(cmbDB) - InStrRev(cmbDB, ".")) = "xls" Or _
   Mid(cmbDB, InStrRev(cmbDB, ".") + 1, Len(cmbDB) - InStrRev(cmbDB, ".")) = "xlsx" Or _
   Mid(cmbDB, InStrRev(cmbDB, ".") + 1, Len(cmbDB) - InStrRev(cmbDB, ".")) = "xlsm" Or _
   Mid(cmbDB, InStrRev(cmbDB, ".") + 1, Len(cmbDB) - InStrRev(cmbDB, ".")) = "xlsb" Then
  
        strConn = "Provider=                      Microsoft.ACE.OLEDB.12.0;" & _
                     "Data Source=                " & cmbDB & ";" & _
                     "Extended Properties=   Excel 12.0;"
                            
        adoConn.Open strConn
        
     strSQL = "SELECT top 1 * FROM [Excel 12.0;HDR=YES;DATABASE=" & cmbDB & "]." & cmbTable
     If adoConn.State = adStateOpen And cmbTable <> "" Then
            adoRS.Open strSQL, strConn, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not adoRS.EOF Then
                 With Sheets("List")
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
                             
  ElseIf Mid(cmbDB, InStrRev(cmbDB, ".") + 1, Len(cmbDB) - InStrRev(cmbDB, ".")) = "accdb" Or _
          Mid(cmbDB, InStrRev(cmbDB, ".") + 1, Len(cmbDB) - InStrRev(cmbDB, ".")) = "mdb" Then
         
        strConn = "PROVIDER=Microsoft.ACE.OLEDB.12.0; " & _
                         "DATA SOURCE=" & cmbDB
                            
        adoConn.Open strConn
        
     strSQL = "SELECT top 1 * FROM " & cmbTable
     If adoConn.State = adStateOpen And cmbTable <> "" Then
            adoRS.Open strSQL, strConn, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not adoRS.EOF Then
                 With Sheets("List")
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
    
    Else
    
    MsgBox "Unsupported file format"
    
 End If

    lastrow_col = Sheets("List").Cells(Rows.Count, 4).End(3).Row
        With Sheets("List")
            lbxCol.Clear
                For i = 2 To lastrow_col
                        lbxCol.AddItem .Cells(i, 4)
                Next i
        End With
        
    With Sheets("Query")
        .Cells(11, 1) = cmbTable
    End With
             
End Sub

Private Sub cmdInsert_Click()

    Dim adoConn As ADODB.Connection
    Dim adoRS As ADODB.Recordset
    Dim strConn As String
    Dim strSQL As String
    Dim i As Long
    
   
   Cells(1, 1) = txtstrSQL
   strSQL = txtstrSQL
   

   If InStr(1, txtstrSQL, "INTO") <> 0 Or InStr(1, txtstrSQL, "INSERT") <> 0 Or InStr(1, txtstrSQL, "DELETE") <> 0 Or InStr(1, txtstrSQL, "DROP") <> 0 Then
   
           Set adoConn = New ADODB.Connection
           Set adoRS = New ADODB.Recordset
        
            strConn = "PROVIDER=Microsoft.ACE.OLEDB.12.0; " & _
                 "DATA SOURCE=" & cmbDB

                            
        adoConn.Open strConn

        adoConn.Execute (txtstrSQL)
        
        adoRS.Close
        adoConn.Close
        Set adoConn = Nothing
        Set adoRS = Nothing
        MsgBox "        Done"
        Exit Sub
   Else
   End If
   
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Cells(1, 1) = txtstrSQL
   strSQL = txtstrSQL
   
If InStr(1, txtstrSQL, "$") = 0 Then
       
        Set adoConn = New ADODB.Connection
        Set adoRS = New ADODB.Recordset
             
            strConn = "PROVIDER=Microsoft.ACE.OLEDB.12.0; " & _
                 "DATA SOURCE=" & cmbDB
                            
        adoConn.Open strConn
        
     If adoConn.State = adStateOpen Then
            adoRS.Open strSQL, strConn, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not adoRS.EOF Then
                 With Sheets("Query")
                    .Range("a20").CurrentRegion.ClearContents
                        For i = 0 To adoRS.Fields.Count - 1
                           .Cells(20, i + 1).Value = adoRS.Fields(i).Name
                        Next i
                    .Range("a21").CopyFromRecordset adoRS
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
    
    Set adoConn = Nothing
    Set adoRS = Nothing

        
Else
      
        Set adoRS = New ADODB.Recordset
        strConn = "Provider=                      Microsoft.ACE.OLEDB.12.0;" & _
                     "Data Source=                " & cmbDB & ";" & _
                     "Extended Properties=   Excel 12.0;"
                     
        adoRS.Open strSQL, strConn, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not adoRS.EOF Then
                 With Sheets("Query")
                    .Range("a20").CurrentRegion.ClearContents
                        For i = 0 To adoRS.Fields.Count - 1
                           .Cells(20, i + 1).Value = adoRS.Fields(i).Name
                        Next i
                    .Range("a21").CopyFromRecordset adoRS
                    .Range("B:ZZ").Columns.AutoFit
                    .Activate
                End With
            Else
                MsgBox "No data", 64, "Error"
            End If
    adoRS.Close
    Set adoRS = Nothing
    
End If

Cells(21, 1).Select
End Sub

Private Sub cmdInsert2_Click()

    Dim adoConn As ADODB.Connection
    Dim adoRS As ADODB.Recordset
    Dim strConn As String
    Dim strSQL As String
    Dim i As Long
    
   
   Cells(1, 1) = txtstrSQLCollection
   txtstrSQL = txtstrSQLCollection
   strSQL = txtstrSQL
   

   If InStr(1, txtstrSQL, "INTO") <> 0 Or InStr(1, txtstrSQL, "INSERT") <> 0 Or InStr(1, txtstrSQL, "DELETE") <> 0 Or InStr(1, txtstrSQL, "DROP") <> 0 Then
   
           Set adoConn = New ADODB.Connection
           Set adoRS = New ADODB.Recordset
        
            strConn = "PROVIDER=Microsoft.ACE.OLEDB.12.0; " & _
                 "DATA SOURCE=" & cmbDB
                            
                            
        adoConn.Open strConn

        adoConn.Execute (txtstrSQL)
        
        adoRS.Close
        adoConn.Close
        Set adoConn = Nothing
        Set adoRS = Nothing
        MsgBox "        Done"
        Exit Sub
   Else
   End If
   
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Cells(1, 1) = txtstrSQLCollection
   txtstrSQL = txtstrSQLCollection
   strSQL = txtstrSQL
   
If InStr(1, txtstrSQL, "$") = 0 Then
       
        Set adoConn = New ADODB.Connection
        Set adoRS = New ADODB.Recordset
             
            strConn = "PROVIDER=Microsoft.ACE.OLEDB.12.0; " & _
                 "DATA SOURCE=" & cmbDB
                            
        adoConn.Open strConn
        
     If adoConn.State = adStateOpen Then
            adoRS.Open strSQL, strConn, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not adoRS.EOF Then
                 With Sheets("Query")
                    .Range("a20").CurrentRegion.ClearContents
                        For i = 0 To adoRS.Fields.Count - 1
                           .Cells(20, i + 1).Value = adoRS.Fields(i).Name
                        Next i
                    .Range("a21").CopyFromRecordset adoRS
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
    
    Set adoConn = Nothing
    Set adoRS = Nothing

        
Else
      
        Set adoRS = New ADODB.Recordset
        strConn = "Provider=                      Microsoft.ACE.OLEDB.12.0;" & _
                     "Data Source=                " & cmbDB & ";" & _
                     "Extended Properties=   Excel 12.0;"
                     
        adoRS.Open strSQL, strConn, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not adoRS.EOF Then
                 With Sheets("Query")
                    .Range("a20").CurrentRegion.ClearContents
                        For i = 0 To adoRS.Fields.Count - 1
                           .Cells(20, i + 1).Value = adoRS.Fields(i).Name
                        Next i
                    .Range("a21").CopyFromRecordset adoRS
                    .Range("B:ZZ").Columns.AutoFit
                    .Activate
                End With
            Else
                MsgBox "No data", 64, "Error"
            End If
    adoRS.Close
    Set adoRS = Nothing
    
End If

Cells(21, 1).Select

End Sub

Private Sub cmdDesc_Click()
    txtstrSQL.SelText = "DESC "
    txtstrSQL.SetFocus
End Sub

Private Sub cmdDrop_Click()
    txtstrSQL.SelText = "DROP TABLE "
    txtstrSQL.SetFocus
End Sub

Private Sub cmdHaving_Click()
    txtstrSQL.SelText = "Having "
    txtstrSQL.SetFocus
End Sub

Private Sub cmdInto_Click()
    txtstrSQL.SelText = "SELECT * INTO " & vbCrLf & "FROM [Excel 12.0;HDR=YES;DATABASE=" & ThisWorkbook.FullName & "].[DB$]"
    txtstrSQL.SetFocus
End Sub

Private Sub cmdJoin_Click()
    txtstrSQL.SelText = "LEFT OUTER JOIN "
    txtstrSQL.SetFocus
End Sub
Private Sub cmdOn_Click()
    txtstrSQL.SelText = "ON "
    txtstrSQL.SetFocus
End Sub

Private Sub cmdOrder_Click()
    txtstrSQL.SelText = "ORDER BY"
    txtstrSQL.SetFocus
End Sub

Private Sub cmdPer_Click()
    txtstrSQL.SelText = "%"
    txtstrSQL.SetFocus
End Sub

Private Sub cmdA_Click()
    txtstrSQL.SelText = "A"
    txtstrSQL.SetFocus
End Sub

Private Sub cmdAnd_Click()
    txtstrSQL.SelText = "AND "
    txtstrSQL.SetFocus
End Sub

Private Sub cmdAs_Click()
    txtstrSQL.SelText = "AS "
    txtstrSQL.SetFocus
End Sub

Private Sub cmdB_Click()
    txtstrSQL.SelText = "B"
    txtstrSQL.SetFocus
End Sub

Private Sub cmdBase_Click()
    txtstrSQL.SelText = "SELECT "
    txtstrSQL.SetFocus
End Sub

Private Sub cmdCol_Click()
    txtstrSQL.SelText = cmbCol
    txtstrSQL.SetFocus
End Sub

Private Sub cmdComma_Click()
     txtstrSQL.SelText = ", "
     txtstrSQL.SetFocus
End Sub

Private Sub cmdCount_Click()
         txtstrSQL.SelText = "COUNT"
         txtstrSQL.SetFocus
End Sub

Private Sub cmdDistinct_Click()
         txtstrSQL.SelText = "DISTINCT "
         txtstrSQL.SetFocus
End Sub

Private Sub cmdEqual_Click()
         txtstrSQL.SelText = "="
         txtstrSQL.SetFocus
End Sub

Private Sub cmdFrom_Click()
     txtstrSQL.SelText = "FROM "
     txtstrSQL.SetFocus
End Sub

Private Sub cmdGroup_Click()
     txtstrSQL.SelText = "GROUP BY "
     txtstrSQL.SetFocus
End Sub

Private Sub cmdL_Click()
         txtstrSQL.SelText = ">"
         txtstrSQL.SetFocus
End Sub

Private Sub cmdOr_Click()
     txtstrSQL.SelText = "OR "
     txtstrSQL.SetFocus
End Sub

Private Sub cmdP_Click()
     txtstrSQL.SelText = "."
     txtstrSQL.SetFocus
End Sub

Private Sub cmdR_Click()
         txtstrSQL.SelText = "<"
         txtstrSQL.SetFocus
End Sub

Private Sub cmdStar_Click()
    txtstrSQL.SelText = "* "
    txtstrSQL.SetFocus
End Sub

Private Sub cmdSub_Click()

    Dim i As Long
    Dim k As Long
    Dim lastrow As Long
    
            lastrow = Sheets("strSQL").Cells(Rows.Count, 3).Row
    
        With Sheets("strSQL")
            .Range(.Cells(2, 3), .Cells(lastrow, 3)).Replace What:=txtOld, Replacement:=txtNew, LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False
        End With
        
End Sub

Private Sub cmdSum_Click()
    txtstrSQL.SelText = "SUM"
    txtstrSQL.SetFocus
End Sub

Private Sub cmdTable_Click()
If Left(cmbTable, 1) <> "[" Then
    txtstrSQL.SelText = cmbTable
    txtstrSQL.SetFocus
ElseIf cmbDB <> ThisWorkbook.FullName Then
   txtstrSQL.SelText = "[Excel 12.0;HDR=YES;DATABASE=" & cmbDB & "]." & cmbTable
   txtstrSQL.SetFocus
Else
   txtstrSQL.SelText = cmbTable
   txtstrSQL.SetFocus
End If
End Sub

Private Sub cmdTop_Click()
    txtstrSQL.SelText = "SELECT TOP 20 * " & vbCrLf & "FROM "
    txtstrSQL.SetFocus
End Sub

Private Sub cmdWhere_Click()
     txtstrSQL.SelText = "WHERE "
     txtstrSQL.SetFocus
End Sub

Private Sub cmd¿À¸¥°ýÈ£_Click()
     txtstrSQL.SelText = ") "
     txtstrSQL.SetFocus
End Sub

Private Sub cmd¿Þ°ýÈ£_Click()
     txtstrSQL.SelText = "("
     txtstrSQL.SetFocus
End Sub

Private Sub lbxCol_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
        Dim i As Long
    Dim c As Range
    Dim lastrow_col  As Long
    
        lastrow_col = Sheets("List").Cells(Rows.Count, 3).End(3).Row
    
    For i = 0 To lbxCol.ListCount - 1
        If lbxCol.Selected(i) = True Then
            txtstrSQL.SelText = lbxCol & ", "
            txtstrSQL.SetFocus
        End If
    Next i

End Sub

Private Sub listboxstrSQLCollection_Change()

    Dim i As Long
    Dim c As Range
    Dim lastrow As Long
    
        lastrow = Sheets("strSQL").Cells(Rows.Count, 2).End(3).Row
    
    For i = 0 To listboxstrSQLCollection.ListCount - 1
        If listboxstrSQLCollection.Selected(i) = True Then
            txtstrSQLCollection = listboxstrSQLCollection
                With Sheets("strSQL")
                    Set c = .Range(.Cells(1, 1), .Cells(lastrow, 3)).Find(listboxstrSQLCollection, LookAt:=xlWhole)
                        If Not c Is Nothing Then
                            txtstrSQLCollection = c.Offset(0, 1)
                        End If
                End With
        End If
    Next i

End Sub

Private Sub cmdEXIT_Click()
    Unload Me
End Sub

Private Sub cmdExit2_Click()
    Unload Me
End Sub

