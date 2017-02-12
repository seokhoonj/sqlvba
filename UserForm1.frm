VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "INSERT strSQL"
   ClientHeight    =   9583
   ClientLeft      =   42
   ClientTop       =   378
   ClientWidth     =   10563
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbDB_Change()

    Dim i As Long
    Dim lastrow_tbl As Long
    Dim adoConn As ADODB.Connection
    Dim adoRS As ADODB.Recordset
    
    lastrow_tbl = Sheets("List").Cells(rows.Count, 2).End(3).row
    
    With Sheets("List")
        .Range(.Cells(2, 2), .Cells(lastrow_tbl, 2)).ClearContents
    End With
    
    
       Set adoConn = New ADODB.Connection

   ' Open the connection.
   With adoConn
      .Provider = "Microsoft.ACE.OLEDB.12.0"
      .Open "C:\myDB\" & cmbDB & ".accdb"
   End With

   ' Open the tables schema rowset.
   Set adoRS = adoConn.OpenSchema(adSchemaTables)

   ' Loop through the results and print the
   ' names and types in the Immediate pane.
   
   i = 2
   With adoRS
      Do While Not .EOF
      lastrow_tbl = Sheets("List").Cells(rows.Count, 2).End(3).row
         If .Fields("TABLE_TYPE") <> "VIEW" And Left(.Fields("TABLE_NAME"), 4) <> "MSys" Then
               Sheets("List").Cells(lastrow_tbl + 1, 2) = .Fields("TABLE_NAME") '& vbTab & .Fields("TABLE_TYPE")
         End If
         .MoveNext
      Loop
   End With
   adoRS.Close
   adoConn.Close
   
   Set adoRS = Nothing
   Set adoConn = Nothing
   
   lastrow_tbl = Sheets("List").Cells(rows.Count, 2).End(3).row
        With Sheets("List")
                cmbTable.Clear
                    For i = 2 To lastrow_tbl
                            cmbTable.AddItem .Cells(i, 2)
                    Next i
        End With
            
End Sub
Private Sub cmbTable_Change()
        
  Dim adoConn As New ADODB.Connection
  Dim adoRS As ADODB.Recordset
  Dim strConn As String
  Dim strSQL As String
  Dim i As Long, lastrow_col As Long
  
      lastrow_col = Sheets("List").Cells(rows.Count, 3).End(3).row
    
    With Sheets("List")
        .Range(.Cells(2, 3), .Cells(lastrow_col, 3)).ClearContents
    End With

        Set adoConn = New ADODB.Connection
        Set adoRS = New ADODB.Recordset
             
        strConn = "PROVIDER=Microsoft.ACE.OLEDB.12.0; " & _
                 "DATA SOURCE=" & _
                 "C:\myDB\" & _
                 cmbDB & ".accdb"
                            
        adoConn.Open strConn
        
     strSQL = "SELECT top 1 * FROM " & cmbTable
     If adoConn.State = adStateOpen Then
            adoRS.Open strSQL, strConn, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not adoRS.EOF Then
                 With Sheets("List")
                        For i = 0 To adoRS.Fields.Count - 1
                           .Cells(i + 2, 3).value = adoRS.Fields(i).Name
                        Next i
                End With
            Else
                MsgBox "자료가 없습니다", 64, "데이터 오류"
            End If
    adoRS.Close
    Else
    End If
    adoConn.Close
    
    Set adoConn = Nothing
    Set adoRS = Nothing

    lastrow_col = Sheets("List").Cells(rows.Count, 3).End(3).row
        With Sheets("List")
            lbxCol.Clear
                    For i = 2 To lastrow_col
                            lbxCol.AddItem .Cells(i, 3)
                    Next i
        End With
             
End Sub


Private Sub cmdDesc_Click()
    txtstrSQL.SelText = "DESC "
    txtstrSQL.SetFocus
End Sub

Private Sub cmdDrop_Click()
    txtstrSQL.SelText = "DROP TABLE "
    txtstrSQL.SetFocus
    Sheets("Query").Cells(13, 1).ClearContents
End Sub

Private Sub cmdHaving_Click()
    txtstrSQL.SelText = "HAVING "
    txtstrSQL.SetFocus
End Sub

Private Sub cmdInto_Click()
    txtstrSQL.SelText = "SELECT * INTO " & vbCr & "FROM [Excel 8.0;HDR=YES;DATABASE=C:\Users\JOO\Desktop\Analysis.xlsb].[DB$]" & vbCr & "WHERE 일자 >= 0"
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
    txtstrSQL.SelText = "ORDER BY "
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
    txtstrSQL.SelText = "and "
    txtstrSQL.SetFocus
End Sub

Private Sub cmdAs_Click()
    txtstrSQL.SelText = "as "
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
         txtstrSQL.SelText = "Count"
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
     txtstrSQL.SelText = "or "
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
    
            lastrow = Sheets("strSQL").Cells(rows.Count, 3).row
    
        With Sheets("strSQL")
            .Range(.Cells(2, 3), .Cells(lastrow, 3)).Replace What:=txtOld, Replacement:=txtNew, LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False
        End With
        
End Sub

Private Sub cmdSum_Click()
    txtstrSQL.SelText = "Sum"
    txtstrSQL.SetFocus
End Sub

Private Sub cmdTable_Click()
    txtstrSQL.SelText = cmbTable
    txtstrSQL.SetFocus
End Sub

Private Sub cmdTop_Click()
     txtstrSQL.SelText = "SELECT Top 20 * " & vbCr & "FROM  "
     txtstrSQL.SetFocus
End Sub

Private Sub cmdWhere_Click()
     txtstrSQL.SelText = "WHERE "
     txtstrSQL.SetFocus
End Sub

Private Sub cmd오른괄호_Click()
     txtstrSQL.SelText = ") "
     txtstrSQL.SetFocus
End Sub

Private Sub cmd왼괄호_Click()
     txtstrSQL.SelText = "("
     txtstrSQL.SetFocus
End Sub



Private Sub lbxCol_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim i As Long
    Dim c As Range
    Dim lastrow_col  As Long
    
        lastrow_col = Sheets("List").Cells(rows.Count, 3).End(3).row
    
    For i = 0 To lbxCol.ListCount - 1
        If lbxCol.Selected(i) = True Then
            txtstrSQL.SelText = lbxCol & ", "
            txtstrSQL.SetFocus
        End If
    Next i

End Sub

Private Sub MultiPage1_Change()

End Sub

Private Sub txtstrSQL_Change()

End Sub

Private Sub UserForm_Initialize()

    Dim i As Long
    Dim lastrow_strSQL As Long
    Dim lastrow_DBList As Long
    Dim adoConn As ADODB.Connection
    Dim adoRS As ADODB.Recordset
    
        lastrow_strSQL = Sheets("strSQL").Cells(rows.Count, 2).End(3).row
        lastrow_DBList = Sheets("List").Cells(rows.Count, 1).End(3).row

    
    For i = 2 To lastrow_strSQL
        With Sheets("strSQL")
            listboxstrSQLCollection.AddItem .Cells(i, 2)
        End With
    Next i


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        With Sheets("List")
                    For i = 2 To lastrow_DBList
                            cmbDB.AddItem .Cells(i, 1)
                    Next i
        End With
        
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Set adoConn = New ADODB.Connection

   ' Open the connection.
   With adoConn
      .Provider = "Microsoft.ACE.OLEDB.12.0"
      .Open "C:\myDB\DB.accdb"
   End With

   ' Open the tables schema rowset.
   Set adoRS = adoConn.OpenSchema(adSchemaTables)

   ' Loop through the results and print the
   ' names and types in the Immediate pane.
   
   i = 2
   With adoRS
      Do While Not .EOF
         If .Fields("TABLE_TYPE") <> "VIEW" And Left(.Fields("TABLE_NAME"), 4) <> "MSys" Then
               Sheets("List").Cells(i, 2) = .Fields("TABLE_NAME") '& vbTab & .Fields("TABLE_TYPE")
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
        cmbDB = .Cells(10, 1)
        cmbTable = .Cells(13, 1)
        txtstrSQL = .Cells(1, 1)
    End With
                
End Sub
Private Sub listboxstrSQLCollection_Change()

    Dim i As Long
    Dim c As Range
    Dim lastrow As Long
    
        lastrow = Sheets("strSQL").Cells(rows.Count, 2).End(3).row
    
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
                 "DATA SOURCE=" & _
                 "C:\myDB\" & _
                 cmbDB & ".accdb"
                            
        adoConn.Open strConn

        adoConn.Execute (txtstrSQL)
        
        adoRS.Close
        adoConn.Close
        Set adoConn = Nothing
        Set adoRS = Nothing
        MsgBox "        처리되었습니다."
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
                 "DATA SOURCE=" & _
                 "C:\myDB\" & _
                 cmbDB & ".accdb"
                            
        adoConn.Open strConn
        
     If adoConn.State = adStateOpen Then
            adoRS.Open strSQL, strConn, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not adoRS.EOF Then
                 With Sheets("Query")
                    .Range("a30").CurrentRegion.ClearContents
                        For i = 0 To adoRS.Fields.Count - 1
                           .Cells(30, i + 1).value = adoRS.Fields(i).Name
                        Next i
                    .Range("a31").CopyFromRecordset adoRS
                    .Range("B:ZZ").Columns.AutoFit
                    .Activate
                End With
            Else
                MsgBox "자료가 없습니다", 64, "데이터 오류"
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
                     "Data Source=                " & ThisWorkbook.FullName & ";" & _
                     "Extended Properties=   Excel 12.0;"
                     
        adoRS.Open strSQL, strConn, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not adoRS.EOF Then
                 With Sheets("Query")
                    .Range("a30").CurrentRegion.ClearContents
                        For i = 0 To adoRS.Fields.Count - 1
                           .Cells(30, i + 1).value = adoRS.Fields(i).Name
                        Next i
                    .Range("a31").CopyFromRecordset adoRS
                    .Range("B:ZZ").Columns.AutoFit
                    .Activate
                End With
            Else
                MsgBox "자료가 없습니다", 64, "데이터 오류"
            End If
    adoRS.Close
    Set adoRS = Nothing
    
End If

Cells(31, 1).Select
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
                 "DATA SOURCE=" & _
                 "C:\myDB\" & _
                 cmbDB & ".accdb"
                            
        adoConn.Open strConn

        adoConn.Execute (txtstrSQL)
        
        adoRS.Close
        adoConn.Close
        Set adoConn = Nothing
        Set adoRS = Nothing
        MsgBox "        처리되었습니다."
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
                 "DATA SOURCE=" & _
                 "C:\myDB\" & _
                 cmbDB & ".accdb"
                            
        adoConn.Open strConn
        
     If adoConn.State = adStateOpen Then
            adoRS.Open strSQL, strConn, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not adoRS.EOF Then
                 With Sheets("Query")
                    .Range("a30").CurrentRegion.ClearContents
                        For i = 0 To adoRS.Fields.Count - 1
                           .Cells(30, i + 1).value = adoRS.Fields(i).Name
                        Next i
                    .Range("a31").CopyFromRecordset adoRS
                    .Range("B:ZZ").Columns.AutoFit
                    .Activate
                End With
            Else
                MsgBox "자료가 없습니다", 64, "데이터 오류"
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
                     "Data Source=                " & ThisWorkbook.FullName & ";" & _
                     "Extended Properties=   Excel 12.0;"
                     
        adoRS.Open strSQL, strConn, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not adoRS.EOF Then
                 With Sheets("Query")
                    .Range("a30").CurrentRegion.ClearContents
                        For i = 0 To adoRS.Fields.Count - 1
                           .Cells(30, i + 1).value = adoRS.Fields(i).Name
                        Next i
                    .Range("a31").CopyFromRecordset adoRS
                    .Range("B:ZZ").Columns.AutoFit
                    .Activate
                End With
            Else
                MsgBox "자료가 없습니다", 64, "데이터 오류"
            End If
    adoRS.Close
    Set adoRS = Nothing
    
End If

Cells(31, 1).Select
End Sub
