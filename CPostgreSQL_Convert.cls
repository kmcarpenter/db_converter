VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPostgreSQL_Convert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' *  Copyright 2002-2005 Michael Carpenter and Zenwerx Custom Programming  *
' ***************************************************************************
' *                                                                         *
' *  Mailing Address:                                                       *
' *                                                                         *
' *  Zenwerx Custom Programming                                             *
' *  c/o Michael Carpenter                                                  *
' *  10 Madison Ave                                                         *
' *  Brantford , Ontario, Canada                                            *
' *  N3T 5X3                                                                *
' *                                                                         *
' ***************************************************************************
' *                                                                         *
' *  Email Address:                                                         *
' *                                                                         *
' *  zenwerx@zenwerx.com                                                    *
' *                                                                         *
' ***************************************************************************
' *                                                                         *
' *  Web Address:                                                           *
' *                                                                         *
' *  http://www.zenwerx.com                                                 *
' *                                                                         *
' ***************************************************************************
'
'    This file is part of DB Converter 1.6.0.0
'
'    DB Converter 1.6.0.0 is free software; you can redistribute it and/or
'    modify it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; either version 2 of the License, or
'    (at your option) any later version.
'
'    DB Converter 1.6.0.0 is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.'
'
'    You should have received a copy of the GNU General Public License
'    along with DB Converter 1.6.0.0; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA


' For full documentation see CMySQL_Convert
' The only difference is in data and SQL structure

Option Explicit

Public Sub MakePostgreSQL(FileName As String, outPath As String, outFile As String, includeData As Integer, Optional cmdLine As Boolean = False)
    On Error GoTo err_handle
    Dim wrkJet As Workspace
    Dim d As Database
    Dim w As Workspace
    Dim i As Integer
    Dim j As Integer
    Dim HeaderPrinted As Boolean
    Dim newfile As String
    Dim Path As String
    
    offset = 1
    
    Set wrkJet = CreateWorkspace("PostgreSQL", "admin", "", dbUseJet)
    Set d = wrkJet.OpenDatabase(FileName, False, True, ";UID=" & DB_User & ";PWD=" & DB_Pass)
    
    Path = outPath
    newfile = LCase(outFile & "_postgre.SQL")
    
    FTP_Path = Path
    Post_Name = newfile
    
    Kill (Path & newfile)
    
    For i = 0 To (d.TableDefs.count - 1)
        If InStr(UCase(d.TableDefs(i).Name), "MSYS") = 0 Then
            If d.TableDefs(i).SourceTableName <> "" And includeData = vbChecked And (IS_SET(Opt.GenOptions, OPT_WARNLINK) And Not cmdLine) Then
                Dim answer As Integer
                Dim strMessage As String
                strMessage = lang.GetString(341) & " " & Chr(34) & d.TableDefs(i).Name & Chr(34) & " " & lang.GetString(342) & vbCrLf & lang.GetString(343)
                strMessage = strMessage & Chr(13) & lang.GetString(344) & Chr(13) & Chr(13) & lang.GetString(345)
                answer = MsgBox(strMessage, vbInformation + vbYesNo, App.Title)
                If answer = vbYes Then
                    If IS_SET(Opt.GenOptions, OPT_TABLES) = 1 Then
                        WritePostgreSQL d.TableDefs(i), (Path & newfile), cmdLine
                        If includeData = vbChecked Then
                            ConvertDataPostgreSQL d.TableDefs(i), (Path & newfile), cmdLine
                        End If
                    Else
                        For j = 0 To (frmTables.lstTables.ListCount - 1)
                            If d.TableDefs(i).Name = frmTables.lstTables.List(j) Then
                                If frmTables.lstTables.Selected(j) = True Then
                                    WritePostgreSQL d.TableDefs(i), (Path & newfile), cmdLine
                                    If includeData = vbChecked Then
                                        ConvertDataPostgreSQL d.TableDefs(i), (Path & newfile), cmdLine
                                    End If
                                    Exit For
                                End If
                                Exit For
                            End If
                        Next j
                    End If
                End If
            Else
                If IS_SET(Opt.GenOptions, OPT_TABLES) = 1 Or cmdLine Then
                    WritePostgreSQL d.TableDefs(i), (Path & newfile), cmdLine
                    If includeData = vbChecked Then
                        ConvertDataPostgreSQL d.TableDefs(i), (Path & newfile), cmdLine
                    End If
                Else
                    For j = 0 To (frmTables.lstTables.ListCount - 1)
                        If d.TableDefs(i).Name = frmTables.lstTables.List(j) Then
                            If frmTables.lstTables.Selected(j) = True Then
                                WritePostgreSQL d.TableDefs(i), (Path & newfile), cmdLine
                                If includeData = vbChecked Then
                                    ConvertDataPostgreSQL d.TableDefs(i), (Path & newfile), cmdLine
                                End If
                                Exit For
                            End If
                            Exit For
                        End If
                    Next j
                End If
            End If
        End If
    Next i
    Set wrkJet = Nothing
    Set d = Nothing
    Exit Sub
err_handle:
    If err.Number = 53 Then
        Resume Next
    Else
        If Not cmdLine Then
            MsgBox err.Number & " : " & err.Description, vbInformation, App.Title
        End If
    End If
    Set wrkJet = Nothing
    Set d = Nothing
End Sub

Private Sub WritePostgreSQL(d As TableDef, newfile As String, cmdLine As Boolean)
    On Error Resume Next
    
    Open LCase(newfile) For Binary As #1
    
    Dim strOutput As String
    Dim KeyArray() As String
    Dim KeyCount As Integer
    
    KeyCount = 0
    
    strOutput = "DROP TABLE IF EXISTS " & ParseFieldName(d.Name) & ";" & Chr(10)
    Put #1, offset, strOutput
    offset = offset + Len(strOutput)
    
    strOutput = "CREATE TABLE " & ParseFieldName(d.Name) & "(" & Chr(10)
    Put #1, offset, strOutput
    offset = offset + Len(strOutput)
        
    Dim i As Integer
    Dim j As Integer
    
    ' Check indexes...basically for primary key
    If d.Indexes.count > 0 Then
        Dim strIndex As String
        Dim IndexCount As Integer
        Dim startpos As Integer
        Dim endpos As Integer
        
        IndexCount = 1
        startpos = 1
        endpos = 1
        
        strOutput = String(5, " ") & "PRIMARY KEY ("
        
        For i = 0 To (d.Indexes.count - 1)
            If d.Indexes(i).Primary = True Then
                ' split indexes
                For j = 1 To Len(d.Indexes(i).Fields)
                    If Mid(d.Indexes(i).Fields, j, 1) = ";" Then
                        IndexCount = IndexCount + 1
                    End If
                Next j
                For j = 1 To IndexCount
                    If IndexCount > 1 And j > 1 Then
                        strOutput = strOutput & ","
                        startpos = endpos + 2
                    Else
                        startpos = endpos
                    End If
                    endpos = InStr(startpos, d.Indexes(i).Fields, ";") - 1
                    If endpos = -1 Then endpos = Len(d.Indexes(i).Fields)
                    ReDim Preserve KeyArray(KeyCount)
                    KeyArray(KeyCount) = Mid(d.Indexes(i).Fields, startpos + 1, (endpos - startpos))
                    strOutput = strOutput & KeyArray(KeyCount)
                    KeyCount = KeyCount + 1
                Next j
                strOutput = strOutput & ")" & "," & Chr(10)
            End If
        Next i
        
        Put #1, offset, strOutput
        offset = offset + Len(strOutput)
    End If
 
    
    If d.Fields.count > 0 Then
        For i = 0 To (d.Fields.count - 1)
            strOutput = String(5, " ") & ParseFieldName(d.Fields(i).Name)
            Select Case d.Fields(i).Type
                ' Cases in Microsoft DAO 3.6 help
                ' Conversions based on what I think they should be,
                ' but open for suggestion
                Case dbBigInt
                    If d.Fields(i).Attributes >= dbAutoIncrField Then
                        strOutput = strOutput & " SERIAL8"
                    Else
                        strOutput = strOutput & " INT8"
                    End If
                Case dbBinary
                    strOutput = strOutput & " CHAR(8)"
                Case dbBoolean
                    strOutput = strOutput & " BOOL"
                Case dbByte
                    ' Not supported. Use tinyint
                    strOutput = strOutput & " CHAR(1)"
                Case dbChar
                    strOutput = strOutput & " CHAR(" & d.Fields(i).size & ")"
                Case dbCurrency
                    ' Support up to 99 billion (larger than access supports I believe)
                    strOutput = strOutput & " DECIMAL(12,2)"
                Case dbDate
                    strOutput = strOutput & " DATE"
                Case dbDecimal
                    strOutput = strOutput & " DECIMAL(12,2)"
                Case dbDouble
                    strOutput = strOutput & " FLOAT8"
                Case dbFloat
                    strOutput = strOutput & " FLOAT4"
                Case dbGUID
                    ' Microsft GUI's
                    ' Not really needed, but supported as a 35 byte CHAR
                    strOutput = strOutput & " CHAR(35)"
                Case dbInteger
                    strOutput = strOutput & " INT"
                Case dbLong
                    strOutput = strOutput & " INT8"
                Case dbLongBinary
                    ' Not supported
                    ' Use as 16 "bit" binary
                    strOutput = strOutput & " CHAR(16)"
                Case dbMemo
                    strOutput = strOutput & " BYTEA"
                Case dbNumeric
                    strOutput = strOutput & " NUMERIC(12,2)"
                Case dbSingle
                    strOutput = strOutput & " FLOAT4"
                Case dbText
                    strOutput = strOutput & " TEXT"
                Case dbTime
                    strOutput = strOutput & " TIME"
                Case dbTimeStamp
                    strOutput = strOutput & " TIMESTAMP"
                Case dbVarBinary
                    strOutput = strOutput & " TEXT"
            End Select
            
            For j = 0 To (KeyCount - 1)
                If KeyArray(j) = ParseFieldName(d.Fields(i).Name) Then
                    strOutput = strOutput & " NOT NULL"
                End If
            Next j
            
            If ((i >= 0) And (i < (d.Fields.count - 1))) Then strOutput = strOutput + ", "
            
            strOutput = strOutput + Chr(10)
            Put #1, offset, strOutput
            offset = offset + Len(strOutput)
            
        Next i
    
        strOutput = String(5, " ") & ")" & ";" & Chr(10) & Chr(10)
        Put #1, offset, strOutput
        offset = offset + Len(strOutput)
    End If
    currentstate = currentstate + 1
    If Not cmdLine Then
        mainform.prgStates.Value = (currentstate / totalstatements) * 100
        mainform.lblPct.Caption = Str(Int((currentstate / totalstatements) * 100)) & "%"
    End If
    DoEvents
    Close #1
End Sub

Private Sub ConvertDataPostgreSQL(d As TableDef, file As String, cmdLine As Boolean)
    On Error Resume Next
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim strOutput As String
    Dim r As Recordset
    Dim temp As String
    Dim temp2 As Date
       
    
    Open file For Binary As #1

    Set r = d.OpenRecordset()
    
    For j = 0 To (r.RecordCount - 1)
        currentstate = currentstate + 1
        If Not cmdLine Then
            mainform.prgStates.Value = (currentstate / totalstatements) * 100
            mainform.lblPct.Caption = Str(Int((currentstate / totalstatements) * 100)) & "%"
        End If
        DoEvents
        strOutput = "INSERT INTO " & ParseFieldName(d.Name) & Chr(10)
        Put #1, offset, strOutput
        offset = offset + Len(strOutput)
        
        strOutput = String(8, " ") & "VALUES ("
        Put #1, offset, strOutput
        offset = offset + Len(strOutput)
        
        For k = 0 To (r.Fields.count - 1)
            
            If r.Fields(k).Type = dbBinary Or _
               r.Fields(k).Type = dbChar Or _
               r.Fields(k).Type = dbGUID Or _
               r.Fields(k).Type = dbLongBinary Or _
               r.Fields(k).Type = dbMemo Or _
               r.Fields(k).Type = dbText Or _
               r.Fields(k).Type = dbVarBinary Then
                
                temp = "" & r.Fields(k).Value
                strOutput = Chr(34) & ParseOut(temp) & Chr(34)
            ElseIf r.Fields(k).Type = dbBoolean Then
                If UCase(r.Fields(k).Value) = "FALSE" Then
                    strOutput = 0
                Else
                    strOutput = -1
                End If
            ElseIf r.Fields(k).Type = dbDate Then
                temp2 = r.Fields(k).Value
                strOutput = Chr(34) & Format(temp2, "yyyy-mm-dd") & Chr(34)
            Else
                strOutput = r.Fields(k).Value
            End If
            
                
            If ((k >= 0) And (k < (r.Fields.count - 1))) Then
                strOutput = strOutput + ","
            End If

            Put #1, offset, strOutput
            offset = offset + Len(strOutput)
            
            If k = (r.Fields.count - 1) Then
                strOutput = ")" & ";" & Chr(10) & Chr(10)
                Put #1, offset, strOutput
                offset = offset + Len(strOutput)
            End If
        Next k
        r.MoveNext
    Next j
    
    Close #1
    Set r = Nothing
End Sub
