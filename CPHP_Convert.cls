VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPHP_Convert"
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
' Header and footer functions are documented (sorta)
' Only other difference is you'll see $query, or $result or other server side
' variables pop up in this function

Option Explicit

Public Sub MakePHP(FileName As String, outPath As String, outFile As String, includeData As Integer)
    On Error Resume Next
    Dim wrkJet As Workspace
    Dim d As Database
    Dim w As Workspace
    Dim i As Integer
    Dim j As Integer
    Dim HeaderPrinted As Boolean
    Dim newfile As String
    Dim Path As String
    
    offset = 1
    
    HeaderPrinted = False
    
    Set wrkJet = CreateWorkspace("NewJetWorkspace", "admin", "", dbUseJet)
    Set d = wrkJet.OpenDatabase(FileName, False, True, ";UID=" & DB_User & ";PWD=" & DB_Pass)
    
    Path = outPath
    newfile = LCase(outFile & Opt.IO_PHP_Ext)
    
    FTP_Path = Path
    PHP_Name = newfile
    
    For i = 0 To (d.TableDefs.count - 1)
        If InStr(UCase(d.TableDefs(i).Name), "MSYS") = 0 Then
            If HeaderPrinted = False Then
                PrintHeader (Path & newfile)
                HeaderPrinted = True
            End If
            If d.TableDefs(i).SourceTableName <> "" And includeData = vbChecked And IS_SET(Opt.GenOptions, OPT_WARNLINK) Then
                Dim answer As Integer
                Dim strMessage As String
                strMessage = lang.GetString(341) & " " & Chr(34) & d.TableDefs(i).Name & Chr(34) & " " & lang.GetString(342) & vbCrLf & lang.GetString(343)
                strMessage = strMessage & Chr(13) & lang.GetString(344) & Chr(13) & Chr(13) & lang.GetString(345)
                answer = MsgBox(strMessage, vbInformation + vbYesNo, App.Title)
                If answer = vbYes Then
                    If IS_SET(Opt.GenOptions, OPT_TABLES) = 1 Then
                        WritePHP d.TableDefs(i), (Path & newfile)
                        If includeData = vbChecked Then
                            ConvertDataPHP d.TableDefs(i), (Path & newfile)
                        End If
                    Else
                        For j = 0 To (frmTables.lstTables.ListCount - 1)
                            If d.TableDefs(i).Name = frmTables.lstTables.List(j) Then
                                If frmTables.lstTables.Selected(j) = True Then
                                    WritePHP d.TableDefs(i), (Path & newfile)
                                    If includeData = vbChecked Then
                                        ConvertDataPHP d.TableDefs(i), (Path & newfile)
                                    End If
                                    Exit For
                                End If
                                Exit For
                            End If
                        Next j
                    End If
                    If i = (d.TableDefs.count - 1) And HeaderPrinted = True Then
                        PrintFooter (Path & newfile)
                    End If
                End If
            Else
                If IS_SET(Opt.GenOptions, OPT_TABLES) = 1 Then
                    WritePHP d.TableDefs(i), (Path & newfile)
                    If includeData = vbChecked Then
                        ConvertDataPHP d.TableDefs(i), (Path & newfile)
                    End If
                Else
                    For j = 0 To (frmTables.lstTables.ListCount - 1)
                        If d.TableDefs(i).Name = frmTables.lstTables.List(j) Then
                            If frmTables.lstTables.Selected(j) = True Then
                                WritePHP d.TableDefs(i), (Path & newfile)
                                If includeData = vbChecked Then
                                    ConvertDataPHP d.TableDefs(i), (Path & newfile)
                                End If
                                Exit For
                            End If
                            Exit For
                        End If
                    Next j
                End If
                If i = (d.TableDefs.count - 1) And HeaderPrinted = True Then
                    PrintFooter (Path & newfile)
                End If
            End If
        End If
    Next i
    
    Set wrkJet = Nothing
    Set d = Nothing
End Sub

' Print the PHP header with connection variables for the database
Private Sub PrintHeader(newfile As String)
    On Error Resume Next
    ' Kill the old file, if it exists
    Kill (newfile)

    Open LCase(newfile) For Binary As #1
    
    Dim strOutput As String
    
    strOutput = "<?php" & Chr(10) & Chr(10) _
                & "     $db=" & Chr(34) & mainform.txtDB.Text & Chr(34) & ";" & Chr(10) _
                & "     $host=" & Chr(34) & mainform.txtHost.Text & Chr(34) & ";" & Chr(10) _
                & "     $user=" & Chr(34) & mainform.txtUser.Text & Chr(34) & ";" & Chr(10) _
                & "     $pass=" & Chr(34) & mainform.txtPassword.Text & Chr(34) & ";" & Chr(10) & Chr(10) _
                & "     $link = mysql_connect($host,$user, $pass) or die(" & Chr(34) & "Unable to connect to DB " & mainform.txtDB.Text & Chr(34) & ");" & Chr(10) _
                & "     mysql_select_db($db, $link);" & Chr(10) & Chr(10)

            
    Put #1, 1, strOutput
    
    offset = offset + Len(strOutput)
    
    Close #1
    
End Sub

' Print off the PHP footer so the PHP daemon knows to stop processing
Private Sub PrintFooter(newfile As String)
    On Error Resume Next
    
    Open newfile For Binary As #1
    
    Dim strOutput As String
    strOutput = "?>"
    Put #1, offset, strOutput
    Close #1
End Sub

Private Sub WritePHP(d As TableDef, newfile As String)
    On Error Resume Next
    
    Open newfile For Binary As #1
    
    Dim strOutput As String
    Dim KeyArray() As String
    Dim KeyCount As Integer
    
    KeyCount = 0
    
    strOutput = "     $query = " & Chr(34) & "DROP TABLE IF EXISTS " & ParseFieldName(d.Name) _
                & Chr(34) & ";" & Chr(10)
    Put #1, offset, strOutput
    offset = offset + Len(strOutput)
    
    strOutput = "     $result = mysql_query($query, $link);" & Chr(10)
    Put #1, offset, strOutput
    offset = offset + Len(strOutput)
    
    strOutput = "     if ($result == FALSE)" & Chr(10) & _
                "     {" & Chr(10) & String(10, " ") & "echo 'Could not drop " & _
                ParseFieldName(d.Name) & ", or unexplained error.\n<br>';" & Chr(10) & "     }" & Chr(10) & Chr(10)
    Put #1, offset, strOutput
    offset = offset + Len(strOutput)
    
    strOutput = "     $query = " & Chr(34) & "CREATE TABLE " _
                & ParseFieldName(d.Name) & "(" & Chr(10)
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
        
        For i = 0 To (d.Indexes.count - 1)
            If d.Indexes(i).Primary = True Then
                strOutput = String(20, " ") & "PRIMARY KEY ("
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
                    KeyArray(KeyCount) = ParseFieldName(Mid(d.Indexes(i).Fields, startpos + 1, (endpos - startpos)))
                    strOutput = strOutput & KeyArray(KeyCount)
                    KeyCount = KeyCount + 1
                Next j
                strOutput = strOutput & ")" & "," & Chr(10)
            End If
        Next i
        
        If InStr(strOutput, "CREATE") = 0 Then
            Put #1, offset, strOutput
            offset = offset + Len(strOutput)
        End If
    End If
 
    
    If d.Fields.count > 0 Then
        For i = 0 To (d.Fields.count - 1)
            strOutput = String(20, " ") & ParseFieldName(d.Fields(i).Name)
            Select Case d.Fields(i).Type
                ' Cases in Microsoft DAO 3.6 help
                ' Conversions based on what I think they should be,
                ' but open for suggestion
                Case dbBigInt
                    strOutput = strOutput & "INT(11)"
                    If d.Fields(i).Attributes >= dbAutoIncrField Then
                        strOutput = strOutput & " AUTO_INCREMENT"
                    End If
                Case dbBinary
                    strOutput = strOutput & " CHAR(8) BINARY"
                Case dbBoolean
                    strOutput = strOutput & " SMALLINT"
                Case dbByte
                    ' Not supported. Use tinyint
                    ' Well, tiny int and byte are the same thing
                    strOutput = strOutput & " TINYINT"
                Case dbChar
                    strOutput = strOutput & " CHAR(" & d.Fields(i).size & ")"
                Case dbCurrency
                    ' Support up to 99 billion (larger than access supports I believe)
                    strOutput = strOutput & " DECIMAL(12,2)"
                Case dbDate
                    strOutput = strOutput & " DATETIME"
                Case dbDecimal
                    strOutput = strOutput & " DECIMAL(12,2)"
                Case dbDouble
                    strOutput = strOutput & " DOUBLE"
                Case dbFloat
                    strOutput = strOutput & " FLOAT"
                Case dbGUID
                    ' Microsft GUI's
                    ' Not really needed, but supported as a 35 byte CHAR
                    strOutput = strOutput & " CHAR(35)"
                Case dbInteger
                    strOutput = strOutput & " INT"
                Case dbLong
                    strOutput = strOutput & " INT(11)"
                    If d.Fields(i).Attributes >= dbAutoIncrField Then
                        strOutput = strOutput & " AUTO_INCREMENT"
                    End If
                Case dbLongBinary
                    ' Not supported
                    ' Use as 16 "bit" binary
                    strOutput = strOutput & " CHAR(16) BINARY"
                Case dbMemo
                    strOutput = strOutput & " LONGBLOB"
                Case dbNumeric
                    strOutput = strOutput & " NUMERIC(12,2)"
                Case dbSingle
                    strOutput = strOutput & " FLOAT"
                Case dbText
                    strOutput = strOutput & " VARCHAR(" & d.Fields(i).size & ")"
                Case dbTime
                    strOutput = strOutput & " TIME"
                Case dbTimeStamp
                    strOutput = strOutput & " TIMESTAMP"
                Case dbVarBinary
                    strOutput = strOutput & " VARCHAR(" & d.Fields(i).size & ") BINARY"
            End Select
            
            'NOT NULL CLAUSE
            If d.Fields(i).Required Then
                        strOutput = strOutput & " NOT NULL"
                    
            Else
                For j = 0 To (KeyCount - 1)
                    If KeyArray(j) = ParseFieldName(d.Fields(i).Name) Then
                        strOutput = strOutput & " NOT NULL"
                    End If
                Next j
            End If
            
            'MARCO
            'Default Value
            If Not IsNull(d.Fields(i).DefaultValue) Then
                If CStr(d.Fields(i).DefaultValue) = "" Then
                
                    'MARCO
                    'ACCORDING TO MY EXPERIENCE IS BETTER TO SET
                    'DEFAULT VALUE TO NULL FOR DATETIME FIELDS IN MIXED MYSQL/ACCESS ENVIROMENTS
                    'MYODBC WORKS BETTER
                    
                    'BETTER TO ADD AN USER CHECK FOR THIS
                    
                    Select Case d.Fields(i).Type
                    Case dbDate
                        strOutput = strOutput & " DEFAULT '0000-00-00 00:00:00'"
                    Case dbTime
                        strOutput = strOutput & " DEFAULT '00:00:00'"
                    Case Else
                        If (InStr(1, strOutput, "AUTO_INCREMENT")) = 0 Then ' Can't add default for auto_inc
                            strOutput = strOutput & " DEFAULT ''"
                        End If
                    End Select
                Else
                    If InStr(1, d.Fields(i).DefaultValue, "=") Then
                        ' special cases
                        ' This should have been a timestamp, but you can blame the db developer
                        ' for crap like this
                        If UCase(d.Fields(i).DefaultValue) = "=DATE()" Then
                            strOutput = strOutput & " DEFAULT 'NOW()'"
                        End If
                    Else
                        If UCase(d.Fields(i).DefaultValue) = "YES" Or UCase(d.Fields(i).DefaultValue) = "TRUE" Then
                            strOutput = strOutput & " DEFAULT '1'"
                        ElseIf UCase(d.Fields(i).DefaultValue) = "NO" Or UCase(d.Fields(i).DefaultValue) = "FALSE" Then
                            strOutput = strOutput & " DEFAULT '0'"
                        Else
                            strOutput = strOutput & " DEFAULT '" & Chr(34) & ". mysql_real_escape_string(" & d.Fields(i).DefaultValue & ") ." & Chr(34) & "'"
                        End If
                    End If
                End If
                
                'TO DO: THERE CAN BE LOCALE CONVERSION ERRORS, CHECK
            End If
            
            If ((i >= 0) And (i < (d.Fields.count - 1))) Then strOutput = strOutput + ", "
            
            strOutput = strOutput + Chr(10)
            Put #1, offset, strOutput
            offset = offset + Len(strOutput)
            
        Next i
    
        strOutput = String(20, " ") & ")" & Chr(34) & ";" & Chr(10)
        Put #1, offset, strOutput
        offset = offset + Len(strOutput)
    End If
    
    
    strOutput = "     $result = mysql_query($query, $link);" & Chr(10)
    Put #1, offset, strOutput
    offset = offset + Len(strOutput)
    
    strOutput = "     if ($result == FALSE)" & Chr(10) & _
                "     {" & Chr(10) & String(10, " ") & "echo mysql_error().'<BR/>';" & Chr(10) & String(10, " ") & "echo " & Chr(34) & "Creation of " & _
                ParseFieldName(d.Name) & " FAILED!!!\n<br>" & Chr(34) & ";" & Chr(10) & "     } else {" & Chr(10) & _
                String(10, " ") & "echo " & Chr(34) & "Creation of " & ParseFieldName(d.Name) & " Succeeded!\n<br>" & Chr(34) & ";" & _
                Chr(10) & "     }"
    Put #1, offset, strOutput
    offset = offset + Len(strOutput)
    
    strOutput = Chr(10) & Chr(10)
    Put #1, offset, strOutput
    offset = offset + Len(strOutput)
    
    currentstate = currentstate + 2
    mainform.prgStates.Value = (currentstate / totalstatements) * 100
    mainform.lblPct.Caption = Str(Int((currentstate / totalstatements) * 100)) & "%"
    DoEvents
    Close #1
End Sub

Private Sub ConvertDataPHP(d As TableDef, file As String)
    On Error Resume Next
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim strOutput As String
    Dim r As Recordset
    
    Open file For Binary As #1
            
    Set r = d.OpenRecordset()
    
    strOutput = "     echo '<br>Inserting into: " & ParseFieldName(r.Name) & "<br>';" & Chr(10)
    Put #1, offset, strOutput
    offset = offset + Len(strOutput)
    
    strOutput = "     $countgood=0;" & Chr(10) & "     $countbad=0;" & Chr(10)
    Put #1, offset, strOutput
    offset = offset + Len(strOutput)
    
    For j = 0 To (r.RecordCount - 1)
        currentstate = currentstate + 2
        mainform.prgStates.Value = (currentstate / totalstatements) * 100
        mainform.lblPct.Caption = Str(Int((currentstate / totalstatements) * 100)) & "%"
        DoEvents
        strOutput = "     $query=" & Chr(34) & "INSERT INTO " & ParseFieldName(d.Name)
        Put #1, offset, strOutput
        offset = offset + Len(strOutput)
        
        strOutput = " VALUES ("
        Put #1, offset, strOutput
        offset = offset + Len(strOutput)
        
        For k = 0 To (r.Fields.count - 1)
            If Not IsNull(r.Fields(k).Value) Then
                If r.Fields(k).Type = dbBinary Or _
                    r.Fields(k).Type = dbChar Or _
                    r.Fields(k).Type = dbGUID Or _
                    r.Fields(k).Type = dbLongBinary Or _
                    r.Fields(k).Type = dbMemo Or _
                    r.Fields(k).Type = dbText Or _
                    r.Fields(k).Type = dbVarBinary Then
                
                    strOutput = "'" & Chr(34) & ". mysql_real_escape_string(" & Chr(34) & ParseOut(r.Fields(k).Value) & Chr(34) & ") . " & Chr(34) & "'"
            
                Else
                    strOutput = r.Fields(k).Value
                End If
            Else
                strOutput = "NULL"
            End If
            
            If ((k >= 0) And (k < (r.Fields.count - 1))) Then
                strOutput = strOutput + ","
            End If
            
            strOutput = strOutput
            Put #1, offset, strOutput
            offset = offset + Len(strOutput)
            
            If k = (r.Fields.count - 1) Then
                strOutput = ")" & Chr(34) & ";" & Chr(10)
                Put #1, offset, strOutput
                offset = offset + Len(strOutput)
                
                strOutput = String(5, " ") & "$result = mysql_query($query, $link);" & Chr(10) & Chr(10)
                Put #1, offset, strOutput
                offset = offset + Len(strOutput)
                
                strOutput = "     if ($result == FALSE)" & Chr(10) & _
                            "     {" & Chr(10) & String(10, " ") & "$countbad+=1;" & Chr(10) & String(10, " ") & "echo mysql_error().'<BR/>';" & Chr(10) & "     } else {" & Chr(10) & _
                            String(10, " ") & "$countgood+=1;" & _
                            Chr(10) & "     }" & Chr(10) & Chr(10)
                Put #1, offset, strOutput
                offset = offset + Len(strOutput)
            End If
        Next k
        If j = (r.RecordCount - 1) Then
            strOutput = "     echo " & Chr(34) & "$countgood records inserted successfully in table " & ParseFieldName(r.Name) & "!<br>$countbad records failed insertion into " & ParseFieldName(r.Name) & "!<br>" & Chr(34) & ";" & Chr(10)
            Put #1, offset, strOutput
            offset = offset + Len(strOutput)
        End If
        r.MoveNext
        
    Next j
    
    Close #1
    Set r = Nothing
End Sub


