VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMySQL_Convert"
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



' This class converts tables structure and data into MySQL format.
' This will be the only class fully documented for converting since they
' are all very close in structure. Any anomolies or extra functions will
' be documented. But, three classes with similar structure is better than
' a Module with several thousand lines of code loaded. In that case they
' are all loaded whether needed or not.

'ADDED FEATURES FROM MARCO:
' INNODB
' LITTLE IMPROVEMENT IN DATA EXPORT (INSERTS)
' DEFAULT VALUE SUPPORT
' INDEXES, ASCENDING/DESCENDING
' FOREIGN KEYS (INNODB)

' Replaced with OPT_INNODB option - MC
' Const UseInnoDB As Boolean = True 'marco, should be a check for this

Option Explicit

Public Sub MakeMySQL(FileName As String, outPath As String, outFile As String, includeData As Integer, Optional cmdLine As Boolean = False)
    On Error GoTo err_handle
    Dim wrkJet As Workspace
    Dim d As Database
    Dim w As Workspace
    Dim i As Integer
    Dim j As Integer
    Dim Index As Long
    Dim HeaderPrinted As Boolean
    Dim newfile As String
    Dim Path As String
    
    Dim Converted As New Collection
    
    ' Set original offset
    offset = 1
    
    ' Open workspace and database
    Set wrkJet = CreateWorkspace("NewJetWorkspace", "admin", "", dbUseJet)
    Set d = wrkJet.OpenDatabase(FileName, False, True, ";UID=" & DB_User & ";PWD=" & DB_Pass)
    
    ' Create our path
    Path = outPath
    newfile = LCase(outFile & Opt.IO_MySQL_Ext)
    
    FTP_Path = Path
    SQL_Name = newfile
    
    ' Kill the old file if it exists
    Call Kill(Path & newfile)
    
    
    
    EnableForeignKeys False, (Path & newfile) ' MARCO
    
    ' Now, this is the BITCH function
    ' Loop on all our tables
    For i = 0 To (d.TableDefs.count - 1)
        ' Make sure it's not a system table
        If InStr(UCase(d.TableDefs(i).Name), "MSYS") = 0 Then
            ' If it's a linked table, warn the user... but only if they asked us to
            If d.TableDefs(i).SourceTableName <> "" And includeData = vbChecked And (IS_SET(Opt.GenOptions, OPT_WARNLINK) And Not cmdLine) Then
                Dim answer As Integer
                Dim strMessage As String
                strMessage = lang.GetString(341) & " " & Chr(34) & d.TableDefs(i).Name & Chr(34) & " " & lang.GetString(342) & vbCrLf & lang.GetString(343)
                strMessage = strMessage & Chr(13) & lang.GetString(344) & Chr(13) & Chr(13) & lang.GetString(345)
                answer = MsgBox(strMessage, vbInformation + vbYesNo, App.Title)
                ' It was linked, but they still want it converted
                If answer = vbYes Then
                    ' If all tables is checked, convert it right away
                    If IS_SET(Opt.GenOptions, OPT_TABLES) = 1 Then
                        WriteMySQL d.TableDefs(i), (Path & newfile), cmdLine
                        ' If data is requested, convert it as well
                        If includeData = vbChecked Then
                            ConvertDataMySQL d.TableDefs(i), (Path & newfile), cmdLine
                        End If
                    ' Otherwise loop through the tables and see if it's checked off
                    ' in our list
                    Else
                        For j = 0 To (frmTables.lstTables.ListCount - 1)
                            If d.TableDefs(i).Name = frmTables.lstTables.List(j) Then
                                ' Did we find a match, and is it checked?
                                If frmTables.lstTables.Selected(j) = True Then
                                    ' Convert
                                    WriteMySQL d.TableDefs(i), (Path & newfile), cmdLine
                                        
                                    'mantain a list of traslated tables
                                    Converted.Add d.TableDefs(i).Name ' marco
                                    
                                    ' If data is requested, do that too
                                    If includeData = vbChecked Then
                                        ConvertDataMySQL d.TableDefs(i), (Path & newfile), cmdLine
                                    End If
                                    Exit For
                                End If
                                Exit For
                            End If
                        Next j
                    End If
                End If
            ' Ok, in this case they don't want us to warn them
            Else
                ' Are we converting all tables?
                If IS_SET(Opt.GenOptions, OPT_TABLES) = 1 Or cmdLine Then
                    'Do your stuff
                    WriteMySQL d.TableDefs(i), (Path & newfile), cmdLine
                                        
                    'mantain a list of traslated tables
                    Converted.Add d.TableDefs(i).Name ' marco
                    
                    If includeData = vbChecked Then
                        ConvertDataMySQL d.TableDefs(i), (Path & newfile), cmdLine
                    End If
                Else
                    ' Loop through the tables and see if we can find a match
                    For j = 0 To (frmTables.lstTables.ListCount - 1)
                        If d.TableDefs(i).Name = frmTables.lstTables.List(j) Then
                            ' Matched and selected?
                            If frmTables.lstTables.Selected(j) = True Then
                                ' Convert
                                WriteMySQL d.TableDefs(i), (Path & newfile), cmdLine
                                'mantain a list of traslated tables
                                Converted.Add d.TableDefs(i).Name ' marco
                                
                                If includeData = vbChecked Then
                                    ConvertDataMySQL d.TableDefs(i), (Path & newfile), cmdLine
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
    
    On Error GoTo 0
    
    AddForeignKeys Converted, d, (Path & newfile) ' MARCO
    EnableForeignKeys True, (Path & newfile) ' MARCO
    
    ' Clean up
    Set wrkJet = Nothing
    Set d = Nothing
    Exit Sub
err_handle:
    ' Kill didn't find the file to delete, just resume where you left off
    If err.Number = 53 Then
        Resume Next
    Else
        If Not cmdLine Then
            MsgBox err.Number & " : " & err.Description, vbInformation, App.Title
        End If
    End If
    ' Clean up
    Set wrkJet = Nothing
    Set d = Nothing
End Sub




Public Sub EnableForeignKeys(doEnable As Boolean, newfile As String)
        
'___________________________________________________________________
'MARCO
'SET DATABASE FOREIGN_KEY_CHECKS STATE
    'mysql> SET FOREIGN_KEY_CHECKS = 0;
    'mysql> SOURCE dump_file_name
    'mysql> SET FOREIGN_KEY_CHECKS = 1
'___________________________________________________________________

    Dim strOutput As String

    If IS_SET(Opt.GenOptions, OPT_INNODB) Then
        
        On Error Resume Next
        Open LCase(newfile) For Binary As #1
        On Error GoTo 0
        
        strOutput = "SET FOREIGN_KEY_CHECKS = " & IIf(doEnable, "1", "0") & ";" & Chr(10)
        Put #1, offset, strOutput
        offset = offset + Len(strOutput)
    
        Close #1
    Else
        'skip
    End If
End Sub


Public Sub AddForeignKeys(Converted As Collection, myDatabase As Database, newfile As String)

'MARCO
'THIS FUNCTION DOESN'T CHECK FOR EXISTENCE OF INDEXES ON FOREING KEYS ... THEY ARE MANDATORY
' -------------------------------------------------------
' This does not seem to be the case, unfortunately
' Added index checking, le sigh
' - MC

    Dim d As Long
    Dim s As Long
    Dim t As Long
    
    Dim i As Integer
    Dim j As Integer
    
    Dim Stat As String
    
   If IS_SET(Opt.GenOptions, OPT_INNODB) Then
        
        On Error Resume Next
        Open LCase(newfile) For Binary As #1
        On Error GoTo 0
        
        
        For d = 1 To Converted.count
            'some fast coding here
            Stat = ""
            For s = 0 To myDatabase.Relations.count - 1
                With myDatabase.Relations(s)
                    If .Table = Converted(d) Then
                        
                        ' Quick and dirty index check
                        ' This could possibly fail if an index is made up of the foreign key
                        ' fields, plus more
                        
                        Dim gotIndex As Boolean
                        
                        ' Assume failure
                        gotIndex = False
                        
                        ' Check all the indexes
                        For i = 0 To myDatabase.TableDefs(.ForeignTable).Indexes.count - 1
                            
                            Dim cnt As Integer
                            cnt = 0
                            ' Find the specified field name in the index fields
                            For j = 0 To .Fields.count - 1
                                If InStr(1, myDatabase.TableDefs(.ForeignTable).Indexes(i).Fields, .Fields(j).ForeignName) <> 0 Then
                                    cnt = cnt + 1
                                End If
                            Next j
                            
                            ' If the counted fields = foreign key field count, we're good. we found an index
                            If cnt = .Fields.count Then
                                gotIndex = True
                                Exit For
                            End If
                        Next i
                        
                        ' No index found? Make one.
                        If Not gotIndex Then
                        
                            Stat = "ALTER TABLE " & .ForeignTable & " ADD INDEX _" & ParseFieldName(.Name) & "("
                            For t = 0 To .Fields.count - 1
                                Stat = Stat & ParseFieldName(.Fields(t).ForeignName)
                                If t < (.Fields.count - 1) Then Stat = Stat & ", "
                            Next
                            Stat = Stat & ");"
                        
                            Put #1, offset, Stat
                            'Debug.Print Stat
                            
                            offset = offset + Len(Stat)
                        End If
                        
                       ' Begin modify  June 19, 2005 by JCarlos@ceis.cujae.edu.cu

'Original CODE
                        ' We now return you to Marco's code (with some small changes)
'                        Stat = "ALTER TABLE " & ParseFieldName(.Table) & " ADD FOREIGN KEY " & ParseFieldName(.Name) & " ("
'
'                        For t = 0 To .Fields.count - 1
'                            Stat = Stat & ParseFieldName(.Fields(t).Name)
'                            If t < (.Fields.count - 1) Then Stat = Stat & ", "
'                        Next
'
'                        Stat = Stat & ") REFERENCES " & ParseFieldName(.ForeignTable) & " ("
'
'                        For t = 0 To .Fields.count - 1
'                            Stat = Stat & ParseFieldName(.Fields(t).ForeignName)
'                            If t < (.Fields.count - 1) Then Stat = Stat & ", "
'                        Next
'
'                        Stat = Stat & ")"
                        

' Code of JCarlos
                        
                       
                        Stat = "ALTER TABLE " & ParseFieldName(.ForeignTable) & " ADD FOREIGN KEY _" & ParseFieldName(.Name) & " ("
                        For t = 0 To .Fields.count - 1
                            Stat = Stat & ParseFieldName(.Fields(t).ForeignName)
                            If t < (.Fields.count - 1) Then Stat = Stat & ", "
                        Next
                        
                        Stat = Stat & ") REFERENCES " & ParseFieldName(.Table) & " ("
                        
                        For t = 0 To .Fields.count - 1
                            Stat = Stat & ParseFieldName(.Fields(t).Name)
                            If t < (.Fields.count - 1) Then Stat = Stat & ", "
                        Next
                        
                        Stat = Stat & ")"
                        
                        
 ' END modify  June 19, 2005 by JCarlos@ceis.cujae.edu.cu
                        
                        If .Attributes And dbRelationDeleteCascade Then
                            Stat = Stat & Chr(10) & "ON DELETE CASCADE"
                        End If
                        
                        If .Attributes And dbRelationUpdateCascade Then
                            Stat = Stat & Chr(10) & "ON UPDATE CASCADE"
                        End If
                        
                        Stat = Stat & ";" & Chr(10) & Chr(10)
                        Put #1, offset, Stat
                        offset = offset + Len(Stat)
                        
                    End If
                End With
            Next
        Next
    
    
        Close #1
    Else
        'skip
    End If
    
End Sub

' Write the table structure
Private Sub WriteMySQL(d As TableDef, newfile As String, cmdLine As Boolean)
    On Error Resume Next
    
    Open LCase(newfile) For Binary As #1
    
    Dim strOutput As String
    Dim KeyArray() As String
    Dim INDEXARRAY() As String
    Dim KeyCount As Integer
    Dim ICOUNT As Integer
    
    KeyCount = 0
    
    ' Drop the table if it's there
    strOutput = "DROP TABLE IF EXISTS " & ParseFieldName(d.Name) & ";" & Chr(10)
    Put #1, offset, strOutput
    offset = offset + Len(strOutput)
    
    ' Start making the SQL for a new table
    strOutput = "CREATE TABLE " & ParseFieldName(d.Name) & "(" & Chr(10)
    Put #1, offset, strOutput
    offset = offset + Len(strOutput)
        
    Dim i As Integer
    Dim j As Integer
    Dim INDICES As String
    
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
            IndexCount = 1
            startpos = 1
            endpos = 1
            KeyCount = 0
            ICOUNT = 0
            
            If d.Indexes(i).Primary = True Then
                strOutput = String(5, " ") & "PRIMARY KEY ("
                ' split indexes, they come in one big string
                For j = 1 To Len(d.Indexes(i).Fields)
                    If Mid(d.Indexes(i).Fields, j, 1) = ";" Then
                        IndexCount = IndexCount + 1
                    End If
                Next j
                ' Go through the indexes, and make an array
                For j = 1 To IndexCount
                    If IndexCount > 1 And j > 1 Then
                        strOutput = strOutput & ","
                        startpos = endpos + 2
                    Else
                        startpos = endpos
                    End If
                    endpos = InStr(startpos, d.Indexes(i).Fields, ";") - 1
                    If endpos = -1 Then endpos = Len(d.Indexes(i).Fields)
                    ' I'm not exactly sure why we make an array, and then use it
                    ' right away. But I'm scared to change it
                    ReDim Preserve KeyArray(KeyCount)
                    KeyArray(KeyCount) = Mid(d.Indexes(i).Fields, startpos + 1, (endpos - startpos))
                    strOutput = strOutput & ParseFieldName(KeyArray(KeyCount))
                    KeyCount = KeyCount + 1
                Next j
                strOutput = strOutput & ")" & "," & Chr(10)
            Else
            'MARCO
            'NON KEY INDEXES **********************************+
                INDICES = INDICES & String(5, " ") & "INDEX _" & ParseFieldName(d.Indexes(i).Name) & " ("
                ' split indexes, they come in one big string
                For j = 1 To Len(d.Indexes(i).Fields)
                    If Mid(d.Indexes(i).Fields, j, 1) = ";" Then
                        IndexCount = IndexCount + 1
                    End If
                Next j
                ' Go through the indexes, and make an array
                For j = 1 To IndexCount
                    If IndexCount > 1 And j > 1 Then
                        INDICES = INDICES & ","
                        startpos = endpos + 2
                    Else
                        startpos = endpos
                    End If
                    endpos = InStr(startpos, d.Indexes(i).Fields, ";") - 1

                    If endpos = -1 Then endpos = Len(d.Indexes(i).Fields)
                    
                    ReDim Preserve INDEXARRAY(ICOUNT)
                    INDEXARRAY(ICOUNT) = Mid(d.Indexes(i).Fields, startpos + 1, (endpos - startpos))
                    INDICES = INDICES & ParseFieldName(INDEXARRAY(ICOUNT))
                    
                    'ASC/DESC
                    Select Case Mid(d.Indexes(i).Fields, startpos, 1)
                        Case "+": INDICES = INDICES & " ASC"
                        Case "-": INDICES = INDICES & " DESC"
                    End Select
                    
                    ICOUNT = ICOUNT + 1
                    
                Next j
                
                
                'UNIQUE
                If d.Indexes(i).Unique Then
                    INDICES = INDICES & " UNIQUE"
                End If
                    
                
                INDICES = INDICES & "" & ")," & Chr(10)
            End If
            
        Next i
        
        If INDICES > "" Then
            INDICES = Left(INDICES, Len(INDICES) - 2)
        End If
        
        ' If there were no indexes... output the string
        ' This cleared a bug with tables that had no indexes not getting
        ' created... or I think it was that we got two "CREATE TABLE"
        ' lines right after each other
        If InStr(strOutput, "CREATE") = 0 Then
            Put #1, offset, strOutput
            offset = offset + Len(strOutput)
        End If
    End If
 
    ' Loop on our field count
    If d.Fields.count > 0 Then
        For i = 0 To (d.Fields.count - 1)
            strOutput = String(5, " ") & ParseFieldName(d.Fields(i).Name)
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
            
            ' If this field is a key, make sure it's not null
            ' ooooo... this is what the key array was for

            'MARCO
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
                            strOutput = strOutput & " DEFAULT " & d.Fields(i).DefaultValue
                        End If
                    End If
                End If
                
                'TO DO: THERE CAN BE LOCALE CONVERSION ERRORS, CHECK
            End If
            
            ' Add a comma if needed
            If ((i >= 0) And (i < (d.Fields.count - 1))) Then
                strOutput = strOutput & ", "
            Else
                strOutput = strOutput & IIf(INDICES > "", ",", "")
            End If
            
            ' Output the string
            strOutput = strOutput + Chr(10)
            Put #1, offset, strOutput
            offset = offset + Len(strOutput)
            
        Next i
    
        ' Put our finishing brace on, and some CRs (or is that LF, I forget)
        strOutput = INDICES
        strOutput = strOutput & String(5, " ") & ") "
        
        'MARCO, ADD INNODB TYPE
        If IS_SET(Opt.GenOptions, OPT_INNODB) Then
            strOutput = strOutput & Chr(10) & "TYPE = InnoDB "
        Else
            'SKIP
            'USE DEFAULT MYSQL SETTINGS
        End If
        
        strOutput = strOutput & ";" & Chr(10) & Chr(10)
        Put #1, offset, strOutput
        offset = offset + Len(strOutput)
    End If
    
    ' Count the statements for the progress
    currentstate = currentstate + 1
    If Not cmdLine Then
        mainform.prgStates.Value = (currentstate / totalstatements) * 100
        mainform.lblPct.Caption = Str(Int((currentstate / totalstatements) * 100)) & "%"
    End If
    
    DoEvents
    Close #1
End Sub

' Convert our data
Private Sub ConvertDataMySQL(d As TableDef, file As String, cmdLine As Boolean)
    On Error Resume Next
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim strOutput As String
    Dim r As Recordset
    Dim temp As String
    Dim temp2 As Date
    
    Dim num As String
    Dim pos As Long
       
    ' Open our file
    Open file For Binary As #1
    
    ' Open our recordset
    Set r = d.OpenRecordset()
    
    ' Loop on our records
    ' When using a linked table, this comes out as 0
    ' so we can't convert our data properly.
    For j = 0 To (r.RecordCount - 1)
        ' Set the current statement and update our progress
        currentstate = currentstate + 1
        If Not cmdLine Then
            mainform.prgStates.Value = (currentstate / totalstatements) * 100
            mainform.lblPct.Caption = Str(Int((currentstate / totalstatements) * 100)) & "%"
        End If
        DoEvents
        ' Start the insert statement
        strOutput = "INSERT INTO " & ParseFieldName(d.Name) & Chr(10)
        ' Write
        Put #1, offset, strOutput
        offset = offset + Len(strOutput)
        
        strOutput = String(8, " ") & "VALUES ("
        Put #1, offset, strOutput
        offset = offset + Len(strOutput)
        
        ' Loop on each filed

        For k = 0 To (r.Fields.count - 1)
        
            'SI ES NULL -> PONER NULL
            If IsNull(r.Fields(k).Value) Then
                strOutput = "NULL"
            Else
                ' If text of any kind, add quotes
                Select Case r.Fields(k).Type
                Case dbBinary, dbChar, dbGUID, dbLongBinary, dbMemo, dbText, dbVarBinary:
                    
                    temp = "" & r.Fields(k).Value
                    strOutput = Chr(34) & ParseOut(temp) & Chr(34)
                
                
                Case dbBoolean:
                    
                    'GOOD FOR C++ ' If boolean, convert to 0/-1
                    
                    'MYSQL & PHP & VB
                    'TRUE = 1
                    'FALSE =0
                                        
                    If r.Fields(k).Value Then
                        strOutput = 1
                    End If
                    
                    If Not r.Fields(k).Value Then
                        strOutput = 0
                    End If
                    
                    
                Case dbDate, dbTimeStamp:
                    temp2 = r.Fields(k).Value
                    strOutput = Chr(34) & Format(temp2, "yyyy-mm-dd hh:mm:ss") & Chr(34)
                Case dbDate, dbTimeStamp, dbTime:
                    temp2 = r.Fields(k).Value
                    strOutput = Chr(34) & Format(temp2, "hh:mm:ss") & Chr(34)
                Case dbCurrency, dbFloat, dbDouble, dbDecimal:
                    num = CStr(r.Fields(k).Value)
                    pos = InStr(1, num, ",")
                    If pos > 0 Then
                        num = Left(num, pos - 1) & "." & Mid(num, pos + 1)
                    End If
                    strOutput = num
                Case Else
                    strOutput = CStr(r.Fields(k).Value)
                    
                End Select
                
            End If
            
            ' Add a comma if needed
            If ((k >= 0) And (k < (r.Fields.count - 1))) Then
                strOutput = strOutput & ","
            End If

            ' Write the SQL
            Put #1, offset, strOutput
            offset = offset + Len(strOutput)
            
            strOutput = "" 'FIX
            
            ' Add closing brace
            If k = (r.Fields.count - 1) Then
                strOutput = ")" & ";" & Chr(10) & Chr(10)
                Put #1, offset, strOutput
                offset = offset + Len(strOutput)
            End If
        Next k
        ' Move next adn repeat
        r.MoveNext
    Next j
    
    ' Clean up
    Close #1
    Set r = Nothing
End Sub
