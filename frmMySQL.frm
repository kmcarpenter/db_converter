VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMySQL 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "156"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
   Icon            =   "frmMySQL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   5100
   StartUpPosition =   1  'CenterOwner
   Tag             =   "156"
   Begin VB.TextBox txtDB 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      TabIndex        =   2
      ToolTipText     =   "153"
      Top             =   840
      Width           =   2655
   End
   Begin VB.CommandButton cmdUpload 
      Caption         =   "152"
      Height          =   495
      Left            =   3960
      TabIndex        =   5
      Tag             =   "152"
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "132"
      Height          =   495
      Left            =   3960
      TabIndex        =   6
      Tag             =   "132"
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtHost 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      ToolTipText     =   "144"
      Top             =   120
      Width           =   2655
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      ToolTipText     =   "145"
      Top             =   480
      Width           =   2655
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      ToolTipText     =   "154"
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox txtPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   4
      ToolTipText     =   "155"
      Top             =   1560
      Width           =   2655
   End
   Begin MSComctlLib.ProgressBar prgUpload 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   2265
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "151"
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Tag             =   "151"
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblStatus 
      Caption         =   "143"
      Height          =   255
      Left            =   1200
      TabIndex        =   13
      Tag             =   "143"
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "138"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Tag             =   "138"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "139"
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Tag             =   "139"
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "140"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Tag             =   "140"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "141"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Tag             =   "141"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "142"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Tag             =   "142"
      Top             =   1920
      Width           =   1095
   End
End
Attribute VB_Name = "frmMySQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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


Option Explicit

Private Sub cmdCancel_Click()
    ' Kill the form
    Unload Me
End Sub
' Start uploading
Private Sub cmdUpload_Click()
    On Error GoTo err_handle
    
    Dim got_err As Boolean
    
    got_err = False
    
    ' Reset the form
    If cmdUpload.Caption = lang.GetString(308) Then
        txtHost.Enabled = True
        txtUser.Enabled = True
        txtPass.Enabled = True
        txtPort.Enabled = True
        txtDB.Enabled = True
        cmdUpload.Caption = lang.GetString(152)
        Exit Sub
    End If
    ' Check for errors
    If txtPort.Text = "" Then
        txtPort.Text = "3306"
    End If
    If txtHost.Text = "" Then
        MsgBox lang.GetString(310), vbInformation, App.Title
        txtHost.SetFocus
        Exit Sub
    End If
    If txtUser.Text = "" Then
        MsgBox lang.GetString(311), vbInformation, App.Title
        txtUser.SetFocus
        Exit Sub
    End If
    If txtDB.Text = "" Then
        MsgBox lang.GetString(312), vbInformation, App.Title
        txtDB.SetFocus
        Exit Sub
    End If
    
    ' Disable controls so they can't be changed while working
    txtUser.Enabled = False
    txtHost.Enabled = False
    txtPass.Enabled = False
    txtPort.Enabled = False
    txtDB.Enabled = False

    cmdCancel.Enabled = False
    cmdUpload.Enabled = False

    ' Create connection and Recordset objects
    Dim DB As MYSQL_CONNECTION
    Dim RS As MYSQL_RS
    Set DB = New MYSQL_CONNECTION
    
    lblStatus.Caption = lang.GetString(313)
    DoEvents
    
    ' Open connection
    DB.OpenConnection txtHost.Text, txtUser.Text, txtPass.Text, "", txtPort.Text
    
    ' We got an error, kick out
    If DB.error.Number <> 0 Then
        lblStatus.Caption = lang.GetString(314)
        DoEvents
        MsgBox lang.GetString(315) & " " & txtHost.Text, vbInformation, App.Title
        cmdUpload.Caption = lang.GetString(308)
        cmdUpload.Enabled = True
        cmdCancel.Caption = lang.GetString(132)
        cmdCancel.Enabled = True
        Exit Sub
    Else
        ' Select DB specified by user
        DB.SelectDb txtDB.Text
    End If
    
    ' If we got an error, kick out
    If DB.error.Number <> 0 Then
        lblStatus.Caption = lang.GetString(316)
        DoEvents
        MsgBox lang.GetString(317) & " " & txtDB.Text, vbInformation, App.Title
        cmdUpload.Caption = "&Try Again"
        cmdUpload.Enabled = True
        cmdCancel.Caption = lang.GetString(132)
        cmdCancel.Enabled = True
        Exit Sub
    Else
        lblStatus.Caption = lang.GetString(318)
        DoEvents
    End If
    
    Dim endpos As Long
    Dim startpos As Long
    Dim strQuery As String
    Dim data As String
    Dim fname As String
    Dim statements As Long
    Dim curstate As Long
    statements = 0
    curstate = 0

    ' Generate path and open file
    fname = FTP_Path & SQL_Name
    data = String(FileLen(fname), " ")
    Open fname For Binary As #1
    Get #1, , data
    Close #1
    endpos = 0
    
    Kill App.Path & "\ERROR.TXT"
    
    ' Loop and count our SQL statements
    lblStatus.Caption = lang.GetString(319)
    Do
        endpos = endpos + 1
        startpos = endpos
        endpos = InStr(startpos, data, ";")
        If endpos <> 0 Then
            If Mid(data, endpos - 1, 1) <> "\" Then
                statements = statements + 1
                DoEvents
            End If
        End If
    Loop Until endpos = 0
    
    endpos = InStr(1, data, ";")
    startpos = 1
    Dim lstring As String
    Dim rstring As String
    Do
        If Mid(data, endpos - 1, 1) <> "\" Then
            ' Update status (although it won't really be readable as it flashes by)
            strQuery = Mid(data, startpos, (endpos - startpos + 1))
            If UCase(InStr(1, strQuery, "DROP")) <> 0 Then
                lblStatus.Caption = lang.GetString(320)
            ElseIf UCase(InStr(1, strQuery, "CREATE")) <> 0 Then
                lblStatus.Caption = lang.GetString(321)
            ElseIf UCase(InStr(1, strQuery, "INSERT")) <> 0 Then
                lblStatus.Caption = lang.GetString(322)
            ElseIf UCase(InStr(1, strQuery, "ALTER")) <> 0 Then
                lblStatus.Caption = lang.GetString(323)
            End If
            ' Set current statement number and calculate % finished
            curstate = curstate + 1
            prgUpload.Value = (curstate / statements) * 100
            ' Execute query
            Set RS = DB.Execute(strQuery)
            If DB.error.Number <> 0 Then
                Dim ff As Integer
                Dim errStr As String
                Dim offset As Long
                
                got_err = True
                
                If (Dir(App.Path & "\ERROR.TXT") = "") Then
                    offset = 1
                Else
                    offset = FileLen(App.Path & "\ERROR.TXT")
                End If
                ff = FreeFile()
                errStr = "SQL: " & strQuery & vbCrLf & "ErrNo: " & DB.error.Number & vbCrLf & "ErrDesc: " & DB.error.Description & vbCrLf
                
                Select Case DB.error.Number
                    Case 1005
                        errStr = errStr & lang.GetString(324)
                End Select
                
                errStr = errStr & vbCrLf & vbCrLf
                
                Open App.Path & "\ERROR.TXT" For Binary As #ff
                Put #ff, offset, errStr
                Close #ff
            End If
            startpos = endpos + 1
            endpos = InStr(startpos, data, ";")
        Else
            ' I believe this throws away carriage returns and tabs
            lstring = Mid(data, 1, endpos - 2)
            rstring = Mid(data, endpos, Len(data) - endpos + 1)
            data = lstring + rstring
            endpos = InStr(endpos, data, ";")
        End If
        DoEvents
    Loop Until endpos = 0
    
    ' Make sure progress is 100
    prgUpload.Value = 100
    
    If got_err Then
        Dim answer As Integer
        
        answer = MsgBox(lang.GetString(325) & vbCrLf & lang.GetString(326) & vbCrLf & vbCrLf & lang.GetString(327) & vbCrLf & vbCrLf & vbTab & App.Path & "\ERROR.TXT" & vbCrLf & vbCrLf & lang.GetString(328) & vbCrLf & vbCrLf & vbTab & "DB Converter Issues" & vbCrLf & vbTab & "dbconverter@zenwerx.com" & vbCrLf & vbCrLf & lang.GetString(329), vbExclamation + vbYesNo, App.Title)
        If answer = vbYes Then
            ShellExecute Me.hwnd, "open", "wordpad.exe", Chr(34) & App.Path & "\ERROR.TXT" & Chr(34), vbNullString, 1
        End If
    End If
    
    lblStatus.Caption = lang.GetString(305)
    cmdCancel.Caption = lang.GetString(132)
    cmdCancel.Enabled = True
    Exit Sub
err_handle:
    ' Got an error, this one means I can't connect
    If err.Number = 3146 Then
        MsgBox lang.GetString(330), vbInformation, App.Title
    
    ' This means we tried to kill the error log, and it wasn't there
    ElseIf err.Number = 53 Then
        Resume Next '
    
    ' Other type of error. It's baaaaaaaad stuff.
    Else
        Close ' Make sure files are closed
        MsgBox err.Number & ": " & err.Description, vbCritical, App.Title
    End If
    ' Reset form to retry
    lblStatus.Caption = lang.GetString(331)
    cmdUpload.Caption = lang.GetString(308)
    cmdUpload.Enabled = True
    cmdCancel.Caption = lang.GetString(132)
    cmdCancel.Enabled = True
End Sub

Private Sub Form_Activate()
    ' Set host as control with focus
    txtHost.SetFocus
End Sub

Private Sub Form_Load()
    lang.LoadFormStrings Me
    
    ' Load default values from the options object
    txtHost.Text = Opt.SQL_Host
    txtUser.Text = Opt.SQL_User
    txtPass.Text = Opt.SQL_Pass
    txtPort.Text = Opt.SQL_Port
    txtDB.Text = Opt.SQL_DB
End Sub

' Code below this line handles highlighting and return to tab conversion
Private Sub txtHost_KeyPress(KeyAscii As Integer)
    MyTab KeyAscii
End Sub
Private Sub txtHost_GotFocus()
    Highlight txtHost
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    MyTab KeyAscii
End Sub
Private Sub txtPass_GotFocus()
    Highlight txtPass
End Sub

Private Sub txtPort_KeyPress(KeyAscii As Integer)
    MyTab KeyAscii
End Sub
Private Sub txtPort_GotFocus()
    Highlight txtPort
End Sub

Private Sub txtuser_KeyPress(KeyAscii As Integer)
    MyTab KeyAscii
End Sub
Private Sub txtUser_gotfocus()
    Highlight txtUser
End Sub



