VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFTP 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "157"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   Icon            =   "frmFTP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   5115
   StartUpPosition =   1  'CenterOwner
   Tag             =   "157"
   Begin VB.Timer tmr 
      Interval        =   1
      Left            =   4380
      Top             =   1440
   End
   Begin MSWinsockLib.Winsock pasv 
      Left            =   4380
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1020
      PasswordChar    =   "*"
      TabIndex        =   12
      ToolTipText     =   "147"
      Top             =   1200
      Width           =   2775
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   1020
      TabIndex        =   11
      ToolTipText     =   "146"
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   1020
      TabIndex        =   10
      ToolTipText     =   "145"
      Top             =   480
      Width           =   2775
   End
   Begin VB.TextBox txtHost 
      Height          =   285
      Left            =   1020
      TabIndex        =   9
      ToolTipText     =   "144"
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "125"
      Height          =   495
      Left            =   3900
      TabIndex        =   2
      Tag             =   "125"
      ToolTipText     =   "150"
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpload 
      Caption         =   "149"
      Height          =   495
      Left            =   3900
      TabIndex        =   1
      Tag             =   "149"
      ToolTipText     =   "148"
      Top             =   120
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar prgUpload 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   0
      Top             =   1992
      Width           =   5112
      _ExtentX        =   9022
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSWinsockLib.Winsock ftpUpload 
      Left            =   4860
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   21
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "142"
      Height          =   255
      Left            =   60
      TabIndex        =   8
      Tag             =   "142"
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "141"
      Height          =   255
      Left            =   60
      TabIndex        =   7
      Tag             =   "141"
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "140"
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Tag             =   "140"
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "139"
      Height          =   255
      Left            =   60
      TabIndex        =   5
      Tag             =   "139"
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "138"
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Tag             =   "138"
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblStatus 
      Caption         =   "143"
      Height          =   252
      Left            =   1020
      TabIndex        =   3
      Tag             =   "143"
      Top             =   1560
      Width           =   2772
   End
End
Attribute VB_Name = "frmFTP"
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



' This is old old code. I'm not exactly sure anyone uses this
' I think I'm going to add a survey into the next version to
' see if I should remove this in order to improve size/performance

' Open an FTP form. Connect and upload your converted sql/php scripts
Option Explicit

' form variables
Dim Path As String
Dim DoSend As Boolean
Dim FileCount As Integer
Dim FileName As String

' boolean flags (could probably be changed to bytes)
Dim sentSQL As Boolean
Dim sentPHP As Boolean
Dim sentOra As Boolean
Dim sentPost As Boolean

' For timeouts if no internet connection or ftp is down
Dim timeout_possible As Boolean

' Totals for %
Dim totalsent As Long
Dim totalbytes As Long

' Times
Dim stime As Long
Dim ltime As Long

' Error flag
Dim error As Boolean

Private Sub cmdUpload_Click()
    ' Error checking
    If txtHost.Text = "" Then
        MsgBox "Remote host needed for upload!", vbInformation, App.Title
        txtHost.SetFocus
        Exit Sub
    ElseIf txtPort.Text = "" Then
        txtPort.Text = "21"
    ElseIf txtUser.Text = "" Then
        MsgBox "Remote user needed for upload!", vbInformation, App.Title
        txtUser.SetFocus
        Exit Sub
    ElseIf txtPass.Text = "" Then
        MsgBox "Remote password needed for upload!", vbInformation, App.Title
        txtPass.SetFocus
        Exit Sub
    End If
    
    ' Set default values
    error = False
    sentSQL = False
    sentPHP = False
    sentOra = False
    sentPost = False
    totalbytes = 0
    totalsent = 0
    prgUpload.Value = 0
    stime = Timer()
    lblStatus.Caption = "Connecting..."
    cmdUpload.Enabled = False
    cmdCancel.Enabled = False
    FileCount = 0
    
    txtUser.Enabled = False
    txtHost.Enabled = False
    txtPass.Enabled = False
    txtPort.Enabled = False
    
    ' Sum total bytes
    If mainform.chkMySQL.Value = vbChecked Then
        totalbytes = totalbytes + FileLen(FTP_Path & SQL_Name)
    End If
    If mainform.chkOracle.Value = vbChecked Then
        totalbytes = totalbytes + FileLen(FTP_Path & Ora_Name)
    End If
    If mainform.chkPHP.Value = vbChecked Then
        totalbytes = totalbytes + FileLen(FTP_Path & PHP_Name)
    End If
    If mainform.chkPost.Value = vbChecked Then
        totalbytes = totalbytes + FileLen(FTP_Path & Post_Name)
    End If
    
    ' Set timeout flag
    timeout_possible = True
    
    ' Close connections, if open and set the input values
    pasv.Close
    ftpUpload.Close
    ftpUpload.RemoteHost = txtHost.Text
    ftpUpload.RemotePort = txtPort.Text
    ftpUpload.Connect
    lblStatus.Caption = "Connecting"
End Sub

Private Sub cmdCancel_Click()
    ' unload form
    Unload Me
End Sub

Private Sub Form_Activate()
    ' set default focus
    txtHost.SetFocus
End Sub

Private Sub Form_Load()
        
    lang.LoadFormStrings Me
    
    ' Load default values from the options object
    txtHost.Text = Opt.FTP_Host
    txtPass.Text = Opt.FTP_Pass
    txtUser.Text = Opt.FTP_User
    txtPort.Text = Opt.FTP_Port
    
    ' Make sure sending flags are off
    sentSQL = False
    sentPHP = False
    sentOra = False
    sentPost = False
    
    DoSend = False
    FileCount = 0
End Sub

Private Sub ftpUpload_DataArrival(ByVal bytesTotal As Long)
    ' Meat and potatoes.
    ' Declare the sub variables
    On Error GoTo err_handle
    Dim strData As String
    Dim strSend As String
    Dim strPasvPort As String
    Dim startpos As Integer
    Dim endpos As Integer
    
    ' If we got data, no more timeouts possible
    timeout_possible = False
    
    ' Get the data from the winsock control
    ftpUpload.GetData strData, vbString
    
    ' Server asked for user info
    If Mid(strData, 1, 3) = "220" Then
        lblStatus.Caption = lang.GetString(298)
        strSend = "USER " & txtUser.Text & Chr(10)
        ftpUpload.SendData strSend
    ' Server asked for password
    ElseIf Mid(strData, 1, 3) = "331" Then
        strSend = "PASS " & txtPass.Text & Chr(10)
        ftpUpload.SendData strSend
    ' Server asked for something. I don't remember!
    ' Looks like we're sending it a request for "present working directory"
    ElseIf Mid(strData, 1, 3) = "230" Then
        strSend = "PWD" & Chr(10)
        ftpUpload.SendData strSend
    ' Set our transfer mode to passive
    ElseIf Mid(strData, 1, 3) = "257" Then
        lblStatus.Caption = lang.GetString(299)
        startpos = InStr(1, strData, Chr(34))
        endpos = InStr(startpos + 1, strData, Chr(34))
        Path = Mid(strData, startpos + 1, endpos - startpos - 1)
        strSend = "PASV" & Chr(10)
        ftpUpload.SendData strSend
    ' Everything seems ok. Find out what port the server wants us to use
    ElseIf Mid(strData, 1, 3) = "227" Then
        Dim i As Integer
        startpos = 0
        endpos = 0
        For i = 1 To 4
            startpos = InStr(startpos + 1, strData, ",")
            If i = 4 Then
                ' Get the piece of data we need
                endpos = InStr(startpos + 1, strData, ",")
                strPasvPort = Mid(strData, startpos + 1, (endpos - startpos - 1))
                startpos = endpos
                endpos = InStr(startpos + 1, strData, ")")
                ' Decode the port
                strPasvPort = Str(Val(strPasvPort) * 256 + Val(Mid(strData, startpos + 1, (endpos - startpos + 1))))
                ' Connect to the port the server is waiting for us to connect on
                pasv.Close
                pasv.RemoteHost = txtHost.Text
                pasv.RemotePort = strPasvPort
                pasv.Connect
                ' Set transfer mode to Ascii
                strSend = "TYPE A" & Chr(10)
                ftpUpload.SendData strSend
            End If
        Next i
    ' We're good and connected. Set our send flag
    ElseIf Mid(strData, 1, 3) = "200" Then
        DoSend = True
    ' Got Mark - Send file
    ElseIf Val(Mid(strData, 1, 3)) >= 100 And Val(Mid(strData, 1, 3)) <= 199 Then
        sendfile
    ' File is finished
    ElseIf Mid(strData, 1, 3) = "226" Then
        lblStatus.Caption = lang.GetString(300)
    ' Got a user/password error
    ElseIf Mid(strData, 1, 3) = "530" Then
        MsgBox lang.GetString(301), vbInformation, App.Title
        GoTo err_handle:
    ' Some other funky error
    ElseIf Mid(strData, 1, 3) = "550" Then
        lblStatus.Caption = lang.GetString(302)
        MsgBox lang.GetString(303), vbInformation, App.Title
        GoTo err_handle
    End If
    Exit Sub
err_handle:
    ' Handle errors
    lblStatus.Caption = lang.GetString(304)
    cmdCancel.Enabled = True
    cmdUpload.Enabled = True
    txtUser.Enabled = True
    txtHost.Enabled = True
    txtPass.Enabled = True
    txtPort.Enabled = True
    prgUpload.Value = 0
    DoSend = False
    FileCount = 0
    pasv.Close
    ftpUpload.Close
End Sub

Private Sub ChooseFiles()
    ' Depending on what you converted (the checkboxes in the main form)
    ' This sub picks which files you want to send and then tells
    ' the ftp server to "STOR" (store) them
    
    ' I believe filecount is more of a flag than an actual count
    Dim strSend As String
    If mainform.chkMySQL.Value = vbChecked And sentSQL = False Then
        strSend = "STOR " & SQL_Name & Chr(10)
        ftpUpload.SendData strSend
        sentSQL = True
        FileCount = 1
        Exit Sub
    Else
        sentSQL = True
    End If
    If mainform.chkOracle.Value = vbChecked And sentOra = False Then
        strSend = "STOR " & Ora_Name & Chr(10)
        ftpUpload.SendData strSend
        sentOra = True
        FileCount = 2
        Exit Sub
    Else
        sentOra = True
    End If
    If mainform.chkPHP.Value = vbChecked And sentPHP = False Then
        strSend = "STOR " & PHP_Name & Chr(10)
        ftpUpload.SendData strSend
        sentPHP = True
        FileCount = 3
        Exit Sub
    Else
        sentPHP = True
    End If
    If mainform.chkPost.Value = vbChecked And sentPost = False Then
        strSend = "STOR " & Post_Name & Chr(10)
        ftpUpload.SendData strSend
        sentPost = True
        FileCount = 4
        Exit Sub
    Else
        sentPost = True
    End If
    FileCount = 5
End Sub

Private Sub pasv_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    ' Update our little progress bar
    Dim progress As Double
    totalsent = totalsent + bytesSent
    progress = (totalsent / totalbytes) * 100
    If progress > 100 Then
        progress = 100
    End If
    prgUpload.Value = progress
    If bytesRemaining <= 0 Then
        pasv.Close
        If FileCount >= 4 Then
            ftpUpload.Close
            prgUpload.Value = 100
            cmdCancel.Caption = lang.GetString(132)
            cmdCancel.Enabled = True
            lblStatus.Caption = lang.GetString(305)
            FileCount = 0
        Else
            ftpUpload.SendData "PASV" & Chr(10)
        End If
    End If
End Sub

' This sub is required to check our status with the FTP server
' The winsock object doesn't actually generate any events unless
' it is sending or receiving data. If it's in an idle state waiting
' for us, we need to figure out what it's doing and what we should be doing
Private Sub tmr_Timer()
    ' If we're sending, check what we're doing
    If DoSend = True Then
        If FileCount < 4 Then
            ChooseFiles
        End If
        DoSend = False
    End If
    ' If we got an error, set the flag and close the connections
    If error = True Then
        error = False
        pasv.Close
        If sentPHP = True Then
            sentPHP = False
        ElseIf sentOra = True Then
            sentOra = False
        ElseIf sentSQL = True Then
            sentSQL = False
        End If
        ftpUpload.SendData "PASV " & Chr(10)
    End If
    ' If there's a possible timeout, check what status we're in
    If timeout_possible = True Then
        ' If we're connected or disconnected, get out
        If ftpUpload.State = 7 Or ftpUpload.State = 0 Then
            timeout_possible = False
        Else
            ' If the timer is over 120 seconds (wow long timeout)
            ' stop doing anything
            If Timer > (stime + 120) Then
                MsgBox lang.GetString(306), vbInformation, App.Title
                lblStatus.Caption = lang.GetString(304)
                cmdUpload.Enabled = True
                cmdCancel.Enabled = True
                txtUser.Enabled = True
                txtHost.Enabled = True
                txtPass.Enabled = True
                txtPort.Enabled = True
                prgUpload.Value = 0
                DoSend = False
                FileCount = 0
                pasv.Close
                ftpUpload.Close
            Else
                ' Update the status bar so that the user knows we're doing something
                If Timer() > (ltime + 1) Then
                    lblStatus.Caption = lblStatus.Caption & "."
                    ltime = Timer()
                End If
            End If
        End If
    End If
    ' This has something to do with the progress bar
    If FileCount > 4 Then
        pasv_SendProgress 0, 0
    End If
End Sub

Private Sub sendfile()
    On Error GoTo err_handle:
    
    ' If we're in here and we're not connected, something is wrong
    If pasv.State <> 7 Then
        error = True
        Exit Sub
    End If
    
    ' Pick which file to send and update the status bar
    If FileCount = 1 Then
        FileName = FTP_Path & SQL_Name
        lblStatus.Caption = lang.GetString(307) & " " & Chr(34) & SQL_Name & Chr(34) & "..."
    ElseIf FileCount = 2 Then
        FileName = FTP_Path & Ora_Name
        lblStatus.Caption = lang.GetString(307) & " " & Chr(34) & Ora_Name & Chr(34) & "..."
    ElseIf FileCount = 3 Then
        FileName = FTP_Path & PHP_Name
        lblStatus.Caption = lang.GetString(307) & " " & Chr(34) & PHP_Name & Chr(34) & "..."
    ElseIf FileCount = 4 Then
        FileName = FTP_Path & Post_Name
        lblStatus.Caption = lang.GetString(307) & " " & Chr(34) & Post_Name & Chr(34) & "..."
    End If
    ' Process any incoming events
    DoEvents
    ' Open the file and send the data to the server
    Open FileName For Binary As #1
    Dim data As String
    data = String(FileLen(FileName), " ")
    Get #1, , data
    Close #1
    pasv.SendData data
    Exit Sub
    
    ' We got an error, bail! bail!
err_handle:
    lblStatus.Caption = lang.GetString(304)
    cmdUpload.Enabled = True
    cmdCancel.Enabled = True
    txtUser.Enabled = True
    txtHost.Enabled = True
    txtPass.Enabled = True
    txtPort.Enabled = True
    prgUpload.Value = 0
    DoSend = False
    FileCount = 0
    pasv.Close
    ftpUpload.Close
End Sub

' All subs below this line handle highlighting and return to tab processing
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
