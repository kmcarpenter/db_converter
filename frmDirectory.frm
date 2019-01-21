VERSION 5.00
Begin VB.Form frmDirectory 
   Caption         =   "123"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4455
   ControlBox      =   0   'False
   HelpContextID   =   250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "123"
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   3
      ToolTipText     =   "129"
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "125"
      Height          =   315
      Left            =   3240
      TabIndex        =   2
      Tag             =   "125"
      ToolTipText     =   "128"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "124"
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Tag             =   "124"
      ToolTipText     =   "127"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.DirListBox Dir1 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "126"
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmDirectory"
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

Private Sub cmdSelect_Click()
    ' Return the directory selected
    If Right(Dir1.Path, 1) = "\" Then
        ReturnDir = Dir1.Path
    Else
        ReturnDir = Dir1.Path & "\"
    End If
    Unload Me
End Sub

Private Sub cmdSelect_KeyPress(KeyAscii As Integer)
    ' Send escape to the directory box
    dir1_KeyPress (KeyAscii)
End Sub

Private Sub cmdCancel_Click()
    ' get rid of me
    Unload Me
End Sub

Private Sub cmdCancel_KeyPress(KeyAscii As Integer)
    ' send escape to the directory box
    dir1_KeyPress (KeyAscii)
End Sub

Private Sub Drive1_Change()
    ' Change drive in dir box
    Dir1.Path = Drive1.Drive
End Sub

Private Sub dir1_KeyPress(KeyAscii As Integer)
    ' if we get an escape, close the form
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Drive1_KeyPress(KeyAscii As Integer)
    ' send escape to the dir box
    dir1_KeyPress (KeyAscii)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ' send escape to teh dir box
    dir1_KeyPress (KeyAscii)
End Sub

Private Sub Form_Load()
    On Error GoTo err_handle
    
    lang.LoadFormStrings Me
    
    ' Load the defaults from the options object
    Dir1.Path = Opt.IO_BrowseDir
    Drive1.Drive = Mid(Opt.IO_BrowseDir, 1, 3)
    Exit Sub
err_handle:
    If err.Number = 76 Then
        Dim old As String
        old = Opt.IO_BrowseDir
        Opt.IO_BrowseDir = App.Path
        Call modGlobals.SaveOpt
        MsgBox lang.GetString(295) & vbCrLf & old & vbCrLf & lang.GetString(280) & vbCrLf & App.Path & vbCrLf & lang.GetString(296), vbInformation + vbOKOnly, App.Title
        Dir1.Path = Opt.IO_BrowseDir
    End If
End Sub
