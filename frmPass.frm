VERSION 5.00
Begin VB.Form frmPass 
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4110
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1245
   ScaleWidth      =   4110
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtUser 
      Height          =   288
      IMEMode         =   3  'DISABLE
      Left            =   1560
      TabIndex        =   1
      Text            =   "Admin"
      Top             =   120
      Width           =   2412
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "125"
      Height          =   372
      Left            =   2880
      TabIndex        =   4
      Tag             =   "125"
      Top             =   840
      Width           =   1092
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "121"
      Height          =   372
      Left            =   1560
      TabIndex        =   3
      Tag             =   "121"
      Top             =   840
      Width           =   1092
   End
   Begin VB.TextBox txtPass 
      Height          =   288
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   480
      Width           =   2412
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "220"
      Height          =   252
      Left            =   0
      TabIndex        =   5
      Tag             =   "220"
      Top             =   120
      Width           =   1332
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "221"
      Height          =   252
      Left            =   0
      TabIndex        =   0
      Tag             =   "221"
      Top             =   480
      Width           =   1332
   End
End
Attribute VB_Name = "frmPass"
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

' This form is shown if we happen to get a password error on the database

' Set the global password and user
Private Sub cmdOK_Click()
    If txtPass.Text <> "" Or txtUser.Text <> "" Then
        DB_Pass = txtPass.Text
        DB_User = txtUser.Text
        Unload Me
    End If
End Sub

' Set the password to nothing
Private Sub cmdCancel_Click()
    DB_Pass = ""
    Unload Me
End Sub

' Set focus to the user input
Private Sub Form_Activate()
    txtUser.SetFocus
End Sub

' Set form title
Private Sub Form_Load()

    lang.LoadFormStrings Me
        
    Me.Caption = App.Title & ": " & lang.GetString(332)
End Sub

' Handle highlight/return to tab conversion
Private Sub txtUser_gotfocus()
    Highlight txtUser
End Sub

Private Sub txtuser_KeyPress(KeyAscii As Integer)
    MyTab KeyAscii
End Sub

