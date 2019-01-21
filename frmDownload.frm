VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDownload 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "136"
   ClientHeight    =   960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4785
   Icon            =   "frmDownload.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   960
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "136"
   Begin VB.CommandButton cmdClose 
      Caption         =   "132"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Tag             =   "132"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "137"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Tag             =   "137"
      Top             =   120
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar prgDownload 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   705
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblSatus 
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmDownload"
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



' Simple download form. Used to run the msi that gets downloaded
' All controls are updated from the download functions
Option Explicit

Private Sub cmdClose_Click()
    ' Close form
    Unload Me
End Sub

Private Sub cmdRun_Click()
    ' Run the installer
    ShellExecute Me.hwnd, "open", (App.Path & "\" & "DBConverter.msi"), "", "", 0
    ' End app
    End
End Sub

Private Sub Form_Load()
    lang.LoadFormStrings Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim answer As Integer
    
    answer = MsgBox(lang.GetString(297), vbQuestion + vbYesNo, App.Title)
    
    If answer = vbNo Then
        Cancel = 1
    Else
        Cancel = 0
    End If
End Sub
