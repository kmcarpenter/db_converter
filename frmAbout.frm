VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "122"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5865
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HelpContextID   =   240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5865
   StartUpPosition =   1  'CenterOwner
   Tag             =   "122"
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   480
      Left            =   240
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   465.344
      ScaleMode       =   0  'User
      ScaleWidth      =   465.344
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "133"
      Top             =   240
      Width           =   480
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "121"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   4245
      TabIndex        =   0
      Tag             =   "121"
      ToolTipText     =   "134"
      Top             =   2385
      Width           =   1467
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "120"
      Height          =   252
      Left            =   1056
      TabIndex        =   7
      Tag             =   "120"
      Top             =   1920
      Width           =   4092
   End
   Begin VB.Label lblWebPage 
      Caption         =   " http://www.zenwerx.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2160
      MouseIcon       =   "frmAbout.frx":0442
      MousePointer    =   99  'Custom
      TabIndex        =   6
      ToolTipText     =   "135"
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Zen Werx"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label lblDescription 
      Caption         =   "119"
      ForeColor       =   &H00000000&
      Height          =   684
      Left            =   1056
      TabIndex        =   4
      Tag             =   "119"
      Top             =   1128
      Width           =   4092
   End
   Begin VB.Label lblTitle 
      Caption         =   "117"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1080
      TabIndex        =   3
      Tag             =   "117"
      Top             =   240
      Width           =   4095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   240
      X2              =   5672
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   240
      X2              =   5657
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label lblVersion 
      Caption         =   "118"
      Height          =   225
      Left            =   1050
      TabIndex        =   2
      Tag             =   "118"
      Top             =   780
      Width           =   4092
   End
End
Attribute VB_Name = "frmAbout"
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

' Declare API Call for IE Shell
Private Declare Function ShellExecute Lib "shell32.dll" _
      Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation _
      As String, ByVal lpFile As String, ByVal lpParameters As String, _
      ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long



' Clicked the link. Take me to the webpage
Private Sub lblWebPage_Click()
    Call ShellExecute(0&, vbNullString, "http://www.zenwerx.com", vbNullString, _
                        vbNullString, vbNormalFocus)
End Sub

' Set the version information for the about form
Private Sub Form_Load()
    lang.LoadFormStrings Me
    lblVersion.Caption = lblVersion.Caption & " " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

' Close the form
Private Sub cmdOK_Click()
        Unload Me
End Sub
