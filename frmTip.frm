VERSION 5.00
Begin VB.Form frmTip 
   Caption         =   "228"
   ClientHeight    =   3285
   ClientLeft      =   2370
   ClientTop       =   2400
   ClientWidth     =   5415
   Icon            =   "frmTip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   5415
   StartUpPosition =   1  'CenterOwner
   Tag             =   "228"
   Begin VB.CheckBox chkLoadTipsAtStartup 
      Caption         =   "225"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Tag             =   "225"
      Top             =   2940
      Width           =   2055
   End
   Begin VB.CommandButton cmdNextTip 
      Caption         =   "226"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Tag             =   "226"
      Top             =   600
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2715
      Left            =   120
      Picture         =   "frmTip.frx":0442
      ScaleHeight     =   2655
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   120
      Width           =   3735
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "227"
         Height          =   255
         Left            =   540
         TabIndex        =   5
         Tag             =   "227"
         Top             =   180
         Width           =   2655
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         Height          =   1635
         Left            =   180
         TabIndex        =   4
         Top             =   840
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "121"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Tag             =   "121"
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmTip"
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



' Modified version of MS's default tip form
Option Explicit

' The in-memory database of tips.
Dim Tips As New Collection

' Index in collection of tip currently being displayed.
Dim CurrentTip As Long


Private Sub DoNextTip()
    CurrentTip = CurrentTip + 1
    If Tips.count < CurrentTip Then
        CurrentTip = 1
    End If
    
    ' Show it.
    frmTip.DisplayCurrentTip
    
End Sub

Function LoadTips() As Boolean
    Dim NextTip As String   ' Each tip read in from file.
    Dim InFile As Integer   ' Descriptor for file.
    Dim i As Integer
    ' Read the collection from a text file.
    
    For i = 1 To 16
        NextTip = lang.GetString(100 + i)
        Tips.Add NextTip
    Next i

    ' Display a tip at random.
    DoNextTip
    
    LoadTips = True
    
End Function

Private Sub chkLoadTipsAtStartup_Click()
    ' save whether or not this form should be displayed at startup
    If chkLoadTipsAtStartup.Value = vbChecked Then
        Call SET_BIT(Opt.GenOptions, OPT_SHOWTIP)
    Else
        Call REMOVE_BIT(Opt.GenOptions, OPT_SHOWTIP)
    End If
    SaveOpt
End Sub

Private Sub cmdNextTip_Click()
    DoNextTip
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim ShowAtStartup As Long
        
    lang.LoadFormStrings Me
    
    ' Set the checkbox, this will force the value to be written back out to the registry
    Me.chkLoadTipsAtStartup.Value = vbChecked
    
    ' Seed Rnd
    Randomize
    
    CurrentTip = Opt.Gen_LastTip
    ' Read in the tips file and display a tip at random.
    If LoadTips() = False Then
        Me.Hide
    End If

    
End Sub

Public Sub DisplayCurrentTip()
    If Tips.count > 0 Then
        lblTipText.Caption = Tips.Item(CurrentTip)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Opt.Gen_LastTip = (CurrentTip)
    SaveOpt
    Set Tips = Nothing
End Sub
