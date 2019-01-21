VERSION 5.00
Begin VB.Form frmDonation 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3195
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "132"
      Height          =   420
      Left            =   3840
      TabIndex        =   2
      Tag             =   "132"
      Top             =   2640
      Width           =   735
   End
   Begin VB.CheckBox chkShow 
      Caption         =   "131"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Tag             =   "131"
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "130"
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Tag             =   "130"
      ToolTipText     =   "130"
      Top             =   120
      Width           =   4455
   End
   Begin VB.Image imgDonate 
      Height          =   660
      Left            =   1800
      MouseIcon       =   "frmDonation.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmDonation.frx":030A
      Top             =   1680
      Width           =   1095
   End
End
Attribute VB_Name = "frmDonation"
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

Private Sub chkShow_Click()
    ' Set show on startup bit if checkbox used
    If chkShow.Value = vbChecked Then
        Call SET_BIT(Opt.GenOptions, OPT_DONATE)
    Else
        Call REMOVE_BIT(Opt.GenOptions, OPT_DONATE)
    End If
    SaveOpt
End Sub

Private Sub cmdClose_Click()
    ' close me
    Unload Me
End Sub

Private Sub Form_Load()
    
    lang.LoadFormStrings Me
    
    ' Grab the show on startup option for the checkbox
    chkShow.Value = IS_SET(Opt.GenOptions, OPT_DONATE)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Make sure to clean up and kill the html file
    On Error Resume Next
    Kill App.Path & "\donatetemp.html"
End Sub

Private Sub imgDonate_Click()
    ' Generate an on the fly HTML page that points to my paypal account
    On Error Resume Next
    Dim HTMLString As String
    HTMLString = "<HTML><HEAD><meta HTTP-EQUIV='Content-Type' content='text/html; charset=iso-8859-1'><META HTTP-EQUIV='REFRESH' CONTENT='0; URL=https://www.paypal.com/cgi-bin/webscr?cmd=_xclick&business=zenwerx@zenwerx.com&item_name=Donation for DB Converter&no_note=1&currency_code=USD&tax=0'></HEAD></HTML>"
    Kill App.Path & "\donatetemp.html"
    Open App.Path & "\donatetemp.html" For Binary As #1
    Put #1, , HTMLString
    Call ShellExecute(0&, vbNullString, "file://" & App.Path & "\donatetemp.html", vbNullString, _
                        vbNullString, vbNormalFocus)
    Close #1
End Sub
