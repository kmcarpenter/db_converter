VERSION 5.00
Begin VB.Form frmTables 
   Caption         =   "222"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "222"
   Begin VB.CommandButton cmdNone 
      Caption         =   "223"
      Height          =   375
      Left            =   180
      TabIndex        =   3
      Tag             =   "223"
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "121"
      Height          =   375
      Left            =   2340
      TabIndex        =   2
      Tag             =   "121"
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "224"
      Height          =   375
      Left            =   1260
      TabIndex        =   1
      Tag             =   "224"
      Top             =   2640
      Width           =   975
   End
   Begin VB.ListBox lstTables 
      Height          =   2535
      ItemData        =   "frmTables.frx":0000
      Left            =   0
      List            =   "frmTables.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "frmTables"
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
' Flag for first activation
Dim firstAct As Boolean

Private Sub cmdAll_Click()
    ' Check all tables
    Dim i As Integer
    For i = 0 To (lstTables.ListCount - 1)
        lstTables.Selected(i) = True
    Next i
    DoEvents
    'lstTables.Refresh
End Sub

Private Sub cmdNone_Click()
    ' Uncheck all tables
    Dim i As Integer
    For i = 0 To (lstTables.ListCount - 1)
        lstTables.Selected(i) = False
    Next i
    DoEvents
    'lstTables.Refresh
End Sub

Private Sub cmdOK_Click()
    ' Double check at least one table is chosen
    ' Then hide the form. It's not unloaded because
    ' other processes use this list
    If lstTables.SelCount = 0 Then
        MsgBox lang.GetString(333), vbInformation, App.Title
    Else
        Me.Hide
    End If
End Sub

Private Sub Form_Activate()
    ' If first time, check all tables
    If firstAct = True Then
        cmdAll_Click
        firstAct = False
    End If
End Sub

Private Sub Form_Load()

    lang.LoadFormStrings Me
    
    ' Load icon
    Me.Icon = LoadResPicture(CInt(101), vbResIcon)
    ' Set flag
    firstAct = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' Check to see a table is selected
    ' Hide the form instead of unloading
    If lstTables.SelCount = 0 Then
        MsgBox lang.GetString(333), vbInformation, App.Title
        Cancel = vbTrue
    ' Double check the form is visible. Otherwise we wouldn't be able to ever
    ' close it if the form was hidden and we got an unload event
    ElseIf Me.Visible = True Then
        Me.Hide
        Cancel = True
    End If
End Sub
