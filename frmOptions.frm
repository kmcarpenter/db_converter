VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   Caption         =   "210"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5130
   HelpContextID   =   260
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "210"
   Begin VB.CommandButton cmdDefaults 
      Caption         =   "180"
      Height          =   405
      Left            =   2606
      TabIndex        =   29
      Tag             =   "180"
      ToolTipText     =   "181"
      Top             =   3240
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3120
      Left            =   60
      TabIndex        =   3
      Tag             =   "158;185;186;187"
      Top             =   15
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5503
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "158"
      TabPicture(0)   =   "frmOptions.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "chkRConv"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "chkAssume"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkReturn"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkHighlight"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkTips"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkLinked"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkUpdate"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chkDonate"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "chkAllTables"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "chkINNODB"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "185"
      TabPicture(1)   =   "frmOptions.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdBrowse2"
      Tab(1).Control(1)=   "cmdBrowse"
      Tab(1).Control(2)=   "Frame1"
      Tab(1).Control(3)=   "txtOutput"
      Tab(1).Control(4)=   "txtBrowse"
      Tab(1).Control(5)=   "Label2"
      Tab(1).Control(6)=   "Label1"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "186"
      TabPicture(2)   =   "frmOptions.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraSettings"
      Tab(2).Control(1)=   "chkPHP_Settings"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "187"
      TabPicture(3)   =   "frmOptions.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame3"
      Tab(3).Control(1)=   "Frame2"
      Tab(3).ControlCount=   2
      Begin VB.CheckBox chkINNODB 
         Caption         =   "177"
         Height          =   255
         Left            =   240
         TabIndex        =   57
         Tag             =   "177"
         ToolTipText     =   "178"
         Top             =   2700
         Width           =   4692
      End
      Begin VB.CheckBox chkAllTables 
         Caption         =   "173"
         Height          =   255
         Left            =   240
         TabIndex        =   56
         Tag             =   "173"
         ToolTipText     =   "174"
         Top             =   2220
         Width           =   4692
      End
      Begin VB.CheckBox chkDonate 
         Caption         =   "169"
         Height          =   255
         Left            =   240
         TabIndex        =   55
         Tag             =   "169"
         ToolTipText     =   "170"
         Top             =   1740
         Width           =   4692
      End
      Begin VB.CheckBox chkUpdate 
         Caption         =   "175"
         Height          =   255
         Left            =   240
         TabIndex        =   54
         Tag             =   "175"
         ToolTipText     =   "176"
         Top             =   2460
         Width           =   4692
      End
      Begin VB.CheckBox chkLinked 
         Caption         =   "171"
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Tag             =   "171"
         ToolTipText     =   "172"
         Top             =   1980
         Width           =   4692
      End
      Begin VB.CheckBox chkTips 
         Caption         =   "167"
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Tag             =   "167"
         ToolTipText     =   "168"
         Top             =   1500
         Width           =   4692
      End
      Begin VB.CheckBox chkHighlight 
         Caption         =   "165"
         Height          =   255
         Left            =   240
         TabIndex        =   51
         Tag             =   "165"
         ToolTipText     =   "166"
         Top             =   1260
         Width           =   4692
      End
      Begin VB.CheckBox chkReturn 
         Caption         =   "163"
         Height          =   255
         Left            =   240
         TabIndex        =   50
         Tag             =   "163"
         ToolTipText     =   "164"
         Top             =   1020
         Width           =   4692
      End
      Begin VB.Frame Frame3 
         Height          =   1335
         Left            =   -74880
         TabIndex        =   31
         Top             =   1680
         Width           =   4815
         Begin VB.TextBox txtSQLDB 
            Height          =   285
            Left            =   3120
            TabIndex        =   42
            ToolTipText     =   "211"
            Top             =   960
            Width           =   1575
         End
         Begin VB.TextBox txtSQLUser 
            Height          =   285
            Left            =   600
            TabIndex        =   38
            ToolTipText     =   "215"
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txtSQLHost 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   600
            TabIndex        =   40
            ToolTipText     =   "214"
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox txtSQLPort 
            Height          =   285
            Left            =   3120
            TabIndex        =   41
            ToolTipText     =   "212"
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox txtSQLPass 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   3120
            PasswordChar    =   "*"
            TabIndex        =   39
            ToolTipText     =   "213"
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label18 
            Caption         =   "151"
            Height          =   255
            Left            =   2280
            TabIndex        =   47
            Tag             =   "151"
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label17 
            Caption         =   "140"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Tag             =   "140"
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label16 
            Caption         =   "138"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Tag             =   "138"
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label15 
            Caption         =   "139"
            Height          =   255
            Left            =   2640
            TabIndex        =   44
            Tag             =   "139"
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label14 
            Caption         =   "141"
            Height          =   255
            Left            =   2280
            TabIndex        =   43
            Tag             =   "141"
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "208"
         Height          =   1095
         Left            =   -74880
         TabIndex        =   30
         Tag             =   "208"
         Top             =   480
         Width           =   4815
         Begin VB.TextBox txtFTPUser 
            Height          =   285
            Left            =   600
            TabIndex        =   34
            ToolTipText     =   "219"
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txtFTPHost 
            Height          =   285
            Left            =   600
            TabIndex        =   36
            ToolTipText     =   "218"
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox txtFTPPass 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   3120
            PasswordChar    =   "*"
            TabIndex        =   35
            ToolTipText     =   "217"
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txtFTPPort 
            Height          =   285
            Left            =   3120
            TabIndex        =   37
            ToolTipText     =   "216"
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label10 
            Caption         =   "140"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Tag             =   "140"
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label11 
            Caption         =   "138"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Tag             =   "138"
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label13 
            Caption         =   "141"
            Height          =   255
            Left            =   2280
            TabIndex        =   33
            Tag             =   "141"
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label12 
            Caption         =   "139"
            Height          =   255
            Left            =   2640
            TabIndex        =   32
            Tag             =   "139"
            Top             =   600
            Width           =   375
         End
      End
      Begin VB.CommandButton cmdBrowse2 
         Caption         =   "190"
         Height          =   285
         Left            =   -71040
         TabIndex        =   21
         Tag             =   "190"
         ToolTipText     =   "192"
         Top             =   810
         Width           =   975
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "190"
         Height          =   285
         Left            =   -71040
         TabIndex        =   19
         Tag             =   "190"
         ToolTipText     =   "191"
         Top             =   450
         Width           =   975
      End
      Begin VB.Frame Frame1 
         Caption         =   "195"
         Height          =   1695
         Left            =   -74895
         TabIndex        =   22
         Tag             =   "195"
         Top             =   1170
         Width           =   4815
         Begin VB.TextBox txtOracle 
            Height          =   285
            Left            =   1440
            TabIndex        =   28
            ToolTipText     =   "201"
            Top             =   1320
            Width           =   3255
         End
         Begin VB.TextBox txtMySQL 
            Height          =   285
            Left            =   1440
            TabIndex        =   27
            ToolTipText     =   "200"
            Top             =   840
            Width           =   3255
         End
         Begin VB.TextBox txtPHP 
            Height          =   285
            Left            =   1440
            TabIndex        =   26
            ToolTipText     =   "199"
            Top             =   360
            Width           =   3255
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "197"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Tag             =   "197"
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "198"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Tag             =   "198"
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "196"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Tag             =   "196"
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.CheckBox chkAssume 
         Caption         =   "161"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Tag             =   "161"
         ToolTipText     =   "162"
         Top             =   780
         Width           =   4692
      End
      Begin VB.TextBox txtOutput 
         Height          =   285
         Left            =   -73455
         Locked          =   -1  'True
         TabIndex        =   20
         ToolTipText     =   "193"
         Top             =   810
         Width           =   2295
      End
      Begin VB.TextBox txtBrowse 
         Height          =   285
         Left            =   -73455
         Locked          =   -1  'True
         TabIndex        =   18
         ToolTipText     =   "189"
         Top             =   450
         Width           =   2295
      End
      Begin VB.Frame fraSettings 
         Enabled         =   0   'False
         Height          =   1815
         Left            =   -74880
         TabIndex        =   9
         Top             =   840
         Width           =   4800
         Begin VB.TextBox txtPassword 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1080
            PasswordChar    =   "*"
            TabIndex        =   17
            ToolTipText     =   "207"
            Top             =   1320
            Width           =   3615
         End
         Begin VB.TextBox txtUser 
            Height          =   285
            Left            =   1080
            TabIndex        =   16
            ToolTipText     =   "206"
            Top             =   960
            Width           =   3615
         End
         Begin VB.TextBox txtHost 
            Height          =   285
            Left            =   1080
            TabIndex        =   15
            ToolTipText     =   "205"
            Top             =   600
            Width           =   3615
         End
         Begin VB.TextBox txtDB 
            Height          =   285
            Left            =   1080
            TabIndex        =   14
            ToolTipText     =   "204"
            Top             =   240
            Width           =   3615
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "138"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Tag             =   "138"
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "140"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Tag             =   "140"
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "141"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Tag             =   "141"
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "151"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Tag             =   "151"
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.CheckBox chkPHP_Settings 
         Caption         =   "202"
         Height          =   255
         Left            =   -74865
         TabIndex        =   8
         Tag             =   "202"
         ToolTipText     =   "203"
         Top             =   450
         Width           =   2055
      End
      Begin VB.CheckBox chkRConv 
         Caption         =   "159"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Tag             =   "159"
         ToolTipText     =   "160"
         Top             =   540
         Width           =   4692
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "194"
         Height          =   255
         Left            =   -74895
         TabIndex        =   7
         Tag             =   "194"
         Top             =   810
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "188"
         Height          =   255
         Left            =   -74895
         TabIndex        =   6
         Tag             =   "188"
         Top             =   450
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "125"
      Height          =   405
      Left            =   3784
      TabIndex        =   2
      Tag             =   "125"
      ToolTipText     =   "184"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "179"
      Height          =   405
      Left            =   1429
      TabIndex        =   1
      Tag             =   "179"
      ToolTipText     =   "182"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "121"
      Height          =   405
      Left            =   252
      TabIndex        =   0
      Tag             =   "121"
      ToolTipText     =   "183"
      Top             =   3240
      Width           =   1095
   End
End
Attribute VB_Name = "frmOptions"
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

' I can't remember why I didn't use toggle bit on these
' but there was a reason!!!

Private Sub chkAllTables_Click()
    ' Check or remove the bit as needed
    If chkAllTables.Value = vbChecked Then
        Call SET_BIT(Opt.GenOptions, OPT_TABLES)
    Else
        Call REMOVE_BIT(Opt.GenOptions, OPT_TABLES)
    End If
End Sub

Private Sub chkAssume_Click()
    ' Check or remove the bit as needed
    If chkAssume.Value = vbChecked Then
        Call SET_BIT(Opt.GenOptions, OPT_ASSUME)
    Else
        Call REMOVE_BIT(Opt.GenOptions, OPT_ASSUME)
    End If
End Sub

Private Sub chkDonate_Click()
    ' Check or remove the bit as needed
    If chkDonate.Value = vbChecked Then
        Call SET_BIT(Opt.GenOptions, OPT_DONATE)
    Else
        Call REMOVE_BIT(Opt.GenOptions, OPT_DONATE)
    End If
End Sub

Private Sub chkHighlight_Click()
    ' Check or remove the bit as needed
    If chkHighlight.Value = vbChecked Then
        Call SET_BIT(Opt.GenOptions, OPT_HIGHLIGHT)
    Else
        Call REMOVE_BIT(Opt.GenOptions, OPT_HIGHLIGHT)
    End If
End Sub

Private Sub chkINNODB_Click()
    ' Check or remove the bit as needed
    If chkRConv.Value = vbChecked Then
        Call REMOVE_BIT(Opt.GenOptions, OPT_INNODB)
    Else
        Call SET_BIT(Opt.GenOptions, OPT_INNODB)
    End If
End Sub

Private Sub chkLinked_Click()
    ' Check or remove the bit as needed
    If chkLinked.Value = vbChecked Then
        Call SET_BIT(Opt.GenOptions, OPT_WARNLINK)
    Else
        Call REMOVE_BIT(Opt.GenOptions, OPT_WARNLINK)
    End If
End Sub

Private Sub chkPHP_Settings_Click()
    ' Enable the PHP settings frame
    If chkPHP_Settings.Value = vbChecked Then
        fraSettings.Enabled = True
    Else
        txtDB.Text = ""
        txtPassword.Text = ""
        txtUser.Text = ""
        txtHost.Text = ""
        fraSettings.Enabled = False
    End If
End Sub

Private Sub chkRConv_Click()
    ' Check or remove the bit as needed
    If chkRConv.Value = vbChecked Then
        Call SET_BIT(Opt.GenOptions, OPT_REMEMBER)
    Else
        Call REMOVE_BIT(Opt.GenOptions, OPT_REMEMBER)
    End If
End Sub

Private Sub chkReturn_Click()
    ' Check or remove the bit as needed
    If chkReturn.Value = vbChecked Then
        Call SET_BIT(Opt.GenOptions, OPT_ENTER)
    Else
        Call REMOVE_BIT(Opt.GenOptions, OPT_ENTER)
    End If
End Sub

Private Sub chkTips_Click()
    ' Check or remove the bit as needed
    If chkTips.Value = vbChecked Then
        Call SET_BIT(Opt.GenOptions, OPT_SHOWTIP)
    Else
        Call REMOVE_BIT(Opt.GenOptions, OPT_SHOWTIP)
    End If
End Sub

Private Sub chkUpdate_Click()
    ' Check or remove the bit as needed
    If chkUpdate.Value = vbChecked Then
        Call SET_BIT(Opt.GenOptions, OPT_UPDATE)
    Else
        Call REMOVE_BIT(Opt.GenOptions, OPT_UPDATE)
    End If
End Sub

Private Sub cmdApply_Click()
    ' Set the general options
    Opt.IO_BrowseDir = txtBrowse.Text
    Opt.IO_OutputDir = txtOutput.Text
    Opt.IO_PHP_Ext = txtPHP.Text
    Opt.IO_MySQL_Ext = txtMySQL.Text
    Opt.IO_Oracle_Ext = txtOracle.Text
    Opt.PHP_RememberPHP = chkPHP_Settings.Value
    Opt.PHP_DB = txtDB.Text
    Opt.PHP_Host = txtHost.Text
    Opt.PHP_User = txtUser.Text
    Opt.PHP_Pass = txtPassword.Text
    Opt.FTP_Host = txtFTPHost.Text
    Opt.FTP_Pass = txtFTPPass.Text
    Opt.FTP_Port = txtFTPPort.Text
    Opt.FTP_User = txtFTPUser.Text
    Opt.SQL_Host = txtSQLHost.Text
    Opt.SQL_Pass = txtSQLPass.Text
    Opt.SQL_Port = txtSQLPort.Text
    Opt.SQL_User = txtSQLUser.Text
    Opt.SQL_DB = txtSQLDB.Text
    modGlobals.SaveOpt
End Sub

Private Sub cmdBrowse_Click()
    ' Find the default browse directory
    frmDirectory.Show vbModal, Me
    txtBrowse.Text = ReturnDir
End Sub

Private Sub cmdBrowse2_Click()
    ' Find the default output directory
    frmDirectory.Show vbModal, Me
    txtOutput.Text = ReturnDir
End Sub

Private Sub cmdCancel_Click()
    ' Kill the form
    Unload Me
End Sub

Private Sub cmdDefaults_Click()
    
    ' General options are easily reset because of bitvector
    Opt.GenOptions = OPT_ASSUME + OPT_ENTER + OPT_HIGHLIGHT + OPT_SHOWTIP + OPT_DONATE + OPT_UPDATE
    
    ' Reset the rest of the options to defaults
    Opt.Gen_LastTip = 0
    Opt.IO_BrowseDir = App.Path & "\"
    Opt.IO_MySQL_Ext = "_mysql.sql"
    Opt.IO_Oracle_Ext = "_oracle.sql"
    Opt.IO_OutputDir = App.Path & "\"
    Opt.IO_PHP_Ext = "_create.php"
    Opt.PHP_DB = ""
    Opt.PHP_Host = ""
    Opt.PHP_Pass = ""
    Opt.PHP_User = ""
    Opt.PHP_RememberPHP = vbUnchecked
    Opt.FTP_Host = "localhost"
    Opt.FTP_Pass = ""
    Opt.FTP_Port = "21"
    Opt.FTP_User = ""
    Opt.SQL_Host = "localhost"
    Opt.SQL_Pass = ""
    Opt.SQL_Port = "3306"
    Opt.SQL_User = "root"
    Opt.SQL_DB = "mydb"
    Form_Load
End Sub

Private Sub cmdOK_Click()
    ' Apply changes, kill the form
    cmdApply_Click
    Unload Me
End Sub

' Load the values from the options object
Private Sub Form_Load()
    lang.LoadFormStrings Me
    
    chkAssume.Value = IS_SET(Opt.GenOptions, OPT_ASSUME)
    chkRConv.Value = IS_SET(Opt.GenOptions, OPT_REMEMBER)
    chkReturn.Value = IS_SET(Opt.GenOptions, OPT_ENTER)
    chkHighlight.Value = IS_SET(Opt.GenOptions, OPT_HIGHLIGHT)
    chkTips.Value = IS_SET(Opt.GenOptions, OPT_SHOWTIP)
    chkLinked.Value = IS_SET(Opt.GenOptions, OPT_WARNLINK)
    chkDonate.Value = IS_SET(Opt.GenOptions, OPT_DONATE)
    chkUpdate.Value = IS_SET(Opt.GenOptions, OPT_UPDATE)
    chkAllTables.Value = IS_SET(Opt.GenOptions, OPT_TABLES)
    chkINNODB.Value = IS_SET(Opt.GenOptions, OPT_INNODB)
    txtBrowse.Text = Opt.IO_BrowseDir
    txtOutput.Text = Opt.IO_OutputDir
    txtPHP.Text = Opt.IO_PHP_Ext
    txtMySQL.Text = Opt.IO_MySQL_Ext
    txtOracle.Text = Opt.IO_Oracle_Ext
    chkPHP_Settings.Value = Opt.PHP_RememberPHP
    txtDB.Text = Opt.PHP_DB
    txtHost.Text = Opt.PHP_Host
    txtUser.Text = Opt.PHP_User
    txtPassword.Text = Opt.PHP_Pass
    txtFTPHost.Text = Opt.FTP_Host
    txtFTPPass.Text = Opt.FTP_Pass
    txtFTPPort.Text = Opt.FTP_Port
    txtFTPUser.Text = Opt.FTP_User
    txtSQLHost.Text = Opt.SQL_Host
    txtSQLPass.Text = Opt.SQL_Pass
    txtSQLPort.Text = Opt.SQL_Port
    txtSQLUser.Text = Opt.SQL_User
    txtSQLDB.Text = Opt.SQL_DB
End Sub

' Code below this line handles highlighting and return to tab conversion
Private Sub txtBrowse_GotFocus()
    Highlight txtBrowse
End Sub

Private Sub txtBrowse_KeyPress(KeyAscii As Integer)
    MyTab KeyAscii
End Sub

Private Sub txtOutput_GotFocus()
    Highlight txtOutput
End Sub

Private Sub txtOutput_KeyPress(KeyAscii As Integer)
    MyTab KeyAscii
End Sub
Private Sub txtPHP_GotFocus()
    Highlight txtPHP
End Sub

Private Sub txtPHP_KeyPress(KeyAscii As Integer)
    MyTab KeyAscii
End Sub
Private Sub txtMySQL_GotFocus()
    Highlight txtMySQL
End Sub

Private Sub txtMySQL_KeyPress(KeyAscii As Integer)
    MyTab KeyAscii
End Sub
Private Sub txtOracle_GotFocus()
    Highlight txtOracle
End Sub

Private Sub txtOracle_KeyPress(KeyAscii As Integer)
    MyTab KeyAscii
End Sub

Private Sub txtDB_GotFocus()
    Highlight txtDB
End Sub

Private Sub txtDB_KeyPress(KeyAscii As Integer)
    MyTab KeyAscii
End Sub
Private Sub txtHost_GotFocus()
    Highlight txtHost
End Sub

Private Sub txtHost_KeyPress(KeyAscii As Integer)
    MyTab KeyAscii
End Sub
Private Sub txtPassword_GotFocus()
    Highlight txtPassword
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    MyTab KeyAscii
End Sub

Private Sub txtUser_gotfocus()
    Highlight txtUser
End Sub

Private Sub txtuser_KeyPress(KeyAscii As Integer)
    MyTab KeyAscii
End Sub

Private Sub txtftphost_GotFocus()
    Highlight txtFTPHost
End Sub

Private Sub txtftphost_KeyPress(KeyAscii As Integer)
    MyTab KeyAscii
End Sub

Private Sub txtftpUser_GotFocus()
    Highlight txtFTPUser
End Sub

Private Sub txtftpUser_KeyPress(KeyAscii As Integer)
    MyTab KeyAscii
End Sub

Private Sub txtftppass_GotFocus()
    Highlight txtFTPPass
End Sub

Private Sub txtftppass_KeyPress(KeyAscii As Integer)
    MyTab KeyAscii
End Sub

Private Sub txtftpport_GotFocus()
    Highlight txtFTPPort
End Sub

Private Sub txtftpport_KeyPress(KeyAscii As Integer)
    MyTab KeyAscii
End Sub

Private Sub txtsqlUser_GotFocus()
    Highlight txtSQLUser
End Sub

Private Sub txtsqlUser_KeyPress(KeyAscii As Integer)
    MyTab KeyAscii
End Sub

Private Sub txtsqlpass_GotFocus()
    Highlight txtSQLPass
End Sub

Private Sub txtsqlpass_KeyPress(KeyAscii As Integer)
    MyTab KeyAscii
End Sub

Private Sub txtsqlhost_GotFocus()
    Highlight txtSQLHost
End Sub

Private Sub txtsqlhost_KeyPress(KeyAscii As Integer)
    MyTab KeyAscii
End Sub

Private Sub txtsqldb_GotFocus()
    Highlight txtSQLDB
End Sub

Private Sub txtsqldb_KeyPress(KeyAscii As Integer)
    MyTab KeyAscii
End Sub

Private Sub txtsqlport_GotFocus()
    Highlight txtSQLPort
End Sub

Private Sub txtsqlport_KeyPress(KeyAscii As Integer)
    MyTab KeyAscii
End Sub


