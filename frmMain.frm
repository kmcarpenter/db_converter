VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "229"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8070
   HelpContextID   =   10
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   8070
   StartUpPosition =   2  'CenterScreen
   Tag             =   "229"
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin VB.PictureBox picConversions 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   120
         ScaleHeight     =   2295
         ScaleWidth      =   5535
         TabIndex        =   13
         Top             =   120
         Visible         =   0   'False
         Width           =   5535
         Begin VB.CheckBox chkOracle 
            Caption         =   "260"
            Height          =   240
            Left            =   120
            TabIndex        =   57
            Tag             =   "260"
            ToolTipText     =   "267"
            Top             =   1680
            Width           =   1932
         End
         Begin VB.CheckBox chkConnect 
            Caption         =   "264"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3600
            TabIndex        =   56
            Tag             =   "264"
            ToolTipText     =   "272"
            Top             =   1920
            Width           =   1935
         End
         Begin VB.CheckBox chkUpload 
            Caption         =   "263"
            Height          =   255
            Left            =   3600
            TabIndex        =   55
            Tag             =   "263"
            ToolTipText     =   "271"
            Top             =   1680
            Width           =   1935
         End
         Begin VB.CheckBox chkPost 
            Caption         =   "268"
            Height          =   240
            Left            =   120
            TabIndex        =   54
            Tag             =   "268"
            ToolTipText     =   "269"
            Top             =   1920
            Width           =   1932
         End
         Begin VB.CheckBox chkPHP 
            Caption         =   "258"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Tag             =   "258"
            ToolTipText     =   "265"
            Top             =   1200
            Width           =   1692
         End
         Begin VB.CheckBox chkMySQL 
            Caption         =   "259"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Tag             =   "259"
            ToolTipText     =   "266"
            Top             =   1440
            Width           =   1812
         End
         Begin VB.CheckBox chkData 
            Caption         =   "261"
            Height          =   255
            Left            =   3600
            TabIndex        =   16
            Tag             =   "261"
            ToolTipText     =   "270"
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label21 
            Caption         =   "256"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Tag             =   "256"
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label19 
            Caption         =   "262"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3600
            TabIndex        =   47
            Tag             =   "262"
            Top             =   1440
            Width           =   1695
         End
         Begin VB.Image Image5 
            Height          =   1080
            Left            =   2040
            Picture         =   "frmMain.frx":0442
            Stretch         =   -1  'True
            Top             =   1200
            Width           =   1065
         End
         Begin VB.Label Label20 
            Caption         =   "257"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3600
            TabIndex        =   48
            Tag             =   "257"
            Top             =   960
            Width           =   1695
         End
      End
      Begin VB.PictureBox picIntro 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   120
         ScaleHeight     =   2295
         ScaleWidth      =   5535
         TabIndex        =   2
         Top             =   120
         Width           =   5535
         Begin VB.Image Image4 
            Height          =   255
            Left            =   3840
            Picture         =   "frmMain.frx":10A0
            Stretch         =   -1  'True
            Top             =   1200
            Width           =   1560
         End
         Begin VB.Image Image3 
            Height          =   705
            Left            =   480
            Picture         =   "frmMain.frx":1598
            Top             =   1395
            Width           =   705
         End
         Begin VB.Image Image2 
            Height          =   585
            Left            =   3840
            Picture         =   "frmMain.frx":1A01
            Stretch         =   -1  'True
            Top             =   1695
            Width           =   1125
         End
         Begin VB.Image Image1 
            Height          =   1080
            Left            =   2040
            Picture         =   "frmMain.frx":20B8
            Stretch         =   -1  'True
            Top             =   1200
            Width           =   1065
         End
      End
      Begin VB.PictureBox picDB 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   120
         ScaleHeight     =   2295
         ScaleWidth      =   5535
         TabIndex        =   9
         Top             =   120
         Visible         =   0   'False
         Width           =   5535
         Begin VB.TextBox txtDatabase 
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   11
            ToolTipText     =   "283"
            Top             =   1920
            Width           =   3015
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "190"
            Height          =   285
            Left            =   4320
            TabIndex        =   12
            Tag             =   "190"
            ToolTipText     =   "273"
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "151"
            Height          =   255
            Left            =   0
            TabIndex        =   10
            Tag             =   "151"
            Top             =   1920
            Width           =   1095
         End
      End
      Begin VB.PictureBox picOutput 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   120
         ScaleHeight     =   2295
         ScaleWidth      =   5535
         TabIndex        =   17
         Top             =   120
         Visible         =   0   'False
         Width           =   5535
         Begin VB.TextBox txtDirectory 
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   19
            ToolTipText     =   "276"
            Top             =   1560
            Width           =   2535
         End
         Begin VB.CommandButton cmdBrowse2 
            Caption         =   "190"
            Enabled         =   0   'False
            Height          =   285
            Left            =   4320
            TabIndex        =   20
            Tag             =   "190"
            ToolTipText     =   "275"
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txtFileName 
            Height          =   285
            Left            =   1680
            TabIndex        =   21
            ToolTipText     =   "277"
            Top             =   1920
            Width           =   3855
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "194"
            Height          =   255
            Left            =   0
            TabIndex        =   22
            Tag             =   "194"
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "274"
            Height          =   255
            Left            =   0
            TabIndex        =   18
            Tag             =   "274"
            Top             =   1920
            Width           =   1575
         End
      End
      Begin VB.PictureBox picFinished 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   120
         ScaleHeight     =   2295
         ScaleWidth      =   5535
         TabIndex        =   42
         Top             =   120
         Visible         =   0   'False
         Width           =   5535
         Begin MSWinsockLib.Winsock sock 
            Left            =   9999
            Top             =   840
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   393216
         End
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   255
            Left            =   0
            TabIndex        =   51
            Top             =   2040
            Width           =   5000
            _ExtentX        =   8811
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin VB.Label lblNewName1 
            Height          =   255
            Left            =   2760
            TabIndex        =   64
            Top             =   720
            Width           =   2775
         End
         Begin VB.Label lblNewName2 
            Height          =   252
            Left            =   2760
            TabIndex        =   63
            Top             =   960
            Width           =   2775
         End
         Begin VB.Label lblNewName3 
            Height          =   252
            Left            =   2760
            TabIndex        =   62
            Top             =   1200
            Width           =   2775
         End
         Begin VB.Label lblNewName4 
            Height          =   252
            Left            =   2760
            TabIndex        =   61
            Top             =   1440
            Width           =   2775
         End
         Begin VB.Label lblFileName 
            Height          =   255
            Left            =   960
            TabIndex        =   58
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            Caption         =   "100%"
            Height          =   255
            Left            =   5040
            TabIndex        =   53
            Top             =   2040
            Width           =   495
         End
         Begin VB.Label Label18 
            Caption         =   "280"
            Height          =   375
            Left            =   2520
            TabIndex        =   45
            Tag             =   "280"
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label15 
            Caption         =   "279"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Tag             =   "279"
            Top             =   720
            Width           =   855
         End
         Begin VB.Image Image7 
            Height          =   1080
            Left            =   120
            Picture         =   "frmMain.frx":2D16
            Stretch         =   -1  'True
            Top             =   925
            Width           =   1065
         End
         Begin VB.Label Label14 
            Caption         =   "278"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   43
            Tag             =   "278"
            Top             =   120
            Width           =   5295
         End
      End
      Begin VB.PictureBox picConvert 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   120
         ScaleHeight     =   2295
         ScaleWidth      =   5535
         TabIndex        =   36
         Top             =   120
         Visible         =   0   'False
         Width           =   5535
         Begin MSComctlLib.ProgressBar prgStates 
            Height          =   255
            Left            =   0
            TabIndex        =   50
            Top             =   2040
            Width           =   5000
            _ExtentX        =   8811
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin VB.Label lblConvert5 
            BackStyle       =   0  'Transparent
            Height          =   252
            Left            =   1680
            TabIndex        =   59
            Top             =   1680
            Width           =   1452
         End
         Begin VB.Label lblPct 
            Alignment       =   1  'Right Justify
            Caption         =   "0%"
            Height          =   255
            Left            =   5040
            TabIndex        =   52
            Top             =   2040
            Width           =   495
         End
         Begin VB.Label lblConvert4 
            BackStyle       =   0  'Transparent
            Height          =   252
            Left            =   1680
            TabIndex        =   46
            Top             =   1440
            Width           =   1452
         End
         Begin VB.Image Image6 
            Height          =   1080
            Left            =   120
            Picture         =   "frmMain.frx":3974
            Stretch         =   -1  'True
            Top             =   925
            Width           =   1065
         End
         Begin VB.Label Label16 
            Caption         =   "282"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Tag             =   "282"
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label lblConvert1 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   1680
            TabIndex        =   39
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label lblConvert2 
            BackStyle       =   0  'Transparent
            Height          =   252
            Left            =   1680
            TabIndex        =   38
            Top             =   960
            Width           =   1452
         End
         Begin VB.Label lblConvert3 
            BackStyle       =   0  'Transparent
            Height          =   252
            Left            =   1680
            TabIndex        =   37
            Top             =   1200
            Width           =   1452
         End
         Begin VB.Label Label17 
            Caption         =   "281"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   41
            Tag             =   "281"
            Top             =   120
            Width           =   5295
         End
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "237"
      Height          =   405
      Left            =   71
      TabIndex        =   60
      Tag             =   "237"
      ToolTipText     =   "242"
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "241"
      Enabled         =   0   'False
      Height          =   405
      Left            =   6480
      TabIndex        =   35
      Tag             =   "241"
      ToolTipText     =   "246"
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "238"
      Enabled         =   0   'False
      Height          =   405
      Left            =   1673
      TabIndex        =   32
      Tag             =   "238"
      ToolTipText     =   "243"
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "239"
      Height          =   405
      Left            =   3275
      TabIndex        =   33
      Tag             =   "239"
      ToolTipText     =   "244"
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdMake 
      Caption         =   "240"
      Enabled         =   0   'False
      Height          =   405
      Left            =   4877
      TabIndex        =   34
      Tag             =   "240"
      ToolTipText     =   "245"
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Frame fraPHP 
      Caption         =   "247"
      Enabled         =   0   'False
      Height          =   975
      Left            =   0
      TabIndex        =   23
      Tag             =   "247"
      Top             =   2520
      Visible         =   0   'False
      Width           =   8055
      Begin VB.TextBox txtHost 
         Height          =   285
         Left            =   1560
         TabIndex        =   25
         ToolTipText     =   "253"
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   5880
         TabIndex        =   26
         ToolTipText     =   "254"
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   5880
         PasswordChar    =   "*"
         TabIndex        =   27
         ToolTipText     =   "255"
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtDB 
         Height          =   285
         Left            =   1560
         TabIndex        =   24
         ToolTipText     =   "252"
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "249"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Tag             =   "249"
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "250"
         Height          =   255
         Left            =   4560
         TabIndex        =   30
         Tag             =   "250"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "251"
         Height          =   255
         Left            =   4320
         TabIndex        =   29
         Tag             =   "251"
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "248"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Tag             =   "248"
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "230"
      Height          =   2535
      Left            =   5760
      TabIndex        =   1
      Tag             =   "230"
      Top             =   0
      Width           =   2295
      Begin VB.Image chkDB 
         Height          =   195
         Left            =   120
         Picture         =   "frmMain.frx":45D2
         Stretch         =   -1  'True
         Top             =   600
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Image chkConversions 
         Height          =   195
         Left            =   120
         Picture         =   "frmMain.frx":49C4
         Stretch         =   -1  'True
         Top             =   1008
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Image chkOutput 
         Height          =   195
         Left            =   120
         Picture         =   "frmMain.frx":4DB6
         Stretch         =   -1  'True
         Top             =   1392
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Image chkConvert 
         Height          =   195
         Left            =   120
         Picture         =   "frmMain.frx":51A8
         Stretch         =   -1  'True
         Top             =   1776
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Image chkFinished 
         Height          =   195
         Left            =   120
         Picture         =   "frmMain.frx":559A
         Stretch         =   -1  'True
         Top             =   2160
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Image chkIntro 
         Height          =   195
         Left            =   120
         Picture         =   "frmMain.frx":598C
         Stretch         =   -1  'True
         Top             =   240
         Width           =   195
      End
      Begin VB.Image xDB 
         Height          =   195
         Left            =   120
         Picture         =   "frmMain.frx":5D7E
         Stretch         =   -1  'True
         Top             =   624
         Width           =   195
      End
      Begin VB.Image xConversions 
         Height          =   195
         Left            =   120
         Picture         =   "frmMain.frx":6194
         Stretch         =   -1  'True
         Top             =   1008
         Width           =   195
      End
      Begin VB.Image xOutput 
         Height          =   195
         Left            =   120
         Picture         =   "frmMain.frx":65AA
         Stretch         =   -1  'True
         Top             =   1392
         Width           =   195
      End
      Begin VB.Image xConvert 
         Height          =   195
         Left            =   120
         Picture         =   "frmMain.frx":69C0
         Stretch         =   -1  'True
         Top             =   1776
         Width           =   195
      End
      Begin VB.Image xFinished 
         Height          =   195
         Left            =   120
         Picture         =   "frmMain.frx":6DD6
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   195
      End
      Begin VB.Label Label6 
         Caption         =   "236"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Tag             =   "236"
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "235"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Tag             =   "235"
         Top             =   1776
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "234"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Tag             =   "234"
         Top             =   1392
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "233"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Tag             =   "233"
         Top             =   1008
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "232"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Tag             =   "232"
         Top             =   624
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "231"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Tag             =   "231"
         Top             =   240
         Width           =   1815
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   3960
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "290"
      HelpContextID   =   210
      Tag             =   "290"
      Begin VB.Menu mnuFile_Exit 
         Caption         =   "291"
         Shortcut        =   ^Q
         Tag             =   "291"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "285"
      HelpContextID   =   220
      Tag             =   "285"
      Begin VB.Menu mnuEdit_Cut 
         Caption         =   "287"
         Shortcut        =   ^X
         Tag             =   "287"
      End
      Begin VB.Menu mnuEdit_Copy 
         Caption         =   "286"
         Shortcut        =   ^C
         Tag             =   "286"
      End
      Begin VB.Menu mnuEdit_Paste 
         Caption         =   "289"
         Shortcut        =   ^V
         Tag             =   "289"
      End
      Begin VB.Menu Seperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Options 
         Caption         =   "288"
         Shortcut        =   ^O
         Tag             =   "288"
      End
   End
   Begin VB.Menu mnuLanguages 
      Caption         =   "393"
      Tag             =   "393"
      Begin VB.Menu mnuLangSub 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "237"
      HelpContextID   =   230
      Tag             =   "237"
      Begin VB.Menu mnuHelp_Contents 
         Caption         =   "293"
         Shortcut        =   {F1}
         Tag             =   "293"
      End
      Begin VB.Menu mnuSep12345 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCheck 
         Caption         =   "284"
         Tag             =   "284"
      End
      Begin VB.Menu mnuCheckLang 
         Caption         =   "394"
         Tag             =   "394"
      End
      Begin VB.Menu mnuHelpDonate 
         Caption         =   "294"
         Checked         =   -1  'True
         Tag             =   "294"
      End
      Begin VB.Menu mnuSep23456 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp_About 
         Caption         =   "292"
         Tag             =   "292"
      End
   End
End
Attribute VB_Name = "frmMain"
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


' Formwide timestamp variable
' Not sure why it's here, I can't find reference to it anywhere
' *******************
' Dim mynow As Long
' *******************

' Show tips variable
Dim showtips As Boolean

' Extra check for showing modal forms
' Form_Activate was firing after modal forms (tips/donate/download) were
' destroyed and causing errors.
' (Introduced due to DoEvents addition???)
Dim doingtips As Boolean

' Frame index
Dim CurIdx As Integer

Option Explicit

' Handle mysql checkbox click
Private Sub chkMySQL_Click()
    If chkMySQL.Value = vbChecked Then
        chkConnect.Enabled = True
    Else
        chkConnect.Value = vbUnchecked
        chkConnect.Enabled = False
    End If
End Sub

' Handle php checkbox click
Private Sub chkPHP_Click()
    ' Set the default values of the PHP text boxes
    ' and handle the showing/hiding of the extra information
    ' boxes
    
    ' The command button moves should be handled more like the
    ' form height. And on the topic of form height, it should be
    ' calculated on control height rather than hard coded.
    If chkPHP.Value = vbChecked Then
        If Opt.PHP_RememberPHP = vbChecked Then
            txtDB.Text = Opt.PHP_DB
            txtHost.Text = Opt.PHP_Host
            txtUser.Text = Opt.PHP_User
            txtPassword.Text = Opt.PHP_Pass
        Else
            txtDB.Text = ""
            txtHost.Text = ""
            txtUser.Text = ""
            txtPassword.Text = ""
        End If
        If picConversions.Visible = True Then
            Me.Height = Me.Height + 950
            cmdHelp.top = 3600
            cmdBack.top = 3600
            cmdNext.top = 3600
            cmdMake.top = 3600
            cmdStart.top = 3600
            fraPHP.Visible = True
            fraPHP.Enabled = True
        End If
    Else
        fraPHP.Enabled = False
        If picConversions.Visible = True Then
            Me.Height = Me.Height - 950
            fraPHP.Visible = False
            cmdHelp.top = 2640
            cmdBack.top = 2640
            cmdNext.top = 2640
            cmdMake.top = 2640
            cmdStart.top = 2640
        End If
    End If
End Sub

Private Sub cmdAbout_Click()
    ' Show the about form
    frmAbout.Show vbModal, Me
End Sub

' Return to the previous step
Private Sub cmdBack_Click()
    CurIdx = CurIdx - 1
    If picDB.Visible = True Then
        PrintStrings CurIdx
        picIntro.Visible = True
        picDB.Visible = False
        picDB.Cls
        chkDB.Visible = False
        xDB.Visible = True
    ElseIf picConversions.Visible = True Then
        PrintStrings CurIdx
        picDB.Visible = True
        picConversions.Visible = False
        picConversions.Cls
        chkConversions.Visible = False
        xConversions.Visible = True
        fraPHP.Enabled = False
        Me.Height = 3735
        fraPHP.Visible = False
        cmdHelp.top = 2640
        cmdBack.top = 2640
        cmdNext.top = 2640
        cmdMake.top = 2640
        cmdStart.top = 2640
    ElseIf picOutput.Visible = True Then
        PrintStrings CurIdx
        picConversions.Visible = True
        picOutput.Visible = False
        picOutput.Cls
        chkOutput.Visible = False
        xOutput.Visible = True
        cmdNext.Enabled = True
        cmdMake.Enabled = False
        If chkPHP.Value = vbChecked Then
            Me.Height = Me.Height + 950
            cmdHelp.top = 3600
            cmdBack.top = 3600
            cmdNext.top = 3600
            cmdMake.top = 3600
            cmdStart.top = 3600
            fraPHP.Visible = True
            fraPHP.Enabled = True
        End If
    End If
End Sub

' Invoke the choose file common dialog
Private Sub cmdBrowse_Click()
    Dim i As Integer
    Dim newfile As String

    On Error GoTo err_handle
    
    CD.CancelError = True
    CD.Filter = "Microsoft Access Databases (*.mdb)|*.mdb|All Files (*.*)|*.*"
    CD.InitDir = Opt.IO_BrowseDir
    CD.FileName = "*.mdb"
    CD.ShowOpen
    txtDatabase = CD.FileName
    For i = 0 To Len(CD.FileName)
        If Mid(CD.FileName, (Len(CD.FileName) - i), 1) = "\" Then
            newfile = Mid(CD.FileName, (Len(CD.FileName) - i + 1), i - 4)
            If IS_SET(Opt.GenOptions, OPT_ASSUME) Then
                If Dir(Opt.IO_OutputDir) = "" Then
                    Opt.IO_OutputDir = App.Path
                    modGlobals.SaveOpt
                End If
                txtDirectory.Text = Opt.IO_OutputDir
                txtFileName.Text = newfile
            End If
            Exit For
        End If
    Next i

err_handle:
End Sub

Private Sub cmdBrowse2_Click()
    ' Choose the directory
    frmDirectory.Show vbModal, Me
    txtDirectory.Text = ReturnDir
End Sub

Private Sub cmdHelp_Click()
    ' Invoke help
    Call ShellExecute(0&, vbNullString, "file://" & Opt.Gen_Path & "\help\index.htm", vbNullString, _
                        vbNullString, vbNormalFocus)
End Sub

' Make our SQL
Private Sub cmdMake_Click()
    On Error GoTo error_handle
    ' Check for errors
    If FileLen(txtDatabase.Text) > 0 Then
        If txtDirectory.Text = "" Then
            MsgBox lang.GetString(353) & "!", vbInformation, App.Title
            cmdBrowse2.SetFocus
            Exit Sub
        ElseIf txtFileName.Text = "" Then
            MsgBox lang.GetString(354) & "!", vbInformation, App.Title
            txtFileName.SetFocus
            Exit Sub
        End If
        DB_Pass = ""
        DB_User = ""
        cmdMake.Enabled = False
        xConvert.Visible = False
        chkConvert.Visible = True
        picConvert.Visible = True
        picOutput.Visible = False
        DoEvents
        cmdBack.Enabled = False
        cmdNext.Enabled = False
        working = True
        goterr = False
        
        ' Choose tables to convert if you have that option chosen
        If IS_SET(Opt.GenOptions, OPT_TABLES) <> 1 Then
            ChooseTables txtDatabase.Text
        End If
        
        ' Count the SQL statements we need
        CountStatements txtDatabase.Text
        
        ' Check for error
        If goterr = True Then
            Exit Sub
        End If
        
        ' This is where the actual conversion takes place.
        ' If the checkboxes are on, declare the appropriate
        ' class variable, convert the Database, and then
        ' Destroy the class
        If chkPHP.Value = vbChecked Then
            lblConvert1.Caption = "PHP"
            DoEvents
            Dim PHP As New CPHP_Convert
            PHP.MakePHP txtDatabase.Text, txtDirectory.Text, txtFileName.Text, chkData.Value
            Set PHP = Nothing
        End If
        
        ' Not sure why this variable is here
        ' Have a feeling it was used for timeouts (maybe) in an older version
        ' *****************
        ' Dim mynow As Long
        ' *****************
        
        If chkMySQL.Value = vbChecked Then
            If lblConvert1.Caption = "" Then
                lblConvert1.Caption = "SQL - MySQL"
            Else
                lblConvert2.Caption = "SQL - MySQL"
            End If
            DoEvents
            Dim MySQL As New CMySQL_Convert
            MySQL.MakeMySQL txtDatabase.Text, txtDirectory.Text, txtFileName.Text, chkData.Value
            Set MySQL = Nothing
        End If
        
        If chkOracle.Value = vbChecked Then
            If lblConvert1.Caption = "" Then
                lblConvert1.Caption = "SQL - Oracle"
            ElseIf lblConvert2.Caption = "" Then
                lblConvert2.Caption = "SQL - Oracle"
            Else
                lblConvert3.Caption = "SQL - Oracle"
            End If
            DoEvents
            Dim Oracle As New COracle_Convert
            Oracle.MakeOracle txtDatabase.Text, txtDirectory.Text, txtFileName.Text, chkData.Value
        End If
        If chkPost.Value = vbChecked Then
            If lblConvert1.Caption = "" Then
                lblConvert1.Caption = "SQL - PostgreSQL"
            ElseIf lblConvert2.Caption = "" Then
                lblConvert2.Caption = "SQL - PostgreSQL"
            ElseIf lblConvert3.Caption = "" Then
                lblConvert3.Caption = "SQL - PostgreSQL"
            Else
                lblConvert4.Caption = "SQL - PostgreSQL"
            End If
            DoEvents
            Dim PostgreSQL As New CPostgreSQL_Convert
            PostgreSQL.MakePostgreSQL txtDatabase.Text, txtDirectory.Text, txtFileName.Text, chkData.Value
            Set PostgreSQL = Nothing
        End If
        ' We're no longer working/Converting
        working = False
        
        ' Find a blank label and tell the user we're done
        If lblConvert2.Caption = "" Then
            lblConvert2.Caption = lang.GetString(355)
        ElseIf lblConvert3.Caption = "" Then
            lblConvert3.Caption = lang.GetString(355)
        ElseIf lblConvert4.Caption = "" Then
            lblConvert4.Caption = lang.GetString(355)
        Else
            lblConvert5.Caption = lang.GetString(355)
        End If
        
        ' If you chose to use the ftp option, invoke the form
        If chkUpload.Value = vbChecked Then
            frmFTP.Show vbModal, Me
        End If
        ' If you chose to use the MySQL connection, invoke the form
        If chkConnect.Value = vbChecked Then
            frmMySQL.Show vbModal, Me
        End If
        
        ' Reset the form so we can start again
        cmdStart.Enabled = True
        cmdStart.SetFocus
    End If
    
    ' Destory the tables form (it was kept alive for use by the other forms)
    Unload frmTables
    Exit Sub

' Handle errors
error_handle:
    ' The file you want to convert doesn't exist
    ' Possible deleted after you chose it
    If err.Number = 76 Or err.Number = 53 Then
        MsgBox lang.GetString(356), vbInformation, App.Title
        txtDatabase.Enabled = True
    ' Can't open your file
    ElseIf err.Number = 3051 Then
        MsgBox lang.GetString(357), vbInformation, App.Title
        cmdStart.Caption = lang.GetString(339)
        cmdStart.Enabled = True
        prgStates.Value = 100
        lblPct.Caption = lang.GetString(340)
        working = False
    ' Wrong password
    ElseIf err.Number = 3031 Then
        MsgBox lang.GetString(358), vbInformation, App.Title
        cmdStart.Caption = lang.GetString(339)
        cmdStart.Enabled = True
        prgStates.Value = 100
        lblPct.Caption = lang.GetString(340)
        working = False
    ' Other error. You should be mighty concerned about this one
    Else
        MsgBox err.Number & " : " & err.Description, vbCritical, App.Title
        cmdStart.Caption = "lang.GetString(339))"
        cmdStart.Enabled = True
        prgStates.Value = 100
        lblPct.Caption = lang.GetString(340)
        working = False
    End If
    
    ' If we haven't already (which is probably the case if you made it this far)
    ' unload the tables form to release the memory.
    Unload frmTables
    Exit Sub
End Sub

' Move to the next step and check data integrity before moving
' this doesn't need to be done when we move backwards because the
' data should already have passed these rules to get to the next step
Public Sub cmdNext_Click()
    CurIdx = CurIdx + 1
    If picIntro.Visible = True Then
        picDB.Cls
        xDB.Visible = False
        chkDB.Visible = True
        PrintStrings CurIdx
        picIntro.Visible = False
        cmdBack.Enabled = True
    ElseIf picDB.Visible = True Then
        If txtDatabase.Text = "" Then
            MsgBox lang.GetString(364) & "!", vbInformation, App.Title
            cmdBrowse.SetFocus
            Exit Sub
        End If
        PrintStrings CurIdx
        xConversions.Visible = False
        chkConversions.Visible = True
        picConversions.Visible = True
        picDB.Visible = False
        If chkPHP.Value = vbChecked Then
            fraPHP.Enabled = True
            Me.Height = Me.Height + 950
            cmdHelp.top = 3600
            cmdBack.top = 3600
            cmdNext.top = 3600
            cmdMake.top = 3600
            cmdStart.top = 3600
            fraPHP.Visible = True
            fraPHP.Enabled = True
        End If
    ElseIf picConversions.Visible = True Then
        Dim checked As Boolean
        checked = False
        If chkPHP.Value = vbChecked Then
            If txtDB.Text = "" Then
                MsgBox lang.GetString(374), vbInformation, App.Title
                txtDB.SetFocus
                Exit Sub
            ElseIf txtHost.Text = "" Then
                MsgBox lang.GetString(375), vbInformation, App.Title
                txtHost.SetFocus
                Exit Sub
            ElseIf txtUser.Text = "" Then
                MsgBox lang.GetString(376), vbInformation, App.Title
                txtUser.SetFocus
                Exit Sub
            ElseIf txtPassword.Text = "" Then
                MsgBox lang.GetString(377), vbInformation, App.Title
                txtPassword.SetFocus
                Exit Sub
            End If
            checked = True
        End If
        If chkMySQL.Value = vbChecked Then
            checked = True
        End If
        If chkOracle.Value = vbChecked Then
            checked = True
        End If
        If chkPost.Value = vbChecked Then
            checked = True
        End If
        If checked = False Then
            MsgBox lang.GetString(378), vbInformation, App.Title
            chkPHP.SetFocus
            Exit Sub
        End If
        fraPHP.Enabled = False
        Me.Height = 3735
        fraPHP.Visible = False
        cmdHelp.top = 2640
        cmdBack.top = 2640
        cmdNext.top = 2640
        cmdMake.top = 2640
        cmdStart.top = 2640
        xOutput.Visible = False
        chkOutput.Visible = True
        PrintStrings CurIdx
        picOutput.Visible = True
        picConversions.Visible = False
        cmdNext.Enabled = False
        cmdMake.Enabled = True
        cmdMake.SetFocus
    End If
End Sub

Private Sub cmdStart_Click()
    ' If the caption says restart, reset the form
    If cmdStart.Caption = lang.GetString(339) Then
        ' If the remember option isn't set, reset otherwise keep it there
        If Not IS_SET(Opt.GenOptions, OPT_REMEMBER) Then
            txtDatabase.Text = ""
            txtFileName.Text = ""
            txtDirectory.Text = ""
            lblNewName1.Caption = ""
            lblNewName2.Caption = ""
            lblNewName3.Caption = ""
            lblFileName.Caption = ""
            lblConvert1.Caption = ""
            lblConvert2.Caption = ""
            lblConvert3.Caption = ""
            lblConvert4.Caption = ""
            chkPHP.Value = vbUnchecked
            chkMySQL.Value = vbUnchecked
            chkOracle.Value = vbUnchecked
            chkData.Value = vbUnchecked
            chkUpload.Value = vbUnchecked
            chkPost.Value = vbUnchecked
        End If
        ' Reset checkboxes/percentages/buttons
        lblPct.Caption = "0%"
        chkDB.Visible = False
        xDB.Visible = True
        chkConversions.Visible = False
        xConversions.Visible = True
        chkConvert.Visible = False
        xConvert.Visible = True
        chkOutput.Visible = False
        xOutput.Visible = True
        chkFinished.Visible = False
        xFinished.Visible = True
        cmdNext.Enabled = True
        cmdStart.Enabled = False
        cmdMake.Enabled = False
        cmdBrowse.Enabled = True
        cmdBrowse2.Enabled = True
        picIntro.Visible = True
        picConvert.Visible = False
        picFinished.Visible = False
        cmdStart.Caption = lang.GetString(241)
        cmdStart.Enabled = False
    Else
        ' Set up the last panel to show you what we actually did
        ProgressBar1.Value = 100
        xFinished.Visible = False
        chkFinished.Visible = True
        picFinished.Visible = True
        picConvert.Visible = False
        Dim i As Integer
        For i = 0 To Len(txtDatabase.Text)
            If Mid(txtDatabase.Text, (Len(txtDatabase.Text) - i), 1) = "\" Then
                lblFileName = Mid(txtDatabase.Text, Len(txtDatabase.Text) - i + 1, i)
                Exit For
            End If
        Next i
        If chkPHP.Value = vbChecked Then
            lblNewName1.Caption = LCase(txtFileName.Text & Opt.IO_PHP_Ext)
        End If
        If chkMySQL.Value = vbChecked Then
            If lblNewName1.Caption = "" Then
                lblNewName1.Caption = LCase(txtFileName.Text & Opt.IO_MySQL_Ext)
            Else
                lblNewName2.Caption = LCase(txtFileName.Text & Opt.IO_MySQL_Ext)
            End If
        End If
        If chkOracle.Value = vbChecked Then
            If lblNewName1.Caption = "" Then
                lblNewName1.Caption = LCase(txtFileName.Text & Opt.IO_Oracle_Ext)
            ElseIf lblNewName2.Caption = "" Then
                lblNewName2.Caption = LCase(txtFileName.Text & Opt.IO_Oracle_Ext)
            Else
                lblNewName3.Caption = LCase(txtFileName.Text & Opt.IO_Oracle_Ext)
            End If
        End If
        If chkPost.Value = vbChecked Then
            If lblNewName1.Caption = "" Then
                lblNewName1.Caption = LCase(txtFileName.Text & "_postgre.sql")
            ElseIf lblNewName2.Caption = "" Then
                lblNewName2.Caption = LCase(txtFileName.Text & "_postgre.sql")
            ElseIf lblNewName3.Caption = "" Then
                lblNewName3.Caption = LCase(txtFileName.Text & "_postgre.sql")
            Else
                lblNewName4.Caption = LCase(txtFileName.Text & "_postgre.sql")
            End If
        End If
        cmdStart.Enabled = True
        cmdBrowse.Enabled = False
        cmdBrowse2.Enabled = False
        cmdStart.Caption = lang.GetString(339)
    End If
End Sub

Private Sub Form_Activate()
    mnuLangSub(0).Visible = False
    
    ' If the intro form is visible
    If picIntro.Visible = True Then
        ' This flag was originally for the tips, but since then there are a few
        ' other startup events that happen. Now we have auto update and donation
        ' forms to show
        If showtips = True Then
            If doingtips = False Then
                doingtips = True
                ' Check autoupdate
                If IS_SET(Opt.GenOptions, OPT_UPDATE) Then
                    mnuCheck_Click
                End If
                ' Show the donation form
                If IS_SET(Opt.GenOptions, OPT_DONATE) Then
                    If frmDonation.Visible = False Then
                        frmDonation.Show vbModal
                    End If
                End If
                ' Show our tips
                If IS_SET(Opt.GenOptions, OPT_SHOWTIP) Then
                    If frmTip.Visible = False Then
                        frmTip.Show vbModal, Me
                    End If
                End If
                ' Set the focus to the next button
                cmdNext.SetFocus
                showtips = False
            End If
        End If
    ElseIf picOutput.Visible = True Then
        ' Set the "Make" button to be our focus
        cmdMake.SetFocus
    End If
    DoEvents
    showtips = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' Make sure we can't close while converting
    If working = True Then
        MsgBox lang.GetString(385), vbCritical, App.Title
        Cancel = vbCancel
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Close form
    modGlobals.SaveOpt
    Unload Me
End Sub

Private Sub mnuCheck_Click()
    ' Check internet connection and then check updates
    If checkInternet Then
        Dim newversion As Boolean
        newversion = checkVersion
        If Not newversion Then
            Dim comp As Integer
            comp = compareVersions(modGlobals.cur_version, modGlobals.new_version)
            If comp > 0 Then
                Dim answer As Integer
                answer = MsgBox(lang.GetString(386) & vbCrLf & vbCrLf & lang.GetString(387) & " " & modGlobals.cur_version & vbCrLf & lang.GetString(388) & "  " & modGlobals.new_version & vbCrLf & vbCrLf & lang.GetString(389), vbInformation + vbYesNo, App.Title)
                If answer = vbYes Then
                    dl_New
                End If
            ElseIf comp < 0 And showtips = False Then
                MsgBox lang.GetString(390) & vbCrLf & vbCrLf & lang.GetString(387) & " " & modGlobals.cur_version & vbCrLf & lang.GetString(391) & " " & modGlobals.new_version, vbInformation, App.Title
            End If
        ElseIf ((newversion = True) And (showtips = False)) Then
            MsgBox lang.GetString(392) & " " & vbCrLf & vbCrLf & lang.GetString(387) & "   " & modGlobals.cur_version & vbCrLf & lang.GetString(391) & " " & modGlobals.new_version, vbInformation, App.Title
        End If
    End If
End Sub

Private Sub mnuCheckLang_Click()
    frmGetLanguages.Show vbModal, Me
End Sub

' Cut/copy/paste functions
Private Sub mnuEdit_Copy_Click()
    On Error Resume Next
    Dim ctl As Control
    Set ctl = Me.ActiveControl
    If ctl.SelText <> "" Then
        Clipboard.SetText ctl.SelText
    End If
    Set ctl = Nothing

End Sub

Private Sub mnuEdit_Cut_Click()
    On Error GoTo err
    Dim ctl As Control
    Set ctl = Me.ActiveControl
    If ctl.SelText <> "" Then
        Clipboard.SetText ctl.SelText
        ctl.SelText = vbNullString
    End If
    Set ctl = Nothing

err:
End Sub

Private Sub mnuEdit_Paste_Click()
    On Error Resume Next
    Dim ctl As Control
    Set ctl = Me.ActiveControl
    ctl.SelText = Clipboard.GetText
    Set ctl = Nothing
End Sub

' Show our options
Private Sub mnuEdit_Options_Click()
    frmOptions.Show vbModal, Me
End Sub

' Kill the form
Private Sub mnuFile_Exit_Click()
    Unload Me
End Sub

' Show the about form
Private Sub mnuHelp_About_Click()
    cmdAbout_Click
End Sub

' Show the help contents
Private Sub mnuHelp_Contents_Click()
    Call ShellExecute(0&, vbNullString, "file://" & App.Path & "\" & lang.GetHelpIndex, vbNullString, _
                        vbNullString, vbNormalFocus)
End Sub

' Show the donation form
Private Sub mnuHelpDonate_Click()
    frmDonation.Show
End Sub



Private Sub mnuLangSub_Click(Index As Integer)
    On Error GoTo err_handle
        
    ' Default to the first language (if possible)
    If Index = 0 Then
        Index = 1
    End If
    
    ' Set the Language Option
    Opt.Gen_Lang = lang.GetFileByIndex(Index)
    modGlobals.SaveOpt
    ' Set the Language for real now
    lang.SetLanguageByIndex Index
    
    ' Now swap out the strings for the new language
    lang.LoadFormStrings Me
    PrintStrings CurIdx
    
    ' Change the menus
    Dim i As Integer
    For i = 1 To mnuLangSub.count - 1
        mnuLangSub(i).Enabled = True
    Next i
    mnuLangSub(Index).Enabled = False
    Exit Sub
err_handle:
    ' Our only hard coded string
    MsgBox "DB Converter has hit a fatal error while attempting to load your language settings." & vbCrLf & "Is it possible you have no language packs currently installed?" & vbCrLf & vbCrLf & "DB Converter cannot continue and will shut down.", vbCritical + vbOKOnly, App.Title + " :: Unrecoverable Error!"
    End
    End Sub

' Make sure we chose an MDB file
Private Sub txtDatabase_Change()
        If UCase(Right(txtDatabase.Text, 3)) = "MDB" Then
            cmdBrowse2.Enabled = True
        Else
            cmdBrowse2.Enabled = False
        End If
End Sub


Private Sub Form_Load()
    
    CurIdx = 0
    
    LoadMenuImages
    
    ' Do all our multilingual stuff
    lang.LoadLanguages
        
    mnuLangSub_Click lang.GetIndexByFile(Opt.Gen_Lang)
    lang.LoadFormStrings Me
    PrintStrings CurIdx
    
    ' Set the original showtips flag
    showtips = True
    working = True
    
    working = False
    
    Me.Show
    DoEvents

    ' Kill the reg object to save mem
    Set reg = Nothing
    
End Sub

' All code below this line handles highlighting and return to tab processing
Private Sub txtDB_GotFocus()
    Highlight txtDB
End Sub

Private Sub txtDB_KeyPress(KeyAscii As Integer)
    MyTab KeyAscii
End Sub

Private Sub txtDirectory_GotFocus()
    Highlight txtDirectory
End Sub
Private Sub txtDirectory_KeyPress(KeyAscii As Integer)
    MyTab KeyAscii
End Sub

Private Sub txtFileName_GotFocus()
    Highlight txtFileName
End Sub
Private Sub txtFileName_KeyPress(KeyAscii As Integer)
    MyTab KeyAscii
End Sub

Private Sub txtDatabase_GotFocus()
    Highlight txtDatabase
End Sub
Private Sub txtDatabase_KeyPress(KeyAscii As Integer)
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

Private Sub PrintStrings(Index As Integer)
    If Index = 0 Then
        ' Print the first screen
        picIntro.Cls
        picIntro.Print lang.GetString(346) & " ";
        picIntro.Font.Bold = True
        picIntro.Print lang.GetString(347);
        picIntro.Font.Bold = False
        picIntro.Print "! " & lang.GetString(348)
        picIntro.Print lang.GetString(349)
        picIntro.Print lang.GetString(350)
        picIntro.Print lang.GetString(351) & "!"
        picIntro.Print
        picIntro.Print lang.GetString(352) & "!"
        picIntro.Visible = True
    ElseIf Index = 1 Then
        picDB.Cls
        picDB.Print lang.GetString(359)
        picDB.Print lang.GetString(360)
        picDB.Print
        picDB.Print lang.GetString(361) & " ";
        picDB.FontBold = True
        picDB.Print lang.GetString(362) & " ";
        picDB.FontBold = False
        picDB.Print lang.GetString(363) & "."
        picDB.Visible = True
    ElseIf Index = 2 Then
        picConversions.Cls
        picConversions.Print lang.GetString(365)
        picConversions.Print lang.GetString(366) & " ";
        picConversions.FontBold = True
        picConversions.Print lang.GetString(367);
        picConversions.FontBold = False
        picConversions.Print ",";
        picConversions.FontBold = True
        picConversions.Print " " & lang.GetString(368);
        picConversions.FontBold = False
        picConversions.Print ","
        picConversions.FontBold = True
        picConversions.Print lang.GetString(369);
        picConversions.FontBold = False
        picConversions.Print ",";
        picConversions.Print " " & lang.GetString(370) & " ";
        picConversions.FontBold = True
        picConversions.Print lang.GetString(371);
        picConversions.FontBold = False
        picConversions.Print ". " & lang.GetString(372)
        picConversions.Print lang.GetString(373)
    ElseIf Index = 3 Then
        picOutput.Cls
        picOutput.Print lang.GetString(379) & " ";
        picOutput.FontBold = True
        picOutput.Print lang.GetString(380) & " ";
        picOutput.FontBold = False
        picOutput.Print lang.GetString(381) & " "
        picOutput.FontBold = True
        picOutput.Print lang.GetString(382)
        picOutput.FontBold = False
        picOutput.Print
        picOutput.Print lang.GetString(383)
        picOutput.Print lang.GetString(384)
    End If
End Sub
