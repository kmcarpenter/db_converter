VERSION 5.00
Begin VB.Form frmGetLanguages 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "394"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4530
   Icon            =   "frmGetLanguages.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "394"
   Begin VB.CommandButton cmdClose 
      Caption         =   "132"
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Tag             =   "132"
      ToolTipText     =   "134"
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdDownload 
      Caption         =   "397"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Tag             =   "397"
      ToolTipText     =   "396"
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "398"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Tag             =   "398"
      ToolTipText     =   "395"
      Top             =   2640
      Width           =   1335
   End
   Begin VB.ListBox lstPacks 
      Height          =   1860
      ItemData        =   "frmGetLanguages.frx":0442
      Left            =   0
      List            =   "frmGetLanguages.frx":0444
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label lblStatus 
      Caption         =   "399"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Tag             =   "399"
      Top             =   2040
      Width           =   4215
   End
End
Attribute VB_Name = "frmGetLanguages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Type LangDownload
    file_name As String
    file_size As Long
    db_index As Long
    lang_name As String
End Type

Private downloading As Boolean
Private avail() As LangDownload

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDownload_Click()
    Dim hInet As Long
    Dim hUrl As Long
    Dim hResult As Long
    Dim Flags As Long
    Dim url As Variant
    Dim i As Integer
    Dim unZip As New CGUnzipFiles
    
    lblStatus.Caption = lang.GetString(403)
    DoEvents
    
    ' Download the language packs which are picked
    For i = 0 To lstPacks.ListCount - 1
        If lstPacks.Selected(i) Then
            lblStatus.Caption = lang.GetString(405) & lstPacks.List(i)
            
            hInet = InternetOpen(App.Title, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0&)

            If hInet Then
                Dim bRead As Long
                Dim bToRead As Long
                Dim bLeft As Long
                Dim sBuffer As String
                
                Dim size As String
                
                Flags = INTERNET_FLAG_KEEP_CONNECTION Or INTERNET_FLAG_NO_CACHE_WRITE Or INTERNET_FLAG_RELOAD
            
                hUrl = InternetOpenUrl(hInet, avail(i + 1).file_name, vbNullString, 0, Flags, 0)
                If hUrl Then
                    buf = ""
                    bToRead = 1024 ' 1024 byte chunk. Should be WAY more than necessary for this
                    
                    Do
                    
                        sBuffer = String(bToRead, 0)
                        hResult = InternetReadFile(hUrl, sBuffer, bToRead, bRead)
                        
                        buf = buf & Left(sBuffer, bRead)
                        
                    Loop Until bRead < bToRead
            
                    Call InternetCloseHandle(hUrl)
                    
                    Open App.Path & "\" & avail(i + 1).lang_name & ".zip" For Binary As #1
                    Put #1, , buf
                    Close #1
                    
                    lblStatus.Caption = lang.GetString(406) & lstPacks.List(i)
                    ' Unzip the files
                    unZip.ZipFileName = App.Path & "\" & avail(i + 1).lang_name & ".zip"
                    unZip.ExtractDir = App.Path & "\"
                    unZip.HonorDirectories = True
                    unZip.unZip
                    
                    Kill App.Path & "\" & avail(i + 1).lang_name & ".zip"
                    
                End If
                
                Call InternetCloseHandle(hInet)
            End If
        End If
    Next i

    lstPacks.Clear
    
    cmdDownload.Enabled = False
    cmdSearch.Enabled = True

    lblStatus.Caption = lang.GetString(355)
End Sub

Private Sub cmdSearch_Click()
    Dim langCount As Integer
    Dim list_first() As String
    Dim list_second() As String
    Dim i As Integer
    
    lblStatus.Caption = lang.GetString(401)
    DoEvents
        
    ' Then download a list of currently available language packs
    ' http://www.zenwerx.com/languages.php
    Dim hInet As Long
    Dim hUrl As Long
    Dim hResult As Long
    Dim Flags As Long
    Dim url As Variant
    
    Dim arr() As String
    Dim buf As String
    
    ' Generate version string
    cur_version = App.Major & "." & App.Minor & "." & App.Revision
    
    hInet = InternetOpen(App.Title, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0&)

    If hInet Then
        Dim bRead As Long
        Dim bToRead As Long
        Dim bLeft As Long
        Dim sBuffer As String
        
        Dim size As String
        
        Flags = INTERNET_FLAG_KEEP_CONNECTION Or INTERNET_FLAG_NO_CACHE_WRITE Or INTERNET_FLAG_RELOAD
    
        hUrl = InternetOpenUrl(hInet, "http://www.zenwerx.com/languages.php?action=0&version=" & cur_version, vbNullString, 0, Flags, 0)
        If hUrl Then
            buf = ""
            bToRead = 1024 ' 1024 byte chunk. Should be WAY more than necessary for this
            
            Do
            
                sBuffer = String(bToRead, 0)
                hResult = InternetReadFile(hUrl, sBuffer, bToRead, bRead)
                
                buf = buf & Left(sBuffer, bRead)
                
            Loop Until bRead < bToRead
    
            Call InternetCloseHandle(hUrl)
            
            
            ' Loop through the two lists, and add those which do not exist
            list_first = Split(buf, "|")
            For i = 0 To UBound(list_first) - 1
                list_second = Split(list_first(i), ",")
                If (lang.GetIndexByName(list_second(1)) = 0) Then
                    ' Add to list
                    Dim idx As Integer
                    idx = UBound(avail) + 1
                    ReDim Preserve avail(idx)
                    avail(idx).lang_name = list_second(1)
                    avail(idx).db_index = Val(list_second(0))
                    
                    ' Get the rest of the pack info right now
                    hUrl = InternetOpenUrl(hInet, "http://www.zenwerx.com/languages.php?action=1&lang=" & avail(idx).db_index, vbNullString, 0, Flags, 0)
                    If hUrl Then
                        buf = ""
                        Do
            
                            sBuffer = String(bToRead, 0)
                            hResult = InternetReadFile(hUrl, sBuffer, bToRead, bRead)
                            
                            buf = buf & Left(sBuffer, bRead)
                            
                        Loop Until bRead < bToRead
    
                        Call InternetCloseHandle(hUrl)
                        
                        list_first = Split(buf, ",")
                        avail(idx).file_name = list_first(0)
                        avail(idx).file_size = Val(list_first(1))
                        
                        lstPacks.AddItem avail(idx).lang_name & " ( " & Int(avail(idx).file_size / 1024) & "KB )"
                        langCount = langCount + 1
                    End If
                End If
            Next i
            
            
            
            Call InternetCloseHandle(hInet)
        End If
    Else
        MsgBox lang.GetString(335), vbCritical + vbOKOnly, App.Title
        Exit Sub
    End If
    
    If (langCount > 0) Then
        cmdDownload.Enabled = True
        cmdSearch.Enabled = False
        lblStatus.Caption = lang.GetString(402)
    Else
        lblStatus.Caption = lang.GetString(404)
    End If
    
End Sub

Private Sub Form_Load()
    downloading = False
    lang.LoadFormStrings Me
    ReDim avail(0)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If downloading = True Then
        MsgBox lang.GetString(400), vbInformation + vbOKOnly, App.Title
        Cancel = 1
    End If
End Sub
