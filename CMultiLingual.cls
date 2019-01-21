VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMultiLingual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type LanguageType
    LanguageName As String
    FileName As String
    Icon As StdPicture
    Help As String
End Type

Private Type StringType
    Index As Integer
    StringVal As String
End Type

Private Current_Lang As Integer
Private Languages() As LanguageType
Private Strings() As StringType

Public Function LoadLanguages()
    Dim fname As String
    Dim idx As Long
        
    idx = 0
    
    fname = Dir(App.Path & "\Languages\*.lang")

    While fname <> ""
        Open App.Path & "\Languages\" & fname For Input As #1
        Dim line As String
        Dim vals() As String

        
        Line Input #1, line
        If UCase(line) = "[LANGUAGE]" Then
            idx = idx + 1
            ReDim Preserve Languages(idx)
            
            Languages(idx).FileName = fname
            
            Line Input #1, line
            vals = Split(line, "=")
            Languages(idx).LanguageName = vals(1)
    
            Line Input #1, line
            vals = Split(line, "=")
            Set Languages(idx).Icon = LoadPicture(App.Path & "\Languages\" & vals(1))
            
            Line Input #1, line
            vals = Split(line, "=")
            Languages(idx).Help = vals(1)
            
            With mainform
                Dim hMenu As Long
                Dim hSubMenu As Long
                Dim ret As Long
                                
                Load .mnuLangSub(idx)
                .mnuLangSub(idx).Caption = Languages(idx).LanguageName
                
                hMenu = modGlobals.GetMenu(mainform.hwnd)
                hSubMenu = modGlobals.GetSubMenu(hMenu, 2)
                Call SetMenuItemBitmaps(hSubMenu, idx, MF_BYPOSITION, Languages(idx).Icon, Languages(idx).Icon)
            End With
            
        End If
        Close
        fname = Dir
    Wend
End Function

Public Sub SetLanguageByIndex(Index As Integer)
    On Error GoTo err_handle:
    Call SetLanguageByFile(Languages(Index).FileName)
err_handle:
    Close
    Exit Sub
End Sub

Public Sub SetLanguageByFile(FileName As String)
    On Error GoTo err_handle:
    Call LoadStrings(FileName)
err_handle:
    Close
    Exit Sub
End Sub

Private Sub LoadStrings(FileName As String)
    
    Dim idx As Integer
    Dim line As String
    Dim vals() As String
    
    Current_Lang = GetIndexByFile(FileName)
    
    idx = 0
    ReDim Strings(1)
    
    Open App.Path & "\Languages\" & FileName For Input As #1
    Do
        Line Input #1, line
    Loop Until UCase(line) = "[STRINGS]"
    While Not EOF(1)
        Line Input #1, line
        vals = Split(line, "=")
        idx = idx + 1
        
        ReDim Preserve Strings(idx)
        
        Strings(idx).Index = Val(vals(0))
        Strings(idx).StringVal = vals(1)
    Wend
    Close
End Sub

Public Function GetString(Index As Integer) As String
    On Error GoTo err_handle:
    GetString = "Unknown String"
    
    Dim i As Integer
    
    For i = 1 To UBound(Strings)
        If Strings(i).Index = Index Then
            GetString = Strings(i).StringVal
            Exit Function
        End If
    Next i
     
    Exit Function
err_handle:
    GetString = "Unknown String"
End Function

Public Function GetFileByIndex(Index As Integer) As String
    On Error GoTo err_handle
    GetFileByIndex = "ERROR"
    GetFileByIndex = Languages(Index).FileName
err_handle:
End Function

Public Function GetIndexByFile(FileName As String) As Integer
    On Error GoTo err_handle
    Dim i As Integer
    GetIndexByFile = 0
    For i = 1 To UBound(Languages)
        If Languages(i).FileName = FileName Then
            GetIndexByFile = i
            Exit Function
        End If
    Next i
err_handle:
End Function

Public Function GetIndexByName(Name As String) As Integer
    On Error GoTo err_handle
    Dim i As Integer
    GetIndexByName = 0
    For i = 1 To UBound(Languages)
        If UCase(Languages(i).LanguageName) = UCase(Name) Then
            GetIndexByName = i
            Exit Function
        End If
    Next i
err_handle:
End Function

Public Sub LoadFormStrings(f As Form)
    On Error GoTo err_handle

    Dim c As Control
    
    If f.Tag <> "" Then f.Caption = lang.GetString(Val(f.Tag))
    For Each c In f.Controls
        If c.Name = "SSTab1" Then
            Dim i As Integer
            Dim Strings() As String
            Strings = Split(c.Tag, ";")
            For i = 0 To UBound(Strings)
                c.Tab = i
                c.Caption = GetString(Val(Strings(i)))
            Next i
            c.Tab = 0
        Else
            If c.Tag <> "" Then
                c.Caption = GetString(Val(c.Tag))
            End If
            If c.ToolTipText <> "" Then
                c.ToolTipText = GetString(Val(c.ToolTipText))
            End If
        End If
    Next
    Exit Sub
err_handle:
    If err.Number = 438 Then
        Resume Next
    Else: MsgBox GetString(334) & " :: " & err.Description, vbCritical + vbOKOnly, App.Title
    End If
End Sub

Public Function GetHelpIndex() As String
    GetHelpIndex = Languages(Current_Lang).Help
End Function

Public Function GetCount() As Integer
    GetCount = UBound(Languages)
End Function
