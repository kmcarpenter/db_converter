VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AccessWiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public Function Create(Optional Path As String = "")
    Load frmMain
    If Path <> "" Then
        Dim temp_path As String
        temp_path = Mid(Path, 2, Len(Path) - 2)
            If FileLen(temp_path) > 0 Then
                LoadFile = True
                frmMain.cmdNext_Click
                LoadFile = False
                frmMain.txtDatabase.Text = temp_path
                Dim i As Integer
                Dim newfile As String
                For i = 0 To Len(temp_path)
                    If Mid(temp_path, (Len(temp_path) - i), 1) = "\" Then
                        newfile = Mid(temp_path, (Len(temp_path) - i + 1), i - 4)
                        If Opt.Gen_AssumeOutput = vbChecked Then
                            frmMain.txtDirectory.Text = Opt.IO_OutputDir
                            frmMain.txtFileName.Text = newfile
                        End If
                        Exit For
                    End If
                Next i
            End If
    End If
    frmMain.Show vbModal
End Function
