Attribute VB_Name = "modGlobals"
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
' *
' *  This file is part of DB Converter 1.6.0.0
' *
' *  DB Converter 1.6.0.0 is free software; you can redistribute it and/or
' *  modify it under the terms of the GNU General Public License as published by
' *  the Free Software Foundation; either version 2 of the License, or
' *  (at your option) any later version.
' *
' *  DB Converter 1.6.0.0 is distributed in the hope that it will be useful,
' *  but WITHOUT ANY WARRANTY; without even the implied warranty of
' *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' *  GNU General Public License for more details.'
' *
' *  You should have received a copy of the GNU General Public License
' *  along with DB Converter 1.6.0.0; if not, write to the Free Software
' *  Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
' *
' ***************************************************************************

' ***************************************************************************
' Force variable declaration
' ***************************************************************************
Option Explicit

' ***************************************************************************
' * Declares for direct ping/download
' ***************************************************************************
Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal lpszAgent As String, ByVal dwAccessType As Long, ByVal lpszProxyName As String, ByVal lpszProxyBypass As String, ByVal dwFlags As Long) As Long
Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInet As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Long
Declare Function InternetReadFile Lib "wininet" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function gethostname Lib "WSOCK32" (ByVal szHost As String, ByVal dwHostLen As Long) As Long

' ***************************************************************************
' * SendMessage Delcare... various uses
' ***************************************************************************
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

' ***************************************************************************
' * Declares for setting icons in the menus
' ***************************************************************************
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long


' ***************************************************************************
' * Constants for graphical menus
' ***************************************************************************
Public Const MF_BYPOSITION = &H400&

' ***************************************************************************
' * Constants for internet use
' ***************************************************************************
Public Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Public Const INTERNET_FLAG_RELOAD = &H80000000
Public Const INTERNET_FLAG_KEEP_CONNECTION = &H400000
Public Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000

' ***************************************************************************
' * Constant for finding strings in the tables listbox
' ***************************************************************************
Public Const LB_FINDSTRING = &H18C
Public Const LB_ERR = (-1)

' ***************************************************************************
' * Options enum (options bitvectors)
' ***************************************************************************
Public Enum OPT_ENUM
    OPT_REMEMBER = 1
    OPT_ASSUME = 2
    OPT_ENTER = 4
    OPT_HIGHLIGHT = 8
    OPT_SHOWTIP = 16
    OPT_WARNLINK = 32
    OPT_DONATE = 64
    OPT_UPDATE = 128
    OPT_TABLES = 256
    OPT_SURVEY = 512 ' Defunct
    OPT_INNODB = 1024
End Enum

' ***************************************************************************
' * Structure for holding options information
' ***************************************************************************
Type OptType
    'Options user doesn't see
    Gen_LastTip As Integer
    Gen_Path As String
    Gen_Lang As String
    'General Options
    GenOptions As Long
    'Input/Output
    IO_BrowseDir As String
    IO_OutputDir As String
    IO_PHP_Ext As String
    IO_MySQL_Ext As String
    IO_Oracle_Ext As String
    'PHP Settings
    PHP_RememberPHP As Integer
    PHP_DB As String
    PHP_Host As String
    PHP_User As String
    PHP_Pass As String
    'Upload settings
    FTP_Host As String
    FTP_User As String
    FTP_Pass As String
    FTP_Port As String
    'MySQL Settings
    SQL_Host As String
    SQL_User As String
    SQL_Pass As String
    SQL_Port As String
    SQL_DB As String
End Type

' ***************************************************************************
' * Global scope for options
' ***************************************************************************
Global Opt As OptType

' ***************************************************************************
' * Global scope for registry editing
' ***************************************************************************
Global reg As CRegSettings

' ***************************************************************************
' * Global scope for Language Data
' ***************************************************************************
Global lang As CMultiLingual

' ***************************************************************************
' * Global variables for returning data between forms
' ***************************************************************************
Global ReturnDir As String
Global DB_Pass As String
Global DB_User As String

' ***************************************************************************
' * Global FTP Form
' ***************************************************************************
Global FTP_Path As String

' ***************************************************************************
' * Global filenames for FTP Form to read
' ***************************************************************************
Global PHP_Name As String
Global SQL_Name As String
Global Ora_Name As String
Global Post_Name As String

' ***************************************************************************
' * Global declarations for dummy download form to read
' ***************************************************************************
Global cur_version As String
Global new_version As String
Global filesize As Long
Global currentstate As Long
Global totalstatements As Long
Global offset As Long

' ***************************************************************************
' * Global flag for computing
' ***************************************************************************
Global working As Boolean

' ***************************************************************************
' * Global error checking
' ***************************************************************************
Global goterr As Boolean

' ***************************************************************************
' * Global variable for main form
' ***************************************************************************
Global mainform As frmMain

' ***************************************************************************
' * Module variables
' ***************************************************************************
Dim seed As Variant
Dim checkType As Integer
Dim x As Long

Private Declare Function FormatMessage Lib "kernel32" _
  Alias "FormatMessageA" _
  (ByVal dwFlags As Long, _
   lpSource As Long, _
   ByVal dwMessageId As Long, _
   ByVal dwLanguageId As Long, _
   ByVal lpBuffer As String, _
   ByVal nSize As Long, _
   Args As Any) As Long

Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000

Private Const ID_DONATE = 101
Private Const ID_COPY = 102
Private Const ID_CUT = 103
Private Const ID_HELP = 104
Private Const ID_PASTE = 105
Private Const ID_OPTIONS = 106
Private Const ID_QUIT = 107

Private Const MENU_PIC_COUNT = 6 ' 0 based

Private Pictures() As StdPicture


' ***************************************************************************
' * Convert enter (return) to a tab
' ***************************************************************************
Public Sub MyTab(Key As Integer)
    If IS_SET(Opt.GenOptions, OPT_ENTER) Then
        If Key = 13 Then
            ' Not a perfect solution with sendkeys, but it works
            SendKeys (vbTab)
        End If
    End If
End Sub

' ***************************************************************************
' * Highilght text on focus
' ***************************************************************************
Public Sub Highlight(Box As TextBox)
    If IS_SET(Opt.GenOptions, OPT_HIGHLIGHT) Then
        Box.SelStart = 0
        Box.SelLength = Len(Box.Text)
    End If
End Sub

' ***************************************************************************
' * Save options to registry
' ***************************************************************************
Public Sub SaveOpt()
    Set reg = New CRegSettings
        
    ' Save general settings
    reg.SaveSetting "Software\Zenwerx\DB Converter\Settings\", "Path", Opt.Gen_Path
    reg.SaveSetting "Software\Zenwerx\DB Converter\Settings\", "Language", Opt.Gen_Lang
    reg.SaveSetting "Software\Zenwerx\DB Converter\Settings\General", "General", Str(Opt.GenOptions)
    ' Save input/output settings
    reg.SaveSetting "Software\Zenwerx\DB Converter\Settings\IO", "BrowseDir", Opt.IO_BrowseDir
    reg.SaveSetting "Software\Zenwerx\DB Converter\Settings\IO", "MySQL_Ext", Opt.IO_MySQL_Ext
    reg.SaveSetting "Software\Zenwerx\DB Converter\Settings\IO", "Oracle_Ext", Opt.IO_Oracle_Ext
    reg.SaveSetting "Software\Zenwerx\DB Converter\Settings\IO", "OutputDir", Opt.IO_OutputDir
    reg.SaveSetting "Software\Zenwerx\DB Converter\Settings\IO", "PHP_Ext", Opt.IO_PHP_Ext
    ' Save php settings
    reg.SaveSetting "Software\Zenwerx\DB Converter\Settings\PHP", "DB", Opt.PHP_DB
    reg.SaveSetting "Software\Zenwerx\DB Converter\Settings\PHP", "Host", Opt.PHP_Host
    reg.SaveSetting "Software\Zenwerx\DB Converter\Settings\PHP", "Pass", CryptLine(Opt.PHP_Pass, True)
    reg.SaveSetting "Software\Zenwerx\DB Converter\Settings\PHP", "User", Opt.PHP_User
    reg.SaveSetting "Software\Zenwerx\DB Converter\Settings\PHP", "RememberPHP", Str(Opt.PHP_RememberPHP)
    ' Save ftp settings
    reg.SaveSetting "Software\Zenwerx\DB Converter\Settings\FTP", "Host", Opt.FTP_Host
    reg.SaveSetting "Software\Zenwerx\DB Converter\Settings\FTP", "Pass", CryptLine(Opt.FTP_Pass, True)
    reg.SaveSetting "Software\Zenwerx\DB Converter\Settings\FTP", "Port", Opt.FTP_Port
    reg.SaveSetting "Software\Zenwerx\DB Converter\Settings\FTP", "User", Opt.FTP_User
    ' Save MySQL settings
    reg.SaveSetting "Software\Zenwerx\DB Converter\Settings\MySQL", "Host", Opt.SQL_Host
    reg.SaveSetting "Software\Zenwerx\DB Converter\Settings\MySQL", "Pass", CryptLine(Opt.SQL_Pass, True)
    reg.SaveSetting "Software\Zenwerx\DB Converter\Settings\MySQL", "Port", Opt.SQL_Port
    reg.SaveSetting "Software\Zenwerx\DB Converter\Settings\MySQL", "User", Opt.SQL_User
    reg.SaveSetting "Software\Zenwerx\DB Converter\Settings\MySQL", "DB", Opt.SQL_DB
    
    ' Decrypt passwords
    ' * Note:
    ' *       This is because the "Cryptline" functions works with a reference variable
    ' *       so when the password gets encrypted, it changes the options variable that
    ' *       we are working with. It has to be manually decrypted again (here for lack of
    ' *       a better place) for use later.
    Opt.SQL_Pass = CryptLine(reg.GetSetting("Software\Zenwerx\DB Converter\Settings\MySQL", "Pass", ""), False)
    Opt.FTP_Pass = CryptLine(reg.GetSetting("Software\Zenwerx\DB Converter\Settings\FTP", "Pass", ""), False)
    Opt.PHP_Pass = CryptLine(reg.GetSetting("Software\Zenwerx\DB Converter\Settings\PHP", "Pass", ""), False)
    
    Set reg = Nothing
End Sub

' ***************************************************************************
' * Get options from registry
' ***************************************************************************
Public Sub GetOpt()
    Set reg = New CRegSettings
    
    ' Get general options
    Opt.GenOptions = Val(reg.GetSetting("Software\Zenwerx\DB Converter\Settings\General", "General", Str(OPT_ASSUME + OPT_ENTER + OPT_HIGHLIGHT + OPT_SHOWTIP + OPT_DONATE + OPT_UPDATE)))
    Opt.Gen_Lang = reg.GetSetting("Software\Zenwerx\DB Converter\Settings\", "Language", "English.lang")
    Opt.Gen_Path = reg.GetSetting("Software\Zenwerx\DB Converter\Settings\", "Path", "")
    ' Get input/output options
    Opt.IO_BrowseDir = reg.GetSetting("Software\Zenwerx\DB Converter\Settings\IO", "BrowseDir", App.Path & "\")
    Opt.IO_MySQL_Ext = reg.GetSetting("Software\Zenwerx\DB Converter\Settings\IO", "MySQL_Ext", "_mysql.sql")
    Opt.IO_Oracle_Ext = reg.GetSetting("Software\Zenwerx\DB Converter\Settings\IO", "Oracle_Ext", "_oracle.sql")
    Opt.IO_OutputDir = reg.GetSetting("Software\Zenwerx\DB Converter\Settings\IO", "OutputDir", App.Path & "\")
    Opt.IO_PHP_Ext = reg.GetSetting("Software\Zenwerx\DB Converter\Settings\IO", "PHP_Ext", "_create.php")
    ' get php options
    Opt.PHP_DB = reg.GetSetting("Software\Zenwerx\DB Converter\Settings\PHP", "DB", "")
    Opt.PHP_Host = reg.GetSetting("Software\Zenwerx\DB Converter\Settings\PHP", "Host", "")
    Opt.PHP_Pass = CryptLine(reg.GetSetting("Software\Zenwerx\DB Converter\Settings\PHP", "Pass", ""), False)
    Opt.PHP_User = reg.GetSetting("Software\Zenwerx\DB Converter\Settings\PHP", "User", "")
    Opt.PHP_RememberPHP = Val(reg.GetSetting("Software\Zenwerx\DB Converter\Settings\PHP", "RememberPHP", vbUnchecked))
    ' get ftp options
    Opt.FTP_Host = reg.GetSetting("Software\Zenwerx\DB Converter\Settings\FTP", "Host", "localhost")
    Opt.FTP_Pass = CryptLine(reg.GetSetting("Software\Zenwerx\DB Converter\Settings\FTP", "Pass", ""), False)
    Opt.FTP_Port = reg.GetSetting("Software\Zenwerx\DB Converter\Settings\FTP", "Port", "21")
    Opt.FTP_User = reg.GetSetting("Software\Zenwerx\DB Converter\Settings\FTP", "User", "")
    ' get sql options
    Opt.SQL_Host = reg.GetSetting("Software\Zenwerx\DB Converter\Settings\MySQL", "Host", "localhost")
    Opt.SQL_Pass = CryptLine(reg.GetSetting("Software\Zenwerx\DB Converter\Settings\MySQL", "Pass", ""), False)
    Opt.SQL_Port = reg.GetSetting("Software\Zenwerx\DB Converter\Settings\MySQL", "Port", "3306")
    Opt.SQL_User = reg.GetSetting("Software\Zenwerx\DB Converter\Settings\MySQL", "User", "root")
    Opt.SQL_DB = reg.GetSetting("Software\Zenwerx\DB Converter\Settings\MySQL", "DB", "mydb")
    
    Set reg = Nothing
End Sub

' ***************************************************************************
' * Generate encryption key
' ***************************************************************************
Private Function MyKey() As String
    Dim stupidkey As String
    Dim crapkey As String
    
    ' Note:
    ' *     Hardcoded key generated by single characters. This way the key isn't stored in
    ' *     the string table as one large string. Storing the string this way makes it
    ' *     harder to read with a hex editor.
    
    ' Note 2:
    ' *     Being released under GNU GPL will probably undermine this encryption technique
    ' *     because my keys will be released publicly. The encryption isn't all that
    ' *     strong anyway... But, for backwards capability please do not change the keys
    ' *     or at least make a version checking encryption routine
    
    MyKey = "S" & "0" & "T" & "P" & "9" & "W" & "8" & "7" & "7" & "5" & "J" & "D" & "B" & "I" & "C" & "M" & "3" & "1" & "6" & "7" & "X" & "D" & "Y" & "C"
    MyKey = MyKey & "R" & "G" & "2" & "B" & "N" & "S" & "T" & "1" & "9" & "J" & "1" & "Z" & "8" & "K" & "F" & "8" & "C" & "O" & "Z" & "7" & "O" & "7" & "M" & "S" & "S" & "F" & "N"
    stupidkey = "B" & "T" & "J" & "S" & "8" & "8" & "N" & "T" & "Z" & "C" & "8" & "B" & "N" & "J" & "Y" & "Z" & "U" & "H" & "Y" & "V" & "3" & "A" & "G" & "B" & "7" & "A" & "0" & "G"
    crapkey = "S" & "L" & "V" & "D" & "I" & "M" & "6" & "9" & "T" & "0" & "5" & "S" & "5" & "2" & "P" & "V" & "Y" & "V" & "Z" & "Q" & "B"
    MyKey = MyKey & stupidkey & crapkey
End Function

' ***************************************************************************
' * Crypto algorithm
' ***************************************************************************
' *
' * I do not take credit for this code (any of the encryption code).
' * I cut and paste it into my project long ago, and can't seem to find
' * the original project anymore. It definately was free, but I just can't
' * give credit to the original author. If anyone recognizes this, please
' * feel free to add the author's name.
' *
' ***************************************************************************
Private Function CryptLine(StrIn As String, flag As Boolean) As String
    'Crypto algorithm
    Dim Sdvig As Byte
    Dim strout As String, StrPas As String, StrA As String
    Dim k As Integer, j As Integer
    Dim Lstr As Long, Pstr As Long
    Dim SeedStr As String
    Rnd (-1)
    For k = 1 To Len(MyKey())
        SeedStr = SeedStr & CStr(Asc(Mid(MyKey(), k, 1)))
    Next k
    seed = CVar(SeedStr)
    Randomize (seed)
    Lstr = Len(StrIn)
    
    If Lstr <> 0 Then
    strout = ""
    
    StrPas = MyKey()
    Pstr = Len(StrPas)
    k = 0
    'back interposition
    
    If Not flag Then
    For j = 1 To Lstr - 1 Step 2
        StrA = Mid(StrIn, j + 1, 1)
        Mid(StrIn, j + 1, 1) = Mid(StrIn, j, 1)
        Mid(StrIn, j, 1) = StrA
    Next j
    End If
    
    'RN & LC
    For j = 1 To Lstr ' Main
            k = k + 1
            If k > Pstr Then k = k - Pstr
              
            Sdvig = Delta(255)
            
            If flag Then 'code
                
                Mid(StrIn, j, 1) = Chr(SumMod(Asc(Mid(StrIn, j, 1)), Asc(Mid(StrPas, k, 1)), 255))
                Mid(StrIn, j, 1) = Chr(SumMod(Asc(Mid(StrIn, j, 1)), Sdvig, 255))
                                    
                strout = strout & Mid(StrIn, j, 1)
              
            Else 'decode
    
                Mid(StrIn, j, 1) = Chr(SubstrMod(Asc(Mid(StrIn, j, 1)), Sdvig, 255))
                Mid(StrIn, j, 1) = Chr(SubstrMod(Asc(Mid(StrIn, j, 1)), Asc(Mid(StrPas, k, 1)), 255))
              
                strout = strout & Mid(StrIn, j, 1)
              
            End If
          
    Next j
    'forward interposition
    
    If flag Then
    For j = 1 To Lstr - 1 Step 2
        StrA = Mid(strout, j + 1, 1)
        Mid(strout, j + 1, 1) = Mid(strout, j, 1)
        Mid(strout, j, 1) = StrA
    Next j
    End If
    
    CryptLine = strout
    End If
End Function

Private Function SumMod(Num1 As Byte, Num2 As Byte, Modul As Byte) As Byte
'(Num1 + Num2) mod Modul
Dim Sum As Integer
Sum = CInt(Num1) + CInt(Num2)
Select Case Sum
    Case Is <= Modul
    SumMod = Sum
    Case Is > Modul
    SumMod = Sum - Modul
End Select
End Function
Private Function SubstrMod(Num1 As Byte, Num2 As Byte, Modul As Byte) As Byte
'(Num1 - Num2) mod Modul
Dim Substr As Integer
Substr = CInt(Num1) - CInt(Num2)
Select Case Substr
    Case Is <= 0
    SubstrMod = Substr + Modul
    Case Is > 0
    SubstrMod = Substr
End Select
End Function

' ***************************************************************************
' * Random number for adding
' ***************************************************************************
' *
' * This number is always the same, but it's better than hard coding a value
' *
' ***************************************************************************
Private Function Delta(Lim As Byte) As Byte
    Delta = Int((Rnd * Lim) + 1)
End Function

' ***************************************************************************
' * End Crypto Algorithm
' ***************************************************************************


' ***************************************************************************
' * Check for active internet connection
' ***************************************************************************
Public Function checkInternet() As Boolean
   Dim hInet As Long
   Dim hUrl As Long
   Dim Flags As Long
   Dim url As Variant
   'Let events fire
   DoEvents
   ' Open internet handle
   hInet = InternetOpen(App.Title, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0&)
   ' If we got a valid internet handle, try pinging our website
   ' This method isn't foolproof, but if our site's down you can't download the
   ' new version anyway
   If hInet Then
      Flags = INTERNET_FLAG_KEEP_CONNECTION Or INTERNET_FLAG_NO_CACHE_WRITE Or INTERNET_FLAG_RELOAD
      hUrl = InternetOpenUrl(hInet, "http://www.zenwerx.com", vbNullString, 0, Flags, 0)
      If hUrl Then
        ' We got a valid handle, internet connection is good
        ' Set the flag and close the handle
         checkInternet = True
         Call InternetCloseHandle(hUrl)
      Else
         checkInternet = False
      End If
      ' Close internet handle
      Call InternetCloseHandle(hInet)
   End If
End Function

' ***************************************************************************
' * Check current version of DB Converter
' ***************************************************************************
Public Function checkVersion() As Boolean
    Dim hInet As Long
    Dim hUrl As Long
    Dim hResult As Long
    Dim Flags As Long
    Dim url As Variant
    
    checkVersion = False
    
    ' Generate version string
    cur_version = App.Major & "." & App.Minor & "." & App.Revision & ".0"
    
    hInet = InternetOpen(App.Title, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0&)

    If hInet Then
        Dim bRead As Long
        Dim bToRead As Long
        Dim bLeft As Long
        Dim sBuffer As String
        
        Dim i As Integer
        Dim size As String
        
        Flags = INTERNET_FLAG_KEEP_CONNECTION Or INTERNET_FLAG_NO_CACHE_WRITE Or INTERNET_FLAG_RELOAD
    
        hUrl = InternetOpenUrl(hInet, "http://www.zenwerx.com/version.php?ver_id=" & cur_version, vbNullString, 0, Flags, 0)
        If hUrl Then
            bToRead = 1024 ' 1024 byte chunk. Should be WAY more than necessary for this
            sBuffer = String(bToRead, 0)
            hResult = InternetReadFile(hUrl, sBuffer, bToRead, bRead)
            
            If bRead < bToRead Then
                size = ""
                
                checkVersion = Val(Left(sBuffer, 1))
                
                For i = 3 To Len(sBuffer)
                    If Mid(sBuffer, i, 1) = Chr(0) Then
                        Exit For
                    End If
                    
                    size = size & Mid(sBuffer, i, 1)
                Next i
            End If
         
            modGlobals.filesize = Val(size)
            
            Call InternetCloseHandle(hUrl)
            
            hUrl = InternetOpenUrl(hInet, "http://www.zenwerx.com/version.php", vbNullString, 0, Flags, 0)
            If hUrl Then
                
                bToRead = 1024 ' 1024 byte chunk. Should be WAY more than necessary for this
                sBuffer = String(bToRead, 0)
                hResult = InternetReadFile(hUrl, sBuffer, bToRead, bRead)
                
                If bRead < bToRead Then
                    size = ""
                    
                    For i = 1 To Len(sBuffer)
                        If Mid(sBuffer, i, 1) = Chr(0) Then
                            Exit For
                        End If
                        
                        size = size & Mid(sBuffer, i, 1)
                    Next i
                End If
             
                modGlobals.new_version = size
                Call InternetCloseHandle(hUrl)
            End If
        End If
    Else
        MsgBox lang.GetString(335), vbCritical + vbOKOnly, App.Title
        Exit Function
    End If
        
End Function

Function compareVersions(v1 As String, v2 As String) As Integer

    Dim i As Integer
    Dim c1 As String
    Dim va_one() As String
    Dim va_two() As String
    
    va_one = Split(v1, ".")
    va_two = Split(v2, ".")
    
    For i = 0 To 3
        If Val(va_one(i)) > Val(va_two(i)) Then
            compareVersions = -1
            Exit Function
        ElseIf Val(va_one(i)) < Val(va_two(i)) Then
            compareVersions = 1
            Exit Function
        End If
    Next i
    compareVersions = 0
End Function

' ***************************************************************************
' * Download new version of DB Converter
' ***************************************************************************
Public Function dl_New() As Boolean
    On Error Resume Next
    Dim hInet As Long
    Dim hUrl As Long
    Dim Flags As Long
    Dim url As Variant
    
    ' Open an internet connection
    hInet = InternetOpen(App.Title, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0&)
    If hInet Then
        ' Set our flags, and open the update url
        Flags = INTERNET_FLAG_KEEP_CONNECTION Or INTERNET_FLAG_NO_CACHE_WRITE Or INTERNET_FLAG_RELOAD
        hUrl = InternetOpenUrl(hInet, "http://www.zenwerx.com/DBConverter.msi", vbNullString, 0, Flags, 0)
        If hUrl Then
            ' Show the download form
            frmDownload.Show
            frmDownload.prgDownload.Value = 0
            
            ' local variables
            Dim sBuffer As String
            Dim bToRead As Long
            Dim bRead As Long
            Dim bLeft As Long
            Dim hResult As Boolean
            
            Dim total As Long
            
            ' read in 1024 byte chunks, a nice power of 2. I hope the computer likes it
            bToRead = 1024
            
            ' Make sure our buffer is EXACTLY the size of how many bytes we're reading
            sBuffer = Space(bToRead)
            
            ' If the file has been downloaded in the past, kill it, and make a new one
            Kill App.Path & "\" & "DBConverter.msi"
            Open App.Path & "\" & "DBConverter.msi" For Binary As #1
            Do
                ' Read the file
                hResult = InternetReadFile(hUrl, sBuffer, bToRead, bRead)
                If hResult = True Then
                    ' Update our progress bar
                    frmDownload.prgDownload.Value = (total / filesize) * 100
                    frmDownload.lblSatus.Caption = lang.GetString(336) & String(Int((total / filesize) * 10), ".") & Chr(13) & Int((total / filesize) * 100) & "% (" & Int(total / 1024) & "KB/" & Int(filesize / 1024) & "KB)"
                    ' Write to file
                    Put #1, , sBuffer
                    ' Get our total and leftover
                    total = total + bRead
                    bLeft = filesize - total
                    ' This is unnecessary now that I know how this function works
                    ' better, but we'll leave it here for grandfather purposes
                    ' (basically, I don't want to break it)
                    If bLeft < 1024 Then
                        bToRead = bLeft
                        sBuffer = ""
                        sBuffer = Space(bToRead)
                    End If
                Else
                    ' Show that we got an error on teh download form
                    frmDownload.cmdClose.Enabled = True
                    frmDownload.cmdRun.Enabled = False
                    frmDownload.lblSatus = "Download Error!" & Chr(13) & Int(total / filesize) & "%"
                    DoEvents
                    Exit Do
                End If
                DoEvents
                If frmDownload.Visible = False Then
                    Close #1
                    Call InternetCloseHandle(hUrl)
                    Call InternetCloseHandle(hInet)
                    Kill App.Path & "\" & "DBConverter.msi"
                    MsgBox lang.GetString(337), vbInformation + vbOKOnly, App.Title
                    Unload frmDownload
                    dl_New = False
                    Exit Function
                End If
                    
            Loop Until bLeft <= 0
            Close #1
            ' Clean up
            Call InternetCloseHandle(hUrl)
        Else
            dl_New = False
        End If
    End If
    ' Show that we were successful, and enable the run button
    frmDownload.lblSatus.Caption = lang.GetString(305) & Chr(13) & Int((total / filesize) * 100) & "% (" & Int(total / 1024) & "KB/" & Int(filesize / 1024) & "KB)"
    frmDownload.cmdClose.Enabled = True
    frmDownload.cmdRun.Enabled = True
    ' Clean up
    Call InternetCloseHandle(hInet)
End Function

' ***************************************************************************
' * The IS_SET function is used extensively for the bitvector checking
' * in the options structure. By default, this function returns either 0 or
' * a number greater than 0. To make it compatible with the VB version of
' * true and false, it's been modified to return only a 0 or a 1 to correspond
' * with vbChecked and vbUnchecked. If it was to adhere to true/false, which
' * I think vbChecked and vbUnchecked SHOULD, it would be 0, -1... or work fine
' * with the "C" versions of boolean (0 = FALSE, TRUE = ANYTHING NOT FALSE).
' *
' * This return 0             This return 1 because one
' * because no bits           bit matches
' * match
' *                                *
' * 00000000                  00000100
' * 11111111                  11111111
' * --------                  --------
' * 00000000 <- Result        00000100 <- Result
' *
' ***************************************************************************
Public Function IS_SET(flag As Long, bit As Long) As Integer
    IS_SET = flag And bit
    If IS_SET > 1 Then
        IS_SET = 1
    End If
End Function

' ***************************************************************************
' * Toggle a bit on or off. We don't care what the old value was
' * just make it the opposite
' ***************************************************************************
Public Sub TOGGLE_BIT(flag As Long, bit As Long)
    flag = flag Xor bit
End Sub

' ***************************************************************************
' * These two functions aren't used at the moment, but could possibly
' * be used somewhere down the line
' *
' * SET_BIT will turn a bit on. It doesn't matter if it was already on
' * The or fires if either bit is true, or both
' ***************************************************************************
Public Sub SET_BIT(flag As Long, bit As Long)
    flag = flag Or bit
End Sub

' ***************************************************************************
' * REMOVE_BIT will turn a bit off, even if it was already off
' * (I have my doubts that this one works...
' * I have a feeling it will turn a bit on if it was off
' * but the main problem is that VB doesn't have bitwise operators
' * and this code was translated from C)
' ***************************************************************************
Public Sub REMOVE_BIT(flag As Long, bit As Long)
    flag = flag And (Not bit)
End Sub

' ***************************************************************************
' * Global parsing function
' *
' * It will look for:
' *  semicolons
' *  quotes
' *  backslashes
' ***************************************************************************
Public Function ParseOut(Optional parse As String) As String
    If parse = "" Then
        ParseOut = ""
        Exit Function
    End If

    Dim charpos As Integer
    Dim lstring As String
    Dim rstring As String
    
    charpos = -1
    
    ' I have a feeling the first two could be put together
    ' The third loop most likely has to stay seperate to make sure we don't
    ' add infinite backslashes
    
    ' If we're connecting to MySQL, this one adds the \ before ;'s
    If mainform.chkConnect.Enabled = True Then
        
        Do
            charpos = InStr(charpos + 2, parse, ";")
            If charpos = 0 Then Exit Do
            lstring = Mid(parse, 1, charpos - 1)
            rstring = Right(parse, (Len(parse) - charpos + 1))
            lstring = lstring & "\"
            parse = lstring + rstring
        Loop
        
        charpos = -1
    End If
       
    ' This one adds the \ before "'s
    ' I'm not sure why it's not in the connection "if" above
    Do
        charpos = InStr(charpos + 2, parse, Chr(34))
        If charpos = 0 Then Exit Do
        lstring = Mid(parse, 1, charpos - 1)
        rstring = Right(parse, (Len(parse) - charpos + 1))
        lstring = lstring & "\"
        parse = lstring + rstring
    Loop
    
    charpos = -1
    
    ' This one looks for escaped characters
    ' because MySQL recognizes \;, and \", so we don't want
    ' to add a \ before those one's
    Do
top:
        charpos = InStr(charpos + 2, parse, "\")
        If charpos = 0 Then Exit Do
        Select Case Mid(parse, charpos + 1, 1)
        Case ";", "\", Chr(34)
            ' for the lack of a continue statement in VB, use a goto
            GoTo top
        End Select
        lstring = Mid(parse, 1, charpos - 1)
        rstring = Right(parse, (Len(parse) - charpos + 1))
        lstring = lstring & "\"
        parse = lstring + rstring
    Loop
    
    ParseOut = parse
End Function

' ***************************************************************************
' * Get the list of valid tables in the database, and add them to the list
' ***************************************************************************
Public Sub ChooseTables(FileName As String)
    On Error GoTo err_handle:
    Dim wrkJet As Workspace
    Dim d As Database
    Dim w As Workspace
    Dim i As Integer
    
do_password:
    Set wrkJet = CreateWorkspace("temp", "admin", "", dbUseJet)
    Set d = wrkJet.OpenDatabase(FileName, False, True, ";UID=" & DB_User & ";PWD=" & DB_Pass)
    
    frmTables.lstTables.Clear
    
    For i = 0 To (d.TableDefs.count - 1)
        If InStr(UCase(d.TableDefs(i).Name), "MSYS") = 0 Then
            frmTables.lstTables.AddItem d.TableDefs(i).Name
        End If
    Next i

    frmTables.Show vbModal, mainform
    
    Set wrkJet = Nothing
    Set d = Nothing
    Exit Sub
err_handle:
    If err.Number = 3031 Or err.Number = 3146 Then
        working = False
        frmPass.Show vbModal, mainform
        GoTo do_password:
    Else
        MsgBox err.Number & err.Description, vbCritical, App.Title
    '    Exit Sub
    End If
    Set wrkJet = Nothing
    Set d = Nothing
End Sub

' ***************************************************************************
' * Public countstatements function
' * It's used by all the conversion routines
' * I figured a global sub that really didn't do any damage was better
' * than recreating code in all the classes (where's my inheritance!!!)
' *
' * This function isn't exactly accurate. It's more of a estimate
' * For the most part it works ok though. If anything it overshoots
' * the number of records we need
' ***************************************************************************
Public Sub CountStatements(FileName As String)
    On Error GoTo err_handle:
    Dim wrkJet As Workspace
    Dim d As Database
    Dim w As Workspace
    Dim i As Integer
    Dim j As Integer
    Dim multiply As Integer
    
    currentstate = 0
    totalstatements = 0
    mainform.prgStates.Value = 0
    multiply = 0
    
do_password:
    Set wrkJet = CreateWorkspace("temp", "admin", "", dbUseJet)
    Set d = wrkJet.OpenDatabase(FileName, False, True, ";UID=" & DB_User & ";PWD=" & DB_Pass)
    
    ' Loop on the tables
    For i = 0 To (d.TableDefs.count - 1)
        If InStr(UCase(d.TableDefs(i).Name), "MSYS") = 0 Then
            ' If "Convert All Tables" is checked, just loop through
            ' 2 statements for each table - drop and create
            If IS_SET(Opt.GenOptions, OPT_TABLES) = 1 Then
                totalstatements = totalstatements + 1
                If mainform.chkData.Value = vbChecked Then
                    ' 1 statement for each record
                    totalstatements = totalstatements + d.TableDefs(i).RecordCount
                End If
            ' Otherwise, we actually have to see if it matches a record
            ' in our checked tables
            Else
                ' Loopity loop
                For j = 0 To frmTables.lstTables.ListCount
                    ' Check if you'er on the list
                    If frmTables.lstTables.List(j) = d.TableDefs(i).Name Then
                        If frmTables.lstTables.Selected(j) = True Then
                            totalstatements = totalstatements + 1
                            If mainform.chkData.Value = vbChecked Then
                                ' 1 statement for each record
                                totalstatements = totalstatements + d.TableDefs(i).RecordCount
                            End If
                        End If
                        Exit For
                    End If
                Next j
            End If
        End If
    Next i
    
    ' Figure out a multiplier for the conversions we want
    If mainform.chkPHP.Value = vbChecked Then
        ' Gonna add 2 for mysql because it takes extra statements
        multiply = multiply + 2
    End If
    If mainform.chkOracle.Value = vbChecked Then
        multiply = multiply + 1
    End If
    If mainform.chkPost.Value = vbChecked Then
        multiply = multiply + 1
    End If
    If mainform.chkMySQL.Value = vbChecked Then
        multiply = multiply + 1
    End If
    
    ' Multiply the total statements by our multiplier
    totalstatements = multiply * totalstatements
    ' Clean up
    Set wrkJet = Nothing
    Set d = Nothing
    Exit Sub
err_handle:
    ' Handle password error
    If err.Number = 3031 Or err.Number = 3146 Then
        working = False
        frmPass.Show vbModal, mainform
        GoTo do_password:
        MsgBox lang.GetString(338), vbInformation, App.Title
        goterr = True
        mainform.cmdStart.Caption = lang.GetString(339)
        mainform.cmdStart.Enabled = True
        mainform.prgStates.Value = 100
        mainform.lblPct.Caption = lang.GetString(340)
    Else
        MsgBox err.Number & err.Description, vbCritical, App.Title
    '    Exit Sub
    End If
    ' Clean up
    Set wrkJet = Nothing
    Set d = Nothing
End Sub

' ***************************************************************************
' * MySQL doesn't like some characters, ditch em
' ***************************************************************************
Public Function ParseFieldName(inString As String) As String
    ' Remove incompatible characters (for mysql table names)
    ' and replace them
    Dim i As Integer
    Dim c As String
    Dim tmpStr As String
    
    tmpStr = UCase(inString)
    
    ' surround special cases in reverse quotes (phpMyAdmin Style)
    If tmpStr = "ALL" Or tmpStr = "GROUP" Or tmpStr = "UPDATE" Or tmpStr = "FROM" Or tmpStr = "HAVING" Or tmpStr = "SELECT" Or tmpStr = "ORDER" Or tmpStr = "INTO" Or tmpStr = "WHERE" Or tmpStr = "DECIMAL" Then
        inString = "`" & inString & "`"
    End If
    
    For i = 1 To (Len(inString))
        c = Mid(inString, i, 1)
        If c = "-" Then
            ParseFieldName = ParseFieldName & "_"
        ElseIf c = "#" Then
            ParseFieldName = ParseFieldName & "_num_"
        ElseIf c = "?" Then
            ParseFieldName = ParseFieldName & "_question_"
        ElseIf c = "~" Then
            ParseFieldName = ParseFieldName & "_tilde_"
        ElseIf c = "+" Then
            ParseFieldName = ParseFieldName & "_plus_"
        ElseIf c = "/" Or c = "\" Then
            ParseFieldName = ParseFieldName & "_slash_"
        ElseIf c = ">" Then
            ParseFieldName = ParseFieldName & "_gt_"
        ElseIf c = "<" Then
            ParseFieldName = ParseFieldName & "_lt_"
        ElseIf c = "=" Then
            ParseFieldName = ParseFieldName & "_eq_"
        ElseIf c = "@" Then
            ParseFieldName = ParseFieldName & "_at_"
        ElseIf c = "$" Then
            ParseFieldName = ParseFieldName & "_dollar_"
        ElseIf c = "&" Then
            ParseFieldName = ParseFieldName & "_and_"
        ElseIf c = "*" Then
            ParseFieldName = ParseFieldName & "_star_"
        ElseIf c = " " Then
            ParseFieldName = ParseFieldName & "_"
        ElseIf c <> "[" And c <> "]" And c <> ")" And c <> "{" And c <> "}" And c <> "(" And c <> ":" And c <> "'" And c <> Chr(34) Then
            ParseFieldName = ParseFieldName & c
        End If
    Next i
End Function

Public Sub Main()
    ' Handle the uninstall option
    Dim strCommands As String
    
    ' Grab our registry options
    Set reg = New CRegSettings
    GetOpt
    
    ' Create the language class
    Set lang = New CMultiLingual
    
    strCommands = Command()
    If InStr(1, strCommands, "/uninstall") <> 0 Then
        ShellExecute 0, "open", "MsiExec.exe", "/I{32A247FE-9858-4939-A7E3-4937D896DC14}", "", 0
        End
    ElseIf InStr(1, strCommands, "/convert") <> 0 Then
        Call ParseAndConvert(Right(strCommands, Len(strCommands) - InStr(1, strCommands, "/convert") - 7))
    ElseIf InStr(1, strCommands, "/help") <> 0 Or strCommands <> "" Then
        Dim strout As String
        strout = "DB Converter Command Line Options:" & vbCrLf & vbCrLf
        strout = strout & "/help" & vbTab & "-- This message" & vbCrLf
        strout = strout & "/uninstall" & vbTab & "-- Uninstall DB Converter" & vbCrLf
        strout = strout & "/convert" & vbTab & "-- convert from the command line" & vbCrLf & vbCrLf
        strout = strout & "Conversion options:" & vbCrLf
        strout = strout & "    [convert type]" & vbTab & "-- MySQL/Oracle/PostgreSQL (REQUIRED)" & vbCrLf
        strout = strout & vbTab & vbTab & "-- PHP is not supported on the commandline" & vbCrLf
        strout = strout & "    [access db]" & vbTab & "-- Access Database File (REQUIRED)" & vbCrLf
        strout = strout & "    [sql file]" & vbTab & vbTab & "-- SQL Output File (OPTIONAL, Can be assumed from Access DB )" & vbCrLf
        strout = strout & "    [data]" & vbTab & vbTab & "-- Include Data in Conversion - Yes/No (OPTIONAL, YES by default)" & vbCrLf
        
        MsgBox strout, vbOKOnly + vbInformation, App.Title
        End
    Else
        
        Set mainform = New frmMain
        
        mainform.Show
    End If
End Sub

Public Function ParseAndConvert(Commands As String) As Boolean
    Dim i As Integer
    Dim count As Integer
    Dim words() As String
    Dim outPath As String
    Dim outFile As String
    Dim InFile As String
    Dim includeData As Integer
    Dim convertType As Integer
    
    
    count = 0
    includeData = vbChecked
    
    ParseAndConvert = False ' Assume failure
    
    Commands = Trim(Commands)
    
    ' kill double spacing, we use spaces for seperating
    While InStr(1, Commands, "  ") <> 0
        Commands = Replace(Commands, "  ", " ")
    Wend
    
    For i = 1 To Len(Commands)
        If i = 1 Or Mid(Commands, i, 1) = " " Then
            count = count + 1
            ReDim Preserve words(count)
            If (i = 1) Then
                words(count) = Mid(Commands, i, 1)
            Else
                words(count) = "" ' just to be sure, set it to a blank string
            End If
        Else
            words(count) = words(count) & Mid(Commands, i, 1)
        End If
    Next i
    
    If count < 2 Then
        MsgBox "Not enough options for conversion!", vbCritical + vbOKOnly, App.Title
        Exit Function
    End If
    
    words(1) = UCase(words(1))
    
    If words(1) <> "MYSQL" And words(1) <> "POSTGRESQL" And words(1) <> "ORACLE" Then
        MsgBox "I don't know how to convert to " & Chr(34) & words(1) & Chr(34) & "." & vbCrLf & vbCrLf & "Please use one of the following options:" & vbCrLf & vbTab & "ORACLE" & vbCrLf & vbTab & "POSTGRESQL" & vbCrLf & vbTab & "MYSQL", vbCritical + vbOKOnly, App.Title
        Exit Function
    Else
        If words(1) = "ORACLE" Then
            convertType = 1
        ElseIf words(1) = "POSTGRESQL" Then
            convertType = 2
        ElseIf words(1) = "MYSQL" Then
            convertType = 3
        End If
    End If
    
    words(2) = UCase(words(2))
    
    If InStr(1, words(2), ".MDB") = 0 Then
        words(2) = words(2) & ".MDB"
    End If
    
    If Dir(words(2)) = "" Then
        MsgBox "Can't find " & Chr(34) & words(2) & Chr(34), vbCritical + vbOKOnly, App.Title
        Exit Function
    Else
        InFile = words(2)
    End If
    
    If count > 2 Then
        outPath = words(3)
        
        ' Find the backslash
        For i = Len(outPath) To 1 Step -1
            If (Mid(outPath, i, 1) = "\") Then
                outPath = Left(outPath, i)
                Exit For
            End If
        Next i
        
        ' Check the folder exists. If it doesn't, break out
        If Dir(outPath) = "" Then
            MsgBox "The path " & Chr(34) & outPath & Chr(34) & " does not exist. It cannot be output to.", vbCritical + vbOKOnly, App.Title
            Exit Function
        End If
        
        ' Strip off the sql extension, if it exists
        If (InStr(1, words(3), ".sql")) <> 0 Then
            words(3) = Left(words(3), InStr(1, words(3), ".sql") - 1)
        End If
        
        outFile = Right(words(3), Len(words(3)) - i)
        
        If count > 3 Then
            words(4) = UCase(words(4))
            If words(4) <> "YES" And words(4) <> "NO" Then
                MsgBox "Data options are YES or NO.", vbCritical + vbOKOnly, App.Title
                Exit Function
            Else
                If words(4) = "YES" Then
                    includeData = vbChecked
                Else
                    includeData = vbUnchecked
                End If
            End If
        Else
            includeData = vbChecked
        End If
    Else
        ' No output file given
        ' Assume it = Access FileName + Path
        outPath = words(2)
        
        For i = Len(outPath) To 1 Step -1
            If (Mid(outPath, i, 1) = "\") Then
                outPath = Left(outPath, i)
                Exit For
            End If
        Next i
                
        outFile = Right(words(2), Len(words(2)) - i)
        outFile = Left(outFile, InStr(1, outFile, ".mdb") - 1)
        
        ' Set the data option
        includeData = vbChecked
    End If
    
    If convertType = 1 Then ' oracle
        Dim coracle As New COracle_Convert
        Call coracle.MakeOracle(InFile, outPath, outFile, includeData, True)
        Set coracle = Nothing
    ElseIf convertType = 2 Then ' postgresql
        Dim cpost As New CPostgreSQL_Convert
        Call cpost.MakePostgreSQL(InFile, outPath, outFile, includeData, True)
        Set cpost = Nothing
    ElseIf convertType = 3 Then ' mysql
        Dim cmysql As New CMySQL_Convert
        Call cmysql.MakeMySQL(InFile, outPath, outFile, includeData, True)
        Set cmysql = Nothing
    End If
    
    ParseAndConvert = True
End Function

Public Sub LoadMenuImages()
    Dim hMenu As Long
    Dim hSubMenu As Long
    Dim sbuff As String
    Dim ret As Long
    
    ReDim Pictures(MENU_PIC_COUNT)
        
    hMenu = modGlobals.GetMenu(mainform.hwnd)
    
    ' Help Submenu
    hSubMenu = modGlobals.GetSubMenu(hMenu, 3)
    
    ' Dollar sign. Donate
    Set Pictures(0) = GetTempMenuGIF(ID_DONATE)
    ret = SetMenuItemBitmaps(hSubMenu, 4, MF_BYPOSITION, Pictures(0), Pictures(0))
    ' Life Preserver. Help
    Set Pictures(1) = GetTempMenuGIF(ID_HELP)
    ret = SetMenuItemBitmaps(hSubMenu, 0, MF_BYPOSITION, Pictures(1), Pictures(1))
    
    ' Edit Submenu
    hSubMenu = modGlobals.GetSubMenu(hMenu, 1)
    'Cut
    Set Pictures(2) = GetTempMenuGIF(ID_CUT)
    ret = SetMenuItemBitmaps(hSubMenu, 0, MF_BYPOSITION, Pictures(2), Pictures(2))
    'Copy
    Set Pictures(3) = GetTempMenuGIF(ID_COPY)
    ret = SetMenuItemBitmaps(hSubMenu, 1, MF_BYPOSITION, Pictures(3), Pictures(3))
    'Paste
    Set Pictures(4) = GetTempMenuGIF(ID_PASTE)
    ret = SetMenuItemBitmaps(hSubMenu, 2, MF_BYPOSITION, Pictures(4), Pictures(4))
    'Options
    Set Pictures(5) = GetTempMenuGIF(ID_OPTIONS)
    ret = SetMenuItemBitmaps(hSubMenu, 4, MF_BYPOSITION, Pictures(5), Pictures(5))
    
    ' File Submenu
    hSubMenu = modGlobals.GetSubMenu(hMenu, 0)
    
    ' Quit
    Set Pictures(6) = GetTempMenuGIF(ID_QUIT)
    ret = SetMenuItemBitmaps(hSubMenu, 0, MF_BYPOSITION, Pictures(6), Pictures(6))
    
    
                    
End Sub

Private Function GetTempMenuGIF(ID As Integer) As StdPicture
    Dim data() As Byte
    data() = LoadResData(ID, "MENU")
    
    Open App.Path & "\temp.gif" For Binary As #1
    Put #1, , data
    Close #1
    
    Set GetTempMenuGIF = LoadPicture(App.Path & "\temp.gif")
    
    Kill App.Path & "\temp.gif"
    
End Function
