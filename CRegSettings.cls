VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRegSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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

' Registry API Calls

' Open/close
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
' Create/Delete keys
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
' Get value
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
' Set/Delete values
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

' Reg Type Needed
Private Const REG_OPTION_NON_VOLATILE = 0       ' Key is preserved when system is rebooted

' Constants for API Access
Private Const HKEY_LOCAL_MACHINE = &H80000002




' Reg Key Security Options
Private Const DELETE = &H10000
Private Const READ_CONTROL = &H20000
Private Const WRITE_DAC = &H40000
Private Const WRITE_OWNER = &H80000
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const SPECIFIC_RIGHTS_ALL = &HFFFF
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))

Private Const ERROR_SUCCESS = 0&

Private Const REG_SZ = 1                         ' Unicode nul terminated string
Private Const REG_DWORD = 4

Public Function GetSetting(ByVal Section As String, ByVal Key As String, Optional ByVal Default As String = "") As String
   Dim nRet As Long
   Dim hKey As Long
   Dim nType As Long
   Dim nBytes As Long
   Dim Buffer As String
   
   ' Assume failure
   GetSetting = Default

   ' Open key
   nRet = RegOpenKeyEx(HKEY_LOCAL_MACHINE, Section, 0&, KEY_ALL_ACCESS, hKey)
   If nRet = ERROR_SUCCESS Then
      ' Set appropriate value for default query
      If Key = "*" Then Key = vbNullString
      
      ' Get buffer size
      nRet = RegQueryValueEx(hKey, Key, 0&, nType, ByVal Buffer, nBytes)
      If nRet = ERROR_SUCCESS Then
         ' Build buffer and get data
         If nBytes > 0 Then
            Buffer = Space(nBytes)
            nRet = RegQueryValueEx(hKey, Key, 0&, nType, ByVal Buffer, Len(Buffer))
            If nRet = ERROR_SUCCESS Then
               ' Trim NULL and return successful query!
               GetSetting = Left(Buffer, nBytes - 1)
            End If
         End If
      Call RegCloseKey(hKey)
      End If
   End If
End Function

Public Function SaveSetting(ByVal Section As String, ByVal Key As String, ByVal Setting As Variant) As Boolean
   Dim nRet As Long
   Dim hKey As Long
   Dim nResult As Long
   
   ' Open (or create and open) key
   nRet = RegCreateKeyEx(HKEY_LOCAL_MACHINE, Section, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, hKey, nResult)
   If nRet = ERROR_SUCCESS Then
      ' Set appropriate value for default query
      If Key = "*" Then Key = vbNullString
      ' Null-terminate setting, in case it's empty.
      If TypeName(Setting) = "Integer" Then
        Dim x As Integer
        x = Val(Setting)
        nRet = RegSetValueEx(hKey, Key, 0&, REG_DWORD, x, 4)
      Else
        Setting = Setting & vbNullChar
        Dim temp As String
        temp = Setting
        nRet = RegSetValueEx(hKey, Key, 0&, REG_SZ, ByVal temp, Len(temp))
      End If
      Call RegCloseKey(hKey)
   End If
   SaveSetting = (nRet = ERROR_SUCCESS)
End Function

' Not really used in this program, but I port this class around to different
' applications
Public Function DeleteSetting(ByVal Section As String, Optional ByVal Key As String = "") As Boolean
   Dim nRet As Long
   Dim hKey As Long

   If Len(Key) Then
      ' Open key
      nRet = RegOpenKeyEx(HKEY_LOCAL_MACHINE, Section, 0&, KEY_ALL_ACCESS, hKey)
      If nRet = ERROR_SUCCESS Then
         ' Set appropriate value for default query
         If Key = "*" Then Key = vbNullString
         ' Delete the requested value
         nRet = RegDeleteValue(hKey, Key)
         Call RegCloseKey(hKey)
      End If
   Else
      nRet = RegOpenKeyEx(HKEY_LOCAL_MACHINE, Section, 0&, KEY_ALL_ACCESS, hKey)
      If nRet = ERROR_SUCCESS Then
         ' Attempt to delete whole section
         nRet = RegDeleteKey(hKey, Section)
         Call RegCloseKey(hKey)
      End If
   End If
   DeleteSetting = (nRet = ERROR_SUCCESS)
End Function


