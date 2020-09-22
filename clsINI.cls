VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String)
Private Declare Function GetPrivateProfileInt& Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String)
Private Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String)

Public Function GetKeyStr(strFile As String, strSection As String, strKey As String)
    Dim strInfo As String
    Dim lngRet As Long
    
    strInfo = String(260, " ")
    
    GetPrivateProfileString strSection, strKey, "", strInfo, 260, strFile
    
    GetKeyStr = Trim(strInfo)
End Function


Public Function GetKeyInt(strFile As String, strSection As String, strKey As String)
    GetKeyInt = GetPrivateProfileInt(strSection, strKey, 0, strFile)
End Function


Public Sub SetKeyString(strFile As String, strSection As String, strKey As String, strData As String)
    WritePrivateProfileString strSection, strKey, strData, strFile
End Sub


Public Sub DeleteSection(strFile As String, strSection As String)
    WritePrivateProfileString strSection, vbNullString, vbNullString, strFile
End Sub


Public Sub DeleteKey(strFile As String, strSection As String, strKey As String)
    WritePrivateProfileString strSection, strKey, vbNullString, strFile
End Sub

