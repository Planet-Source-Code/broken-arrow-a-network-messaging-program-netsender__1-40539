Attribute VB_Name = "modMain"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOW = 5

Sub GrowForm(frmObj As Form, Optional FramesPerSec As Long = 600)
Dim vWidth As Long, vHeight As Long, Count As Long

Load frmObj

vWidth = frmObj.Width
vHeight = frmObj.Height

With frmObj
    frmObj.Move -1 * frmObj.Width, -1 * frmObj.Height
    frmObj.Show
    
    For Count = 1 To vWidth Step FramesPerSec
        frmObj.Move (Screen.Width - Count) / 2, (Screen.Height - (vHeight * Count / vWidth)) / 2, Count, vHeight * Count / vWidth
        frmObj.Refresh
        DoEvents
    Next
    
    frmObj.Move (Screen.Width - vWidth) / 2, (Screen.Height - vHeight) / 2, vWidth, vHeight
End With
End Sub

Sub ShrinkForm(frmObj As Form, Optional FramesPerSec As Long = 600, Optional UnloadForm As Boolean = False)
Dim vWidth As Long, vHeight As Long, Count As Long

vWidth = frmObj.Width
vHeight = frmObj.Height

With frmObj
    For Count = vWidth To 1 Step -1 * FramesPerSec
        frmObj.Move (Screen.Width - Count) / 2, (Screen.Height - (vHeight * Count / vWidth)) / 2, Count, vHeight * Count / vWidth
        frmObj.Refresh
        DoEvents
    Next
    
    frmObj.Hide
    If UnloadForm Then Unload frmObj
End With
End Sub

Sub Main()
GrowForm frmMain

If frmMain.chkStartInSysTray.Value = vbChecked Then
    CreateIcon
    frmMain.Hide
End If
End Sub
