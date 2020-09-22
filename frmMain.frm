VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "NetSender"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9495
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   243
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   633
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkStartInSysTray 
      Caption         =   "Start in system tray"
      Height          =   195
      Left            =   6240
      TabIndex        =   21
      Top             =   1560
      Width           =   1815
   End
   Begin VB.PictureBox picTray 
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   6120
      Picture         =   "frmMain.frx":08CA
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   20
      Top             =   2640
      Width           =   480
   End
   Begin VB.CheckBox chkSendToGroup 
      Caption         =   "Send to groupe:"
      Height          =   195
      Left            =   3480
      TabIndex        =   19
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdMessage 
      Caption         =   ">"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3120
      Width           =   255
   End
   Begin VB.ListBox lstMessage 
      Height          =   2205
      ItemData        =   "frmMain.frx":1194
      Left            =   360
      List            =   "frmMain.frx":1196
      TabIndex        =   17
      Top             =   960
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.TextBox txtInterval 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6960
      TabIndex        =   15
      Text            =   "0"
      Top             =   1155
      Width           =   615
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   7800
      TabIndex        =   12
      Top             =   3165
      Width           =   1575
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "Hide"
      Height          =   375
      Left            =   6120
      TabIndex        =   11
      Top             =   3165
      Width           =   1575
   End
   Begin VB.TextBox txtMessageCount 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7320
      TabIndex        =   9
      Text            =   "1"
      Top             =   795
      Width           =   495
   End
   Begin VB.ListBox lstTarget 
      Enabled         =   0   'False
      Height          =   3180
      IntegralHeight  =   0   'False
      Left            =   3480
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   360
      Width           =   2535
   End
   Begin VB.OptionButton optPCNameRandom 
      Caption         =   "Generate random name"
      Height          =   195
      Left            =   6240
      TabIndex        =   8
      Top             =   480
      Width           =   2055
   End
   Begin VB.OptionButton optPCNameUser 
      Height          =   195
      Left            =   6240
      TabIndex        =   7
      Top             =   120
      Value           =   -1  'True
      Width           =   255
   End
   Begin VB.TextBox txtPCName 
      Height          =   285
      Left            =   6600
      TabIndex        =   6
      Text            =   "Anonymous"
      Top             =   75
      Width           =   2775
   End
   Begin VB.ComboBox cboTarget 
      Height          =   315
      Left            =   840
      TabIndex        =   4
      Top             =   60
      Width           =   2535
   End
   Begin VB.TextBox txtMessage 
      Height          =   2535
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   600
      Width           =   3255
   End
   Begin ComctlLib.ProgressBar pbarMessageSend 
      Height          =   255
      Left            =   2100
      TabIndex        =   25
      Top             =   3225
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send Message"
      Height          =   371
      Left            =   2040
      TabIndex        =   3
      Top             =   3165
      Width           =   1331
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SKabir@HotPOP.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   315
      Left            =   6675
      MouseIcon       =   "frmMain.frx":1198
      MousePointer    =   99  'Custom
      TabIndex        =   24
      Top             =   2730
      Width           =   2685
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SKabir@HotPOP.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   6660
      TabIndex        =   23
      Top             =   2715
      Width           =   2685
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SKabir@HotPOP.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   6690
      TabIndex        =   22
      Top             =   2745
      Width           =   2685
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "sec(s)"
      Height          =   195
      Left            =   7680
      TabIndex        =   16
      Top             =   1200
      Width           =   435
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Interval"
      Height          =   195
      Left            =   6240
      TabIndex        =   14
      Top             =   1200
      Width           =   570
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Send message"
      Height          =   195
      Left            =   6240
      TabIndex        =   13
      Top             =   840
      Width           =   1035
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Time(s)"
      Height          =   195
      Left            =   7920
      TabIndex        =   10
      Top             =   840
      Width           =   525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Message"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Terget"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   480
   End
   Begin VB.Menu mnuLstTargetRightClick 
      Caption         =   "mnuLstTargetRightClick"
      Visible         =   0   'False
      Begin VB.Menu mnuLstTargetRightClickRemove 
         Caption         =   "Remove"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private NetRoot As NetResource

'API Declarations
Private Declare Function NetMessageBufferSend Lib "netapi32.dll" (ByVal ServerName As String, ByVal msgname As String, ByVal FromName As String, ByVal Buffer As String, ByVal BufSize As Long) As Long

Private Const NERR_SUCCESS As Long = 0
Private Const NERR_BASE As Long = 2100
Private Const NERR_NetworkError As Long = (NERR_BASE + 36)
Private Const NERR_NameNotFound As Long = (NERR_BASE + 173)
Private Const NERR_UseNotFound As Long = (NERR_BASE + 150)
Private Const ERROR_ACCESS_DENIED As Long = 5
Private Const ERROR_BAD_NETPATH As Long = 53
Private Const ERROR_NOT_SUPPORTED As Long = 50
Private Const ERROR_INVALID_PARAMETER As Long = 87
Private Const ERROR_INVALID_NAME As Long = 123

Function NetSendMessage(sSendTo As String, sMessage As String, Optional ServerName As String = vbNullString, Optional FromName As String = vbNullString) As Long
    'convert ANSI strings to UNICODE and send the message
    NetSendMessage = NetMessageBufferSend(StrConv(ServerName, vbUnicode), StrConv(sSendTo, vbUnicode), StrConv(FromName & vbNullString, vbUnicode), StrConv(sMessage, vbUnicode), Len(StrConv(sMessage, vbUnicode)))
End Function

'returns the description of the Netapi Error Code
Function NetSendErrorMessage(ErrNum As Long) As String
    Select Case ErrNum
        Case NERR_SUCCESS
            NetSendErrorMessage = "The message was successfully sent"
        Case NERR_NameNotFound
            NetSendErrorMessage = "Send To not found"
        Case NERR_NetworkError
            NetSendErrorMessage = "General network error occurred"
        Case NERR_UseNotFound
            NetSendErrorMessage = "Network connection not found"
        Case ERROR_ACCESS_DENIED
            NetSendErrorMessage = "Access to computer denied"
        Case ERROR_BAD_NETPATH
            NetSendErrorMessage = "Sent From server name not found."
        Case ERROR_INVALID_PARAMETER
            NetSendErrorMessage = "Invalid parameter(s) specified."
        Case ERROR_NOT_SUPPORTED
            NetSendErrorMessage = "Network request not supported."
        Case ERROR_INVALID_NAME
            NetSendErrorMessage = "Illegal character or malformed name."
        Case Else
            NetSendErrorMessage = "Unknown error executing command."
   End Select
End Function

Private Sub chkSendToGroup_Click()
lstTarget.Enabled = chkSendToGroup.Value
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdHide_Click()
CreateIcon
Me.Hide
End Sub

Private Sub cmdMessage_Click()
lstMessage.Visible = True
lstMessage.SetFocus
End Sub

Private Sub cmdSend_Click()
    Dim ret As Long, C As Long, FromName As String, D As Long, SuppressError As Boolean
    
    If optPCNameUser Then FromName = txtPCName Else FromName = GeneratePCName
    
    pbarMessageSend.Visible = True
    pbarMessageSend.Max = Val(txtMessageCount)
    For C = 1 To Val(txtMessageCount)
        If chkSendToGroup.Value = vbChecked Then
            If lstTarget.ListCount = 0 Then
                MsgBox "Target group empty!", vbCritical, "Trouble!"
                pbarMessageSend.Visible = False
                Exit Sub
            End If
            For D = 0 To lstTarget.ListCount - 1
                If lstTarget.Selected(D) = True Then
                    ret = NetSendMessage(lstTarget.List(D), txtMessage, vbNullString, FromName)
                    If ret <> 0 Then
                        If Not SuppressError Then
                            If MsgBox(NetSendErrorMessage(ret) & vbCrLf & vbCrLf & "Do you want to suppress further error messages?", vbCritical + vbYesNo, "Error!") = vbYes Then SuppressError = True Else SuppressError = False
                        End If
                    End If
                End If
            Next
        Else
            If cboTarget.Text = "" Then
                MsgBox "No target specified!", vbCritical, "Trouble!"
                pbarMessageSend.Visible = False
                Exit Sub
            End If
            ret = NetSendMessage(cboTarget.Text, txtMessage, vbNullString, FromName)
            If ret <> 0 Then
                If Not SuppressError Then
                    If MsgBox(NetSendErrorMessage(ret) & vbCrLf & vbCrLf & "Do you want to suppress further error messages?", vbCritical + vbYesNo, "Error!") = vbYes Then SuppressError = True Else SuppressError = False
                End If
            End If
        End If
        pbarMessageSend.Value = C
        
        If Val(txtInterval) > 0 Then
            Dim TimeElapsed As Date
            TimeElapsed = Now
            Do While DateDiff("s", TimeElapsed, Now) < Val(txtInterval)
                DoEvents
            Loop
        End If
    Next
    pbarMessageSend.Visible = False
    
    If Not IsItemInList(cboTarget.Text, cboTarget) Then
        cboTarget.AddItem cboTarget.Text
        lstTarget.AddItem cboTarget.Text
        lstTarget.Selected(lstTarget.ListCount - 1) = True
    End If
    
    If lstMessage.ListCount > 0 And lstMessage.ListCount = 20 Then
        For C = 1 To lstMessage.ListCount - 1
            lstMessage.List(C - 1) = lstMessage.List(C)
        Next
    End If
    If lstMessage.ListCount = 20 Then lstMessage.List(20) = txtMessage Else lstMessage.AddItem txtMessage
End Sub

Private Sub Form_Load()
LoadTarget
End Sub

Private Sub SaveTarget()
Dim clsINI As clsINI, a As Long
Set clsINI = New clsINI

clsINI.SetKeyString App.Path & "\NetSender.INI", "Setting", "Start in system tray", chkStartInSysTray.Value

clsINI.SetKeyString App.Path & "\NetSender.INI", "Target", "Total", lstTarget.ListCount
If lstTarget.ListCount < 1 Then Exit Sub
For a = 0 To lstTarget.ListCount - 1
    clsINI.SetKeyString App.Path & "\NetSender.INI", "Target", "Target (" & a + 1 & ")", lstTarget.List(a) & "," & lstTarget.Selected(a)
Next

clsINI.SetKeyString App.Path & "\NetSender.INI", "Message", "Total", lstMessage.ListCount
If lstMessage.ListCount < 1 Then Exit Sub
For a = 0 To lstMessage.ListCount - 1
    clsINI.SetKeyString App.Path & "\NetSender.INI", "Message", "Message (" & a + 1 & ")", lstMessage.List(a)
Next

End Sub

Private Sub LoadTarget()
Dim Total As Long, a As Long, clsINI As clsINI, Target As String, TotalMessage As Long
Set clsINI = New clsINI

chkStartInSysTray.Value = clsINI.GetKeyInt(App.Path & "\NetSender.INI", "Setting", "Start in system tray")

Total = clsINI.GetKeyInt(App.Path & "\NetSender.INI", "Target", "Total")
If Total < 1 Then Exit Sub
For a = 1 To Total
    Target = clsINI.GetKeyStr(App.Path & "\NetSender.INI", "Target", "Target (" & a & ")")
    lstTarget.AddItem Left(Target, InStr(Target, ",") - 1)
    cboTarget.AddItem Left(Target, InStr(Target, ",") - 1)
    If Mid(Target, InStr(Target, ",") + 1) = True Then lstTarget.Selected(a - 1) = True
Next

Total = clsINI.GetKeyInt(App.Path & "\NetSender.INI", "Message", "Total")
If TotalMessage < 1 Then Exit Sub
For a = 1 To TotalMessage
    lstTarget.AddItem clsINI.GetKeyStr(App.Path & "\NetSender.INI", "Message", "Message (" & a & ")")
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveTarget

DeleteIcon

ShrinkForm Me
End Sub

Private Sub Label9_Click()
ShellExecute hwnd, "open", "mailto:SKabir@HotPOP.com?Subject=About NetSender", vbNullString, vbNullString, SW_SHOW
End Sub

Private Sub lstMessage_Click()
txtMessage.Text = lstMessage.List(lstMessage.ListIndex)
End Sub

Private Sub lstMessage_LostFocus()
lstMessage.Visible = False
End Sub

Private Sub lstMessage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then lstMessage.Visible = False
End Sub

Private Sub lstTarget_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 And lstTarget.ListIndex > -1 And lstTarget.ListCount > 0 Then PopupMenu mnuLstTargetRightClick
End Sub

Private Sub mnuLstTargetRightClickRemove_Click()
lstTarget.RemoveItem lstTarget.ListIndex
End Sub

Private Sub optPCNameRandom_Click()
optPCNameUser_Click
End Sub

Private Sub optPCNameUser_Click()
txtPCName.Enabled = optPCNameUser.Value
End Sub

Private Function GeneratePCName() As String
Dim a As Long, NameLength As Long, PCName As String

Randomize Timer
NameLength = Int(Rnd * 15)

For a = 1 To NameLength
    PCName = PCName & Chr(64 + Rnd * 52)
Next

GeneratePCName = PCName
End Function

Private Function IsItemInList(ItemToCheck As String, ListObj As Object) As Boolean
Dim C As Long, Found As Boolean

If ListObj.ListCount = 0 Then
    IsItemInList = False
    Exit Function
End If

For C = 0 To ListObj.ListCount - 1
    If ItemToCheck = ListObj.List(C) Then Found = True
    If Found Then Exit For
Next

IsItemInList = Found
End Function

Private Sub picTray_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.Visible Then Exit Sub
    Select Case X '/ Screen.TwipsPerPixelX
    Case Is = WM_LBUTTONDOWN
        Me.Visible = True
    Case Is = WM_RBUTTONDOWN
        'Add the code for the left mouse click on the tray icon
        '
    End Select
End Sub


