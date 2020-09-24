Attribute VB_Name = "tray"
Option Explicit


Const NIF_MESSAGE    As Long = &H1
Const NIF_ICON       As Long = &H2
Const NIF_TIP        As Long = &H4
Const NIM_ADD        As Long = &H0
Const NIM_MODIFY     As Long = &H1
Const NIM_DELETE     As Long = &H2

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public Enum TrayRetunEventEnum
    MouseMove = &H200
    LeftUp = &H202
    LeftDown = &H201
    LeftDbClick = &H203
    RightUp = &H205
    RightDown = &H204
    RightDbClick = &H206
    MiddleUp = &H208
    MiddleDown = &H207
    MiddleDbClick = &H209
End Enum

Public Enum ModifyItemEnum
    ToolTip = 1
    Icon = 2
End Enum


Private TrayIcon As NOTIFYICONDATA
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean


Public Sub TrayAdd(hwnd As Long, Icon As Picture, _
                    ToolTip As String, ReturnCallEvent As TrayRetunEventEnum)
    With TrayIcon
        .cbSize = Len(TrayIcon)
        .hwnd = hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = ReturnCallEvent
        .hIcon = Icon
        .szTip = ToolTip & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, TrayIcon
End Sub

Public Sub TrayDelete()
    Shell_NotifyIcon NIM_DELETE, TrayIcon
End Sub

Public Sub TrayModify(Item As ModifyItemEnum, vNewValue As Variant)
    Select Case Item
        Case ToolTip
            TrayIcon.szTip = vNewValue & vbNullChar
        Case Icon
            TrayIcon.hIcon = vNewValue
    End Select
    Shell_NotifyIcon NIM_MODIFY, TrayIcon
End Sub


