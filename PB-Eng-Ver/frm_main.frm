VERSION 5.00
Begin VB.Form frm_main 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7950
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer ilang 
      Interval        =   500
      Left            =   1080
      Top             =   240
   End
   Begin VB.PictureBox Picture1 
      Enabled         =   0   'False
      Height          =   615
      Left            =   360
      Picture         =   "frm_main.frx":4D4A
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   24
      Top             =   9240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer kill_task 
      Interval        =   100
      Left            =   360
      Top             =   240
   End
   Begin VB.Frame Frame_Menu 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Main menu"
      Height          =   4335
      Left            =   5280
      TabIndex        =   7
      Top             =   960
      Width           =   2535
      Begin Porn_Blocker.XpButton cmd_atur 
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " Block base website address "
         ForeColor       =   -2147483630
         ForeHover       =   0
      End
      Begin Porn_Blocker.XpButton cmd_admin 
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Admin"
         ForeColor       =   -2147483630
         ForeHover       =   0
      End
      Begin Porn_Blocker.XpButton cmd_exit 
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   3600
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Exit"
         ForeColor       =   -2147483630
         ForeHover       =   0
      End
      Begin Porn_Blocker.XpButton cmd_about 
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   3000
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "About"
         ForeColor       =   -2147483630
         ForeHover       =   0
      End
      Begin Porn_Blocker.XpButton cmd_hide 
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   2400
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Hide Applicaion"
         ForeColor       =   -2147483630
         ForeHover       =   0
      End
      Begin Porn_Blocker.XpButton cmd_aturcap 
         Height          =   615
         Left            =   120
         TabIndex        =   37
         Top             =   1080
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " Block base Caption"
         ForeColor       =   -2147483630
         ForeHover       =   0
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   " Admin Menu"
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   120
      TabIndex        =   17
      Top             =   960
      Visible         =   0   'False
      Width           =   4935
      Begin VB.TextBox txtcryp 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   2400
         TabIndex        =   38
         Top             =   5520
         Visible         =   0   'False
         Width           =   2415
      End
      Begin Porn_Blocker.XpButton cmd_save_seting 
         Height          =   495
         Left            =   720
         TabIndex        =   25
         Top             =   3840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Save "
         ForeColor       =   -2147483630
         ForeHover       =   0
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Left            =   240
         TabIndex        =   21
         Top             =   1320
         Width           =   4455
         Begin VB.CheckBox Check2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Run on Start-UP"
            Height          =   375
            Left            =   120
            TabIndex        =   23
            Top             =   1200
            Value           =   1  'Checked
            Width           =   3855
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Disable Task Manager, Command Prompt, and TaskKill during blocker active"
            Height          =   855
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Value           =   1  'Checked
            Width           =   4215
         End
      End
      Begin Porn_Blocker.XpButton cmd_set 
         Height          =   375
         Left            =   3600
         TabIndex        =   20
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Set"
         ForeColor       =   -2147483630
         ForeHover       =   0
      End
      Begin VB.TextBox txtpwd 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   3135
      End
      Begin Porn_Blocker.XpButton cmd_default 
         Height          =   495
         Left            =   2520
         TabIndex        =   26
         Top             =   3840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Default"
         ForeColor       =   -2147483630
         ForeHover       =   0
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Set Admin Password"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Block base website address "
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   4935
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   2670
         Left            =   240
         TabIndex        =   16
         Top             =   1680
         Width           =   4335
      End
      Begin Porn_Blocker.XpButton cmd_hapus 
         Height          =   615
         Left            =   2520
         TabIndex        =   5
         Top             =   5040
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Delete"
         ForeColor       =   -2147483630
         ForeHover       =   0
      End
      Begin Porn_Blocker.XpButton cmd_add 
         Height          =   375
         Left            =   3480
         TabIndex        =   2
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Add"
         ForeColor       =   -2147483630
         ForeHover       =   0
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   3015
      End
      Begin Porn_Blocker.XpButton cmd_refresh 
         Height          =   615
         Left            =   600
         TabIndex        =   6
         Top             =   5040
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Refresh"
         ForeColor       =   -2147483630
         ForeHover       =   0
      End
      Begin VB.Label lbl_jml 
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   15
         Top             =   4560
         Width           =   615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Website Blocked :"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   4560
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Add Website "
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "List of Blocked Websites"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   4095
      End
   End
   Begin VB.Frame Framecap 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "   Block Base Caption"
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   120
      TabIndex        =   27
      Top             =   960
      Width           =   4935
      Begin VB.ListBox lst_cap 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   2670
         Left            =   240
         TabIndex        =   28
         Top             =   1680
         Width           =   4335
      End
      Begin VB.TextBox txt_blokcap 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   240
         TabIndex        =   31
         Top             =   720
         Width           =   3015
      End
      Begin Porn_Blocker.XpButton cmd_delcap 
         Height          =   615
         Left            =   2520
         TabIndex        =   29
         Top             =   5040
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Delete"
         ForeColor       =   -2147483630
         ForeHover       =   0
      End
      Begin Porn_Blocker.XpButton cmd_addcap 
         Height          =   375
         Left            =   3480
         TabIndex        =   30
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Add"
         ForeColor       =   -2147483630
         ForeHover       =   0
      End
      Begin Porn_Blocker.XpButton cmd_refreshcap 
         Height          =   615
         Left            =   600
         TabIndex        =   32
         Top             =   5040
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Refresh"
         ForeColor       =   -2147483630
         ForeHover       =   0
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "List of blocked caption"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Add Caption"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Total blocked caption :"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   4560
         Width           =   1935
      End
      Begin VB.Label lbl_jmlcap 
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   33
         Top             =   4560
         Width           =   615
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (C) YaDoY SoFtWaRe DeVeLoPmEnT, 2007. All right reserved."
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   7560
      Width           =   5655
   End
   Begin VB.Image Image2 
      Height          =   1575
      Left            =   5760
      Picture         =   "frm_main.frx":5F0C
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   885
      Left            =   3000
      Picture         =   "frm_main.frx":B10EE
      Top             =   120
      Width           =   4755
   End
   Begin VB.Menu mnu 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu show 
         Caption         =   "Show"
      End
      Begin VB.Menu about 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tujuan As String

Private Sub about_Click()
frm_about.show
End Sub

Private Sub cmd_about_Click()
Load frm_about
frm_about.show

End Sub

Private Sub cmd_add_Click()
Dim cari As Long
If Text1.Text = "" Then
MsgBox "Please enter website address", vbInformation + vbOKOnly, "The P * * N Blocker"
Exit Sub
End If

For cari = 0 To List1.ListCount - 1

If Text1.Text = List1.list(cari) Then
MsgBox "Web site address is present in the list", vbInformation + vbOKOnly, "The P * * N Blocker"
Exit Sub
End If
Text1.SetFocus
Next


List1.AddItem Text1.Text
Text1.Text = ""
SaveFileHost List1, GetSystemPath & "\drivers\etc\Hosts"
lbl_jml.Caption = List1.ListCount

End Sub

Private Sub cmd_addcap_Click()
Dim cari As Long
If txt_blokcap.Text = "" Then
MsgBox "Please enter caption", vbInformation + vbOKOnly, "The P * * N Blocker"
Exit Sub
End If

For cari = 0 To lst_cap.ListCount - 1
If txt_blokcap.Text = lst_cap.list(cari) Then
MsgBox "Caption present in the list", vbInformation + vbOKOnly, "The P * * N Blocker"
txt_blokcap.SetFocus
Exit Sub
End If
Next

lst_cap.AddItem txt_blokcap.Text
txt_blokcap.Text = ""
SaveCaption lst_cap, App.Path & "\list.txt"
lbl_jmlcap.Caption = lst_cap.ListCount


End Sub

Private Sub cmd_admin_Click()
Frame1.Visible = False
Frame1.Enabled = False
Frame2.Enabled = True
Frame2.Visible = True
Framecap.Visible = False
Framecap.Enabled = False

End Sub


Private Sub cmd_atur_Click()
Frame1.Visible = True
Frame1.Enabled = True
Frame2.Enabled = False
Frame2.Visible = False
Framecap.Visible = False
Framecap.Enabled = False
lbl_jml.Caption = List1.ListCount

End Sub

Private Sub cmd_aturcap_Click()
Frame1.Visible = False
Frame1.Enabled = False
Frame2.Enabled = False
Frame2.Visible = False
Framecap.Visible = True
Framecap.Enabled = True
lbl_jmlcap.Caption = lst_cap.ListCount

End Sub

Private Sub cmd_default_Click()

If MsgBox("Are you sure?", vbInformation + vbYesNo, "The P * * N Blocker") = vbYes Then
Open App.Path & "\pass.txt" For Output As #1
    Print #1, "ªe•aªggg"
Close #1
Check1.Value = 1
Check2.Value = 1
kill_task.Enabled = True
kill_task.Interval = 100
CreateStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run\", REG_SZ, "Porn Blocker", Getprogramfile & "Porn Blocker.exe"

MsgBox "Restore setting success" & vbNewLine & "Now your password is y4d0y666" & vbNewLine & "Change password early", vbInformation + vbOKOnly, "The P * * N Blocker"
Else
Exit Sub
End If

End Sub

Private Sub cmd_delcap_Click()
If lst_cap.ListIndex = -1 Then
MsgBox "Please select caption", vbInformation + vbOKOnly, "The P * * N Blocker"
Exit Sub
End If
lst_cap.RemoveItem (lst_cap.ListIndex)
HapusCaption lst_cap, App.Path & "\list.txt"
Call cmd_refreshcap_Click
lbl_jmlcap.Caption = lst_cap.ListCount


End Sub

Private Sub cmd_exit_Click()
If MsgBox("Are you sure?" & vbNewLine & "If you exit, the blocker will not active", vbInformation + vbYesNo, "The P * * N Blocker") = vbYes Then
TrayDelete
backup
Kill App.Path & "\kill.bat"
End
Else
Exit Sub
End If

End Sub

Private Sub cmd_hapus_Click()
If List1.ListIndex = -1 Then
MsgBox "Please select website address", vbInformation + vbOKOnly, "The P * * N Blocker"
Exit Sub
End If
List1.RemoveItem (List1.ListIndex)
hapus List1, GetSystemPath & "\drivers\etc\Hosts"
Call cmd_refresh_Click
lbl_jml.Caption = List1.ListCount

End Sub

Private Sub cmd_hide_Click()
Me.Hide
App.TaskVisible = False
ilang.Enabled = True
End Sub

Private Sub cmd_refresh_Click()
List1.Clear
LoadFileHost List1, GetSystemPath & "\drivers\etc\Hosts"
Text1.Text = ""
Text1.SetFocus
lbl_jml.Caption = List1.ListCount

End Sub

Private Sub cmd_refreshcap_Click()
lst_cap.Clear
Load_Caption lst_cap, App.Path & "\list.txt"
txt_blokcap.Text = ""
txt_blokcap.SetFocus
lbl_jmlcap.Caption = lst_cap.ListCount


End Sub


Private Sub cmd_save_seting_Click()
If Check1.Value = 0 Then
kill_task.Enabled = False
kill_task.Interval = 0
Else
kill_task.Enabled = True
kill_task.Interval = 100
End If

If Check2.Value = 1 Then
CreateStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run\", REG_SZ, "Porn Blocker", "C:\Program Files\Porn_Blocker\Porn Blocker.exe"
Else
DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Porn Blocker"
End If
End Sub

Private Sub cmd_set_Click()
If txtpwd.Text = "" Then
MsgBox "Please enter password", vbInformation + vbOKOnly, "The P * * N Blocker"
Exit Sub
End If
txtcryp.Text = crypt(txtpwd.Text, True)

Open App.Path & "\pass.txt" For Output As #1
    Print #1, txtcryp.Text
Close #1
MsgBox "Password saving success" & vbNewLine & "Now your password is :" & txtpwd.Text, vbInformation + vbOKOnly, "The P * * N Blocker"
txtpwd.Text = ""
txtpwd.SetFocus
End Sub


Private Sub Form_DblClick()
MsgBox "Copyright (C) YaDoY SofTwaRe DeVeLoPmEnT 2007", vbOKOnly + vbInformation, "The P * * N Blocker"
End Sub

Private Sub Form_Load()
mulai
TrayAdd hwnd, Picture1.Picture, "The Porn Blocker", MouseMove

Frame1.Visible = True
Frame1.Enabled = True
Frame2.Enabled = False
Frame2.Visible = False
Framecap.Visible = False
Framecap.Enabled = False


LoadFileHost List1, GetSystemPath & "\drivers\etc\Hosts"
lbl_jml.Caption = List1.ListCount

lst_cap.Clear
Load_Caption lst_cap, App.Path & "\list.txt"
lbl_jmlcap.Caption = lst_cap.ListCount

CreateStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run\", REG_SZ, "Porn Blocker", "C:\Program Files\Porn_Blocker\Porn Blocker.exe"

buat_kill
End Sub


Private Sub Form_Unload(Cancel As Integer)
TrayDelete
backup
Kill App.Path & "\kill.bat"
End

End Sub

Private Sub Frame1_Click()
Text1.SetFocus
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim cEvent As Single
cEvent = x / Screen.TwipsPerPixelX
Select Case cEvent
    Case MouseMove
        Debug.Print "MouseMove"
    Case LeftUp
        Debug.Print "Left Up"
    Case LeftDown
        Debug.Print "LeftDown"
    Case LeftDbClick
        Debug.Print "LeftDbClick"
    Case MiddleUp
        Debug.Print "MiddleUp"
    Case MiddleDown
        Debug.Print "MiddleDown"
    Case MiddleDbClick
        Debug.Print "MiddleDbClick"
    Case RightUp
        Debug.Print "RightUp": PopupMenu mnu
    Case RightDown
        Debug.Print "RightDown"
    Case RightDbClick
        Debug.Print "RightDbClick"
End Select
End Sub

Private Sub ilang_Timer()
On Error Resume Next
Dim bunuh As Long
frm_main.Hide
App.TaskVisible = False
For bunuh = 0 To lst_cap.ListCount - 1
kill_IE (lst_cap.list(bunuh))
Tonjok (lst_cap.list(bunuh))
Next

End Sub


Private Sub kill_task_Timer()
Hajar "TASK MANAGER"
Hajar "CMD"
Hajar "Command Prompt"
End Sub

Private Sub show_Click()
 frm_pass.show
End Sub


Private Sub buat_kill()
Open App.Path & "\kill.bat" For Output As #1
 Print #1, "taskkill /f /im iexplore.exe"
Close #1
End Sub



