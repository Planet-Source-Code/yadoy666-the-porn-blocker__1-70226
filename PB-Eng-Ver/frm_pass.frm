VERSION 5.00
Begin VB.Form frm_pass 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "  Login"
   ClientHeight    =   1680
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4005
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   4005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt2 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   360
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.TextBox txt1 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   3375
   End
   Begin Porn_Blocker.XpButton cmd_login 
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   1200
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
      Caption         =   "Login"
      ForeColor       =   -2147483630
      ForeHover       =   0
   End
   Begin VB.TextBox TxtPass 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "|"
      TabIndex        =   0
      Top             =   720
      Width           =   3375
   End
   Begin Porn_Blocker.XpButton cmd_cancel 
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   1200
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
      Caption         =   "Cancel"
      ForeColor       =   -2147483630
      ForeHover       =   0
   End
   Begin VB.Image Image1 
      Height          =   405
      Left            =   120
      Picture         =   "frm_pass.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   405
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Password"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   0
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "frm_pass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_cancel_Click()
Unload Me
End Sub

Private Sub cmd_login_Click()
On Error Resume Next
If TxtPass.Text = txt2.Text Then
 frm_main.ilang.Enabled = False
 frm_main.show
 frm_main.Text1.SetFocus
Else
MsgBox "Password not match, Login Failed", vbInformation + vbOKOnly, "The P * * N Blocker"
TxtPass.Text = ""
TxtPass.SetFocus
Exit Sub
End If
TxtPass.Text = ""
Unload Me
End Sub

Private Sub Form_Activate()
On Error Resume Next
Dim linestr As String
Open App.Path & "\pass.txt" For Input As #1
Line Input #1, linestr
Close #1
txt1.Text = linestr
txt2.Text = crypt(txt1.Text, False)

TxtPass.SetFocus
End Sub

