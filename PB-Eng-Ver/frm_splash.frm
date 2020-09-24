VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_splash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   2580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   2520
      TabIndex        =   0
      Top             =   1680
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   6120
      Top             =   0
   End
   Begin VB.Image Image2 
      Height          =   885
      Left            =   2160
      Picture         =   "frm_splash.frx":0000
      Top             =   480
      Width           =   4755
   End
   Begin VB.Image Image1 
      Height          =   1695
      Left            =   360
      Picture         =   "frm_splash.frx":DBAA
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   1
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Height          =   2295
      Left            =   120
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "frm_splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
ProgressBar1.Value = ProgressBar1.Min
End Sub


Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 5
If ProgressBar1.Value = 10 Then
Label3.Caption = "Application Initialazing"
End If
If ProgressBar1.Value = 40 Then
Label3.Caption = "Loading Database"
End If
If ProgressBar1.Value = 80 Then
Label3.Caption = "Loading Complete"
End If

If ProgressBar1.Value >= ProgressBar1.Max Then
Unload Me
frm_main.Hide

End If

End Sub
