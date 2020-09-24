VERSION 5.00
Begin VB.Form frm_about 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  About"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   840
      LinkTimeout     =   35
      ScaleHeight     =   2655
      ScaleWidth      =   4695
      TabIndex        =   1
      Top             =   360
      Width           =   4695
      Begin VB.VScrollBar VScroll1 
         Height          =   6315
         Left            =   0
         TabIndex        =   3
         Top             =   2040
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   6735
         HideSelection   =   0   'False
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   "frm_about.frx":0000
         Top             =   2520
         Width           =   4335
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   0
      Top             =   1440
   End
   Begin VB.Image Image1 
      Height          =   885
      Left            =   1680
      Picture         =   "frm_about.frx":034D
      Top             =   3960
      Width           =   4755
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "The Porn Blocker is freeware but without any warranty. Use it with your own risk. Bugs please send to yadoy666@gmail.com."
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   3240
      Width           =   6015
   End
End
Attribute VB_Name = "frm_about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()

    Timer1.Enabled = True
End Sub


Private Sub Form_Load()
Dim lReturn As Long
frm_about.show
    Timer1.Interval = 35
    VScroll1.Max = Picture1.Height
    VScroll1.Min = 0 - Text1.Height
    VScroll1.Value = VScroll1.Max


End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
GotoVal = Me.Height / 2


For Gointo = 1 To GotoVal
    'NEW ADDITION NEXT LINE


    DoEvents
        Me.Height = Me.Height - 10
        'Me.Top = (Screen.Height - Me.Height) \ 2
        If Me.Height <= 11 Then GoTo horiz
    Next Gointo


    'This is the width part of the same sequence above
horiz:
    Me.Height = 30
    GotoVal = Me.Width / 2


    For Gointo = 1 To GotoVal
        'NEW ADDITION NEXT LINE


        DoEvents
            Me.Width = Me.Width - 10
            'Me.Left = (Screen.Width - Me.Width) \ 2
            If Me.Width <= 11 Then End
        Next Gointo
        
Unload Me

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Timer1_Timer()


    If VScroll1.Value >= VScroll1.Min + 20 Then
         VScroll1.Value = VScroll1.Value - 35
    Else
         VScroll1.Value = VScroll1.Max

         DoEvents
        End If

            Text1.Top = VScroll1.Value
            Text1.Visible = True

            DoEvents
        End Sub





