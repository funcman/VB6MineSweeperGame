VERSION 5.00
Begin VB.Form frmWin 
   Caption         =   "Win!  : )"
   ClientHeight    =   2025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5640
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2025
   ScaleWidth      =   5640
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblWin 
      Caption         =   "扫完了！ T_T"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5295
   End
End
Attribute VB_Name = "frmWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnOK_Click()
    Call NewGame
    start = False
    frmWin.Hide

End Sub

Private Sub Form_Load()
    start = False

End Sub
