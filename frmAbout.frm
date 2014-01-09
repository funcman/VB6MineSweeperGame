VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   "关于"
   ClientHeight    =   2145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2955
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   2955
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "何晏清 徐大伟 制作"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "扫雷 V 1.0"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnOK_Click()
    frmAbout.Hide
End Sub
