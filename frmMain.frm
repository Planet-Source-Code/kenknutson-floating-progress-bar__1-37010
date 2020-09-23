VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   600
      Top             =   2520
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   975
      Left            =   840
      TabIndex        =   0
      Top             =   1200
      Width           =   3015
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  frmFloat.MaxValue = 100
  frmFloat.CurrentValue = 0
  frmFloat.Show
  Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
  If frmFloat.CurrentValue = 99 Then
    Timer1.Enabled = False
    frmFloat.CurrentValue = frmFloat.MaxValue
  End If
  frmFloat.CurrentValue = frmFloat.CurrentValue + 1
End Sub
