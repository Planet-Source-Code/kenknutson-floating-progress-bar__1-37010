VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFloat 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1200
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   270
   ScaleWidth      =   1200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar pbarCurrent 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmFloat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type POINTAPI
        x As Long
        y As Long
End Type

Public XSizeAsTwips As Integer
Public YSizeAsTwips As Integer
Public CurrentValue As Double
Public MaxValue As Double

Private Sub PositionFormAtCursor()
  Dim x As Long
  Dim y As Long
  Dim PA As POINTAPI
  Dim pos As String
  Dim FormWidthPix As Single
  Dim FormHeightPix As Single
  Dim FormWidthTwips As Single
  Dim FormHeightTwips As Single
  Dim FormPositionX As Single
  Dim FormPositionY As Single
  Dim ScreenWidthPix As Single
  Dim ScreenHeightPix As Single
  Dim CursorXPix As Single
  Dim CursorYPix As Single
  
  ScreenWidthPix = frmFloat.ScaleX(Screen.Width, vbTwips, vbPixels)
  ScreenHeightPix = frmFloat.ScaleY(Screen.Height, vbTwips, vbPixels)
  
  lngRET = GetCursorPos(PA)
  CursorXPix = PA.x
  CursorYPix = PA.y
  
  FormWidthTwips = frmFloat.pbarCurrent.Width '+ 100
  FormHeightTwips = frmFloat.pbarCurrent.Height '+ 100
  FormWidthPix = frmFloat.ScaleX(FormWidthTwips, vbTwips, vbPixels)
  FormHeightPix = frmFloat.ScaleY(FormHeightTwips, vbTwips, vbPixels)
  
  If CursorXPix + FormWidthPix + 20 > ScreenWidthPix Then
    FormPositionX = CursorXPix - FormWidthPix - 20
  Else
    FormPositionX = CursorXPix + 20
  End If
  
  If CursorYPix + FormHeightPix + 20 > ScreenHeightPix Then
    FormPositionY = CursorYPix - FormHeightPix - 20
  Else
    FormPositionY = CursorYPix + 20
  End If
  
  FormPositionX = frmFloat.ScaleX(FormPositionX, vbPixels, vbTwips)
  FormPositionY = frmFloat.ScaleY(FormPositionY, vbPixels, vbTwips)
  
  frmFloat.Move FormPositionX, FormPositionY, FormWidthTwips, FormHeightTwips
  StayOnTop frmFloat
  
End Sub

Sub StayOnTop(TheForm As Form)
'This Sub will keep your form(s) on top of everything
'else. Use this Sub like this:
'"StayOnTop Me"
'Put the code above in the Form_Load Sub.
  Const HWND_TOPMOST = -1
  Const SWP_NOMOVE = 2
  Const SWP_NOSIZE = 1
  Const Flags = SWP_NOMOVE Or SWP_NOSIZE
  
  SetWinOnTop = SetWindowPos(TheForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
End Sub

Private Sub Form_Activate()
  If XSizeAsTwips = 0 Then XSizeAsTwips = 1200
  If YSizeAsTwips = 0 Then YSizeAsTwips = 270
  Me.Width = XSizeAsTwips
  Me.Height = YSizeAsTwips
  pbarCurrent.Width = Me.Width
  pbarCurrent.Height = Me.Height
  
  pbarCurrent.Value = 0
  pbarCurrent.Max = MaxValue
  
  On Error Resume Next
  
  Do While CurrentValue <= MaxValue: DoEvents
    PositionFormAtCursor
    pbarCurrent.Value = CurrentValue
  Loop 'Until CurrentValue >= MaxValue - 0.1
  Me.Hide
  Unload Me
  Set frmFloat = Nothing
End Sub

