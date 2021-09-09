VERSION 5.00
Begin VB.Form Resizer 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   8460
   ClientLeft      =   5205
   ClientTop       =   -8580
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   10860
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "Resizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type FormSize
  Width As Long
  Height As Long
  ScaleWidth As Long
  ScaleHeight As Long
  BorderWidth As Long
  BorderHeight As Long
  zError As Boolean
End Type

Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Dim TMP As FormSize, rectClient As RECT, rectWindow As RECT

Private Sub Form_Activate()
    MainWindow.ZOrder 0
End Sub

Public Sub Form_Resize()
    GetClientRect hWnd, rectClient
    GetWindowRect hWnd, rectWindow
    If WindowState <> 2 Then
        MainWindow.shpBorder.Move 0, 0, Width - 30, Height - 30
        MainWindow.Move Left + 15, Top + 15, Width - 30, Height - 30
        'shpToolBar.Width = Width + 15
        'shpStatusBar.Move 0, Height - shpStatusBar.Height + 15, Width + 15, shpStatusBar.Height
        'lblTitleBar.Width = Width
    Else
        'MainWindow.Move Left + (Width - Screen.TwipsPerPixelX * rectClient.Right + Screen.TwipsPerPixelX * rectClient.Left) / 2, Top + (Height - Screen.TwipsPerPixelY * rectClient.Bottom + Screen.TwipsPerPixelY * rectClient.Top) / 2, Width, Height
        MainWindow.Move Screen.TwipsPerPixelX * rectClient.Left, Screen.TwipsPerPixelY * rectClient.Top, Screen.TwipsPerPixelX * (rectClient.Right - rectClient.Left), Height
    End If
End Sub

Private Function GetFormSize(ByVal hWnd As Long) As FormSize

  If (GetClientRect(hWnd, rectClient) <> 0) And (GetWindowRect(hWnd, rectWindow) <> 0) Then
    TMP.Width = rectWindow.Right - rectWindow.Left
    TMP.Height = rectWindow.Bottom - rectWindow.Top
    TMP.ScaleWidth = rectClient.Right - rectClient.Left
    TMP.ScaleHeight = rectClient.Bottom - rectClient.Top
    TMP.BorderWidth = TMP.Width - TMP.ScaleWidth
    TMP.BorderHeight = TMP.Height - TMP.ScaleHeight
    TMP.zError = False
  Else
    TMP.Width = 0
    TMP.Height = 0
    TMP.ScaleWidth = 0
    TMP.ScaleHeight = 0
    TMP.BorderWidth = 0
    TMP.BorderHeight = 0
    TMP.zError = True
  End If
  GetFormSize = TMP
End Function

Public Function ResizeForm(ByVal hWnd As Long, Optional ByVal zWidth As Long = -1, Optional ByVal zHeight As Long = -1) As Boolean
  Const SWP_NOACTIVATE = &H10
  Const SWP_NOMOVE = &H2
  Const SWP_NOOWNERZORDER = &H200
'  Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
  Const SWP_NOZORDER = &H4
  Const zFLAGS = SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOOWNERZORDER Or SWP_NOZORDER
  Dim TMP As FormSize
  ResizeForm = False
  TMP = GetFormSize(hWnd)
  If TMP.zError = False Then
    If zWidth = -1 Then zWidth = TMP.Width
    If zHeight = -1 Then zHeight = TMP.Height
    If SetWindowPos(hWnd, 0, 0, 0, zWidth + TMP.BorderWidth, zHeight + TMP.BorderHeight, zFLAGS) <> 0 Then ResizeForm = True
  End If
End Function
