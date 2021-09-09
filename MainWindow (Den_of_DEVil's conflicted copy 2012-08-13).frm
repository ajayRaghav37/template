VERSION 5.00
Begin VB.Form MainWindow 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "ProgramName"
   ClientHeight    =   8610
   ClientLeft      =   4095
   ClientTop       =   -9045
   ClientWidth     =   11010
   ForeColor       =   &H00E0E0E0&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   11010
   Begin VB.PictureBox picMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00F9F9F9&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   2400
      ScaleHeight     =   4215
      ScaleWidth      =   2175
      TabIndex        =   7
      Top             =   2040
      Width           =   2175
      Begin VB.Label lblShortcut 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Del"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007F7F7F&
         Height          =   195
         Index           =   0
         Left            =   1725
         TabIndex        =   9
         Top             =   30
         Width           =   255
      End
      Begin VB.Label lblSubMenu 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00998200&
         BackStyle       =   0  'Transparent
         Caption         =   "New File"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H007F7F7F&
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   8
         Top             =   30
         Width           =   660
      End
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.Label lblMenu 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00998200&
      BackStyle       =   0  'Transparent
      Caption         =   " Help "
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007F7F7F&
      Height          =   195
      Index           =   4
      Left            =   1905
      TabIndex        =   12
      Top             =   60
      Width           =   450
   End
   Begin VB.Label lblMenu 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00998200&
      BackStyle       =   0  'Transparent
      Caption         =   " Media "
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007F7F7F&
      Height          =   195
      Index           =   3
      Left            =   1335
      TabIndex        =   11
      Top             =   60
      Width           =   570
   End
   Begin VB.Label lblMenu 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00998200&
      BackStyle       =   0  'Transparent
      Caption         =   " View "
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007F7F7F&
      Height          =   195
      Index           =   2
      Left            =   870
      TabIndex        =   10
      Top             =   60
      Width           =   465
   End
   Begin VB.Label lblMenu 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00998200&
      BackStyle       =   0  'Transparent
      Caption         =   " Edit "
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007F7F7F&
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   6
      Top             =   60
      Width           =   390
   End
   Begin VB.Label lblMenu 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00998200&
      BackStyle       =   0  'Transparent
      Caption         =   " File "
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007F7F7F&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   60
      Width           =   360
   End
   Begin VB.Label lblProgramName 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ProgramName Major.Minor.Build"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00606060&
      Height          =   180
      Left            =   8145
      TabIndex        =   4
      Top             =   60
      Width           =   1950
   End
   Begin VB.Label lblControlBox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007F7F7F&
      Height          =   270
      Index           =   0
      Left            =   10200
      TabIndex        =   3
      Top             =   0
      Width           =   225
   End
   Begin VB.Label lblControlBox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007F7F7F&
      Height          =   285
      Index           =   1
      Left            =   10455
      TabIndex        =   2
      Top             =   0
      Width           =   195
   End
   Begin VB.Label lblControlBox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "r"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007F7F7F&
      Height          =   270
      Index           =   2
      Left            =   10680
      TabIndex        =   1
      Top             =   0
      Width           =   225
   End
   Begin VB.Label lblTitleBar 
      BackColor       =   &H00F2F2F2&
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11055
   End
   Begin VB.Shape shpToolBar 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00998200&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   0
      Top             =   285
      Width           =   11055
   End
   Begin VB.Shape shpStatusBar 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00998200&
      FillStyle       =   0  'Solid
      Height          =   1815
      Left            =   0
      Top             =   7200
      Width           =   11055
   End
End
Attribute VB_Name = "MainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MenuItems() As Variant
Dim Shortcuts() As Variant
Dim TempNum As Integer
Dim MenuUnderlined As Boolean
Dim PrevLeft As Integer, PrevWidth As Integer, PrevTop As Integer, PrevHeight As Integer
Dim IsMaximized As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyMenu Then
        If MenuUnderlined Then
            For TempNum = lblMenu.LBound To lblMenu.UBound
                lblMenu(TempNum).Caption = Replace$(lblMenu(TempNum).Caption, "&", vbNullString)
            Next
            MenuUnderlined = False
        Else
            For TempNum = lblMenu.LBound To lblMenu.UBound
                lblMenu(TempNum).Caption = Rnd
            Next
            MenuUnderlined = True
        End If
    End If
End Sub

Private Sub Form_Load()
    Load Resizer
    Resizer.Show
    Resizer.Visible = False
    Resizer.Height = Screen.Height * 0.72
    Resizer.Width = Resizer.Height * 4 / 3
    Show
    MainWindow.ZOrder 0
    Resizer.Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2, Width + 30, Height + 30
    MainWindow.Move Resizer.Left + 15, Resizer.Top + 15, Resizer.Width - 30, Resizer.Height - 30
    Resizer.Visible = True
    MainWindow.ZOrder 0
    lblProgramName.Caption = App.ProductName & " " & Trim(Str(App.Major)) & "." & Trim(Str(App.Minor)) & "." & Trim(Str(App.Revision))
    shpStatusBar.Height = ((Height / 5.4) + (4.7 * (19.05 - Height)) / (Height ^ 2))
    shpToolBar.Height = shpStatusBar.Height / 2
    MenuItems = Array("New File", "Technology Ahead Always......", "Hakuna Matata", "Zindagi Toofani hai", "Poke'mon", "Let's do it")
    Shortcuts = Array("Ctrl+N", "Del", "Backspace", ">", "", "")
    CreateMenu 300, 120
    Form_Resize
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    DeHover
End Sub

Private Sub Form_Resize()
    shpToolBar.Width = Width + 15
    shpStatusBar.Move 0, Height - shpStatusBar.Height + 15, Width + 15, shpStatusBar.Height
    lblTitleBar.Width = Width
    lblControlBox(2).Left = Width - lblControlBox(2).Width - 60
    lblControlBox(1).Left = lblControlBox(2).Left - lblControlBox(1).Width - 45
    lblControlBox(0).Left = lblControlBox(1).Left - lblControlBox(0).Width - 45
    lblProgramName.Left = lblControlBox(0).Left - lblProgramName.Width - 120
    shpBorder.Move 0, 0, Width, Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Resizer
End Sub

Private Sub lblControlBox_Click(Index As Integer)
    Select Case Index
        Case 0
            WindowState = 1
        Case 1
            MaxResMe
        Case 2
            Unload Me
    End Select
End Sub

Private Sub lblControlBox_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    DeHover
    lblControlBox(Index).ForeColor = vbBlack
End Sub

Private Sub DeHover()
    If lblControlBox(0).ForeColor = vbBlack Then
        lblControlBox(0).ForeColor = &H7F7F7F
    End If
    If lblControlBox(1).ForeColor = vbBlack Then
        lblControlBox(1).ForeColor = &H7F7F7F
    End If
    If lblControlBox(2).ForeColor = vbBlack Then
        lblControlBox(2).ForeColor = &H7F7F7F
    End If
    For TempNum = lblMenu.LBound To lblMenu.UBound
        lblMenu(TempNum).BackStyle = 0
        lblMenu(TempNum).ForeColor = &H7F7F7F
    Next
    For TempNum = lblSubMenu.LBound To lblSubMenu.UBound
        lblSubMenu(TempNum).BackStyle = 0
        lblSubMenu(TempNum).ForeColor = &H7F7F7F
        lblShortcut(TempNum).ForeColor = &H7F7F7F
    Next
End Sub

Private Sub lblMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    DeHover
    lblMenu(Index).BackStyle = 1
    lblMenu(Index).ForeColor = vbWhite
End Sub

Private Sub lblProgramName_DblClick()
    lblTitleBar_DblClick
End Sub

Private Sub lblProgramName_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblTitleBar_MouseMove Button, Shift, x, y
End Sub

Private Sub lblShortcut_Click(Index As Integer)
    lblSubMenu_Click Index
End Sub

Private Sub lblShortcut_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    lblSubMenu_MouseMove Index, Button, Shift, x, y
End Sub

Private Sub lblSubMenu_Click(Index As Integer)
    MsgBox lblSubMenu(Index).Caption
End Sub

Private Sub lblSubMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    DeHover
    lblSubMenu(Index).BackStyle = 1
    lblSubMenu(Index).ForeColor = vbWhite
    lblShortcut(Index).ForeColor = vbWhite
End Sub

Private Sub lblTitleBar_DblClick()
    MaxResMe
End Sub

Private Sub lblTitleBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    DeHover
End Sub

Private Sub CreateMenu(MenuTop As Integer, MenuLeft As Integer)
    Dim MaxWidth As Integer
    picMenu.Top = MenuTop
    picMenu.Left = MenuLeft
    lblSubMenu(0).Caption = "  " & MenuItems(0)
    If Shortcuts(0) = ">" Then
        lblShortcut(0).Caption = "4"
        lblShortcut(0).FontName = "Webdings"
        lblShortcut(0).FontSize = 10
    Else
        lblShortcut(0).Caption = Shortcuts(0)
        lblShortcut(0).FontName = "Segoe UI"
        lblShortcut(0).FontSize = 8
    End If
    MaxWidth = 1800
    For TempNum = 1 To UBound(MenuItems)
        Load lblSubMenu(TempNum)
        Load lblShortcut(TempNum)
        lblSubMenu(TempNum).Caption = "  " & MenuItems(TempNum)
        If Shortcuts(TempNum) = ">" Then
            lblShortcut(TempNum).Caption = "4"
            lblShortcut(TempNum).FontName = "Webdings"
            lblShortcut(TempNum).FontSize = 10
        Else
            lblShortcut(TempNum).Caption = Shortcuts(TempNum)
            lblShortcut(TempNum).FontName = "Segoe UI"
            lblShortcut(0).FontSize = 8
        End If
        lblSubMenu(TempNum).Left = lblSubMenu(0).Left
        lblSubMenu(TempNum).Top = lblSubMenu(TempNum - 1).Top + lblSubMenu(TempNum - 1).Height + 30
        lblShortcut(TempNum).Top = lblSubMenu(TempNum).Top
        If lblSubMenu(TempNum).Width + 480 + lblShortcut(TempNum).Width > MaxWidth Then
            MaxWidth = lblSubMenu(TempNum).Width + 480 + lblShortcut(TempNum).Width
        End If
        lblSubMenu(TempNum).Visible = True
        lblShortcut(TempNum).Visible = True
        lblShortcut(TempNum).ZOrder 0
    Next
    For TempNum = 0 To UBound(MenuItems)
        lblSubMenu(TempNum).Width = MaxWidth
        lblShortcut(TempNum).Left = lblSubMenu(TempNum).Left + MaxWidth - lblShortcut(TempNum).Width - 120
    Next
    picMenu.Width = MaxWidth
    picMenu.Height = lblSubMenu(UBound(MenuItems)).Height + lblSubMenu(UBound(MenuItems)).Top + 30
End Sub

Private Sub MaxResMe()
    If IsMaximized Then
        'Left = PrevLeft
        'Top = PrevTop
        'Width = PrevWidth
        'Height = PrevHeight
        Resizer.WindowState = 0
        IsMaximized = False
        lblControlBox(1).Caption = 1
        shpBorder.Visible = True
        DeHover
    Else
        Resizer.WindowState = 2
        'PrevLeft = Left
        'PrevTop = Top
        'PrevWidth = Width
        'PrevHeight = Height
        'Left = SysInfo.WorkAreaLeft
        'Top = SysInfo.WorkAreaTop
        'Width = SysInfo.WorkAreaWidth
        'Height = SysInfo.WorkAreaHeight
        IsMaximized = True
        lblControlBox(1).Caption = 2
        shpBorder.Visible = False
    End If
End Sub
