VERSION 5.00
Begin VB.Form frmTaskSample 
   Caption         =   "Animated Task Tray Sample"
   ClientHeight    =   5010
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5070
   Icon            =   "frmTaskSample.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   1200
      Top             =   120
   End
   Begin VB.Image imgMoon 
      Height          =   480
      Index           =   7
      Left            =   3360
      Picture         =   "frmTaskSample.frx":0442
      Top             =   3600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMoon 
      Height          =   480
      Index           =   6
      Left            =   2880
      Picture         =   "frmTaskSample.frx":0884
      Top             =   3600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMoon 
      Height          =   480
      Index           =   5
      Left            =   2400
      Picture         =   "frmTaskSample.frx":0CC6
      Top             =   3600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMoon 
      Height          =   480
      Index           =   4
      Left            =   1920
      Picture         =   "frmTaskSample.frx":1108
      Top             =   3600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMoon 
      Height          =   480
      Index           =   3
      Left            =   1440
      Picture         =   "frmTaskSample.frx":154A
      Top             =   3600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMoon 
      Height          =   480
      Index           =   2
      Left            =   960
      Picture         =   "frmTaskSample.frx":198C
      Top             =   3600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMoon 
      Height          =   480
      Index           =   1
      Left            =   480
      Picture         =   "frmTaskSample.frx":1DCE
      Top             =   3600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMoon 
      Height          =   480
      Index           =   0
      Left            =   0
      Picture         =   "frmTaskSample.frx":2210
      Top             =   3600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Menu mnuLeftClick 
      Caption         =   "Left Click"
      Begin VB.Menu mnuBeep 
         Caption         =   "Siren"
      End
   End
   Begin VB.Menu mnuRightClick 
      Caption         =   "Right Click"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmTaskSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Declare Function SetForegroundWindow Lib "user32" ( _
    ByVal hwnd As Long) As Long

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" ( _
    ByVal dwMessage As Long, _
    pnid As NOTIFYICONDATA) As Boolean
    
Private Declare Function Beep Lib "kernel32" ( _
    ByVal dwFreq As Long, _
    ByVal dwDuration As Long) As Long
    
Const NIM_ADD = &H0
Const NIM_MODIFY = &H1
Const NIM_DELETE = &H2
Const NIF_MESSAGE = &H1
Const NIF_ICON = &H2
Const NIF_TIP = &H4
Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_LBUTTONDBLCLK = &H203
Const WM_RBUTTONDOWN = &H204
Const WM_RBUTTONUP = &H205
Const WM_RBUTTONDBLCLK = &H206
    
Dim MoonPhase As Single
Dim NID As NOTIFYICONDATA

Private Sub Form_Load()
MoonPhase = 0
mnuLeftClick.Visible = False
mnuRightClick.Visible = False
With NID
        .cbSize = Len(NID)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = imgMoon(0).Picture
        .szTip = " Double-Click to open! " & vbNullChar
    End With

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Result As Long
    Dim msg As Long
    If Me.ScaleMode = vbPixels Then
        msg = x
    Else
        msg = x / Screen.TwipsPerPixelX
    End If

    Select Case msg
        Case WM_LBUTTONUP
            Me.PopupMenu Me.mnuLeftClick
        Case WM_LBUTTONDBLCLK
            Me.WindowState = vbNormal
            Me.Show
        Case WM_RBUTTONUP
            Me.PopupMenu Me.mnuRightClick
    End Select
End Sub

Private Sub Form_Resize()
If Me.WindowState = 1 Then
    Me.Hide
    Timer1.Enabled = True
    Shell_NotifyIcon NIM_ADD, NID
Else
    Timer1.Enabled = False
    Shell_NotifyIcon NIM_DELETE, NID
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Shell_NotifyIcon NIM_DELETE, NID
End Sub

Private Sub mnuBeep_Click()
Dim x As Integer, y As Integer
For y = 1 To 10
    For x = 0 To 2000 Step 100
        Beep x, 1
    Next x
    For x = 2000 To 0 Step -100
        Beep x, 1
    Next x
Next y
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
NID.hIcon = imgMoon(MoonPhase).Picture
Shell_NotifyIcon NIM_MODIFY, NID

If MoonPhase > 6 Then MoonPhase = 0 Else MoonPhase = MoonPhase + 1
End Sub
