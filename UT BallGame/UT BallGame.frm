VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6975
   Icon            =   "UT BallGame.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox utBall 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   1080
      Picture         =   "UT BallGame.frx":0882
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   1
      Top             =   0
      Width           =   450
   End
   Begin VB.PictureBox utStick 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   0
      Picture         =   "UT BallGame.frx":0C52
      ScaleHeight     =   135
      ScaleWidth      =   930
      TabIndex        =   0
      Top             =   120
      Width           =   930
   End
   Begin VB.Timer utTimer 
      Interval        =   10
      Left            =   360
      Top             =   840
   End
   Begin VB.Shape utBlack 
      BackColor       =   &H00000000&
      BorderWidth     =   3
      Height          =   495
      Left            =   1680
      Top             =   2520
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim X       As Long
Dim GerakX  As Byte
Dim GerakY  As Byte
Dim GerakXY As Byte
Dim Turun   As Boolean
Dim Kanan   As Boolean
Dim Pesan   As String

Const utPix = 120
Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long

Const HWND_TOP = 0
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const HWND_DESKTOP = 0
Const HWND_BOTTOM = 1

Private Function utGerak(ByVal Gbr As PictureBox, ByVal X As _
    Byte, ByVal Y As Byte)
Dim Xx, Yy  As Long
Select Case X
    Case 1
        Xx = Gbr.Left + utPix
    Case 2
        Xx = Gbr.Left - utPix
End Select
Select Case Y
    Case 1
        Yy = Gbr.Top + utPix
    Case 2
        Yy = Gbr.Top - utPix
End Select
    Gbr.Move (Xx), Yy
End Function
Private Sub Form_Activate()
    utBlack.Height = Me.Height
    utBlack.Width = Me.Width
    utBlack.Move (0), 0
    utStick.Left = 0.5 * (Me.Width - utStick.Width)
    utStick.Top = Me.Height - utStick.Height - utPix
    X = Int(Rnd * Me.Width) + 1
    utBall.Move (X), utPix
    utBall.BackColor = Me.BackColor
    GerakX = Int(Rnd * 2) + 1
    GerakY = 1
    Dim Ctrl      As Control
    On Error Resume Next
    For Each Ctrl In Controls
        Ctrl.Refresh
    Next Ctrl
End Sub

Private Sub Form_Initialize()
    Pesan = "UltimaTech® Ball Game...Created by Junian Triajianto..."
End Sub

Private Sub Form_Load()
    Randomize
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    utStick.Move X - 0.5 * utStick.Width
If utStick.Left <= Me.Left Then
    utStick.Left = 0
ElseIf utStick.Left >= Me.Width - utStick.Width Then
    utStick.Left = Me.Width - utStick.Width
End If
End Sub

Private Sub utTimer_Timer()
Dim PosBall   As Long
PosBall = (utBall.Top + utBall.Height)
Call utGerak(utBall, GerakX, GerakY)
If utBall.Left <= Me.Left Then
    utBall.Left = Me.Left
    GerakX = 1
ElseIf utBall.Left + utBall.Width >= Me.Width Then
    utBall.Left = Me.Width - utBall.Width
    GerakX = 2
End If
If utBall.Top <= Me.Top Then
    utBall.Top = Me.Top
    GerakY = 1
ElseIf utStick.Top <= PosBall And (utStick.Left <= utBall.Left _
    And utStick.Left + utStick.Width >= utBall.Left + utBall.Width) Then
    utBall.Top = utStick.Top - utBall.Height
    GerakY = 2
End If
End Sub

