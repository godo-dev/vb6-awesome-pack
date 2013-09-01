VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lucky Seven Game"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7395
   Icon            =   "Lucky Seven.frx":0000
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "Lucky Seven.frx":030A
   ScaleHeight     =   371
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   493
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Cmd 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Spin"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3030
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Luck 
      BackStyle       =   0  'Transparent
      Caption         =   "Luck: 0,00%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Label Spin 
      BackStyle       =   0  'Transparent
      Caption         =   "Spin: 0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   3480
      Width           =   2775
   End
   Begin VB.Label Skor 
      BackStyle       =   0  'Transparent
      Caption         =   "Skor: 0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   4
      DrawMode        =   6  'Mask Pen Not
      Height          =   1695
      Left            =   2130
      Shape           =   4  'Rounded Rectangle
      Top             =   2880
      Width           =   3135
   End
   Begin VB.Label Msg 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lucky Seven Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   615
      Left            =   1470
      TabIndex        =   4
      Top             =   2160
      Width           =   4455
   End
   Begin VB.Label Lbl7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   45.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1335
      Index           =   2
      Left            =   4470
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Lbl7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   45.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1335
      Index           =   1
      Left            =   3030
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Lbl7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   45.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1335
      Index           =   0
      Left            =   1590
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim X, Y As Long

Private Sub Cmd_Click()
Dim I As Byte
If Cmd.Caption = "&Spin" Then
    Y = Y + 1
    Timer1.Enabled = True
    Cmd.Caption = "&Stop"
ElseIf Cmd.Caption = "&Stop" Then
    Timer1.Enabled = False
    Cmd.Caption = "&Spin"
    For I = 0 To 2
        If Lbl7(I).Caption = "7" Then
            X = X + 1
            Skor = "Skor: " & X
        End If
    Next I
    Spin = "Spin: " & Y
    Luck = "Luck: " & FormatPercent(X / Y, 2)
End If
End Sub

Private Sub Form_Load()
Dim I As Byte
    Randomize
For I = 0 To 2
    Lbl7(I).ForeColor = vbBlack
Next I
End Sub

Private Sub Form_Unload(Cancel As Integer)
    X = 0
    Y = 0
    MDIForm1.Mnu7.Checked = False
End Sub

Private Sub Lbl7_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Lbl7(Index).ToolTipText = Lbl7(Index).Caption
End Sub

Private Sub Luck_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Luck.ToolTipText = Luck.Caption
End Sub

Private Sub Skor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Skor.ToolTipText = Skor.Caption
End Sub

Private Sub Spin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Spin.ToolTipText = Spin.Caption
End Sub

Private Sub Timer1_Timer()
Dim I As Byte
For I = 0 To 2
    Lbl7(I) = Int(Rnd * 10)
    If Lbl7(I) = "7" Then
        Lbl7(I).ForeColor = vbRed
    Else
        Lbl7(I).ForeColor = vbBlack
    End If
Next I
End Sub
