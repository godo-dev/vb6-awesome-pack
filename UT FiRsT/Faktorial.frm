VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Faktorial"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4755
   Icon            =   "Faktorial.frx":0000
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "Faktorial.frx":030A
   ScaleHeight     =   166
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   317
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Proses"
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
      Left            =   1770
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1043
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   210
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1763
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   210
      TabIndex        =   1
      Top             =   683
      Width           =   4335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Faktorial Dari Angka Ini Adalah:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   210
      TabIndex        =   3
      Top             =   1530
      Width           =   2730
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Masukkan Suatu Angka:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   210
      TabIndex        =   0
      Top             =   450
      Width           =   2100
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim X As Double
Dim I As Byte
If Text1 <> "" Then
    On Error GoTo Keluar
    Text2 = "<Unknown>"
    X = Val(Text1)
If X > 1 Then
    For I = 1 To X - 1
        X = X * I
    Next I
End If
    Text2 = X
Keluar:
    Text1.SetFocus
    utHomEnd
Else
    Text2 = ""
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIForm1.MnuFak.Checked = False
End Sub

Private Sub Text1_GotFocus()
    utHomEnd
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
    KeyAscii = 0
    Beep
End If
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Text1.ToolTipText = Text1.Text
End Sub

Private Sub Text2_GotFocus()
    utHomEnd
End Sub

Private Sub Text2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Text2.ToolTipText = Text2.Text
End Sub
