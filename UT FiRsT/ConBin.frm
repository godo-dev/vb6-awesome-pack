VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convert Binary"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4755
   Icon            =   "ConBin.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "ConBin.frx":030A
   ScaleHeight     =   166
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   317
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   1
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1943
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   135
      TabIndex        =   1
      Top             =   503
      Width           =   3255
   End
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
      Left            =   2175
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   863
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   0
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1463
      Width           =   3255
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HexaDecimal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3495
      TabIndex        =   7
      Top             =   1943
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Masukkan Suatu Angka Binary:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   135
      TabIndex        =   0
      Top             =   270
      Width           =   2685
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Angka Ini Sama Dengan:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   135
      TabIndex        =   3
      Top             =   1223
      Width           =   2130
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Decimal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3495
      TabIndex        =   5
      Top             =   1463
      Width           =   690
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim X   As String * 1
Dim I   As Byte
Dim Y   As Double
If Text1 <> "" Then
    On Error GoTo Keluar
    Text2(0) = "<Unknown>"
    Text2(1) = "<Unknown>"
    For I = 1 To Len(Text1)
        X = Mid(Text1, I, 1)
        Y = Y + ((2 ^ (Len(Text1) - I)) * Val(X))
    Next I
    Text2(0) = Y
    Text2(1) = Hex(Y)
Keluar:
    Text1.SetFocus
    utHomEnd
Else
    Text2(0) = ""
    Text2(1) = ""
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIForm1.MnuBiner.Checked = False
End Sub

Private Sub Text1_GotFocus()
    utHomEnd
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Not (KeyAscii = Asc("0") Or KeyAscii = Asc("1") Or KeyAscii = vbKeyBack) Then
    KeyAscii = 0
    Beep
End If
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Text1.ToolTipText = Text1.Text
End Sub

Private Sub Text2_GotFocus(Index As Integer)
    utHomEnd
End Sub

Private Sub Text2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Text2(Index).ToolTipText = Text2(Index).Text
End Sub
