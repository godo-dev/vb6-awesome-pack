VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convert HexaDecimal"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   4755
   Icon            =   "ConHex.frx":0000
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "ConHex.frx":030A
   ScaleHeight     =   166
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   317
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   138
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
      Left            =   2178
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   863
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   0
      Left            =   138
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1463
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   1
      Left            =   138
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1943
      Width           =   3255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Masukkan Suatu Angka HexaDecimal:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   270
      Width           =   3270
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   127
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3495
      TabIndex        =   5
      Top             =   1470
      Width           =   690
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Binary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3495
      TabIndex        =   7
      Top             =   1950
      Width           =   540
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim I, I0   As Byte
Dim J, N    As Double
Dim S, X, B As String
If Text1 <> "" Then
    Text2(0) = "<Unknown>"
    Text2(1) = "<Unknown>"
    On Error GoTo Keluar
    For I = 1 To Len(Text1)
        S = ""
        X = Mid(Text1, I, 1)
        Select Case LCase(X)
            Case Is = "a"
                N = 10
            Case Is = "b"
                N = 11
            Case Is = "c"
                N = 12
            Case Is = "d"
                N = 13
            Case Is = "e"
                N = 14
            Case Is = "f"
                N = 15
            Case Else
                N = Val(X)
        End Select
        S = utBin(N) & S
        If Len(S) < 4 And I <> 1 Then
            For I0 = 1 To (4 - Len(S))
                S = "0" & S
            Next I0
        End If
        B = B & S
    Next I
    For I = 1 To Len(B)
        J = Val(Mid(B, I, 1)) * (2 ^ (Len(B) - I)) + J
    Next I
    Text2(0) = J
    Text2(1) = utBin(J)
Keluar:
    Text1.SetFocus
    utHomEnd
Else
    Text2(0) = ""
    Text2(1) = ""
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIForm1.MnuHex.Checked = False
End Sub

Private Sub Text1_GotFocus()
    utHomEnd
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or _
    (KeyAscii >= Asc("A") And KeyAscii <= Asc("F")) Or _
    KeyAscii = vbKeyBack) Then
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

