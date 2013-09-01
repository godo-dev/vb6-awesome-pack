VERSION 5.00
Begin VB.Form Form9 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fungsi Kuadrat dan Linear"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   3870
   Icon            =   "LinKua.frx":0000
   LinkTopic       =   "Form9"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "LinKua.frx":030A
   ScaleHeight     =   225
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   258
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   1628
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   1628
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   1628
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   3068
      TabIndex        =   13
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   1988
      TabIndex        =   10
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   788
      TabIndex        =   6
      Text            =   "1"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   1988
      TabIndex        =   4
      Top             =   240
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   788
      TabIndex        =   1
      Text            =   "1"
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Hitung!!!"
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
      Left            =   1328
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Solusi (x, y):"
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
      Index           =   9
      Left            =   308
      TabIndex        =   17
      Top             =   2520
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Diskriminan: "
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
      Index           =   6
      Left            =   308
      TabIndex        =   15
      Top             =   2160
      Width           =   1110
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      DrawMode        =   6  'Mask Pen Not
      Height          =   1335
      Left            =   128
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   3615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y  ="
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
      Index           =   5
      Left            =   308
      TabIndex        =   5
      Top             =   840
      Width           =   360
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
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
      Index           =   4
      Left            =   2813
      TabIndex        =   12
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
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
      Index           =   3
      Left            =   2573
      TabIndex        =   11
      Top             =   840
      Width           =   165
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
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
      Index           =   2
      Left            =   1733
      TabIndex        =   9
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
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
      Index           =   1
      Left            =   1508
      TabIndex        =   8
      Top             =   720
      Width           =   135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
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
      Index           =   0
      Left            =   1373
      TabIndex        =   7
      Top             =   840
      Width           =   165
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y  ="
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
      Index           =   12
      Left            =   308
      TabIndex        =   0
      Top             =   240
      Width           =   360
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
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
      Index           =   13
      Left            =   1733
      TabIndex        =   3
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
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
      Index           =   14
      Left            =   1373
      TabIndex        =   2
      Top             =   240
      Width           =   165
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      DrawMode        =   6  'Mask Pen Not
      Height          =   1215
      Left            =   128
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim A, B, C, P, Q As Double

Private Function Y(ByVal X As Double) As Double
    Y = (P * X) + Q
End Function

Private Sub Command1_Click()
Dim I                   As Byte
Dim X1, X2, Y1, Y2, Dis As Double
For I = 0 To 4
    Text1(I) = Val(Text1(I))
Next I
If Val(Text1(3)) = 0 Then
    MsgBox "Nilai 'p' tidak boleh sama dengan '0'!", vbExclamation, "UltimaTech"
    Text1(3).SetFocus
    utHomEnd
    Exit Sub
ElseIf Val(Text1(0)) = 0 Then
    MsgBox "Nilai 'a' tidak boleh sama dengan '0'!", vbExclamation, "UltimaTech"
    Text1(0).SetFocus
    utHomEnd
    Exit Sub
End If
For I = 0 To 2
    Text2(I) = "<Unknown>"
Next I
On Error GoTo Keluar
    A = Val(Text1(0))
    B = Val(Text1(1))
    C = Val(Text1(2))
    P = Val(Text1(3))
    Q = Val(Text1(4))
    Dis = utD(A, B - P, C - Q)
    Text2(0) = Dis
    X1 = utAlpha(Dis, A, B - P)
    Y1 = Y(X1)
    X2 = utBetha(Dis, A, B - P)
    Y2 = Y(X2)
    Text2(1) = "( " & X1 & " , " & Y1 & " )"
    Text2(2) = "( " & X2 & " , " & Y2 & " )"
    utHomEnd
Keluar:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIForm1.MnuLinKu.Checked = False
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    utHomEnd
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack _
    Or KeyAscii = Asc("-")) Then
    KeyAscii = 0
    Beep
End If
End Sub

Private Sub Text1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Text1(Index).ToolTipText = Text1(Index).Text
End Sub

Private Sub Text2_GotFocus(Index As Integer)
    utHomEnd
End Sub

Private Sub Text2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Text2(Index).ToolTipText = Text2(Index).Text
End Sub
