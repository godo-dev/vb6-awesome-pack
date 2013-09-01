VERSION 5.00
Begin VB.Form Form7 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Persamaan Kuadrat"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   4470
   DrawStyle       =   6  'Inside Solid
   Icon            =   "PersKuadrat.frx":0000
   LinkTopic       =   "Form7"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "PersKuadrat.frx":030A
   ScaleHeight     =   298
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   298
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   1208
      TabIndex        =   1
      Text            =   "1"
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   2408
      TabIndex        =   5
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   3488
      TabIndex        =   8
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   2168
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   2168
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   2168
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   2168
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   2168
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   5
      Left            =   2168
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   6
      Left            =   2168
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1935
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
      Left            =   1628
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Titik Potong Sb. X: "
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
      Index           =   4
      Left            =   360
      TabIndex        =   23
      Top             =   3240
      Visible         =   0   'False
      Width           =   1695
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
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   20
      Left            =   1808
      TabIndex        =   2
      Top             =   360
      Width           =   255
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
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   21
      Left            =   2048
      TabIndex        =   3
      Top             =   240
      Width           =   135
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
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   22
      Left            =   2168
      TabIndex        =   4
      Top             =   360
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
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   23
      Left            =   3008
      TabIndex        =   6
      Top             =   360
      Width           =   255
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
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   24
      Left            =   3248
      TabIndex        =   7
      Top             =   360
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F(X)  ="
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
      Height          =   255
      Index           =   25
      Left            =   488
      TabIndex        =   0
      Top             =   360
      Width           =   615
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      DrawMode        =   6  'Mask Pen Not
      Height          =   735
      Left            =   188
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pers. Sumbu Simetri: "
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
      Index           =   1
      Left            =   375
      TabIndex        =   12
      Top             =   2160
      Width           =   1830
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nilai Extrim: "
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
      Index           =   2
      Left            =   375
      TabIndex        =   14
      Top             =   2520
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Titik Potong Sb. X: "
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
      Index           =   3
      Left            =   375
      TabIndex        =   16
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Titik Potong Sb. Y: "
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
      Index           =   5
      Left            =   375
      TabIndex        =   19
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Koord. Titik Puncak: "
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
      Index           =   6
      Left            =   375
      TabIndex        =   21
      Top             =   3960
      Width           =   1830
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
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   368
      TabIndex        =   10
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      DrawMode        =   6  'Mask Pen Not
      Height          =   2655
      Left            =   128
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   4215
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim A, B, C As Double

Private Function utF(ByVal X As Double, ByVal A As Double, _
    ByVal B As Double, ByVal C As Double) As Double
    utF = (A * X ^ 2) + (B * X) + C
End Function

Private Function utX_Puncak(ByVal A As Double, ByVal B As Double _
    ) As Double
    utX_Puncak = B / -(2 * A)
End Function

Private Function utY_Puncak(ByVal utD As Double, ByVal A As Double) _
    As Double
    utY_Puncak = utD / -(4 * A)
End Function

Private Sub Command1_Click()
Dim I As Byte
For I = 0 To 2
    Text1(I) = Val(Text1(I))
Next I
If Val(Text1(0)) = 0 Then
    MsgBox "Nilai 'a' tidak boleh sama dengan '0'!!!", vbExclamation, "UltimaTech"
    Text1(0).SetFocus
    utHomEnd
    Exit Sub
End If
    utHomEnd
For I = 0 To 6
    Text2(I) = "<Unknown>"
Next I
On Error GoTo Keluar
    A = Val(Text1(0))
    B = Val(Text1(1))
    C = Val(Text1(2))
    Text2(0) = utD(A, B, C)
    Text2(1) = utX_Puncak(A, B)
    Text2(2) = utY_Puncak(utD(A, B, C), A)
    Text2(3) = "( " & utAlpha(utD(A, B, C), A, B) & " , 0 )"
    Text2(4) = "( " & utBetha(utD(A, B, C), A, B) & " , 0 )"
    Text2(5) = "( 0 , " & utF(0, A, B, C) & " )"
    Text2(6) = "( " & Text2(1) & " , " & Text2(2) & " )"
Keluar:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIForm1.MnuPers.Checked = False
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    utHomEnd
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or _
    KeyAscii = vbKeyBack Or KeyAscii = Asc("-")) Then
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
    Text2(Index).ToolTipText = Label1(Index).Caption & Text2(Index).Text
End Sub
