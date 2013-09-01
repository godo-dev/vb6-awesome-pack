VERSION 5.00
Begin VB.Form Form11 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convert Measurement"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   Icon            =   "ConMeasurement.frx":0000
   LinkTopic       =   "Form11"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "ConMeasurement.frx":0442
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
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
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2520
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1800
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   2400
      TabIndex        =   6
      Top             =   480
      Width           =   2655
   End
   Begin VB.OptionButton Opt1 
      BackColor       =   &H0000FF00&
      Caption         =   "&Twip"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Tag             =   "Twip"
      Top             =   1560
      Width           =   855
   End
   Begin VB.OptionButton Opt1 
      BackColor       =   &H0000FF00&
      Caption         =   "&Inch"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Tag             =   "Inch"
      Top             =   1200
      Width           =   855
   End
   Begin VB.OptionButton Opt1 
      BackColor       =   &H0000FF00&
      Caption         =   "&Pixel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Tag             =   "Pixel"
      Top             =   840
      Width           =   855
   End
   Begin VB.OptionButton Opt1 
      BackColor       =   &H0000FF00&
      Caption         =   "&Centimeter (cm)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Tag             =   "Cm"
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nilai Ini Sama Dengan:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   2
      Left            =   2400
      TabIndex        =   9
      Top             =   1440
      Width           =   1965
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Masukkan Nilai Ukuran:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Index           =   1
      Left            =   2400
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
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
      Index           =   3
      Left            =   5160
      TabIndex        =   15
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
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
      Index           =   2
      Left            =   5160
      TabIndex        =   13
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
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
      Index           =   1
      Left            =   5160
      TabIndex        =   11
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   5160
      TabIndex        =   7
      Top             =   480
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   1575
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pilih Besaran Ukuran:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      FillColor       =   &H00FF0000&
      Height          =   495
      Left            =   2280
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Width           =   3615
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   1215
      Left            =   2280
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   3615
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim I As Byte
Dim X, Y, Z, Q As Double
For I = 1 To 3
    Text1(I) = "<Unknown>"
Next I
On Error GoTo Keluar
    X = Val(Text1(0))
If Opt1(0).Value = True Then
    Y = Round(100 / 1.27 * X, 0)
    Z = Round(1 / 2.54 * X, 3)
    Q = Round(15 * Y, 0)
ElseIf Opt1(1).Value = True Then
    X = Int(X)
    Text1(0) = X
    Y = Round(1.27 / 100 * X, 2)
    Z = Round(1 / 2.54 * Y, 3)
    Q = Round(15 * X, 0)
ElseIf Opt1(2).Value = True Then
    Y = Round(2.54 * X, 2)
    Z = Round(100 / 1.27 * Y, 0)
    Q = Round(15 * Z, 0)
ElseIf Opt1(3).Value = True Then
    X = Int(X)
    Text1(0) = X
    Z = Round(X / 15, 0)
    Y = Round(1.27 / 100 * Z, 2)
    Q = Round(1 / 2.54 * Y, 3)
End If
    Text1(1) = Y
    Text1(2) = Z
    Text1(3) = Q
Keluar:
    Text1(0).SetFocus
    utHomEnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIForm1.MnuConMea.Checked = False
End Sub

Private Sub Lbl1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Lbl1(Index).ToolTipText = Lbl1(Index).Caption
End Sub

Private Sub Opt1_Click(Index As Integer)
Dim I As Byte
Lbl1(0).Caption = Opt1(Index).Tag
Select Case Index
    Case 0
        For I = 1 To 3
            Lbl1(I).Caption = Opt1(I).Tag
        Next I
    Case 1
        Lbl1(1).Caption = Opt1(0).Tag
        Lbl1(2).Caption = Opt1(2).Tag
        Lbl1(3).Caption = Opt1(3).Tag
    Case 2
        Lbl1(1).Caption = Opt1(0).Tag
        Lbl1(2).Caption = Opt1(1).Tag
        Lbl1(3).Caption = Opt1(3).Tag
    Case 3
        For I = 0 To 2
            Lbl1(I + 1).Caption = Opt1(I).Tag
        Next I
End Select
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    utHomEnd
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or _
    KeyAscii = vbKeyBack Or KeyAscii = Asc(".")) Then
    KeyAscii = 0
    Beep
End If
End Sub

Private Sub Text1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Text1(Index).ToolTipText = Text1(Index).Text
End Sub
