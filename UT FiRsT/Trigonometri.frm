VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trigonometri"
   ClientHeight    =   2490
   ClientLeft      =   315
   ClientTop       =   375
   ClientWidth     =   4755
   Icon            =   "Trigonometri.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "Trigonometri.frx":030A
   ScaleHeight     =   166
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   317
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   5
      Left            =   3030
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1999
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   4
      Left            =   3030
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1639
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   3
      Left            =   3030
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1279
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   2
      Left            =   870
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1999
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   1
      Left            =   870
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1639
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Index           =   0
      Left            =   870
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1279
      Width           =   1335
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
      Left            =   1770
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   799
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   390
      TabIndex        =   1
      Top             =   439
      Width           =   3975
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ctg:"
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
      Left            =   2430
      TabIndex        =   13
      Top             =   2006
      Width           =   360
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sec:"
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
      Left            =   2430
      TabIndex        =   11
      Top             =   1646
      Width           =   405
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cosec: "
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
      Left            =   2430
      TabIndex        =   9
      Top             =   1286
      Width           =   660
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Tan: "
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
      Height          =   255
      Left            =   390
      TabIndex        =   7
      Top             =   1999
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cos: "
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
      Height          =   255
      Left            =   390
      TabIndex        =   5
      Top             =   1639
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sin: "
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
      Left            =   390
      TabIndex        =   3
      Top             =   1286
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Masukkan Suatu Besaran Sudut:"
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
      Left            =   390
      TabIndex        =   0
      Top             =   206
      Width           =   2805
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim X As Double
Dim I As Byte
If Text1 <> "" Then
    On Error Resume Next
For I = 0 To 5
    Text2(I) = "<Unknown>"
Next I
    X = Val(Text1)
    Text1 = X
    Text2(0) = utSin(X)
    Text2(1) = utCos(X)
    Text2(2) = utTan(X)
    Text2(3) = utCosec(X)
    Text2(4) = utSec(X)
    Text2(5) = utCtg(X)
    Text1.SetFocus
    utHomEnd
Else
    For I = 0 To 5
        Text2(I) = ""
    Next I
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIForm1.MnuTri.Checked = False
End Sub

Private Sub Text1_GotFocus()
    utHomEnd
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = Asc(".") _
    Or KeyAscii = Asc("-")) Then
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
