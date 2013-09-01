VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00000000&
   Caption         =   "UltimaTech FiRsT"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   11880
   Icon            =   "UltimaTech.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7860
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   6800
            TextSave        =   "06/05/2006"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   6800
            TextSave        =   "5:46"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   6800
            Text            =   "Created by: Junian Triajianto"
            TextSave        =   "Created by: Junian Triajianto"
            Object.ToolTipText     =   "X-5 / 17"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu MnuMenu 
      Caption         =   "&Menu Utama"
      Begin VB.Menu MnuKeluar 
         Caption         =   "K&eluar"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu MnuProPro 
      Caption         =   "&FiRsT Programs"
      Begin VB.Menu MnuPro 
         Caption         =   "Program M&atematika"
         Begin VB.Menu MnuProCon 
            Caption         =   "Program &Converting"
            Begin VB.Menu MnuDec 
               Caption         =   "Convert &Decimal"
            End
            Begin VB.Menu MnuBiner 
               Caption         =   "Convert &Binary"
            End
            Begin VB.Menu MnuHex 
               Caption         =   "Convert He&xaDecimal"
            End
         End
         Begin VB.Menu MnuFungsi 
            Caption         =   "F&ungsi"
            Begin VB.Menu MnuPers 
               Caption         =   "Persamaan &Kuadrat"
            End
            Begin VB.Menu MnuLinKu 
               Caption         =   "&Persamaan Linear dan Kuadrat"
            End
         End
         Begin VB.Menu MnuLain2 
            Caption         =   "&Lain-Lain"
            Begin VB.Menu MnuTri 
               Caption         =   "&Trigonometri"
            End
            Begin VB.Menu MnuFak 
               Caption         =   "&Faktorial"
            End
            Begin VB.Menu MnuHari 
               Caption         =   "Nama &Hari"
            End
         End
      End
      Begin VB.Menu MnuHibur 
         Caption         =   "Program &Hiburan"
         Begin VB.Menu MnuGame 
            Caption         =   "&Game"
            Begin VB.Menu Mnu7 
               Caption         =   "Lucky &7 Game"
            End
         End
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbPopupMenuRightButton Then
    Me.PopupMenu MnuProPro
End If
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    Call MnuKeluar_Click
End Sub

Private Sub MDIForm_Resize()
If Me.WindowState = 0 Then
    Me.WindowState = 2
End If
End Sub

Private Sub Mnu7_Click()
    Call utMenuSub(Mnu7, Form5)
End Sub

Private Sub MnuDec_Click()
    Call utMenuSub(MnuDec, Form1)
End Sub

Private Sub MnuBiner_Click()
    Call utMenuSub(MnuBiner, Form2)
End Sub

Private Sub MnuFak_Click()
    Call utMenuSub(MnuFak, Form4)
End Sub

Private Sub MnuHari_Click()
    Call utMenuSub(MnuHari, Form8)
End Sub

Private Sub MnuHex_Click()
    Call utMenuSub(MnuHex, Form6)
End Sub

Private Sub MnuLinKu_Click()
    Call utMenuSub(MnuLinKu, Form9)
End Sub

Private Sub MnuPers_Click()
    Call utMenuSub(MnuPers, Form7)
End Sub

Private Sub MnuTri_Click()
    Call utMenuSub(MnuTri, Form3)
End Sub

Private Sub MnuKeluar_Click()
Dim X As VbMsgBoxResult
    X = MsgBox("Yakin nich pengen keluar?", vbQuestion + vbYesNo, "UltimaTech")
If X = vbYes Then
    End
End If
End Sub

