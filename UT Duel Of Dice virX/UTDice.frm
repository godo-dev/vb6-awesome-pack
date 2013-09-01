VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "UltimaTech Duel of Dice VirX"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   Icon            =   "UTDice.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "UTDice.frx":0882
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picStart 
      BackColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   3413
      Picture         =   "UTDice.frx":593E1
      ScaleHeight     =   4155
      ScaleWidth      =   5115
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "x"
      Top             =   2880
      Width           =   5175
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00C0FFFF&
         Caption         =   "K&eluar...!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "x"
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton cmdStart 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Mulai...!"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "x"
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox txtPlayer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   720
         MaxLength       =   17
         TabIndex        =   6
         Tag             =   "x"
         Top             =   2640
         Width           =   3615
      End
      Begin VB.TextBox txtPlayer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   720
         MaxLength       =   17
         TabIndex        =   4
         Tag             =   "x"
         Top             =   1680
         Width           =   3615
      End
      Begin VB.OptionButton optPlayer 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Player vs. Player"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "x"
         Top             =   840
         Width           =   3615
      End
      Begin VB.OptionButton optPlayer 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Player vs. &Computer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   1
         Tag             =   "x"
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label lblPlayer 
         BackStyle       =   0  'Transparent
         Caption         =   "Player&2 Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   5
         Tag             =   "x"
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label lblPlayer 
         BackStyle       =   0  'Transparent
         Caption         =   "Player&1 Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   3
         Tag             =   "x"
         Top             =   1320
         Width           =   2055
      End
   End
   Begin VB.Timer tmrChar 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   2500
      Left            =   0
      Top             =   960
   End
   Begin VB.Timer tmrControl 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   0
      Top             =   480
   End
   Begin VB.CommandButton cmdRoll 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Roll!"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   9233
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2220
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer tmrRoller 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdRoll 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Roll!"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1553
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2220
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picChar 
      Height          =   2295
      Index           =   0
      Left            =   1553
      Picture         =   "UTDice.frx":6E96E
      ScaleHeight     =   2235
      ScaleWidth      =   2955
      TabIndex        =   14
      Top             =   5280
      Visible         =   0   'False
      Width           =   3015
      Begin VB.Image imgHeal 
         Height          =   810
         Index           =   0
         Left            =   0
         Picture         =   "UTDice.frx":79E87
         Tag             =   "x"
         Top             =   0
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Image imgSword 
         Height          =   1620
         Index           =   0
         Left            =   0
         Picture         =   "UTDice.frx":7A2D0
         Tag             =   "x"
         Top             =   0
         Visible         =   0   'False
         Width           =   1650
      End
   End
   Begin VB.PictureBox picChar 
      Height          =   2295
      Index           =   1
      Left            =   7440
      Picture         =   "UTDice.frx":7AC70
      ScaleHeight     =   2235
      ScaleWidth      =   2955
      TabIndex        =   15
      Top             =   5280
      Visible         =   0   'False
      Width           =   3015
      Begin VB.Image imgHeal 
         Height          =   810
         Index           =   1
         Left            =   0
         Picture         =   "UTDice.frx":80F4E
         Tag             =   "x"
         Top             =   0
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Image imgSword 
         Height          =   1620
         Index           =   1
         Left            =   0
         Picture         =   "UTDice.frx":81397
         Tag             =   "x"
         Top             =   0
         Visible         =   0   'False
         Width           =   1650
      End
   End
   Begin VB.CommandButton cmdBaru 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Mulai B&aru...!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5033
      Style           =   1  'Graphical
      TabIndex        =   11
      Tag             =   "x"
      Top             =   5280
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdUlang 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Ulangi Duel...!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   12
      Tag             =   "x"
      Top             =   6180
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0FFFF&
      Caption         =   "K&eluar...!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   13
      Tag             =   "x"
      Top             =   7080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblDamage 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   8512
      TabIndex        =   24
      Top             =   7680
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label lblDamage 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   2618
      TabIndex        =   23
      Top             =   7680
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label lblCommand 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Skill of the Dice"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   4208
      TabIndex        =   22
      Top             =   3720
      Visible         =   0   'False
      Width           =   3585
   End
   Begin VB.Label lblHp 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HP: 9999"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   9113
      TabIndex        =   21
      Top             =   4920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblHp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HP: 9999"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   1560
      TabIndex        =   20
      Top             =   4920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblChar 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Char2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   9593
      TabIndex        =   19
      Top             =   4560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblChar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Char1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   1560
      TabIndex        =   18
      Top             =   4560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblDice2 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   1815
      Left            =   7193
      TabIndex        =   17
      Top             =   1560
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblDice1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1815
      Left            =   2993
      TabIndex        =   16
      Top             =   1560
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type Karakter
    Nama As String
    HP As Double
End Type

Const MaxHP = 9999
Const xHP = "HP: "

Private Char(1) As Karakter

Dim Damage(1) As Integer
Dim Command As Integer
Dim Pilihan As VbMsgBoxResult
Dim I As Integer

Private Function Executing(Dmg1 As Integer, Dmg2 As Integer, HP1 As Double, HP2 As Double)
    Damage(0) = Dmg1
    Damage(1) = Dmg2
    HP1 = HP1 + Dmg1
    If HP1 > MaxHP Then
        HP1 = MaxHP
    ElseIf HP1 <= 0 Then
        HP1 = HP1 * 0
    End If
    HP2 = HP2 + Dmg2
    If HP2 > MaxHP Then
        HP2 = MaxHP
    ElseIf HP2 <= 0 Then
        HP2 = HP2 * 0
    End If
End Function

Private Function Inv(X As Integer) As Integer
    Inv = Abs(Not (-X))
End Function

Private Function Dice(DiceNum As Integer, HPx As Double, HPy As Double) As String
Select Case DiceNum
Case 11
    Dice = "Critical Bash!"
    Call Executing(0, -2000, HPx, HPy)
Case 12
    Dice = "Fire Attack!"
    Call Executing(0, -500, HPx, HPy)
Case 13
    Dice = "Thunder Storm!"
    Call Executing(0, -750, HPx, HPy)
Case 14
    Dice = "Somersault!"
    Call Executing(0, -400, HPx, HPy)
Case 15
    Dice = "Heal!"
    Call Executing(500, 0, HPx, HPy)
Case 16
    Dice = "Fire Ball!"
    Call Executing(0, -750, HPx, HPy)
Case 21
    Dice = "Meteor Crash!"
    Call Executing(-500, -1000, HPx, HPy)
Case 22
    Dice = "Curse of The Lord of Shadow!"
    Call Executing(-Round(0.1 * HPx), -Round(0.4 * HPy), HPx, HPy)
Case 23
    Dice = "Earthquake!"
    Call Executing(0, -750, HPx, HPy)
Case 24
    Dice = "Absorb HP!"
    Call Executing(Round(0.1 * HPy), -Round(0.1 * HPy), HPx, HPy)
Case 25
    Dice = "Aqua Breath!"
    Call Executing(0, -750, HPx, HPy)
Case 26
    Dice = "Limited Miracle Strikes!"
    Call Executing(-Round(0.1 * HPx), -Round(0.25 * HPy), HPx, HPy)
Case 31
    Dice = "Sky of Dust!"
    Call Executing(-500, -1000, HPx, HPy)
Case 32
    Dice = "Melody from the Hell!"
    Call Executing(0, -800, HPx, HPy)
Case 33
    Dice = "Miracle from The Blue Heaven!"
    Call Executing(MaxHP, 0, HPx, HPy)
Case 34
    Dice = "Hyper Shot!"
    Call Executing(0, -900, HPx, HPy)
Case 35
    Dice = "Unpowered Attack!"
    Call Executing(0, -50, HPx, HPy)
Case 36
    Dice = "Tornado Flame!"
    Call Executing(0, -1250, HPx, HPy)
Case 41
    Dice = "Prince of War Zone!"
    Call Executing(0, -1000, HPx, HPy)
Case 42
    Dice = "Attack with Broken Sword!"
    Call Executing(0, -250, HPx, HPy)
Case 43
    Dice = "Combination of Angel and Demon!"
    Call Executing(500, -600, HPx, HPy)
Case 44
    Dice = "Big Bang!"
    Call Executing(-Round(0.5 * HPx), -Round(0.75 * HPy), HPx, HPy)
Case 45
    Dice = "Cure!"
    Call Executing(1000, 0, HPx, HPy)
Case 46
    Dice = "The Reborn of Dark Creatures!"
    Call Executing(0, -1500, HPx, HPy)
Case 51
    Dice = "Drink a Bottle of Elixir!"
    Call Executing(1250, 0, HPx, HPy)
Case 52
    Dice = "Tribute for Destruction!"
    Call Executing(-1000, -1750, HPx, HPy)
Case 53
    Dice = "Light-Speed Attack!"
    Call Executing(0, -1200, HPx, HPy)
Case 54
    Dice = "Refresh the Energy!"
    Call Executing(1150, 0, HPx, HPy)
Case 55
    Dice = "Sound-Speed Attack with Odin's Sword!"
    Call Executing(0, -Round(0.5 * HPy), HPx, HPy)
Case 56
    Dice = "Exchange HP!"
    Call Executing(HPy - HPx, HPx - HPy, HPx, HPy)
Case 61
    Dice = "Arrows from the Sky!"
    Call Executing(0, -1000, HPx, HPy)
Case 62
    Dice = "Summon The Heaven Guardian!"
    Call Executing(0, -1500, HPx, HPy)
Case 63
    Dice = "Reborn!"
    Call Executing(0.5 * HPx, 0, HPx, HPy)
Case 64
    Dice = "Medicine from the Angel!"
    Call Executing(2000, 0, HPx, HPy)
Case 65
    Dice = "The Last Strike!"
    Call Executing(0, -1500, HPx, HPy)
Case 66
    Dice = "The End of The World!"
    Call Executing(-(HPx - 1), -(HPy - 1), HPx, HPy)
End Select
End Function

Private Sub cmdBaru_Click()
cmdBaru.Visible = False
cmdUlang.Visible = False
cmdExit(1).Visible = False
Dim Ctrl As Control
For Each Ctrl In Controls
    On Error Resume Next
    If Ctrl.Tag <> "x" Then Ctrl.Visible = False
Next Ctrl
picStart.Visible = True
txtPlayer(0).Text = ""
txtPlayer(1).Text = ""
optPlayer(0).Value = True
End Sub

Private Sub cmdExit_Click(Index As Integer)
    Pilihan = MsgBox("Apakah anda yakin ingin keluar?", vbQuestion + vbYesNo, "DoD virX")
    If Pilihan = vbYes Then End
End Sub

Private Sub cmdStart_Click()
Dim Ctrl As Control
    txtPlayer(0).Text = Trim(txtPlayer(0).Text)
    txtPlayer(1).Text = Trim(txtPlayer(1).Text)
If Trim(txtPlayer(0)) = "" Then
    MsgBox "Isi dulu donk Nama Player1-nya...!", vbExclamation, "DoD virX"
    txtPlayer(0).SetFocus
ElseIf Trim(txtPlayer(1)) = "" Then
    MsgBox "Isi dulu donk Nama Player2-nya...!", vbExclamation, "DoD virX"
    txtPlayer(1).SetFocus
Else
    picStart.Visible = False
    For I = 0 To 1
        imgSword(I).Move (0.5 * (picChar(I).Width - imgSword(I).Width)), _
        0.5 * (picChar(I).Height - imgSword(I).Height)
        imgHeal(I).Move (0.5 * (picChar(I).Width - imgHeal(I).Width)), _
        0.5 * (picChar(I).Height - imgHeal(I).Height)
        Char(I).HP = MaxHP
        lblHp(I).ForeColor = vbWhite
        Char(I).Nama = txtPlayer(I).Text
        lblChar(I).Caption = Char(I).Nama
        cmdRoll(I).Tag = ""
        lblHp(I).Caption = xHP & Char(I).HP
        lblDamage(I).Caption = ""
        cmdRoll(I).Caption = "&Roll!"
    Next I
    lblDice1.Caption = ""
    lblDice2.Caption = ""
    For Each Ctrl In Controls
        On Error Resume Next
        If Ctrl.Tag <> "x" Then Ctrl.Visible = True
    Next Ctrl
    lblCommand.Caption = Char(0).Nama & " Turn!"
    cmdRoll(0).Enabled = True
    If optPlayer(0).Value = True Then cmdRoll(1).Visible = False
    cmdBaru.Visible = False
    cmdUlang.Visible = False
    cmdExit(1).Visible = False
    cmdRoll(0).SetFocus
End If
End Sub

Private Sub cmdUlang_Click()
    Call cmdStart_Click
End Sub

Private Sub Form_Load()
If Screen.Height <> 600 * Screen.TwipsPerPixelY And Screen.Width <> 800 * Screen.TwipsPerPixelX Then
    Pilihan = MsgBox("Untuk tampilan yang lebih memuaskan gunakan resolusi monitor" & _
    " 800 x 600!" & Chr(13) & "Masih mau melanjutkan?", vbExclamation + vbYesNo, "DoD virX")
    If Pilihan = vbNo Then End
End If
    Randomize
End Sub

Private Sub lblDamage_Change(Index As Integer)
If lblDamage(Index).Caption = "0" Then
    lblDamage(Index).Caption = ""
ElseIf Val(lblDamage(Index).Caption) < 0 Then
    lblDamage(Index).ForeColor = vbRed
    imgSword(Index).Visible = True
ElseIf Val(lblDamage(Index).Caption) > 0 Then
    lblDamage(Index).Caption = "+" & Val(lblDamage(Index).Caption)
    lblDamage(Index).ForeColor = vbGreen
    imgHeal(Index).Visible = True
End If
End Sub

Private Sub lblHp_Change(Index As Integer)
If Char(Index).HP <= 1000 Then
    lblHp(Index).ForeColor = vbRed
ElseIf Char(Index).HP <= 2500 Then
    lblHp(Index).ForeColor = vbMagenta
Else
    lblHp(Index).ForeColor = vbWhite
End If
If Index = 1 Then
    If Char(1).HP <= 0 And Char(0).HP <= 0 Then
        lblDamage(0).ForeColor = vbGreen
        lblDamage(0) = "DRAW...!"
        lblDamage(1).ForeColor = vbGreen
        lblDamage(1) = "DRAW...!"
        GoTo Langsung
    ElseIf Char(1).HP <= 0 Then
        lblDamage(1).ForeColor = vbRed
        lblDamage(1).Caption = "LOSER...!"
        lblDamage(0).ForeColor = vbWhite
        lblDamage(0).Caption = "CHAMPION...!"
        GoTo Langsung
    End If
ElseIf Index = 0 And Char(0).HP <= 0 Then
    lblDamage(0).ForeColor = vbRed
    lblDamage(0).Caption = "LOSER...!"
    lblDamage(1).ForeColor = vbWhite
    lblDamage(1).Caption = "CHAMPION...!"
Langsung:
    lblCommand.Caption = "The End of DoD virX...!"
    lblDice1.Caption = ""
    lblDice2.Caption = ""
    tmrRoller.Enabled = False
    tmrChar(1).Enabled = False
    tmrControl.Enabled = False
    cmdRoll(0).Enabled = False
    cmdRoll(1).Enabled = False
    cmdBaru.Visible = True
    cmdUlang.Visible = True
    cmdExit(1).Visible = True
    cmdBaru.SetFocus
End If
End Sub

Private Sub cmdRoll_Click(Index As Integer)
For I = 0 To 1
    imgSword(I).Visible = False
    lblDamage(I).Caption = ""
    imgHeal(I).Visible = False
Next I
If cmdRoll(Index).Caption = "&Roll!" Then
    cmdRoll(Index).Caption = "&Stop!"
    tmrRoller.Enabled = True
ElseIf cmdRoll(Index).Caption = "&Stop!" Then
    cmdRoll(Index).Caption = "&Roll!"
    tmrRoller.Enabled = False
    lblCommand = Dice(Command, Char(Index).HP, Char(Inv(Index)).HP)
    lblDamage(Index).Caption = Damage(0)
    lblDamage(Inv(Index)).Caption = Damage(1)
    cmdRoll(Index).Enabled = False
    cmdRoll(Index).Tag = "q"
    tmrControl.Enabled = True
End If
End Sub

Private Sub tmrChar_Timer(Index As Integer)
    Call cmdRoll_Click(1)
    tmrChar(1).Enabled = False
End Sub

Private Sub tmrControl_Timer()
    lblDice1.Caption = ""
    lblDice2.Caption = ""
If cmdRoll(0).Tag = "q" Then
    cmdRoll(0).Tag = ""
    cmdRoll(1).Enabled = True
    lblCommand.Caption = Char(1).Nama & " Turn!"
    If optPlayer(0).Value = True Then
        Call cmdRoll_Click(1)
        tmrChar(1).Enabled = True
    ElseIf optPlayer(1).Value = True Then
        cmdRoll(1).SetFocus
    End If
ElseIf cmdRoll(1).Tag = "q" Then
    cmdRoll(1).Tag = ""
    cmdRoll(0).Enabled = True
    lblCommand.Caption = Char(0).Nama & " Turn!"
    cmdRoll(0).SetFocus
End If
For I = 0 To 1
    lblDamage(I).Caption = ""
    imgSword(I).Visible = False
    imgHeal(I).Visible = False
Next I
    lblHp(0) = xHP & Char(0).HP
    lblHp(1) = xHP & Char(1).HP
    tmrControl.Enabled = False
End Sub

Private Sub tmrRoller_Timer()
    Dim Dice1 As Integer, Dice2 As Integer
    Dice1 = Int(Rnd * 6) + 1
    Dice2 = Int(Rnd * 6) + 1
    lblDice1.Caption = Dice1
    lblDice2.Caption = Dice2
    Command = Dice1 * 10 + Dice2
End Sub

Private Sub txtPlayer_GotFocus(Index As Integer)
    SendKeys "{Home}+{End}"
End Sub
