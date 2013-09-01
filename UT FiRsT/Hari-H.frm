VERSION 5.00
Begin VB.Form Form8 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nama Hari"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   4080
   Icon            =   "Hari-H.frx":0000
   LinkTopic       =   "Form8"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "Hari-H.frx":030A
   ScaleHeight     =   139
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   272
   Begin VB.TextBox TxtTahun 
      Height          =   315
      Left            =   2753
      MaxLength       =   4
      TabIndex        =   5
      Top             =   360
      Width           =   1215
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
      Left            =   1433
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   773
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1560
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   1
      Left            =   1433
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   113
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pada Saat Itu Adalah Hari:"
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
      Left            =   878
      TabIndex        =   7
      Top             =   1320
      Width           =   2325
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ta&hun:"
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
      Height          =   255
      Index           =   2
      Left            =   2753
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Bulan:"
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
      Height          =   255
      Index           =   1
      Left            =   1433
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Tanggal:"
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
      Height          =   255
      Index           =   0
      Left            =   113
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Bulan(13)   As String
Dim Hari(7)     As String
Dim JmlBln(13)  As Byte
Dim Bln         As Byte
Dim Hr          As Byte
Dim Thn         As Integer

Private Sub Form_Activate()
    Combo1(0).Text = Format(Day(Now), "0#")
    Combo1(1).Text = Bulan(Month(Now))
    TxtTahun.Text = Year(Now)
    Command1_Click
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Text1.ToolTipText = Text1.Text
End Sub

Private Sub TxtTahun_Change()
    Thn = Val(TxtTahun.Text)
    Call Combo1_Click(1)
End Sub

Private Sub Combo1_Click(Index As Integer)
Dim I, II   As Byte
Dim HrMax   As Byte
Select Case Index
    Case 0
        Hr = Val(Combo1(0).Text)
    Case 1
        Thn = Val(TxtTahun.Text)
        If Combo1(1).Text = "" Then
            Exit Sub
        End If
        Bln = Combo1(1).ItemData(Combo1(1).ListIndex)
        Combo1(0).Clear
        For I = 1 To 7 Step 2
            If Bln = I Then
                For II = 1 To 31
                    Combo1(0).AddItem Format(II, "0#")
                    HrMax = II
                Next II
            End If
        Next I
        For I = 8 To 12 Step 2
            If Bln = I Then
                For II = 1 To 31
                    Combo1(0).AddItem Format(II, "0#")
                    HrMax = II
                Next II
            End If
        Next I
        If Bln = 4 Or Bln = 6 Or Bln = 9 Or Bln = 11 Then
            For I = 1 To 30
                Combo1(0).AddItem Format(I, "0#")
                HrMax = I
            Next I
        End If
        If Bln = 2 And Thn Mod 4 = 0 Then
            For I = 1 To 29
                Combo1(0).AddItem Format(I, "0#")
                HrMax = I
            Next I
        ElseIf Bln = 2 And Thn Mod 4 <> 0 Then
            For I = 1 To 28
                Combo1(0).AddItem Format(I, "0#")
                HrMax = I
            Next I
        End If
        If Hr > HrMax Then
            Hr = HrMax
        End If
        If Hr <> 0 Then
            Combo1(0).Text = Format(Hr, "0#")
        End If
End Select
    Combo1(Index).ToolTipText = Combo1(Index).Text
End Sub

Private Sub TxtTahun_Click()
    Thn = Val(TxtTahun.Text)
    TxtTahun.Text = Format(Thn, "000#")
    Combo1_Click (1)
End Sub

Private Sub TxtTahun_GotFocus()
    utHomEnd
End Sub

Private Sub TxtTahun_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc(0) And KeyAscii <= Asc("9") Or _
    KeyAscii = vbKeyBack) Then
    KeyAscii = 0
    Beep
End If
End Sub

Private Sub TxtTahun_LostFocus()
    TxtTahun.Text = Format(Thn, "000#")
End Sub

Private Sub Command1_Click()
Dim NamaHari        As Double
Dim Abad            As Long
Dim HAbad           As Long
Dim H, B            As Byte
Dim T               As Double
On Error GoTo Keluar
    TxtTahun.Text = Format(Thn, "000#")
If Hr = 0 Then
    MsgBox "Anda harus memasukkan tanggal yang anda tuju!", vbExclamation, "UltimaTech"
    Combo1(0).SetFocus
ElseIf Bln = 0 Then
    MsgBox "Anda harus memasukkan bulan yang anda tuju!", vbExclamation, "UltimaTech"
    Combo1(1).SetFocus
ElseIf Thn = 0 Then
    MsgBox "Tahun tidak boleh sama dengan '0'!", vbExclamation, "UltimaTech"
    TxtTahun.SetFocus
    utHomEnd
End If
HAbad = 36525
T = Thn - 1
Abad = T \ 100
T = T Mod 100
B = Bln - 1
H = Hr - 1
NamaHari = ((Abad * HAbad) + ((T * 365) + (T \ 4)) + JumlahBulan(B) + H) Mod 7
Text1.Text = Hari(NamaHari)
Keluar:
End Sub

Private Sub Form_Initialize()
    Bulan(1) = "Januari"
    Bulan(2) = "Februari"
    Bulan(3) = "Maret"
    Bulan(4) = "April"
    Bulan(5) = "Mei"
    Bulan(6) = "Juni"
    Bulan(7) = "Juli"
    Bulan(8) = "Agustus"
    Bulan(9) = "September"
    Bulan(10) = "Oktober"
    Bulan(11) = "November"
    Bulan(12) = "Desember"
    Hari(0) = "Minggu"
    Hari(1) = "Senin"
    Hari(2) = "Selasa"
    Hari(3) = "Rabu"
    Hari(4) = "Kamis"
    Hari(5) = "Jum'at"
    Hari(6) = "Sabtu"
End Sub

Private Sub Form_Load()
Dim I As Integer
For I = 1 To 12
    Combo1(1).AddItem Bulan(I)
    Combo1(1).ItemData(I - 1) = I
Next I
For I = 1 To 31
    Combo1(0).AddItem Format(I, "0#")
Next I
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIForm1.MnuHari.Checked = False
End Sub

Private Function JumlahBulan(ByVal Input_Bulan As Byte) As Integer
Dim I As Byte
For I = 1 To 7 Step 2
    JmlBln(I) = 31
Next I
For I = 8 To 12 Step 2
    JmlBln(I) = 31
Next I
    JmlBln(4) = 30
    JmlBln(6) = 30
    JmlBln(9) = 30
    JmlBln(11) = 30
If Thn Mod 4 = 0 Then
    JmlBln(2) = 29
ElseIf Thn Mod 4 <> 0 Then
    JmlBln(2) = 28
End If
For I = 1 To Input_Bulan
    JumlahBulan = JumlahBulan + JmlBln(I)
Next I
End Function

Private Sub Text1_GotFocus()
    utHomEnd
End Sub

Private Sub TxtTahun_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    TxtTahun.ToolTipText = TxtTahun.Text
End Sub
