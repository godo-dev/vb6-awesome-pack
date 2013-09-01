VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UltimaTech DaTetiMe"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   Icon            =   "Jam.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer utTmr 
      Interval        =   1
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox utPesan 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   -45
      ScaleHeight     =   195
      ScaleWidth      =   7005
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2520
      Width           =   7065
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Format &Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Index           =   1
      Left            =   3480
      TabIndex        =   12
      Tag             =   "x"
      Top             =   120
      Visible         =   0   'False
      Width           =   3375
      Begin VB.TextBox utText 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   15
         Top             =   750
         Width           =   495
      End
      Begin VB.TextBox utText 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   720
         MaxLength       =   2
         TabIndex        =   13
         Top             =   750
         Width           =   495
      End
      Begin VB.TextBox utText 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   17
         Top             =   750
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   1920
         TabIndex        =   16
         Top             =   750
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   1200
         TabIndex        =   14
         Top             =   750
         Width           =   255
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderWidth     =   3
         Height          =   495
         Left            =   600
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Format &Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Tag             =   "x"
      Top             =   120
      Visible         =   0   'False
      Width           =   3375
      Begin VB.ComboBox utCombo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   720
         Width           =   855
      End
      Begin VB.ComboBox utCombo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox utCombo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Th&n:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   2280
         TabIndex        =   10
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Bln:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   8
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "T&gl:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   375
      End
   End
   Begin VB.CommandButton utCmd 
      BackColor       =   &H00FFC0C0&
      Caption         =   "View Flash &Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   4500
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "x"
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton utCmd 
      BackColor       =   &H00FFC0C0&
      Caption         =   "View Flash &Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   2580
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton utCmd 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Format Date/Time"
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
      Height          =   495
      Index           =   0
      Left            =   660
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   135
      Left            =   120
      Tag             =   "x"
      Top             =   840
      Width           =   6735
   End
   Begin VB.Label utTimeNow 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   555
      Left            =   3345
      TabIndex        =   1
      Tag             =   "x"
      Top             =   1080
      Width           =   285
   End
   Begin VB.Label utDayNow 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   3345
      TabIndex        =   0
      Tag             =   "x"
      Top             =   120
      Width           =   285
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000C000&
      BorderWidth     =   10
      Height          =   2535
      Left            =   0
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim utBln(11)   As String
Dim utHari(6)   As String
Dim X           As SYSTEMTIME
Dim Nama        As String

Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Private Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Declare Function SetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME) As Long

Private Sub utHomEnd()
    SendKeys "{Home}+{End}"
End Sub

Private Sub Form_Initialize()
    Nama = "UltimaTech DaTetiMe...Created by: Junian Triajianto...Copyright © 2006...DILARANG KERAS MEMBAJAK PROGRAM INI!..."
    utBln(0) = "Januari"
    utBln(1) = "Februari"
    utBln(2) = "Maret"
    utBln(3) = "April"
    utBln(4) = "Mei"
    utBln(5) = "Juni"
    utBln(6) = "Juli"
    utBln(7) = "Agustus"
    utBln(8) = "September"
    utBln(9) = "Oktober"
    utBln(10) = "November"
    utBln(11) = "Desember"
    utHari(0) = "Minggu"
    utHari(1) = "Senin"
    utHari(2) = "Selasa"
    utHari(3) = "Rabu"
    utHari(4) = "Kamis"
    utHari(5) = "Jumat"
    utHari(6) = "Sabtu"
End Sub

Private Sub Form_Load()
Dim I As Integer
For I = 1 To 31
    utCombo(1).AddItem Format(I, "0#")
Next I
For I = 0 To 11
    utCombo(2).AddItem utBln(I)
    utCombo(2).ItemData(I) = I + 1
Next I
For I = 1980 To 2099
    utCombo(3).AddItem I
Next I
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim X As VbMsgBoxResult
Cancel = 1
X = MsgBox("Kamu Sudah Yakin 100% Nich Untuk Keluar?", vbQuestion + vbYesNo, "UltimaTech")
If X = vbYes Then
    End
End If
End Sub

Private Sub Timer1_Timer()
    utPesan.Cls
    utPesan.Print Nama
    Nama = Mid(Nama, 2, Len(Nama) - 1) & Mid(Nama, 1, 1)
End Sub

Private Sub utCmd_Click(Index As Integer)
Dim Ctrl    As Control
Dim Y       As SYSTEMTIME

Select Case Index

'/// utCmd(0)_Click() ///

    Case 0
    If utCmd(0).Caption = "&Format Date/Time" Then
        On Error Resume Next
        utCmd(0).Caption = "&Save Format"
        utCmd(1).Caption = "B&atal"
        Call utCmd_Click(1)
        utCombo(1).Text = utCombo(1).List(X.wDay - 1)
        utCombo(2).Text = utCombo(2).List(X.wMonth - 1)
        utCombo(3).Text = X.wYear
        utText(0).Text = Format(X.wHour, "0#")
        utText(1).Text = Format(X.wMinute, "0#")
        utText(2).Text = Format(X.wSecond, "0#")
    ElseIf utCmd(0).Caption = "&Save Format" Then
        With Y
            .wDay = Val(utCombo(1).Text)
            .wMonth = utCombo(2).ItemData(utCombo(2).ListIndex)
            .wYear = Val(utCombo(3).Text)
            .wHour = Val(utText(0).Text)
            .wMinute = Val(utText(1).Text)
            .wSecond = Val(utText(2).Text)
        End With
        utTmr.Enabled = False
        Call SetLocalTime(Y)
        utTmr.Enabled = True
        Call utCmd_Click(1)
    End If
    
'/// utCmd(1)_Click() ///

    Case 1
    If utCmd(1).Caption = "View Flash &Time" Then
        On Error GoTo Keluar
        Shell ("Time.exe")
    ElseIf utCmd(1).Caption = "B&atal" Then
        For Each Ctrl In Controls
            If Ctrl.Tag = "x" Then
                Ctrl.Visible = Not (Ctrl.Visible)
            End If
        Next Ctrl
    End If
    If Frame1(0).Visible = False Then
        utCmd(1).Caption = "View Flash &Time"
        utCmd(0).Caption = "&Format Date/Time"
    End If
    
'/// utCmd(2)_Click() ///

    Case 2
    On Error GoTo Keluar
    Shell ("Date.exe")
End Select
Exit Sub
Keluar:
    MsgBox "Maaf! Program ini masih belum ter-Install!", vbExclamation, "UltimaTech"
End Sub

Private Sub utCombo_Click(Index As Integer)
If Index = 2 Or Index = 3 Then
Dim Thn As Integer
Dim Bln As Byte
Dim Hr  As Byte
Dim I   As Integer, II As Integer
        If utCombo(2).Text = "" Then Exit Sub
        Hr = Val(utCombo(1).Text)
        Thn = Val(utCombo(3).Text)
        Bln = utCombo(2).ItemData(utCombo(2).ListIndex)
        utCombo(1).Clear
        For I = 1 To 7 Step 2
            If Bln = I Then
                For II = 1 To 31
                    utCombo(1).AddItem Format(II, "0#")
                Next II
            End If
        Next I
        For I = 8 To 12 Step 2
            If Bln = I Then
                For II = 1 To 31
                    utCombo(1).AddItem Format(II, "0#")
                Next II
            End If
        Next I
        If Bln = 4 Or Bln = 6 Or Bln = 9 Or Bln = 11 Then
            For II = 1 To 30
                utCombo(1).AddItem Format(II, "0#")
            Next II
        End If
        If Bln = 2 And Thn Mod 4 = 0 Then
            For II = 1 To 29
                utCombo(1).AddItem Format(II, "0#")
            Next II
        ElseIf Bln = 2 And Thn Mod 4 <> 0 Then
            For II = 1 To 28
                utCombo(1).AddItem Format(II, "0#")
            Next II
        End If
        If Hr > II - 1 Then Hr = II - 1
        If Hr <> 0 Then utCombo(1).Text = Format(Hr, "0#")
End If
End Sub

Private Sub utText_GotFocus(Index As Integer)
    utHomEnd
End Sub

Private Sub utText_KeyPress(Index As Integer, KeyAscii As Integer)
If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then
    Beep
    KeyAscii = 0
End If
End Sub

Private Sub utTimeNow_Change()
If utDayNow.ForeColor = &HFF0000 Then
    utDayNow.ForeColor = &HFF&
    utTimeNow.ForeColor = &HFF&
ElseIf utDayNow.ForeColor = &HFF& Then
    utDayNow.ForeColor = 16711935
    utTimeNow.ForeColor = 16711935
Else
    utDayNow.ForeColor = &HFF0000
    utTimeNow.ForeColor = &HFF0000
End If
End Sub

Private Sub utTmr_Timer()
Call GetLocalTime(X)
utDayNow.Caption = utHari(X.wDayOfWeek) & ", " & X.wDay & " " _
    & utBln(X.wMonth - 1) & " " & X.wYear
utTimeNow.Caption = Format(X.wHour, "0#:") & Format(X.wMinute, "0#:") _
    & Format(X.wSecond, "0#")
End Sub
