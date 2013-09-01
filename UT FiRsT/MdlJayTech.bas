Attribute VB_Name = "Module1"
Option Explicit

Public Sub utHomEnd()
    SendKeys "{Home}+{End}"
End Sub

Public Sub utMenuSub(Ax As Menu, Frm As Form)
On Error GoTo Keluar
If Ax.Checked = False Then
    Ax.Checked = True
    Load Frm
    Frm.Show
Else
    Ax.Checked = False
    Unload Frm
End If
Keluar:
End Sub

Public Function utBin(ByVal Nil As Double) As String
Dim S As String
Dim B As Byte
If Nil > 1 Then
    Do
        B = Nil Mod 2
        S = B & S
        Nil = (Nil - B) / 2
    Loop Until Nil = 1
End If
    S = Nil & S
    utBin = S
End Function

Public Function utD(ByVal A As Double, ByVal B As Double, _
    ByVal C As Double) As Double
    utD = (B ^ 2) - (4 * A * C)
End Function

Public Function utAlpha(ByVal utD As Double, ByVal A As Double, _
    ByVal B As Double) As Double
    utAlpha = (-B + (utD ^ 0.5)) / (2 * A)
End Function

Public Function utBetha(ByVal utD As Double, ByVal A As Double, _
    ByVal B As Double) As Double
    utBetha = (-B - (utD ^ 0.5)) / (2 * A)
End Function

Public Function utSudut(ByVal X As Double) As Double
    utSudut = X * (22 / 7 / 180)
End Function

Public Function utSin(ByVal X As Double) As Double
    utSin = Round(Sin(utSudut(X)), 2)
End Function

Public Function utCos(ByVal X As Double) As Double
    utCos = Round(Cos(utSudut(X)), 2)
End Function

Public Function utTan(ByVal X As Double) As Double
    utTan = Round(Tan(utSudut(X)), 2)
End Function

Public Function utCosec(ByVal X As Double) As Double
    utCosec = Round(1 / Sin(utSudut(X)), 2)
End Function

Public Function utSec(ByVal X As Double) As Double
    utSec = Round(1 / Cos(utSudut(X)), 2)
End Function

Public Function utCtg(ByVal X As Double) As Double
    utCtg = Round(1 / Tan(utSudut(X)), 2)
End Function
