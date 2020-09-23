Attribute VB_Name = "ArcSine"
  Option Explicit

' ----------------------------------------------------------------
' Define arc sine function to return degree or radian values.
' The default is radians unless degrees is specifically indicated.

  Public Function ArcSin(ArgX, Deg_or_Rad)

  Dim Q As Double

  If Abs(ArgX) = 1 Then
     Q = 2 * Atn(1) * ArgX
  Else
     Q = Atn(ArgX / Sqr(-ArgX * ArgX + 1))
  End If
  If Left(UCase(Deg_or_Rad), 1) = "D" Then Q = Q * 45 / Atn(1)
     ArcSin = Q

  End Function

