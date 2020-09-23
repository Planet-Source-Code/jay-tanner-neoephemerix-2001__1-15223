Attribute VB_Name = "Sidereal_Time"
Option Explicit

' This module contains a function to compute the mean or apparent
' sidereal time.

' ==============================================================================
' Compute the local mean or true apparent sidereal time angle in degrees for JDE
' value on a given date at any given longitude.
'
' The historical astronomical convention is:
' East = (0 to -180) and West = (0 to +180)

  Public Function LST_For _
  (Dd_Mmm_Yyyy_BCAD, UT_HHMMSS, Lng, Apparent As Boolean)

  Dim Q   As Variant
  Dim JDE As Double
  Dim T   As Double
  Dim ST  As Double
  Dim Obl As Double
  
      
  Q = JDE_For(Dd_Mmm_Yyyy_BCAD, 0)
  If Error_In(Q) Then LST_For = "ERROR: Invalid calendar date": Exit Function

  JDE = Q

    T = (JDE - 2451545) / 36525
  
  Obl = Ecliptic_Obliquity(JDE, "Apparent")
  
   ST = 100.46061837 + 36000.770053608 * T _
      + 0.000387933 * T * T _
      + T * T * T / 38710000
      
' Compute mean ST at specified longitude and UT
  ST = ST + Day_Frac_Equiv_To(UT_HHMMSS) * 360.985647366 - Lng
  
' Correct for nutation if true ST mode indicated, otherwise use mean value
  If Apparent Then ST = ST + Delta_Psi(JDE + Day_Frac_Equiv_To(UT_HHMMSS)) _
         * Cos(Obl * Atn(1) / 45)
      
' Modulate sidereal time angle to fall between 0 and 360 degrees
  If Abs(ST) > 360 Then ST = ST - 360 * Int(ST / 360)
  If ST < 0 Then ST = ST + 360

' Return computed ST value
  LST_For = ST

  End Function

