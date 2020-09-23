Attribute VB_Name = "Nutation_In_Obliquity"
  Option Explicit

' ====================================================================

' This function computes the nutation in obliquity of the ecliptic.
' It is also included within the Ecliptic_Obliquity() function.

' Compute nutational correction for obliquity of the ecliptic in degrees.
' In terms of accuracy, the value in arc seconds is to about ±0.001"

' This correction is applied to the mean obliquity to obtain the true or
' apparent obliquity at any given moment.

' This computation is based on the "1980 IAU Theory of Nutation" and
' includes only the terms with coefficients > 0.0003 arcsecond.

  Public Function Delta_e(At_JDE)
' LEVEL 0

  Dim JD As String ' JD number for date and time
  Dim T  As Double ' Julian centuries since J2000.0
  Dim T2 As Double ' T to the power of 2
  Dim T3 As Double ' T to the power of 3
  
  Dim Q  As Double  ' Nutation series accumulator
  
  Dim V  As Double  ' Mean elongation of the moon from the sun
  Dim W  As Double  ' Mean anomaly of the sun
  Dim x  As Double  ' Mean anomaly of the moon
  Dim y  As Double  ' Moon's argument of latitude
  
' Longitude of ascending node of lunar orbit on the ecliptic
' as measured from the mean equinox of date.
  Dim Z  As Double
    
  T = (At_JDE - 2451545#) / 36525
   
' Compute the mean elongation of the moon in radians
  V = 297.85036 + 445267.11148 * T - 0.0019142 * T2 + T3 / 189474
  V = V * Atn(1) / 45
  
' Compute the mean anomaly of the sun in radians
  W = 357.52772 + 35999.05034 * T - 0.0001603 * T2 - T3 / 300000
  W = W * Atn(1) / 45
  
' Compute the mean anomaly of the moon in radians
  x = 134.96298 + 477198.867398 * T + 0.0086972 * T2 + T3 / 56250
  x = x * Atn(1) / 45
  
' Compute the moon's argument of latitude in radians
  y = 93.27191 + 483202.017538 * T - 0.0036825 * T2 + T3 / 327270
  y = y * Atn(1) / 45
  
' Compute the longitude of moon's ascending node in radians
  Z = 125.04452 - 1934.136261 * T + 0.0020708 * T2 + T3 / 450000
  Z = Z * Atn(1) / 45
  
' Proceed to compute the nutation in obliquity
  Q = Cos(Z) * (92025 + 8.9 * T)
  Q = Q + Cos(2 * (y - V + Z)) * (5736 - 3.1 * T)
  Q = Q + Cos(2 * (y + Z)) * (977 - 0.5 * T)
  Q = Q + Cos(2 * Z) * (0.5 * T - 895)
  Q = Q + Cos(W) * (54 - 0.1 * T)
  Q = Q - 7 * Cos(x)
  Q = Q + Cos(W + 2 * (y - V + Z)) * (224 - 0.6 * T)
  Q = Q + 200 * Cos(2 * y + Z)
  Q = Q + Cos(x + 2 * (y + Z)) * (129 - 0.1 * T)
  Q = Q + Cos(2 * (y - V + Z) - W) * (0.3 * T - 95)
  Q = Q - 70 * Cos(2 * (y - V) + Z)
  Q = Q - 53 * Cos(2 * (y + Z) - x)
  Q = Q - 33 * Cos(x + Z)
  Q = Q + 26 * Cos(2 * (V + y + Z) - x)
  Q = Q + 32 * Cos(Z - x)
  Q = Q + 27 * Cos(x + 2 * y + Z)
  Q = Q - 24 * Cos(2 * (y - x) + Z)
  Q = Q + 16 * Cos(2 * (V + y + Z))
  Q = Q + 13 * Cos(2 * (x + y + Z))
  Q = Q - 12 * Cos(x + 2 * (y - V + Z))
  Q = Q - 10 * Cos(2 * y + Z - x)
  Q = Q - 8 * Cos(2 * V - x + Z)
  Q = Q + 7 * Cos(2 * (W - V + y + Z))
  Q = Q + 9 * Cos(W + Z)
  Q = Q + 7 * Cos(x + Z - 2 * V)
  Q = Q + 6 * Cos(Z - W)
  Q = Q + 5 * Cos(2 * (V + y) - x + Z)
  Q = Q + 3 * Cos(x + 2 * (y + V + Z))
  Q = Q - 3 * Cos(W + 2 * (y + Z))
  Q = Q + 3 * Cos(2 * (y + Z) - W)
  Q = Q + 3 * Cos(2 * (V + y) + Z)
  Q = Q - 3 * Cos(2 * (x + y + Z - V))
  Q = Q - 3 * Cos(x + 2 * (y - V) + Z)
  Q = Q + 3 * Cos(2 * (V - x) + Z)
  Q = Q + 3 * Cos(2 * V + Z)
  Q = Q + 3 * Cos(2 * (y - V) + Z - W)
  Q = Q + 3 * Cos(Z - 2 * V)
  Q = Q + 3 * Cos(2 * (x + y) + Z)

' Return result in decimal degrees
  Delta_e = Q / 36000000

  End Function

' ====================================================================

