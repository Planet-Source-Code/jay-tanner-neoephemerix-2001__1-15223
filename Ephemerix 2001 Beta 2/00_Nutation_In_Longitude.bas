Attribute VB_Name = "Nutation_In_Longitude"
  Option Explicit

' ====================================================================

' Nutational correction ecliptical longitude

' Compute nutation in ecliptical longitude in decimal degrees.  In terms
' of accuracy, the value in arc seconds is to about ±0.001"
'
  Public Function Delta_Psi(At_JDE)
' LEVEL 0

  Dim T  As Double ' Julian centuries since J2000.0
  Dim T2 As Double ' T to the power of 2
  Dim T3 As Double ' T to the power of 3
  
  Dim Q As Double  ' Nutation series accumulator
  
  Dim V As Double  ' Mean elongation of the moon from the sun
  Dim W As Double  ' Mean anomaly of the sun
  Dim x As Double  ' Mean anomaly of the moon
  Dim y As Double  ' Moon's argument of latitude
  
' Longitude of ascending node of lunar orbit on the ecliptic
' measured from the mean equinox of date.
  Dim Z As Double
  
  T = (At_JDE - 2451545#) / 36525
   
' Compute Mean elongation of moon in radians
  V = 297.85036 + 445267.11148 * T - 0.0019142 * T2 + T3 / 189474
  V = V * Atn(1) / 45
  
' Compute mean anomaly of the sun in radians
  W = 357.52772 + 35999.05034 * T - 0.0001603 * T2 - T3 / 300000
  W = W * Atn(1) / 45
  
' Compute mean anomaly of moon in radians
  x = 134.96298 + 477198.867398 * T + 0.0086972 * T2 + T3 / 56250
  x = x * Atn(1) / 45
  
' Compute moon's argument of latitude in radians
  y = 93.27191 + 483202.017538 * T - 0.0036825 * T2 + T3 / 327270
  y = y * Atn(1) / 45
  
' Compute longitude of moon's ascending node in radians
  Z = 125.04452 - 1934.136261 * T + 0.0020708 * T2 + T3 / 450000
  Z = Z * Atn(1) / 45
  
' Proceed to compute the nutation in longitude in arc seconds
  Q = Sin(Z) * (-174.2 * T - 171996)
  Q = Q + Sin(2 * (y + Z - V)) * (-1.6 * T - 13187)
  Q = Q + Sin(2 * (y + Z)) * (-2274 - 0.2 * T)
  Q = Q + Sin(2 * Z) * (0.2 * T + 2062)
  Q = Q + Sin(W) * (1426 - 3.4 * T)
  Q = Q + Sin(x) * (0.1 * T + 712)
  Q = Q + Sin(W + 2 * (y + Z - V)) * (1.2 * T - 517)
  Q = Q + Sin(2 * y + Z) * (-0.4 * T - 386)
  Q = Q - 301 * Sin(x + 2 * (y + Z))
  Q = Q + Sin(2 * (y + Z - V) - W) * (217 - 0.5 * T)
  Q = Q - 158 * Sin(x - 2 * V)
  Q = Q + Sin(2 * (y - V) + Z) * (129 + 0.1 * T)
  Q = Q + 123 * Sin(2 * (y + Z) - x)
  Q = Q + 63 * Sin(2 * V)
  Q = Q + Sin(x + Z) * (0.1 * T + 63)
  Q = Q - 59 * Sin(2 * (V + y + Z) - x)
  Q = Q + Sin(Z - x) * (-0.1 * T - 58)
  Q = Q - 51 * Sin(x + 2 * y + Z)
  Q = Q + 48 * Sin(2 * (x - V))
  Q = Q + 46 * Sin(2 * (y - x) + Z)
  Q = Q - 38 * Sin(2 * (V + y + Z))
  Q = Q - 31 * Sin(2 * (x + y + Z))
  Q = Q + 29 * Sin(2 * x)
  Q = Q + 29 * Sin(x + 2 * (y + Z - V))
  Q = Q + 26 * Sin(2 * y)
  Q = Q - 22 * Sin(2 * (y - V))
  Q = Q + 21 * Sin(2 * y + Z - x)
  Q = Q + Sin(2 * W) * (17 - 0.1 * T)
  Q = Q + 16 * Sin(2 * V - x + Z)
  Q = Q + Sin(2 * (W + y + Z - V)) * (0.1 * T - 16)
  Q = Q - 15 * Sin(W + Z)
  Q = Q - 13 * Sin(x + Z - 2 * V)
  Q = Q - 12 * Sin(Z - W)
  Q = Q + 11 * Sin(2 * (x - y))
  Q = Q - 10 * Sin(2 * (y + V) + Z - x)
  Q = Q - 8 * Sin(x + 2 * (y + V + Z))
  Q = Q + 7 * Sin(W + 2 * (y + Z))
  Q = Q - 7 * Sin(W + x - 2 * V)
  Q = Q - 7 * Sin(2 * (y + Z) - W)
  Q = Q - 7 * Sin(2 * V + 2 * y + Z)
  Q = Q + 6 * Sin(2 * V + x)
  Q = Q + 6 * Sin(2 * (x + y + Z - V))
  Q = Q + 6 * Sin(x + 2 * (y - V) + Z)
  Q = Q - 6 * Sin(2 * (V - x) + Z)
  Q = Q - 6 * Sin(2 * V + Z)
  Q = Q + 5 * Sin(x - W)
  Q = Q - 5 * Sin(2 * (y - V) + Z - W)
  Q = Q - 5 * Sin(Z - 2 * V)
  Q = Q - 5 * Sin(2 * (x + y) + Z)
  Q = Q + 4 * Sin(2 * (x - V) + Z)
  Q = Q + 4 * Sin(W + 2 * (y - V) + Z)
  Q = Q + 4 * Sin(x - 2 * y)
  Q = Q - 4 * Sin(x - V)
  Q = Q - 4 * Sin(W - 2 * V)
  Q = Q - 4 * Sin(V)
  Q = Q + 3 * Sin(x + 2 * y)
  Q = Q - 3 * Sin(2 * (y + Z - x))
  Q = Q - 3 * Sin(x - V - W)
  Q = Q - 3 * Sin(W + x)
  Q = Q - 3 * Sin(x + 2 * (y + Z) - W)
  Q = Q - 3 * Sin(2 * (V + y + Z) - W - x)
  Q = Q - 3 * Sin(3 * x + 2 * (y + Z))
  Q = Q - 3 * Sin(2 * (V + y + Z) - W)

' Return result in degrees
  Delta_Psi = Q / 36000000

  End Function

' ====================================================================

