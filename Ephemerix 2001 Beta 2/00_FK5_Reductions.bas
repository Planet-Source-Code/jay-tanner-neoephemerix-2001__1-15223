Attribute VB_Name = "FK5_Reductions"
  Option Explicit

' ====================================================================

' Compute the correction required to convert VSOP87 dynamical
' longitude into the FK5 system longitude.

' The arguments longitude (Lng_Deg) and latitude (Lat_Deg) are in
' decimal degrees and so is the returned correction value.

  Public Function FK5_Lng_Corr(At_JDE, Lng_Deg, Lat_Deg)
' LEVEL 0

  Dim T As Double
      T = (At_JDE - 2451545) / 36525
      
  Dim Q As Double

  Dim Lprime As Double
  Dim B      As Double

  B = Lat_Deg * Atn(1) / 45
  Lprime = (Lng_Deg - 1.397 * T - 0.00031 * T * T) * Atn(1) / 45

  Q = -0.09033 + 0.03916 * (Cos(Lprime) + Sin(Lprime)) * Tan(B)
  
  FK5_Lng_Corr = Q / 3600

  End Function


' Compute the correction required to convert VSOP87 dynamical
' latitude into the FK5 system latitude.

' The argument (Lng_Deg) is in decimal degrees and so is
' the returned correction value.

  Public Function FK5_Lat_Corr(At_JDE, Lng_Deg)
' LEVEL 0

  Dim T As Double
      T = (At_JDE - 2451545) / 36525
      
  Dim Lprime As Double
  
  Lprime = (Lng_Deg - 1.397 * T - 0.00031 * T * T) * Atn(1) / 45

  FK5_Lat_Corr = (0.03916 * (Cos(Lprime) - Sin(Lprime))) / 3600
  
  End Function
  
