Attribute VB_Name = "Geocentric_Position"
  Option Explicit

' Function to compute the complete apparent geocentric position of
' a planet using the interrelated function modules.  The coordinates
' are returned as a delimited data vector in the format RA|Decl|Dist
'
' The right ascension and declination are returned as decimal degrees
' and the geocentric distance is in astronomical units.

  Public Function Geocentric_RA_Decl_Dist_For(Object_Name, At_JDE)

  Dim Q   As Variant ' Random work variable

  Dim LBR As String  ' Heliocentric spherical coordinates
  Dim XYZ As String  ' Rectangular coordinates

' Heliocentric & geocentric rectangular coordinates of Object
  Dim x   As Double
  Dim y   As Double
  Dim Z   As Double

' Heliocentric rectangular coordinates of Earth
  Dim Xe  As Double
  Dim Ye  As Double
  Dim Ze  As Double

' True distance between Earth and Object
  Dim D   As Double

' Light time between Earth and Object
  Dim LT  As Double

' Geocentric ecliptical longitude and latitude of Object
  Dim Ecl_Lng As Double
  Dim Ecl_Lat As Double

' Apparent obliquity of the ecliptic
  Dim Obliquity As Double

' Compute heliocentric coordinates of Earth
  LBR = LBR_For("Earth", At_JDE)
  XYZ = HXYZ_Coords_From(LBR)
   Xe = Val_of_Coord("X", XYZ)
   Ye = Val_of_Coord("Y", XYZ)
   Ze = Val_of_Coord("Z", XYZ)

' Compute heliocentric coordinates of Object
  LBR = LBR_For(Object_Name, At_JDE)

  If Error_In(LBR) Then
  Geocentric_RA_Decl_Dist_For = "ERROR: " & Object_Name & " = Invalid object name"
  Beep
  Exit Function
  End If


  XYZ = HXYZ_Coords_From(LBR)
    x = Val_of_Coord("X", XYZ)
    y = Val_of_Coord("Y", XYZ)
    Z = Val_of_Coord("Z", XYZ)

' Replace heliocentric rectangular coordinates with
' geocentric rectangular coordinates of Object
  x = x - Xe
  y = y - Ye
  Z = Z - Ze

' Compute true distance between Earth and Object
  D = Sqr(x * x + y * y + Z * Z)

' Compute light time between Earth and object in days
  LT = D * 5.77551830441213E-03

' Recompute heliocentric coordinates of Earth at
' original time minus the light time.
  LBR = LBR_For("Earth", At_JDE - LT)
  XYZ = HXYZ_Coords_From(LBR)
   Xe = Val_of_Coord("X", XYZ)
   Ye = Val_of_Coord("Y", XYZ)
   Ze = Val_of_Coord("Z", XYZ)

' Heliocentric coordinates of Object at original
' time minus the light time.
  LBR = LBR_For(Object_Name, At_JDE - LT)
  XYZ = HXYZ_Coords_From(LBR)
    x = Val_of_Coord("X", XYZ)
    y = Val_of_Coord("Y", XYZ)
    Z = Val_of_Coord("Z", XYZ)

' Replace heliocentric rectangular coordinates with
' geocentric rectangular coordinates of Object
  x = x - Xe
  y = y - Ye
  Z = Z - Ze

' Compute geocentric ecliptical longitude and latitude
  Ecl_Lng = ArcTan2(y, x, "Deg")
  Ecl_Lat = Atn(Z / Sqr(x * x + y * y)) * 45 / Atn(1)
  
' Apply FK5 corrections to ecliptical coordinates
  Q = Ecl_Lng + FK5_Lng_Corr(At_JDE, Ecl_Lng, Ecl_Lat)
  Ecl_Lat = Ecl_Lat + FK5_Lat_Corr(At_JDE, Ecl_Lng)
  Ecl_Lng = Q

' Apply correction for nutation in longitude
  Ecl_Lng = Ecl_Lng + Delta_Psi(At_JDE)

' Compute apparent ecliptic obliquity
  Obliquity = Ecliptic_Obliquity(At_JDE, True)

' Convert ecliptic coordinates into apparent RA and Decl
   Q = EQU_Lng_Lat_Equiv_To(Ecl_Lng, Ecl_Lat, Obliquity)

' Return the computed apparent geocentric RA, Decl and Distance
  Geocentric_RA_Decl_Dist_For = Q & "|" & D

  End Function

