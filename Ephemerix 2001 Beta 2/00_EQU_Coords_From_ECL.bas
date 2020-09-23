Attribute VB_Name = "EQU_Coords_From_ECL"
  Option Explicit

' ====================================================================

' Function to convert geocentric ecliptical coordinates in degrees
' into corresponding right ascension and declination and return
' coordinates as a delimited 2D vector in the form  RA|Decl

  Public Function EQU_Lng_Lat_Equiv_To(Ecl_Lng, Ecl_Lat, Obliquity)
' LEVEL 0

  Dim Q As Variant

  Dim Sin_e As Double
      Sin_e = Sin(Obliquity * Atn(1) / 45)

  Dim Cos_e As Double
      Cos_e = Cos(Obliquity * Atn(1) / 45)

  Dim x As Double

  Dim y As Double

  x = Sin(Ecl_Lng * Atn(1) / 45) * Cos_e _
    - Tan(Ecl_Lat * Atn(1) / 45) * Sin_e

  y = Cos(Ecl_Lng * Atn(1) / 45)

  EQU_Lng_Lat_Equiv_To = Atn(y / x) * 45 / Atn(1)
  If x < 0 Then EQU_Lng_Lat_Equiv_To = EQU_Lng_Lat_Equiv_To + 180

     Q = Sin(Ecl_Lat * Atn(1) / 45) * Cos_e _
       + Cos(Ecl_Lat * Atn(1) / 45) * Sin_e _
       * Sin(Ecl_Lng * Atn(1) / 45)

  EQU_Lng_Lat_Equiv_To = EQU_Lng_Lat_Equiv_To & "|" _
  & Atn(Q / Sqr(-Q * Q + 1)) * 45 / Atn(1)
  
  End Function
  
' ====================================================================



