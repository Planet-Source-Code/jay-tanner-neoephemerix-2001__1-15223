Attribute VB_Name = "HXYZ_Coords_From_HLBR_Coords"
  Option Explicit

' ====================================================================

' This function computes the heliocentric rectangular X,Y,Z coords
' from a given set of L,B,R coordinates given a delimited data vector
' in the format L|B|R with angles expressed in degrees and R expressed
' in astronomical units.
'
' The L|B|R| arguments are the values returned by the above function.
'
' The X,Y,Z coordinates are also returned as a delimited data vector
' in the format X|Y|Z

  Public Function HXYZ_Coords_From(HLBR_Vector)
' LEVEL 1
' DEPENDENCY: Val_of_Coord()

  Dim L As Double
  Dim B As Double
  Dim R As Double

  L = Val_of_Coord("L", HLBR_Vector) * Atn(1) / 45
  B = Val_of_Coord("B", HLBR_Vector) * Atn(1) / 45
  R = Val_of_Coord("R", HLBR_Vector)
  
  HXYZ_Coords_From = R * Cos(B) * Cos(L) & "|" _
  & R * Cos(B) * Sin(L) & "|" & R * Sin(B)

  End Function

' ====================================================================



