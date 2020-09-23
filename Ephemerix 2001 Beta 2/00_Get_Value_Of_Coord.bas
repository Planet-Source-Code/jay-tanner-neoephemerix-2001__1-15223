Attribute VB_Name = "Get_Value_of_Coord"
  Option Explicit

' ====================================================================

' Function to extract an individual delimited coordinate value
' from a 3D coordinate vector for subsequent computation.
'
' The delimiter is the "|" character.
'
' Spherical coordinates are in the format L|B|R
'
' Cylindrical coordinates are in the format X|Y|Z
'
' Geocentric coordinates are in the format RA|Decl|Dist
'
' This function is NOT case sensitive.
'
  Public Function Val_of_Coord(LBR_or_XYZ_or_RA_Decl_Dist, From_Vector)
' LEVEL 0
  
  Dim i As Integer       ' Delimiter pointer

  Dim Coord_ID As String ' Coordinate symbol (L,B,R or X,Y,Z)
      Coord_ID = UCase(LBR_or_XYZ_or_RA_Decl_Dist)

' Check for valid coordinate symbol argument
  If Coord_ID = "L" Or Coord_ID = "B" Or Coord_ID = "R" _
  Or Coord_ID = "X" Or Coord_ID = "Y" Or Coord_ID = "Z" _
  Or Coord_ID = "RA" Or Coord_ID = "DECL" Or Coord_ID = "DIST" Then _
     GoTo COORD_OK

' Drop through here if error
  Val_of_Coord = "ERROR: """ & Coord_ID & """ = Invalid Coord ID" _
  & " - Must be L,B,R  or  X,Y,Z  or  RA,Decl,Dist"
  Beep
  Exit Function

COORD_OK:
    
    If Coord_ID = "L" Or Coord_ID = "X" Or Coord_ID = "RA" Then _
       Val_of_Coord = Val(From_Vector)

    i = InStr(1, From_Vector, "|") + 1
    If Coord_ID = "B" Or Coord_ID = "Y" Or Coord_ID = "DECL" Then _
       Val_of_Coord = Val(Mid(From_Vector, i, Len(From_Vector)))

    i = InStr(i, From_Vector, "|") + 1
    If Coord_ID = "R" Or Coord_ID = "Z" Or Coord_ID = "DIST" Then _
    Val_of_Coord = Val(Mid(From_Vector, i, Len(From_Vector)))

  End Function

' ====================================================================



