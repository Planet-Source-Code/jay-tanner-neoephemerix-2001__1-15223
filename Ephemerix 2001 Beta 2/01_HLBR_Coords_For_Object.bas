Attribute VB_Name = "HLBR_Coords_For_Object"
  Option Explicit

' ====================================================================

' Function to compute VSOP87 heliocentric L,B,R coordinates for a
' given object.
'
' The spelling of the name of the object must be exact, but is NOT
' case sensitive.
'
' This function calls the individual planetary L,B,R modules.

  Public Function LBR_For(Object_Name, At_JDE)
' VERSION 1.01
' LEVEL 1
' DEPENDENCY: LBR_For_Earth_or_Sun()
'             LBR_For_Mercury()
'             LBR_For_Venus()
'             LBR_For_Mars()
'             LBR_For_Jupiter()
'             LBR_For_Saturn()
'             LBR_For_Uranus()
'             LBR_For_Neptune()

  Dim LBR    As String
      LBR = "ERROR"

  Dim Object As String
      Object = UCase(Trim(Object_Name))
  
  If Object = "SUN" Then LBR = LBR_For_Earth_or_Sun(At_JDE, "S")
  If Object = "MERCURY" Then LBR = LBR_For_Mercury(At_JDE)
  If Object = "VENUS" Then LBR = LBR_For_Venus(At_JDE)
  If Object = "EARTH" Then LBR = LBR_For_Earth_or_Sun(At_JDE, "E")

  If Object = "MARS" Then LBR = LBR_For_Mars(At_JDE)
  If Object = "JUPITER" Then LBR = LBR_For_Jupiter(At_JDE)
  If Object = "SATURN" Then LBR = LBR_For_Saturn(At_JDE)
  If Object = "URANUS" Then LBR = LBR_For_Uranus(At_JDE)
  If Object = "NEPTUNE" Then LBR = LBR_For_Neptune(At_JDE)

  If LBR = "ERROR" Then
     LBR_For = LBR & ": """ & Object & """ = Invalid object name"
     Beep
     Exit Function
  End If

  LBR_For = LBR

  End Function

' ====================================================================

