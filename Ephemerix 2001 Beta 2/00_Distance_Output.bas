Attribute VB_Name = "Distance_Output"
  Option Explicit

' ===================================================================

' Function to return distance value according to interface settings.
' The (Pos_sign) set to true will attach a (+) to any non-negative
' values returned.  Kilometers and miles are displayed to the nearest
' whole unit.
'
' Modified version:  (km) and (mi) are in millions
'

  Public Function Dist_Out(AU_In, D_Units, Pos_Sign As Boolean)
  
' Read the raw data value in AUs
  Dim D As String
      D = Val(AU_In)

' Format control string
  Dim Fmt$
      
' Read the distance units setings
  Dim DU As String
      DU = UCase(Trim(D_Units))

' Convert raw AUs into equivalent km or mi according to mode
  If DU = "KM" Then D = D * 149597870: GoTo KM_OUT
  If DU = "MI" Then D = D * 92955806.8380657: GoTo MI_OUT
  If DU = "AU" Then GoTo AU_OUT

' Drop through if invalid units
  Dist_Out = "ERROR: """ & D_Units & """ = Invalid distance units"
  Exit Function
 
AU_OUT:
     If D < 10 Then Fmt$ = "#0.######0" Else Fmt$ = "#0.#####0"
     D = Format(D, Fmt$)
  If Pos_Sign = True And Val(D) >= 0 Then D = "+" & D
     Dist_Out = Right(Space(15) & D & " AU", 15)
  Exit Function

KM_OUT:
     If D < 10 Then Fmt$ = "#0.######0" Else Fmt$ = "#0.#####0"
     D = Format(D / 1000000#, Fmt$)
  If Pos_Sign = True And Val(D) >= 0 Then D = "+" & D
     Dist_Out = Right(Space(15) & D & " km*", 15)
  Exit Function
 
MI_OUT:
     If D < 10 Then Fmt$ = "#0.######0" Else Fmt$ = "#0.#####0"
     D = Format(D / 1000000#, Fmt$)
  If Pos_Sign = True And Val(D) >= 0 Then D = "+" & D
     Dist_Out = Right(Space(15) & D & " mi*", 15)
  
  End Function

' ===================================================================

