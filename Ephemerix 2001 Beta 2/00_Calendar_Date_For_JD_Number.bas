Attribute VB_Name = "Calendar_Date_For_JD_Number"
  Option Explicit

' ====================================================================

' This function is the opposite of the JDE_For() function.  It will
' return the calendar date string corresponding to the JD number and
' return it as a date string in the same format as 1 Jan 2000 BC|AD
'
' This function assumes the JD value is in the astronomical convention.

  Public Function Calendar_Date_For(JD_Num)
' LEVEL 0

  Dim D    As Double ' Proper day value (1 to 31)
  Dim M    As Double ' Proper month value (1 to 12)
  Dim Mmm  As String ' Output month abbreviation (Jan to Dec)
  Dim y    As Double ' Proper year value
  Dim Yyyy As String ' Output year string with "BC|AD" suffix
  Dim G    As Double ' Julian/Gregorian flag

  Dim JD  As Double  ' The given JD argument

' Auxiliary variables
  Dim Q   As Double
  Dim R   As Double
  Dim S   As Double
  Dim T   As Double
  Dim U   As Double

  Dim V   As Double

' Adjust astronomical JD argument value
  JD = Int(Val(Trim(JD_Num)) + 0.5)

' Determine calendar mode to use
' All dates up to 4 Oct 1582 AD use the Julian calendar.  After
' that date, the Gregorian calendar is used for computations.
  If JD < 2299161 Then G = 0 Else G = 1

' Compute auxiliary values
  Q = G * Int((JD / 36524.25) - 51.12264)
  R = JD + G + Q - Int(Q / 4)
  S = R + 1524
  T = Int((S / 365.25) - 0.3343)
  U = Int(T * 365.25)
  V = Int((S - U) / 30.61)

' Compute the raw, numerical calendar date elements
  D = S - U - Int(V * 30.61)
  M = (V - 1) + 12 * (V > 13.5)
  y = T - (M < 2.5) - 4716

' At this point the raw numerical values of D, M and Y have
' been computed.  Now they must be converted into the standard
' date format, "Dd Mmm Yyyy BC|AD", for output.

' Day of the month (1 to 31)
  D = Trim(D)

' Determine English month abbreviation (Jan to Dec)
  Mmm = " " & _
  Mid("JanFebMarAprMayJunJulAugSepOctNovDec", 3 * (M - 1) + 1, 3)
  Mmm = Mmm & " "

' Determine the year in BC|AD format
  If y < 0 Then
     Yyyy = Trim(1 - y) & " BC"
  Else
     Yyyy = Trim(y) & " AD"
  End If

' Finally, return the computed standard date string in
' the same format as  "12 Jan 2000 BC|AD"
  Calendar_Date_For = D & Mmm & Yyyy

  End Function

' ====================================================================

