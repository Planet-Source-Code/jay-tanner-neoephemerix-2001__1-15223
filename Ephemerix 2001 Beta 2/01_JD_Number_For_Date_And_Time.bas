Attribute VB_Name = "JD_Number_For_Date_And_Time"
  Option Explicit

' ===================================================================
'
' Compute JDE value for given date string and time of day.
' The date string is NOT case sensitive.
'
' This function checks for invalid dates.
'
' The JDE value is the ephemeris JD value often seen in astronomical
' computations.  In astronomy, the day begins at noon instead of at
' midnight as on the civil calendar which means that the JD value
' is 12 hours (0.5 day) behind the civil JD number value.
'
' For example, noon of 31 Dec 1996 marks the instant of the beginning
' of JD 2450449.
'
' On the civil calendar, since dates are reckoned from midnight,
' instead of noon, the actual value of JD for 0h on 31 Dec 1996 is
' (2450449 - 0.5) = 2450448.5, which is the JD value that would be
' used in astronomical computations referring to 0h on that calendar
' date.
'
' The date argument has the general format:  "Dd Mmm Yyyy BC|AD"
' and the day fraction value ranges from 0.0 to 1.0
'
' Typical valid date string examples are:
' "1 Jan 4713 BC"   or   "20 May 1066 AD"   or   "4 Jul 1776"
'
' The "BC" is optional as required.  The "AD" is always implied and
' assumed unless "BC" is specifically indicated.
'
' This function automatically selects the Julian or Gregorian
' calendar mode depending on the given date.

  Public Function JDE_For(Dd_Mmm_Yyyy_BCAD, Time_or_Frac)
' LEVEL 1
' DEPENDENCY LIST: Day_Frac_Equiv_To()
'                  Calendar_Date_For()

  Dim D  As Single  ' Day
  Dim M  As String  ' Month
  Dim y  As String  ' Year
  Dim DS As String  ' Date string
  
' Auxiliary working variables
  Dim i  As Integer
  Dim j  As Integer
  Dim k  As Integer
  Dim Q  As String
  
  Dim NumChars As String  ' Valid numerical ASCII characters
  Dim MAbbrevs As String  ' Valid month name abbreviations
  Dim JD As Double        ' Julian day number starting at noon
  Dim Day_Frac As Double  ' Fraction of a day corresponding time

' This is the reconstructed date and is used to check if the input
' date was valid.  The computed JDE value is used to compute the
' corresponding calendar date.  If the dates do NOT match, then the
' input date argument was invalid and an error is returned.
  Dim Test_Date As String

          
' Define month abbreviations string
  MAbbrevs = "JANFEBMARAPRMAYJUNJULAUGSEPOCTNOVDEC"
 
' Define numerical character string
  NumChars = "0123456789"
 
' Read input date string argument
  DS = Trim(UCase(Dd_Mmm_Yyyy_BCAD))

  If Val(DS) = 0 Then GoTo ERROR_HANDLER
  
' Determine whether 2nd argument is the time or fraction of a day.
' If it is a time, then convert it into a fraction of the day.
' Otherwise, leave it as-is.
' If there is NO colon in the (Time_or_Frac) string, then it is assumed
' to be a fraction of a day.  A colon in the string means it is a time
' value and it will be converted into the equivalent fraction of a day.
  If InStr(Time_or_Frac, ":") <> 0 Then
     Day_Frac = Day_Frac_Equiv_To(Time_or_Frac)
  Else
     Day_Frac = Val(Time_or_Frac)
  End If

  If InStr(DS, "-") > 0 Then GoTo ERROR_HANDLER
  
' Extract numerical value of the month day from (DS).
' If the day is less than 1 or greater than 31, then it
' is definitely invalid in any case.
     D = Val(DS)
  If D < 1 Or D > 31 Then GoTo ERROR_HANDLER
     
  Test_Date = Trim(D) & " "
 
' Extract the three letter month abbreviation from the date string
' and determine the corresponding month number (1 to 12).
      Q = ""
  For i = 1 To Len(DS)
      If InStr(NumChars, Mid(DS, i, 1)) = 0 Then Exit For
  Next i
       Q = Trim(Mid(DS, i, Len(DS)))
       M = Trim(Left(Q, 1) & Mid(Q, 2, 2))
           Test_Date = Test_Date & M & " "
       M = 1 + Int(InStr(1, MAbbrevs, M) - 1) / 3
       If (M < 1 Or M > 12) Then GoTo ERROR_HANDLER
               
' Extract value of the year from the date string and normalize the
' numerical value for BC era if required.
  For i = 1 To Len(Q)
      If InStr(NumChars, Mid(Q, i, 1)) <> 0 Then Exit For
  Next i
       y = Trim(Mid(Q, i, Len(Q)))
           If Right(y, 2) <> "BC" Then y = Val(y) _
           Else y = 1 - Val(y)

 If y <= 0 Then
    Test_Date = Test_Date & (1 - y) & " BC"
 Else
    Test_Date = Test_Date & y & " AD"
 End If

' At this point, the three numerical date variables, D, M and Y,
' should now be ready for use in the subsequent JD computation.

' First compute the JD number according to the old Julian calendar.
  k = Int((14 - M) / 12)
 JD = D + Int(367 * (M + (k * 12) - 2) / 12) _
    + Int(1461 * (y + 4800 - k) / 4) - 32113

' Auto-select the proper calendar mode. If the date is prior
' to 15 Oct 1582, then use the Julian calendar, otherwise
' use the Gregorian calendar.
' The official final date on the old Julian calendar was
' Thursday, 4 Oct, 1582, which was followed by the first official
' date on the Gregorian calendar, Friday, 15 Oct, 1582.
  If JD > 2299160 Then
     JD = JD - (Int(3 * Int((y + 100 - k) / 100) / 4) - 2)
  End If

  JD = JD - 0.5

  Q = UCase(Trim(Calendar_Date_For(JD)))
  If Q <> Trim(Test_Date) Then GoTo ERROR_HANDLER
  
Q = Q

' Done - Return the astronomical JD value for given
' date and time of day.
  JDE_For = JD + Val(Day_Frac)

  Exit Function

' Handle erroneous date argument
ERROR_HANDLER:
  JDE_For = "ERROR: """ & Dd_Mmm_Yyyy_BCAD _
  & """ = Invalid calendar date"
  
  End Function

' ===================================================================

