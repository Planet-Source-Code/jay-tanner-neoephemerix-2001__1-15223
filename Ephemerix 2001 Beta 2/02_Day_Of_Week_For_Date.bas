Attribute VB_Name = "Day_Of_Week_For_Date"
  Option Explicit

' ===================================================================
' Compute calendar day of the week for a given date.  It automatically
'  handles both the Julian and Gregorian calendars.
'
' The returned value is a three letter abbreviation such as, Sun, Mon,
' Tue, etc.
'
' The date is in the format:  Dd Mmm Yyyy BC|AD
'
' This function checks for invalid dates.

  Public Function Day_Of_Week_For(Dd_Mmm_Yyyy_BCAD)
' LEVEL 2
' DEPENDENCY: JDE_For()

  Dim Q As String   ' Working variable
    
  Dim DoW As String ' Day of week index (0 to 6)
    
' Compute day of week index (DoW).  This value ranges from 0 to 6
' where: 0=Sun, 1=Mon, 2=Tue, 3=Wed, 4=Thu, 5=Fri and 6=Sat
  Q = JDE_For(Dd_Mmm_Yyyy_BCAD, 0)

' Check for invalid date
  If InStr(UCase(Q), "ERROR") > 0 Then
     Day_Of_Week_For = Q
     Exit Function
  End If

' Drop through here if date was OK
  DoW = Int(JDE_For(Dd_Mmm_Yyyy_BCAD, 0) + 1.5) Mod 7

' Adjust if (DoW) is negative to prevent discontinuity.
  If DoW < 0 Then DoW = DoW + 7
  
' Return the weekday abbreviation corresponding to the (DoW) index.
  Day_Of_Week_For = Mid("SunMonTueWedThuFriSat", 1 + 3 * DoW, 3)
  
  End Function
  
' ===================================================================

