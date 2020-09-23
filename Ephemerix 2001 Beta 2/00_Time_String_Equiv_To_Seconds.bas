Attribute VB_Name = "Time_String_Equiv_To_Seconds"
  Option Explicit

' ===================================================================

' Given time as seconds, convert to time in the standard format
' of  10:02:34 (Modified to 0 decimals for this program)

  Public Function Time_Equiv_To(Seconds)
' LEVEL 0

  Dim HH As String
  Dim MM As String
  Dim SS As String
  Dim Q  As String
  Dim S  As Double

  Dim Sign As String
  
  S = Val(Seconds) ' Read the seconds argument

  If S >= 0 Then Sign = "" Else Sign = "-": S = -S
  
' Compute hours
  HH = Int(S / 3600): S = S - 3600 * HH
' Compute minutes
  MM = Int(S / 60): S = S - 60 * MM
' Compute seconds
  SS = Format(S, "0#") ' Modified to 0 decimals or this program
  
' Correct for any values of 60
  If Val(SS) = 60 Then MM = MM + 1: SS = "00"
  If MM = 60 Then HH = HH + 1: MM = "00"
  
' Format and output the equivalent time string
  If HH = 0 Then HH = "00:" Else HH = Format(HH, "0#") & ":"
  If MM <> "" Then MM = Format(MM, "0#") & ":"
  If SS <> "" Then SS = SS
  Time_Equiv_To = Trim(Sign & HH & MM & SS)
  
  End Function

' ===================================================================


