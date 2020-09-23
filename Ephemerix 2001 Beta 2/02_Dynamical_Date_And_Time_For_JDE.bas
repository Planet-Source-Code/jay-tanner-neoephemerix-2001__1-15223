Attribute VB_Name = "Dynamical_Date_And_Time_For_JDE"
Option Explicit

' Compute full dynamical date and time given complete JDE value
' which assumes all the necessary adjustments are included in the
' JDE value.
'
' NOTE: Result will be returned to the nearest whole second.

  Public Function Dynamical_Date_and_Time(At_JDE)
' LEVEL 2
' DEPENDENCY: Time_Equiv_To()
'             Calendar_Date_For()
'             Day_Of_Week_For

  Dim Q As Variant
  Dim U As Variant
  Dim W As Variant
 
' Read the JDE value
  Q = Val(Trim(At_JDE)) - 0.5

' Get fractional part of JDE value
     U = InStr(Q, ".")
  If U = 0 Then Q = Q & ".0"
     U = InStr(Q, ".")
     U = Val(Mid(Q, U, Len(Q)))
  
' Convert fraction of day into time string
  U = Time_Equiv_To(86400 * U)
   
' Compute calendar date for given JDE
  Q = Calendar_Date_For(At_JDE)

' Determine calendar week day
  W = Day_Of_Week_For(Q)

' Return dynamical date & corresponding time as data vector
  Dynamical_Date_and_Time = W & " - " & Q & "  at TD " & U

  End Function


