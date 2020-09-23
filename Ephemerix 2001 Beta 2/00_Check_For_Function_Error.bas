Attribute VB_Name = "Check_For_Function_Error"
  Option Explicit

' ====================================================================

  Public Function Error_In(Returned_Value) As Boolean
' V1.0
' Return error status of returned value of a function.
'
' This function is NOT case sensitive.
' This makes it easier to detect if an error occured within one
' of the functions.

' Just pass the returned value to this function as an argument to
' find out if an error was returned.
'
' If the returned string from a function contains the substring
' "ERROR", then return boolean "True", otherwise return "False".

  If InStr(UCase(Returned_Value), "ERROR") > 0 Then
     Error_In = True
  Else
     Error_In = False
  End If
  
  End Function

' ====================================================================


