Attribute VB_Name = "Interface_Globals"
  Option Explicit

' Global variables read and computed from the interface settings
' each time the interface is read.

' The full interface date setting in "Dd Mmm Yyyy BC|AD" format
  Global The_Date As String

' The Delta T (dT) value in decimal
' This is the difference between dynamical time and UT
' where TD = UT + Delta T
  Global The_dT As Double

' The JDE value for the given date and time and dT
  Global The_JDE As Double

' The computation mode flag (HC|EC|EQ|STATS)
  Global COMP_MODE As String
  
' The hour angle display mode setting (HMS|DH|DMS|DD)
' for equatorial mode and sidereal time angles
  Global HA_MODE As String

' Longitude display mode (DMS|DD) for heliocentric and ecliptical modes
  Global LNG_MODE As String

' The latitude angle display mode setting (DMS|DD) for all modes
  Global DECL_LAT_MODE As String

' The distance units display mode setting (AU|KM|MI)
  Global DIST_UNITS As String

' This variable holds the text of the most recent interface error
  Global DATA_ERROR As String

' This is the hourly ephemeris mode flag (H|M)
' If value = H, then the hourly table has already been computed
' for the selected day, so, then generate a listing for each minute instead.
' If value = M, then ignore the double click.
  Global HM_TABLE_MODE As String

' Temporary day storage variable for hour/minute table date
  Global TEMP_D As String

