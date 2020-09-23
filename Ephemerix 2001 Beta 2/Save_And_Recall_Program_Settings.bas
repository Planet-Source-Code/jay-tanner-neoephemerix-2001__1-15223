Attribute VB_Name = "Save_And_Recall_Program_Settings"

  Option Explicit

' =========================================================================
'
' Saves and recalls a form's control values and window size/position.
'
' =========================================================================
'
' This segment of code was written by David Edlen - davidnedlen@cs.com
' Nov-19-2000
'
' A one-step process to save all control values on a form to the Windows
' registry and to retrieve previously saved values and re-apply them to the
' form.  It will automatically save the values for each text box, check
' box, option button, list box, and combo list.  The window size and
' position are also saved and retrieved.  No special coding is required for
' individual controls.
'
' No additional coding is required when new controls are added to the form.
' Especially useful for complex "Options" forms.  All settings remain in
' the Windows registry after the application terminates.
'
' =========================================================================
'
' Usage:
'
' SaveFormSettings [PROGRAM_NAME], [FORM_NAME], True
'
' Saves all control values plus the window size and position for
' form frmOptions.
'
' GetFormSettings [PROGRAM_NAME], [FORM_NAME], True
'
' Retrieves all previously saved control values for form frmOptions and
' assigns those values to their respective controls; sets the form size
' and position to the last saved settings.
'
' =========================================================================
'
' Public Procedures:
'
' SaveFormSettings Statement
' ==========================
' Saves or creates a series of entries in the application's entry in the
' Windows registry.  The entries will consist of the values for each of
' the form's text boxes, check boxes, option buttons, list boxes, and combo
' lists.  Optionally, the form's window size and position are included in
' the entries.
'
' Syntax:
'   SaveFormSettings appname, form, [winsettings], [errnumber]
'   ----------------------------------------------------------
'   appname     =  String expression containing the name of the application

'   form        =  A form object to be saved

'   winsettings =  Boolean expression; window size and position settings
'                  will be saved if true

'   errnumber   =  Returned Long value: the code number of any procedure
'                  error, where 0 = no error.
'
' =========================================================================
'
' GetFormSettings Statement
' =========================
' Retrieves all settings for a form saved by the SaveFormSettings statement
' and assigns them to their respective controls.  The window size and
' position are also retrieved and applied.
'
' Syntax:
' GetFormSettings ["APP_NAME"], [FORM_NAME], [WINSETTINGS], [ERRNUMBER]
'
' APP_NAME    =  String expression containing the name of the application
'
' FORM_NAME   =  The name of the form object itself
'
' WINSETTINGS =  Boolean expression; The last saved window size and position
'                will be applied to the form if true.
' ERRNUMBER   =  Returned Long value: the code number of any procedure error.
'                0 = no error.
'
' =========================================================================
'
' Notes:
' Uses the VB SaveSettings and GetSettings statements.
' Registry entries are stored in:
' HKEY_CURRENT_USER\Software\VB and VBA Program Settings
'
' The registry entries are named as follows:
' ------------------------------------------
' Application name = appname parameter value
' Section name = name of the form object
' Key names = names of each of the form's controls
' Settings = the text, value, or list items of each control.  List items
' are saved as a single string with each item delimited by a Chr(11).
'
' =========================================================================
'
  Const WinHeight = "@Height@"
  Const WinWidth = "@Width@"
  Const WinTop = "@Top@"
  Const WinLeft = "@Left@"
'
' =========================================================================
'
  Public Sub SaveFormSettings _
 (ByVal pAppName As String, pForm As Form, _
  Optional pFormPosition As Boolean, _
  Optional pError As Long)

  Dim ix As Long
  Dim vName As String
  Dim vControl As Control
  Dim vError As Long
    
  On Error GoTo errSaveFormSettings
    
' Windows Settings
  If pFormPosition = True Then
     SaveSetting pAppName, pForm.Name, WinHeight, pForm.Height
     SaveSetting pAppName, pForm.Name, WinWidth, pForm.Width
     SaveSetting pAppName, pForm.Name, WinTop, pForm.Top
     SaveSetting pAppName, pForm.Name, WinLeft, pForm.Left
  End If
    
' Loop through the form's control collection.
' Save the value parameter for each control.
    
  For Each vControl In pForm.Controls
        
  On Error Resume Next
     With vControl
            ix = .Index
            If Err.Number = 343 Then
                vName = .Name
                Err.Clear
            Else
                vName = .Name & ":" & Trim(CStr(ix))
            End If
        End With
        
        On Error GoTo errSaveFormSettings
        If TypeOf vControl Is TextBox Then
            SaveSetting pAppName, pForm.Name, vName, vControl.Text
        ElseIf TypeOf vControl Is CheckBox _
        Or TypeOf vControl Is OptionButton Then
            SaveSetting pAppName, pForm.Name, vName, vControl.Value
        ElseIf TypeOf vControl Is ListBox _
        Or TypeOf vControl Is ComboBox Then
            SaveSetting pAppName, pForm.Name, vName, GetListString(vControl, vError)
            If vError <> 0 Then Err.Raise vError
        End If
    
    Next vControl

    Set vControl = Nothing

errSaveFormSettings:
    
    pError = Err.Number

End Sub
'
' =========================================================================
'
  Public Sub GetFormSettings(ByVal pAppName As String, pForm As Form, Optional pFormPosition As Boolean, Optional pError As Long)

  Dim ix As Long
  Dim vName As String
  Dim vControl As Control
  Dim vError As Long
    
  On Error GoTo errGetFormSettings
    
' Windows Settings
    
  If pFormPosition = True Then
     pForm.Height = GetSetting(pAppName, pForm.Name, WinHeight, pForm.Height)
     pForm.Width = GetSetting(pAppName, pForm.Name, WinWidth, pForm.Width)
     pForm.Top = GetSetting(pAppName, pForm.Name, WinTop, pForm.Top)
     pForm.Left = GetSetting(pAppName, pForm.Name, WinLeft, pForm.Left)
  End If
    
' Loop through the form's control collection.
' Retrieve the value parameter for each control.
    
  For Each vControl In pForm.Controls
        
        On Error Resume Next
        With vControl
            ix = .Index
            If Err.Number = 343 Then
                vName = .Name
                Err.Clear
            Else
                vName = .Name & ":" & Trim(CStr(ix))
            End If
        End With

        On Error GoTo errGetFormSettings
        If TypeOf vControl Is TextBox Then
            vControl.Text = GetSetting(pAppName, pForm.Name, vName, vControl.Text)
        ElseIf TypeOf vControl Is CheckBox _
        Or TypeOf vControl Is OptionButton Then
            vControl.Value = GetSetting(pAppName, pForm.Name, vName, vControl.Value)
        ElseIf TypeOf vControl Is ListBox _
        Or TypeOf vControl Is ComboBox Then
            PopulateList vControl, GetSetting(pAppName, pForm.Name, vName, ""), vError
            If vError <> 0 Then Err.Raise vError
        End If
    
  Next vControl

    Set vControl = Nothing

errGetFormSettings:
    
    pError = Err.Number

End Sub

Private Function GetListString(pControl As Control, pError As Long) As String

' Convert the contents of the specified list control to a string expression.
' The string will consist of all list items delimited by a (vbVerticalTab).

  Dim strList As Variant
  Dim ix As Long
    
    On Error GoTo errGetListString
    strList = ""
    
    If TypeOf pControl Is ListBox _
    Or TypeOf pControl Is ComboBox Then
        With pControl
            For ix = 0 To .ListCount - 1
                If strList <> "" Then
                    strList = strList & vbVerticalTab
                End If
                strList = strList & .List(ix)
            Next ix
        End With
    End If

    GetListString = strList
    
errGetListString:

    pError = Err.Number

End Function

 Private Sub PopulateList _
(pControl As Control, pListString As String, pError As Long)

' Convert a list string to list items and populate the specified list control.

  Dim arList As Variant
  Dim ix As Integer
    
    On Error GoTo errPopulateList

    If TypeOf pControl Is ListBox _
    Or TypeOf pControl Is ComboBox Then
    
        pControl.Clear
        arList = Split(pListString, vbVerticalTab)
        
        If IsArray(arList) Then
            For ix = LBound(arList) To UBound(arList)
                pControl.AddItem arList(ix)
            Next ix
        Else
            pControl.AddItem arList
        End If
        
    End If
    
errPopulateList:

    pError = Err.Number

End Sub

' =========================================================================

' This segment of code writen by Jay Tanner to apply the above program
' code written by David Edlen.
'
'
' The routines defined below will STORE and RECALL the form and control
' settings between each run of the program.
'
' AppName   = NeoEphemerix_2001
' Form name = NeoEphemerix_2001_Interface
'
' The Windows registry entries are stored in:
' HKEY_CURRENT_USER\Software\VB and VBA Program Settings

  Public Sub STORE_PROGRAM_SETTINGS()
  SaveFormSettings "NeoEphemerix_2001", NeoEphemerix_2001_Interface, True
  End Sub

  Public Sub RECALL_PROGRAM_SETTINGS()
  GetFormSettings "NeoEphemerix_2001", NeoEphemerix_2001_Interface, True
  End Sub

' =========================================================================

