Attribute VB_Name = "modError"
'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    Note:  "modError.bas" was created by "Mel Grubb II".         ****
'   ****                                                                 ****
'   ****    It came from the program called "FormShaper" which can be    ****
'   ****    found at:                                                    ****
'   ****                                                                 ****
'   ****    http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=28524&lngWId=1
'   ****                                                                 ****
'   ****    Using it helped me to eliminate a couple of errors that I    ****
'   ****    didn't previously know existed.                              ****
'   ****                                                                 ****
'   ****    Thank you Mel Grubb II.                                      ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'
'===============================================================================
'   modError - Central error handling support module
'   Provides centralized error handling and support for logging errors to the
'   event log.
'
'   Version   Date        User            Notes
'   1.0     11/16/00    Mel Grubb II    Initial version
'   1.1     11/29/00    Mel Grubb II    Added error handlers
'   Applied new coding standards
'   1.2     09/05/01    Mel Grubb II    Added Trace command
'   Removed Error enumerations
'===============================================================================

Option Explicit

'===============================================================================
'   Constants
'===============================================================================

Private Const mc_strModuleID   As String = "modError."     'Used to identify the location of errors

'===============================================================================
'   Global variables
'===============================================================================

Private g_blnDebug             As Boolean   'Whether or not the program is in debug mode

'===============================================================================
'   AppVersion - Standardize the formatting of the application version number
'
'   Arguments: None
'
'   Notes:
'===============================================================================

Public Function AppVersion() As String

    On Error GoTo ExitHandler
    AppVersion = App.Major & "." & Format$(App.Minor, "00") & "." & Format$(App.Revision, "0000")
    Exit Function

ExitHandler:
    AppVersion = "<Error>"

End Function

'==============================================================================
'   ProcessError - Logs the specified error to the NT error log.
'
'   Parameters:
'   objErr (IN) - the error to be logged
'   ProcedureID (IN) - the module or method name where the error occurred.
'   blnReraiseError (IN) - True if the error should be reraised; False otherwise.
'
'   Notes:
'==============================================================================

Public Sub ProcessError(ByRef objErr As ErrObject, Optional ByVal ProcedureID As String, Optional ByVal blnReraiseError As Boolean = False)

Dim strMessage                 As String
Dim strTitle                   As String

    On Error GoTo ExitHandler
'   Build the simple error string for the dialog
    strMessage = "Error Number = " & Err.Number & " (0x" & Hex$(Err.Number) & ")" & vbNewLine & _
                 "Description = " & Err.Description & vbNewLine & _
                 "Source = " & objErr.Source
    If Len(ProcedureID) > 0 Then
        strMessage = strMessage & vbNewLine & "Module = " & ProcedureID
    End If
    If Erl <> 0 Then
        strMessage = strMessage & vbNewLine & "Line = " & Erl
    End If

'   Show the error dialog
    strTitle = App.Title & " [" & AppVersion() & "]"
    MsgBox strMessage, vbOKOnly, strTitle

'   Expand the error before logging
    strMessage = strTitle & vbNewLine & strMessage

'   Log the error to the event log or log file, and the debug window
    App.LogEvent strMessage, vbLogEventTypeError
    Debug.Print vbNewLine & strMessage

'   Reraise the error if necessary
    If blnReraiseError Then
        ReraiseError objErr, ProcedureID
    End If

'   The next line will only be executed in Debug mode while in the IDE.
'   It causes the application to stop so that the programmer can debug.
    Debug.Assert StopInIDE() = True

ExitHandler:
'   Release any screen locks
    Screen.MousePointer = vbDefault

End Sub

'==============================================================================
'   ReraiseError - reraises the specified error.
'
'   Parameters:
'   objErr (IN) - the error to be reraised
'   strModuleID (IN) - the module or method name where the error occurred.
'
'   Notes:
'==============================================================================

Private Sub ReraiseError(objErr As ErrObject, Optional ByVal strModuleID As String = vbNullString)

    On Error Resume Next
    If Len(strModuleID) > 0 Then
        Err.Raise objErr.Number, strModuleID & vbNewLine & objErr.Source, objErr.Description, objErr.HelpFile, objErr.HelpContext
    Else
        Err.Raise objErr.Number, objErr.Source, objErr.Description, objErr.HelpFile, objErr.HelpContext
    End If
    On Error GoTo 0

End Sub

'===========================================================================
'   StartLogging
'===========================================================================

Public Sub StartLogging(ByVal LogTarget As String, LogMode As LogModeConstants)

    App.StartLogging LogTarget, LogMode

End Sub

'===========================================================================
'   StopInIDE - Causes a stop, but only in development mode
'
'   Arguments: None
'
'   Notes:
'===========================================================================

Private Function StopInIDE() As Boolean

    On Error GoTo ExitHandler
    Stop
    StopInIDE = True

ExitHandler:

End Function

'===============================================================================
'   Trace - Adds statements to trace log
'
'   Arguments:
'   Expression - String to append to trace log
'
'   Notes: Used to build a trace log in a finished executable since there is no
'   debug window.  The trace log will be appended to the Error log in the event an
'   error is trapped down the line.
'
'   g_blnDebug is checked here, but the calling application will probably benefit
'   if it is also checked before any string concatenations are performed like this
'   If g_blnDebug Then Trace "ProcName('" & Param1 & "')"
'===============================================================================

Public Sub Trace(ByRef Expression As String)

    If g_blnDebug Then
        Debug.Print Expression
        App.LogEvent Expression, vbLogEventTypeInformation
    End If

End Sub
