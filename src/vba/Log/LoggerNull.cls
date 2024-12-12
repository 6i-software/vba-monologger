VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoggerNull"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ------------------------------------- '
'                                       '
'    VBA Monologger                     '
'    Copyright © 2024, 6i software      '
'                                       '
' ------------------------------------- '
'
'@Exposed
'@Folder("VBAMonologger.Log")
'@FQCN("VBAMonologger.Log.LoggerNull")
'@ModuleDescription("This Logger can be used to avoid conditional log calls.")
'
' The **LoggerNull** is useful in the following situations:
' - **Temporarily disabling logging**: If you have a component or feature where you want to disable
'   logging without modifying the code or reconfiguring the logging system, you can replace the
'   active logger with a **LoggerNull**.
' - **Default or fallback behavior**: If a part of your application always expects a logger object,
'   you can pass a **LoggerNull** to prevent errors or exceptions when logging is not configured or
'   needed. This avoids `null` values or additional checks in the code to determine if a logger is
'   available.
' - **Unit testing**: During testing, you can use a **LoggerNull** to suppress log output while
'   focusing on testing other aspects of the system.
'
' Unlike a standard logger (which writes to a file, database, or sends notifications), the **LoggerNull**
' uses no resources to process log messages. It is a practical solution for handling cases where logging
' is unnecessary, while maintaining a consistent interface throughout your code.
''

Option Explicit

Implements VBAMonologger.LoggerInterface

'@inheritdoc
Private Sub LoggerInterface_log(ByVal level As LOG_LEVELS, ByVal message As String, Optional context As Scripting.IDictionary = Nothing, Optional extra As Scripting.IDictionary = Nothing)
End Sub

'@inheritdoc
Private Sub LoggerInterface_alert(ByVal message As String, Optional context As Scripting.IDictionary = Nothing, Optional extra As Scripting.IDictionary = Nothing)
End Sub

'@inheritdoc
Private Sub LoggerInterface_critical(ByVal message As String, Optional context As Scripting.IDictionary = Nothing, Optional extra As Scripting.IDictionary = Nothing)
End Sub

'@inheritdoc
Private Sub LoggerInterface_emergency(ByVal message As String, Optional context As Scripting.IDictionary = Nothing, Optional extra As Scripting.IDictionary = Nothing)
End Sub

'@inheritdoc
Private Sub LoggerInterface_error(ByVal message As String, Optional context As Scripting.IDictionary = Nothing, Optional extra As Scripting.IDictionary = Nothing)
End Sub

'@inheritdoc
Private Sub LoggerInterface_info(ByVal message As String, Optional context As Scripting.IDictionary = Nothing, Optional extra As Scripting.IDictionary = Nothing)
End Sub

'@inheritdoc
Private Sub LoggerInterface_notice(ByVal message As String, Optional context As Scripting.IDictionary = Nothing, Optional extra As Scripting.IDictionary = Nothing)
End Sub

'@inheritdoc
Private Sub LoggerInterface_trace(ByVal message As String, Optional context As Scripting.IDictionary = Nothing, Optional extra As Scripting.IDictionary = Nothing)
End Sub

'@inheritdoc
Private Sub LoggerInterface_warning(ByVal message As String, Optional context As Scripting.IDictionary = Nothing, Optional extra As Scripting.IDictionary = Nothing)
End Sub
