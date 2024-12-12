VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoggerInterface"
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
'@FQCN("VBAMonologger.Log.LoggerInterface")
'@ModuleDescription("Describes a logger instance.")
'@Interface
'
' This interface is (*supposed to be*) designed according to the PSR-3 standard (see: https://www.php-fig.org/psr/psr-3/).
'
' Some remarks:
'
' - The 'debug' method exposes presents in PSR-3 is rename into 'trace'
' - The message MUST be a string.
' - The message MAY contain placeholders in the form. The placeholder {foo} will be replaced by the context data in dictionnary key "foo".
'   ```
'     Dim context as New Scripting.Dictionary
'     context Add "username" "'alice.cooper'"
'     logger.Info("User {username} created", context)
'     '' Ouput result with default line formatter
'     '' [2024/10/24 15:52:12] INFO: User 'alice.cooper' created
'   ```
''

Option Explicit

'@Description("Logs with an arbitrary level.")
Public Sub log(ByVal level As VBAMonologger.LOG_LEVELS, ByVal message As String, Optional ByRef context As Scripting.Dictionary = Nothing, Optional ByRef extra As Scripting.Dictionary = Nothing)
End Sub

'@Description("System is unusable.")
Public Sub emergency(ByVal message As String, Optional ByRef context As Scripting.Dictionary = Nothing, Optional ByRef extra As Scripting.Dictionary = Nothing)
End Sub

'@Description("Action must be taken immediately.")
Public Sub alert(ByVal message As String, Optional ByRef context As Scripting.Dictionary = Nothing, Optional ByRef extra As Scripting.Dictionary = Nothing)
End Sub

'@Description("Critical conditions.")
Public Sub critical(ByVal message As String, Optional ByRef context As Scripting.Dictionary = Nothing, Optional ByRef extra As Scripting.Dictionary = Nothing)
End Sub

'@Description("Runtime errors that do not require immediate action but should typically be logged and monitored.")
Public Sub error(ByVal message As String, Optional ByRef context As Scripting.Dictionary = Nothing, Optional ByRef extra As Scripting.Dictionary = Nothing)
End Sub

'@Description("Exceptional occurrences that are not errors.")
Public Sub warning(ByVal message As String, Optional ByRef context As Scripting.Dictionary = Nothing, Optional ByRef extra As Scripting.Dictionary = Nothing)
End Sub

'@Description("Normal but significant events.")
Public Sub notice(ByVal message As String, Optional ByRef context As Scripting.Dictionary = Nothing, Optional ByRef extra As Scripting.Dictionary = Nothing)
End Sub

'@Description("Interesting events.")
Public Sub info(ByVal message As String, Optional ByRef context As Scripting.Dictionary = Nothing, Optional ByRef extra As Scripting.Dictionary = Nothing)
End Sub

'@Description("Detailed debug information. The 'debug' method exposes presents in PSR-3 is rename into 'trace'.")
Public Sub trace(ByVal message As String, Optional ByRef context As Scripting.Dictionary = Nothing, Optional ByRef extra As Scripting.Dictionary = Nothing)
End Sub
