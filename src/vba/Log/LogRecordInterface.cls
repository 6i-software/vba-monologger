VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogRecordInterface"
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
'@FQCN("VBAMonologger.Log.LogRecordInterface")
'@ModuleDescription("Interface for a log record.")
'@Interface
'
' A log record refers to an entry that captures specific information about an event or an
' action that occurs within your application. This can be useful for debugging, tracking
' user actions, or auditing processes.
'
' A log record contains several key elements:
'
'  - Message: The main text describing the event (e.g., "User attempted to log in.").
'  - Level: The severity level of the log entry (e.g., Info, Warning, Error).
'  - Channel: It is a way to categorize or segment logs based on their source or purpose.
'  - Datetime: The date and time when the log entry was created.
'  - Context: Additional relevant information about the event (e.g., user ID).
'  - Extra: Optional additional information that might be helpful (e.g., memory usage).
'
' Only the message property have a setter. The other properties are read only.
''

Option Explicit

'@Description("Get the log record creation date and time.")
Public Property Get dateTime() As Date
End Property

'@Description("Get the log message.")
Public Property Get message() As String
End Property

'@Description("Set the log message.")
Public Property Let message(ByVal value As String)
End Property

'@Description("Get the channel name.")
Public Property Get channel() As String
End Property

'@Description("Get the level of the logger.")
Public Property Get level() As Variant
End Property

'@Description("Get the logging context, refers to additionnal information to add into the log message")
Public Property Get context() As Variant
End Property

'@Description("Get the additional log data to be used or produced by processors or formatters")
Public Property Get extra() As Variant
End Property
