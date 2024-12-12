VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormatterInterface"
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
'@Folder("VBAMonologger.Formatter")
'@FQCN("VBAMonologger.Formatter.FormatterInterface")
'@ModuleDescription("Interface for formatters logging.")
'@Interface
'
' Defines the behavior a formatter must have to format log messages. We uses formatters
' to transform log data (such as messages, log levels, context, etc.) into a readable
' representation or a specific format, like JSON, plain text, or others.
''

Option Explicit

'@Description("Format a log record.")
Public Function format(ByRef record As LogRecordInterface) As Variant
End Function

'@Description("Formats a collection of log record.")
Public Function formatBatch(ByRef records() As LogRecordInterface) As Variant
End Function

