VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HandlerInterface"
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
'@Folder("VBAMonologger.Handler")
'@FQCN("VBAMonologger.Handler.HandlerInterface")
'@ModuleDescription("Interface of handler log, responsible for determining where and how a log message will be recorded.")
'@Interface
'
' Defines the behavior of log handlers. Handlers are responsible for how log messages are
' processed and where they are sent, whether that’s to a file, an external service, the
' VBA console, or other destinations.
''

Option Explicit

''
' Determines the current log level of handler.
'
'@Description("The minimum logging level at which this handler will be triggered.")
''
Public level As VBAMonologger.LOG_LEVELS

''
' Determines whether or not the log record should "bubble up" to other handlers in the stack after being processed by
' the current handler. When you have multiple handlers in the stack of your logger, and want a specific handler to be
' the only one processing certain log levels or messages, you have to set bubble to false for that handler. This is
' useful if you need specific messages to go only to a certain location.
'
'  - If bubble is set to true (the default), the log record will be processed by the current handler and continue to
'    "bubble up" to other handlers in the stack. This allows multiple handlers to process the same log message.
'  - If bubble is set to false, the log record will stop propagating after being processed by the current handler. This
'    essentially "catches" the record, preventing it from being handled by other handlers in the stack.
'
'@Description("Whether the messages that are handled can bubble up the stack or not.")
''
Public bubble As Boolean

''
' Get the associated formatter of current handler.
'
'@Description("Get the formatter of handler.")
''
Public Property Get formatter() As VBAMonologger.FormatterInterface
End Property

''
' By ensuring that handlers can define their own formatters, its provides significant
' flexibility for managing logs in various applications.
'
'@Description("Set the formatter of handler.")
''
Public Property Set formatter(ByRef paramFormatter As VBAMonologger.FormatterInterface)
End Property

''
' Provides a way for handlers to filter log messages. This ensures that a handler only
' processes messages that are relevant to it, which can improve performance and reduce
' unnecessary processing. It returns true if the handler is capable of processing the
' log message and false if the handler should not process the log message.
'
'@Description("Checks whether the given log record will be handled by this handler.")
''
Public Function isHandling(ByRef LogRecord As VBAMonologger.LogRecordInterface) As Boolean
End Function

''
' All records may be passed to this method, and the handler should discard those that it does not want to handle.
' The return value of this function controls the bubbling process of the handler stack. Unless the bubbling is interrupted
' (by returning true), the Logger class will keep on calling further handlers in the stack with a given log record.
'
'@Description("Handles a log record.")
''
Public Function handle(ByRef LogRecord As VBAMonologger.LogRecordInterface) As Boolean
End Function

'@Description("Handles a set of records at once.")
Public Function handleBatch(ByRef logRecords() As VBAMonologger.LogRecordInterface) As Boolean
End Function

''
' Closing a Handler means flushing all buffers and freeing any open resources/handles.
'
'@Description("Closes handler means, ends a log cycle and frees all resources used by the handler.")
''
Public Sub closeHandler()
End Sub
