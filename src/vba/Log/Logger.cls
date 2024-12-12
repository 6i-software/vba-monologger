VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger"
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
'@FQCN("VBAMonologger.Log.Logger")
'@ModuleDescription("Logger, the main component of logging system.")
''

Option Explicit

Implements VBAMonologger.LoggerInterface

'@Description("The logging channel, a simple descriptive name that is attached to all log records.")
Public name As String

'@Description("Stack of handlers, the first one in the array is called first, etc.")
Private pHandlers As VBA.Collection

'@Description("Stack of processors")
Private pProcessors As VBA.Collection


' ____________ '
'              '
'  Initialize  '
' ____________ '
'              '

Private Sub Class_Initialize()
    name = ""
    Set pHandlers = New VBA.Collection
    Set pProcessors = New VBA.Collection
End Sub


'@Description('Emulate a class constructor.')
Public Function construct( _
    ByVal paramName As String, _
    Optional paramHandlers As VBA.Collection = Nothing, _
    Optional paramProcessors As VBA.Collection = Nothing _
) As VBAMonologger.LoggerInterface
    name = paramName
    Set pHandlers = IIf(Not paramHandlers Is Nothing, paramHandlers, New VBA.Collection)
    Set pProcessors = IIf(Not paramProcessors Is Nothing, paramProcessors, New VBA.Collection)
    
    Set construct = Me
End Function



' _______________ '
'                 '
'  Public method  '
' _______________ '
'                 '

'@Description("Get all handlers referenced into logger.")
Public Property Get handlers() As VBA.Collection
    Set handlers = pHandlers
End Property

'@Description("Set handlers, replacing all existing ones.")
Public Property Set handlers(ByVal newHandlers As VBA.Collection)
    Set pHandlers = newHandlers
End Property

'@Description("Returns True if there is at least one item in the Handlers collection, and False if it’s empty.")
Public Property Get hasHandlers() As Boolean
    hasHandlers = (handlers.count > 0)
End Property

'@Description("Pushes a handler on to the stack.")
Public Function pushHandler(ByVal paramHandler As VBAMonologger.HandlerInterface) As VBAMonologger.LoggerInterface
    pHandlers.Add paramHandler
    Set pushHandler = Me
End Function

'@Description("Pops a handler from the stack.")
Public Function popHandler() As VBAMonologger.LoggerInterface
    If pHandlers.count = 0 Then
        Err.Raise vbObjectError + 1000, "Logger::PopHandler", "You tried to pop from an empty handlers stack."
    End If
    ' Get the last added handler, to remove it
    pHandlers.Remove pHandlers.count
    
    Set popHandler = Me
End Function

'@Description("Get all pre-processors referenced into logger.")
Public Property Get processors() As VBA.Collection
    Set processors = pProcessors
End Property

'@Description("Set pre-processors, replacing all existing ones.")
Public Property Set processors(ByVal newProcessors As VBA.Collection)
    Set pProcessors = newProcessors
End Property

'@Description("Returns True if there is at least one item in the pre-processors collection, and False if it’s empty.")
Public Property Get hasProcessors() As Boolean
    hasProcessors = (processors.count > 0)
End Property

'@Description("Pushes a pre-processor on to the stack.")
Public Function pushProcessor(ByVal paramProcessor As VBAMonologger.ProcessorInterface) As VBAMonologger.LoggerInterface
    pProcessors.Add paramProcessor
    Set pushProcessor = Me
End Function

'@Description("Removes the pre-processor on top of the stack and returns remaining pre-processors.")
Public Function popProcessor() As VBAMonologger.LoggerInterface
    If pProcessors.count = 0 Then
        Err.Raise vbObjectError + 1000, "Logger::PopProcessor", "You tried to pop from an empty processors stack."
    End If
    ' Get the last added handler, to remove it
    pProcessors.Remove pProcessors.count
    
    Set popProcessor = Me
End Function



' ________________________________ '
'                                  '
'  Implementation: LoggerInterface '
' ________________________________ '
'                                  '

'@inheritdoc
Private Sub LoggerInterface_log( _
    ByVal level As VBAMonologger.LOG_LEVELS, _
    ByVal message As String, _
    Optional ByRef context As Scripting.IDictionary = Nothing, _
    Optional ByRef extra As Scripting.IDictionary = Nothing _
)
    ' Initialize log record
    Dim record As VBAMonologger.LogRecord
    Set record = New VBAMonologger.LogRecord
    Set record = record.construct(message, level, Me.name, context, extra)
        
    Dim recordHandled As Boolean
    Dim recordWithProcessorsInvoked As Boolean
    Dim currentHandler As VBAMonologger.HandlerInterface
    Dim currentProcessor As VBAMonologger.ProcessorInterface
    
    recordHandled = False
    recordWithProcessorsInvoked = (hasProcessors() = 0)
    For Each currentHandler In handlers()
        ' Skip invoking processors on the log record as long as no handler is going to handle it
        If currentHandler.isHandling(record) Then
        
            ' Invoking processors on the log record only once in foreach iteration
            If recordWithProcessorsInvoked = False Then
                For Each currentProcessor In processors
                    Call currentProcessor.perform(record)
                Next currentProcessor
                
                ' Ok processors are invoked on the log record
                recordWithProcessorsInvoked = True
            End If
            
            ' Once the log record is initialized, send it to all handlers as long as the bubbling chain is not interrupted
            recordHandled = currentHandler.handle(record)
            
            ' Interrupted the bubbling when the Handle method returning false (see HandlerInterface)
            If Not recordHandled Then
                Exit Sub
            End If
        End If
    Next currentHandler
End Sub

'@inheritdoc
Private Sub LoggerInterface_emergency(ByVal message As String, Optional context As Scripting.IDictionary = Nothing, Optional extra As Scripting.IDictionary = Nothing)
    Call LoggerInterface_log(VBAMonologger.LEVEL_EMERGENCY, message, context, extra)
End Sub

'@inheritdoc
Private Sub LoggerInterface_alert(ByVal message As String, Optional context As Scripting.IDictionary = Nothing, Optional ByRef extra As Scripting.IDictionary = Nothing)
    Call LoggerInterface_log(VBAMonologger.LEVEL_ALERT, message, context, extra)
End Sub

'@inheritdoc
Private Sub LoggerInterface_critical(ByVal message As String, Optional context As Scripting.IDictionary = Nothing, Optional ByRef extra As Scripting.IDictionary = Nothing)
    Call LoggerInterface_log(VBAMonologger.LEVEL_CRITICAL, message, context, extra)
End Sub

'@inheritdoc
Private Sub LoggerInterface_error(ByVal message As String, Optional context As Scripting.IDictionary = Nothing, Optional ByRef extra As Scripting.IDictionary = Nothing)
    Call LoggerInterface_log(VBAMonologger.LEVEL_ERROR, message, context, extra)
End Sub

'@inheritdoc
Private Sub LoggerInterface_warning(ByVal message As String, Optional context As Scripting.IDictionary = Nothing, Optional ByRef extra As Scripting.IDictionary = Nothing)
    Call LoggerInterface_log(VBAMonologger.LEVEL_WARNING, message, context, extra)
End Sub

'@inheritdoc
Private Sub LoggerInterface_notice(ByVal message As String, Optional context As Scripting.IDictionary = Nothing, Optional ByRef extra As Scripting.IDictionary = Nothing)
    Call LoggerInterface_log(VBAMonologger.LEVEL_NOTICE, message, context, extra)
End Sub

'@inheritdoc
Private Sub LoggerInterface_info(ByVal message As String, Optional context As Scripting.IDictionary = Nothing, Optional ByRef extra As Scripting.IDictionary = Nothing)
    Call LoggerInterface_log(VBAMonologger.LEVEL_INFO, message, context, extra)
End Sub

'@inheritdoc
Private Sub LoggerInterface_trace(ByVal message As String, Optional context As Scripting.IDictionary = Nothing, Optional ByRef extra As Scripting.IDictionary = Nothing)
    Call LoggerInterface_log(VBAMonologger.LEVEL_DEBUG, message, context, extra)
End Sub



' _____________________________________________ '
'                                               '
'  Public exposition methods of LoggerInterface '
' _____________________________________________ '
'                                               '

'@Description("Proxy method for public exposition.")
Public Sub log( _
    ByVal level As VBAMonologger.LOG_LEVELS, _
    ByVal message As String, _
    Optional ByRef context As Scripting.IDictionary = Nothing, _
    Optional ByRef extra As Scripting.IDictionary = Nothing _
)
    Call LoggerInterface_log(level, message, context, extra)
End Sub

'@Description("Proxy method for public exposition.")
Public Sub emergency(ByVal message As String, Optional context As Scripting.IDictionary = Nothing, Optional ByRef extra As Scripting.IDictionary = Nothing)
    Call LoggerInterface_emergency(message, context, extra)
End Sub

'@Description("Proxy method for public exposition.")
Public Sub alert(ByVal message As String, Optional context As Scripting.IDictionary = Nothing, Optional ByRef extra As Scripting.IDictionary = Nothing)
    Call LoggerInterface_alert(message, context, extra)
End Sub

'@Description("Proxy method for public exposition.")
Public Sub critical(ByVal message As String, Optional context As Scripting.IDictionary = Nothing, Optional ByRef extra As Scripting.IDictionary = Nothing)
    Call LoggerInterface_critical(message, context, extra)
End Sub

'@Description("Proxy method for public exposition.")
Public Sub error(ByVal message As String, Optional context As Scripting.IDictionary = Nothing, Optional ByRef extra As Scripting.IDictionary = Nothing)
    Call LoggerInterface_error(message, context, extra)
End Sub

'@Description("Proxy method for public exposition.")
Public Sub warning(ByVal message As String, Optional context As Scripting.IDictionary = Nothing, Optional ByRef extra As Scripting.IDictionary = Nothing)
    Call LoggerInterface_warning(message, context, extra)
End Sub

'@Description("Proxy method for public exposition.")
Public Sub notice(ByVal message As String, Optional context As Scripting.IDictionary = Nothing, Optional ByRef extra As Scripting.IDictionary = Nothing)
    Call LoggerInterface_notice(message, context, extra)
End Sub

'@Description("Proxy method for public exposition.")
Public Sub info(ByVal message As String, Optional context As Scripting.IDictionary = Nothing, Optional ByRef extra As Scripting.IDictionary = Nothing)
    Call LoggerInterface_info(message, context, extra)
End Sub

'@Description("Proxy method for public exposition.")
Public Sub trace(ByVal message As String, Optional context As Scripting.IDictionary = Nothing, Optional ByRef extra As Scripting.IDictionary = Nothing)
    Call LoggerInterface_trace(message, context, extra)
End Sub

