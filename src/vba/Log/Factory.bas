Attribute VB_Name = "Factory"
' ------------------------------------- '
'                                       '
'    VBA Monologger                         '
'    Copyright © 2024, 6i software      '
'                                       '
' ------------------------------------- '
'
'@Exposed
'@Folder("VBAMonologger.Log")
'@FQCN("VBAMonologger.Log.Factory")
'@ModuleDescription("To build VBAMonologger.logger instance.")


Option Explicit


' ________ '
'          '
'  Logger  '
' ________ '
'          '

'@Description("Create a simple logger without handlers and pre-processors, as an empty logger")
Public Function createLogger() As VBAMonologger.Logger
    Dim Logger As VBAMonologger.Logger
    Set Logger = New VBAMonologger.Logger
    
    Set createLogger = Logger
End Function

'@Description("Create a logger instance that outputs log messages to the VBA console (Excel's Immediate Window).")
Public Function createLoggerConsoleVBA( _
    Optional ByVal paramLoggerName As String = vbNullString, _
    Optional ByRef paramFormatter As VBAMonologger.FormatterInterface = Nothing _
) As VBAMonologger.LoggerInterface
    Dim Logger As VBAMonologger.Logger
    Set Logger = New VBAMonologger.Logger
    
    ' Naming the logger, i.e. add a channel to the logger
    Logger.name = paramLoggerName
    
    ' Creating a console handler
    Dim handler As VBAMonologger.HandlerInterface
    Set handler = New VBAMonologger.HandlerConsoleVBA
    If (Nothing Is paramFormatter) Then
        Set handler.formatter = New VBAMonologger.FormatterLine
    Else
        Set handler.formatter = paramFormatter
    End If
    
    ' Add console handler into logger
    Call Logger.pushHandler(handler)
    
    ' Add pre-processors placeholders
    Call pushProcessorPlaceholders(Logger)
    
    Set createLoggerConsoleVBA = Logger
End Function


'@Description("Create a logger instance that outputs log messages to a file")
Public Function createLoggerFile( _
    Optional ByVal paramLoggerName As String = vbNullString, _
    Optional ByRef paramFormatter As VBAMonologger.FormatterInterface = Nothing, _
    Optional ByVal paramLogFileName As String = vbNullString, _
    Optional ByVal paramLogFileFolder As String = vbNullString _
) As VBAMonologger.Logger
    Dim Logger As VBAMonologger.Logger
    Set Logger = New VBAMonologger.Logger
    
    ' Naming the logger, i.e. add a channel to the logger
    Logger.name = paramLoggerName
    
    ' Creating a file handler
    Dim HandlerFile As VBAMonologger.HandlerFile
    Set HandlerFile = New VBAMonologger.HandlerFile
    If (Nothing Is paramFormatter) Then
        Set HandlerFile.formatter = New VBAMonologger.FormatterLine
    Else
        Set HandlerFile.formatter = paramFormatter
    End If

    ' Changing the name and the destination of the log file
    If (paramLogFileName <> vbNullString) Then
        HandlerFile.logFileName = paramLogFileName
    End If
    If (paramLogFileFolder <> vbNullString) Then
        HandlerFile.logFileFolder = paramLogFileFolder
    End If
    
    ' Add console handler into logger
    Call Logger.pushHandler(HandlerFile)
    
    ' Add pre-processors placeholders
    Call pushProcessorPlaceholders(Logger)
    
    Set createLoggerFile = Logger
End Function


'@Description("Create a logger instance that outputs log messages to the windows console (by using VBAMonologger HTTP server logs viewer).")
Public Function createLoggerConsole( _
    Optional ByVal paramLoggerName As String = vbNullString, _
    Optional ByRef paramFormatter As VBAMonologger.FormatterInterface = Nothing, _
    Optional ByRef paramWithANSIColorSupport As Boolean = True, _
    Optional ByRef paramWithNewlineForContextAndExtra As Boolean = True, _
    Optional ByRef paramWithProcessorUID As Boolean = False, _
    Optional ByRef paramWithProcessorUsageMemory As Boolean = False, _
    Optional ByRef paramWithProcessorUsageCPU As Boolean = False, _
    Optional ByRef paramWithDebugServer As Boolean = False, _
    Optional ByRef paramWithDebugClient As Boolean = False _
) As VBAMonologger.Logger
    Dim Logger As VBAMonologger.Logger
    Set Logger = New VBAMonologger.Logger
    
    ' Naming the logger, i.e. add a channel to the logger
    Logger.name = paramLoggerName
    
    ' Creating a console handler
    Dim handler As VBAMonologger.HandlerConsole
    Set handler = New VBAMonologger.HandlerConsole
    handler.withDebug = paramWithDebugClient
    
    If Not (Nothing Is paramFormatter) Then
        Set handler.formatter = paramFormatter
    Else
        ' Configure the default line formatter (if no formatter given)
        handler.withANSIColorSupport = paramWithANSIColorSupport
        handler.withNewlineForContextAndExtra = paramWithNewlineForContextAndExtra
    End If
    
    ' Start the VBAMonloger HTTP server logs viewer (with or without debug for dev only)
    handler.startServerLogsViewer paramWithDebugServer
    
    ' Add console handler into logger
    Call Logger.pushHandler(handler)
    
    ' Add pre-processors placeholders
    Call pushProcessorPlaceholders(Logger)
    If paramWithProcessorUID Then Call pushProcessorUID(Logger)
    If paramWithProcessorUsageCPU Then Call pushProcessorUsageCPU(Logger)
    If paramWithProcessorUsageMemory Then Call pushProcessorUsageMemory(Logger)
           
    Set createLoggerConsole = Logger
End Function

'@Description("Create a null logger instance that does nothing!")
Public Function createLoggerNull() As VBAMonologger.LoggerInterface
    Dim Logger As VBAMonologger.LoggerInterface
    Set Logger = New VBAMonologger.LoggerNull
    
    Set createLoggerNull = Logger
End Function



' ___________ '
'             '
'  Formatter  '
' ___________ '
'             '

'@Description("Create a FormatterLine instance.")
Public Function createFormatterLine( _
    Optional paramTemplateLine As String = vbNullString, _
    Optional paramWhitespace As Variant = Nothing _
) As VBAMonologger.FormatterLine
    Dim formatter As New VBAMonologger.FormatterLine
    Set formatter = formatter.construct(paramTemplateLine, paramWhitespace)
    
    Set createFormatterLine = formatter
End Function

'@Description("Create a FormatterANSIColoredLine instance.")
Public Function createFormatterANSIColoredLine( _
    Optional paramTemplateLine As String = vbNullString _
) As VBAMonologger.FormatterANSIColoredLine
    Dim formatter As New VBAMonologger.FormatterANSIColoredLine
    Set formatter = formatter.construct(paramTemplateLine)
    
    Set createFormatterLine = formatter
End Function



' _________ '
'           '
'  Handler  '
' _________ '
'           '

'@Description("Create a HandlerConsoleVBA instance.")
Public Function createHandlerConsoleVBA( _
    Optional ByVal paramBubble As Boolean = True, _
    Optional ByVal paramLevel As VBAMonologger.LOG_LEVELS = VBAMonologger.LOG_LEVELS.LEVEL_DEBUG _
) As VBAMonologger.HandlerConsoleVBA
    Dim handler As New VBAMonologger.HandlerConsoleVBA
    Set createHandlerConsoleVBA = handler.construct(paramBubble, paramLevel)
End Function

'@Description("Create a HandlerFile instance.")
Public Function createHandlerFile( _
    Optional ByVal paramBubble As Boolean = True, _
    Optional ByVal paramLevel As VBAMonologger.LOG_LEVELS = VBAMonologger.LOG_LEVELS.LEVEL_DEBUG, _
    Optional ByVal paramLogFileName As String = vbNullString, _
    Optional ByVal paramLogFileFolder As String = vbNullString _
) As VBAMonologger.HandlerFile
    Dim handler As New VBAMonologger.HandlerFile
    Set createHandlerFile = handler.construct(paramBubble, paramLevel, paramLogFileName, paramLogFileFolder)
End Function

'@Description("Create a HandlerConsole instance.")
Public Function createHandlerConsole( _
    Optional ByVal paramBubble As Boolean = True, _
    Optional ByVal paramLevel As VBAMonologger.LOG_LEVELS = VBAMonologger.LOG_LEVELS.LEVEL_DEBUG, _
    Optional ByVal paramHostnameServer As String = "localhost", _
    Optional ByVal paramPortServer As Integer = 20100 _
) As VBAMonologger.HandlerConsole
    Dim handler As New VBAMonologger.HandlerConsole
    Set createHandlerConsole = handler.construct(paramBubble, paramLevel, paramHostnameServer, paramPortServer)
End Function



' _______________ '
'                 '
'  Pre-processor  '
' _______________ '
'                 '

'@Description("Push a new pre-processors into given logger that allow to use placeholders in log message, according to PSR-3 standard.")
Public Sub pushProcessorPlaceholders(ByRef Logger As VBAMonologger.Logger)
    Dim processor As New VBAMonologger.ProcessorPlaceholders
    Logger.pushProcessor processor
End Sub

'@Description("Push a new UID pre-processors into given logger, in order to add a generated and random UID into log records.")
Public Sub pushProcessorUID( _
    ByRef Logger As VBAMonologger.Logger, _
    Optional ByVal paramLengthUID As Integer = 10 _
)
    Dim processor As New VBAMonologger.ProcessorUid
    Call processor.setLengthUid(paramLengthUID)
    
    Logger.pushProcessor processor
End Sub

'@Description("Push a new memory usage pre-processors into given logger.")
Public Sub pushProcessorUsageMemory( _
    ByRef Logger As VBAMonologger.Logger, _
    Optional ByVal paramWithDetails As Boolean = False _
)
    Dim processor As New VBAMonologger.ProcessorUsageMemory
    processor.withDetails = paramWithDetails
    
    Logger.pushProcessor processor
End Sub

'@Description("Push a new CPU usage pre-processors into given logger.")
Public Sub pushProcessorUsageCPU(ByRef Logger As VBAMonologger.Logger)
    Dim processor As New VBAMonologger.ProcessorUsageCPU
    Logger.pushProcessor processor
End Sub

'@Description("Push a new Tags pre-processors into given logger.")
Public Sub pushProcessorTags( _
    ByRef Logger As VBAMonologger.Logger, _
    Optional paramTags As Scripting.IDictionary = Nothing, _
    Optional paramTagsDestination As TAGS_DESTINATION = TAGS_DESTINATION.LOG_CONTEXT _
)
    Dim processor As New VBAMonologger.ProcessorTags
    processor.tagsDestination = paramTagsDestination
    Call processor.setTags(paramTags)

    Logger.pushProcessor processor
End Sub
