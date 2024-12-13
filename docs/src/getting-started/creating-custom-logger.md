

### Create a custom logger from scratch

In this example, we create an empty logger, and we push into multiples handlers with differnts formatters.
    - a handler for ouptut log into console VBA
    - a handler for ouptut log into console
    - a handler for ouptut log into file only for error log records (level >= error)

```vbscript
Public Sub howto_create_a_custom_Logger_from_scratch()
    Dim customLogger As VBAMonologger.Logger
    Set customLogger = VBAMonologger.Factory.createLogger
    
    ' Create a custom line formatter
    Dim customFormatterLine As VBAMonologger.FormatterLine
    Set customFormatterLine = VBAMonologger.Factory.createFormatterLine
    customFormatterLine.showContext = True
    customFormatterLine.showExtra = True
    customFormatterLine.withAllowingInlineLineBreaks = False
    customFormatterLine.templateLine = ":: {{ channel }}{{ level_name }} - {{ message }}"
  
    ' Add a custom console VBA handler (use custom formatter)
    Dim customHandlerConsoleVBA As VBAMonologger.HandlerConsoleVBA
    Set customHandlerConsoleVBA = VBAMonologger.Factory.createHandlerConsoleVBA
    Set customHandlerConsoleVBA.formatter = customFormatterLine
    customLogger.pushHandler customHandlerConsoleVBA
    
    ' Add a custom console handler (use default formatter with ANSI support)
    Dim customHandlerConsole As VBAMonologger.HandlerConsole
    Set customHandlerConsole = VBAMonologger.Factory.createHandlerConsole
    customHandlerConsole.portServer = 20101
    customHandlerConsole.hostnameServer = "127.0.0.1"
    customHandlerConsole.withANSIColorSupport = True
    customHandlerConsole.withDebug = False
    customHandlerConsole.withNewlineForContextAndExtra = True
    customHandlerConsole.startServerLogsViewer
    customLogger.pushHandler customHandlerConsole
    
    ' Add a custom file handler (use a custom formatter and capture only error log records, i.e. with level >= LEVEL_ERROR)
    Dim customHandlerFile As VBAMonologger.handlerFile
    Set customHandlerFile = VBAMonologger.Factory.createHandlerFile
    customHandlerFile.logFileName = "error_" & Format(Now, "yyyy-mm-dd") & ".log"
    customHandlerFile.logFileFolder = ThisWorkbook.Path & "\logs"
    customHandlerFile.Level = LEVEL_ERROR
    Dim formatter As VBAMonologger.FormatterLine
    Set formatter = customHandlerFile.formatter
    formatter.setTemplateLineWithNewlineForContextAndExtra
    formatter.withWhitespace = True
    formatter.withAllowingInlineLineBreaks = True
    customLogger.pushHandler customHandlerFile
    
    ' Add pre-processors
    VBAMonologger.Factory.pushProcessorPlaceholders customLogger
    VBAMonologger.Factory.pushProcessorUID customLogger, 8
    VBAMonologger.Factory.pushProcessorUsageCPU customLogger
    VBAMonologger.Factory.pushProcessorUsageMemory customLogger
    Dim tags As Object
    Set tags = CreateObject("Scripting.Dictionary")
    tags.Add "environment", "production"
    VBAMonologger.Factory.pushProcessorTags customLogger, tags, TAGS_DESTINATION.LOG_EXTRA
    
    ' Use the custom logger
    customLogger.trace "Authentication function call for user 'Bob Morane'." ' The 'debug' method exposes presents in PSR-3 is rename into 'trace' in order to be compatible in VBA ecosystem
    customLogger.info "User 'Ultra Vomit' has logged in successfully."
    customLogger.notice "Process completed successfully with minor issues."
    customLogger.warning "The user 'Beetlejuice' should not be called more than 3 times."
    customLogger.Error "An error occurred when the user 'DeadRobotZombieCopFromOuterspace' tried to read the dashboard file."
    customLogger.critical "System is in an unstable state. Unable to authenticate the user 'Skjalg Skagen'."
    customLogger.alert "Action required: unable to generate the dashboard."
    customLogger.emergency "A critical failure occurred in the application for moving files."
    
    Dim context As Object: Set context = CreateObject("Scripting.Dictionary")
    context.Add "UserName", "Bob Morane"
    context.Add "UserID", 342527
    customLogger.trace "Authentication function call for user '{UserName}' with id '{UserID}'.", context
    customLogger.Error "User id '{UserID}' does not exist. Unable to create dashboard file.", context
End Sub
```

![VBAMonologger-multiples-handlers.png](VBAMonologger-multiples-handlers.png)

