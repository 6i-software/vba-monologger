Attribute VB_Name = "Usecases"
' ------------------------------------- '
'                                       '
'    VBA Monologger                     '
'    Copyright © 2024, 6i software      '
'                                       '
' ------------------------------------- '
'
'@Exposed
'@Folder("VBAMonologger.Tests")
'@FQCN("VBAMonologger.Tests.Usecases")
'@ModuleDescription("Use cases of VBAMonologger for testing purposes.")
''

Option Explicit


' API Windows
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If


Private Sub run_all_usecases()
    Usecases_LogLevel
    Usecases_LogRecord
    Usecases_FormatterLine
    Usecases_HandlerConsoleVBA
    Usecases_HandlerFile
    Usecases_FormatterANSIColoredLine
    Usecases_Logger
    Usecases_LoggerWithProcessors
    Usecases_HandlerConsole
    Usecases_FactoryCreateLoggerConsoleVBA
    Usecases_FactoryCreateLoggerFile
    Usecases_FactoryCreateLoggerConsole
    Usecases_DummyClassWithLoggerAware
    Usecases_DummyClassWithLoggerNull
End Sub

Private Sub Usecases_DummyClassWithLoggerNull()
    UsecasePrintTitle ("DummyClassWithLoggerNull")
    
    Dim dummyClass As DummyClassUsingLoggerAware
    Set dummyClass = New DummyClassUsingLoggerAware
    
    ' Create a null logger
    Dim Logger As New VBAMonologger.LoggerNull
    
    ' Inject logger into class
    Call dummyClass.setLogger(Logger)
    
    ' In dummyAction we use the logger, but in this case, nothing happens! That’s normal because it’s a null logger.
    Debug.Print "=== Use the logger by dependencies injection but in this case, nothing happens! That’s normal because it’s a null logger ==="
    dummyClass.foo
    Debug.Print ""

    UsecasePrintEnd
End Sub

Private Sub Usecases_DummyClassWithLoggerAware()
    UsecasePrintTitle ("DummyClassWithLoggerAware")
    
    Dim dummyClass As DummyClassUsingLoggerAware
    Set dummyClass = New DummyClassUsingLoggerAware
    
    ' Create a custom logger
    Dim Logger As VBAMonologger.Logger
    Set Logger = VBAMonologger.Factory.createLoggerConsoleVBA("App")
    
    ' Inject logger into class
    Call dummyClass.setLogger(Logger)
    
    ' In foo method we use the logger
    Debug.Print "=== Use the logger by dependencies injection (custom class implements LoggerAwareInterface) ==="
    dummyClass.foo
    Debug.Print ""

    UsecasePrintEnd
End Sub

Private Sub Usecases_FactoryCreateLoggerConsole()
    UsecasePrintTitle ("VBAMonologger.Factory.createLoggerConsole")
    
    Dim Logger As VBAMonologger.Logger
    Set Logger = VBAMonologger.Factory.createLoggerConsole("App", Nothing, False, False, False)
    
    ' Fetch the handlerConsole in order to use sendRequest procedure
    Dim handler As VBAMonologger.HandlerConsole
    Set handler = Logger.handlers(1)

    ' Test with no ANSI support and no newline for context and extra informations
    handler.sendRequest vbCrLf & ANSI.BG_YELLOW & ANSI.BLACK & "--- Test logger console with no ANSI support and no newlines for context and extra informations. ---" & ANSI.RESET & vbCrLf
    Call useLoggerForEachLevel(Logger, False) ' Use logger methods for all levels (logger.info, logger.notice, logger.notice...)
    
    ' Test with no ANSI support and newlines for context and extra informations
    Set Logger = VBAMonologger.Factory.createLoggerConsole("App", Nothing, False, True)
    handler.sendRequest vbCrLf & vbCrLf & ANSI.BG_YELLOW & ANSI.BLACK & "--- Test logger console with no ANSI support and newlines for context and extra informations. ---" & ANSI.RESET & vbCrLf
    Call useLoggerForEachLevel(Logger, False) ' Use logger methods for all levels (logger.info, logger.notice, logger.notice...)
    
    ' Test with ANSI support and no newline for context and extra informations
    Set Logger = VBAMonologger.Factory.createLoggerConsole("App", Nothing, True, False)
    handler.sendRequest vbCrLf & vbCrLf & ANSI.BG_YELLOW & ANSI.BLACK & "--- Test logger console with no ANSI support and no newlines for context and extra informations. ---" & ANSI.RESET & vbCrLf
    Call useLoggerForEachLevel(Logger, False) ' Use logger methods for all levels (logger.info, logger.notice, logger.notice...)

    ' Test with ANSI support and newline for context and extra infos (default mode of handlerConsole)
    Set Logger = VBAMonologger.Factory.createLoggerConsole("App", Nothing, True, True)
    handler.sendRequest vbCrLf & vbCrLf & ANSI.BG_YELLOW & ANSI.BLACK & "--- Test logger console with ANSI support and newlines for context and extra informations. ---" & ANSI.RESET & vbCrLf
    Call useLoggerForEachLevel(Logger, False)
    
    Logger.emergency "A critical failure occurred in the application for {Operation} operation."
    Logger.alert "Action required for {Operation} process can not be done."
    
    ' Wait 1 seconds and stop the server
    handler.sendRequest vbCrLf & vbCrLf & ANSI.BG_YELLOW & ANSI.BLACK & "--- Wait 1 seconds before stopping the server ---" & ANSI.RESET & vbCrLf
    Sleep 1000
    ' Send command to stop the server
    handler.sendRequest "stop-server"
    ' Send keyboard-command instruction "exit" in order to close the active console
    handler.sendExitCommand

    UsecasePrintEnd
End Sub

Private Sub Usecases_FactoryCreateLoggerConsoleWithVerboseMode()
    UsecasePrintTitle ("VBAMonologger.Factory.createLoggerConsole with verbose mode")
    
    Dim Logger As VBAMonologger.Logger
    Set Logger = VBAMonologger.Factory.createLoggerConsole("App", Nothing, True, False, True, True)
    
    ' Fetch the handlerConsole in order to use sendRequest procedure
    Dim handler As VBAMonologger.HandlerConsole
    Set handler = Logger.handlers(1)

    ' Test with no ANSI support and no newline for context and extra informations
    handler.sendRequest vbCrLf & ANSI.BG_YELLOW & ANSI.BLACK & "--- Test logger console with verbose mode on server and client. ---" & ANSI.RESET & vbCrLf
    Call useLoggerForEachLevel(Logger) ' Use logger methods for all levels (logger.info, logger.notice, logger.notice...)
        
    UsecasePrintEnd
End Sub

Private Sub Usecases_FactoryCreateLoggerFile()
    UsecasePrintTitle ("VBAMonologger.Factory.createLoggerFile")
    
    Dim Logger As VBAMonologger.Logger
    Set Logger = VBAMonologger.Factory.createLoggerFile("App")
    Call useLoggerForEachLevel(Logger)
    Debug.Print "See the content of `./var/log/logfile_xxxx-yy-zz.log` file!"
    
    UsecasePrintEnd
End Sub

Private Sub Usecases_FactoryCreateLoggerConsoleVBA()
    UsecasePrintTitle ("VBAMonologger.Factory.createLoggerConsoleVBA")
    
    Dim Logger As VBAMonologger.Logger
    Set Logger = VBAMonologger.Factory.createLoggerConsoleVBA("App")
    Call useLoggerForEachLevel(Logger)
    
    UsecasePrintEnd
End Sub

Private Sub Usecases_HandlerConsole()
    UsecasePrintTitle ("VBAMonologger.HandlerConsole")
    
    Debug.Print "=== Create a new HandlerConsole with a formatterLine ==="
    Dim handler As VBAMonologger.HandlerConsole
    Set handler = New VBAMonologger.HandlerConsole
    Dim handlerClone As VBAMonologger.HandlerConsole
    Set handlerClone = handler
    Debug.Print handlerClone.toString
            
    ' Start the server
    handler.startServerLogsViewer True
        
    ' Send request tot the server
    handler.sendRequest "I belive I can fly (éèöôàç) !"
    handler.sendRequest "Really?"
    
    ' Handle a single log record
    Debug.Print "=== Handle a log record with HandlerConsole ==="
    Dim dummyRecord As VBAMonologger.LogRecord
    Set dummyRecord = randomLogRecord(VBAMonologger.LEVEL_NOTICE)
    handler.handle dummyRecord
    
    ' Handle a collection fo log records
    Debug.Print "=== Handle log records collection ==="
    Dim records() As VBAMonologger.LogRecordInterface
    ReDim records(1 To 8)
    records = randomLogRecordsForEachLevel
    handler.handleBatch records
    
    ' Send command to stop the server
    handler.sendRequest "stop-server"
    
    ' Send keyboard-command instruction "exit" in order to close the active console
    handler.sendExitCommand
    
    UsecasePrintEnd
End Sub

Private Sub Usescases_ConsoleWrapper()
    Dim console As ConsoleWrapper
    Set console = New ConsoleWrapper
    Dim pTempFolderPowershellPrograms As String
    Dim pPowershellProgramServerFilepath As String
    pTempFolderPowershellPrograms = Environ("TEMP") & "\VBAMonologger\powershell"
    pPowershellProgramServerFilepath = pTempFolderPowershellPrograms & "\VBAMonologgerHTTPServerLogsViewer.ps1"
    
    Dim command As String
    command = "cmd.exe /K"
    command = command & " powershell.exe -File """ & pPowershellProgramServerFilepath & """"
    command = command & " -port 20100"
    command = command & " -hostname localhost"
    command = command & " -titleConsole ""VBAMonologger server logs viewer"""
    command = command & " -Verbose"
        
    console.withDebug = True
    console.titleConsoleWindow = "VBAMonologger server logs viewer"
    
    ' The console is created only once. If a console window with that title already exists, it does nothing!
    console.createConsole command
End Sub

Private Sub Usecases_LoggerWithProcessors()
    UsecasePrintTitle ("VBAMonologger.Processor")
    
    ' Creating a logger
    Dim Logger As VBAMonologger.Logger
    Set Logger = New VBAMonologger.Logger
    Logger.name = "App"
                
    ' Creating and customizing a formatter line
    Dim lineFormatter As VBAMonologger.FormatterLine
    Set lineFormatter = New VBAMonologger.FormatterLine
    lineFormatter.setTemplateLineWithNewlineForContextAndExtra
    lineFormatter.withWhitespace = True
    lineFormatter.withAllowingInlineLineBreaks = True
                
    ' Add console handler into logger
    Dim consoleHandler As VBAMonologger.HandlerConsoleVBA
    Set consoleHandler = New VBAMonologger.HandlerConsoleVBA
    Set consoleHandler.formatter = lineFormatter
    Logger.pushHandler consoleHandler
    
    ' Add file handler into logger
    Dim fileHandler As VBAMonologger.HandlerFile
    Set fileHandler = New VBAMonologger.HandlerFile
    fileHandler.logFileName = "logfileWithProcessors.log"
    Set fileHandler.formatter = lineFormatter
    Logger.pushHandler fileHandler
    
    ' Add pre-processors "Placeholders" into logger
    Dim preprocessorPlaceholders As VBAMonologger.ProcessorPlaceholders
    Set preprocessorPlaceholders = New VBAMonologger.ProcessorPlaceholders
    
    ' Add pre-processors "UIG" into logger
    Dim preprocessorUid As VBAMonologger.ProcessorUid
    Set preprocessorUid = New VBAMonologger.ProcessorUid
    Call preprocessorUid.setLengthUid(10)
        
    ' Add pre-processors "Tags" into logger
    Dim preprocessorTags As VBAMonologger.ProcessorTags
    Set preprocessorTags = New VBAMonologger.ProcessorTags
    Dim tags As Object
    Set tags = CreateObject("Scripting.Dictionary")
    tags.Add "environment", "production"
    tags.Add "user_role", "admin"
    Call preprocessorTags.addTags(tags)
    preprocessorTags.tagsDestination = LOG_CONTEXT
    preprocessorTags.keepKeyTags = False
    
    ' Add pre-processors "UsageMemory" into logger
    Dim preprocessorUsageMemory As VBAMonologger.ProcessorUsageMemory
    Set preprocessorUsageMemory = New VBAMonologger.ProcessorUsageMemory
    preprocessorUsageMemory.withDetails = True
    
    ' Add pre-processors "UsageCPU" into logger
    Dim preprocessorUsageCPU As VBAMonologger.ProcessorUsageCPU
    Set preprocessorUsageCPU = New VBAMonologger.ProcessorUsageCPU
    
    Logger.pushProcessor preprocessorPlaceholders
    Logger.pushProcessor preprocessorUid
    Logger.pushProcessor preprocessorTags
    Logger.pushProcessor preprocessorUsageMemory
    Logger.pushProcessor preprocessorUsageCPU
    
    ' Perform logger's method
    Call useLoggerForEachLevel(Logger, False)
    Debug.Print ""
    
    UsecasePrintEnd
End Sub

Private Sub Usecases_Logger()
    UsecasePrintTitle ("VBAMonologger.Log.Logger")
    
    Dim Logger As VBAMonologger.Logger
    Set Logger = New VBAMonologger.Logger
    
    ' Change logger channel name
    Logger.name = "App"
    
    ' Create a default formatter
    Dim formatter As VBAMonologger.FormatterInterface
    Set formatter = New VBAMonologger.FormatterLine
        
    ' Add handler console for all levels
    Dim handler As VBAMonologger.HandlerConsoleVBA
    Set handler = New VBAMonologger.HandlerConsoleVBA
    Set handler.formatter = formatter
    Logger.pushHandler handler
    
    ' Add handler for alert and emergency log (with no propagation)
    Dim handlerFileAlert As VBAMonologger.HandlerFile
    Set handlerFileAlert = New VBAMonologger.HandlerFile
    Set handlerFileAlert.formatter = formatter
    handlerFileAlert.logFileName = "alert.log"
    handlerFileAlert.level = LOG_LEVELS.LEVEL_ALERT
    handlerFileAlert.bubble = False
    Logger.pushHandler handlerFileAlert

    ' Add handler file for another levels
    Dim HandlerFile As VBAMonologger.HandlerFile
    Set HandlerFile = New VBAMonologger.HandlerFile
    Set HandlerFile.formatter = formatter
    HandlerFile.logFileName = "debug.log"
    Logger.pushHandler HandlerFile
    
    Debug.Print "Logger has handlers? " & Logger.hasHandlers
    Debug.Print ""
        
    ' Use logger's method
    Call useLoggerForEachLevel(Logger)
    Debug.Print ""
    
    UsecasePrintEnd
End Sub

Private Sub Usecases_FormatterANSIColoredLine()
    UsecasePrintTitle ("VBAMonologger.Formatter.FormatterANSIColoredLine")
    
    Debug.Print "=== Create a new FormatterANSIColoredLine (with default line template) ==="
    Dim FormatterANSIColoredLine As VBAMonologger.FormatterANSIColoredLine
    Set FormatterANSIColoredLine = New VBAMonologger.FormatterANSIColoredLine
    Debug.Print FormatterANSIColoredLine.toString
    
    Debug.Print "=== Output record minimal ==="
    Dim DummyRecord1 As VBAMonologger.LogRecord
    Set DummyRecord1 = New VBAMonologger.LogRecord
    Set DummyRecord1 = DummyRecord1.construct( _
        "I believe I can fly", _
        VBAMonologger.LEVEL_INFO _
    )
    Debug.Print FormatterANSIColoredLine.format(DummyRecord1) & vbCrLf
    
    Debug.Print "=== Output by using FormatBatch on records collection ==="
    Dim records() As VBAMonologger.LogRecordInterface
    ReDim records(1 To 8)
    records = randomLogRecordsForEachLevel
    Debug.Print " >>> Result of formatterLine(records)"
    Debug.Print FormatterANSIColoredLine.formatBatch(records) & vbCrLf
    
    Debug.Print "=== Output records with HandlerFile ==="
    Dim handler As VBAMonologger.HandlerInterface
    Set handler = New VBAMonologger.HandlerFile
    Debug.Print " >>> Ok, HandlerFile is just created."
    Set handler.formatter = FormatterANSIColoredLine
    Debug.Print " >>> Ok, a FormatterANSIColoredLine is added into HandlerFile as default formatter."
    Dim handlerClone As VBAMonologger.HandlerFile
    Set handlerClone = handler
    handlerClone.logFileName = "ANSI-colorized-log-file-name.log"
    Dim resultHandleBatch As Boolean
    resultHandleBatch = handler.handleBatch(randomLogRecordsForEachLevel)
    Debug.Print " >>> Result of handler.handleBatch(records): " & resultHandleBatch
    Debug.Print ""
    
    Debug.Print "=== Change color scheme for output records with HandlerFile ==="
    Set FormatterANSIColoredLine.colorScheme = FormatterANSIColoredLine.getTrafficLightColorScheme
    FormatterANSIColoredLine.showContext = False
    FormatterANSIColoredLine.showExtra = False
    resultHandleBatch = handler.handleBatch(randomLogRecordsForEachLevel)
    Debug.Print " >>> Result of handler.handleBatch(records): " & resultHandleBatch
    Debug.Print ""
End Sub


Private Sub Usecases_HandlerFile()
    UsecasePrintTitle ("VBAMonologger.HandlerFile")
    
    Debug.Print "=== Create a new HandlerFile with a formatterLine ==="
    Dim handler As VBAMonologger.HandlerInterface
    Set handler = New VBAMonologger.HandlerFile
    Debug.Print " >>> Ok, HandlerFile is just created."
    Set handler.formatter = New VBAMonologger.FormatterLine
    Debug.Print " >>> Ok, a new formatterLine is added into HandlerFile as default formatter."
    Dim handlerClone As VBAMonologger.HandlerFile
    Set handlerClone = handler
    Debug.Print handlerClone.toString
    
    Debug.Print "=== Change the logfile name ==="
    handlerClone.logFileName = "Amazing-log-file-name.log"
    Debug.Print handlerClone.toString
    
    Debug.Print "=== Change the logfile folder ==="
    handlerClone.logFileFolder = VBA.Environ$("USERPROFILE") & "\VBAMonologger\logs"
    Debug.Print handlerClone.toString
    
    Debug.Print "=== Initialize handlerFile with custom logfile name and folder ==="
    Set handlerClone = handlerClone.construct( _
         paramLogFileName:="log-file___" & format(Now, "yyyy-mm-dd") & ".log", _
         paramLogFileFolder:=ThisWorkbook.Path & "\var\log" _
    )
    Debug.Print handlerClone.toString
    
    Debug.Print "=== Handle a log record with HandlerFile ==="
    Dim dummyRecord As New VBAMonologger.LogRecord
    Set dummyRecord = dummyRecord.construct( _
        "I believe I can fly!", _
        VBAMonologger.LEVEL_NOTICE, _
        "App.BusinessIntelligence" _
    )
    Dim isHandling As Boolean
    isHandling = handler.isHandling(dummyRecord)
    Debug.Print " >>> Result of handler.isHandling(dummyRecord): " & isHandling
    Dim resultHandle As Boolean
    resultHandle = handler.handle(dummyRecord)
    Debug.Print " >>> Result of handler.handle(dummyRecord): " & resultHandle
    Debug.Print ""
    
    Debug.Print "=== Handle a collection of log records with handlerFile ==="
    Debug.Print " >>> Result of handler.handleBatch(records): "
    Dim resultHandleBatch As Boolean
    resultHandleBatch = handler.handleBatch(randomLogRecordsForEachLevel)
    Debug.Print " >>> Result of handler.handleBatch(records): " & resultHandleBatch
    Debug.Print ""
    
    Debug.Print "=== Change level of handler to LEVEL_ERROR ==="
    handler.level = VBAMonologger.LOG_LEVELS.LEVEL_ERROR
    resultHandleBatch = handler.handleBatch(randomLogRecordsForEachLevel)
    Debug.Print " >>> Result of handler.handleBatch(records): " & resultHandleBatch
    Set handlerClone = handler
    Debug.Print handlerClone.toString
    
    Debug.Print "=== Close the handler ==="
    handlerClone.closeHandler
    Debug.Print " >>> Ok, the handler is closed." & vbCrLf

    UsecasePrintEnd
End Sub

Private Sub Usecases_HandlerConsoleVBA()
    UsecasePrintTitle ("VBAMonologger.HandlerConsoleVBA")
    
    Debug.Print "=== Create a new HandlerConsoleVBA with a formatterLine ==="
    Dim handler As VBAMonologger.HandlerConsoleVBA
    Set handler = New VBAMonologger.HandlerConsoleVBA
    Debug.Print " >>> Ok, HandlerConsoleVBA is just created."
    Set handler.formatter = New VBAMonologger.FormatterLine
    Debug.Print " >>> Ok, a new formatterLine is added into HandlerConsoleVBA as default formatter."
    Dim handlerClone As VBAMonologger.HandlerConsoleVBA
    Set handlerClone = handler
    Debug.Print handlerClone.toString
        
    Debug.Print "=== Handle a log record with HandlerConsoleVBA ==="
    Dim dummyRecord As VBAMonologger.LogRecord
    Set dummyRecord = randomLogRecord(VBAMonologger.LEVEL_NOTICE)
    Dim isHandling As Boolean
    isHandling = handler.isHandling(dummyRecord)
    Debug.Print " >>> Result of handler.isHandling(dummyRecord): " & isHandling
    Debug.Print " >>> Result of handler.handle(dummyRecord): "
    handler.handle dummyRecord
    Debug.Print ""
    
    Debug.Print "=== Handle a collection of log records with HandlerConsoleVBA ==="
    Debug.Print " >>> Result of handler.handleBatch(records): "
    handler.handleBatch randomLogRecordsForEachLevel
    Debug.Print ""
    
    Debug.Print "=== Change level of handler to LEVEL_CRITICAL ==="
    handler.level = VBAMonologger.LOG_LEVELS.LEVEL_CRITICAL
    Debug.Print " >>> Result of handler.handleBatch(dummyLogRecords): "
    handler.handleBatch randomLogRecordsForEachLevel
    Set handlerClone = handler
    Debug.Print handlerClone.toString

    Debug.Print "=== Close the handler ==="
    handler.closeHandler
    Debug.Print " >>> Ok, the handler is closed." & vbCrLf
    
    UsecasePrintEnd
End Sub


Private Sub Usecases_FormatterLine()
    UsecasePrintTitle ("VBAMonologger.Formatter.FormatterLine")
    
    Debug.Print "=== Create a new FormatterLine (with default line template) ==="
    Dim formatter As VBAMonologger.FormatterLine
    Set formatter = New VBAMonologger.FormatterLine
    Debug.Print formatter.toString
    
    Debug.Print "=== Output record minimal ==="
    Dim dummyRecord As VBAMonologger.LogRecord
    Set dummyRecord = New VBAMonologger.LogRecord
    Set dummyRecord = dummyRecord.construct( _
        "I believe I can fly", _
        VBAMonologger.LEVEL_INFO _
    )
    Debug.Print formatter.format(dummyRecord) & vbCrLf
    
    Debug.Print "=== Output record with channel name ==="
    Set dummyRecord = dummyRecord.construct( _
        "I believe I can fly!", _
        VBAMonologger.LEVEL_NOTICE, _
        "App.BusinessIntelligence" _
    )
    Debug.Print formatter.format(dummyRecord) & vbCrLf
       
    Debug.Print "=== Output record with log context and extra metadatas ==="
    Set dummyRecord = randomLogRecord(VBAMonologger.LEVEL_EMERGENCY)
    Debug.Print formatter.format(dummyRecord) & vbCrLf
    
    Debug.Print "=== Output by using FormatBatch on records collection ==="
    Dim records() As VBAMonologger.LogRecordInterface
    ReDim records(1 To 8)
    records = randomLogRecordsForEachLevel
    Debug.Print " >>> Result of formatterLine(records)"
    Debug.Print formatter.formatBatch(records) & vbCrLf

    Debug.Print "=== Remove context in line formatted (showContext is False) ==="
    formatter.showContext = False
    Debug.Print formatter.format(dummyRecord) & vbCrLf
    
    Debug.Print "=== Remove extra in line formatted (showExtra is False) ==="
    formatter.showContext = True
    formatter.showExtra = False
    Debug.Print formatter.format(dummyRecord) & vbCrLf
    
    Debug.Print "=== Remove context and extra in line formatted ==="
    formatter.showContext = False
    formatter.showExtra = False
    Debug.Print formatter.format(dummyRecord) & vbCrLf

    Debug.Print "=== Change the line template of FormatterLine (which contains premotif and postmotif placeholder) ==="
    formatter.templateLine = "[{{datetime}}] {{ <Chanel:/>channel< ; /> }}{{ level_name }}: {{ message }}{{< ; ctx: /> context}}{{< ; extra: /> extra < ;/> }}"
    formatter.showContext = True
    formatter.showExtra = True
    Debug.Print formatter.toString

    Debug.Print "===  Output records with this new line template ==="
    Debug.Print " >>> Result of formatterLine(records)"
    Debug.Print formatter.formatBatch(records) & vbCrLf
    
    Debug.Print "=== Change the line template of FormatterLine (which contains premotif and postmotif placeholder) ==="
    formatter.templateLine = "{{ <Chanel:/>channel< ; /> }}{{ level_name }}: {{ message }}}"
    formatter.showContext = True
    formatter.showExtra = True
    Debug.Print formatter.toString

    Debug.Print "===  Output records with this new line template ==="
    Debug.Print " >>> Result of formatterLine(records)"
    Debug.Print formatter.formatBatch(records) & vbCrLf
    
    UsecasePrintEnd
End Sub


Private Sub Usecases_LogRecord()
    UsecasePrintTitle ("LogRecord")
    
    Debug.Print "=== Init LogRecord with minimal parameters ==="
    Dim record As VBAMonologger.LogRecord
    Set record = New VBAMonologger.LogRecord
    Set record = record.construct( _
        "I believe I can fly", _
        VBAMonologger.LEVEL_EMERGENCY _
    )
    Debug.Print record.toString
    

    Debug.Print "=== With channel naming ==="
    Set record = record.construct( _
        "I believe I can fly", _
        VBAMonologger.LEVEL_EMERGENCY, _
        "Channel one" _
    )
    Debug.Print record.toString
    
    
    Debug.Print "=== With log context and extra metadatas ==="
    Set record = randomLogRecord(VBAMonologger.LEVEL_EMERGENCY)
    Debug.Print record.toString
    
    
    Debug.Print "=== With a nested dictionnary in log context ==="
    Set record = record.construct( _
        "I believe I can fly", _
        VBAMonologger.LEVEL_EMERGENCY, _
        "Channel one", _
        randomStudents _
    )
    Debug.Print record.toString
    
    UsecasePrintEnd
End Sub


Private Sub Usecases_LogLevel()
    UsecasePrintTitle ("LogLevel")
    
    Dim testLevel As VBAMonologger.LogLevel
    Set testLevel = New LogLevel
    
    Debug.Print "=== Init the LogLevel with 'Info' level ==="
    testLevel.construct LEVEL_INFO
    Debug.Print testLevel.toString
    
    Debug.Print "=== Change level to 'Notice' ==="
    testLevel.currentLogLevel = VBAMonologger.LOG_LEVELS.LEVEL_NOTICE
    Debug.Print testLevel.toString
    
    Debug.Print "=== Change level with fromName methode to 'Warning' ==="
    Call testLevel.fromName("WARNING")
    Debug.Print testLevel.toString
    
    Debug.Print "=== Comparaison beetwen two levels ==="
    Dim otherLevel As LogLevel
    Set otherLevel = New LogLevel
    
    ' Compare WARNING with INFO
    Debug.Print "> Compare WARNING with INFO"
    otherLevel.construct LOG_LEVELS.LEVEL_INFO
    Debug.Print "Is the current level '" & testLevel.name & "' is higher than '" & otherLevel.name & "'? "; testLevel.isHigherThan(otherLevel) ' Result: True
    Debug.Print "Is the current level '" & testLevel.name & "' is lower than '" & otherLevel.name & "'? "; testLevel.isLowerThan(otherLevel) ' Result: False
    Debug.Print "Is the current level '" & testLevel.name & "' includes '" & otherLevel.name & "'? "; testLevel.includes(otherLevel) ' Result: False
    Debug.Print ""
    
    ' Compare WARNING with ALERT
    Debug.Print "> Compare WARNING with ALERT"
    otherLevel.currentLogLevel = LOG_LEVELS.LEVEL_ALERT
    Debug.Print "Is the current level '" & testLevel.name & "' is higher than '" & otherLevel.name & "'? "; testLevel.isHigherThan(otherLevel) ' Result: True
    Debug.Print "Is the current level '" & testLevel.name & "' is lower than '" & otherLevel.name & "'? "; testLevel.isLowerThan(otherLevel) ' Result: False
    Debug.Print "Is the current level '" & testLevel.name & "' includes '" & otherLevel.name & "'? "; testLevel.includes(otherLevel) ' Result: False
    Debug.Print ""
    
    UsecasePrintEnd
End Sub




' ____________________________ '
'                              '
'  Helpers for usecase runner  '
' _____________________________'
'                              '
Private Sub UsecasePrintTitle(title As String)
    Dim boxWidth As Integer
    Dim borderLine As String
    Dim titleLine As String
    Dim spaceBeforeTitle As Integer
    Dim spaceAfterTitle As Integer
    
    title = "Usecases - " & title
    boxWidth = Len(title) + 4
    borderLine = " " & String(boxWidth, "_")
    spaceBeforeTitle = (boxWidth - Len(title)) \ 2
    spaceAfterTitle = boxWidth - Len(title) - spaceBeforeTitle
    titleLine = "|" & String(spaceBeforeTitle, " ") & title & String(spaceAfterTitle, " ") & "|"
    
    Debug.Print borderLine
    Debug.Print "|" & String(boxWidth, " ") & "|"
    Debug.Print titleLine
    Debug.Print "|" & String(boxWidth, "_") & "|"
    Debug.Print ""
    Debug.Print "-------------------------------------------------<<< BEGIN >>>-------------------------------------------------" & vbCrLf
End Sub

Private Sub UsecasePrintEnd()
    Debug.Print "--------------------------------------------------<<< END >>>--------------------------------------------------" & vbCrLf
End Sub




' __________'
'           '
'  Fixtures '
' _________ '
'           '

Public Function randomLogRecordsForEachLevel() As VBAMonologger.LogRecordInterface()
    Dim record As VBAMonologger.LogRecordInterface
    Dim records() As VBAMonologger.LogRecordInterface

    Dim levels As Variant
    levels = Array( _
        VBAMonologger.LEVEL_EMERGENCY, _
        VBAMonologger.LEVEL_ALERT, _
        VBAMonologger.LEVEL_CRITICAL, _
        VBAMonologger.LEVEL_ERROR, _
        VBAMonologger.LEVEL_WARNING, _
        VBAMonologger.LEVEL_NOTICE, _
        VBAMonologger.LEVEL_INFO, _
        VBAMonologger.LEVEL_DEBUG _
    )
    
    ReDim records(LBound(levels) To UBound(levels))
    
    Dim i As Integer
    For i = LBound(levels) To UBound(levels)
        Set records(i) = randomLogRecord(levels(i))
    Next i

    randomLogRecordsForEachLevel = records
End Function

Public Function randomLogRecord(level As Variant) As VBAMonologger.LogRecord
    Dim record As VBAMonologger.LogRecord
    Set record = New VBAMonologger.LogRecord
        
    Dim logMessage As String
    Select Case level
        Case VBAMonologger.LEVEL_EMERGENCY
            logMessage = "A critical failure occurred in the application for {Operation} process"
        Case VBAMonologger.LEVEL_ALERT
            logMessage = "Action required for process {Operation} failure."
        Case VBAMonologger.LEVEL_CRITICAL
            logMessage = "System is in an unstable state. Unable to authenticate {UserId}."
        Case VBAMonologger.LEVEL_ERROR
            logMessage = "An error occurred when the user {UserId} try to {Operation} the file {file}."
        Case VBAMonologger.LEVEL_WARNING
            logMessage = "The user {UserId} does not exist. Unable to perform '{Operation}' user file."
        Case VBAMonologger.LEVEL_NOTICE
            logMessage = "Process completed successfully with minor issues for {UserId}."
        Case VBAMonologger.LEVEL_INFO
            logMessage = "User {UserId} has logged in successfully."
        Case VBAMonologger.LEVEL_DEBUG
            logMessage = "Authentification function call for user {UserId}."
    End Select
    
    Set randomLogRecord = record.construct( _
        logMessage, _
        level, _
        "App.Authentification", _
        randomLogContext, _
        randomLogExtra _
    )
End Function

Public Function randomLogContext() As Scripting.Dictionary
    Dim dummyContext As Scripting.Dictionary
    Set dummyContext = New Scripting.Dictionary
    
    Dim availableUserName As Variant
    Dim availableOperations As Variant
    availableUserName = Array("Bob", "Alice", "Arthur", "v20100v", "CaravanPalace", "2o8o")
    availableOperations = Array("create", "read", "update", "delete")
    
    Randomize
    dummyContext.Add "UserName", availableUserName(Int(Rnd * (UBound(availableUserName) + 1)))
    dummyContext.Add "UserID", Int((99999 - 10000 + 1) * Rnd + 10000)
    dummyContext.Add "Operation", availableOperations(Int(Rnd * (UBound(availableOperations) + 1)))

    Set randomLogContext = dummyContext
End Function

Public Function randomLogExtra() As Scripting.Dictionary
    Dim dummyExtra As Scripting.Dictionary
    Set dummyExtra = New Scripting.Dictionary
    
    Randomize
    dummyExtra.Add "ExecutionTime", Round((Rnd * 10), 4) & " seconds"
        
    Set randomLogExtra = dummyExtra
End Function

Public Function randomStudents(Optional count As Integer = 3) As Scripting.Dictionary
    Dim students As Scripting.Dictionary
    Set students = New Scripting.Dictionary
    
    Dim fullNames As Variant
    fullNames = Array( _
        "Isaac Newton", "Albert Einstein", "Galileo Galilei", "Pythagore", _
        "Alan Turing", "Stephen Hawking", "Marie Curie", "Leonhard Euler", _
        "Pierre de Fermat", "Bernhard Riemann", "Nicolaus Copernicus", "Pierre-Simon Laplace", _
        "Niels Bohr", "Paul Dirac", "Felix Klein", "Nikola Tesla", _
        "Werner Heisenberg", "Max Planck", "Hendrik Lorentz", "Richard Feynman", _
        "James Clerk Maxwell", "Joseph-Louis Lagrange", "Évariste Galois", _
        "Charles Darwin", "Gregor Mendel", "James Watson", "Francis Crick", _
        "Richard Dawkins", "Jane Goodall", "Carl Sagan", "Rita Levi-Montalcini" _
    )
    Dim randomFullName As String
    randomFullName = fullNames(Int(Rnd * (UBound(fullNames) + 1)))
    
    Dim subjects As Variant
    subjects = Array( _
        "Physics", "Mathematics", "Biology", "Chemistry", "Astrophysics", _
        "Computer Science", "Statistics", "Genetics", "Quantum Mechanics", _
        "Thermodynamics", "Ecology", "Evolutionary Biology", "Organic Chemistry", _
        "Linear Algebra", "Calculus", "Differential Equations", "Biochemistry" _
    )
     
    Dim student As Scripting.Dictionary
    Dim grades As Scripting.Dictionary
    Dim selectedSubjects As Collection
    Dim randomNumberSubjects As Integer
    Dim randomSubject As String

    Dim i As Integer
    Dim j As Integer
    For i = 1 To count
        Set student = New Scripting.Dictionary
        student.Add "FullName", randomFullName
        student.Add "Age", Int((55 - 18 + 1) * Rnd + 16)
        
        Set grades = New Scripting.Dictionary
        ' Add random number of subjects (2 at 6)
        randomNumberSubjects = Int((6 - 2 + 1) * Rnd + 1)
        Set selectedSubjects = New Collection
        Do While selectedSubjects.count < randomNumberSubjects
            randomSubject = subjects(Int(Rnd * (UBound(subjects) + 1)))
            ' Avoid if subject is already selected
            On Error Resume Next
            selectedSubjects.Add randomSubject, randomSubject
            On Error GoTo 0
        Loop
        For j = 1 To selectedSubjects.count
            Dim subjectName As String
            subjectName = selectedSubjects(j)
            grades.Add subjectName, Int((100 - 11 + 1) * Rnd + 50)
        Next j
        
        student.Add "Grades", grades
        students.Add "Student" & i, student
    Next i

    Set randomStudents = students
End Function

Public Sub useLoggerForEachLevel(Logger As VBAMonologger.LoggerInterface, Optional withRandomLogExtra As Boolean = True)
    Debug.Print "=== Calling logger's method for each level ==="
    
    Logger.emergency "A critical failure occurred in the application for {Operation} operation.", randomLogContext, IIf(withRandomLogExtra, randomLogExtra, Nothing)
    Logger.alert "Action required for {Operation} process can not be done.", randomLogContext, IIf(withRandomLogExtra, randomLogExtra, Nothing)
    Logger.critical "System is in an unstable state. Unable to authenticate the user '{UserName}' with id '{UserID}'.", randomLogContext, IIf(withRandomLogExtra, randomLogExtra, Nothing)
    Logger.error "An error occurred when the user '{UserID}' try to {Operation} the dashboard file.", randomLogContext, IIf(withRandomLogExtra, randomLogExtra, Nothing)
    Logger.warning "The user '{UserName}' does not exist. Unable to perform {Operation} dashboard file.", randomLogContext, IIf(withRandomLogExtra, randomLogExtra, Nothing)
    Logger.notice "Process completed successfully with minor issues for the user '{UserName}'.", randomLogContext, IIf(withRandomLogExtra, randomLogExtra, Nothing)
    Logger.info "User '{UserID}' has logged in successfully.", randomLogContext, IIf(withRandomLogExtra, randomLogExtra, Nothing)
    Logger.trace "Authentification function call for user '{UserID}'.", randomLogContext, IIf(withRandomLogExtra, randomLogExtra, Nothing)
End Sub
