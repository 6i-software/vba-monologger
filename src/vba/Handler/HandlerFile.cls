VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HandlerFile"
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
'@FQCN("VBAMonologger.Handler.HandlerFile")
'@ModuleDescription("Handle log to a file.")
''

Option Explicit

Implements VBAMonologger.HandlerInterface
Implements VBAMonologger.StringableInterface

'@Description("The associated formatter of current handler.")
Private pformatter As VBAMonologger.FormatterInterface
'@Description("The bubbling behavior. True means that this handler allows bubbling, and false means that bubbling is not permitted.")
Private pBubble As Boolean
'@Description("The minimum logging level at which this handler will be triggered as a class module LogLevel")
Private pLogLevel As VBAMonologger.LogLevel
'@Description("The minimum logging level at which this handler will be triggered as an value of enumeration of LOG_LEVELS.")
Private pLevel As VBAMonologger.LOG_LEVELS

'@Description("The name of the logfile.")
Private pLogFileName As String
'@Description("The full path where the logfile is save.")
Private pLogfilePath As String
'@Description("The folder where the logfile is save.")
Private pLogFileFolder As String
'@Description("Denotes the folder is created.")
Private pFolderCreated As Variant
'@Description("To interact with file system (create folder, delete file...).")
Private pFileSystemObject As New Scripting.FileSystemObject
'@Description("To write file in UTF-8 with ADODB.Stream object available in Microsoft ActiveX Data Objects.")
Private pLogFileStream As Object

Private pLockFile As Boolean


' ____________ '
'              '
'  Initialize  '
' ____________ '
'              '

Private Sub Class_Initialize()
    ' Default propagation enable
    pBubble = True
    
    ' Default logging level
    Set pLogLevel = New VBAMonologger.LogLevel
    pLogLevel.currentLogLevel = VBAMonologger.LOG_LEVELS.LEVEL_DEBUG
    pLevel = VBAMonologger.LOG_LEVELS.LEVEL_DEBUG
        
    ' Default formatter
    Set pformatter = New VBAMonologger.FormatterLine
    
    ' Default path where logfile is stored.
    logFileName = "logfile_" & format(Now, "yyyy-mm-dd") & ".log"
    logFileFolder = ThisWorkbook.Path & "\var\log"
    pFolderCreated = Null
    
    ' Configure the default stream
    Set pLogFileStream = CreateObject("ADODB.Stream")
    With pLogFileStream
        .Type = 2 ' ADODB.Stream type texte
        .Charset = "UTF-8" ' Default encoding
    End With
    
    pLockFile = False
End Sub

'@Description('Emulate a class constructor.')
Public Function construct( _
    Optional ByVal paramBubble As Boolean = True, _
    Optional ByVal paramLevel As VBAMonologger.LOG_LEVELS = VBAMonologger.LOG_LEVELS.LEVEL_DEBUG, _
    Optional ByVal paramLogFileName As String = vbNullString, _
    Optional ByVal paramLogFileFolder As String = vbNullString _
) As VBAMonologger.HandlerInterface
    pBubble = paramBubble
    pLogLevel.currentLogLevel = paramLevel
    pLevel = paramLevel
    
    If (paramLogFileName <> vbNullString) Then
        logFileName = paramLogFileName
    End If
    If (paramLogFileFolder <> vbNullString) Then
        logFileFolder = paramLogFileFolder
        pFolderCreated = Null
    End If
        
    Set construct = Me
End Function


' ___________________ '
'                     '
'  Public properties  '
' ___________________ '
'                     '

'@Description("Get the name of the log file.")
Public Property Get logFileName() As String
    logFileName = pLogFileName
End Property

'@Description("Set a new name for the log file.")
Public Property Let logFileName(ByVal newLogFileName As String)
    If Not pIsValidFileName(newLogFileName) Then
        Err.Raise vbObjectError + 1001, "VBAMonologger.Handler.HandlerFile::SetLogFileName", "The new name for logfile (" & newLogFileName & ") contains invalid characters !"
    End If
    
    pLogFileName = newLogFileName
    pLogfilePath = pLogFileFolder & "\" & pLogFileName
End Property

'@Description("Check the validity of a file name.")
Private Function pIsValidFileName(ByVal fileName As String) As Boolean
    Dim invalidChars As String
    Dim i As Integer
    
    ' Set unvalid characters for file name
    invalidChars = "\/:*?""<>|"
    For i = 1 To Len(invalidChars)
        If InStr(fileName, Mid(invalidChars, i, 1)) > 0 Then
            pIsValidFileName = False
            Exit Function
        End If
    Next i
    
    pIsValidFileName = True
End Function

'@Description("Get the folder where the log file is stored.")
Public Property Get logFileFolder() As String
    logFileFolder = pLogFileFolder
End Property

'@Description("Set the folder where the log file is stored.")
Public Property Let logFileFolder(ByVal newLogFileFolder As String)
    Dim pathParts() As String
    Dim currentPath As String
    Dim i As Integer
    
    ' Check validity of each part of logFileFolder
    On Error Resume Next
    pathParts = Split(newLogFileFolder, "\")
    For i = 0 To UBound(pathParts)
        If Not pIsValidFileName(pathParts(i)) Then
            Err.Raise vbObjectError + 1002, "VBAMonologger.Handler.HandlerFile::LetLogFileFolder", "The new folder for logfile (" & newLogFileFolder & ") contains invalid characters in part (" & pathParts(i) & ")!"
        End If
    Next i
    On Error GoTo 0

    pLogFileFolder = newLogFileFolder
    pLogfilePath = newLogFileFolder & "\" & pLogFileName
    pFolderCreated = Null
End Property

'@Description("Get the full path where the log file is stored.")
Public Property Get logFilePath() As String
    logFilePath = pLogfilePath
End Property





' _________________________________ '
'                                   '
'  Implementation: HandlerInterface '
' _________________________________ '
'                                   '

'@inheritdoc
Private Property Get HandlerInterface_formatter() As VBAMonologger.FormatterInterface
    Set HandlerInterface_formatter = pformatter
End Property

'@inheritdoc
Private Property Set HandlerInterface_formatter(ByRef newFormatter As VBAMonologger.FormatterInterface)
    Set pformatter = newFormatter
End Property

'@inheritdoc
Private Property Get HandlerInterface_bubble() As Boolean
    HandlerInterface_bubble = pBubble
End Property

'@inheritdoc
Private Property Let HandlerInterface_bubble(ByVal value As Boolean)
    pBubble = value
End Property

'@inheritdoc
Private Property Get HandlerInterface_level() As VBAMonologger.LOG_LEVELS
    HandlerInterface_level = pLogLevel.currentLogLevel
End Property

'@inheritdoc
Private Property Let HandlerInterface_level(ByVal newLevel As VBAMonologger.LOG_LEVELS)
    pLogLevel.currentLogLevel = newLevel
    pLevel = newLevel
End Property

'@inheritdoc
Private Function HandlerInterface_isHandling(paramLogRecord As VBAMonologger.LogRecordInterface) As Boolean
    ' To avoid calling processors for nothing.
    HandlerInterface_isHandling = pLogLevel.includes(paramLogRecord.level)
End Function

'@inheritdoc
Private Function HandlerInterface_handle(paramLogRecord As LogRecordInterface) As Boolean
    Dim result As Boolean
    Dim formattedMessage As String
    Dim lines() As String
    Dim i As Long
    
    If Not HandlerInterface_isHandling(paramLogRecord) Then
        result = False
    Else
        ' Create folder if does not exist
        pCreateFolder logFileFolder
        
        ' Ignore errors if logfile does not exist
        On Error Resume Next
            pLogFileStream.Open
            pLogFileStream.LoadFromFile logFilePath
            pLogFileStream.Position = pLogFileStream.Size
        On Error GoTo 0
        
        ' Format the message
        formattedMessage = formatter.format(paramLogRecord)
        
       ' Process the message to split it when a \n (newline) is encountered.
        lines = Split(formattedMessage, "\n")
        
        ' Write each line to the log, adding an empty line for each \n
        For i = LBound(lines) To UBound(lines)
            pLogFileStream.WriteText lines(i), 1 ' 1: Writes a text string and a line separator character to stream
        Next i
        
        ' Write new content into stream at position 0 if file does not exist or at the end if file exists
        ' pLogFileStream.WriteText formatter.format(paramLogRecord), 1 ' 1: Writes a text string and a line separator character to stream
        pLogFileStream.SaveToFile logFilePath, 2 ' 2: Overwrite file
        
        result = True
    End If
    
    result = (True = pBubble)
    
    HandlerInterface_handle = result
End Function

'@inheritdoc
Private Function HandlerInterface_handleBatch(paramLogRecords() As LogRecordInterface) As Boolean
    Dim i As Integer
    Dim record As VBAMonologger.LogRecordInterface
    Dim result As Boolean
    
    HandlerInterface_handleBatch = True
    For i = LBound(paramLogRecords) To UBound(paramLogRecords)
        Set record = paramLogRecords(i)
        result = HandlerInterface_handle(record)

        If result = False Then
            HandlerInterface_handleBatch = False
        End If
    Next i
End Function

'@inheritdoc
Private Sub HandlerInterface_closeHandler()
    pLogFileStream.Close
End Sub

'@Description("Creates folder and its parent folders if they does not exist, where the log file is stored.")
Private Sub pCreateFolder(ByVal newLogFileFolder As String)
    ' To skip creation
    If pFolderCreated = True Then
        Exit Sub
    End If
    
    ' Try to create folders recursvily (i.e. create all sub-folders of logFileFolder)
    Dim pathParts() As String
    Dim currentPath As String
    Dim i As Integer
    On Error Resume Next
    pathParts = Split(newLogFileFolder, "\")
    currentPath = pathParts(0)
    For i = 1 To UBound(pathParts)
        currentPath = currentPath & "\" & pathParts(i)
        If Not pFileSystemObject.FolderExists(currentPath) Then
            pFileSystemObject.CreateFolder currentPath
        End If
    Next i
    On Error GoTo 0

    ' Check final folder
    If Not pFileSystemObject.FolderExists(newLogFileFolder) Then
        pFolderCreated = False
        Err.Raise vbObjectError + 1003, "VBAMonologger.Handler.HandlerFile::SetLogFileFolder", "Failed to create the new folder (" & newLogFileFolder & ") !"
    Else
        pFolderCreated = True
    End If
End Sub

'@Description("Proxy method for public exposition.")
Public Property Get formatter() As VBAMonologger.FormatterInterface
    Set formatter = pformatter
End Property

'@Description("Proxy method for public exposition.")
Public Property Set formatter(ByRef newFormatter As VBAMonologger.FormatterInterface)
    Set pformatter = newFormatter
End Property

'@Description("Proxy method for public exposition.")
Public Property Get bubble() As Boolean
    bubble = pBubble
End Property

'@Description("Proxy method for public exposition.")
Public Property Let bubble(ByVal value As Boolean)
    pBubble = value
End Property

'@Description("Proxy method for public exposition.")
Public Function isHandling(paramLogRecord As VBAMonologger.LogRecordInterface) As Boolean
    isHandling = HandlerInterface_isHandling(paramLogRecord)
End Function

'@Description("Proxy method for public exposition.")
Public Function handle(paramLogRecord As LogRecordInterface) As Boolean
    handle = HandlerInterface_handle(paramLogRecord)
End Function

'@Description("Proxy method for public exposition.")
Public Function handleBatch(paramLogRecords() As LogRecordInterface) As Boolean
    handleBatch = HandlerInterface_handleBatch(paramLogRecords)
End Function

'@Description("Proxy method for public exposition.")
Public Sub closeHandler()
    HandlerInterface_closeHandler
End Sub

'@Description("Proxy method for public exposition.")
Public Property Get level() As VBAMonologger.LOG_LEVELS
    level = pLogLevel.currentLogLevel
End Property

'@Description("Proxy method for public exposition.")
Public Property Let level(ByVal newLevel As VBAMonologger.LOG_LEVELS)
    pLogLevel.currentLogLevel = newLevel
    pLevel = newLevel
End Property


' ____________________________________ '
'                                      '
'  Implementation: StringableInterface '
' ____________________________________ '
'                                      '

'@inheritdoc
Private Function StringableInterface_toString() As String
    Dim output As String
    output = "[VBAMonologger.Handler.HandlerFile]" & vbCrLf
    output = output & " | level: " & pLogLevel.name & vbCrLf
    output = output & " | bubble: " & pBubble & vbCrLf
    output = output & " | logFileName: " & logFileName & vbCrLf
    output = output & " | logfilePath: " & logFilePath & vbCrLf
    output = output & " | logFileFolder: " & logFileFolder & vbCrLf

    StringableInterface_toString = output
End Function

'@Description("Proxy method for public exposition.")
Public Function toString() As String
    toString = StringableInterface_toString
End Function
