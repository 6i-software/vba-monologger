VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HandlerConsoleNamedPipeServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ------------------------------------- '
'                                       '
'    VBA Monologger                     '
'    Copyright � 2024, 6i software      '
'                                       '
' ------------------------------------- '
'
'@Deprecated("Replace by a HTTP server into the HandlerConsole")
'@Exposed
'@Folder("VBAMonologger.Handler")
'@FQCN("VBAMonologger.Handler.HandlerConsole")
'@ModuleDescription("Handles logs message to Windows console (cmd.exe) by streaming logs beetwen a named piped stream (architecture client/server).")
''

Option Explicit

Implements VBAMonologger.HandlerInterface
Implements VBAMonologger.StringableInterface

'@Description("The associated formatter of current handler.")
Private pformatter As VBAMonologger.FormatterInterface
'@Description("The bubbling behavior. True means that this handler allows bubbling, and false means that bubbling is not permitted.")
Private pBubble As Boolean
'@Description("The minimum logging level at which this handler will be triggered.")
Public logLevel As VBAMonologger.logLevel
'@Description("To interact with file system (create folder, delete file...).")
Private pFileSystemObject As New Scripting.FileSystemObject
'@Description("Folder where powershell programs, for creating server logs viewer and client logs sender, are generated and stored.")
Private pTempFolderPowershellPrograms As String
'@Description("File which contains the powershell code source of 'VBAMonologger server logs viewer' program")
Private pPowershellProgramServerFilepath As String
'@Description("File which contains the powershell code source of 'VBAMonologger client log sender' program")
Private pPowershellProgramClientFilepath As String
'@Description("To write file in UTF-8 with ADODB.Stream object available in Microsoft ActiveX Data Objects.")
Private pWriterFileStream As Object

' ____________ '
'              '
'  Initialize  '
' ____________ '
'              '
Private Sub Class_Initialize()
    ' Default propagation enable
    pBubble = True
    
    ' Default logging level
    Set logLevel = New VBAMonologger.logLevel
    logLevel.currentLogLevel = VBAMonologger.LOG_LEVELS.LEVEL_DEBUG
    
    ' Default formatter
    Set pformatter = New VBAMonologger.FormatterLine
    
    ' Configure the writer stream
    Set pWriterFileStream = CreateObject("ADODB.Stream")
    With pWriterFileStream
        .Type = 2 ' ADODB.Stream type texte
        .Charset = "UTF-8" ' Default encoding
    End With
    
    ' Set temp folder where saved powershell programs
    pTempFolderPowershellPrograms = Environ("TEMP") & "\VBAMonologger\powershell"
    pPowershellProgramServerFilepath = pTempFolderPowershellPrograms & "\VBAMonologgerServerLogsViewer.ps1"
    pPowershellProgramClientFilepath = pTempFolderPowershellPrograms & "\VBAMonologgerClientLogSender.ps1"
    
    ' Create temp folder used for saving powershell program files (server & client)
    createPowershellTempfolder
    
    ' Create PowerShell program files (if they do not exist)
    createPowershellFileVBAMonologgerServerLogsViewer
    createPowershellFileVBAMonologgerClientLogSender
End Sub

'@Description('Emulate a class constructor.')
Public Function construct( _
    Optional ByVal paramBubble As Boolean = True, _
    Optional ByVal paramLevel As VBAMonologger.LOG_LEVELS = VBAMonologger.LOG_LEVELS.LEVEL_DEBUG _
) As VBAMonologger.HandlerInterface
    pBubble = paramBubble
    logLevel.currentLogLevel = paramLevel
    
    Set construct = Me
End Function



' __________________________________ '
'                                    '
'  VBAMonologger server logs viewer  '
' __________________________________ '
'                                    '

'@Desciption("Create the temp folder for save powershell program files")
Private Sub createPowershellTempfolder()
    Dim pathParts() As String
    Dim currentPath As String
    Dim i As Integer
    
    ' Try to create folders recursvily (i.e. create all sub-folders of pFolderPowershellPrograms)
    On Error Resume Next
    pathParts = Split(pTempFolderPowershellPrograms, "\")
    currentPath = pathParts(0)
    For i = 1 To UBound(pathParts)
        currentPath = currentPath & "\" & pathParts(i)
        If Not pFileSystemObject.FolderExists(currentPath) Then
            pFileSystemObject.CreateFolder currentPath
        End If
    Next i
    On Error GoTo 0
    
    ' Check final folder
    If Not pFileSystemObject.FolderExists(pTempFolderPowershellPrograms) Then
        Err.Raise vbObjectError + 1003, "VBAMonologger.Handler.HandlerConsole::createPowershellTempfolder", "Failed to create the temp folder (" & pTempFolderPowershellPrograms & ") !"
    Else
        Debug.Print "Temp powershell folder successfully created in: " & pTempFolderPowershellPrograms
    End If
End Sub

Private Sub createPowershellFileVBAMonologgerServerLogsViewer()
    Dim serverFile As Object
    Dim lines() As String
    Dim i As Long
    
    ' Create file code source, only if does not exist
    If Not pFileSystemObject.FileExists(pPowershellProgramServerFilepath) Then
        ' Get powershell program of 'VBAMonologger server logs viewer'
        Dim serverPowershellCode As String
        serverPowershellCode = getPowershellCodeServerLogsViewer()
                
        ' Ignore errors if logfile does not exist
        On Error Resume Next
            pWriterFileStream.Open
            pWriterFileStream.LoadFromFile pPowershellProgramServerFilepath
        On Error GoTo 0
        
        ' Split code source when a \n (newline) is encountered.
        lines = Split(serverPowershellCode, "\n")
        ' Write each line code, adding an empty line for each \n
        For i = LBound(lines) To UBound(lines)
            pWriterFileStream.WriteText lines(i), 1 ' 1: Writes a text string and a line separator character to stream
        Next i
        pWriterFileStream.SaveToFile pPowershellProgramServerFilepath, 2 ' 2: Overwrite file

        ' Final check
        If Not pFileSystemObject.FileExists(pPowershellProgramServerFilepath) Then
            Err.Raise vbObjectError + 1003, "VBAMonologger.Handler.HandlerConsole::createPowershellFileVBAMonologgerServerLogsViewer", "Failed to create the powershell file: " & pPowershellProgramServerFilepath
        Else
            Debug.Print "Powershell file for VBAMonologger server logs viewer was successfully created in: " & pPowershellProgramServerFilepath
        End If
    Else
        Debug.Print "Powershell file for VBAMonologger server logs viewer already exists in: " & pPowershellProgramServerFilepath
    End If
End Sub

Private Sub createPowershellFileVBAMonologgerClientLogSender()
    Dim clientFile As Object
    Dim lines() As String
    Dim i As Long
    
    ' Create file code source, only if does not exist
    If Not pFileSystemObject.FileExists(pPowershellProgramClientFilepath) Then
        ' Get powershell client code source
        Dim clientPowershellCode As String
        clientPowershellCode = getPowershellCodeClientLogsSender()
                
        ' Ignore errors if logfile does not exist
        On Error Resume Next
            pWriterFileStream.Open
            pWriterFileStream.LoadFromFile pPowershellProgramClientFilepath
        On Error GoTo 0
        
        ' Split code source when a \n (newline) is encountered.
        lines = Split(clientPowershellCode, "\n")
        ' Write each line code, adding an empty line for each \n
        For i = LBound(lines) To UBound(lines)
            pWriterFileStream.WriteText lines(i), 1 ' 1: Writes a text string and a line separator character to stream
        Next i
        pWriterFileStream.SaveToFile pPowershellProgramClientFilepath, 2 ' 2: Overwrite file

        ' Final check
        If Not pFileSystemObject.FileExists(pPowershellProgramClientFilepath) Then
            Err.Raise vbObjectError + 1003, "VBAMonologger.Handler.HandlerConsole::createPowershellFileVBAMonologgerServerLogsViewer", "Failed to create the powershell fil: " & pPowershellProgramClientFilepath
        Else
            Debug.Print "Powershell file for VBAMonologger client log sender was successfully created in: " & pPowershellProgramClientFilepath
        End If
    Else
        Debug.Print "Powershell file for VBAMonologger client log sender already exists in: " & pPowershellProgramClientFilepath
    End If
End Sub

Public Sub startVBAMonologgerServerLogsViewer( _
    Optional ByVal paramVerbose As Boolean = False _
)
    Dim command As String
    command = "cmd.exe /K"
    command = command & " powershell.exe -File """ & pPowershellProgramServerFilepath & """"
    If paramVerbose = True Then
        command = command & " -Verbose"
    End If
    
    Debug.Print command
    shell command, vbMaximizedFocus
End Sub

Public Sub startVBAMonologgerClientLogSender( _
    paramMessage As String, _
    Optional ByVal paramVerbose As Boolean = False _
)
    'Dim command As String
    'command = "cmd.exe /K"
    'command = command & " powershell.exe -File """ & pPowershellProgramClientFilepath & """"
    'command = command & " -Message """ & paramMessage & """"
    'If paramVerbose = True Then
    '    command = command & " -Verbose"
    '    Debug.Print command
    '    shell command, vbMaximizedFocus
    'Else
    '    shell command, vbHide
    'End If
    
    Dim command As String
    Dim shell As Object
    Dim returnCode As Integer
    
    command = "powershell.exe -File """ & pPowershellProgramClientFilepath & """ -Message """ & paramMessage & """"
    Set shell = CreateObject("WScript.Shell")
    returnCode = shell.Run(command, 0, True)
    
End Sub


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
    HandlerInterface_level = logLevel.currentLogLevel
End Property

'@inheritdoc
Private Property Let HandlerInterface_level(ByVal newLevel As VBAMonologger.LOG_LEVELS)
    logLevel.currentLogLevel = newLevel
End Property

'@inheritdoc
Private Property Get HandlerInterface_logLevel() As VBAMonologger.logLevel
    Set HandlerInterface_logLevel = logLevel
End Property

'@inheritdoc
Private Property Set HandlerInterface_logLevel(ByVal newLogLevel As VBAMonologger.logLevel)
    Set logLevel = newLogLevel
End Property

'@inheritdoc
Private Function HandlerInterface_isHandling(paramLogRecord As VBAMonologger.LogRecordInterface) As Boolean
    ' To avoid calling processors for nothing.
    HandlerInterface_isHandling = logLevel.includes(paramLogRecord.level)
End Function

'@inheritdoc
Private Function HandlerInterface_handle(paramLogRecord As LogRecordInterface) As Boolean
    Dim result As Boolean
    If Not HandlerInterface_isHandling(paramLogRecord) Then
        result = False
    Else
        Dim formattedMessage As String
        formattedMessage = VBA.CStr$(pformatter.format(paramLogRecord))
        formattedMessage = VBA.Replace$(formattedMessage, "\n", vbCrLf)
        
        ' Call VBAMonologger client logs sender in order to make a request to VBAMonologger server logs viewer (i.e. send a message to show into server's output)
        
        ' Debug.Print formattedMessage
        
        
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
    ' Nothing to release
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
    level = logLevel.currentLogLevel
End Property

'@Description("Proxy method for public exposition.")
Public Property Let level(ByVal newLevel As VBAMonologger.LOG_LEVELS)
    logLevel.currentLogLevel = newLevel
End Property



' ____________________________________ '
'                                      '
'  Implementation: StringableInterface '
' ____________________________________ '
'                                      '

'@inheritdoc
Private Function StringableInterface_toString() As String
    Dim output As String
    output = "[VBAMonologger.Handler.HandlerConsole]" & vbCrLf
    output = output & " | level: " & logLevel.name & vbCrLf
    output = output & " | bubble: " & pBubble & vbCrLf
    
    StringableInterface_toString = output
End Function

'@Description("Proxy method for public exposition.")
Public Function toString() As String
    toString = StringableInterface_toString
End Function



' ___________________________________________________________________________ '
'                                                                             '
'  Powershell code for VBAMonologger server logs viewer & client logs sender  '
' ___________________________________________________________________________ '
'                                                                             '
Private Function getPowershellCodeServerLogsViewer()
    Dim code As String
    code = ""
    
    ' Header code source file
    code = code & "# ------------------------------------- #" & vbCrLf
    code = code & "#                                       #" & vbCrLf
    code = code & "#    VBA Monologger                     #" & vbCrLf
    code = code & "#    Copyright � 2024, 6i software      #" & vbCrLf
    code = code & "#                                       #" & vbCrLf
    code = code & "# ------------------------------------- #" & vbCrLf
    code = code & "#" & vbCrLf
    code = code & "# Server for displaying log messages sent by clients using a named pipe stream." & vbCrLf
    code = code & "#" & vbCrLf
    code = code & "# The server will do nothing more than display the log message sent by the client in the console." & vbCrLf
    code = code & "#" & vbCrLf
    code = code & "# Named pipes are a stream-based mechanism for inter-process communication (IPC) within a architecture client" & vbCrLf
    code = code & "# server. In the .NET framework, they enable one-way or duplex communication between a 'pipe server' and one or" & vbCrLf
    code = code & "# more 'pipe clients.' Named pipes can facilitate interprocess communication both locally and across a network." & vbCrLf
    code = code & "# They provide a reliable method for communication between separate PowerShell processes, whether on the same" & vbCrLf
    code = code & "# machine or different machines." & vbCrLf
    code = code & "##" & vbCrLf
    code = code & vbCrLf
    
    ' Cmdlet declaration
    code = code & "# To enable CLI option -v/-Verbose (and more...)" & vbCrLf
    code = code & "[CmdletBinding()]" & vbCrLf
    code = code & "param()" & vbCrLf
    code = code & vbCrLf

    ' Region global variables
    code = code & "# region - Setting of global variables" & vbCrLf
    code = code & "$Global:pipeName = '\\.\pipe\StreamVBAMonologger'" & vbCrLf
    code = code & "$Global:prefixConsoleOutput = '[VBAMonologger server]'" & vbCrLf
    code = code & "$Global:defaultForegroundColor = 'White'" & vbCrLf
    code = code & "$Global:defaultBackgroundColor = 'Black'" & vbCrLf
    code = code & "$Global:highlightForegroundColor = 'DarkGreen'" & vbCrLf
    code = code & "$Global:highlightBackgroundColor = 'Black'" & vbCrLf
    code = code & "$Global:pipeReader = $null" & vbCrLf
    code = code & "$Global:pipeWriter = $null" & vbCrLf
    code = code & "# endregion" & vbCrLf
    code = code & vbCrLf
    
    ' Region Helpers console
    code = code & "# region - Helpers console" & vbCrLf
    
    ' Function _consoleWriteStyles
    code = code & "function _consoleWriteStyles {" & vbCrLf
    code = code & "    param (" & vbCrLf
    code = code & "        [Parameter(Mandatory = $true)] [string]$message" & vbCrLf
    code = code & "    )" & vbCrLf
    code = code & vbCrLf
    code = code & "    # Search tag {h}{/h} and replace it by its default styles" & vbCrLf
    code = code & "    $message = $message -replace '<h>(.*?)</h>', ""<style=`""foregroundColor:$highlightForegroundColor; backgroundColor:$highlightBackgroundColor;`"">`${1}`</s>""" & vbCrLf
    code = code & vbCrLf
    code = code & "    # Regex used to capture styles" & vbCrLf
    code = code & "    $regex = '<style=""([^""]+)"">(.+?)</s>'" & vbCrLf
    code = code & "    $regexStyleColor = '(?:foregroundColor:(?<fgColor>[^;]+);?)?(?:\s*backgroundColor:(?<bgColor>[^;]+);?)?'" & vbCrLf
    code = code & vbCrLf
    code = code & "    $lastIndex = 0" & vbCrLf
    code = code & "    $matchesStyles = [regex]::Matches($message, $regex)" & vbCrLf
    code = code & "    if ($matchesStyles.Count -eq 0) {" & vbCrLf
    code = code & "        # If no matches, write text with default style" & vbCrLf
    code = code & "        [console]::WriteLine($message)" & vbCrLf
    code = code & "    } else {" & vbCrLf
    code = code & "        # Saved current console colors" & vbCrLf
    code = code & "        $currentForegroundColor = [console]::ForegroundColor" & vbCrLf
    code = code & "        $currentBackgroundColor = [console]::BackgroundColor" & vbCrLf
    code = code & vbCrLf
    code = code & "        foreach ($match in $matchesStyles) {" & vbCrLf
    code = code & "            $captureStyles = $match.Groups[1].Value" & vbCrLf
    code = code & "            $captureTextStyled = $match.Groups[2].Value" & vbCrLf
    code = code & "            $startIndex = $match.Index" & vbCrLf
    code = code & vbCrLf
    code = code & "            # Write text with default style before style bloc" & vbCrLf
    code = code & "            if ($startIndex - $lastIndex -gt 0) {" & vbCrLf
    code = code & "                $textBefore = $message.Substring($lastIndex, $startIndex - $lastIndex)" & vbCrLf
    code = code & "                [console]::Write($textBefore)" & vbCrLf
    code = code & "            }" & vbCrLf
    code = code & vbCrLf
    code = code & "            # Capture styles properties colors" & vbCrLf
    code = code & "            $styleMatch = [regex]::Match($captureStyles, $regexStyleColor)" & vbCrLf
    code = code & "            $foregroundColor = if ($styleMatch.Groups[""fgColor""].Value) {" & vbCrLf
    code = code & "                $styleMatch.Groups[""fgColor""].Value" & vbCrLf
    code = code & "            } else {" & vbCrLf
    code = code & "                $currentForegroundColor" & vbCrLf
    code = code & "            }" & vbCrLf
    code = code & "            $backgroundColor = if ($styleMatch.Groups[""bgColor""].Value) {" & vbCrLf
    code = code & "                $styleMatch.Groups[""bgColor""].Value" & vbCrLf
    code = code & "            } else {" & vbCrLf
    code = code & "                $currentBackgroundColor" & vbCrLf
    code = code & "            }" & vbCrLf
    code = code & vbCrLf
    code = code & "            # Write capture text with colors" & vbCrLf
    code = code & "            [console]::ForegroundColor = $foregroundColor" & vbCrLf
    code = code & "            [console]::BackgroundColor = $backgroundColor" & vbCrLf
    code = code & "            [console]::Write($captureTextStyled)" & vbCrLf
    code = code & "            [console]::ForegroundColor = $currentForegroundColor" & vbCrLf
    code = code & "            [console]::BackgroundColor = $currentBackgroundColor" & vbCrLf
    code = code & vbCrLf
    code = code & "            # Update position after regexp match" & vbCrLf
    code = code & "            $lastIndex = $startIndex + $match.Length" & vbCrLf
    code = code & "        }" & vbCrLf
    code = code & vbCrLf
    code = code & "        # Write remaining text after the last matches with default style" & vbCrLf
    code = code & "        if ($lastIndex -lt $message.Length) {" & vbCrLf
    code = code & "            $remainingText = $message.Substring($lastIndex)" & vbCrLf
    code = code & "            [console]::Write($remainingText)" & vbCrLf
    code = code & "        }" & vbCrLf
    code = code & "        [console]::WriteLine("""")" & vbCrLf
    code = code & "    }" & vbCrLf
    code = code & "}" & vbCrLf
    code = code & vbCrLf
    
    ' Function _consoleLog
    code = code & "function _consoleLog" & vbCrLf
    code = code & "{" & vbCrLf
    code = code & "    param (" & vbCrLf
    code = code & "        [Parameter(Mandatory = $true)] [string] $message," & vbCrLf
    code = code & "        [string] $type" & vbCrLf
    code = code & "    )" & vbCrLf
    code = code & vbCrLf
    code = code & "    # Save current colors" & vbCrLf
    code = code & "    $currentForegroundColor = [console]::ForegroundColor" & vbCrLf
    code = code & "    $currentBackgroundColor = [console]::BackgroundColor" & vbCrLf
    code = code & "    $currentHighlightForegroundColor = $Global:highlightForegroundColor" & vbCrLf
    code = code & "    $currentHighlightBackgroundColor = $Global:highlightBackgroundColor" & vbCrLf
    code = code & vbCrLf
    code = code & "    # Prepare output message" & vbCrLf
    code = code & "    $type = $type.ToLower()" & vbCrLf
    code = code & "    switch ($type)" & vbCrLf
    code = code & "    {" & vbCrLf
    code = code & "        'debug' {" & vbCrLf
    code = code & "            if ($VerbosePreference -eq 'Continue')" & vbCrLf
    code = code & "            {" & vbCrLf
    code = code & "                [console]::ForegroundColor = 'DarkGray'" & vbCrLf
    code = code & "                $Global:highlightForegroundColor = 'DarkGreen'" & vbCrLf
    code = code & "                $message = $Global:prefixConsoleOutput + ' [{0}] {1}' -f (Get-Date).ToString(), $message" & vbCrLf
    code = code & "            }" & vbCrLf
    code = code & "        }" & vbCrLf
    code = code & "        'error' {" & vbCrLf
    code = code & "            [console]::ForegroundColor = 'Red'" & vbCrLf
    code = code & "            $Global:highlightForegroundColor = 'Yellow'" & vbCrLf
    code = code & "            $Global:highlightBackgroundColor = 'DarkRed'" & vbCrLf
    code = code & "            $message = $Global:prefixConsoleOutput + ' [{0}] [ERROR] {1}' -f (Get-Date).ToString(), $message" & vbCrLf
    code = code & "        }" & vbCrLf
    code = code & "        'warning' {" & vbCrLf
    code = code & "            [console]::ForegroundColor = 'DarkYellow'" & vbCrLf
    code = code & "            $Global:highlightBackgroundColor = 'DarkMagenta'" & vbCrLf
    code = code & "            $message = $Global:prefixConsoleOutput + ' [{0}] [WARNING] {1}' -f (Get-Date).ToString(), $message" & vbCrLf
    code = code & "        }" & vbCrLf
    code = code & "        default {" & vbCrLf
    code = code & "            [console]::ForegroundColor = 'White'" & vbCrLf
    code = code & "            $message = $Global:prefixConsoleOutput + ' [{0}] {1}' -f (Get-Date).ToString(), $message" & vbCrLf
    code = code & "        }" & vbCrLf
    code = code & "    }" & vbCrLf
    code = code & "    _consoleWriteStyles($message)" & vbCrLf
    code = code & vbCrLf
    code = code & "    # Restore previous colors" & vbCrLf
    code = code & "    [console]::ForegroundColor = $currentForegroundColor" & vbCrLf
    code = code & "    [console]::BackgroundColor = $currentBackgroundColor" & vbCrLf
    code = code & "    $Global:highlightForegroundColor = $currentHighlightForegroundColor" & vbCrLf
    code = code & "    $Global:highlightBackgroundColor = $currentHighlightBackgroundColor" & vbCrLf
    code = code & "}" & vbCrLf
    code = code & vbCrLf
    
    ' Function _consoleDebug
    code = code & "function _consoleDebug {" & vbCrLf
    code = code & "    param ([Parameter(Mandatory = $true)] [string]$message)" & vbCrLf
    code = code & "    if ($VerbosePreference -eq 'Continue')" & vbCrLf
    code = code & "    {" & vbCrLf
    code = code & "        _consoleLog -message $message -type 'debug'" & vbCrLf
    code = code & "    }" & vbCrLf
    code = code & "}" & vbCrLf
    code = code & vbCrLf

    ' Function _consoleError
    code = code & "function _consoleError()" & vbCrLf
    code = code & "{" & vbCrLf
    code = code & "    param ([Parameter(Mandatory = $true)] [string]$message)" & vbCrLf
    code = code & "    _consoleLog -message $message -type 'error'" & vbCrLf
    code = code & "}" & vbCrLf
    code = code & vbCrLf
    
    ' Function _consoleWarning
    code = code & "function _consoleWarning()" & vbCrLf
    code = code & "{" & vbCrLf
    code = code & "    param ([Parameter(Mandatory = $true)] [string]$message)" & vbCrLf
    code = code & "    _consoleLog -message $message -type 'warning'" & vbCrLf
    code = code & "}" & vbCrLf
    code = code & vbCrLf
    
    ' Function _splashscreen
    code = code & "function _splashscreen()" & vbCrLf
    code = code & "{" & vbCrLf
    code = code & "    # Clear output console" & vbCrLf
    code = code & "    [System.Console]::Clear()" & vbCrLf
    code = code & "    $Host.UI.RawUI.WindowTitle = ""VBAMonologger server logs viewer""" & vbCrLf
    code = code & vbCrLf
    code = code & "    # Write splashscreen message" & vbCrLf
    code = code & "    $currentForegroundColor = [console]::ForegroundColor" & vbCrLf
    code = code & "    [console]::ForegroundColor = 'Blue'" & vbCrLf
    code = code & "    [console]::WriteLine(""`n=== Start VBAMonologger server logs viewer ===`n"")" & vbCrLf
    code = code & vbCrLf
    code = code & "    # Restore previous color" & vbCrLf
    code = code & "    [console]::ForegroundColor = $currentForegroundColor" & vbCrLf
    code = code & "}" & vbCrLf
    code = code & vbCrLf

    ' Close region: Helpers console
    code = code & "# endregion" & vbCrLf
    code = code & vbCrLf
    
    ' Region: Server
    code = code & "# region - Server logs viewer" & vbCrLf
    code = code & vbCrLf
    
    ' Function _checkAdminSecurityPermissions
    code = code & "function _checkAdminSecurityPermissions" & vbCrLf
    code = code & "{" & vbCrLf
    code = code & "    # Check if current user has administrator privileges, only in verbose mode (-v)" & vbCrLf
    code = code & "    if ($VerbosePreference -eq 'Continue')" & vbCrLf
    code = code & "    {" & vbCrLf
    code = code & "        try" & vbCrLf
    code = code & "        {" & vbCrLf
    code = code & "            $user = [Security.Principal.WindowsIdentity]::GetCurrent()" & vbCrLf
    code = code & "            $principal = New-Object Security.Principal.WindowsPrincipal($user)" & vbCrLf
    code = code & "            $isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)" & vbCrLf
    code = code & "            if ($isAdmin)" & vbCrLf
    code = code & "            {" & vbCrLf
    code = code & "                _consoleDebug('Current user running the server has administrator privileges.')" & vbCrLf
    code = code & "            }" & vbCrLf
    code = code & "            else" & vbCrLf
    code = code & "            {" & vbCrLf
    code = code & "                _consoleWarning('Current user running the server does not have administrator privileges.')" & vbCrLf
    code = code & "            }" & vbCrLf
    code = code & "        }" & vbCrLf
    code = code & "        catch" & vbCrLf
    code = code & "        {" & vbCrLf
    code = code & "            _consoleError('Unable to verify the current user''s permissions: $_')" & vbCrLf
    code = code & "        }" & vbCrLf
    code = code & "    }" & vbCrLf
    code = code & "}" & vbCrLf
    code = code & vbCrLf
    
    ' Function _readRequest
    code = code & "function _readRequest" & vbCrLf
    code = code & "{" & vbCrLf
    code = code & "    $request = $pipeReader.ReadLine()" & vbCrLf
    code = code & "    _consoleDebug('Received data from client''s request.')" & vbCrLf
    code = code & "    return $request" & vbCrLf
    code = code & "}" & vbCrLf
    code = code & vbCrLf
    
    ' Function _startServer
    code = code & "function _startServer()" & vbCrLf
    code = code & "{" & vbCrLf
    code = code & "    [console]::ForegroundColor = $defaultForegroundColor" & vbCrLf
    code = code & "    [console]::BackgroundColor = $defaultBackgroundColor" & vbCrLf
    code = code & vbCrLf
    code = code & "    _splashscreen" & vbCrLf
    code = code & "    _checkAdminSecurityPermissions" & vbCrLf
    code = code & vbCrLf
    code = code & "    try" & vbCrLf
    code = code & "    {" & vbCrLf
    code = code & "        # Create a bi-directionnal named pipe server" & vbCrLf
    code = code & "        $pipeServer = New-object System.IO.Pipes.NamedPipeServerStream($pipeName, [System.IO.Pipes.PipeDirection]::InOut)" & vbCrLf
    code = code & "        _consoleDebug('Server is listening on pipe: <h>""$pipeName""</h>')" & vbCrLf
    code = code & vbCrLf
    code = code & "        While ($true)" & vbCrLf
    code = code & "        {" & vbCrLf
    code = code & "            _consoleDebug('Waiting for new VBAMonologger client logs sender connection...')" & vbCrLf
    code = code & "            $pipeServer.WaitForConnection()" & vbCrLf
    code = code & "            _consoleDebug('Connection established by a new client.')" & vbCrLf
    code = code & vbCrLf
    code = code & "            # Create stream writer (for writing request) and stream reader (for reading response)" & vbCrLf
    code = code & "            $Global:pipeReader = New-object System.IO.StreamReader($pipeServer)" & vbCrLf
    code = code & "            $Global:pipeWriter = New-Object System.IO.StreamWriter($pipeServer)" & vbCrLf
    code = code & "            $pipeWriter.AutoFlush = $true" & vbCrLf
    code = code & vbCrLf
    code = code & "            # Read client''s request" & vbCrLf
    code = code & "            $request = _readRequest" & vbCrLf
    code = code & vbCrLf
    code = code & "            # Prepare server''s response" & vbCrLf
    code = code & "            $response = New-Object PSObject -Property @{" & vbCrLf
    code = code & "                id = [guid]::NewGuid()" & vbCrLf
    code = code & "            }" & vbCrLf
    code = code & vbCrLf
    code = code & "            # Process client''s request" & vbCrLf
    code = code & "            $command = $request.ToLower()" & vbCrLf
    code = code & "            if ($command -eq 'exit' -or $command -eq 'stop' -or $command -eq 'stop-server')" & vbCrLf
    code = code & "            {" & vbCrLf
    code = code & "                $response | Add-Member -MemberType NoteProperty -Name body -Value 'Command exit received and executed by the server.'" & vbCrLf
    code = code & "                $reponseString = $response | ConvertTo-Json" & vbCrLf
    code = code & "                $pipeWriter.Write($reponseString)" & vbCrLf
    code = code & "                _consoleDebug('Send a server response to the client.')" & vbCrLf
    code = code & vbCrLf
    code = code & "                _consoleDebug('Executed the <h>""exit""</h> command to stop the server.')" & vbCrLf
    code = code & "                $pipeReader.Dispose()" & vbCrLf
    code = code & "                break" & vbCrLf
    code = code & "            }" & vbCrLf
    code = code & "            else" & vbCrLf
    code = code & "            {" & vbCrLf
    code = code & "                $response | Add-Member -MemberType NoteProperty -Name body -Value 'Request received successfully by the server.'" & vbCrLf
    code = code & "                $reponseString = $response | ConvertTo-Json" & vbCrLf
    code = code & "                $pipeWriter.Write($reponseString)" & vbCrLf
    code = code & vbCrLf
    code = code & "                # Simply output data''s request into console" & vbCrLf
    code = code & "                [console]::WriteLine($request)" & vbCrLf
    code = code & "            }" & vbCrLf
    code = code & vbCrLf
    code = code & "            # Disconnects the client from the server" & vbCrLf
    code = code & "            $pipeServer.Disconnect()" & vbCrLf
    code = code & "            _consoleDebug('Closed the connection with the current client.')" & vbCrLf
    code = code & "        }" & vbCrLf
    code = code & "    }" & vbCrLf
    code = code & "    Catch" & vbCrLf
    code = code & "    {" & vbCrLf
    code = code & "        _consoleError(""The server encountered an <h>critical error</h>, preventing it from continuing to operate.`n$_"")" & vbCrLf
    code = code & "        Exit" & vbCrLf
    code = code & "    }" & vbCrLf
    code = code & "    Finally" & vbCrLf
    code = code & "    {" & vbCrLf
    code = code & "        [console]::WriteLine(""`nServer shutdown, bye bye !"")" & vbCrLf
    code = code & "        $pipeServer.Dispose()" & vbCrLf
    code = code & "    }" & vbCrLf
    code = code & "}" & vbCrLf
    code = code & "# endregion" & vbCrLf
    code = code & vbCrLf
    
    ' Start app
    code = code & "_startServer" & vbCrLf
 
    getPowershellCodeServerLogsViewer = code
End Function

Private Function getPowershellCodeClientLogsSender()
    Dim code As String
    code = ""
    
    ' Header code source file
    code = code & "# ------------------------------------- #" & vbCrLf
    code = code & "#                                       #" & vbCrLf
    code = code & "#    VBA Monologger                     #" & vbCrLf
    code = code & "#    Copyright � 2024, 6i software      #" & vbCrLf
    code = code & "#                                       #" & vbCrLf
    code = code & "# ------------------------------------- #" & vbCrLf
    code = code & "#" & vbCrLf
    code = code & "# Client for sending log message to a server by using a named pipe stream." & vbCrLf
    code = code & "#" & vbCrLf
    code = code & "# The server will do nothing more than display the log message sent by the client in the console." & vbCrLf
    code = code & "#" & vbCrLf
    code = code & "# Named pipes are a stream-based mechanism for inter-process communication (IPC) within a architecture client " & vbCrLf
    code = code & "# server. In the .NET framework, they enable one-way or duplex communication between a ""pipe server"" and one or " & vbCrLf
    code = code & "# more ""pipe clients."" Named pipes can facilitate interprocess communication both locally and across a network. " & vbCrLf
    code = code & "# They provide a reliable method for communication between separate PowerShell processes, whether on the same " & vbCrLf
    code = code & "# machine or different machines." & vbCrLf
    code = code & "##" & vbCrLf
    code = code & vbCrLf

    ' CLI parameters
    code = code & "param (" & vbCrLf
    code = code & "# The log message that will be sent to the server." & vbCrLf
    code = code & "    [Parameter(Position = 0, Mandatory = $true)][string] $message," & vbCrLf
    code = code & "# Defines the maximum time (in milliseconds) to wait for a response or connection." & vbCrLf
    code = code & "    [Parameter(Position = 1)][int] $timeout = 2000," & vbCrLf
    code = code & "# Defines the interval (in milliseconds) between retry attempts during the timeout period." & vbCrLf
    code = code & "    [Parameter(Position = 2)][int] $timeRetryInterval = 250" & vbCrLf
    code = code & ")" & vbCrLf
    code = code & vbCrLf

    ' Global variables
    code = code & "# region - Setting of global variables" & vbCrLf
    code = code & "$Global:pipeName = '\\.\pipe\StreamVBAMonologger'" & vbCrLf
    code = code & "$Global:prefixConsoleOutput = ""[VBAMonologger client]""" & vbCrLf
    code = code & "$Global:defaultForegroundColor = ""White""" & vbCrLf
    code = code & "$Global:defaultBackgroundColor = ""Black""" & vbCrLf
    code = code & "$Global:highlightForegroundColor = ""Green""" & vbCrLf
    code = code & "$Global:highlightBackgroundColor = ""Black""" & vbCrLf
    code = code & "$Global:pipeReader = $null" & vbCrLf
    code = code & "$Global:pipeWriter = $null" & vbCrLf
    code = code & "# endregion" & vbCrLf

    ' Region Helpers console
    code = code & "# region - Helpers console" & vbCrLf
    
    ' Function _consoleWriteStyles
    code = code & "function _consoleWriteStyles {" & vbCrLf
    code = code & "    param (" & vbCrLf
    code = code & "        [Parameter(Mandatory = $true)] [string]$message" & vbCrLf
    code = code & "    )" & vbCrLf
    code = code & vbCrLf
    code = code & "    # Search tag {h}{/h} and replace it by its default styles" & vbCrLf
    code = code & "    $message = $message -replace '<h>(.*?)</h>', ""<style=`""foregroundColor:$highlightForegroundColor; backgroundColor:$highlightBackgroundColor;`"">`${1}`</s>""" & vbCrLf
    code = code & vbCrLf
    code = code & "    # Regex used to capture styles" & vbCrLf
    code = code & "    $regex = '<style=""([^""]+)"">(.+?)</s>'" & vbCrLf
    code = code & "    $regexStyleColor = '(?:foregroundColor:(?<fgColor>[^;]+);?)?(?:\s*backgroundColor:(?<bgColor>[^;]+);?)?'" & vbCrLf
    code = code & vbCrLf
    code = code & "    $lastIndex = 0" & vbCrLf
    code = code & "    $matchesStyles = [regex]::Matches($message, $regex)" & vbCrLf
    code = code & "    if ($matchesStyles.Count -eq 0) {" & vbCrLf
    code = code & "        # If no matches, write text with default style" & vbCrLf
    code = code & "        [console]::WriteLine($message)" & vbCrLf
    code = code & "    } else {" & vbCrLf
    code = code & "        # Saved current console colors" & vbCrLf
    code = code & "        $currentForegroundColor = [console]::ForegroundColor" & vbCrLf
    code = code & "        $currentBackgroundColor = [console]::BackgroundColor" & vbCrLf
    code = code & vbCrLf
    code = code & "        foreach ($match in $matchesStyles) {" & vbCrLf
    code = code & "            $captureStyles = $match.Groups[1].Value" & vbCrLf
    code = code & "            $captureTextStyled = $match.Groups[2].Value" & vbCrLf
    code = code & "            $startIndex = $match.Index" & vbCrLf
    code = code & vbCrLf
    code = code & "            # Write text with default style before style bloc" & vbCrLf
    code = code & "            if ($startIndex - $lastIndex -gt 0) {" & vbCrLf
    code = code & "                $textBefore = $message.Substring($lastIndex, $startIndex - $lastIndex)" & vbCrLf
    code = code & "                [console]::Write($textBefore)" & vbCrLf
    code = code & "            }" & vbCrLf
    code = code & vbCrLf
    code = code & "            # Capture styles properties colors" & vbCrLf
    code = code & "            $styleMatch = [regex]::Match($captureStyles, $regexStyleColor)" & vbCrLf
    code = code & "            $foregroundColor = if ($styleMatch.Groups[""fgColor""].Value) {" & vbCrLf
    code = code & "                $styleMatch.Groups[""fgColor""].Value" & vbCrLf
    code = code & "            } else {" & vbCrLf
    code = code & "                $currentForegroundColor" & vbCrLf
    code = code & "            }" & vbCrLf
    code = code & "            $backgroundColor = if ($styleMatch.Groups[""bgColor""].Value) {" & vbCrLf
    code = code & "                $styleMatch.Groups[""bgColor""].Value" & vbCrLf
    code = code & "            } else {" & vbCrLf
    code = code & "                $currentBackgroundColor" & vbCrLf
    code = code & "            }" & vbCrLf
    code = code & vbCrLf
    code = code & "            # Write capture text with colors" & vbCrLf
    code = code & "            [console]::ForegroundColor = $foregroundColor" & vbCrLf
    code = code & "            [console]::BackgroundColor = $backgroundColor" & vbCrLf
    code = code & "            [console]::Write($captureTextStyled)" & vbCrLf
    code = code & "            [console]::ForegroundColor = $currentForegroundColor" & vbCrLf
    code = code & "            [console]::BackgroundColor = $currentBackgroundColor" & vbCrLf
    code = code & vbCrLf
    code = code & "            # Update position after regexp match" & vbCrLf
    code = code & "            $lastIndex = $startIndex + $match.Length" & vbCrLf
    code = code & "        }" & vbCrLf
    code = code & vbCrLf
    code = code & "        # Write remaining text after the last matches with default style" & vbCrLf
    code = code & "        if ($lastIndex -lt $message.Length) {" & vbCrLf
    code = code & "            $remainingText = $message.Substring($lastIndex)" & vbCrLf
    code = code & "            [console]::Write($remainingText)" & vbCrLf
    code = code & "        }" & vbCrLf
    code = code & "        [console]::WriteLine("""")" & vbCrLf
    code = code & "    }" & vbCrLf
    code = code & "}" & vbCrLf
    code = code & vbCrLf
    
    ' Function _consoleLog
    code = code & "function _consoleLog" & vbCrLf
    code = code & "{" & vbCrLf
    code = code & "    param (" & vbCrLf
    code = code & "        [Parameter(Mandatory = $true)] [string] $message," & vbCrLf
    code = code & "        [string] $type" & vbCrLf
    code = code & "    )" & vbCrLf
    code = code & vbCrLf
    code = code & "    # Save current colors" & vbCrLf
    code = code & "    $currentForegroundColor = [console]::ForegroundColor" & vbCrLf
    code = code & "    $currentBackgroundColor = [console]::BackgroundColor" & vbCrLf
    code = code & "    $currentHighlightForegroundColor = $Global:highlightForegroundColor" & vbCrLf
    code = code & "    $currentHighlightBackgroundColor = $Global:highlightBackgroundColor" & vbCrLf
    code = code & vbCrLf
    code = code & "    # Prepare output message" & vbCrLf
    code = code & "    $type = $type.ToLower()" & vbCrLf
    code = code & "    switch ($type)" & vbCrLf
    code = code & "    {" & vbCrLf
    code = code & "        'debug' {" & vbCrLf
    code = code & "            if ($VerbosePreference -eq 'Continue')" & vbCrLf
    code = code & "            {" & vbCrLf
    code = code & "                [console]::ForegroundColor = 'DarkGray'" & vbCrLf
    code = code & "                $Global:highlightForegroundColor = 'DarkGreen'" & vbCrLf
    code = code & "                $message = $Global:prefixConsoleOutput + ' [{0}] {1}' -f (Get-Date).ToString(), $message" & vbCrLf
    code = code & "            }" & vbCrLf
    code = code & "        }" & vbCrLf
    code = code & "        'error' {" & vbCrLf
    code = code & "            [console]::ForegroundColor = 'Red'" & vbCrLf
    code = code & "            $Global:highlightForegroundColor = 'Yellow'" & vbCrLf
    code = code & "            $Global:highlightBackgroundColor = 'DarkRed'" & vbCrLf
    code = code & "            $message = $Global:prefixConsoleOutput + ' [{0}] [ERROR] {1}' -f (Get-Date).ToString(), $message" & vbCrLf
    code = code & "        }" & vbCrLf
    code = code & "        'warning' {" & vbCrLf
    code = code & "            [console]::ForegroundColor = 'DarkYellow'" & vbCrLf
    code = code & "            $Global:highlightBackgroundColor = 'DarkMagenta'" & vbCrLf
    code = code & "            $message = $Global:prefixConsoleOutput + ' [{0}] [WARNING] {1}' -f (Get-Date).ToString(), $message" & vbCrLf
    code = code & "        }" & vbCrLf
    code = code & "        default {" & vbCrLf
    code = code & "            [console]::ForegroundColor = 'White'" & vbCrLf
    code = code & "            $message = $Global:prefixConsoleOutput + ' [{0}] {1}' -f (Get-Date).ToString(), $message" & vbCrLf
    code = code & "        }" & vbCrLf
    code = code & "    }" & vbCrLf
    code = code & "    _consoleWriteStyles($message)" & vbCrLf
    code = code & vbCrLf
    code = code & "    # Restore previous colors" & vbCrLf
    code = code & "    [console]::ForegroundColor = $currentForegroundColor" & vbCrLf
    code = code & "    [console]::BackgroundColor = $currentBackgroundColor" & vbCrLf
    code = code & "    $Global:highlightForegroundColor = $currentHighlightForegroundColor" & vbCrLf
    code = code & "    $Global:highlightBackgroundColor = $currentHighlightBackgroundColor" & vbCrLf
    code = code & "}" & vbCrLf
    
    ' Function _consoleDebug
    code = code & "function _consoleDebug {" & vbCrLf
    code = code & "    param ([Parameter(Mandatory = $true)] [string]$message)" & vbCrLf
    code = code & "    if ($VerbosePreference -eq 'Continue')" & vbCrLf
    code = code & "    {" & vbCrLf
    code = code & "        _consoleLog -message $message -type 'debug'" & vbCrLf
    code = code & "    }" & vbCrLf
    code = code & "}" & vbCrLf
    code = code & vbCrLf

    ' Function _consoleError
    code = code & "function _consoleError()" & vbCrLf
    code = code & "{" & vbCrLf
    code = code & "    param ([Parameter(Mandatory = $true)] [string]$message)" & vbCrLf
    code = code & "    _consoleLog -message $message -type 'error'" & vbCrLf
    code = code & "}" & vbCrLf
    code = code & vbCrLf
    
    ' Function _consoleWarning
    code = code & "function _consoleWarning()" & vbCrLf
    code = code & "{" & vbCrLf
    code = code & "    param ([Parameter(Mandatory = $true)] [string]$message)" & vbCrLf
    code = code & "    _consoleLog -message $message -type 'warning'" & vbCrLf
    code = code & "}" & vbCrLf
    code = code & vbCrLf
    
    ' Function _splashscreen
    code = code & "function _splashscreen()" & vbCrLf
    code = code & "{" & vbCrLf
    code = code & "    # Clear output console" & vbCrLf
    code = code & "    [System.Console]::Clear()" & vbCrLf
    code = code & vbCrLf
    code = code & "    # Write splashscreen message" & vbCrLf
    code = code & "    $currentForegroundColor = [console]::ForegroundColor" & vbCrLf
    code = code & "    [console]::ForegroundColor = 'Blue'" & vbCrLf
    code = code & "    [console]::WriteLine(""`n=== Start VBAMonologger client logs sender ===`n"")" & vbCrLf
    code = code & vbCrLf
    code = code & "    # Restore previous color" & vbCrLf
    code = code & "    [console]::ForegroundColor = $currentForegroundColor" & vbCrLf
    code = code & "}" & vbCrLf
    code = code & vbCrLf
    
    ' Function _setConsoleDefaultColors
    code = code & "function _setConsoleDefaultColors" & vbCrLf
    code = code & "{" & vbCrLf
    code = code & "    # Apply default colors from global" & vbCrLf
    code = code & "    [console]::ForegroundColor = $defaultForegroundColor" & vbCrLf
    code = code & "    [console]::BackgroundColor = $defaultBackgroundColor" & vbCrLf
    code = code & "}" & vbCrLf
    code = code & vbCrLf

    ' Close region: Helpers console
    code = code & "# endregion" & vbCrLf
    code = code & vbCrLf
    
    ' Region Client
    code = code & "# region - Client logs sender" & vbCrLf
    code = code & vbCrLf
    
    ' Custom exception
    code = code & "# region - Custom exception class" & vbCrLf
    code = code & "class CustomTimeoutException: System.Exception" & vbCrLf
    code = code & "{" & vbCrLf
    code = code & "    CustomTimeoutException([string]$message): base($message)" & vbCrLf
    code = code & "    {" & vbCrLf
    code = code & "    }" & vbCrLf
    code = code & "}" & vbCrLf
    code = code & "# endregion" & vbCrLf
    code = code & vbCrLf
    
    ' Function _checkAdminSecurityPermissions
    code = code & "function _checkAdminSecurityPermissions" & vbCrLf
    code = code & "{" & vbCrLf
    code = code & "    # Check if current user has administrator privileges, only in verbose mode (-v)" & vbCrLf
    code = code & "    if ($VerbosePreference -eq ""Continue"")" & vbCrLf
    code = code & "    {" & vbCrLf
    code = code & "        try" & vbCrLf
    code = code & "        {" & vbCrLf
    code = code & "            $user = [Security.Principal.WindowsIdentity]::GetCurrent()" & vbCrLf
    code = code & "            $principal = New-Object Security.Principal.WindowsPrincipal($user)" & vbCrLf
    code = code & "            $isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)" & vbCrLf
    code = code & "            if ($isAdmin)" & vbCrLf
    code = code & "            {" & vbCrLf
    code = code & "                _consoleDebug(""Current client has administrator privileges."")" & vbCrLf
    code = code & "            }" & vbCrLf
    code = code & "            else" & vbCrLf
    code = code & "            {" & vbCrLf
    code = code & "                _consoleWarning(""Current client does not have administrator privileges."")" & vbCrLf
    code = code & "            }" & vbCrLf
    code = code & "        }" & vbCrLf
    code = code & "        catch" & vbCrLf
    code = code & "        {" & vbCrLf
    code = code & "            _consoleError(""Unable to verify the current user's permissions: $_"")" & vbCrLf
    code = code & "        }" & vbCrLf
    code = code & "    }" & vbCrLf
    code = code & "}" & vbCrLf
    
    ' Function _connectToServer
    code = code & "function _connectToServer" & vbCrLf
    code = code & "{" & vbCrLf
    code = code & "    $pipeClient = New-Object System.IO.Pipes.NamedPipeClientStream($pipeName)" & vbCrLf
    code = code & "    _consoleDebug(""Client is trying to connect on server with pipe name: <h>""""$pipeName""""</h>"")" & vbCrLf
    code = code & vbCrLf
    code = code & "    $stopWatch = [System.Diagnostics.Stopwatch]::StartNew()" & vbCrLf
    code = code & "    $attempts = 0" & vbCrLf
    code = code & "    while ($stopWatch.ElapsedMilliseconds -lt $timeout)" & vbCrLf
    code = code & "    {" & vbCrLf
    code = code & "        try" & vbCrLf
    code = code & "        {" & vbCrLf
    code = code & "            $attempts++" & vbCrLf
    code = code & vbCrLf
    code = code & "            # No delay on first try connection" & vbCrLf
    code = code & "            if ($attempts -eq 1)" & vbCrLf
    code = code & "            {" & vbCrLf
    code = code & "                $pipeClient.Connect(0)" & vbCrLf
    code = code & "            }" & vbCrLf
    code = code & "            else" & vbCrLf
    code = code & "            {" & vbCrLf
    code = code & "                $pipeClient.Connect($timeRetryInterval)" & vbCrLf
    code = code & "            }" & vbCrLf
    code = code & vbCrLf
    code = code & "            if ($pipeClient.IsConnected)" & vbCrLf
    code = code & "            {" & vbCrLf
    code = code & "                _consoleDebug(""Connection established with the server."")" & vbCrLf
    code = code & "                return $pipeClient" & vbCrLf
    code = code & "            }" & vbCrLf
    code = code & "        }" & vbCrLf
    code = code & "        catch" & vbCrLf
    code = code & "        {" & vbCrLf
    code = code & "            _consoleWarning(""Attempt $attempts failed, unable to connect to this server. Try another attempt."")" & vbCrLf
    code = code & "            Start-Sleep -Milliseconds $timeRetryInterval" & vbCrLf
    code = code & "        }" & vbCrLf
    code = code & "    }" & vbCrLf
    code = code & vbCrLf
    code = code & "    throw [CustomTimeoutException]::new(""Failed to connect to server within the timeout period."")" & vbCrLf
    code = code & "}" & vbCrLf
    code = code & vbCrLf
    code = code & "function _sendRequest" & vbCrLf
    code = code & "{" & vbCrLf
    code = code & "    param (" & vbCrLf
    code = code & "        [Parameter(Mandatory = $true)] [string]$message" & vbCrLf
    code = code & "    )" & vbCrLf
    code = code & vbCrLf
    code = code & "    _consoleDebug(""Try to send a request to the server."")" & vbCrLf
    code = code & "    $Global:pipeWriter.WriteLine($message)" & vbCrLf
    code = code & "    _consoleLog(""Client's request was successfully sent to the server: """"<h>"" + $message + ""</h>"""" "")" & vbCrLf
    code = code & "}" & vbCrLf
    code = code & vbCrLf
    
    ' Function _readResponse
    code = code & "function _readResponse" & vbCrLf
    code = code & "{" & vbCrLf
    code = code & "    $responseJson = $Global:pipeReader.ReadToEnd().TrimEnd()" & vbCrLf
    code = code & "    $response = $responseJson | ConvertFrom-Json" & vbCrLf
    code = code & "    _consoleLog(""Server's response was successfully received: """"<h>"" + $response.body + ""</h>"""" "")" & vbCrLf
    code = code & "    _consoleDebug(""Server's JSON response:`n"" + $responseJson)" & vbCrLf
    code = code & "}" & vbCrLf
    code = code & vbCrLf
    
    ' Function _startClient
    code = code & "function _startClient" & vbCrLf
    code = code & "{" & vbCrLf
    code = code & "    param (" & vbCrLf
    code = code & "        [Parameter(Mandatory)][string]$message" & vbCrLf
    code = code & "    )" & vbCrLf
    code = code & vbCrLf
    code = code & "    _setConsoleDefaultColors" & vbCrLf
    code = code & "    _splashscreen" & vbCrLf
    code = code & "    _checkAdminSecurityPermissions" & vbCrLf
    code = code & vbCrLf
    code = code & "    try" & vbCrLf
    code = code & "    {" & vbCrLf
    code = code & "        # Try to connecting to named pipe stream server within multiple attempts on timeout period" & vbCrLf
    code = code & "        $pipeClient = _connectToServer" & vbCrLf
    code = code & vbCrLf
    code = code & "        # Create stream writer (for writing request) and stream reader (for reading response)" & vbCrLf
    code = code & "        $Global:pipeReader = New-object System.IO.StreamReader($pipeClient)" & vbCrLf
    code = code & "        $Global:pipeWriter = New-Object System.IO.StreamWriter($pipeClient)" & vbCrLf
    code = code & "        $pipeWriter.AutoFlush = $true" & vbCrLf
    code = code & vbCrLf
    code = code & "        try" & vbCrLf
    code = code & "        {" & vbCrLf
    code = code & "            # Send a request to the server" & vbCrLf
    code = code & "            _sendRequest($message)" & vbCrLf
    code = code & "            # Read the server response" & vbCrLf
    code = code & "            _readResponse" & vbCrLf
    code = code & "            # Dispose of the pipeReader to release resources" & vbCrLf
    code = code & "            $Global:pipeReader.Dispose()" & vbCrLf
    code = code & "        }" & vbCrLf
    code = code & "        catch" & vbCrLf
    code = code & "        {" & vbCrLf
    code = code & "            _consoleError(""Error sending the client's message to the server:`n$_"")" & vbCrLf
    code = code & "        }" & vbCrLf
    code = code & "        finally" & vbCrLf
    code = code & "        {" & vbCrLf
    code = code & "            [console]::WriteLine(""`nClient shutdown, bye bye !"")" & vbCrLf
    code = code & "            $pipeClient.Dispose()" & vbCrLf
    code = code & "        }" & vbCrLf
    code = code & "    }" & vbCrLf
    code = code & "    catch [CustomTimeoutException]" & vbCrLf
    code = code & "    {" & vbCrLf
    code = code & "        _consoleError($_)" & vbCrLf
    code = code & "    }" & vbCrLf
    code = code & "}" & vbCrLf
    code = code & "# endregion" & vbCrLf
    code = code & vbCrLf
    
    ' Start client application
    code = code & "_startClient -message $message" & vbCrLf
   
    getPowershellCodeClientLogsSender = code
End Function
