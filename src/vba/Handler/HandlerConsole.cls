VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HandlerConsole"
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
'@FQCN("VBAMonologger.Handler.HandlerConsole")
'@ModuleDescription("Handles logs message to Windows console (cmd.exe) by streaming logs into an HTTP architecture client/server. Client sends request (log entry) and the server show into output console.")
''

Option Explicit

Implements VBAMonologger.HandlerInterface
Implements VBAMonologger.StringableInterface

Private Const TITLE_CONSOLE_WINDOW As String = "VBAMonologger logs viewer"

'@Description("The associated formatter of current handler.")
Private pformatter As VBAMonologger.FormatterInterface
'@Description("The bubbling behavior. True means that this handler allows bubbling, and false means that bubbling is not permitted.")
Private pBubble As Boolean
'@Description("The minimum logging level at which this handler will be triggered as a LogLevel")
Private pLogLevel As VBAMonologger.LogLevel
'@Description("The minimum logging level at which this handler will be triggered as an value of enumeration of LOG_LEVELS.")
Private pLevel As VBAMonologger.LOG_LEVELS

'@Description("To interact with file system (create folder, delete file...).")
Private pFileSystemObject As New Scripting.FileSystemObject
'@Description("Folder where powershell programs, for creating server logs viewer and client logs sender, are generated and stored.")
Private pTempFolderPowershellPrograms As String
'@Description("File which contains the powershell code source of 'VBAMonologger HTTP server logs viewer' program.")
Private pPowershellProgramServerFilepath As String
'@Description("To write file in UTF-8 with ADODB.Stream object available in Microsoft ActiveX Data Objects.")
Private pWriterFileStream As Object
'@Description("Wraps a Windows console (cmd.exe) in order to simplify its creation and interaction with it.")
Private pConsoleWrapper As VBAMonologger.ConsoleWrapper
'@Description("Url of VBAMonologger HTTP server. By default, we use http://localhost:20100.")
Private pUrlServer As String
'@Description("To show debug information when we use its instance (for dev only)")
Private pWithDebug As Boolean
'@Description("To enable support ANSI colors into default formatter by using FormatterANSIColoredLine.")
Private pWithANSIColorSupport As Boolean
'@Description("To write context and extra information on newlines into default line formatter.")
Private pWithNewlineForContextAndExtra As Boolean
'@Description("Hostname or IP adresse of VBAMonologger HTTP server. By default, we use localhost.")
Public hostnameServer As String
'@Description("Port of VBAMonologger HTTP server. By default, we use 20100 port number.")
Public portServer As Integer



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
            
    ' Configuration of VBAMonologger HTTP server logs viewer.
    hostnameServer = "localhost"
    portServer = 20100
    
    ' Create console wrapper
    Set pConsoleWrapper = New VBAMonologger.ConsoleWrapper
    pConsoleWrapper.titleConsoleWindow = TITLE_CONSOLE_WINDOW
    pConsoleWrapper.consoleCreationWait = 2000
    
    ' Enable debug information (for dev only)
    withDebug = False
    
    ' Configure the ADODB writer stream used for generating the PowerShell program VBAMonologger HTTP server logs viewer.
    Set pWriterFileStream = CreateObject("ADODB.Stream")
    With pWriterFileStream
        .Type = 2 ' ADODB.Stream type texte
        .Charset = "UTF-8" ' Default encoding
    End With
    
    ' Create a temporary folder to generate the PowerShell program for the VBAMonologger HTTP server logs viewer.
    pTempFolderPowershellPrograms = Environ("TEMP") & "\VBAMonologger\powershell"
    pPowershellProgramServerFilepath = pTempFolderPowershellPrograms & "\VBAMonologgerHTTPServerLogsViewer.ps1"
    createPowershellTempfolder
    createPowershellFileVBAMonologgerServer
    
    ' Add default formatter
    Set pformatter = New VBAMonologger.FormatterANSIColoredLine
    withANSIColorSupport = True
    withNewlineForContextAndExtra = True
End Sub

'@Description('Emulate a class constructor.')
Public Function construct( _
    Optional ByVal paramBubble As Boolean = True, _
    Optional ByVal paramLevel As VBAMonologger.LOG_LEVELS = VBAMonologger.LOG_LEVELS.LEVEL_DEBUG, _
    Optional ByVal paramHostnameServer As String = "localhost", _
    Optional ByVal paramPortServer As Integer = 20100 _
) As VBAMonologger.HandlerInterface
    pBubble = paramBubble
    pLogLevel.currentLogLevel = paramLevel
    pLevel = paramLevel
    hostnameServer = paramHostnameServer
    portServer = paramPortServer
    
    Set construct = Me
End Function



' ___________________ '
'                     '
'  Public properties  '
' ___________________ '
'                     '

'@Description("Get the url of VBAMonologger HTTP server.")
Public Property Get urlServer() As String
    urlServer = "http://" & hostnameServer & ":" & portServer & "/"
End Property

'@Description("To show debug information.")
Public Property Get withDebug() As Boolean
    withDebug = pWithDebug
End Property

'@Description("Sets the value to show debug information.")
Public Property Let withDebug(ByVal newValueForWithDebug As Boolean)
    pWithDebug = newValueForWithDebug
    pConsoleWrapper.withDebug = newValueForWithDebug
End Property

'@Description("Get value of option withANSIColorSupport.")
Public Property Get withANSIColorSupport() As Boolean
    withANSIColorSupport = pWithANSIColorSupport
End Property

'@Description("Sets the value of option withANSIColorSupport.")
Public Property Let withANSIColorSupport(ByVal newValueForWithANSIColorSupport As Boolean)
    pWithANSIColorSupport = newValueForWithANSIColorSupport
    Call addDefaultFormatter(withANSIColorSupport, withNewlineForContextAndExtra)
End Property

'@Description("Get value of option withNewlineForContextAndExtra.")
Public Property Get withNewlineForContextAndExtra() As Boolean
    withNewlineForContextAndExtra = pWithNewlineForContextAndExtra
End Property

'@Description("Sets the value of option withANSIColorSupport.")
Public Property Let withNewlineForContextAndExtra(ByVal newValueForWithNewlineForContextAndExtra As Boolean)
    pWithNewlineForContextAndExtra = newValueForWithNewlineForContextAndExtra
    Call addDefaultFormatter(withANSIColorSupport, withNewlineForContextAndExtra)
End Property

'@Description("To add a default formatter according to options ANSI color support, new lines for showing context and extra informations.")
Private Sub addDefaultFormatter(paramWithANSIColorSupport As Boolean, paramWithNewlineForContextAndExtra As Boolean)
    If paramWithANSIColorSupport Then
        Dim formatterAnsiColor As New VBAMonologger.FormatterANSIColoredLine
        Set formatterAnsiColor = New VBAMonologger.FormatterANSIColoredLine

        If paramWithNewlineForContextAndExtra Then
            formatterAnsiColor.setTemplateLineWithNewlineForContextAndExtra
            formatterAnsiColor.withWhitespace = True
            formatterAnsiColor.withAllowingInlineLineBreaks = True
        End If
        
        Set pformatter = formatterAnsiColor
    Else
        Dim formatter As New VBAMonologger.FormatterLine
        Set formatter = New VBAMonologger.FormatterLine
        
        If paramWithNewlineForContextAndExtra Then
            formatter.setTemplateLineWithNewlineForContextAndExtra
            formatter.withWhitespace = True
            formatter.withAllowingInlineLineBreaks = True
        End If
            
        Set pformatter = formatter
    End If
End Sub



' _______________________________________ '
'                                         '
'  VBAMonologger HTTP server logs viewer  '
' _______________________________________ '
'                                         '

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
        Err.Raise vbObjectError + 1000, "VBAMonologger.Handler.HandlerConsole::createPowershellTempfolder", "Failed to create the temp folder (" & pTempFolderPowershellPrograms & ") !"
    Else
        If withDebug Then
            Debug.Print "[DEBUG] HandlerConsole::createPowershellTempfolder | The temp folder for PowerShell programs exists: " & pTempFolderPowershellPrograms
        End If
    End If
End Sub

'@Desciption("Create the file wich contains the code of powershell program of VBAMonologger HTTP server logs viewer.")
Private Sub createPowershellFileVBAMonologgerServer()
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
            Err.Raise vbObjectError + 1000, "VBAMonologger.Handler.HandlerConsole::createPowershellFileVBAMonologgerServer", "Failed to create the powershell file: " & pPowershellProgramServerFilepath
        Else
            If withDebug Then Debug.Print "[DEBUG] HandlerConsole::createPowershellFileVBAMonologgerServer | The powershell file for VBAMonlogger server was successfully created in: " & pPowershellProgramServerFilepath
        End If
    Else
        If withDebug Then Debug.Print "[DEBUG] HandlerConsole::createPowershellFileVBAMonologgerServer | The powershell file for VBAMonlogger server already exists in: " & pPowershellProgramServerFilepath
    End If
End Sub

'@Desciption("To start the VBAMonologger HTTP server logs viewer.")
Public Sub startServerLogsViewer( _
    Optional ByVal paramVerbose As Boolean = False _
)
    ' Check if the file exists using FileSystemObject
    If Not pFileSystemObject.FileExists(pPowershellProgramServerFilepath) Then
        Err.Raise vbObjectError + 1000, "HandlerConsole::startServerLogsViewer", "Powershell program file for *VBAMonologger HTTP server logs viewer* was not found: " & pPowershellProgramServerFilepath
        Exit Sub
    End If
    
    Dim shellCommand As String
    shellCommand = "cmd.exe /K"
    shellCommand = shellCommand & " powershell.exe -File """ & pPowershellProgramServerFilepath & """"
    shellCommand = shellCommand & " -port" & " " & portServer
    shellCommand = shellCommand & " -hostname" & " " & hostnameServer
    shellCommand = shellCommand & " -titleConsole """ & TITLE_CONSOLE_WINDOW & """"
    
    If paramVerbose = True Then
        shellCommand = shellCommand & " -Verbose"
        pConsoleWrapper.withDebug = True
    End If
       
    ' If console (cmd.exe) with the title "TitleConsoleWindow" already exist... we do nothing with createConsole!
    pConsoleWrapper.createConsole shellCommand
        
    ' Check server status
    If isServerListening() Then
        If withDebug Then Debug.Print "[DEBUG] HandlerConsole::startServerLogsViewer | The VBAMonologger server is listening on: " & urlServer
    Else
        Err.Raise vbObjectError + 1000, "HandlerConsole::startServerLogsViewer", "Unable to start the VBAMonologger server. It's not listening on: " & urlServer
    End If
End Sub

Public Function isServerListening() As Boolean
    Dim http As Object
    Dim serverStatus As Boolean
    
    On Error GoTo ErrorHandler
    
    ' Try to connect to server
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", urlServer, False
    http.send
    
    ' If open and send method return a value, the server is listening
    serverStatus = (http.Status = 200 Or http.Status = 405) ' 200 = OK, 405 = Method Not Allowed but server is listening
    isServerListening = serverStatus
    If withDebug Then
        If isServerListening Then
            Debug.Print "[DEBUG] HandlerConsole::isServerListening | VBAMonologger server respond with code: " & http.Status
        Else
            Debug.Print "[DEBUG] HandlerConsole::isServerListening | VBAMonologger server does not respond."
        End If
    End If
    
    Exit Function
ErrorHandler:
    ' If an error occured, the server is not listening
    If withDebug Then Debug.Print "[DEBUG] HandlerConsole::isServerListening | The VBAMonologger server is not started. Cannot establish communication with it: " & Err.Description
    isServerListening = False
End Function

'@Description("To send a request to HTTP server")
Public Sub sendRequest(paramMessage As String)
    Dim http As Object
    Dim message As String

    ' Try to connect to VBAMonologger HTTP server logs viewer
    If withDebug Then Debug.Print "[DEBUG] HandlerConsole::sendRequest | Try to send request to VBAMonologger server on: " & urlServer
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", urlServer, False ' False = Request should be handled synchronously
    
    ' Send request to server
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.setRequestHeader "User-Agent", "VBAMonologger VBA client"
    http.setRequestHeader "Cache-Control", "no-cache"
    http.setRequestHeader "Connection", "keep-alive"
    http.send paramMessage
    If withDebug Then Debug.Print "[DEBUG] HandlerConsole::sendRequest | The request was sent to the server with body message: """ & paramMessage & """"
    
    ' Received and read the response's server
    Dim responseBody As String
    Dim responseCode As Integer
    responseBody = http.responseText
    responseCode = http.Status
    If withDebug Then
        Select Case responseCode
            Case 200
                Debug.Print "[DEBUG] HandlerConsole::sendRequest | The server has responded successfully (200 - OK): " & responseBody
            Case 404
                Debug.Print "[DEBUG] HandlerConsole::sendRequest | The server responded an error (404 - resource not found)."
            Case Else
                Debug.Print "[DEBUG] HandlerConsole::sendRequest | Unexpected server response with HTTP code: " & responseCode & " and body response: " & responseBody
        End Select
    End If
    
    If responseCode <> 200 Then
        Err.Raise vbObjectError + 1000, "VBAMonologger.Handler.HandlerConsole::sendRequest", "Unexpected server response with HTTP code: " & responseCode & " and body response: " & responseBody
    End If
End Sub

'@Description("To send a command (keyboard-simulating) to the Window console, in order to stop it.")
Public Sub sendExitCommand()
    pConsoleWrapper.sendCommand "exit"
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
    If Not HandlerInterface_isHandling(paramLogRecord) Then
        result = False
    Else
        Dim formattedMessage As String
        formattedMessage = VBA.CStr$(pformatter.format(paramLogRecord))
        formattedMessage = VBA.Replace$(formattedMessage, "\n", vbCrLf)
        
        ' Call VBAMonologger client logs sender in order to send a request to VBAMonologger server logs viewer (i.e. send a message to show into server's output)
        Call sendRequest(formattedMessage)
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
    output = "[VBAMonologger.Handler.HandlerConsole]" & vbCrLf
    output = output & " | level: " & pLogLevel.name & vbCrLf
    output = output & " | bubble: " & pBubble & vbCrLf
    output = output & " | VBAMonologger temp powershell folder: " & pTempFolderPowershellPrograms & vbCrLf
    output = output & " | VBAMonologger server program : " & pPowershellProgramServerFilepath & vbCrLf
    output = output & " | VBAMonologger server hostname: " & hostnameServer & vbCrLf
    output = output & " | VBAMonologger server port: " & portServer & vbCrLf
    output = output & " | Option withDebug: " & pWithDebug & vbCrLf
    output = output & " | Option withANSIColorSupport: " & pWithANSIColorSupport & vbCrLf
    output = output & " | Option withNewlineForContextAndExtra: " & pWithNewlineForContextAndExtra & vbCrLf
    
    
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
Private Function getPowershellCodeServerLogsViewer() As String
    Dim code As String
    code = ""
    
    ' Header code source file
    code = code & "# ------------------------------------- #" & vbCrLf
    code = code & "#                                       #" & vbCrLf
    code = code & "#    VBA Monologger                     #" & vbCrLf
    code = code & "#    Copyright © 2024, 6i software      #" & vbCrLf
    code = code & "#                                       #" & vbCrLf
    code = code & "# ------------------------------------- #" & vbCrLf
    code = code & "#" & vbCrLf
    code = code & "# HTTP Server for displaying log message sent by request's clients." & vbCrLf
    code = code & "#" & vbCrLf
    code = code & "# The server will do nothing more than display the log message sent by the client in the console." & vbCrLf
    code = code & "##" & vbCrLf
    code = code & vbCrLf
    
    ' Cmdlet declaration
    code = code & "# To enable CLI option -v/-Verbose (and more...)" & vbCrLf
    code = code & "[CmdletBinding()]" & vbCrLf
    code = code & "param (" & vbCrLf
    code = code & "    [Parameter(Position = 0)][string] $hostname = ""localhost""," & vbCrLf
    code = code & "    [Parameter(Position = 1)][int] $port = 20100," & vbCrLf
    code = code & "    [Parameter(Position = 2)][string] $titleConsole = ""VBAMonologger server logs viewer""" & vbCrLf
    code = code & ")" & vbCrLf
    code = code & vbCrLf
    code = code & "Add-Type -AssemblyName System.Net.Http" & vbCrLf
    code = code & vbCrLf

    ' Region global variables
    code = code & "# region - Setting of global variables" & vbCrLf
    code = code & "$Global:prefixConsoleOutput = '[VBAMonologger server]'" & vbCrLf
    code = code & "$Global:defaultForegroundColor = 'White'" & vbCrLf
    code = code & "$Global:defaultBackgroundColor = 'Black'" & vbCrLf
    code = code & "$Global:highlightForegroundColor = 'DarkGreen'" & vbCrLf
    code = code & "$Global:highlightBackgroundColor = 'Black'" & vbCrLf
    code = code & "$Global:hostname = $hostname" & vbCrLf
    code = code & "$Global:port = $port" & vbCrLf
    code = code & "$Global:titleConsole = $titleConsole" & vbCrLf
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
    code = code & "    $Host.UI.RawUI.WindowTitle = $Global:titleConsole" & vbCrLf
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
    
    ' Create HTTP Server
    code = code & "function _createHTTPServer" & vbCrLf
    code = code & "{" & vbCrLf
    code = code & "    try" & vbCrLf
    code = code & "    {" & vbCrLf
    code = code & "        $server = [System.Net.HttpListener]::new()" & vbCrLf
    code = code & "        $server.Prefixes.Add(""http://"" + $Global:hostname + "":"" + $Global:port + ""/"")" & vbCrLf
    code = code & "        $server.Start()" & vbCrLf
    code = code & "        _consoleDebug(""Server is listening on : <h>""""http://"" + $Global:hostname + "":"" + $Global:port + """"""</h>"")" & vbCrLf
    code = code & "    }" & vbCrLf
    code = code & "    catch" & vbCrLf
    code = code & "    {" & vbCrLf
    code = code & "        _consoleError(""Creation of server encountered an <h>critical error</h>. It is possible that the HTTP server's port: <h>"" + $Global:port + ""</h>, is already in use by another application or process.`n$_"")" & vbCrLf
    code = code & "        Exit 1" & vbCrLf
    code = code & "    }" & vbCrLf
    code = code & vbCrLf
    code = code & "    return $server" & vbCrLf
    code = code & "}" & vbCrLf
    code = code & vbCrLf
    
    ' Stop HTTP Server
    code = code & "function _stopHTTPServer" & vbCrLf
    code = code & "{" & vbCrLf
    code = code & "    param ([Parameter(Mandatory = $true)] [System.Net.HttpListener] $server)" & vbCrLf
    code = code & "    $server.Stop()" & vbCrLf
    code = code & "    [console]::WriteLine(""`nServer shutdown, bye bye !"")" & vbCrLf
    code = code & "}" & vbCrLf
    code = code & vbCrLf
    
    ' Start HTTP Server
    code = code & "function _startHTTPServer" & vbCrLf
    code = code & "{" & vbCrLf
    code = code & "    param ([Parameter(Mandatory = $true)] [System.Net.HttpListener] $server)" & vbCrLf
    code = code & "" & vbCrLf
    code = code & "    try" & vbCrLf
    code = code & "    {" & vbCrLf
    code = code & "        $continue = $true" & vbCrLf
    code = code & "        while ($continue)" & vbCrLf
    code = code & "        {" & vbCrLf
    code = code & "            _consoleDebug(""Waiting for new client connection..."")" & vbCrLf
    code = code & "            $context = $server.GetContext()" & vbCrLf
    code = code & "            $request = $context.Request" & vbCrLf
    code = code & "            $response = $context.Response" & vbCrLf
    code = code & "            _consoleDebug(""Connection established by a new client."")" & vbCrLf
    code = code & "" & vbCrLf
    code = code & "            # Read client request (with support encoding UTF-8)" & vbCrLf
    code = code & "            $reader = [System.IO.StreamReader]::new($request.InputStream, [System.Text.Encoding]::UTF8)" & vbCrLf
    code = code & "            $message = $reader.ReadToEnd()" & vbCrLf
    code = code & "            _consoleDebug(""Request received from the new client."")" & vbCrLf
    code = code & "            " & vbCrLf
    code = code & "            # Process client's request" & vbCrLf
    code = code & "            $command = $message.ToLower()" & vbCrLf
    code = code & "            if ($command -eq 'exit' -or $command -eq 'stop' -or $command -eq 'stop-server')" & vbCrLf
    code = code & "            {" & vbCrLf
    code = code & "                $responseString = ""Stop command received and executed by the server.""" & vbCrLf
    code = code & "                $continue = $false" & vbCrLf
    code = code & "            }" & vbCrLf
    code = code & "            else" & vbCrLf
    code = code & "            {" & vbCrLf
    code = code & "                $responseString = ""Log entry received !""" & vbCrLf
    code = code & "                # Simply output message's request into console, in order to show log record" & vbCrLf
    code = code & "                [console]::WriteLine($message)" & vbCrLf
    code = code & "            }" & vbCrLf
    code = code & "" & vbCrLf
    code = code & "            # Add custom headers"
    code = code & "            $response.Headers.Add(""Server"", ""VBAMonologger HTTP Server"")"
    code = code & "            $response.Headers.Add(""X-Powered-By"", ""PowerShell 5"")"
    code = code & "            $response.Headers.Add(""X-Request-Received"", (Get-Date).ToString(""yyyy-MM-ddTHH:mm:ss""))"
    code = code & "" & vbCrLf
    code = code & "            # Send server response (with support encoding UTF-8)" & vbCrLf
    code = code & "            $response.StatusCode = 200" & vbCrLf
    code = code & "            $response.ContentType = ""text/plain; charset=utf-8""" & vbCrLf
    code = code & "            $buffer = [System.Text.Encoding]::UTF8.GetBytes($responseString)" & vbCrLf
    code = code & "            $response.OutputStream.Write($buffer, 0, $buffer.Length)" & vbCrLf
    code = code & "            $response.Close()" & vbCrLf
    code = code & "        }" & vbCrLf
    code = code & "    }" & vbCrLf
    code = code & "    catch" & vbCrLf
    code = code & "    {" & vbCrLf
    code = code & "        _consoleError(""The server encountered an <h>critical error</h>, preventing it from continuing to operate.`n$_"")" & vbCrLf
    code = code & "    }" & vbCrLf
    code = code & "    finally" & vbCrLf
    code = code & "    {" & vbCrLf
    code = code & "        _stopHTTPServer -Server $server" & vbCrLf
    code = code & "    }" & vbCrLf
    code = code & "}" & vbCrLf
    
    ' Close region: Server
    code = code & "# endregion" & vbCrLf
    code = code & "" & vbCrLf
    
    ' Start app
    code = code & "function _main" & vbCrLf
    code = code & "{" & vbCrLf
    code = code & "    _splashscreen" & vbCrLf
    code = code & "" & vbCrLf
    code = code & "    $server = _createHTTPServer" & vbCrLf
    code = code & "    _startHTTPServer -Server $server" & vbCrLf
    code = code & "}" & vbCrLf
    code = code & "" & vbCrLf
    code = code & "_main" & vbCrLf
    
    getPowershellCodeServerLogsViewer = code
End Function
