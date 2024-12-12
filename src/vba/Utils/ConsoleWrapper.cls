VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConsoleWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ------------------------------------- '
'                                       '
'    VBA Monologger                     '
'    Copyright © 2024, 6i software      '
'                                       '
' ------------------------------------- '
'
'@Exposed
'@Folder("VBAMonologger.Utils")
'@FQCN("VBAMonologger.Utils.ConsoleWrapper")
'@ModuleDescription("Wraps console functionalities to simplify creation and interaction with a console.")
'
' Wraps console functionalities to interact with it. You can create a new or fetch a console window by its title,
' and also send keyboard-simulated commands to the console. The window search is performed using its title and
' class (e.g., for the cmd.execonsole, the class `ConsoleWindowClass` is used).
''

Option Explicit

' Windows API declarations
#If VBA7 Then
    ' For 64-bit VBA (Excel 64 bits)
    Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As LongPtr, ByVal hWndChildAfter As LongPtr, ByVal lpszClass As String, ByVal lpszWindow As String) As LongPtr
    Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" ( _
        ByVal hwnd As LongPtr, _
        ByVal msg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As String _
    ) As Long
    Private Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" ( _
        ByVal hwnd As LongPtr, _
        ByVal msg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long _
    ) As Long
#Else
    ' For 32-bit VBA (Excel 32 bits)
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
    Private Declare Function GetLastError Lib "kernel32" () As Long
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
        ByVal hwnd As Long, _
        ByVal msg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As String _
    ) As Long
    Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" ( _
        ByVal hwnd As Long, _
        ByVal msg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long _
    ) As Long
#End If

Private Const WM_SETTEXT As Long = &HC
Private Const WM_CHAR As Long = &H102
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101
Private Const VK_RETURN As Long = &HD


' Handle of console wrapper window
Private pHandleConsoleWindow As LongPtr

' The title of console wrapper window
Public titleConsoleWindow As String
' The class of console wrapper window
Public classConsoleWindow As String
' Tempo for waiting the creation of console (in milliseconds)
Public consoleCreationWait As Integer
' Show debug information when we use its instance (for dev only)
Public withDebug As Boolean



' ____________ '
'              '
'  Initialize  '
' ____________ '
'              '
Private Sub Class_Initialize()
    pHandleConsoleWindow = 0
    titleConsoleWindow = "VBAMonologger console wrapper"
    classConsoleWindow = "ConsoleWindowClass"
    consoleCreationWait = 300
    withDebug = True
End Sub

'@Description("Create a new console window, or fetch it if it already exists.")
Public Sub createConsole( _
    Optional paramCommand As String = "", _
    Optional withFindingConsoleBeforeCreation As Boolean = True _
)
    On Error GoTo ErrorHandler
    
    If withFindingConsoleBeforeCreation Then
        Call findConsole(False)
        If Not pHandleConsoleWindow = 0 Then
            If withDebug Then Debug.Print "[DEBUG] ConsoleWrapper::createConsole | Console window with title """ & titleConsoleWindow & """ already exists, so do nothing."
            Exit Sub
        End If
        If withDebug Then Debug.Print "[DEBUG] ConsoleWrapper::createConsole | Console window with title """ & titleConsoleWindow & """ does not exist, need initiating creation."
    End If

    If paramCommand = "" Then
        ' Default creation if no command parameter is given
        Dim defaultCommand As String
        defaultCommand = "cmd.exe /K title " & titleConsoleWindow
        shell defaultCommand, vbMaximizedFocus
    Else
        ' Creation console with given command (don't forget to name the console window with a title)
        shell paramCommand, vbMaximizedFocus
    End If
    
    ' Tempo for waiting the creation of console
    Sleep consoleCreationWait
    
    ' Set pHandleConsoleWindow by calling findConsole
    Call findConsole(False)

    ' In case of default console creation, initialization of the prompt's console
    If paramCommand = "" Then
        Call sendCommand("PROMPT $S" & vbCrLf)
        Call sendCommand("@echo off" & vbCrLf)
        Call sendCommand("cls" & vbCrLf)
    End If

    Exit Sub

ErrorHandler:
    If withDebug Then Debug.Print "[ERROR] ConsoleWrapper::createConsole | Error on creation of console (cmd.exe): " & Err.Description
    Err.Raise vbObjectError + 1000, "ConsoleWrapper::createNewConsole", "Error on creation of console (cmd.exe): " & Err.Description
End Sub

'@Description("Find the console Window, return and set its handle window.")
Public Function findConsole(Optional withRaiseError As Boolean = True) As LongPtr
    Dim handleWindow As LongPtr
    handleWindow = findWindowByClassAndTitle(classConsoleWindow, titleConsoleWindow, withRaiseError)
    
    If withDebug Then
        If handleWindow = 0 Then
            Debug.Print "[DEBUG] ConsoleWrapper::findConsole | Console window with title """ & titleConsoleWindow & """ was not found."
        Else
            Debug.Print "[DEBUG] ConsoleWrapper::findConsole | Console window with title """ & titleConsoleWindow & """ was found."
        End If
    End If
    
    pHandleConsoleWindow = handleWindow
    findConsole = handleWindow
End Function

'@Description("Send a command to the VBAMonologger console Window (simulate keyboard)")
Public Sub sendCommand(command As String)
    Dim charsWritten As LongPtr
    If pHandleConsoleWindow = 0 Then
        Err.Raise vbObjectError + 1000, "ConsoleWrapper::writeConsole", "Unable to write into the console, the private handleConsoleWindow equals zero (not found)."
        Exit Sub
    End If
    Dim i As Integer
    
    command = command & vbCrLf
    For i = 1 To Len(command)
        PostMessage pHandleConsoleWindow, WM_CHAR, Asc(Mid(command, i, 1)), 0
    Next i
    
    ' Simulate press touch return (i.e. VK_RETURN)'
    PostMessage pHandleConsoleWindow, WM_KEYDOWN, VK_RETURN, 0
    PostMessage pHandleConsoleWindow, WM_KEYUP, VK_RETURN, 0
End Sub

Private Function findWindowByClassAndTitle(windowClass As String, windowTitle As String, Optional withRaiseError As Boolean = True) As LongPtr
    ' A handle to the parent window to search the child windows of. To search all windows, specify 0 for this parameter.
    Dim handleWindowParent As LongPtr
    handleWindowParent = 0
    ' A handle to the child window specifying a place to begin searching. Searching begins with the child window immediately after this window in the Z-order. If this is 0, searching begins with the child window at the top of the Z-order.
    Dim handleWindowChild As LongPtr
    handleWindowChild = 0
    ' Find window by its class name and title window
    Dim handleWindow As LongPtr
    handleWindow = FindWindowEx(handleWindowParent, handleWindowChild, windowClass, windowTitle)
    
    ' Debugging error on FindWindowEx, i.e. when it returns zero.
    If withRaiseError Then
        If handleWindow = 0 Then
            Dim lastError As Long
            lastError = GetLastError()
            Err.Raise vbObjectError + 1000 + lastError, "ConsoleWrapper::findConsole", "Failed with FindWindowEx, the handleWindow returns is equal to zero. Code error : " & lastError
        End If
    End If
    
    findWindowByClassAndTitle = handleWindow
End Function
