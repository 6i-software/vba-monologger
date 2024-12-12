VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HandlerConsoleVBA"
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
'@FQCN("VBAMonologger.Handler.HandlerConsoleVBA")
'@ModuleDescription("Handles logs message to the VBA Console (a.k.a. the Excel's immediate Window).")
''

Option Explicit

Implements VBAMonologger.HandlerInterface
Implements VBAMonologger.StringableInterface

'@Description("The associated formatter of current handler.")
Private pformatter As VBAMonologger.FormatterInterface
'@Description("The bubbling behavior. True means that this handler allows bubbling, and false means that bubbling is not permitted.")
Private pBubble As Boolean
'@Description("The minimum logging level at which this handler will be triggered as a LogLevel")
Private pLogLevel As VBAMonologger.LogLevel
'@Description("The minimum logging level at which this handler will be triggered as an value of enumeration of LOG_LEVELS.")
Private pLevel As VBAMonologger.LOG_LEVELS


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
End Sub

'@Description('Emulate a class constructor.')
Public Function construct( _
    Optional ByVal paramBubble As Boolean = True, _
    Optional ByVal paramLevel As VBAMonologger.LOG_LEVELS = VBAMonologger.LOG_LEVELS.LEVEL_DEBUG _
) As VBAMonologger.HandlerInterface
    pBubble = paramBubble
    pLogLevel.currentLogLevel = paramLevel
    pLevel = paramLevel
    
    Set construct = Me
End Function



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
        
        Debug.Print formattedMessage
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
    output = "[VBAMonologger.Handler.HandlerConsoleVBA]" & vbCrLf
    output = output & " | level: " & pLogLevel.name & vbCrLf
    output = output & " | bubble: " & pBubble & vbCrLf
    
    StringableInterface_toString = output
End Function

'@Description("Proxy method for public exposition.")
Public Function toString() As String
    toString = StringableInterface_toString
End Function
