VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogRecord"
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
'@Folder("VBAMonologger.Log")
'@FQCN("VBAMonologger.Log.LogRecord")
'@ModuleDescription("Represents a log record.")
'
' A log record refers to an entry that captures specific information about an event or an
' action that occurs within your application (See LogRecordInterface).
''

Option Explicit

Implements VBAMonologger.LogRecordInterface
Implements VBAMonologger.StringableInterface

'@Description("The date and time when the log entry was created.")
Private pDateTime As Date
'@Description("The main text describing the event (e.g., "User attempted to log in.").")
Private pMessage As String
'@Description("A channel to categorize and organize log.")
Private pChannel As String
'@Description("The severity level of the log entry (e.g., Info, Warning, Error).")
Private pLevel As VBAMonologger.LogLevel
'@Description("Additional relevant information about the event (e.g., user ID).")
Private pContext As Scripting.Dictionary
'@Description("Optional additional information that might be helpful (e.g., memory usage).")
Private pExtra As Scripting.Dictionary



' ____________ '
'              '
'  Initialize  '
' ____________ '
'              '

Private Sub Class_Initialize()
    pDateTime = VBA.Now()
    pMessage = vbNullString
    pChannel = vbNullString
    Set pLevel = New VBAMonologger.LogLevel
    pLevel.currentLogLevel = LEVEL_DEBUG
    Set pContext = New Scripting.Dictionary
    Set pExtra = New Scripting.Dictionary
End Sub

'@Description('Emulate a class constructor.')
Public Function construct( _
    ByVal paramMessage As String, _
    ByVal paramLevel As VBAMonologger.LOG_LEVELS, _
    Optional paramChannel As String = vbNullString, _
    Optional paramContext As Scripting.IDictionary = Nothing, _
    Optional paramExtra As Scripting.IDictionary = Nothing _
) As VBAMonologger.LogRecordInterface
    pMessage = paramMessage
    pLevel.currentLogLevel = paramLevel
    pChannel = IIf(paramChannel <> vbNullString, paramChannel, pChannel)
    Set pContext = IIf(Not paramContext Is Nothing, paramContext, New Scripting.Dictionary)
    Set pExtra = IIf(Not paramExtra Is Nothing, paramExtra, New Scripting.Dictionary)

    Set construct = Me
End Function



' ___________________________________ '
'                                     '
'  Implementation: LogRecordInterface '
' ___________________________________ '
'                                     '

'@inheritdoc
Private Property Get LogRecordInterface_dateTime() As Date
    LogRecordInterface_dateTime = pDateTime
End Property

'@inheritdoc
Private Property Get LogRecordInterface_message() As String
    LogRecordInterface_message = pMessage
End Property

'@inheritdoc
Private Property Let LogRecordInterface_message(ByVal value As String)
    pMessage = value
End Property

'@inheritdoc
Private Property Get LogRecordInterface_channel() As String
    LogRecordInterface_channel = pChannel
End Property

'@inheritdoc
Private Property Get LogRecordInterface_level() As Variant
    Set LogRecordInterface_level = pLevel
End Property

'@inheritdoc
Private Property Get LogRecordInterface_context() As Variant
    If IsObject(pContext) Then
        Set LogRecordInterface_context = pContext
    Else
        LogRecordInterface_context = pContext
    End If
End Property

'@inheritdoc
Private Property Get LogRecordInterface_extra() As Variant
    If IsObject(pExtra) Then
        Set LogRecordInterface_extra = pExtra
    Else
        LogRecordInterface_extra = pExtra
    End If
End Property

'@Description("Proxy method for public exposition.")
Public Property Get dateTime() As Date
    dateTime = pDateTime
End Property

'@Description("Proxy method for public exposition.")
Public Property Get message() As String
    message = pMessage
End Property

'@Description("Proxy method for public exposition.")
Public Property Let message(ByVal value As String)
    LogRecordInterface_message value
End Property

'@Description("Proxy method for public exposition.")
Public Property Get channel() As String
    channel = pChannel
End Property

'@Description("Proxy method for public exposition.")
Public Property Get level() As LogLevel
    Set level = pLevel
End Property

'@Description("Proxy method for public exposition.")
Public Property Get context() As Variant
    If IsObject(pContext) Then
        Set context = pContext
    Else
        context = pContext
    End If
End Property

'@Description("Proxy method for public exposition.")
Public Property Get extra() As Variant
    If IsObject(pExtra) Then
        Set extra = pExtra
    Else
        extra = pExtra
    End If
End Property



' ____________________________________ '
'                                      '
'  Implementation: StringableInterface '
' ____________________________________ '
'                                      '

'@inheritdoc
Private Function StringableInterface_toString() As String
    Dim output As String
    output = "[VBAMonologger.Log.LogRecord]" & vbCrLf
    output = output & " | dateTime: " & pDateTime & vbCrLf
    output = output & " | message: " & pMessage & vbCrLf
    output = output & " | channel: " & pChannel & vbCrLf
    output = output & " | level: " & pLevel.name & " (level value=" & pLevel.currentLogLevel & ")" & vbCrLf
    output = output & " | context: " & VBAMonologger.ConvertToJson(pContext) & vbCrLf
    output = output & " | extra: " & VBAMonologger.ConvertToJson(pExtra) & vbCrLf
    
    StringableInterface_toString = output
End Function

'@Description("Proxy method for public exposition.")
Public Function toString() As String
    toString = StringableInterface_toString
End Function
