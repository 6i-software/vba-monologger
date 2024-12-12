VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogLevel"
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
'@FQCN("VBAMonologger.Log.LogLevel")
'@ModuleDescription("Represents a log level.")
'
' To get the level name or value out of a LogLevel there are several options:
'
' - Use the 'Name' property in order to get the standard name which is full uppercased (e.g. "INFO")
' - Use ->ToPsrLogLevel() to get the standard PSR-3 name which is full lowercased (e.g. "notice")
' - Use ->ToRFC5424Level() to get the standard RFC 5424 value (e.g. 0 for "emergency")
'
' To get the internal value for filtering, if the includes/ isLowerThan/ isHigherThan methods
' are not enough, you can use the property 'CurrentLogLevel' to get the enum case's integer value.
''

Option Explicit

Implements VBAMonologger.LogLeveLInterface
Implements VBAMonologger.StringableInterface

'@Description("Represents the current log level.")
Public currentLogLevel As VBAMonologger.LOG_LEVELS


' ____________ '
'              '
'  Initialize  '
' ____________ '
'              '
Private Sub Class_Initialize()
    currentLogLevel = LEVEL_DEBUG
End Sub

'@Description('Emulate a class constructor.')
Public Function construct(paramLogLevel As VBAMonologger.LOG_LEVELS)
    currentLogLevel = paramLogLevel
    
    Set construct = Me
End Function


' _______________ '
'                 '
'  Public method  '
' _______________ '
'                 '
'@Description("Get the current level as standardized all-capitals human readable format.")
Public Property Get name() As String
    Select Case currentLogLevel
        Case VBAMonologger.LOG_LEVELS.LEVEL_DEBUG
            name = "DEBUG"
        Case VBAMonologger.LOG_LEVELS.LEVEL_INFO
            name = "INFO"
        Case VBAMonologger.LOG_LEVELS.LEVEL_NOTICE
            name = "NOTICE"
        Case VBAMonologger.LOG_LEVELS.LEVEL_WARNING
            name = "WARNING"
        Case VBAMonologger.LOG_LEVELS.LEVEL_ERROR
            name = "ERROR"
        Case VBAMonologger.LOG_LEVELS.LEVEL_CRITICAL
            name = "CRITICAL"
        Case VBAMonologger.LOG_LEVELS.LEVEL_ALERT
            name = "ALERT"
        Case VBAMonologger.LOG_LEVELS.LEVEL_EMERGENCY
            name = "EMERGENCY"
        Case Else
            Err.Raise vbObjectError + 1000, "LoggerLevel::Name", "Unrecognized current log level: " & currentLogLevel
    End Select
End Property

'@Description("Set the current level from a given string.")
Public Sub fromName(name As String)
    Select Case UCase(name)
        Case "DEBUG"
            currentLogLevel = VBAMonologger.LOG_LEVELS.LEVEL_DEBUG
        Case "INFO"
            currentLogLevel = VBAMonologger.LOG_LEVELS.LEVEL_INFO
        Case "NOTICE"
            currentLogLevel = VBAMonologger.LOG_LEVELS.LEVEL_NOTICE
        Case "WARNING"
            currentLogLevel = VBAMonologger.LOG_LEVELS.LEVEL_WARNING
        Case "ERROR"
            currentLogLevel = VBAMonologger.LOG_LEVELS.LEVEL_ERROR
        Case "CRITICAL"
            currentLogLevel = VBAMonologger.LOG_LEVELS.LEVEL_CRITICAL
        Case "ALERT"
            currentLogLevel = VBAMonologger.LOG_LEVELS.LEVEL_ALERT
        Case "EMERGENCY"
            currentLogLevel = VBAMonologger.LOG_LEVELS.LEVEL_EMERGENCY
        Case Else
            Err.Raise vbObjectError + 1000, "LoggerLevel::fromName", "Unrecognized given level name: " & name
    End Select
End Sub

'@Description("Get the current level corresponding to RFC5424 standards.")
Public Function toRFC5424Level() As Long
    Select Case currentLogLevel
        Case VBAMonologger.LOG_LEVELS.LEVEL_DEBUG
            toRFC5424Level = 7
        Case VBAMonologger.LOG_LEVELS.LEVEL_INFO
            toRFC5424Level = 6
        Case VBAMonologger.LOG_LEVELS.LEVEL_NOTICE
            toRFC5424Level = 5
        Case VBAMonologger.LOG_LEVELS.LEVEL_WARNING
            toRFC5424Level = 4
        Case VBAMonologger.LOG_LEVELS.LEVEL_ERROR
            toRFC5424Level = 3
        Case VBAMonologger.LOG_LEVELS.LEVEL_CRITICAL
            toRFC5424Level = 2
        Case VBAMonologger.LOG_LEVELS.LEVEL_ALERT
            toRFC5424Level = 1
        Case VBAMonologger.LOG_LEVELS.LEVEL_EMERGENCY
            toRFC5424Level = 0
        Case Else
            Err.Raise vbObjectError + 1000, "LoggerLevel::ToRFC5424Level", "Unrecognized current log level: " & currentLogLevel
    End Select
End Function

'@Description("Get the current level corresponding to PSRLog standards.")
Public Function toPsrLogLevel() As String
    Select Case currentLogLevel
        Case VBAMonologger.LOG_LEVELS.LEVEL_DEBUG
            toPsrLogLevel = "debug"
        Case VBAMonologger.LOG_LEVELS.LEVEL_INFO
            toPsrLogLevel = "info"
        Case VBAMonologger.LOG_LEVELS.LEVEL_NOTICE
            toPsrLogLevel = "notice"
        Case VBAMonologger.LOG_LEVELS.LEVEL_WARNING
            toPsrLogLevel = "warning"
        Case VBAMonologger.LOG_LEVELS.LEVEL_ERROR
            toPsrLogLevel = "error"
        Case VBAMonologger.LOG_LEVELS.LEVEL_CRITICAL
            toPsrLogLevel = "critical"
        Case VBAMonologger.LOG_LEVELS.LEVEL_ALERT
            toPsrLogLevel = "alert"
        Case VBAMonologger.LOG_LEVELS.LEVEL_EMERGENCY
            toPsrLogLevel = "emergency"
        Case Else
            Err.Raise vbObjectError + 1000, "LoggerLevel::ToPsrLogLevel", "Unrecognized current log level: " & currentLogLevel
    End Select
End Function

'@Description("Checks if the current log level is higher or equal to the given level.")
Public Function includes(level As LogLevel) As Boolean
    includes = (currentLogLevel <= level.currentLogLevel)
End Function

'@Description("Checks if the current log level is higher than the given level.")
Public Function isHigherThan(level As LogLevel) As Boolean
    isHigherThan = (currentLogLevel > level.currentLogLevel)
End Function

'@Description("Checks if the current log level is lower than the given level.")
Public Function isLowerThan(level As LogLevel) As Boolean
    isLowerThan = (currentLogLevel < level.currentLogLevel)
End Function



' ____________________________________ '
'                                      '
'  Implementation: StringableInterface '
' ____________________________________ '
'                                      '

'@inheritdoc
Private Function StringableInterface_toString() As String
    Dim output As String
    output = ""
    output = output & "[VBAMonologger.Log.LogLevel]" & vbCrLf
    output = output & " | currentLogLevel: " & currentLogLevel & vbCrLf
    output = output & " | name: '" & name() & "'" & vbCrLf
    'output = output & " | Object type: '" & TypeName(Me) & "'" & vbCrLf
    'output = output & " | Memory address: " & VarPtr(Me) & vbCrLf
    
    StringableInterface_toString = output
End Function

'@Description("Proxy method for public exposition.")
Public Function toString() As String
    toString = StringableInterface_toString
End Function
