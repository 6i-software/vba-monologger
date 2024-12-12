VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormatterAnsiColoredLine"
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
'@Folder("VBAMonologger.Formatter")
'@FQCN("VBAMonologger.Formatter.FormatterANSIColoredLine")
'@ModuleDescription("Formats incoming log records into a one-line string representation.")
'
' Formatter of log messages as plain text with ANSI escape sequences color foreach log level.
''

Option Explicit

Implements VBAMonologger.StringableInterface
Implements VBAMonologger.FormatterInterface

Private pFormatterLine As VBAMonologger.FormatterLine
Private pColorScheme As Scripting.Dictionary



' ____________ '
'              '
'  Initialize  '
' ____________ '
'              '

Private Sub Class_Initialize()
    Set pFormatterLine = New VBAMonologger.FormatterLine
    Set pColorScheme = getDefaultColorScheme
End Sub

'@Description('Emulate a class constructor.')
Public Function construct( _
    Optional paramTemplateLine As String = vbNullString _
) As VBAMonologger.FormatterInterface
    If paramTemplateLine <> vbNullString Then
        pFormatterLine.templateLine = paramTemplateLine
    End If
    
    ' Fluent
    Set construct = Me
End Function

'@Description("Get the default color scheme.)
Public Function getDefaultColorScheme() As Scripting.Dictionary
    Dim defaultColorScheme As Scripting.Dictionary
    Set defaultColorScheme = New Scripting.Dictionary
    
    defaultColorScheme.Add "DEBUG", VBAMonologger.ANSI.WHITE
    defaultColorScheme.Add "INFO", VBAMonologger.ANSI.GREEN
    defaultColorScheme.Add "NOTICE", VBAMonologger.ANSI.CYAN
    defaultColorScheme.Add "WARNING", VBAMonologger.ANSI.YELLOW
    defaultColorScheme.Add "ERROR", VBAMonologger.ANSI.BRIGHT_RED
    defaultColorScheme.Add "CRITICAL", VBAMonologger.ANSI.RED
    defaultColorScheme.Add "ALERT", VBAMonologger.ANSI.BG_BRIGHT_RED & VBAMonologger.ANSI.WHITE
    defaultColorScheme.Add "EMERGENCY", VBAMonologger.ANSI.BG_RED & VBAMonologger.ANSI.WHITE
    
    defaultColorScheme.Add "RESET", VBAMonologger.ANSI.RESET
    
    Set getDefaultColorScheme = defaultColorScheme
End Function

'@Description("Get the traffic light color scheme.")
Public Function getTrafficLightColorScheme() As Scripting.Dictionary
    Dim trafficLightColorScheme As Scripting.Dictionary
    Set trafficLightColorScheme = New Scripting.Dictionary
    
    trafficLightColorScheme.Add "DEBUG", Chr$(27) & "[2;32m" ' FG Green faint
    trafficLightColorScheme.Add "INFO", Chr$(27) & "[32m" ' FG Green normal
    trafficLightColorScheme.Add "NOTICE", Chr$(27) & "[1;32m" ' FG Green bright
    trafficLightColorScheme.Add "WARNING", Chr$(27) & "[2;33m" ' FG Yellow faint
    trafficLightColorScheme.Add "ERROR", Chr$(27) & "[33m" ' FG Yellow normal
    trafficLightColorScheme.Add "CRITICAL", Chr$(27) & "[31m" ' FG Red normal
    trafficLightColorScheme.Add "ALERT", Chr$(27) & "[1;31m" ' FG Red bright
    trafficLightColorScheme.Add "EMERGENCY", Chr$(27) & "[41;37;5m"
    trafficLightColorScheme.Add "RESET", Chr$(27) & "[0m"
    
    Set getTrafficLightColorScheme = trafficLightColorScheme
End Function



' ________________ '
'                  '
'  Public methods  '
' ________________ '
'                  '

'@Description("Get the template line used in this formatter.")
Public Property Get templateLine() As String
    templateLine = pFormatterLine.templateLine
End Property

'@Description("Set a new template line to used in this formatter.")
Public Property Let templateLine(ByVal newTemplateLine As String)
    pFormatterLine.templateLine = newTemplateLine
End Property

'@Description("Get the color scheme.")
Public Property Get colorScheme() As Scripting.Dictionary
    colorScheme = pColorScheme
End Property

'@Description("Set the color scheme.")
Public Property Set colorScheme(ByVal newColorScheme As Scripting.Dictionary)
    Set pColorScheme = newColorScheme
End Property

'@Description("Get the ANSI color sequence of given log level according to the current color scheme.")
Public Function GetANSIEscapeSequenceColorOfLogLevel(levelName As String) As String
    levelName = UCase(levelName)
    If pColorScheme.Exists(levelName) Then
        GetANSIEscapeSequenceColorOfLogLevel = pColorScheme(levelName)
    Else
        GetANSIEscapeSequenceColorOfLogLevel = GetANSIEscapeSequenceColorReset ' Default color for reinit
    End If
End Function

'@Description("Get the ANSI color sequence that finishes the line styling.")
Public Function GetANSIEscapeSequenceColorReset() As String
    GetANSIEscapeSequenceColorReset = pColorScheme("RESET")
End Function

'@Description("Get the showContext property from FormatterLine.")
Public Property Get showContext() As Boolean
    showContext = pFormatterLine.showContext
End Property

'@Description("Set the showContext property to used in this formatter.")
Public Property Let showContext(ByVal value As Boolean)
    pFormatterLine.showContext = value
End Property

'@Description("Get the showExtra property from FormatterLine.")
Public Property Get showExtra() As Boolean
    showExtra = pFormatterLine.showExtra
End Property

'@Description("Set the showExtra property to used in this formatter.")
Public Property Let showExtra(ByVal newValue As Boolean)
    pFormatterLine.showExtra = newValue
End Property

'@Description("Get the withWhitespace property from FormatterLine.")
Public Property Get withWhitespace() As Boolean
    withWhitespace = pFormatterLine.withWhitespace
End Property

'@Description("Set the withWhitespace property to used in this formatter.")
Public Property Let withWhitespace(ByVal newValue As Boolean)
    pFormatterLine.withWhitespace = newValue
End Property

'@Description("Get the withAllowingInlineLineBreaks property from FormatterLine.")
Public Property Get withAllowingInlineLineBreaks() As Boolean
    withAllowingInlineLineBreaks = pFormatterLine.withAllowingInlineLineBreaks
End Property

'@Description("Set the withAllowingInlineLineBreaks property to be used in this formatter.")
Public Property Let withAllowingInlineLineBreaks(ByVal newValue As Boolean)
    pFormatterLine.withAllowingInlineLineBreaks = newValue
End Property

'@Description("Set a template line with newline for log context and extra.")
Public Sub setTemplateLineWithNewlineForContextAndExtra()
    pFormatterLine.setTemplateLineWithNewlineForContextAndExtra
End Sub



' ___________________________________ '
'                                     '
'  Implementation: FormatterInterface '
' ___________________________________ '
'                                     '

'@inheritdoc
Private Function FormatterInterface_format(ByRef paramLogRecord As LogRecordInterface) As Variant
    FormatterInterface_format = GetANSIEscapeSequenceColorOfLogLevel(paramLogRecord.level.name) & Trim$(pFormatterLine.format(paramLogRecord)) & GetANSIEscapeSequenceColorReset
End Function

'@inheritdoc
Private Function FormatterInterface_formatBatch(ByRef records() As LogRecordInterface) As Variant
    Dim output As String
    output = ""
    
    Dim i As Integer
    Dim record As VBAMonologger.LogRecordInterface
    For i = LBound(records) To UBound(records)
        Set record = records(i)
        output = output & FormatterInterface_format(record)
        If i <> UBound(records) Then
            output = output & vbCrLf
        End If
    Next i
    
    FormatterInterface_formatBatch = output
End Function

'@Description("Proxy method for public exposition.")
Public Function format(ByRef paramLogRecord As LogRecordInterface) As Variant
    format = FormatterInterface_format(paramLogRecord)
End Function

'@Description("Proxy method for public exposition.")
Public Function formatBatch(ByRef records() As LogRecordInterface) As Variant
    formatBatch = FormatterInterface_formatBatch(records)
End Function



' ____________________________________ '
'                                      '
'  Implementation: StringableInterface '
' ____________________________________ '
'                                      '

'@inheritdoc
Private Function StringableInterface_toString() As String
    Dim output As String
    output = output & "[VBAMonologger.Formatter.FormatterANSIColoredLine]" & vbCrLf
    output = output & " | Template line: " & pFormatterLine.templateLine & vbCrLf
    output = output & " | Show context: " & IIf(pFormatterLine.showContext, "True", "False") & vbCrLf
    output = output & " | Show extra: " & IIf(pFormatterLine.showExtra, "True", "False") & vbCrLf
    
    Dim outputColorScheme As String
    Dim rows() As String
    Dim i As Integer
    outputColorScheme = VBAMonologger.ConvertToJson(pColorScheme, 2)
    rows = Split(outputColorScheme, vbCrLf)
    For i = LBound(rows) To UBound(rows)
        rows(i) = " | " & rows(i)
    Next i
    outputColorScheme = Join(rows, vbCrLf)
    output = output & " | Color sheme: " & vbCrLf & outputColorScheme & vbCrLf
    
    StringableInterface_toString = output
End Function

'@Description("Proxy method for public exposition.")
Public Function toString() As String
    toString = StringableInterface_toString
End Function
