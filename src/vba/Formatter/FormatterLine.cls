VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormatterLine"
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
'@FQCN("VBAMonologger.Formatter.FormatterLine")
'@ModuleDescription("Formats incoming log records into a one-line string representation.")
'
' Formatter of log messages as plain text, typically outputting each log entry on a single line.
' It is designed to make logs readable and structured, facilitating easier review in log files
' or console outputs (Excel's VBA Immediate Window).
''

Option Explicit

Implements VBAMonologger.StringableInterface
Implements VBAMonologger.FormatterInterface

Private Const TEMPLATE_LINE_DEFAULT As String = "[{{ datetime }}] {{ channel <./>}}{{ level_name }}: {{ message }}{{< | context: /> context }}{{< | extra: /> extra }}"
Private Const TEMPLATE_LINE_WITH_NEWLINE_FOR_CONTEXT_AND_EXTRA = "[{{ datetime }}] {{ channel <./>}}{{ level_name }}: {{ message }}{{<\n | context: /> context }}{{<\n | extra: /> extra }}"

'@Description("The template used to format the line")
Private pTemplateLine As String
'@Description("Whitespace for "pretty" print json with given number of spaces per indentation")
Private pWhitespace As Variant
'@Description("Show or hide context in line formatted (default: True).")
Public showContext As Boolean
'@Description("Show or hide extra metadatas in line formatted (default: True).")
Public showExtra As Boolean
'@Description("Whether to allow inline line breaks in log entries")
Public withAllowingInlineLineBreaks As Boolean
'@Description("Enable formatted JSON output with whitespace for log context and extra. They are displaying on multiple lines for better readability.")
Public withWhitespace As Boolean
'@Description("Add a prefix to each line of the JSON string representation of log context and extra")
Public withPrefixToJsonLines As String



' ____________ '
'              '
'  Initialize  '
' ____________ '
'              '

Private Sub Class_Initialize()
    pTemplateLine = TEMPLATE_LINE_DEFAULT
    showContext = True
    showExtra = True
    withWhitespace = False
    withAllowingInlineLineBreaks = False
    withPrefixToJsonLines = " | "
End Sub

'@Description('Emulate a class constructor.')
Public Function construct( _
    Optional paramTemplateLine As String = vbNullString, _
    Optional paramWhitespace As Variant = Nothing _
) As VBAMonologger.FormatterInterface
    pTemplateLine = IIf(paramTemplateLine <> vbNullString, paramTemplateLine, pTemplateLine)
    
    If Not IsMissing(paramWhitespace) Then
        If VBA.VarType(paramWhitespace) = VBA.vbInteger Then
            pWhitespace = paramWhitespace
        End If
    End If
    
    ' Fluant class
    Set construct = Me
End Function



' ________________ '
'                  '
'  Public methods  '
' ________________ '
'                  '

'@Description("Get the template line used in this formatter.")
Public Property Get templateLine() As String
    templateLine = pTemplateLine
End Property

'@Description("Set a new template line to used in this formatter.")
Public Property Let templateLine(ByVal newTemplateLine As String)
    pTemplateLine = newTemplateLine
End Property

'@Description("Reset to the default template line.")
Public Sub setTemplateLineDefault()
    pTemplateLine = TEMPLATE_LINE_DEFAULT
End Sub

'@Description("Set a template line with newline for log context and extra.")
Public Sub setTemplateLineWithNewlineForContextAndExtra()
    pTemplateLine = TEMPLATE_LINE_WITH_NEWLINE_FOR_CONTEXT_AND_EXTRA
End Sub


' ___________________________________ '
'                                     '
'  Implementation: FormatterInterface '
' ___________________________________ '
'                                     '

'@inheritdoc
Private Function FormatterInterface_format(ByRef paramLogRecord As LogRecordInterface) As Variant
    Dim output As String
    output = ""
    
    ' Remove new lines
    If Not withAllowingInlineLineBreaks Then
        pTemplateLine = Replace$(pTemplateLine, "\n", "")
    End If

    ' Put template pattern in output line, and after apply values for each pattern ({{datetime}}, {{message}} ...)
    output = pTemplateLine
    output = pApplyPlaceholder(output, "datetime", VBA.format$(paramLogRecord.dateTime, "yyyy\/mm\/dd hh\:nn\:ss"))
    output = pApplyPlaceholder(output, "level_name", paramLogRecord.level.name)
    output = pApplyPlaceholder(output, "message", paramLogRecord.message)
    
    ' Handle {{channel}} placeholder
    If paramLogRecord.channel = "" Then
        ' If no channel value given replace its pre-motif, post-motif and placeholder by an empty string
        output = pRemoveEmptyPlaceholder(output, "channel")
    Else
        output = pApplyPlaceholder(output, "channel", paramLogRecord.channel)
    End If
    
    ' Handle {{context}} placeholder
    If (paramLogRecord.context.count = 0) Or (showContext = False) Then
        output = pRemoveEmptyPlaceholder(output, "context")
    Else
        Dim contextFormatted As String
        If withWhitespace And withAllowingInlineLineBreaks Then
            contextFormatted = VBAMonologger.ConvertToJson(paramLogRecord.context, 3)
            If Not withPrefixToJsonLines = "" Then
                contextFormatted = pAddPrefixToJsonLines(contextFormatted, withPrefixToJsonLines)
            End If
        Else
            contextFormatted = VBAMonologger.ConvertToJson(paramLogRecord.context)
        End If
                
        output = pApplyPlaceholder(output, "context", contextFormatted)
    End If
    
    ' Handle {{extra}} placeholder
    If (paramLogRecord.extra.count = 0) Or (showExtra = False) Then
        output = pRemoveEmptyPlaceholder(output, "extra")
    Else
        Dim extraFormatted As String
        If withWhitespace And withAllowingInlineLineBreaks Then
            extraFormatted = VBAMonologger.ConvertToJson(paramLogRecord.extra, 3)
            If Not withPrefixToJsonLines = "" Then
                extraFormatted = pAddPrefixToJsonLines(extraFormatted, withPrefixToJsonLines)
            End If
        Else
            extraFormatted = VBAMonologger.ConvertToJson(paramLogRecord.extra)
        End If

        output = pApplyPlaceholder(output, "extra", extraFormatted)
    End If
                    
    FormatterInterface_format = output
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

'@Description("Replaces placeholder, premotif and postmotif in template line with its associating value.")
Private Function pApplyPlaceholder(ByVal output As String, placeholder As String, value As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = True
        .IgnoreCase = True
        .pattern = "({{\s*)(<([^}{]*)\/>)?(\s*" & placeholder & "\s*)(<([^}{]*)\/>)?(\s*}})"
    End With
    output = regex.Replace(output, "$3" & value & "$6")
    
    pApplyPlaceholder = output
End Function

'@Description("Remove placeholder, premotif and postmotif in template line if no value is given for this placeholder.")
Private Function pRemoveEmptyPlaceholder(ByVal output As String, placeholder As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = True
        .IgnoreCase = True
        .pattern = "({{\s*(<[^}{]*\/>)?(\s*" & placeholder & "\s*)(<[^}{]*\/>)?\s*}})"
    End With
    
    pRemoveEmptyPlaceholder = regex.Replace(output, "")
End Function

'@Description("Helper to fetch the prefix from the template line based on the placeholder name.")
Private Function pGetPrefixFromTemplate(ByVal templateLine As String, ByVal placeholder As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    Dim match As Object
    
    regex.IgnoreCase = True
    regex.Global = True
    regex.pattern = "({{\s*)(<([^}{]*)\/>)?(\s*" & placeholder & "\s*)(<([^}{]*)\/>)?(\s*}})"
    
    ' Search for the prefix in the templateLine
    Set match = regex.execute(templateLine)
    If match.count > 0 Then
        pGetPrefixFromTemplate = match(0).Submatches(2)
    Else
        pGetPrefixFromTemplate = ""
    End If
End Function

'@Description("Helper to add a custom prefix prefix on each lines of JSON representation.")
Private Function pAddPrefixToJsonLines(ByVal json As String, ByVal prefix As String) As String
    Dim lines() As String
    Dim i As Integer
    Dim output As String
        
    ' Get all JSON lines
    lines = Split(json, vbCrLf)
    
    ' Add the prefix to each line
    output = vbCrLf & ""
    For i = LBound(lines) To UBound(lines)
        output = output & prefix & lines(i)
        
        If i < UBound(lines) Then
            output = output & vbCrLf
        End If
    Next i
    
    pAddPrefixToJsonLines = output
End Function



' ____________________________________ '
'                                      '
'  Implementation: StringableInterface '
' ____________________________________ '
'                                      '

'@inheritdoc
Private Function StringableInterface_toString() As String
    Dim output As String
    output = output & "[VBAMonologger.Formatter.FormatterLine]" & vbCrLf
    output = output & " | Template line: " & pTemplateLine & vbCrLf
    output = output & " | Show context: " & IIf(showContext, "True", "False") & vbCrLf
    output = output & " | Show extra: " & IIf(showExtra, "True", "False") & vbCrLf
    
    StringableInterface_toString = output
End Function

'@Description("Proxy method for public exposition.")
Public Function toString() As String
    toString = StringableInterface_toString
End Function

