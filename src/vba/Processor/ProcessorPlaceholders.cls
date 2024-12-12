VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProcessorPlaceholders"
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
'@Folder("VBAMonologger.Processor")
'@FQCN("VBAMonologger.Processor.ProcessorPlaceholders")
'@ModuleDescription("Allow to use placeholders in log message that will be replaced with value provided in the log record context.")
'
' Processes a record's message according to PSR-3 rules. It replaces the placeholder {foo} with the value
' from log record context dictionnary `context`.
'
' @TODO: Handle nested placeholoders {foo.bar.team}
''

Option Explicit

Implements ProcessorInterface


'_____________________________________________________'
'                                                     '
'          Implementation ProcessorInterface          '
'_____________________________________________________'

'@inheritdoc
Private Sub ProcessorInterface_perform(ByRef paramLogRecord As VBAMonologger.LogRecordInterface)
    ' Nothing to do, if log record does not have a context or a relevant placeholders
    If (paramLogRecord.context Is Nothing) Or (paramLogRecord.message Like "*{*" = False) Then
        Exit Sub
    End If
    
    Dim context As Scripting.Dictionary
    Set context = paramLogRecord.context

    Dim messageWithReplacements As String
    messageWithReplacements = paramLogRecord.message
    
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
    End With
    
    Dim key As Variant
    Dim value As Variant
    Dim placeholder As String
    Dim replacement As String
    
    For Each key In context.Keys
        placeholder = "{" & key & "}"
        
        ' Check if placeholder exists in the message
        If InStr(1, messageWithReplacements, placeholder) > 0 Then
            value = context.item(key)
            
            If IsNull(value) Then
                replacement = "{null}"
            ElseIf IsDate(value) Then
                replacement = format(value, "yyyy-mm-dd HH:MM:SS")
            ElseIf IsArray(value) Then
                replacement = "[" & Join(value, ", ") & "]"
            ElseIf IsObject(value) Then
                replacement = "{(object: " & TypeName(value) & ")}"
            Else
                replacement = VBA.CStr(value)
            End If
                
            ' Replace the placeholder with the formatted value
            regex.pattern = "\{" & key & "\}"
            messageWithReplacements = regex.Replace(messageWithReplacements, replacement)
        End If
    Next key
    
    ' Update the log record message with replacements
    paramLogRecord.message = messageWithReplacements
End Sub
 

