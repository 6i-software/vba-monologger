VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProcessorTags"
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
'@Folder("VBAMonologger.Processor")
'@FQCN("VBAMonologger.Processor.ProcessorTags")
'@ModuleDescription("Adds tags into log records.")
'
' The `ProcessorTags` class is a VBA pre-processors component designed to add tags to log records.
' It stores tags as key-value pairs in a `Scripting.Dictionary`, allowing each log entry to be
' enriched with contextual information.
''

Option Explicit

Implements ProcessorInterface

Public Enum TAGS_DESTINATION
    LOG_CONTEXT = 1
    LOG_EXTRA = 2
    LOG_BOTH = LOG_CONTEXT Or LOG_EXTRA
End Enum

Private pTags As Scripting.Dictionary
Public tagsDestination As TAGS_DESTINATION
Public keepKeyTags As Boolean


'_____________________________'
'                             '
'          Initialize         '
'_____________________________'

Private Sub Class_Initialize()
    Set pTags = New Scripting.Dictionary
    tagsDestination = TAGS_DESTINATION.LOG_CONTEXT
    keepKeyTags = True
End Sub

'@Description('Emulate a class constructor.')
Public Function construct( _
    Optional paramTags As Scripting.IDictionary = Nothing, _
    Optional paramTagsDestination As TAGS_DESTINATION = TAGS_DESTINATION.LOG_CONTEXT _
) As VBAMonologger.ProcessorInterface
    Set pTags = IIf(Not paramTags Is Nothing, paramTags, pTags)
    tagsDestination = paramTagsDestination

    ' Fluent class
    Set construct = Me
End Function


'_________________________'
'                         '
'          Public         '
'_________________________'

'@Description("Add new tags into log processor.")
Public Function addTags(ByRef newTags As Scripting.Dictionary) As VBAMonologger.ProcessorInterface
    Dim key As Variant
    
    For Each key In newTags.Keys
        If Not pTags.Exists(key) Then
            pTags.Add key, newTags(key)
        End If
    Next key

    ' Fluent class
    Set addTags = Me
End Function

'@Description("Set tags into log processor by overwritting existing tags.")
Public Function setTags(ByRef newTags As Scripting.Dictionary) As VBAMonologger.ProcessorInterface
    Set pTags = New Scripting.Dictionary
    Dim key As Variant
    For Each key In newTags.Keys
        pTags.Add key, newTags(key)
    Next key

    ' Fluent class
    Set setTags = Me
End Function


'_____________________________________________________'
'                                                     '
'          Implementation ProcessorInterface          '
'_____________________________________________________'

'@inheritdoc
Private Sub ProcessorInterface_perform(ByRef paramLogRecord As VBAMonologger.LogRecordInterface)
    Dim key As Variant
    
    If tagsDestination = TAGS_DESTINATION.LOG_CONTEXT Or tagsDestination = TAGS_DESTINATION.LOG_BOTH Then
        If keepKeyTags Then
            paramLogRecord.context.Add "tags", pTags
        Else
            For Each key In pTags.Keys
                paramLogRecord.context.Add key, pTags(key)
            Next key
        End If
    End If

    If tagsDestination = TAGS_DESTINATION.LOG_EXTRA Or tagsDestination = TAGS_DESTINATION.LOG_BOTH Then
        If keepKeyTags Then
            paramLogRecord.extra.Add "tags", pTags
        Else
            For Each key In pTags.Keys
                paramLogRecord.extra.Add key, pTags(key)
            Next key
        End If
    End If
End Sub
