VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProcessorInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
'@Folder("VBAMonologger.Processor")
'@FQCN("VBAMonologger.Processor.ProcessorInterface")
'@ModuleDescription("An optional interface to allow labelling processors.")
'@Interface

Option Explicit

Public Sub perform(ByRef paramLogRecord As VBAMonologger.LogRecordInterface)
End Sub


