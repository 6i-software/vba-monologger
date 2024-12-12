VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProcessorUid"
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
'@FQCN("VBAMonologger.Processor.ProcessorUid")
'@ModuleDescription("Adds a unique identifier (UID) into log records.")
'
' Pre-processors to generate and add a random UID into log records. The purpose of this process is to uniquely
' tag each log record with a hexadecimal UID, ensuring traceability and consistency in logging.
'
' The UID remains the same for the lifetime of the **UidProcessor**, unless it is manually recreated. This means
' all log messages recorded by a logger using the same **UidProcessor** will share the same UID, which is useful
' for tracking all actions within a single user session.
''

Option Explicit

Implements ProcessorInterface

Private pUid As String


'_____________________________'
'                             '
'          Initialize         '
'_____________________________'

Private Sub Class_Initialize()
    ' By default, Uid have 8 hexadecimal characters
    pUid = pGenerateUid(8)
End Sub

'@Description('Emulate a class constructor.')
Public Function construct(ByVal Length As Integer) As VBAMonologger.ProcessorInterface
    pUid = pGenerateUid(Length)

    ' Fluant class
    Set construct = Me
End Function

'@Description('Alias method to generate Uid with given length')
Public Sub setLengthUid(ByVal Length As Integer)
    pUid = pGenerateUid(Length)
End Sub
'_____________________________________________________'
'                                                     '
'          Implementation InterfaceProcessor          '
'_____________________________________________________'

'@inheritdoc
Private Sub ProcessorInterface_perform(ByRef paramLogRecord As VBAMonologger.LogRecordInterface)
    Call paramLogRecord.extra.Add("session-UID", pUid)
End Sub

'@Description("Generates a random unique identifier (UID) of specified length, consisting of hexadecimal characters (e.g. A3F7B1C4).")
Private Function pGenerateUid(ByVal Length As Integer) As String
    Dim bytes() As Byte
    Dim uidHex As String
    Dim i As Integer
        
    ReDim bytes(0 To (Length + 1) \ 2 - 1)
    For i = LBound(bytes) To UBound(bytes)
        bytes(i) = Int((256) * Rnd)
    Next i

    uidHex = ""
    For i = LBound(bytes) To UBound(bytes)
        uidHex = uidHex & Right("0" & Hex(bytes(i)), 2)
    Next i

    pGenerateUid = Left(uidHex, Length)
End Function



