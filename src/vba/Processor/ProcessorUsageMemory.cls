VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProcessorUsageMemory"
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
'@FQCN("VBAMonologger.Processor.ProcessorUsageMemory")
'@ModuleDescription("Adds memory usage into log records.")
'
' The `ProcessorUsageMemory` class is a VBA pre-processor designed to log system memory usage
' data. It uses the Windows `GlobalMemoryStatusEx` API to retrieve details about the system's
' total, available, and percentage of used memory. Users can customize which memory metrics—total
' memory, available memory, and memory usage percentage—are included in each log entry. This
' metadata is then added as a dictionary to the log records into extra property, making it easier
' to monitor and diagnose system memory usage.
''

Option Explicit

Implements ProcessorInterface


'___________________________________________________________________'
'                                                                   '
'          Windows API declaration - GlobalMemoryStatusEx           '
'___________________________________________________________________'

Private Type MEMORYSTATUSEX
    dwLength As Long
    dwMemoryLoad As Long
    ullTotalPhys As LongLong
    ullAvailPhys As LongLong
    ullTotalPageFile As LongLong
    ullAvailPageFile As LongLong
    ullTotalVirtual As LongLong
    ullAvailVirtual As LongLong
    ullAvailExtendedVirtual As LongLong
End Type

Private Declare PtrSafe Function GlobalMemoryStatusEx Lib "kernel32" (ByRef lpBuffer As MEMORYSTATUSEX) As Long

'@Description("Adds or remove details information on memory usage (total memory, available memory)")
Private pWithDetails As Boolean
'@Description("Adds or remove total memory in extra log, only when details is enabled.")
Private pWithTotalMemory As Boolean
'@Description("Adds or remove available memory in extra log, only when details is enabled.")
Private pWithAvailableMemory  As Boolean


'_____________________________'
'                             '
'          Initialize         '
'_____________________________'

Private Sub Class_Initialize()
    withDetails = False
    withTotalMemory = True
    withAvailableMemory = True
End Sub

'@Description('Emulate a class constructor.')
Public Function construct( _
    Optional ByVal paramWithDetails As Boolean = True, _
    Optional ByVal paramWithTotalMemory As Boolean = True, _
    Optional ByVal paramWithAvailableMemory As Boolean = True _
) As VBAMonologger.ProcessorInterface
    withDetails = paramWithDetails
    withTotalMemory = paramWithTotalMemory
    withAvailableMemory = paramWithAvailableMemory

    ' Fluant
    Set construct = Me
End Function


'______________________________'
'                              '
'          Properties          '
'______________________________'


Public Property Get withDetails() As Boolean
    withDetails = pWithDetails
End Property

Public Property Let withDetails(ByVal value As Boolean)
    pWithDetails = value
End Property

Public Property Get withTotalMemory() As Boolean
    withTotalMemory = pWithTotalMemory
End Property

Public Property Let withTotalMemory(ByVal value As Boolean)
    pWithTotalMemory = value
End Property

Public Property Get withAvailableMemory() As Boolean
    withAvailableMemory = pWithAvailableMemory
End Property

Public Property Let withAvailableMemory(ByVal value As Boolean)
    pWithAvailableMemory = value
End Property


'_____________________________________________________'
'                                                     '
'          Implementation ProcessorInterface          '
'_____________________________________________________'

'@inheritdoc
Private Sub ProcessorInterface_perform(ByRef paramLogRecord As VBAMonologger.LogRecordInterface)
    Dim memoryInfo As Scripting.Dictionary
    Set memoryInfo = GetMemoryStatus()
        
    If memoryInfo.Exists("error") Then
        Err.Raise vbObjectError + 1000, "ProcessorUsageMemory::ProcessorInterface_Invoke", memoryInfo("error")
    Else
        If pWithDetails Then
            Dim memoryInfoChoosen As Scripting.Dictionary
            Set memoryInfoChoosen = New Scripting.Dictionary
            
            memoryInfoChoosen.Add "memory-used", memoryInfo("memory-used")
            
            If withTotalMemory Then
                memoryInfoChoosen.Add "memory-total", memoryInfo("memory-total")
            End If
            
            If withAvailableMemory Then
                memoryInfoChoosen.Add "memory-available", memoryInfo("memory-available")
            End If
            
            ' Add a memory dictionnary for details
            paramLogRecord.extra.Add "memory", memoryInfoChoosen
        Else
            ' Add simply memory-used (no details). It's the default mode of this pre-processor.
            paramLogRecord.extra.Add "memory-used", memoryInfo("memory-used")
        End If
    End If
End Sub


' ---
' This function retrieves the current memory status of the system by calling the
' GlobalMemoryStatusEx API from Windows. It returns a dictionary containing
' formatted information about the total memory, available memory, and the
' percentage of memory used.
'
' Returns:
'   A Scripting.Dictionary containing the following key-value pairs:
'   - "total-Memory": Total physical memory available in gigabytes (GB), formatted to two decimal places.
'   - "available-Memory": Available physical memory in gigabytes (GB),formatted to two decimal places.
'   - "used-memory": Percentage of memory currently in use.
'   - "error": A string indicating an error if the memory status cannot be retrieved.
'
'@Description("Retrieves the current memory status.")
'---
Private Function GetMemoryStatus() As Scripting.Dictionary
    Dim memoryStatusFormatted As Scripting.Dictionary
    Set memoryStatusFormatted = New Scripting.Dictionary
    
    Dim memoryStatus As MEMORYSTATUSEX
    memoryStatus.dwLength = Len(memoryStatus)
    
    If GlobalMemoryStatusEx(memoryStatus) Then
        Dim totalMemory As Double
        Dim availableMemory As Double
        Dim usedMemory As Double
        
        totalMemory = CDbl(memoryStatus.ullTotalPhys) / (1024# * 1024# * 1024#)
        availableMemory = CDbl(memoryStatus.ullAvailPhys) / (1024# * 1024# * 1024#)
        usedMemory = memoryStatus.dwMemoryLoad
                
        memoryStatusFormatted.Add "memory-total", FormatNumber(totalMemory, 2)
        memoryStatusFormatted.Add "memory-available", FormatNumber(availableMemory, 2)
        memoryStatusFormatted.Add "memory-used", usedMemory & "%"
    Else
        memoryStatusFormatted.Add "error", "Windows API GlobalMemoryStatusEx error - Unable to retrieve the memory status."
    End If
    
    Set GetMemoryStatus = memoryStatusFormatted
End Function
