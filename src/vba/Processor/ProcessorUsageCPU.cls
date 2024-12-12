VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProcessorUsageCPU"
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
'@FQCN("VBAMonologger.Processor.ProcessorCPUUsage")
'@ModuleDescription("Adds CPU usage into log records.")
'
' The `ProcessorCPUUsage` class is a VBA pre-processor designed to log system CPU usage data.
' It uses the Windows `GetSystemTimes` API to retrieve details about the system's total, idle,
' and usage CPU times. Users can customize which CPU metrics—total, idle, and usage—are included
' in each log entry. This metadata is then added as a dictionary to the log records into extra property,
' making it easier to monitor and diagnose system CPU usage.
''

Option Explicit

Implements ProcessorInterface


'___________________________________________________________________'
'                                                                   '
'          Windows API declaration - GetSystemTimes                 '
'___________________________________________________________________'

Private Declare PtrSafe Function GetSystemTimes Lib "kernel32.dll" ( _
    lpIdleTime As Currency, _
    lpKernelTime As Currency, _
    lpUserTime As Currency) As Long
      

'_____________________________________________________'
'                                                     '
'          Implementation ProcessorInterface          '
'_____________________________________________________'

'@inheritdoc
Private Sub ProcessorInterface_perform(ByRef paramLogRecord As VBAMonologger.LogRecordInterface)
    Dim currentCPUUsage As Double
    Dim currentCPUUsageFormatted As String
    
    currentCPUUsage = GetCPUUsage()
    If currentCPUUsage >= 0 Then
        currentCPUUsageFormatted = format(currentCPUUsage, "0.0") & "%"
    Else
        currentCPUUsageFormatted = "Unable to retrieve CPU usage."
    End If
    
    Call paramLogRecord.extra.Add("CPU-used", currentCPUUsageFormatted)
End Sub

'@Description("Get the CPU usage.")
Private Function GetCPUUsage() As Double
    Static IdleTime1 As Currency, KernelTime1 As Currency, UserTime1 As Currency
    Dim IdleTime2 As Currency, KernelTime2 As Currency, UserTime2 As Currency
    Dim IdleDiff As Double, KernelDiff As Double, UserDiff As Double, TotalDiff As Double

    ' Récupère les valeurs actuelles
    If GetSystemTimes(IdleTime2, KernelTime2, UserTime2) Then
        If IdleTime1 <> 0 And KernelTime1 <> 0 And UserTime1 <> 0 Then
            ' Calcul des différences
            IdleDiff = (IdleTime2 - IdleTime1) * 10000
            KernelDiff = (KernelTime2 - KernelTime1) * 10000
            UserDiff = (UserTime2 - UserTime1) * 10000

            ' Calcul du total
            TotalDiff = KernelDiff + UserDiff

            ' Calcul de l'utilisation CPU (en pourcentage)
            If TotalDiff > 0 Then
                GetCPUUsage = 100 - (IdleDiff / TotalDiff * 100)
            End If
        End If

        ' Stocke les valeurs actuelles pour le prochain appel
        IdleTime1 = IdleTime2
        KernelTime1 = KernelTime2
        UserTime1 = UserTime2
    Else
        ' Si la fonction échoue
        GetCPUUsage = -1
    End If
End Function


