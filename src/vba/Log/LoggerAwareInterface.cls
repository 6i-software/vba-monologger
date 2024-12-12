VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoggerAwareInterface"
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
'@Folder("VBAMonologger.Log")
'@FQCN("VBAMonologger.Log.LoggerAwareInterface")
'@ModuleDescription("Interface for describes a logger-aware instance.")
'@Interface
'
' This interface is used to enable dependency injection for logging capabilities in a custom class.
' Its purpose is to standardize how a logger (or logging handler) is injected into a class. This
' allows any class that implements the interface to receive a logger instance, which it can then
' use to log messages.
'
' Usage:
'
' ```
'   ' CustomClass.cls
'   Option Explicit
'   Implements LoggerAwareInterface
'
'   Private logger As VBAMonologger.LoggerInterface
'
'   Private Sub LoggerAwareInterface_setLogger(paramLogger As LoggerInterface)
'       Set logger = paramLogger
'   End Sub
'
'   Public Sub setLogger(paramLogger As LoggerInterface)
'       Call LoggerAwareInterface_setLogger(paramLogger)
'   End Sub
'
'   Public Sub foo()
'      logger.Info "The logger is available into your class module."
'   End Sub
' ```
'
' ```
'   ' Main.cls
'   Dim instance as New CustomClass
'
'   ' Dependency injection of logger (inject logger into class)
'   instance.setLogger VBAMonologger.Factory.createLoggerConsoleVBA("App)
'
'   ' In foo method we use the logger
'   instance.foo()
' ```
''

Option Explicit

'@Description("Set a logger instance in the object.")
Public Sub setLogger(ByRef Logger As VBAMonologger.LoggerInterface)
End Sub

