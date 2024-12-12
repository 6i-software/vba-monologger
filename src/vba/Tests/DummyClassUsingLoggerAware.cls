VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DummyClassUsingLoggerAware"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements VBAMonologger.LoggerAwareInterface

Private Logger As VBAMonologger.LoggerInterface

' _____________________________________ '
'                                       '
'  Implementation: LoggerAwareInterface '
' _____________________________________ '
'                                       '
'@inheritdoc
Private Sub LoggerAwareInterface_setLogger(paramLogger As LoggerInterface)
    Set Logger = paramLogger
End Sub

'@Description("Proxy method for public exposition.")
Public Sub setLogger(paramLogger As LoggerInterface)
    Call LoggerAwareInterface_setLogger(paramLogger)
End Sub


' ________________ '
'                  '
'  Public methods  '
' ________________ '
'                  '
Public Sub foo()
    ' Minimal exemple of using the "injected" logger
    Logger.info "I believe I can fly!"
    
    ' Example of using the logger with a specified context and pre-processors placeholders (if enabled)
    Dim context As New Scripting.Dictionary
    context.Add "User", "Bob"
    context.Add "Operation", "fly"
    Logger.notice "I believe {User} can {Operation} in the sky", context
End Sub
