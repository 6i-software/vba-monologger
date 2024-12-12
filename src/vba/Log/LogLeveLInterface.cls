VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogLeveLInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("Availables log level.")
Public Enum LOG_LEVELS
    LEVEL_EMERGENCY = 600
    LEVEL_ALERT = 550
    LEVEL_CRITICAL = 500
    LEVEL_ERROR = 400
    LEVEL_WARNING = 300
    LEVEL_NOTICE = 250
    LEVEL_INFO = 200
    LEVEL_DEBUG = 100
End Enum
