Attribute VB_Name = "ANSI"
' ------------------------------------- '
'                                       '
'    VBA Monologger                     '
'    Copyright © 2024, 6i software      '
'                                       '
' ------------------------------------- '
'
'@Exposed
'@Folder("VBAMonologger.Utils")
'@FQCN("VBAMonologger.Utils.ANSI")
'@ModuleDescription("Utils to get ANSI escape sequence.")
'
' Using by the formatter `VBAMonologger.Formatter.FormatterANSIColoredLine`
''

Option Explicit

Public Function GetANSIEscapeSequence(ByVal styleName As String) As String
    Dim ESC As String
    ESC = Chr$(27)

    Select Case UCase(styleName)
        ' Reset
        Case "RESET": GetANSIEscapeSequence = ESC & "[0m"
    
        ' Styles
        Case "BOLD": GetANSIEscapeSequence = ESC & "[1m"
        Case "WEAK": GetANSIEscapeSequence = ESC & "[2m"
        Case "UNDERLINE": GetANSIEscapeSequence = ESC & "[4m"
        Case "BLINK": GetANSIEscapeSequence = ESC & "[5m"
        Case "REVERSE": GetANSIEscapeSequence = ESC & "[7m"
        Case "HIDDEN": GetANSIEscapeSequence = ESC & "[8m"
        
        ' Foreground normal
        Case "BLACK": GetANSIEscapeSequence = ESC & "[30m"
        Case "RED": GetANSIEscapeSequence = ESC & "[31m"
        Case "GREEN": GetANSIEscapeSequence = ESC & "[32m"
        Case "YELLOW": GetANSIEscapeSequence = ESC & "[33m"
        Case "BLUE": GetANSIEscapeSequence = ESC & "[34m"
        Case "MAGENTA": GetANSIEscapeSequence = ESC & "[35m"
        Case "CYAN": GetANSIEscapeSequence = ESC & "[36m"
        Case "WHITE": GetANSIEscapeSequence = ESC & "[37m"
        
        ' Foreground High-Intensity
        Case "BRIGHT_BLACK": GetANSIEscapeSequence = ESC & "[90m"
        Case "BRIGHT_RED": GetANSIEscapeSequence = ESC & "[91m"
        Case "BRIGHT_GREEN": GetANSIEscapeSequence = ESC & "[92m"
        Case "BRIGHT_YELLOW": GetANSIEscapeSequence = ESC & "[93m"
        Case "BRIGHT_BLUE": GetANSIEscapeSequence = ESC & "[94m"
        Case "BRIGHT_MAGENTA": GetANSIEscapeSequence = ESC & "[95m"
        Case "BRIGHT_CYAN": GetANSIEscapeSequence = ESC & "[96m"
        Case "BRIGHT_WHITE": GetANSIEscapeSequence = ESC & "[97m"
        
        ' Background normal
        Case "BG_BLACK": GetANSIEscapeSequence = ESC & "[40m"
        Case "BG_RED": GetANSIEscapeSequence = ESC & "[41m"
        Case "BG_GREEN": GetANSIEscapeSequence = ESC & "[42m"
        Case "BG_YELLOW": GetANSIEscapeSequence = ESC & "[43m"
        Case "BG_BLUE": GetANSIEscapeSequence = ESC & "[44m"
        Case "BG_MAGENTA": GetANSIEscapeSequence = ESC & "[45m"
        Case "BG_CYAN": GetANSIEscapeSequence = ESC & "[46m"
        Case "BG_WHITE": GetANSIEscapeSequence = ESC & "[47m"
        
        ' Background High-Intensity
        Case "BG_BRIGHT_BLACK": GetANSIEscapeSequence = ESC & "[100m"
        Case "BG_BRIGHT_RED": GetANSIEscapeSequence = ESC & "[101m"
        Case "BG_BRIGHT_GREEN": GetANSIEscapeSequence = ESC & "[102m"
        Case "BG_BRIGHT_YELLOW": GetANSIEscapeSequence = ESC & "[103m"
        Case "BG_BRIGHT_BLUE": GetANSIEscapeSequence = ESC & "[104m"
        Case "BG_BRIGHT_MAGENTA": GetANSIEscapeSequence = ESC & "[105m"
        Case "BG_BRIGHT_CYAN": GetANSIEscapeSequence = ESC & "[106m"
        Case "BG_BRIGHT_WHITE": GetANSIEscapeSequence = ESC & "[107m"

        ' Error
        Case Else:
            Err.Raise vbObjectError + 1000, "VBAMonologger.ANSI", "Unknown ANSI sequence: " & styleName
    End Select
End Function

' Reset
Public Function RESET() As String: RESET = GetANSIEscapeSequence("RESET"): End Function

' Styles
Public Function BOLD() As String: BOLD = GetANSIEscapeSequence("BOLD"): End Function
Public Function WEAK() As String: WEAK = GetANSIEscapeSequence("WEAK"): End Function
Public Function UNDERLINE() As String: UNDERLINE = GetANSIEscapeSequence("UNDERLINE"): End Function
Public Function BLINK() As String: BLINK = GetANSIEscapeSequence("BLINK"): End Function
Public Function REVERSE() As String: REVERSE = GetANSIEscapeSequence("REVERSE"): End Function
Public Function HIDDEN() As String: HIDDEN = GetANSIEscapeSequence("HIDDEN"): End Function

' Foreground normal
Public Function BLACK() As String: BLACK = GetANSIEscapeSequence("BLACK"): End Function
Public Function RED() As String: RED = GetANSIEscapeSequence("RED"): End Function
Public Function GREEN() As String: GREEN = GetANSIEscapeSequence("GREEN"): End Function
Public Function YELLOW() As String: YELLOW = GetANSIEscapeSequence("YELLOW"): End Function
Public Function BLUE() As String: BLUE = GetANSIEscapeSequence("BLUE"): End Function
Public Function MAGENTA() As String: MAGENTA = GetANSIEscapeSequence("MAGENTA"): End Function
Public Function CYAN() As String: CYAN = GetANSIEscapeSequence("CYAN"): End Function
Public Function WHITE() As String: WHITE = GetANSIEscapeSequence("WHITE"): End Function

' Foreground high-Intensity
Public Function BRIGHT_BLACK() As String: BRIGHT_BLACK = GetANSIEscapeSequence("BRIGHT_BLACK"): End Function
Public Function BRIGHT_RED() As String: BRIGHT_RED = GetANSIEscapeSequence("BRIGHT_RED"): End Function
Public Function BRIGHT_GREEN() As String: BRIGHT_GREEN = GetANSIEscapeSequence("BRIGHT_GREEN"): End Function
Public Function BRIGHT_YELLOW() As String: BRIGHT_YELLOW = GetANSIEscapeSequence("BRIGHT_YELLOW"): End Function
Public Function BRIGHT_BLUE() As String: BRIGHT_BLUE = GetANSIEscapeSequence("BRIGHT_BLUE"): End Function
Public Function BRIGHT_MAGENTA() As String: BRIGHT_MAGENTA = GetANSIEscapeSequence("BRIGHT_MAGENTA"): End Function
Public Function BRIGHT_CYAN() As String: BRIGHT_CYAN = GetANSIEscapeSequence("BRIGHT_CYAN"): End Function
Public Function BRIGHT_WHITE() As String: BRIGHT_WHITE = GetANSIEscapeSequence("BRIGHT_WHITE"): End Function

' Background normal
Public Function BG_BLACK() As String: BG_BLACK = GetANSIEscapeSequence("BG_BLACK"): End Function
Public Function BG_RED() As String: BG_RED = GetANSIEscapeSequence("BG_RED"): End Function
Public Function BG_GREEN() As String: BG_GREEN = GetANSIEscapeSequence("BG_GREEN"): End Function
Public Function BG_YELLOW() As String: BG_YELLOW = GetANSIEscapeSequence("BG_YELLOW"): End Function
Public Function BG_BLUE() As String: BG_BLUE = GetANSIEscapeSequence("BG_BLUE"): End Function
Public Function BG_MAGENTA() As String: BG_MAGENTA = GetANSIEscapeSequence("BG_MAGENTA"): End Function
Public Function BG_CYAN() As String: BG_CYAN = GetANSIEscapeSequence("BG_CYAN"): End Function
Public Function BG_WHITE() As String: BG_WHITE = GetANSIEscapeSequence("BG_WHITE"): End Function

' Background high-Intensity
Public Function BG_BRIGHT_BLACK() As String: BG_BRIGHT_BLACK = GetANSIEscapeSequence("BG_BRIGHT_BLACK"): End Function
Public Function BG_BRIGHT_RED() As String: BG_BRIGHT_RED = GetANSIEscapeSequence("BG_BRIGHT_RED"): End Function
Public Function BG_BRIGHT_GREEN() As String: BG_BRIGHT_GREEN = GetANSIEscapeSequence("BG_BRIGHT_GREEN"): End Function
Public Function BG_BRIGHT_YELLOW() As String: BG_BRIGHT_YELLOW = GetANSIEscapeSequence("BG_BRIGHT_YELLOW"): End Function
Public Function BG_BRIGHT_BLUE() As String: BG_BRIGHT_BLUE = GetANSIEscapeSequence("BG_BRIGHT_BLUE"): End Function
Public Function BG_BRIGHT_MAGENTA() As String: BG_BRIGHT_MAGENTA = GetANSIEscapeSequence("BG_BRIGHT_MAGENTA"): End Function
Public Function BG_BRIGHT_CYAN() As String: BG_BRIGHT_CYAN = GetANSIEscapeSequence("BG_BRIGHT_CYAN"): End Function
Public Function BG_BRIGHT_WHITE() As String: BG_BRIGHT_WHITE = GetANSIEscapeSequence("BG_BRIGHT_WHITE"): End Function

