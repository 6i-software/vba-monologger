VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringableInterface"
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
'@Folder("VBAMonologger.Utils")
'@FQCN("VBAMonologger.Utils.StringableInterface")
'@ModuleDescription("Denotes the class as having a toString() method.")
'@Interface
''

'@Description("Give a string representation of instance.")
Public Function toString() As String
End Function
