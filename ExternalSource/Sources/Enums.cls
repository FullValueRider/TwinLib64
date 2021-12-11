VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Enums"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Provides a central location for providing complex enumerations"
Option Explicit
'@ModuleDescription("Provides a central location for providing complex enumerations")
'@Folder("Enums")
'@PredeclaredId
'@exposed

' Complex enumerations are enumerations embedded in a class
Public Property Get Message() As enumMessages
    Set Message = enumMessages
End Property

Public Property Get TableToLystActions() As EnumTableToLystActions
    Set TableToLystActions = EnumTableToLystActions
End Property
