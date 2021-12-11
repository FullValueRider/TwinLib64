VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ResultItemsEnums"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "This module holds enumerations that give the indexing for the items lyst in different Result objects"
'@Exposed
'@PredeclaredId
Option Explicit
'@Folder("VBALib")
'@ModuleDescription("This module holds enumerations that give the indexing for the items lyst in different Result objects")

'We'd prefer to not use prefixes but unfortunately these are required in VBA
Public Enum ItemExtent

    ieFirstIndex = 0
    ieLbound = 0
    ieLastIndex = 1
    ieUbound = 1
    ieCount = 2
        
End Enum

Public Enum StartRun
        
    srStartIndex = 0
    srRunLength = 1
        
End Enum

Public Enum ParseParamArray
        
    ppCount = 0
    ppItems = 1
    ppKeys = 2
        
End Enum

' Funtions below allow intellisense for the enums above
Public Function ItemExtent(ByVal ipItemExtent As ItemExtent) As ItemExtent
    ItemExtent = ipItemExtent
End Function

Public Function StartRun(ByVal ipStartRun As StartRun) As StartRun
    StartRun = ipStartRun
End Function

Public Function ParseParamArray(ByVal ipParseParamArray As ParseParamArray) As ParseParamArray
    ParseParamArray = ipParseParamArray
End Function
