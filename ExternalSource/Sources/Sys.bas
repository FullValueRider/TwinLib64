Attribute VB_Name = "Sys"
Attribute VB_Description = "A place for useful VBA things not explicitly provided by VBA"
   

Option Explicit
'@ModuleDescription("A place for useful VBA things not explicitly provided by VBA")
'@Folder("VBALib")

'@Ignore ConstantNotUsed
Public Const MaxLong                        As Long = &H7FFFFFFF
Public Const MinLong                        As Long = &HFFFFFFFF

Public Const NotOkay                        As Boolean = False
Public Const Okay                           As Boolean = True

Public Const NotANumber                     As String = "NaN"

Public Function AsOneItem(ByVal ipIterable As Variant) As Variant
    AsOneItem = Array(ipIterable)
End Function

