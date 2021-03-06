VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ftIfIt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@PredeclaredId
'@Exposed
'@Folder("VBALib.Filters")
Implements IFilter
    
Private Type State
        
    Comparer                        As IComparer
    
        
End Type

Private s                           As State

Public Function Deb(ByRef ipComparer As IComparer) As ftIfIt
        
    With New ftIfIt
        
        Set Deb = .ReadyToUseInstance(ipComparer)
            
    End With
        
End Function

Friend Function ReadyToUseInstance(ByRef ipComparer As IComparer) As ftIfIt
        
    Set s.Comparer = ipComparer
    Set ReadyToUseInstance = Me
        
End Function

'@Description("Adds the value in ipArgsArray(0) to each item in the Lyst.  If the host item is not a number the string NAN is returned")
Private Function IFilter_ExecFilter(ByRef ipHostItems As Lyst) As Lyst
Attribute IFilter_ExecFilter.VB_Description = "Adds the value in ipArgsArray(0) to each item in the Lyst.  If the host item is not a number the string NAN is returned"

    Set IFilter_ExecFilter = ExecRed(ipHostItems)
End Function

Public Function ExecRed(ByVal ipHostItems As Lyst) As Lyst
        
    Dim myLyst As Lyst
    Set myLyst = Lyst.Deb
        
   Dim myItem As Variant
   For Each myItem In ipHostItems
   
        If s.Comparer.ExecCmp(myItem) Then myLyst.Add myItem
   
   Next
   
   Set ExecRed = myLyst
        
End Function
