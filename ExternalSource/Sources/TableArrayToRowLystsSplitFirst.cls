VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableArrayToRowLystsSplitFirst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("VBALib")
Option Explicit
Implements ITableToLystAction

Private Function ITableToLystAction_execAction(ByVal ipTablearray As Variant) As Lyst
    Set ITableToLystAction_execAction = ExecAction(ipTablearray)
End Function

Public Function Deb() As TableArrayToRowLystsSplitFirst
    Set Deb = Me
End Function

Public Function ExecAction(ByVal ipTablearray As Variant) As Lyst
        
    Dim myLystOfRankLysts As Lyst
    Set myLystOfRankLysts = TableArrayToRowLysts.ExecAction(ipTablearray)
        
    Dim myRankFirstItemsLyst As Lyst
    Set myRankFirstItemsLyst = Lyst.Deb
        
    Dim myRankLysts As Lyst
    Set myRankLysts = Lyst.Deb
        
    Dim myItem As Variant
    For Each myItem In myLystOfRankLysts
            
        Dim myInnerLyst As Lyst
        Set myInnerLyst = myItem
        myRankFirstItemsLyst.AddKnownItemsAsItems myInnerLyst.First
            
        Dim myExtent As Result
        If Types.TryExtent(myInnerLyst, myExtent).NotOkay Then Exit Function
                
                
        ' Dim myStartrun As Result
        ' Set myStartrun = Ranges.TryStartRunFromAnyStartRunEnd(myInnerLyst, myStartrun, 1)  'trystartrunfromanystartrunend(myinnerlyst)
        myRankLysts.AddKnownItemsAsItems myInnerLyst.GetRange(1, myInnerLyst.Count - 1)
            
    Next
        
    Set ExecAction = Lyst.Deb.AddKnownItemsAsItems(myRankFirstItemsLyst, myRankLysts)
        
End Function
