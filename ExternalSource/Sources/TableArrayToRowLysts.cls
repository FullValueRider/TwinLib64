VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableArrayToRowLysts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Exposed

'@predeclaredId
'
'@Folder("VBALib")
Option Explicit

Implements ITableToLystAction

Private Function ITableToLystAction_execAction(ByVal ipTablearray As Variant) As Lyst
    Set ITableToLystAction_execAction = ExecAction(ipTablearray)
End Function

Public Function Deb() As TableArrayToRowLysts
    Set Deb = Me
End Function

Public Function ExecAction(ByVal ipTablearray As Variant) As Lyst
        
    If BailOut.When(Arrays.IsNotTableArray(ipTablearray), Msg.IsNotTableArray) Then Exit Function
        
    Dim myFirstRow As Long
    myFirstRow = LBound(ipTablearray, 1)
        
    Dim myLastRow As Long
    myLastRow = UBound(ipTablearray, 1)
        
    Dim myFirstColumn As Long
    myFirstColumn = LBound(ipTablearray, 2)
        
    Dim myLastColumn As Long
    myLastColumn = UBound(ipTablearray, 2)
        
    Dim myLystOfLyst As Lyst
    Set myLystOfLyst = Lyst.Deb
    Dim myRow As Long
    Dim myColumn As Long
    For myRow = myFirstRow To myLastRow
        
        Dim myRankLyst As Lyst
        Set myRankLyst = Lyst.Deb
        'myLystOfLyst.Add Lyst.Deb
            
        For myColumn = myFirstColumn To myLastColumn
                
            myRankLyst.AddKnownItemsAsItems ipTablearray(myRow, myColumn)
            'myLystOfLyst.Item(myRow - myFirstRow).Add ipTableArray(myRow, myColumn)
                
        Next
        
        myLystOfLyst.AddKnownItemsAsItems myRankLyst.Clone
            
    Next
    
    Globals.Res.Init (Msg.okay)
    Set ExecAction = myLystOfLyst
        
End Function

