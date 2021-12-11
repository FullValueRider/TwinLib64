VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mpToLong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@PredeclaredId
'@Exposed
'@Folder("VBALib.Mappers")
    
Implements IMapper

'TwinBasic requirement
Private Sub Class_Initialize()
End Sub

Public Function Deb() As mpToLong
        
    With New mpToLong
        
        Set Deb = .ReadyToUseInstance
            
    End With
        
End Function

Friend Function ReadyToUseInstance() As mpToLong
        
    Set ReadyToUseInstance = Me
        
End Function

Private Function IMapper_ExecMapper(ByVal ipHostItem As Variant) As Variant
    IMapper_ExecMapper = ExecMapper(ipHostItem)
End Function

'@Description("Converts tthe items in the Lyst to Long Integers.  If the input item cannot be converted the string "NaN" is returned")
Public Function ExecMapper(ByVal ipHostItem As Variant) As Variant
        
    On Error Resume Next
    Dim myResult As Long
    myResult = CLng(ipHostItem)
    Dim myErrorNo As Long
    myErrorNo = Err.Number
    On Error GoTo 0
        
    ExecMapper = IIf(myErrorNo = 0, myResult, Sys.NotANumber)
        
End Function
