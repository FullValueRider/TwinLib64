VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BailOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@PredeclaredId
'@Exposed
'@Folder("VBALib.Filters")


' Guard statements
' Guard statements imply allowing execution to continue if the test is true
' This is incompatible with a global result value because we wish to say
' the result is NOTOKAY when the test is true
' for this reason the exit when class decouples setting the result
' from the return value and so we can just return a simple
' boolean
'@Ignore EmptyMethod
Private Sub Class_Initialize()
End Sub

'@Description("Sets message and global.res.status to not okay. Returns ipBailout ")
Public Function When _
( _
    ByVal ipBailOut As Boolean, _
    Optional ByVal ipMsg As Msg = Msg.Unspecified, _
    Optional ByRef ipMsgArgsLyst As Lyst, _
    Optional ByVal ipAltMessage As String _
) As Boolean
    
    If ipBailOut Then
        
        SetGlobalResMsg ipMsg, ipMsgArgsLyst, ipAltMessage
       
        Globals.Res.SetNotOkay
        
    End If
 
    When = ipBailOut
    
End Function

'@Description("Sets message and global.res.status to okay. Returns ipNotify.")
Public Function Notify _
( _
    ByVal ipNotify As Boolean, _
    Optional ByVal ipMsg As Msg = Msg.Unspecified, _
    Optional ByRef ipMsgArgsLyst As Lyst, _
    Optional ByVal ipAltMessage As String _
) As Boolean
        
    If ipNotify Then
        
        SetGlobalResMsg ipMsg, ipMsgArgsLyst, ipAltMessage
        Globals.Res.SetOkay
        
    End If
    
    Notify = ipNotify
    
End Function


Friend Sub SetGlobalResMsg _
( _
    ByVal ipMessageEnum As Msg, _
    Optional ByVal ipArgLyst As Lyst, _
    Optional ByVal ipAltMessage As String _
)
    
    Dim myMessage As String
    If VBA.Len(ipAltMessage) = 0 Then
        
        myMessage = enums.Message.ToString(ipMessageEnum)
            
        
    Else
            
        myMessage = ipAltMessage
            
        
    End If
    
    Dim myargs As Variant
    If ipArgLyst Is Nothing Then
        
        myargs = Array()
        
    Else
        
        myargs = ipArgLyst.ToArray
        
    End If
    
    Dim myErrorText As String
    myErrorText = Fmt.TxtArr(myMessage, myargs)
    Globals.Res.SetText myErrorText
    
End Sub

