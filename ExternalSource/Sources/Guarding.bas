Attribute VB_Name = "Guarding"
Option Explicit
'@Folder("Guarding")

    
'    Public Const UnexpectedResult                   As Boolean = True
'    Public Const NoGuardTextParams                  As Variant = Empty

Public Sub Guard _
       ( _
       ByVal ipMessageId As Id, _
       ByVal ipThrow As Boolean, _
       ByVal ipLocation As String, _
       Optional ByVal ipArgs As Variant = Empty, _
       Optional ByVal ipAltMessage As String = vbNullString _
       )
            
    If ipThrow Then
            
        Dim myargs As Variant
        If Arrays.HasItems(ipArgs) Then
                
            myargs = ipArgs
                
        Else
                
            myargs = Array(ipArgs)
                
        End If
            
        Dim myMessage As String
        If VBA.Len(ipAltMessage) = 0 Then
            
            myMessage = Enums.Message.ToString(ipMessageId)
                
        Else
                
            myMessage = ipAltMessage
                
        End If
            
        If Arrays.HasItems(ipArgs) Then
                
            myMessage = Fmt.TxtArr(myMessage, myargs)
                
        End If

        VBA.Information.Err.Raise ipMessageId, ipLocation, myMessage
                        
    End If
        
End Sub

Public Sub GuardIf _
       ( _
       ByVal ipTest As Boolean, _
       ByVal ipMessageId As Id, _
       ByVal ipThrow As Boolean, _
       ByVal ipLocation As String, _
       Optional ByVal ipArgs As Variant = Empty, _
       Optional ByVal ipAltMessage As String = vbNullString _
       )
                
    If ipTest Then
                
        Guard ipMessageId, ipThrow, ipLocation, ipArgs, ipAltMessage
            
            
    End If
            
            
End Sub


