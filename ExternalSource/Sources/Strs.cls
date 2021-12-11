VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Strs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@PredeclaredId
'@Exposed
'@Folder("VBALib")
        
Private Type Properties
        
    Separator               As String
        
End Type

Private p                   As Properties

Private Sub Class_Initialize()
    p.Separator = Char.twComma
End Sub

Public Function WhiteSpace() As Variant: WhiteSpace = Array(Char.twSpace, Char.twPeriod, Char.twSemiColon, Char.twColon, Char.twComma, vbTab, vbCrLf): End Function
Public Function NumberChars() As String: NumberChars = "0123456789": End Function

Public Function LCaseChars() As String: LCaseChars = "abcdefghijklmnopqrstuvwxyz": End Function
Public Function UCaseChars() As String: UCaseChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ": End Function

Friend Function TryExtent(ByVal ipItem As String, ByRef iopResult As Result) As Result
        
    If iopResult Is Nothing Then Set iopResult = Result.Deb
    

    Dim myCount As Long
    myCount = VBA.Len(ipItem)
    If myCount = 0 Then
            
        iopResult.MessageEnum = Enums.Message.AsEnum(Id.StringIsEmpty)
        Set TryExtent = iopResult
        Exit Function
            
            
    End If
        
    iopResult.SetOkay 1, myCount, myCount
    Set TryExtent = iopResult
     
End Function

'@Description("Returny the long value of a string containing a binary number")
Public Function BinToLong(ByVal ipBinary As String) As Long
Attribute BinToLong.VB_Description = "Returny the long value of a string containing a binary number"

        
    Dim myValue As Long
    myValue = 0
    Dim myIndex As Long
    For myIndex = Len(ipBinary) To 1 Step -1
        
        Dim myChar As String
        myChar = VBA.Mid$(ipBinary, myIndex, 1)
        
        If myChar = "1" Then
            
            myValue = myValue + 2 ^ (Len(ipBinary) - myIndex)
                
        End If
            
    Next
        
    BinToLong = myValue
        
End Function

' Public Function dedup(ByVal ipsource As String, ByRef ipDedupChars As Scripting.Dictionary) As String
        
' End Function
' Public Function dedup(ByVal ipsource As String, ByRef ipDedupChars As ArrayList) As String
        
' End Function
    
' Public Function Dedup(ByVal ipsource As String, ByRef ipDedupArray As Collection) As String
    
        
' 'End Function
    
Public Property Get DefaultSeparator() As String
    DefaultSeparator = p.Separator
End Property

Public Property Let DefaultSeparator(ByVal ipSeparator As String)
    p.Separator = ipSeparator
End Property

Public Function Dedup(ByVal ipSource As String, ParamArray ipDedupArray() As Variant) As String

    Dedup = ipSource
    If VBA.Len(ipSource) = 0 Then Exit Function
    If VBA.IsMissing(ipDedupArray) Then ipDedupArray = Array(Char.twSpace)
    If Arrays.HasNoItems(ipDedupArray) Then ipDedupArray = Array(Char.twSpace)
        
    ' DO some dodgy stuff if only one parameter is provided
        
    Dim myResult As Result
    Dim myDedupItems As Variant
        
    If Arrays.HasOneItem(ipDedupArray) Then
        
        If ParamArrayParsing.OneItem(CVar(ipDedupArray), myResult).NotOkay Then Exit Function
        If Types.IsNotString(myResult.Item(1).Item(0)) Then Exit Function
        
        myDedupItems = myResult.Item(1).ToArray
           
                
        If VBA.Len(myDedupItems(0)) = 1 Then
                    
            myDedupItems = Array(myDedupItems(0))
                        
        ElseIf VBA.InStr(myDedupItems(0), p.Separator) > 0 Then
                        
            myDedupItems = VBA.Split(myDedupItems(0), p.Separator)
                    
        End If
            
    Else
           
        myDedupItems = ipDedupArray
                          
            
    End If
        
        
    Dim mySource As String
    mySource = ipSource
       
    Dim myDeDup As Variant
    For Each myDeDup In myDedupItems
        
        Dim MyDedupDedup As String
        MyDedupDedup = myDeDup & myDeDup
            
        Do
            
            Dim myLen As Long
            myLen = Len(mySource)
            mySource = Replace(mySource, MyDedupDedup, myDeDup)
                
        Loop Until myLen = Len(mySource)
            
    Next
        
    Dedup = mySource
        
End Function

'@Description("Takes a list of characters and removed them from both ends of the target string.  If the list contains strings >1 character then these are ignored")
Public Function Trimmer(ByVal ipSource As String, ParamArray ipTrimChars() As Variant) As String
Attribute Trimmer.VB_Description = "Takes a list of characters and removed them from both ends of the target string.  If the list contains strings >1 character then these are ignored"

    Trimmer = ipSource
    If VBA.Len(ipSource) = 0 Then Exit Function
    If VBA.IsMissing(ipTrimChars) Then ipTrimChars = Strs.WhiteSpace
    If Arrays.HasNoItems(ipTrimChars) Then ipTrimChars = Strs.WhiteSpace
        
    Dim myTrimArray As Variant
    If Arrays.HasOneItem(ipTrimChars) Then
        
        Dim myResult As Result
        If ParamArrayParsing.OneItem(ipTrimChars(0), myResult).NotOkay Then Exit Function
        myTrimArray = myResult.Item(0)
            
    Else
            
        myTrimArray = ipTrimChars
            
    End If
        
    Dim myTrimLyst As Lyst
    Set myTrimLyst = Lyst.Deb.AddRange(myTrimArray)
    Dim myString As String
    myString = ipSource
      
    Dim myLen As Long
    myLen = VBA.Len(myString)
    Do
  
        myLen = VBA.Len(myString)
        DoEvents                                 ' Always put a do event statement in a do loop
        If myTrimLyst.HoldsItem(VBA.Left$(myString, 1)) Then myString = VBA.Mid$(myString, 2)
       
                
    Loop Until myLen = VBA.Len(myString)
         
    Do
        
        myLen = VBA.Len(myString)
        DoEvents                                 ' Always put a do event statement in a do loop
        If myTrimLyst.HoldsItem(VBA.Right$(myString, 1)) Then myString = VBA.Mid$(myString, 1, myLen - 1)

            
    Loop Until myLen = VBA.Len(myString)
            
    Trimmer = myString
        
End Function

Public Function PadLeft(ByVal ipString As String, ByVal ipWidth As Long) As String

    PadLeft = ipString
    If Len(ipString) >= ipWidth Then Exit Function
        
    Dim myReturn As String
    myReturn = VBA.String$(ipWidth, Char.twSpace)
    LSet myReturn = ipString
    PadLeft = myReturn
        
End Function

Public Function PadRight(ByVal ipString As String, ByVal ipWidth As Long, Optional ByVal ipChar As String = " ") As String
    PadRight = ipString
    If Len(ipString) >= ipWidth Then Exit Function
        
    Dim myReturn As String
    myReturn = String$(ipWidth, ipChar)
    RSet myReturn = ipString
    PadRight = myReturn
        
End Function

Public Function Count(ByVal ipString As String, ByVal ipChar As String) As Long
    Count = Len(ipString) - Len(Replace(ipString, ipChar, vbNullString))
End Function

'Public Function NullStrArr(ByVal ipCount As Long) As Variant
'
'    'Dim myKvp As Kvp
'    ' vbNullString gives an EMpty variant, "" gives ""
'    '@Ignore EmptyStringLiteral
'    NullStrArr = Kvp.Deb.Add("", ipCount).GetValues
'
'End Function


'@Description("Takes string in the form of X,Y and returns array containing X Long, Y Long")
Public Function CoordsToXY(ByVal ipCoord As String) As Variant
Attribute CoordsToXY.VB_Description = "Takes string in the form of X,Y and returns array containing X Long, Y Long"
    CoordsToXY = Array(CLng(Split(ipCoord, ",")(0)), CLng(Split(ipCoord, ",")(1)))
End Function

'@Description("Converts a string to an Lyst of trimmed substrings")
Public Function ToSubStrLyst(ByVal ipString As String, Optional ByVal ipSeparator_twComma As Variant, Optional ByVal ipDeleteChars_WhiteSpace As Variant) As Lyst
Attribute ToSubStrLyst.VB_Description = "Converts a string to an Lyst of trimmed substrings"

    Dim myArray As Variant
    Dim myString As String
    ipSeparator_twComma = IIf(VBA.IsMissing(ipSeparator_twComma), Char.twComma, ipSeparator_twComma)
    ipDeleteChars_WhiteSpace = IIf(VBA.IsMissing(ipDeleteChars_WhiteSpace), Strs.WhiteSpace, ipDeleteChars_WhiteSpace)
        
    myString = Strs.Trimmer(ipString, ipDeleteChars_WhiteSpace)
    If InStr(myString, ipSeparator_twComma) = 0 Then
            
        myArray = Array(myString)
            
    Else
            
        myArray = VBA.Split(myString, ipSeparator_twComma)
            
    End If
        
    Dim myItem As Variant
    Dim myLyst As Lyst
    Set myLyst = Lyst.Deb
    For Each myItem In myArray
        
        myLyst.Add myItem
            
    Next

    Set ToSubStrLyst = myLyst
        
End Function

Public Function Replacer(ByVal ipString As String, Optional ByVal ipReplaceChars_whiteSpace As Variant) As String
        
    ipReplaceChars_whiteSpace = IIf(VBA.IsMissing(ipReplaceChars_whiteSpace), Strs.WhiteSpace, ipReplaceChars_whiteSpace)
    Dim myString As String
    myString = ipString
        
    If VBA.Len(ipReplaceChars_whiteSpace) = 0 Then Exit Function
        
    'Dim myResult As String
    Dim myIndex As Long
    For myIndex = 1 To Len(ipReplaceChars_whiteSpace)
            
        
        myString = VBA.Replace(myString, VBA.Mid$(ipReplaceChars_whiteSpace, myIndex, 1), vbNullString)
            
    Next

    Replacer = myString
        
End Function

Public Function ToAnsiBytes(ByVal ipString As String) As Lyst

    If VBA.Len(ipString) = 0 Then
        
        Set ToAnsiBytes = Lyst.Deb
            
    Else
        
        Set ToAnsiBytes = Lyst.Deb(VBA.Split(StrConv(ipString, vbFromUnicode)))
            
    End If
        
End Function

Public Function ToUnicodeBytes(ByVal ipString As String) As Lyst

    If VBA.Len(ipString) = 0 Then
        
        Set ToUnicodeBytes = Lyst.Deb
            
    Else
        
        Dim myBytes() As Byte
        myBytes = ipString
        Set ToUnicodeBytes = Lyst.Deb.AddRange(myBytes)
            
    End If
        
End Function

Public Function ToCharLyst(ByVal ipString As String) As Lyst

    Dim myLyst As Lyst
    Set myLyst = Lyst.Deb
    Set ToCharLyst = myLyst
        
    Dim myLen As Long
    myLen = VBA.Len(ipString)
    If myLen = 0 Then Exit Function
        
    Dim myIndex As Long
    For myIndex = 1 To myLen
        
        myLyst.Add VBA.Mid$(ipString, myIndex, 1)
            
    Next
        
    Set ToCharLyst = myLyst

End Function
