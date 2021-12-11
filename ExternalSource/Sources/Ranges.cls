VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ranges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@PredeclaredId
'@Exposed
'@Folder("VBALib")

Private Type ReadOnly                            ' Pseudo constants

    Location                        As String
        

End Type

'@Ignore MoveFieldCloserToUsage
Private r                           As ReadOnly

Private Sub Class_Initialize()
    r.Location = "VBALib.Ranges."
End Sub

Public Function GetSeries(ByVal ipStart As Variant, ByVal ipCount As Long, Optional ByVal ipStep As Variant = 1) As Lyst

    Dim myLyst As Lyst
    Set myLyst = Lyst.Deb
    Set GetSeries = myLyst
        
    If Types.Group.IsNotNumber(ipStart, ipCount, ipStep) Then Exit Function
        
    Dim myItem As Variant
    For myItem = 1 To ipCount
        
        myLyst.Host.add ipStart
        ipStart = ipStart + ipStep
            
    Next
        
End Function

Public Function GetConstSeries(ByVal ipCount As Long, ByVal ipValue As Variant) As Lyst

    Dim myLyst As Lyst
    Set myLyst = Lyst.Deb
    Set GetConstSeries = myLyst
        
    Dim myItem As Long
    For myItem = 1 To ipCount
        
        myLyst.Host.add ipValue
            
    Next
        
End Function

'@Description("Takes an iterable or string and returns a Result (see Result.fields.StartRun). The Input StartIndex and EndIndex  use a 1 based index and may be negative.")
Friend Function TryStartRunFromAnyStartRunEnd _
       ( _
       ByVal ipItem As Variant, _
       Optional ByRef iopResult As Result, _
       Optional ByVal ipStartIndex As Variant, _
       Optional ByVal ipRun As Variant, _
       Optional ByVal ipEndIndex As Variant, _
       Optional ByVal ipRank As Long = 1 _
       ) As Result
Attribute TryStartRunFromAnyStartRunEnd.VB_Description = "Takes an iterable or string and returns a Result (see Result.fields.StartRun). The Input StartIndex and EndIndex  use a 1 based index and may be negative."
    ' Item must be an iterable or a string
    ' If ipRun is provided, any value for ipEndIndex is ignored
    ' to the start (positive numbers) or end(negativenumbers)
    ' The method attempts to return valid start and run values so may
    ' coerce input values to do so
    ' Where coercion is used a message retuned about the coercion applied
    ' Only the most recent coercion message is returned.
    '
    ' To emulate other programming languages, values
    ' for the start, end or run can be negative to indicate counting down from a high to low index.
    ' e.g. in terms of this library an ipstartindex of -1 would correspond to the LastIndex property.
    ' This is the reason why the input startindex and end index must be relative values
    ' The start and end index values are the ordinal position starting the first or last index
    ' Thus positive indeces start at 1 and proceeed to a value of '.count' and for negative indeces indeces start at -1
    ' and proceeeds to a value of '-.count'
    ' The output startindex is givem in terms of the indexing of the input iterable
        
        
    If iopResult Is Nothing Then
            
        Set iopResult = Result.Deb
            
            
    Else
            
        iopResult.SetNotOkay
            
            
    End If
        
    Set TryStartRunFromAnyStartRunEnd = iopResult
        
    If Types.Group.IsNotIterable(ipItem) Then
            
        ' we want to allow strings
        If Types.IsNotString(ipItem) Then
                
            iopResult.MessageEnum = Enums.Message.AsEnum(Id.IsNotIterable)
            Exit Function
            
                
        End If
            
    End If
        
    If Types.IsNotArray(ipItem) Then
        
        If ipRank > 1 Then
            
            iopResult.MessageEnum = Enums.Message.AsEnum(Id.ObjectDoesNotSupportRanks)
            Exit Function
                
                
        End If
        
    End If
        
    If Types.TryExtent(ipItem, iopResult, ipRank).NotOkay Then Exit Function
        
    ' Preserve the first index for later use
    Dim myFirstIndex As Long
    myFirstIndex = iopResult.Item(ResultItemsEnums.ItemExtent(ieFirstIndex))
        
    ' Guard against invalid conditions
    ' Guard against a run of zero
        
    If Types.IsNotMissing(ipRun) Then
            
        If ipRun = 0 Then
                
            iopResult.MessageEnum = Enums.Message.AsEnum(Id.RunLengthIsZero)
            Exit Function
                
                
        End If
            
    End If
        
    ' Ensure a legal start index
    If Types.IsNotMissing(ipStartIndex) Then
            
        If VBA.Abs(ipStartIndex) > iopResult.Item(ResultItemsEnums.ItemExtent(ieCount)) Then
                
            iopResult.SetNotOkay
            iopResult.MessageEnum = Enums.Message.AsEnum(Id.StartIndexExceedsItemCount)
            Set TryStartRunFromAnyStartRunEnd = iopResult
            Exit Function
                
            
        ElseIf ipStartIndex = 0 Then
                
            ipStartIndex = 1
            iopResult.MessageEnum = Enums.Message.AsEnum(Id.StartIndexWasZeroResetToOne)
                
            
        End If
            
    End If
        
    ' Ensure a legal end index
    If Types.IsNotMissing(ipEndIndex) Then
            
        If VBA.Abs(ipEndIndex) > iopResult.Item(ResultItemsEnums.ItemExtent(ieCount)) Then
                
            ipEndIndex = iopResult.Item(ResultItemsEnums.ItemExtent(ieCount)) * VBA.Sgn(ipEndIndex)
            iopResult.MessageEnum = Enums.Message.AsEnum(Id.EndIndexExceedsItemCount)
                
                
        ElseIf ipEndIndex = 0 Then
                
            ipEndIndex = 1
            iopResult.MessageEnum = Enums.Message.AsEnum(Id.EndIndexWasZeroResetToOne)
            

        End If
                
    End If
    
    ' There are 8 combinations of three variables (ipStartIndex, ipRun, and ipEndIndex)
    ' however when ipRun is provided ipEndIndex can be ignored
    ' So we can simplify to 6 cases
    ' iopResult is being used to pass in the results of Tyes.Extent
    Select Case True
            
    Case Types.IsNotMissing(ipStartIndex, ipRun)
            
        Set iopResult.Items = GetStartRunFromAnyStartRun(ipStartIndex, ipRun, iopResult)
                
            
    Case Types.IsNotMissing(ipStartIndex, ipEndIndex)
            
        Set iopResult.Items = GetStartRunFromAnyStartEnd(ipStartIndex, ipEndIndex, iopResult)
                
            
    Case Types.IsNotMissing(ipStartIndex)
            
        Set iopResult.Items = GetStartRunFromAnyStartOnly(ipStartIndex, iopResult)
                
            
    Case Types.IsNotMissing(ipRun)
            
        Set iopResult.Items = GetStartRunFromAnyRunOnly(ipRun, iopResult)
                
            
    Case Types.IsNotMissing(ipEndIndex)
            
        Set iopResult.Items = GetStartRunFromAnyEndOnly(ipEndIndex, iopResult)
                
            
    Case Types.IsMissing(ipStartIndex, ipRun, ipEndIndex)
            
        Set iopResult.Items = GetStartRunFromNoStartRunOrEnd(iopResult)
                
                
    End Select
        
    iopResult.Item(ResultItemsEnums.StartRun(srStartIndex)) = iopResult.Item(ResultItemsEnums.StartRun(srStartIndex)) + myFirstIndex - 1
    'iopresult.setokay    
End Function

Friend Function GetStartRunFromAnyStartRun _
       ( _
       ByVal ipStartIndex As Long, _
       ByVal ipRun As Long, _
       Optional ByRef iopResult As Result _
       ) As ArrayList
        
    Dim myStartIndex As Long
    Dim myRun As Long
    If ipStartIndex > 0 And ipRun > 0 Then
            
        If ipStartIndex + ipRun - 1 <= iopResult.Item(ResultItemsEnums.ItemExtent(ieCount)) Then
                
            myRun = ipRun
            myStartIndex = ipStartIndex
                
                
        Else
                
            myRun = iopResult.Item(ResultItemsEnums.ItemExtent(ieCount)) - ipStartIndex + 1
            myStartIndex = ipStartIndex
            iopResult.MessageEnum = Enums.Message.AsEnum(Id.RunLengthTooLong)
                
                
        End If
            
    ElseIf ipStartIndex > 0 And ipRun < 0 Then
            
        If ipStartIndex + ipRun + 1 > 0 Then
                
            myStartIndex = ipStartIndex + ipRun + 1
            myRun = VBA.Abs(ipRun)
                
        Else
                
            myStartIndex = 1
            myRun = ipStartIndex
            iopResult.MessageEnum = Enums.Message.AsEnum(Id.RunLengthTooLong)
                
        End If
            
            
    ElseIf ipStartIndex < 0 And ipRun > 0 Then
            
        myStartIndex = iopResult.Item(ResultItemsEnums.ItemExtent(ieCount)) + ipStartIndex + 1
        If myStartIndex + ipRun - 1 <= iopResult.Item(ResultItemsEnums.ItemExtent(ieCount)) Then
                
            myRun = ipRun
                
        Else
                
            myRun = iopResult.Item(ResultItemsEnums.ItemExtent(ieCount)) - myStartIndex + 1
            iopResult.MessageEnum = Enums.Message.AsEnum(Id.RunLengthTooLong)
                
                
        End If
            
    Else                                         ' ipstartindex <0 and ipRun <0 then
            
        myStartIndex = iopResult.Item(ResultItemsEnums.ItemExtent(ieCount)) + ipStartIndex + 1
        If myStartIndex + ipRun + 1 > 0 Then
                
            myStartIndex = myStartIndex + ipRun + 1
            myRun = Abs(ipRun)
                
        Else
                
            myStartIndex = 1
            myRun = iopResult.Item(ResultItemsEnums.ItemExtent(ieCount)) - myStartIndex + 1
                
        End If
            
            
    End If
        
    ' ' Adjust startindex to align witth the true first index
    ' myStartIndex = myStartIndex - 1 + iopResult.Item(ResultItemsEnums.ItemExtent(IsFirstIndex))
    Set GetStartRunFromAnyStartRun = Types.Iterable.ToArrayList(myStartIndex, myRun)
        
End Function

Friend Function GetStartRunFromAnyStartEnd _
       ( _
       ByVal ipStartIndex As Long, _
       ByVal ipEndIndex As Long, _
       Optional ByRef iopResult As Result _
       ) As ArrayList
            
    Dim myStartIndex As Long
    Dim myEndIndex As Long
    Dim myRun As Long
            
    myStartIndex = IIf(ipStartIndex > 0, ipStartIndex, iopResult.Item(ResultItemsEnums.ItemExtent(ieCount)) + ipStartIndex + 1)
    myEndIndex = IIf(ipEndIndex > 0, ipEndIndex, iopResult.Item(ResultItemsEnums.ItemExtent(ieCount)) + ipEndIndex + 1)
            
    If myStartIndex < myEndIndex Then
                
        myRun = myEndIndex - myStartIndex + 1
                
    Else
                
        myRun = myStartIndex - myEndIndex + 1
        myStartIndex = myEndIndex
                
    End If
            
    ' ' Adjust startindex to align witth the true first index
    ' myStartIndex = myStartIndex - 1 + iopResult.Item(ResultItemsEnums.ItemExtent(IsFirstIndex))
    Set GetStartRunFromAnyStartEnd = Types.Iterable.ToArrayList(myStartIndex, myRun)
            
End Function

Friend Function GetStartRunFromAnyStartOnly _
       ( _
       ByVal ipStartIndex As Long, _
       Optional ByRef iopResult As Result _
       ) As ArrayList

    Dim myStartIndex As Long
    Dim myRun  As Long
    If ipStartIndex > 0 Then
            
        myStartIndex = ipStartIndex
        myRun = iopResult.Item(ResultItemsEnums.ItemExtent(ieCount)) - ipStartIndex + 1
            
    Else
        
        myStartIndex = iopResult.Item(ResultItemsEnums.ItemExtent(ieCount)) + ipStartIndex + 1
        myRun = VBA.Abs(ipStartIndex)
            
    End If

    ' ' Adjust startindex to align witth the true first index
    ' myStartIndex = myStartIndex - 1 + iopResult.Item(ResultItemsEnums.ItemExtent(IsFirstIndex))
    Set GetStartRunFromAnyStartOnly = Types.Iterable.ToArrayList(myStartIndex, myRun)

End Function

Friend Function GetStartRunFromAnyRunOnly(ByVal ipRun As Long, ByRef iopResult As Result) As ArrayList
        
    Dim myStartIndex As Long
    Dim myRun As Long
    If VBA.Abs(ipRun) > iopResult.Item(ResultItemsEnums.ItemExtent(ieCount)) Then
            
        myRun = iopResult.Item(ResultItemsEnums.ItemExtent(ieCount))
        myStartIndex = 1

            
    Else
            
        If ipRun < 0 Then
                
            myStartIndex = iopResult.Item(ResultItemsEnums.ItemExtent(ieCount)) + ipRun + 1
            myRun = VBA.Abs(ipRun)
                
                
        Else
                
            myStartIndex = 1
            myRun = ipRun
                
                
        End If
            
    End If
        
    ' ' Adjust startindex to align witth the true first index
    ' myStartIndex = myStartIndex - 1 + iopResult.Item(ResultItemsEnums.ItemExtent(IsFirstIndex))
    Set GetStartRunFromAnyRunOnly = Types.Iterable.ToArrayList(myStartIndex, myRun)
        
End Function

Friend Function GetStartRunFromAnyEndOnly(ByVal ipEndIndex As Long, ByRef iopResult As Result) As ArrayList
        
    'Dim myStartIndex As Long
    Dim myRun As Long
        
    'myStartIndex = iopResult.Item(ResultItemsEnums.ItemExtent(IsFirstIndex))
            
    If ipEndIndex < 0 Then
        
            
        myRun = iopResult.Item(ResultItemsEnums.ItemExtent(ieCount)) + ipEndIndex + 1
            
    Else
            
        myRun = ipEndIndex
            
            
    End If
        
    ' Adjust startindex to align witth the true first index
        
    Set GetStartRunFromAnyEndOnly = Types.Iterable.ToArrayList(1, myRun)
        
End Function

Friend Function GetStartRunFromNoStartRunOrEnd(ByRef iopExtent As Result) As ArrayList
        
    ' Dim myStartIndex As Long
    ' myStartIndex = iopExtent.Item(ResultItemsEnums.ItemExtent(IsFirstIndex))
        
    Dim myRun As Long
    myRun = iopExtent.Item(ResultItemsEnums.ItemExtent(ieCount))
        
    ' Adjust startindex to align witth the true first index
    'myStartIndex = myStartIndex - 1 + iopResult.Item(ResultItemsEnums.ItemExtent(IsFirstIndex))
    Set GetStartRunFromNoStartRunOrEnd = Types.Iterable.ToArrayList(1, myRun)

End Function

'@Description("Returns the specified subrange in Item(0) of the result")
Public Function TryRange _
       ( _
       ByVal ipIterable As Variant, _
       ByRef iopResult As Result, _
       Optional ByVal ipStartIndex As Variant, _
       Optional ByVal ipRun As Variant, _
       Optional ByVal ipEndIndex As Variant, _
       Optional ByVal ipRank As Variant = 1 _
       ) As Result
Attribute TryRange.VB_Description = "Returns the specified subrange in Item(0) of the result"
        
    If iopResult Is Nothing Then Set iopResult = Result.Deb
       
    If Types.Iterable.TryExtent(ipIterable, iopResult, ipRank).NotOkay Then
        
        Set TryRange = iopResult
        Exit Function
    
        
    End If
        
    ' Preserve the original first index to use later when converting back from start run values
    Dim myOriginalStartIndex As Long
    myOriginalStartIndex = iopResult.Item(ResultItemsEnums.ItemExtent(ieFirstIndex))
        
    If Ranges.TryStartRunFromAnyStartRunEnd(ipIterable, iopResult, ipStartIndex, ipRun, ipEndIndex, ipRank).NotOkay Then
        
        Set TryRange = iopResult
        Exit Function
        
        
    End If
        
    ' Convert the virtual start run into actual start end
    Dim myStartIndex As Long
    myStartIndex = iopResult.Item(ResultItemsEnums.StartRun(srStartIndex)) + myOriginalStartIndex
        
    Dim myEndIndex As Long
    myEndIndex = myStartIndex + iopResult.Item(ResultItemsEnums.StartRun(srRunLength)) + myOriginalStartIndex - 1
        
        
    Dim myList As Lyst = Lyst.Deb
    iopResult.Items.Clear
    Dim myIndex As Long
    For myIndex = myStartIndex To myEndIndex
            
        iopResult.Items.add ipIterable(myIndex)
            
    Next
        
    iopResult.SetOkay
    Set TryRange = iopResult
    
End Function