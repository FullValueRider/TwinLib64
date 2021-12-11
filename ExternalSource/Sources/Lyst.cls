VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lyst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "A Wrapper for the ArrayList object which provides intellisense and extends functionality"
'@IgnoreModule FunctionReturnValueAlwaysDiscarded, SuspiciousPredeclaredInstanceAccess
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
' Code line limit should be 120 characters.
' Comment line limit should be 80 characters
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
' Lyst Class
'
' This class is a wrapper for the ArrayList class
' the principle differences being
'
' 1.    It omits parts of the arraylistt class that cannot be acheved in VBA
' 2.    Methods return the Me instance so that methods can be chained
' 3.    It allows intellisense for the class
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
Option Explicit
'@ModuleDescription("A Wrapper for the ArrayList object which provides intellisense and extends functionality")
'@PredeclaredId
'@exposed
'@Folder("VBALib")

Private Type ReadOnly
        
    Location                                 As String
        
        
End Type

Private r                                   As ReadOnly


Private Type State

    Host                                    As ArrayList

        
End Type

Private s                                   As State


Private Type Properties
        
    Result                                   As Result
        
        
End Type

Private p                                   As Properties

Public Sub Class_Initialize()
        
    If VBA.Len(r.Location) = 0 Then
            
        r.Location = "VBALib.Lyst."
            
            
    End If
        
    If p.Result Is Nothing Then
            
        Set p.Result = Result.Deb.SetOkay
            
    End If
        
End Sub

'@Description("Returns a new instance of the Lyst Class. Optionally populates the instance with the result of For Each applied to ipIterable.")
Public Function Deb(Optional ByVal ipIterable As Variant) As Lyst
Attribute Deb.VB_Description = "Returns a new instance of the Lyst Class. Optionally populates the instance with the result of For Each applied to ipIterable."
        
    'Guard UsePredeclaredIdInstance, Not (Me Is Lyst), c.Location & "Deb"
        
    If Not VBA.IsMissing(ipIterable) Then
            
        Guard IsNotIterable, Types.Group.Iterables.LacksItem(ipIterable), r.Location & "Deb"
            
            
    End If
        
    With New Lyst
        
        Set Deb = .ReadyToUseInstance(ipIterable)
            
            
    End With
        
End Function

'@Description("For internal use only.  Instantiates a new instance of the lyst class")
Public Function ReadyToUseInstance(ByVal ipIterable As Variant) As Lyst
Attribute ReadyToUseInstance.VB_Description = "For internal use only.  Instantiates a new instance of the lyst class"
        
    Set ReadyToUseInstance = Me
    Set s.Host = New ArrayList
        
    If VBA.IsMissing(ipIterable) Then Exit Function
        
    Dim myItem As Variant
    For Each myItem In ipIterable
            
        s.Host.Add myItem
            
            
    Next
            
End Function

'@Enumerator
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4

    Dim enumerator As IUnknown
    '@Ignore DefaultMemberRequired
    Set enumerator = s.Host.GetEnumerator(0, s.Host.Count)
    Set NewEnum = enumerator

End Function

Public Function GetEnumerator() As Variant
    GetEnumerator = s.Host.GetEnumerator
End Function

'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
' Properties
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'@Description("Adds a comma separated list of items to the Lyst.  )")
Public Function Add(ParamArray ipArgs() As Variant) As Lyst
Attribute Add.VB_Description = "Adds a comma separated list of items to the Lyst.  )"
        
    If Arrays.HasNoItems(ipArgs) Then Exit Function
    Dim myItem As Variant
    For Each myItem In ipArgs
            
        s.Host.Add myItem
            
    Next
    
    Set Add = Me
        
End Function

' Friend Function InternalAdd(ByRef ipArgs As Variant) As Lyst
            
'     Dim myItem As Variant
'     For Each myItem In ipArgs
                
'         s.Host.Add myItem
        
'     Next
        
'     Return Me
        
' End function

'@Description("Adds the Items of a single iterable to the Lyst.")
Public Function AddRange(ByVal ipIterable As Variant) As Lyst
Attribute AddRange.VB_Description = "Adds the Items of a single iterable to the Lyst."
    ' The arraylist equivalent requires that ipiterable supports ICollection
        
    'Guard IsNotIterable, Types.Group.IsNotIterable(ipIterable), c.Location & "AddRange"
        
    Dim myItem As Variant
    For Each myItem In ipIterable
            
        s.Host.Add myItem
            
    Next
        
    Set AddRange = Me
        
End Function

'@Description("Returns a message indicating a the status of the called method")
Public Property Get Res() As Result

    Set Res = p.Result
End Property

Private Property Set Res(ByVal ipResult As Result)
    Set p.Result = ipResult
End Property

'@Description("Returns the number of Items in the Lyst instance)")
Public Function Count() As Long
Attribute Count.VB_Description = "Returns the number of Items in the Lyst instance)"
    Count = s.Host.Count
End Function

'@Description("Returns the number of Items in the Lyst instance equal to ipValue)")
Public Function CountItem(ByVal ipValue As Variant) As Long
Attribute CountItem.VB_Description = "Returns the number of Items in the Lyst instance equal to ipValue)"
    
    Dim myCount As Long
    myCount = 0
    Dim myItem As Variant
    For Each myItem In s.Host
            
        If myItem = ipValue Then myCount = myCount + 1
           
    Next
        
    CountItem = myCount
        
End Function

' '@Description("Counts the number of items in the lyst where the result of ipCmp is true")
' Public Function Count(ByVal ipValue As Variant, ByVal ipCmp As IComparer) As Long
    
    
'     Dim myCount As Long = 0
'     Dim myItem As Variant
'     For Each myItem In s.Host
            
'         if ipcmp(myitem,ipvalue) = ipValue then count+=1
            
'     Next
        
'     Return myCount
        
' End Function

      
' '@Description("Counts the number of items in the spacified range")
' Public Function Count _
' ( _
'     ByVal ipValue1 As Variant, _
'     ByRef ipCmp1 As IComparer, _
'     byval ipValue2 as variant, _
'     ByRef ipCmp2 As IComparer _
' ) As Long

'     Dim myCount As Long = 0
'     Dim myItem As Variant
        
'     For Each myItem In s.Host
            
'         if ipcmp1(myitem,ipvalue1) and ipCmp2(myitem,ipvalue2) then mycount+=1
            
'     Next
        
'     Return myCount
        
' End Function
      
   

''@Description("Returns the first item in the list.  Remove the item from the list")
'Public Function Dequeue() As Variant
'
'    Dequeue = Empty
'    If s.Host.Count = 0 Then Exit Function
'    If VBA.IsObject(s.Host.Item(0)) Then
'
'        Set Dequeue = s.Host.Item(0)
'
'
'    Else
'
'        Dequeue = s.Host.Item(0)
'
'    End If
'
'    s.Host.RemoveAt 0
'
'End Function

'@Description("Returns values taken sequentially from position 0 of the Lyst. If count 1 then an array of one item returned")
Public Function Dequeue(Optional ByVal ipCount As Long = 1) As Lyst
Attribute Dequeue.VB_Description = "Returns values taken sequentially from position 0 of the Lyst. If count 1 then an array of one item returned"
        
       
    If s.Host.Count < 1 Then
            
        p.Result.SetOkay.MessageEnum = Enums.Message.AsEnum(Id.RunLengthIsZero)
        Set Dequeue = Lyst.Deb
            
    End If

    If ipCount > Count Then
            
        p.Result.SetOkay.MessageEnum = Enums.Message.AsEnum(Id.RunLengthTooLong)
            
        'Check if what is happening here is legal
        Set Dequeue = Me.Reverse
        s.Host.Clear
        Exit Function
            
    End If
        
    Set Dequeue = s.Host.GetRange(0, IIf(ipCount > Count, Count, ipCount)).Reverse
    s.Host.RemoveRange 0, ipCount
        
End Function

'@Description("Alternative syntax for AddRange")
Public Function Enqueue(ParamArray ipArgs() As Variant) As Lyst
Attribute Enqueue.VB_Description = "Alternative syntax for AddRange"
        
    p.Result.SetOkay
        
        
    If Arrays.HasNoItems(ipArgs) Then
            
        p.Result.SetOkay.MessageEnum = Enums.Message.AsEnum(Id.IterableHasNoItems)
        Set Enqueue = Me
        Exit Function
            
    End If
        
    Dim myIterable As Variant
    If Arrays.HasOneItem(ipArgs) Then
    
        If ParamArrayParsing.OneItem(CVar(ipArgs), p.Result).NotOkay Then
        
            Set Enqueue = Me
            Exit Function
            
            
        End If
        
        s.Host.AddRange (p.Result.Item(1).ToArrayList)
                
            
    Else
            
        s.Host.AddRange Types.Iterable.ToArrayList(ipArgs)
            
            
    End If
    
    Set Enqueue = Me
            
End Function

'@Description("Alternative syntax for .Item(x))")
Public Function Peek(ByVal ipIndex As Long) As Variant
Attribute Peek.VB_Description = "Alternative syntax for .Item(x))"
    
    If VBA.IsObject(Item(ipIndex)) Then
    
        Set Peek = Item(ipIndex)
        
    Else
    
        Peek = Item(ipIndex)
        
    End If
        
End Function

Public Function Push(ParamArray ipArgs() As Variant) As Lyst
        
    p.Result.SetOkay
        
    If Arrays.HasNoItems(ipArgs) Then
            
        p.Result.SetOkay.MessageEnum = Enums.Message.AsEnum(Id.IterableHasNoItems)
        Set Push = Me
        Exit Function
            
    End If
        
    Dim myIterable As Variant
    If Arrays.HasOneItem(ipArgs) Then
    
        If ParamArrayParsing.OneItem(CVar(ipArgs), p.Result).NotOkay Then
        
            Set Push = Me
            Exit Function
            
        End If
        AddKnownIterableAsItems p.Result.Item(1)
        Set Push = Me
        
    Else
        AddKnownIterableAsItems ipArgs
        
            
           
    End If
        
    Set Push = Me
        
End Function

'@DefaultMember
'@Description("Allows read and write of an item value")
Public Property Get Item(ByVal ipIndex As Long) As Variant
Attribute Item.VB_Description = "Allows read and write of an item value"
Attribute Item.VB_UserMemId = 0

    'Todo: Update the guard statements below
    ' Guard Id.StartIndexSetToOne, ipIndex < ArrayListFirstIndex, i.Location & "Set Item"
    ' Guard Id.IndexMoreThanLastIndex, ipIndex > Me.LastIndex, i.Location & "Get Item"
    Dim myList As ArrayList
    Set myList = s.Host

    If VBA.IsObject(myList.Item(ipIndex)) Then
            
        Set Item = s.Host.Item(ipIndex)
                
    Else
            
        Item = s.Host.Item(ipIndex)
                
    End If
        
End Property

Public Property Let Item(ByVal ipIndex As Long, ByVal ipValue As Variant)
    'Todo: Update the guard statements below
    ' Guard Id.StartIndexSetToOne, ipIndex < ArrayListFirstIndex, i.Location & "Set Item"
    ' Guard Id.IndexMoreThanLastIndex, ipIndex > Me.LastIndex, i.Location & "Set Item"
        
    If VBA.IsObject(ipValue) Then
        
        Set s.Host.Item(ipIndex) = ipValue
            
    Else
        
        s.Host.Item(ipIndex) = ipValue
            
    End If
        
End Property

Public Property Set Item(ByVal ipIndex As Long, ByVal ipValue As Variant)
        
    Set s.Host.Item(ipIndex) = ipValue
        
End Property

'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
' Methods
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
' Lyst cannot use Is<Group> methods of the TypesGroup class as this leads to using
' Lysts of typenames before they have been populated
' thus in the Lyst class we use the pvLacksItem and pvHoldsItem Methods
' Public Function pvHoldsItem(ByVal ipGroupString As String, ByVal ipItem As Variant) As Boolean
'     pvHoldsItem = VBA.InStr(ipGroupString, TypeName(ipItem)) > 0
' End Function

' Public Function pvLacksItem(ByVal ipGroupString As String, ByVal ipItem As Variant) As Boolean
'     pvLacksItem = Not pvHoldsItem(ipGroupString, ipItem)
' End Function




Friend Function AddKnownItemsAsItems(ParamArray ipItems() As Variant) As Lyst
        
    Dim myItem As Variant
    For Each myItem In ipItems
        
        s.Host.Add myItem
            
                
    Next
        
    Set AddKnownItemsAsItems = Me
        
End Function

'@Description("Implements add for a single item produced by ParseParamArrayToLyst<test> thus does not retest via ParseParaArray... thereby avoiding an infinite recursion")
Friend Function AddKnownIterableAsItems(ByVal ipIterable As Variant) As Lyst
Attribute AddKnownIterableAsItems.VB_Description = "Implements add for a single item produced by ParseParamArrayToLyst<test> thus does not retest via ParseParaArray... thereby avoiding an infinite recursion"

    Dim myItem As Variant
    For Each myItem In ipIterable
            
        s.Host.Add myItem
            
            
    Next
        
    Set AddKnownIterableAsItems = Me
        
End Function

'**Error**
'Public Function BinarySearch(ByVal ipValue As Variant) As Long
'' Produces Invalid Procedure call or argument
'' Probably needs a comparer which cannot be provided by VBA
'
'    s.Host.Sort
'    BinarySearch = s.Host.BinarySearch(ipValue)
'End Function

'@Description("Removes all items from the Lyst")
Public Function Clear() As Lyst
Attribute Clear.VB_Description = "Removes all items from the Lyst"
        
    s.Host.Clear
    Set Clear = Me
   
        
End Function

'@Description("Returns a shallow copy of the Lyst")
Public Function Clone() As Lyst
Attribute Clone.VB_Description = "Returns a shallow copy of the Lyst"
    Set Clone = Lyst.Deb.AddKnownIterableAsItems(s.Host)
End Function

'@Description("Returns True if the Item is found in the Lyst. For compatibility with ArrayList syntax")
Public Function Contains(ByVal ipItem As Variant) As Boolean
Attribute Contains.VB_Description = "Returns True if the Item is found in the Lyst. For compatibility with ArrayList syntax"
    Contains = s.Host.Contains(ipItem)
End Function

'@Description("Alternative syntax for Contains(Item)")
Public Function HoldsItem(ByVal ipValue As Variant) As Boolean
Attribute HoldsItem.VB_Description = "Alternative syntax for Contains(Item)"
    HoldsItem = s.Host.Contains(ipValue)         ' Does this need to be typename
End Function

'@Description("Alternative syntax for 'Not Contains(item)'")
Public Function LacksItem(ByVal ipValue As Variant) As Boolean
Attribute LacksItem.VB_Description = "Alternative syntax for 'Not Contains(item)'"
    LacksItem = Not HoldsItem(ipValue)
End Function

' ArrayList:CopyTo:Overloads
' CopyTo(Array)
' Copies the entire ArrayList to a compatible one-dimensional Array, starting at the beginning of the target array.

' CopyTo(Array, Int32)
' Copies the entire ArrayList to a compatible one-dimensional Array, starting at the specified index of the target array.

' CopyTo(Int32, Array, Int32, Int32)
' Copies a range of elements from the ArrayList to a compatible one-dimensional Array, starting at the specified index of the target array.
'
'@Description("Copies the specified range to the byref array provided as the first parameter. Default is to copy the whole Lyst. For compatibility, the ToArray method should be preferred")
Public Sub CopyTo _
       ( _
       ByRef iparray As Variant, _
       Optional ByVal ipStartIndex As Variant, _
       Optional ByVal ipRun As Variant, _
       Optional ByVal ipEndIndex As Variant _
       )
Attribute CopyTo.VB_Description = "Copies the specified range to the byref array provided as the first parameter. Default is to copy the whole Lyst. For compatibility, the ToArray method should be preferred"
    ' ipStartIndex, tpRun and ipEndIndex need to be variants to allow the IsMissing property
    ' p.MsgId = enums.Message.AsEnum(Id.NotOkay)
    ' p.Status = Sys.NotOkay
    If p.Result Is Nothing Then
        
        Set p.Result = Result.Deb
        
    Else
        
        p.Result.SetToFullNotOkay
        
    End If
           
    If Arrays.IsNotListArray(iparray) Then
                
        p.Result.MessageEnum = Enums.Message.AsEnum(Id.IsNotAListArray)
        Exit Sub
            
    End If
        
    ' The parameters for TryStartRunFromAnyStartRunEnd are 1 based indexes
    ' so as to allow negative numbers
    ' Consequently, the real values for ipStartIndex/ipEndIndex need to be converted
    ' to a 1 based index
    ' for Lyst, the host will always be a 0 based ArrayList
    ' This we convert by adding 1 to startindex and endindex if they are provided
        
    ' If Types.IsNotMissing(ipStartIndex) Then ipStartIndex = ipStartIndex + 1
    ' If Types.IsNotMissing(ipEndIndex) Then ipEndIndex = ipEndIndex + 1
            
    Dim myStartRun As Result
    If Ranges.TryStartRunFromAnyStartRunEnd(s.Host, myStartRun, ipStartIndex, ipRun, ipEndIndex).NotOkay Then Exit Sub
            
    Dim myStartIndex As Long
    myStartIndex = myStartRun.Item(ResultItemsEnums.StartRun(srStartIndex))
        
    Dim myRun As Long
    myRun = myStartRun.Item(ResultItemsEnums.StartRun(srRunLength))

    Dim myUBound As Long
    myUBound = UBound(iparray)
        
    Dim myoffset As Long
    For myoffset = 0 To myRun - 1
            
        If myoffset > myUBound Then
                
            p.Result.SetOkay
            p.Result.MessageEnum = Enums.Message.AsEnum(Id.DestinationTooSmall)
            Exit Sub
                
                
        Else
                
            iparray(myoffset) = Me.Item(myoffset + myStartIndex)
                
        End If
            
    Next
        
    p.Result.SetOkay
        
End Sub

'@Description("Returns a new Lyst instance containing a shallow copy of the specified range")
Public Function GetRange(Optional ByVal ipStartIndex As Variant, Optional ByVal ipRun As Variant, Optional ByVal ipEndIndex As Variant) As Lyst
Attribute GetRange.VB_Description = "Returns a new Lyst instance containing a shallow copy of the specified range"

    Dim myList As Lyst
    Set myList = Lyst.Deb
    Set GetRange = myList
        
    If Types.IsNotMissing(ipStartIndex) Then ipStartIndex = ipStartIndex + 1
    If Types.IsNotMissing(ipEndIndex) Then ipEndIndex = ipEndIndex + 1
            
    Dim myStartRun As Result
    If Ranges.TryStartRunFromAnyStartRunEnd(Me, myStartRun, ipStartIndex, ipRun, ipEndIndex).NotOkay Then Exit Function
        
        
    Dim myAL As ArrayList
    Set myAL = s.Host.GetRange(myStartRun.Item(ResultItemsEnums.StartRun(srStartIndex)), myStartRun.Item(ResultItemsEnums.StartRun(srRunLength))).Clone
        
    myList.InjectHostArrayList myAL
        
End Function

Friend Property Get Host() As ArrayList
    Set Host = s.Host
End Property

Friend Property Set Host(ByRef ipArrayList As ArrayList)
    s.Host = ipArrayList
End Property

'@Description("Returns the Index of the first Item found in the subsection of the Lyst starting at ipStartIndex, through ipCount.   the whole list")
Public Function IndexOf(ByVal ipItem As Variant, Optional ByVal ipStartIndex As Variant, Optional ByVal ipCount As Variant) As Long
Attribute IndexOf.VB_Description = "Returns the Index of the first Item found in the subsection of the Lyst starting at ipStartIndex, through ipCount.   the whole list"
    ' Updated to allow index of objects to be returned using the ObjPtr method.
        
    IndexOf = -1
    
    If Types.IsMissing(ipStartIndex) Then
            
        ipStartIndex = 0
    
    End If

    If Types.IsMissing(ipCount) Then
            
        ipCount = Me.Count - ipStartIndex
            
        
    End If
        
    If VBA.IsObject(ipItem) Then
            
        IndexOf = IndexOfObject(ipItem, ipStartIndex, ipCount)
            
            
    Else
            
        IndexOf = IndexOfValue(ipItem, ipStartIndex, ipCount)
            
            
    End If
    
End Function

Friend Function IndexOfValue(ByVal ipValue As Variant, Optional ByVal ipStartIndex As Variant, Optional ByVal ipCount As Variant) As Long
    ' The following caveats apply for using ArrayList from twinbasic
    ' TB/VBA gives an error if the method call does not contain a startindex value
    ' TB/VBA gives an error if the count parameter is used
    ' The Index returned is counted from start of list not start of range
    ' if ipValue is not found  then -1 is returned
    ' ipCount is used to limit the range over which the search is conducted
        
    ' As the count parameter is not supported a subrange is extracted so we can just use Index of(value)
    Dim myRange As ArrayList
    Set myRange = s.Host.GetRange(ipStartIndex, ipCount)
        
    Dim myIndex As Long
    ' Remember that a start index is required in TB/VBA
    myIndex = myRange.IndexOf(ipValue, 0)
    If myIndex = -1 Then Exit Function
    IndexOfValue = myIndex + ipStartIndex
        
End Function

Friend Function IndexOfObject(ByVal ipItem As Variant, Optional ByVal ipStartIndex As Long = 0, Optional ByVal ipCount As Long = -1) As Long
        
    IndexOfObject = -1
    Dim myRange As ArrayList
    Set myRange = s.Host.GetRange(ipStartIndex, ipCount)
        
    Dim myIndex As Long
    ' Remember that a start index is required in TB/VBA
        
    For myIndex = 0 To myRange.Count - 1
        
        If VBA.ObjPtr(ipItem) = VBA.ObjPtr(s.Host.Item(myIndex)) Then
                
            IndexOfObject = myIndex + ipStartIndex
            Exit Function
                
                
        End If
            
    
    Next
        
End Function

Public Function InsertAt(ByVal ipIndex As Long, ParamArray ipArgs() As Variant) As Lyst

    Set InsertAt = Me
    If ipIndex < 0 Then Exit Function
    If ipIndex > s.Host.Count Then Exit Function
    If Arrays.HasNoItems(ipArgs) Then Exit Function
        
    
    ' If UBound(ipArgs) = 0 Then
            
    '     If Types.Group.IsNotIterable(ipArgs(0)) Then
            
    '         Set myArrayList = Types.Iterable.ToArrayList(ipArgs(0))
                
                
    '     Else
                
    '         Set myArrayList = Types.Iterable.ToArrayList(ipArgs(0))
                
            
    '     End If
            
    ' Else
            
    '     Set myArrayList = Types.Iterable.ToArrayList(ipArgs)
            
    ' End If
    Dim myArrayList As ArrayList
    Set myArrayList = Types.Iterable.ToArrayList(ipArgs)
    If ipIndex = s.Host.Count Then
        
        s.Host.AddRange myArrayList
               
    Else
        
        s.Host.InsertRange ipIndex - 1, myArrayList
            
    End If
        
End Function

'@Description("Returns a modified Lyst with values from the defined subrange of ipIterable inserted at ipInsertIndex.")
Public Function InsertRange _
       ( _
       ByVal ipInsertIndex As Variant, _
       ByVal ipIterable As Variant, _
       Optional ByVal ipStartIndex As Variant, _
       Optional ByVal ipRun As Variant, _
       Optional ByVal ipEndIndex As Variant, _
       Optional ByVal ipRank As Long = ArrayAdmin.FirstRank _
       ) As Lyst
Attribute InsertRange.VB_Description = "Returns a modified Lyst with values from the defined subrange of ipIterable inserted at ipInsertIndex."

    ' ipstartindex,iprun and ipEndIndex refer to to the desired range from ipiterable
    ' and are expressed as relative to the first or last item
    ' For the purposes of the ranges methods, ranges are always considered to start from one.
    ' ipInsertIndex is the index at which insertion starts
        
    p.Result.SetOkay
        
    If Types.Group.IsNotIterable(ipIterable) Then
            
        p.Result.SetNotOkay.MessageEnum = (Enums.Message.AsEnum(Id.IsNotIterable))
        Set InsertRange = Me
        Exit Function
            
    End If
        
    If Types.Iterable.HasNoItems(ipIterable) Then
            
        p.Result.SetOkay.MessageEnum = Enums.Message.AsEnum(Id.IterableHasNoItems)
        Set InsertRange = Me
        Exit Function
            
    End If
       
    ' Get the insertion index in terms of the one based indexing used for get range
    ' the +1 is essentially .firstindex
    Dim myStartRun As Result
    If Ranges.TryStartRunFromAnyStartRunEnd(s.Host, myStartRun, ipInsertIndex + 1).NotOkay Then
            
        Set InsertRange = Me
        Exit Function
            
            
    End If
            
    Dim myInsertIndex As Variant
    myInsertIndex = myStartRun.Item(ResultItemsEnums.StartRun(srStartIndex))
        
    ' Adjust start and end indexes from 0 based indexing to 1 based indexing
    ' If Types.IsNotMissing(ipStartIndex) Then ipStartIndex = ipStartIndex + 1
    ' If Types.IsNotMissing(ipEndIndex) Then ipEndIndex = ipEndIndex + 1
        
    'Get the range to insert from ipIterable
    Dim myInsertRange As Result
    If Ranges.TryRange(ipIterable, myInsertRange, ipStartIndex, ipRun, ipEndIndex).NotOkay Then
            
        Set p.Result = myInsertRange
        Set InsertRange = Me
        Exit Function
            
    End If
            
    p.Result.SetOkay.MessageEnum = myInsertRange.MessageEnum
        
    s.Host.InsertRange myInsertIndex - 1, myInsertRange.Items
        
    Set InsertRange = Me
        
End Function

'@Description("Returns the index of the last item in the Lyst or -1 if the Lyst count is 0")
Public Property Get LastIndex() As Long
Attribute LastIndex.VB_Description = "Returns the index of the last item in the Lyst or -1 if the Lyst count is 0"
        
    Dim myCount As Long
    myCount = s.Host.Count
    LastIndex = IIf(myCount = 0, -1, myCount - 1)
        
End Property

Public Function LastIndexOf(Optional ByVal ipValue As Variant, Optional ByVal ipStartIndex As Long = 0, Optional ByVal ipCount As Long = -1) As Long
        
    '@Ignore IIfSideEffect
    LastIndexOf = IIf(Count = 0, 0, Count - 1)
    If VBA.IsMissing(ipValue) Then Exit Function
            
    LastIndexOf = -1
    If ipStartIndex < 0 Then Exit Function
    If ipStartIndex >= s.Host.Count Then Exit Function
    If ipStartIndex + ipCount - 1 >= s.Host.Count Then Exit Function
            
    LastIndexOf = 0
    If ipCount = 0 Then Exit Function
        
        
    Dim myCount As Long
    '@Ignore IIfSideEffect
    myCount = IIf(ipCount = -1, s.Host.Count, ipCount + 1) - ipStartIndex
    ' ArrayList.Lastindexof will only accept a value from VBA
    ' so we use GetRange to first extract the range for lastindexof
    ' for the
    ' object,int32 and
    ' object, int32, int32 overloads
    '
    LastIndexOf = s.Host.GetRange(ipStartIndex, myCount).LastIndexOf(ipValue) + ipStartIndex
        
End Function

'@Description("Map applies the method defined by the function object to each item of the array and returns a new lyst")
Public Function MapIt(ByVal ipMapper As IMapper) As Lyst
Attribute Map.VB_Description = "Map applies the method defined by the function object to each item of the array and returns a new lyst"
        
    Dim myLyst As Lyst
    Set myLyst = Lyst.Deb
        
    Dim myItem As Variant
    For Each myItem In s.Host
        
        myLyst.Add ipMapper.ExecMapper(myItem)
        
    Next
        
    Set MapIt = myLyst
        
End Function

Public Function ReduceIt(ByVal ipReduce As IReduce) As Variant
    
   ReduceIt = ipReduce.ExecRed(Me)
    
End Function


Public Function RemoveValue(ByVal ipValue As Variant) As Lyst

    Set RemoveValue = Me
    s.Host.Remove ipValue
        
End Function


Public Function RemoveAt(ByVal ipIndex As Variant) As Lyst

    Set RemoveAt = Me
    If Types.Group.IsNotNumber(ipIndex) Then Exit Function
        
    If ipIndex < 0 Then Exit Function
    If ipIndex >= s.Host.Count Then Exit Function
    s.Host.RemoveAt ipIndex
        
End Function


Public Function RemoveRange(Optional ByVal ipStartIndex As Long = 0, Optional ByVal ipCount As Long = -1) As Lyst
        
    Set RemoveRange = Me
    If ipCount < 0 Then
            
        s.Host.Clear
        Exit Function
            
            
    End If
        
    If ipCount = 0 Then Exit Function
    If ipStartIndex < 0 Then Exit Function
    If ipStartIndex >= s.Host.Count Then Exit Function
        
    Dim myCount As Long
        
    If ipStartIndex + ipCount - 1 >= s.Host.Count Then
            
        myCount = s.Host.Count - ipStartIndex + 1
            
            
    Else
            
        myCount = ipCount
            
        
    End If
        
    s.Host.RemoveRange ipStartIndex, myCount
        
End Function

' See Sequence.GetSeries, Sequence.GetConstSeries
''@Description("Adds a repeated value to a Lyst (ArrayList equivalent is Lyst.Deb.Repeat)"
'Public Function Repeat(ByVal ipValue As Variant, ByVal ipCount As Long) As Lyst
'
'    Set Repeat = Me
'    If ipCount < 0 Then Exit Function
'
'    Dim myCount As Long
'    For myCount = 1 To ipCount
'
'        s.Host.Add ipValue
'
'    Next
'
'End Function

Public Function Reverse(Optional ByVal ipStartIndex As Long = 0, Optional ByVal ipCount As Long = 0) As Lyst
        
    Set Reverse = Me
        
    If ipStartIndex < 0 Then Exit Function
    If ipCount < 0 Then Exit Function
        
    If ipStartIndex = 0 And ipCount = 0 Then
        
        s.Host.Reverse
        Exit Function
            
            
    End If
        
    Dim myArrayList As ArrayList
    Set myArrayList = s.Host.GetRange(ipStartIndex, ipCount)
        
    '@Ignore MemberNotOnInterface
    myArrayList.Reverse
    s.Host.SetRange ipStartIndex, myArrayList
        
End Function

Public Function SetRange(ByVal ipStartIndex As Long, ByVal ipIterable As Variant) As Lyst
        
    Dim myStartIndex As Long
    myStartIndex = ipStartIndex
    Set SetRange = Me
        
    Dim myArrayList As ArrayList
    Set myArrayList = New ArrayList
    If myStartIndex < 0 Then Exit Function
    Dim mySize As Long
    mySize = Types.Iterable.Count(ipIterable)
    If mySize = Sys.NotOkay Then Exit Function
    If myStartIndex > s.Host.Count - mySize + 1 Then Exit Function
            
    If Types.Group.IsNotIterable(ipIterable) Then
            
        myArrayList.Add ipIterable
            
            
    ElseIf Types.IsTypeByName(Types.Name.OfArrayList, ipIterable) Then
        
        Set myArrayList = ipIterable
            
            
    Else
            
        Set myArrayList = Types.Iterable.ToArrayList(ipIterable)
        
            
    End If

    Dim myItem As Variant
    For Each myItem In myArrayList
            
        's.Host.SetRange myStartIndex, myArrayList
        s.Host.Item(myStartIndex) = myItem
        myStartIndex = myStartIndex + 1
            
            
    Next
        
End Function

'@Description("Sorts the whole list. Does not support sorting part of the list")
'ToDo: implement the missing ArrayList sort overloads
' 1 sort from index to end
' 2 sort from index through count
Public Function Sort(Optional ByVal ipStartIndex As Long = 0, Optional ByVal ipCount As Long = -1) As Lyst
Attribute Sort.VB_Description = "Sorts the whole list. Does not support sorting part of the list"
        
    Set Sort = Me
    If ipStartIndex < 0 Then Exit Function
    If ipStartIndex >= s.Host.Count Then Exit Function
    If ipCount = 0 Then Exit Function

    s.Host.Sort                                  'ipStartIndex, myCount
            
        
End Function

Public Function ToArray() As Variant
    ToArray = s.Host.ToArray
End Function

Public Function ToArrayList() As ArrayList
    Set ToArrayList = s.Host.Clone
End Function

Friend Function InjectHostArrayList(ByVal ipArrayList As ArrayList) As Lyst
    Set s.Host = ipArrayList.Clone
    Set InjectHostArrayList = Me
End Function

Public Function ToString(Optional ByVal ipSeparator As String = vbNullString) As String

    Stringifier.ResetItemSeparator ipSeparator
    Stringifier.ResetArrayMarkup vbNullString, vbNullString, vbNullString
    Stringifier.ResetObjectMarkup vbNullString, vbNullString, vbNullString
    ToString = Stringifier.ToString(s.Host)      ' myString
        
        
End Function

'@Description("Transposes a lyst of lysts")
Public Function Transpose() As Lyst
Attribute Transpose.VB_Description = "Transposes a lyst of lysts"
        
    Dim myTransposed As Lyst
    Set myTransposed = Lyst.Deb
    Dim myIndex As Long
    For myIndex = 0 To s.Host.Item(0).LastIndex
        
        myTransposed.Add Lyst.Deb
            
    Next
        
    Dim myItem As Variant
    For Each myItem In s.Host
            
        For myIndex = 0 To myItem.LastIndex
                
            myTransposed.Item(myIndex).Add myItem(myIndex)
                
        Next
            
    Next
        
    Set Transpose = myTransposed
        
End Function

Public Function First() As Variant
        
    Guard NullReference, s.Host Is Nothing, r.Location & "Last"
        
    If VBA.IsObject(s.Host.Item(0)) Then
            
        Set First = s.Host.Item(0)
            
            
    Else
            
        First = s.Host.Item(0)
            
            
    End If
            
End Function

'Public Function Filter(ByRef ipValue As Variant, ) As Lyst
'
'    Dim myLyst As Lyst = Lyst.Deb
'    Dim myItem As Variant
'    For Each myItem In s.Host
'
'        If myItem = ipValue Then myLyst.Add myItem
'
'    Next
'
'End Function

Public Function Filter(ipComparer As IComparer, Optional ByVal ipIndex As Variant = -1) As Lyst
        
    Dim myLyst As Lyst
    Set myLyst = Lyst.Deb
    Dim myItem As Variant
    
    If VBA.IsMissing(ipIndex) Or ipIndex = -1 Then
    
        For Each myItem In s.Host
            
            If ipComparer.ExecCmp(myItem) Then myLyst.Add myItem
                   
        Next
        
    Else
    
        For Each myItem In s.Host
           
            If ipComparer(myItem.Item(ipIndex)) Then myLyst.Add myItem
                    
        Next
        
    End If
        
    Set Filter = myLyst
            
End Function

Public Function FilterIt(ByVal ipFilter As IFilter) As Lyst
    Attribute Map.VB_Description = "Map applies the method defined by the function object to each item of the array and returns a new lyst"
            
    Set FilterIt = ipFilter.ExecFilter(Me)
        
            
End Function

' Public Function Filter(ByVal ipTestItem As Variant, ByRef ipComparer As IComparer) As Lyst
        
'     dim myLyst as Lyst = Lyst.Deb
'     Dim myItem as variant
'     for each myItem in s.Host
            
'          If ipComparer.ExecCmp(myItem, ipTestItem) Then myLyst.Add myItem
                
'     Next
        
'     Return myLyst
        
' End Function
    
Public Function Last() As Variant
        
    Guard NullReference, Types.IsSameInstance(s.Host, Nothing), r.Location & "Last"
            
    If VBA.IsObject(s.Host.Item(Me.LastIndex)) Then
            
        Set Last = s.Host.Item(Me.LastIndex)
            
            
    Else
            
        Last = s.Host.Item(Me.LastIndex)
            
            
    End If
            
End Function

'@Description("Clearer intent compared to 'Not Result.Status'")
Public Function NotOkay() As Boolean
Attribute NotOkay.VB_Description = "Clearer intent compared to 'Not Result.Status'"
    NotOkay = Not p.Result.Status
End Function

'@Description("Clearer intent compared to 'Result.Status')
Public Function Okay() As Boolean
    Okay = p.Result.Status
End Function

' '@DEscription("Sets the status to 'True' and MessageEnum to Okay only if is NotOkay")
' Private Function SetOkay(Optional ipMessage As enumMessages.Id) As Lyst
    
'     ' Preserve any message other than Not Okay
'     If VBA.IsMissing(ipMessage) Then
            
'         ' Preserve any message other than  NotOkay
'         If p.MsgId = Id.NotOkay Then p.MsgId = Id.Okay
                
'     Else
            
'         p.MsgId = ipMessage
      
'     End If
        
'     p.Status = Sys.Okay
        
'     return Me
'     Return
' End Function


' '@Description("Sets Status and Messageenum to NotOkay")
' Private Function SetNotOkay(Optional ipMessage As enumMessages.Id) As Lyst
    
'     ' Preserve any message other than Okay
'     If VBA.IsMissing(ipMessage) Then
            
'         If p.MsgId = Id.Okay Then p.MsgId = Id.NotOkay
                
'     Else
            
'         p.MsgId = ipMessage
      
'     End If
        
'     p.Status = Sys.NotOkay
        
'     Return Me
        
' End Function
    
' Private Function SetStatus(ByRef ipResult As Result) As Lyst
'     p.MsgId = ipResult.MessageEnum
'     p.Status = ipResult.Status
' End Function
