VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TypesIterables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
Option Explicit
'@PredecalrediD
'@eXPOSED
'@Folder("VBALib.Types")
Private Type ReadOnly
        
    Location                    As String
    NotCountable                As Long

        
End Type

Private r                       As ReadOnly


' Public Enum IterableRangeInputType
        
'     IsAdd = -1          ' Input at last index. No change in order of input. Special Case of Insert
'     IsInsert = -2       ' Input at any index/key. No change in order of input
'     IsPush = -3         ' Input at any Index/Key. Order of input is reversed
        
        
' End Enum

Private Sub Class_Initialize()
    Guard InvalidNew, Types.IsNotSameInstance(Me, TypesIterables), r.Location & "New"
    r.Location = "VBALib.TypesIterables."
    r.NotCountable = 1
End Sub

'@Description("Result.items 0,1,2 are firstindex, lastindex and count if available")
Public Function TryExtent _
       ( _
       ByVal ipIterable As Variant, _
       Optional ByRef iopResult As Result, _
       Optional ByVal ipRank_one As Variant _
       ) As Result
Attribute TryExtent.VB_Description = "Result.items 0,1,2 are firstindex, lastindex and count if available"
        
    ipRank_one = IIf(VBA.IsMissing(ipRank_one), 1, ipRank_one)
    If iopResult Is Nothing Then Set iopResult = Result.Deb
    Set TryExtent = iopResult
        
    If Types.Group.IsNotIterable(ipIterable) Then
            
        iopResult.MessageEnum = Enums.Message.AsEnum(Id.IsNotIterable)
        Exit Function
            
            
    End If
        
    Select Case True
            
    Case VBA.IsArray(ipIterable)
            
        Arrays.TryExtent ipIterable, iopResult, ipRank_one
        Exit Function
                    
    Case Types.Group.IsIterableByMethod(ipIterable)
        ' Queues and stack type iterables don't allow indexes
                
        If ipIterable.Count = 0& Then
                    
            iopResult.MessageEnum = Enums.Message.AsEnum(IterableHasNoItems)
            Exit Function
                    
        Else
                    
            iopResult.SetOkay Empty, Empty, ipIterable.Count
                    
                    
        End If

                
    Case Types.Group.IsIterable(ipIterable)
                
                
        If ipIterable.Count = 0& Then
                    
            iopResult.MessageEnum = Enums.Message.AsEnum(IterableHasNoItems)
            Exit Function
                    
                    
        End If
                
        Dim myFirstIndex As Variant
        Dim myLastIndex As Variant
        If Types.IsTypeByName(Types.Name.OfCollection, ipIterable) Then

            myFirstIndex = 1&
            myLastIndex = ipIterable.Count
                    
                    
        Else
                    
            myFirstIndex = 0&
            myLastIndex = ipIterable.Count - 1
                
        End If
                
    End Select
            
    iopResult.SetOkay myFirstIndex, myLastIndex, ipIterable.Count
        
End Function

'@Description("Returns the number of items in a valid array or colllection object. An unitialised array returns -1. A collection object returns -1 if nothing and 0 if unpoulated0")")
Public Function Count(ByVal ipIterable As Variant, Optional ByVal ipRank_one As Variant) As Long

    ipRank_one = IIf(VBA.IsMissing(ipRank_one), 1, ipRank_one)
    Count = r.NotCountable
        
    'If VBA.IsEmpty(ipIterable) Then Exit Function
        
    Select Case True
                
    Case VBA.IsArray(ipIterable)
            
        If Arrays.HasNoItems(ipIterable) Then Exit Function
            
        Count = Arrays.Count(ipIterable, ipRank_one)
                    
            
    Case Types.Group.IsIterable(ipIterable)
            
        If ipIterable Is Nothing Then Exit Function
        Count = ipIterable.Count
            
    End Select
        
End Function

Public Function Sum(ByVal ipIterable As Variant) As Variant
        
    If Types.Group.IsNotIterable(ipIterable) Then
        
        If Types.Group.IsNotNumber(ipIterable) Then Exit Function
            
        Sum = -1
        Exit Function
            
    End If
        
    Dim myItem As Variant
    Dim mySum As Double
    For Each myItem In ipIterable
        
        mySum = mySum + myItem
            
    Next
        
    Sum = mySum
        
End Function

Public Function MinMax(ByVal ipIterable As Variant) As Variant
            
    If Types.Group.IsNotIterable(ipIterable) Then
        
        If Types.Group.IsNotNumber(ipIterable) Then Exit Function
            
        MinMax = Array(ipIterable, ipIterable)
        Exit Function
            
    End If
    
    Dim myItem As Variant
    Dim myMax As Variant
    Dim myMin As Variant
    For Each myItem In ipIterable
            
        If VBA.IsEmpty(myMin) Then
            
            myMax = myItem
            myMin = myItem
                
        End If
            
        If myItem < myMin Then myMin = myItem
        If myItem > myMax Then myMax = myItem
            
    Next

    MinMax = Array(myMin, myMax)
        
End Function

'Public Function IsSameSize _
'       ( _
'       ByRef ipTestIterable As Variant, _
'       ByRef ipRefIterable As Variant _
'       ) As Boolean
'
'
'    Dim myLocation As String
'    myLocation = r.Location & "IsSameSize"
'    Guard IsNotIterable, Types.Group.IsNotIterable(ipRefIterable, ipTestIterable), myLocation
'
'    IsSameSize = False
'
'    Dim myRefSize As Long
'    Dim myTestSize As Long
'
'    myRefSize = Count(ipRefIterable, Arrays.FirstRank)
'    If myRefSize = Sys.NotOkay Then Exit Function
'
'    myTestSize = Count(ipTestIterable, Arrays.FirstRank)
'    If myTestSize = Sys.NotOkay Then Exit Function
'
'    IsSameSize = myTestSize = myRefSize
'
'End Function

'Public Function IsSameSize _
'       ( _
'       ByRef ipTestIterable As Variant, _
'       ByRef ipRefIterable As Variant, _
'       Optional ipRefRank As Long = 1 _
'       ) As Boolean
'
'    Dim myLocation As String
'    myLocation = r.Location & "IsSameSize"
'    Guard IsNotIterable, Types.Group.IsNotIterable(ipRefIterable, ipTestIterable), myLocation
'
'    IsSameSize = False
'
'    Dim myRefSize As Long
'    Dim myTestSize As Long
'
'    myRefSize = Count(ipRefIterable, ipRefRank)
'    If myRefSize = Sys.NotOkay Then Exit Function
'
'    myTestSize = Count(ipTestIterable, Arrays.FirstRank)
'    If myTestSize = Sys.NotOkay Then Exit Function
'
'    IsSameSize = myTestSize = myRefSize
'
'End Function

'@Description("If only one rank dimesion is provided than the value used for refRank is arrays.firstrank and TestRank is set to")
Public Function IsSameSize _
       ( _
       ByRef ipTestIterable As Variant, _
       ByRef ipRefIterable As Variant, _
       ipRefRank As Long, _
       ipTestRank As Long _
       ) As Boolean
Attribute IsSameSize.VB_Description = "If only one rank dimesion is provided than the value used for refRank is arrays.firstrank and TestRank is set to"
            
    'If iptestrank is zero then iprefrank=arrays.firsstrank and iptestrank =iprefrank
    Dim myLocation As String
    myLocation = r.Location & "IsSameSize"
    Guard IsNotIterable, Types.Group.IsNotIterable(ipRefIterable, ipTestIterable), myLocation
            
    IsSameSize = False
            
    Dim myRefSize As Long
    Dim myTestSize As Long
        
    myRefSize = Count(ipRefIterable, ipRefRank)
    If myRefSize = Sys.NotOkay Then Exit Function
                
    myTestSize = Count(ipTestIterable, ipTestRank)
    If myTestSize = Sys.NotOkay Then Exit Function
            
    IsSameSize = myTestSize = myRefSize
            
End Function

'Public Function IsNotSameSize _
'       ( _
'       ipTestIterable As Variant, _
'       ipRefIterable As Variant _
'       ) As Boolean
'    IsNotSameSize = Not IsSameSize(ipTestIterable, ipRefIterable, Arrays.FirstRank, Arrays.FirstRank)
'End Function
'
'Public Function IsNotSameSize _
'       ( _
'       ipTestIterable As Variant, _
'       ipRefIterable As Variant, _
'       ipRefRank As Long _
'       ) As Boolean
'    IsNotSameSize = Not IsSameSize(ipTestIterable, ipRefIterable, ipRefRank, Arrays.FirstRank)
'End Function

Public Function IsNotSameSize _
       ( _
       ipTestIterable As Variant, _
       ipRefIterable As Variant, _
       Optional ipTestRank As Long = 1, _
       Optional ipRefRank As Long = 1 _
       ) As Boolean
    ipTestRank = IIf(VBA.IsMissing(ipTestRank), 1, ipTestRank)
    ipRefRank = IIf(VBA.IsMissing(ipRefRank), 1, ipRefRank)
    IsNotSameSize = Not IsSameSize(ipTestIterable, ipRefIterable, ipTestRank, ipRefRank)
End Function

'@Description("Returns true if each item in the iterable is the same type")
Public Function IsUniform(ByVal ipIterable As Variant) As Boolean
Attribute IsUniform.VB_Description = "Returns true if each item in the iterable is the same type"

    IsUniform = False
        
    If Not Types.Group.IsIterable(ipIterable) Then
        
        Err.Raise _
        17, _
        "VBALib.Types.IsIterableUniform", _
        "Expected an Iterable by value type"
                
        Exit Function
            
    End If
        
    ' Typename is used to avoid issues of vartype getting the type of the default
    ' member if ipIterable(x) is an object with a default member
    Dim myType  As String
    myType = TypeName(ipIterable(0))
        
    Dim myItem As Variant
    For Each myItem In ipIterable
        
        If TypeName(myItem) <> myType Then Exit Function
        
    Next
        
    IsUniform = True

End Function

Public Function IsNotUniform(ByVal iparray As Variant) As Boolean
    IsNotUniform = Not IsUniform(iparray)
End Function

Public Function HasNoItems(ByVal ipArg As Variant, Optional ByVal ipRank As Long = 1) As Boolean
    HasNoItems = Not HasItems(ipArg, ipRank)
End Function

'@Description("Returns true if arg is an array and has ubound of 0 or more, or is an object with a count greater than 0")
Public Function HasItems(ByVal ipArg As Variant, Optional ByVal ipRank As Long = 1) As Boolean
Attribute HasItems.VB_Description = "Returns true if arg is an array and has ubound of 0 or more, or is an object with a count greater than 0"
    
    ' Uninitialised variables will be objects or arrays
    ' not initialised means
    ' an object that is nothing
    ' an array that gives a ubound error
    ' and array that gives a ubound of -1
        
    HasItems = False
        
    If Types.Group.IsNotIterable(ipArg) Then Exit Function

    If VBA.IsArray(ipArg) Then
        
        If Arrays.HasNoItems(ipArg) Then Exit Function
            
    Else
        
            
        'Guard IsNotArray, ipRank <> 1, mcLocation & "HasItems", Array(TypeName(ipArg))
            
        If ipArg.Count <= 0 Then Exit Function

    End If
        
    HasItems = True
        
End Function

Public Function ToArray(ParamArray ipArgs() As Variant) As Variant

    ToArray = Empty
        
    If Arrays.HasNoItems(ipArgs) Then Exit Function
            
    Dim myArrayList As ArrayList
    Set myArrayList = New ArrayList
    Dim myItem As Variant
        
    If UBound(ipArgs) > 0 Then
            
        For Each myItem In ipArgs
            
            myArrayList.Add myItem
                
        Next
            
            
    ElseIf VBA.IsArray(ipArgs(0)) Then
            
        ToArray = ipArgs(0)
        Exit Function
            
            
    ElseIf Types.Group.IsIterableItemsByEnum(ipArgs(0)) Then
        
        Dim myargs As Variant
        myargs = CVar(ipArgs(0))
        For Each myItem In myargs
                
            myArrayList.Add myItem
                
        Next
            
            
    Else
        
        myArrayList.Add ipArgs(0)
            
            
    End If
        
    '@Ignore MemberNotOnInterface
    ToArray = myArrayList.ToArray
        
End Function

Public Function ToQueue(ParamArray ipArgs() As Variant) As Queue
        
    Dim myQueue As Queue
    Set myQueue = New Queue
    Set ToQueue = myQueue

    If Arrays.HasNoItems(ipArgs) Then Exit Function
        
    Dim myItem As Variant
    If UBound(ipArgs) > 0 Then
                
        For Each myItem In ipArgs
                
            '@Ignore MemberNotOnInterface
            myQueue.Enqueue myItem
                    
        Next
                
        Exit Function
                
    End If
            
    ' There is only one argument in ipargs
    If Types.IsTypeByName(Types.Name.OfQueue, ipArgs(0)) Then
                
        Set ToQueue = ipArgs(0).Clone
        Exit Function
                
    End If
            
    If Types.Group.IsIterableItemsByEnum(ipArgs(0)) Then
            
        Dim myargs As Variant
        myargs = CVar(ipArgs(0))
        For Each myItem In myargs
                    
            '@Ignore MemberNotOnInterface
            myQueue.Enqueue myItem
                    
        Next
                
        Exit Function
                
    End If
            
    '@Ignore MemberNotOnInterface
    myQueue.Enqueue ipArgs(0)
                
End Function

Public Function ToStack(ParamArray ipArgs() As Variant) As Stack
        
    Dim myStack As Stack
    Set myStack = New Stack
    Set ToStack = myStack

    If Arrays.HasNoItems(ipArgs) Then Exit Function
        
    Dim myItem As Variant
    If UBound(ipArgs) > 0 Then
                
        For Each myItem In ipArgs
                
            '@Ignore MemberNotOnInterface
            myStack.Push myItem
                    
        Next
                
        Exit Function
                
    End If
            
    ' There is only one argument in ipargs
    If Types.IsTypeByName(Types.Name.OfStack, ipArgs(0)) Then
                
        Set ToStack = ipArgs(0).Clone
        Exit Function
                
    End If
            
    If Types.Group.IsIterableItemsByEnum(ipArgs(0)) Then
            
        Dim myargs As Variant
        myargs = CVar(ipArgs(0))
        For Each myItem In myargs
                    
            '@Ignore MemberNotOnInterface
            myStack.Push myItem
                    
        Next
                
        Exit Function
                
    End If
            
    '@Ignore MemberNotOnInterface
    myStack.Push ipArgs(0)
                
End Function

''@Description("For the rare case where we wish to have a single iterable as a single iterable)
'Public Function SingleItemList(ByVal ipValue As Variant) As VBALib.Lyst
'
'    Dim myList As VBALib.Lyst
'    Set myList = New VBALib.Lyst
'    myList.Add ipValue
'    Set SingleItemList = myList
'End Function

Public Function ToArrayList(ParamArray ipArgs() As Variant) As ArrayList

    Dim myAL As ArrayList
    Set myAL = New ArrayList
    Set ToArrayList = myAL
    If VBA.IsMissing(ipArgs) Then Exit Function
    If VBA.IsEmpty(ipArgs) Then Exit Function
        
    
        
    Dim myInterim As ArrayList
    Set myInterim = New ArrayList
        
    Select Case True
            
    Case Arrays.IsNotSingleItem(ipArgs)
            
        myInterim.Add CVar(ipArgs)
        
        
    Case Types.Group.IsNotIterable(ipArgs(0))
            
        myInterim.Add CVar(ipArgs)
                
                
    Case Arrays.IsTableArray(ipArgs(0))
                    
        myInterim.Add Arrays.ToLystOfRanksAsLyst(ipArgs(0), RankIsRowFirstItemActionIsNoAction)
                    
                
    Case Arrays.IsMDArray(ipArgs(0))
                    
        myInterim.Add ipArgs(0)
                        
            
    Case Types.Group.IsIterableItemsByEnum(ipArgs(0))
            
        myInterim.Add ipArgs(0)
                
                
    Case Types.Group.IsIterableKeysByEnum(ipArgs(0))
            
        myInterim.Add ipArgs(0).Items
        'myInterim.Add ipArgs(0).Keys
            
            
    Case Else
            
        Guard Id.UnexpectedItemInParamArray, True, r.Location & "TryParseParamArray"
        Exit Function
                
            
    End Select
        
    Dim myItem As Variant
    For Each myItem In myInterim.Item(0)
            
        myAL.Add myItem
            
    Next
                
End Function

'@Description("Converts input to Lyst. If ipArgs is one iterable then the individual items of the iterable are added. Thus to add one iterable as an iterable, encapsulate it in an array")
Public Function ToLyst(ParamArray ipArgs() As Variant) As Lyst
Attribute ToLyst.VB_Description = "Converts input to Lyst. If ipArgs is one iterable then the individual items of the iterable are added. Thus to add one iterable as an iterable, encapsulate it in an array"
        
    Dim myLyst As Lyst
    Set myLyst = Lyst.Deb
    Set ToLyst = myLyst
        
    If Arrays.HasNoItems(ipArgs) Then Exit Function
        
    Dim myResult As Result
    If UBound(ipArgs) = 0 Then
            
        If Types.IsTypeByName(Types.Name.OfLyst, ipArgs(0)) Then
                
            Set ToLyst = ipArgs(0)
            Exit Function

        End If
            
        If ParamArrayParsing.OneItem(CVar(ipArgs), myResult).NotOkay Then Exit Function
                
            
    Else
            
        Set ToLyst = Lyst.Deb.AddKnownIterableAsItems(ipArgs)
        Exit Function
            
    End If
        
    ' the identification of lyst types probably only offers a very minor increase in speed
    If Types.IsTypeByName(Types.Name.OfLyst, myResult.Item(ResultItemsEnums.ParseParamArray(ppItems))) Then
        
        Set ToLyst = myResult.Item(ResultItemsEnums.ParseParamArray(ppItems))
                
            
    Else
                
        Set ToLyst = Lyst.Deb.AddKnownIterableAsItems(myResult.Item(ResultItemsEnums.ParseParamArray(ppItems)))

            
    End If
        
End Function

Public Function ToCollection(ParamArray ipArgs() As Variant) As Collection
        
    Dim myCollection As Collection
    Set myCollection = New Collection
    Set ToCollection = myCollection

    If Arrays.HasNoItems(ipArgs) Then Exit Function
        
    Dim myItem As Variant
    If UBound(ipArgs) > 0 Then
                
        For Each myItem In ipArgs
                
            myCollection.Add myItem
                    
        Next
                
        Exit Function
                
    End If
            
    ' There is only one argument in ipargs
    ' collection doesn't offer the clone method so no quick copy here
    If Types.Group.IsIterableItemsByEnum(ipArgs(0)) Then
            
        Dim myargs As Variant
        myargs = CVar(ipArgs(0))
        For Each myItem In myargs
                    
            myCollection.Add myItem
                    
        Next
                
        Exit Function
                
    End If
            
    myCollection.Add ipArgs(0)
        
End Function

' Public Function pvGetToThingArray(ByVal ipArgs As Variant) As Variant
        
'     ' ipArgs has already been triaged for empty
'     ' ipArgs is a parameter list from a 'ToThing' method where thing is an iterable array/object
'     ' the intent of this function is to process ipArgs to an array
'     ' in item 1 of the list.
'     ' In the case of a single item being provided to ipargs, and the single item is an
'     ' iterable type, the contents of the iterable type are added item by item
'     '
'     ' In the odd case where a single item is an iterable, which we want to add as
'     ' as single item, then that item should be encapsulated into an array
'     ' when provided as a parameter
        
'     Dim myList As ArrayList
'     Set myList = New ArrayList
        
'     ' if ipArgs is a single item
'     If UBound(ipArgs) = 0 Then
        
'         If Types.Group.IsPrimitive(ipArgs(0)) Then
            
'             myList.Add Array(ipArgs(0))
                
'         ElseIf Types.Group.IsIterableByMethod(ipArgs(0)) Then
'             ' captures object like stack and queue
'             Dim myarray As Variant
'             myarray = ipArgs(0).ToArray
'             myList.Add myarray
                
'         ElseIf Types.Group.IsIterableKeysByEnum(ipArgs(0)) Then
'             ' captures objects like Scripting.Dictionary and Kvp
'             myList.Add ipArgs(0).Items
                
'         ElseIf Types.Group.IsIterableItemsByEnum(ipArgs(0)) Then
'             ' Captures arrays and other object that can be enumerated
'             myList.Add ipArgs(0)
                
'         End If
            
'     Else
        
'         myList.Add ipArgs
            
'     End If
        
'     pvGetToThingArray = myList.ToArray

' End Function