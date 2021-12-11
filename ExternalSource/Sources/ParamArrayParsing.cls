VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParamArrayParsing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Parses a ParamArray. Used for inputs to the Lyst and Kvp objects"
Option Explicit
'@PredeclaredId
'@Exposed
'@Folder("Helpers")
'@ModuleDescription("Parses a ParamArray. Used for inputs to the Lyst and Kvp objects")

'@Description("Parses an array of variants to a result value where items(0) is the count of iterables returned from item 1 onwards")
' Friend Function ParseToIterables(ByVal ipForwardedParamArray As Variant, ByRef iopResult As Result) As Result
        
'     If iopResult Is Nothing Then Set iopResult = Result.Deb
'     Set ParseToIterables = iopResult
        
'     iopResult.Item(ResultItemsEnums.ParseParamarray.IsCount) = Arrays.Count(ipForwardedParamArray)
        
'     Select Case iopResult.Item(ResultItemsEnums.ParseParamarray.IsCount)
        
'         Case 0
                
'             iopResult.MessageEnum = Enums.Message.AsEnum(Id.IterableHasNoItems)
'             Exit Function
                
                
'         Case 1
            
'             If Me.OneItem(ipForwardedParamArray(0), iopResult).NotOkay Then Exit Function
                
                
'         Case 2
                
'             If Me.TwoItems(ipForwardedParamArray, iopResult).NotOkay Then Exit Function
                
            
'         Case Else
                
'             iopResult.SetOkay
'             iopResult.Items(0) = ipForwardedParamArray
                
                    
'     End Select
        
'     Set ParseToIterables = iopResult
    
'   End Function

'@Descscription("Interprets the single Item in the ParamArray. The generated iterables are accessed through the Items Method of ParseResult.)
Friend Function OneItem(ByRef ipForwardedParamArray As Variant, iopResult As Result) As Result
Attribute OneItem.VB_Description = "Parses an array of variants to a result value where items(0) is the count of iterables returned from item 1 onwards"
    ' set item 1 of the .items lyst of the result object to an iterable
    If iopResult Is Nothing Then Set iopResult = Result.Deb
    Set OneItem = iopResult
       
    With iopResult
        
        Select Case True
                    
        Case Types.Group.IsNotIterable(ipForwardedParamArray)
                    
            .Item(1) = Lyst.Deb.Add(ipForwardedParamArray)
                    
                
        Case Types.Group.IsNotIterable(ipForwardedParamArray(0))
                
            .Item(1) = Lyst.Deb.Add(ipForwardedParamArray(0))
                    
                    
        Case Arrays.IsMDArray(ipForwardedParamArray(0))
            ' There is no special iterable action for arrays
            ' with more then two dimensions so return as single item
                
            .Item(1) = Lyst.Deb.Add(ipForwardedParamArray(0))
                    
                    
            'process a 2d array into a lyst of lysts
        Case Arrays.IsTableArray(ipForwardedParamArray(0))

            Set .Item(1) = Arrays.ToLystOfRanksAsLyst(ipForwardedParamArray(0), RankIsRowFirstItemActionIsNoAction)
                    
                
        Case Arrays.IsListArray(ipForwardedParamArray(0))
            ' Note: Not a Net ArrayList
            .Item(1) = Lyst.Deb.AddRange(ipForwardedParamArray(0))
                
                
        Case Types.Group.IsIterableItemsByEnum(ipForwardedParamArray(0))
            'Debug.Print TypeName(ipForwardedParamArray(0))
            .Item(1) = Lyst.Deb.AddRange(ipForwardedParamArray(0)) ' was set ToDo: Why does this work without set
                    
                    
        Case Types.Group.IsIterableKeysByEnum(ipForwardedParamArray(0))
                
            .Item(1) = Lyst.Deb.AddRange(ipForwardedParamArray(0).Items)
            .Item(2) = Lyst.Deb.AddRange(ipForwardedParamArray(0).Keys)
                
                    
        Case Else
                
            Set OneItem = iopResult.SetNotOkay(Enums.Message.AsEnum(UnexpectedItemInParamArray))
            ' .MessageEnum = enums.Message.AsEnum(UnexpectedItemInParamArray)
            Exit Function
                
                    
        End Select
            
        .SetOkay
            
            
    End With
        
    Set OneItem = iopResult
        
End Function

Friend Function TwoItems(ByVal ipForwardedParamArray As Variant, ByRef iopResult As Result) As Result
        
    If iopResult Is Nothing Then Set iopResult = Result.Deb
    Set TwoItems = iopResult
            
    Select Case True
            
    Case Types.Group.IsAdmin(ipForwardedParamArray(0))
            
        iopResult.MessageEnum = Enums.Message.AsEnum(Id.InvalidKey)
        Exit Function
        
                    
    Case Types.Group.IsNotIterable(ipForwardedParamArray(0)) 'Imlemented
        
        iopResult.SetOkay 2, Array(ipForwardedParamArray(0)), Array(ipForwardedParamArray(1))
                
                
    Case Arrays.IsMDArray(ipForwardedParamArray(0))
            
        iopResult.SetOkay 1, Array(ipForwardedParamArray(0))
        Exit Function
                
                
    Case Arrays.IsTableArray(ipForwardedParamArray(0))
                
        If TwoItemsWhereFirstItemIsTableArray(ipForwardedParamArray, iopResult).NotOkay Then Exit Function
                
            
    Case Types.Group.IsIterableItemsByEnum(ipForwardedParamArray(0)) '
            
        If TwoItemsWhereFirstItemIsIterableItemsByEnum(ipForwardedParamArray, iopResult).NotOkay Then Exit Function
                
                
    Case Types.Group.IsIterableKeysByEnum(ipForwardedParamArray(0))
            
        If TwoItemsWhereFirstItemIsIterableKeysByENum(ipForwardedParamArray, iopResult).NotOkay Then Exit Function
        
            
    Case Else
            
        Set TwoItems = iopResult.SetNotOkay(Enums.Message.AsEnum(Id.UnexpectedItemInParamArray))
        'iopResult.MessageEnum = enums.Message.AsEnum(Id.UnexpectedItemInParamArray)

            
    End Select
        
End Function

' Friend Function TwoItemsWhereFirstItemIsNotIterable(ByVal ipForwardedParamArray As Variant, ByRef iopresult As Result) As Result
        
'     If iopresult Is Nothing Then Set iopresult = Result.Deb
'     Set TwoItemsWhereFirstItemIsNotIterable = iopresult
            
'     iopresult.Item(0) = 2
'     iopresult.Item(1) = Array(ipForwardedParamArray(0))
'     iopresult.Item(2) = Array(ipForwardedParamArray(1))
            
        
' End Function


Friend Function TwoItemsWhereFirstItemIsTableArray(ByVal ipForwardedParamArray As Variant, ByRef iopResult As Result) As Result
        
    'determine what to do based on second item
    Select Case True
            
    Case Enums.TableToLystActions.HoldsMember(ipForwardedParamArray(1))
        ' ipForwardedParamArray(1) is an TableToLystActions enumeration member that says how to process the table
        Dim myLyst As Lyst
        Set myLyst = Enums.TableToLystActions.Action(ipForwardedParamArray(1)).ExecAction(ipForwardedParamArray(0))
                
        If myLyst.Count = 2 Then
                    
            iopResult.SetOkay 2, myLyst.Item(0), myLyst.Item(1)
                
                    
        Else
                
            iopResult.SetOkay 1, myLyst.Item(0)
                    
                    
        End If

            
    Case Arrays.IsMDArray(ipForwardedParamArray(1))
            
        Set TwoItemsWhereFirstItemIsTableArray = iopResult.SetNotOkay(Enums.Message.AsEnum(Id.InvalidType))
        ' iopresult.MessageEnum = enums.Message.AsEnum(Id.InvalidType)
        ' Exit Function
                
                
    Case Arrays.IsTableArray(ipForwardedParamArray(1))
        'Two tables so merge first column of table 1(0) with rows of table 2(1)
                
        iopResult.SetOkay _
            2, _
            Enums.TableToLystActions.Action(RankIsRowFirstItemActionIsSplit).ExecAction(ipForwardedParamArray(0)).Item(0), _
            Enums.TableToLystActions.Action(RankIsRowFirstItemActionIsNoAction).ExecAction(ipForwardedParamArray(1)).Item(0)
                
                
    Case Types.Group.IterableKeysByEnum(ipForwardedParamArray)
        ' table vs dictionary, first column of table to dictionary items
        iopResult.SetOkay _
            2, _
            Enums.TableToLystActions.Action(RankIsRowFirstItemActionIsSplit).ExecAction(ipForwardedParamArray(0)).Item(0), _
            ipForwardedParamArray(1).Items
            
    Case Types.Group.IsIterableItemsByEnum(ipForwardedParamArray(1))
        'table to iterable, first column of table to iterable
        iopResult.SetOkay _
            2, _
            Enums.TableToLystActions.Action(RankIsRowFirstItemActionIsSplit).ExecAction(ipForwardedParamArray(0)).Item(0), _
            ipForwardedParamArray(1)
                
                
    Case Else
            
        iopResult.SetNotOkay Enums.Message.AsEnum(Id.UnexpectedItemInParamArray)
        ' iopresult.MessageEnum =
        Exit Function
                
    End Select
    
    Set TwoItemsWhereFirstItemIsTableArray = iopResult
        
End Function

Friend Function TwoItemsWhereFirstItemIsIterableItemsByEnum(ByVal ipForwardedParamArray As Variant, ByRef iopResult As Result) As Result
    'determine what to do based on second item
    Select Case True
            
    Case Types.Group.IsPrimitive(ipForwardedParamArray(1))
        'iterable to primitive is iterable to const series of primitive
        iopResult.SetOkay _
            2, _
            ipForwardedParamArray(0), _
            Ranges.GetConstSeries(Types.Iterable.TryExtent(ipForwardedParamArray(0)).Item(ResultItemsEnums.ItemExtent(ieCount)), VBA.CVar(ipForwardedParamArray(1)))
            
    Case Arrays.IsMDArray(ipForwardedParamArray(1))
        'MDarrays have no special action with an iterable so error out
        Set TwoItemsWhereFirstItemIsIterableItemsByEnum = iopResult.SetNotOkay(Enums.Message.AsEnum(Id.InvalidType))
        ' iopresult.MessageEnum = enums.Message.AsEnum(Id.InvalidType)
        Exit Function
                
                
    Case Arrays.IsTableArray(ipForwardedParamArray(1))
        'Iterable vs Table is iterable vs table rows with rows of table 2(1)
                
        iopResult.SetOkay _
            2, _
            ipForwardedParamArray(0), _
            Enums.TableToLystActions.Action(RankIsRowFirstItemActionIsNoAction).ExecAction(ipForwardedParamArray).Item(0)
                
                
    Case Types.Group.IterableKeysByEnum(ipForwardedParamArray)
        ' iterable vs dictionary is iterable to dictionary items
        iopResult.SetOkay _
            2, _
            Enums.TableToLystActions.Action(RankIsRowFirstItemActionIsSplit).ExecAction(ipForwardedParamArray(0)).Item(0), _
            ipForwardedParamArray(1).Items
            
    Case Types.Group.IsIterableItemsByEnum(ipForwardedParamArray(1))
        'iterable to iterable is iterable to iterable
        iopResult.SetOkay _
            2, _
            ipForwardedParamArray(0), _
            ipForwardedParamArray(1)
                
                
    Case Else
            
        Set TwoItemsWhereFirstItemIsIterableItemsByEnum = iopResult.SetNotOkay(Enums.Message.AsEnum(Id.UnexpectedItemInParamArray))
        'iopresult.MessageEnum = enums.Message.AsEnum(Id.UnexpectedItemInParamArray)
         Exit Function
                
    End Select
        
End Function

Friend Function TwoItemsWhereFirstItemIsIterableKeysByENum(ByVal ipForwardedParamArray As Variant, ByRef iopResult As Result) As Result
        
    Select Case True
            
            
            
    End Select
        
End Function
