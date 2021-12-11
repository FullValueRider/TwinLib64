VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Arrays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@Folder("VBALib")
'@PredeclaredId
'@Exposed

'12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
' Code line limit should be 120 characters.
' Comment line limit should be 80 characters
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
    
Public Enum ArrayAdmin
    
    NotCountable = 0
    FirstRank = 1
        
End Enum

Private Type ReadOnly
        
    Location As String

End Type

Private r                                     As ReadOnly


Public Enum ArrayDirections

    FourWay
    Eightway
        
End Enum

Public Enum RankOrientation

    ByRow
    ByColumn
        
End Enum

Private Sub Class_Initialize()
        
    If VBA.Len(r.Location) = 0 Then
            
        r.Location = "VBALib.Arrays."
            
            
    End If
        
    Guard Enums.Message.AsEnum(Id.InvalidNew), Not (Me Is Arrays), r.Location & "Class_Initialise"
        
End Sub

' Public Sub Class_Initialize()
' End Sub

'Public Function FirstRank() As Long: FirstRank = 1: End Function
'Public Function ArrayListFirstIndex() As Long: ArrayListFirstIndex = 0: End Function
    'ToDo: Move perimiter and associated functions to its own Area class
    ' ToDO use Lbound and Ubound to get min and max values
    '@Description("Returns an array of valid coordinates based on the provided x,y coordinate and the available area.")
Public Function GetPerimeterCoordinates _
       ( _
       ByVal ipCurRow As Long, _
       ByVal ipRowMin As Long, _
       ByVal ipRowMax As Long, _
       ByVal ipCurCol As Long, _
       ByVal ipColMin As Long, _
       ByVal ipColMax As Long, _
       ByVal ipType As ArrayDirections _
       ) As Variant
Attribute GetPerimeterCoordinates.VB_Description = "Returns an array of valid coordinates based on the provided x,y coordinate and the available area."

    Dim myLyst As Lyst
    Set myLyst = Lyst.Deb
        
    If ipCurRow - 1 >= ipRowMin Then myLyst.Add Array(ipCurRow - 1, ipCurCol)
    If ipCurRow + 1 <= ipRowMax Then myLyst.Add Array(ipCurRow + 1, ipCurCol)
        
    If ipCurCol - 1 >= ipColMin Then myLyst.Add Array(ipCurRow, ipCurCol - 1)
    If ipCurCol + 1 <= ipColMax Then myLyst.Add Array(ipCurRow, ipCurCol + 1)

    If ipType = ArrayDirections.Eightway Then
        
        If ipCurRow - 1 >= ipRowMin And ipCurCol - 1 >= ipColMin Then myLyst.Add Array(ipCurRow - 1, ipCurCol - 1)
        If ipCurRow - 1 >= ipRowMin And ipCurCol + 1 <= ipColMax Then myLyst.Add Array(ipCurRow - 1, ipCurCol + 1)
            
        If ipCurRow + 1 <= ipRowMax And ipCurCol - 1 >= ipColMin Then myLyst.Add Array(ipCurRow + 1, ipCurCol - 1)
        If ipCurRow + 1 <= ipRowMax And ipCurCol + 1 <= ipColMax Then myLyst.Add Array(ipCurRow + 1, ipCurCol + 1)
            
                
    End If
        
    GetPerimeterCoordinates = myLyst.ToArrayList.ToArray
        
End Function

Public Function GetMoveCoordinates(ByVal ipType As ArrayDirections) As Scripting.Dictionary

    Dim myDirections As Scripting.Dictionary
    Set myDirections = New Scripting.Dictionary
        
    With myDirections
                
        .Add "N", Array(1, 0)
        .Add "E", Array(0, 1)
        .Add "S", Array(-1, 0)
        .Add "W", Array(0, -1)
                
        If ipType = ArrayDirections.Eightway Then
                
            .Add "NE", Array(1, 1)
            .Add "NW", Array(1, -1)
            .Add "SW", Array(-1, -1)
            .Add "SE", Array(-1, 1)
                    
                
        End If
            
    End With
        
    Set GetMoveCoordinates = myDirections

End Function

' Public Function IsEmptyForwardedParamArray(ByRef ipForwardedParamarray As Variant) As Boolean
'     IsEmptyForwardedParamArray = (LBound(ipForwardedParamarray) = 0) And (UBound(ipForwardedParamarray) = -1)
' End Function

' Public Function IsSameSize _
' ( _
'     ByVal ipTestArray As Variant, _
'     ByVal ipRefArray As Variant, _
'     Optional ByVal ipTestRank_One As Variant, _
'     Optional ByVal ipRefRank_One As Variant _
' ) As Boolean
        
'     IsSameSize = False
'     ipTestRank_One = If(VBA.IsMissing(ipTestRank_One), 1, ipTestRank_One)
'     ipRefRank_One = If(VBA.IsMissing(ipTestRank_One), 1, ipRefRank_One)
        
'     If (ipTestRank_One = 0) And (ipRefRank_One = 0) Then
            
'         IsSameSize = pvIsArraySameSize(ipTestArray, ipRefArray)
            
            
'     Else
            
'         IsSameSize = pvIsRankSameSize(ipTestArray, ipRefArray, ipTestRank_One, ipRefRank_One)
            
            
'     End If

' End Function


Public Function pvIsRankSameSize _
       ( _
       ByVal ipTestArray As Variant, _
       ByVal ipRefArray As Variant, _
       Optional ByVal ipTestRank_One As Variant, _
       Optional ByVal ipRefRank_One As Variant _
       ) As Boolean
        
    pvIsRankSameSize = False
    ipTestRank_One = IIf(VBA.IsMissing(ipTestRank_One), 1, ipTestRank_One)
    ipRefRank_One = IIf(VBA.IsMissing(ipTestRank_One), 1, ipRefRank_One)
    If LacksRank(ipTestArray, ipTestRank_One) Then Exit Function
    If LacksRank(ipRefArray, ipRefRank_One) Then Exit Function
        
    pvIsRankSameSize = (UBound(ipTestArray) - LBound(ipTestArray)) = (UBound(ipRefArray) = LBound(ipRefArray))
        
End Function

Public Function pvIsArraySameSize _
       ( _
       ByVal ipTestArray As Variant, _
       ByVal ipRefArray As Variant _
       ) As Boolean

    pvIsArraySameSize = False
        
    Dim myTestRanks As Long
    myTestRanks = Ranks(ipTestArray)
    If myTestRanks = 0 Then Exit Function
                
    Dim myRefRanks As Long
    myRefRanks = Ranks(ipRefArray)
    If myRefRanks = 0 Then Exit Function
            
    If myRefRanks <> myTestRanks Then Exit Function
            
    Dim mySize As Long
    mySize = MinLong
        
    Dim myRank As Long
    For myRank = FirstRank To myRefRanks
        
        Dim myRefSize As Long
        myRefSize = Count(ipRefArray, myRank)
            
        If myRefSize = Sys.NotOkay Then Exit Function
        If myRefRanks = MinLong Then myRefRanks = myRefSize
        If mySize <> myRefSize Then Exit Function
                
        Dim myTestSize As Long
        myTestSize = Count(ipTestArray(myRank))
        If mySize <> myTestSize Then Exit Function
                
                
    Next
        
    pvIsArraySameSize = True
        
End Function

'@Description("Returns the number of dimensions of an array.  A return value of -1 means the input could not be treated as an array")
Public Function Ranks(ByVal iparray As Variant) As Long
Attribute Ranks.VB_Description = "Returns the number of dimensions of an array.  A return value of -1 means the input could not be treated as an array"
        
    Ranks = 0
    ' Can't call HasNoItems as this will introlduce circular dependencies.
    'If LacksItems(iparray) Then Exit Function
        
    'Visual Basic for Applications arrays can have up to 60000 dimensions; this allows for that
    Dim myDims As Long
    For myDims = 1 To 60000
        
        '@Ignore VariableNotUsed
        Dim myUBound As Long
        On Error Resume Next
        myUBound = UBound(iparray, myDims)
        If Err.Number <> 0 Then
                
            On Error GoTo 0
            Ranks = myDims - 1
            Exit Function
                
                
        End If
            
    Next
        
    On Error GoTo 0
        
End Function

'@Description("Returns True is an array has the requested dimension")
Public Function HasRank(ByVal iparray As Variant, Optional ByVal ipRank As Long = 1) As Boolean
Attribute HasRank.VB_Description = "Returns True is an array has the requested dimension"
        
    HasRank = False
    'If LacksItems(iparray) Then Exit Function
    If ipRank < 1 Then Exit Function
    If ipRank > Ranks(iparray) Then Exit Function
    HasRank = True
        
End Function

Public Function LacksRank(ByVal iparray As Variant, Optional ByVal ipRank As Long = 1) As Boolean
    LacksRank = Not HasRank(iparray, ipRank)
End Function

Public Function IsSingleItem(ByVal iparray As Variant) As Boolean
    IsSingleItem = (UBound(iparray) - LBound(iparray) + 1) = 1
End Function

Public Function IsNotSingleItem(ByVal iparray As Variant) As Boolean
    IsNotSingleItem = Not IsSingleItem(iparray)
End Function

Public Function IsListArray(ParamArray ipArgs() As Variant) As Boolean
    IsListArray = HasRanks(CVar(ipArgs), cmpEQ.Deb(1))
End Function

Public Function IsNotListArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotListArray = Not HasRanks(CVar(ipArgs), cmpEQ.Deb(1))
End Function

Public Function IsTableArray(ParamArray ipArgs() As Variant) As Boolean
    IsTableArray = HasRanks(CVar(ipArgs), cmpEQ.Deb(2))
End Function

Public Function IsNotTableArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotTableArray = Not HasRanks(CVar(ipArgs), cmpEQ.Deb(2))
End Function

Public Function IsMDArray(ParamArray ipArgs() As Variant) As Boolean
    IsMDArray = HasRanks(ipArgs, cmpMT.Deb(2))
End Function

Public Function IsNotMDArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotMDArray = Not IsMDArray(ipArgs, cmpMT.Deb(2))
End Function

Friend Function HasRanks(ByVal ipArgs As Variant, ByVal ipComparer As IComparer) As Boolean

    HasRanks = False
    If HasNoItems(ipArgs) Then Exit Function

    Dim myArg As Variant
    For Each myArg In ipArgs

        Dim myRanks As Long
        myRanks = Ranks(myArg)
        If Not ipComparer.ExecCmp(myRanks) Then Exit Function

    Next

    HasRanks = True

End Function

Friend Function LacksRanks(ByVal ipArgs As Variant, ByVal ipComparer As IComparer) As Boolean
    LacksRanks = Not HasRanks(ipArgs, ipComparer)
End Function

'@Description("Returns -1 for an unitialised array/not array otherwise the number of items in the rank of the array")
Public Function Count(ByRef iparray As Variant, Optional ByVal ipRank As ArrayAdmin = ArrayAdmin.FirstRank) As Long
Attribute Count.VB_Description = "Returns -1 for an unitialised array/not array otherwise the number of items in the rank of the array"
        
   
    If HasNoItems(iparray, ipRank) Then
            
        Count = ArrayAdmin.NotCountable
        
    Else
        
        Count = UBound(iparray, ipRank) - LBound(iparray, ipRank) + 1
            
    End If
            
End Function

'@Description("Returns -1 for an unitialised array/not array otherwise the total number of items in the array")
Public Function CountAll(ByRef iparray As Variant) As Long
Attribute CountAll.VB_Description = "Returns -1 for an unitialised array/not array otherwise the total number of items in the array"

    CountAll = ArrayAdmin.NotCountable
        
    If HasNoItems(iparray) Then Exit Function
        
    Dim myRank As Long
    Dim myRanks As Long
    myRanks = Ranks(iparray)
    Dim myCount As Long
    myCount = 1
    For myRank = 1 To myRanks
            
        myCount = myCount * Count(iparray, myRank)
            
            
    Next
            
    CountAll = myCount
            
End Function

Public Function HasItems(ByVal ipArray As Variant, Optional ByVal ipRank As ArrayAdmin = ArrayAdmin.FirstRank) As Boolean
    
    ' First test for having no ranks
    HasItems = False
    On Error Resume Next
    Dim myUbound As Long
    myUbound = UBound(ipArray)
    If Err.Number <> 0 Then
        On Error GoTo 0
        Exit Function
    End If
    
    'Does ipRank exist
    If ipRank < 1 Or ipRank > Ranks(ipArray) Then Exit Function
        
    ' Is it  variant containing an an uninitialised array    
    If UBound(ipArray, ipRank) < LBound(ipArray, ipRank) Then Exit Function
    
    HasItems = True
    
End Function

'

' Public Function HasNoItems(ByVal iparray As Variant) As Boolean
'     HasNoItems = Not TryExtent(iparray, ipRank_one:=Arrays.FirstRank)
' End Function
    
Public Function HasNoItems(ByVal ipArray As Variant, Optional ByVal ipRank As ArrayAdmin = ArrayAdmin.FirstRank) As Boolean
    HasNoItems = Not HasItems(ipArray, ipRank:=ipRank)
End Function

Public Function HasOneItem(ByVal ipArray As Variant, Optional ByVal ipRank As ArrayAdmin = ArrayAdmin.FirstRank) As Boolean
        
    HasOneItem = Count(ipArray, ipRank) = 1
End Function

Public Function TryExtent(ByRef ipArray As Variant, Optional ByRef iopResult As Result, Optional ByVal ipRank As ArrayAdmin = ArrayAdmin.FirstRank) As Result
        
    If iopResult Is Nothing Then Set iopResult = Result.Deb
    Set TryExtent = iopResult
        
    If LacksRank(ipArray, ipRank) Then
    
        iopResult.SetNotOkay
        iopResult.MessageEnum = Enums.Message.AsEnum(Id.InvalidRank)
        Exit Function
        
            
    End If
        
    Dim myErrorNo As Long
    On Error Resume Next
    Dim myUBound As Long
    myUBound = UBound(ipArray, ipRank)
    myErrorNo = Err.Number
        
    Err.Clear
    On Error GoTo 0
        
    If myErrorNo <> 0 Then
    
        iopResult.SetNotOkay (Enums.Message.AsEnum(Id.ArrayNotInitialised))
        Exit Function
            
            
     End If
        
    'An uninitialised variant array has Ubound of -1 and Lbound of 0
    Dim myLbound As Long
    myLbound = LBound(ipArray, ipRank)
    If myUBound < myLbound Then
            
         iopResult.SetNotOkay Empty, Empty, Empty, Empty, Empty
         iopResult.MessageEnum = Enums.Message.AsEnum(Id.ArrayNotInitialised)
         Exit Function
            
            
    End If
            
    iopResult.SetOkay myLbound, myUBound, myUBound - myLbound + 1
        
End Function

'@Description("Returns a Lyst of Lyst where the inner lyst is a lyst the selected rank(row or column) items")
Public Function ToLystOfRanksAsLyst _
       ( _
       ByVal ipTablearray As Variant, _
       Optional ByVal ipAction As TableToLystAction = TableToLystAction.RankIsRowFirstItemActionIsNoAction _
       ) As Lyst
Attribute ToLystOfRanksAsLyst.VB_Description = "Returns a Lyst of Lyst where the inner lyst is a lyst the selected rank(row or column) items"

    Guard NotTableArray, Me.IsNotTableArray(ipTablearray), r.Location & "ToLystOfRankLysts"
    Guard NotEnumerationMember, Enums.TableToLystActions.LacksMember(ipAction), r.Location & "TryToLystOfLysts"
    Dim myAction As ITableToLystAction
    Set myAction = Enums.TableToLystActions.Action(ipAction)
    Set ToLystOfRanksAsLyst = myAction.ExecAction(ipTablearray)
        
End Function

'@Description("transposes rows and colums of a 2d array")
Public Function Transpose(ByVal ipTablearray As Variant) As Variant
Attribute Transpose.VB_Description = "transposes rows and colums of a 2d array"
    
    ' 'Check for a jaggedarray
    ' If Ranks(ipTablearray) = 1 And Ranks(ipTablearray(LBound(ipTablearray))) = 1 Then
            
    '     Return TransposeJagged(ipTablearray)
            
            
    ' End If

    If IsNotTableArray(ipTablearray) Then
    
        Transpose = Empty
        Exit Function
        
    End If
        
    Dim myFirstRow As Long
    myFirstRow = LBound(ipTablearray, 1)
        
    Dim myLastRow As Long
    myLastRow = UBound(ipTablearray, 1)

    Dim myFirstCol As Long
    myFirstCol = LBound(ipTablearray, 2)
        
    Dim myLastCol As Long
    myLastCol = UBound(ipTablearray, 2)

    '@Ignore VariableNotAssigned
    Dim myTableArray As Variant
    ReDim myTableArray(myFirstCol To myLastCol, myFirstRow To myLastRow)
        
    Dim myRow As Long
    For myRow = myFirstRow To myLastRow
        
        Dim myCol As Long
        For myCol = myFirstCol To myLastCol
            
            myTableArray(myCol, myRow) = ipTablearray(myRow, myCol)
                
        Next
        
    Next
        
    '@Ignore UnassignedVariableUsage
    Transpose = myTableArray
        
End Function

' '@Description("Transposes a JaggedArray")
' Private Function TransposeJagged(ByVal ipJagged As Variant) As Variant
    
'     'as we are working with variants we dont care if we have an array of arrays of arrays
'     Dim myFirstRow As Long
'     myFirstRow = LBound(ipJagged, 1)
        
'     Dim myLastRow As Long
'     myLastRow = UBound(ipJagged, 1)

'     Dim myFirstCol As Long
'     myFirstCol = LBound(ipJagged(0), 1)
        
'     Dim myLastCol As Long
'     myLastCol = UBound(ipJagged(0), 1)
        
'     ' make a 2d array
'     Dim myTableArray As Variant
'     ReDim myTableArray(myFirstCol To myLastCol, myFirstRow To myLastRow)
        
'     Dim myRow As Long
'     For myRow = myFirstRow To myLastRow
            
'         Dim myCol As Long
'         For myRow = myFirstCol To myLastCol
                
'             myTableArray(myCol, myRow) = ipJagged(myRow)(myCol)
                
'         Next
            
'     Next
            
'     Dim mylyst As Lyst = enums.TableToLystActions.Action(RankIsColumnFirstItemActionIsNoAction).ExecAction(myTableArray)
        
'     Dim mytransposed As Variant = mylyst.ToArray
'     Dim myItem As Variant
'     For Each myitem In
        
' End Function
