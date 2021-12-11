VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Result"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@PredeclaredId
'@Exposed
'@Folder("VBALib")


Private Type State
        
    Status                              As Boolean
    MsgId                               As Id
    ResultItems                         As ArrayList
        
End Type

Private s                               As State

Public Function Deb() As Result
        
    With New Result
        
        Set Deb = .ReadyToUseInstance
            
        
    End With
        
        
End Function

'@Description("Returns a Result object in its full NotOkey state")
Friend Function ReadyToUseInstance() As Result
Attribute ReadyToUseInstance.VB_Description = "Returns a Result object in its full NotOkey state"
        
    ' This method sets up Result in the NotOkay state with items being 5 empty values
    ' Methods requiring multivalue results should use an
    ' appropriate enumeration (see ResultFieldsEnums) to access the result fields
    s.Status = Sys.NotOkay
    s.MsgId = Enums.Message.AsEnum(NotOkay)
    ' Can't use Lyst here otherwise we get into a circular reference as Lyst uses Result
    Set s.ResultItems = New ArrayList
    s.ResultItems.Add Empty
    s.ResultItems.Add Empty
    s.ResultItems.Add Empty
    s.ResultItems.Add Empty
    s.ResultItems.Add Empty
        
    Set ReadyToUseInstance = Me
        
End Function

'@DefaultMember
Public Property Get Status() As Boolean
Attribute Status.VB_UserMemId = 0
    Status = s.Status
End Property

Public Property Let Status(ByVal ipStatus As Boolean)
        
    s.Status = ipStatus
    s.MsgId = IIf(s.Status, Enums.Message.AsEnum(Okay), Enums.Message.AsEnum(NotOkay))
            
End Property

'@Description("Returns a message of any issue causing a deviation from the happy path.  Messages can occur even if the operation was successful.")
Public Property Get MessageEnum() As Id
Attribute MessageEnum.VB_Description = "Returns a message of any issue causing a deviation from the happy path.  Messages can occur even if the operation was successful."
    MessageEnum = s.MsgId
End Property

Public Property Let MessageEnum(ByVal ipEnumMember As Id)
    s.MsgId = ipEnumMember
End Property

Public Property Get Item(ByVal ipResultFieldsEnum As Long) As Variant
        
    If VBA.IsObject(s.ResultItems.Item(ipResultFieldsEnum)) Then
            
        Set Item = s.ResultItems.Item(ipResultFieldsEnum)
            
            
    Else
            
        Item = s.ResultItems.Item(ipResultFieldsEnum)
            
            
    End If

End Property

Public Property Let Item(ByVal ipResultFieldsEnum As Long, ByVal ipItem As Variant)
        
    If VBA.IsObject(ipItem) Then
            
        Set s.ResultItems.Item(ipResultFieldsEnum) = ipItem
            
        
    Else
            
        s.ResultItems.Item(ipResultFieldsEnum) = ipItem
            
            
    End If
        
End Property

'@Description("Adds an Item to the Items collection of a Result object")
Public Sub Add(ByVal ipItem As Variant)
Attribute Add.VB_Description = "Adds an Item to the Items collection of a Result object"
    s.ResultItems.Add ipItem
End Sub

'@Description("Clearer intent compared to 'Not Result.Status'")
Public Function NotOkay() As Boolean
Attribute NotOkay.VB_Description = "Clearer intent compared to 'Not Result.Status'"
    NotOkay = Not s.Status
End Function

' Public Sub ClearItems()
'     s.ResultItems.clear
' End Sub

'@Description("Clearer intent compared to 'Result.Status')
Public Function Okay() As Boolean
    Okay = s.Status
End Function

'@Description("Returns a Lyst of the result items collection")
Public Property Get Items() As ArrayList
Attribute Items.VB_Description = "Returns a Lyst of the result items collection"
    
    Set Items = s.ResultItems
End Property

Public Property Set Items(ByVal ipItems As ArrayList)
    Set s.ResultItems = ipItems
End Property
' Friend Property Get Host() As ArrayList
'     Set Host = s.ResultItems
    
' End Property

' Friend Property Set Host(ByRef ipArrayList As ArrayList)
'     s.ResultItems = ipArrayList
' End Property


'@Enumerator
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4

    Dim enumerator As IUnknown
    Set enumerator = s.ResultItems.GetEnumerator '  (0, s.Host.Count)
    Set NewEnum = enumerator

End Function

'@DEscription("Sets the status to 'True' and MessageEnum to Okay only if is NotOkay")
Public Function SetOkay(ParamArray ipArgs() As Variant) As Result
Attribute SetOkay.VB_Description = "Sets the status to 'True' and MessageEnum to Okay only if is NotOkay"
        
    ' Preserve any message other than Not Okay
    If s.MsgId = Id.NotOkay Then s.MsgId = Id.Okay
    s.Status = Sys.Okay
        
    Set SetOkay = Me
    If VBA.IsMissing(ipArgs) Then Exit Function
    If VBA.IsEmpty(ipArgs) Then Exit Function
    If UBound(ipArgs) = -1 Then Exit Function
            
    Set s.ResultItems = Types.Iterable.ToArrayList(ipArgs)
    
    Set SetOkay = Me
            
End Function

'@Description("Sets Status and Messageenum to NotOkay")
Public Function SetNotOkay(ParamArray ipArgs() As Variant) As Result
Attribute SetNotOkay.VB_Description = "Sets Status and Messageenum to NotOkay"
        
    If s.MsgId = Id.Okay Then s.MsgId = Id.NotOkay
    s.Status = Sys.NotOkay
        
        
    Set SetNotOkay = Me
    If VBA.IsMissing(ipArgs) Then Exit Function
    If VBA.IsEmpty(ipArgs) Then Exit Function
    If UBound(ipArgs) = -1 Then Exit Function
            
    Set s.ResultItems = Types.Iterable.ToArrayList(ipArgs)
    
    Set SetNotOkay = Me
        
End Function

'@Description("Sets the Result to the 'NotOkay' or 'Failed' Result.Deb state")
Public Function SetToFullNotOkay() As Result
Attribute SetToFullNotOkay.VB_Description = "Sets the Result to the 'NotOkay' or 'Failed' Result.Deb state"
    s.MsgId = Id.NotOkay
    s.Status = Sys.NotOkay
    Set s.ResultItems = Lyst.Deb.AddKnownItemsAsItems(Empty, Empty, Empty, Empty, Empty).ToArrayList
    Set SetToFullNotOkay = Me
End Function
