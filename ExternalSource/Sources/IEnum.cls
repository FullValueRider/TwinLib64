VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEnum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
Option Explicit
'@Folder("Enums")
'@Interface
Public Function Strings() As Variant
End Function

Public Function Members() As Variant
End Function

Public Function Actions() As Variant
End Function

Public Function Ord(ByVal ipResultCode As Id) As Long
End Function

Public Property Get AsEnum(ByVal ipResultCode As Id) As Id
End Property

Public Function CanExec(ByVal ipResultCode As Id) As Boolean
End Function

Public Property Get ToString(ByVal ipResultCode As Id) As String
End Property

Public Property Get Exec(ByVal ipResultCode As Id, ByVal ipTablearray As Variant) As Object
End Property

Public Property Get IsMember(ByVal ipValue As Long) As Boolean
End Property

Public Property Get Name() As String
End Property

Public Property Get Count() As Long
End Property
