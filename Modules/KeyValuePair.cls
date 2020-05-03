VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyValuePair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Class to hold Key/Value pairs, updated 27-Apr-18 for Excel 16.13 for Mac
' to avoid bug where user-defined objects passed to procs return Null for Public properties, but methods work OK

Private Key_ As String
Private Value_ As Variant
Private IsObject_ As Boolean

Public Property Get Key() As String
   Key = Key_
End Property

Public Property Let Key(KeyString As String)
    Key_ = KeyString
End Property

Public Property Get Value() As Variant
   If IsObject_ Then
      Set Value = Value_
   Else
      Let Value = Value_
   End If
End Property

Public Property Let Value(Var As Variant)
   Let Value_ = Var
   IsObject_ = False
End Property

Public Property Set Value(Obj As Object)
   Set Value_ = Obj
   IsObject_ = True
End Property

Public Property Get IsObject() As Boolean
   IsObject = IsObject_
End Property

