Attribute VB_Name = "OPTIONAL_VALUE"
Public Type Optional_long
    Has_value As Boolean
    Value As Long
End Type

Public Type Optional_string
    Has_value As Boolean
    Value As String
End Type

Public Type Optional_variant
    Has_value As Boolean
    Value As Variant
End Type
