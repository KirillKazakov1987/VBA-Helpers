Attribute VB_Name = "OPTIONAL_VALUE"
Public Type Optional_int32
    Has_value As Boolean
    Value As Long
End Type

Public Type Optional_float64
    Has_value As Boolean
    Value As Double
End Type

Public Type Optional_bool
    Has_value As Boolean
    Value As Boolean
End Type

Public Type Optional_string
    Has_value As Boolean
    Value As String
End Type

Public Type Optional_variant
    Has_value As Boolean
    Value As Variant
End Type
