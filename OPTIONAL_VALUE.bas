Attribute VB_Name = "OPTIONAL_VALUE"
Public Type Optional_int32
    Has_value As Boolean
    value As Long
End Type

Public Type Optional_float64
    Has_value As Boolean
    value As Double
End Type

Public Type Optional_bool
    Has_value As Boolean
    value As Boolean
End Type

Public Type Optional_string
    Has_value As Boolean
    value As String
End Type

Public Type Optional_variant
    Has_value As Boolean
    value As Variant
End Type

Public Type Optional_object
    Has_value As Boolean
    value As Object
End Type

