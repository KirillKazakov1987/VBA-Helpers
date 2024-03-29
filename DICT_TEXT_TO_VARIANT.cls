VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DICT_TEXT_TO_VARIANT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private dict_base As New DICT_BASE_FOR_TEXT_KEY

Private values() As Variant


Private Sub Class_Initialize()
    ReDim values(0 To dict_base.Capacity - 1)
End Sub


Public Property Get Count() As Long
    Count = dict_base.Count
End Property


Public Property Get Capacity() As Long
    Capacity = dict_base.Capacity
End Property



Public Property Get Case_sensitivity() As Boolean
    Case_sensitivity = dict_base.Case_sensitivity
End Property



Public Property Let Case_sensitivity(val As Boolean)
    dict_base.Case_sensitivity = val
End Property


Private Sub Ensure_capacity(required_capacity As Long)
    dict_base.Ensure_capacity required_capacity
    ReDim Preserve values(0 To dict_base.Capacity - 1)
End Sub



Public Function Get_key(record_index As Long) As String
    Get_key = dict_base.Get_key(record_index)
End Function



Public Function Get_value_by_index(record_index As Long) As Variant
    Debug.Assert record_index >= 0
    Debug.Assert record_index < Count
    
    If IsObject(values(record_index)) Then
        Set Get_value_by_index = values(record_index)
    Else
        Get_value_by_index = values(record_index)
    End If
End Function


Public Function Try_get_value(key As String) As Optional_variant
    Dim optional_record_index As Optional_int32
    optional_record_index = dict_base.Try_get_record_index(key)
    
    If optional_record_index.Has_value Then
        Try_get_value.Has_value = True
        If IsObject(values(optional_record_index.value)) Then
            Set Try_get_value.value = values(optional_record_index.value)
        Else
             Try_get_value.value = values(optional_record_index.value)
        End If
    Else
        Try_get_value.Has_value = False
    End If
End Function



Public Function Get_value(key As String) As Variant
    Dim OPTIONAL_VALUE As Optional_variant
    OPTIONAL_VALUE = Me.Try_get_value(key)
    
    If OPTIONAL_VALUE.Has_value Then
        If IsObject(OPTIONAL_VALUE.value) Then
            Set Get_value = OPTIONAL_VALUE.value
        Else
            Get_value = OPTIONAL_VALUE.value
        End If
    Else
        Err.Raise 999999
    End If
End Function



Public Sub Add_or_replace(key As String, val As Variant)
    Dim cap0 As Long: cap0 = Me.Capacity
    
    Dim idx As Long
    idx = dict_base.Add_key_and_get_record_index(key)

    If Me.Capacity > cap0 Then
        Ensure_capacity Me.Count + 1
    End If
    
    If IsObject(val) Then
        Set values(idx) = val
    Else
        values(idx) = val
    End If
End Sub


Public Function Contains_key(key As String) As Boolean
    Contains_key = dict_base.Contains_key(key)
End Function

