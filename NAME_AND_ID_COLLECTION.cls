VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NAME_AND_ID_COLLECTION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private dict_id_to_name As New DICT_INT32_TO_VARIANT
Private dict_name_to_id As New DICT_TEXT_TO_VARIANT
Private id_counter  As Long


Private Sub Class_Initialize()
    id_counter = 1
End Sub


Public Property Get Count()
    Count = dict_id_to_name.Count
End Property


Public Property Get Case_sensitivity() As Boolean
    Case_sensitivity = dict_name_to_id.Case_sensitivity
End Property

Public Property Let Case_sensitivity(val As Boolean)
    dict_name_to_id.Case_sensitivity = val
End Property


Public Function Add(name As String) As Long
    If dict_name_to_id.Contains_key(name) Then
        Add = dict_name_to_id.Get_value(name)
    Else
        dict_name_to_id.Add_or_replace name, id_counter
        dict_id_to_name.Add_or_replace id_counter, name
        Add = id_counter
        id_counter = id_counter + 1
    End If
End Function


Public Function Get_name(id As Long) As String
    Get_name = dict_id_to_name.Get_value(id)
End Function

Public Function Get_id(name As String) As Long
    Get_id = dict_name_to_id.Get_value(name)
End Function


Public Function Get_name_by_index(record_index As Long) As String
    Get_name_by_index = dict_name_to_id.Get_key(record_index)
End Function

Public Function Get_id_by_index(record_index As Long) As Long
    Get_id_by_index = dict_id_to_name.Get_key(record_index)
End Function


Public Function Contains_id(id As Long) As Boolean
    Contains_id = dict_id_to_name.Contains_key(id)
End Function

Public Function Contains_name(name As String) As Boolean
    Contains_name = dict_name_to_id.Contains_key(name)
End Function

