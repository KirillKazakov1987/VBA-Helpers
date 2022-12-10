VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QUANTITY_TYPE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_dict_of_units As New DICT_TEXT_TO_VARIANT

Private m_name As String
Private m_is_readonly As Boolean


Private Sub Class_Initialize()
    m_is_readonly = False
    
    m_dict_of_units.Case_sensitivity = True
End Sub


Public Sub Set_as_readonly()
    m_is_readonly = True
End Sub


Public Property Get Name() As String
    Name = m_name
End Property

Public Property Let Name(value As String)
    Debug.Assert (m_is_readonly = False)
    m_name = value
End Property



Public Property Get Count()
    Count = m_dict_of_units.Count
End Property


Public Function Contains_unit_with_name(unit_name As String) As Boolean
    Contains_unit_with_name = m_dict_of_units.Contains_key(unit_name)
End Function


Public Function Get_unit_by_name(unit_name As String) As QUANTITY_UNIT
     Dim u As New QUANTITY_UNIT
     
     Set u = m_dict_of_units.Get_value(unit_name)
     
     Set Get_unit_by_name = u
End Function



Public Sub Add_new_unit_linearly_based_on_single_point( _
    unit_name As String, _
    v_this As Double, _
    v_SI As Double)

    Debug.Assert (m_is_readonly = False)
    Debug.Assert (m_dict_of_units.Contains_key(unit_name) = False)
    
    Dim u As New QUANTITY_UNIT
    u.Init_linearly_based_on_single_point unit_name, v_this, v_SI

    m_dict_of_units.Add_or_replace u.Name, u
End Sub


Public Sub Add_new_unit_linearly_based_on_two_points( _
    unit_name As String, _
    v_this_1 As Double, _
    v_SI_1 As Double, _
    v_this_2 As Double, _
    v_SI_2 As Double)
    
    Debug.Assert (m_is_readonly = False)
    Debug.Assert (m_dict_of_units.Contains_key(unit_name) = False)
    
    Dim u As New QUANTITY_UNIT
    u.Init_linearly_based_on_two_points unit_name, v_this_1, v_SI_1, v_this_2, v_SI_2

    m_dict_of_units.Add_or_replace u.Name, u
End Sub

