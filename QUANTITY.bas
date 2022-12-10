Attribute VB_Name = "QUANTITY"
Private m_initialized As Boolean
Private m_mass_units As QUANTITY_TYPE

Public Type Quantity
    value As Double
    Unit As QUANTITY_UNIT
End Type


Private Sub Init_module()
    Set m_mass_units = New QUANTITY_TYPE
    
    m_mass_units.Add_new_unit_linearly_based_on_single_point "ò", 1, 1000
    
    
    
    m_mass_units.Set_as_readonly

End Sub









