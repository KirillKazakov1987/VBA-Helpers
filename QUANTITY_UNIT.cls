VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QUANTITY_UNIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CONVERTING_BY_LINEAR_DEPENDENCY  As Long = 1

Private m_converting_method As Long
Private m_name As String
Private m_is_initialized As Boolean

Private mult_from_SI As Double
Private mult_to_SI As Double
Private add_after_mult_from_SI As Double
Private add_after_mult_to_SI As Double


Private Sub Class_Initialize()
    m_converting_method = 0
    
    mult_from_SI = 1
    mult_to_SI = 1
    
    add_after_mult_from_SI = 0
    add_after_mult_to_SI = 0
    
    m_is_initialized = False
End Sub


Public Sub Init_linearly_based_on_single_point( _
    unit_name As String, _
    v_this As Double, _
    v_SI As Double)
    
    Debug.Assert Abs(v_this) > 0
    Debug.Assert Abs(v_SI) > 0

    m_converting_method = CONVERTING_BY_LINEAR_DEPENDENCY


    mult_to_SI = v_SI / v_this
    mult_from_SI = v_this / v_SI
    
    add_after_mult_to_SI = 0
    add_after_mult_from_SI = 0
    
    m_name = unit_name
    
    m_is_initialized = True

End Sub



Public Sub Init_linearly_based_on_two_points( _
    unit_name As String, _
    v_this_1 As Double, _
    v_SI_1 As Double, _
    v_this_2 As Double, _
    v_SI_2 As Double)
    
    Debug.Assert Abs(v_this_1 - v_this_2) > 0
    Debug.Assert Abs(v_SI_1 - v_SI_2) > 0

    m_converting_method = CONVERTING_BY_LINEAR_DEPENDENCY

    ' ===============================
    ' SIU = a * TU + b
    ' SIU1 = a * TU1 + b
    ' SIU2 = a * TU2 + b
    ' (SIU2 - SIU1) = a * (TU2 - TU1)
    ' a = (SIU2 - SIU1) / (TU2 - TU1)
    ' b = SIU1 - a * TU1
    ' ===============================
    Dim SIU1 As Double, SIU2 As Double, TU1 As Double, TU2 As Double
    Dim a As Double, b As Double
    
    SIU1 = v_SI_1
    SIU2 = v_SI_2
    TU1 = v_this_1
    TU2 = v_this_2
    a = (SIU2 - SIU1) / (TU2 - TU1)
    b = SIU1 - a * TU1
    
    
    ' ================
    ' TU = c * SIU + d
    ' ================
    Dim c As Double, d As Double
    
    c = 1 / a
    d = -1 * b / a


    mult_to_SI = a
    mult_from_SI = c
    
    add_after_mult_to_SI = b
    add_after_mult_from_SI = d
    
    m_name = unit_name
    
    m_is_initialized = True

End Sub



Public Property Get Name() As String
    Debug.Assert m_is_initialized
    Name = m_name
End Property


Public Function Convert_to_SI(quantity_value_in_this_unit As Double) As Double
    Debug.Assert m_is_initialized
    
    Select Case (m_converting_method)
        
        Case CONVERTING_BY_LINEAR_DEPENDENCY
            Convert_to_SI = quantity_value_in_this_unit * mult_to_SI + add_after_mult_to_SI

        Case Else
            Debug.Assert False
    
    End Select

End Function


Public Function Convert_from_SI(quantity_value_in_SI As Double) As Double
    Debug.Assert m_is_initialized
    
    Select Case (m_converting_method)
        
        Case CONVERTING_BY_LINEAR_DEPENDENCY
            Convert_from_SI = quantity_value_in_SI * mult_from_SI + add_after_mult_from_SI
            
        Case Else
            Debug.Assert False
    
    End Select
    
End Function


