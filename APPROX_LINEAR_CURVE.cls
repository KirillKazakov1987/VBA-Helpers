VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "APPROX_LINEAR_CURVE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0


Private n As Long
Private aX() As Double
Private aY() As Double


Public Sub Init(x_values() As Double, y_values() As Double)
    aX = ARRAY_HELPER.Change_lbound_of_array1D_f64(x_values)
    aY = ARRAY_HELPER.Change_lbound_of_array1D_f64(y_values)

    Debug.Assert UBound(aX) = UBound(aY)
    Debug.Assert Is_sorted_f64(aX)
    
    n = UBound(aX)
End Sub



Public Property Get First_x()
    First_x = aX(0)
End Property

Public Property Get Last_x()
    Last_x = aX(n - 1)
End Property



Public Sub Rescale_in_x(new_first_x As Double, new_last_x As Double)
    aX = MATH_HELPER.Rescale_1D_array_tying_to_boundaries(aX, new_first_x, new_last_x)
End Sub

Public Sub Rescale_in_y(new_first_y As Double, new_last_y As Double)
    aY = MATH_HELPER.Rescale_1D_array_tying_to_boundaries(aY, new_first_y, new_last_y)
End Sub



Public Function Clone() As APPROX_LINEAR_CURVE
    Dim instanse As New APPROX_LINEAR_CURVE
    instanse.Init aX, aY
    Clone = instanse
End Function


Public Function Calculate_y(x As Double) As Double
    Dim m1, m2, M3 As Long
    m1 = 0
    M3 = n

    Dim y As Double

    Do While ((M3 - m1) > 1)
        m2 = (M3 + m1) \ 2
        If x > aX(m2) Then m1 = m2 Else M3 = m2
    Loop

    If x < aX(m1) Then
        y = aY(m1)
    ElseIf x > aX(M3) Then
        y = aY(M3)
    Else
        y = aY(m1) + (aY(M3) - aY(m1)) / (aX(M3) - aX(m1)) * (x - aX(m1))
    End If

    Calculate_y = y
End Function
