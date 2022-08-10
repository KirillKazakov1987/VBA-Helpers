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


Private N As Long
Private aX() As Double
Private aY() As Double


Public Sub Init(x_values() As Double, y_values() As Double)
    aX = ARRAY_HELPER.Convert_to_zero_base_array_f64(x_values)
    aY = ARRAY_HELPER.Convert_to_zero_base_array_f64(y_values)

    Debug.Assert UBound(aX) = UBound(aY)
    Debug.Assert Is_sorted_f64(aX)
    
    N = UBound(aX)
End Sub



Public Property Get First_x()
    First_x = aX(0)
End Property

Public Property Get Last_x()
    Last_x = aX(N - 1)
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
    Dim M1, M2, M3 As Long
    M1 = 0
    M3 = N

    Dim y As Double

    Do While ((M3 - M1) > 1)
        M2 = (M3 + M1) \ 2
        If x > aX(M2) Then M1 = M2 Else M3 = M2
    Loop

    If x < aX(M1) Then
        y = aY(M1)
    ElseIf x > aX(M3) Then
        y = aY(M3)
    Else
        y = aY(M1) + (aY(M3) - aY(M1)) / (aX(M3) - aX(M1)) * (x - aX(M1))
    End If

    Calculate_y = y
End Function