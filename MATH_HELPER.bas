Attribute VB_Name = "MATH_HELPER"

Private Sub Test()
    Dim aX() As Double
    aX = ARRAY_HELPER.Provide_array_f64(0, 50, 100)
    
    Dim aY() As Double
    aY = ARRAY_HELPER.Provide_array_f64(0, 25, 50)
    
    
    Dim curve As New APPROX_LINEAR_CURVE
    curve.Init aX, aY
    
    curve.Rescale_in_y 50, 100
    
    
    Dim aXT() As Double
    aXT = ARRAY_HELPER.Provide_array_f64(-10, -1, 0, 1, 25, 49, 50, 51, 52, 75, 100, 101, 1000, 10000)
    
    Dim N As Long: N = UBound(aXT) - LBound(aXT) + 1
    Dim aYT() As Double: ReDim aYT(0 To N - 1)
    
    Dim i As Long: For i = 0 To N - 1
        Dim x As Double: x = aXT(i)
        Dim y As Double: y = curve.Calculate_y(x)
        aYT(i) = y
    Next i


    'DEBUG_HELPER.Print_into_debug Zip_1D_arrays(aX, aY, aXT, aYT)
End Sub



Public Function Max_f64(x As Double, y As Double) As Double
    If x > y Then Max_f64 = x Else Max_f64 = y
End Function

Public Function Max_i32(x As Long, y As Long) As Long
    If x > y Then Max_i32 = x Else Max_i32 = y
End Function

Public Function Min_i32(x As Long, y As Long) As Long
    If x < y Then Min_i32 = x Else Min_i32 = y
End Function


Public Function Clamp_i32(x As Long, min As Long, max As Long) As Long
    If x < min Then
        Clamp_i32 = min
    ElseIf x > max Then
        Clamp_i32 = max
    Else
        Clamp_i32 = x
    End If
End Function




Public Function Rescale_1D_array_tying_to_boundaries( _
        old_x_values() As Double, _
        new_first_x As Double, _
        new_last_x As Double) As Double()

    Dim first_index As Long: first_index = LBound(old_x_values)
    Dim last_index As Long: last_index = UBound(old_x_values)

    Dim old_first_x As Double: old_first_x = old_x_values(first_index)
    Dim old_last_x As Double: old_last_x = old_x_values(last_index)

    Dim old_span_x As Double: old_span_x = old_last_x - old_first_x
    Dim new_span_x As Double: new_span_x = new_last_x - new_first_x

    Debug.Assert Abs(old_span_x) > 0.000000000000001
    Debug.Assert Abs(new_span_x) > 0.000000000000001

    Dim new_x_values() As Double: ReDim new_x_values(first_index To last_index)
    
    Dim x_old As Double
    Dim x_new As Double

    Dim scaler As Double: scaler = new_span_x / old_span_x

    Dim i As Long: For i = first_index To last_index
        x_old = old_x_values(i)
        x_new = new_first_x + (x_old - old_first_x) * scaler
        new_x_values(i) = x_new
    Next i
    
    Rescale_1D_array_tying_to_boundaries = new_x_values
End Function

