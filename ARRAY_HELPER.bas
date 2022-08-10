Attribute VB_Name = "ARRAY_HELPER"

Option Explicit
Option Base 0



Public Function Provide_array_f64(ParamArray values() As Variant) As Double()
    Dim ib As Long: ib = LBound(values)
    Dim ie As Long: ie = UBound(values)
    Dim N As Long: N = ie - ib + 1
    Dim result() As Double: ReDim result(0 To N - 1)
    
    Dim i As Long: For i = ib To ie
        result(0 + i - ib) = values(i)
    Next i
    
    Provide_array_f64 = result
End Function





Public Function Rank(arr As Variant) As Long
    If IsArray(arr) = False Then
        Rank = 0
    Else
        Dim i As Long: i = 0
        Dim N As Long
        On Error Resume Next
            Do While (Err = 0)
                i = i + 1
                N = UBound(arr, i)
            Loop
        On Error GoTo 0
        Rank = i - 1
    End If
End Function



Public Function Zip_1D_arrays(ParamArray single_dimension_arrays() As Variant)
    Debug.Assert IsArray(single_dimension_arrays)
    
    Dim first_j As Long: first_j = LBound(single_dimension_arrays)
    Dim last_j As Long: last_j = UBound(single_dimension_arrays)
    
    Dim min_first_i As Long: min_first_i = 1
    Dim max_last_i As Long: max_last_i = 0
    
    
    
    Dim arr As Variant
    Dim first_i As Long, last_i As Long
    Dim j As Long: For j = first_j To last_j
        arr = single_dimension_arrays(j)
        Debug.Assert IsArray(arr)
        Debug.Assert Rank(arr) = 1
        
        first_i = LBound(arr)
        last_i = UBound(arr)
        
        If first_i < min_first_i Then min_first_i = first_i
        If last_i > max_last_i Then max_last_i = last_i
    Next j
    
    
    
    Dim result As Variant: ReDim result(min_first_i To max_last_i, first_j To last_j)
    
    For j = first_j To last_j
        arr = single_dimension_arrays(j)
        first_i = LBound(arr)
        last_i = UBound(arr)
        
        Dim i As Long: For i = first_i To last_i
            result(i, j) = arr(i)
        Next i
    Next j
    
    Zip_1D_arrays = result
End Function


Public Function Convert_to_zero_base_array_f64(source_array() As Double) As Double()
    Dim li As Long: li = LBound(source_array)
    Dim ui As Long: ui = UBound(source_array)
    
    Dim N As Long: N = ui - li + 1
    
    Dim dst_arr() As Double: ReDim dst_arr(0 To N - 1)
    
    Dim i As Long: For i = 0 To N - 1
        dst_arr(i) = source_array(i + li)
    Next i
    
    Convert_to_zero_base_array_f64 = dst_arr
End Function



Public Function Is_sorted_f64(values() As Double) As Boolean
    Is_sorted_f64 = True

    Dim first_index As Long: first_index = LBound(values)
    Dim last_index As Long: last_index = UBound(values)

    Dim i As Long: For i = first_index + 1 To last_index
        If values(i) < values(i - 1) Then
            Is_sorted_f64 = False
            Exit Function
        End If
    Next i
End Function


Public Function Is_sorted_i32(values() As Long) As Boolean
    Is_sorted_i32 = True

    Dim first_index As Long: first_index = LBound(values)
    Dim last_index As Long: last_index = UBound(values)

    Dim i As Long: For i = first_index + 1 To last_index
        If values(i) < values(i - 1) Then
            Is_sorted_i32 = False
            Exit Function
        End If
    Next i
End Function
