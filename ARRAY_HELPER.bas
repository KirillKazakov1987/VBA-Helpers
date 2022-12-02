Attribute VB_Name = "ARRAY_HELPER"
Option Explicit
Option Base 0

Public Function Repeat( _
    v As Variant, _
    first_inclusive_index As Long, _
    last_inclusive_index As Long) As Variant
    
    Dim n As Long
    n = last_inclusive_index - first_inclusive_index + 1
    Debug.Assert n > 0
    
    Dim result As Variant
    ReDim result(first_inclusive_index To last_inclusive_index)
    
    Dim i As Long
    For i = first_inclusive_index To last_inclusive_index
        result(i) = v
    Next i
    
    Repeat = v
End Function



Public Function Provide_array(ParamArray arrays() As Variant) As Variant
    Dim n1 As Long: n1 = LBound(arrays)
    Dim n2 As Long: n2 = UBound(arrays)
    Dim result As Variant: ReDim result(n1 To n2)
    Dim i As Long
    For i = n1 To n2
        result(i) = arrays(i)
    Next i
    Provide_array = result
End Function


Public Function Provide_array_f64(ParamArray values() As Variant) As Double()
    Dim ib As Long: ib = LBound(values)
    Dim ie As Long: ie = UBound(values)
    Dim n As Long: n = ie - ib + 1
    Dim result() As Double: ReDim result(0 To n - 1)
    
    Dim i As Long: For i = ib To ie
        result(0 + i - ib) = values(i)
    Next i
    
    Provide_array_f64 = result
End Function



Public Function Get_item_of_array1D( _
    array1D As Variant, _
    zero_based_index As Long) As Variant
    
    Get_item_of_array1D = array1D(LBound(array1D) + zero_based_index)
End Function

Public Sub Set_item_of_array1D( _
    array1D As Variant, _
    zero_based_index As Long, _
    new_value As Variant)
    
    array1D(LBound(array1D) + zero_based_index) = new_value
End Sub


Public Function Get_item_of_array2D( _
    array2D As Variant, _
    zero_based_row_index As Long, _
    zero_based_column_index As Long) As Variant
    
    Dim i As Long
    Dim j As Long
    i = LBound(array2D, 1) + zero_based_row_index
    j = LBound(array2D, 2) + zero_based_column_index
    
    Get_item_of_array2D = array2D(i, j)
End Function


Public Sub Set_item_of_array2D( _
    array2D As Variant, _
    zero_based_row_index As Long, _
    zero_based_column_index As Long, _
    new_value As Variant)
    
    Dim i As Long
    Dim j As Long
    i = LBound(array2D, 1) + zero_based_row_index
    j = LBound(array2D, 2) + zero_based_column_index
    
    If IsObject(new_value) Then
        Set array2D(i, j) = new_value
    Else
        array2D(i, j) = new_value
    End If
End Sub



Public Sub Set_item_of_array2D_of_byte( _
    array2D() As Byte, _
    zero_based_row_index As Long, _
    zero_based_column_index As Long, _
    new_value As Byte)
    
    Dim i As Long
    Dim j As Long
    i = LBound(array2D, 1) + zero_based_row_index
    j = LBound(array2D, 2) + zero_based_column_index
    
    array2D(i, j) = new_value
End Sub



Public Sub Set_item_of_array2D_of_int16( _
    array2D() As Integer, _
    zero_based_row_index As Long, _
    zero_based_column_index As Long, _
    new_value As Integer)
    
    Dim i As Long
    Dim j As Long
    i = LBound(array2D, 1) + zero_based_row_index
    j = LBound(array2D, 2) + zero_based_column_index
    
    array2D(i, j) = new_value
End Sub


Public Sub Set_item_of_array2D_of_int32( _
    array2D() As Long, _
    zero_based_row_index As Long, _
    zero_based_column_index As Long, _
    new_value As Long)
    
    Dim i As Long
    Dim j As Long
    i = LBound(array2D, 1) + zero_based_row_index
    j = LBound(array2D, 2) + zero_based_column_index
    
    array2D(i, j) = new_value
End Sub



Public Sub Set_item_of_array2D_of_int64( _
    array2D() As LongLong, _
    zero_based_row_index As Long, _
    zero_based_column_index As Long, _
    new_value As LongLong)
    
    Dim i As Long
    Dim j As Long
    i = LBound(array2D, 1) + zero_based_row_index
    j = LBound(array2D, 2) + zero_based_column_index
    
    array2D(i, j) = new_value
End Sub


Public Sub Set_item_of_array2D_of_float32( _
    array2D() As Single, _
    zero_based_row_index As Long, _
    zero_based_column_index As Long, _
    new_value As Single)
    
    Dim i As Long
    Dim j As Long
    i = LBound(array2D, 1) + zero_based_row_index
    j = LBound(array2D, 2) + zero_based_column_index
    
    array2D(i, j) = new_value
End Sub



Public Sub Set_item_of_array2D_of_float64( _
    array2D() As Double, _
    zero_based_row_index As Long, _
    zero_based_column_index As Long, _
    new_value As Double)
    
    Dim i As Long
    Dim j As Long
    i = LBound(array2D, 1) + zero_based_row_index
    j = LBound(array2D, 2) + zero_based_column_index
    
    array2D(i, j) = new_value
End Sub



Public Sub Set_item_of_array2D_of_string( _
    array2D() As String, _
    zero_based_row_index As Long, _
    zero_based_column_index As Long, _
    new_value As String)
    
    Dim i As Long
    Dim j As Long
    i = LBound(array2D, 1) + zero_based_row_index
    j = LBound(array2D, 2) + zero_based_column_index
    
    array2D(i, j) = new_value
End Sub



Public Sub Set_item_of_array2D_of_date( _
    array2D() As Date, _
    zero_based_row_index As Long, _
    zero_based_column_index As Long, _
    new_value As Date)
    
    Dim i As Long
    Dim j As Long
    i = LBound(array2D, 1) + zero_based_row_index
    j = LBound(array2D, 2) + zero_based_column_index
    
    array2D(i, j) = new_value
End Sub


Public Sub Set_item_of_array2D_of_objects( _
    array2D() As Object, _
    zero_based_row_index As Long, _
    zero_based_column_index As Long, _
    new_value As Object)
    
    Dim i As Long
    Dim j As Long
    i = LBound(array2D, 1) + zero_based_row_index
    j = LBound(array2D, 2) + zero_based_column_index
    
    Set array2D(i, j) = new_value
End Sub



Public Sub Set_item_of_array2D_of_bool( _
    array2D() As Boolean, _
    zero_based_row_index As Long, _
    zero_based_column_index As Long, _
    new_value As Boolean)
    
    Dim i As Long
    Dim j As Long
    i = LBound(array2D, 1) + zero_based_row_index
    j = LBound(array2D, 2) + zero_based_column_index
    
    array2D(i, j) = new_value
End Sub



Public Function Repeat_as_array1D( _
    val As Variant, _
    Count As Long, _
    Optional first_index As Long = 0) As Variant
    
    Debug.Assert Count > 0
    
    Dim result As Variant: ReDim result(first_index To first_index + Count - 1)
    Dim i As Long
    For i = first_index To first_index + Count - 1
        result(i) = val
    Next i
    
    Repeat_as_array1D = result
End Function


Public Function Repeat_as_array2D( _
    val As Variant, _
    count_rows As Long, _
    count_columns As Long, _
    Optional first_index As Long = 0) As Variant
    
    Debug.Assert count_rows > 0
    Debug.Assert count_columns > 0
    
    Dim r1 As Long: r1 = first_index
    Dim r2 As Long: r2 = first_index + count_rows - 1
    Dim c1 As Long: c1 = first_index
    Dim c2 As Long: c2 = first_index + count_columns - 1
    
    Dim result As Variant: ReDim result(r1 To r2, c1 To c2)
    Dim i As Long
    Dim j As Long
    For i = r1 To r2
        For j = c1 To c2
            result(i, j) = val
        Next j
    Next i
    
    Repeat_as_array2D = result
End Function


Public Function Transform_array1D_to_array2D(src_arr1D As Variant) As Variant
    Debug.Assert IsArray(src_arr1D)
    Debug.Assert ARRAY_HELPER.Get_array_rank(src_arr1D) = 1

    Dim r1 As Long: r1 = LBound(src_arr1D)
    Dim r2 As Long: r2 = UBound(src_arr1D)
    Dim c1 As Long: c1 = r1
    Dim c2 As Long: c2 = c1
    
    Dim result As Variant: ReDim result(r1 To r2, c1 To c2)
    Dim i As Long
    For i = r1 To r2
        result(i, c1) = src_arr1D(i)
    Next i
    
    Transform_array1D_to_array2D = result
End Function


Public Function Transpose_array2D(src_arr2D As Variant) As Variant
    Debug.Assert IsArray(src_arr2D)
    Debug.Assert ARRAY_HELPER.Get_array_rank(src_arr2D) = 2
    
    Dim r1 As Long: r1 = LBound(src_arr2D, 1)
    Dim r2 As Long: r2 = UBound(src_arr2D, 1)
    Dim c1 As Long: c1 = LBound(src_arr2D, 2)
    Dim c2 As Long: c2 = UBound(src_arr2D, 2)
    
    Dim result As Variant: ReDim result(c1 To c2, r1 To r2)
    
    Dim i As Long
    Dim j As Long
    For i = r1 To r2
        For j = c1 To c2
            result(j, i) = src_arr2D(i, j)
        Next j
    Next i
    
    Transpose_array2D = result
End Function


Public Function Transform_array2D_to_array1D( _
    src_arr As Variant, _
    Optional first_index As Long = 0) As Variant
    Debug.Assert IsArray(src_arr)
    Debug.Assert ARRAY_HELPER.Get_array_rank(src_arr) = 2
    
    Dim src_r1 As Long: src_r1 = LBound(src_arr, 1)
    Dim src_r2 As Long: src_r2 = UBound(src_arr, 1)
    Dim src_c1 As Long: src_c1 = LBound(src_arr, 2)
    Dim src_c2 As Long: src_c2 = UBound(src_arr, 2)
    
    Dim dst_size As Long: dst_size = (src_r2 - src_r1 + 1) * (src_c2 - src_c1 + 1)
    Dim result As Variant: ReDim result(first_index To first_index + dst_size - 1)

    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    k = first_index
    For i = src_r1 To src_r2
        For j = src_c1 To src_c2
            result(k) = src_arr(i, j)
            k = k + 1
        Next j
    Next i
    
    Transform_array2D_to_array1D = result
End Function




Public Function Transform_array2D_of_int32_to_array1D( _
    src_arr() As Long, _
    Optional first_index As Long = 0) As Long()
    
    Debug.Assert IsArray(src_arr)
    Debug.Assert ARRAY_HELPER.Get_array_rank(src_arr) = 2
    
    Dim src_r1 As Long: src_r1 = LBound(src_arr, 1)
    Dim src_r2 As Long: src_r2 = UBound(src_arr, 1)
    Dim src_c1 As Long: src_c1 = LBound(src_arr, 2)
    Dim src_c2 As Long: src_c2 = UBound(src_arr, 2)
    
    Dim dst_size As Long: dst_size = (src_r2 - src_r1 + 1) * (src_c2 - src_c1 + 1)
    Dim result() As Long: ReDim result(first_index To first_index + dst_size - 1)

    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    k = first_index
    For i = src_r1 To src_r2
        For j = src_c1 To src_c2
            result(k) = src_arr(i, j)
            k = k + 1
        Next j
    Next i
    
    Transform_array2D_of_int32_to_array1D = result
End Function



Public Function Transform_array2D_of_float64_to_array1D( _
    src_arr() As Double, _
    Optional first_index As Long = 0) As Double()
    
    Debug.Assert IsArray(src_arr)
    Debug.Assert ARRAY_HELPER.Get_array_rank(src_arr) = 2
    
    Dim src_r1 As Long: src_r1 = LBound(src_arr, 1)
    Dim src_r2 As Long: src_r2 = UBound(src_arr, 1)
    Dim src_c1 As Long: src_c1 = LBound(src_arr, 2)
    Dim src_c2 As Long: src_c2 = UBound(src_arr, 2)
    
    Dim dst_size As Long: dst_size = (src_r2 - src_r1 + 1) * (src_c2 - src_c1 + 1)
    Dim result() As Double: ReDim result(first_index To first_index + dst_size - 1)

    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    k = first_index
    For i = src_r1 To src_r2
        For j = src_c1 To src_c2
            result(k) = src_arr(i, j)
            k = k + 1
        Next j
    Next i
    
    Transform_array2D_of_float64_to_array1D = result
End Function



Public Function Transform_array2D_of_string_to_array1D( _
    src_arr() As String, _
    Optional first_index As Long = 0) As String()
    
    Debug.Assert IsArray(src_arr)
    Debug.Assert ARRAY_HELPER.Get_array_rank(src_arr) = 2
    
    Dim src_r1 As Long: src_r1 = LBound(src_arr, 1)
    Dim src_r2 As Long: src_r2 = UBound(src_arr, 1)
    Dim src_c1 As Long: src_c1 = LBound(src_arr, 2)
    Dim src_c2 As Long: src_c2 = UBound(src_arr, 2)
    
    Dim dst_size As Long: dst_size = (src_r2 - src_r1 + 1) * (src_c2 - src_c1 + 1)
    Dim result() As String: ReDim result(first_index To first_index + dst_size - 1)

    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    k = first_index
    For i = src_r1 To src_r2
        For j = src_c1 To src_c2
            result(k) = src_arr(i, j)
            k = k + 1
        Next j
    Next i
    
    Transform_array2D_of_string_to_array1D = result
End Function



Public Function Transform_array2D_of_bool_to_array1D( _
    src_arr() As Boolean, _
    Optional first_index As Long = 0) As Boolean()
    
    Debug.Assert IsArray(src_arr)
    Debug.Assert ARRAY_HELPER.Get_array_rank(src_arr) = 2
    
    Dim src_r1 As Long: src_r1 = LBound(src_arr, 1)
    Dim src_r2 As Long: src_r2 = UBound(src_arr, 1)
    Dim src_c1 As Long: src_c1 = LBound(src_arr, 2)
    Dim src_c2 As Long: src_c2 = UBound(src_arr, 2)
    
    Dim dst_size As Long: dst_size = (src_r2 - src_r1 + 1) * (src_c2 - src_c1 + 1)
    Dim result() As Boolean: ReDim result(first_index To first_index + dst_size - 1)

    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    k = first_index
    For i = src_r1 To src_r2
        For j = src_c1 To src_c2
            result(k) = src_arr(i, j)
            k = k + 1
        Next j
    Next i
    
    Transform_array2D_of_bool_to_array1D = result
End Function


Public Function Get_array_rank(arr As Variant) As Long
    If IsArray(arr) = False Then
        Get_array_rank = 0
    Else
        Dim i As Long: i = 0
        Dim n As Long
        On Error Resume Next
            Do While (Err = 0)
                i = i + 1
                n = UBound(arr, i)
            Loop
        On Error GoTo 0
        Get_array_rank = i - 1
    End If
End Function




Private Function Zip_1D_arrays(single_dimension_arrays As Variant)
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
        Debug.Assert Get_array_rank(arr) = 1
        
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



Public Function Get_min_lbound(bound_rank As Long, arrays As Variant) As Long
    Debug.Assert bound_rank > 0
    Debug.Assert UBound(arrays) - LBound(arrays) >= 0
    
    Dim min_lbound As Long: min_lbound = LBound(arrays(LBound(arrays)), bound_rank)
    Dim i As Long
    For i = LBound(arrays) + 1 To UBound(arrays)
        Dim current_lbound As Long
        current_lbound = LBound(arrays(i), bound_rank)
        min_lbound = MATH_HELPER.Min_i32(min_lbound, current_lbound)
    Next i
    
    Get_min_lbound = min_lbound
End Function

Public Function Get_max_lbound(bound_rank As Long, arrays As Variant) As Long
    Debug.Assert bound_rank > 0
    Debug.Assert UBound(arrays) - LBound(arrays) >= 0
    
    Dim max_lbound As Long: max_lbound = LBound(arrays(LBound(arrays)), bound_rank)
    Dim i As Long
    For i = LBound(arrays) + 1 To UBound(arrays)
        Dim current_lbound As Long
        current_lbound = LBound(arrays(i), bound_rank)
        max_lbound = MATH_HELPER.Max_i32(max_lbound, current_lbound)
    Next i
    
    Get_max_lbound = max_lbound
End Function


Public Function Get_max_array_size(bound_rank As Long, arrays As Variant) As Long
    Debug.Assert bound_rank > 0
    Debug.Assert UBound(arrays) - LBound(arrays) >= 0
    
    Dim max_size As Long: max_size = UBound(arrays(LBound(arrays)), bound_rank) - LBound(arrays(LBound(arrays)), bound_rank) + 1
    Dim i As Long
    For i = LBound(arrays) + 1 To UBound(arrays)
        Dim current_size As Long
        current_size = UBound(arrays(i), bound_rank) - LBound(arrays(i), bound_rank) + 1
        max_size = MATH_HELPER.Max_i32(max_size, current_size)
    Next i
    
    Get_max_array_size = max_size
End Function


Public Function Get_sum_array_size(bound_rank As Long, arrays As Variant) As Long
    Debug.Assert bound_rank > 0
    Debug.Assert UBound(arrays) - LBound(arrays) >= 0
    
    Dim sum_sizes As Long: sum_sizes = 0
    Dim i As Long
    For i = LBound(arrays) To UBound(arrays)
        Dim current_size As Long
        current_size = Get_array_size(arrays(i), bound_rank)
        sum_sizes = sum_sizes + current_size
    Next i
    
    Get_sum_array_size = sum_sizes
End Function


Public Function Get_array_size(arr As Variant, Optional rnk As Long = 1)
    Debug.Assert rnk > 0
    Get_array_size = UBound(arr, rnk) - LBound(arr, rnk) + 1
End Function


Public Function Is_array_of_arrays(subj As Variant) As Boolean
    If IsArray(subj) = False Then
        Is_array_of_arrays = False
        Exit Function
    End If
    
    Dim v As Variant
    v = Get_first_element_of_any_array(subj)
    Is_array_of_arrays = IsArray(v)
End Function


Public Function Get_first_element_of_any_array(subj As Variant) As Variant
    If IsArray(subj) = False Then
        Get_first_element_of_any_array = subj
        Exit Function
    End If
    
    Dim rank As Long: rank = Get_array_rank(subj)
    
    Debug.Assert rank > 0
    Debug.Assert rank < 8
    
    Dim i As Long, j As Long, k As Long, l As Long, m As Long, n As Long, o As Long
    If rank = 1 Then
        i = LBound(subj, 1)
        Get_first_element_of_any_array = subj(i)
    ElseIf rank = 2 Then
        i = LBound(subj, 1)
        j = LBound(subj, 2)
        Get_first_element_of_any_array = subj(i, j)
    ElseIf rank = 3 Then
        i = LBound(subj, 1)
        j = LBound(subj, 2)
        k = LBound(subj, 3)
        Get_first_element_of_any_array = subj(i, j, k)
    ElseIf rank = 4 Then
        i = LBound(subj, 1)
        j = LBound(subj, 2)
        k = LBound(subj, 3)
        l = LBound(subj, 4)
        Get_first_element_of_any_array = subj(i, j, k, l)
    ElseIf rank = 5 Then
        i = LBound(subj, 1)
        j = LBound(subj, 2)
        k = LBound(subj, 3)
        l = LBound(subj, 4)
        m = LBound(subj, 5)
        Get_first_element_of_any_array = subj(i, j, k, l, m)
    ElseIf rank = 6 Then
        i = LBound(subj, 1)
        j = LBound(subj, 2)
        k = LBound(subj, 3)
        l = LBound(subj, 4)
        m = LBound(subj, 5)
        n = LBound(subj, 6)
        Get_first_element_of_any_array = subj(i, j, k, l, m, n)
    ElseIf rank = 7 Then
        i = LBound(subj, 1)
        j = LBound(subj, 2)
        k = LBound(subj, 3)
        l = LBound(subj, 4)
        m = LBound(subj, 5)
        n = LBound(subj, 6)
        o = LBound(subj, 7)
        Get_first_element_of_any_array = subj(i, j, k, l, m, n, o)
    End If
    
End Function



Public Function Zip_arrays_as_columns( _
    ParamArray array_of_arrays() As Variant _
    ) As Variant
    
    Dim size As Long: size = Get_array_size(CVar(array_of_arrays))
    Dim rank As Long: rank = Get_array_rank(CVar(array_of_arrays))

    If size = 1 And rank = 1 And Is_array_of_arrays(CVar(array_of_arrays)) Then
        Dim new_array_of_arrays As Variant
        new_array_of_arrays = ARRAY_HELPER.Get_item_of_array1D(CVar(array_of_arrays), 0)
        Zip_arrays_as_columns = Zip_arrays_as_columns_internal(new_array_of_arrays)
    Else
        Zip_arrays_as_columns = Zip_arrays_as_columns_internal(CVar(array_of_arrays))
    End If
    
End Function



Private Function Zip_arrays_as_columns_internal(array_of_arrays As Variant) As Variant
    
    Dim na As Long: na = Get_array_size(array_of_arrays)
    Dim new_array_of_arrays As Variant: ReDim new_array_of_arrays(0 To na - 1)
    
    Dim i As Long
    For i = 0 To na - 1
        Debug.Assert IsArray(array_of_arrays(i))
        Dim r As Long: r = Get_array_rank(array_of_arrays(i))
        Debug.Assert r < 3
        If r = 1 Then
            new_array_of_arrays(i) = ARRAY_HELPER.Transform_array1D_to_array2D(array_of_arrays(i))
        Else
            new_array_of_arrays(i) = array_of_arrays(i)
        End If
    Next i
    
    Zip_arrays_as_columns_internal = Zip_2D_arrays_as_columns(new_array_of_arrays)
End Function



Private Function Zip_2D_arrays_as_columns(arrays As Variant)
    Debug.Assert IsArray(arrays)

    Dim result_r1 As Long: result_r1 = Get_min_lbound(1, arrays)
    Dim result_r2 As Long: result_r2 = result_r1 + Get_max_array_size(1, arrays) - 1
    Dim result_nr As Long: result_nr = result_r2 - result_r1 + 1
    
    Dim result_c1 As Long: result_c1 = Get_min_lbound(2, arrays)
    Dim result_c2 As Long: result_c2 = result_c1 + Get_sum_array_size(2, arrays) - 1
    Dim result_nc As Long: result_nc = result_c2 - result_c1 + 1
    
    Dim result As Variant: ReDim result(result_r1 To result_r2, result_c1 To result_c2)
    
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim result_j As Long
    result_j = 0
    
    For k = LBound(arrays) To UBound(arrays)
        For i = 0 To result_nr - 1
            Dim nc As Long: nc = ARRAY_HELPER.Get_array_size(arrays(k), 2)
            For j = 0 To nc - 1
                Dim v As Variant
                v = Get_item_of_array2D(arrays(k), i, j)
                Set_item_of_array2D result, i, result_j + j, v
            Next j
        Next i
        
        result_j = result_j + nc
     Next k

    Zip_2D_arrays_as_columns = result
End Function



Public Function Change_lbound_of_array1D_f64( _
    source_array() As Double, _
    Optional new_lbound As Long = 0) As Double()
    
    'Dim li As Long: li = LBound(source_array)
    'Dim ui As Long: ui = UBound(source_array)
    
    Dim n As Long: n = ARRAY_HELPER.Get_array_size(source_array, 1) 'N = ui - li + 1
    
    Dim dst_arr() As Double: ReDim dst_arr(new_lbound To new_lbound + n - 1)
    
    Dim i As Long
    For i = new_lbound To new_lbound + n - 1
        dst_arr(i) = ARRAY_HELPER.Get_item_of_array1D(source_array, i - new_lbound)
    Next i
    
    Change_lbound_of_array1D_f64 = dst_arr
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


Public Function Convert_from_collection_to_array1D( _
    coll As Collection, _
    Optional first_index As Long = 0)
    
    Debug.Assert coll.Count > 0

    Dim n1 As Long: n1 = first_index
    Dim n2 As Long: n2 = first_index + coll.Count - 1
    Dim result As Variant: ReDim result(n1 To n2)
    
    Dim el As Variant
    Dim i As Long
    
    i = n1
    For Each el In coll
         result(i) = el
         i = i + 1
    Next el
    
    Convert_from_collection_to_array1D = result
End Function





Public Function Distinct_array2D(src_arr As Variant) As Variant
    Dim nr1 As Long: nr1 = LBound(src_arr, 1)
    Dim nr2 As Long: nr2 = UBound(src_arr, 1)
    Dim nc1 As Long: nc1 = LBound(src_arr, 2)
    Dim nc2 As Long: nc2 = UBound(src_arr, 2)

    Dim bLocalDif As Boolean
    Dim bFound As Boolean
    
    Dim ans As Variant: ReDim ans(nc1 To nc2, 1 To 1)
    
    Dim k As Long
    For k = nc1 To nc2
        ans(k, 1) = src_arr(nr1, k)
    Next k
    
    Dim Count As Long: Count = 1

    Dim i As Long
    Dim j As Long
    For i = nr1 + 1 To nr2
    
        bFound = False
        
        For j = 1 To UBound(ans, 2)
            
            bLocalDif = False
            
            For k = nc1 To nc2
                If Not src_arr(i, k) = ans(k, j) Then
                    bLocalDif = True
                End If
            Next k
            
            If bLocalDif = False Then
                bFound = True
                Exit For
            End If
            
        Next j
    
        If bFound = False Then
            Count = Count + 1
            ReDim Preserve ans(nc1 To nc2, 1 To Count)
            For k = nc1 To nc2
                ans(k, Count) = src_arr(i, k)
            Next k
            
        End If
    
    Next i
    
    Distinct_array2D = ARRAY_HELPER.Transpose_array2D(ans)
End Function




Public Function Is_arrays_equals(arr1 As Variant, arr2 As Variant) As Boolean
    Dim rank1 As Long: rank1 = ARRAY_HELPER.Get_array_rank(arr1)
    Dim rank2 As Long: rank2 = ARRAY_HELPER.Get_array_rank(arr2)
    
    If Not rank1 = rank2 Then
        Is_arrays_equals = False
        Exit Function
    End If
    
    
    
    Debug.Assert rank1 > 0
    Debug.Assert rank1 < 3
    
    Dim i1 As Long, j1 As Long, k1 As Long
    Dim i2 As Long, j2 As Long, k2 As Long
    Dim n1 As Long, n2 As Long
    Dim m1 As Long, m2 As Long
    
    If rank1 = 1 Then
        n1 = ARRAY_HELPER.Get_array_size(arr1)
        n2 = ARRAY_HELPER.Get_array_size(arr2)
    
        If Not n1 = n2 Then
            Is_arrays_equals = False
            Exit Function
        End If
    
        For i1 = LBound(arr1) To UBound(arr2)
            i2 = i1 - LBound(arr1) + LBound(arr2)
        
            If Not arr1(i1) = arr2(i2) Then
                Is_arrays_equals = False
                Exit Function
            End If
        Next i1
    
    ElseIf rank1 = 2 Then
        n1 = ARRAY_HELPER.Get_array_size(arr1, 1)
        n2 = ARRAY_HELPER.Get_array_size(arr2, 1)
        m1 = ARRAY_HELPER.Get_array_size(arr1, 2)
        m2 = ARRAY_HELPER.Get_array_size(arr2, 2)
    
        If Not n1 = n2 Then
            Is_arrays_equals = False
            Exit Function
        End If
        
        If Not m1 = m2 Then
            Is_arrays_equals = False
            Exit Function
        End If
    
        For i1 = LBound(arr1, 1) To UBound(arr2, 1)
            i2 = i1 - LBound(arr1, 1) + LBound(arr2, 1)
        
            For j1 = LBound(arr1, 2) To UBound(arr2, 2)
                j2 = j1 - LBound(arr1, 2) + LBound(arr2, 2)

                If Not arr1(i1, j1) = arr2(i2, j1) Then
                    Is_arrays_equals = False
                    Exit Function
                End If
            Next j1
        Next i1
    
    End If
    
    
    Is_arrays_equals = True
    
End Function


