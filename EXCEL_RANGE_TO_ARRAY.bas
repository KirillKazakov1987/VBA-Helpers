Attribute VB_Name = "EXCEL_RANGE_TO_ARRAY"
Option Explicit





Public Function Read_array1D_from_worksheet( _
    source_worksheet As Worksheet, _
    box As RANGE_BOX.RANGE_BOX, _
    Optional trim_on_used_range As Boolean = True, _
    Optional first_index As Long = 0) As Variant
    
    Dim arr2D As Variant: arr2D = Read_array2D_from_worksheet(source_worksheet, box, trim_on_used_range, first_index)
    
    Dim arr1D As Variant: arr1D = ARRAY_HELPER.Transform_array2D_to_array1D(arr2D)
    
    Read_array1D_from_worksheet = arr1D
End Function


Public Function Read_array2D_from_worksheet( _
    source_worksheet As Worksheet, _
    box As RANGE_BOX.RANGE_BOX, _
    Optional trim_on_used_range As Boolean = True, _
    Optional first_index As Long = 0) As Variant
    
    Dim rng As Range
    Set rng = RANGE_BOX.Get_excel_range_from_range_box(source_worksheet, box, trim_on_used_range)
    
    Read_array2D_from_worksheet = Read_array2D_from_range(rng, False, first_index)
    
End Function


Public Function Read_array1D_from_range( _
    rng As Range, _
    Optional trim_on_used_range As Boolean = True, _
    Optional first_index As Long = 0) As Variant
    
    Dim result As Variant
    result = Read_array2D_from_range(rng, trim_on_used_range, first_index)
    
    Read_array1D_from_range = ARRAY_HELPER.Transform_array2D_to_array1D(result)
End Function



Public Function Read_array1D_of_int32_from_range( _
    rng As Range, _
    Optional trim_on_used_range As Boolean = True, _
    Optional first_index As Long = 0, _
    Optional substitute_for_empty As Long = 0, _
    Optional substitute_for_uncovertable As Long = 0, _
    Optional ByRef counted_empty_values As Long, _
    Optional ByRef counted_unconvertable_values As Long) As Long()
    
    Dim result() As Long
    
    result = Read_array2D_of_int32_from_range( _
        rng, _
        trim_on_used_range, _
        first_index, _
        substitute_for_empty, substitute_for_uncovertable, _
        counted_empty_values, counted_unconvertable_values)
    
    Read_array1D_of_int32_from_range = ARRAY_HELPER.Transform_array2D_of_int32_to_array1D(result)
End Function



Public Function Read_array1D_of_float64_from_range( _
    rng As Range, _
    Optional trim_on_used_range As Boolean = True, _
    Optional first_index As Long = 0, _
    Optional substitute_for_empty As Double = 0, _
    Optional substitute_for_uncovertable As Double = 0, _
    Optional ByRef counted_empty_values As Long, _
    Optional ByRef counted_unconvertable_values As Long) As Double()
    
    Dim result() As Double
    
    result = Read_array2D_of_float64_from_range( _
        rng, _
        trim_on_used_range, _
        first_index, _
        substitute_for_empty, substitute_for_uncovertable, _
        counted_empty_values, counted_unconvertable_values)
    
    Read_array1D_of_float64_from_range = ARRAY_HELPER.Transform_array2D_of_float64_to_array1D(result)
End Function



Public Function Read_array1D_of_bool_from_range( _
    rng As Range, _
    Optional trim_on_used_range As Boolean = True, _
    Optional first_index As Long = 0, _
    Optional substitute_for_empty As Boolean = 0, _
    Optional substitute_for_uncovertable As Boolean = 0, _
    Optional ByRef counted_empty_values As Long, _
    Optional ByRef counted_unconvertable_values As Long) As Boolean()
    
    Dim result() As Boolean
    
    result = Read_array2D_of_bool_from_range( _
        rng, _
        trim_on_used_range, _
        first_index, _
        substitute_for_empty, substitute_for_uncovertable, _
        counted_empty_values, counted_unconvertable_values)
    
    Read_array1D_of_bool_from_range = ARRAY_HELPER.Transform_array2D_of_bool_to_array1D(result)
End Function



Public Function Read_array1D_of_string_from_range( _
    rng As Range, _
    Optional trim_on_used_range As Boolean = True, _
    Optional first_index As Long = 0, _
    Optional substitute_for_empty As String = 0, _
    Optional ByRef counted_empty_values As Long) As String()
    
    Dim result() As String
    
    result = Read_array2D_of_string_from_range( _
        rng, _
        trim_on_used_range, _
        first_index, _
        substitute_for_empty, _
        counted_empty_values)
    
    Read_array1D_of_string_from_range = ARRAY_HELPER.Transform_array2D_of_string_to_array1D(result)
End Function



Public Function Read_array2D_from_range( _
    rng As Range, _
    Optional trim_on_used_range As Boolean = True, _
    Optional first_index As Long = 0) As Variant
    
    If trim_on_used_range Then
        Set rng = Excel.Application.Intersect(rng, rng.Worksheet.UsedRange)
    End If

    Dim nr As Long: nr = rng.Rows.Count
    Dim nc As Long: nc = rng.Columns.Count
    
    Dim arr As Variant
    If nr * nc = 0 Then
        Read_array2D_from_range = Empty
    ElseIf nr * nc = 1 Then
        ReDim arr(first_index To first_index, first_index To first_index)
        arr(first_index, first_index) = rng.Value
        Read_array2D_from_range = arr
    Else
        Dim raw_arr As Variant: raw_arr = rng.Value
        
        ReDim arr(first_index To first_index + nr - 1, first_index To first_index + nc - 1)
    
        Dim i As Long
        Dim j As Long
        For i = 0 To nr - 1
            For j = 0 To nc - 1
                Dim v As Variant: v = ARRAY_HELPER.Get_item_of_array2D(raw_arr, i, j)
                ARRAY_HELPER.Set_item_of_array2D arr, i, j, v
            Next j
        Next i
        
        Read_array2D_from_range = arr
    End If

End Function




Public Function Read_array2D_of_int32_from_range( _
    rng As Range, _
    Optional trim_on_used_range As Boolean = True, _
    Optional first_index As Long = 0, _
    Optional substitute_for_empty As Long = 0, _
    Optional substitute_for_uncovertable As Long = 0, _
    Optional ByRef counted_empty_values As Long, _
    Optional ByRef counted_unconvertable_values As Long) As Long()
    
    counted_empty_values = 0
    counted_unconvertable_values = 0
    
    If trim_on_used_range Then
        Set rng = Excel.Application.Intersect(rng, rng.Worksheet.UsedRange)
    End If

    Dim nr As Long: nr = rng.Rows.Count
    Dim nc As Long: nc = rng.Columns.Count
    
    Dim arr() As Long
    Dim opi32 As Optional_int32
    
    If nr * nc = 0 Then
        Read_array2D_of_int32_from_range = Empty
    ElseIf nr * nc = 1 Then
        ReDim arr(first_index To first_index, first_index To first_index)
        
        If IsEmpty(rng.Value) Then
            counted_empty_values = counted_empty_values + 1
            arr(first_index, first_index) = substitute_for_empty
        Else
            opi32 = TYPE_HELPER.Try_convert_to_int32(rng.Value)
            
            If opi32.Has_value Then
                arr(first_index, first_index) = opi32.Value
            Else
                counted_unconvertable_values = counted_unconvertable_values + 1
                arr(first_index, first_index) = substitute_for_uncovertable
            End If
        End If
        
        Read_array2D_of_int32_from_range = arr
    Else
        Dim raw_arr As Variant: raw_arr = rng.Value
        
        ReDim arr(first_index To first_index + nr - 1, first_index To first_index + nc - 1)
    
        Dim i As Long
        Dim j As Long
        For i = 0 To nr - 1
            For j = 0 To nc - 1
                Dim v As Variant: v = ARRAY_HELPER.Get_item_of_array2D(raw_arr, i, j)
                
                If IsEmpty(v) Then
                    counted_empty_values = counted_empty_values + 1
                    ARRAY_HELPER.Set_item_of_array2D_of_int32 arr, i, j, substitute_for_empty
                Else
                    opi32 = TYPE_HELPER.Try_convert_to_int32(v)
                    
                    If opi32.Has_value Then
                        ARRAY_HELPER.Set_item_of_array2D_of_int32 arr, i, j, opi32.Value
                    Else
                        counted_unconvertable_values = counted_unconvertable_values + 1
                        ARRAY_HELPER.Set_item_of_array2D_of_int32 arr, i, j, substitute_for_uncovertable
                    End If
                End If
                
            Next j
        Next i
        
        Read_array2D_of_int32_from_range = arr
    End If

End Function




Public Function Read_array2D_of_float64_from_range( _
    rng As Range, _
    Optional trim_on_used_range As Boolean = True, _
    Optional first_index As Long = 0, _
    Optional substitute_for_empty As Double = 0, _
    Optional substitute_for_uncovertable As Double = 0, _
    Optional ByRef counted_empty_values As Long, _
    Optional ByRef counted_unconvertable_values As Long) As Double()
    
    counted_empty_values = 0
    counted_unconvertable_values = 0
    
    If trim_on_used_range Then
        Set rng = Excel.Application.Intersect(rng, rng.Worksheet.UsedRange)
    End If

    Dim nr As Long: nr = rng.Rows.Count
    Dim nc As Long: nc = rng.Columns.Count
    
    Dim arr() As Double
    Dim opf64 As Optional_float64
    
    If nr * nc = 0 Then
        Read_array2D_of_float64_from_range = Empty
    ElseIf nr * nc = 1 Then
        ReDim arr(first_index To first_index, first_index To first_index)
        
        If IsEmpty(rng.Value) Then
            counted_empty_values = counted_empty_values + 1
            arr(first_index, first_index) = substitute_for_empty
        Else
            opf64 = TYPE_HELPER.Try_convert_to_float64(rng.Value)
            
            If opf64.Has_value Then
                arr(first_index, first_index) = opf64.Value
            Else
                counted_unconvertable_values = counted_unconvertable_values + 1
                arr(first_index, first_index) = substitute_for_uncovertable
            End If
        End If
        
        Read_array2D_of_float64_from_range = arr
    Else
        Dim raw_arr As Variant: raw_arr = rng.Value
        
        ReDim arr(first_index To first_index + nr - 1, first_index To first_index + nc - 1)
    
        Dim i As Long
        Dim j As Long
        For i = 0 To nr - 1
            For j = 0 To nc - 1
                Dim v As Variant: v = ARRAY_HELPER.Get_item_of_array2D(raw_arr, i, j)
                
                If IsEmpty(v) Then
                    counted_empty_values = counted_empty_values + 1
                    ARRAY_HELPER.Set_item_of_array2D_of_float64 arr, i, j, substitute_for_empty
                Else
                    opf64 = TYPE_HELPER.Try_convert_to_float64(v)
                    
                    If opf64.Has_value Then
                        ARRAY_HELPER.Set_item_of_array2D_of_float64 arr, i, j, opf64.Value
                    Else
                        counted_unconvertable_values = counted_unconvertable_values + 1
                        ARRAY_HELPER.Set_item_of_array2D_of_float64 arr, i, j, substitute_for_uncovertable
                    End If
                End If
                
            Next j
        Next i
        
        Read_array2D_of_float64_from_range = arr
    End If

End Function




Public Function Read_array2D_of_string_from_range( _
    rng As Range, _
    Optional trim_on_used_range As Boolean = True, _
    Optional first_index As Long = 0, _
    Optional substitute_for_empty As String = 0, _
    Optional ByRef counted_empty_values As Long) As String()
    
    counted_empty_values = 0
    
    If trim_on_used_range Then
        Set rng = Excel.Application.Intersect(rng, rng.Worksheet.UsedRange)
    End If

    Dim nr As Long: nr = rng.Rows.Count
    Dim nc As Long: nc = rng.Columns.Count
    
    Dim arr() As String

    If nr * nc = 0 Then
        Read_array2D_of_string_from_range = Empty
    ElseIf nr * nc = 1 Then
        ReDim arr(first_index To first_index, first_index To first_index)
        
        If IsEmpty(rng.Value) Then
            counted_empty_values = counted_empty_values + 1
            arr(first_index, first_index) = substitute_for_empty
        Else
            arr(first_index, first_index) = CStr(rng.Value)
        End If
        
        Read_array2D_of_string_from_range = arr
    Else
        Dim raw_arr As Variant: raw_arr = rng.Value
        
        ReDim arr(first_index To first_index + nr - 1, first_index To first_index + nc - 1)
    
        Dim i As Long
        Dim j As Long
        For i = 0 To nr - 1
            For j = 0 To nc - 1
            
                Dim v As Variant: v = ARRAY_HELPER.Get_item_of_array2D(raw_arr, i, j)
                
                If IsEmpty(v) Then
                    counted_empty_values = counted_empty_values + 1
                    ARRAY_HELPER.Set_item_of_array2D_of_string arr, i, j, substitute_for_empty
                Else
                    ARRAY_HELPER.Set_item_of_array2D_of_string arr, i, j, CStr(v)
                End If
                
            Next j
        Next i
        
        Read_array2D_of_string_from_range = arr
    End If

End Function




Public Function Read_array2D_of_bool_from_range( _
    rng As Range, _
    Optional trim_on_used_range As Boolean = True, _
    Optional first_index As Long = 0, _
    Optional substitute_for_empty As Boolean = 0, _
    Optional substitute_for_uncovertable As Boolean = 0, _
    Optional ByRef counted_empty_values As Long, _
    Optional ByRef counted_unconvertable_values As Long) As Boolean()
    
    counted_empty_values = 0
    counted_unconvertable_values = 0
    
    If trim_on_used_range Then
        Set rng = Excel.Application.Intersect(rng, rng.Worksheet.UsedRange)
    End If

    Dim nr As Long: nr = rng.Rows.Count
    Dim nc As Long: nc = rng.Columns.Count
    
    Dim arr() As Boolean
    Dim opb As Optional_bool
    
    If nr * nc = 0 Then
        Read_array2D_of_bool_from_range = Empty
    ElseIf nr * nc = 1 Then
        ReDim arr(first_index To first_index, first_index To first_index)
        
        If IsEmpty(rng.Value) Then
            counted_empty_values = counted_empty_values + 1
            arr(first_index, first_index) = substitute_for_empty
        Else
            opb = TYPE_HELPER.Try_convert_to_bool(rng.Value)
            
            If opb.Has_value Then
                arr(first_index, first_index) = opb.Value
            Else
                counted_unconvertable_values = counted_unconvertable_values + 1
                arr(first_index, first_index) = substitute_for_uncovertable
            End If
        End If
        
        Read_array2D_of_bool_from_range = arr
    Else
        Dim raw_arr As Variant: raw_arr = rng.Value
        
        ReDim arr(first_index To first_index + nr - 1, first_index To first_index + nc - 1)
    
        Dim i As Long
        Dim j As Long
        For i = 0 To nr - 1
            For j = 0 To nc - 1
                Dim v As Variant: v = ARRAY_HELPER.Get_item_of_array2D(raw_arr, i, j)
                
                If IsEmpty(v) Then
                    counted_empty_values = counted_empty_values + 1
                    ARRAY_HELPER.Set_item_of_array2D_of_bool arr, i, j, substitute_for_empty
                Else
                    opb = TYPE_HELPER.Try_convert_to_bool(v)
                    
                    If opb.Has_value Then
                        ARRAY_HELPER.Set_item_of_array2D_of_bool arr, i, j, opb.Value
                    Else
                        counted_unconvertable_values = counted_unconvertable_values + 1
                        ARRAY_HELPER.Set_item_of_array2D_of_bool arr, i, j, substitute_for_uncovertable
                    End If
                End If
                
            Next j
        Next i
        
        Read_array2D_of_bool_from_range = arr
    End If

End Function



Public Function Get_interior_colors_of_range_cells_as_array2D_of_int32( _
    rng As Range, _
    Optional trim_on_used_range As Boolean = True, _
    Optional first_index As Long = 0) As Long()
    
    If trim_on_used_range Then
        Set rng = Excel.Application.Intersect(rng, rng.Worksheet.UsedRange)
    End If

    Dim nr As Long: nr = rng.Rows.Count
    Dim nc As Long: nc = rng.Columns.Count
    
    Dim arr() As Long
    
    If nr * nc = 0 Then
        Get_interior_colors_of_range_cells_as_array2D_of_int32 = Empty
    ElseIf nr * nc = 1 Then
        ReDim arr(first_index To first_index, first_index To first_index)
        arr(first_index, first_index) = rng.Interior.Color
        Get_interior_colors_of_range_cells_as_array2D_of_int32 = arr
    Else
        
        ReDim arr(first_index To first_index + nr - 1, first_index To first_index + nc - 1)
    
        Dim i As Long
        Dim j As Long
        For i = 0 To nr - 1
            For j = 0 To nc - 1
                Dim cell_color As Long
                cell_color = rng((i - first_index + 1), (j - first_index + 1)).Interior.Color
                arr(first_index + i, first_index + j) = cell_color
            Next j
        Next i
        
        Get_interior_colors_of_range_cells_as_array2D_of_int32 = arr
    End If
End Function


Public Function Get_interior_colors_of_range_cells_as_array1D_of_int32( _
    rng As Range, _
    Optional trim_on_used_range As Boolean = True, _
    Optional first_index As Long = 0) As Long()
    
    
    Dim arr2D() As Long
    arr2D = Get_interior_colors_of_range_cells_as_array2D_of_int32(rng, trim_on_used_range, first_index)
    
    Dim arr1D() As Long
    arr1D = ARRAY_HELPER.Transform_array2D_of_int32_to_array1D(arr2D, first_index)
    
    Get_interior_colors_of_range_cells_as_array1D_of_int32 = arr1D
End Function
