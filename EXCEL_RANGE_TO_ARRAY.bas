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



Public Function Read_array2D_from_range( _
    rng As Range, _
    Optional trim_on_used_range As Boolean = True, _
    Optional first_index As Long = 0) As Variant
    
    If trim_on_used_range Then
        Set rng = Excel.Application.Intersect(rng, rng.Worksheet.UsedRange)
    End If

    Dim nr As Long: nr = rng.Rows.count
    Dim nc As Long: nc = rng.Columns.count
    
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


