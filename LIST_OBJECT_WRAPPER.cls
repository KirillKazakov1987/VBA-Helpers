VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LIST_OBJECT_WRAPPER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private lo As ListObject

Public Sub Init(list_object As ListObject)
    If IsNull(list_object) Then
        Err.Raise 999999, , "Try to initialize LIST_OBJECT_WRAPPER instance with argument list_object that have null value."
    End If
    
    Set lo = list_object
End Sub


Public Property Get Data_body_range()
    Set Data_body_range = lo.DataBodyRange
End Property



Public Sub Init_by_name_only(book As Workbook, list_object_name As String)
    Dim sh As Worksheet
    Dim llo As ListObject
    
    Dim list_object_name_lc As String
    list_object_name_lc = LCase(list_object_name)
    
    For Each sh In book.Worksheets
        For Each llo In sh.ListObjects
            If LCase(llo.Name) = list_object_name_lc Then
                Init llo
                Exit Sub
            End If
        Next llo
    Next sh
    
    Err.Raise 999999, , "Provided name of list object " + list_object_name + " is not found in the book " + book.Name + "!"
End Sub





Public Property Get Name()
    Name = lo.Name
End Property


Public Property Get count_rows()
    count_rows = lo.ListRows.Count
End Property




Public Property Get Real_cell_value(column_name As String, row_ondex As Long) As Double
    Dim x As Double: x = lo.ListColumns(column_name).Range(row_ondex + 1, 1).Value2
    Real_cell_value = x
End Property

Public Property Let Real_cell_value(column_name As String, row_ondex As Long, val As Double)
    lo.ListColumns(column_name).Range(row_ondex + 1, 1).Value2 = val
End Property


Public Property Get Int_cell_value(column_name As String, row_ondex As Long) As Long
    Dim x As Long: x = lo.ListColumns(column_name).Range(row_ondex + 1, 1).Value2
    Int_cell_value = x
End Property

Public Property Let Int_cell_value(column_name As String, row_ondex As Long, val As Long)
    lo.ListColumns(column_name).Range(row_ondex + 1, 1).Value2 = val
End Property


Public Property Get String_cell_value(column_name As String, row_ondex As Long) As String
    Dim x As String: x = lo.ListColumns(column_name).Range(row_ondex + 1, 1).Value2
    String_cell_value = x
End Property

Public Property Let String_cell_value(column_name As String, row_ondex As Long, val As String)
    lo.ListColumns(column_name).Range(row_ondex + 1, 1).Value2 = val
End Property



Public Property Get Bool_cell_value(column_name As String, row_ondex As Long) As Boolean
    Dim x As Long: x = lo.ListColumns(column_name).Range(row_ondex + 1, 1).Value2
    Bool_cell_value = IIf(x = 0, False, True)
End Property

Public Property Let Bool_cell_value(column_name As String, row_ondex As Long, val As Boolean)
    lo.ListColumns(column_name).Range(row_ondex + 1, 1).Value2 = IIf(val, 1, 0)
End Property



Public Property Get Is_row_empty(row_ondex As Long) As Boolean
    Dim c As Long
    c = WorksheetFunction.CountA(lo.ListRows.Item(row_ondex).Range)
    
    If c = 0 Then
        Is_row_empty = True
    Else
        Is_row_empty = False
    End If
End Property


Public Property Get Is_cell_empty(column_name As String, row_ondex As Long) As Boolean
    Is_cell_empty = IsEmpty(lo.ListColumns(column_name).Range(row_ondex + 1, 1).Value)
End Property


Public Sub Clear_column(column_name As String)
    lo.ListColumns(column_name).DataBodyRange.ClearContents
End Sub


Public Function Get_column_as_array1D(column_name As String) As Variant
    Debug.Assert Not lo.DataBodyRange Is Nothing
    
    Dim locol As ListColumn
    Set locol = lo.ListColumns.Item(column_name)
    
    Dim result As Variant
    result = EXCEL_RANGE_TO_ARRAY.Read_array1D_from_range(lo.DataBodyRange)
    
    Get_column_as_array1D = result
End Function




Public Function Get_interior_colors_of_column_cells_as_array1D_of_int32( _
    column_name As String, _
    Optional first_index As Long = 0) As Long()
    
    Debug.Assert Not lo.DataBodyRange Is Nothing
    Debug.Assert Not IsEmpty(column_name)
    Debug.Assert Not column_name = ""

    Dim locol As ListColumn
    Set locol = lo.ListColumns.Item(column_name)
    
    Dim result() As Long
    result = EXCEL_RANGE_TO_ARRAY.Get_interior_colors_of_range_cells_as_array1D_of_int32( _
        lo.DataBodyRange, _
        True, _
        first_index)

    Get_interior_colors_of_column_cells_as_array1D_of_int32 = result
End Function



Public Function Get_column_as_array1D_of_int32( _
    column_name As String, _
    Optional first_index As Long = 0, _
    Optional substitute_for_empty As Long = 0, _
    Optional substitute_for_uncovertable As Long = 0, _
    Optional ByRef counted_empty_values As Long, _
    Optional ByRef counted_unconvertable_values As Long) As Long()
    
    
    Debug.Assert Not lo.DataBodyRange Is Nothing
    
    Dim locol As ListColumn
    Set locol = lo.ListColumns.Item(column_name)
    
    Dim result() As Long
    result = EXCEL_RANGE_TO_ARRAY.Read_array1D_of_int32_from_range( _
        lo.DataBodyRange, _
        True, _
        first_index, _
        substitute_for_empty, substitute_for_uncovertable, _
        counted_empty_values, counted_unconvertable_values)

    Get_column_as_array1D_of_int32 = result
End Function


Public Function Get_column_as_array1D_of_float64( _
    column_name As String, _
    Optional first_index As Long = 0, _
    Optional substitute_for_empty As Double = 0, _
    Optional substitute_for_uncovertable As Double = 0, _
    Optional ByRef counted_empty_values As Long, _
    Optional ByRef counted_unconvertable_values As Long) As Double()
    
    
    Debug.Assert Not lo.DataBodyRange Is Nothing
    
    Dim locol As ListColumn
    Set locol = lo.ListColumns.Item(column_name)
    
    Dim result() As Double
    result = EXCEL_RANGE_TO_ARRAY.Read_array1D_of_float64_from_range( _
        lo.DataBodyRange, _
        True, _
        first_index, _
        substitute_for_empty, substitute_for_uncovertable, _
        counted_empty_values, counted_unconvertable_values)

    Get_column_as_array1D_of_float64 = result
End Function



Public Function Get_column_as_array1D_of_bool( _
    column_name As String, _
    Optional first_index As Long = 0, _
    Optional substitute_for_empty As Boolean = 0, _
    Optional substitute_for_uncovertable As Boolean = 0, _
    Optional ByRef counted_empty_values As Long, _
    Optional ByRef counted_unconvertable_values As Long) As Boolean()
    
    
    Debug.Assert Not lo.DataBodyRange Is Nothing
    
    Dim locol As ListColumn
    Set locol = lo.ListColumns.Item(column_name)
    
    Dim result() As Boolean
    result = EXCEL_RANGE_TO_ARRAY.Read_array1D_of_bool_from_range( _
        lo.DataBodyRange, _
        True, _
        first_index, _
        substitute_for_empty, substitute_for_uncovertable, _
        counted_empty_values, counted_unconvertable_values)

    Get_column_as_array1D_of_bool = result
End Function



Public Function Get_column_as_array1D_of_string( _
    column_name As String, _
    Optional first_index As Long = 0, _
    Optional substitute_for_empty As String = "", _
    Optional ByRef counted_empty_values As Long) As String()
    
    
    Debug.Assert Not lo.DataBodyRange Is Nothing
    
    Dim locol As ListColumn
    Set locol = lo.ListColumns.Item(column_name)
    
    Dim result() As String
    result = EXCEL_RANGE_TO_ARRAY.Read_array1D_of_string_from_range( _
        lo.DataBodyRange, _
        True, _
        first_index, _
        substitute_for_empty, _
        counted_empty_values)

    Get_column_as_array1D_of_string = result
End Function



Public Function Get_columns_as_array2D( _
    ParamArray column_names() As Variant _
    ) As Variant

    Debug.Assert Not lo.DataBodyRange Is Nothing
    
    Dim new_column_names As Variant
    If IsArray(ARRAY_HELPER.Get_first_element_of_any_array(CVar(column_names))) Then
        Debug.Assert ARRAY_HELPER.Get_array_size(CVar(column_names)) = 1
        Debug.Assert ARRAY_HELPER.Get_array_rank(CVar(column_names)) = 1
        new_column_names = ARRAY_HELPER.Get_first_element_of_any_array(CVar(column_names))
    Else
        new_column_names = CVar(column_names)
    End If
    
    Get_columns_as_array2D = Get_columns_as_array2D_internal(new_column_names)
    
End Function


Private Function Get_columns_as_array2D_internal( _
    column_names As Variant _
    ) As Variant
    
    Debug.Assert IsArray(column_names)
    
    Debug.Assert ARRAY_HELPER.Get_array_rank(column_names) = 1
    
    Dim n As Long: n = ARRAY_HELPER.Get_array_size(column_names)
    Debug.Assert n > 0
    
    Dim coll As New Collection
    
    Dim i As Long
    For i = LBound(column_names) To UBound(column_names)
        Dim column_name As Variant
        column_name = column_names(i)
        
        Dim locol As ListColumn
        Set locol = lo.ListColumns.Item(column_name)
        
        Dim column_content As Variant
        column_content = EXCEL_RANGE_TO_ARRAY.Read_array1D_from_range(locol.DataBodyRange)
        
        coll.Add column_content
    Next i
    
    Dim array_of_column_contents As Variant
    array_of_column_contents = ARRAY_HELPER.Convert_from_collection_to_array1D(coll)
    
    Dim result As Variant
    result = ARRAY_HELPER.Zip_arrays_as_columns(array_of_column_contents)

    Get_columns_as_array2D_internal = result
End Function



Public Sub Clear_content()
    If Not lo.DataBodyRange Is Nothing Then
        lo.DataBodyRange.ClearContents
    End If
End Sub


Public Sub Delete_all_content_rows()
    If Not lo.DataBodyRange Is Nothing Then
        lo.DataBodyRange.Rows.Delete
    End If
End Sub


Public Property Get Free_row_index()
    If lo.DataBodyRange Is Nothing Then
        Free_row_index = 0
    Else
        Free_row_index = lo.DataBodyRange.Rows.Count - 1
    End If
End Property


Public Sub Append_2D_array( _
    array2D As Variant, _
    Optional leftmost_column_name As String = "", _
    Optional uppermost_row_index As Long = -1)

    Debug.Assert IsArray(array2D)
    Debug.Assert ARRAY_HELPER.Get_array_rank(array2D) = 2

    If uppermost_row_index < 0 Then uppermost_row_index = Me.Free_row_index
    
    If leftmost_column_name = "" Then leftmost_column_name = lo.ListColumns.Item(1).Name
    
    Dim locol As ListColumn: Set locol = lo.ListColumns.Item(leftmost_column_name)
    
    Dim dst_upper_left_cell As Range: Set dst_upper_left_cell = locol.Range.Cells(uppermost_row_index + 2, 1)
    
    ARRAY_TO_EXCEL_RANGE.Write_2D_array_to_excel_range array2D, dst_upper_left_cell
End Sub




