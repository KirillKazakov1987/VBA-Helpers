Attribute VB_Name = "RANGE_BOX"
Option Explicit

Private Const ERR_INTERNAL_CODE = 999999

Private Const RANGE_ADDRESS_BLOCK_TYPE_UNDEFINED As Long = 0
Private Const RANGE_ADDRESS_BLOCK_TYPE_COLUMN As Long = 1
Private Const RANGE_ADDRESS_BLOCK_TYPE_ROW As Long = 2
Private Const RANGE_ADDRESS_BLOCK_TYPE_SEPARATOR As Long = 3

Private Const APLHABET As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Private Const APLHABET_LENGTH As Long = 26

Private Const EXCEL_WORKSHEET_MAX_ROWS As Long = 1048576
Private Const EXCEL_WORKSHEET_MAX_COLUMNS As Long = 16384

Public Type RANGE_BOX

    row_onedex As Long
    
    column_onedex As Long
    
    count_rows As Long
    
    count_columns As Long
    
End Type


Public Function Get_range_box_from_excel_range(rng As Excel.Range) As RANGE_BOX.RANGE_BOX
    
    Dim result As RANGE_BOX.RANGE_BOX
    
    With result
        result.row_onedex = rng.Row
        result.column_onedex = rng.Column
        result.count_rows = rng.Rows.Count
        result.count_columns = rng.Columns.Count
    End With
    
    Get_range_box_from_excel_range = result

End Function



Public Function Get_excel_range_from_range_box( _
    sh As Excel.Worksheet, _
    box As RANGE_BOX, _
    Optional trim_on_used_range As Boolean = True) As Range
    
    Dim result As Range: Set result = sh.Range(RANGE_BOX.Get_address_of_range_box(box))
    
    If trim_on_used_range Then
        Dim used_range As Range: Set used_range = sh.UsedRange
        Set result = Excel.Application.Intersect(result, used_range)
    End If
    
    Set Get_excel_range_from_range_box = result

End Function


Public Function Get_range_box_lower_and_upper_onedexes( _
    first_row_onedex As Long, _
    last_row_onedex As Long, _
    first_column_onedex As Long, _
    last_column_onedex As Long) As RANGE_BOX.RANGE_BOX
    
    Dim count_rows As Long: count_rows = last_row_onedex - first_row_onedex + 1
    Dim count_columns As Long: count_columns = last_column_onedex - first_column_onedex + 1
    
    Debug.Assert count_rows > 0
    Debug.Assert count_columns > 0
    Debug.Assert first_row_onedex > 0
    Debug.Assert first_column_onedex > 0
    
    Dim result As RANGE_BOX.RANGE_BOX
    With result
        .row_onedex = first_row_onedex
        .column_onedex = first_column_onedex
        .count_rows = count_rows
        .count_columns = count_columns
    End With
    
    Get_range_box_lower_and_upper_onedexes = result
End Function
    


Public Sub Deconstruct_to_lower_and_upper_inclusive_onedexes( _
    src As RANGE_BOX, _
    ByRef row1 As Long, _
    ByRef row2 As Long, _
    ByRef col1 As Long, _
    ByRef col2 As Long)
    
    row1 = src.row_onedex
    row2 = src.row_onedex + src.count_rows - 1
    col1 = src.column_onedex
    col2 = src.column_onedex + src.count_columns - 1
End Sub
    

Public Function Clamp_by_bounding_range_box(subj As RANGE_BOX, bounding_box As RANGE_BOX) As RANGE_BOX

    Dim r1 As Long, r2 As Long, c1 As Long, c2 As Long
    Deconstruct_to_lower_and_upper_inclusive_onedexes subj, r1, r2, c1, c2
    
    Dim lr1 As Long, lr2 As Long, lc1 As Long, lc2 As Long
    Deconstruct_to_lower_and_upper_inclusive_onedexes bounding_box, lr1, lr2, lc1, lc2
    
    
    r1 = MATH_HELPER.Clamp_i32(r1, lr1, lr2)
    r2 = MATH_HELPER.Clamp_i32(r2, lr1, lr2)
    c1 = MATH_HELPER.Clamp_i32(c1, lc1, lc2)
    c2 = MATH_HELPER.Clamp_i32(c2, lc1, lc2)

    Dim result As RANGE_BOX
    With result
        .row_onedex = r1
        .column_onedex = c1
        .count_rows = r2 - r1 + 1
        .count_columns = c2 - c1 + 1
    End With

    Clamp_by_bounding_range_box = result

End Function


Public Function Get_cell_address( _
    row_onedex As Long, _
    column_onedex As Long)
    
    If row_onedex < 0 Or column_onedex < 0 Then Err.Raise ERR_INTERNAL_CODE
    
    Dim row_number As Long: row_number = row_onedex
    Dim column_address As String: column_address = Unchecked_express_onedex_in_AZ(column_onedex)
    
    Get_cell_address = column_address & row_number
End Function


Public Function Get_range_address( _
    row_onedex As Long, _
    column_onedex As Long, _
    Optional count_rows As Long = 1, _
    Optional count_columns As Long = 1)
    
    If count_rows = 1 And count_columns = 1 Then
        Get_range_address = Get_cell_address(row_onedex, column_onedex)
    Else
        Dim upper_left_cell_address As String
        Dim lower_right_cell_address As String
    
        upper_left_cell_address = Get_cell_address(row_onedex, column_onedex)
        
        lower_right_cell_address = Get_cell_address(row_onedex + count_rows - 1, column_onedex + count_columns - 1)
    
        Get_range_address = upper_left_cell_address & ":" & lower_right_cell_address
    End If
End Function



Public Function Get_address_of_range_box(box As RANGE_BOX) As String
    Get_address_of_range_box = Get_range_address(box.row_onedex, box.column_onedex, box.count_rows, box.count_columns)
End Function


Public Function Get_range_box_from_address(address As String) As RANGE_BOX.RANGE_BOX
    Dim result As RANGE_BOX.RANGE_BOX
    
    ' 1. Parse chars into blocks
    Const max_blocks As Long = 5
    Dim parsed_number_buffer(max_blocks) As Long
    Dim parsed_block_type_buffer(max_blocks) As Long
    
    
    address = Replace(address, " ", "")
    address = Replace(address, "$", "")
    
    If Len(address) < 2 Then GoTo PARSING_FAILED

    Dim first_char As String: first_char = Mid(address, 1, 1)


    Dim previous_block_type As Long
    Dim cumulative_number As Long
    Parse_char first_char, cumulative_number, previous_block_type
    If previous_block_type = RANGE_ADDRESS_BLOCK_TYPE_UNDEFINED Then GoTo PARSING_FAILED

    
    Dim count_filled_blocks As Long: count_filled_blocks = 0
    
    Dim char_onedex As Long
    For char_onedex = 2 To Len(address)
        Dim ch As String: ch = Mid(address, char_onedex, 1)
    
        Dim current_number As Long
        Dim current_block_type As Long
        Parse_char ch, current_number, current_block_type
        
        If current_block_type = RANGE_ADDRESS_BLOCK_TYPE_UNDEFINED Then GoTo PARSING_FAILED
        
        If current_block_type = previous_block_type Then
            If current_block_type = RANGE_ADDRESS_BLOCK_TYPE_ROW Then
                cumulative_number = cumulative_number * 10
            ElseIf current_block_type = RANGE_ADDRESS_BLOCK_TYPE_COLUMN Then
                cumulative_number = cumulative_number * APLHABET_LENGTH
            Else
                GoTo PARSING_FAILED
            End If
            
            cumulative_number = cumulative_number + current_number
        Else
            parsed_number_buffer(count_filled_blocks + 1) = cumulative_number
            parsed_block_type_buffer(count_filled_blocks + 1) = previous_block_type
            count_filled_blocks = count_filled_blocks + 1
            
            If count_filled_blocks >= max_blocks Then GoTo PARSING_FAILED
            
            cumulative_number = current_number
        End If
        
        previous_block_type = current_block_type
        
    Next char_onedex
    
    parsed_number_buffer(count_filled_blocks + 1) = cumulative_number
    parsed_block_type_buffer(count_filled_blocks + 1) = previous_block_type
    count_filled_blocks = count_filled_blocks + 1
    
    
    
    ' 2. Prepare return
    Dim column_onedex As Long
    Dim row_onedex As Long
    Dim first_column_onedex As Long
    Dim first_row_onedex As Long
    Dim last_column_onedex As Long
    Dim last_row_onedex As Long
    Dim count_columns As Long
    Dim count_rows As Long
    

    If count_filled_blocks = 2 And _
            parsed_block_type_buffer(1) = RANGE_ADDRESS_BLOCK_TYPE_COLUMN And _
            parsed_block_type_buffer(2) = RANGE_ADDRESS_BLOCK_TYPE_ROW _
    Then ' A1
        
        column_onedex = parsed_number_buffer(1)
        row_onedex = parsed_number_buffer(2)

        result.column_onedex = column_onedex
        result.row_onedex = row_onedex
        result.count_columns = 1
        result.count_rows = 1
        
    ElseIf count_filled_blocks = 3 And _
            parsed_block_type_buffer(2) = RANGE_ADDRESS_BLOCK_TYPE_SEPARATOR And _
            parsed_block_type_buffer(1) = parsed_block_type_buffer(3) _
    Then ' 1:1 or A:A
        Dim block_type As Long: block_type = parsed_block_type_buffer(1)
        
        If block_type = RANGE_ADDRESS_BLOCK_TYPE_COLUMN Then 'A:A
            first_column_onedex = parsed_number_buffer(1)
            last_column_onedex = parsed_number_buffer(3)
            
            Ensure_ascending_order first_column_onedex, last_column_onedex
            
            count_columns = last_column_onedex - first_column_onedex + 1
            
            result.column_onedex = first_column_onedex
            result.row_onedex = 1
            result.count_columns = count_columns
            result.count_rows = EXCEL_WORKSHEET_MAX_ROWS
        Else '1:1
            first_row_onedex = parsed_number_buffer(1)
            last_row_onedex = parsed_number_buffer(3)
            
            Ensure_ascending_order first_row_onedex, last_row_onedex
            
            count_rows = last_row_onedex - first_row_onedex + 1
            
            result.column_onedex = 1
            result.row_onedex = first_row_onedex
            result.count_columns = EXCEL_WORKSHEET_MAX_COLUMNS
            result.count_rows = count_rows
        End If
        
    ElseIf count_filled_blocks = 5 And _
            parsed_block_type_buffer(1) = RANGE_ADDRESS_BLOCK_TYPE_COLUMN And _
            parsed_block_type_buffer(2) = RANGE_ADDRESS_BLOCK_TYPE_ROW And _
            parsed_block_type_buffer(3) = RANGE_ADDRESS_BLOCK_TYPE_SEPARATOR And _
            parsed_block_type_buffer(4) = RANGE_ADDRESS_BLOCK_TYPE_COLUMN And _
            parsed_block_type_buffer(5) = RANGE_ADDRESS_BLOCK_TYPE_ROW _
    Then ' A1:B2
        first_column_onedex = parsed_number_buffer(1)
        first_row_onedex = parsed_number_buffer(2)
        last_column_onedex = parsed_number_buffer(4)
        last_row_onedex = parsed_number_buffer(5)
        
        Ensure_ascending_order first_row_onedex, last_row_onedex
        Ensure_ascending_order first_column_onedex, last_column_onedex
        
        count_rows = last_row_onedex - first_row_onedex + 1
        count_columns = last_column_onedex - first_column_onedex + 1
        
        result.column_onedex = first_column_onedex
        result.row_onedex = first_row_onedex
        result.count_columns = count_columns
        result.count_rows = count_rows
    Else
        GoTo PARSING_FAILED
    End If
    
    Get_range_box_from_address = result
    
    
PARSING_FAILED:
    ' no actions
End Function




Private Function Unchecked_express_onedex_in_AZ(onedex As Long) As String
    If (onedex <= APLHABET_LENGTH) Then
        Unchecked_express_onedex_in_AZ = Mid(APLHABET, onedex, 1)
    Else
        Dim q As Long: q = (onedex - 1) \ APLHABET_LENGTH
        Dim r As Long: r = (onedex - 1) Mod APLHABET_LENGTH

        Unchecked_express_onedex_in_AZ = Unchecked_express_onedex_in_AZ(q) + Mid(APLHABET, r + 1, 1)
    End If
End Function


Private Sub Ensure_ascending_order(ByRef x As Long, ByRef y As Long)
    If x > y Then
        Dim z As Long
        z = x
        x = y
        y = z
    End If
End Sub



Private Sub Parse_char(ch As String, ByRef number As Long, ByRef block_type As Long)
    
    Dim char_code_ch As Long: char_code_ch = Asc(ch)
    
    If Asc(ch) >= Asc("A") And Asc(ch) <= Asc("Z") Then
        
        number = Asc(ch) - Asc("A") + 1
        block_type = RANGE_ADDRESS_BLOCK_TYPE_COLUMN
        
    ElseIf Asc(ch) >= Asc("0") And Asc(ch) <= Asc("9") Then
        
        number = Asc(ch) - Asc("0")
        block_type = RANGE_ADDRESS_BLOCK_TYPE_ROW
        
    ElseIf Asc(ch) = Asc(":") Then
        
        number = 0
        block_type = RANGE_ADDRESS_BLOCK_TYPE_SEPARATOR
        
    ElseIf Asc(ch) >= Asc("a") And Asc(ch) <= Asc("z") Then
        
        number = Asc(ch) - Asc("a") + 1
        block_type = RANGE_ADDRESS_BLOCK_TYPE_COLUMN
    Else
        
        number = 0
        block_type = RANGE_ADDRESS_BLOCK_TYPE_UNDEFINED
        
    End If
    
End Sub
