Attribute VB_Name = "RANGE_BOX_MODULE"
Option Explicit

Public Type Range_box
    row_onedex As Long
    column_onedex As Long
    count_rows As Long
    count_columns As Long
    address As String
End Type



Private Const RANGE_ADDRESS_BLOCK_TYPE_UNDEFINED As Long = 0
Private Const RANGE_ADDRESS_BLOCK_TYPE_COLUMN As Long = 1
Private Const RANGE_ADDRESS_BLOCK_TYPE_ROW As Long = 2
Private Const RANGE_ADDRESS_BLOCK_TYPE_SEPARATOR As Long = 3

Private Const APLHABET As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Private Const APLHABET_LENGTH As Long = 26


Public Function Get_cell_address( _
    row_onedex As Long, _
    column_onedex As Long)
    
    If row_onedex < 0 Or column_onedex < 0 Then Err.Raise 999999
    
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


Function Unchecked_express_onedex_in_AZ(onedex As Long) As String
        If (onedex <= APLHABET_LENGTH) Then
            Unchecked_express_onedex_in_AZ = Mid(APLHABET, onedex, 1)
        Else
            Dim q As Long: q = (onedex - 1) / APLHABET_LENGTH
            Dim r As Long: r = (onedex - 1) Mod APLHABET_LENGTH

            Unchecked_express_onedex_in_AZ = Unchecked_express_onedex_in_AZ(q) + Mid(APLHABET, r + 1, 1)
        End If
End Function











Public Function Get_range_box_from_address(address As String) As Range_box
    Dim result As Range_box
    
    ' 1. Parse chars into blocks
    Const max_blocks As Long = 5
    Dim parsed_number_buffer(max_blocks) As Long
    Dim parsed_block_type_buffer(parsed_block_type_buffer) As Long
    
    
    address = Replace(address, " ", "")
    address = Replace(address, "$", "")
    
    If Len(address) < 2 Then GoTo PARSING_FAILED

    Dim first_char As String: first_char = Mid(address, 1, 1)


    Dim previous_block_type As Long
    Dim cumulative_number As Long
    Parse_char first_char, cumulative_number, previous_block_type
    If previous_block_type = Address_block_type.Undefined Then GoTo PARSING_FAILED

    
    Dim count_filled_blocks As Long: count_filled_blocks = 0
    
    Dim char_onedex As Long:
    For char_onedex = 2 To Len(address)
        Dim ch As String: ch = Mid(address, char_onedex, 1)
    
        current_number As Long
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
    If count_filled_blocks = 2 And _
            parsed_block_type_buffer(1) = RANGE_ADDRESS_BLOCK_TYPE_COLUMN And _
            parsed_block_type_buffer(2) = RANGE_ADDRESS_BLOCK_TYPE_ROW Then ' A1
        {
            Dim column_index As Long: column_index = parsed_number_buffer(0) - 1
            Dim row_index As Long: row_index = parsed_number_buffer(1) - 1

            result = new Range_box(row_index, column_index);

            return true;
        }





    Get_range_box_from_address = result
PARSING_FAILED:
    
End Function




Private Sub Parse_char(ch As String, ByRef number As Long, ByRef Address_block_type As Long)
    Const char_code_A As Long = Asc("A")
    Const char_code_Z As Long = Asc("Z")
    
    
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
