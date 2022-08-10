Attribute VB_Name = "ARRAY_TO_TEXT"
Public Function Print_1D_array( _
        values As Variant, _
        Optional fmt As String = "") As String
        
    Debug.Assert IsArray(values)
    Debug.Assert ARRAY_HELPER.Rank(values) = 1

    Dim sb As New STRING_BUILDER
    
    Dim first_index As Long: first_index = LBound(values)
    Dim last_index As Long: last_index = UBound(values)

    Dim cap_i As String: cap_i = "Index"
    Dim cap_v As String: cap_v = "Value"

    Dim value_column_required_length As Long
    value_column_required_length = TEXT_HELPER.Required_width_of_1D_array(fmt, values)
    value_column_required_length = MATH_HELPER.Max_i32(value_column_required_length, Len(cap_v))
    
    
    Dim idx_column_required_length As Long
    idx_column_required_length = Len(CStr(last_index))
    idx_column_required_length = MATH_HELPER.Max_i32(idx_column_required_length, Len(cap_i))
    
    
    Dim width_int As Long: width_int = idx_column_required_length + 3 + value_column_required_length
    Dim width_ext As Long: width_ext = 2 + width_int + 2
    
    
    Dim s1 As String: s1 = String(idx_column_required_length, "-")
    Dim s2 As String: s2 = String(value_column_required_length, "-")
    Dim hsep As String: hsep = "+-" & s1 & "-+-" & s2 & "-+"
    sb.Append_line hsep

    s1 = Pad_left(cap_i, idx_column_required_length)
    s2 = Pad_left(cap_v, value_column_required_length)
    sb.Append_line "| " & s1 & " | " & s2 & " |"
    
    sb.Append_line hsep
    
    
    Dim i As Long: For i = first_index To last_index
        Dim v As Variant: v = values(i)
        s1 = Pad_left(i, idx_column_required_length)
        s2 = Pad_left(Format(v, fmt), value_column_required_length)
        
        sb.Append_line "| " & s1 & " | " & s2 & " |"
    Next i

    sb.Append_line hsep
    
    Print_1D_array = sb.Get_string()
    
End Function


Public Function Print_2D_array( _
        values As Variant, _
        Optional fmt As String = "") As String
        
    Debug.Assert IsArray(values)
    Debug.Assert ARRAY_HELPER.Rank(values) = 2

    Dim sb As New STRING_BUILDER
    
    Dim i_first As Long: i_first = LBound(values, 1)
    Dim i_last As Long: i_last = UBound(values, 1)
    Dim j_first As Long: j_first = LBound(values, 2)
    Dim j_last As Long: j_last = UBound(values, 2)

    Dim corner_cap As String: corner_cap = "R\C"
    
    Dim widths() As Long
    widths = TEXT_HELPER.Required_widths_of_2D_array(fmt, values)
    
    Dim j As Long: For j = LBound(widths) To UBound(widths)
        Dim l As Long: l = Len(CStr(j))
        If l > widths(j) Then widths(j) = l
    Next j
    
    Dim idx_col_width As Long
    idx_col_width = Len(CStr(i_last))
    idx_col_width = MATH_HELPER.Max_i32(idx_col_width, Len(corner_cap))
    
    
    Dim hsep As String
    hsep = "+-" & String(idx_col_width, "-")
    
    For j = LBound(widths) To UBound(widths)
        Dim w As Long: w = widths(j)
        hsep = hsep & "-+-" & String(w, "-")
    Next j
    
    hsep = hsep & "-+"

    sb.Append_line hsep

    
    
    Dim cap_row As String
    
    cap_row = "| " & Pad_left(corner_cap, idx_col_width)
    
    For j = LBound(widths) To UBound(widths)
        Dim col_idx As Long: col_idx = j_first + j - LBound(widths)
        cap_row = cap_row & " | " & Pad_left(col_idx, widths(j))
    Next j
    cap_row = cap_row & " |"
    
    sb.Append_line cap_row
    
    sb.Append_line hsep
    
    
    Dim i As Long: For i = i_first To i_last
        Dim data_row As String
        data_row = "| " & Pad_left(i, idx_col_width)
        
        For j = j_first To j_last
            Dim v As Variant: v = values(i, j)
            Dim vf As String: vf = Format(v, fmt)
            
            data_row = data_row & " | " & Pad_left(vf, widths(j))
        Next j
        
        data_row = data_row & " |"
        
        sb.Append_line data_row
    Next i

    sb.Append_line hsep
    
    Print_2D_array = sb.Get_string()
    
End Function
