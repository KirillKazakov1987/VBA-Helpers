Attribute VB_Name = "TEXT_HELPER"
Option Explicit
Option Base 0


Public Function Pad_left( _
        v As Variant, _
        total_length As Long, _
        Optional pad_char As String = " ") As String
        
    Pad_left = Pad(v, total_length, pad_char, True)
        
End Function


Public Function Pad_right( _
        v As Variant, _
        total_length As Long, _
        Optional pad_char As String = " ") As String
        
    Pad_right = Pad(v, total_length, pad_char, False)
        
End Function


Public Function Required_width_of_1D_array( _
        fmt As String, _
        values As Variant _
        ) As Long
    
    Debug.Assert IsArray(values)
    Debug.Assert ARRAY_HELPER.Rank(values) = 1
    
    Dim max_len As Long: max_len = 0
    
    Dim first_index As Long: first_index = LBound(values)
    Dim last_index As Long: last_index = UBound(values)
    
    Dim i As Long: For i = first_index To last_index
        Dim v As Variant: v = values(i)
        Dim s As String: s = CStr(v)
        Dim fs As String: fs = Format(v, fmt)
        Dim l As Long: l = Len(fs)
        
        If l > max_len Then max_len = l
    Next i
    
    Required_width_of_1D_array = max_len
        
End Function




Public Function Required_widths_of_2D_array( _
        fmt As String, _
        values As Variant _
        ) As Long()
        
    Debug.Assert IsArray(values)
    Debug.Assert ARRAY_HELPER.Rank(values) = 2

    Dim i_first As Long: i_first = LBound(values, 1)
    Dim i_last As Long: i_last = UBound(values, 1)
    Dim j_first As Long: j_first = LBound(values, 2)
    Dim j_last As Long: j_last = UBound(values, 2)

    Dim count_columns: count_columns = j_last - j_first + 1

    Dim widths() As Long: ReDim widths(0 To count_columns - 1)
    

    Dim i As Long: For i = i_first To i_last
        
        Dim j As Long: For j = j_first To j_last
            Dim v As Variant: v = values(i, j)
            Dim s As String: s = CStr(v)
            Dim fs As String: fs = Format(v, fmt)
            Dim l As Long: l = Len(fs)
        
            If l > widths(j) Then widths(j) = l
            
        Next j
    Next i
    
    Required_widths_of_2D_array = widths
        
End Function


Private Function Pad( _
        v As Variant, _
        total_length As Long, _
        pad_char As String, _
        is_left_pad As Boolean) As String
    
    Dim v_text As String: v_text = CStr(v)
    Dim v_len As Long: v_len = Len(v_text)
    
    Dim required_p_len As Long: required_p_len = total_length - v_len
    
    If required_p_len > 0 Then
        Dim padding_text As String: padding_text = String(required_p_len, pad_char)
        Dim p_len As Long: p_len = Len(padding_text)
        
        If p_len <= 0 Then
            padding_text = String(required_p_len, " ")
        ElseIf p_len > required_p_len Then
            padding_text = Mid(padding_text, 1, required_p_len)
        End If
        
        If is_left_pad Then
            Pad = padding_text & v_text
        Else
            Pad = v_text & padding_text
        End If
    Else
        Pad = v_text
    End If
    
End Function

