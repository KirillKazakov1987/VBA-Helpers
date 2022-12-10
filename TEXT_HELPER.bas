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
    Debug.Assert ARRAY_HELPER.Get_array_rank(values) = 1
    
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
    Debug.Assert ARRAY_HELPER.Get_array_rank(values) = 2

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
        
            If l > widths(j - j_first) Then widths(j - j_first) = l
            
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




Public Function Count_substrings(search_space As String, substr As String) As Long
    Dim sublen As Long, splen As Long
    sublen = Len(substr)
    splen = Len(search_space)
    
    If sublen = 0 Or splen = 0 Then
        Count_substrings = 0
        Exit Function
    End If

    
    
    Dim fragment As String
    
    Dim i As Long
    i = 1
    
    Dim counter As Long
    counter = 0
    
    Do While i + sublen <= splen + 1
        fragment = Mid(search_space, i, sublen)
        If fragment = substr Then
            counter = counter + 1
            i = i + sublen
        Else
            i = i + 1
        End If
        
    Loop
    
    Count_substrings = counter

End Function




Public Function Onedex_of(search_space As String, substr As String) As Optional_int32
    Dim sublen As Long, splen As Long
    sublen = Len(substr)
    splen = Len(search_space)
    
    If sublen = 0 Or splen = 0 Then
        Onedex_of.Has_value = False
        Exit Function
    End If

    
    
    Dim fragment As String
    
    Dim i As Long
    i = 1
    
    Dim counter As Long
    counter = 0
    
    Do While i + sublen <= splen + 1
        fragment = Mid(search_space, i, sublen)
        If fragment = substr Then
            Onedex_of.Has_value = True
            Onedex_of.value = i
            Exit Function
        Else
            i = i + 1
        End If
    Loop
    
    Onedex_of.Has_value = False
End Function



Public Function LTrim_custom(subj As String, trim_fragment As String) As String
    Dim ns As Long: ns = Len(subj)
    Dim nt As Long: nt = Len(trim_fragment)
    
    If ns = 0 Then GoTo NOT_CHANGED
    If nt = 0 Then GoTo NOT_CHANGED
    
    Dim i As Long: i = 1
    
    Dim s As String
    Do While i + nt < ns + 1
        s = Mid(subj, i, nt)
        
        If s = trim_fragment Then
            i = i + nt
        Else
            LTrim_custom = Mid(subj, i)
            Exit Function
        End If
    Loop
    
    LTrim_custom = Mid(subj, i)
    Exit Function
    
NOT_CHANGED:
    LTrim_custom = subj
     
End Function



Public Function RTrim_custom(subj As String, trim_fragment As String) As String
    subj = StrReverse(subj)
    trim_fragment = StrReverse(trim_fragment)
    RTrim_custom = LTrim_custom(subj, trim_fragment)
    RTrim_custom = StrReverse(RTrim_custom)
End Function


Public Function Trim_custom(subj As String, trim_fragment As String) As String
    subj = LTrim_custom(subj, trim_fragment)
    subj = RTrim_custom(subj, trim_fragment)
    Trim_custom = subj
End Function

Public Function Trim_custom2(subj As String, trim_fragment1 As String, trim_fragment2 As String) As String
    subj = Trim_custom(subj, trim_fragment1)
    subj = Trim_custom(subj, trim_fragment2)
    Trim_custom2 = subj
End Function

Public Function Trim_custom3(subj As String, trim_fragment1 As String, trim_fragment2 As String, trim_fragment3 As String) As String
    subj = Trim_custom(subj, trim_fragment1)
    subj = Trim_custom(subj, trim_fragment2)
    subj = Trim_custom(subj, trim_fragment3)
    Trim_custom3 = subj
End Function

Public Function Trim_custom4(subj As String, trim_fragment1 As String, trim_fragment2 As String, trim_fragment3 As String, trim_fragment4 As String) As String
    subj = Trim_custom(subj, trim_fragment1)
    subj = Trim_custom(subj, trim_fragment2)
    subj = Trim_custom(subj, trim_fragment3)
    subj = Trim_custom(subj, trim_fragment4)
    Trim_custom4 = subj
End Function



Public Function Starts_with(subj As String, expected_in_start As String) As Boolean
    Dim ns As Long
    Dim ne As Long
    
    ns = Len(subj)
    ne = Len(expected_in_start)
    
    
    If ne > ns Then GoTo NOT_STARTS
    
    If ns = 0 Then
        Starts_with = True
        Exit Function
    End If
    
    
    Dim s As String
    s = Mid(subj, 1, ne)
    
    Starts_with = (s = expected_in_start)
    Exit Function
    
NOT_STARTS:
    expected_in_start = False

End Function



