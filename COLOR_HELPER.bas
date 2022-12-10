Attribute VB_Name = "COLOR_HELPER"
Option Explicit
Option Base 0

Private dictionary_from_color_name_to_color_rgb As DICT_TEXT_TO_VARIANT


Public Function Get_color_as_checking_and_correcting_rgb_integers( _
    red As Long, _
    green As Long, _
    blue As Long) As Long
    
    red = MATH_HELPER.Clamp_i32(red, 0, 255)
    green = MATH_HELPER.Clamp_i32(green, 0, 255)
    blue = MATH_HELPER.Clamp_i32(blue, 0, 255)
    
    Dim color As Long: color = RGB(red, green, blue)
    
    Get_color_as_checking_and_correcting_rgb_integers = color
End Function


Public Function Get_color_as_checking_and_correcting_rgb_fractions( _
    red As Double, _
    green As Double, _
    blue As Double) As Long
    
    red = MATH_HELPER.Clamp_f64(red, 0, 1)
    green = MATH_HELPER.Clamp_f64(green, 0, 1)
    blue = MATH_HELPER.Clamp_f64(blue, 0, 1)
    
    Dim r As Long: r = CLng(255 * red)
    Dim g As Long: g = CLng(255 * green)
    Dim b As Long: b = CLng(255 * blue)
    
    Get_color_as_checking_and_correcting_rgb_fractions = RGB(r, g, b)
End Function



Public Function Create_new_color_dictionary() As DICT_TEXT_TO_VARIANT
    Dim dict As New DICT_TEXT_TO_VARIANT
    dict.Case_sensitivity = False
    
    dict.Add_or_replace " –¿—Õ€…", RGB(255, 0, 0)
    dict.Add_or_replace "RED", RGB(255, 0, 0)
    
    dict.Add_or_replace "—¬≈“ÀŒ-«≈À≈Õ€…", RGB(0, 255, 0)
    dict.Add_or_replace "ﬂ– Œ-«≈À≈Õ€…", RGB(0, 255, 0)
    dict.Add_or_replace "—¬≈“ÀŒ «≈À≈Õ€…", RGB(0, 255, 0)
    dict.Add_or_replace "ﬂ– Œ «≈À≈Õ€…", RGB(0, 255, 0)
    dict.Add_or_replace "—¬≈“ÀŒ-«≈À®Õ€…", RGB(0, 255, 0)
    dict.Add_or_replace "ﬂ– Œ-«≈À®Õ€…", RGB(0, 255, 0)
    dict.Add_or_replace "—¬≈“ÀŒ «≈À®Õ€…", RGB(0, 255, 0)
    dict.Add_or_replace "ﬂ– Œ «≈À®Õ€…", RGB(0, 255, 0)
    dict.Add_or_replace "LIME", RGB(0, 255, 0)
    dict.Add_or_replace "À¿…Ã", RGB(0, 255, 0)
    dict.Add_or_replace "À»ÃŒÕÕ€…", RGB(0, 255, 0)
    dict.Add_or_replace "À»ÃŒÕŒ¬€…", RGB(0, 255, 0)
    
    dict.Add_or_replace "«≈À≈Õ€…", RGB(0, 127, 0)
    dict.Add_or_replace "GREEN", RGB(0, 127, 0)
    
    dict.Add_or_replace "—»Õ»…", RGB(0, 0, 255)
    dict.Add_or_replace "BLUE", RGB(0, 0, 255)
    
    dict.Add_or_replace "“≈ÃÕŒ-—»Õ»…", RGB(0, 0, 127)
    dict.Add_or_replace "“®ÃÕŒ —»Õ»…", RGB(0, 0, 127)
    dict.Add_or_replace "DARK-BLUE", RGB(0, 0, 127)
    dict.Add_or_replace "DARK BLUE", RGB(0, 0, 127)
    dict.Add_or_replace "NAVY", RGB(0, 0, 127)
    
    dict.Add_or_replace "¡≈À€…", RGB(255, 255, 255)
    dict.Add_or_replace "WHITE", RGB(255, 255, 255)
    
    dict.Add_or_replace "◊≈–Õ€…", RGB(0, 0, 0)
    dict.Add_or_replace "◊®–Õ€…", RGB(0, 0, 0)
    dict.Add_or_replace "BLACK", RGB(0, 0, 0)
    
    dict.Add_or_replace "—≈–€…", RGB(127, 127, 127)
    dict.Add_or_replace "GRAY", RGB(127, 127, 127)
    
    dict.Add_or_replace "√ŒÀ”¡Œ…", RGB(0, 255, 255)
    dict.Add_or_replace "AQUA", RGB(0, 255, 255)
    
    dict.Add_or_replace "Ã¿À»ÕŒ¬€…", RGB(255, 0, 255)
    dict.Add_or_replace "–Œ«Œ¬€…", RGB(255, 0, 255)
    dict.Add_or_replace "‘” —»Õ", RGB(255, 0, 255)
    dict.Add_or_replace "MAGNETA", RGB(255, 0, 255)
    dict.Add_or_replace "FUCHSIA", RGB(255, 0, 255)
    
    dict.Add_or_replace "‘»ŒÀ≈“Œ¬€…", RGB(127, 0, 127)
    dict.Add_or_replace "PURPLE", RGB(127, 0, 127)
    
    dict.Add_or_replace "∆≈À“€…", RGB(255, 255, 0)
    dict.Add_or_replace "YELLOW", RGB(255, 255, 0)
    
    dict.Add_or_replace "—≈–≈¡–ﬂÕ€…", RGB(191, 191, 191)
    dict.Add_or_replace "SILVER", RGB(191, 191, 191)
    
    dict.Add_or_replace "MAROON", RGB(127, 0, 0)
    dict.Add_or_replace "¡Œ–ƒŒ¬€…", RGB(127, 0, 0)
    
    
    dict.Add_or_replace "TEAL", RGB(0, 127, 127)
    
    dict.Add_or_replace "OLIVE", RGB(127, 127, 0)
    dict.Add_or_replace "ŒÀ»¬ Œ¬€…", RGB(127, 127, 0)
    
    Set Create_new_color_dictionary = dict
End Function


Public Function Try_get_color_from_its_name(color As String) As Optional_int32
    color = UCase(Trim(color))
    color = Replace(color, "®", "≈")

    If dictionary_from_color_name_to_color_rgb Is Nothing Then
        Set dictionary_from_color_name_to_color_rgb = Create_new_color_dictionary()
    End If
    
    Dim result As Optional_variant
    result = dictionary_from_color_name_to_color_rgb.Try_get_value(color)
    
    
    If result.Has_value Then
        Try_get_color_from_its_name.Has_value = True
        Try_get_color_from_its_name.value = CLng(result.value)
    Else
        Try_get_color_from_its_name.Has_value = False
    End If
    
End Function



Public Function Try_get_color_from_rgb_separated_byte_values( _
    rgb_separated_byte_values As String, _
    Optional separator As String = ",") As Optional_int32
    
    Dim values() As String
    values = Split(rgb_separated_byte_values, separator, 3, vbTextCompare)
    
    If Not ARRAY_HELPER.Get_array_size(values) = 3 Then
        Try_get_color_from_rgb_separated_byte_values.Has_value = False
        Exit Function
    End If
    
    Dim red_str As String
    red_str = ARRAY_HELPER.Get_item_of_array1D(values, 0)
    
    Dim green_str As String
    green_str = ARRAY_HELPER.Get_item_of_array1D(values, 1)
    
    Dim blue_str As String
    blue_str = ARRAY_HELPER.Get_item_of_array1D(values, 2)
    
    
    Dim red_opt_val As Optional_int32
    red_opt_val = TYPE_HELPER.Try_convert_to_int32(red_str)
    
    Dim green_opt_val As Optional_int32
    green_opt_val = TYPE_HELPER.Try_convert_to_int32(green_str)
    
    Dim blue_opt_val As Optional_int32
    blue_opt_val = TYPE_HELPER.Try_convert_to_int32(blue_str)
    
    Dim result As Long
    Dim r As Long, g As Long, b As Long
    If red_opt_val.Has_value And green_opt_val.Has_value And blue_opt_val.Has_value Then
        r = red_opt_val.value
        g = green_opt_val.value
        b = blue_opt_val.value
        result = Get_color_as_checking_and_correcting_rgb_integers(r, g, b)
        
        Try_get_color_from_rgb_separated_byte_values.Has_value = True
        Try_get_color_from_rgb_separated_byte_values.value = result
    Else
        Try_get_color_from_rgb_separated_byte_values.Has_value = False
    End If

End Function


Public Function Try_get_color_from_rgb_percents(rgb_separated_percents As String) As Optional_int32
    If Not TEXT_HELPER.Count_substrings(rgb_separated_percents, "%") = 3 Then
        Try_get_color_from_rgb_percents.Has_value = False
        Exit Function
    End If
    
    Dim values() As String
    values = Split(rgb_separated_percents, "%", 3, vbTextCompare)
    
    If Not ARRAY_HELPER.Get_array_size(values) = 3 Then
        Try_get_color_from_rgb_percents.Has_value = False
        Exit Function
    End If
    
    Dim red_str As String
    red_str = ARRAY_HELPER.Get_item_of_array1D(values, 0)
    red_str = Trim(red_str)
    red_str = TEXT_HELPER.Trim_custom4(red_str, "%", "-", ",", ";")
    
    Dim green_str As String
    green_str = ARRAY_HELPER.Get_item_of_array1D(values, 1)
    green_str = Trim(green_str)
    green_str = TEXT_HELPER.Trim_custom4(green_str, "%", "-", ",", ";")
    
    Dim blue_str As String
    blue_str = ARRAY_HELPER.Get_item_of_array1D(values, 2)
    blue_str = Trim(blue_str)
    blue_str = TEXT_HELPER.Trim_custom4(blue_str, "%", "-", ",", ";")
    
    
    Dim red_opt_val As Optional_float64
    red_opt_val = TYPE_HELPER.Try_convert_to_float64(red_str)
    
    Dim green_opt_val As Optional_float64
    green_opt_val = TYPE_HELPER.Try_convert_to_float64(green_str)
    
    Dim blue_opt_val As Optional_float64
    blue_opt_val = TYPE_HELPER.Try_convert_to_float64(blue_str)
    
    Dim result As Long
    Dim r As Double, g As Double, b As Double
    If red_opt_val.Has_value And green_opt_val.Has_value And blue_opt_val.Has_value Then
        r = red_opt_val.value * 0.01
        g = green_opt_val.value * 0.01
        b = blue_opt_val.value * 0.01
        result = Get_color_as_checking_and_correcting_rgb_fractions(r, g, b)
        
        Try_get_color_from_rgb_percents.Has_value = True
        Try_get_color_from_rgb_percents.value = result
    Else
        Try_get_color_from_rgb_percents.Has_value = False
    End If

End Function


Public Function Try_get_color_from_rgb_hex(rgb_hex As String) As Optional_int32
    rgb_hex = TEXT_HELPER.LTrim_custom(rgb_hex, "#")
    
    rgb_hex = Trim(LCase(rgb_hex))
    
    If TEXT_HELPER.Starts_with(rgb_hex, "0X") Then
        rgb_hex = Replace(rgb_hex, "0x", "&H")
    End If
    If TEXT_HELPER.Starts_with(rgb_hex, "&H") = False Then
        rgb_hex = "&H" + rgb_hex
    End If
    
    If IsNumeric(rgb_hex) = False Then GoTo FAILED
    
    Dim v As Long
    v = CLng(rgb_hex)
    v = v And &HFFFFFF
    Try_get_color_from_rgb_hex.value = v
    Try_get_color_from_rgb_hex.Has_value = True
    Exit Function

FAILED:
    Try_get_color_from_rgb_hex.Has_value = False
End Function






Public Function Try_get_color_from_arbitraty_string(color_as_text As String) As Optional_int32
    Dim result As Optional_int32
    
    result = COLOR_HELPER.Try_get_color_from_its_name(color_as_text)
    If result.Has_value Then
        Try_get_color_from_arbitraty_string = result
        Exit Function
    End If
    
    result = COLOR_HELPER.Try_get_color_from_rgb_separated_byte_values(color_as_text, "-")
    If result.Has_value Then
        Try_get_color_from_arbitraty_string = result
        Exit Function
    End If
    
    result = COLOR_HELPER.Try_get_color_from_rgb_separated_byte_values(color_as_text, ";")
    If result.Has_value Then
        Try_get_color_from_arbitraty_string = result
        Exit Function
    End If
    
    result = COLOR_HELPER.Try_get_color_from_rgb_separated_byte_values(color_as_text, ",")
    If result.Has_value Then
        Try_get_color_from_arbitraty_string = result
        Exit Function
    End If
    
    result = COLOR_HELPER.Try_get_color_from_rgb_separated_byte_values(color_as_text, " ")
    If result.Has_value Then
        Try_get_color_from_arbitraty_string = result
        Exit Function
    End If
    
    result = COLOR_HELPER.Try_get_color_from_rgb_percents(color_as_text)
    If result.Has_value Then
        Try_get_color_from_arbitraty_string = result
        Exit Function
    End If
    
    result = COLOR_HELPER.Try_get_color_from_rgb_hex(color_as_text)
    If result.Has_value Then
        Try_get_color_from_arbitraty_string = result
        Exit Function
    End If
    
    Try_get_color_from_arbitraty_string.Has_value = False
End Function
