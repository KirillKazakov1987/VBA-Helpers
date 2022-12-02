Attribute VB_Name = "STYLE_INFO"
Option Explicit


Private Type Style_replacement_pair
    Old_value As String
    new_value As String
End Type

Private Const wsname As String = "StyleInfo"
Private Const loname_list_of_styles As String = "StyleList"
Private Const loname_style_replacement As String = "StyleRepl"
Private Const loname_style_deleting As String = "StyleDel"



Sub Create_worksheet_with_style_information()

    Dim ws As Worksheet
    Dim ws_found As Boolean: ws_found = False
    For Each ws In ThisWorkbook.Worksheets
        If wsname = ws.Name Then
            ws.Cells.Clear
            ws.Rows.Delete
            ws.Columns.Delete
            ws_found = True
            ws.Activate
            Exit For
        End If
    Next ws
    
    If ws_found = False Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = wsname
    End If

    Dim Nst As Long: Nst = ThisWorkbook.Styles.Count
    Dim aStData() As Variant: ReDim aStData(1 To Nst + 1, 1 To 30)
    Dim j As Long: j = 0

    j = j + 1: aStData(1, j) = "Onedex"
    j = j + 1: aStData(1, j) = "Name"
    j = j + 1: aStData(1, j) = "NameLocal"
    j = j + 1: aStData(1, j) = "IsBuiltIn"
    j = j + 1: aStData(1, j) = "IsLocked"
    j = j + 1: aStData(1, j) = "FontName"
    j = j + 1: aStData(1, j) = "FontSize"
    j = j + 1: aStData(1, j) = "IsFontBold"
    j = j + 1: aStData(1, j) = "IsFontItalic"
    j = j + 1: aStData(1, j) = "FontUnderline"
    j = j + 1: aStData(1, j) = "FontColorRGB"
    j = j + 1: aStData(1, j) = "NumberFormat"
    j = j + 1: aStData(1, j) = "NumberFormatLocal"
    
    Dim st As Style
    Dim i As Long: i = 1
    For Each st In ThisWorkbook.Styles
        i = i + 1
        j = 0
        
        j = j + 1: aStData(i, j) = i - 1 'number
        j = j + 1: aStData(i, j) = st.Name 'name
        j = j + 1: aStData(i, j) = st.NameLocal 'name local
        j = j + 1: aStData(i, j) = IIf(st.BuiltIn, 1, 0) 'IsBuiltIn
        j = j + 1: aStData(i, j) = IIf(st.Locked, 1, 0) 'IsLocked
        j = j + 1: aStData(i, j) = st.Font.Name 'FontName
        j = j + 1: aStData(i, j) = st.Font.size 'FontSize
        j = j + 1: aStData(i, j) = IIf(st.Font.Bold, 1, 0) 'FontBold
        j = j + 1: aStData(i, j) = IIf(st.Font.Italic, 1, 0) 'FontItalic
        j = j + 1: aStData(i, j) = IIf(st.Font.Underline, 1, 0) 'FontUnderline
        j = j + 1: aStData(i, j) = st.Font.Color 'FontColorRGB
        j = j + 1: aStData(i, j) = st.NumberFormat 'NumberFormat
        j = j + 1: aStData(i, j) = st.NumberFormatLocal 'NumberFormatLocal
        
    Next st
    
    Dim Nc_sl As Long: Nc_sl = j
    
    [a2].Value = "List of styles"
    Dim dst_rng_address As String
    dst_rng_address = Get_range_address(3, 1, Nst + 1, Nc_sl)
    
    Dim filled_rng As Range: Set filled_rng = ws.Range(dst_rng_address)
    filled_rng.Value2 = aStData

    Dim lo As ListObject: Set lo = ws.ListObjects.Add(xlSrcRange, filled_rng, , xlYes)
    lo.Name = loname_list_of_styles
    
    Dim bool_number_fmt As String
    bool_number_fmt = "[=1][Color10]" & Chr(34) & "V" & Chr(34) & ";[=0][Red]" & Chr(34) & "X" & Chr(34) & ";General"


    With lo.ListColumns.Item("IsBuiltIn").DataBodyRange
        .NumberFormat = bool_number_fmt
        .Font.Bold = True
    End With
    
    With lo.ListColumns.Item("IsLocked").DataBodyRange
        .NumberFormat = bool_number_fmt
        .Font.Bold = True
    End With
    
    With lo.ListColumns.Item("IsFontBold").DataBodyRange
        .NumberFormat = bool_number_fmt
        .Font.Bold = True
    End With
    
    With lo.ListColumns.Item("IsFontItalic").DataBodyRange
        .NumberFormat = bool_number_fmt
        .Font.Bold = True
    End With
    
    With lo.ListColumns.Item("FontUnderline").DataBodyRange
        .NumberFormat = bool_number_fmt
        .Font.Bold = True
    End With
    
    
    
    ' Create table for style replacement
    dst_rng_address = Get_range_address(3, Nc_sl + 3)
    
    Dim dst_rng As Range
    Set dst_rng = ws.Range(dst_rng_address)
    
    dst_rng.Offset(-1, 0).Value = "Style replacement instructions"

    Set lo = Create_empty_list_object(dst_rng, loname_style_replacement, 10, _
        "WorksheetName", "OldStyleName", "NewStyleName")
    
    Add_button_to_range dst_rng.Offset(-2, 0), "REPLACE", "Replace_styles_from_list_object_instructions"
    
    
    ' Create table for style deletion
    dst_rng_address = Get_range_address(3, dst_rng.Column + lo.ListColumns.Count + 2)
    Set dst_rng = ws.Range(dst_rng_address)
    
    dst_rng.Offset(-1, 0).Value = "Style deleting instructions"

    Set lo = Create_empty_list_object(dst_rng, loname_style_deleting, 10, "StyleName")
    
    Add_button_to_range dst_rng.Offset(-2, 0), "DELETE", "Delete_styles_from_list_object_instructions"

End Sub



Public Sub Delete_styles_from_list_object_instructions()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim lo As ListObject: Set lo = ws.ListObjects(loname_style_deleting)
    
    Dim content As Variant: content = lo.DataBodyRange.Value2
    Dim s As String
    Dim st As Style
    
    If IsArray(content) Then
        Dim i As Long
        For i = LBound(content, 1) To UBound(content, 1)
            
            s = content(i, LBound(content, 2))
            If Not s = "" Then
                On Error Resume Next
                    Set st = ThisWorkbook.Styles.Item(s)
                    st.Delete
                    
                    If Err.number = 0 Then
                        Debug.Print "STYLE " & s & " DELETED SUCSESSFULLY."
                    Else
                        Debug.Print "STYLE " & s & " DELETING FAILED."
                    End If
                On Error GoTo 0
            End If
        Next i
    Else
        s = content
        If Not s = "" Then
            On Error Resume Next
                Set st = ThisWorkbook.Styles.Item(s)
                st.Delete
                
                If Err.number = 0 Then
                    Debug.Print "STYLE " & s & " DELETED SUCSESSFULLY."
                Else
                    Debug.Print "STYLE " & s & " DELETING FAILED."
                End If
            On Error GoTo 0
        End If
    
    End If

End Sub





Private Function Create_empty_list_object( _
    upper_left_cell As Range, _
    lo_name As String, _
    inserting_empty_rows As Long, _
    ParamArray field_names() As Variant) As ListObject
    
    Dim ws As Worksheet: Set ws = upper_left_cell.Worksheet
    
    
    inserting_empty_rows = IIf(inserting_empty_rows < 1, 1, inserting_empty_rows)
    
    Dim nc As Long: nc = UBound(field_names) - LBound(field_names) + 1
    Dim nr As Long: nr = 1 + inserting_empty_rows
    
    Dim filling_range_address As String
    filling_range_address = Get_range_address(upper_left_cell.Row, upper_left_cell.Column, nr, nc)
    
    Dim filling_range As Range: Set filling_range = ws.Range(filling_range_address)
    filling_range.Clear
    
    Dim headers() As String: ReDim headers(1 To 1, 1 To nc)
    Dim i As Long
    For i = 1 To nc
        headers(1, i) = field_names(LBound(field_names) + i - 1)
    Next i
    
    ARRAY_TO_EXCEL_RANGE.Write_2D_array_to_excel_range headers, upper_left_cell
    
    Dim lo As ListObject
    Set lo = ws.ListObjects.Add(xlSrcRange, filling_range, , xlYes)
    lo.Name = lo_name
    
    Set Create_empty_list_object = lo
End Function


Private Function Add_button_to_range( _
    rng As Range, _
    button_name As String, _
    Optional action_name As String = "")

    Dim dx As Double: dx = rng.Width
    Dim dy As Double: dy = rng.Height

    Dim x0 As Double: x0 = 0
    Dim y0 As Double: y0 = 0
    
    If rng.Column > 1 Then
        Dim c1 As Range: Set c1 = rng.Worksheet.Range("A1")
        Dim c2 As Range: Set c2 = rng.Offset(0, -1)
        Dim rng_sz1 As Range: Set rng_sz1 = rng.Worksheet.Range(c1, c2)
        x0 = rng_sz1.Width
    End If
    
    If rng.Row > 1 Then
        Dim c3 As Range: Set c3 = rng.Worksheet.Range("A1")
        Dim c4 As Range: Set c4 = rng.Offset(-1, 0)
        Dim rng_sz2 As Range: Set rng_sz2 = rng.Worksheet.Range(c3, c4)
        y0 = rng_sz2.Height
    End If
    
    Dim btn As Button: Set btn = rng.Worksheet.Buttons.Add(x0, y0, dx, dy)
    btn.Name = button_name
    btn.text = button_name
    
    If Not action_name = "" Then
        btn.OnAction = action_name
    End If
    
    Set Add_button_to_range = btn
End Function
    
    
    
Public Sub Replace_styles_from_list_object_instructions()
    Dim lo As ListObject
    Set lo = ActiveSheet.ListObjects(loname_style_replacement)
    
    Dim nr As Long: nr = lo.DataBodyRange.Rows.Count

    Dim worksheet_names() As String: ReDim worksheet_names(1 To nr)
    Dim style_old_names() As String: ReDim style_old_names(1 To nr)
    Dim style_new_names() As String: ReDim style_new_names(1 To nr)

    Dim r As ListRow
    Dim i As Long: i = 1
    For Each r In lo.ListRows
        worksheet_names(i) = r.Range(1, 1).Value
        style_old_names(i) = r.Range(1, 2).Value
        style_new_names(i) = r.Range(1, 3).Value
        i = i + 1
    Next r


    Replace_styles_in_worksheets worksheet_names, style_old_names, style_new_names
End Sub
    

    
Private Sub Replace_styles_in_worksheets( _
    worksheet_names() As String, _
    style_old_names() As String, _
    style_new_names() As String)
    
    Debug.Assert LBound(worksheet_names) = LBound(style_old_names)
    Debug.Assert LBound(style_old_names) = LBound(style_new_names)
    
    Debug.Assert UBound(worksheet_names) = UBound(style_old_names)
    Debug.Assert UBound(style_old_names) = UBound(style_new_names)
    
    
    Dim dict As Variant
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim col As Collection
    Dim pair As Style_replacement_pair

    Dim i As Long
    For i = LBound(worksheet_names) To UBound(worksheet_names)
        Dim ws_name As String: ws_name = worksheet_names(i)
        pair.Old_value = style_old_names(i)
        pair.new_value = style_new_names(i)
        
        If ws_name = "" Or pair.Old_value = "" Or pair.new_value = "" Then
            GoTo END_OF_LOOP_ITERATION
        End If
            
        If dict.Exists(LCase(ws_name)) Then
            Set col = dict(LCase(ws_name))
            col.Add pair.Old_value & "@" & pair.new_value
        Else
            Set col = New Collection
            col.Add pair.Old_value & "@" & pair.new_value
            dict.Add LCase(ws_name), col
        End If
        
END_OF_LOOP_ITERATION:
    Next i
    
    Dim k As Variant
    For Each k In dict.keys
        Set col = dict(k)
        
        Dim pairs() As Style_replacement_pair: ReDim pairs(1 To col.Count)
        
        For i = 1 To col.Count
            Dim onv As String: onv = col(i)
            Dim arr_str() As String: arr_str = Split(onv, "@")

            pairs(i).Old_value = arr_str(LBound(arr_str) + 0)
            pairs(i).new_value = arr_str(LBound(arr_str) + 1)
        Next i
        
        Dim ws As Worksheet
        On Error Resume Next
            Set ws = Worksheets(k)
            
            If Err.number = 0 Then
                Replace_styles pairs, ws.Cells
            End If
            
        On Error GoTo 0
    Next k
End Sub
    
    

    
Private Sub Replace_styles( _
    pairs() As Style_replacement_pair, _
    subj As Range, _
    Optional too_much_cells_value As Long = 10000000)

    Dim n As Long: n = UBound(pairs) - LBound(pairs) + 1
    Dim old_styles() As Style: ReDim old_styles(1 To n)
    Dim new_styles() As Style: ReDim new_styles(1 To n)
    Dim filters() As Boolean: ReDim filters(1 To n)
    
    Dim ws As Worksheet: Set ws = subj.Worksheet
    Dim wb As Workbook: Set wb = ws.Parent
    Dim all_styles As Styles: Set all_styles = wb.Styles
    
    Dim i As Long
    For i = 1 To n
        Dim pair As Style_replacement_pair: pair = pairs(i - 1 + LBound(pairs))
    
        On Error Resume Next
            Dim old_style As Style: Set old_style = all_styles.Item(pair.Old_value)
            Dim new_style As Style: Set new_style = all_styles(pair.new_value)
            
            If Not Err.number = 0 Then
                filters(i) = False
                Debug.Print "MISSED: " & pair.Old_value & " AND " & pair.new_value
            End If
        On Error GoTo 0
        
        Set old_styles(i) = old_style
        Set new_styles(i) = new_style
        filters(i) = True
    Next i
    
    Dim used_range As Range
    Set used_range = ws.UsedRange
    Set subj = Application.Intersect(subj, used_range)


    Dim cell As Range
    For Each cell In subj
        For i = 1 To n
            If cell.Style = old_styles(i) Then cell.Style = new_styles(i)
        Next i
    Next cell

End Sub
