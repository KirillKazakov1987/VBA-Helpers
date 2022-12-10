VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EXCEL_CHART_SERIES_VIEW_SETTING"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_line_color As Optional_int32
Private m_interior_color As Optional_int32
Private m_line_style As Optional_int32
Private m_line_dash_style As Optional_int32
Private m_marker_style As Optional_int32
Private m_line_weight As Double
Private m_marker_size As Double
Private m_line_transparency As Optional_float64
Private m_interior_transparency As Optional_float64
Private m_line_visibility As Boolean


Public Sub Set_line_color_as_int32(color As Long)
    m_line_color.value = color
    m_line_color.Has_value = True
End Sub

Public Sub Set_line_color_as_rgb_bytes(red As Long, green As Long, blue As Long)
    Set_line_color_as_int32 COLOR_HELPER.Get_color_as_checking_and_correcting_rgb_integers(red, green, blue)
End Sub

Public Sub Set_line_color_as_rgb_fractions(red As Double, green As Double, blue As Double)
    Set_line_color_as_int32 COLOR_HELPER.Get_color_as_checking_and_correcting_rgb_fractions(red, green, blue)
End Sub

Public Sub Try_set_line_color_as_string(color_as_text As String)
    Dim result As Optional_int32
    result = COLOR_HELPER.Try_get_color_from_arbitraty_string(color_as_text)
    If result.Has_value Then
        Set_line_color_as_int32 result.value
    End If
End Sub


Public Sub Set_interior_color_as_int32(color As Long)
    m_interior_color.value = color
    m_interior_color.Has_value = True
End Sub

Public Sub Set_interior_color_as_rgb_bytes(red As Long, green As Long, blue As Long)
    Set_interior_color_as_int32 COLOR_HELPER.Get_color_as_checking_and_correcting_rgb_integers(red, green, blue)
End Sub

Public Sub Set_interior_color_as_rgb_fractions(red As Double, green As Double, blue As Double)
    Set_interior_color_as_int32 COLOR_HELPER.Get_color_as_checking_and_correcting_rgb_fractions(red, green, blue)
End Sub

Public Sub Try_set_interior_color_as_string(color As String)
    Dim result As Optional_int32
    result = COLOR_HELPER.Try_get_color_from_arbitraty_string(color)
    If result.Has_value Then
        Set_interior_color_as_int32 result.value
    End If
End Sub



Public Sub Set_line_dash_style_as_int32(style As Long)
    
End Sub

Public Sub Set_line_dash_style_as_string(style As String)
    
End Sub

Public Sub Set_line_style_as_int32(style As Long)
    
End Sub

Public Sub Set_line_style_as_string(style As String)
    
End Sub


Public Sub Set_marker_style_as_int32(style As Long)
    
End Sub

Public Sub Set_marker_style_as_string(style As String)
    
End Sub

Public Property Get Line_weigth() As Double

End Property
Public Property Let Line_weigth(w As Double)

End Property


Public Property Get Marker_size() As Double

End Property
Public Property Let Marker_size(w As Double)

End Property

Public Property Get Line_transparency() As Double

End Property
Public Property Let Line_transparency(t As Double)

End Property

Public Property Get Interior_transparency() As Double

End Property
Public Property Let Interior_transparency(t As Double)

End Property



Public Property Get Line_visibility() As Boolean

End Property
Public Property Let Line_visibility(vis As Boolean)
    
        
End Property


Public Property Get Marker_visibility() As Boolean
    If m_marker_style.value = xlMarkerStyleNone Or m_marker_size = 0 Then
        Marker_visibility = False
    Else
        Marker_visibility = True
    End If
End Property
Public Property Let Marker_visibility(vis As Boolean)
    If vis = False Then
        m_marker_style.value = xlMarkerStyleNone
    Else
        If m_marker_style.value = xlMarkerStyleNone Then
            m_marker_style.value = xlMarkerStyleAutomatic
        End If
    End If
End Property


Public Property Get Data_label_font_size() As Double

End Property
Public Property Let Data_label_font_size(sz As Double)

End Property



Public Sub Set_view(s As Excel.Series)
    
    's.LeaderLines
    
    
    s.Format.Line.Visible = IIf(m_line_visibility, msoTrue, msoFalse)
    
    
    If s.Format.Line.Visible = msoTrue And m_line_weight > 0 Then
        s.Format.Line.Weight = m_line_weight
    
        If m_line_transparency.Has_value Then
            s.Format.Line.Transparency = m_line_transparency.value
        End If
        
        If m_line_style.Has_value Then
            s.Format.Line.style = m_line_style.value
        End If
        
        If m_line_dash_style.Has_value Then
            s.Format.Line.DashStyle = m_line_dash_style.value
        End If
        
        If m_line_color.Has_value Then
            s.Format.Line.ForeColor.RGB = m_line_color.value
        End If
    End If

    
    
    If m_marker_style.Has_value Then
        s.MarkerStyle = m_marker_style.value
    End If
    
    If m_marker_size = 0 Then
        s.MarkerStyle = xlMarkerStyleNone
    End If


    If Not s.MarkerStyle = xlMarkerStyleNone Then
        If m_interior_color.Has_value Then
            's.Fill.ForeColor.RGB = 1 ' = m_interior_color.Value
        End If
        
        s.MarkerSize = m_marker_size
        
        If m_interior_transparency.Has_value Then
            s.Format.Fill.Transparency = m_interior_transparency.value
        
        End If
        
        'm_interior_transparency
    
    End If

 

    ' Удаляем все даталейблы
    Dim p As Point
    Dim n As Long: n = UBound(s.values)
    Dim j As Long
    For j = 1 To n
        Set p = s.Points(j)
        If p.HasDataLabel Then p.DataLabel.Delete
    Next j
    
    
    
'    ' Добавляем подпись самого большого значения
'    Dim y_max As Double: y_max = 0
'    Dim y As Double
'    Dim j_max As Long: j_max = n
'    Dim все_y_целые As Boolean: все_y_целые = True
'
'    For j = aНомерПервойТочкиСКоторойИщемМаксЗнач(i) To n
'        y = s.values(j)
'
'        If y >= y_max Then
'            y_max = y
'            j_max = j
'        End If
'
'
'        If Not y = Round(y, 0) Then все_y_целые = False
'    Next j
'
'
'    Dim y_axis As Axis: Set y_axis = cho.Chart.Axes(xlValue)
'    y_axis_number_format = y_axis.TickLabels.NumberFormat
'    Dim y_ax_has_dot As Boolean
'    If InStr(1, y_axis_number_format, ".") > 0 Or InStr(1, y_axis_number_format, ",") > 0 Then
'        y_ax_has_dot = True
'    Else
'        y_ax_has_dot = False
'    End If
'
'
'
'    ' Подписываем все точки
'    Dim data_labels As DataLabels
'    If aПодписиДанных(i) = 1 Then
'        s.ApplyDataLabels
'        Set data_labels = s.DataLabels
'        data_labels.Orientation = xlUpward
'
'        ' Определяем формат лейблов
'        If InStr(1, y_axis_number_format, "%") > 0 Then
'            If y_max < 10 Then
'                data_labels.NumberFormat = "0.0%;-0.0%;0%"
'            Else
'                data_labels.NumberFormat = "0%;-0%;0%"
'            End If
'        ElseIf y_axis.MaximumScale <= 1 Then
'            data_labels.NumberFormat = "0.00;-0;"""""
'        ElseIf y_axis.MaximumScale <= 30 And Not все_y_целые Then
'            data_labels.NumberFormat = "0.0;-0;"""""
'        Else
'            data_labels.NumberFormat = "# ##0;-0;"""""
'        End If
'
'
'        data_labels.Position = xlLabelPositionAbove
'        s.HasLeaderLines = False
'        data_labels.Format.TextFrame2.TextRange.Font.Bold = msoTrue
'        data_labels.Format.TextFrame2.TextRange.Font.size = aРазмерШрифтаПодписейДанных(i)
'
'        With data_labels.Format.TextFrame2.TextRange.Font.Fill
'            .Visible = msoTrue
'            .ForeColor.RGB = aForeColor(i)
'            .Transparency = 0
'            .Solid
'        End With
'
'    Else
'        If s.HasDataLabels Then
'            Set data_labels = s.DataLabels
'            data_labels.Delete
'        End If
'    End If
'
'
'
'
'
'
'    If y_max > 0 And aПоказыватьМаксЗнач(i) = 1 Then
'        Set p = s.Points(j_max)
'        p.ApplyDataLabels
'
'        Dim point_data_label As DataLabel: Set point_data_label = p.DataLabel
'
'        point_data_label.ShowValue = True
'        point_data_label.AutoText = True
'
'
'        point_data_label.Orientation = xlHorizontal
'
'        ' Определяем формат лейблов
'        If InStr(1, y_axis_number_format, "%") > 0 Then
'            If y_max < 10 Then
'                point_data_label.NumberFormat = "0.0%;-0.0%;0%"
'            Else
'                point_data_label.NumberFormat = "0%;-0%;0%"
'            End If
'        ElseIf y_axis.MaximumScale <= 1 Then
'            point_data_label.NumberFormat = "0.00;-0;"""""
'        ElseIf y_axis.MaximumScale <= 30 And Not все_y_целые Then
'            point_data_label.NumberFormat = "0.0;-0;"""""
'        Else
'            point_data_label.NumberFormat = "# ##0;-0;"""""
'        End If
'
'        Debug.Print point_data_label.text
'
'        point_data_label.Position = xlLabelPositionAbove
'        's.HasLeaderLines = False
'        point_data_label.Format.TextFrame2.TextRange.Font.Bold = msoTrue
'        point_data_label.Format.TextFrame2.TextRange.Font.size = aРазмерШрифтаПодписейДанных(i)
'
'        With point_data_label.Format.TextFrame2.TextRange.Font.Fill
'            .Visible = msoTrue
'            .ForeColor.RGB = aForeColor(i)
'            .Transparency = 0
'            .Solid
'        End With
'    End If

End Sub



'Sub Раскрасить_график()
'    Dim aForeColor() As Long: ReDim aForeColor(10)
'    Dim aBackColor() As Long: ReDim aBackColor(10)
'    Dim aLineWeight() As Double: ReDim aLineWeight(10)
'    Dim aLineVisible() As Long: ReDim aLineVisible(10)
'    Dim aMarkerStyle() As Long: ReDim aMarkerStyle(10)
'    Dim aMarkerSize() As Long: ReDim aMarkerSize(10)
'    Dim aLineStyle() As Long: ReDim aLineStyle(10)
'    Dim aПоказыватьМаксЗнач() As Long: ReDim aПоказыватьМаксЗнач(10)
'    Dim aНомерПервойТочкиСКоторойИщемМаксЗнач() As Long: ReDim aНомерПервойТочкиСКоторойИщемМаксЗнач(10)
'
'
'    Dim aПодписиДанных() As Long: ReDim aПодписиДанных(10)
'    Dim aРазмерШрифтаПодписейДанных() As Double: ReDim aРазмерШрифтаПодписейДанных(10)
'
'
'    Dim rng As Range
'    Dim cell As Range
'
'
'    Dim i As Long
'
'
'
'
'
'    Set rng = Range("K4:K13")
'    i = 1
'    For Each cell In rng
'        aForeColor(i) = cell.Interior.color
'        i = i + 1
'    Next cell
'
'    Set rng = Range("L4:L13")
'    i = 1
'    For Each cell In rng
'        aBackColor(i) = cell.Interior.color
'        i = i + 1
'    Next cell
'
'    Set rng = Range("M4:M13")
'    i = 1
'    For Each cell In rng
'        aLineWeight(i) = cell.Value
'        i = i + 1
'    Next cell
'
'    Set rng = Range("N4:N13")
'    i = 1
'    For Each cell In rng
'        aLineVisible(i) = cell.Value
'        i = i + 1
'    Next cell
'
'    Set rng = Range("O4:O13")
'    i = 1
'    For Each cell In rng
'        aMarkerStyle(i) = cell.Value
'        i = i + 1
'    Next cell
'
'    Set rng = Range("P4:P13")
'    i = 1
'    For Each cell In rng
'        aMarkerSize(i) = cell.Value
'        i = i + 1
'    Next cell
'
'    Set rng = Range("S4:S13")
'    i = 1
'    For Each cell In rng
'        aLineStyle(i) = cell.Value
'        i = i + 1
'    Next cell
'
'
'    Set rng = Range("T4:T13")
'    i = 1
'    For Each cell In rng
'        aПодписиДанных(i) = cell.Value
'        i = i + 1
'    Next cell
'
'
'    Set rng = Range("U4:U13")
'    i = 1
'    For Each cell In rng
'        aРазмерШрифтаПодписейДанных(i) = cell.Value
'        i = i + 1
'    Next cell
'
'    Set rng = Range("V4:V13")
'    i = 1
'    For Each cell In rng
'        aПоказыватьМаксЗнач(i) = cell.Value
'        i = i + 1
'    Next cell
'
'    Set rng = Range("W4:W13")
'    i = 1
'    For Each cell In rng
'        aНомерПервойТочкиСКоторойИщемМаксЗнач(i) = cell.Value
'        i = i + 1
'    Next cell
'
'
'
'    Dim cho As ChartObject
'    Dim fsc As FullSeriesCollection
'    Dim s As Series
'
'
'
'    For Each cho In ActiveSheet.ChartObjects
'
'        If cho.Chart.ChartType = xlColumnClustered Then GoTo 1
'        If cho.Chart.ChartType = -4111 Then GoTo 1
'
'        i = 1
'        For Each s In cho.Chart.FullSeriesCollection
'
'            s.Format.Line.ForeColor.RGB = aForeColor(i)
'
'            s.Format.Line.Weight = aLineWeight(i)
'
'            s.Format.Fill.ForeColor.RGB = aBackColor(i)
'
'            s.MarkerSize = aMarkerSize(i)
'
'
'
'            If aLineVisible(i) = 1 Then
'                s.Format.Line.Visible = msoTrue
'            Else
'                s.Format.Line.Visible = msoFalse
'            End If
'
'
'            s.MarkerStyle = aMarkerStyle(i)
'
'
'            s.Format.Line.DashStyle = aLineStyle(i)
'            s.Format.Line.style = msoLineSingle
'
'
'            ' Удаляем все даталейблы
'            Dim p As Point
'            Dim n As Long: n = UBound(s.values)
'            Dim j As Long
'            For j = 1 To n
'                Set p = s.Points(j)
'                If p.HasDataLabel Then p.DataLabel.Delete
'            Next j
'
'
'
'            ' Добавляем подпись самого большого значения
'            Dim y_max As Double: y_max = 0
'            Dim y As Double
'            Dim j_max As Long: j_max = n
'            Dim все_y_целые As Boolean: все_y_целые = True
'
'            For j = aНомерПервойТочкиСКоторойИщемМаксЗнач(i) To n
'                y = s.values(j)
'
'                If y >= y_max Then
'                    y_max = y
'                    j_max = j
'                End If
'
'
'                If Not y = Round(y, 0) Then все_y_целые = False
'            Next j
'
'
'            Dim y_axis As Axis: Set y_axis = cho.Chart.Axes(xlValue)
'            y_axis_number_format = y_axis.TickLabels.NumberFormat
'            Dim y_ax_has_dot As Boolean
'            If InStr(1, y_axis_number_format, ".") > 0 Or InStr(1, y_axis_number_format, ",") > 0 Then
'                y_ax_has_dot = True
'            Else
'                y_ax_has_dot = False
'            End If
'
'
'
'            ' Подписываем все точки
'            Dim data_labels As DataLabels
'            If aПодписиДанных(i) = 1 Then
'                s.ApplyDataLabels
'                Set data_labels = s.DataLabels
'                data_labels.Orientation = xlUpward
'
'                ' Определяем формат лейблов
'                If InStr(1, y_axis_number_format, "%") > 0 Then
'                    If y_max < 10 Then
'                        data_labels.NumberFormat = "0.0%;-0.0%;0%"
'                    Else
'                        data_labels.NumberFormat = "0%;-0%;0%"
'                    End If
'                ElseIf y_axis.MaximumScale <= 1 Then
'                    data_labels.NumberFormat = "0.00;-0;"""""
'                ElseIf y_axis.MaximumScale <= 30 And Not все_y_целые Then
'                    data_labels.NumberFormat = "0.0;-0;"""""
'                Else
'                    data_labels.NumberFormat = "# ##0;-0;"""""
'                End If
'
'
'                data_labels.Position = xlLabelPositionAbove
'                s.HasLeaderLines = False
'                data_labels.Format.TextFrame2.TextRange.Font.Bold = msoTrue
'                data_labels.Format.TextFrame2.TextRange.Font.size = aРазмерШрифтаПодписейДанных(i)
'
'                With data_labels.Format.TextFrame2.TextRange.Font.Fill
'                    .Visible = msoTrue
'                    .ForeColor.RGB = aForeColor(i)
'                    .Transparency = 0
'                    .Solid
'                End With
'
'            Else
'                If s.HasDataLabels Then
'                    Set data_labels = s.DataLabels
'                    data_labels.Delete
'                End If
'            End If
'
'
'
'
'
'
'            If y_max > 0 And aПоказыватьМаксЗнач(i) = 1 Then
'                Set p = s.Points(j_max)
'                p.ApplyDataLabels
'
'                Dim point_data_label As DataLabel: Set point_data_label = p.DataLabel
'
'                point_data_label.ShowValue = True
'                point_data_label.AutoText = True
'
'
'                point_data_label.Orientation = xlHorizontal
'
'                ' Определяем формат лейблов
'                If InStr(1, y_axis_number_format, "%") > 0 Then
'                    If y_max < 10 Then
'                        point_data_label.NumberFormat = "0.0%;-0.0%;0%"
'                    Else
'                        point_data_label.NumberFormat = "0%;-0%;0%"
'                    End If
'                ElseIf y_axis.MaximumScale <= 1 Then
'                    point_data_label.NumberFormat = "0.00;-0;"""""
'                ElseIf y_axis.MaximumScale <= 30 And Not все_y_целые Then
'                    point_data_label.NumberFormat = "0.0;-0;"""""
'                Else
'                    point_data_label.NumberFormat = "# ##0;-0;"""""
'                End If
'
'                Debug.Print point_data_label.text
'
'                point_data_label.Position = xlLabelPositionAbove
'                's.HasLeaderLines = False
'                point_data_label.Format.TextFrame2.TextRange.Font.Bold = msoTrue
'                point_data_label.Format.TextFrame2.TextRange.Font.size = aРазмерШрифтаПодписейДанных(i)
'
'                With point_data_label.Format.TextFrame2.TextRange.Font.Fill
'                    .Visible = msoTrue
'                    .ForeColor.RGB = aForeColor(i)
'                    .Transparency = 0
'                    .Solid
'                End With
'            End If
'
'
'            i = i + 1
'
'        Next s
'
'1:
'
'    Next cho
'
'
'End Sub
'
