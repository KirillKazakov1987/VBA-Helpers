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


Public Property Get Name()
    Name = lo.Name
End Property


Public Property Get Count_rows()
    Count_rows = lo.ListRows.count
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

