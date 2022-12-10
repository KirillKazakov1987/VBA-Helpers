Attribute VB_Name = "ARRAY_TO_EXCEL_RANGE"
Option Explicit
Option Base 1

Private Const ERR_INTERNAL_CODE = 999999



Public Sub Write_2D_array_to_excel_range(arr As Variant, dst_rng As Range)
    If Not Get_array_rank(arr) = 2 Then Err.Raise ERR_INTERNAL_CODE
    
    Dim n1 As Long, n2 As Long
    n1 = LBound(arr, 1): n2 = UBound(arr, 1)

    Dim m1 As Long, m2 As Long
    m1 = LBound(arr, 2): m2 = UBound(arr, 2)

    Dim cell_TL As Range: Set cell_TL = dst_rng.Cells(1, 1)
    Dim cell_BR As Range: Set cell_BR = dst_rng.Cells(n2 - n1 + 1, m2 - m1 + 1)
    
    Dim r As Range: Set r = Range(cell_TL, cell_BR)

    r.value = arr
End Sub



Private Function Get_array_rank(arr As Variant) As Long
    If IsArray(arr) = False Then
        Get_array_rank = 0
    Else
        Dim i As Long: i = 0
        Dim n As Long
        On Error Resume Next
            Do While (Err = 0)
                i = i + 1
                n = UBound(arr, i)
            Loop
        On Error GoTo 0
        Get_array_rank = i - 1
    End If
End Function
