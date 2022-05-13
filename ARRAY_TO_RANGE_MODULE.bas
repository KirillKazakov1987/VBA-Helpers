Attribute VB_Name = "ARRAY_TO_RANGE_MODULE"
Option Explicit
Option Base 1

Private Const ERR_INTERNAL_CODE = 999999



Public Sub Write_2D_array_to_excel_range(arr As Variant, dst_rng As Range)
    If Not Get_array_rank(arr) = 2 Then Err.Raise ERR_INTERNAL_CODE
    
    Dim N1 As Long, N2 As Long
    N1 = LBound(arr, 1): N2 = UBound(arr, 1)

    Dim M1 As Long, M2 As Long
    M1 = LBound(arr, 2): M2 = UBound(arr, 2)

    Dim cell_TL As Range: Set cell_TL = dst_rng.Cells(1, 1)
    Dim cell_BR As Range: Set cell_BR = dst_rng.Cells(N2 - N1 + 1, M2 - M1 + 1)
    
    Dim r As Range: Set r = Range(cell_TL, cell_BR)

    r.Value = arr
End Sub



Private Function Get_array_rank(arr As Variant) As Long
    If IsArray(arr) = False Then
        Get_array_rank = 0
    Else
        Dim i As Long: i = 0
        Dim N As Long
        On Error Resume Next
            Do While (Err = 0)
                i = i + 1
                N = UBound(arr, i)
            Loop
        On Error GoTo 0
        Get_array_rank = i - 1
    End If
End Function
