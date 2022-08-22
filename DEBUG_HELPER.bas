Attribute VB_Name = "DEBUG_HELPER"
Option Explicit
Option Base 0

Public Sub Print_into_debug( _
        v As Variant, _
        Optional fmt As String = "")
        
    If IsArray(v) Then
        Dim r As Long: r = ARRAY_HELPER.Get_array_rank(v)
        
        If r = 1 Then
            Print_1D_array v, fmt
        ElseIf r = 2 Then
            Print_2D_array v, fmt
        Else
            Debug.Print "Can't implicitly print value(s) of type: " & TypeName(v) & "!"
        End If
    Else
        
        Debug.Print Format(v, fmt)
        
    End If
    
    
End Sub



Public Sub ccccccccc()
    Dim a1() As Long: ReDim a1(0 To 5)
    a1(0) = 1100000
    a1(1) = 2200000
    a1(2) = 3300000
    a1(3) = -1
    a1(4) = -2
    a1(5) = -999999999
    
    Dim a2() As Variant: ReDim a2(0 To 4, 0 To 4)
    a2(0, 0) = 100100
    a2(0, 1) = 100200
    a2(1, 0) = 200100
    a2(1, 1) = 200200
    a2(3, 3) = 123456789
    a2(2, 2) = "ABCzyx"
    
    Dim a3() As String: ReDim a3(0 To 2)
    a3(0) = "abc abc abc abc abc"
    a3(1) = "zzzzzzzzzzz"
    a3(2) = "aaa"


    Print_into_debug a1, "0"
    
    Print_into_debug a2
    
    Print_into_debug a3
End Sub


Private Sub Print_1D_array( _
        values As Variant, _
        Optional fmt As String = "")
        
    Debug.Print ARRAY_TO_TEXT.Print_1D_array(values, fmt)
    
End Sub


Private Sub Print_2D_array( _
        values As Variant, _
        Optional fmt As String = "")
        
    Debug.Print ARRAY_TO_TEXT.Print_2D_array(values, fmt)
    
End Sub


