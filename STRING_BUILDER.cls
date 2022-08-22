VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "STRING_BUILDER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private Chunks() As String
Private Capacity_of_chunks As Long
Private Count_chunks As Long
Private Count_chars As Long

Private Sub Class_Initialize()
    Capacity_of_chunks = 8
    Count_chunks = 0
    Count_chars = 0
    ReDim Chunks(0 To Capacity_of_chunks - 1)
End Sub


Private Function Get_free_chunk_index() As Long
    If Count_chunks >= Capacity_of_chunks Then
        Dim capacity_shift As Long: capacity_shift = Capacity_of_chunks
        capacity_shift = MATH_HELPER.Clamp_i32(capacity_shift, 8, 256)
        Capacity_of_chunks = Capacity_of_chunks + capacity_shift
        ReDim Preserve Chunks(0 To Capacity_of_chunks - 1)
    End If
    
    Get_free_chunk_index = Count_chunks
End Function


Public Sub Append(text As String)
    Dim N As Long: N = Len(text)
    
    If N > 0 Then
        Count_chars = Count_chars + N
        Count_chunks = Count_chunks + 1
        
        Dim i As Long: i = Get_free_chunk_index()
        
        Chunks(i) = text
        
    End If

End Sub


Public Sub Append_line(text As String)
    Append text & vbNewLine
End Sub



Public Function Get_string() As String
    Dim accum As String
    accum = ""
    
    Dim i As Long: For i = 0 To Count_chunks
        accum = accum & Chunks(i)
    Next i
    
    Get_string = accum
End Function

