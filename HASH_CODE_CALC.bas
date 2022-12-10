Attribute VB_Name = "HASH_CODE_CALC"
Option Explicit



Public Function Get_hash_code_of_string(s As String) As Long
    Dim n As Long
    n = Len(s)
    
    Dim char_code As Integer
    
    Dim result As Long
    result = n
    
    Dim seed1 As Integer
    seed1 = 113
    
    Dim seed2 As Long
    seed2 = 17
    
    Dim i As Long
    For i = 1 To n
        Dim ch As String
        ch = Mid$(s, i, 1)
        
        char_code = AscW(ch)
        
        result = result + (char_code Xor seed1) - (i Xor seed2)
        
    Next i
    
    Get_hash_code_of_string = result


End Function
