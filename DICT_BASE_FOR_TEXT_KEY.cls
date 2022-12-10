VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DICT_BASE_FOR_TEXT_KEY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private case_sens As Boolean
Private cap As Long
Private count_records As Long

Private bucket_first_record_onedexes() As Long ' zero is empty

Private hash_codes() As Long
Private keys() As String
Private next_record_onedexes() As Long


Private Sub Class_Initialize()
    Case_sensitivity = True
    
    count_records = 0
    cap = 7
    
    ReDim bucket_first_record_onedexes(0 To cap - 1)
    ReDim hash_codes(0 To cap - 1)
    ReDim keys(0 To cap - 1)
    ReDim next_record_onedexes(0 To cap - 1)
End Sub



Public Property Get Count() As Boolean
    Count = count_records
End Property

Public Property Get Capacity() As Long
    Capacity = cap
End Property

Private Property Let Capacity(val As Long)
    cap = val
End Property

Public Function Get_key(record_index As Long)
    Debug.Assert record_index >= 0
    Debug.Assert record_index < Count
    
    Get_key = keys(record_index)
End Function


Public Property Get Case_sensitivity() As Boolean
    Case_sensitivity = case_sens
End Property

Public Property Let Case_sensitivity(val As Boolean)
    If count_records > 0 Then Err.Raise 999999
    case_sens = val
End Property




Private Function Check_key_equality(k1 As String, k2 As String) As Boolean
    If Case_sensitivity Then
        Check_key_equality = (k1 = k2)
    Else
        Check_key_equality = (LCase(k1) = LCase(k2))
    End If
End Function


Private Function Get_hash_code(k As String)
    If Case_sensitivity Then
        Get_hash_code = HASH_CODE_CALC.Get_hash_code_of_string(k)
    Else
        Get_hash_code = HASH_CODE_CALC.Get_hash_code_of_string(LCase(k))
    End If
End Function



Public Sub Ensure_capacity(required_capacity As Long)
    If Capacity >= required_capacity Then Exit Sub
    
    Dim new_capacity As Long
    new_capacity = MATH_HELPER.Max_i32(new_capacity, Capacity * 2 - 1)
    Capacity = new_capacity
    
    If Capacity > 100000000 Then Err.Raise 999999


    ReDim Preserve hash_codes(0 To Capacity - 1)
    ReDim Preserve keys(0 To Capacity - 1)
    
    ReDim bucket_first_record_onedexes(0 To Capacity - 1)
    ReDim next_record_onedexes(0 To Capacity - 1)
    
    
    
    Dim i As Long
    For i = 0 To count_records - 1
        Dim hc As Long
        hc = hash_codes(i)
        
        Dim bucket_header_index As Long
        bucket_header_index = Abs(hc) Mod Capacity
        
        Dim bucket_first_record_onedex As Long
        bucket_first_record_onedex = bucket_first_record_onedexes(bucket_header_index)
        
        If bucket_first_record_onedex = 0 Then
            bucket_first_record_onedexes(bucket_header_index) = i + 1
        Else
            Dim rec_onedex As Long
            rec_onedex = bucket_first_record_onedex
            
            Do While (True)
                Dim j As Long
                j = next_record_onedexes(rec_onedex - 1)
                If j = 0 Then
                    next_record_onedexes(rec_onedex - 1) = i + 1
                    Exit Do
                Else
                    rec_onedex = j
                End If
            Loop
        End If
    
    Next i
End Sub


Public Function Contains_key(key As String) As Boolean
    Contains_key = Try_get_record_index(key).Has_value
End Function


Public Function Add_key_and_get_record_index(key As String) As Long

    Dim optional_record_index As Optional_int32
    
    optional_record_index = Try_get_record_index(key)
    
    If optional_record_index.Has_value Then
        Add_key_and_get_record_index = optional_record_index.value
        Exit Function
    End If
    
    
    Ensure_capacity count_records + 1
    
    
    Dim key_hash_code As Long
    key_hash_code = Get_hash_code(key)
    
    Dim bucket_index As Long
    bucket_index = Abs(key_hash_code) Mod Capacity

    Dim onedex As Long: onedex = bucket_first_record_onedexes(bucket_index)
    
    Dim step_counter As Long: step_counter = 0
    
    Dim record_index As Long
    
    If onedex = 0 Then
        count_records = count_records + 1
        onedex = count_records
        
        record_index = onedex - 1

        hash_codes(record_index) = key_hash_code
        keys(record_index) = key
        
        bucket_first_record_onedexes(bucket_index) = onedex


        Add_key_and_get_record_index = record_index
        
        Exit Function
    End If
    
    
    
    Do While (onedex > 0)
    
        step_counter = step_counter + 1
        
        Dim next_record_onedex As Long
        next_record_onedex = next_record_onedexes(onedex - 1)
        
        
        If next_record_onedex = 0 Then
            
            count_records = count_records + 1
        
            record_index = count_records - 1
            
            next_record_onedexes(onedex - 1) = record_index + 1
            
            hash_codes(record_index) = key_hash_code
            
            keys(record_index) = key
            
            Add_key_and_get_record_index = record_index
            
            Exit Function
            
        Else
            onedex = next_record_onedex
        End If
        
        If step_counter > Capacity Then
            Err.Raise 999999
        End If
    Loop
    
End Function



Public Function Try_get_record_index(key As String) As Optional_int32
   
    Dim key_hash_code As Long
    key_hash_code = Get_hash_code(key)
    
    
    Dim bucket_index As Long
    bucket_index = Abs(key_hash_code) Mod Capacity

    Dim bucket_first_record_onedex As Long
    bucket_first_record_onedex = bucket_first_record_onedexes(bucket_index)
    
    
    Dim onedex As Long: onedex = bucket_first_record_onedex
    

    Do While (onedex > 0)
        Dim hash_code As Long
        hash_code = hash_codes(onedex - 1)
       
        If (hash_code = key_hash_code) And Check_key_equality(key, keys(onedex - 1)) Then
            
            Try_get_record_index.Has_value = True
            Try_get_record_index.value = onedex - 1
            
            Exit Function
        Else
            onedex = next_record_onedexes(onedex - 1)
        End If
    Loop
    
    
    Try_get_record_index.Has_value = False
    Try_get_record_index.value = -1
    
End Function


