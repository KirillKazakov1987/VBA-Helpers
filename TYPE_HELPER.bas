Attribute VB_Name = "TYPE_HELPER"
Option Explicit

Public Const FLOAT64_MAX As Double = 1.79769313486231E+308
Public Const FLOAT64_MIN As Double = -1.79769313486231E+308

Public Const FLOAT32_MAX As Double = -3.402823E+38
Public Const FLOAT32_MIN As Double = 3.402823E+38

Public Const INT32_MAX As Long = 2147483647
Public Const INT32_MIN As Long = -2147483648#

Public Const INT16_MAX As Integer = 32767
Public Const INT16_MIN As Integer = -32768



Public Function Is_boolean(v As Variant)
    Dim result As Boolean
    result = (VarType(v) = vbBoolean)
    Is_boolean = result
End Function


Public Function Is_float32(v As Variant)
    Dim result As Boolean
    result = (VarType(v) = vbSingle)
    Is_float32 = result
End Function



Public Function Is_float64(v As Variant)
    Dim result As Boolean
    result = (VarType(v) = vbDouble)
    Is_float64 = result
End Function



Public Function Is_int16(v As Variant)
    Dim result As Boolean
    result = (VarType(v) = vbLongLong)
    Is_int16 = result
End Function



Public Function Is_int32(v As Variant)
    Dim result As Boolean
    result = (VarType(v) = vbInteger)
    Is_int32 = result
End Function



Public Function Is_int64(v As Variant)
    Dim result As Boolean
    result = (VarType(v) = vbLongLong)
    Is_int64 = result
End Function


Public Function Is_date(v As Variant)
    Dim result As Boolean
    result = (VarType(v) = vbDate)
    Is_date = result
End Function



Public Function Is_string(v As Variant)
    Dim result As Boolean
    result = (VarType(v) = vbString)
    Is_string = result
End Function




Public Function Try_convert_to_int32(v As Variant) As Optional_int32
    On Error GoTo FAILED_IN__Try_convert_to_int32
    
    Select Case VarType(v)
        Case vbLong:
            Try_convert_to_int32.Has_value = True
            Try_convert_to_int32.value = v
            Exit Function
            
        Case vbInteger:
            Try_convert_to_int32.Has_value = True
            Try_convert_to_int32.value = CLng(v)
            Exit Function
            
        Case vbByte:
            Try_convert_to_int32.Has_value = True
            Try_convert_to_int32.value = CLng(v)
            Exit Function
            
        Case vbBoolean:
            Try_convert_to_int32.Has_value = True
            Try_convert_to_int32.value = IIf(CBool(v), 1, 0)
            Exit Function
            
        Case vbDouble:
            Dim f64 As Double
            f64 = v
            If f64 >= INT32_MIN And f64 <= INT32_MAX Then
                Try_convert_to_int32.Has_value = True
                Try_convert_to_int32.value = CLng(f64)
            Else
                Try_convert_to_int32.Has_value = False
            End If
            Exit Function

        Case vbSingle:
            Dim f32 As Single
            f32 = v
            If f32 >= INT32_MIN And f32 <= INT32_MAX Then
                Try_convert_to_int32.Has_value = True
                Try_convert_to_int32.value = CLng(f32)
            Else
                Try_convert_to_int32.Has_value = False
            End If
            Exit Function
            
         Case vbString:
            Dim s As String
            s = v
            
            If IsNumeric(s) Then
                Dim num As Double
                num = CDbl(s)
                
                If num >= INT32_MIN And num <= INT32_MAX Then
                    Try_convert_to_int32.Has_value = True
                    Try_convert_to_int32.value = CLng(num)
                Else
                    Try_convert_to_int32.Has_value = False
                End If
            Else
                Try_convert_to_int32.Has_value = False
            End If

            Exit Function

        Case Else:
            Try_convert_to_int32.Has_value = False
            Exit Function
    End Select
    
    
FAILED_IN__Try_convert_to_int32:
    Try_convert_to_int32.Has_value = False
End Function



Public Function Try_convert_to_float64(v As Variant) As Optional_float64
    On Error GoTo FAILED_IN__Try_convert_to_float64

    Select Case VarType(v)
        Case vbDouble:
            Try_convert_to_float64.Has_value = True
            Try_convert_to_float64.value = v
            Exit Function
            
        Case vbLong:
            Try_convert_to_float64.Has_value = True
            Try_convert_to_float64.value = CDbl(v)
            Exit Function
            
        Case vbSingle:
            Try_convert_to_float64.Has_value = True
            Try_convert_to_float64.value = CDbl(v)
            Exit Function
            
        Case vbInteger:
            Try_convert_to_float64.Has_value = True
            Try_convert_to_float64.value = CDbl(v)
            Exit Function
            
        Case vbByte:
            Try_convert_to_float64.Has_value = True
            Try_convert_to_float64.value = CDbl(v)
            Exit Function
            
        Case vbBoolean:
            Try_convert_to_float64.Has_value = True
            Try_convert_to_float64.value = IIf(CBool(v), 1, 0)
            Exit Function

        Case vbString:
            Dim s As String
            s = v
            
            If IsNumeric(s) Then
                Try_convert_to_float64.Has_value = True
                Try_convert_to_float64.value = CDbl(s)
            Else
                Try_convert_to_float64.Has_value = False
            End If

            Exit Function

        Case Else:
            Try_convert_to_float64.Has_value = False
            Exit Function
    End Select
    
    
FAILED_IN__Try_convert_to_float64:
    Try_convert_to_float64.Has_value = False
End Function




Public Function Try_convert_to_string(v As Variant) As Optional_string
    On Error GoTo FAILED_IN__Try_convert_to_string
    
    Dim result As String
    
    Select Case VarType(v)
        Case vbEmpty:
            Try_convert_to_string.Has_value = True
            Try_convert_to_string.value = ""
            Exit Function
            
        Case vbNull:
            Try_convert_to_string.Has_value = True
            Try_convert_to_string.value = ""
            Exit Function
        
        Case vbArray:
            Try_convert_to_string.Has_value = False
            Exit Function
            
        Case vbUserDefinedType:
            Try_convert_to_string.Has_value = False
            Exit Function
            
        Case Else:
            Try_convert_to_string.Has_value = True
            Try_convert_to_string.value = CStr(v)
            Exit Function
    End Select
    
    
FAILED_IN__Try_convert_to_string:
    Try_convert_to_string.Has_value = False
End Function





Public Function Try_convert_to_bool(v As Variant) As Optional_bool
    On Error GoTo FAILED_IN__Try_convert_to_bool
    
    Select Case VarType(v)
        Case vbLong:
            Try_convert_to_bool.Has_value = True
            Try_convert_to_bool.value = Not (CLng(v) = 0)
            Exit Function
            
        Case vbInteger:
            Try_convert_to_bool.Has_value = True
            Try_convert_to_bool.value = Not (CInt(v) = 0)
            Exit Function
            
        Case vbByte:
            Try_convert_to_bool.Has_value = True
            Try_convert_to_bool.value = Not (CByte(v) = 0)
            Exit Function
            
        Case vbBoolean:
            Try_convert_to_bool.Has_value = True
            Try_convert_to_bool.value = v
            Exit Function
            
        Case vbDouble:
            Try_convert_to_bool.Has_value = True
            Try_convert_to_bool.value = Not (CDbl(v) = 0)
            Exit Function

        Case vbSingle:
            Try_convert_to_bool.Has_value = True
            Try_convert_to_bool.value = Not (CSng(v) = 0)
            Exit Function
            
         Case vbString:
            Dim s As String
            s = v
            
            If IsNumeric(s) Then
                Try_convert_to_bool.Has_value = True
                Try_convert_to_bool.value = Not (CDbl(s) = 0)
            Else
                If UCase(s) = "ÈÑÒÈÍÀ" _
                    Or UCase(s) = "TRUE" _
                    Or UCase(s) = "YES" _
                    Or UCase(s) = "ÄÀ" _
                Then
                    Try_convert_to_bool.Has_value = True
                    Try_convert_to_bool.value = True
                
                ElseIf UCase(s) = "ÈÑÒÈÍÀ" _
                    Or UCase(s) = "TRUE" _
                    Or UCase(s) = "YES" _
                    Or UCase(s) = "ÄÀ" _
                Then
                    Try_convert_to_bool.Has_value = True
                    Try_convert_to_bool.value = False
                Else
                    Try_convert_to_bool.Has_value = False
                End If
            
                Exit Function
            End If
        
        Case Else:
            Try_convert_to_bool.Has_value = False
            Exit Function
    End Select
    
    
FAILED_IN__Try_convert_to_bool:
    Try_convert_to_bool.Has_value = False
End Function
