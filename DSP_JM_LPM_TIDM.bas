Attribute VB_Name = "DSP_JM_LPM_TIDM"
Option Explicit


Public Function DSP_TIDM_Cal( _
       ByVal in_CapturedWave_DSP As DSPWave, _
       ByVal in_DataBits_lng As Long, _
       ByVal in_ValidBits_lng As Long, _
       ByVal i_ValidIndex_DSP As DSPWave, _
       ByVal i_tidmStartIndex_DSP As DSPWave, _
       ByVal i_stdcellStartIndex_DSP As DSPWave, _
       ByRef out_valid_DSP As DSPWave, _
       ByRef out_tidm_DSP As DSPWave, _
       ByRef out_stdcell_DSP As DSPWave) As Long
    

    Dim i                   As Long
    Dim i_SampleSize_lng    As Long
    
    i_SampleSize_lng = in_CapturedWave_DSP.SampleSize / 21
    
    out_valid_DSP.CreateConstant 0, i_SampleSize_lng, DspLong
    out_tidm_DSP.CreateConstant 0, i_SampleSize_lng, DspLong
    out_stdcell_DSP.CreateConstant 0, i_SampleSize_lng, DspLong
    
    On Error Resume Next    ' DSP does not support <On Error Goto ...>, but you can use <On Error Resume Next>
 
    For i = 0 To i_SampleSize_lng - 1
        out_valid_DSP.Element(i) = in_CapturedWave_DSP.Select(i_ValidIndex_DSP.Element(i), 1, in_ValidBits_lng).Element(0)
        out_tidm_DSP.Element(i) = in_CapturedWave_DSP.Select(i_tidmStartIndex_DSP.Element(i), 1, in_DataBits_lng).ConvertStreamTo(tldspParallel, in_DataBits_lng, 0, Bit0IsMsb).Element(0)
        out_stdcell_DSP.Element(i) = in_CapturedWave_DSP.Select(i_stdcellStartIndex_DSP.Element(i), 1, in_DataBits_lng).ConvertStreamTo(tldspParallel, in_DataBits_lng, 0, Bit0IsMsb).Element(0)
    Next i
  
End Function


Public Function DSP_LPM_Cal( _
       ByVal in_CapturedWave_DSP As DSPWave, _
       ByVal in_LoopCount_lng As Long, _
       ByVal in_DataBits_lng As Long, _
       ByRef out_Data_DSP As DSPWave) As Long
    
    Dim i                   As Long
    Dim j                   As Long
    Dim i_SampleSize_lng    As Long
    
    i_SampleSize_lng = in_CapturedWave_DSP.SampleSize
    
     ' concatenate all the captured wave segment
    out_Data_DSP.CreateConstant 0, 0, DspLong
    
    On Error Resume Next    ' DSP does not support <On Error Goto ...>, but you can use <On Error Resume Next>
    
     For i = 0 To in_LoopCount_lng
        If i > 0 Then in_CapturedWave_DSP.Next
            out_Data_DSP = out_Data_DSP.Concatenate(in_CapturedWave_DSP)
    Next i
    
       ' check integrity of captured wave
    If Not (out_Data_DSP.SampleSize = (in_LoopCount_lng + 1) * i_SampleSize_lng) Then
        DSP_LPM_Cal = 404
        out_Data_DSP.CreateConstant 999, i_SampleSize_lng * (in_LoopCount_lng + 1), DspLong
        Exit Function
    End If
    

End Function

Public Function DSP_JM_Calc(ByVal CaptureWave As DSPWave, _
                            ByVal in_LoopCount_lng As Long, _
                            ByRef overflow_dsp As DSPWave, _
                            ByRef underflow_dsp As DSPWave, _
                            ByRef dtc_meadone_dsp As DSPWave, _
                            ByRef a_dsp As DSPWave, _
                            ByRef b_dsp As DSPWave, _
                            ByRef c_dsp As DSPWave, _
                            ByRef d_dsp As DSPWave, _
                            ByRef e_dsp As DSPWave, _
                            ByRef min_dsp As DSPWave, _
                            ByRef max_dsp As DSPWave, _
                            ByRef jm_dsp As DSPWave) As Long
                        


' MSB first mode
'BirPerWord is a const (this value in 6280 is 32)


Dim Slot As Long

Dim Temp_a_dsp As New DSPWave
Dim Temp_b_dsp As New DSPWave
Dim Temp_c_dsp As New DSPWave
Dim Temp_d_dsp As New DSPWave

Dim Temp_overflow_dsp As New DSPWave
Dim Temp_underflow_dsp As New DSPWave
Dim Temp_dtc_meadone_dsp As New DSPWave

Dim Temp_e_dsp As New DSPWave
Dim Temp_min_dsp As New DSPWave
Dim Temp_max_dsp As New DSPWave

Dim hex_a_dsp As New DSPWave
Dim hex_b_dsp As New DSPWave
Dim hex_c_dsp As New DSPWave
Dim hex_d_dsp As New DSPWave

Dim hex_overflow_dsp As New DSPWave
Dim hex_underflow_dsp As New DSPWave
Dim hex_dtc_meadone_dsp As New DSPWave

Dim hex_e_dsp As New DSPWave
Dim hex_min_dsp As New DSPWave
Dim hex_max_dsp As New DSPWave

Dim i_EntireCapture_DSP As New DSPWave

Dim i As Long

Dim slot_half As Long


i_EntireCapture_DSP.CreateConstant 0, 0, DspLong

On Error Resume Next    ' DSP does not support <On Error Goto ...>, but you can use <On Error Resume Next>


For i = 0 To in_LoopCount_lng
    If i > 0 Then CaptureWave.Next
    i_EntireCapture_DSP = i_EntireCapture_DSP.Concatenate(CaptureWave)
Next i



Slot = i_EntireCapture_DSP.SampleSize

jm_dsp.CreateConstant 9999, Slot / 2, DspDouble

Temp_a_dsp.CreateConstant 999, Slot, DspLong
Temp_b_dsp.CreateConstant 999, Slot, DspLong
Temp_c_dsp.CreateConstant 999, Slot, DspLong
Temp_d_dsp.CreateConstant 999, Slot, DspLong

Temp_overflow_dsp.CreateConstant 999, Slot, DspLong
Temp_underflow_dsp.CreateConstant 999, Slot, DspLong
Temp_dtc_meadone_dsp.CreateConstant 999, Slot, DspLong

Temp_e_dsp.CreateConstant 999, Slot, DspLong
Temp_min_dsp.CreateConstant 999, Slot, DspLong
Temp_max_dsp.CreateConstant 999, Slot, DspLong


Temp_a_dsp = i_EntireCapture_DSP.BitwiseAnd(&HFF00).BitwiseShiftRight(8)
Temp_b_dsp = i_EntireCapture_DSP.BitwiseAnd(&HFF)
Temp_c_dsp = i_EntireCapture_DSP.BitwiseAnd(&HFF000000).BitwiseShiftRight(24)
Temp_d_dsp = i_EntireCapture_DSP.BitwiseAnd(&HFF0000).BitwiseShiftRight(16)

Temp_e_dsp = i_EntireCapture_DSP.BitwiseAnd(&HFF00).BitwiseShiftRight(8)
Temp_min_dsp = i_EntireCapture_DSP.BitwiseAnd(&HFF0000).BitwiseShiftRight(16)
Temp_max_dsp = i_EntireCapture_DSP.BitwiseAnd(&HFF000000).BitwiseShiftRight(24)
Temp_overflow_dsp = i_EntireCapture_DSP.BitwiseAnd(&H8).BitwiseShiftRight(3)
Temp_underflow_dsp = i_EntireCapture_DSP.BitwiseAnd(&H4).BitwiseShiftRight(2)
Temp_dtc_meadone_dsp = i_EntireCapture_DSP.BitwiseAnd(&H1)
        
For i = 0 To Slot - 1
    If i Mod 2 = 0 Then
        Temp_e_dsp.Element(i) = 0
        Temp_min_dsp.Element(i) = 0
        Temp_max_dsp.Element(i) = 0
        Temp_overflow_dsp.Element(i) = 0
        Temp_underflow_dsp.Element(i) = 0
        Temp_dtc_meadone_dsp.Element(i) = 0
    Else
        Temp_a_dsp.Element(i) = 0
        Temp_b_dsp.Element(i) = 0
        Temp_c_dsp.Element(i) = 0
        Temp_d_dsp.Element(i) = 0
        
        If (Temp_d_dsp.Element(i - 1) - Temp_a_dsp.Element(i - 1)) * (Temp_b_dsp.Element(i - 1) - Temp_c_dsp.Element(i - 1)) + Temp_b_dsp.Element(i - 1) - Temp_e_dsp.Element(i) <> 0 Then
            jm_dsp.Element((i - 1) / 2) = (Temp_max_dsp.Element(i) - Temp_min_dsp.Element(i)) / (Sqr(2) * ((Temp_d_dsp.Element(i - 1) - Temp_a_dsp.Element(i - 1)) * (Temp_b_dsp.Element(i - 1) - Temp_c_dsp.Element(i - 1)) + Temp_b_dsp.Element(i - 1) - Temp_e_dsp.Element(i)))
            If jm_dsp.Element((i - 1) / 2) = 0 Then
                jm_dsp.Element((i - 1) / 2) = 0
            End If
        Else
            jm_dsp.Element((i - 1) / 2) = 999
        End If
            
    End If
    
Next i

a_dsp.CreateConstant 999, Slot / 2, DspLong
b_dsp.CreateConstant 999, Slot / 2, DspLong
c_dsp.CreateConstant 999, Slot / 2, DspLong
d_dsp.CreateConstant 999, Slot / 2, DspLong
e_dsp.CreateConstant 999, Slot / 2, DspLong
min_dsp.CreateConstant 999, Slot / 2, DspLong
max_dsp.CreateConstant 999, Slot / 2, DspLong
overflow_dsp.CreateConstant 999, Slot / 2, DspLong
underflow_dsp.CreateConstant 999, Slot / 2, DspLong
dtc_meadone_dsp.CreateConstant 999, Slot / 2, DspLong


a_dsp = Temp_a_dsp.Select(0, 2, Slot / 2).Copy
b_dsp = Temp_b_dsp.Select(0, 2, Slot / 2).Copy
c_dsp = Temp_c_dsp.Select(0, 2, Slot / 2).Copy
d_dsp = Temp_d_dsp.Select(0, 2, Slot / 2).Copy
overflow_dsp = Temp_overflow_dsp.Select(1, 2, Slot / 2).Copy
underflow_dsp = Temp_underflow_dsp.Select(1, 2, Slot / 2).Copy
dtc_meadone_dsp = Temp_dtc_meadone_dsp.Select(1, 2, Slot / 2).Copy
e_dsp = Temp_e_dsp.Select(1, 2, Slot / 2).Copy
min_dsp = Temp_min_dsp.Select(1, 2, Slot / 2).Copy
max_dsp = Temp_max_dsp.Select(1, 2, Slot / 2).Copy

slot_half = Slot / 2

For i = 0 To slot_half - 1

    a_dsp.Element(i) = Temp_a_dsp.Element(2 * i)
    b_dsp.Element(i) = Temp_b_dsp.Element(2 * i)
    c_dsp.Element(i) = Temp_c_dsp.Element(2 * i)
    d_dsp.Element(i) = Temp_d_dsp.Element(2 * i)

    overflow_dsp.Element(i) = Temp_overflow_dsp.Element(2 * i + 1)
    underflow_dsp.Element(i) = Temp_underflow_dsp.Element(2 * i + 1)
    dtc_meadone_dsp.Element(i) = Temp_dtc_meadone_dsp.Element(2 * i + 1)
    
    e_dsp.Element(i) = Temp_e_dsp.Element(2 * i + 1)
    min_dsp.Element(i) = Temp_min_dsp.Element(2 * i + 1)
    max_dsp.Element(i) = Temp_max_dsp.Element(2 * i + 1)

    '' 1. check if overflow =0, if not,output error value 9999
    If overflow_dsp.Element(i) <> 0 Then
        overflow_dsp.Element(i) = 9999
        jm_dsp.Element(i) = 9999
    End If
    '' 2. check if underflow =0, if not,output error value 9999
    If underflow_dsp.Element(i) <> 0 Then
        underflow_dsp.Element(i) = 9999
        jm_dsp.Element(i) = 9999
    End If
    '' 3. check if dtc_meadone =1, if not,output error value 9999
    If dtc_meadone_dsp.Element(i) <> 1 Then
        dtc_meadone_dsp.Element(i) = 9999
        jm_dsp.Element(i) = 9999
    End If
    '' 4. check if code max > code e > code min, if not,output error value 9999
    If max_dsp.Element(i) - e_dsp.Element(i) <= 0 Or max_dsp.Element(i) - min_dsp.Element(i) <= 0 Or e_dsp.Element(i) - min_dsp.Element(i) <= 0 Then
        max_dsp.Element(i) = 9999
        min_dsp.Element(i) = 9999
        e_dsp.Element(i) = 9999
        jm_dsp.Element(i) = 9999
    End If
    '' 5. check if code C > code B, if not,output error value 9999
    If c_dsp.Element(i) - b_dsp.Element(i) <= 0 Then
         b_dsp.Element(i) = 9999
         c_dsp.Element(i) = 9999
        jm_dsp.Element(i) = 9999
    End If
    '' 6. check if code A > code D, if not,output error value 9999
    If a_dsp.Element(i) - d_dsp.Element(i) <= 0 Then
         a_dsp.Element(i) = 9999
         d_dsp.Element(i) = 9999
        jm_dsp.Element(i) = 9999
    End If
    
        
    
Next i


End Function

