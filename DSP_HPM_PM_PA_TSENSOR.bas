Attribute VB_Name = "DSP_HPM_PM_PA_TSENSOR"
Option Explicit

' This module should be used only for DSP Procedure code.  Functions in this
' module will be available to be called to perform DSP in all DSP modes.
' Additional modules may be added as needed (all starting with "DSP_").
'
' The required signature for a DSP Procedure is:
'
' Public Function FuncName(<arglist>) as Long
'   where <arglist> is any list of arguments supported by DSP code.
'
' See online help for supported types and other restrictions.
'
' Notes from TaiShan:
'   1. in_LoopCount_lng is intentionally greater than the actual looping count by (1).
'   2. Don't Trust anyone, including yourself of one hour ago.
'   3.
'

Public Function E_01_HPM_Cal( _
       ByVal in_CapturedWave_DSP As DSPWave, _
       ByVal in_LoopCount_lng As Long, _
       ByVal in_DataBits_lng As Long, _
       ByRef out_Data_DSP As DSPWave, _
       ByVal in_CheckValidFlag_bool As Boolean, _
       ByVal in_ValidBits_lng As Long, _
       ByRef out_valid_DSP As DSPWave) As Long
' get data array and valid bits, called in Function: <HPM_ReadCode_vbt>

    Dim i As Long
    Dim i_EntireCapture_DSP As New DSPWave

    ' concatenate all the captured wave segment
    i_EntireCapture_DSP.CreateConstant 0, 0, DspLong
    On Error Resume Next    ' DSP does not support <On Error Goto ...>, but you can use <On Error Resume Next>
    For i = 0 To in_LoopCount_lng
        If i > 0 Then in_CapturedWave_DSP.Next
        i_EntireCapture_DSP = i_EntireCapture_DSP.Concatenate(in_CapturedWave_DSP)
    Next

    ' check integrity of captured wave
    If Not (i_EntireCapture_DSP.SampleSize = (in_LoopCount_lng + 1) * in_CapturedWave_DSP.SampleSize) Then
        E_01_HPM_Cal = 404
        out_Data_DSP.CreateConstant 999, 999, DspLong
        out_valid_DSP.CreateConstant 999, 999, DspLong
        Exit Function
    End If
        
    ' calculate them
    If in_CheckValidFlag_bool = True Then
        out_Data_DSP = i_EntireCapture_DSP.BitwiseAnd(2 ^ in_DataBits_lng - 1)
        out_valid_DSP = i_EntireCapture_DSP.BitwiseShiftRight(in_DataBits_lng).BitwiseAnd(2 ^ in_ValidBits_lng - 1)
    Else
        out_Data_DSP = i_EntireCapture_DSP.BitwiseAnd(2 ^ in_DataBits_lng - 1)
        out_valid_DSP.CreateConstant 0, out_Data_DSP.SampleSize, DspLong
    End If

End Function


Public Function E_01_HPM_Cal_All( _
       ByVal in_CapturedWave_DSP As DSPWave, _
       ByVal in_LoopCount_lng As Long, _
       ByVal In_RegIndex_DSP As DSPWave, _
       ByVal In_ValidIndex_DSP As DSPWave, _
       ByVal in_DataBits_lng As Long, _
       ByRef out_Data_DSP As DSPWave, _
       ByVal in_ValidBits_lng As Long, _
       ByRef out_valid_DSP As DSPWave) As Long
' get data array and valid bits, called in Function: <HPM_ReadCode_vbt>

    Dim i As Long
    Dim j As Long
    Dim i_EntireCapture_DSP As New DSPWave
    Dim i_CapSampleSize_lng As Long

    Dim i_tmpdata_DSP As New DSPWave
    Dim i_tmpvalid_DSP As New DSPWave

    ' concatenate all the captured wave segment
    i_EntireCapture_DSP.CreateConstant 0, 0, DspLong
    out_Data_DSP.CreateConstant 0, 0, DspLong
    out_valid_DSP.CreateConstant 0, 0, DspLong

    i_tmpdata_DSP.CreateConstant 0, 50 * (in_LoopCount_lng + 1), DspLong
    i_tmpvalid_DSP.CreateConstant 0, 50 * (in_LoopCount_lng + 1), DspLong

    On Error Resume Next    ' DSP does not support <On Error Goto ...>, but you can use <On Error Resume Next>
    i_CapSampleSize_lng = in_CapturedWave_DSP.SampleSize
    For i = 0 To in_LoopCount_lng
        If i > 0 Then in_CapturedWave_DSP.Next
        i_EntireCapture_DSP = i_EntireCapture_DSP.Concatenate(in_CapturedWave_DSP)
        
    Next

    ' check integrity of captured wave
    If Not (i_EntireCapture_DSP.SampleSize = (in_LoopCount_lng + 1) * i_CapSampleSize_lng) Then
        E_01_HPM_Cal_All = 404
        out_Data_DSP.CreateConstant 999, 50 * (in_LoopCount_lng + 1), DspLong
        out_valid_DSP.CreateConstant 999, 50 * (in_LoopCount_lng + 1), DspLong
        Exit Function
    End If


    ' modify for denver hpm

    For i = 0 To in_LoopCount_lng
        For j = 0 To 49
            i_tmpdata_DSP.Element(j + 50 * i) = i_EntireCapture_DSP.Select _
                                                ((In_RegIndex_DSP.Element(j) + 550 * i), 1, in_DataBits_lng) _
                                                .ConvertStreamTo(tldspParallel, in_DataBits_lng, 0, Bit0IsMsb).Element(0)
            i_tmpvalid_DSP.Element(j + 50 * i) = i_EntireCapture_DSP.Select _
                                                 ((In_ValidIndex_DSP.Element(j) + 550 * i), 1, in_ValidBits_lng).Element(0)
        Next j
    Next i

    out_Data_DSP = i_tmpdata_DSP.Copy
    out_valid_DSP = i_tmpvalid_DSP.Copy

End Function

''''
''''''PASENSOR_CPUB sampleSize 144 + 4 =148 (37*4)
''''''PASENSOR_GPU sampleSize 504 + 14 = 518 (37*14)
''''''
'''''' DSSC cycle: 9, 9, 9, 9, 1, 9, 9, 9, 9, 1, ......
'''''' ~~~~~~~~
'''''' 37 bit
'''''
Public Function E_02_PASENSOR_Cal( _
       ByVal in_CapturedWave_DSP As DSPWave, _
       ByVal in_LoopCount_lng As Long, _
       ByRef out_Data_DSP As DSPWave, _
       ByRef out_valid_DSP As DSPWave) As Long
' get data array and valid bits, called in Function: <VBT_PASensor_ReadCode>

    Dim i As Long
    Dim i_EntireCapture_DSP As New DSPWave
    Dim i_tmpIdx_DSP As New DSPWave

    ' concatenate all the captured wave segment
    i_EntireCapture_DSP.CreateConstant 0, 0, DspLong
    On Error Resume Next    ' DSP does not support <On Error Goto ...>, but you can use <On Error Resume Next>
    For i = 0 To in_LoopCount_lng
        If i > 0 Then in_CapturedWave_DSP.Next
        i_EntireCapture_DSP = i_EntireCapture_DSP.Concatenate(in_CapturedWave_DSP)
    Next

    ' check integrity of captured wave
    If Not (i_EntireCapture_DSP.SampleSize = (in_LoopCount_lng + 1) * in_CapturedWave_DSP.SampleSize) Then
        E_02_PASENSOR_Cal = 404
        out_Data_DSP.CreateConstant 999, 999, DspLong
        out_valid_DSP.CreateConstant 999, 999, DspLong
        Exit Function
    End If

    ' calculate them
    '''    Call dsp_PASENSOR_wave_divide(i_EntireCapture_DSP, out_Data_DSP, out_Valid_DSP)
    '''    Exit Function

    out_valid_DSP = i_EntireCapture_DSP.Select(36, 37).Copy
    i_tmpIdx_DSP.CreateConstant 1, out_valid_DSP.SampleSize * 36, DspLong
    i_tmpIdx_DSP.Select(36, 36).Replace (2)
    i_tmpIdx_DSP.Element(0) = 0
    i_tmpIdx_DSP = i_tmpIdx_DSP.IntegrateElements
    out_Data_DSP = i_EntireCapture_DSP.Lookup(i_tmpIdx_DSP)
    out_Data_DSP = out_Data_DSP.ConvertStreamTo(tldspParallel, 9, 0, Bit0IsMsb)
    ' THERE IS A BUG IN IGXL DSP, THE Bit0IsMSB is actually LSB first.


End Function

''''
'''' PM_H240_0/1, PM_H360_0/1/2/3 sampleSize 1312 = 41 * 32
''''
'''' DSSC cycle: 1, 1, 30, 1, 1, 30, 1, 1, 30, ......
''''
'''' 32 bit
''''
Public Function E_03_PM_Cal( _
       ByVal in_CapturedWave_DSP As DSPWave, _
       ByVal in_LoopCount_lng As Long, _
       ByRef out_Data_DSP As DSPWave, _
       ByRef out_valid_DSP As DSPWave) As Long
' get data array and valid bits, called in Function: <PM_ReadCode_vbt>

    Dim i As Long
    Dim i_EntireCapture_DSP As New DSPWave
    Dim i_OneCapture_DSP As New DSPWave

    ' concatenate all the captured wave segment
    i_EntireCapture_DSP.CreateConstant 0, 0, DspLong
    On Error Resume Next    ' DSP does not support <On Error Goto ...>, but you can use <On Error Resume Next>

    For i = 0 To in_LoopCount_lng
        If i > 0 Then in_CapturedWave_DSP.Next
        'i_OneCapture_DSP = in_CapturedWave_DSP.ConvertStreamTo(tldspParallel, 32, 0, Bit0IsLsb).ConvertDataTypeTo(DspLong)
        i_OneCapture_DSP = in_CapturedWave_DSP
        i_EntireCapture_DSP = i_EntireCapture_DSP.Concatenate(i_OneCapture_DSP)
    Next

    ' check integrity of captured wave
    If Not (i_EntireCapture_DSP.SampleSize = (in_LoopCount_lng + 1) * in_CapturedWave_DSP.SampleSize) Then
        E_03_PM_Cal = 404
        out_Data_DSP.CreateConstant 999, 999, DspLong
        out_valid_DSP.CreateConstant 999, 999, DspLong
        Exit Function
    End If

    '     Calculate them
    out_Data_DSP = i_EntireCapture_DSP.BitwiseShiftRight(2).BitwiseAnd(&H3FFFFFFF)
    out_valid_DSP = i_EntireCapture_DSP.BitwiseAnd(&H3).ConvertStreamTo(tldspSerial, 2, 0, Bit0IsMsb)

End Function



'
Public Function E_06_TSensor_SoC_Cal( _
       ByVal in_CapturedWave_DSP As DSPWave, _
       ByVal in_Chuck_Temp_dbl As Double, _
       ByVal in_MoudleCount_lng As Long, _
       ByRef out_Temp_Ready_1_lng As Long, _
       ByRef out_Temp_Mean_1_dbl As Double, _
       ByRef out_Temp_Delta_Mean_1_dbl As Double, _
       ByRef out_Temp_Delta_Min_1_dbl As Double, _
       ByRef out_Temp_Delta_Max_1_dbl As Double _
     ) As Long
' read Tsensor in SoC through PinGroups, called in Function: <TSensor_SOC_DSSC_vbt>

    Dim i                                                       As Long
    Dim i_Ready_1_DSP                                           As New DSPWave
    Dim i_Tempwave_1_DSP                                        As New DSPWave
    Dim i_CapReady_DSP                                          As New DSPWave
    Dim i_CapTemp_DSP                                           As New DSPWave
  

    On Error Resume Next    ' DSP does not support <On Error Goto ...>, but you can use <On Error Resume Next>

    i_CapReady_DSP = in_CapturedWave_DSP.Select(0, 2).Copy
    i_CapTemp_DSP = in_CapturedWave_DSP.Select(1, 2).Copy
    
    
    i_Ready_1_DSP = i_CapReady_DSP.BitwiseExtract(1, 0)
    '(TempValue - 414#) / (700# - 414#) * (125# - (-40#)) + (-40#)
    i_Tempwave_1_DSP = i_CapTemp_DSP.BitwiseExtract(16, 1).Subtract(414).ConvertDataTypeTo(DspDouble)
    i_Tempwave_1_DSP = i_Tempwave_1_DSP.Multiply((125 + 40) / (700 - 414)).Subtract(40#)

    out_Temp_Ready_1_lng = i_Ready_1_DSP.CalcMinimumValue
    out_Temp_Mean_1_dbl = i_Tempwave_1_DSP.CalcMean
    out_Temp_Delta_Mean_1_dbl = out_Temp_Mean_1_dbl - in_Chuck_Temp_dbl
    out_Temp_Delta_Min_1_dbl = i_Tempwave_1_DSP.CalcMinimumValue - in_Chuck_Temp_dbl
    out_Temp_Delta_Max_1_dbl = i_Tempwave_1_DSP.CalcMaximumValue - in_Chuck_Temp_dbl

End Function



'
Public Function E_06_TSensor_DJTAG_Cal( _
       ByVal in_CapturedWave_DSP As DSPWave, _
       ByVal in_Chuck_Temp_dbl As Double, _
       ByRef out_Temp_Mean_DSP As DSPWave, _
       ByRef out_Temp_Delta_Mean_DSP As DSPWave) As Long
' read Tsensor in SoC through JTAG_TDO, called in Function: <TSensor_DJTAG_DSSC_vbt>

    On Error Resume Next    ' DSP does not support <On Error Goto ...>, but you can use <On Error Resume Next>

'(TempValue - 109#) / (907# - 109#) * (125# - (-40#)) + (-40#)
'(TempValue - 414#) / (700# - 414#) * (125# - (-40#)) + (-40#)
    out_Temp_Mean_DSP = in_CapturedWave_DSP.BitwiseExtract(16, 0).Subtract(414).ConvertDataTypeTo(DspDouble)
    out_Temp_Mean_DSP = out_Temp_Mean_DSP.Multiply((125 + 40) / (700 - 414)).Subtract(40#)

    out_Temp_Delta_Mean_DSP = out_Temp_Mean_DSP.Subtract(in_Chuck_Temp_dbl)

End Function


Public Function D_02_MTCMOS_Delay( _
       ByVal in_capWave_DSP As DSPWave, _
       ByRef out_MTCMOS_Delay_L2H_DSP As DSPWave, _
       ByRef out_MTCMOS_Delay_H2L_DSP As DSPWave) As Long
'

    On Error Resume Next    ' DSP does not support <On Error Goto ...>, but you can use <On Error Resume Next>

    Dim i As Long

    Dim i_bitWave_DSP As New DSPWave

    Dim i_delay_L2H_dbl(24) As Double
    Dim i_delay_H2L_dbl(24) As Double

    '    out_MTCMOS_Delay_L2H_DSP.CreateConstant 0, 24, DspDouble
    '    out_MTCMOS_Delay_H2L_DSP.CreateConstant 0, 24, DspDouble

    For i = 0 To 24
        i_bitWave_DSP = in_capWave_DSP.BitwiseExtract(1, i)
        'dsp_1pin_MTCMOS_Delay i_bitWave_DSP, i_delay_L2H_dbl, i_delay_H2L_dbl
    Next i

    out_MTCMOS_Delay_L2H_DSP.Data = i_delay_L2H_dbl
    out_MTCMOS_Delay_H2L_DSP.Data = i_delay_H2L_dbl

End Function

Public Function E_06_Tsensor_Glitch(ByVal in_capWave_DSP As DSPWave, _
                                    ByRef out_TempMean_db As Double, _
                                    ByRef out_TempMin_db As Double, _
                                    ByRef out_TempMax_db As Double, _
                                    ByRef out_TempGap_db As Double) As Long
    
    out_TempMean_db = in_capWave_DSP.CalcMean
    out_TempMin_db = in_capWave_DSP.CalcMinimumValue
    out_TempMax_db = in_capWave_DSP.CalcMaximumValue
    out_TempGap_db = out_TempMax_db - out_TempMean_db
    
End Function


