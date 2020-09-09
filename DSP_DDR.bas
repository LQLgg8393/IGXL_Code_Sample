Attribute VB_Name = "DSP_DDR"
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


Public Function D_03_CalculateSPO( _
       ByVal in_CaptureWave_DSP As DSPWave, _
       ByVal in_DataRate_dbl As Double, _
       ByVal in_DelayMeas_dbl As Double, _
       ByRef out_RdqsCyc_dbl As Double, _
       ByRef out_SPO_dbl As Double, _
       ByRef out_Jitter_dbl As Double, _
       ByRef out_Jitter_ref_dbl As Double, _
       ByRef out_Jitter_fb_dbl As Double) As Long

    Dim i As Long
    Dim i_Delay_dbl As Double

    Dim i_RdqsCyc_DSP As New DSPWave

    Dim i_RefClk_bdl_E_DSP As New DSPWave
    Dim i_FbClk_bdl_E_DSP As New DSPWave
    Dim i_All_bdl_E_DSP As New DSPWave
    Dim i_All_bdl_Minus50_DSP As New DSPWave

    Dim I_SPO_lng As Long
    Dim i_BeginNot0_lng As Long
    Dim i_EndNot100_lng As Long
    Dim i_Jitter_lng As Long
    Dim i_Jitter_ref_lng As Long
    Dim i_Jitter_fb_lng As Long
    
    Dim i_BeginJitterIndex_lng As Long
    Dim i_EndJitterIndex_lng As Long

    On Error Resume Next    ' DSP does not support <On Error Goto ...>, but you can use <On Error Resume Next>
    
    ' I_SPO_lng init value = 15
    I_SPO_lng = 15
    

    ' check out
    If in_CaptureWave_DSP.SampleSize <> 500 Then
        out_RdqsCyc_dbl = 9999
        out_SPO_dbl = 9999
        out_Jitter_dbl = 9999
        Exit Function
    End If


    '''-----Calculate rdqscyc value-----
    '''RdqsCyc(0)    0-9
    '''RdqsCyc(1)    10-19

    i_RdqsCyc_DSP = in_CaptureWave_DSP.Select(0, , 20).ConvertStreamTo(tldspParallel, 10, 0, Bit0IsMsb)

    out_RdqsCyc_dbl = i_RdqsCyc_DSP.Element(0)

    If out_RdqsCyc_dbl = 0 Then
        out_RdqsCyc_dbl = 8888
        out_SPO_dbl = 8888
        out_Jitter_dbl = 8888
        Exit Function
    End If

    i_Delay_dbl = in_DelayMeas_dbl * 4 * (1000000 / in_DataRate_dbl) / out_RdqsCyc_dbl


    '''-----Calculate RefClk_bdl and FbClk_bdl-----
    '''RefClk_bdl   (0-14, 15-29,..., 225-239)+20
    '''FbClk_bdl    (0-14, 16-29,..., 225-239)+240+20

    i_All_bdl_E_DSP = in_CaptureWave_DSP.Select(20, , 15 * 16 * 2).ConvertStreamTo(tldspParallel, 15, 0, Bit0IsMsb)
    i_All_bdl_E_DSP = i_All_bdl_E_DSP.Divide(10)    '.Divide(iSpoSample).Multiply(100), iSpoSample=1000
    
''''    'for SPO debug
''''    i_All_bdl_E_DSP.CreateConstant 100, 32, DspDouble
''''    For i = 0 To 17
''''        i_All_bdl_E_DSP.Element(i) = 0
''''    Next i

    i_All_bdl_Minus50_DSP = i_All_bdl_E_DSP.Subtract(50)

    i_RefClk_bdl_E_DSP = i_All_bdl_E_DSP.Select(15, -1, 16).Copy    ' reordered the first 16 samples
    i_FbClk_bdl_E_DSP = i_All_bdl_E_DSP.Select(16, , 16).Copy
    ' NEW ALGO!!!  do not reorder
    'i_All_bdl_E_DSP = i_RefClk_bdl_E_DSP.Concatenate(i_FbClk_bdl_E_DSP)


    '''Cal the nearest to 50 data's index
    '''Cal SPO_n
    Call i_All_bdl_Minus50_DSP.Abs.CalcMinimumValue(I_SPO_lng)
    out_SPO_dbl = I_SPO_lng * i_Delay_dbl
    
    If i_All_bdl_Minus50_DSP.CalcMinimumValue(I_SPO_lng) = 50 Then
        I_SPO_lng = 15
        out_SPO_dbl = I_SPO_lng * i_Delay_dbl  'out_SPO_dbl = 15 * i_Delay_dbl
         
    ElseIf i_All_bdl_Minus50_DSP.CalcMinimumValue(I_SPO_lng) = -50 And _
           i_All_bdl_Minus50_DSP.Abs.CalcMinimumValue(I_SPO_lng) = 50 And _
           i_FbClk_bdl_E_DSP.CalcMinimumValue(I_SPO_lng) = 100 And _
           i_RefClk_bdl_E_DSP.CalcMaximumValue(I_SPO_lng) = 100 Then
        
       I_SPO_lng = 15 - i_RefClk_bdl_E_DSP.FindIndex(OfFirstElement, EqualTo, 100)
       out_SPO_dbl = I_SPO_lng * i_Delay_dbl
       
    ElseIf i_All_bdl_Minus50_DSP.CalcMinimumValue(I_SPO_lng) = -50 And _
           i_All_bdl_Minus50_DSP.Abs.CalcMinimumValue(I_SPO_lng) = 50 And _
           i_RefClk_bdl_E_DSP.CalcMaximumValue(I_SPO_lng) = 0 And _
           i_FbClk_bdl_E_DSP.CalcMinimumValue(I_SPO_lng) = 0 And _
           i_FbClk_bdl_E_DSP.CalcMaximumValue(I_SPO_lng) = 100 Then
           
       I_SPO_lng = i_FbClk_bdl_E_DSP.FindIndex(OfFirstElement, EqualTo, 100) - 1
       out_SPO_dbl = I_SPO_lng * i_Delay_dbl
       
    ElseIf i_All_bdl_Minus50_DSP.Abs.CalcMinimumValue(I_SPO_lng) < 50 And I_SPO_lng > 15 Then
    
          I_SPO_lng = I_SPO_lng - 16
          out_SPO_dbl = I_SPO_lng * i_Delay_dbl
    ElseIf i_All_bdl_E_DSP.CalcMaximumValue(I_SPO_lng) = 0 Then
          I_SPO_lng = 15
          out_SPO_dbl = I_SPO_lng * i_Delay_dbl
          
    End If
    
    '''Cal Jitter_ref
    '''Cal the first and last non 0 non 100 index
    i_BeginJitterIndex_lng = i_RefClk_bdl_E_DSP.Subtract(50).Abs.FindIndex(OfFirstElement, NotEqualTo, 50)
    i_EndJitterIndex_lng = i_RefClk_bdl_E_DSP.Subtract(50).Abs.FindIndex(OfLastElement, NotEqualTo, 50)
    If i_BeginJitterIndex_lng < 0 Then
        i_Jitter_ref_lng = 0
    Else
        i_Jitter_ref_lng = i_EndJitterIndex_lng - i_BeginJitterIndex_lng + 1
    End If
    
    out_Jitter_ref_dbl = i_Jitter_ref_lng * i_Delay_dbl

    '''Cal Jitter_fb
    '''Cal the first and last non 0 non 100 index
    i_BeginJitterIndex_lng = i_FbClk_bdl_E_DSP.Subtract(50).Abs.FindIndex(OfFirstElement, NotEqualTo, 50)
    i_EndJitterIndex_lng = i_FbClk_bdl_E_DSP.Subtract(50).Abs.FindIndex(OfLastElement, NotEqualTo, 50)
    If i_BeginJitterIndex_lng < 0 Then
        i_Jitter_fb_lng = 0
    Else
        i_Jitter_fb_lng = i_EndJitterIndex_lng - i_BeginJitterIndex_lng + 1
    End If
    
    out_Jitter_fb_dbl = i_Jitter_fb_lng * i_Delay_dbl
    


    '''Cal the first non 0 index and last non 100 index
    '''Cal Jitter_n
    i_BeginNot0_lng = i_RefClk_bdl_E_DSP.Subtract(50).Abs.FindIndex(OfFirstElement, NotEqualTo, 50)
    i_EndNot100_lng = i_FbClk_bdl_E_DSP.Subtract(50).Abs.FindIndex(OfLastElement, NotEqualTo, 50)
    
'    If i_BeginNot0_lng < 0 Then
'        i_BeginNot0_lng = 0
'    End If
'    If i_EndNot100_lng < 0 Then
'        i_EndNot100_lng = 0
'    End If
    
    ''' Judge if index 15 and index 16 are in the section between first non 0 index and last non 100 index, if so, i_Jitter_lng=i_EndNot100_lng - i_BeginNot0_lng + 16
    If i_BeginNot0_lng > -1 And i_EndNot100_lng > -1 Then
       i_Jitter_lng = i_EndNot100_lng - i_BeginNot0_lng + 16
          
    ElseIf i_EndNot100_lng = -1 And i_BeginNot0_lng > -1 Then
            i_Jitter_lng = 16 - i_BeginNot0_lng             ''' Lower Section (Index 16 to Index 31) doesn't contain non 100 data
   
    ElseIf i_EndNot100_lng > -1 And i_BeginNot0_lng = -1 Then
            i_Jitter_lng = i_EndNot100_lng + 1              ''' Upper Section (Index 0 to Index 15) doesn't contain non 0 data
  
    ElseIf i_BeginNot0_lng = -1 And i_EndNot100_lng = -1 Then
        i_Jitter_lng = 0                                    ''' Upper Section (Index 0 to Index 15) doesn't contain non 0 data and Lower Section (Index 16 to Index 31) doesn't contain non 100 data
    End If
    
    out_Jitter_dbl = i_Jitter_lng * i_Delay_dbl

End Function

Public Function D_07_CalDDRLPBK( _
       ByVal CaptureWave As DSPWave, _
       ByRef zcal_result As Long) As Long

    On Error Resume Next    ' DSP does not support <On Error Goto ...>, but you can use <On Error Resume Next>

    Dim zcal008 As Long
    Dim zcal028 As Long

    zcal008 = CaptureWave.Element(0)
    zcal028 = CaptureWave.Element(1)

    If zcal008 = &H10000000 Then
        zcal_result = 1
    ElseIf zcal008 = &H10000008 Then
        If (zcal028 And &H80) > 0 Then
            zcal_result = 0
        ElseIf (zcal028 And &H7F0000) Then
            zcal_result = 1
        Else
            zcal_result = 0
        End If
    Else
        zcal_result = 0
    End If

End Function




Public Function D_06_Cal_PLL_Unlock( _
       ByVal in_CaptureWave_DSP As DSPWave, _
       ByRef out_RegData_DSP As DSPWave) As Long
' should use DSSC serial mode!!!
    On Error Resume Next    ' DSP does not support <On Error Goto ...>, but you can use <On Error Resume Next>

    If in_CaptureWave_DSP.SampleSize <> 192 Then
        out_RegData_DSP.CreateConstant 999, 8
        Exit Function
    End If

    out_RegData_DSP = in_CaptureWave_DSP.ConvertStreamTo(tldspParallel, 32, 0, Bit0IsMsb)

End Function

