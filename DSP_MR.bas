Attribute VB_Name = "DSP_MR"

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

Public Function dspwaveSum(inDspwave As DSPWave, sumResult As Long, sumTop As Long, sumBot As Long) As Long
    
    sumResult = inDspwave.CalcSum / inDspwave.SampleSize
    sumTop = inDspwave.Select(0, 1, inDspwave.SampleSize / 2).CalcSum / inDspwave.SampleSize
    sumBot = inDspwave.Select(inDspwave.SampleSize / 2, 1, inDspwave.SampleSize / 2).CalcSum / inDspwave.SampleSize
    inDspwave = inDspwave.ConvertStreamTo(tldspSerial, 32, 0, Bit0IsLsb)  ' reverse for MRBsDatarocess--> decompress()
End Function




Public Function dsp_bisr_decompress(dw_data As DSPWave, dw_result As DSPWave, dw_chainlength As DSPWave, i_BotAreaFresh As Long, i_decompressFunctionerror As Long, _
MRB_INFO_BASE As Long, MRB_INFO_LEN As Long, MAX_MRB_NUM As Long) As Long


' for dsp function
    Dim i_MRBNum As Long
    i_decompressFunctionerror = 0
    If i_BotAreaFresh <> 0 Then Exit Function
    i_MRBNum = MAX_MRB_NUM
    Dim chainLength() As Long
    ReDim chainLength(i_MRBNum - 1)
    chainLength = dw_chainlength.Data

'On Error GoTo errhandler
On Error Resume Next
    Dim dw_result_even As New DSPWave: dw_result_even.CreateConstant -99, dw_result.SampleSize / 2, DspLong
    Dim dw_result_odd As New DSPWave: dw_result_odd.CreateConstant -99, dw_result.SampleSize / 2, DspLong
    Dim dw_result_xor As New DSPWave: dw_result_xor.CreateConstant -99, dw_result.SampleSize / 2, DspLong
    
    Dim res(127) As Long
    Dim tmp_dw_res_16bit As New DSPWave: tmp_dw_res_16bit.CreateConstant -99, dw_result.SampleSize / 2 / 16, DspLong
    Dim tmp_dw_res_8bit As New DSPWave: tmp_dw_res_8bit.CreateConstant -99, dw_result.SampleSize / 2 / 8, DspLong
    Dim i_length As Long
    
    Dim dw_res As New DSPWave
    Dim tmp1 As Long
    Dim i As Long, j As Long, k As Long, ptr_addr As Long, zero_cnt As Long
    
    
    ' check
    If dw_data.SampleSize < 0 Or dw_result.SampleSize < 0 Then
        i_decompressFunctionerror = -1
        Exit Function
    End If
    
    'check
    If dw_result.SampleSize Mod 32 <> 0 Then
        i_decompressFunctionerror = 3
        Exit Function
    End If
    'check
    If MAX_MRB_NUM <= 0 Then
        i_decompressFunctionerror = 3
        Exit Function
    End If
    
    'check
    If dw_chainlength.SampleSize <> MAX_MRB_NUM Then
        i_decompressFunctionerror = 3
        Exit Function
    End If
    
    'check
    If dw_result.SampleSize <> 4096 Then
        i_decompressFunctionerror = 3
        Exit Function
    End If
    For i = 0 To i_MRBNum - 1
        dw_data.CreateConstant -99, chainLength(i), DspLong
    Next i
    
    ' double bit check
    dw_result_even = dw_result.Select(1, 2, dw_result.SampleSize / 2).ConvertDataTypeTo(DspLong) ' exchange even to odd for match 93k
    dw_result_odd = dw_result.Select(0, 2, dw_result.SampleSize / 2).ConvertDataTypeTo(DspLong)
    
    dw_result_xor = dw_result_odd.BitwiseXor(dw_result_even)
    'dw_result_xor = dw_result_odd.Subtract(dw_result_even)
    
    If dw_result_xor.CalcSum <> 0 Then
        'theexec.Datalog.WriteComment " Error: Double Bit Check Failed "
        i_decompressFunctionerror = 5
        Exit Function
    End If

    
    ' to align 93k code line number 290
'    dw_res = dw_result_even.Copy

    ' to get every second bit, then store to res()
    tmp_dw_res_16bit = dw_result_even.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspParallel, 16, 0, Bit0IsLsb) ' 128 samples
    
    For i = 0 To 127
        res(i) = 0
        res(i) = tmp_dw_res_16bit.Element(i)        ' i_dw_result1_even_16bit.Element(i)
        'TheExec.Datalog.WriteComment " res" + CStr(i) + "  : " + CStr(res(i))
    Next i
    
    ' ############## (START) code check-03 ######################################################
'            For i = 0 To dw_res.SampleSize - 1
'                TheExec.Datalog.WriteComment "dw_res.Element" + CStr(i) + ": " + CStr(dw_res.Element(i))
'            Next i
    ' ################## (END) code check-03 ####################################################


    Dim tmp_dw_CurrentChain As New DSPWave
    Dim arr_CurrentChain() As Long

    Dim i_validBitsCnt As Long
    Dim tmp_dw_data As New DSPWave
    tmp_dw_data.CreateConstant 0, 0, DspLong

    Dim first_address As Long
    first_address = res(0)
    first_address = first_address And &HFF
    If first_address <> MRB_INFO_BASE Then
        i_decompressFunctionerror = 7
        Exit Function           ' break
    End If
    

    ' start to decompress
    For i = 0 To i_MRBNum - 1
        tmp_dw_CurrentChain.CreateConstant -99, chainLength(i), DspLong
        ReDim arr_CurrentChain(chainLength(i) - 1)
        'read address
        ptr_addr = res(i \ 2) '/
        If i Mod 2 = 1 Then
            ptr_addr = ptr_addr / 2 ^ 8  ' right shift 8 bit
        Else
            ptr_addr = ptr_addr And 255 '&HFF
            If ptr_addr = &HFF Then
                i_decompressFunctionerror = 6
                Exit Function     'break
            End If
        End If
        
        If ptr_addr > 127 Then Exit For

        For j = 0 To chainLength(i) - 1
            If ptr_addr > 127 Then
                
                i_decompressFunctionerror = 4 ' defined by licha
                i_validBitsCnt = j
                'theexec.Datalog.WriteComment " too many 1s in the block(chain)"
                Exit For    'too many 1s in chain will caused ptr_addr exceeds 128
            End If
            zero_cnt = res(ptr_addr) And 8191 ' &H1FFF&
            'TheExec.Datalog.WriteComment " zero cnt: " + CStr(zero_cnt) + " " + CStr(j)
            For k = 0 To zero_cnt - 1
                'dw_data.Element(i_length + j) = 0
                If Not (j < chainLength(i)) Then Exit For
                arr_CurrentChain(j) = 0 ' >>> instead of dw_data.Element(), for ttr
                j = j + 1
            Next k
           
            If Not (j < chainLength(i)) Then Exit For

            i_length = i_length + chainLength(i)
            
            'If (res(ptr_addr) And &HE000&) = &H2000& Then
            If (res(ptr_addr) And 57344) = 8192 Then
                ptr_addr = ptr_addr + 1
                
                Dim i_ValidBit As Long: i_ValidBit = 0
                For k = 0 To 16 - 1   'copy the data from LSB to MSB
'                For k = 15 To 0 Step -1 ' 20190418
'                    Dim i_ValidBit As Long: i_ValidBit = 0
                    If j < chainLength(i) Then
                        If (res(ptr_addr) And (2 ^ k)) <> 0 Then ' to judge if the bit is 1 or 0
                            'dw_data.Element(i_length + j) = 1
                            arr_CurrentChain(j) = 1   ' >>> instead of dw_data.Element(), for ttr
                            i_ValidBit = 1
                            j = j + 1
                        Else
                            'dw_data.Element(i_length + j) = 0
                            If i_ValidBit = 1 Then
                                arr_CurrentChain(j) = 0   ' >>> instead of dw_data.Element(), for ttr
                                j = j + 1
                            End If
                            'j = j + 1
                        End If
                        'j = j + 1
                    End If
                Next k
                If k = 16 Then j = j - 1
                ptr_addr = ptr_addr + 1
            Else
                ptr_addr = ptr_addr + 1
            End If
        Next j
        tmp_dw_CurrentChain.Data = arr_CurrentChain
        tmp_dw_data = tmp_dw_CurrentChain.Concatenate(tmp_dw_data)
'        tmp_dw_data = tmp_dw_data.Concatenate(tmp_dw_CurrentChain)
    Next i

    
    If 0 Then
        dw_data = tmp_dw_data.Copy  ' error happening
   
    Else  ' workaround with below method
        Dim tmpArr() As Long: ReDim tmpArr(tmp_dw_data.SampleSize - 1)
        tmpArr = tmp_dw_data.Data
        dw_data.Data = tmpArr
    End If
    
End Function


Public Function dsp_bisr_compress(dw_data As DSPWave, dw_result As DSPWave, dw_chainlength As DSPWave, i_MRBNum As Long, i_MRB_INFO_BASE As Long, i_TotalChainLength As Long, i_compressFunctionerror As Long, i_MRBInfoZeroCount As Long, i_AreaFresh As Long, i_runCompressFlag As Long, i_dw_data As DSPWave) As Long

    i_compressFunctionerror = 0
    
    ' when no 1s in MRB info or the area is unfresh, won't execute compress action
    If i_MRBInfoZeroCount <> 1 Or i_AreaFresh <> 1 Then i_runCompressFlag = 0: Exit Function
    If dw_chainlength.SampleSize <> i_MRBNum Then i_compressFunctionerror = -1: Exit Function
    If i_TotalChainLength <= 0 Then i_compressFunctionerror = -1: Exit Function
    
    Dim i As Long, j As Long, k As Long
    Dim ptr_res_info As Long
    Dim cnt_zero As Long
    Dim tmp_dw_res As New DSPWave
    Dim tmp_res(31) As Long

    Dim dw_dataPerMRB As New DSPWave

    Dim tmp_dw_result As New DSPWave
    Dim tmp_Length As Long
    Dim res(127) As Long
    Dim ii As Long, jj As Long
    Dim i_HAVE_RAW_DATA As Long

On Error Resume Next
    Dim chainLength() As Long
    ReDim chainLength(i_MRBNum - 1)
    chainLength = dw_chainlength.Data
  
    i_HAVE_RAW_DATA = 8192 ' &H2000
    
    tmp_dw_result.CreateConstant 0, 0, DspLong
    tmp_dw_res.CreateConstant 0, 32, DspLong
    
    ' check
    If dw_data.SampleSize < 0 Or dw_result.SampleSize < 0 Then
        i_compressFunctionerror = -1
        Exit Function
    End If
    
    For i = 0 To 127
        res(i) = 0
    Next i
    
    'init all the pointers
    ptr_res_info = i_MRB_INFO_BASE
    
    For i = 0 To i_MRBNum - 1
        If dw_data.SampleSize <> i_TotalChainLength Then
            i_compressFunctionerror = 3
            Exit Function
        End If
                
        If i Mod 2 = 1 Then
            res(i \ 2) = ptr_res_info * 2 ^ 8 Or res(i \ 2)
        Else
            res(i \ 2) = ptr_res_info
        End If
        
        ' compress
        cnt_zero = 0
        
        dw_dataPerMRB = dw_data.Select(i_dw_data.Element(i), 1, chainLength(i)).ConvertDataTypeTo(DspLong)
        tmp_Length = tmp_Length + chainLength(i)

        Dim tmp_dw_16bit As New DSPWave
        Dim tmp_dw_lessthan15bit As New DSPWave
        If ptr_res_info >= 128 Then i_compressFunctionerror = 4: Exit For  ' Check if Bisr Efuse All 128 lines have all data 2019/10/23
        
        If chainLength(i) < 1 Then i_compressFunctionerror = -1: Exit Function
        
        For ii = 0 To chainLength(i) - 1
            jj = dw_dataPerMRB.FindIndex(OfFirstElement, EqualTo, 1)
            'Debug.Print ptr_res_info
            If jj = -1 Then
                cnt_zero = dw_dataPerMRB.SampleSize
                res(ptr_res_info) = cnt_zero
                ptr_res_info = ptr_res_info + 1
                'If ptr_res_info >= MAX_BISR_LEN Then bisr_compress0 = 4: Exit For
                'If ptr_res_info >= 128 Then i_compressFunctionerror = 4: Exit For  'nop @ 20190605
                Exit For
            Else
                cnt_zero = jj
                res(ptr_res_info) = cnt_zero Or i_HAVE_RAW_DATA
                ptr_res_info = ptr_res_info + 1
'                If ptr_res_info > MRB_INFO_LEN Then bisr_compress0 = 4: Exit For
                If ptr_res_info >= 128 Then i_compressFunctionerror = 4: Exit For
                
                If dw_dataPerMRB.SampleSize - jj < 16 Then  ' 15 Then 'UPDATED
                    tmp_dw_16bit = dw_dataPerMRB.Select(jj, 1, dw_dataPerMRB.SampleSize - jj).ConvertDataTypeTo(DspLong)
                    tmp_dw_lessthan15bit.CreateConstant 0, 16 - (dw_dataPerMRB.SampleSize - jj), DspLong
                    tmp_dw_16bit = tmp_dw_16bit.Concatenate(tmp_dw_lessthan15bit)    ' add 0s in front of 1,wron desciption in Hisi document
'                    res(ptr_res_info) = tmp_dw_16bit.ConvertStreamTo(tldspParallel, 16, 0, Bit0IsLsb).Element(0)
                    For k = 0 To 15
                        res(ptr_res_info) = res(ptr_res_info) + tmp_dw_16bit.Element(k) * 2 ^ k
                    Next k
                    ptr_res_info = ptr_res_info + 1
                    If ptr_res_info >= 128 Then i_compressFunctionerror = 4: Exit For
                    'If ptr_res_info > MRB_INFO_LEN Then bisr_compress0 = 4: Exit For
                    Exit For
                    
                ElseIf dw_dataPerMRB.SampleSize - jj >= 16 Then
                    tmp_dw_16bit = dw_dataPerMRB.Select(jj, 1, 16).ConvertDataTypeTo(DspLong)   ' need cover when the count of bits after "1" less than 15, need add mores "0" to ... to be added
                    If tmp_dw_16bit.SampleSize < 16 Then i_compressFunctionerror = 4: Exit For
'                    res(ptr_res_info) = tmp_dw_16bit.ConvertStreamTo(tldspParallel, 16, 0, Bit0IsLsb).Element(0)
                    For k = 0 To 15
                        res(ptr_res_info) = res(ptr_res_info) + tmp_dw_16bit.Element(k) * 2 ^ k
                    Next k
                    ii = jj + 15
                    ptr_res_info = ptr_res_info + 1
                    If ptr_res_info >= 128 Then i_compressFunctionerror = 4: Exit For
                    
                    If ii + 1 < dw_dataPerMRB.SampleSize Then
                        dw_dataPerMRB = dw_dataPerMRB.Select(ii + 1, 1, dw_dataPerMRB.SampleSize - ii - 1).ConvertDataTypeTo(DspLong)
                    Else
                        Exit For
                    End If
                    
                Else
                    i_compressFunctionerror = 4: Exit For
                
                End If
            End If
        Next ii
    Next i

    ' move the format to efuse mode
    For i = 0 To 127
        For j = 15 To 0 Step -1
            If (res(i) And (2 ^ j)) <> 0 Then
                tmp_res(j * 2) = 1
                tmp_res(j * 2 + 1) = 1
            Else
                tmp_res(j * 2) = 0
                tmp_res(j * 2 + 1) = 0
            End If
        Next j
        tmp_dw_res.Data = tmp_res
'        tmp_dw_res = tmp_dw_res.ConvertStreamTo(tldspParallel, 32, 0, Bit0IsMsb).ConvertDataTypeTo(DspLong)
'        tmp_dw_res = tmp_dw_res.ConvertStreamTo(tldspSerial, 32, 0, Bit0IsLsb).ConvertDataTypeTo(DspLong)
        tmp_dw_result = tmp_dw_result.Concatenate(tmp_dw_res)
    Next i
    
    
    If 1 Then
        dw_result = tmp_dw_result.Copy  ' error happening
   
    Else  ' workaround with below method
        Dim tmpArr() As Long: ReDim tmpArr(tmp_dw_result.SampleSize - 1)
        tmpArr = tmp_dw_result.Data
        dw_result.Data = tmpArr
    End If
    
    i_runCompressFlag = 1
                                          
End Function


