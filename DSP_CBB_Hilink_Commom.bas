Attribute VB_Name = "DSP_CBB_Hilink_Commom"
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

Public Function Calc_Res_MinMax(ByVal ResArray As DSPWave, ByRef Res_Min As Double, ByRef Res_Max As Double, ByVal LBRes As Double) As Long

'    Dim VoltageArray As New DSPWave
'    VoltageArray.CreateRamp 0, 0.05, ResArray.SampleSize
'
'    ResArray = VoltageArray.Divide(ResArray).Select(1)
'    Res_Min = ResArray.CalcMinimumValue
'    Res_Max = ResArray.CalcMaximumValue

    Dim VoltageArray As New DSPWave
    Dim ResArraytemp As New DSPWave
    VoltageArray.CreateRamp 0.1, 0.1, ResArray.SampleSize
'    VoltageArray.CreateConstant 0.1, ResArray.SampleSize, DspDouble
'    ResArray = ResArray.Differentiate
    ResArray = VoltageArray.Divide(ResArray).Select(0).Copy
    ResArray = ResArray.Subtract(LBRes)
    Res_Min = ResArray.CalcMinimumValue
    Res_Max = ResArray.CalcMaximumValue

End Function

Public Function Calc_DeltaV_MinMax(ResArray As DSPWave, Res_Min As Double, Res_Max As Double) As Long

    Res_Min = ResArray.CalcMinimumValue
    Res_Max = ResArray.CalcMaximumValue

End Function


Public Function INL_DNL_GAIN_OFFSET(P_wave As DSPWave, N_wave As DSPWave, INL As Double, DNL As Double, gain As Double, Offset As Double) As Long

    Dim Coeff As New DSPWave
    Dim DIFF_wave As New DSPWave
    Dim DIFF_wave2 As New DSPWave
    Dim DIFF_wave3 As New DSPWave
    Dim DNL_Wave As New DSPWave
    Dim INL_Wave As New DSPWave
    Dim Ideal_Wave As New DSPWave

    Dim n As Long
    Dim MeanVal As Double

    Dim InitalVal As Double

    DIFF_wave = P_wave.Subtract(N_wave)
    'DIFF_wave.Plot "TxFIR_DC_MAIN_" + "All1"
    n = DIFF_wave.SampleSize

    Offset = DIFF_wave.Element(0)
    gain = DIFF_wave.Element(n - 1)

    Coeff = DIFF_wave.FitPolynomial(1)
    MeanVal = Coeff.Element(1)

    Call DIFF_wave2.CreatePolynomial(Coeff, n)
    DIFF_wave2 = DIFF_wave.Subtract(DIFF_wave2)
    'DIFF_wave2.Plot "+INL"

    INL = DIFF_wave2.Divide(MeanVal).CalcMaximumMagnitude

    DIFF_wave3 = DIFF_wave2.Differentiate
    'DIFF_wave3.Plot "+DNL"

    DNL = DIFF_wave3.Divide(MeanVal).CalcMaximumMagnitude


    'Dim MaxVal As Double
    'Dim MinVal As Double
    'MaxVal = DIFF_wave.CalcMaximumValue
    'MinVal = DIFF_wave.CalcMinimumValue
    'IDEAL_wave.CreateRamp MinVal,


End Function

Public Function INL_DNL_GAIN_OFFSET_PREPOINT(P_wave As DSPWave, N_wave As DSPWave, INL As Double, DNL As Double, gain As Double, Offset As Double, DIFF_wave As DSPWave) As Long

    'Dim DIFF_wave As New DSPWave
    Dim DIFF_wave2 As New DSPWave
    Dim DNL_Wave As New DSPWave
    Dim INL_Wave As New DSPWave
    Dim Ideal_Wave As New DSPWave

    Dim n As Long
    Dim MeanVal As Double

    Dim InitalVal As Double

    DIFF_wave = P_wave.Subtract(N_wave)
    n = DIFF_wave.SampleSize

    DIFF_wave2 = DIFF_wave.Differentiate
    MeanVal = DIFF_wave2.CalcMean
    DIFF_wave2 = DIFF_wave2.Subtract(MeanVal)
    DNL = DIFF_wave2.Divide(MeanVal).CalcMaximumMagnitude
    Offset = DIFF_wave.Element(0)
    gain = DIFF_wave.Element(n - 1)

    InitalVal = DIFF_wave2.Divide(MeanVal).Element(0)
    DIFF_wave2 = DIFF_wave2.Divide(MeanVal).Integrate(InitalVal)
    INL = DIFF_wave2.CalcMaximumValue
End Function

Public Function MeasureJitter_DSP_UP1600(ByVal DSPWavesAllPinsAllSites As DSPWave, ByRef DDj As Double, ByRef Rj As Double, ByRef Tj As Double, ByRef Meas_UI As Double, ByRef dspStatus As Long) As Long

    Call DSPWavesAllPinsAllSites.MeasureJitter(Rj, DDj, Meas_UI, dspStatus)
    
    Tj = DDj + Rj * 14
    
   ' Call TheHdw.Digital.Jitter.FileExport(DSPWavesAllPinsAllSites, ".\ExportFile\FileExport_Jitter.txt")
    
End Function

Public Function MeasureEye_DSP_UP1600(ByVal DSPWavesAllPinsAllSites As DSPWave, ByRef DDj As Double, ByRef Rj As Double, ByRef Tj As Double, ByRef Meas_UI As Double, _
                                      ByRef JitterPKPK As Double, ByRef P2Pj_WoDDj As Double, ByRef dspStatus As Long) As Long

    Dim RiseTime As Double
    Dim FallTime As Double
    Dim EarlyLow As Double
    Dim LateLow As Double
    Dim EarlyMid As Double
    Dim LatMid As Double
    Dim EarlyHigh As Double
    Dim LateHigh As Double
    
    DSPWavesAllPinsAllSites.MeasureEye Rj, DDj, Meas_UI, RiseTime, FallTime, EarlyLow, LateLow, EarlyMid, LatMid, EarlyHigh, LateHigh, dspStatus

    If dspStatus = 0 Then
        Tj = DDj + Rj * 14
        JitterPKPK = LatMid - EarlyMid
        P2Pj_WoDDj = LatMid - EarlyMid - DDj
    Else
        Tj = dspStatus
        JitterPKPK = dspStatus
    End If
    
    'Call TheHdw.Digital.Jitter.FileExport(Wave, ".\Datalog_DumpStatus\20180418\FDCLK_Jitter\FileExport_FDCLKJitter_LB.txt")
End Function














