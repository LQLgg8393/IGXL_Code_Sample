Attribute VB_Name = "DSP_PLL_FLL"

Option Explicit


Public Function CalcFreq(ByVal in_capWave_DSP As DSPWave, ByVal in_period_dbl As Double, ByRef out_Freq_dbl As Double) As Long

Dim i_cycles_lng As Long
Dim i_FirstEdge_lng As Long
Dim i_LastEdge_lng As Long


in_capWave_DSP = in_capWave_DSP.ConvertDataTypeTo(DspLong)
in_capWave_DSP = in_capWave_DSP.ConvertStreamTo(tldspSerial, 32, 0, Bit0IsMsb)


in_capWave_DSP = in_capWave_DSP.Differentiate
i_FirstEdge_lng = in_capWave_DSP.FindIndex(OfFirstElement, EqualTo, 1)
i_LastEdge_lng = in_capWave_DSP.FindIndex(OfLastElement, EqualTo, 1)
i_cycles_lng = in_capWave_DSP.CountElements(EqualTo, 1) - 1
If i_LastEdge_lng = i_FirstEdge_lng Then
    out_Freq_dbl = 0
Else
    out_Freq_dbl = 1# / in_period_dbl / (i_LastEdge_lng - i_FirstEdge_lng)
    out_Freq_dbl = out_Freq_dbl * CDbl(i_cycles_lng)
End If
End Function
