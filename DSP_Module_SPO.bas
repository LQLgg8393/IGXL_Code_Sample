Attribute VB_Name = "DSP_Module_SPO"
Option Explicit

' ### Teradyne HISI UFLEX Basic Templates V13.00 ###

Public Function UTAH_TMU_T_Delay(ByVal Capwave_PinA As DSPWave, ByVal Capwave_PinB As DSPWave, ByVal WindowWidth As Long, ByVal EFS As Double, T_delay As Double, Duty_DTO0 As Double, Duty_DTO1 As Double, ByVal RisingEdge As Long) As Long

    Dim PinA_wave As New DSPWave
    Dim PinB_wave As New DSPWave

    Dim PinA_Edge As New DSPWave
    Dim PinB_Edge As New DSPWave
    Dim EdgeMean As Double

    Dim tmpMaxA As Double
    Dim tmpMaxB As Double
    Dim tmpMinA As Double
    Dim tmpMinB As Double

    Dim MovingAverage As New DSPWave
    Dim Max_Samples As Long

    PinA_wave = Capwave_PinA.ConvertStreamTo(tldspSerial, 32, 0, Bit0IsMsb).ConvertDataTypeTo(DspLong).ConvertDataTypeTo(DspDouble)
    PinB_wave = Capwave_PinB.ConvertStreamTo(tldspSerial, 32, 0, Bit0IsMsb).ConvertDataTypeTo(DspLong).ConvertDataTypeTo(DspDouble)

    Dim tmpMeanA As Double
    Dim tmpMeanB As Double

    ' Get duty Capwave_PinA and Capwave_PinB
    tmpMeanA = PinA_wave.CalcMean
    tmpMeanB = PinB_wave.CalcMean
    Duty_DTO0 = tmpMeanA
    Duty_DTO1 = tmpMeanB

    ' Duty should be around 0.5, check whether it > 0.8 or < 0.2, if yes, the clock is not reasonable, set T_delay to -901
    If tmpMeanA > 0.8 Or tmpMeanA < 0.2 Or tmpMeanB > 0.8 Or tmpMeanB < 0.2 Then
        T_delay = -901
        Exit Function
    End If

    MovingAverage.CreateConstant 1, WindowWidth, DspDouble

    PinA_wave = PinA_wave.Convolve(MovingAverage).ConvertDataTypeTo(DspLong).Select(WindowWidth, 1, PinA_wave.SampleSize - WindowWidth * 2).Copy  ' bug fixed in 0.2% cases, WindowWidth/2 would be a line
    PinB_wave = PinB_wave.Convolve(MovingAverage).ConvertDataTypeTo(DspLong).Select(WindowWidth, 1, PinB_wave.SampleSize - WindowWidth * 2).Copy  ' bug fixed in 0.2% cases, WindowWidth/2 would be a line

    tmpMaxA = PinA_wave.CalcMaximumValue
    tmpMaxB = PinB_wave.CalcMaximumValue
    tmpMinA = PinA_wave.CalcMinimumValue
    tmpMinB = PinB_wave.CalcMinimumValue

    ' Check whether Window width is reasonable
    If tmpMaxA < WindowWidth * 0.6 Or tmpMaxB < WindowWidth * 0.6 Or tmpMinA > WindowWidth * 0.4 Or tmpMinB > WindowWidth * 0.4 Then
        T_delay = -902
        Exit Function
    End If

    PinA_Edge = PinA_wave.FindIndices(EqualTo, WindowWidth / 2).Select(3).Copy
    PinB_Edge = PinB_wave.FindIndices(EqualTo, WindowWidth / 2).Select(3).Copy

    If RisingEdge = 1 Then

        If PinA_wave.Element(PinA_Edge.Element(0) - 1) < WindowWidth / 2 Then    ' find only rising Edge
            PinA_Edge = PinA_Edge.Select(0, 2).Copy
        Else
            PinA_Edge = PinA_Edge.Select(1, 2).Copy
        End If

        If PinB_wave.Element(PinB_Edge.Element(0) - 1) < WindowWidth / 2 Then    ' find only rising Edge
            PinB_Edge = PinB_Edge.Select(0, 2).Copy
        Else
            PinB_Edge = PinB_Edge.Select(1, 2).Copy
        End If

    Else
        If PinA_wave.Element(PinA_Edge.Element(0) - 1) > WindowWidth / 2 Then    ' find only falling Edge
            PinA_Edge = PinA_Edge.Select(0, 2).Copy
        Else
            PinA_Edge = PinA_Edge.Select(1, 2).Copy
        End If

        If PinB_wave.Element(PinB_Edge.Element(0) - 1) > WindowWidth / 2 Then    ' find only falling Edge
            PinB_Edge = PinB_Edge.Select(0, 2).Copy
        Else
            PinB_Edge = PinB_Edge.Select(1, 2).Copy
        End If

    End If

    Max_Samples = 100

    If Max_Samples > PinA_Edge.SampleSize Then
        Max_Samples = PinA_Edge.SampleSize
    End If

    If Max_Samples > PinB_Edge.SampleSize Then
        Max_Samples = PinB_Edge.SampleSize
    End If

    PinA_Edge = PinA_Edge.Select(0, 1, Max_Samples).Copy
    PinB_Edge = PinB_Edge.Select(0, 1, Max_Samples).Copy


    T_delay = PinA_Edge.Subtract(PinB_Edge).CalcMean
    EdgeMean = PinB_Edge.Differentiate.CalcMean

    While T_delay > EdgeMean * 0.5
        T_delay = T_delay - EdgeMean
    Wend

    While T_delay < (EdgeMean * -1) * 0.5
        T_delay = T_delay + EdgeMean
    Wend

    T_delay = T_delay * EFS

End Function
