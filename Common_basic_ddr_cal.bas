Attribute VB_Name = "Common_basic_ddr_cal"
Option Explicit

' ### Teradyne HISI UFLEX Basic Templates V13.00 ###

' HISI DDR Test Template
'
' Description:
'       Common function for DDR test with semi-automatic calibration.
'       For cases like SD6183 & SD6185, DDR IO neet to be calibrated before DC testing.
'       But the calibration result can not be passed to practical application in DFT mode.
'       So it's essential for ATE to transfer the calibration result to DC test configure.
'       This template provide the a flexiable function for calib-value preparation.
' This Template include DDR_CalibValue_Prepation function.
'
' DDR_CalibValue_Prepation
' Argument:
'       DigitalSourcePin: generally this pin is JTAG_TDI.
'       PreConditionPat: the pre-condition pattern for DDR configuration.
'       DigitalSourceSig: the digital source signal name. it shoule match the one specified in PreConditionPat.
'       ConfigValueArray: all the calib-values are stored in the DSPwave g_DDR_CAL_DSP defined in VBT_basic_ddr_cal_read.bas.
'                         you need to specify how to organise these values to config DDR IOs.
'                         when you fill the parameter with "0,2,2,1,3,3", the specified elements in g_DDR_CAL_DSP will be put in order to configure.


' Revision History:
' Date              Description                                                             Author
' 2017/6/17          Initial version                                                        Mason Song/Zoe Song
'
'
Public Function DDR_CalibValue_Prepation(DigitalSourcePin As PinList, _
                                         PreConditionPat As Pattern, _
                                         DigitalSourceSig As String, _
                                         ConfigValueArray As String)

    On Error GoTo errHandler

    Dim WaveDefinitionName As String
    WaveDefinitionName = "Src"

    Dim ConfigValueIndexArray() As String
    ConfigValueIndexArray = Split(ConfigValueArray, ",")

    Dim tmp_ValueSize As Long
    tmp_ValueSize = UBound(ConfigValueIndexArray)
    Dim tmp_DDR_CAL_WRITE As New DSPWave
    Call tmp_DDR_CAL_WRITE.CreateConstant(0, tmp_ValueSize + 1)

    Dim CalDataSampleSize As Long
    Dim i As Long
    Dim Site As Variant
    For Each Site In TheExec.Sites.Active
        CalDataSampleSize = g_DDR_CAL_READ_DSP(Site).SampleSize
        For i = 0 To tmp_ValueSize
            If Int(ConfigValueIndexArray(i)) > CalDataSampleSize - 1 Then

                If TheExec.RunMode = runModeProduction Then
                    TheExec.AddOutput "Please check if you have calibrated the DDRIO or if the data is stored in the correct position."
                    GoTo errHandler
                Else
                    MsgBox "Please check if you have calibrated the DDRIO or if the data is stored in the correct position."
                    Stop
                End If

            Else
                tmp_DDR_CAL_WRITE(Site).Element(i) = g_DDR_CAL_READ_DSP(Site).Element(Int(ConfigValueIndexArray(i)))
            End If
        Next i

        WaveDefinitionName = "Src_" & CStr(Site)
        TheExec.WaveDefinitions.CreateWaveDefinition WaveDefinitionName, tmp_DDR_CAL_WRITE, True

        With TheHdw.DSSC.Pins(DigitalSourcePin).Pattern(PreConditionPat).Source
            .Signals.Add (DigitalSourceSig)
            .Signals(DigitalSourceSig).WaveDefinitionName = WaveDefinitionName
            .Signals(DigitalSourceSig).LoadSettings
        End With

    Next Site

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function
