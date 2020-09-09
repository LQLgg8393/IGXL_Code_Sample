Attribute VB_Name = "VBT_basic_dc_fimv"
Option Explicit

' ### Teradyne HISI UFLEX Basic Templates V13.00 ###

' HISI OpenShort Test Template
' (c) Teradyne, Inc, 2016-2017
'     All Rights Reserved
' Inclusion of a copyright notice does not imply that this software has been
' published.
' This software is the trade secret information of Teradyne, Inc.
' Use of this software is only in accordance with the terms of a license
' agreement from Teradyne, Inc.
'
'
'Description:This template is designed for OpenShort test,it measure pins in serially.
'            It supports UP1600/UltraPAC80/US10G OpenShort measure.
'            The clamp voltage will protect the device.
'
'
'Arguments:
'          MeasPins     measure pins
'          ForceCurrentVal   force current value
'          VClampHi & VClampLo    clamp voltage
'
'
' Revision History:
' Version Number                    Date              Description                                                                                     Author
' V10.00.01                         2016/6/23                                                                                           Sunny   sunny.wang@teradyne.com
' V10.00.02                         2016/7/21        obtain the voltage: first strobe, then readback together                           Sunny   sunny.wang@teradyne.com
' V11.00.00                         2016/12/1        support sds                                                                        Sunny   sunny.wang@teradyne.com
' V11.02.00                         2017/05/25       support the clamp voltage settings of different instrument                         Sunny   sunny.wang@teradyne.com
'
'
'
' ================================================================================
'                             OpenShort_VDD_T_vbt
' ================================================================================
Public Function Basic_02_OS_VDD_PPMU(MeasPins As PinList, _
                                ForceCurrentVal As Double, _
                                Optional ForceCurrRange As String = "tl_GetPPMUForceCurrentRanges()", _
                                Optional InitWaitTime As Double = 0.0001, _
                                Optional SettlingTime As Double = 0.0001, _
                                Optional DriveLoPins As PinList, _
                                Optional DriveHiPins As PinList, _
                                Optional DriveZPins As PinList, _
                                Optional FloatPins As PinList, _
                                Optional Util1Pins As PinList, _
                                Optional Util0Pins As PinList) As Long
' EDITFORMAT1 1,,PinList,,Digital Pins Under Test,MeasPins|2,,Double,,Force current value,ForceCurrentVal|3,,String,,Force current range,ForceCurrRange|4,,Double,,Waitting time after applyleveltiming,InitWaitTime|5,,Double,,Waitting time after instrument setting,SettlingTime|6,,PinList,,Set pins to low,DriveLoPins|7,,PinList,,Set pins to high,DriveHiPins|8,,PinList,,Set pins to Z,DriveZPins|9,,PinList,,Disconnect pins,FloatPins|10,,PinList,,Relay connect to 1,Util1Pins|11,,PinList,,Relay connect to 0,Util0Pins

    On Error GoTo errHandler
    
    
    Dim i_VoltMeas_PLD As New PinListData
    Dim i_ForceCurrRange_dbl As Double
    
    If ForceCurrentVal <= 0 Then
        If TheExec.RunMode = runModeProduction Then
            TheExec.AddOutput "FIMV Error 0004: ForceCurrentVal is wrong!"
            Exit Function
        Else
            MsgBox "FIMV Error 0004: ForceCurrentVal is wrong!"
            Stop
        End If
    End If
    
    
    i_ForceCurrRange_dbl = StrToDbl(ForceCurrRange)
    
' ================================================================================
'                             Initial Setting
' ================================================================================
    'Set drive state on specified utility pins
    Call tl_SetTestState(True, True, True, 1, InitWaitTime, DriveLoPins, DriveHiPins, DriveZPins, FloatPins, Util1Pins, Util0Pins)
    
    'call function FIMV
    DC_FIMV_Method MeasPins, ForceCurrentVal, in_VClampHi:=1.5, in_VClampLo:=-1, out_VoltMeas:=i_VoltMeas_PLD, _
                   in_ForceCurrRange:=i_ForceCurrRange_dbl, in_SampleSize:=8, in_SettlingTime:=SettlingTime, in_PE_Connect:=False, in_ForceV_GND:=True
    
    ' For Open result output
    TheExec.Flow.TestLimit ResultVal:=i_VoltMeas_PLD, forceVal:=ForceCurrentVal, unit:=unitVolt, _
        scaletype:=scaleNone, forceResults:=tlForceFlow, Forceunit:=unitAmp
    
'    ' For Short result output
'    TheExec.Flow.TestLimit ResultVal:=i_VoltMeas_PLD, forceVal:=ForceCurrentVal, unit:=unitVolt, _
'        scaletype:=scaleNone, ForceResults:=tlForceFlow, forceUnit:=unitAmp
    
    
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function
' HISI OpenShort Test Template
' (c) Teradyne, Inc, 2016-2017
'     All Rights Reserved
' Inclusion of a copyright notice does not imply that this software has been
' published.
' This software is the trade secret information of Teradyne, Inc.
' Use of this software is only in accordance with the terms of a license
' agreement from Teradyne, Inc.
'
'
'Description:This template is designed for OpenShort test,it measure pins in serially.
'            It supports UP1600/UltraPAC80/US10G OpenShort measure.
'            The clamp voltage will protect the device.
'
'
'Arguments:
'          MeasPins     measure pins
'          ForceCurrentVal   force current value
'          VClampHi & VClampLo    clamp voltage
'
'
' Revision History:
' Version Number                    Date              Description                                                                                     Author
' V10.00.01                         2016/6/23                                                                                           Sunny   sunny.wang@teradyne.com
' V10.00.02                         2016/7/21        obtain the voltage: first strobe, then readback together                           Sunny   sunny.wang@teradyne.com
' V11.00.00                         2016/12/1        support sds                                                                        Sunny   sunny.wang@teradyne.com
' V11.02.00                         2017/05/25       support the clamp voltage settings of different instrument                         Sunny   sunny.wang@teradyne.com
' V11.03.00                         2017/09/25       support the force current                                                          Sunny   sunny.wang@teradyne.com
'
'
' ================================================================================
'                             OpenShort_VSS_T_vbt
' ================================================================================
Public Function Basic_02_OS_VSS_PPMU(MeasPins As PinList, _
                                ForceCurrentVal As Double, _
                                Optional ForceCurrRange As String = "tl_GetPPMUForceCurrentRanges()", _
                                Optional InitWaitTime As Double = 0.0001, _
                                Optional SettlingTime As Double = 0.0001, _
                                Optional DriveLoPins As PinList, _
                                Optional DriveHiPins As PinList, _
                                Optional DriveZPins As PinList, _
                                Optional FloatPins As PinList, _
                                Optional Util1Pins As PinList, _
                                Optional Util0Pins As PinList) As Long
' EDITFORMAT1 1,,PinList,,Digital Pins Under Test,MeasPins|2,,Double,,Force current value,ForceCurrentVal|3,,String,,Force current range,ForceCurrRange|4,,Double,,Waitting time after applyleveltiming,InitWaitTime|5,,Double,,Waitting time after instrument setting,SettlingTime|6,,PinList,,Set pins to low,DriveLoPins|7,,PinList,,Set pins to high,DriveHiPins|8,,PinList,,Set pins to Z,DriveZPins|9,,PinList,,Disconnect pins,FloatPins|10,,PinList,,Relay connect to 1,Util1Pins|11,,PinList,,Relay connect to 0,Util0Pins

    On Error GoTo errHandler
    
    
    Dim i_VoltMeas_PLD As New PinListData
    Dim i_ForceCurrRange_dbl As Double
    
    
     If ForceCurrentVal >= 0 Then
        If TheExec.RunMode = runModeProduction Then
            TheExec.AddOutput "FIMV Error 0004: ForceCurrentVal is wrong!"
            Exit Function
        Else
            MsgBox "FIMV Error 0004: ForceCurrentVal is wrong!"
            Stop
        End If
    End If
    
    
    i_ForceCurrRange_dbl = StrToDbl(ForceCurrRange)
    
' ================================================================================
'                             Initial Setting
' ================================================================================
    'Set drive state on specified utility pins
    Call tl_SetTestState(True, True, True, 1, InitWaitTime, DriveLoPins, DriveHiPins, DriveZPins, FloatPins, Util1Pins, Util0Pins)
    
    'call function FIMV
    DC_FIMV_Method MeasPins, ForceCurrentVal, in_VClampHi:=1.5, in_VClampLo:=-1, out_VoltMeas:=i_VoltMeas_PLD, _
                   in_ForceCurrRange:=i_ForceCurrRange_dbl, in_SampleSize:=8, in_SettlingTime:=SettlingTime, in_PE_Connect:=False, in_ForceV_GND:=True
    
    ' For open result output
    TheExec.Flow.TestLimit ResultVal:=i_VoltMeas_PLD, forceVal:=ForceCurrentVal, unit:=unitVolt, _
        scaletype:=scaleNone, forceResults:=tlForceFlow, Forceunit:=unitAmp
        
'    ' For short result output
'    TheExec.Flow.TestLimit ResultVal:=i_VoltMeas_PLD, forceVal:=ForceCurrentVal, unit:=unitVolt, _
'        scaletype:=scaleNone, ForceResults:=tlForceFlow, forceUnit:=unitAmp
        
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

' HISI FIMV Test Template
' (c) Teradyne, Inc, 2017-2017
'     All Rights Reserved
' Inclusion of a copyright notice does not imply that this software has been
' published.
' This software is the trade secret information of Teradyne, Inc.
' Use of this software is only in accordance with the terms of a license
' agreement from Teradyne, Inc.
'
'
'Description:This template is designed for FIMV test,it measure pins in serially.
'            It supports UP1600/UltraPAC80/US10G FIMV measurement.
'            The clamp voltage will protect the device.
'
'
'Arguments:
'          MeasPins     measure pins
'          ForceCurrentVal   force current value
'          VClampHi & VClampLo    clamp voltage
'
'
' Revision History:
' Version Number                    Date                                Description                                                                                       Author
' V11.02.00                      2017/05/25                                                                                                            Sunny   sunny.wang@teradyne.com
' V11.03.00                      2017/09/25                         support the force current                                                          Sunny   sunny.wang@teradyne.com
'


' ================================================================================
'                             DDR_VREF_FIMV
' ================================================================================
Public Function Basic_02_DDR_VREF(MeasPins As PinList, _
                                Optional PatName As Pattern, _
                                Optional MeasCount As Long = 1, _
                                Optional InitWaitTime As Double = 0.0001, _
                                Optional SettlingTime As Double = 0.0001, _
                                Optional WaitTime As Double = 0.001, _
                                Optional PE_Connect As Boolean = True, _
                                Optional PatResult As Boolean = True, _
                                Optional DriveLoPins As PinList, _
                                Optional DriveHiPins As PinList, _
                                Optional DriveZPins As PinList, _
                                Optional FloatPins As PinList, _
                                Optional Util1Pins As PinList, _
                                Optional Util0Pins As PinList) As Long
' EDITFORMAT1 1,,PinList,,Digital Pins Under Test,MeasPins|2,,Pattern,,Choose pattern,PatName|3,,Long,,Measure count/should equal to CPUFlagA count in pattern,MeasCount|4,,Double,,Waitting time after applyleveltiming,InitWaitTime|5,,Double,,Waitting time after instrument setting,SettlingTime|6,,Double,,WaitTime after CPU Flag Set,WaitTime|7,,Boolean,,Whether connect PE,PE_Connect|8,,Boolean,,Whether print pattern result,PatResult|9,,PinList,,Set pins to low,DriveLoPins|10,,PinList,,Set pins to high,DriveHiPins|12,,PinList,,Set pins to Z,DriveZPins|11,,PinList,,Disconnect pins,FloatPins|13,,PinList,,Relay connect to 1,Util1Pins|14,,PinList,,Relay connect to 0,Util0Pins

    On Error GoTo errHandler
    
    Dim i_VoltMeas_PLD As New PinListData
    Dim i_Func_Result_SBOOL As New SiteBoolean
    Dim i_VClampHi_dbl As Double
    Dim i_VClampLo_dbl As Double
    Dim i As Long
    
' ================================================================================
'                             Initial Setting
' ================================================================================
    'Set drive state on specified utility pins
    Call tl_SetTestState(True, True, True, 1, InitWaitTime, DriveLoPins, DriveHiPins, DriveZPins, FloatPins, Util1Pins, Util0Pins)
    
    ' Get default clamp voltage
    i_VClampHi_dbl = TheHdw.PPMU.Pins(MeasPins).ClampVHi
    i_VClampLo_dbl = TheHdw.PPMU.Pins(MeasPins).ClampVLo
    
    If nonblank(PatName) Then TheHdw.Patterns(PatName).Start
' ================================================================================
'                             Measurement
' ================================================================================
    ' Disconnect PE
    TheHdw.Digital.Pins(MeasPins).Disconnect
    ' PPMU Setup
    With TheHdw.PPMU.Pins(MeasPins)
         .ForceI 0, 0.000005
         .ClampVHi = 1.7
         .ClampVLo = 0
         .Gate = tlOn
         .Connect
    End With
    TheHdw.Wait SettlingTime
    
    ' Measurement
    If MeasCount = 1 Then
        i_VoltMeas_PLD = TheHdw.PPMU.Pins(MeasPins).Read(tlPPMUReadMeasurements, 8)
        TheExec.Flow.TestLimit ResultVal:=i_VoltMeas_PLD, forceVal:=0, unit:=unitVolt, scaletype:=scaleNone, forceResults:=tlForceFlow, Forceunit:=unitAmp
    Else
        For i = 0 To MeasCount - 1
            TheHdw.Digital.Patgen.FlagWait cpuA, 0
            TheHdw.Wait WaitTime
            
            i_VoltMeas_PLD = TheHdw.PPMU.Pins(MeasPins).Read(tlPPMUReadMeasurements, 8)
            TheExec.Flow.TestLimit ResultVal:=i_VoltMeas_PLD, forceVal:=0, unit:=unitVolt, scaletype:=scaleNone, forceResults:=tlForceFlow, Forceunit:=unitAmp
            
            TheExec.Flow.TestLimitIndex = 0
            TheHdw.Digital.Patgen.Continue 0, cpuA
        Next i
    End If
    
    TheHdw.Digital.Patgen.HaltWait
   
' ================================================================================
'                             Reset and DataLog
' ================================================================================
    TheHdw.PPMU.Pins(MeasPins).ClampVHi = i_VClampHi_dbl
    If i_VClampLo_dbl < -1 Then
        TheHdw.PPMU.Pins(MeasPins).ClampVLo = -1
    Else
        TheHdw.PPMU.Pins(MeasPins).ClampVLo = i_VClampLo_dbl
    End If
    ' disconnect PPMU
    TheHdw.PPMU.Pins(MeasPins).Disconnect
    
    ' Test Instances sequence decide the PE connection
    If PE_Connect Then TheHdw.Digital.Pins(MeasPins).Connect
        
    'Func Reslut
    If PatResult Then
        If nonblank(PatName) Then
            i_Func_Result_SBOOL = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
            TheExec.Flow.TestLimit ResultVal:=i_Func_Result_SBOOL, PinName:="FuncResults", lowval:=-1, hival:=-1
        End If
    End If
    
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function


' HISI FIMV Test Template
' (c) Teradyne, Inc, 2017-2017
'     All Rights Reserved
' Inclusion of a copyright notice does not imply that this software has been
' published.
' This software is the trade secret information of Teradyne, Inc.
' Use of this software is only in accordance with the terms of a license
' agreement from Teradyne, Inc.
'
'
'Description:This template is designed for FIMV test,it measure pins in serially.
'            It supports UP1600/UltraPAC80/US10G FIMV measurement.
'            The clamp voltage will protect the device.
'
'
'Arguments:
'          MeasPins     measure pins
'          ForceCurrentVal   force current value
'          VClampHi & VClampLo    clamp voltage
'
'
' Revision History:
' Version Number                    Date                         Description                                                                                       Author
' V11.02.00                      2017/05/25                                                                                                            Sunny   sunny.wang@teradyne.com
' V11.03.00                      2017/09/25                                                                                                            Sunny   sunny.wang@teradyne.com
'


' ================================================================================
'                             DCM_FIMV
' ================================================================================
Public Function Basic_02_DCM_FIMV(MeasPins As PinList, _
                                Optional PatName As Pattern, _
                                Optional InitWaitTime As Double = 0.0001, _
                                Optional SettlingTime As Double = 0.0001, _
                                Optional PE_Connect As Boolean = True, _
                                Optional PatResult As Boolean = True, _
                                Optional DriveLoPins As PinList, _
                                Optional DriveHiPins As PinList, _
                                Optional DriveZPins As PinList, _
                                Optional FloatPins As PinList, _
                                Optional Util1Pins As PinList, _
                                Optional Util0Pins As PinList) As Long
' EDITFORMAT1 1,,PinList,,Digital Pins Under Test,MeasPins|2,,Pattern,,Choose pattern,PatName|3,,Double,,Waitting time after applyleveltiming,InitWaitTime|4,,Double,,Waitting time after instrument setting,SettlingTime|5,,Boolean,,Whether connect PE,PE_Connect|6,,Boolean,,Whether print pattern result,PatResult|7,,PinList,,Set pins to low,DriveLoPins|8,,PinList,,Set pins to high,DriveHiPins|9,,PinList,,Set pins to Z,DriveZPins|10,,PinList,,Disconnect pins,FloatPins|11,,PinList,,Relay connect to 1,Util1Pins|12,,PinList,,Relay connect to 0,Util0Pins

    On Error GoTo errHandler
    
    
    Dim i_VoltMeas_PLD As New PinListData
    Dim i_Func_Result_SBOOL As New SiteBoolean
 

' ================================================================================
'                             Initial Setting
' ================================================================================
    'Set drive state on specified utility pins
    Call tl_SetTestState(True, True, True, 1, InitWaitTime, DriveLoPins, DriveHiPins, DriveZPins, FloatPins, Util1Pins, Util0Pins)
    
    ' call function FIMV
    DC_FIMV_Method MeasPins, in_ForceCurrentVal:=0, in_VClampHi:=2, in_VClampLo:=0, out_VoltMeas:=i_VoltMeas_PLD, out_FuncResult:=i_Func_Result_SBOOL, _
                   in_PatName:=PatName, in_ForceCurrRange:=0.000005, in_SampleSize:=8, in_SettlingTime:=SettlingTime, in_PatResult:=PatResult, in_PE_Connect:=PE_Connect, in_ForceV_GND:=False
    
    TheExec.Flow.TestLimit ResultVal:=i_VoltMeas_PLD, forceVal:=0, unit:=unitVolt, _
        scaletype:=scaleNone, forceResults:=tlForceFlow, Forceunit:=unitAmp
        'Func Reslut
    If PatResult Then
        If nonblank(PatName) Then
           TheExec.Flow.TestLimit ResultVal:=i_Func_Result_SBOOL, PinName:="FuncResults", lowval:=-1, hival:=-1
        End If
    End If
    
    
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

' HISI FIMV Test Template
' (c) Teradyne, Inc, 2017-2017
'     All Rights Reserved
' Inclusion of a copyright notice does not imply that this software has been
' published.
' This software is the trade secret information of Teradyne, Inc.
' Use of this software is only in accordance with the terms of a license
' agreement from Teradyne, Inc.
'
'
'Description:This template is designed for FIMV test,it measure pins in serially.
'            It supports UP1600/UltraPAC80/US10G FIMV measurement.
'            The clamp voltage will protect the device.
'
'
'Arguments:
'          MeasPins     measure pins
'          ForceCurrentVal   force current value
'          VClampHi & VClampLo    clamp voltage
'
'
' Revision History:
' Version Number                    Date                         Description                                                                                       Author
' V11.02.00                      2017/06/05         After halting the PatGen, use PPMU to measure the current.                                         Sunny   sunny.wang@teradyne.com
' V11.03.00                      2017/09/25                                                                                                            Sunny   sunny.wang@teradyne.com'


' ================================================================================
'                             FIMV
' ================================================================================
Public Function Basic_02_DC_FIMV(MeasPins As PinList, ForceCurrentVal As Double, _
                                    VClampHi As Double, _
                                    VClampLo As Double, _
                                    Optional PatName As Pattern, _
                                    Optional ForceCurrRange As String = "tl_GetPPMUForceCurrentRanges()", _
                                    Optional InitWaitTime As Double = 0.0001, _
                                    Optional SettlingTime As Double = 0.0001, _
                                    Optional SampleSize As Long = 8, _
                                    Optional PE_Connect As Boolean = True, _
                                    Optional ForceV_GND As Boolean = False, _
                                    Optional PatResult As Boolean = True, _
                                    Optional ConnectAllPins As Boolean = True, _
                                    Optional LoadLevels As Boolean = True, _
                                    Optional LoadTiming As Boolean = True, _
                                    Optional relayMode As tlRelayMode = 1, _
                                    Optional DriveLoPins As PinList, _
                                    Optional DriveHiPins As PinList, _
                                    Optional DriveZPins As PinList, _
                                    Optional FloatPins As PinList, _
                                    Optional Util1Pins As PinList, _
                                    Optional Util0Pins As PinList) As Long
' EDITFORMAT1 1,,PinList,,Digital Pins Under Test,MeasPins|3,,Double,,Force current value,ForceCurrentVal|8,,Double,,Maximum Output Voltage,VClampHi|9,,Double,,Minimum Output Voltage,VClampLo|2,,Pattern,,Choose pattern,PatName|4,,String,,Force current range,ForceCurrRange|6,,Double,,Waitting time after applyleveltiming,InitWaitTime|7,,Double,,Waitting time after instrumet setting,SettlingTime|5,,Long,,Samples in each strobe,SampleSize|10,,Boolean,,Whether connect PE,PE_Connect|11,,Boolean,,Whether force 0V,ForceV_GND|12,,Boolean,,Whether print pattern result,PatResult|16,,Boolean,Pins State,All pins connect,ConnectAllPins|13,,Boolean,Level Timing,Level setting,LoadLevels|14,,Boolean,,Timing setting,LoadTiming|15,,tlRelayMode,,Power_Up mode,RelayMode|17,,PinList,,Set pins to low,DriveLoPins|18,,PinList,,Set pins to high,DriveHiPins|19,,PinList,,Set pins to Z,DriveZPins|20,,PinList,,Disconnect pins,FloatPins|21,,PinList,,Relay connect to 1,Util1Pins|22,,PinList,,Relay connect to 0,Util0Pins

    On Error GoTo errHandler
    
    
    Dim i_VoltMeas_PLD As New PinListData
    Dim i_Func_Result_SBOOL As New SiteBoolean
    Dim i_ForceCurrRange_dbl As Double

' ================================================================================
'                             Initial Setting
' ================================================================================
    'Set drive state on specified utility pins
    Call tl_SetTestState(ConnectAllPins, LoadLevels, LoadTiming, relayMode, InitWaitTime, DriveLoPins, DriveHiPins, DriveZPins, FloatPins, Util1Pins, Util0Pins)
    
    i_ForceCurrRange_dbl = StrToDbl(ForceCurrRange)
    
    'call function FIMV
    DC_FIMV_Method MeasPins, ForceCurrentVal, VClampHi, VClampLo, i_VoltMeas_PLD, i_Func_Result_SBOOL, PatName, _
                   i_ForceCurrRange_dbl, SampleSize, SettlingTime, PatResult, PE_Connect, ForceV_GND
    
    TheExec.Flow.TestLimit ResultVal:=i_VoltMeas_PLD, forceVal:=ForceCurrentVal, unit:=unitVolt, _
        scaletype:=scaleNone, forceResults:=tlForceFlow, Forceunit:=unitAmp
        
    'Func Reslut
    If PatResult Then
        If nonblank(PatName) Then
           TheExec.Flow.TestLimit ResultVal:=i_Func_Result_SBOOL, PinName:="FuncResults", lowval:=-1, hival:=-1
        End If
    End If
    
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

' ================================================================================
'                             FIMV
' ================================================================================
Private Function DC_FIMV_Method(in_MeasPins As PinList, in_ForceCurrentVal As Double, _
                                in_VClampHi As Double, in_VClampLo As Double, _
                                out_VoltMeas As PinListData, _
                                Optional out_FuncResult As SiteBoolean, _
                                Optional in_PatName As Pattern, _
                                Optional in_ForceCurrRange As Double, _
                                Optional in_SampleSize As Long, _
                                Optional in_SettlingTime As Double, _
                                Optional in_PatResult As Boolean, _
                                Optional in_PE_Connect As Boolean, _
                                Optional in_ForceV_GND As Boolean) As Long

    On Error GoTo errHandler

    Dim i_PerPin_str() As String
    Dim i_PinNum_lng As Long
    Dim i_PinIndex_lng As Long
    Dim i_DigtPin_str As String
    Dim i_UPAC_CapPin_str As String
    Dim i_UPAC_SrcPin_str As String
    Dim i_US10GPin_str As String
    Dim i_VClampHi_UP1600_dbl As Double
    Dim i_VClampLo_UP1600_dbl As Double
    Dim i_VClampHi_UPAC_dbl As Double
    Dim i_VClampLo_UPAC_dbl As Double
    
' ================================================================================
'                             Initial Setting
' ================================================================================
    ' run pattern and ensure the pattern setting before measurement
    If Not (in_PatName Is Nothing) Then
       If in_PatName <> "" Then
        TheHdw.Patterns(in_PatName).Load
        TheHdw.Patterns(in_PatName).Start
        TheHdw.Digital.Patgen.HaltWait
       End If
    End If
        
    ' handle input pins
    Call TheExec.DataManager.DecomposePinList(in_MeasPins, i_PerPin_str(), i_PinNum_lng)
    SortPinType in_MeasPins, i_DigtPin_str, i_UPAC_CapPin_str, i_UPAC_SrcPin_str, i_US10GPin_str
    
    ' disconnect PE
    If i_DigtPin_str <> "" Then
         TheHdw.Digital.Pins(i_DigtPin_str).Disconnect
    End If
    If i_US10GPin_str <> "" Then
         TheHdw.Serial.Pins(i_US10GPin_str).Disconnect
    End If
    If i_UPAC_CapPin_str <> "" Then
         TheHdw.UltraCapture.Pins(i_UPAC_CapPin_str).Disconnect
    End If
    If i_UPAC_SrcPin_str <> "" Then
         TheHdw.UltraSource.Pins(i_UPAC_SrcPin_str).Disconnect
    End If
    
    ' connect PPMU
    TheHdw.PPMU.Pins(in_MeasPins).Connect
    If in_ForceV_GND Then
       TheHdw.PPMU.Pins(in_MeasPins).ForceV 0
       TheHdw.PPMU.Pins(in_MeasPins).Gate = tlOn
    End If

' ================================================================================
'                             Parallel Measurement
' ================================================================================
    If in_VClampLo = in_VClampHi Then TheExec.AddOutput "FIMV Error 0002:VClampHi=VClampLo. Please ensure the clamp voltage setting.", vbRed, True

    ' Get default clamp voltage
    If i_DigtPin_str <> "" Then
       If in_VClampLo = 0 And in_VClampHi = 0 Then
         ' keep instrument default value
         TheExec.AddOutput "FIMV Error 0003:Please set clamp voltage. If you don't do this, use instrument default clamp voltage.", vbBlue, True
       Else
         i_VClampHi_UP1600_dbl = TheHdw.PPMU.Pins(i_DigtPin_str).ClampVHi
         i_VClampLo_UP1600_dbl = TheHdw.PPMU.Pins(i_DigtPin_str).ClampVLo
       End If
    End If
    
    If i_UPAC_CapPin_str <> "" Then
       If in_VClampLo = 0 And in_VClampHi = 0 Then
       ' keep instrument default value
        TheExec.AddOutput "FIMV Error 0003:Please set clamp voltage. If you don't do this, use instrument default clamp voltage.", vbBlue, True
       Else
        i_VClampHi_UPAC_dbl = TheHdw.PPMU.Pins(i_UPAC_CapPin_str).ClampVHi
        i_VClampLo_UPAC_dbl = TheHdw.PPMU.Pins(i_UPAC_CapPin_str).ClampVLo
       End If
    End If
    
    If i_UPAC_SrcPin_str <> "" Then
       If in_VClampLo = 0 And in_VClampHi = 0 Then
       ' keep instrument default value
        TheExec.AddOutput "FIMV Error 0003:Please set clamp voltage. If you don't do this, use instrument default clamp voltage.", vbBlue, True
       Else
        i_VClampHi_UPAC_dbl = TheHdw.PPMU.Pins(i_UPAC_SrcPin_str).ClampVHi
        i_VClampLo_UPAC_dbl = TheHdw.PPMU.Pins(i_UPAC_SrcPin_str).ClampVLo
       End If
    End If
    
    If in_VClampLo = 0 And in_VClampHi = 0 Then
    ' keep instrument default value
    Else
    ' set clamp voltage
      SetClampVoltage i_DigtPin_str, i_UPAC_CapPin_str, i_UPAC_SrcPin_str, in_VClampHi, in_VClampLo, in_VClampHi, in_VClampLo
    End If
    
    'modify to parallel for AP project
    With TheHdw.PPMU.Pins(in_MeasPins)
         .ForceI in_ForceCurrentVal, in_ForceCurrRange
         .Gate = tlOn
    End With
    TheHdw.Wait in_SettlingTime

    out_VoltMeas = TheHdw.PPMU.Pins(in_MeasPins).Read(tlPPMUReadMeasurements, in_SampleSize)
    ' OS, connect pin to GND
    If in_ForceV_GND Then TheHdw.PPMU.Pins(in_MeasPins).ForceV 0
    
' ================================================================================
'                             Reset and DataLog
' ================================================================================
    If in_VClampLo = 0 And in_VClampHi = 0 Then
    ' keep instrument default value
    Else
    ' Restore default clamp voltage
      SetClampVoltage i_DigtPin_str, i_UPAC_CapPin_str, i_UPAC_SrcPin_str, i_VClampHi_UP1600_dbl, i_VClampLo_UP1600_dbl, i_VClampHi_UPAC_dbl, i_VClampLo_UPAC_dbl
    End If

    ' disconnect PPMU
    TheHdw.PPMU.Pins(in_MeasPins).Disconnect
    
    ' Test Instances sequence decide the PE connection
    If in_PE_Connect Then
        If i_DigtPin_str <> "" Then
             TheHdw.Digital.Pins(i_DigtPin_str).Connect
        End If
        If i_US10GPin_str <> "" Then
             TheHdw.Serial.Pins(i_US10GPin_str).Connect
        End If
    End If
    
    If in_PatResult Then
        If Not (in_PatName Is Nothing) Then
           If in_PatName <> "" Then
              out_FuncResult = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
           End If
        End If
    End If
    
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function SortPinType(in_MeasPins As PinList, out_DigtPin As String, _
                             out_UPAC_CapPin As String, out_UPAC_SrcPin As String, out_US10GPin As String) As Long

    On Error GoTo errHandler
       
    Dim i_PerPin_str() As String
    Dim i_PinNum_lng As Long
    Dim i_PinIndex_lng As Long
    Dim i_PinInstrument_str As String
    
        ' handle the input pins
    Call TheExec.DataManager.DecomposePinList(in_MeasPins, i_PerPin_str(), i_PinNum_lng)
    For i_PinIndex_lng = LBound(i_PerPin_str) To UBound(i_PerPin_str)
       i_PinInstrument_str = TheExec.DataManager.ChannelType(i_PerPin_str(i_PinIndex_lng))
       If i_PinInstrument_str = "I/O" Then
            If out_DigtPin = "" Then
                out_DigtPin = i_PerPin_str(i_PinIndex_lng)
            Else
                out_DigtPin = out_DigtPin + "," + i_PerPin_str(i_PinIndex_lng)
            End If
       ElseIf i_PinInstrument_str = "UltraCapture" Then
             If out_UPAC_CapPin = "" Then
                 out_UPAC_CapPin = i_PerPin_str(i_PinIndex_lng)
             Else
                 out_UPAC_CapPin = out_UPAC_CapPin + "," + i_PerPin_str(i_PinIndex_lng)
             End If
       ElseIf i_PinInstrument_str = "UltraSource" Then
             If out_UPAC_SrcPin = "" Then
                 out_UPAC_SrcPin = i_PerPin_str(i_PinIndex_lng)
             Else
                 out_UPAC_SrcPin = out_UPAC_SrcPin + "," + i_PerPin_str(i_PinIndex_lng)
             End If
       ElseIf i_PinInstrument_str = "Serial10G" Then
            If out_US10GPin = "" Then
                out_US10GPin = i_PerPin_str(i_PinIndex_lng)
            Else
                out_US10GPin = out_US10GPin + "," + i_PerPin_str(i_PinIndex_lng)
            End If
       Else
            If TheExec.RunMode = runModeProduction Then
                TheExec.AddOutput "FIMV Error 0001: Wrong Input Pin!"
                Exit Function
            Else
                'Error message here if get unexpexted PinType
                MsgBox "FIMV Error 0001: Wrong Input Pin!"
                Stop
            End If
       End If
    Next i_PinIndex_lng
   

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function


Private Function SetClampVoltage(in_DigtPin As String, in_UPAC_CapPin As String, in_UPAC_SrcPin As String, _
                             in_VClampHi_UP1600 As Double, in_VClampLo_UP1600 As Double, _
                             in_VClampHi_UPAC As Double, in_VClampLo_UPAC As Double) As Long

    On Error GoTo errHandler

    If in_DigtPin <> "" Then
         TheHdw.PPMU.Pins(in_DigtPin).ClampVHi = in_VClampHi_UP1600
         If in_VClampLo_UP1600 < -1 Then
            TheHdw.PPMU.Pins(in_DigtPin).ClampVLo = -1
         Else
            TheHdw.PPMU.Pins(in_DigtPin).ClampVLo = in_VClampLo_UP1600
         End If
    End If
    If in_UPAC_CapPin <> "" Then
        TheHdw.PPMU.Pins(in_UPAC_CapPin).ClampVHi = in_VClampHi_UPAC
        TheHdw.PPMU.Pins(in_UPAC_CapPin).ClampVLo = in_VClampLo_UPAC
    End If
    If in_UPAC_SrcPin <> "" Then
        TheHdw.PPMU.Pins(in_UPAC_SrcPin).ClampVHi = in_VClampHi_UPAC
        TheHdw.PPMU.Pins(in_UPAC_SrcPin).ClampVLo = in_VClampLo_UPAC
    End If

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function SetCase(Site As String, InstanceName As String, _
                        PinList As String, Util0Pins As String, Util1Pins As String, _
                        ChannelType As String, TheExec As String, TheHdw As String)



End Function



