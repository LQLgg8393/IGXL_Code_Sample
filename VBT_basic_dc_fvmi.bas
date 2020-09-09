Attribute VB_Name = "VBT_basic_dc_fvmi"
Option Explicit

' ### Teradyne HISI UFLEX Basic Templates V12.00 ###

' HISI IDDQ Test Template
' (c) Teradyne, Inc, 2016-
'     All Rights Reserved
' Inclusion of a copyright notice does not imply that this software has been
' published.
' This software is the trade secret information of Teradyne, Inc.
' Use of this software is only in accordance with the terms of a license
' agreement from Teradyne, Inc.
'


'Description:This template is designed for Leakage test:
'                VDD:Force High voltage and test leakage
'                VSS:Force Low  voltage and test leakage
'                If VForceHigh not be set will not test VDD and VForceLow not be set will not test VSS(For PullUP and PullDown)
'                If the value of power is not equal to the high force voltage,it will give an message and sotp test
'Arguments:
' Basic_03_DC_Leakage
'                PatName:PreCondition Pat
'                PowerPin:The Source Power of the pin being test
'                InitWaitTime:Relay Wait Time
'                WaitTime:PPMU Wait Time
'                VForceHigh: If Blank Will not Test VDD
'                VForceLow: If Blank Will not Test VSS
' Basic_03_DC_IOHL
'                VForceValue: PPMU force voltage for IOHL test(VOH or VOL)
'                PinsNumPerGrp: Pins number per group ( for IOHL group test as Current load capacity of power is limited)
'                MeasWaitTime: PPMU wait time after force voltage
' Basic_03_DCM_RES
'                DCM_OUT_Pin: pin to measure, DCM OUT pin name
'                CORE_POWER_Pin: The Source Power name for the pin being test
'
'
' Revision History:
' Date        Description                                                                                               Author
' 20160714    Subroutines within the "PowerPin Number Check" process
'             Did not consider the situation of DCVS, only considered the DCVI                                          Alex Sun
'             Added DCVS to the situation

' 20160715    Detect Pin Type And Judge "section in the leakage_ParCheck subfunction
'             DCVI or DCVS read back a value of 0 when offline, and VForceHigh                                          Alex Sun
'             may be left blank during test VDD, adding judgment conditions to circumvent these two cases

' 20171125    Added DCM divider network pull-down resistance test method                                                shizhenhua

' 20171126    Added five functions, one of which is PUBLIC: Basic_03_DC_IHOL for testing IOHL
'             four additional functions: Pri_PPMU_fvmi, Pri_GetMeterRangeFromLimit, Pri_Cut_Pinlist, Pri_DatalogOutput
'                         Pri_PPMU_fvmi as ppmu fvmi underlying function
'                         Pri_GetMeterRangeFromLimit Used to find the appropriate measurement range from Use-Limit
'                         Pri_Cut_Pinlist is used to split the given PinList
'                         Pri_DatalogOutput is used as a datalog output format function to report errors                Xuyiming
'
' 20180117    Edited Chinese characters into English for resolving string missing(e.g. exit function)
'                                              when importing assic file to generate new project file                   Zoe Song
'             Unified error msg. output format as production mode and debug mode
'             Added to check if measure current range is reasonable of leakage test, if it less than the measure current, the current be be clamp






' ================================================================================
'                        Leakage_T_vbt
' ================================================================================
Public Function Basic_03_DC_Leakage(PatName As Pattern, _
                            powerPin As PinList, _
                            MeasPins As PinList, _
                            VForceHigh As String, _
                            VForceLow As String, _
                            Optional MeasIRangeHi As String = 30 * ma, _
                            Optional MeasIRangeLo As String = 30 * ma, _
                            Optional SampleSize As Long = 10, _
                            Optional WaitTime As Double = 0.0001, _
                            Optional CheckPG As PFType = pfAlways, _
                            Optional InitWaitTime As Double = 0.001, _
                            Optional ConnectAllPins As Boolean = True, _
                            Optional LoadLevels As Boolean = True, _
                            Optional LoadTiming As Boolean = True, _
                            Optional relayMode As tlRelayMode = 1, _
                            Optional DriveLoPins As PinList, _
                            Optional DriveHiPins As PinList, _
                            Optional DriveZPins As PinList, _
                            Optional FloatPins As PinList, _
                            Optional Util0Pins As PinList, _
                            Optional Util1Pins As PinList, _
                            Optional SortPin As Boolean = False) As Long
' EDITFORMAT1 1,,Pattern,,,PatName|
' EDITFORMAT1 2,,PinList,,Only One,PowerPin|
' EDITFORMAT1 3,,PFType,,,CheckPG|
' EDITFORMAT1 4,,Boolean,,,ApplyLeveltiming|
' EDITFORMAT1 5,,Boolean,,,ConnectAllPins|
' EDITFORMAT1 6,,Boolean,,,LoadLevels|
' EDITFORMAT1 7,,Boolean,,,LoadTiming|
' EDITFORMAT1 8,,tlRelayMode,,,RelayMode|
' EDITFORMAT1 9,,Double,,Apply lvl timing and Relay Setup Time,InitWaitTime|
' EDITFORMAT1 10,,PinList,PPMU,,MeasPins|
' EDITFORMAT1 11,,Long,,PPMU,SampleSize|
' EDITFORMAT1 12,,String,,Blank For PullUp Test,VForceHigh|
' EDITFORMAT1 13,,String,,Blank For PullDown Test,VForceLow|
' EDITFORMAT1 14,,String,,Must Be Set,MeasIRangeHi|
' EDITFORMAT1 15,,String,,Must Be Set,MeasIRangeLo|
' EDITFORMAT1 16,,Double,,PPMU Setup Time,WaitTime|
' EDITFORMAT1 17,,PinList,Pin States,,DriveLoPins|
' EDITFORMAT1 18,,PinList,,,DriveHiPins|
' EDITFORMAT1 19,,PinList,,,DriveZPins|
' EDITFORMAT1 20,,PinList,,,DisablePins|
' EDITFORMAT1 21,,PinList,,,FloatPins|
' EDITFORMAT1 22,,PinList,,,Util0Pins|
' EDITFORMAT1 23,,PinList,,,Util1Pins
    '"tl_GetPPMUMeasureCurrentRanges()",
    '"tl_GetPPMUMeasureCurrentRanges()",
    On Error GoTo errHandler
    ' ================================================================================
    '                        Declare variables
    ' ================================================================================
    Dim i_PPMUMeasureHigh_PLD As New PinListData
    Dim i_PPMUMeasureLow_PLD As New PinListData
    Dim i_MeasPins_str() As String
    Dim i_PinNum_lng As Long
    Dim Site As Variant
    
    ' ================================================================================
    '                        Initialize Settings
    ' ================================================================================
    'Set drive state on specified utility pins
    Call tl_SetTestState(ConnectAllPins, LoadLevels, LoadTiming, relayMode, InitWaitTime, DriveLoPins, DriveHiPins, DriveZPins, FloatPins, Util1Pins, Util0Pins)
    
    'DecomposePinList: Pinlist --> String Array
    Call TheExec.DataManager.DecomposePinList(powerPin, i_MeasPins_str(), i_PinNum_lng)
    
    'Disable PinlistDataSort
    Call tl_pinlistdatasort(SortPin)
    
    'Instance Value Check
    'Call Leakage_ParCheck(PowerPin, VForceHigh, VForceLow)
    
    ' ================================================================================
    '                        Start Testing
    ' ================================================================================
    If nonblank(PatName) Then Call TheHdw.Patterns(PatName).test(CheckPG, 0)
    
 
    'Disconnect Digitial pins in PE(PPMU and Digital Pins can connect to the pins in the same time)
    Call TheHdw.Digital.Pins(MeasPins).Disconnect
        
 
    'Setup Instrument & Measure
    With TheHdw.PPMU(MeasPins)
        .Gate = tlOff
        .ForceV 0
        .Connect
        .Gate = tlOn
        
        If nonblank(VForceHigh) Then
            .ForceV CDbl(VForceHigh), StrToDbl(Trim(MeasIRangeHi))
            TheHdw.Wait WaitTime
            
           
            i_PPMUMeasureHigh_PLD = .Read(tlPPMUReadMeasurements, SampleSize)
            
            
            ' Added by Zoe for checking if measure current range is reasonable, if it less than the measure current, the current be be clamp
'            If TheExec.TesterMode = testModeOnline Then
'                For Each site In TheExec.Sites.Selected
'                    If i_PPMUMeasureHigh_PLD.Analyze.Maximum.Abs(site) > StrToDbl(Trim(MeasIRangeHi)) Then
'
'                        If TheExec.RunMode = runModeProduction Then
'                            TheExec.AddOutput "Leakage Error 0005 : MeasIRangeHi is less than measure current, current is clamp!"
'                            'GoTo errHandler
'                        Else
'                            TheExec.AddOutput "Leakage Error 0005 : MeasIRangeHi is less than measure current, current is clamp!"
'                          'MsgBox "Leakage Error 0005 : MeasIRangeHi is less than measure current, current is clamp!"
'                          'Stop
'                        '''''''''''''''''''''''MASK BY JANE0511 FOR SITE0 FAIL WILL EFFECT SITE1 TEST'''''''''''''''''''''''''
'                        End If
'
'                    End If
'                Next site
'            End If
                        
        End If
        
        If nonblank(VForceLow) Then
            .ForceV CDbl(VForceLow), StrToDbl(Trim(MeasIRangeLo))
            TheHdw.Wait WaitTime
            
            
            i_PPMUMeasureLow_PLD = .Read(tlPPMUReadMeasurements, SampleSize)
            
            ' Added by Zoe for checking if measure current range is reasonable, if it less than the measure current, the current be be clamp
'            If TheExec.TesterMode = testModeOnline Then
'                For Each site In TheExec.Sites.Selected
'                    If i_PPMUMeasureLow_PLD.Analyze.Maximum.Abs(site) > StrToDbl(Trim(MeasIRangeLo)) Then
'
'                        If TheExec.RunMode = runModeProduction Then
'                            TheExec.AddOutput "Leakage Error 0006 : MeasIRangeLo is less than measure current, current is clamp!"
'                            GoTo errHandler
'                        Else
'                            MsgBox "Leakage Error 0006 : MeasIRangeLo is less than measure current, current is clamp!"
'                            Stop
'                        End If
'
'                    End If
'                Next site
'            End If
            
        End If
        .Disconnect
    End With

    'Connect Digitial pins in PE(PPMU and Digital Pins can connect to the pins in the same time)
    Call TheHdw.Digital.Pins(MeasPins).Connect

    ' ================================================================================
    '                        Datalog
    ' ================================================================================
    If nonblank(VForceLow) Then
        TheExec.Flow.TestLimit ResultVal:=i_PPMUMeasureLow_PLD, unit:=unitAmp, forceVal:=CDbl(VForceLow), Forceunit:=unitVolt, forceResults:=tlForceFlow
    End If
    If nonblank(VForceHigh) Then
        TheExec.Flow.TestLimit ResultVal:=i_PPMUMeasureHigh_PLD, unit:=unitAmp, forceVal:=CDbl(VForceHigh), Forceunit:=unitVolt, forceResults:=tlForceFlow
    End If
    'Reset
    'Set PinlistdataSort Default
    Call tl_pinlistdatasort(True)
        
Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function



Public Function Basic_03_DC_Leakage_JTAG_SEL(PatName As Pattern, _
                            powerPin As PinList, _
                            MeasPins As PinList, _
                            VForceHigh As String, _
                            VForceLow As String, _
                            Optional MeasIRangeHi As String = 30 * ma, _
                            Optional MeasIRangeLo As String = 30 * ma, _
                            Optional SampleSize As Long = 10, _
                            Optional WaitTime As Double = 0.0001, _
                            Optional CheckPG As PFType = pfAlways, _
                            Optional InitWaitTime As Double = 0.001, _
                            Optional ConnectAllPins As Boolean = True, _
                            Optional LoadLevels As Boolean = True, _
                            Optional LoadTiming As Boolean = True, _
                            Optional relayMode As tlRelayMode = 1, _
                            Optional DriveLoPins As PinList, _
                            Optional DriveHiPins As PinList, _
                            Optional DriveZPins As PinList, _
                            Optional FloatPins As PinList, _
                            Optional Util0Pins As PinList, _
                            Optional Util1Pins As PinList, _
                            Optional SortPin As Boolean = False) As Long
' EDITFORMAT1 1,,Pattern,,,PatName|
' EDITFORMAT1 2,,PinList,,Only One,PowerPin|
' EDITFORMAT1 3,,PFType,,,CheckPG|
' EDITFORMAT1 4,,Boolean,,,ApplyLeveltiming|
' EDITFORMAT1 5,,Boolean,,,ConnectAllPins|
' EDITFORMAT1 6,,Boolean,,,LoadLevels|
' EDITFORMAT1 7,,Boolean,,,LoadTiming|
' EDITFORMAT1 8,,tlRelayMode,,,RelayMode|
' EDITFORMAT1 9,,Double,,Apply lvl timing and Relay Setup Time,InitWaitTime|
' EDITFORMAT1 10,,PinList,PPMU,,MeasPins|
' EDITFORMAT1 11,,Long,,PPMU,SampleSize|
' EDITFORMAT1 12,,String,,Blank For PullUp Test,VForceHigh|
' EDITFORMAT1 13,,String,,Blank For PullDown Test,VForceLow|
' EDITFORMAT1 14,,String,,Must Be Set,MeasIRangeHi|
' EDITFORMAT1 15,,String,,Must Be Set,MeasIRangeLo|
' EDITFORMAT1 16,,Double,,PPMU Setup Time,WaitTime|
' EDITFORMAT1 17,,PinList,Pin States,,DriveLoPins|
' EDITFORMAT1 18,,PinList,,,DriveHiPins|
' EDITFORMAT1 19,,PinList,,,DriveZPins|
' EDITFORMAT1 20,,PinList,,,DisablePins|
' EDITFORMAT1 21,,PinList,,,FloatPins|
' EDITFORMAT1 22,,PinList,,,Util0Pins|
' EDITFORMAT1 23,,PinList,,,Util1Pins
    '"tl_GetPPMUMeasureCurrentRanges()",
    '"tl_GetPPMUMeasureCurrentRanges()",
    On Error GoTo errHandler
    ' ================================================================================
    '                        Declare variables
    ' ================================================================================
    Dim i_PPMUMeasureHigh_PLD As New PinListData
    Dim i_PPMUMeasureLow_PLD As New PinListData
    Dim i_MeasPins_str() As String
    Dim i_PinNum_lng As Long
    Dim Site As Variant
    
    TheHdw.Digital.Pins("DIG_1V8,DIG_1V2").Disconnect
    ' ================================================================================
    '                        Initialize Settings
    ' ================================================================================
    'Set drive state on specified utility pins
    Call tl_SetTestState(ConnectAllPins, LoadLevels, LoadTiming, relayMode, InitWaitTime, DriveLoPins, DriveHiPins, DriveZPins, FloatPins, Util1Pins, Util0Pins)
    
    'DecomposePinList: Pinlist --> String Array
    Call TheExec.DataManager.DecomposePinList(powerPin, i_MeasPins_str(), i_PinNum_lng)
    
    'Disable PinlistDataSort
    Call tl_pinlistdatasort(SortPin)
    
    'Instance Value Check
    'Call Leakage_ParCheck(PowerPin, VForceHigh, VForceLow)
    
    ' ================================================================================
    '                        Start Testing
    ' ================================================================================
    If nonblank(PatName) Then Call TheHdw.Patterns(PatName).test(CheckPG, 0)

 
    'Disconnect Digitial pins in PE(PPMU and Digital Pins can connect to the pins in the same time)
    Call TheHdw.Digital.Pins(MeasPins).Disconnect
        
 
    'Setup Instrument & Measure
    With TheHdw.PPMU(MeasPins)
        .Gate = tlOff
        .ForceV 0
        .Connect
        .Gate = tlOn
        
        If nonblank(VForceHigh) Then
            .ForceV CDbl(VForceHigh), StrToDbl(Trim(MeasIRangeHi))
            TheHdw.Wait WaitTime
            
           
            i_PPMUMeasureHigh_PLD = .Read(tlPPMUReadMeasurements, SampleSize)
            
            
            ' Added by Zoe for checking if measure current range is reasonable, if it less than the measure current, the current be be clamp
'            If TheExec.TesterMode = testModeOnline Then
'                For Each site In TheExec.Sites.Selected
'                    If i_PPMUMeasureHigh_PLD.Analyze.Maximum.Abs(site) > StrToDbl(Trim(MeasIRangeHi)) Then
'
'                        If TheExec.RunMode = runModeProduction Then
'                            TheExec.AddOutput "Leakage Error 0005 : MeasIRangeHi is less than measure current, current is clamp!"
'                            'GoTo errHandler
'                        Else
'                            TheExec.AddOutput "Leakage Error 0005 : MeasIRangeHi is less than measure current, current is clamp!"
'                          'MsgBox "Leakage Error 0005 : MeasIRangeHi is less than measure current, current is clamp!"
'                          'Stop
'                        '''''''''''''''''''''''MASK BY JANE0511 FOR SITE0 FAIL WILL EFFECT SITE1 TEST'''''''''''''''''''''''''
'                        End If
'
'                    End If
'                Next site
'            End If
                        
        End If
        
        If nonblank(VForceLow) Then
            .ForceV CDbl(VForceLow), StrToDbl(Trim(MeasIRangeLo))
            TheHdw.Wait WaitTime
            
            
            i_PPMUMeasureLow_PLD = .Read(tlPPMUReadMeasurements, SampleSize)
            
            ' Added by Zoe for checking if measure current range is reasonable, if it less than the measure current, the current be be clamp
'            If TheExec.TesterMode = testModeOnline Then
'                For Each site In TheExec.Sites.Selected
'                    If i_PPMUMeasureLow_PLD.Analyze.Maximum.Abs(site) > StrToDbl(Trim(MeasIRangeLo)) Then
'
'                        If TheExec.RunMode = runModeProduction Then
'                            TheExec.AddOutput "Leakage Error 0006 : MeasIRangeLo is less than measure current, current is clamp!"
'                            GoTo errHandler
'                        Else
'                            MsgBox "Leakage Error 0006 : MeasIRangeLo is less than measure current, current is clamp!"
'                            Stop
'                        End If
'
'                    End If
'                Next site
'            End If
            
        End If
        .Disconnect
    End With

    'Connect Digitial pins in PE(PPMU and Digital Pins can connect to the pins in the same time)
    Call TheHdw.Digital.Pins(MeasPins).Connect

    ' ================================================================================
    '                        Datalog
    ' ================================================================================
    If nonblank(VForceLow) Then
        TheExec.Flow.TestLimit ResultVal:=i_PPMUMeasureLow_PLD, unit:=unitAmp, forceVal:=CDbl(VForceLow), Forceunit:=unitVolt, forceResults:=tlForceFlow
    End If
    If nonblank(VForceHigh) Then
        TheExec.Flow.TestLimit ResultVal:=i_PPMUMeasureHigh_PLD, unit:=unitAmp, forceVal:=CDbl(VForceHigh), Forceunit:=unitVolt, forceResults:=tlForceFlow
    End If
    'Reset
    'Set PinlistdataSort Default
    Call tl_pinlistdatasort(True)
        
Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function Leakage_ParCheck(powerPin As PinList, VForceHigh As String, VForceLow As String)

    On Error GoTo errHandler
    
    Dim i_PinType_str As String
    Dim i_DCVIValue_PLD As New PinListData
    
    i_PinType_str = TheExec.DataManager.ChannelType(powerPin)
    
    If MID(i_PinType_str, 1, 4) = "DCVI" Then
        i_DCVIValue_PLD = TheHdw.DCVI.Pins(powerPin).Meter.Read(tlStrobe, 10, 1000, tlDCVIMeterReadingFormatAverage)
    Else
        i_DCVIValue_PLD = TheHdw.DCVS.Pins(powerPin).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
    End If
    
    ' ================================================================================
    '                        At Least One Of VForceHigh and VForceLow Should Be Set
    ' ================================================================================
    If (Not nonblank(VForceHigh)) And (Not nonblank(VForceLow)) Then
        If TheExec.RunMode = runModeProduction Then
            TheExec.AddOutput "Leakage Error 0001 :  At Least One Of VForceHigh and VForceLow Should Be Set"
            GoTo errHandler
        Else
            MsgBox "Leakage Error 0001 :  At Least One Of VForceHigh and VForceLow Should Be Set"
            Stop
        End If
    End If
    
    ' ================================================================================
    '                        PowerPin Number Check
    ' ================================================================================
    
    If i_DCVIValue_PLD.Pins.count <> 1 Then
        If TheExec.RunMode = runModeProduction Then
            TheExec.AddOutput "Leakage Error 0002 : PowerPin Number Should Not Exceed 1 Or Blank"
            GoTo errHandler
        Else
            MsgBox "Leakage Error 0002 : PowerPin Number Should Not Exceed 1 Or Blank"
            Stop
        End If
    End If
    ' ================================================================================
    '                        Detect Pin Type And Judge
    ' ================================================================================
    '   1.Detect Pin Type(DCVI or DCVS)
    '   2.Judge The Force Value Equal to The Power Pin HiValue

'    If NonBlank(VForceHigh) And (TheExec.TesterMode = testModeOnline) Then
'        If Mid(i_PinType_str, 1, 4) = "DCVI" Then
'            If (Abs(VForceHigh - TheHdw.DCVI.Pins(PowerPin).Voltage) / VForceHigh) > 0.01 Then
'                If TheExec.RunMode = runModeProduction Then
'                    TheExec.AddOutput "Leakage Error 0003 : Force Value Not Equal to The Power Pin HiValue(DCVI)"
'                    GoTo errHandler
'                Else
'                    MsgBox "Leakage Error 0003 : Force Value Not Equal to The Power Pin HiValue(DCVI)"
'                    Stop
'                End If
'            End If
'        ElseIf Mid(i_PinType_str, 1, 4) = "DCVS" Then
'            If (Abs(VForceHigh - TheHdw.DCVS.Pins(PowerPin).Voltage.Value) / VForceHigh) > 0.01 Then
'                    If TheExec.RunMode = runModeProduction Then
'                    TheExec.AddOutput "Leakage Error 0004 : Force Value Not Equal to The Power Pin HiValue(DCVS)"
'                    GoTo errHandler
'                Else
'                    MsgBox "Leakage Error 0004 : Force Value Not Equal to The Power Pin HiValue(DCVS)"
'                    Stop
'                End If
'            End If
'        End If
'    End If
Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

' ================================================================================
'                        Basic_03_DCM_RES
' ================================================================================
Public Function Basic_03_DCM_RES(DCM_OUT_Pin As PinList, _
                                CORE_POWER_Pin As PinList, _
                                Optional WaitTime As Double = 0.0001, _
                                Optional InitWaitTime As Double = 0.001, _
                                Optional ConnectAllPins As Boolean = True, _
                                Optional LoadLevels As Boolean = True, _
                                Optional LoadTiming As Boolean = True, _
                                Optional relayMode As tlRelayMode = 1, _
                                Optional DriveLoPins As PinList, _
                                Optional DriveHiPins As PinList, _
                                Optional DriveZPins As PinList, _
                                Optional FloatPins As PinList, _
                                Optional Util0Pins As PinList, _
                                Optional Util1Pins As PinList)
' EDITFORMAT1 1,,PinList,,DCM OUT Pin,DCM_OUT_Pin|
' EDITFORMAT1 2,,PinList,,,CORE_POWER_Pin|
' EDITFORMAT1 3,,Double,,PPMU settling time,WaitTime|
' EDITFORMAT1 4,,Double,,wait time after apply lvl/timing and relay setup,InitWaitTime|
' EDITFORMAT1 5,,Boolean,,,ConnectAllPins|
' EDITFORMAT1 6,,Boolean,,,LoadLevels|
' EDITFORMAT1 7,,Boolean,,,LoadTiming|
' EDITFORMAT1 8,,tlRelayMode,,,RelayMode|
' EDITFORMAT1 9,,PinList,Pin States,,DriveLoPins|
' EDITFORMAT1 10,,PinList,,,DriveHiPins|
' EDITFORMAT1 11,,PinList,,,DriveZPins|
' EDITFORMAT1 12,,PinList,,,FloatPins|
' EDITFORMAT1 13,,PinList,,,Util0Pins|
' EDITFORMAT1 14,,PinList,,,Util1Pins
    On Error GoTo errHandler
    
    Dim Core_Power_Voltage As New SiteDouble
    Dim Force_Low_Voltage As New SiteDouble
    Dim ForceH_Current As New PinListData
    Dim ForceL_Current As New PinListData
    Dim Upper_Res As New PinListData
    Dim Lower_Res As New PinListData

    Call tl_SetTestState(ConnectAllPins, LoadLevels, LoadTiming, relayMode, InitWaitTime, DriveLoPins, DriveHiPins, DriveZPins, FloatPins, Util1Pins, Util0Pins)
    
    Core_Power_Voltage = Power_voltage(CORE_POWER_Pin)
    Force_Low_Voltage = 0
    
    Call FVMI_PREPARE(DCM_OUT_Pin)
    
    ' 2uA and 20uA measure current range for normal condition of DCM RES test
    ForceH_Current = FVMI_READ(DCM_OUT_Pin, Core_Power_Voltage, 10 * US, 2 * uA)
    ForceL_Current = FVMI_READ(DCM_OUT_Pin, Force_Low_Voltage, 10 * US, 200 * uA)
     
    Call FVMI_END(DCM_OUT_Pin)
    
    Upper_Res = ForceL_Current.Math.Invert.Multiply(Force_Low_Voltage - Core_Power_Voltage).Abs
    Lower_Res = ForceH_Current.Math.Invert.Multiply(Core_Power_Voltage).Abs
    
    TheExec.Flow.TestLimit Lower_Res, TName:="DCM_RES_LOWER", forceResults:=tlForceFlow, unit:=unitCustom, customUnit:="Ohm", forceVal:=Core_Power_Voltage, Forceunit:=unitVolt
    TheExec.Flow.TestLimit Upper_Res, TName:="DCM_RES_UPPER", forceResults:=tlForceFlow, unit:=unitCustom, customUnit:="Ohm", forceVal:=Force_Low_Voltage, Forceunit:=unitVolt
    
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function FVMI_READ(MeasPins As PinList, ForceVoltage As SiteDouble, WaitTime As Double, MeasCurRange As Double) As PinListData
    
    On Error GoTo errHandler
    
    Set FVMI_READ = New PinListData
    Dim ActualCurrentRange As New SiteDouble
    Dim MaxCurrent As New SiteDouble
    Dim Site As Variant
    Dim InstanceName As String
    
    With TheHdw.PPMU(MeasPins)
    
        .ForceV ForceVoltage, MeasCurRange
        TheHdw.Wait WaitTime
        FVMI_READ = .Read(tlPPMUReadMeasurements, 10)
        ActualCurrentRange = .MeasureCurrentRange.Value
        
    End With

    MaxCurrent = FVMI_READ.Analyze.Maximum.Abs
    
    If TheExec.TesterMode = testModeOnline Then
    
        For Each Site In TheExec.Sites.Active
            If MaxCurrent(Site) > ActualCurrentRange(Site) Then
                
                InstanceName = TheExec.DataManager.InstanceName
                
                If TheExec.RunMode = runModeProduction Then
                    TheExec.AddOutput "DCM RES Error 0001:" + InstanceName + ": Measure Current(" + CStr(MaxCurrent(Site)) + ") Larger than Current Range(" + CStr(ActualCurrentRange(Site)) + ") for Site(" + CStr(Site) + ")!", vbRed
                    GoTo errHandler
                Else
                    MsgBox "DCM RES Error 0001:" + InstanceName + ": Measure Current(" + CStr(MaxCurrent(Site)) + ") Larger than Current Range(" + CStr(ActualCurrentRange(Site)) + ") for Site(" + CStr(Site) + ")!"
                    Stop
                End If
                
            End If
        Next
        
    End If
    
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function FVMI_PREPARE(MeasPins As PinList)
    
    Call TheHdw.Digital.Pins(MeasPins).Disconnect

    With TheHdw.PPMU(MeasPins)
        .Gate = tlOff
        .ForceV 0
        .Connect
        .Gate = tlOn
    End With
    
End Function

Private Function FVMI_END(MeasPins As PinList)

    Call TheHdw.PPMU(MeasPins).Disconnect
    Call TheHdw.Digital.Pins(MeasPins).Connect
    
End Function

Private Function Power_voltage(powerPin As PinList) As SiteDouble
    
    Dim PowerPinChannelType As String
    Set Power_voltage = New SiteDouble
    
    PowerPinChannelType = TheExec.DataManager.ChannelType(powerPin)
    
    If MID(PowerPinChannelType, 1, 4) = "DCVI" Then
        Power_voltage = TheHdw.DCVI.Pins(powerPin).voltage
    Else
        Power_voltage = TheHdw.DCVS.Pins(powerPin).voltage.Main.Value
    End If

End Function


' ================================================================================
'                        Basic_03_DC_IOHL
' ================================================================================

Public Function Basic_03_DC_IOHL(PatName As Pattern, _
                                MeasPins As PinList, _
                                VForceValue As Double, _
                                Optional CheckPG As PFType = pfAlways, _
                                Optional PinsNumPerGrp As Long = 1, _
                                Optional relayMode As tlRelayMode = 1, _
                                Optional SampleSize As Long = 10, _
                                Optional InitWaitTime As Double = 0.0001, _
                                Optional MeasWaitTime As Double = 0.002, _
                                Optional ConnectAllPins As Boolean = True, _
                                Optional LoadLevels As Boolean = True, _
                                Optional LoadTiming As Boolean = True, _
                                Optional DriveLoPins As PinList, _
                                Optional DriveHiPins As PinList, _
                                Optional DriveZPins As PinList, _
                                Optional FloatPins As PinList, _
                                Optional Util1Pins As PinList, _
                                Optional Util0Pins As PinList) As Long
' EDITFORMAT1 1,,Pattern,,Can not be blank,PatName|
' EDITFORMAT1 8,,PinList,PPMU,Pin group or comma separated pin list,MeasPins|
' EDITFORMAT1 9,,Double,,PPMU Force value(VOH or VOL),VForceValue|
' EDITFORMAT1 2,,PFType,,,checkPG|
' EDITFORMAT1 10,,Long,,Pin Numbers Per Group(test all pins when equals to 0 or large than total pin number),PinsNumPerGrp|
' EDITFORMAT1 6,,tlRelayMode,,,RelayMode|
' EDITFORMAT1 11,,Long,,,SampleSize|
' EDITFORMAT1 7,,Double,,apply lvl timing and relay set up time,InitWaitTime|
' EDITFORMAT1 12,,Double,, PPMU Setup Time,MeasWaitTime|
' EDITFORMAT1 3,,Boolean,,,ConnectAllPins|
' EDITFORMAT1 4,,Boolean,,,LoadLevels|
' EDITFORMAT1 5,,Boolean,,,LoadTiming|
' EDITFORMAT1 13,,PinList,Pin States,,DriveLoPins|
' EDITFORMAT1 14,,PinList,,,DriveHiPins|
' EDITFORMAT1 15,,PinList,,,DriveZPins|
' EDITFORMAT1 16,,PinList,,,FloatPins|
' EDITFORMAT1 18,,PinList,,,Util1Pins|
' EDITFORMAT1 17,,PinList,,,Util0Pins

    On Error GoTo errHandler
    
    ' ================================================================================
    '                        Declare variables
    ' ================================================================================
    Dim PPMUMeasure_array() As New PinListData
    Dim Site As Variant
    Dim GroupNum_lng As Long
    Dim Pinname_list() As String
    Dim i_PinIndex_lng As Long
    Dim Measrange() As Double
    
    '===============================================================================
    '               Initialize Setting
    '===============================================================================


    Call tl_SetTestState(ConnectAllPins, LoadLevels, LoadTiming, relayMode, InitWaitTime, DriveLoPins, DriveHiPins, DriveZPins, FloatPins, Util1Pins, Util0Pins)
    
    '===============================================================================
    '               Get Current Range From Use-Limits
    '===============================================================================

    Call Pri_GetMeterRangeFromLimit(Measrange(), 1)

    '===============================================================================
    '               Run Pattern
    '===============================================================================
    Call tl_pinlistdatasort(False)

    If nonblank(PatName) Then
        TheHdw.Patterns(PatName).test CheckPG, 0
    Else
        If TheExec.RunMode = runModeProduction Then
            TheExec.AddOutput "IOHL Error 0002: Empty Pattern Input! "
            GoTo errHandler
        Else
            MsgBox "IOHL Error 0002: Empty Pattern Input! "
            Stop
        End If
    End If
    
    '===============================================================================
    '               SmartParallel PPMU test
    '===============================================================================
    
    Call Pri_Cut_Pinlist(MeasPins, PinsNumPerGrp, Pinname_list(), GroupNum_lng)

    ReDim PPMUMeasure_array(GroupNum_lng - 1) As New PinListData
           
    For i_PinIndex_lng = 0 To GroupNum_lng - 1
    
        PPMUMeasure_array(i_PinIndex_lng) = Pri_PPMU_fvmi(Pinname_list(i_PinIndex_lng), VForceValue, True, MeasWaitTime, Measrange(0), SampleSize)
        
        With TheExec.Flow
            .TestLimitIndex = 0
            .TestLimit ResultVal:=PPMUMeasure_array(i_PinIndex_lng), unit:=unitAmp, _
            forceVal:=VForceValue, _
            Forceunit:=unitVolt, _
            forceResults:=tlForceFlow
        End With
        
    Next i_PinIndex_lng
    
    Call tl_pinlistdatasort(True)

    Exit Function
    
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function Pri_Cut_Pinlist(MeasPins_all As PinList, _
                              PinsNumPerGrp As Long, _
                              ResultString() As String, _
                              Optional GroupNum As Long, _
                              Optional MeasPinsNum As Long) As Long   ''the last THREE parameter would RETURN back
                              

    On Error GoTo errHandler

    Dim i_index_lng As Long
    Dim j_index_lng As Long
    Dim i_PerPin_str() As String
    Dim i_PinNum_lng As Long
    Dim i_PinIndex_lng As Long
    
    '===============================================================================
    '               Decompose the PinList Given
    '===============================================================================
    
    Call TheExec.DataManager.DecomposePinList(MeasPins_all, i_PerPin_str(), i_PinNum_lng)
    
    If PinsNumPerGrp > i_PinNum_lng Or PinsNumPerGrp = 0 Then
        ReDim ResultString(0)
        ResultString(0) = MeasPins_all.Value
        MeasPinsNum = i_PinNum_lng
        GroupNum = 1
        Exit Function
    Else
        If PinsNumPerGrp < 0 Then
        
            If TheExec.RunMode = runModeProduction Then
                TheExec.AddOutput "IOHL Error 0001: PinsNumPerGrp should large than 0 "
                GoTo errHandler
            Else
                MsgBox "IOHL Error 0001: PinsNumPerGrp should large than 0 "
                Stop
            End If
        End If
                
    End If
    
    GroupNum = Ceiling(i_PinNum_lng / PinsNumPerGrp)
    
    '===============================================================================
    '               Cut the PinList into pieces
    '===============================================================================
    ReDim ResultString(GroupNum - 1)
    For i_index_lng = 0 To GroupNum - 1
    
        ResultString(i_index_lng) = ""
        For j_index_lng = PinsNumPerGrp * i_index_lng To IIf(PinsNumPerGrp * (i_index_lng + 1) - 2 > i_PinNum_lng - 2, i_PinNum_lng - 2, PinsNumPerGrp * (i_index_lng + 1) - 2)
            ResultString(i_index_lng) = ResultString(i_index_lng) & i_PerPin_str(j_index_lng) & ","
        Next j_index_lng
        ResultString(i_index_lng) = ResultString(i_index_lng) & i_PerPin_str(j_index_lng)
        
    Next i_index_lng
    
    MeasPinsNum = i_PinNum_lng

    Exit Function
    
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function Pri_GetMeterRangeFromLimit(out_LimitVal() As Double, _
                                            in_PinNum As Long) As Long
On Error GoTo errHandler

    Dim FlowLimitsInfo As IFlowLimitsInfo
    Dim i_HighLimit_str() As String
    Dim i_LowLimit_str() As String
    Dim i As Long
    
    ReDim out_LimitVal(in_PinNum - 1) As Double
    
    Call TheExec.Flow.GetTestLimits(FlowLimitsInfo)
    Call FlowLimitsInfo.GetHighLimits(i_HighLimit_str())
    Call FlowLimitsInfo.GetLowLimits(i_LowLimit_str())
    
    '===============================================================================
    '               Give Range From Use-Limit
    '===============================================================================
    If UBound(i_HighLimit_str) <> in_PinNum - 1 Or UBound(i_LowLimit_str) <> in_PinNum - 1 Then
        
        If TheExec.RunMode = runModeProduction Then
            TheExec.AddOutput "Use-Limit Error 0001: Number of limits and pins does not match, please check your limits in flow table!"
            GoTo errHandler
        Else
            MsgBox "Use-Limit Error 0001: Number of limits and pins does not match, please check your limits in flow table!"
            Stop
        End If
        
    Else
    
        For i = 0 To in_PinNum - 1
            If Abs(i_HighLimit_str(i)) > Abs(i_LowLimit_str(i)) Then
                out_LimitVal(i) = StrToDbl(Abs(i_HighLimit_str(i))) * 1.01
            Else
                out_LimitVal(i) = StrToDbl(Abs(i_LowLimit_str(i))) * 1.01
            End If
        Next i
        
    End If
    
    Exit Function
    
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function Pri_PPMU_fvmi(MeasPins As String, _
                          ForceV_value As Double, _
                          NeedConnectPE As Boolean, _
                          Optional WaitMeasTime As Double = 0.001, _
                          Optional currentRange As Double, _
                          Optional SampleSize As Long = 10) As PinListData
                          
On Error GoTo errHandler

    TheHdw.Digital.Pins(MeasPins).Disconnect
    With TheHdw.PPMU.Pins(MeasPins)
        .ForceI 0
        .Connect
        .ForceV ForceV_value, currentRange
        .Gate = tlOn
        TheHdw.Wait WaitMeasTime
        Set Pri_PPMU_fvmi = .Read(tlPPMUReadMeasurements, SampleSize)
        .ForceI 0
        .Gate = tlOff
        .Disconnect
    End With

    If NeedConnectPE = True Then
        TheHdw.Digital.Pins(MeasPins).Connect
    End If
    
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function


