Attribute VB_Name = "HiBench_IP"
Option Explicit



' This module should be used for VBT Tests.  All functions in this module

' will be available to be used from the Test Instance sheet.

' Additional modules may be added as needed (all starting with "VBT_").

'

' The required signature for a VBT Test is:

'

' Public Function FuncName(<arglist>) As Long

'   where <arglist> is any list of arguments supported by VBT Tests.

'

' See online help for supported argument types in VBT Tests.

'

'

' It is highly suggested to use error handlers in VBT Tests.  A sample

' VBT Test with a suggeseted error handler is shown below:

'

' Function FuncName() As Long

'     On Error GoTo errHandler

'

'     Exit Function

' errHandler:

'     If AbortTest Then Exit Function Else Resume Next

' End Function

Public hpmmap As New Collection
Public hpmmapx As New Collection
Public Const match_time = 40 'ms

Public Type STR_HIBENCH_RESULT

    HiBenchResultName As String
    CapPinNameList As String
    ExpectValue As Long
    ReadValue As New SiteLong
    DigCap_Pos As String

End Type


Public Type STR_HIBENCHTESTINFO

    ItemName As String
    ItemCnt As Long
    ExecuteTime_ms As Double
    VolType As String
    MaxStructLine As Long
    HiBenchResult() As STR_HIBENCH_RESULT
    Item_Pass As New SiteBoolean
    per_site_result  As New SiteBoolean
    per_site_result_high As New SiteBoolean
    MatchLoopTime As New SiteLong
    all_site_pass As Boolean
    all_site_high As Boolean
    comparedValue As Long
    PatName As String

End Type


Public STR_HIBENCH_ITEM_LIST As String
Public TheHiBench() As STR_HIBENCHTESTINFO




Public Function HiBenchInfo_Initilize() As Long  ' to be called in "onProgrammedValidation"
    
    
    On Error GoTo errHandler
    
    HiBenchInfo_Initilize = 0
    
    Dim str_HiBenchItem As String
    Dim arr_HiBenchItem() As String
    Dim str_ExecuteTime_ms As String
    Dim arr_ExecuteTime_ms() As String
    Dim str_ExpectValue As String
    Dim arr_ExpectValue() As String
    Dim str_MaxStructLine As String
    Dim arr_MaxStructLine() As String
    Dim arr_subExpectValue() As String
    Dim str_HiBenchResultName As String
    Dim arr_HiBenchResultName() As String
    Dim arr_subHibenchResultName() As String

    str_HiBenchItem = "LV_DDR_0V6,LV_DDR_0V65,LV_DDR_0V7,LV_DDR_0V75,DDR,IP,CPU_L,L3,CPU_M,CPU_B,LV_IP"
        arr_HiBenchItem = Split(str_HiBenchItem, ",")
        
    str_ExecuteTime_ms = "2400,3000,3000,3500,3000,6000,4000,4000,4000,4000,8000"
        arr_ExecuteTime_ms = Split(str_ExecuteTime_ms, ",")

    str_ExpectValue = "3,5,7_3,31,11,63_3,15,9,15,15_1,63_1" '??
        arr_ExpectValue = Split(str_ExpectValue, ",")

    str_MaxStructLine = "1,1,2,1,1,2,1,1,1,2,2"
        arr_MaxStructLine = Split(str_MaxStructLine, ",")

    str_HiBenchResultName = "LV_DDR_0V6_Result,LV_DDR_0V65_Result,LV_DDR_0V7_Result/IMG_Version_xloader,LV_DDR_0V75_Result,DDR_Result,Function_Reuslt/IMG_Version_fastboot,CPU_L_Result,L3_Result,CPU_M_Result,CPU_B_Result/DDR_down_Freq_judge,LV_FUNCTION_Result/EFUSE_Result"
        arr_HiBenchResultName = Split(str_HiBenchResultName, ",")

    ReDim TheHiBench(UBound(arr_HiBenchItem))

    ' to initialize test info

    Dim i As Long
    Dim j As Long

    For i = 0 To UBound(arr_HiBenchItem)
        TheHiBench(i).ItemName = arr_HiBenchItem(i)
        TheHiBench(i).ItemCnt = UBound(arr_HiBenchItem) + 1
        TheHiBench(i).ExecuteTime_ms = CLng(arr_ExecuteTime_ms(i))
        TheHiBench(i).VolType = "NORMAL"  ' "NORMAL|AVS"
        TheHiBench(i).MaxStructLine = arr_MaxStructLine(i)
        ReDim TheHiBench(i).HiBenchResult(arr_MaxStructLine(i))
        arr_subExpectValue = Split(arr_ExpectValue(i), "_")
        arr_subHibenchResultName = Split(arr_HiBenchResultName(i), "/")

        For j = 0 To TheHiBench(i).MaxStructLine - 1
            TheHiBench(i).HiBenchResult(j).ExpectValue = arr_subExpectValue(j)
            TheHiBench(i).HiBenchResult(j).HiBenchResultName = arr_subHibenchResultName(j)
            TheHiBench(i).HiBenchResult(j).ReadValue = 0
            TheHiBench(i).HiBenchResult(j).DigCap_Pos = 1
        Next j

        TheHiBench(i).Item_Pass = False
        TheHiBench(i).per_site_result = False
        TheHiBench(i).per_site_result_high = False
        TheHiBench(i).MatchLoopTime = 0
        TheHiBench(i).PatName = "HiBench_IP.patx"

        Select Case i

            Case 0
                TheHiBench(i).HiBenchResult(0).CapPinNameList = "RF0_RESET_N,GPIO_BBA_2,GPIO_BBA_0,GPIO_052,GPIO_051,GPIO_050,GPIO_049"
            Case 1
                TheHiBench(i).HiBenchResult(0).CapPinNameList = "RF0_RESET_N,GPIO_BBA_2,GPIO_BBA_0,GPIO_052,GPIO_051,GPIO_050,GPIO_049"
            Case 2
                TheHiBench(i).HiBenchResult(0).CapPinNameList = "RF0_RESET_N,GPIO_BBA_2,GPIO_BBA_0,GPIO_052,GPIO_051,GPIO_050,GPIO_049"
                TheHiBench(i).HiBenchResult(1).CapPinNameList = "GPIO_048,GPIO_042,GPIO_041,GPIO_040,GPIO_039"
            Case 3
                TheHiBench(i).HiBenchResult(0).CapPinNameList = "RF0_RESET_N,GPIO_BBA_2,GPIO_BBA_0,GPIO_052,GPIO_051,GPIO_050,GPIO_049"
            Case 4
                TheHiBench(i).HiBenchResult(0).CapPinNameList = "RF0_RESET_N,GPIO_BBA_2,GPIO_BBA_0,GPIO_052,GPIO_051,GPIO_050,GPIO_049"
            Case 5
                TheHiBench(i).HiBenchResult(0).CapPinNameList = "RF0_RESET_N,GPIO_BBA_2,GPIO_BBA_0,GPIO_052,GPIO_051,GPIO_050,GPIO_049"
            Case 6
                TheHiBench(i).HiBenchResult(0).CapPinNameList = "GPIO_068,LTE_GPS_TIM_IND,GPIO_BBA_7,GPIO_BBA_6,GPIO_BBA_5,GPIO_BBA_4,GPIO_BBA_3"
                TheHiBench(i).HiBenchResult(1).CapPinNameList = "GPIO_048,GPIO_042,GPIO_041,GPIO_040,GPIO_039"
            Case 7
                TheHiBench(i).HiBenchResult(0).CapPinNameList = "GPIO_069,GPIO_070,GPIO_071,GPIO_072,GPIO_073"
            Case 8
                TheHiBench(i).HiBenchResult(0).CapPinNameList = "GPIO_069,GPIO_070,GPIO_071,GPIO_072,GPIO_073"
            Case 9
                TheHiBench(i).HiBenchResult(0).CapPinNameList = "UART4_CTS_N,UART4_RTS_N,UART4_RXD,UART4_TXD,I2C5_SCL"
            Case 10
                TheHiBench(i).HiBenchResult(0).CapPinNameList = "SIF0_DOUT1,SIF0_DIN1,SPI4_DO,SPI1_CLK,SPI1_D"
                TheHiBench(i).HiBenchResult(1).CapPinNameList = "GPIO_068,LTE_GPS_TIM_IND,GPIO_BBA_7,GPIO_BBA_6,GPIO_BBA_5,GPIO_BBA_4,GPIO_BBA_3"
            Case 11
                TheHiBench(i).HiBenchResult(0).CapPinNameList = "GPIO_068,LTE_GPS_TIM_IND,GPIO_BBA_7,GPIO_BBA_6,GPIO_BBA_5,GPIO_BBA_4,GPIO_BBA_3"
                TheHiBench(i).HiBenchResult(1).CapPinNameList = "SPI1_DO"
            Case Else
        

        End Select

    Next i
    
    HiBenchInfo_Initilize = 1
    
    Exit Function
    
errHandler:
        TheExec.Datalog.WriteComment "error occurred in HiBenchInfo_Initilize "
        HiBenchInfo_Initilize = -999

End Function





Public Function HiBench_DigCap_ResultCheck() As Long


    Dim i As Long, j As Long
    Dim i_DebugMode As Boolean
    Dim dw_cap As New DSPWave
    Dim Site As Variant
    Dim i_lpcnt As Long
    Dim tmpStr As String
    
    Dim lev_spec As String
    Dim SPEC_NAME As String
    Dim label As String
    Dim lev_voltage As New SiteDouble

    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered

    On Error GoTo errHandler
    
    If HiBenchInfo_Initilize = 0 Then ' to initilize settings
        Call HiBenchInfo_Initilize
    End If

    If i_DebugMode = True Then

        Dim CurrDCCategory As String
        Dim CurrDCSelector As String
        Dim CurrLevel As String
        Dim CurrACCategory As String
        Dim CurrACSelector As String
        Dim CurrTimeSet As String
        Dim CurrEdgeSet As String
        Dim CurrOverlay As String

        Call TheExec.DataManager.GetInstanceContext(CurrDCCategory, CurrDCSelector, CurrACCategory, CurrACSelector, CurrTimeSet, CurrEdgeSet, CurrLevel, CurrOverlay)
        TheExec.Datalog.WriteComment "CurrDCCategory:  " + vbTab + CurrDCCategory
        TheExec.Datalog.WriteComment "CurrDCSelector:  " + vbTab + CurrDCSelector
        TheExec.Datalog.WriteComment "CurrACCategory:  " + vbTab + CurrACCategory
        TheExec.Datalog.WriteComment "CurrACSelector:  " + vbTab + CurrACSelector
        TheExec.Datalog.WriteComment "CurrTimeSet:  " + vbTab + CurrTimeSet
        TheExec.Datalog.WriteComment "CurrLevel:  " + vbTab + CurrLevel
        TheExec.Datalog.WriteComment "CurrLevel:  " + vbTab + CurrLevel

    End If
    

    ' to get initial value of dc/ac spec

    Dim Vspec_VDD08_CPU_BM As Double: Vspec_VDD08_CPU_BM = TheExec.Specs.DC.Item("VDD08_CPU_BM").CurrentValue
    Dim Vspec_VDD07_GPU As Double: Vspec_VDD08_CPU_BM = TheExec.Specs.DC.Item("VDD07_GPU").CurrentValue
    Dim Vspec_VDDC08_MEM_CPUM As Double: Vspec_VDD08_CPU_BM = TheExec.Specs.DC.Item("VDDC08_MEM_CPUM").CurrentValue
    Dim Vspec_VDD08_CPU_L As Double: Vspec_VDD08_CPU_BM = TheExec.Specs.DC.Item("VDD08_CPU_L").CurrentValue
    Dim Vspec_VDD075_PERI As Double: Vspec_VDD08_CPU_BM = TheExec.Specs.DC.Item("VDD075_PERI").CurrentValue
    Dim Vspec_VDD08_DDR As Double: Vspec_VDD08_CPU_BM = TheExec.Specs.DC.Item("VDD08_DDR").CurrentValue
    Dim Vspec_VDD08_NPU As Double: Vspec_VDD08_CPU_BM = TheExec.Specs.DC.Item("VDD08_NPU").CurrentValue

    
    ' to start 32K,19M2...
    
    Call TheHdw.Digital.ApplyLevelsTiming(True, True, True)
    
    With TheHdw.Digital.Pins("CLK_SLEEP").FreeRunningClock '32K
        .Enabled = True
        .Frequency = 32000#
        .Start
    End With
    
    With TheHdw.Digital.Pins("CLK_SYSTEM").FreeRunningClock '19M2
        .Enabled = True
        .Frequency = 19200000#
        .Start
    End With
    
    With TheHdw.Digital.Pins("UFS_REF_CLK").FreeRunningClock '19M2 UFS
        .Enabled = True
        .Frequency = 19200000#
        .Start
    End With
    
    ' for offline
    
    Dim tmpVal As New SiteDouble
    If TheExec.TesterMode = testModeOffline Then
        hpmmap.Add tmpVal, "VDD08_DDR_0V7"
    End If
    
    ' end
    
    '....................................................................................................................................
    ' setup patterns waiting for handan inputs
    
            '        PORT_HiBench_IP_DDR_0P6_INT
            '
            '        PORT_HiBench_IP_DDR_0P65_INT
            '
            '        PORT_HiBench_IP_DDR_0P7_INT
            '
            '        PORT_HiBench_IP_0P8_INT
            '
            '        PORT_HiBench_IP_DDR_0P75_INT
            '
            '        XXX
            '
            '        PORT_HiBench_IP_UFS_CPU_L_INT
            '
            '        PORT_HiBench_IP_UFS_L3_INT
            '
            '        PORT_HiBench_IP_UFS_CPU_M_INT
            '
            '        PORT_HiBench_IP_UFS_CPU_B_INT
            '
            '        PORT_HiBench_IP_UFS_LV_INT
            
            
            '        RDI_BEGIN( (rdiRunMode==0)?(TA::PROD):(TA::BURST) );
            '                rdi.port(PortName_19M2).func().label("PORT_HiBench_IP_DDR_0P65_INT").execute();
            '                rdi.port(PortName_32K).func().label("PORT_32K").execute();
            '                rdi.port(PortName_19M2_CLK).func().label("PORT_19P2M").execute();
            '                rdi.port(PortName_19M2_UFS).func().label("PORT_19P2M_UFS").execute();
            '        RDI_END();
    '....................................................................................................................................
    
    '*********HiBench_LV_DDR_0V6 start**********************************************************************************************************************
    
    i = 0
    TheHiBench(i).Item_Pass = True
    TheHiBench(i).per_site_result = False
    TheHiBench(i).per_site_result_high = False
    TheHiBench(i).comparedValue = 32

    If TheHiBench(i).VolType = "AVS" Then
        lev_spec = "VDD08_DDR_0V7"
        SPEC_NAME = "DDR"
        label = "0P6V"
        For Each Site In TheExec.Sites
            lev_voltage(Site) = GET_VMIN_FCK_BY_AVS_FORMULA(CDbl(hpmmap.Item(lev_spec)(Site)), CDbl(hpmmapx(lev_spec)(Site)), SPEC_NAME, label, CLng(Site))
        Next Site
        Call TheExec.Overlays.ApplyUniformSpecToHW(lev_spec, lev_voltage, False, False)
    ElseIf TheHiBench(i).VolType = "NORMAL" Then
    
    
    End If
    
    Call TheExec.Overlays.ApplyUniformSpecToHW("VDD075_PERI", "0.6", False, False)
    Call TheExec.Overlays.ApplyUniformSpecToHW("VDD08_DDR", "0.6", False, False)
    TheHdw.Wait 3 * mS

    TheHdw.Patterns("PORT_HiBench_IP_DDR_0P6_INT").Start   ' to setup some pins status to control device
    TheHdw.Wait 2000 * mS
    
    Call ResultCheck(i, TheHiBench(i).comparedValue)
    TheHdw.Wait 50 * mS
    
    
    '#############HiBench_LV_DDR_0V6 end ###################
    
    
    '*********HiBench_LV_DDR_0V65 start**********************************************************************************************************************
    i = 1
    TheHiBench(i).Item_Pass = True
    TheHiBench(i).per_site_result = False
    TheHiBench(i).per_site_result_high = False
    TheHiBench(i).comparedValue = 32
    
    If TheHiBench(i).VolType = "AVS" Then
        lev_spec = "VDD08_DDR_0V7"
        SPEC_NAME = "DDR"
        label = "0P65V"
        For Each Site In TheExec.Sites
            lev_voltage(Site) = GET_VMIN_FCK_BY_AVS_FORMULA(CDbl(hpmmap.Item(lev_spec)(Site)), CDbl(hpmmapx(lev_spec)(Site)), SPEC_NAME, label, CLng(Site))
        Next Site
        Call TheExec.Overlays.ApplyUniformSpecToHW(lev_spec, lev_voltage, False, False)
    ElseIf TheHiBench(i).VolType = "NORMAL" Then

        
    End If
    
    Call TheExec.Overlays.ApplyUniformSpecToHW("VDD075_PERI", "0.65", False, False)
    Call TheExec.Overlays.ApplyUniformSpecToHW("VDD08_DDR", "0.65", False, False)
    TheHdw.Wait 3 * mS
    
    TheHdw.Patterns("PORT_HiBench_IP_DDR_0P65_INT").Start   ' to setup some pins status to control device
    TheHdw.Wait 50 * mS
    
    Call ResultCheck(i, TheHiBench(i).comparedValue)
    
    
    
    ' ###################HiBench_LV_DDR_0V65 end###################
    
    
    '*********HiBench_LV_DDR_0V7 start**********************************************************************************************************************
    i = 2
    TheHiBench(i).Item_Pass = True
    TheHiBench(i).per_site_result = False
    TheHiBench(i).per_site_result_high = False
    TheHiBench(i).comparedValue = 32
    If TheHiBench(i).VolType = "AVS" Then
        lev_spec = "VDD08_DDR_0V7"
        SPEC_NAME = "DDR"
        label = "0P7V"
        For Each Site In TheExec.Sites
            lev_voltage(Site) = GET_VMIN_FCK_BY_AVS_FORMULA(CDbl(hpmmap.Item(lev_spec)(Site)), CDbl(hpmmapx(lev_spec)(Site)), SPEC_NAME, label, CLng(Site))
        Next Site
        Call TheExec.Overlays.ApplyUniformSpecToHW(lev_spec, lev_voltage, False, False)

    ElseIf TheHiBench(i).VolType = "NORMAL" Then
        
    End If
    
    Call TheExec.Overlays.ApplyUniformSpecToHW("VDD08_DDR", "0.7", False, False)
    TheHdw.Wait 3 * mS
    
    TheHdw.Patterns("PORT_HiBench_IP_DDR_0P7_INT").Start   ' to setup some pins status to control device
    TheHdw.Wait 50 * mS
    
    Call ResultCheck(i, TheHiBench(i).comparedValue)
    
    ' ###################HiBench_LV_DDR_0V7 end###################
    
    ' waiting to add code here
    
    Dim idx_StructLine As Long
    Dim HardBinResult As New SiteLong
    For idx_StructLine = 0 To TheHiBench(i).MaxStructLine - 1
        If TheHiBench(i).HiBenchResult(idx_StructLine).HiBenchResultName = "IMG_Version_xloader" Then
            For Each Site In TheExec.Sites
                'HardBinResult(Site) = GetHardBinResult(Site)
                TheExec.Datalog.WriteComment "Site" + CStr(Site) + "_" + "*HardBinResultis*:" + HardBinResult
                If (TheHiBench(i).HiBenchResult(idx_StructLine).ReadValue(Site) = TheHiBench(i).HiBenchResult(idx_StructLine).ExpectValue) And HardBinResult = -1 Then
                    Call TheExec.Overlays.ApplyUniformSpecToHW("VDD18_EFUSE", "1.8", False, False)
                    TheHdw.Wait 3 * mS
                Else
                    Call TheExec.Overlays.ApplyUniformSpecToHW("VDD18_EFUSE", "0", False, False)
                    TheHdw.Wait 3 * mS
                End If
            Next Site
        End If
    Next idx_StructLine
    '.....
    
    
    '*********HiBench_DDR start**********************************************************************************************************************
    
    i = 3
    TheHiBench(i).Item_Pass = True
    TheHiBench(i).per_site_result = False
    TheHiBench(i).per_site_result_high = False
    TheHiBench(i).comparedValue = 32
    If TheHiBench(i).VolType = "AVS" Then

        lev_spec = "VDD08_DDR_0V7"
        SPEC_NAME = "DDR"
        label = "0P8V"
        For Each Site In TheExec.Sites
            lev_voltage(Site) = GET_VMIN_FCK_BY_AVS_FORMULA(CDbl(hpmmap.Item(lev_spec)(Site)), CDbl(hpmmapx(lev_spec)(Site)), SPEC_NAME, label, CLng(Site))
        Next Site
        Call TheExec.Overlays.ApplyUniformSpecToHW(lev_spec, lev_voltage, False, False)

    ElseIf TheHiBench(i).VolType = "NORMAL" Then

        
    End If
    
    Call TheExec.Overlays.ApplyUniformSpecToHW("VDD075_PERI", "0.8", False, False)
    Call TheExec.Overlays.ApplyUniformSpecToHW("VDD08_DDR", "0.8", False, False)
    TheHdw.Wait 3 * mS
    
    TheHdw.Patterns("HiBench_IP_DDR_0P8_INT_Burst").Start   ' to setup some pins status to control device
    TheHdw.Wait 50 * mS
    
    Call ResultCheck(i, TheHiBench(i).comparedValue)
    ' ################### HiBench_DDR end ###################
    
    '*********HiBench_LV_DDR_0V75 start**********************************************************************************************************************
    i = 4
    TheHiBench(i).Item_Pass = True
    TheHiBench(i).per_site_result = False
    TheHiBench(i).per_site_result_high = False
    TheHiBench(i).comparedValue = 32
    If TheHiBench(i).VolType = "AVS" Then

        lev_spec = "VDD08_DDR_0V7"
        SPEC_NAME = "DDR"
        label = "0P75V"
        For Each Site In TheExec.Sites
            lev_voltage(Site) = GET_VMIN_FCK_BY_AVS_FORMULA(CDbl(hpmmap.Item(lev_spec)(Site)), CDbl(hpmmapx(lev_spec)(Site)), SPEC_NAME, label, CLng(Site))
        Next Site
        Call TheExec.Overlays.ApplyUniformSpecToHW(lev_spec, lev_voltage, False, False)
        

    ElseIf TheHiBench(i).VolType = "NORMAL" Then

        
    End If
    
    Call TheExec.Overlays.ApplyUniformSpecToHW("VDD075_PERI", "0.8", False, False)
    Call TheExec.Overlays.ApplyUniformSpecToHW("VDD08_DDR", "0.75", False, False)
    TheHdw.Wait 3 * mS
    
    TheHdw.Patterns("HiBench_IP_DDR_0P75_INT_Burst").Start   ' to setup some pins status to control device
    TheHdw.Wait 50 * mS
    
    Call ResultCheck(i, TheHiBench(i).comparedValue)
    
    
    '######### HiBench_LV_DDR_0V75 end#############################################
    
    
    
    '************HiBench_IP start**********************************************************************************************************************
    i = 5
    TheHiBench(i).Item_Pass = True
    TheHiBench(i).per_site_result = False
    TheHiBench(i).per_site_result_high = False
    TheHiBench(i).comparedValue = 64


    TheHdw.Patterns("HiBench_IP_Result").Start   ' to setup some pins status to control device
    TheHdw.Wait 50 * mS
    
    tmpStr = "3,4"
    
    Call ResultCheck(i, TheHiBench(i).comparedValue, tmpStr)
    
    '************HiBench_IP end*******************
    
    
    '******************CPU_L start**********************************************************************************************************************
    i = 6
    TheHiBench(i).Item_Pass = True
    TheHiBench(i).per_site_result = False
    TheHiBench(i).per_site_result_high = False
    TheHiBench(i).comparedValue = 16
   

    Call TheExec.Overlays.ApplyUniformSpecToHW("VDD08_CPU_L", "0.85", False, False)
    Call TheExec.Overlays.ApplyUniformSpecToHW("VDD08_CPU_BM", "1.05", False, False)
    Call TheExec.Overlays.ApplyUniformSpecToHW("VDDC08_MEM_CPUM", "1", False, False)
    TheHdw.Wait 3 * mS
    
    TheHdw.Patterns("PORT_HiBench_IP_UFS_CPU_L_INT").Start   ' to setup some pins status to control device
    TheHdw.Wait 50 * mS
  
    tmpStr = "3,4,5"
    Call ResultCheck(i, TheHiBench(i).comparedValue, tmpStr)
    ' ###################CPU_L end ###################
    
    '******************L3 start**********************************************************************************************************************   ' need modify
    i = 7
    TheHiBench(i).Item_Pass = True
    TheHiBench(i).per_site_result = False
    TheHiBench(i).per_site_result_high = False
    
    TheHdw.Patterns("PORT_HiBench_IP_UFS_L3_INT").Start   ' to setup some pins status to control device
    TheHdw.Wait 50 * mS
    
    Dim i_time As Long
    Dim i_match_time As Long
    Dim Idx As Long

    For i_time = 0 To TheHiBench(i).ExecuteTime_ms / i_match_time - 1
        Call CMEM_CAPTURE(i)
        For Each Site In TheExec.Sites
            If TheHiBench(i).per_site_result_high(Site) = False Then
                TheHiBench(i).per_site_result_high(Site) = True
                For Idx = 0 To TheHiBench(i).MaxStructLine - 1
                    If TheHiBench(i).HiBenchResult(Idx).ExpectValue <> -1 Then
                        'If TheHiBench(Idx).HiBenchResult(i).HiBenchResultName = TheHiBench(i).HiBenchResult(0).HiBenchResultName Then
                            'If TheHiBench(Idx).HiBenchResult(i).ReadValue(Site) < i_comparedVal Then
                                TheHiBench(i).per_site_result(Site) = TheHiBench(i).per_site_result(Site) And (TheHiBench(i).HiBenchResult(Idx).ReadValue(Site) = TheHiBench(i).HiBenchResult(Idx).ExpectValue)
                                'TheHiBench(Idx).per_site_result_high(Site) = False
                            'Else
                                'TheHiBench(Idx).per_site_result(Site) = False
                                'TheHiBench(Idx).per_site_result_high(Site) = TheHiBench(Idx).per_site_result_high(Site) And True
                            'End If
                        'Else
                            'TheHiBench(Idx).per_site_result(Site) = TheHiBench(Idx).per_site_result(Site) And (TheHiBench(Idx).HiBenchResult(i).ReadValue(Site) = TheHiBench(Idx).HiBenchResult(i).ExpectValue)
                        'End If
                    End If
                Next Idx
                TheHiBench(i).MatchLoopTime(Site) = (i_time + 1) * i_match_time
                If (i_time = TheHiBench(i).ExecuteTime_ms / i_match_time - 1) Then 'Or (TheHiBench(i).HiBenchResult(0).ExpectValue <> -1 And (TheHiBench(i).per_site_result_high = True)) Then
                    TheHiBench(i).Item_Pass(Site) = False
                End If
            End If
            TheHiBench(i).all_site_pass = TheHiBench(i).all_site_pass And (TheHiBench(i).per_site_result(Site) = True Or TheHiBench(4).per_site_result(Site) = False Or TheHiBench(3).per_site_result(Site) = False Or TheHiBench(5).per_site_result(Site) = False Or TheHiBench(6).per_site_result(Site) = False)
            TheHiBench(i).all_site_high = TheHiBench(i).all_site_high And TheHiBench(i).per_site_result_high(Site)
        Next Site
        
        If TheHiBench(i).all_site_pass = True Or TheHiBench(i).all_site_high = True Then
            Exit For
        End If
    Next i_time
    
    ' ###################L3 end###################
    
    
    
    '******************CPU_M start**********************************************************************************************************************
    
    i = 8
    TheHiBench(i).Item_Pass = True
    TheHiBench(i).per_site_result = False
    TheHiBench(i).per_site_result_high = False
    TheHiBench(i).comparedValue = 16

    TheHdw.Patterns("PORT_HiBench_IP_UFS_CPU_M_INT").Start   ' to setup some pins status to control device
    TheHdw.Wait 50 * mS
    
    tmpStr = "3,4,5,6,7"
    Call ResultCheck(i, TheHiBench(i).comparedValue, tmpStr)
    
    
    ' ###################CPU_M end###################
    
    
    
    '******************CPU_B start**********************************************************************************************************************
    
    i = 9
    TheHiBench(i).Item_Pass = True
    TheHiBench(i).per_site_result = False
    TheHiBench(i).per_site_result_high = False
    TheHiBench(i).comparedValue = 16
    
    TheHdw.Patterns("PORT_HiBench_IP_UFS_CPU_B_INT").Start   ' to setup some pins status to control device
    TheHdw.Wait 50 * mS

    
    tmpStr = "3,4,5,6,7,8"
    Call ResultCheck(i, TheHiBench(i).comparedValue, tmpStr)
    
     ' ###################CPU_B end###################
    
    
    '******************GPU_LV_HiBench_IP startt**********************************************************************************************************************
    
    i = 10
    TheHiBench(i).Item_Pass = True
    TheHiBench(i).per_site_result = False
    TheHiBench(i).per_site_result_high = False
    TheHiBench(i).comparedValue = 64

    Call TheExec.Overlays.ApplyUniformSpecToHW("VDD075_PERI", "0.6", False, False)
    Call TheExec.Overlays.ApplyUniformSpecToHW("VDD07_GPU", "0.55", False, False)
    TheHdw.Wait 3 * mS
    
    TheHdw.Patterns("PORT_HiBench_IP_UFS_LV_INT").Start   ' to setup some pins status to control device
    TheHdw.Wait 50 * mS
    
    tmpStr = "3,4,5,6,7,8,9"
    Call ResultCheck(i, TheHiBench(i).comparedValue, tmpStr)
    TheHdw.Wait 50 * mS
    
    
    
    ' ###################GPU_LV_HiBench_IP end###################
    
   
    ' to restore DC spec
    
    Call TheExec.Overlays.ApplyUniformSpecToHW("VDD18_EFUSE", "0", False, False)
    Call TheExec.Overlays.ApplyUniformSpecToHW("VDD08_CPU_BM", Vspec_VDD08_CPU_BM, False, False)
    Call TheExec.Overlays.ApplyUniformSpecToHW("VDDC08_MEM_CPUM", Vspec_VDDC08_MEM_CPUM, False, False)
    Call TheExec.Overlays.ApplyUniformSpecToHW("VDD08_CPU_L", Vspec_VDD08_CPU_L, False, False)
    Call TheExec.Overlays.ApplyUniformSpecToHW("VDD075_PERI", Vspec_VDD075_PERI, False, False)
    Call TheExec.Overlays.ApplyUniformSpecToHW("VDD08_DDR", Vspec_VDD08_DDR, False, False)
    Call TheExec.Overlays.ApplyUniformSpecToHW("VDD08_NPU", Vspec_VDD08_NPU, False, False)
    Call TheExec.Overlays.ApplyUniformSpecToHW("VDD07_GPU", Vspec_VDD07_GPU, False, False)
    
    TheHdw.Digital.Pins("CLK_SLEEP").FreeRunningClock.Enabled = False
    TheHdw.Digital.Pins("CLK_SYSTEM").FreeRunningClock.Enabled = False
    TheHdw.Digital.Pins("UFS_REF_CLK").FreeRunningClock.Enabled = False

    TheHdw.Wait 10 * mS
    ' to print debug info

    If i_DebugMode = True Then
        Dim match_flag As String
        
        
        For Each Site In TheExec.Sites
        
            TheExec.Datalog.WriteComment Space(5) + "Site" + Space(15) + "IP_NAME" + Space(25) + "SubIp_Name" + Space(15) + "Execute_Time" + Space(15) + "Expect_Value" + Space(15) + "Read_Value" + Space(15) + "Match_Flag" + Space(15)
            TheExec.Datalog.WriteComment "-------------------------------------------------------------------------------------------------------------------------"
            i = 0
            For Idx = 0 To TheHiBench(i).MaxStructLine - 1
                If TheHiBench(i).HiBenchResult(Idx).ExpectValue = -1 Or (TheHiBench(i).HiBenchResult(Idx).ExpectValue = TheHiBench(i).HiBenchResult(Idx).ReadValue(Site)) Then
                    match_flag = "P"
                End If
                TheExec.Datalog.WriteComment Space(5) + CStr(Site) + Space(15) + "DDR_0V60" + Space(25) + TheHiBench(i).ItemName + Space(15) + TheHiBench(i).MatchLoopTime(Site) + Space(15) + TheHiBench(i).HiBenchResult(Idx).ExpectValue + Space(15) + _
                                            TheHiBench(i).HiBenchResult(Idx).ReadValue(Site) + Space(15) + match_flag + Space(5)
                match_flag = "F"
            Next Idx
            
            i = 1
            For Idx = 0 To TheHiBench(i).MaxStructLine - 1
                If TheHiBench(i).HiBenchResult(Idx).ExpectValue = -1 Or (TheHiBench(i).HiBenchResult(Idx).ExpectValue = TheHiBench(i).HiBenchResult(Idx).ReadValue(Site)) Then
                    match_flag = "P"
                End If
                TheExec.Datalog.WriteComment Space(5) + CStr(Site) + Space(15) + "DDR_0V65" + Space(25) + TheHiBench(i).ItemName + Space(15) + TheHiBench(i).MatchLoopTime(Site) + Space(15) + TheHiBench(i).HiBenchResult(Idx).ExpectValue + Space(15) + _
                                            TheHiBench(i).HiBenchResult(Idx).ReadValue(Site) + Space(15) + match_flag + Space(5)
                match_flag = "F"
            Next Idx
            
            i = 2
            For Idx = 0 To TheHiBench(i).MaxStructLine - 1
                If TheHiBench(i).HiBenchResult(Idx).ExpectValue = -1 Or (TheHiBench(i).HiBenchResult(Idx).ExpectValue = TheHiBench(i).HiBenchResult(Idx).ReadValue(Site)) Then
                    match_flag = "P"
                End If
                TheExec.Datalog.WriteComment Space(5) + CStr(Site) + Space(15) + "DDR_0V70" + Space(25) + TheHiBench(i).ItemName + Space(15) + TheHiBench(i).MatchLoopTime(Site) + Space(15) + TheHiBench(i).HiBenchResult(Idx).ExpectValue + Space(15) + _
                                            TheHiBench(i).HiBenchResult(Idx).ReadValue(Site) + Space(15) + match_flag + Space(5)
                match_flag = "F"
            Next Idx
            
            i = 3
            For Idx = 0 To TheHiBench(i).MaxStructLine - 1
                If TheHiBench(i).HiBenchResult(Idx).ExpectValue = -1 Or (TheHiBench(i).HiBenchResult(Idx).ExpectValue = TheHiBench(i).HiBenchResult(Idx).ReadValue(Site)) Then
                    match_flag = "P"
                End If
                TheExec.Datalog.WriteComment Space(5) + CStr(Site) + Space(15) + "DDR_0V80" + Space(25) + TheHiBench(i).ItemName + Space(15) + TheHiBench(i).MatchLoopTime(Site) + Space(15) + TheHiBench(i).HiBenchResult(Idx).ExpectValue + Space(15) + _
                                            TheHiBench(i).HiBenchResult(Idx).ReadValue(Site) + Space(15) + match_flag + Space(5)
                match_flag = "F"
            Next Idx
            
            i = 4
            For Idx = 0 To TheHiBench(i).MaxStructLine - 1
                If TheHiBench(i).HiBenchResult(Idx).ExpectValue = -1 Or (TheHiBench(i).HiBenchResult(Idx).ExpectValue = TheHiBench(i).HiBenchResult(Idx).ReadValue(Site)) Then
                    match_flag = "P"
                End If
                TheExec.Datalog.WriteComment Space(5) + CStr(Site) + Space(15) + "DDR_0V75" + Space(25) + TheHiBench(i).ItemName + Space(15) + TheHiBench(i).MatchLoopTime(Site) + Space(15) + TheHiBench(i).HiBenchResult(Idx).ExpectValue + Space(15) + _
                                            TheHiBench(i).HiBenchResult(Idx).ReadValue(Site) + Space(15) + match_flag + Space(5)
                match_flag = "F"
            Next Idx
            
            
            i = 5
            For Idx = 0 To TheHiBench(i).MaxStructLine - 1
                If TheHiBench(i).HiBenchResult(Idx).ExpectValue = -1 Or (TheHiBench(i).HiBenchResult(Idx).ExpectValue = TheHiBench(i).HiBenchResult(Idx).ReadValue(Site)) Then
                    match_flag = "P"
                End If
                TheExec.Datalog.WriteComment Space(5) + CStr(Site) + Space(15) + "HIBENCH_IP" + Space(25) + TheHiBench(i).ItemName + Space(15) + TheHiBench(i).MatchLoopTime(Site) + Space(15) + TheHiBench(i).HiBenchResult(Idx).ExpectValue + Space(15) + _
                                            TheHiBench(i).HiBenchResult(Idx).ReadValue(Site) + Space(15) + match_flag + Space(5)
                match_flag = "F"
            Next Idx
            
            i = 6
            For Idx = 0 To TheHiBench(i).MaxStructLine - 1
                If TheHiBench(i).HiBenchResult(Idx).ExpectValue = -1 Or (TheHiBench(i).HiBenchResult(Idx).ExpectValue = TheHiBench(i).HiBenchResult(Idx).ReadValue(Site)) Then
                    match_flag = "P"
                End If
                TheExec.Datalog.WriteComment Space(5) + CStr(Site) + Space(15) + "CPUL" + Space(25) + TheHiBench(i).ItemName + Space(15) + TheHiBench(i).MatchLoopTime(Site) + Space(15) + TheHiBench(i).HiBenchResult(Idx).ExpectValue + Space(15) + _
                                            TheHiBench(i).HiBenchResult(Idx).ReadValue(Site) + Space(15) + match_flag + Space(5)
                match_flag = "F"
            Next Idx
            
            i = 7
            For Idx = 0 To TheHiBench(i).MaxStructLine - 1
                If TheHiBench(i).HiBenchResult(Idx).ExpectValue = -1 Or (TheHiBench(i).HiBenchResult(Idx).ExpectValue = TheHiBench(i).HiBenchResult(Idx).ReadValue(Site)) Then
                    match_flag = "P"
                End If
                TheExec.Datalog.WriteComment Space(5) + CStr(Site) + Space(15) + "L3" + Space(25) + TheHiBench(i).ItemName + Space(15) + TheHiBench(i).MatchLoopTime(Site) + Space(15) + TheHiBench(i).HiBenchResult(Idx).ExpectValue + Space(15) + _
                                            TheHiBench(i).HiBenchResult(Idx).ReadValue(Site) + Space(15) + match_flag + Space(5)
                match_flag = "F"
            Next Idx
            
            i = 8
            For Idx = 0 To TheHiBench(i).MaxStructLine - 1
                If TheHiBench(i).HiBenchResult(Idx).ExpectValue = -1 Or (TheHiBench(i).HiBenchResult(Idx).ExpectValue = TheHiBench(i).HiBenchResult(Idx).ReadValue(Site)) Then
                    match_flag = "P"
                End If
                TheExec.Datalog.WriteComment Space(5) + CStr(Site) + Space(15) + "L3" + Space(25) + TheHiBench(i).ItemName + Space(15) + TheHiBench(i).MatchLoopTime(Site) + Space(15) + TheHiBench(i).HiBenchResult(Idx).ExpectValue + Space(15) + _
                                            TheHiBench(i).HiBenchResult(Idx).ReadValue(Site) + Space(15) + match_flag + Space(5)
                match_flag = "F"
            Next Idx
            
            i = 9
            For Idx = 0 To TheHiBench(i).MaxStructLine - 1
                If TheHiBench(i).HiBenchResult(Idx).ExpectValue = -1 Or (TheHiBench(i).HiBenchResult(Idx).ExpectValue = TheHiBench(i).HiBenchResult(Idx).ReadValue(Site)) Then
                    match_flag = "P"
                End If
                TheExec.Datalog.WriteComment Space(5) + CStr(Site) + Space(15) + "CPUM" + Space(25) + TheHiBench(i).ItemName + Space(15) + TheHiBench(i).MatchLoopTime(Site) + Space(15) + TheHiBench(i).HiBenchResult(Idx).ExpectValue + Space(15) + _
                                            TheHiBench(i).HiBenchResult(Idx).ReadValue(Site) + Space(15) + match_flag + Space(5)
                match_flag = "F"
            Next Idx
            
            i = 10
            For Idx = 0 To TheHiBench(i).MaxStructLine - 1
                If TheHiBench(i).HiBenchResult(Idx).ExpectValue = -1 Or (TheHiBench(i).HiBenchResult(Idx).ExpectValue = TheHiBench(i).HiBenchResult(Idx).ReadValue(Site)) Then
                    match_flag = "P"
                End If
                TheExec.Datalog.WriteComment Space(5) + CStr(Site) + Space(15) + "CPUB" + Space(25) + TheHiBench(i).ItemName + Space(15) + TheHiBench(i).MatchLoopTime(Site) + Space(15) + TheHiBench(i).HiBenchResult(Idx).ExpectValue + Space(15) + _
                                            TheHiBench(i).HiBenchResult(Idx).ReadValue(Site) + Space(15) + match_flag + Space(5)
                match_flag = "F"
            Next Idx
            
            i = 11
            For Idx = 0 To TheHiBench(i).MaxStructLine - 1
                If TheHiBench(i).HiBenchResult(Idx).ExpectValue = -1 Or (TheHiBench(i).HiBenchResult(Idx).ExpectValue = TheHiBench(i).HiBenchResult(Idx).ReadValue(Site)) Then
                    match_flag = "P"
                End If
                TheExec.Datalog.WriteComment Space(5) + CStr(Site) + Space(15) + "HIBENCH_IP_LV" + Space(25) + TheHiBench(i).ItemName + Space(15) + TheHiBench(i).MatchLoopTime(Site) + Space(15) + TheHiBench(i).HiBenchResult(Idx).ExpectValue + Space(15) + _
                                            TheHiBench(i).HiBenchResult(Idx).ReadValue(Site) + Space(15) + match_flag + Space(5)
                match_flag = "F"
            Next Idx
        
        Next Site
    
    
    End If

    

    ' to testlimit, if specified print for some item, select ... case could be used

    

    For i = 0 To UBound(TheHiBench)

        TheExec.Flow.TestLimit ResultVal:=TheHiBench(i).MatchLoopTime, TName:=TheHiBench(i).ItemName + "_Execute_Time"

        For j = 0 To TheHiBench(i).MaxStructLine - 1

            If TheHiBench(i).HiBenchResult(j).ExpectValue <> -1 Then

                TheExec.Flow.TestLimit ResultVal:=TheHiBench(i).HiBenchResult(j).ReadValue, TName:=TheHiBench(i).HiBenchResult(j).HiBenchResultName

            End If

        Next j

    Next i

    

errHandler:

    

    

End Function

Public Function CMEM_CAPTURE(Idx As Long) As Long
        
    Dim i As Long
    Dim Site As Variant
    Dim j As Long
    
    Dim i_CmemCapSize As Long
    Dim i_tempData As Long
    i_CmemCapSize = 1

    Dim i_CapCmem_PLD() As New PinListData
    ReDim i_CapCmem_PLD(TheHiBench(Idx).MaxStructLine - 1)
    TheHdw.Digital.CMEM.SetCaptureConfig 0, CmemCaptNone
    TheHdw.Digital.CMEM.SetCaptureConfig -1, CmemCaptSTV, tlCMEMCaptureSourceDutData
    TheHdw.Patterns(TheHiBench(Idx).PatName).Start
    TheHdw.Digital.Patgen.HaltWait
    
    For i = 0 To TheHiBench(Idx).MaxStructLine - 1
        i_CapCmem_PLD(i) = TheHdw.Digital.Pins(TheHiBench(Idx).HiBenchResult(i).CapPinNameList).CMEM.Data(0, 1, tlCMEMPackData)
        For Each Site In TheExec.Sites
            i_tempData = 0
            For j = 0 To UBound(Split(TheHiBench(Idx).HiBenchResult(i).CapPinNameList, ","))
                i_tempData = i_tempData * (2) + i_CapCmem_PLD(i)(Site)  ' the seq: pin(0) is MSB
            Next j
            TheHiBench(Idx).HiBenchResult(i).ReadValue(Site) = i_tempData
        Next Site
    Next i
       
End Function

Sub ResultCheck(Idx As Long, i_comparedVal As Long, Optional tmpStr As String)

    Dim i_time As Long
    Dim i_match_time As Long
    Dim i As Long
    Dim Site As Variant
    Dim tmpArr() As Long
    Dim tmp_RST As Boolean
    Dim tmp_i As Long
    
    If tmpStr <> "" Then
        tmpArr = Split(tmpStr, ",")
    End If
    
    i_match_time = match_time

    For i_time = 0 To TheHiBench(i).ExecuteTime_ms / i_match_time - 1
        Call CMEM_CAPTURE(Idx)
        For Each Site In TheExec.Sites
            If TheHiBench(Idx).per_site_result_high(Site) = False Then
                TheHiBench(Idx).per_site_result_high(Site) = True
                If TheHiBench(Idx).per_site_result(Site) = False Then
                    TheHiBench(Idx).per_site_result(Site) = True
                    For i = 0 To TheHiBench(Idx).MaxStructLine - 1
                        If TheHiBench(Idx).HiBenchResult(i).ExpectValue <> -1 Then
                            If TheHiBench(Idx).HiBenchResult(i).HiBenchResultName = TheHiBench(i).HiBenchResult(0).HiBenchResultName Then
                                If TheHiBench(Idx).HiBenchResult(i).ReadValue(Site) < i_comparedVal Then
                                    TheHiBench(Idx).per_site_result(Site) = TheHiBench(Idx).per_site_result(Site) And (TheHiBench(Idx).HiBenchResult(i).ReadValue(Site) = TheHiBench(Idx).HiBenchResult(i).ExpectValue)
                                    TheHiBench(Idx).per_site_result_high(Site) = False
                                Else
                                    TheHiBench(Idx).per_site_result(Site) = False
                                    TheHiBench(Idx).per_site_result_high(Site) = TheHiBench(Idx).per_site_result_high(Site) And True
                                End If
                            Else
                                TheHiBench(Idx).per_site_result(Site) = TheHiBench(Idx).per_site_result(Site) And (TheHiBench(Idx).HiBenchResult(i).ReadValue(Site) = TheHiBench(Idx).HiBenchResult(i).ExpectValue)
                            End If
                        End If
                    Next i
                    TheHiBench(Idx).MatchLoopTime(Site) = (i_time + 1) * i_match_time
                    If (i_time = TheHiBench(Idx).ExecuteTime_ms / i_match_time - 1) Or (TheHiBench(Idx).HiBenchResult(0).ExpectValue <> -1 And (TheHiBench(Idx).per_site_result_high = True)) Then
                        TheHiBench(Idx).Item_Pass(Site) = False
                    End If
                End If
            End If
            
            tmp_RST = False
            
            If tmpStr <> "" Then
                For tmp_i = 0 To UBound(tmpArr)
                    tmp_RST = tmp_RST Or (TheHiBench(tmp_i).Item_Pass = False)
                Next tmp_i
            End If
            
            TheHiBench(Idx).all_site_pass = TheHiBench(Idx).all_site_pass And (TheHiBench(Idx).per_site_result(Site) = True Or tmp_RST)
            TheHiBench(Idx).all_site_high = TheHiBench(Idx).all_site_high And TheHiBench(Idx).per_site_result_high(Site)
        Next Site
        
        If TheHiBench(Idx).all_site_pass = True Or TheHiBench(Idx).all_site_high = True Then
            Exit For
        End If
    Next i_time
    
    TheExec.Datalog.WriteComment "****" + TheHiBench(Idx).ItemName + " has been executed ! ****  " + "Item Index is " + CStr(Idx) + "  ****"

End Sub




Public Function GET_VMIN_FCK_BY_AVS_FORMULA(HpmVolt As Double, HpmxVolt As Double, AVS_IP_NAME As String, label As String, Site As Long) As Double

    Dim Volt As Double: Volt = 0.8
    Dim hpm_h300_l8_svt As Double
    
    
    If AVS_IP_NAME = "DDR" Then
        hpm_h300_l8_svt = hpmmap("DDR,300svt")(Site) * 1000
        TheExec.Datalog.WriteComment "u" + CStr(Site) + "=" + hpm_h300_l8_svt
        If hpm_h300_l8_svt > 870 Then
            If InStr(1, label, "0P8V", vbTextCompare) > 0 Then Volt = 760 * 0.001
            If InStr(1, label, "0P75V", vbTextCompare) > 0 Then Volt = 712.5 * 0.001
            If InStr(1, label, "0P7V", vbTextCompare) > 0 Then Volt = 665 * 0.001
            If InStr(1, label, "0P65V", vbTextCompare) > 0 Then Volt = 617.5 * 0.001
            If InStr(1, label, "0P6V", vbTextCompare) > 0 Then Volt = 760 * 0.001
        ElseIf hpm_h300_l8_svt <= 840 Then
            If InStr(1, label, "0P8V", vbTextCompare) > 0 Then Volt = 740 * 0.001
            If InStr(1, label, "0P75V", vbTextCompare) > 0 Then Volt = 690 * 0.001
            If InStr(1, label, "0P7V", vbTextCompare) > 0 Then Volt = 650 * 0.001
            If InStr(1, label, "0P65V", vbTextCompare) > 0 Then Volt = 610 * 0.001
            If InStr(1, label, "0P6V", vbTextCompare) > 0 Then Volt = 589.5 * 0.001
        Else
            If InStr(1, label, "0P8V", vbTextCompare) > 0 Then Volt = 760 * 0.001
            If InStr(1, label, "0P75V", vbTextCompare) > 0 Then Volt = 710 * 0.001
            If InStr(1, label, "0P7V", vbTextCompare) > 0 Then Volt = 660 * 0.001
            If InStr(1, label, "0P65V", vbTextCompare) > 0 Then Volt = 617.5 * 0.001
            If InStr(1, label, "0P6V", vbTextCompare) > 0 Then Volt = 589 * 0.001
        End If
            
    End If
    
    GET_VMIN_FCK_BY_AVS_FORMULA = Volt
    
End Function


