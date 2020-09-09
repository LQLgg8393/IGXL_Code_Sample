Attribute VB_Name = "HiBenchLeakage"
Option Explicit



Public Function HiBench_Leakage_Current_Denver(i_DebugMode As Boolean, i_PowerPins As String, i_HiBench_Leakage_ExecuteTime_ms As Double, i_SetupPat As String, i_DigCapPat As String) As Long


    Dim i As Long
    Dim i_match_time As Long: i_match_time = 40 'ms
    Dim i_time As Long
    Dim j As Long
    Dim Site As Variant

    

    Dim HiBenchLeakage As STR_HIBENCHTESTINFO

    HiBenchLeakage.MatchLoopTime = 0
    HiBenchLeakage.MaxStructLine = 1
    HiBenchLeakage.per_site_result = False
    HiBenchLeakage.all_site_pass = False

    For i = 0 To HiBenchLeakage.MaxStructLine - 1
        HiBenchLeakage.HiBenchResult(i).ReadValue = 0
        HiBenchLeakage.HiBenchResult(i).HiBenchResultName = "HiBench_Leakage_Result"
        HiBenchLeakage.HiBenchResult(i).CapPinNameList = "PMU_PER_EN"
        HiBenchLeakage.HiBenchResult(i).ExpectValue = 0
        HiBenchLeakage.HiBenchResult(i).DigCap_Pos = "1"
        HiBenchLeakage.ExecuteTime_ms = i_HiBench_Leakage_ExecuteTime_ms
    Next i
    
    ' to get initial value of dc/ac spec

    Dim Vspec_VDD075_SYS As Double: Vspec_VDD075_SYS = TheExec.Specs.DC.Item("VDD075_SYS").CurrentValue
    Dim Vspec_VDD18_IO As Double: Vspec_VDD18_IO = TheExec.Specs.DC.Item("VDD18_IO").CurrentValue
    Dim Vspec_VDD075_MEM_MODEM As Double: Vspec_VDD075_MEM_MODEM = TheExec.Specs.DC.Item("VDD075_MEM_MODEM").CurrentValue
    Dim Vspec_VDD11_DDR As Double: Vspec_VDD11_DDR = TheExec.Specs.DC.Item("VDD11_DDR").CurrentValue
    Dim Vspec_AVDD12_SYS_PERI As Double: Vspec_AVDD12_SYS_PERI = TheExec.Specs.DC.Item("AVDD12_SYS_PERI").CurrentValue
    Dim Vspec_VDD075_MEM_SYS As Double: Vspec_VDD075_MEM_SYS = TheExec.Specs.DC.Item("VDD075_MEM_SYS").CurrentValue
  

    Dim i_CmemCapSize As Long
    Dim i_tempData As Long
    i_CmemCapSize = 1
    Dim i_CapCmem_PLD() As New PinListData
    ReDim i_CapCmem_PLD(HiBenchLeakage.MaxStructLine - 1)
    
    
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
    
    
    TheHdw.Patterns(i_SetupPat).Load
    TheHdw.Patterns(i_SetupPat).Start
    TheHdw.Digital.Patgen.HaltWait
    
    For i_time = 0 To HiBenchLeakage.ExecuteTime_ms / i_match_time - 1
        
        TheHdw.Digital.CMEM.SetCaptureConfig 0, CmemCaptNone
        TheHdw.Digital.CMEM.SetCaptureConfig -1, CmemCaptSTV, tlCMEMCaptureSourceDutData
        
    ''    rdi.setPatBurst("HiBench_IP_INT_LEAKAGE_Burst",TA::MD5_OFF);
    ''    RDI_BEGIN( (rdiRunMode==0)?(TA::PROD):(TA::BURST) );
    ''        rdi.port(PortName_19M2).func().label("PORT_HiBench_IP_UFS_INT_LEAKAGE").execute();
    ''        rdi.port(PortName_19M2_UFS).func().label("PORT_19P2M_UFS").execute();
    ''        rdi.port(PortName_32K).func().label("PORT_32K").execute();
    ''        rdi.port(PortName_19M2_CLK).func().label("PORT_19P2M").execute();
    ''    RDI_END();
        TheHdw.Patterns(i_DigCapPat).Start  'to burst pattern "HiBench_IP_DigCap_LEAKAGE_Burst"
        TheHdw.Digital.Patgen.HaltWait
        
        'to get ReadValue from captured data
        For i = 0 To HiBenchLeakage.MaxStructLine - 1
            i_CapCmem_PLD(i) = TheHdw.Digital.Pins(HiBenchLeakage.HiBenchResult(i).CapPinNameList).CMEM.Data(0, 1, tlCMEMPackData)
            For Each Site In TheExec.Sites
                i_tempData = 0
                For j = 0 To UBound(Split(HiBenchLeakage.HiBenchResult(i).CapPinNameList, ","))
                    i_tempData = i_tempData * (2) + i_CapCmem_PLD(i)(Site)  ' the seq: pin(0) is MSB
                Next j
                HiBenchLeakage.HiBenchResult(i).ReadValue(Site) = i_tempData
            Next Site
        Next i
        
        For Each Site In TheExec.Sites
            If HiBenchLeakage.per_site_result(Site) = False Then
                HiBenchLeakage.per_site_result(Site) = True
                For i = 0 To HiBenchLeakage.MaxStructLine - 1
                    If HiBenchLeakage.HiBenchResult(i).ExpectValue <> -1 Then
                        HiBenchLeakage.per_site_result(Site) = HiBenchLeakage.per_site_result(Site) And (HiBenchLeakage.HiBenchResult(i).ReadValue(Site) = HiBenchLeakage.HiBenchResult(i).ExpectValue)
                    End If
                Next i
                HiBenchLeakage.MatchLoopTime(Site) = (i_time + 1) * i_match_time
            End If
            
            HiBenchLeakage.all_site_pass = HiBenchLeakage.all_site_pass And HiBenchLeakage.per_site_result(Site)
        Next Site
        
        If HiBenchLeakage.all_site_pass = True Then
            Exit For
        End If
            
    Next i_time
    
    
    '------------------------------------------RETENTION MODE---------------------------------------------
    ' to measure VDD leakage
    

    With TheHdw.DCVS.Pins(i_PowerPins)
        .currentRange.Value = 5
        .Meter.Mode = tlDCVSMeterCurrent
        .Meter.currentRange.Value = 50 * ma
        
    End With
    
    TheHdw.Wait 100 * mS
    Dim PLD As New PinListData
    'dpsTask.samples(1024).wait(100 ms).trigMode(TM::INTERNAL).execMode(TM::PVAL).execute();
    PLD = TheHdw.DCVS.Pins(i_PowerPins).Meter.Read(tlStrobe, 10)
    
   
    TheExec.Flow.TestLimit ResultVal:=PLD
    
 ' to restore dc spec
    Call TheExec.Overlays.ApplyUniformSpecToHW("VDD075_SYS", Vspec_VDD075_SYS, False, False)
    Call TheExec.Overlays.ApplyUniformSpecToHW("VDD18_IO", Vspec_VDD18_IO, False, False)
    Call TheExec.Overlays.ApplyUniformSpecToHW("VDD075_MEM_MODEM", Vspec_VDD075_MEM_MODEM, False, False)
    Call TheExec.Overlays.ApplyUniformSpecToHW("VDD11_DDR", Vspec_VDD11_DDR, False, False)
    Call TheExec.Overlays.ApplyUniformSpecToHW("AVDD12_SYS_PERI", Vspec_AVDD12_SYS_PERI, False, False)
    Call TheExec.Overlays.ApplyUniformSpecToHW("VDD075_MEM_SYS", Vspec_VDD075_MEM_SYS, False, False)
    
    TheHdw.Digital.Pins("CLK_SLEEP").FreeRunningClock.Enabled = False
    TheHdw.Digital.Pins("CLK_SYSTEM").FreeRunningClock.Enabled = False
    TheHdw.Digital.Pins("UFS_REF_CLK").FreeRunningClock.Enabled = False


End Function


