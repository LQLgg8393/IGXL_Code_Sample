Attribute VB_Name = "CBB_H20USR_Declaration"
' This module is part of HiLink CBBs(H20USR)
' This module is generated by CBB_Code_Gen_Tool_for_H16_H30, contact GSO Shanghai Team for more details
' Alpha 100  2015/09/30  initial release

Option Explicit

Global Const MacroCnt_H20USR As Long = 1

Global Const SubMacroCnt_H20USR As Long = 1
Global Const TxLaneCnt_H20USR As Long = 1
Global Const RxLaneCnt_H20USR As Long = 2
''''    ENUM START  >>>>

''' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ '

'''++     This H20USR_Enum_MacroName is generated by auto-gen tool     ++'
'''++     It should be checked mannually and confirmed by project     ++'
'''++     owner when the CBB is applied to a new project              ++'
''' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ '

Public Enum H20USR_Enum_MacroName

    'Macro name and macro ID

    Macro_SDS = 0

    Default_H20USR_Macro = Macro_SDS

End Enum

''''Below lines are fixed by CBB, please do not put your code under this line.
''''Falure to do so will make your CBB update hard to perform and time comsuming.

'>> CS


'>> TX


'>> RX


Public Enum H20USR_DataIndex
    H20USR_14G7456 = 1
    H20USR_7G3728 = 2
    H20USR_3G6864 = 3
    H20USR_1G8432 = 4
    H20USR_0G9216 = 5
    H20USR_0G4608 = 6
End Enum


Public Enum H20USR_PRBSPATSEL
    H20USR_PRBS_7 = 0
    H20USR_PRBS_9 = 1
    H20USR_PRBS_10 = 2
    H20USR_PRBS_11 = 3
    H20USR_PRBS_15 = 4
    H20USR_PRBS_20 = 5
    H20USR_PRBS_23 = 6
    H20USR_PRBS_31 = 7
    H20USR_CustomerPattern = 16
End Enum

Public Enum H20USR_Custom_Data
    H20USR_All0 = 0
    H20USR_All1 = 1
End Enum

Public Enum H20USR_LoopBackMode
    H20USR_ExternalLpbk = 0
    H20USR_InternalLpbk = 1
End Enum

Public Enum H20USR_DATA_SliceNumber
    H20USR_DS_00 = 0
    H20USR_DS_01 = 1
    H20USR_DS_02 = 2
    H20USR_DS_03 = 3
    H20USR_DS_04 = 4
    H20USR_DS_05 = 5
    H20USR_DS_06 = 6
    H20USR_DS_07 = 7
    H20USR_DS_08 = 8
    H20USR_DS_09 = 9
    H20USR_DS_10 = 10
    H20USR_DS_11 = 11
    H20USR_DS_12 = 12
    H20USR_DS_13 = 13
    H20USR_DS_14 = 14
    H20USR_DS_15 = 15
    H20USR_DS_BroadCast = -1
End Enum

Public Enum H20USR_Jitter_Inj_Idx
    H20USR_Amp0_078UI = 0
    H20USR_Amp0_157UI = 1
    H20USR_Amp0_235UI = 2
    H20USR_Amp0_313UI = 3
    H20USR_Amp0_392UI = 4
    H20USR_Amp0_470UI = 5
    H20USR_Amp0_548UI = 6
    H20USR_Amp0_626UI = 7
    H20USR_Amp0_705UI = 8
    H20USR_Amp0_783UI = 9
    H20USR_Amp0_861UI = 10
    H20USR_Amp0_940UI = 11
    H20USR_Amp1_018UI = 12
    H20USR_Amp1_096UI = 13
    H20USR_Amp1_175UI = 14
    H20USR_Amp1_253UI = 15
    H20USR_NoJitter = -1         'no jitter
End Enum


Public Enum H20USR_Result_Type
    H20USR_ErrorCount = 0
    H20USR_PrbsErrCheck = 1
    H20USR_EyeWidth = 2
    H20USR_EyeWidth_A = 3
    H20USR_EyeHeight = 4
    H20USR_EyeHeight_A = 5
    H20USR_Eyecenter = 6
    H20USR_EyeCenter_A = 7
    H20USR_TestResults_PLD = 8
    H20USR_TestResults_A_PLD = 9
    H20USR_TestResults_B_PLD = 10
    H20USR_Customer_Datalog = 11
End Enum

Public Enum H20USR_Rx_RelayMode
    Rx_RF1_ExtLpbk = 0
    Rx_RF2_ExtLpbk = 1
    Rx_RF3_ExtLpbk = 2
    Rx_RF4_ExtLpbk = 3
End Enum

Public Enum H20USR_Tx_RelayMode
    Tx_RF1_NA = 0
    Tx_RF2_NA = 1
    Tx_RF3_NA = 2
    Tx_RF4_NA = 3
End Enum

Public Enum H20USR_RegSelect
    Dump_All = 0
    Dump_All_Except_TXD_RXD = 1
End Enum


Public Enum H20USR_StressEnb
    H20USR_StressDisable = 0
    H20USR_StressEnable = 1
End Enum

Public Enum SwitchTerm
    Term_Extlpbk = 0
    Term_UP1600 = 1
    Term_US10G = 2
    Term_Floating = 3
    Term_NA = 4
End Enum

'Add by ANYUE for API tool
Public Enum H20USR_API_Type
    H20USR_DS0_API = 0
    H20USR_DS1_API = 1
    H20USR_DS2_API = 2
    H20USR_DS3_API = 3
    H20USR_DS4_API = 4
    H20USR_DS5_API = 5
    H20USR_DS6_API = 6
    H20USR_DS7_API = 7
    H20USR_DS8_API = 8
    H20USR_DS9_API = 9
    H20USR_DS10_API = 10
    H20USR_DS11_API = 11
    H20USR_DS12_API = 12
    H20USR_DS13_API = 13
    H20USR_DS14_API = 14
    H20USR_DS15_API = 15
    H20USR_TOP_API = 16
    H20USR_CS_API = 17
    H20USR_ABIST_API = 18
    H20USR_DS_API_BroadCast = -1
End Enum


'Add by Charles for API tool  0521 update
Public Enum H20USR_API_Class
    H20USR_TOP_API_Class = 0
    H20USR_CS_API_Class = 1
    H20USR_DS_API_Class = 2
    H20USR_ABIST_API_Class = 3
End Enum


Public Enum H20USR_SubMacro_Sel
    SubMacro0_Sel = 0
    SubMacro1_Sel = 1
    SubMacro_AllSel = -1
End Enum

Public Enum H20USR_TxClkDiv_Type
    DIVM = 0
    DIV5 = 1
    DIV7 = 2
End Enum

Public Enum H20USR_SwitchRelayMode
    Switch_Port1 = 0
    Switch_Port2 = 1
    Switch_Port3 = 2
    Switch_Port4 = 3
    Switch_Isolation1 = 4
    Switch_50oumu = 5
    Switch_Isolation2 = 6
    Switch_Shutdown = 7
End Enum

Global glb_H20USR_Current_DataIndex As Long
''''ENUM END >>>>


''''Type START>>>>
''Type H20USR_MacroInfo is only applicable if the DUT has H20USR IP(s) in it.
Type Term_Info
    Selected As Boolean
    Terminal As SwitchTerm
End Type

Type Switch
    TxSwitchPath(0 To 3) As Term_Info
    RxSwitchPath(0 To 3) As Term_Info
End Type

Type H20USR_LaneInfo
    Selected As Boolean
    TxP As String
    TxN As String
    RxP As String
    RxN As String
    RF1_ExtLpbk_FromTxP As String
    RF2_ExtLpbk_FromTxP As String
End Type

Type H20USR_SubMacroInfo
    Selected As Boolean
    Name As String
    TX_LanCnt As Long
    RX_LanCnt As Long
End Type

Type H20USR_MacroInfo
    Selected As Boolean
    Name As String
    SDR_TDI_ID As String
    SDR_TDO_ID As String
    TX_LanCnt As Long
    RX_LanCnt As Long
    CR_LanCnt As Long
    Lane(0 To 15) As H20USR_LaneInfo
    EnlaneCnt As Long
    SubMacro(0 To 1) As H20USR_SubMacroInfo
End Type

Type H20USR_IP
    SIR_TDI_com As String
    SIR_TDO_com As String
    SIR_AHB_com As String
    SDR_TDI_ID_All As String
    SDR_TDI_ID_Enabled As String
    SDR_TDI_bitwidth As Long
    SDR_TDO_bitwidth As Long
    Macro(MacroCnt_H20USR - 1) As H20USR_MacroInfo
End Type

'''for datarate config
Type H20USR_TxTargetDataRate
    Lane(0 To TxLaneCnt_H20USR - 1) As Long
End Type
Type H20USR_TxCurrentDataRate
    Lane(0 To TxLaneCnt_H20USR - 1) As Long
End Type
Type H20USR_RxTargetDataRate
    Lane(0 To RxLaneCnt_H20USR - 1) As Long
End Type
Type H20USR_RxCurrentDataRate
    Lane(0 To RxLaneCnt_H20USR - 1) As Long
End Type
''''Type End>>>>

'*********************For New Abist Test*********************

''AdcFunction
Public Enum H20USR_DataLog_Mode
    H20USR_PassFail_mode = 0
    H20USR_Detail_mode = 1
End Enum


''''Global Variables Declaration START>>>>
Global H20USR As New Cls_H20USR
Global H20USR_SerialPinMap_PLD As New PinListData
Global H20USR_RXSerialPinMap_PLD As New PinListData
Global H20USR_TXSerialPinMap_PLD As New PinListData
Global H20USR_SubMacro0_RXSerialPinMap_PLD As New PinListData
Global H20USR_SubMacro0_TXSerialPinMap_PLD As New PinListData
Global H20USR_SubMacro1_RXSerialPinMap_PLD As New PinListData
Global H20USR_SubMacro1_TXSerialPinMap_PLD As New PinListData
Global H20USR_Cal_Done As New PinListData
Global glb_H20USR_Datarate_Gbps As Double

'Datalog Related
Global glb_H20USR_LogType As Long
Global glb_H20USR_SimpleLog As Boolean
Global glb_H20USR_RegLog_ON As Boolean

'CSR File Related
Global glb_H20USR_CSRDSP_OK As Long
Global glb_H20USR_Released As Long
Global glb_H20USR_CSRDSP_SiteAvailable As New SiteBoolean

'Verify Related
Global glb_H20USR_CSR_Verify_En As Boolean
Global glb_H20USR_CSR_Verify_DebugMode_En As Boolean
Global glb_H20USR_CSR_Verify_1Lane_En As Boolean
Global glb_H20USR_CSR_Verify_Rst As New SiteLong

'CSR Recording Related
Global glb_H20USR_Refresh_PAM_Recall As Boolean
Global glb_H20USR_NPAM_FNum As Integer
Global glb_H20USR_NPAM_FName As String
Global glb_H20USR_NPAM_FName_1 As String
Global glb_H20USR_NPAM_FName_2 As String
Global glb_H20USR_NPAM_FName_3 As String
Global glb_H20USR_NPAM_FName_4 As String
Global glb_H20USR_NPAM_FName_IsUsing As String
Global glb_H20USR_CSR_Rcd_ING As Boolean


'Other
Global glb_H20USR_RecordModule_ON As Boolean
Global glb_H20USR_tlNWireCMEMMoveMode As tlNWireCMEMMoveMode
Global glb_H20USR_CSRDirPath As String
Global glb_H20USR_FWVer_LogOn As Boolean

Global glb_H20USR_NewAbistTestInitializedIsOK As Boolean

'' For ATB Test
Global glb_H20USR_Test_FW_Version As String
Global glb_H20USR_CAL_FW_Version As String

Global mUI_Pt As Long          'UI Steps
Global LBRes As New SiteDouble           'Compensation for Trace Res
Global H20USR_TxRegDrvVrefSelCalCode_PLD As New PinListData

Global glb_H20USR_Refclk_MHz As Double

Global glb_current_page_addr As Long

''''END OF CBB_H20USR_Declaration
