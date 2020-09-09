Attribute VB_Name = "GlobalVariables"
Option Explicit
Global Site As Variant
Global GlobalVariable_EXT_Temp As New SiteDouble        'For TSENSOR, ADT7240 Reading
Global TheAlarm As New CLS_AlarmMonitor

'=================For DVS Module=============================================================
Public Type STR_DVS_ITEM
    DVS_FLAG_RD As New SiteLong
    DVS_RESULT_RD As New SiteLong
    DVS_FLAG_WR As New SiteLong
    DVS_RESULT_WR As New SiteLong
    dvs_allsite_execute As Boolean
End Type

Public Type STR_HISI_CHAR_SETUP_PAT
    SetupPatSetName As String
    DC_Category As String
    AC_Selector As String
    DC_Selector As String
    Level_Sheet As String
    AC_Category As String
    Timing_Sheet As String
End Type

Public Type STR_HISI_DVS_PAT
    BlockName As String
    StressPattern As String
    OriginalLevel As String
    StressPinGroup() As String
    StressTimes() As String
    LoopTimes As Long
End Type

Public Type HISI_DVS_FLOW_CTRL
    DvsExecuteFlag As New SiteLong
    BiningBeforeDvsFlag As New SiteLong
End Type

Public Enum BinActionType
    Initial = 0
    Update = 1
    Judgment = 2
End Enum

Public DVS_EFUSE_ITEM As STR_DVS_ITEM
Public aSTR_HISI_DVS_PAT(30) As STR_HISI_DVS_PAT
Public TestInstanceName As New SiteVariant
Public DisableInitialFlag As Boolean
Public ProgramInitial As Boolean
Public DVS_FLOW_CTRL_ITEM As HISI_DVS_FLOW_CTRL
Public aSTR_HISI_CHAR_PAT(12) As STR_HISI_CHAR_SETUP_PAT

'GlobalVaribles used in HISEE_OSC
Public glob_Pins_CSL As New CSL
Public glob_Freq_PLD As New PinListData
Public glob_MeasVolt_PLD As New PinListData
Public i_CurVtValue_dbl As Double
Public glob_CurVtValue_dbl As Double

Public GlobalVariable_Chuck_Temp As Double
'Glabal Variable for SHM/Vmin 32K
Public Set_32K_FLAG As Long

Public ActivateSheetFlag As Boolean
