Attribute VB_Name = "CBB_Hilink_Declaration"
Option Explicit

Global Const CBB_Code_Gen_Tool_VER As String = "2.0.0.0"


''''    TYPE START  >>>>

Type DUTInfo
    H20USR As H20USR_IP      'H20USR is only applicable if the DUT has H20USR IP(s) in it.
End Type


''''    <<<<    TYPE END

Global TheDUT As DUTInfo
Global HiLink_ExecIP As New Cls_HiLink_Exec_IP_Module
'Global Site As Variant
Global CSRDSP As New DSPWave
Global LanCnt As Long
Global PAM_Name As String
Global glb_DebugLog_ON As Boolean
Global Ttemp As Long

Global TheSwitch As Switch
Global Term(0 To 3) As Term_Info
























