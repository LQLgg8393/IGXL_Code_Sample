Attribute VB_Name = "DSP_VBT_DC"
Option Explicit

' This module should be used only for DSP Procedure code.  Functions in this
' module will be available to be called to perform DSP in all DSP modes.
' Additional modules may be added as needed (all starting with "DSP_").
'
' The required signature for a DSP Procedure is:
'
' Public Function FuncName(<arglist>) as Long
'   where <arglist> is any list of arguments supported by DSP code.
'
' See online help for supported types and other restrictions.

Function USB_Calc_R( _
         ByVal in_V1_dbl As Double, _
         ByVal in_V2_dbl As Double, _
         ByVal in_I1_dbl As Double, _
         ByVal in_I2_dbl As Double, _
         ByRef out_R_dbl As Double _
       ) As Long

    On Error Resume Next    ' DSP does not support <On Error Goto ...>, but you can use <On Error Resume Next>
    If (in_I2_dbl - in_I1_dbl) = 0 Then
        out_R_dbl = 999999.999999
    Else
        out_R_dbl = (in_V2_dbl - in_V1_dbl) / (in_I2_dbl - in_I1_dbl)
    End If

End Function
