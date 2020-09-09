Attribute VB_Name = "VBU_T_Utils"
Option Explicit

Public TheDPS As New CLS_DPS_Instrument

Private m_UsedVBTModules_DCT As Object
Private m_UsedProcedures_DCT As Object
Private m_TestFlowRunLog_DCT As Object

Public Sub LogCalledFunctions( _
       ByVal in_VBModuleName_str As String, _
       ByVal in_FunctionName_str As String, _
       Optional ByVal in_ParameterVal_v As Variant = "")

    Dim i_Log_str As String
    Dim i_ParameterVal_str As String

    in_FunctionName_str = in_VBModuleName_str + "::" + in_FunctionName_str

    If in_ParameterVal_v <> "" Then
        i_ParameterVal_str = CStr(in_ParameterVal_v)
        i_Log_str = in_FunctionName_str + "::" + i_ParameterVal_str
    Else
        i_Log_str = in_FunctionName_str
    End If

    If m_UsedVBTModules_DCT Is Nothing Then
        Set m_UsedVBTModules_DCT = CreateObject("Scripting.Dictionary")
    ElseIf Not m_UsedVBTModules_DCT.Exists(in_VBModuleName_str) Then
        Call m_UsedVBTModules_DCT.Add(in_VBModuleName_str, 0)
    Else
        ' do nothing
    End If

    If m_UsedProcedures_DCT Is Nothing Then
        Set m_UsedProcedures_DCT = CreateObject("Scripting.Dictionary")
    ElseIf Not m_UsedProcedures_DCT.Exists(in_FunctionName_str) Then
        Call m_UsedProcedures_DCT.Add(in_FunctionName_str, 0)
        m_UsedVBTModules_DCT.Item(in_VBModuleName_str) = _
        m_UsedVBTModules_DCT.Item(in_VBModuleName_str) + 1
    Else
        m_UsedProcedures_DCT.Item(in_FunctionName_str) = _
        m_UsedProcedures_DCT.Item(in_FunctionName_str) + 1
    End If

    If m_TestFlowRunLog_DCT Is Nothing Then
        Set m_TestFlowRunLog_DCT = CreateObject("Scripting.Dictionary")
    ElseIf Not m_TestFlowRunLog_DCT.Exists(i_Log_str) Then
        Call m_TestFlowRunLog_DCT.Add(i_Log_str, 1)
    Else
        m_TestFlowRunLog_DCT.Item(i_Log_str) = _
        m_TestFlowRunLog_DCT.Item(i_Log_str) + 1
    End If

End Sub

Sub PrintLog2File( _
    in_FileName_str As String)

    Dim FN As Long
    Dim i_key_str As Variant

    If in_FileName_str = "" Then in_FileName_str = "c:\VBT_Called_Log.txt"
    FN = FreeFile
    Open in_FileName_str For Output As #FN

    Print #FN, vbNewLine + "Modules:" + vbNewLine
    For Each i_key_str In m_UsedVBTModules_DCT.keys
        Print #FN, m_UsedVBTModules_DCT.Item(i_key_str) + vbTab + i_key_str + vbNewLine
    Next
    Print #FN, String(4, vbNewLine)

    Print #FN, vbNewLine + "Procedures:" + vbNewLine
    For Each i_key_str In m_UsedProcedures_DCT.keys
        Print #FN, m_UsedProcedures_DCT.Item(i_key_str) + vbTab + i_key_str + vbNewLine
    Next
    Print #FN, String(4, vbNewLine)

    Print #FN, vbNewLine + "Function and Arguments:" + vbNewLine
    For Each i_key_str In m_TestFlowRunLog_DCT.keys
        Print #FN, m_TestFlowRunLog_DCT.Item(i_key_str) + vbTab + i_key_str + vbNewLine
    Next

    Close #FN

End Sub




Public Function UT() As Long
''''    Dim i_tmpIdx_DSP As New DSPWave
''''
''''    i_tmpIdx_DSP.CreateConstant 1, 10 * 36, DspLong
''''    i_tmpIdx_DSP.Select(36, 36).Replace (2)
''''    i_tmpIdx_DSP.Element(0) = 0
''''    i_tmpIdx_DSP = i_tmpIdx_DSP.IntegrateElements
''''    i_tmpIdx_DSP.Plot "i_tmpIdx_DSP"

End Function


Public Function Write_to_file( _
       ByVal FileName_str As String, _
       ByVal writeString As String) As Long

    Dim FN As Long
    On Error GoTo exit_function

    If FileName_str = "" Then FileName_str = "c:\Temp_File.txt"
    FN = FreeFile
    Open FileName_str For Output As #FN

    Write #FN, writeString

exit_function:
    Close #FN
End Function

Public Function Read_from_file( _
       ByVal FileName_str As String, _
       ByRef readString As String) As Long

    Dim FN As Long
    On Error GoTo exit_function

    If FileName_str = "" Then FileName_str = "c:\Temp_File.txt"
    FN = FreeFile
    Open FileName_str For Input As #FN

    Input #FN, readString

exit_function:
    Close #FN
End Function

