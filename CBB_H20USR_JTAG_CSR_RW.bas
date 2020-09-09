Attribute VB_Name = "CBB_H20USR_JTAG_CSR_RW"

Option Explicit

Private mPLD As New PinListData
Private mDSPwave As New DSPWave
Private Site As Variant

Public Function JTAG_Write_H20USR(ByVal mADDR As Long, ByVal mdata As Variant) As Long
    On Error GoTo errHandler
'''**************************************************************************
    Dim FrameName As String
    FrameName = "AHB_Write_SDS"
'''**************************************************************************
    mADDR = (mADDR And &HFFFE&) \ 2

    With TheHdw.Protocol.Ports("JTAG_Pins")
        With .NWire.Frames(FrameName)
            .Fields("Addr").Value = mADDR
            .Fields("Data").Value = mdata
            Call .Execute
        End With
        .IdleWait
    End With
    
    Call LogJTAG_AHB_H20USR("JTAG_Write_H20USR", mADDR * 2, mdata)
    Exit Function
errHandler:
    Call TheExec.ErrorLogMessage("Test " & TL_C_ERRORSTR & ", Instance: " & TheExec.DataManager.InstanceName)
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function JTAG_Read_H20USR(ByVal mADDR As Long, ByRef mdata As SiteLong) As Long
    On Error GoTo errHandler

    Dim ExicutionType As tlNWireExecutionType

    Set mdata = New SiteLong
    
'''**************************************************************************

    Dim FrameName As String
    
    FrameName = "AHB_Read_SDS"
    
'''**************************************************************************

    If TheHdw.Protocol.Ports("JTAG_Pins").ModuleFiles.IsLoading Then
        ExicutionType = tlNWireExecutionType_PushToStack
    Else
        With TheHdw.Protocol.Ports("JTAG_Pins").NWire.CMEM
            .MoveMode = glb_H20USR_tlNWireCMEMMoveMode
            mPLD = .DSPWave
            mDSPwave = mPLD.Pins("JTAG_Pins")
        End With
        ExicutionType = tlNWireExecutionType_CaptureInCMEM
    End If

    mADDR = (mADDR And &HFFFE) \ 2

    With TheHdw.Protocol.Ports("JTAG_Pins")

        With .NWire.Frames(FrameName)
            .Fields("Addr").Value = mADDR
            .Execute tlNWireExecutionType_CaptureInCMEM
        End With
        Call .IdleWait
    End With

    If TheHdw.Protocol.Ports("JTAG_Pins").ModuleFiles.IsLoading Then     ' skip checking pass/fail if in recording PA Module
        JTAG_Read_H20USR = 0
        mdata = 0
        Exit Function
    End If

    For Each Site In TheExec.Sites
        If TheHdw.Protocol.Ports("JTAG_Pins").NWire.CMEM.Transactions.count = 0 Then
            ' no reading, sth wrong with HW PA engines
            mdata = -1
            JTAG_Read_H20USR = -1
            TheExec.AddOutput "Site " + CStr(Site) + " <JTAG_Read_H20USR> Error! Instance <" + TheExec.DataManager.InstanceName + "> has abnormal PA reading, " + _
                              "check TN:" + CStr(TheExec.Datalog.LastTestNumLogged + 1), vbRed
            TheExec.AddOutput "AHB Address: &H" + Hex(mADDR * 2), vbRed
        ElseIf TheHdw.Protocol.Ports("JTAG_Pins").Passed = False Then
            ' if trapped here, it means JTAG reading is invalid
            mdata = -1
            JTAG_Read_H20USR = -1
            TheExec.AddOutput "Site " + CStr(Site) + " <JTAG_Read_H20USR> Error! Instance <" + TheExec.DataManager.InstanceName + "> has invalid PA reading, " + _
                              "check TN:" + CStr(TheExec.Datalog.LastTestNumLogged + 1), vbRed
            TheExec.AddOutput "AHB Address: &H" + Hex(mADDR * 2), vbRed
        Else
            mdata = mDSPwave.Element(0)
            JTAG_Read_H20USR = mdata()
        End If
    Next Site


    ' Log the Read transaction
    Call LogJTAG_AHB_H20USR("JTAG_Read_H20USR", mADDR * 2, mdata)

    Exit Function
errHandler:
    Call TheExec.ErrorLogMessage("Test " & TL_C_ERRORSTR & ", Instance: " & TheExec.DataManager.InstanceName)
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function JTAG_XRead_H20USR(ByVal mADDR As Long, ByRef mdata As SiteLong) As Long
    On Error GoTo errHandler

    If Not TheHdw.Protocol.Ports("JTAG_Pins").ModuleFiles.IsLoading Then
        JTAG_XRead_H20USR = JTAG_Read_H20USR(mADDR, mdata)
        Exit Function
    End If

    mADDR = (mADDR And &HFFFE) \ 2
    Set mdata = New SiteLong

'''**************************************************************************

    Dim FrameName As String

    FrameName = "AHB_xRead_SDS"

'''**************************************************************************

    With TheHdw.Protocol.Ports("JTAG_Pins")

        With .NWire.Frames(FrameName)
            .Fields("Addr").Value = mADDR
            .Execute tlNWireExecutionType_CaptureInCMEM
        End With
        Call .IdleWait
    End With

    JTAG_XRead_H20USR = 0
    mdata = 0

    Exit Function
errHandler:
    Call TheExec.ErrorLogMessage("Test " & TL_C_ERRORSTR & ", Instance: " & TheExec.DataManager.InstanceName)
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function JTAG_Read_Match_H20USR(ByVal mADDR As Long, ByVal mdata As Long, ByVal Mask As Long, _
                                     Optional MaxReadUntilCount As Long = 1000) As SiteBoolean
    On Error GoTo errHandler

    Dim Matched As New SiteBoolean
    Dim ExicutionType As tlNWireExecutionType
    Dim MASK_Name As String
    
'''**************************************************************************

    Dim FrameName As String

    FrameName = "AHB_Read_SDS"
    
'''**************************************************************************
    MASK_Name = "H" + Hex(Mask)

    If TheHdw.Protocol.Ports("JTAG_Pins").ModuleFiles.IsLoading Then
        ExicutionType = tlNWireExecutionType_UntilMatch
    Else
        With TheHdw.Protocol.Ports("JTAG_Pins").NWire.CMEM
            .MoveMode = glb_H20USR_tlNWireCMEMMoveMode
            mPLD = .DSPWave
            mDSPwave = mPLD.Pins("JTAG_Pins")
        End With

        TheHdw.Protocol.Ports("JTAG_Pins").NWire.Frames(FrameName).Fields("Data").Masks(MASK_Name).Value = Not Mask
        ExicutionType = tlNWireExecutionType_UntilMatchCaptureAll
    End If

    mADDR = (mADDR And &HFFFE) \ 2

    With TheHdw.Protocol.Ports("JTAG_Pins")

        .NWire.MaxReadUntilCount = MaxReadUntilCount
        With .NWire.Frames(FrameName)
            .Fields("Addr").Value = mADDR
            .Fields("Data").Value = mdata
            .Execute ExicutionType, MASK_Name
        End With
        Call .IdleWait
    End With

    Set JTAG_Read_Match_H20USR = New SiteBoolean

    If TheHdw.Protocol.Ports("JTAG_Pins").ModuleFiles.IsLoading Then    ' skip checking pass/fail if in recording PA Module
        JTAG_Read_Match_H20USR = True
        Exit Function
    End If

    For Each Site In TheExec.Sites
        JTAG_Read_Match_H20USR = TheHdw.Protocol.LastTest.Passed
    Next Site

    Exit Function
errHandler:
    Call TheExec.ErrorLogMessage("Test " & TL_C_ERRORSTR & ", Instance: " & TheExec.DataManager.InstanceName)
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function JTAG_Read_Match_H20USR_Mask8000(ByVal mADDR As Long, ByVal mdata As Long, _
                                     Optional MaxReadUntilCount As Long = 1000) As SiteBoolean
    On Error GoTo errHandler

    Dim Matched As New SiteBoolean
    Dim ExicutionType As tlNWireExecutionType
    'Dim MASK_Name As String
    
'''**************************************************************************
'''Just for TC Program
    Dim FrameName As String

    FrameName = "AHB_Read_SDS_H8000"
    
'''**************************************************************************
    'MASK_Name = "H" + Hex(Mask)

    If TheHdw.Protocol.Ports("JTAG_Pins").ModuleFiles.IsLoading Then
        ExicutionType = tlNWireExecutionType_UntilMatch
    Else
        With TheHdw.Protocol.Ports("JTAG_Pins").NWire.CMEM
            .MoveMode = glb_H20USR_tlNWireCMEMMoveMode
            mPLD = .DSPWave
            mDSPwave = mPLD.Pins("JTAG_Pins")
        End With

        'TheHdw.Protocol.Ports("JTAG_Pins").NWire.Frames(FrameName).Fields("Data").Masks(MASK_Name).Value = Not Mask
        ExicutionType = tlNWireExecutionType_UntilMatchCaptureAll
    End If

    mADDR = (mADDR And &HFFFE) \ 2

    With TheHdw.Protocol.Ports("JTAG_Pins")

        .NWire.MaxReadUntilCount = MaxReadUntilCount
        With .NWire.Frames(FrameName)
            .Fields("Addr").Value = mADDR
            .Fields("Data").Value = mdata
            .Execute ExicutionType ', MASK_Name
        End With
        Call .IdleWait
    End With

    Set JTAG_Read_Match_H20USR_Mask8000 = New SiteBoolean

    If TheHdw.Protocol.Ports("JTAG_Pins").ModuleFiles.IsLoading Then    ' skip checking pass/fail if in recording PA Module
        JTAG_Read_Match_H20USR_Mask8000 = True
        Exit Function
    End If

    For Each Site In TheExec.Sites
        JTAG_Read_Match_H20USR_Mask8000 = TheHdw.Protocol.LastTest.Passed
    Next Site

    Exit Function
errHandler:
    Call TheExec.ErrorLogMessage("Test " & TL_C_ERRORSTR & ", Instance: " & TheExec.DataManager.InstanceName)
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function JTAG_Verify_H20USR(ByVal mADDR As Long, ByVal mdata As Variant) As SiteBoolean
    On Error GoTo errHandler

'''**************************************************************************

    Dim FrameName As String
    
    FrameName = "AHB_Read_SDS"
'''**************************************************************************

    mADDR = (mADDR And &HFFFE) \ 2
    
    With TheHdw.Protocol.Ports("JTAG_Pins")

        With .NWire.Frames(FrameName)
            .Fields("Addr").Value = mADDR
            .Fields("Data").Value = mdata
            .Execute tlNWireExecutionType_Default
        End With
        Call .IdleWait

    End With

    Exit Function
errHandler:
    Call TheExec.ErrorLogMessage("Test " & TL_C_ERRORSTR & ", Instance: " & TheExec.DataManager.InstanceName)
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function JTAG_Compare_Mask8000_H20USR(ByVal mADDR As Long, ByVal CompareValue As Long) As SiteBoolean
    On Error GoTo errHandler
      
'''**************************************************************************

    Dim FrameName As String
    
    FrameName = "AHB_Read_SDS_H8000"

'''**************************************************************************
    
    mADDR = (mADDR And &HFFFE) \ 2
    For Each Site In TheExec.Sites.Active
        With TheHdw.Protocol.Ports("JTAG_Pins")
            With .NWire.Frames(FrameName)
                .Fields("Addr").Value = mADDR
                .Fields("Data").Value = CompareValue
                .Execute tlNWireExecutionType_Default
            End With
            Call .IdleWait
        End With
    Next Site
    Exit Function
errHandler:
    Call TheExec.ErrorLogMessage("Test " & TL_C_ERRORSTR & ", Instance: " & TheExec.DataManager.InstanceName)
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function LogJTAG_AHB_H20USR(ByVal RW As String, mADDR As Long, mdata As Variant) As Long
    On Error GoTo errHandler

    If glb_DebugLog_ON Then
        For Each Site In TheExec.Sites
            TheExec.Datalog.WriteComment "Site" + CStr(Site) + ": " + RW + "(&H" + Hex(mADDR) + "& , &H" + Hex(mdata) + ")"
        Next Site
    End If

    Exit Function
errHandler:
    Call TheExec.ErrorLogMessage("Test " & TL_C_ERRORSTR & ", Instance: " & TheExec.DataManager.InstanceName)
    If AbortTest Then Exit Function Else Resume Next
End Function

