Attribute VB_Name = "CBB_Hi5345_I2C_RW"
Option Explicit

Private mPLD As New PinListData
Private mDSPwave As New DSPWave
Private Site As Variant

Public Function I2C_Read_H20USR(ByVal mData_ADDR As Long, ByRef mdata As SiteLong) As Long
    On Error GoTo errHandler

    Dim ExicutionType As tlNWireExecutionType
    Set mdata = New SiteLong
    Dim FrameName As String
'''**************************************************************************
'''**************************************************************************
'''**************************************************************************
''''write page addr
    Dim mData_ADDR_H8bits As Long
    Dim mData_ADDR_L8bits As Long
    Dim mPage_ADDR As Long
    
    mData_ADDR_H8bits = Int(mData_ADDR / (2 ^ 8))
    mData_ADDR_L8bits = Int(mData_ADDR - mData_ADDR_H8bits * (2 ^ 8))
    mPage_ADDR = &H1

    If glb_current_page_addr <> mData_ADDR_H8bits Then
        Call I2C_Writepage_H20USR(mPage_ADDR, mData_ADDR_H8bits)
    End If
    
    glb_current_page_addr = mData_ADDR_H8bits
'''**************************************************************************
    FrameName = "Read"
    
    If TheHdw.Protocol.Ports("I2C_Pins").ModuleFiles.IsLoading Then
        ExicutionType = tlNWireExecutionType_PushToStack
    Else
        With TheHdw.Protocol.Ports("I2C_Pins").NWire.CMEM
            .MoveMode = glb_H20USR_tlNWireCMEMMoveMode
            mPLD = .DSPWave
            mDSPwave = mPLD.Pins("I2C_Pins")
        End With
        ExicutionType = tlNWireExecutionType_CaptureInCMEM
    End If

    With TheHdw.Protocol.Ports("I2C_Pins")

        With .NWire.Frames(FrameName)
            .Fields("ADDR").Value = mData_ADDR_L8bits
            .Execute tlNWireExecutionType_CaptureInCMEM
        End With
        Call .IdleWait
    End With

    If TheHdw.Protocol.Ports("I2C_Pins").ModuleFiles.IsLoading Then     ' skip checking pass/fail if in recording PA Module
        I2C_Read_H20USR = 0
        mdata = 0
        Exit Function
    End If

    For Each Site In TheExec.Sites
        If TheHdw.Protocol.Ports("I2C_Pins").NWire.CMEM.Transactions.count = 0 Then
            ' no reading, sth wrong with HW PA engines
            mdata = -1
            I2C_Read_H20USR = -1
            TheExec.AddOutput "Site " + CStr(Site) + " <I2C_Read_H20USR> Error! Instance <" + TheExec.DataManager.InstanceName + "> has abnormal PA reading, " + _
                              "check TN:" + CStr(TheExec.Datalog.LastTestNumLogged + 1), vbRed
            TheExec.AddOutput "AHB Address: &H" + Hex(mData_ADDR_L8bits), vbRed
        ElseIf TheHdw.Protocol.Ports("I2C_Pins").Passed = False Then
            ' if trapped here, it means JTAG reading is invalid
            mdata = -1
            I2C_Read_H20USR = -1
            TheExec.AddOutput "Site " + CStr(Site) + " <I2C_Read_H20USR> Error! Instance <" + TheExec.DataManager.InstanceName + "> has invalid PA reading, " + _
                              "check TN:" + CStr(TheExec.Datalog.LastTestNumLogged + 1), vbRed
            TheExec.AddOutput "AHB Address: &H" + Hex(mData_ADDR_L8bits), vbRed
        Else
            mdata = mDSPwave.Element(0)
            I2C_Read_H20USR = mdata()
        End If
    Next Site


    Exit Function
errHandler:
    Call TheExec.ErrorLogMessage("Test " & TL_C_ERRORSTR & ", Instance: " & TheExec.DataManager.InstanceName)
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function I2C_Write_H20USR(ByVal mData_ADDR As Long, mdata As Variant) As Long
    On Error GoTo errHandler
'''**************************************************************************
    Dim mData_ADDR_H8bits As Long
    Dim mData_ADDR_L8bits As Long
    Dim mPage_ADDR As Long
    
    mData_ADDR_H8bits = Int(mData_ADDR / (2 ^ 8))
    mData_ADDR_L8bits = Int(mData_ADDR - mData_ADDR_H8bits * (2 ^ 8))
    mPage_ADDR = &H1
'''**************************************************************************

    If glb_current_page_addr <> mData_ADDR_H8bits Then
        Call I2C_Writepage_H20USR(mPage_ADDR, mData_ADDR_H8bits)
    End If
    Call I2C_Writedata_H20USR(mData_ADDR_L8bits, mdata)
    
    glb_current_page_addr = mData_ADDR_H8bits
    
    Exit Function
errHandler:
    Call TheExec.ErrorLogMessage("Test " & TL_C_ERRORSTR & ", Instance: " & TheExec.DataManager.InstanceName)
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function I2C_Writepage_H20USR(ByVal mPageADDR As Long, mPageData As Variant) As Long
    On Error GoTo errHandler
'''**************************************************************************
    Dim FrameName As String
    FrameName = "Write"
'''**************************************************************************

    With TheHdw.Protocol.Ports("I2C_Pins")
        With .NWire.Frames(FrameName)
            .Fields("ADDR").Value = mPageADDR
            .Fields("Data").Value = mPageData
            Call .Execute
        End With
        .IdleWait
    End With
    
    Exit Function
errHandler:
    Call TheExec.ErrorLogMessage("Test " & TL_C_ERRORSTR & ", Instance: " & TheExec.DataManager.InstanceName)
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function I2C_Writedata_H20USR(ByVal mDataADDR As Long, mDataData As Variant) As Long
    On Error GoTo errHandler
'''**************************************************************************
    Dim FrameName As String
    FrameName = "Write"
'''**************************************************************************

    With TheHdw.Protocol.Ports("I2C_Pins")
        With .NWire.Frames(FrameName)
            .Fields("ADDR").Value = mDataADDR
            .Fields("Data").Value = mDataData
            Call .Execute
        End With
        .IdleWait
    End With

    Exit Function
errHandler:
    Call TheExec.ErrorLogMessage("Test " & TL_C_ERRORSTR & ", Instance: " & TheExec.DataManager.InstanceName)
    If AbortTest Then Exit Function Else Resume Next
End Function



