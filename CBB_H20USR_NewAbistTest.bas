Attribute VB_Name = "CBB_H20USR_NewAbistTest"
Option Explicit
'''*********************************************************************************
''all FW Table1
Type Table1_H20USR_ALLFW
    Slice(1 To 10) As String
    NodeNUM(1 To 10) As Long
    NodeName(1 To 10) As String
    Address(1 To 10) As String
    Size(1 To 10) As String
    Spec_Low(1 To 10) As Double
    Spec_High(1 To 10) As Double
    TestRST(1 To 10) As New SiteDouble
End Type

Type H20USRALLFWInfo
    Public_Information As Table1_H20USR_ALLFW
End Type

Global H20USRALLFW As H20USRALLFWInfo
'''*********************************************************************************

'''*********************************************************************************
Public Function Initialize_H20USRNewAbistTest() As Long
    On Error GoTo errHandler

    Dim mSheet As Worksheet
    Dim i As Long, j As Long, Idx As Long
    
    ''*************************************************************************************************************************
    Set mSheet = ThisWorkbook.Worksheets("H20USR_ALLFW_Table1")
    mSheet.Activate
    With H20USRALLFW
        With .Public_Information
            For i = 1 To 10
                .Slice(i) = mSheet.Cells(i + 1, 1)
                .NodeNUM(i) = mSheet.Cells(i + 1, 2)
                .NodeName(i) = mSheet.Cells(i + 1, 3)
                .Address(i) = mSheet.Cells(i + 1, 4)
                .Size(i) = mSheet.Cells(i + 1, 5)
                .Spec_High(i) = CDbl(mSheet.Cells(i + 1, 6))
                .Spec_Low(i) = CDbl(mSheet.Cells(i + 1, 7))
                .TestRST(i) = CDbl(mSheet.Cells(i + 1, 8))
            Next i
        End With
    End With

    ''*************************************************************************************************************************
    
    Set mSheet = ThisWorkbook.Worksheets("H20USR_CSTest_Table2")
    mSheet.Activate
    With H20USRCSTest
        With .Fixed_Node_Spec
            For i = 1 To 45
                .Slice(i) = mSheet.Cells(i + 1, 1)
                .NodeNUM(i) = mSheet.Cells(i + 1, 2)
                .NodeName(i) = mSheet.Cells(i + 1, 3)
                .Address(i) = mSheet.Cells(i + 1, 4)
                .Size(i) = mSheet.Cells(i + 1, 5)
                .Spec_High(i) = CDbl(mSheet.Cells(i + 1, 6))
                .Spec_Low(i) = CDbl(mSheet.Cells(i + 1, 7))
                .TestRST(i) = CDbl(mSheet.Cells(i + 1, 8))
                .High_Limit_Address(i) = CLng(mSheet.Cells(i + 1, 9))
                .Low_Limit_Address(i) = CLng(mSheet.Cells(i + 1, 10))
                .Written_Rate(i) = CLng(mSheet.Cells(i + 1, 11))
            Next i
        End With
    End With
    '************************************************************************************************************************

    Set mSheet = ThisWorkbook.Worksheets("H20USR_RxLatchOffset_Table2")
    mSheet.Activate
    With H20USRRxLatchOffset
        With .Fixed_Node_Spec
            For i = 1 To 576
                .Slice(i) = mSheet.Cells(i + 1, 1)
                .NodeNUM(i) = mSheet.Cells(i + 1, 2)
                .NodeName(i) = mSheet.Cells(i + 1, 3)
                .Address(i) = mSheet.Cells(i + 1, 4)
                .Size(i) = mSheet.Cells(i + 1, 5)
                .Spec_High(i) = CDbl(mSheet.Cells(i + 1, 6))
                .Spec_Low(i) = CDbl(mSheet.Cells(i + 1, 7))
                .TestRST(i) = CDbl(mSheet.Cells(i + 1, 8))
            Next i
            For i = 1 To 64
                .High_Limit_Address(i) = CLng(mSheet.Cells(i + 1, 9))
                .Low_Limit_Address(i) = CLng(mSheet.Cells(i + 1, 10))
                .Written_Rate(i) = CLng(mSheet.Cells(i + 1, 11))
            Next i
        End With
    End With
''    ''*************************************************************************************************************************
''
    Set mSheet = ThisWorkbook.Worksheets("H20USR_DigitalAndIdle_Table2")
    mSheet.Activate
    With H20USRDigitalAndIdle
        With .Fixed_Node_Spec
            For i = 1 To 396
                .Slice(i) = mSheet.Cells(i + 1, 1)
                .NodeNUM(i) = mSheet.Cells(i + 1, 2)
                .NodeName(i) = mSheet.Cells(i + 1, 3)
                .Address(i) = mSheet.Cells(i + 1, 4)
                .Size(i) = mSheet.Cells(i + 1, 5)
                .Spec_High(i) = CDbl(mSheet.Cells(i + 1, 6))
                .Spec_Low(i) = CDbl(mSheet.Cells(i + 1, 7))
                .TestRST(i) = CDbl(mSheet.Cells(i + 1, 8))
            Next i
            For i = 1 To 396
                .High_Limit_Address(i) = CLng(mSheet.Cells(i + 1, 9))
                .Low_Limit_Address(i) = CLng(mSheet.Cells(i + 1, 10))
                .Written_Rate(i) = CLng(mSheet.Cells(i + 1, 11))
            Next i
        End With
    End With
 
''    *************************************************************************************************************************

    Set mSheet = ThisWorkbook.Worksheets("H20USR_IntDynamicSwitch_Table2")
    mSheet.Activate
    With H20USRIntDynamicSwitch
        With .Fixed_Node_Spec
            For i = 1 To 342
                .Slice(i) = mSheet.Cells(i + 1, 1)
                .NodeNUM(i) = mSheet.Cells(i + 1, 2)
                .NodeName(i) = mSheet.Cells(i + 1, 3)
                .Address(i) = mSheet.Cells(i + 1, 4)
                .Size(i) = mSheet.Cells(i + 1, 5)
                .Spec_High(i) = CDbl(mSheet.Cells(i + 1, 6))
                .Spec_Low(i) = CDbl(mSheet.Cells(i + 1, 7))
                .TestRST(i) = CDbl(mSheet.Cells(i + 1, 8))
            Next
            For i = 1 To 38
                .High_Limit_Address(i) = CLng(mSheet.Cells(i + 1, 9))
                .Low_Limit_Address(i) = CLng(mSheet.Cells(i + 1, 10))
                .Written_Rate(i) = CLng(mSheet.Cells(i + 1, 11))
            Next i
        End With
    End With

''    *************************************************************************************************************************

    Set mSheet = ThisWorkbook.Worksheets("H20USR_ExtDynamicSwitch_Table2")
    mSheet.Activate
    With H20USRExtDynamicSwitch
        With .Fixed_Node_Spec
            For i = 1 To 342
                .Slice(i) = mSheet.Cells(i + 1, 1)
                .NodeNUM(i) = mSheet.Cells(i + 1, 2)
                .NodeName(i) = mSheet.Cells(i + 1, 3)
                .Address(i) = mSheet.Cells(i + 1, 4)
                .Size(i) = mSheet.Cells(i + 1, 5)
                .Spec_High(i) = CDbl(mSheet.Cells(i + 1, 6))
                .Spec_Low(i) = CDbl(mSheet.Cells(i + 1, 7))
                .TestRST(i) = CDbl(mSheet.Cells(i + 1, 8))
            Next
            For i = 1 To 38
                .High_Limit_Address(i) = CLng(mSheet.Cells(i + 1, 9))
                .Low_Limit_Address(i) = CLng(mSheet.Cells(i + 1, 10))
                .Written_Rate(i) = CLng(mSheet.Cells(i + 1, 11))
            Next i
        End With
    End With
    ''*************************************************************************************************************************

    Set mSheet = ThisWorkbook.Worksheets("H20USR_DSATB_Table2")
    mSheet.Activate
    With H20USRDSATB
        With .Fixed_Node_Spec
            For i = 1 To 423
                .Slice(i) = mSheet.Cells(i + 1, 1)
                .NodeNUM(i) = mSheet.Cells(i + 1, 2)
                .NodeName(i) = mSheet.Cells(i + 1, 3)
                .Address(i) = mSheet.Cells(i + 1, 4)
                .Size(i) = mSheet.Cells(i + 1, 5)
                .Spec_High(i) = CDbl(mSheet.Cells(i + 1, 6))
                .Spec_Low(i) = CDbl(mSheet.Cells(i + 1, 7))
                .TestRST(i) = CDbl(mSheet.Cells(i + 1, 8))
            Next
            For i = 1 To 47
                .High_Limit_Address(i) = CLng(mSheet.Cells(i + 1, 9))
                .Low_Limit_Address(i) = CLng(mSheet.Cells(i + 1, 10))
                .Written_Rate(i) = CLng(mSheet.Cells(i + 1, 11))
            Next i
        End With
    End With


'*************************************************************************************************************************

    Set mSheet = ThisWorkbook.Worksheets("H20USR_MiscExtLpbk_Table2")
    mSheet.Activate
    With H20USRMiscExtLpbk
        With .Fixed_Node_Spec
            For i = 1 To 549
                .Slice(i) = mSheet.Cells(i + 1, 1)
                .NodeNUM(i) = mSheet.Cells(i + 1, 2)
                .NodeName(i) = mSheet.Cells(i + 1, 3)
                .Address(i) = mSheet.Cells(i + 1, 4)
                .Size(i) = mSheet.Cells(i + 1, 5)
                .Spec_High(i) = CDbl(mSheet.Cells(i + 1, 6))
                .Spec_Low(i) = CDbl(mSheet.Cells(i + 1, 7))
                .TestRST(i) = CDbl(mSheet.Cells(i + 1, 8))
            Next i
            For i = 1 To 61
                .High_Limit_Address(i) = CLng(mSheet.Cells(i + 1, 9))
                .Low_Limit_Address(i) = CLng(mSheet.Cells(i + 1, 10))
                .Written_Rate(i) = CLng(mSheet.Cells(i + 1, 11))
            Next i
        End With
    End With

'*************************************************************************************************************************

    Set mSheet = ThisWorkbook.Worksheets("H20USR_MiscIntLpbk_Table2")
    mSheet.Activate
    With H20USRMiscIntLpbk
        With .Fixed_Node_Spec
            For i = 1 To 468
                .Slice(i) = mSheet.Cells(i + 1, 1)
                .NodeNUM(i) = mSheet.Cells(i + 1, 2)
                .NodeName(i) = mSheet.Cells(i + 1, 3)
                .Address(i) = mSheet.Cells(i + 1, 4)
                .Size(i) = mSheet.Cells(i + 1, 5)
                .Spec_High(i) = CDbl(mSheet.Cells(i + 1, 6))
                .Spec_Low(i) = CDbl(mSheet.Cells(i + 1, 7))
                .TestRST(i) = CDbl(mSheet.Cells(i + 1, 8))
            Next i
            For i = 1 To 52
                .High_Limit_Address(i) = CLng(mSheet.Cells(i + 1, 9))
                .Low_Limit_Address(i) = CLng(mSheet.Cells(i + 1, 10))
                .Written_Rate(i) = CLng(mSheet.Cells(i + 1, 11))
            Next i
        End With
    End With
'*************************************************************************************************************************
    Set mSheet = ThisWorkbook.Worksheets("H20USR_ApTunningRange_Table2")
    mSheet.Activate
    With H20USRAdpllTunningRange
        With .Fixed_Node_Spec
            For i = 1 To 10
                .Slice(i) = mSheet.Cells(i + 1, 1)
                .NodeNUM(i) = mSheet.Cells(i + 1, 2)
                .NodeName(i) = mSheet.Cells(i + 1, 3)
                .Address(i) = mSheet.Cells(i + 1, 4)
                .Size(i) = mSheet.Cells(i + 1, 5)
                .Spec_High(i) = CDbl(mSheet.Cells(i + 1, 6))
                .Spec_Low(i) = CDbl(mSheet.Cells(i + 1, 7))
                .TestRST(i) = CDbl(mSheet.Cells(i + 1, 8))
            Next i
            For i = 1 To 10
                .High_Limit_Address(i) = CLng(mSheet.Cells(i + 1, 9))
                .Low_Limit_Address(i) = CLng(mSheet.Cells(i + 1, 10))
                .Written_Rate(i) = CLng(mSheet.Cells(i + 1, 11))
            Next i
        End With
    End With
'*************************************************************************************************************************
    Set mSheet = ThisWorkbook.Worksheets("H20USR_ApVcoLockRange_Table2")
    mSheet.Activate
    With H20USRAdpllVcoLockRange
        With .Fixed_Node_Spec
            For i = 1 To 26
                .Slice(i) = mSheet.Cells(i + 1, 1)
                .NodeNUM(i) = mSheet.Cells(i + 1, 2)
                .NodeName(i) = mSheet.Cells(i + 1, 3)
                .Address(i) = mSheet.Cells(i + 1, 4)
                .Size(i) = mSheet.Cells(i + 1, 5)
                .Spec_High(i) = CDbl(mSheet.Cells(i + 1, 6))
                .Spec_Low(i) = CDbl(mSheet.Cells(i + 1, 7))
                .TestRST(i) = CDbl(mSheet.Cells(i + 1, 8))
            Next i
            For i = 1 To 26
                .High_Limit_Address(i) = CLng(mSheet.Cells(i + 1, 9))
                .Low_Limit_Address(i) = CLng(mSheet.Cells(i + 1, 10))
                .Written_Rate(i) = CLng(mSheet.Cells(i + 1, 11))
            Next i
        End With
    End With
   
    Exit Function
errHandler:
    Call TheExec.ErrorLogMessage("Test " & TL_C_ERRORSTR & ", Instance: " & TheExec.DataManager.InstanceName)
    If AbortTest Then Exit Function Else Resume Next
    
End Function
    



