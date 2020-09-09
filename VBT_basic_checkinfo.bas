Attribute VB_Name = "VBT_basic_checkinfo"
Option Explicit

' ### Teradyne HISI UFLEX Basic Templates V13.00 ###

'Version History
'Version 1.1 This Verison Release 09072015 Eric Peng
'Version 2.1 This Verison Release 09092015 Eric Peng
'Version 3.1 This Verison Release 09152015 Eric Peng
'Version 3.2 This Verison Release 09242015 Eric Peng
'Version 3.3
'This Verison Correct bug when load the log to excel and Split TimeSet error 2015/11/4
'Correct the JobName Lowcase and Upper case compare false trouble 2015/12/09
'Version 3.4
'This Verison Fixed VIHL,VOHL pins are not correct error 2016/06/30

'*****************Manual Start***************************
'Step1:Import the VBT_HISI_CheckInfo_Vxx.bas to VBT module
'Step2:
'>>You can set interpose function in test instance vb codes that you wants the information - Suggested
'    Add Set Interpose fuction before pattern run
'    PrintInfo = True
'    If PrintInfo = True Then
'       Call tl_SetInterpose(TL_C_POSTPATF, "CheckInfoCollect", "")
'    End If
'>>You can set interpose function for the flow.
'>>While this case you should confirm there is not any interpose function clear template or command in the flow
'    Add Set Interpose fuction in Function OnProgramStarted()
'    PrintInfo = True
'    If PrintInfo = True Then
'        Call TheExec.Flow.SetInterpose(TL_C_POSTPATF, "CheckInfoCollect", "")
'    End If

'Step3: Call the print function in function OnProgramEnded
'    Call CheckInfoPrint
'Step4:Run the test program
'Step5:The CheckInfoLog generated in program folder
'*******************Manul End******************************


'Public PrintInfo As Boolean ' Need to Check if the global variable is defined

Dim LogBuffer As String
Dim LogTestNum As String * 10
Dim TestInstName As String * 60
Dim PatName As String * 100
Dim CycleCount As String * 12
Dim TSetName As String * 80
Dim TSetEdge As String * 10
Dim PeriodPare As String * 30

Dim PwrPinName As String * 20
Dim DigPinName As String * 20
Dim DigPinIgnore As String * 20
Dim DigPinValue As String * 20

Dim tempstr As String
Dim InstanceNum As Long
Dim InstanceIndex As Variant
Dim InstanceArray() As String
Dim RepContrl As Long
Dim TestNumIndex As Variant

Dim LogTestInfoTestNum As New PinListData
Dim LogTestInfoVIHLPins As New PinListData


Public Function CheckInfoPrint()

    Dim Header As String
    Dim PowerPinNames() As String
    Dim PinType As String
    Dim i As Variant

    If PrintInfo = True Then

        LogTestNum = "TestNum"
        TestInstName = "TestInstance"
        PatName = "PatternName"
        CycleCount = "CycleCount"
        TSetName = "TimeSetNameAndPeriod/Frequency"
        TSetEdge = "TSetEdge"
        PeriodPare = "PeriodParemeter"

        Header = LogTestNum & TestInstName & PatName & CycleCount

        FindPowerFromPinSheet PowerPinNames

        For i = LBound(PowerPinNames) To UBound(PowerPinNames)
            PinType = TheExec.DataManager.ChannelType(PowerPinNames(i))
            If Left(PinType, 4) = "DCVS" Then
                PwrPinName = PowerPinNames(i)
                Header = Header + PwrPinName
            ElseIf Left(PinType, 4) = "DCVI" Then
                PwrPinName = PowerPinNames(i)
                Header = Header + PwrPinName
            Else
                PwrPinName = PowerPinNames(i)
                Header = Header + PwrPinName
            End If
        Next i

        For i = 0 To LogTestInfoVIHLPins.Pins.count - 1

            DigPinName = LogTestInfoVIHLPins.Pins(i).Name + "_VIH"
            tempstr = tempstr + DigPinName
            DigPinName = LogTestInfoVIHLPins.Pins(i).Name + "_VIL"
            tempstr = tempstr + DigPinName
            DigPinName = LogTestInfoVIHLPins.Pins(i).Name + "_VOH"
            tempstr = tempstr + DigPinName
            DigPinName = LogTestInfoVIHLPins.Pins(i).Name + "_VOL"
            tempstr = tempstr + DigPinName


        Next i



        Header = Header + tempstr + TSetEdge + PeriodPare + TSetName + "ClkPinsAndClkPeriod/Frequency"
        LogBuffer = Header + vbNewLine + LogBuffer

        TheExec.Datalog.WriteComment "*********Checking Log Started**********"

        Close #1
        Open ".\CheckInfoLog.txt" For Output As #1
        ' Open ".\CheckInfoLog.txt" For Append As #1
        Print #1, LogBuffer
        Close #1

        TheExec.Datalog.WriteComment "*********Checking Log Ended**********"
        TheExec.Datalog.WriteComment "******PleaseCheckInTheProgramFolder******"


        LogBuffer = ""
        tempstr = ""
        Header = ""
        PeriodPare = ""


    End If
End Function



Public Function CheckInfoCollect(argc As Long, argv() As String) As Long


'Exit Function ' add by Eric



    Dim LastBurst As String
    Dim SheetName As String
    Dim ws As Worksheet

    Dim Period As Double
    Dim TimeSetName As String
    Dim TimeSetIndex As Variant
    Dim TimeSetList() As String

    Dim DCVSPins As String
    Dim DCVIPins As String

    Dim DCCat As String, DCSel As String, ACCat As String, ACSel As String
    Dim TS As String, ES As String, LevSh As String, Overlay As String
    Dim row As Long, pin As String

    Dim PowerPinNames() As String
    Dim PinType As String
    Dim i As Variant
    Dim Rep As Boolean
    Dim num As Long
    Dim TestNum() As Long
    Dim StrPare As String

    DigPinIgnore = ""

    TestInstName = TheExec.DataManager.InstanceName

    ' Test Number
    LogTestNum = TheExec.Sites(0).TestNumber

    'Vector Module

    Call TheHdw.Digital.TimeDomains("").Patgen.ReadLastStart(LastBurst, False, "")
    PatName = LastBurst

    'Cycle Count
    CycleCount = TheHdw.Digital.Patgen.CycleCount(tlCycleTypeAbsolute)    'thehdw.Digital.TimeDomains("").Patgen.CycleCount

    LogBuffer = LogBuffer + LogTestNum + TestInstName + PatName + CycleCount

    'Power Pins Search
    FindPowerFromPinSheet PowerPinNames

    For i = LBound(PowerPinNames) To UBound(PowerPinNames)
        PinType = TheExec.DataManager.ChannelType(PowerPinNames(i))
        If Left(PinType, 4) = "DCVS" Then
            PwrPinName = Format(TheHdw.DCVS.Pins(PowerPinNames(i)).voltage.Main, "0.####") + "V"
            LogBuffer = LogBuffer + PwrPinName
        ElseIf Left(PinType, 4) = "DCVI" Then
            PwrPinName = Format(TheHdw.DCVI.Pins(PowerPinNames(i)).voltage, "0.####") + "V"
            LogBuffer = LogBuffer + PwrPinName
        Else
            PwrPinName = "UnknowPin"
            LogBuffer = LogBuffer + PwrPinName
        End If

    Next i

    '    '  Read the instance context just to find the level sheet
    Call TheExec.DataManager.GetInstanceContext(DCCat, DCSel, ACCat, _
                                                ACSel, TS, ES, LevSh, Overlay)
    '    '  Read the level sheet, to find pin names
    For Each ws In Worksheets
        If InStr(ws.Cells(1, 1), "DTLevelSheet,") > 0 Then
            SheetName = ws.Name
            row = 4
            Do
                If Application.Worksheets(SheetName).Cells(row, 2).Value = "" Then Exit Do
                '  We look for vil only, so this fails if the digital pin type is Output
                If Application.Worksheets(SheetName).Cells(row, 4).Value = "Vil" Then
                    pin = Application.Worksheets(SheetName).Cells(row, 2).Value

                    ' LogTestInfoVIHLPins.AddPin (Pin)
                    If LogTestInfoVIHLPins.Pins.count = 0 Then
                        LogTestInfoVIHLPins.AddPin (pin)
                    Else
                        For i = 0 To LogTestInfoVIHLPins.Pins.count - 1
                            If LogTestInfoVIHLPins.Pins(i) = pin Then Exit For
                            If i = LogTestInfoVIHLPins.Pins.count - 1 Then
                                LogTestInfoVIHLPins.AddPin (pin)
                            End If
                        Next i
                    End If
                End If
                row = row + 1
            Loop
        End If
    Next ws

    ' VIHL/VOHL
    '    Dim pin As Variant
    '    For Each pin In LogTestInfoVIHLPins.Pins
    '
    '    Next pin

    For i = 0 To LogTestInfoVIHLPins.Pins.count - 1

        DigPinValue = Format(TheHdw.Digital.Pins(LogTestInfoVIHLPins.Pins(i)).Levels.Value(chVih), "0.####") + "V"
        LogBuffer = LogBuffer + DigPinValue

        DigPinValue = Format(TheHdw.Digital.Pins(LogTestInfoVIHLPins.Pins(i)).Levels.Value(chVil), "0.####") + "V"
        LogBuffer = LogBuffer + DigPinValue


        DigPinValue = Format(TheHdw.Digital.Pins(LogTestInfoVIHLPins.Pins(i)).Levels.Value(chVoh), "0.####") + "V"
        LogBuffer = LogBuffer + DigPinValue

        DigPinValue = Format(TheHdw.Digital.Pins(LogTestInfoVIHLPins.Pins(i)).Levels.Value(chVol), "0.####") + "V"
        LogBuffer = LogBuffer + DigPinValue

    Next i



    'TimSet/Frequency
    Dim RowPF As Boolean
    Dim temp As String

    TimeSetName = TheHdw.Digital.Timing.TimeSetNameList
    TSetName = TimeSetName
    Dim ClkPeriod As String
    Dim ClkPinName As String
    RowPF = True

    'Collect TimeSet and Period/Frequency
    temp = ""
    If InStr(1, TimeSetName, ",") > 1 Then
        TimeSetList = Split(TimeSetName, ",")
        For TimeSetIndex = 0 To UBound(TimeSetList)
            If TimeSetList(TimeSetIndex) <> "" Then
                Period = TheHdw.Digital.Timing.Period(TimeSetList(TimeSetIndex)).Value
                If temp <> "" Then
                    temp = temp + TimeSetList(TimeSetIndex) + "(" + Format(Period * 1000000000#, "0.####") + "ns/" + Format(1 / (Period * 1000000#), "0.####") + "MHz" + "),"
                Else
                    temp = TimeSetList(TimeSetIndex) + "(" + Format(Period * 1000000000#, "0.####") + "ns/" + Format(1 / (Period * 1000000#), "0.####") + "MHz" + "),"
                End If
            End If
        Next TimeSetIndex
        TSetName = temp

    Else
        Period = TheHdw.Digital.Timing.Period(TimeSetName).Value
        TSetName = TimeSetName + "(" + Format(Period * 1000000000#, "0.####") + "ns/" + Format(1 / (Period * 1000000#), "0.####") + "MHz" + "),"
    End If

    temp = ""
    StrPare = ""

    For Each ws In Worksheets
        SheetName = ws.Name
        If SheetName = TS Then
            row = 8
            Do
                'Checking Compare Edge
                If Application.Worksheets(TS).Cells(row, 13).Value = "" Then Exit Do
                If Application.Worksheets(TS).Cells(row, 13).Value = "Off" Then
                    If Application.Worksheets(TS).Cells(row, 14).Value = "Disable" Then
                        RowPF = RowPF And 1
                    Else

                        RowPF = RowPF And 0
                    End If
                Else

                    RowPF = RowPF And 1
                End If

                'Checking ClockPin and Period

                If Application.Worksheets(TS).Cells(row, 6).Value = "clock" Then

                    ClkPeriod = Format(Application.Worksheets(TS).Cells(row, 5).Value * 1000000000, "0.####") + "ns/"    ' _
                                                                                                                        + Format(1 / (Application.Worksheets(TS).Cells(row, 5).Value * 1000000), "0.####") + "MHz"
                    ClkPinName = Application.Worksheets(TS).Cells(row, 4).Value
                    temp = temp + ClkPinName + "(" + ClkPeriod + "),"

                End If


                Dim PeriodPareList() As String
                Dim PeriodPareIndex As Variant


                If Trim(StrPare) = "" Then
                    StrPare = Application.Worksheets(TS).Cells(row, 3).Formula
                Else
                    If Trim(StrPare) = Application.Worksheets(TS).Cells(row, 3).Formula Then
                        StrPare = StrPare
                    Else
                        If InStr(1, StrPare, ",") > 1 Then
                            PeriodPareList = Split(StrPare, ",")
                            For PeriodPareIndex = 0 To UBound(PeriodPareList)

                                If PeriodPareList(PeriodPareIndex) = Application.Worksheets(TS).Cells(row, 3).Formula Then Exit For

                                If PeriodPareIndex = UBound(PeriodPareList) Then
                                    PeriodPare = PeriodPare + "," + Application.Worksheets(TS).Cells(row, 3).Formula
                                End If
                            Next PeriodPareIndex
                        Else
                            StrPare = StrPare + "," + Application.Worksheets(TS).Cells(row, 3).Formula

                        End If
                    End If
                End If

                row = row + 1
            Loop

            PeriodPare = StrPare
        End If

    Next

    'Edge Check
    If RowPF Then
        TSetEdge = "P"
    Else
        TSetEdge = "F"
    End If


    LogBuffer = LogBuffer + TSetEdge + PeriodPare + TSetName + temp + vbNewLine

End Function


Public Function FindPowerFromPinSheet(PowerPinNames() As String)

    Dim SheetName As String, PinSheet As String
    Dim Buf As String, row As Long, col As Long
    Dim NumPowPins As Long, i As Long
    Dim ws As Worksheet, HasWorksheet As Boolean
    Dim JobName As String, PartName As String, Environment As String

    For Each ws In Worksheets
        If InStr(ws.Cells(1, 1), "DTJobListSheet,") > 0 Then
            SheetName = ws.Name
            HasWorksheet = True
            TheExec.DataManager.GetJobContext JobName, PartName, Environment
            row = 5
            Do
                Buf = Application.Worksheets(SheetName).Cells(row, 2).Value
                If Buf = "" Then Exit Do
                If UCase(Buf) = UCase(JobName) Then
                    PinSheet = Application.Worksheets(SheetName).Cells(row, 3).Value
                    GoTo HavePinSheet
                End If
                row = row + 1
            Loop
        End If
    Next ws

    '  No Job sheet, find the single pin sheet
    If HasWorksheet = False Then
        For Each ws In Worksheets
            If InStr(ws.Cells(1, 1), "DTPinMap,") = 0 Then
                PinSheet = ws.Name
                GoTo HavePinSheet
            End If
        Next ws
    End If

    '  Read the pin sheet and first count the number of power pins
HavePinSheet:
    row = 4
    Do
        If Application.Worksheets(PinSheet).Cells(row, 2).Value <> "" Then Exit Do
        If LCase(Application.Worksheets(PinSheet).Cells(row, 4).Value) = "power" Then
            NumPowPins = NumPowPins + 1
        End If
        row = row + 1
    Loop

    ' Set the array size, then fill it with a list of power pin names
    ReDim PowerPinNames(NumPowPins - 1)
    row = 4
    i = 0
    Do
        If Application.Worksheets(PinSheet).Cells(row, 2).Value <> "" Then Exit Do
        If LCase(Application.Worksheets(PinSheet).Cells(row, 4).Value) = "power" Then
            PowerPinNames(i) = Application.Worksheets(PinSheet).Cells(row, 3).Value
            i = i + 1
        End If
        row = row + 1
    Loop

End Function



