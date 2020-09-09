Attribute VB_Name = "VBT_basic_template_common"
Option Explicit

' ### Teradyne HISI UFLEX Basic Templates V13.00 ###

' ================================================================================
'                        Template Public Declares
' ================================================================================
' For DDR CAL Read
Global g_DDR_CAL_READ_DSP As New DSPWave

' For SHM/PM
Public Enum Enum_ResultsMode
    ResultsSeparate = 0
    ResultsAnd = 1
    ResultsOr = 2
End Enum

'To Enum Different Test Mode
Public Enum Exec_Modes
    Production_Mode = 0
    Debug_Mode = 1
    AutoSearch_Mode = 2
End Enum

' For SHM/PM
Public Type Type_PointData
    Xval As New SiteDouble
    Yval As New SiteDouble
    Result() As New SiteLong
    Center As New SiteBoolean
End Type


#If Win64 Then
    Public Declare PtrSafe Sub MD5Init Lib "Cryptdll.dll" (ByVal pContex As LongPtr)
    Public Declare PtrSafe Sub MD5Final Lib "Cryptdll.dll" (ByVal pContex As LongPtr)
    Public Declare PtrSafe Sub MD5Update Lib "Cryptdll.dll" (ByVal pContex As LongPtr, ByVal lPtr As LongPtr, ByVal nSize As LongPtr)

#Else

    Public Declare Sub MD5Init Lib "Cryptdll.dll" (ByVal pContex As Long)
    Public Declare Sub MD5Final Lib "Cryptdll.dll" (ByVal pContex As Long)
    Public Declare Sub MD5Update Lib "Cryptdll.dll" (ByVal pContex As Long, ByVal lPtr As Long, ByVal nSize As Long)
#End If

Public Type Type_MD5_CTX
    i(1) As Long
    Buf(3) As Long
    Inc(63) As Byte
    TempMD5(15) As Byte
End Type

Public PrintInfo As Boolean
Public glb_PrintCycleCount_bool As Boolean


' ================================================================================
'                        Template Public Functions
' ================================================================================

' Return true means burst=yes; return false means burst=no
Private Function CheckBurstForPatset(patset As String, Sheet As String) As Boolean
    Dim row As Long
    Dim PatSetVersion As String
    Dim BurstIndex As Long
    row = 4
    Do
        If Application.Worksheets(Sheet).Cells(row, 2).Value = "" Then Exit Do

        ' Debug.Print Application.Worksheets(sheet).Cells(row, 2).Value

        If UCase(Application.Worksheets(Sheet).Cells(row, 2).Value) = UCase(patset) Then

            ' For different Version IG-XL the burst column is different, To support both versions EricPeng 20161221
            If InStr(1, Application.Worksheets(Sheet).Cells(1, 1).Value, "version=2.1:") > 0 Then    ' IGXL 8.10.14
                BurstIndex = 6
            ElseIf InStr(1, Application.Worksheets(Sheet).Cells(1, 1).Value, "version=2.2:") > 0 Then  'IGXL 8.30.02

                BurstIndex = 7
            ElseIf InStr(1, Application.Worksheets(Sheet).Cells(1, 1).Value, "version=2.3:") > 0 Then  ' add for IGXL 10.10
                BurstIndex = 6

            Else
                TheExec.Datalog.WriteComment "Unknown IG-XL PatSet Verion please contact Teradyne GSO for support "

            End If


            If LCase(Application.Worksheets(Sheet).Cells(row, BurstIndex).Value) = "no" Then
                CheckBurstForPatset = False
                Exit Function
            Else
                CheckBurstForPatset = True
                Exit Function
            End If
        End If
        row = row + 1
    Loop
End Function



Public Function PatSetInfo(PatName As String, _
                           PatListName() As String, PatPath() As String, PatMD5() As String, _
                           Optional PatMD5Flag As Boolean = False, _
                           Optional PrintFlag As Boolean = False)
' V10.00.01 Created by Eric Peng
' This function is extract the patterns from the patterns seprated by ", " And Pattern Sets
' V11.01 If pattern Set burst Yes, Then output the pattern name together, Split by ","
'Else the pattern name will display seperately with the mode burst no.
    Dim i, j As Long
    Dim i_TempArray_str() As String
    Dim i_TemPattern_str As String
    Dim i_TempLen_int As Long
    ' Dim i_ModName_str() As String
    Dim i_PatIndex_lng As Long
    Dim i_PatListNameTempi_str() As String
    Dim i_PatListNameTempj_str() As String
    Dim i_OutStr_str As String

    i_PatIndex_lng = -1
    If InStr(1, PatName, ",") > 0 Then
        i_PatListNameTempi_str = Split(PatName, ",")
    Else
        ReDim Preserve i_PatListNameTempi_str(0)
        i_PatListNameTempi_str(0) = PatName
    End If

    Dim i_Wksheet As Excel.Worksheet
    Dim i_PatSetSheetName_str As String
    Dim PatternSetSheetNum As Long

    PatternSetSheetNum = 0
    For Each i_Wksheet In Application.Worksheets
        ' DTPatternSetSheet,version=2.2:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1
        If InStr(i_Wksheet.Cells(1, 1), "DTPatternSetSheet,") > 0 Then
            ' Debug.Print CheckBurstForPatset("patset_bno", wks.Name)
            ' Debug.Print CheckBurstForPatset("patset_byes", wks.Name)
            i_PatSetSheetName_str = i_Wksheet.Name

            PatternSetSheetNum = PatternSetSheetNum + 1

        End If
    Next i_Wksheet


    If PatternSetSheetNum > 1 Then

        TheExec.Datalog.WriteComment "The Template is Only Support one pattern Set Sheet"


    End If


    ' Get Each Pat Name, Pat Path and Pat MD5
    For i = 0 To UBound(i_PatListNameTempi_str)
        If UCase(Right(i_PatListNameTempi_str(i), 4)) = ".PAT" Or UCase(Right(i_PatListNameTempi_str(i), 2)) = "GZ" Then    ' if not PatSet
            i_PatIndex_lng = i_PatIndex_lng + 1
            ReDim Preserve PatListName(i_PatIndex_lng) As String
            ReDim Preserve PatPath(i_PatIndex_lng) As String
            ReDim Preserve PatMD5(i_PatIndex_lng) As String
            i_TempArray_str = Split(i_PatListNameTempi_str(i), "\")
            PatListName(i_PatIndex_lng) = i_TempArray_str(UBound(i_TempArray_str))  ' get the last data is the pattern name
            PatPath(i_PatIndex_lng) = i_PatListNameTempi_str(i)
            If PatMD5Flag Then PatMD5(i_PatIndex_lng) = GetMD5Hash_File(PatPath(i_PatIndex_lng))
            ' ModName = m_stdsvcclient.PatternService.PatSymbolData.PatternSetModules(PatListName(0))
            ' i_OutStr_str = i_OutStr_str & vbTab & PatListName(0) & vbTab & ModName(0) & vbTab & PatPath(0) & vbTab & PatMD5(0)
            If PrintFlag Then i_OutStr_str = i_OutStr_str & vbTab & PatListName(i_PatIndex_lng) & vbTab & PatPath(i_PatIndex_lng) & vbTab & PatMD5(i_PatIndex_lng)
        Else    ' if pattern Set




            If CheckBurstForPatset(i_PatListNameTempi_str(i), i_PatSetSheetName_str) Then    ' For burst Yes Pattern

                i_PatIndex_lng = i_PatIndex_lng + 1

                ReDim Preserve PatListName(i_PatIndex_lng)
                ReDim Preserve PatPath(i_PatIndex_lng)
                ReDim Preserve PatMD5(i_PatIndex_lng)

                ' PatListName(i_PatIndex_lng) = i_PatListNameTempi_str(i)
                PatPath(i_PatIndex_lng) = i_PatListNameTempi_str(i)


                'i_PatListNameTempj_str = m_STDSvcClient.PatternService.PatSymbolData.PatternSetElements(i_PatListNameTempi_str(i))
                
                If TheHdw.Tester.Type = "UltraFLEXplus" Then
                    i_PatListNameTempj_str = CheckPatfileForPatset(i_PatListNameTempi_str(i), i_PatSetSheetName_str)
                Else
                    i_PatListNameTempj_str = m_STDSvcClient.PatternService.PatSymbolData.PatternSetElements(i_PatListNameTempi_str(i))
                End If

                ' ReDim Preserve PatPath(i_PatIndex_lng + UBound(i_PatListNameTempj_str) + 1)
                '  ReDim Preserve PatMD5(i_PatIndex_lng + UBound(i_PatListNameTempj_str) + 1)
                '  ReDim Preserve PatListName(i_PatIndex_lng + UBound(i_PatListNameTempj_str) + 1)

                PatListName(i_PatIndex_lng) = i_PatListNameTempj_str(0)
                For j = 1 To UBound(i_PatListNameTempj_str)
                    ' i_PatIndex_lng = i_PatIndex_lng + 1
                    'PatListName(i_PatIndex_lng) = i_PatListNameTempj_str(j) & "," & i_PatListNameTempj_str(j + 1)
                    PatListName(i_PatIndex_lng) = PatListName(i_PatIndex_lng) & "," & i_PatListNameTempj_str(j)


                    '  PatPath(i_PatIndex_lng) = TheHdw.Patterns(i_PatListNameTempj_str(j)).Path & PatListName(i_PatIndex_lng)
                    ' If PatMD5Flag Then PatMD5(i_PatIndex_lng) = GetMD5Hash_File(PatPath(i_PatIndex_lng))
                    'ModName = m_stdsvcclient.PatternService.PatSymbolData.PatternSetModules(PatListName(i_PatIndex_lng))
                    'i_OutStr_str = i_OutStr_str & vbTab & PatListName(i_PatIndex_lng) & vbTab & ModName(0) & vbTab & PatPath(i_PatIndex_lng) & vbTab & PatMD5(i_PatIndex_lng) & vbCrLf
                    'If PrintFlag Then i_OutStr_str = i_OutStr_str & vbTab & PatListName(i_PatIndex_lng) & vbTab & PatPath(i_PatIndex_lng) & vbTab & PatMD5(i_PatIndex_lng) & vbCrLf
                Next j
            Else
                i_PatListNameTempj_str = m_STDSvcClient.PatternService.PatSymbolData.PatternSetElements(i_PatListNameTempi_str(i))
                'i_PatIndex_lng = i_PatIndex_lng + UBound(i_PatListNameTempj_str) + 1
                ReDim Preserve PatPath(i_PatIndex_lng + UBound(i_PatListNameTempj_str) + 1)
                ReDim Preserve PatMD5(i_PatIndex_lng + UBound(i_PatListNameTempj_str) + 1)
                ReDim Preserve PatListName(i_PatIndex_lng + UBound(i_PatListNameTempj_str) + 1)
                For j = 0 To UBound(i_PatListNameTempj_str)
                    i_PatIndex_lng = i_PatIndex_lng + 1
                    PatListName(i_PatIndex_lng) = i_PatListNameTempj_str(j)
                    PatPath(i_PatIndex_lng) = TheHdw.Patterns(i_PatListNameTempj_str(j)).Path & PatListName(i_PatIndex_lng)
                    If PatMD5Flag Then PatMD5(i_PatIndex_lng) = GetMD5Hash_File(PatPath(i_PatIndex_lng))
                    'ModName = m_stdsvcclient.PatternService.PatSymbolData.PatternSetModules(PatListName(i_PatIndex_lng))
                    'i_OutStr_str = i_OutStr_str & vbTab & PatListName(i_PatIndex_lng) & vbTab & ModName(0) & vbTab & PatPath(i_PatIndex_lng) & vbTab & PatMD5(i_PatIndex_lng) & vbCrLf
                    If PrintFlag Then i_OutStr_str = i_OutStr_str & vbTab & PatListName(i_PatIndex_lng) & vbTab & PatPath(i_PatIndex_lng) & vbTab & PatMD5(i_PatIndex_lng) & vbCrLf
                Next j
            End If
        End If
    Next i

    If PrintFlag Then TheExec.Datalog.WriteComment i_OutStr_str

End Function

Private Function CheckPatfileForPatset(patset As String, Sheet As String) As String()
    Dim row As Long
    Dim i As Long
    Dim PatSetVersion As String
    Dim PatfileIndex As Long
    Dim Patfilename As String
    Dim i_PatListNameTempi_str() As String
    Dim i_PatListName_str() As String
    Dim i_TempArray_str() As String
    Dim bFirstelement As Boolean
    bFirstelement = True
    
    row = 4
    ' For different Version IG-XL the burst column is different, To support both versions EricPeng 20161221
    If InStr(1, Application.Worksheets(Sheet).Cells(1, 1).Value, "version=2.1:") > 0 Then ' IGXL 8.10.14
        PatfileIndex = 5
    ElseIf InStr(1, Application.Worksheets(Sheet).Cells(1, 1).Value, "version=2.2:") > 0 Then  'IGXL 8.30.02
        PatfileIndex = 6
    ElseIf InStr(1, Application.Worksheets(Sheet).Cells(1, 1).Value, "version=2.3:") > 0 Then  'IGXL 10.10.11 update 20190819
        PatfileIndex = 5
    
    Else
        TheExec.Datalog.WriteComment "Unknown IG-XL PatSet Verion please contact Teradyne GSO for support "
    
    End If
    Do
        If Application.Worksheets(Sheet).Cells(row, 2).Value = "" Then Exit Do
        
       ' Debug.Print Application.Worksheets(sheet).Cells(row, 2).Value
        
        If UCase(Application.Worksheets(Sheet).Cells(row, 2).Value) = UCase(patset) Then
                        
            If bFirstelement Then
                Patfilename = Application.Worksheets(Sheet).Cells(row, PatfileIndex).Value
                bFirstelement = False
            Else
                Patfilename = Patfilename + "," + Application.Worksheets(Sheet).Cells(row, PatfileIndex).Value
            End If
            

        End If
        row = row + 1
    Loop

    If InStr(1, Patfilename, ",") > 0 Then
        i_PatListNameTempi_str = Split(Patfilename, ",")
        ReDim i_PatListName_str(UBound(i_PatListNameTempi_str))
    Else
        ReDim Preserve i_PatListNameTempi_str(0)
        ReDim Preserve i_PatListName_str(0)
        i_PatListNameTempi_str(0) = Patfilename
    End If
    
    For i = 0 To UBound(i_PatListNameTempi_str)
        i_TempArray_str = Split(i_PatListNameTempi_str(i), "\")
        i_PatListName_str(i) = i_TempArray_str(UBound(i_TempArray_str))  ' get the last data is the pattern name
    Next i
    
    CheckPatfileForPatset = i_PatListName_str


End Function

'MD5 Codes
Public Function ConvBytesToBinaryString(BytesIn() As Byte) As String
    Dim i As Long
    Dim i_nSize_lng As Long
    Dim i_StrRet_str As String
    i_nSize_lng = UBound(BytesIn)
    For i = 0 To i_nSize_lng
        i_StrRet_str = i_StrRet_str & Right$("0" & Hex(BytesIn(i)), 2)
    Next
    ConvBytesToBinaryString = i_StrRet_str
End Function

Public Function GetMD5Hash(BytesIn() As Byte) As Byte()
    Dim i_CTX As Type_MD5_CTX
    Dim i_nSize_lng As Long
    i_nSize_lng = UBound(BytesIn) + 1
    MD5Init VarPtr(i_CTX)
    MD5Update ByVal VarPtr(i_CTX), ByVal VarPtr(BytesIn(0)), i_nSize_lng
    MD5Final VarPtr(i_CTX)
    GetMD5Hash = i_CTX.TempMD5
End Function

Public Function GetMD5Hash_Bytes(BytesIn() As Byte) As String       'Byte MD5
    GetMD5Hash_Bytes = ConvBytesToBinaryString(GetMD5Hash(BytesIn))
End Function
'
Public Function GetMD5Hash_String(ByVal StrIn As String) As String   ' String MD5
    GetMD5Hash_String = GetMD5Hash_Bytes(StrConv(StrIn, vbFromUnicode))
End Function

Public Function GetMD5Hash_File(ByVal PatName As String) As String
    Dim i_LFile_lng As Long
    Dim i_Bytes_byt() As Byte
    Dim i_LSize_lng As Long
    Dim i_File_str As String
    i_File_str = PatName
    i_LSize_lng = FileLen(i_File_str)
    If (i_LSize_lng) Then
        i_LFile_lng = FreeFile
        ReDim i_Bytes_byt(i_LSize_lng - 1)
        Open i_File_str For Binary As i_LFile_lng
        Get i_LFile_lng, , i_Bytes_byt
        Close i_LFile_lng
        GetMD5Hash_File = GetMD5Hash_Bytes(i_Bytes_byt)
    End If
End Function


Public Function tl_SetTestState(in_ConnectAllPins As Boolean, _
                                in_LoadLevels As Boolean, _
                                in_LoadTiming As Boolean, _
                                in_RelayMode As tlRelayMode, _
                                in_InitWaitTime As Double, _
                                in_DriveLoPins As PinList, _
                                in_DriveHiPins As PinList, _
                                in_DriveZPins As PinList, _
                                in_FloatPins As PinList, _
                                in_Util1Pins As PinList, _
                                in_Util0Pins As PinList)


' Set drive state on specified utility pins
    If nonblank(in_Util0Pins) Then Call tl_SetUtilState(in_Util0Pins, 0)
    If nonblank(in_Util1Pins) Then Call tl_SetUtilState(in_Util1Pins, 1)

    ' ApplyLevelTiming will
    '   Optionally power down instruments and power supplies
    '   Optionally Close Pin-Electronics, High-Voltage, & Power Supply Relays,
    '       of pins noted on the active levels sheet
    '   Optionally load Timing and Levels information
    '   Set init-state driver conditions on specified pins
    '       Setting init state causes the pin to drive the specified value.  Init
    '       state is set once, during the prebody, before the first pattern burst.
    '       Default is to leave the pin driving whatever value it last drove during
    '       the previous pattern burst.
    Call TheHdw.Digital.ApplyLevelsTiming(in_ConnectAllPins, in_LoadLevels, in_LoadTiming, in_RelayMode)

    If nonblank(in_DriveLoPins) Then Call tl_SetInitState(in_DriveLoPins, chInitLo)
    If nonblank(in_DriveHiPins) Then Call tl_SetInitState(in_DriveHiPins, chInitHi)
    If nonblank(in_DriveZPins) Then Call tl_SetInitState(in_DriveZPins, chInitoff)

    ' Remove specified DUT pins, if any, from connection to tester pin-electronics and other resources
    If nonblank(in_FloatPins) Then Call tl_SetFloatState(in_FloatPins)
    ' Set start-state driver conditions on specified pins.
    ' Start state determines the driver value the pin is set to as each pattern burst starts.
    ' Default is to have start state automatically selected appropriately
    '   depending on the Format of the first vector of each pattern burst.
    If nonblank(in_DriveLoPins) Then Call tl_SetStartState(in_DriveLoPins, chStartLo)
    If nonblank(in_DriveHiPins) Then Call tl_SetStartState(in_DriveHiPins, chStartHi)
    If nonblank(in_DriveZPins) Then Call tl_SetStartState(in_DriveZPins, chStartOff)
    TheHdw.Wait in_InitWaitTime

End Function



Public Function PrintCycleCount(argc As Long, argv() As String) As Long

'Format： ####[CountCyle:] <test_number> <SiteNum> <test_instance_name> <pattern name> < xxx>
    Dim i_TestInstName_str As String
    Dim i_LastBurst_str As String
    Dim i_CycleCount_lng As Long
    Dim i_TestNum_lng As Long
    Dim Site As Variant

    i_TestInstName_str = TheExec.DataManager.InstanceName

    Call TheHdw.Digital.TimeDomains("").Patgen.ReadLastStart(i_LastBurst_str, False, "")

    i_CycleCount_lng = TheHdw.Digital.Patgen.CycleCount(tlCycleTypeAbsolute)

    For Each Site In TheExec.Sites.Active
        i_TestNum_lng = TheExec.Sites(Site).TestNumber
        TheExec.Datalog.WriteComment "####[CountCycle:]<" & i_TestNum_lng & "><" & Site & "><" + i_TestInstName_str & "><" & i_LastBurst_str & "><" & i_CycleCount_lng & ">"
    Next

End Function


