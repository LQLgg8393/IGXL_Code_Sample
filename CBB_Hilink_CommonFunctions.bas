Attribute VB_Name = "CBB_Hilink_CommonFunctions"
Option Explicit


Public Function PLDContainsPin(PLD As PinListData, PinName As String) As Boolean
' This function check if pin exist in PLD
' This function is a component of Hilink CBB, contact GSO Shanghai Team for more details
' Public function meant to be called in H15BP/H20USR/H20USR CBB, may also support other HiLink CBB
' ByVal <PLD> As PinListData
' Return <PLDContainsPin> As Boolean
' 2015/03/04; ver.01

    Dim pd As PinData
    
    For Each pd In PLD.Pins
        If LCase(pd.Name) = LCase(PinName) Then
            PLDContainsPin = True
            Exit Function
        End If
    Next
    
    PLDContainsPin = False

End Function

Public Function RefIsFound(RefKeyWord As String) As Boolean
' This function find if Reference with RefKeyWord has been referenced to IGXL program
' This function is a component of Hilink CBB, contact GSO Shanghai Team for more details
' Public function meant to be called in H15BP CBB, may also support other HiLink CBB
' Return <RefIsFound> As Boolean
' 2016/07/11; ver.01
    
    RefIsFound = False
    
    Dim ws As Worksheet
    Dim WB As String
    Dim ws_count As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim ws_firstcell As String
    Dim refsheet_cell As String
    Dim MD5Path As String
    Dim pathnames As String
    Dim MD5Name As String
    
    Dim isRefSheetExist As Boolean
    Dim isMD5Exist_In_Ref As Boolean
    Dim isMD5Exist_In_Folder As Boolean
    
    isRefSheetExist = False
    isMD5Exist_In_Ref = False
    isMD5Exist_In_Folder = False
    
    ws_count = ThisWorkbook.Sheets.count
    For i = 1 To ws_count
        ws_firstcell = ThisWorkbook.Sheets(i).Range("A1")
        If InStr(1, ws_firstcell, "ReferencesSheet") Then
            isRefSheetExist = True
            Set ws = ThisWorkbook.Sheets(i)
            ws.Activate
            For j = 4 To 10
                refsheet_cell = ws.Cells(j, 2)
                If refsheet_cell = "" Then
                    Exit For
                ElseIf InStr(1, LCase(refsheet_cell), LCase(RefKeyWord)) Then
                    isMD5Exist_In_Ref = True
                    
                    MD5Path = ThisWorkbook.Path + Right(refsheet_cell, Len(refsheet_cell) - 1)
                    pathnames = Dir(MD5Path)
                    For k = Len(refsheet_cell) To 1 Step -1
                        If MID(refsheet_cell, k, 1) = "\" Then Exit For
                    Next k
                    MD5Name = Right(refsheet_cell, Len(refsheet_cell) - k)
                    If LCase(Trim(pathnames)) = LCase(Trim(MD5Name)) Then
                        isMD5Exist_In_Folder = True
                    Else
                    'need to be test if logic is optimized
                        isMD5Exist_In_Folder = False
                        'Do not need to check anymore , IGXL will fail validation
                        Exit For
                    End If
                End If
            Next
        End If
    Next
    
    '==========================================Warnings================================================
    If isRefSheetExist = False Then
        TheExec.AddOutput "No Reference Sheet Found In IGXL! Please Add MD5 Check Tool in Reference Sheet!", vbRed
    ElseIf isMD5Exist_In_Ref = False Then
        TheExec.AddOutput "MD5 Tool Not Found In Reference Sheet! Please Add MD5 Check Tool in Reference Sheet!", vbRed
    ElseIf isMD5Exist_In_Folder = False Then
        TheExec.AddOutput "MD5 Tool Not Found Under Folder In Reference Sheet! Please Add MD5 Check Tool Under Folder!", vbRed
    End If

    RefIsFound = isRefSheetExist And isMD5Exist_In_Ref And isMD5Exist_In_Folder
    
    If RefIsFound = False Then
        TheExec.AddOutput ""
        TheExec.AddOutput "MD5 Tool Not Valid! Hilink OnProgramLoaded & OnProgramValidated Function Not Run!", vbRed
        TheExec.AddOutput "Please Add MD5 Tool In Reference Sheet And Load Program Again!!!", vbRed
    End If
  
End Function

Public Function Hilink_InvertHL8Bit(ByVal RawBigEndin As Long) As Long

    Hilink_InvertHL8Bit = ((RawBigEndin And &HFF&) * 2 ^ 8) + ((RawBigEndin And &HFF00&) \ 2 ^ 8)
    
End Function

