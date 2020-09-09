Attribute VB_Name = "AddExternalReference"
Option Explicit

' ### Teradyne HISI UFLEX Basic Templates V13.00 ###

Public Const HKEY_LOCAL_MACHINE = &H80000002

Global RegProcName$

'Define severity codes
Public Const ERROR_SUCCESS = 0&

'Registry Function Prototypes
#If Win64 Then
Declare PtrSafe Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare PtrSafe Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Declare PtrSafe Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

#Else
Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
#End If

'===========================================================================================
' Setup standard IG-XL refs for SFP (taken from ASCII_Utils.xla)
'    - Requires IG-XL 8.10.10 or newer
'===========================================================================================
Public Function SampleAddReferences() As Boolean

    Dim igxlRelPath As String                   '<- IG-XL path


'___ Get the IGXL path and revision info _______________________________________________
    igxlRelPath = GetRegValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Teradyne\IG-XL", "RootPath", "")

    '___ Add IG-XL References ______________________________________________________________
    '    Call AddReference(igxlRelPath + "\bin\Template.xla")
    '    Call AddReference(igxlRelPath + "\bin\DataTool.xla")
    '    Call AddReference(igxlRelPath + "\bin\Scratch.xla")

    '___ Add .DLL Reference__________________________________________________________________
    '    Call AddReference(igxlRelPath + "\bin\JaguarLanguage.dll")
    Call AddReference("C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB")

End Function


'===========================================================================================
' Add a reference for this XLA
'===========================================================================================
Public Function AddReference(AddInFileName As String, Optional AddInProjectName As String) As Long

    Dim WB As Workbook
    Dim wbName As String
    Dim NewXlaFullPath As String
    Dim OldXlaFileName As String
    Dim Item As Long
    Dim XlaName As String
    Dim AddInList As String
    Dim AddInNameArray As Variant
    Dim AddInName As Variant
    Dim fs

    Set fs = CreateObject("Scripting.FileSystemObject")

    On Error GoTo reference_err
    wbName = Application.ThisWorkbook.Name

    ' Get out if it does not exist this can be true for different version of IG-XL
    ' For instance, Scratch does not exist for J750
    If (Not fs.fileExists(AddInFileName)) Then Exit Function

    XlaName = MID(Dir(AddInFileName), 1, Len(Dir(AddInFileName)) - 4)    'Assumes the name of file is the poject name
    NewXlaFullPath = fs.GetAbsolutePathName(AddInFileName)

    AddInList = XlaName + "," + AddInProjectName
    AddInNameArray = Split(AddInList, ",")

    Set WB = Workbooks(wbName)
    With WB.VBProject

        For Each AddInName In AddInNameArray
            If (AddInName = "") Then Exit For
            ' Remove any existing reference to the xla file first
            For Item = 1 To .References.count
                'Debug.Print "RefName: ", .References.item(item).Name

                ' check vba project name as listed in reference table
                If (UCase(.References.Item(Item).Name) = UCase(AddInName)) Then

                    OldXlaFileName = Dir(.References.Item(Item).FullPath)
                    ' The reference exists already but we always want to remove it
                    ' and add it back either by opening the IG-XL program via the .xls file
                    ' or opening the program using ASCII files to ensure that it exexcutes
                    ' workkbook_open which may add menu items to Excel.
                    On Error GoTo RemoveRefError
                    .References.Remove .References(Item)


                    ' NOTE: Cannot do the following because DataTool and other IG-Xl specific
                    '       add-ins are in control of the main application project VBAPRoject

                    ' !!!! You must also close the xla's workbook before changing the reference to a different path
                    ' If file exists, then it is still opened (hidden) in excel
                    On Error GoTo ErrorClosingOldWorkBook
                    '                    Call Workbooks(OldXlaFileName).Close
                    Exit For
                End If
            Next Item
        Next AddInName

        ' All old references removed so now add
        On Error GoTo AddFromFileErr
        '        Stop
        '        Application.VBE.ActiveVBProject.References.AddFromFile NewXlaFullPath
        '        .References.AddFromGuid "{0002E157-0000-0000-C000-000000000046}", 0, 0
        .References.AddFromFile NewXlaFullPath

    End With

    'Fs.Close
    Set fs = Nothing

    ' All done
    Exit Function

    ' Error handling
reference_err:
    On Error GoTo 0
    MsgBox ("Error" + _
          " In Function: " + "AddReference()" + ": " + vbNewLine & vbNewLine + _
            "Could not add Reference: " + vbNewLine + _
            AddInFileName + vbNewLine + _
          " Please verify that Referenced File " + vbNewLine + _
            AddInFileName + vbNewLine + _
          " exists" + vbNewLine), vbCritical, "IG_XL Utils Error"
    Resume Next

RemoveRefError:
    MsgBox ("Error Removing Existing Reference" + AddInFileName)
    On Error GoTo 0
    Resume Next

AddFromFileErr:
    Select Case Err.Number
    Case 32813
        AddReference = 1    ' we have been already there, now it's really added
    Case Else
        AddReference = 0
    End Select
    On Error GoTo 0
    Resume Next   ' Resume execution at line after error.

ErrorClosingOldWorkBook:
    MsgBox ("Error Closing Existing WorkBook / Reference:  " + OldXlaFileName)
    Stop
    On Error GoTo 0
    Resume Next

End Function


' This function was copied from the VB Example included in the MSDN Library:
' Reading and Modifying the Windows Registry Through Visual Basic
'
Function GetRegValue(hKey As Long, ByVal lpszSubKey As String, ByVal szKey As String, szDefault As String) As Variant

    On Error GoTo ErrorTrap
    RegProcName = "GetRegValue()"
    Dim phkResult As Long, lResult As Long, szBuffer As String, lBuffSize As Long

    'Create Buffer
    ' This used to be a piddling 255, but I expanded it to 2048.
    ' I should have fixed this routine instear, but it's used in
    ' too many places.    Vin Shelton 3/19/99
    szBuffer = VBA.Space(2048)
    lBuffSize = VBA.Len(szBuffer)

    'Open the key
    RegOpenKeyEx hKey, lpszSubKey, 0, 1, phkResult

    'Query the value
    lResult = RegQueryValueEx(phkResult, szKey, 0, 0, ByVal szBuffer, lBuffSize)

    'Close the key
    RegCloseKey phkResult

    'Return obtained value
    If lResult = ERROR_SUCCESS Then
        GetRegValue = VBA.Left$(szBuffer, lBuffSize - 1)
    Else
        GetRegValue = szDefault
    End If
    Exit Function

ErrorTrap:
    'ErrorHandler Err.Number, False
    '  MsgBox "ERROR #" & Str$(Err) & " : " & Error & Chr(13) _
       & "Please exit and try again."
    GetRegValue = szDefault

End Function
