Attribute VB_Name = "Exec_IP_Module"
Option Explicit
 
' This module contains empty Exec Interpose functions (see online help
' for details).  These are here for convenience and are completely optional.
' It is not necessary to delete them if they are not being used, nor is it
' necessary that they exist in the program.



' Immediately at the conclusion of the initialization process.
' Do not program test system hardware from this function.
Function OnTesterInitialized()
    On Error GoTo errHandler

        ' Init sub program exec interpose classes
        Dim exec1 As IExecInterposeFunctions
        Set exec1 = New CLS_basic_exec_ip
        exec1.OnTesterInitialized


    
    
    Exit Function
errHandler:
    ' OnTesterInitialized executes before TheExec is even established so nothing
    ' better to do then msgbox in this case.  Note that unhandled errors can allow the
    ' user to press "End" which will result in a DataTool crash.  Errors in this routine
    ' need to be debugged carefully.
    MsgBox "Error encountered in Exec Interpose Function OnTesterInitialized" + vbCrLf + _
        "VBT Error # " + Trim(Str(Err.Number)) + ": " + Err.Description
End Function
 
' Immediately at the conclusion of the load process.
' Do not program test system hardware from this function.
Function OnProgramLoaded()

    On Error GoTo errHandler

        ' Init sub program exec interpose classes
        Dim exec1 As IExecInterposeFunctions
        Set exec1 = New CLS_basic_exec_ip
        exec1.OnProgramLoaded

        Dim exec2 As IExecInterposeFunctions
        Set exec2 = New Cls_HiLink_Exec_IP_Module
        exec2.OnProgramLoaded
   
    Exit Function
errHandler:
    HandleExecIPError "OnProgramLoaded"
End Function
 
' Immediately at the conclusion of the validate process. Called only if validation succeeds.
Function OnProgramValidated()
    On Error GoTo errHandler

        ' Init sub program exec interpose classes
        Dim exec1 As IExecInterposeFunctions
        Set exec1 = New CLS_basic_exec_ip
        exec1.OnProgramValidated

        Dim exec2 As IExecInterposeFunctions
        Set exec2 = New Cls_HiLink_Exec_IP_Module
        exec2.OnProgramValidated
    
    Exit Function
errHandler:
    HandleExecIPError "OnProgramValidated"
End Function
 
' Immediately at the conclusion of the validate process. Called only if validation fails.
Function OnProgramFailedValidation()
    On Error GoTo errHandler

        ' Init sub program exec interpose classes
        Dim exec1 As IExecInterposeFunctions
        Set exec1 = New CLS_basic_exec_ip
        exec1.OnProgramFailedValidation


    
    
    Exit Function
errHandler:
    HandleExecIPError "OnProgramFailedValidation"
End Function
 
' Immediately at the conclusion of the user DIB calibration process (previously
' known as the TDR calibration process). Called only if user DIB calibration succeeds.
Function OnTDRCalibrated()
    On Error GoTo errHandler

        ' Init sub program exec interpose classes
        Dim exec1 As IExecInterposeFunctions
        Set exec1 = New CLS_basic_exec_ip
        exec1.OnTDRCalibrated


    
    
    Exit Function
errHandler:
    HandleExecIPError "OnTDRCalibrated"
End Function
 
' Immediately after "pre-job reset" when the test program starts.
' Note that "first run" actions can be enclosed in
' If TheExec.ExecutionCount = 0 Then...
' (see online help for ExecutionCount)
Function OnProgramStarted()
    On Error GoTo errHandler

        ' Init sub program exec interpose classes
        Dim exec1 As IExecInterposeFunctions
        Set exec1 = New CLS_basic_exec_ip
        exec1.OnProgramStarted

        Dim exec2 As IExecInterposeFunctions
        Set exec2 = New Cls_HiLink_Exec_IP_Module
        exec2.OnProgramStarted

    
    
    Exit Function
errHandler:
    HandleExecIPError "OnProgramStarted"
End Function
 
' Immediately before "post-job reset" when the test program completes.
' Note that any actions taken here with respect to modification of binning
' will affect the binning sent to the Operator Interface, but will not affect
' the binning reported in Datalog.
Function OnProgramEnded()
    On Error GoTo errHandler

        ' Init sub program exec interpose classes
        Dim exec1 As IExecInterposeFunctions
        Set exec1 = New CLS_basic_exec_ip
        exec1.OnProgramEnded


        Dim exec2 As IExecInterposeFunctions
        Set exec2 = New Cls_HiLink_Exec_IP_Module
        exec2.OnProgramEnded


    
    Exit Function
errHandler:
    HandleExecIPError "OnProgramEnded"
End Function
 
' Immediately before a site is disconnected.
' Use TheExec.Sites.SiteNumber to determine which site is being disconnected.
Function OnPreShutDownSite()
    On Error GoTo errHandler

        ' Init sub program exec interpose classes
        Dim exec1 As IExecInterposeFunctions
        Set exec1 = New CLS_basic_exec_ip
        exec1.OnPreShutDownSite


    
    
    Exit Function
errHandler:
    HandleExecIPError "OnPreShutDownSite"
End Function
 
' Use TheExec.Sites.SiteNumber to determine which site is being disconnected.
' Immediately after a site is disconnected.
Function OnPostShutDownSite()
    On Error GoTo errHandler

        ' Init sub program exec interpose classes
        Dim exec1 As IExecInterposeFunctions
        Set exec1 = New CLS_basic_exec_ip
        exec1.OnPostShutDownSite


    
    
    Exit Function
errHandler:
    HandleExecIPError "OnPostShutDownSite"
End Function
 
' Immediately befoe any new calibration factors are loaded
' or new calibrations run.  Not called if no action is taken during AutoCal.
Function OnAutoCalStarted()
    On Error GoTo errHandler

        ' Init sub program exec interpose classes
        Dim exec1 As IExecInterposeFunctions
        Set exec1 = New CLS_basic_exec_ip
        exec1.OnAutoCalStarted


    
    
    Exit Function
errHandler:
    HandleExecIPError "OnAutoCalStarted"
End Function

' Immediately after AutoCal has completed.
' Not called no action has been taken (new factors loaded, or cal performed).
Function OnAutoCalCompleted()
    On Error GoTo errHandler

        ' Init sub program exec interpose classes
        Dim exec1 As IExecInterposeFunctions
        Set exec1 = New CLS_basic_exec_ip
        exec1.OnAutoCalCompleted


    
    
    Exit Function
errHandler:
    HandleExecIPError "OnAutoCalCompleted"
End Function


' Called right before an alarm is reported
' The alarmList is a tab delimited string of alarm error messages
Function OnAlarmOccurred(alarmList As String)

    On Error GoTo errHandler
    
        ' Init sub program exec interpose classes
        Dim exec1 As IExecInterposeFunctions
        Set exec1 = New CLS_basic_exec_ip
        exec1.OnAlarmOccurred alarmList



    Exit Function
errHandler:
    HandleExecIPError "OnAlarmOccurred"
End Function

' When the user pressed the VB Stop button, this interpose function would be called after OnPostShutDownSite was called.
' The user would put code here to make sure global variable are created and contain the correct data.
Function OnGlobalVariableReset()
    On Error GoTo errHandler

        ' Init sub program exec interpose classes
        Dim exec1 As IExecInterposeFunctions
        Set exec1 = New CLS_basic_exec_ip
        exec1.OnGlobalVariableReset


    
    
    Exit Function
errHandler:
    HandleExecIPError "OnGlobalVariableReset"
End Function

' Immediately once Vaildation get started
Function OnValidationStart()
    On Error GoTo errHandler

        ' Init sub program exec interpose classes
        Dim exec1 As IExecInterposeFunctions
        Set exec1 = New CLS_basic_exec_ip
        exec1.OnValidationStart


    
    
    Exit Function
errHandler:
    HandleExecIPError "OnValidationStart"
End Function
' Immediately at the conclusion of the workbook close process. The function is called in any of the following options,
' File->Close
' File->Exit
' Directly triggered the close (“X”) button of the workbook.
Function OnProgramClose()
    On Error GoTo errHandler

        ' Init sub program exec interpose classes
        Dim exec1 As IExecInterposeFunctions
        Set exec1 = New CLS_basic_exec_ip
        exec1.OnProgramClose




    Exit Function
errHandler:

    HandleExecIPError "OnProgramClose"

End Function
