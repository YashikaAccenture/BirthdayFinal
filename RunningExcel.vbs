Option Explicit

Dim ObjExcel 
Dim ObjWB
On Error Resume Next
Dim ObjShell


  Set ObjExcel = CreateObject("excel.Application")
  Set ObjShell = Wscript.CreateObject("Wscript.Shell")


objShell.Run "C:\Users\asha.chauhan\Desktop\WishingTool\Birthday\New.vbs"
If Not ObjExcel Is Nothing Then
   ObjExcel.Visible = True
    Set ObjWB =ObjExcel.Workbooks.Open("C:\Users\asha.chauhan\Desktop\WishingTool\Sheet\Birthday.xlsm")
   ObjExcel.Run("Workbook_Open")
Else
   Msgbox "Unable to retrieve Excel."
End If

ObjWB.Close False 
ObjExcel.Quit
Set ObjExcel = Nothing
Set ObjShell = Nothing

Dim strScriptToKill
    strScriptToKill = "New.vbs"

Dim objWMIService, objProcess, colProcess

    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set colProcess = objWMIService.ExecQuery ( _ 
        "Select * from Win32_Process " & _ 
        "WHERE (Name = 'cscript.exe' OR Name = 'wscript.exe') " & _ 
        "AND Commandline LIKE '%"& strScriptToKill &"%'" _ 
    )
    For Each objProcess in colProcess
        objProcess.Terminate()
    Next