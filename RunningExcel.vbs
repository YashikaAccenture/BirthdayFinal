Option Explicit

Dim ObjExcel 
Dim ObjWB
On Error Resume Next
Dim ObjShell


  Set ObjExcel = CreateObject("excel.Application")
  Set ObjShell = Wscript.CreateObject("Wscript.Shell")


objShell.Run"C:\Users\asha.chauhan\Desktop\WishingTool\Birthday\New.vbs"
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