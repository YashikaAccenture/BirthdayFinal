Option Explicit

Dim ObjExcel 
Dim ObjWB
On Error Resume Next
Dim ObjShell


  Set ObjExcel = CreateObject("excel.Application")
  Set ObjShell = Wscript.CreateObject("Wscript.Shell")


objShell.Run"C:\Users\yashika.a.gupta\Desktop\WishingTool\Birthday\New.vbs"
If Not ObjExcel Is Nothing Then
   ObjExcel.Visible = True
    Set ObjWB =ObjExcel.Workbooks.Open("C:\Users\yashika.a.gupta\Desktop\WishingTool\Sheet\Birthday.xlsm")
   ObjExcel.Run("Workbook_Open")
Else
   Msgbox "Unable to retrieve Excel."
End If

ObjWB.Close False 
ObjExcel.Quit
Set ObjExcel = Nothing
Set ObjShell = Nothing