Option Explicit

Dim ObjExcel 
Dim ObjWB
On Error Resume Next


Set ObjExcel = GetObject(, "excel.application") 'gives error 429 if Word is not open
If Err.Number = 429 Then
  Err.Clear
  Set ObjExcel = CreateObject("excel.Application")
   
End If
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