Dim path
path = "C:\Users\Josenildo\Documents\Stefanini\Vbs\test.xls"
Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open(path)
objExcel.Application.DisplayAlerts = False
objExcel.Application.Visible = True

objExcel.Cells(1, 1).Value = "Test value"
objExcel.Cells(1, 2).Value = "Second value"
objExcel.Cells(2,1).Value = "terceiro value"



objExcel.ActiveWorkbook.SaveAs path
objExcel.ActiveWorkbook.Close
