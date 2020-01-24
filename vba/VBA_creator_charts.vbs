Dim planilha
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = true
objExcel.DisplayAlerts= False
Set wb = objExcel.Workbooks.Open("U:\Pastas pessoais\Pedro\Códigos e experimentos\R\Code_VBA_Chart\xlsx_file\xlsx_file_export.xlsx")
 'MsgBox wb.Worksheets.Count
For planilha = 1 To wb.Worksheets.Count
	wb.Worksheets(planilha).Shapes.AddChart2
	Next
wb.Save
wb.Close
objExcel.quit