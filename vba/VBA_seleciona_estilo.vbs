Dim planilha
Dim numero_grafico
Dim numero_tabela
Set objExcel = CreateObject("Excel.Application")
Set objWord = CreateObject("Word.Application")
Dim rngFind

objExcel.Visible = true
objExcel.DisplayAlerts= False
objWord.Visible = true
objWord.DisplayAlerts= False


'Set wb = objExcel.Workbooks.Open("U:\Pastas pessoais\Pedro\Codigos  e experimentos\R\Code_VBA_Chart\xlsx_file\xlsx_file_export.xlsx")
Set wb = objExcel.Workbooks.Add

Set word = objWord.Documents.Open("U:\Pastas pessoais\Pedro\Codigos  e experimentos\R\Code_VBA_Chart\docx_extract\docx_extract.docx")

numero_grafico = 0
' Este trecho copia os dados inline shape, formas que estão "na linha"
' Estes gráficos estão alinhados com o texto
' on error resume next
' word.Styles("Legenda Gráfico/Tabela").Copy
word.SelectSimilarFormatting
' for documentos = 1 to word.inlineshapes.count
	' if word.inlineshapes(documentos).hasChart then
		' word.inlineshapes(documentos).Select
		' objWord.Selection.Copy
		' numero_grafico = numero_grafico + 1
		' nome_planilha = "Gráfico " & numero_grafico
		' wb.Sheets.Add.Name = _
                ' nome_planilha
		' wb.Worksheets(nome_planilha).Range("D5").Select
		' wb.Worksheets(nome_planilha).Paste
		' wb.Worksheets(nome_planilha).Range("B1").Value = _
			' wb.Worksheets(nome_planilha).shapes(1).Chart.ChartTitle.Text
		' 'MsgBox documentos
	' End If
' Next
' Este trecho copia os dados shape, formas que estão "soltas"
' Estes gráficos não estão alinhados com o texto
' for documentos_shape = 1 to word.shapes.count
	' if word.shapes(documentos_shape).hasChart then
		' numero_grafico = numero_grafico + 1
		' nome_planilha = "Gráfico " & numero_grafico
		' wb.Sheets.Add.Name = _
                ' nome_planilha
		' word.shapes(documentos_shape).Chart.Select
		' objWord.Selection.Copy
		' wb.Worksheets(nome_planilha).Range("D5").Select
		' wb.Worksheets(nome_planilha).Paste
		' wb.Worksheets(nome_planilha).Range("b1").Value = _
			' wb.Worksheets(nome_planilha).shapes(1).Chart.ChartTitle.Text
	' End If
' Next

' for documentos_tabelas = 2 to word.Tables.count
	' numero_tabela = numero_tabela + 1
	' nome_planilha = "Tabela " & (numero_tabela-1)
	' wb.Sheets.Add.Name = _
			' nome_planilha
	' word.Tables(documentos_tabelas).Select
	' objWord.Selection.Copy
	' wb.Worksheets(nome_planilha).Range("A4").Select
	' wb.Worksheets(nome_planilha).Paste
	' 'wb.Worksheets(nome_planilha).Range("b1").Value = _
	' '	word.shapes(documentos_shape).Chart.ChartTitle.Text
' Next


' wb.SaveAs("U:\Pastas pessoais\Pedro\Codigos  e experimentos\R\Code_VBA_Chart\xlsx_file\xlsx_file_export.xlsx")
' wb.Close
'objExcel.quit
' objWord.Quit'