Dim planilha
Dim numero_grafico
Dim numero_tabela
Dim srs
Dim range_xvalues, range_coluna
Set objExcel = CreateObject("Excel.Application")
Set objWord = CreateObject("Word.Application")

objExcel.Visible = true
objExcel.DisplayAlerts= False
objWord.Visible = true
objWord.DisplayAlerts= False


'Set wb = objExcel.Workbooks.Open("U:\Pastas pessoais\Pedro\Codigos e experimentos\R\Code_VBA_Chart\xlsx_file\xlsx_file_export.xlsx")
Set wb = objExcel.Workbooks.Add
Set worksheet_Func = objExcel.WorksheetFunction
Set word = objWord.Documents.Open("U:\Pastas pessoais\Pedro\Codigos e experimentos\R\Code_VBA_Chart\docx_extract\docx_extract.docx")

numero_grafico = 0
' Este trecho copia os dados inline shape, formas que estão "na linha"
' Estes gráficos estão alinhados com o texto
on error resume next
for documentos = 1 to word.inlineshapes.count
	if word.inlineshapes(documentos).hasChart then
		word.inlineshapes(documentos).Select
		objWord.Selection.Copy
		numero_grafico = numero_grafico + 1
		nome_planilha = "Gráfico " & numero_grafico
		wb.Sheets.Add.Name = _
                nome_planilha
		wb.Worksheets(nome_planilha).Cells.Interior.Color = RGB(255, 255, 255)
		wb.Worksheets(nome_planilha).Range("D5").Select
		wb.Worksheets(nome_planilha).Paste
		wb.Worksheets(nome_planilha).Range("B1").Value = _
			wb.Worksheets(nome_planilha).shapes(1).Chart.ChartTitle.Text
		'MsgBox documentos
		wb.Worksheets(nome_planilha).Range("A1").Select
	End If
Next
' Este trecho copia os dados shape, formas que estão "soltas"
' Estes gráficos não estão alinhados com o texto
for documentos_shape = 1 to word.shapes.count
	if word.shapes(documentos_shape).hasChart then
		numero_grafico = numero_grafico + 1
		nome_planilha = "Gráfico " & numero_grafico
		wb.Sheets.Add.Name = _
                nome_planilha
		' Colorizar worksheet
		wb.Worksheets(nome_planilha).Cells.Interior.Color = RGB(255, 255, 255)
		word.shapes(documentos_shape).Chart.Select
		objWord.Selection.Copy
		wb.Worksheets(nome_planilha).Range("D5").Select
		wb.Worksheets(nome_planilha).Paste
		wb.Worksheets(nome_planilha).Range("b1").Value = _
			wb.Worksheets(nome_planilha).shapes(1).Chart.ChartTitle.Text
		wb.Worksheets(nome_planilha).Range("A1").Select
	End If
Next

' Copia os dados dos gráficos e cola na terceira linha
for each wkst in wb.Worksheets
	on error resume next
	set cht = wkst.Shapes(1).Chart
		For iSrs = 1 to cht.SeriesCollection.Count
			Set srs = cht.SeriesCollection(iSrs)
			On Error Resume Next
			If iSrs = 1 Then
				wkst.Cells(3, 2 * iSrs).Value = srs.Name
				' Cópia dos valores
				wkst.Cells(4, 2 * iSrs - 1).Resize(srs.Points.Count).Value = _
					worksheet_Func.Transpose(srs.XValues)
				wkst.Cells(4, 2 * iSrs).Resize(srs.Points.Count).Value = _
					worksheet_Func.Transpose(srs.Values)
				
			Else
				wkst.Cells(3, 1 * iSrs + 1).Value = srs.Name
				' wkst.Cells(2, 1 * iSrs - 1).Resize(srs.Points.Count).Value = _
				 '   WorksheetFunction.Transpose(srs.XValues)
				wkst.Cells(4, 1 * iSrs + 1).Resize(srs.Points.Count).Value = _
					worksheet_Func.Transpose(srs.Values)

		
			End If
		Next

		' Seleciona Elementos
		' Nome da tabela
		nome_tabela = wkst.Name
		nome_tabela = Replace(nome_tabela,"?","a")
		nome_tabela = Replace(nome_tabela," ","_")
		'MsgBox nome_tabela
		' Posições em X e Y da última linha da tabela e endereços
		y_tabela = cht.SeriesCollection(1).Points.Count+3
		x_tabela = cht.SeriesCollection.Count+1
		endereco_comeco = wkst.Cells(3,1).Address
		endereco_fim = wkst.Cells(y_tabela,x_tabela).Address

		wkst.Cells(3,1).Value = "Unidade:"
		wkst.Cells(3,1).Font.Bold = True
		Set range_tabela = wkst.Range(endereco_comeco,endereco_fim)
		range_tabela.Cells.Interior.Pattern = xlNone

		endereco_comeco_xvalue = wkst.Cells(4,1).Address
		endereco_fim_xvalue = wkst.Cells(y_tabela,1).Address
		Set range_xvalues = wkst.Range(endereco_comeco_xvalue,endereco_fim_xvalue)
		range_xvalues.Font.Bold = True

		' Criação de tabela
		'' Cria a tabela na lista de objetos



		' Atribui a tabela estilo
		'wkst.ListObjects(nome_tabela).TableStyle = "TableStyleLight1"

		' Se a média do nome for maior que 4000, então coloca o estilo como data
		if worksheet_Func.Average(range_xvalues) > 4000 Then
			range_xvalues.NumberFormat = "mmm/yyyy"
		end if

		For iSrs = 1 to cht.SeriesCollection.Count
			' Define os endereços para criar o range
			endereco_comeco_grafico = wkst.Cells(4,iSrs+1).Address
			endereco_fim_grafico = wkst.Cells(y_tabela,iSrs+1).Address
			endereco_nome_serie = wkst.Cells(3,iSrs+1).Address
			' Define a série e o ranges
			Set srs = cht.SeriesCollection(iSrs)
			Set range_nome =  wkst.range(endereco_nome_serie)
			Set range_coluna = wkst.Range(endereco_comeco_grafico,endereco_fim_grafico)

			' O nome da série é definido como a fórmula porque é a única forma de manter
			' o link dinâmico
			nome_serie = "='" & wkst.Name & "'!" & range_nome.Address
			' Formatação do texto do título da série para ficar em negrito
			range_nome.Font.Bold = True
			
			' Delimitação das séries com os valores dos ranges na planilha
			srs.XValues = range_xvalues
			srs.Name = nome_serie
			srs.Values = range_coluna
			Set range_coluna = Nothing
			Set range_nome = Nothing
		Next
	range_tabela.select
	wkst.ListObjects.Add(SourceType = xlSrcRange).Name = nome_tabela

	Set tabela_formato = wkst.ListObjects(nome_tabela)

	tabela_formato.TableStyle = "TableStyleMedium7"
	tabela_formato.Unlist

	set tabela_formato = Nothing
	set range_tabela = Nothing
	set cht = Nothing
next

for documentos_tabelas = 2 to word.Tables.count
	numero_tabela = numero_tabela + 1
	nome_planilha = "Tabela " & (numero_tabela)
	wb.Sheets.Add.Name = _
			nome_planilha
	word.Tables(documentos_tabelas).Select
	objWord.Selection.Copy
	wb.Worksheets(nome_planilha).Cells.Interior.Color = RGB(255, 255, 255)
	wb.Worksheets(nome_planilha).Range("A4").Select
	wb.Worksheets(nome_planilha).Paste
	'wb.Worksheets(nome_planilha).Range("b1").Value = _
	'	word.shapes(documentos_shape).Chart.ChartTitle.Text
Next

wb.SaveAs("U:\Pastas pessoais\Pedro\Codigos e experimentos\R\Code_VBA_Chart\xlsx_file\xlsx_file_export.xlsx")
'wb.Close
'objExcel.quit
objWord.Quit'