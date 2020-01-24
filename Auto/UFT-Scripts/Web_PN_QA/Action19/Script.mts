'Call ExcelData()
'DataTable.SetCurrentRow(numIterFlujo)

Call Objetos_Resumen()

Dim fraccionamiento

Call CapturarConstancias()
Function CapturarConstancias()
	Browser("Rentas").Page("Resumen").WebElement("Ver detalle").Click
	Browser("Rentas").Page("Resumen").WebButton("Guardar").Click
	Browser("Rentas").Page("Resumen").WebButton("Cerrar").Click

	Browser("Rentas").Page("Resumen").WebElement("Ver detalle").Click
	Browser("Rentas").Page("Resumen").WebButton("Guardar_2").Click
	Browser("Rentas").Page("Resumen").WebButton("Cancelar").Click

	Browser("Rentas").Page("Resumen").WebElement("Ver todas las constancias").Click
	Browser("Rentas").Page("Resumen").WebButton("Guardar_2").Click
	Browser("Rentas").Page("Resumen").WebButton("Cancelar").Click


End Function

'fraccionamiento		= DataTable("fraccionamiento", 11)
wait 3
'Call Fun_fraccionamiento(fraccionamiento)
Browser("Rentas").CaptureBitmap RutaEvidencias()  & numIter &  "Resumen.png", True
imagenToWord "Resumen ", RutaEvidencias()  & numIter &  "Resumen.png"
WAIT 2


Browser("Rentas").Close 
