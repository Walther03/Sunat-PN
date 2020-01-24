
	Tipo_pago		= DataTable("Tipo_pago", 12)

Set shell = CreateObject("Wscript.Shell") 
	shell.SendKeys "{PGUP}"

CALL Fun_PAGO_Captura1()



'Select Case Tipo_pago
'	
'	Case "BDN"
'		Browser("Rentas").Page("Pagos").WebRadioGroup("BancoNacion").Select "#0"
'
'
'	
'	Case "BCP"
'		Browser("Rentas").Page("Pagos").WebRadioGroup("BCP").Select "#0"
'
'
'	Case "BC"
'		Browser("Rentas").Page("Pagos").WebRadioGroup("BancoComercio").Select "#0"
'
'	
'	Case "SCOTIBANK"
'		Browser("Rentas").Page("Pagos").WebRadioGroup("Scotibank").Select "#0"
'	
'	Case "INTERBANK"
'		Browser("Rentas").Page("Pagos").WebRadioGroup("Interbank").Select "#0"
'
'	Case "PICHINCHA"
'		Browser("Rentas").Page("Pagos").WebRadioGroup("BancoPichincha").Select "#0"
'	
'	Case "GNB"
'		Browser("Rentas").Page("Pagos").WebRadioGroup("BancoGNB").Select "#0"
'
'	Case "BBVA"
'		Browser("Rentas").Page("Pagos").WebRadioGroup("BvaContinental").Select "#0"
'	
'	Case "CITYBANK"
'		Browser("Rentas").Page("Pagos").WebRadioGroup("Citibank").Select "#0"
'
'	Case "BANBIF"
'		Browser("Rentas").Page("Pagos").WebRadioGroup("Banbif").Select "#0"
'	
'	Case "SANTANDER"
'		Browser("Rentas").Page("Pagos").WebRadioGroup("Santander").Select "#0"
'
'	Case "VISA"
'		Browser("Rentas").Page("Pagos").WebRadioGroup("Visa").Select "#0"
'		Browser("Rentas").Page("Pagos").WebButton("Presente/Pague").Click
'
'		Browser("Rentas").Page("Pagos").WebButton("Aceptar").Click
'		Browser("Rentas").Page("Pagos").WebButton("Aceptar").Click
'		
'
'		Browser("Rentas").Page("Pagos").Frame("Frame").WebEdit("txtTarjetaTitular").Set "4220 5200 0329 2565"
'		Browser("Rentas").Page("Pagos").Frame("Frame").WebList("cbMes").Select "#3"
'		Browser("Rentas").Page("Pagos").Frame("Frame").WebList("cbAnho").Select "#3"
'		Browser("Rentas").Page("Pagos").Frame("Frame").WebEdit("txtCodSeguridad").Set 123
'		Browser("Rentas").Page("Pagos").Frame("Frame").WebEdit("txtNomTitular").Set "asdas"
'		Browser("Rentas").Page("Pagos").Frame("Frame").WebEdit("txtApeTitular").Set "asd"
'
'
'	
'	Case "NPS"
'		Browser("Rentas").Page("Pagos").WebRadioGroup("NPS").Select "#0"
'	
'End Select

CALL Fun_PAGO_Captura2()

Function Fun_PAGO_Captura1()

	Browser("Rentas").CaptureBitmap RutaEvidencias()  & numIter &  "pago1.png", True
	imagenToWord "Pago 1", RutaEvidencias()  & numIter &  "pago1.png"
	
	
End Function
		
Function Fun_PAGO_Captura2()
	
	Set shell = CreateObject("Wscript.Shell") 
		shell.SendKeys "{PGDN}"
	Browser("Rentas").CaptureBitmap RutaEvidencias()  & numIter &  "pago2.png", True
	imagenToWord "Pago 2", RutaEvidencias()  & numIter &  "pago2.png"
	
	
End Function

