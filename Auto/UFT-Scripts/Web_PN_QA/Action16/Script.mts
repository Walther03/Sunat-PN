'		Browser("Rentas").Page("R_PrimeraCategoria").WebEdit("RPC_Txt_100").Click
'		Browser("Rentas").Page("R_PrimeraCategoria").WebEdit("RPC_Txt_100").WebButton("RPC_Txt_100_Aceptar").Click
'		Browser("Rentas").Page("R_PrimeraCategoria").WebEdit("RPC_Txt_100").WebButton("RPC_Txt_100_Agregar").Click
'		Browser("Rentas").Page("R_PrimeraCategoria").WebEdit("RPC_Txt_100").WebList("RPC_Txt_100_tipoDocumento").Select "#5"
'		Browser("Rentas").Page("R_PrimeraCategoria").WebEdit("RPC_Txt_100").WebEdit("RPC_Txt_100_nroDocumento").Set 10282298738
'		Browser("Rentas").Page("R_PrimeraCategoria").WebEdit("RPC_Txt_100").WebList("RPC_Txt_100_cmbTipoBien").Select "#2"
'		Browser("Rentas").Page("R_PrimeraCategoria").WebEdit("RPC_Txt_100").WebEdit("RPC_Txt_100_txtNumPart").Set 1423
'		Browser("Rentas").Page("R_PrimeraCategoria").WebEdit("RPC_Txt_100").WebButton("RPC_Txt_100_Grabar").Click
'		Browser("Rentas").Page("R_PrimeraCategoria").WebEdit("RPC_Txt_100").WebButton("RPC_Txt_100_Aceptar_cabecera").Click
'		Browser("Rentas").Page("R_PrimeraCategoria").WebEdit("RPC_Txt_100").WebButton("RPC_Txt_100_Aceptar").Click
'	
'		Browser("Rentas").Page("R_PrimeraCategoria").WebEdit("RPC_Txt_100").WebButton("RPC_Txt_100_Agregar_2").Click
'		Browser("Rentas").Page("R_PrimeraCategoria").WebEdit("RPC_Txt_100").WebElement("RPC_Txt_100_Calendario").Click
'		Browser("Rentas").Page("Det_Deuda").WebButton("WebButton_3").Click
'		Browser("Rentas").Page("Det_Deuda").WebElement("ene_2").Click
'
'		
'		Browser("Rentas").Page("R_PrimeraCategoria").WebEdit("RPC_Txt_100").WebList("RPC_Txt_100_numeroFormulario").Select "#1"
'		Browser("Rentas").Page("R_PrimeraCategoria").WebEdit("RPC_Txt_100").WebEdit("RPC_Txt_100_numOrden").Set 451
'		Browser("Rentas").Page("R_PrimeraCategoria").WebEdit("RPC_Txt_100").WebElement("RPC_Txt_100_Calendario2").Click
'		wait 1
'		Set shell = CreateObject("Wscript.Shell") 
'			shell.SendKeys "{ENTER}"
'		wait 1
'		Browser("Rentas").Page("R_PrimeraCategoria").WebEdit("RPC_Txt_100").WebEdit("RPC_Txt_100_pagoSinintereses").Set monto_renta_por_bienes
'		Browser("Rentas").Page("R_PrimeraCategoria").WebEdit("RPC_Txt_100").WebButton("RPC_Txt_100_Guardar").Click
'		Browser("Rentas").Page("R_PrimeraCategoria").WebEdit("RPC_Txt_100").WebButton("RPC_Txt_100_Aceptar").Click
'		Browser("Rentas").Page("R_PrimeraCategoria").WebEdit("RPC_Txt_100").WebButton("RPC_Txt_100_Cancelar").Click
'		Browser("Rentas").Page("R_PrimeraCategoria").WebEdit("RPC_Txt_100").WebButton("RPC_Txt_Cancelar_2").Click
'		txt_100 = Browser("Rentas").Page("R_PrimeraCategoria").WebEdit("RPC_Txt_100").GetROProperty("value")
'		txt_100 = Replace(txt_100,"S/ " , "")
'		txt_100 = Replace(txt_100,"," , "")
'		txt_100 = txt_100 + 0
