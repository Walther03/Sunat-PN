'Call ExcelData()
'DataTable.SetCurrentRow(numIterFlujo)

Call Objetos_Determinacion_Deuda()

Dim existe_dev_app , saldo , dev_apli , interes_monetario , monto_interes , tipo_renta


existe_dev_app		= DataTable("existe_dev_app", 10)
saldo_primera		= DataTable("saldo_primera", 10)
saldo_segunda		= DataTable("saldo_segunda", 10)
saldo_fuente		= DataTable("saldo_fuente", 10)
dev_apli			= DataTable("dev_apli", 10)
tipo_renta 			= DataTable("tipo_renta", 2)
importe_primera		= DataTable("importe_primera", 10)
importe_segunda		= DataTable("importe_segunda", 10)
importe_trabajo		= DataTable("importe_trabajo", 10)

txt_127 			= DataTable("txt_127", 10)
monto_txt_127 		= DataTable("monto_txt_127", 10)








If txt_127 = "Si" Then
		
	Browser("Rentas").Page("Det_Deuda").WebEdit("S/ 0").Click
	Browser("Rentas").Page("Det_Deuda").WebButton("Agregar_2").Click @@ script infofile_;_ZIP::ssf12.xml_;_
	Browser("Rentas").Page("Det_Deuda").WebButton("WebButton").Click @@ script infofile_;_ZIP::ssf13.xml_;_
	Browser("Rentas").Page("Det_Deuda").WebButton("WebButton_2").Click @@ script infofile_;_ZIP::ssf14.xml_;_
	Browser("Rentas").Page("Det_Deuda").WebElement("ene_2").Click @@ script infofile_;_ZIP::ssf15.xml_;_
	Browser("Rentas").Page("Det_Deuda").WebList("numFormulario").Select "0616" @@ script infofile_;_ZIP::ssf16.xml_;_
	Browser("Rentas").Page("Det_Deuda").WebEdit("numOrden").Set "21" @@ script infofile_;_ZIP::ssf17.xml_;_
	Browser("Rentas").Page("Det_Deuda").WebButton("WebButton_3").Click @@ script infofile_;_ZIP::ssf18.xml_;_
	Browser("Rentas").Page("Det_Deuda").WebElement("3").Click @@ script infofile_;_ZIP::ssf19.xml_;_
	Browser("Rentas").Page("Det_Deuda").WebEdit("saldoFavor").Set monto_txt_127
	Browser("Rentas").Page("Det_Deuda").WebEdit("pagoSI").Set monto_txt_127 @@ script infofile_;_ZIP::ssf20.xml_;_
	Browser("Rentas").Page("Det_Deuda").WebButton("Guardar").Click @@ script infofile_;_ZIP::ssf21.xml_;_
	Browser("Rentas").Page("Det_Deuda").WebButton("DD_Aceptar").Click @@ script infofile_;_ZIP::ssf22.xml_;_
	Browser("Rentas").Page("Det_Deuda").WebButton("Cancelar").Click @@ script infofile_;_ZIP::ssf23.xml_;_
	
End If


 
 If existe_dev_app = "Si" Then 
 
 	If not saldo_primera = "" Then
		Browser("Rentas").Page("Det_Deuda").WebEdit("DD_Saldofavor_1").Set saldo_primera
	End If

	If not saldo_segunda = "" Then
		Browser("Rentas").Page("Det_Deuda").WebEdit("DD_Saldofavor_2").Set saldo_segunda
	End If

	If not saldo_fuente = "" Then
		Browser("Rentas").Page("Det_Deuda").WebEdit("DD_Saldofavor_3").Set saldo_fuente
	End If
 
 
 
 
 
 
	Select Case dev_apli
			Case "Devolucion"
				
				Dim dev1 , dev2
				dev1 = Browser("Rentas").Page("Det_Deuda").WebRadioGroup("DD_devolucionAplicacion1").GetROProperty("disabled") @@ script infofile_;_ZIP::ssf31.xml_;_
				dev2 = Browser("Rentas").Page("Det_Deuda").WebRadioGroup("DD_devolucionAplicacion2").GetROProperty("disabled")
				
				If dev1 = "0" Then
					Browser("Rentas").Page("Det_Deuda").WebRadioGroup("DD_devolucionAplicacion1").Select "#0"
				End If
				
				If dev2 = "0" Then
					Browser("Rentas").Page("Det_Deuda").WebRadioGroup("DD_devolucionAplicacion2").Select "#0"
				End If
				

				
				Set shell = CreateObject("Wscript.Shell") 
					shell.SendKeys "{ENTER}"
'				
				wait 2
				
			Case "Aplicacion"
				
			
				dev1 = Browser("Rentas").Page("Det_Deuda").WebRadioGroup("DD_devolucionAplicacion1").GetROProperty("disabled") @@ script infofile_;_ZIP::ssf31.xml_;_
				dev2 = Browser("Rentas").Page("Det_Deuda").WebRadioGroup("DD_devolucionAplicacion2").GetROProperty("disabled")
				
				If dev1 = "0" Then
					Browser("Rentas").Page("Det_Deuda").WebRadioGroup("DD_devolucionAplicacion1").Select "#1"
				End If
				
				If dev2 = "0" Then
					Browser("Rentas").Page("Det_Deuda").WebRadioGroup("DD_devolucionAplicacion2").Select "#1"
				End If
			
				Set shell = CreateObject("Wscript.Shell") 
					shell.SendKeys "{ENTER}"
'				
				wait 2
				
		End Select
		
 		
 End If

If not importe_primera = "" Then
	Browser("Rentas").Page("Det_Deuda").WebEdit("S/").Set importe_primera
End If

If not importe_segunda = "" Then
	Browser("Rentas").Page("Det_Deuda").WebEdit("S/").Set importe_segunda
End If

If not importe_trabajo = "" Then
	Browser("Rentas").Page("Det_Deuda").WebEdit("S/_3").Set importe_trabajo
End If




'Call MongoDb1(montoDB1)

Call Fun_Det_Captura()
wait 3

Browser("Rentas").Page("Det_Deuda").WebButton("DD_Presente/Pague").Click
Browser("Rentas").Page("Det_Deuda").WebButton("DD_Aceptar").Click
 @@ script infofile_;_ZIP::ssf32.xml_;_


	
