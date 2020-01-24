'Call ExcelData()
'DataTable.SetCurrentRow(numIterFlujo)

Call Objetos_Otros_Ingresos()

Dim ingresos , renta_exonerada , renta_inafectas , resultado_actividades , dividendo_percibidos , fuente_extranjera

ingresos					= DataTable("ingresos", 6)
renta_exonerada 			= DataTable("renta_exonerada", 6)
renta_inafectas				= DataTable("renta_inafectas", 6)
resultado_actividades		= DataTable("resultado_actividades", 6)
monto_resultado_actividades = DataTable("monto_resultado_actividades", 6)
dividendo_percibidos		= DataTable("dividendo_percibidos", 6)
monto_dividendo_percibidos  = DataTable("monto_dividendo_percibidos", 6)
fuente_extranjera			= DataTable("fuente_extranjera", 6)
monto_fuente_extranjera		= DataTable("monto_fuente_extranjera", 6)



'Call Fun_OtrosIngreso(ingresos)

Call Fun_renta_exonerada(renta_exonerada)
Call Fun_renta_inafectas(renta_inafectas)
Call Fun_resultado_actividades(resultado_actividades,monto_resultado_actividades)
Call Fun_dividendo_percibidos(dividendo_percibidos,monto_dividendo_percibidos)
Call Fun_fuente_extranjera(fuente_extranjera,monto_fuente_extranjera)
 @@ script infofile_;_ZIP::ssf2.xml_;_
Call Fun_OTI_Captura()
Call Fun_OTI_Siguiente()
wait 3

Browser("Rentas").Page("Otros Ingresos").WebButton("OTI_Aceptar").Click


