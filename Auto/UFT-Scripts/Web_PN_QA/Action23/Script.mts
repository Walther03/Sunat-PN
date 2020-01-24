'Call ExcelData() 
'DataTable.SetCurrentRow(numIterFlujo)

'Nro pestaña 9
Call Objetos_TrabajoDeterminativa()

Dim  deuda , saldo

deuda 							= DataTable("deuda", 9)
saldo 							= DataTable("saldo", 9)
renta_bruta_trabajo 			= DataTable("renta_bruta_trabajo", 9)
monto_renta_bruta_trabajo 		= DataTable("monto_renta_bruta_trabajo", 9)
monto2_renta_bruta_trabajo 		= DataTable("monto_renta_bruta_trabajo", 9)
renta_cuarta 					= DataTable("renta_cuarta", 9)
monto_renta_cuarta				= DataTable("monto_renta_cuarta", 9)
monto2_renta_cuarta				= DataTable("monto2_renta_cuarta", 9)
renta_quinta 					= DataTable("renta_quinta", 9)
monto_renta_quinta				= DataTable("monto_renta_quinta", 9)
deduccion_itf 					= DataTable("deduccion_itf", 9)
monto_deduccion_itf				= DataTable("monto_deduccion_itf", 9)
monto2_deduccion_itf			= DataTable("monto2_deduccion_itf", 9)
deduccion_dona 					= DataTable("deduccion_dona", 9)
monto_deduccion_dona			= DataTable("monto_deduccion_dona", 9)
renta_fuente_extranjera 		= DataTable("renta_fuente_extranjera", 9)
monto_renta_fuente_extranjera 	= DataTable("monto_renta_fuente_extranjera", 9)

'Call Fun_TrabajoDeterminativa(deuda,saldo)
Call Fun_renta_bruta_trabajo(renta_bruta_trabajo,monto_renta_bruta_trabajo,monto2_renta_bruta_trabajo) @@ script infofile_;_ZIP::ssf2.xml_;_
Call Fun_Validartxt107()
Call Fun_Validartxt_508()
Call Fun_renta_cuarta(renta_cuarta,monto_renta_cuarta,monto2_renta_cuarta)
Call Fun_Validartxt_509()
Call Fun_renta_quinta(renta_quinta,monto_renta_quinta)
Call Fun_Validartxt_510()
Call Fun_Validartxt_511()
Call Fun_Validartxt_514()
Call Fun_Validartxt_512()
Call Fun_deduccion_itf(deduccion_itf,monto_deduccion_itf,monto2_deduccion_itf)
Call Validartxt_513()
Call Fun_renta_fuente_extranjera(renta_fuente_extranjera,monto_renta_fuente_extranjera)
Call Fun_deduccion_dona(deduccion_dona,monto_deduccion_dona)
Call Validartxt_517()
Call Fun_RT_Captura()
Call Fun_RT_Siguiente()
wait 1








