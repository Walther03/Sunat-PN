'Call ExcelData() 
'DataTable.SetCurrentRow(numIterFlujo)
'Nro pestaña 8
Call Objetos_SegundaDeterminativa()

Dim  deuda , saldo , Renta_Bruta_SC , Perdida_capital , Renta_Neta_Fuente_Extranjera

deuda 						 		= DataTable("deuda", 8)
saldo 						 		= DataTable("saldo", 8)
Renta_Bruta_SC  			 		= DataTable("Renta_Bruta_SC", 8)
monto_Renta_Bruta_SC		 		= DataTable("monto_Renta_Bruta_SC", 8)
Perdida_capital 			 		= DataTable("Perdida_capital", 8)
monto_Perdida_capital		 		= DataTable("monto_Perdida_capital", 8)
Renta_Neta_Fuente_Extranjera 		= DataTable("Renta_Neta_Fuente_Extranjera", 8)
monto_Renta_Neta_Fuente_Extranjera  = DataTable("monto_Renta_Neta_Fuente_Extranjera", 8)
monto2_Renta_Neta_Fuente_Extranjera = DataTable("monto2_Renta_Neta_Fuente_Extranjera", 8)


'Call Fun_Segunda_Deuda(deuda,saldo)

Call Fun_Renta_Bruta_SC(Renta_Bruta_SC,monto_Renta_Bruta_SC)
Call ValidarTxt_353()
Call ValidarTxt_354()
Call Fun_Perdida_capital(Perdida_capital,monto_Perdida_capital)
Call Fun_Renta_Neta_Fuente_Extranjera(Renta_Neta_Fuente_Extranjera,monto_Renta_Neta_Fuente_Extranjera,monto2_Renta_Neta_Fuente_Extranjera)
Call ValidarTxt_356()
Call Fun_RSC_Captura()
Call Fun_RSC_Siguiente()

wait 2



		

		

	
