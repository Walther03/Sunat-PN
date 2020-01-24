'Call ExcelData() 
'DataTable.SetCurrentRow(numIterFlujo)
'Nro pestaña 7
Call Objetos_PrimeraDeterminativa()

Dim  deuda , saldo

deuda 			 	   = DataTable("deuda", 7)
saldo 			 	   = DataTable("saldo", 7)
renta_por_bienes	   = DataTable("renta_por_bienes", 7)
monto_renta_por_bienes = DataTable("monto_renta_por_bienes", 7)
Renta_Bruta_PC 	 	   = DataTable("Renta_Bruta_PC", 7)
monto_Renta_Bruta_PC   = DataTable("monto_Renta_Bruta_PC", 7)
Cesion_Gratuita  	   = DataTable("Cesion_Gratuita", 7)
monto_Cesion_Gratuita  = DataTable("monto_Cesion_Gratuita", 7)

'Call Fun_Primera_Deuda(deuda,saldo)
Call Fun_RentaPorBienes(renta_por_bienes,monto_renta_por_bienes)
Call Fun_RentaBruta(Renta_Bruta_PC,monto_Renta_Bruta_PC)
Call Fun_CesionGratuita(Cesion_Gratuita,monto_Cesion_Gratuita)
Call Fun_ValidarCampos()
Call Fun_RPC_Captura()
Call Fun_RPC_Siguiente()

WAIT 1



