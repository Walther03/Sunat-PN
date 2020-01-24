'Call ExcelData()
'DataTable.SetCurrentRow(numIterFlujo)

Call Objetos_Alquileres()

Dim alquiler , cant_alquileres , tipo_doc , num_doc, monto_alquiler , nro_meses , tipo_bien , bien , llenar


alquiler	    		= DataTable("alquiler", 4)
cant_alquileres			= DataTable("cant_alquileres", 4)

Call Fun_Alqui(alquiler)


If alquiler = "Si" Then
	
	For Iterator = 1 To cant_alquileres Step 1
	
		If Iterator = 1 Then
		
			tipo_doc		 		= DataTable("tipo_doc", 4)
			num_doc 				= DataTable("num_doc", 4)
			monto_alquiler			= DataTable("monto_alquiler", 4)
			nro_meses				= DataTable("nro_meses", 4)
			tipo_bien				= DataTable("tipo_bien", 4)
			bien					= DataTable("bien", 4)
			llenar					= DataTable("llenar", 4)
				
			Call Fun_Alquileres(tipo_doc,num_doc,monto_alquiler,nro_meses,tipo_bien,bien,llenar)
				
		Else	
		
			tipo_doc		 		= DataTable("tipo_doc"&Iterator, 4)
			num_doc 				= DataTable("num_doc"&Iterator, 4)
			monto_alquiler			= DataTable("monto_alquiler"&Iterator, 4)
			nro_meses				= DataTable("nro_meses"&Iterator, 4)
			tipo_bien				= DataTable("tipo_bien"&Iterator, 4)
			bien					= DataTable("bien"&Iterator, 4)
			llenar					= DataTable("llenar"&Iterator, 4)
		
			Call Fun_Alquileres(tipo_doc,num_doc,monto_alquiler,nro_meses,tipo_bien,bien,llenar)
		
		
		End If	
	Next

End If


Call Fun_ALQ_Captura()
Call Fun_ALQ_Siguiente()
wait 1


 @@ script infofile_;_ZIP::ssf3.xml_;_
