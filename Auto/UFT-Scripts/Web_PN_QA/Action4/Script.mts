'Call ExcelData()

'DataTable.SetCurrentRow(numIterFlujo)

Call Objetos_Condominos()

Dim incluye_Condominio , cant_participantes , tipo_doc , num_doc, participacion ,partida_registral , valor_bien

incluye_Condominio	    = DataTable("incluye_Condominio", 3)
cant_participantes 	 	= DataTable("cant_participantes", 3)

Call Fun_Condomi(incluye_Condominio)

Dim IteratorParti 
If not Cant_Participantes = "" Then
	For Iterator = 1 To Cant_Participantes Step 1
		IteratorParti = Iterator
		If Iterator = 1 Then
			tipo_doc		 	= DataTable("tipo_doc", 3) @@ hightlight id_;_12_;_script infofile_;_ZIP::ssf2.xml_;_
			num_doc 			= DataTable("num_doc", 3)
			participacion		= DataTable("participacion", 3)
			partida_registral	= DataTable("partida_registral", 3)
			valor_bien 			= DataTable("valor_bien", 3)
					
			Call Participantes( IteratorParti , tipo_doc , num_doc , participacion , partida_registral , valor_bien)
					
		Else	
			tipo_doc		 	= DataTable("tipo_doc"&Iterator, 3)
			num_doc 			= DataTable("num_doc"&Iterator, 3)
			participacion		= DataTable("participacion"&Iterator, 3)
			partida_registral	= DataTable("partida_registral"&Iterator, 3)
			valor_bien 			= DataTable("valor_bien"&Iterator, 3)
			Call Participantes( IteratorParti , tipo_doc , num_doc , participacion , partida_registral , valor_bien)
			
			
	End If	
Next
End If


Call Fun_Captura()
Call Fun_CD_Siguiente()

wait 1



