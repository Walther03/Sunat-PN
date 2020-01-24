'LLama a excel
Call ExcelData()
row = Datatable.getsheet(1).Getrowcount 



For Iterator = 1 To row Step 1	
	
	
	Call ExcelData()
	DataTable.SetCurrentRow(Iterator)
	flujo   = DataTable("Flujos", 1)
	Call NomFlujo(flujo)
	Call ExcelData1(flujo)
	
	NumIterValue(Iterator)
	
	For Iterator2 = 1 To 1 Step 1
		
		RunAction "Login", oneIteration
		If FinishFor = "Si" Then
			FinishFor = "No"
			Exit for 
		End If
	
		Call Objetos_Tipo_Declaracion()

'		DataTable.SetCurrentRow(numIterFlujo)
		Dim declaracion, tipo_renta , tipo_declaracion , parti_conyugal , tipo_doc_conyu , doc_conyugal

		declaracion 		= DataTable("declaracion", 2)
		tipo_renta 			= DataTable("tipo_renta", 2)
		tipo_declaracion	= DataTable("tipo_declaracion", 2)
		parti_conyugal 		= DataTable("parti_conyugal", 2)
		tipo_doc_conyu 		= DataTable("tipo_doc_conyu", 2)
		doc_conyugal 		= DataTable("doc_conyugal", 2)

 
		Call Fun_declara(declaracion)
		Call Fun_TipoRenta(tipo_renta)

		If not tipo_renta = "Renta de Trabajo y/o Fuente Extranjera" Then

			Call Fun_TipoDeclaracion(tipo_declaracion)
	
		End If

		wait 1
		Call Fun_Rent_Obtenidas(parti_conyugal,tipo_doc_conyu,doc_conyugal) 

		Call Fun_TipoDeclaracion_Captura()
		Call Fun_Tipo_Siguiente()

		Select Case tipo_renta
	
'			1 - 1 y 2 Categoria
			Case "Renta de Capital Primera Categoria" , "Renta de Capital Primera/Segunda Categoria"
				RunAction "Condominios", oneIteration
				If FinishFor = "Si" Then
					FinishFor = "No"
					Exit for 
				End If
				RunAction "Alquileres_pagados", oneIteration
				If FinishFor = "Si" Then
					FinishFor = "No"
					Exit for 
				End If
				RunAction "Otros_Ingresos", oneIteration
				If FinishFor = "Si" Then
					FinishFor = "No"
					Exit for 
				End If
				RunAction "Rentas_Primera_Categoría", oneIteration
				If FinishFor = "Si" Then
					FinishFor = "No"
					Exit for 
				End If
				If tipo_renta = "Renta de Capital Primera/Segunda Categoria" Then
					RunAction "Rentas_Segunda_Categoría", oneIteration
					If FinishFor = "Si" Then
						FinishFor = "No"
						Exit for 
					End If
				End If
				
'			2 Categoria	Y Renta Trabajo
			Case "Renta de Capital Segunda Categoria"	,  "Renta de Capital Segunda Categoria/Trabajo Fuente Extranjera"
				RunAction "Alquileres_pagados", oneIteration
				If FinishFor = "Si" Then
					FinishFor = "No"
					Exit for 
				End If
				If tipo_renta = "Renta de Capital Segunda Categoria/Trabajo Fuente Extranjera" Then
					RunAction "Atribuciones", oneIteration
					If FinishFor = "Si" Then
						FinishFor = "No"
						Exit for 
					End If
				End If
				RunAction "Otros_Ingresos", oneIteration
				If FinishFor = "Si" Then
					FinishFor = "No"
					Exit for 
				End If
				RunAction "Rentas_Segunda_Categoría", oneIteration
				If FinishFor = "Si" Then
					FinishFor = "No"
					Exit for 
				End If
				If tipo_renta = "Renta de Capital Segunda Categoria/Trabajo Fuente Extranjera" Then
					RunAction "Rentas_Trabajo", oneIteration
					If FinishFor = "Si" Then
						FinishFor = "No"
						Exit for 
					End If
				End If
		
'			Renta Trabajo
			Case "Renta de Trabajo y/o Fuente Extranjera" 
				RunAction "Alquileres_pagados", oneIteration	
				If FinishFor = "Si" Then
					FinishFor = "No"
					Exit for 
				End If
				RunAction "Atribuciones", oneIteration
				If FinishFor = "Si" Then
					FinishFor = "No"
					Exit for 
				End If
				RunAction "Otros_Ingresos", oneIteration
				If FinishFor = "Si" Then
					FinishFor = "No"
					Exit for 
				End If
				RunAction "Rentas_Trabajo", oneIteration
				If FinishFor = "Si" Then
					FinishFor = "No"
					Exit for 
				End If
		
		
'			1 categoria y renta trabajo
			Case "Renta de Capital Primera/Trabajo Fuente Extranjera"	
				RunAction "Condominios", oneIteration
				If FinishFor = "Si" Then
					FinishFor = "No"
					Exit for 
				End If
				RunAction "Alquileres_pagados", oneIteration
				If FinishFor = "Si" Then
					FinishFor = "No"
					Exit for 
				End If
				RunAction "Atribuciones", oneIteration
				If FinishFor = "Si" Then
					FinishFor = "No"
					Exit for 
				End If
				RunAction "Otros_Ingresos", oneIteration
				If FinishFor = "Si" Then
					FinishFor = "No"
					Exit for 
				End If
				RunAction "Rentas_Primera_Categoría", oneIteration
				If FinishFor = "Si" Then
					FinishFor = "No"
					Exit for 
				End If
				RunAction "Rentas_Trabajo", oneIteration
				If FinishFor = "Si" Then
					FinishFor = "No"
					Exit for 
				End If
	
			Case "Renta de Capital Primera/Segunda Categoria/Trabajo Fuente"
				RunAction "Condominios", oneIteration
				If FinishFor = "Si" Then
					FinishFor = "No"
					Exit for 
				End If
				RunAction "Alquileres_pagados", oneIteration
				If FinishFor = "Si" Then
					FinishFor = "No"
					Exit for 
				End If
				RunAction "Atribuciones", oneIteration
				If FinishFor = "Si" Then
					FinishFor = "No"
					Exit for 
				End If
				RunAction "Otros_Ingresos", oneIteration
				If FinishFor = "Si" Then
					FinishFor = "No"
					Exit for 
				End If
				RunAction "Rentas_Primera_Categoría", oneIteration
				If FinishFor = "Si" Then
					FinishFor = "No"
					Exit for 
				End If
				RunAction "Rentas_Segunda_Categoría", oneIteration
				If FinishFor = "Si" Then
					FinishFor = "No"
					Exit for 
				End If
				RunAction "Rentas_Trabajo", oneIteration
				If FinishFor = "Si" Then
					FinishFor = "No"
					Exit for 
				End If
		
	
		End Select


		RunAction "Determinación de la Deuda", oneIteration
		If FinishFor = "Si" Then
			FinishFor = "No"
			Exit for 
		End If
		
		importe_primera		= DataTable("importe_primera", 10)
		importe_segunda		= DataTable("importe_segunda", 10)
		importe_trabajo		= DataTable("importe_trabajo", 10)
		If not importe_primera = "" or  not importe_segunda = "" or not importe_trabajo = "" Then
			RunAction "Pagos", oneIteration
			If FinishFor = "Si" Then
			FinishFor = "No"
			Exit for 
			End If
		End If
		
		
		
		RunAction "Resumen", oneIteration
		If FinishFor = "Si" Then
			FinishFor = "No"
			Exit for 
		End If
	
		Next
	
Next


'RunAction "VerificarOk", oneIteration

ExitAction
ExitTest

