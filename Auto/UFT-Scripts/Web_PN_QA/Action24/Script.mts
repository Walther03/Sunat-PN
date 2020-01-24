Call ExcelData()
row = Datatable.getsheet(1).Getrowcount 
SystemUtil.Run "C:\Users\Administrador\AppData\Local\Programs\nosqlbooster4mongo\NoSQLBooster for MongoDB.exe"
wait 5
 @@ hightlight id_;_65822_;_script infofile_;_ZIP::ssf67.xml_;_
 For Iterator = 1 To row Step 1
 		
 		DataTable.SetCurrentRow(Iterator)
 		
 		e_doc		= Datatable.Value("doc",1)
' 		tipo_renta 	= DataTable("tipo_renta", 2)
' 		monto		= DataTable("Monto", 12)
 		
' 		DataTable("Ruc", 12) = e_doc
 		 		
		
 		If Window("NoSQLBooster for MongoDB").InsightObject("InsightObject").Exist(3)  = true then
 		
			Window("NoSQLBooster for MongoDB").InsightObject("InsightObject").Click
			wait 2

		Else 
			Window("NoSQLBooster for MongoDB").InsightObject("InsightObject_5").Click @@ hightlight id_;_8_;_script infofile_;_ZIP::ssf97.xml_;_
 		wait 2
 		
 		End If
		Window("NoSQLBooster for MongoDB").InsightObject("cmd").Click @@ hightlight id_;_4_;_script infofile_;_ZIP::ssf88.xml_;_
		wait 2
		
		Window("NoSQLBooster for MongoDB").Type "db.presentaciones.find({numRuc :'"&e_doc&"', indPagRea : '1'})"
		wait 2
		
'		
'		Select Case tipo_renta
'		
'			Case "Renta de Capital Primera Categoria" , "Renta de Capital Primera/Segunda Categoria" , "Renta de Capital Primera/Trabajo Fuente Extranjera" ,  "Renta de Capital Primera/Segunda Categoria/Trabajo Fuente"	
'					
'				Window("NoSQLBooster for MongoDB").Type "db.presentaciones.find({numRuc :'"&e_doc&"', indPagRea : '1'})"
''				Window("NoSQLBooster for MongoDB").Type "db.preDeclaraciones.find({numRuc : '"&e_doc&"',"
''	 			Window("NoSQLBooster for MongoDB").Type " 'declaracion.seccDeterminativa.rentaPrimera.resumenPrimera.mtoCas164' : NumberDecimal('"&monto&"')})"
' 				wait 2
' 		
'			Case "Renta de Capital Segunda Categoria" , "Renta de Capital Segunda Categoria/Trabajo Fuente Extranjera"
'					
'					Window("NoSQLBooster for MongoDB").Type "db.preDeclaraciones.find({numRuc : '"&e_doc&"',"
'	 				Window("NoSQLBooster for MongoDB").Type "'declaracion.seccDeterminativa.rentaSegunda.resumenSegunda.mtoCas365' : NumberDecimal('" &monto&"')})"
' 					wait 2
'			Case "Renta de Trabajo y/o Fuente Extranjera" 
'		
'					Window("NoSQLBooster for MongoDB").Type "db.preDeclaraciones.find({numRuc : '"&e_doc&"',"
'	 				Window("NoSQLBooster for MongoDB").Type "'declaracion.seccDeterminativa.rentaTrabajo.resumenTrabajo.mtoCas146' : NumberDecimal('" &monto&"')})"
' 					wait 2
'		
'		
'		End Select
		
		Window("NoSQLBooster for MongoDB").InsightObject("btn_run").Click
		wait 3
		
		If Window("NoSQLBooster for MongoDB").InsightObject("no_found").Exist(3) = true Then
		
			Window("NoSQLBooster for MongoDB").CaptureBitmap RutaEvidencias()  & numIter &  "MongoDb.png", True
			imagenToWord "MongoDb Consulta : El Ruc "&e_doc&" no termino el flujo ", RutaEvidencias()  & numIter &  "MongoDb.png"
			wait 2
			DataTable("Resultado", 12) = "Flujo Incorrecto"
			
			Window("NoSQLBooster for MongoDB").InsightObject("cmd").Click
			wait 2
		
			Window("NoSQLBooster for MongoDB").Type "db.preDeclaraciones.remove({'numRuc':'"&e_doc&"'})"
			
						
			Window("NoSQLBooster for MongoDB").InsightObject("btn_run").Click
			
			wait 2
			Window("NoSQLBooster for MongoDB").CaptureBitmap RutaEvidencias()  & numIter &  "MongoDbeliminar.png", True
			imagenToWord "Se elimina de MongoDb ", RutaEvidencias()  & numIter &  "MongoDbeliminar.png"
			wait 2
			
			
		else
			
			Window("NoSQLBooster for MongoDB").CaptureBitmap RutaEvidencias()  & numIter &  "MongoDb.png", True
			imagenToWord "MongoDb Consulta : El Ruc "&e_doc&" si termino el flujo ", RutaEvidencias()  & numIter &  "MongoDb.png"
			wait 2
			DataTable("Resultado", 12) = "Flujo Correcto"
			Window("NoSQLBooster for MongoDB").InsightObject("cmd").Click
			wait 2
		
			Window("NoSQLBooster for MongoDB").Type "db.preDeclaraciones.remove({'numRuc':'"&e_doc&"'})"
			
			Window("NoSQLBooster for MongoDB").InsightObject("btn_run").Click
			
			wait 3
			
			Window("NoSQLBooster for MongoDB").CaptureBitmap RutaEvidencias()  & numIter &  "MongoDbeliminar.png", True
			imagenToWord "Se elimina de MongoDb ", RutaEvidencias()  & numIter &  "MongoDbeliminar.png"
			wait 2
			Window("NoSQLBooster for MongoDB").InsightObject("cmd").Click
			wait 2
		
			Window("NoSQLBooster for MongoDB").Type "db.presentaciones.remove({'numRuc':'"&e_doc&"'})"
			
						
			Window("NoSQLBooster for MongoDB").InsightObject("btn_run").Click
			
		End If
 @@ hightlight id_;_2_;_script infofile_;_ZIP::ssf86.xml_;_
		
 	
 Next

While Window("NoSQLBooster for MongoDB").InsightObject("InsightObject_2").Exist = true 
	Window("NoSQLBooster for MongoDB").InsightObject("InsightObject_2").Click
	
Wend
	Window("NoSQLBooster for MongoDB").InsightObject("btn_run").Click
	Window("NoSQLBooster for MongoDB").Close


 @@ hightlight id_;_7_;_script infofile_;_ZIP::ssf96.xml_;_

 @@ hightlight id_;_8_;_script infofile_;_ZIP::ssf81.xml_;_
 @@ hightlight id_;_4_;_script infofile_;_ZIP::ssf94.xml_;_
 @@ hightlight id_;_8_;_script infofile_;_ZIP::ssf95.xml_;_


'db.preDeclaraciones.remove({"numRuc":"10166738631"})


