'Call ExcelData()
'DataTable.SetCurrentRow(numIterFlujo)

Call Objetos_Login()

Dim e_ruta , e_doc , e_url

'numIter = Environment.Value("ActionIteration")
e_ruta		=  Datatable.Value("ruta",1)
e_doc		=  Datatable.Value("doc",1)
e_url		=  Datatable.Value("url",1)
e_usuario	=  Datatable.Value("usuario",1)
e_contra	=  Datatable.Value("contra",1)



'Page Ingreso y seleccion
'Call Ingreso(e_ruta,e_url,e_doc,e_usuario,e_contra)
Call Fun_Login(e_ruta,e_url,e_doc,e_usuario,e_contra) @@ script infofile_;_ZIP::ssf4.xml_;_
wait 1



