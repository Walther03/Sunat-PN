'Call ExcelData()
'DataTable.SetCurrentRow(numIterFlujo)

Call Objetos_Atribuciones()

Dim tipo_gasto, tipo_vinculo , doc_vinculo , tipo_doc , nro_doc , apellidos_nombres
Dim dia , mes , anio , correo , fijo , celular

tipo_gasto 			= DataTable("tipo_gasto", 5)
tipo_vinculo 		= DataTable("tipo_vinculo", 5)
doc_vinculo			= DataTable("doc_vinculo", 5)
nro_doc_vinculo		= DataTable("nro_doc_vinculo", 5)
tipo_doc	 		= DataTable("tipo_doc", 5)
nro_doc 			= DataTable("nro_doc", 5)
apellidos_nombres 	= DataTable("apellidos_nombres", 5)
dia					= DataTable("dia", 5)
mes					= DataTable("mes", 5)
anio				= DataTable("anio", 5)
correo				= DataTable("correo", 5)
fijo				= DataTable("fijo", 5)
celular				= DataTable("celular", 5)

Call Fun_Atribuciones(tipo_gasto,tipo_vinculo,doc_vinculo,nro_doc_vinculo,tipo_doc,nro_doc,apellidos_nombres,correo,fijo,celular)
Call Fun_ATR_Siguiente()
WAIT 1


