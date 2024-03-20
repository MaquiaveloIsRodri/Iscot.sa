sub main
  ' ******* Funcion para dar de alta un Asiento Contable desde Excel ***********
  ' ******* Formato para generar el Excel para cada registro Contable *****
  ' 0 Cuenta
  ' 1 Fecha
  ' 2 Debe
  ' 3 Haber 
  ' 4 Detalle
  ' 5 Observacion
  ' 6 Centro de Costo	
  ' 7 Interno

  'STOP
  set xapp = getappcalipso(aworkspace.value)

  ' Nodificado: 01/03/2013 - José Fantasia.
  ' Para arranque de ISCOT.
  Set xVisualVar = VisualVarEditor ("Migracion de Asientos Contables desde Excel")
  call AddVarString (xVisualVar, "Attr1_xPath", "Path", "Parámetros","Y:\DOCUMENT\CALIPSO\migracion 27-02\")
  call AddVarString (xVisualVar, "Attr2_xNombreArchivo", "Nombre del Archivo", "Parámetros", "AsientosFuturosMarzo.xlsx")
  call AddVarString (xVisualVar, "Attr5_xNombrehoja", "Nombre de la Hoja", "Parámetros", "Hoja3")
  call AddVarString (xVisualVar, "Attr3_xDesdeRegistro", "Desde Registro nº", "Parámetros", "18")
  call AddVarString (xVisualVar, "Attr4_xHastaRegistro", "Hasta Registro nº", "Parámetros", "1046")
  xAceptar = ShowVisualVar (xVisualVar)
  if xAceptar then
	 xPath = GetValueVisualVar (xVisualVar, "Attr1_xPath", "Parámetros")
	 xNombreArchivo = GetValueVisualVar (xVisualVar,"Attr2_xNombreArchivo", "Parámetros")
	 xDesdeRegistro = GetValueVisualVar (xVisualVar,"Attr3_xDesdeRegistro", "Parámetros")
	 xHastaRegistro = GetValueVisualVar (xVisualVar,"Attr4_xHastaRegistro", "Parámetros")
 	 xnombrehoja = GetValueVisualVar (xVisualVar,"Attr5_xnombrehoja", "Parámetros")
	 xDesdeRegistro = int(xDesdeRegistro)
	 xHastaRegistro = int(xHastaRegistro)

	 'Abre el Archivo de Excel
	 set HojaExcel = CreateObject("Excel.Application")
	 xPathCompleto = xPath & xNombreArchivo
	 HojaExcel.WorkBooks.Open xPathCompleto
	 xDebugMsg = "Archivo:" & xNombreArchivo & " Desde Registro:" & xDesdeRegistro & " Hasta Registro:" & xHastaRegistro
	 SendDebug(xDebugMsg)
	 
	i = xDesdeRegistro
	'ENCABEZADO = TRUE
	errores = FALSE
	 
	set xUnidadOperativa=existebo(xapp, "uocontabilidad", "id", "{82FE1A99-19BF-4D0E-B33D-525760782AF0}", nill, false, false, "=")
	set xcompania=instanciarbo("{63E432B5-726E-4195-8F9D-2025C20D7467}", "compania", xapp.workspace)
	
	' ejercicio.
	set xEjercicio = InstanciarBO("{77063E9F-8B97-4FDA-BB9E-AC65A8104ABF}", "EJERCICIO", xapp.Workspace)
	nuevoAsiento   = ""
	errores 	   = False
	seguir 		   = True
	k			   = 0
	
	While (i <= xHastaRegistro)and (HojaExcel.WorkSheets(xnombrehoja).Cells(i, 1).Value <> "")
	   	  If HojaExcel.WorkSheets(xnombrehoja).Cells(i, 10).Value <> "Importado Correctamente" Then 
			 	 Cuenta     = Trim(HojaExcel.WorkSheets(xnombrehoja).Cells(i, 1).Value)
			 	 Fecha		= CDate(Trim(HojaExcel.WorkSheets(xnombrehoja).Cells(i, 2).Value))
				 Debe		= Trim(HojaExcel.WorkSheets(xnombrehoja).Cells(i, 3).Value)
				 Haber		= Trim(HojaExcel.WorkSheets(xnombrehoja).Cells(i, 4).Value)
				 Detalle    = Trim(HojaExcel.WorkSheets(xnombrehoja).Cells(i, 5).Value)
				 Obs    	= Trim(HojaExcel.WorkSheets(xnombrehoja).Cells(i, 6).Value)
				 CCTO    	= Trim(HojaExcel.WorkSheets(xnombrehoja).Cells(i, 7).Value)
				 CodInt		= Trim(HojaExcel.WorkSheets(xnombrehoja).Cells(i, 8).Value)
				 'stop
                 '1 IF ENCABEZADO THEN
				 '1        ENCABEZADO = FALSE
						 '1 para identificar el ejercicio 
	 					 '1 set xview = NewCompoundView( xapp , "EJERCICIO", xapp.Workspace, nil, true )
 	 					 '1 aca va la fecha de aplicacion contable
	 					 '1 xview.addfilter(newfilterspec(xview.columnfrompath("DESDEFECHA"),"<=", CDATE(Fecha))) 
	 					 '1 xview.addfilter(newfilterspec(xview.columnfrompath("HASTAFECHA"),">=", CDATE(Fecha)))
	 					 '1 xview.addfilter(newfilterspec(xview.columnfrompath("ESQUEMAOPERATIVO"),"=", xEsquemaOperativo.ID))
'	 					 xview.addfilter(newfilterspec(xview.columnfrompath("CODIGO"),"LIKE", root.childNodes(0).text&"-%"))
	 					 '1 for each xitem in xview.viewitems
		 				 '1   set xejercicio = xitem.bo
	 					 '1 next
						 '1 set xejercicio=instanciarbo("C363D86E-FE4A-4C86-9547-3DD811F6AA6A", "EJERCICIO", xapp.workspace)
				 stop
				 If nuevoAsiento <> CodInt Then
				   	k = k + 1
					If Not errores And k > 1 Then
					   call WorkSpaceCheck( xapp.WorkSpace )
					End If
					
					nuevoAsiento   = CodInt
					errores 	   = False
					seguir 		   = True
					
					'Creo el asiento	                           
	 				set asiento = CrearTRContable( "AsCon", xejercicio.codigo, cdate(Fecha), xUnidadOperativa )
     				call NoMensaje( asiento, true )
	 				asiento.fechaaplicacion = cdate(Fecha) 
	 				asiento.detalle = Detalle
	 				asiento.nota = Obs
				 Else
				 	' Mismo asiento pero hubo error en alguna cuenta:
					If errores Then
					   seguir = False
					End If
				 END IF 
				 
				 'stop
				 'set xcuenta = existebo(xapp, "cuenta", "codigo", cuenta , nill, false, false, "=")
				 If seguir Then
				    ' Instancio la cuenta según código de AIKON.
				    set xcuenta = Nothing
				    set xviewcta = NewCompoundView( xapp , "cuenta", xapp.Workspace, nil, true )
				    xviewcta.addJoin(newJoinSpec(NewColumnSpec( "CUENTA", "BOEXTENSION", "" ), NewColumnSpec( "UD_CUENTA", "ID", "" ), False))
 	 			    xviewcta.addfilter(newfilterspec(NewColumnSpec( "UD_CUENTA", "CODIGOAIKON", "" ),"=", cuenta))
					xviewcta.addfilter(newfilterspec(NewColumnSpec( "CUENTA", "activestatus", "" ),"=", "0"))
	 			    for each xitemcta in xviewcta.viewitems
		 		        set xcuenta = xitemcta.bo
	 			    next
				 
				 'controlocc = TRUE
				 'If CCTO = "" Then
		   		 '      controlocc = FALSE
	 	         'Else
				 '     set xcentrocosto = existebo(xapp, "centrocostos", "codigo", CCTO, nill, false, false, "=")
				'	  If xcentrocosto is nothing Then
   		      	'  	  	     HojaExcel.WorkSheets(xnombrehoja).Cells(i,11).Value="Centro Costo no existe " & Cenco	 	
				'  			 SendDebug("Centro costo " & Cenco & " no existe")
				'			 errores = TRUE
				'			 controlocc = FALSE
				'	  End if 
				' End If   
				 
				    ' Instancio el CC si es que tiene, según código viejo AIKON.
				    set xcentrocosto = Nothing
				    If CCTO <> "" Then
				 	   set xviewcc = NewCompoundView( xapp , "CENTROCOSTOS", xapp.Workspace, nil, true )
				 	   xviewcc.addfilter(newfilterspec(NewColumnSpec( "CENTROCOSTOS", "DETALLE", "" ),"=", CCTO))
	 			 	   for each xitemcc in xviewcc.viewitems
		 		     	   set xcentrocosto = xitemcc.bo
	 			 	   next
				    End If

				    ' creo los ítems.
				    if not xcuenta is nothing then
				        if Debe ="" Then Debe  = 0
				        if Haber ="" Then Haber  = 0
					    'stop
		   		 	    SET XITEMasiento  = CrearItemTransaccion(asiento)
	 	   			    asiento.ITEMSTRANSACCION.ADD(XITEMasiento)
 	 	   			    XITEMasiento.REFERENCIA  = xcuenta
		   			    XITEMasiento.debeoriginal.importe = CDBL(Debe)
	 	   			    XITEMasiento.haberoriginal.importe = CDBL(Haber)
		   			 ' IF controlocc Then
		    	 	  	      XITEMasiento.centrocostos  = xcentrocosto
		   			 ' End If		   
 	 	   			    XITEMasiento.detalle  = Detalle 
					    HojaExcel.WorkSheets(xnombrehoja).Cells(i,10).Value="Importado Correctamente"
		            else
		    	 	    errores = TRUE
					    asiento.workspace.rollback
					    if xcuenta is nothing then
			   			   SendDebug("Falta la cuenta contable")
  		      	  	  	   HojaExcel.WorkSheets(xnombrehoja).Cells(i,11).Value="Cuenta Contable Inexistente (código AIKON): " & Cuenta	 	
			            end if
				    end if
				 End If
          END IF
		
          HojaExcel.ActiveWorkbook.save	 		
		  I = I + 1
		  '2 stop
		  '2 IF HojaExcel.WorkSheets(xnombrehoja).Cells(i,1).Value="NUEVO" OR  HojaExcel.WorkSheets(xnombrehoja).Cells(i,1).Value="" THEN
		     'ShowBO(asiento)
  		  	 '2 xEstadoMensaje = nomensaje (asiento, True)
		     '2 If not errores Then
	    		'2 FIN = WorkSpaceCheck( XAPP.workspace )
			    '2 IF FIN = 0 THEN
		   	  	   '2 HojaExcel.WorkSheets(xnombrehoja).Cells(i,10).Value="Importado Correctamente"
				   '2 HojaExcel.WorkSheets(xnombrehoja).Cells(i,11).Value=asiento.nombre	  
	  	        '2 ELSE
		      	   '2 HojaExcel.WorkSheets(xnombrehoja).Cells(i,10).Value="Se registraron errores en la importacion"	
				   '2 asiento.workspace.rollback	
                   '2 SET asiento = NOTHING  
		        '2 END IF
			 '2 Else
	      	 	'2 HojaExcel.WorkSheets(xnombrehoja).Cells(i,10).Value="Se registraron errores en la importacion"		
	         '2 End iF 
			 '2 HojaExcel.ActiveWorkbook.save	 		
		  	 '2 I = I + 1
			 '2 ENCABEZADO = TRUE
	 		 '2 errores = FALSE
		 '2 END IF		   
	Wend

   'HojaExcel.ActiveWorkbook.save
   call WorkSpaceCheck( xapp.WorkSpace )	
   HojaExcel.Application.Quit
   Set HojaExcel = Nothing
  
  End if
		  
	MSGBOX "Proceso Finalizado!!! "	  
End sub