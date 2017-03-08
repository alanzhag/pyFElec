Attribute VB_Name = "ModuloWSFEv1"
Sub IniciarWSFEv1()
    On Error GoTo ManejoError
    Debug.Print WSFEv1.version
    
    If Not ModuloWSAA.Conectado Then
        Exit Sub
    End If
    
    If sign <> "" Then
        WSFEv1.token = token
        WSFEv1.sign = sign
    ElseIf WSAA.token And WSAA.sign <> "" Then
        WSFEv1.token = WSAA.token
        WSFEv1.sign = WSAA.sign
    End If
    
    ' CUIT del emisor (debe estar registrado en la AFIP)
    WSFEv1.Cuit = "20267565393"
    
    ' Conectar al Servicio Web de Facturación
    cache = "" ' directorio temporal (usar predeterminado)
    url_wsdl = "https://wswhomo.afip.gov.ar/wsfev1/service.asmx?WSDL" ' usar servicios1 para producción
    proxy = "" ' información de servidor intermedio (si corresponde)
    ok = WSFEv1.Conectar(cache, url_wsdl, proxy) ' homologación
    MsgBox "¡Conectado!", vbOKOnly, "WSFEv1"
    ' Llamo a un servicio nulo, para obtener el estado del servidor (opcional)
    WSFEv1.Dummy
    Debug.Print "appserver status", WSFEv1.AppServerStatus
    Debug.Print "dbserver status", WSFEv1.DbServerStatus
    Debug.Print "authserver status", WSFEv1.AuthServerStatus
    
'    ' Establezco los valores de la factura a autorizar:
'    tipo_cbte = 1
'    punto_vta = 4001
'    cbte_nro = CInt(WSFEv1.CompUltimoAutorizado(tipo_cbte, punto_vta) + 1)
'    fecha = Format(Date, "yyyymmdd")
'    concepto = 1
'    tipo_doc = 80: nro_doc = "30500010912"
'    cbt_desde = cbte_nro: cbt_hasta = cbte_nro
'    imp_total = "122.00": imp_tot_conc = "0.00": imp_neto = "100.00"
'    imp_iva = "21.00": imp_trib = "1.00": imp_op_ex = "0.00"
'    fecha_cbte = fecha
'    ': fecha_venc_pago = fecha
'    ' Fechas del período del servicio facturado (solo si concepto = 1?)
'    'fecha_serv_desde = fecha: fecha_serv_hasta = fecha
'    moneda_id = "PES": moneda_ctz = "1.000"
'
'    ok = WSFEv1.CrearFactura(concepto, tipo_doc, nro_doc, tipo_cbte, punto_vta, _
'        cbt_desde, cbt_hasta, imp_total, imp_tot_conc, imp_neto, _
'        imp_iva, imp_trib, imp_op_ex, fecha_cbte, fecha_venc_pago, _
'        fecha_serv_desde, fecha_serv_hasta, _
'        moneda_id, moneda_ctz)
'
'    ' Agrego los comprobantes asociados:
'    If False Then ' solo nc/nd
'        tipo = 19
'        pto_vta = 2
'        nro = 1234
'        ok = WSFEv1.AgregarCmpAsoc(tipo, pto_vta, nro)
'    End If
'
'    ' Agrego impuestos varios
'    id = 99
'    Desc = "Impuesto Municipal Matanza'"
'    base_imp = "100.00"
'    alic = "1.00"
'    importe = "1.00"
'    ok = WSFEv1.AgregarTributo(id, Desc, base_imp, alic, importe)
'
'    ' Agrego tasas de IVA
'    id = 5 ' 21%
'    base_imp = "100.00"
'    importe = "21.00"
'    ok = WSFEv1.AgregarIva(id, base_imp, importe)
'
'    ' Solicito CAE:
'    cae = WSFEv1.CAESolicitar()
'
'    Debug.Print "Resultado", WSFEv1.Resultado
'    Debug.Print "CAE", WSFEv1.cae
'
'    Debug.Print "Numero de comprobante:", WSFEv1.CbteNro
'
'    ' Imprimo pedido y respuesta XML para depuración (errores de formato)
'    Debug.Print WSFEv1.XmlRequest
'    Debug.Print WSFEv1.XmlResponse
'    Debug.Print "Resultado:" & WSFEv1.Resultado & " CAE: " & cae & " Venc: " & WSFEv1.Vencimiento & " Obs: " & WSFEv1.obs, vbInforma, tion + vbOKOnly
'    Call WriteFile(WSFEv1.XmlRequest, "XmlRequest.xml")
'    Call WriteFile(WSFEv1.XmlResponse, "XmlResponse.xml")
'    Call WriteFile(WSFEv1.ErrMsg, "XmlErrMsg.txt")
'    MsgBox "Resultado:" & WSFEv1.Resultado & " CAE: " & cae & " Venc: " & WSFEv1.Vencimiento & " Obs: " & WSFEv1.obs, vbInforma, tion + vbOKOnly
'    'txtResul.text = "Resultado:" & WSFEv1.Resultado & " CAE: " & cae & " Venc: " & WSFEv1.Vencimiento & " Obs: " & WSFEv1.obs & "," & vbInformation + vbOKOnly

    ' Muestro los eventos (mantenimiento programados y otros mensajes de la AFIP)
    For Each evento In WSFEv1.eventos:
        MsgBox evento, vbInformation, "Evento"
    Next
    
ManejoError:
    
    ' Si hubo error (tradicional, no controlado):
        
    ' Mostrar mensajes de Depuración en ventana de inmediato
    If Not WSAA Is Nothing Then
        If WSAA.version >= "1.02a" Then
            Debug.Print WSAA.Excepcion
            Debug.Print WSAA.Traceback
            Debug.Print WSAA.XmlRequest
            Debug.Print WSAA.XmlResponse
        End If
    End If
    If Not WSFEv1 Is Nothing Then
        If WSFEv1.version >= "1.10a" Then
            Debug.Print WSFEv1.Excepcion
            Debug.Print WSFEv1.Traceback
            Debug.Print WSFEv1.XmlRequest
            Debug.Print WSFEv1.XmlResponse
            Debug.Print WSFEv1.DebugLog()
        End If
    End If
    
    ' Si hubo error:
    
    Debug.Print Err.Description            ' descripción error afip
    Debug.Print Err.Number - vbObjectError ' codigo error afip
    
    If Err.Description <> "" Then
        Select Case MsgBox(Err.Description, vbCritical + vbRetryCancel, "Error:" & Err.Number - vbObjectError & " en " & Err.Source)
            Case vbRetry
                Debug.Assert False
                Resume
            Case vbCancel
                Debug.Print Err.Description
        End Select
    End If
End Sub

Public Function Conectado() As Boolean
On Error GoTo ManejoError

    WSFEv1.Dummy
    Conectado = True
    Exit Function
    
ManejoError:
    MsgBox "No se ha iniciado el WSFEv1. Por favor dirijase a WebService->WSFEv1", vbOKOnly + vbExclamation, "Error"
    Conectado = False
End Function

