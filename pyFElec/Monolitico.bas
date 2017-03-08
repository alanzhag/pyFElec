Attribute VB_Name = "Monolitico"
Sub Main()
    Dim WSFE As Object
    'Dim WSAA As Object,
    On Error GoTo ManejoError
    
    ' Crear objeto interface Web Service Autenticación y Autorización
    Set WSAA = CreateObject("WSAA")
    
    ' Generar un Ticket de Requerimiento de Acceso (TRA)
    tra = WSAA.CreateTRA()
    Debug.Print tra
    
    ' Especificar la ubicacion de los archivos certificado y clave privada
    'Path = CurDir() + "\"
    Path = WSAA.InstallDir + "\"
    ' Certificado: certificado es el firmado por la AFIP
    ' ClavePrivada: la clave privada usada para crear el certificado
    CERTIFICADO = "reingart.crt" ' certificado de prueba
    ClavePrivada = "reingart.key" ' clave privada de prueba
    
    
    ' Generar el mensaje firmado (CMS)
    cms = WSAA.SignTRA(tra, Path + CERTIFICADO, Path + ClavePrivada)
    Debug.Print cms
    
    ' Llamar al web service para autenticar:
    'TA = WSAA.CallWSAA(cms, "https://wsaa.afip.gov.ar/ws/services/LoginCms") ' Hologación
    TA = WSAA.CallWSAA(cms, "https://wsaahomo.afip.gov.ar/ws/services/LoginCms") ' Producción

    ' Imprimir el ticket de acceso, ToKen y Sign de autorización
    Debug.Print TA
    Debug.Print "Token:", WSAA.token
    Debug.Print "Sign:", WSAA.sign
    
    ' Una vez obtenido, se puede usar el mismo token y sign por 6 horas
    ' (este período se puede cambiar)
    ' Crear objeto interface Web Service de Factura Electrónica
    Set WSFE = CreateObject("WSFEv1")
    ' Setear tocken y sing de autorización (pasos previos)
    WSFE.token = WSAA.token
    WSFE.sign = WSAA.sign
    
    ' CUIT del emisor (debe estar registrado en la AFIP)
    WSFE.Cuit = "20267565393"
    
    ' Conectar al Servicio Web de Facturación
    ok = WSFE.Conectar() ' homologación
    'ok = WSFE.Conectar("https://servicios1.afip.gov.ar/feliz") ' homologación
    'ok = WSFE.Conectar("https://servicios1.afip.gov.ar/wsfe/service.asmx") ' producción

    ' Llamo a un servicio nulo, para obtener el estado del servidor (opcional)
    WSFE.Dummy
    Debug.Print "appserver status", WSFE.AppServerStatus
    Debug.Print "dbserver status", WSFE.DbServerStatus
    Debug.Print "authserver status", WSFE.AuthServerStatus
    
    ' Recupera cantidad máxima de registros (opcional)
    qty = WSFE.RecuperarQty()
    
    ' Recupera último número de secuencia ID
    LastId = WSFE.UltNro()
    
    ' Recupero último número de comprobante para un punto de venta y tipo (opcional)
    tipo_cbte = 1: punto_vta = 1
    LastCBTE = WSFE.RecuperaLastCMP(punto_vta, tipo_cbte)
    
    ' Establezco los valores de la factura o lote a autorizar:
    fecha = Format(Date, "yyyymmdd")
    id = LastId + 1: presta_serv = 1
    tipo_doc = 80: nro_doc = "23111111113"
    cbt_desde = LastCBTE + 1: cbt_hasta = LastCBTE + 1
    imp_total = "121.00": imp_tot_conc = "0.00": imp_neto = "100.00"
    impto_liq = "21.00": impto_liq_rni = "0.00": imp_op_ex = "0.00"
    fecha_cbte = fecha: fecha_venc_pago = fecha
    ' Fechas del período del servicio facturado (solo si presta_serv = 1)
    fecha_serv_desde = fecha: fecha_serv_hasta = fecha
    
    ' Llamo al WebService de Autorización para obtener el CAE
    cae = WSFE.Aut(id, presta_serv, _
        tipo_doc, nro_doc, tipo_cbte, punto_vta, _
        cbt_desde, cbt_hasta, imp_total, imp_tot_conc, imp_neto, _
        impto_liq, impto_liq_rni, imp_op_ex, fecha_cbte, fecha_venc_pago, _
        fecha_serv_desde, fecha_serv_hasta) ' si presta_serv = 0 no pasar estas fechas
    
    Debug.Print "Vencimiento ", WSFE.Vencimiento ' Fecha de vencimiento o vencimiento de la autorización
    Debug.Print "Resultado: ", WSFE.Resultado ' A=Aceptado, R=Rechazado
    Debug.Print "Motivo de rechazo o advertencia", WSFE.Motivo ' 00= No hay error
    Debug.Print "Reprocesado?", WSFE.Reproceso ' S=Si, N=No
    
    ' Verifico que no haya rechazo o advertencia al generar el CAE
    If cae = "" Then
        MsgBox "La página esta caida o la respuesta es inválida"
    ElseIf cae = "NULL" Or WSFE.Resultado <> "A" Then
        MsgBox "No se asignó CAE (Rechazado). Motivos: " & WSFE.Motivo, vbInformation + vbOKOnly
    ElseIf WSFE.Motivo <> "NULL" And WSFE.Motivo <> "00" Then
        MsgBox "Se asignó CAE pero con advertencias. Motivos: " & WSFE.Motivo, vbInformation + vbOKOnly
    End If
    
    ' Imprimo respuesta XML para depuración (errores de formato)
    Debug.Print WSFE.XmlResponse
    
    MsgBox "QTY: " & qty & vbCrLf & "LastId: " & LastId & vbCrLf & "LastCBTE:" & LastCBTE & vbCrLf & "CAE: " & cae, vbInformation + vbOKOnly
    MsgBox "Número: " & WSFE.CbtDesde & " - " & WSFE.CbtHasta & vbCrLf & _
           "Fecha: " & WSFE.FechaCbte & vbCrLf & _
           "Total: " & WSFE.ImpTotal & vbCrLf & _
           "Neto: " & WSFE.ImpNeto & vbCrLf & _
           "Iva: " & WSFE.ImptoLiq
    Exit Sub
ManejoError:
    ' Si hubo error:
    Debug.Print Err.Description            ' descripción error afip
    Debug.Print Err.Number - vbObjectError ' codigo error afip
    Select Case MsgBox(Err.Description, vbCritical + vbRetryCancel, "Error:" & Err.Number - vbObjectError & " en " & Err.Source)
        Case vbRetry
            Debug.Assert False
            Resume
        Case vbCancel
            Debug.Print Err.Description
    End Select

End Sub

Sub SolicitarCAE()
    Dim WSAA As Object, WSFEv1 As Object, cbte_nro As Integer
    
    ' Crear objeto interface Web Service Autenticación y Autorización
    Set WSAA = CreateObject("WSAA")
    
    ' Generar un Ticket de Requerimiento de Acceso (TRA) para WSFEv1
    tra = WSAA.CreateTRA("wsfe")
    
    ' Especificar la ubicacion de los archivos certificado y clave privada
    Path = CurDir() + "\"
    ' Certificado: certificado es el firmado por la AFIP
    ' ClavePrivada: la clave privada usada para crear el certificado
    CERTIFICADO = "reingart.crt" ' certificado de prueba
    ClavePrivada = "reingart.key" ' clave privada de prueba
    
    ' Generar el mensaje firmado (CMS)
    cms = WSAA.SignTRA(tra, Path + CERTIFICADO, Path + ClavePrivada)
    
    ' Llamar al web service WSAA para autenticar:
    cache = "" ' directorio temporal (usar predeterminado)
    url_wsdl = "https://wsaahomo.afip.gov.ar/ws/services/LoginCms"  ' usar wsaa.afip.gov.ar en producción
    proxy = "" ' información de servidor intermedio (si corresponde)
    ok = WSAA.Conectar(cache, url_wsdl, proxy, wrapper)
    TA = WSAA.LoginCMS(cms)
    
    ' Una vez obtenido, se puede usar el mismo token y sign por 24 horas
    ' (este período se puede cambiar)
    
    ' Crear objeto interface Web Service de Factura Electrónica de Mercado Interno
    Set WSFEv1 = CreateObject("WSFEv1")
    Debug.Print WSFEv1.version
    
    ' Setear tocken y sing de autorización (pasos previos)
    WSFEv1.token = WSAA.token
    WSFEv1.sign = WSAA.sign
    
    ' CUIT del emisor (debe estar registrado en la AFIP)
    WSFEv1.Cuit = "20267565393"
    
    ' Conectar al Servicio Web de Facturación
    cache = "" ' directorio temporal (usar predeterminado)
    url_wsdl = "https://wswhomo.afip.gov.ar/wsfev1/service.asmx?WSDL" ' usar servicios1 para producción
    proxy = "" ' información de servidor intermedio (si corresponde)
    ok = WSFEv1.Conectar(cache, url_wsdl, proxy) ' homologación
    
    ' Llamo a un servicio nulo, para obtener el estado del servidor (opcional)
    WSFEv1.Dummy
    Debug.Print "appserver status", WSFEv1.AppServerStatus
    Debug.Print "dbserver status", WSFEv1.DbServerStatus
    Debug.Print "authserver status", WSFEv1.AuthServerStatus
       
    ' Establezco los valores de la factura a autorizar:
    tipo_cbte = 1
    punto_vta = 4001
    cbte_nro = WSFEv1.CompUltimoAutorizado(tipo_cbte, punto_vta) + 1
    fecha = Format(Date, "yyyymmdd")
    concepto = 2
    tipo_doc = 80: nro_doc = "20397593046"
    cbt_desde = cbte_nro: cbt_hasta = cbte_nro
    imp_total = "122.00": imp_tot_conc = "0.00": imp_neto = "100.00"
    imp_iva = "21.00": imp_trib = "1.00": imp_op_ex = "0.00"
    fecha_cbte = fecha: fecha_venc_pago = fecha
    ' Fechas del período del servicio facturado (solo si concepto = 1?)
    fecha_serv_desde = fecha: fecha_serv_hasta = fecha
    moneda_id = "PES": moneda_ctz = "1.000"
    
    ok = WSFEv1.CrearFactura(concepto, tipo_doc, nro_doc, tipo_cbte, punto_vta, _
        cbt_desde, cbt_hasta, imp_total, imp_tot_conc, imp_neto, _
        imp_iva, imp_trib, imp_op_ex, fecha_cbte, fecha_venc_pago, _
        fecha_serv_desde, fecha_serv_hasta, _
        moneda_id, moneda_ctz)
    
    ' Agrego los comprobantes asociados:
    If False Then ' solo nc/nd
        tipo = 19
        pto_vta = 2
        nro = 1234
        ok = WSFEv1.AgregarCmpAsoc(tipo, pto_vta, nro)
    End If
        
    ' Agrego impuestos varios
    id = 99
    Desc = "Impuesto Municipal Matanza'"
    base_imp = "100.00"
    alic = "1.00"
    importe = "1.00"
    ok = WSFEv1.AgregarTributo(id, Desc, base_imp, alic, importe)
    
    ' Agrego tasas de IVA
    id = 5 ' 21%
    base_imp = "100.00"
    importe = "21.00"
    ok = WSFEv1.AgregarIva(id, base_imp, importe)
    
    ' Solicito CAE:
    cae = WSFEv1.CAESolicitar()
    
    Debug.Print "Resultado", WSFEv1.Resultado
    Debug.Print "CAE", WSFEv1.cae
    
    Debug.Print "Numero de comprobante:", WSFEv1.CbteNro
    
    ' Imprimo pedido y respuesta XML para depuración (errores de formato)
    Debug.Print WSFEv1.XmlRequest
    Debug.Print WSFEv1.XmlResponse
    
    Dim aux As String
    
    Call Archivos.WriteFile(WSFEv1.XmlRequest, "XmlRequest.xml")
    Call Archivos.WriteFile(WSFEv1.XmlResponse, "XmlResponse.xml")
    
    MsgBox "Resultado:" & WSFEv1.Resultado & " CAE: " & cae & " Venc: " & WSFEv1.Vencimiento & " Obs: " & WSFEv1.obs, vbInforma, tion + vbOKOnly
    txtResul.text = "Resultado:" & WSFEv1.Resultado & " CAE: " & cae & " Venc: " & WSFEv1.Vencimiento & " Obs: " & WSFEv1.obs & "," & vbInformation + vbOKOnly
    Debug.Print WSFEv1.ErrMsg
    
    ' Muestro los eventos (mantenimiento programados y otros mensajes de la AFIP)
    For Each evento In WSFEv1.eventos:
        MsgBox evento, vbInformation, "Evento"
    Next
End Sub
