Attribute VB_Name = "mdlWSFE"
Sub SolicitarCAE()
    Dim cbte_nro As Integer 'este funca
    
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
    concepto = 1
    tipo_doc = 80: nro_doc = "20397593046"
    cbt_desde = cbte_nro: cbt_hasta = cbte_nro
    imp_total = "122.00": imp_tot_conc = "0.00": imp_neto = "100.00"
    imp_iva = "21.00": imp_trib = "1.00": imp_op_ex = "0.00"
    fecha_cbte = fecha: fecha_venc_pago = fecha
    ' Fechas del período del servicio facturado (solo si concepto = 1?)
    'fecha_serv_desde = fecha: fecha_serv_hasta = fecha
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


