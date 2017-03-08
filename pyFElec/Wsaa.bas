Attribute VB_Name = "mdlWSAA"
Public TA As String
Dim WSFEv1 As Object


Public Function SetWsaa()
    TA = ""
    'Set WSAA = CreateObject("WSAA")
    'Set WSFEv1 = CreateObject("WSFEv1")
    
End Function

Function ObtenerRutaCertDemo()
    Dim Ruta As String
    Ruta = WSAA.InstallDir + "\" + "reingart.crt"
    Main.txtCertPath.text = Ruta
End Function
Function UltNro()
    Debug.Print TypeName(WSAA)
    Debug.Print WSAA.UltNro()
End Function

Function Autenticar(TA As String) As String
    ' Procedimiento para autenticar con AFIP y reutilizar el ticket de acceso
    ' Llamar antes de utilizar WSAA.Token y WSAA.Sign (WSAA debe estar definido a nivel de módulo)
    Dim ok, expiracion, solicitar, token, sign

    ' cargar ticket de acceso previo (si no se mantiene WSAA instanciado en memoria)
    If TA <> "" Then
        ok = WSAA.AnalizarXml(TA)
    End If

    ' revisar si el ticket es válido y no ha expirado:
    expiracion = WSAA.ObtenerTagXml("expirationTime")
    Debug.Print "Fecha Expiracion ticket: ", expiracion
    If IsNull(expiracion) Then
        solicitar = True                           ' solicitud inicial
    Else
        solicitar = WSAA.Expirado(expiracion)      ' chequear solicitud previa
    End If

    If solicitar Then
        ' Generar un Ticket de Requerimiento de Acceso (TRA)
        tra = WSAA.CreateTRA()

        ' uso la ruta a la carpeta de instalaciòn con los certificados de prueba
        Ruta = WSAA.InstallDir + "\"
        Debug.Print "ruta", Ruta

        ' Generar el mensaje firmado (CMS)
        cms = WSAA.SignTRA(tra, Ruta + "reingart.crt", Ruta + "reingart.key") ' Cert. Demo
        
        ok = WSAA.Conectar("", "https://wsaahomo.afip.gov.ar/ws/services/LoginCms") ' Homologacion

        ' Llamar al web service para autenticar
        TA = WSAA.LoginCMS(cms)
    Else
        Debug.Print "no expirado!", "Reutilizando!"
    End If
    Debug.Print WSAA.ObtenerTagXml("destination")

    ' Obtener las credenciales del ticket de acceso (desde el XML por si no se conserva el objeto WSAA)
    token = WSAA.ObtenerTagXml("token")
    sign = WSAA.ObtenerTagXml("sign")
    ' Al retornar se puede utilizar token y sign para WSFEv1 o similar
    ' Devuelvo el ticket de acceso (RETURN) para que el programa principal lo almacene si es necesario:
    Autenticar = TA
End Function

Public Function AnalizarCert() As Variant()
    Dim Values() As Variant
    crt = Main.txtCertPath.text
    WSAA.AnalizarCertificado (crt)
    Values = Array("Identidad: " + CStr(WSAA.Identidad), "Fecha de caducidad: " + CStr(WSAA.Caducidad), "Emisor: " + CStr(WSAA.Emisor))
    AnalizarCert = Values
End Function

Public Function ObtenerTicket()
    'ticket = mdlWSAA.TA
    'Call Autenticar(ticket)
    wsaa_url = "https://wsaahomo.afip.gov.ar/ws/services/LoginCms?wsdl"
    
    proxy = ""                           ' solo usar si hay servidor intermedio
    wrapper = ""                         ' httplib2 (default), pycurl (depende proxy)
    cacert = "conf/afip_ca_info.crt"     ' autoridades certificantes (servidores)
    cache = ""                           ' directorio archivos temporales (verificar permisos)
    devug = False                        ' depuración interna (en VB es palabra reservada, usar otro nombre)

    '   obtener el TA para pruebas
    'TA = WSAA.Autenticar("wsfe", "reingart.crt", "reingart.key", wsaa_url, proxy, wrapper, cacert, cache, devug)
    TA = Autenticar(TA)
    '# utilizar las credenciales:
    Debug.Print WSAA.token
    Debug.Print WSAA.sign

    ' establecer Ticket de Acceso en un solo paso (Nuevo método):
    WSFEv1.SetTicketAcceso (TA)
    
End Function
