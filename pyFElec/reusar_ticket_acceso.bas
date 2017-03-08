Attribute VB_Name = "reusar_ticket_acceso"
'Public WSAA As Object

Sub Main()
    ' Crear objeto interface Web Service Autenticación y Autorización
    'Set WSAA = CreateObject("WSAA")
    ' verifico la versión:
    Debug.Assert WSAA.version >= "2.04a"
    ' deshabilito errores no manejados (version 2.04 o superior)
    WSAA.LanzarExcepciones = False
    
    ' datos de prueba del certificado (para depuración):
    Dest = "C=ar, O=pyafipws-sistemas agiles, SERIALNUMBER=CUIT 20267565393, CN=mariano reingart"
    
    ' inicializo las variables:
    token = ""
    sign = ""

    ' busco un ticket de acceso previamente almacenado:
    If Dir("ta.xml") <> "" Then
        ' leo el xml almacenado del archivo
        Open "ta.xml" For Input As #1
        Line Input #1, ta_xml
        Close #1
        ' analizo el ticket de acceso previo:
        ok = WSAA.AnalizarXml(ta_xml)
        ' verifico que el destino corresponda (CUIT)
        Debug.Assert WSAA.ObtenerTagXml("destination") = Dest
        ' verificar CUIT
        If Not WSAA.Expirado() Then
            ' puedo reusar el ticket de acceso:
            token = WSAA.ObtenerTagXml("token")
            sign = WSAA.ObtenerTagXml("sign")
        End If
    End If

    ' Si no reuso un ticket de acceso, solicito uno nuevo:
    If token = "" Or sign = "" Then
        ' Generar un Ticket de Requerimiento de Acceso (TRA)
        tra = WSAA.CreateTRA("wsfe", 43200) ' 3600*12hs
        ' Especificar la ubicacion de los archivos certificado y clave privada
        cert = WSAA.InstallDir + "\" + "reingart.crt" ' certificado de prueba
        clave = WSAA.InstallDir + "\" + "reingart.key" ' clave privada de prueba
        ' Generar el mensaje firmado (CMS)
        cms = WSAA.SignTRA(tra, cert, clave)
        If cms <> "" Then
            ' Llamar al web service para autenticar:
            ok = WSAA.Conectar()
            ta_xml = WSAA.LoginCMS(cms)
            If ta_xml <> "" Then
                ' guardo el ticket de acceso en el archivo
                Open "ta.xml" For Output As #1
                Print #1, ta_xml
                Close #1
            End If
            token = WSAA.token
            sign = WSAA.sign
        End If
        ' reviso que no haya errores:
        Debug.Print "excepcion", WSAA.Excepcion
        If WSAA.Excepcion <> "" Then
            Debug.Print WSAA.Traceback
            MsgBox WSAA.Excepcion, vbCritical, "Excepción"
        End If
    End If
    
    ' Imprimir los datos del ticket de acceso: ToKen y Sign de autorización
    MsgBox "Token: " + token
    MsgBox "Sign: " + sign
        
End Sub
