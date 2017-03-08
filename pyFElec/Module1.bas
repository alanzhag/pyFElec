Attribute VB_Name = "Module1"
Public MyGlobalString As String
Dim WSAA As Object
Set WSAA = CreateObject("WSAA")
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
        ruta = WSAA.InstallDir + "\"
        Debug.Print "ruta", ruta

        ' Generar el mensaje firmado (CMS)
        cms = WSAA.SignTRA(tra, ruta + "reingart.crt", ruta + "reingart.key") ' Cert. Demo
        
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
