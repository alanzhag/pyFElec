VERSION 5.00
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "AfipWebService Tools Z"
   ClientHeight    =   8025
   ClientLeft      =   165
   ClientTop       =   630
   ClientWidth     =   12525
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuVentas 
      Caption         =   "&Ventas"
      Index           =   1
      Begin VB.Menu mnuFactu 
         Caption         =   "Factura"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuCredi 
         Caption         =   "Nota de Credito"
      End
      Begin VB.Menu mnuDebi 
         Caption         =   "Nota de Debito"
      End
   End
   Begin VB.Menu mnuWs 
      Caption         =   "&WebService"
      Index           =   2
      Begin VB.Menu mnuWSAA 
         Caption         =   "WSAA"
         Begin VB.Menu mnuConectarWSAA 
            Caption         =   "Conectar"
            Shortcut        =   {F5}
         End
         Begin VB.Menu mnuToken 
            Caption         =   "Token"
         End
         Begin VB.Menu mnuSign 
            Caption         =   "Sign"
         End
      End
      Begin VB.Menu mnuWSFEv1 
         Caption         =   "WSFEv1"
         Begin VB.Menu mnuConectarWSFEv1 
            Caption         =   "Conectar"
            Shortcut        =   {F6}
         End
         Begin VB.Menu mnuTablasWSFEv1 
            Caption         =   "Tablas"
         End
      End
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "&Configuracion"
      Index           =   3
      Begin VB.Menu mnuCert 
         Caption         =   "Certificados"
         Index           =   1
      End
      Begin VB.Menu mnuPath 
         Caption         =   "Directorios"
         Index           =   2
      End
      Begin VB.Menu mnuDatosFiscales 
         Caption         =   "Datos Fiscales"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "Acerca de"
   End
   Begin VB.Menu mnuDebug 
      Caption         =   "Debug"
      Begin VB.Menu mnuTest1 
         Caption         =   "Test1"
      End
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
    Call InitProgram
    Me.Width = 16000
End Sub

Private Sub mnuAbout_Click()
    MsgBox "Interfaz WebService" & vbCrLf _
    & "Copyright " & Chr$(169) & " 2017 Alan Zhao", , _
    "About"
End Sub

Private Sub mnuCert_Click(Index As Integer)
    frmCertificados.Show
End Sub

Sub InitProgram()
    ArchivoINI
    ObtenerRutas
    Configurar
    StartWS
End Sub

Private Sub mnuConectarWSAA_Click()
    IniciarWSAA
End Sub

Private Sub mnuConectarWSFEv1_Click()
    IniciarWSFEv1
End Sub

Private Sub mnuDatosFiscales_Click()
    frmDatosFiscales.Show
End Sub

Private Sub mnuFactu_Click()
    frmFactura.Show
End Sub

Private Sub mnuPath_Click(Index As Integer)
    frmDirectorios.Show
End Sub

Private Sub mnuSign_Click()
    MsgBox "Sign: " + sign, vbOKOnly, "WSAA"
End Sub

Private Sub mnuTablasWSFEv1_Click()
    frmTablasWSFEv1.Show
End Sub

Private Sub mnuTest1_Click()
    Test2
End Sub

Private Sub mnuToken_Click()
    MsgBox "Token: " + token, vbOKOnly, "WSAA"
End Sub
Sub ArchivoINI()
    Dim existe As String
    sINIFile = App.Path & "\AWSZ.INI"
    existe = IfNotExistsCreateFile(existe, sINIFile, "CONFIG", "RESTORE", "NO")
End Sub

Sub ObtenerRutas()
    'en estos archivos se guardan las configuraciones
    sWSFile = IfNotExistsCreateFile(sWSFile, sINIFile, "Rutas", "WS_PATH", App.Path & "\WebService\WS.INI")
    sDatosFile = IfNotExistsCreateFile(sDatosFile, sINIFile, "Rutas", "DATFIS_PATH", App.Path & "\Datos\DatosFiscales.INI")
End Sub
Sub Configurar()
    'WebService
    sCertPath = sGetINI(sWSFile, "Certificado", "CERT_PATH", "?") 'El que devuelve la afips
    sPrivPath = sGetINI(sWSFile, "Certificado", "PRIV_PATH", "?") 'La clave privada con la que se genero en primera instancia
    sWSXML = sGetINI(sWSFile, "Datos", "XML_PATH", "?")
    sWSTipos = sGetINI(sWSFile, "Datos", "TIPOS_PATH", "?")
    'Datos Fiscales
    sMiCuit = sGetINI(sDatosFile, "Datos", "CUIT", "?")
    sMiRazSoc = sGetINI(sDatosFile, "Datos", "RAZSOC", "?")
End Sub
Sub StartWS()
    Set WSAA = CreateObject("WSAA")
    Set WSFEv1 = CreateObject("WSFEv1")
End Sub
