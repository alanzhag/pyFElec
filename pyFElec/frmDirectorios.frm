VERSION 5.00
Begin VB.Form frmDirectorios 
   Caption         =   "Directorios"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4695
   ScaleWidth      =   9375
   Begin VB.CommandButton btnGuardar 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   5640
      TabIndex        =   8
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Frame fraWebService 
      Caption         =   "WebService"
      Height          =   1455
      Left            =   360
      TabIndex        =   2
      Top             =   2160
      Width           =   8175
      Begin VB.TextBox txtWSTipos 
         Height          =   285
         Left            =   1680
         TabIndex        =   7
         Text            =   "Ruta..."
         Top             =   840
         Width           =   6255
      End
      Begin VB.TextBox txtWSXML 
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Text            =   "Ruta..."
         Top             =   360
         Width           =   6255
      End
      Begin VB.Label lblTipos 
         Caption         =   "Tipos:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblXML 
         Caption         =   "XML:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame fraConfig 
      Caption         =   "Configuracion"
      Height          =   1575
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   8175
      Begin VB.TextBox txtDatFis 
         Height          =   285
         Left            =   1680
         TabIndex        =   12
         Text            =   "Ruta..."
         Top             =   1080
         Width           =   6255
      End
      Begin VB.TextBox txtWSINI 
         Height          =   285
         Left            =   1680
         TabIndex        =   10
         Text            =   "Ruta..."
         Top             =   720
         Width           =   6255
      End
      Begin VB.TextBox txtAWSZINI 
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Text            =   "Ruta..."
         Top             =   360
         Width           =   6255
      End
      Begin VB.Label lblDatFis 
         Caption         =   "DatosFiscales.INI:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblWSINI 
         Caption         =   "WS.INI:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblConfigINI 
         Caption         =   "AWSZ.INI:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmDirectorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnGuardar_Click()
    Dim op As String
    op = MsgBox("¿Desea guardar los cambios?", vbYesNo + vbQuestion, "Guardar")
    If op = vbYes Then
        Guardar
    End If
End Sub

Private Sub Form_Load()
    CentrarForm Me
    
    'Configuracion
    RecuperarContenidoTextbox txtAWSZINI, sINIFile
    RecuperarContenidoTextbox txtWSINI, sWSFile
    RecuperarContenidoTextbox txtDatFis, sDatosFile
    'WS
    RecuperarContenidoTextbox txtWSXML, sWSXML
    RecuperarContenidoTextbox txtWSTipos, sWSTipos
    
End Sub

Private Sub Guardar()
    Dim existe As String
    sINIFile = txtAWSZINI.text
    existe = IfNotExistsCreateFile(existe, sINIFile, "CONFIG", "RESTORE", "NO")
    sWSFile = IfNotExistsCreateFile(sWSFile, sINIFile, "Rutas", "WS_PATH", txtWSINI.text)
    sDatosFile = IfNotExistsCreateFile(sDatosFile, sINIFile, "Rutas", "DATFIS_PATH", txtDatFis.text)

    writeINI sWSFile, "Datos", "XML_PATH", txtWSXML.text
    writeINI sWSFile, "Datos", "TIPOS_PATH", txtWSTipos.text
    
    MDIMain.Configurar
    
End Sub

Private Sub fraConfig_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub lblWSINI_Click()

End Sub
