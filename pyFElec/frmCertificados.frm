VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCertificados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Certificados"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   6555
   Begin VB.CommandButton btnGuardar 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   5160
      TabIndex        =   8
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton btnAuth 
      Caption         =   "Autenticar"
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   360
      Width           =   1215
   End
   Begin VB.Frame frmFiles 
      Caption         =   "Archivos"
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4695
      Begin VB.TextBox txtPrivPath 
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Text            =   "Ruta..."
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox txtCertPath 
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Text            =   "Ruta..."
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton btnCert 
         Caption         =   "Buscar"
         Height          =   255
         Left            =   3840
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton btnPriv 
         Caption         =   "Buscar"
         Height          =   255
         Left            =   3840
         TabIndex        =   1
         Top             =   720
         Width           =   735
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3960
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblCert 
         Caption         =   "Certificado.crt"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblPriv 
         Caption         =   "Privada.key"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmCertificados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAuth_Click()
    Autenticar
End Sub

Private Sub btnCert_Click()
    BuscarRuta CommonDialog1, txtCertPath, "Certificates Files (*.crt)|*.crt|"
End Sub

Private Sub btnPriv_Click()
    BuscarRuta CommonDialog1, txtPrivPath, "Private Key Files (*.Key)|*.key|"
End Sub

Private Sub btnGuardar_Click()
    Dim op As String
    op = MsgBox("¿Desea guardar los cambios?", vbYesNo + vbQuestion, "Guardar")
    If op = vbYes Then
        Guardar
    End If
End Sub

Private Sub Form_Load()
    CentrarForm Me
    RecuperarContenidoTextbox txtCertPath, sCertPath
    RecuperarContenidoTextbox txtPrivPath, sPrivPath
End Sub

Private Sub Guardar()
     writeINI sWSFile, "Certificado", "CERT_PATH", txtCertPath.text
     writeINI sWSFile, "Certificado", "PRIV_PATH", txtPrivPath.text
     sCertPath = txtCertPath.text
     sPrivPath = txtPrivPath.text
End Sub

Private Sub Autenticar()
    crt = txtCertPath.text
    WSAA.AnalizarCertificado (crt)
    MsgBox "Identidad: " + CStr(WSAA.Identidad) + vbNewLine + _
    "Fecha de caducidad: " + CStr(WSAA.Caducidad) + vbNewLine + _
    "Emisor: " + CStr(WSAA.Emisor), vbOKOnly + vbInformation, "Autenticacion"
End Sub

