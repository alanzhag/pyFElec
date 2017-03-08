VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Main 
   Caption         =   "Form1"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCAEReq 
      Caption         =   "Solicitar CAE"
      Height          =   375
      Left            =   1680
      TabIndex        =   14
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton btnDummy 
      Caption         =   "Testeo"
      Height          =   615
      Left            =   5520
      TabIndex        =   13
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton btnUltNro 
      Caption         =   "Ultimo Numero"
      Height          =   375
      Left            =   1680
      TabIndex        =   12
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton btnClear 
      Caption         =   "Limpiar Resul"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton btnAuth 
      Caption         =   "Autenticar"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Frame frmResul 
      Caption         =   "Resultados"
      Height          =   2535
      Left            =   360
      TabIndex        =   3
      Top             =   3120
      Width           =   6015
      Begin VB.TextBox txtResul 
         Height          =   2175
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   4
         Text            =   "FElec.frx":0000
         Top             =   240
         Width           =   5775
      End
   End
   Begin VB.CommandButton btnReqTkt 
      Caption         =   "Solicitar Ticket"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtPrivPath 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Text            =   "Ruta..."
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox txtCertPath 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Text            =   "Ruta..."
      Top             =   360
      Width           =   2415
   End
   Begin VB.Frame frmFiles 
      Caption         =   "Archivos"
      Height          =   1215
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   4695
      Begin VB.CommandButton btnPriv 
         Caption         =   "Buscar"
         Height          =   255
         Left            =   3840
         TabIndex        =   10
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton btnCert 
         Caption         =   "Buscar"
         Height          =   255
         Left            =   3840
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3960
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblPriv 
         Caption         =   "Private.key"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblCert 
         Caption         =   "Certificado.crt"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAuth_Click()
    Dim Values() As Variant
    Values() = AnalizarCert()
    Call MostrarResul("Auth", Values)
End Sub

Private Sub btnCAEReq_Click()
    Monolitico.Main
End Sub

Private Sub btnCert_Click()
   Call BuscarRuta(CommonDialog1, txtCertPath, "Certificates Files (*.crt)|*.crt|")
End Sub

Private Sub btnClear_Click()
    Call LimpiarResultado
End Sub

Private Sub btnDummy_Click()


Private Sub btnPriv_Click()
   ' CancelError is True.
   On Error GoTo ErrHandler
   ' Set filters.
   CommonDialog1.Filter = "Key Files (*.key)|*.key|All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Batch Files (*.bat)|*.bat"
   ' Specify default filter.
   CommonDialog1.FilterIndex = 1

   ' Display the Open dialog box.
   CommonDialog1.ShowOpen
   ' Call the open file procedure.
   txtPrivPath.text = CommonDialog1.filename
   Exit Sub

ErrHandler:
' User pressed Cancel button.
   Exit Sub
End Sub

Private Sub btnReqTkt_Click()
    reusar_ticket_acceso.Main
End Sub

Private Sub btnUltNro_Click()
    Call mdlWSAA.UltNro
End Sub

Private Sub Form_Load()
    Dim Ruta As String
    'Call mdlWSAA.SetWsaa
    'mdlWSAA.ObtenerRutaCertDemo
    'txtCertPath.Text = ruta
End Sub

Private Sub MostrarResul(opc As String, val() As Variant)
    Call LimpiarResultado
    Dim iT
    Select Case opc
        Case "Auth"
            'Debug.Print ("Auth")
            For Each iT In val
                'Debug.Print obj
                txtResul.text = txtResul.text & CStr(iT) & vbNewLine
            Next
    End Select
        
End Sub

Private Sub LimpiarResultado()
    txtResul.text = ""
End Sub

Private Sub BuscarRuta(CommonDialog As CommonDialog, volcarEn As TextBox, tipoArchivo As String)
' CancelError is True.
   On Error GoTo ErrHandler
   ' Set filters.
   CommonDialog.Filter = tipoArchivo + "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|Batch Files (*.bat)|*.bat"
   ' Specify default filter.
   CommonDialog.FilterIndex = 1
   ' Display the Open dialog box.
   CommonDialog.ShowOpen
   ' Call the open file procedure.
   volcarEn.text = CommonDialog.filename
   Exit Sub

ErrHandler:
' User pressed Cancel button.
   Exit Sub
End Sub



Sub InitProgram()
Dim sINIFile As String
Dim sUserName As String
Dim nCount As Integer
Dim i As Integer
'Store the location of the INI file
sINIFile = App.Path & "\MYAPP.INI"
'Read the user name from the INI file
sUserName = sGetINI(sINIFile, "Settings", "UserName", "?")
If sUserName = "?" Then
      'No user name was present – ask for it and save for next time
      sUserName = InputBox$("Enter your name please:")
      writeINI sINIFile, "Settings", "UserName", sUserName
End If
'Fill up combo box list from INI file and select the user's
'last chosen item
nCount = CInt(sGetINI(sINIFile, "Regions", "Count", 0))
For i = 1 To nCount
cmbRegn.AddItem sGetINI(sINIFile, "Regions", "Region" & i, "?")
Next i
cmbRegn.text = sGetINI(sINIFile, "Regions", _
               "LastRegion", cmbRegions.List(0))
End Sub

