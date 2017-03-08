VERSION 5.00
Object = "{13592B48-653C-491D-ACB1-C3140AA12F33}#6.1#0"; "ubGrid.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form frmTablasWSFEv1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tablas WSFEv1"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   14715
   Begin VB.CommandButton btnLimpiar 
      Caption         =   "Limpiar"
      Height          =   375
      Left            =   7440
      TabIndex        =   6
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton btnGuardar 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   11160
      TabIndex        =   4
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton btnConsultar 
      Caption         =   "Consultar"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   6240
      Width           =   1455
   End
   Begin ubGridControl.ubGrid grdTipos 
      Height          =   4695
      Left            =   750
      TabIndex        =   2
      Top             =   960
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   8281
      Rows            =   1
      Cols            =   4
      Redraw          =   -1  'True
      ShowGrid        =   -1  'True
      GridSolid       =   -1  'True
      GridLineColor   =   12632256
      BackColorFixed  =   12632256
      RowHeader       =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontEdit {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AllowUserResizing=   0   'False
   End
   Begin TabDlg.SSTab stbTipos 
      Height          =   5415
      Left            =   510
      TabIndex        =   0
      Top             =   480
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   9
      TabsPerRow      =   9
      TabHeight       =   520
      TabCaption(0)   =   "Comprobantes"
      TabPicture(0)   =   "frmTablasWSFEv1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Documentos"
      TabPicture(1)   =   "frmTablasWSFEv1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Conceptos"
      TabPicture(2)   =   "frmTablasWSFEv1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "IVA"
      TabPicture(3)   =   "frmTablasWSFEv1.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Tributos"
      TabPicture(4)   =   "frmTablasWSFEv1.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Puntos de Venta"
      TabPicture(5)   =   "frmTablasWSFEv1.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "Monedas"
      TabPicture(6)   =   "frmTablasWSFEv1.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).ControlCount=   0
      TabCaption(7)   =   "Opcionales"
      TabPicture(7)   =   "frmTablasWSFEv1.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).ControlCount=   0
      TabCaption(8)   =   "Cotización"
      TabPicture(8)   =   "frmTablasWSFEv1.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).ControlCount=   0
   End
   Begin VB.Frame fraTipos 
      Caption         =   "Tipos de"
      Height          =   5895
      Left            =   270
      TabIndex        =   1
      Top             =   240
      Width           =   14175
   End
   Begin VB.Label Label1 
      Caption         =   "Ultima actualización: "
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   6360
      Width           =   2895
   End
End
Attribute VB_Name = "frmTablasWSFEv1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const T_COMPRO = 0
Const T_DOCU = 1
Const T_CONCEP = 2
Const T_IVA = 3
Const T_TRIBU = 4
Const T_PTOVTA = 5
Const T_MONE = 6
Const T_OPCION = 7
Const T_COTI = 8

Const C_ID = 1
Const C_DESCRI = 2
Const C_FCHDDE = 3
Const C_FCHHST = 4

Private Sub btnConsultar_Click()
    Consultar
End Sub

Private Sub btnGuardar_Click()
    Guardar
End Sub

Private Sub btnLimpiar_Click()
    Limpiar
End Sub

Private Sub Form_Load()
    Call Utilidad.CentrarForm(Me)
    'ParamGetTiposMonedas(),ParamGetTiposCbte(), ParamGetTiposDoc(), ParamGetTiposIva(), ParamGetTiposOpcional(), ParamGetTiposTributos(), ParamGetTiposPaises(): recupera valores referenciales de códigos de las tablas de parámetros, devuelve una lista de strings con el id/código, descripción del parámetro y vigencia -si corresponde- (ver ejemplos). Más información en Tablas de Parámetros. ParamGetTiposPaises agregado para "COMPGv2.6"
    'ParamGetCotizacion(moneda_id): devuelve cotización y fecha de la moneda indicada como parámetro
    'ParamGetPtosVenta(): devuelve los puntos de venta autorizados para factura electrónica
    FormatearGrillaFactura
End Sub
Private Sub FormatearGrillaFactura()
    'Encabezado
    grdTipos.AutoSetup 1, 4, False, True, "Código     |Descripción     |Fch.Vig.Desde     |Fch.Vig.Hasta     "
    
    'Mascaras
    grdTipos.ColMask(C_FCHDDE) = 3
    grdTipos.ColMask(C_FCHHST) = 3
    
    Dim i As Integer
    'Bloquear columna total
    For i = 1 To grdTipos.Cols
        grdTipos.ColAllowEdit(i) = False
    Next
    
    'Alineamiento de columnas
    'grdItems.ColAlign(I_TOTAL) = 2
    
    'Cantidad caracteres
    'grdItems.ColEditWidth(I_PRODUC) = 60
    
    'Anchos
    grdTipos.ColWidth(C_ID) = 80
    grdTipos.ColWidth(C_DESCRI) = 600
    grdTipos.ColWidth(C_FCHDDE) = 95
    grdTipos.ColWidth(C_FCHHST) = 95
    
End Sub

Private Sub Consultar()
    On Error GoTo ManejoError
    Dim tipo As Integer
    
    If Not ModuloWSFEv1.Conectado() Then
        Exit Sub
    End If
    
    Limpiar
    
    tipo = stbTipos.Tab
    
    Select Case tipo
        Case T_COMPRO
            For Each X In WSFEv1.ParamGetTiposCbte()
                AgregarLinea (X)
            Next
        Case T_DOCU
            For Each X In WSFEv1.ParamGetTiposDoc()
                AgregarLinea (X)
            Next
        Case T_CONCEP
            For Each X In WSFEv1.ParamGetTiposConcepto()
                AgregarLinea (X)
            Next
        Case T_IVA
            For Each X In WSFEv1.ParamGetTiposIva()
                AgregarLinea (X)
            Next
        Case T_TRIBU
            For Each X In WSFEv1.ParamGetTiposTributos()
                AgregarLinea (X)
            Next
        Case T_PTOVTA
            For Each X In WSFEv1.ParamGetPtosVenta()
                AgregarLinea (X)
            Next
        Case T_MONE
            For Each X In WSFEv1.ParamGetTiposMonedas()
                AgregarLinea (X)
            Next
        Case T_OPCION
            For Each X In WSFEv1.ParamGetTiposOpcional()
                AgregarLinea (X)
            Next
        Case T_COTI
            ctz = WSFEv1.ParamGetCotizacion("DOL")
            MsgBox "Cotización Dólar: " & ctz
    End Select

    grdTipos.col = 1
    grdTipos.row = 1
    
    Exit Sub
    
ManejoError:
    ' Si hubo error:
    Debug.Print Err.Description            ' descripción error afip
    Debug.Print Err.Number - vbObjectError ' codigo error afip
    Select Case MsgBox(Err.Description, vbCritical + vbRetryCancel, "Error:" & Err.Number - vbObjectError & " en " & Err.Source)
        Case vbRetry
            Debug.Print WSFEv1.Excepcion
            Debug.Print WSFEv1.Traceback
            Debug.Print WSFEv1.XmlRequest
            Debug.Print WSFEv1.XmlResponse
            Debug.Assert False
            Resume
        Case vbCancel
            Debug.Print Err.Description
    End Select
    Debug.Print WSFEv1.XmlRequest
    Debug.Assert False
End Sub
Private Sub AgregarLinea(rawLine As String)
    Dim i As Integer
    Dim cookedLine() As String
    
    cookedLine() = Split(rawLine, "|")
    grdTipos.AddItem ("")
    grdTipos.row = grdTipos.row + 1
    grdTipos.col = 1
    For i = 1 To grdTipos.Cols
        grdTipos.TextMatrix(grdTipos.row, i) = cookedLine(i - 1)
    Next

End Sub

Private Sub Limpiar()
    grdTipos.Clear
End Sub
Private Sub Guardar()
    'paso lo de la grilla a un archivo
    tipo = stbTipos.Tab
    
    Select Case tipo
        Case T_COMPRO
            clave = "COMPROBANTES"
        Case T_DOCU
            clave = "DOCUMENTOS"
        Case T_CONCEP
            clave = "CONCEPTO"
        Case T_IVA

        Case T_TRIBU

        Case T_PTOVTA

        Case T_MONE

        Case T_OPCION

    For i = 1 To grdTipos.Rows
        id = grdTipos.text
        descrip = grdTipos.text
        writeINI sWSTipos & "\" & tipo & ".dat", clave, id, descrip
    Next
End Sub

Private Sub stbTipos_Click(PreviousTab As Integer)
    Consultar
End Sub

