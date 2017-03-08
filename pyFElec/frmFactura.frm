VERSION 5.00
Object = "{13592B48-653C-491D-ACB1-C3140AA12F33}#6.1#0"; "ubGrid.ocx"
Begin VB.Form frmFactura 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Factura"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   5490
   Begin VB.CommandButton btnClear 
      Caption         =   "Limpiar"
      Height          =   375
      Left            =   4200
      TabIndex        =   21
      Top             =   2280
      Width           =   1095
   End
   Begin ubGridControl.ubGrid grdItems 
      Height          =   3015
      Left            =   240
      TabIndex        =   7
      Top             =   2280
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   5318
      Rows            =   1
      Cols            =   5
      Redraw          =   -1  'True
      ShowGrid        =   -1  'True
      GridSolid       =   -1  'True
      GridLineColor   =   12632256
      BackColorFixed  =   12632256
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
   End
   Begin VB.TextBox txtCbtCae 
      Height          =   285
      Left            =   840
      TabIndex        =   27
      Text            =   "67063464461975"
      Top             =   6480
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Recuperar CAE"
      Height          =   375
      Left            =   1080
      TabIndex        =   25
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Solicitar CAE"
      Height          =   375
      Left            =   1080
      TabIndex        =   24
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton btnCalcular 
      Caption         =   "Calcular"
      Height          =   375
      Left            =   4200
      TabIndex        =   23
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtSubTotal 
      Height          =   285
      Left            =   3960
      TabIndex        =   9
      Text            =   "100"
      Top             =   5760
      Width           =   1335
   End
   Begin VB.TextBox txtImpIva 
      Height          =   285
      Left            =   3960
      TabIndex        =   10
      Text            =   "21"
      Top             =   6120
      Width           =   1335
   End
   Begin VB.TextBox txtImpTotal 
      Height          =   285
      Left            =   3960
      TabIndex        =   11
      Text            =   "121"
      Top             =   6450
      Width           =   1335
   End
   Begin VB.Frame fraFechas 
      Caption         =   "Fechas"
      Height          =   615
      Left            =   240
      TabIndex        =   16
      Top             =   1440
      Width           =   2295
      Begin VB.TextBox txtCbtFecha 
         Height          =   285
         Left            =   960
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame fraCbt 
      Caption         =   "Comprobante"
      Height          =   1455
      Left            =   2760
      TabIndex        =   8
      Top             =   240
      Width           =   2535
      Begin VB.TextBox txtCbtNro 
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtCbtPtoVta 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Text            =   "4001"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtCbtTipo 
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Text            =   "1"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblCbtNro 
         Caption         =   "Numero"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblPtoVta 
         Caption         =   "Pto. Venta"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblTipo 
         Caption         =   "Tipo"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame fraCliente 
      Caption         =   "Cliente"
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2295
      Begin VB.TextBox txtTipoDoc 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Text            =   "80"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtNroDoc 
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Text            =   "2039759304"
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblTipoDoc 
         Caption         =   "Tipo Doc."
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblNroDoc 
         Caption         =   "Numero"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Label Label3 
      Caption         =   "CAE"
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   6480
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "SubTotal"
      Height          =   255
      Left            =   3120
      TabIndex        =   22
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label lblIva 
      Caption         =   "IVA"
      Height          =   255
      Left            =   3120
      TabIndex        =   20
      Top             =   6120
      Width           =   735
   End
   Begin VB.Label lblTotal 
      Caption         =   "Total"
      Height          =   255
      Left            =   3120
      TabIndex        =   19
      Top             =   6480
      Width           =   735
   End
End
Attribute VB_Name = "frmFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const I_PRODUC = 1
Const I_CANTI = 2
Const I_PRECIO = 3
Const I_TOTAL = 4
Private Sub btnCalcular_Click()
    SubTotales
    CalcularIVA
    CalcularTotal
End Sub

Private Sub btnClear_Click()
    grdItems.Clear
    grdItems.AddItem ("")
End Sub

Private Sub Form_Load()
    Call Utilidad.CentrarForm(Me)
    FormatearGrillaFactura
    CompletarForm
End Sub

Private Sub FormatearGrillaFactura()
    'Encabezado
    grdItems.AutoSetup 1, 4, False, True, "Producto     |Cantidad     |Precio     |Total     "
    
    'Mascaras
    grdItems.ColMask(I_PRODUC) = 1
    grdItems.ColMask(I_CANTI) = 2
    grdItems.ColMask(I_PRECIO) = 2
    
    'Bloquear columna total
    grdItems.ColAllowEdit(I_TOTAL) = False
    
    'Alineamiento de columnas
    grdItems.ColAlign(I_TOTAL) = 2
    
    'Cantidad caracteres
    grdItems.ColEditWidth(I_PRODUC) = 60
    
    'Anchos
    grdItems.ColWidth(I_PRODUC) = 82
    
    grdItems.AutoNewRow = True
    
    'grdItems.TextMatrix(1, I_CANTI) = CDbl(0#)  'Format("0.00", "#.##")
    'grdItems.TextMatrix(1, I_PRECIO) = CCur(0#)  'Format("0.00", "#.##")
    
    grdItems.AutoRedraw = True
    grdItems.Refresh
    
End Sub
Private Sub CompletarForm()
    txtCbtFecha.text = Format(Date, "yyyymmdd")
End Sub

Private Function RowIsEmpty(isEmpty As Boolean) As Boolean
    Dim i As Integer
    isEmpty = True
    For i = 0 To grdItems.Cols
        If grdItems.text <> "" Then
            isEmpty = False
            Exit For
        End If
    Next
    RowIsEmpty = isEmpty
End Function


Private Sub grdItems_AfterEdit(ByVal row As Long, ByVal col As Long, ByVal NewValue As String)
    Select Case col
        Case I_CANTI
            CalcularTotalLinea (row)
        Case I_PRECIO
            CalcularTotalLinea (row)
    End Select
End Sub

Private Sub grdItems_BeforeAddRow(Cancel As Boolean)
    Dim isEmpty As Boolean
    'isEmpty = RowIsEmpty(isEmpty)
    isEmpty = Utilidad.RowIsEmpty(isEmpty, grdItems, grdItems.row)
    If isEmpty Then
        Cancel = True
    ElseIf Not isEmpty And Utilidad.RowIsEmpty(isEmpty, grdItems, grdItems.row + 1) Then
        Cancel = True
    End If
End Sub

Private Sub grdItems_BeforeDeleteRow(ByVal row As Long, Cancel As Boolean)
    If grdItems.Rows <= 1 Then
        Cancel = True
    End If
End Sub

Private Sub grdItems_KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
    Select Case KeyCode
    Case vbKeyReturn
        MoverFocusGrilla
    Case vbKeyInsert
        grdItems.AddItem ("")
    Case vbKeyDelete
        grdItems.RemoveItem grdItems.row
    End Select
End Sub

Private Sub CalcularTotalLinea(linea As Long)
    Dim canti As Double
    Dim precio, total As Currency
    Dim aux As String
    
    If linea <= 0 Then
        Exit Sub
    End If
    grdItems.col = I_CANTI
    grdItems.row = linea
    canti = CDbl(Utilidad.GetNumberFromString(grdItems.text))
    grdItems.col = I_PRECIO
    precio = CCur(Utilidad.GetNumberFromString(grdItems.text))
    total = canti * precio
    grdItems.TextMatrix(linea, I_TOTAL) = CCur(total)
End Sub
Private Sub MoverFocusGrilla()
    If grdItems.col < grdItems.Cols - 1 Then
        grdItems.col = grdItems.col + 1
    ElseIf grdItems.col = grdItems.Cols - 1 Then
        grdItems.AddItem ("")
        grdItems.row = grdItems.row + 1
        grdItems.col = 1
    End If
End Sub
Private Sub SubTotales()
    Dim i As Integer
    Dim subTotal As Currency
    For i = 1 To grdItems.Rows
        grdItems.col = I_TOTAL
        grdItems.row = i
        subTotal = subTotal + CCur(Utilidad.GetNumberFromString(grdItems.text))
    Next
    txtSubTotal.text = Format(subTotal, "0.00")
End Sub

Private Sub CalcularIVA()
    Dim subTotal As Currency
    subTotal = txtSubTotal.text
    txtImpIva.text = Format(subTotal * 0.21, "0.00")
End Sub
Private Sub CalcularTotal()
    Dim subTotal, impuestos As Currency
    subTotal = txtSubTotal.text
    impuestos = impuestos + txtImpIva.text
    txtImpTotal.text = Format(subTotal + impuestos, "0.00")
End Sub
