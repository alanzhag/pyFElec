VERSION 5.00
Begin VB.Form frmDatosFiscales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos Fiscales"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   6990
   Begin VB.CommandButton btnSave 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Frame fraIdentidad 
      Caption         =   "Identidad"
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6495
      Begin VB.TextBox txtRazSoc 
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox txtCuit 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label lblRazSoc 
         Caption         =   "Razon Social:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblCuit 
         Caption         =   "CUIT:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmDatosFiscales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    CentrarForm Me
    RecuperarContenidoTextbox txtCuit, sMiCuit
    RecuperarContenidoTextbox txtRazSoc, sMiRazSoc
End Sub

Private Sub btnSave_Click()
    Dim op As String
    op = MsgBox("¿Desea guardar los cambios?", vbYesNo + vbQuestion, "Guardar")
    If op = vbYes Then
        Guardar
    End If
End Sub

Private Sub Guardar()
     writeINI sDatosFile, "Datos", "CUIT", txtCuit.text
     writeINI sDatosFile, "Datos", "RAZSOC", txtRazSoc.text
     sMiCuit = txtCuit.text
     sMiRazSoc = txtRazSoc.text
End Sub

