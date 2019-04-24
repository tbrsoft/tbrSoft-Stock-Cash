VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmPedidos 
   BackColor       =   &H00544B45&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pedidos Pendientes"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8460
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPedidos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin tbrFaroButton.fBoton cmdEliminar 
      Height          =   465
      Left            =   6840
      TabIndex        =   2
      Top             =   2130
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "eliminar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdModificar 
      Height          =   465
      Left            =   6840
      TabIndex        =   3
      Top             =   1290
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "modificar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin MSComctlLib.ListView lvFacturasP 
      Height          =   3765
      Left            =   390
      TabIndex        =   1
      Top             =   570
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   6641
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nro Factura"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Fecha"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Proveedor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Importe"
         Object.Width           =   2117
      EndProperty
   End
   Begin tbrFaroButton.fBoton command1 
      Height          =   465
      Left            =   7020
      TabIndex        =   4
      Top             =   4530
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione el pedido"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   390
      TabIndex        =   0
      Top             =   240
      Width           =   2685
   End
End
Attribute VB_Name = "frmPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEliminar_Click()
    If lvFacturasP.ListItems.Count = 0 Then Exit Sub
    
    Dim NroFac As String
    
    NroFac = txtInLvW(lvFacturasP, lvFacturasP.SelectedItem.Index, 0)
    
    If MsgBox("¿Está seguro de borrar el pedido nº " + NroFac + "?", _
        vbOKCancel + vbExclamation, "Atención") = vbCancel Then Exit Sub
    
    DB.EXECUTE "DELETE FROM FacturaCompra WHERE NroFactura = '" + NroFac + "'"
    
    MsgBox "El pedido fue eliminado definitivamente", vbExclamation
    
    Unload Me
    
End Sub

Private Sub cmdModificar_Click()
    If lvFacturasP.ListItems.Count = 0 Then Exit Sub
    
    Dim NroFac As String
    
    NroFac = txtInLvW(lvFacturasP, lvFacturasP.SelectedItem.Index, 0)
    
    frmCompras.AbrirDatos NroFac
    
    Unload Me
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    CargarComboLV lvFacturasP, "SELECT NroFactura, Fecha, Proveedor, Pagado " + _
        "FROM FacturaCompra WHERE EsPedido <> 0", "NroFactura," + _
        "Fecha/f,Proveedor,Pagado/$"
End Sub
