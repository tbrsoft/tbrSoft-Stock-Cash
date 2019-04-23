VERSION 5.00
Object = "{A7FBD38D-2930-49E3-B60C-9E0202D84549}#15.0#0"; "tbrControles.ocx"
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmVerFactura 
   BackColor       =   &H00544B45&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Visualizador de Facturas"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10965
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVerFactura.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   10965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin tbrFaroButton.fBoton cmdEntregar 
      Height          =   435
      Left            =   6660
      TabIndex        =   32
      Top             =   7875
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Entregar"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdImprimir 
      Height          =   435
      Index           =   1
      Left            =   8235
      TabIndex        =   31
      Top             =   6795
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Imprimir Prod. a Entregar"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdLimpiar 
      Height          =   435
      Left            =   6600
      TabIndex        =   30
      Top             =   6795
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Limpiar"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdBorrarFactura 
      Height          =   435
      Left            =   5220
      TabIndex        =   27
      Top             =   3705
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Borrar Factura"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdAdd 
      Height          =   465
      Left            =   6015
      TabIndex        =   26
      Top             =   5460
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   ">>"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin VB.ComboBox cmbSucursales 
      Height          =   315
      Left            =   6870
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   7380
      Width           =   2505
   End
   Begin tbrControles.tbrBuscador tbrBuscador1 
      Height          =   1905
      Left            =   420
      TabIndex        =   22
      Top             =   1740
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   3360
      BackColor       =   5524293
      BeginProperty Fontt {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox lstFacturasC 
      BackColor       =   &H00D8D9D7&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   7470
      TabIndex        =   19
      Top             =   2670
      Width           =   2655
   End
   Begin MSComctlLib.ListView lvCamion 
      Height          =   2415
      Left            =   6900
      TabIndex        =   17
      Top             =   4320
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   4260
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "IdP"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Producto"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Cantidad"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.CheckBox chkEnt 
      BackColor       =   &H00544B45&
      Caption         =   "Seleccionar Facturas sin entregar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   345
      Left            =   3540
      TabIndex        =   16
      Top             =   1440
      Width           =   3255
   End
   Begin VB.ComboBox cmbProveedores 
      Height          =   315
      Left            =   3600
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   630
      Width           =   2505
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00544B45&
      Caption         =   "Seleccione Tipo de Factura"
      ForeColor       =   &H00FFFFFF&
      Height          =   1005
      Left            =   390
      TabIndex        =   5
      Top             =   330
      Width           =   2955
      Begin VB.OptionButton optV 
         BackColor       =   &H00544B45&
         Caption         =   "Factura de Venta"
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   270
         TabIndex        =   1
         Top             =   570
         Width           =   2415
      End
      Begin VB.OptionButton optC 
         BackColor       =   &H00544B45&
         Caption         =   "Factura de Compra"
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   270
         TabIndex        =   0
         Top             =   270
         Value           =   -1  'True
         Width           =   2235
      End
   End
   Begin MSComctlLib.ListView lvFactura 
      Height          =   2025
      Left            =   390
      TabIndex        =   7
      Top             =   4680
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   3572
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cant"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Producto"
         Object.Width           =   4533
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "IdProd"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Importe"
         Object.Width           =   2293
      EndProperty
   End
   Begin tbrFaroButton.fBoton cmdModifCli 
      Height          =   435
      Left            =   4410
      TabIndex        =   28
      Top             =   4200
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Modificar"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdImprimir 
      Height          =   525
      Index           =   0
      Left            =   2700
      TabIndex        =   29
      Top             =   7845
      Width           =   2610
      _ExtentX        =   4604
      _ExtentY        =   926
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Imprimir Orden de Entrega"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdSalir 
      Height          =   435
      Left            =   9180
      TabIndex        =   33
      Top             =   7890
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin MSComCtl2.DTPicker DTFecha 
      Height          =   345
      Left            =   4740
      TabIndex        =   3
      Top             =   1050
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   61276161
      CurrentDate     =   39197
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Desde el día"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   3540
      TabIndex        =   34
      Top             =   1140
      Width           =   1005
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Otros Conceptos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   750
      TabIndex        =   25
      Top             =   7410
      Width           =   1455
   End
   Begin VB.Label lblOtrosConc 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   405
      Left            =   2250
      TabIndex        =   24
      Top             =   7320
      Width           =   1380
   End
   Begin VB.Label lblFueEntr 
      BackStyle       =   0  'Transparent
      Caption         =   "La mercadería de esta factura ya fue entregada"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   510
      TabIndex        =   4
      Top             =   3720
      Width           =   5265
   End
   Begin VB.Label lblDireccion 
      BackStyle       =   0  'Transparent
      Caption         =   "Facturas ya agregadas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   450
      TabIndex        =   21
      Top             =   4320
      Width           =   4095
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Facturas ya agregadas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   7410
      TabIndex        =   20
      Top             =   2340
      Width           =   2565
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Productos a entregar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   6840
      TabIndex        =   18
      Top             =   3990
      Width           =   3675
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione Factura"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   480
      TabIndex        =   15
      Top             =   1500
      Width           =   2475
   End
   Begin VB.Label lblDescuentos 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   405
      Left            =   2250
      TabIndex        =   14
      Top             =   6750
      Width           =   1380
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descuentos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1020
      TabIndex        =   13
      Top             =   6840
      Width           =   1185
   End
   Begin VB.Label lblIVA 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   405
      Left            =   4395
      TabIndex        =   12
      Top             =   6780
      Width           =   1380
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "IVA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   3690
      TabIndex        =   11
      Top             =   6840
      Width           =   525
   End
   Begin VB.Label lblFacturadeQue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Factura adeudada"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   360
      Left            =   180
      TabIndex        =   10
      Top             =   3960
      Width           =   5385
   End
   Begin VB.Label lblPesosF 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   405
      Left            =   4395
      TabIndex        =   9
      Top             =   7320
      Width           =   1380
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   3060
      TabIndex        =   8
      Top             =   7380
      Width           =   1185
   End
   Begin VB.Label lblSelecciona 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione Proveedor"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   3660
      TabIndex        =   6
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "frmVerFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IDVenta As Long
Dim ProductosaEntregar() As String 'con los id agregados a entregar

Private Sub chkEnt_Click()
    If chkEnt Then
        Me.Width = 11200
        lvCamion.Visible = True
        Label5.Visible = True
        Label6.Visible = True
        lstFacturasC.Visible = True
        cmdADD.Visible = True
        cmdLimpiar.Visible = True
        cmdImprimir(1).Visible = True
        cmdEntregar.Visible = True
        cmbSucursales.Visible = True
        
        tbrBuscador1.SqlSinLike = "SELECT Facturas.NroFactura, Clientes.Nombre, " + _
            "Facturas.Fecha, Facturas.Pagado " + _
            "FROM Clientes INNER JOIN Facturas ON Clientes.ID = Facturas.IdCliente " + _
            "WHERE Facturas.Entregado = 0 AND Fecha >= #" + stFechaSQL(DTFecha) + "#"
    Else
        ReDim ProductosaEntregar(0)
        ProductosaEntregar(0) = "Nada"
        lstFacturasC.Clear
        
        lvCamion.ListItems.Clear
        lvCamion.Visible = False
        Label5.Visible = False
        Label6.Visible = False
        lstFacturasC.Visible = False
        cmdADD.Visible = False
        cmdLimpiar.Visible = False
        cmdImprimir(1).Visible = False
        cmdEntregar.Visible = False
        cmbSucursales.Visible = False
        Me.Width = 8200
        
        tbrBuscador1.SqlSinLike = "SELECT Facturas.NroFactura, Clientes.Nombre, " + _
            "Facturas.Fecha, Facturas.Pagado " + _
            "FROM Clientes INNER JOIN Facturas ON Clientes.ID = " + _
            "Facturas.IdCliente WHERE Fecha >= #" + stFechaSQL(DTFecha) + "#"
    End If
    
    tbrBuscador1.OrderBy = "ORDER BY Facturas.ID DESC"
    tbrBuscador1.ColumnasSepPorComasyParentesis = "NroFactura(1650)/Cliente(1550)" + _
        "/Fecha(1100)/Importe(1300)"
    tbrBuscador1.CampoEnQueBuscar = "NroFactura/b,Nombre,Fecha,Pagado/$"
    tbrBuscador1.Recargar
End Sub

Private Sub cmbProveedores_Click()
    'cargo los tbrbuscadores
    
    If optC Then 'es compra
        tbrBuscador1.SqlSinLike = "SELECT NroFactura, Proveedor, Fecha, Pagado " + _
            "FROM FacturaCompra WHERE Fecha >= #" + stFechaSQL(DTFecha) + "#"
            
        If cmbProveedores <> "TODOS" Then
            tbrBuscador1.SqlSinLike = tbrBuscador1.SqlSinLike + _
                " AND Proveedor = '" + cmbProveedores + "'"
        End If
        
        tbrBuscador1.OrderBy = "ORDER BY Fecha DESC"
        tbrBuscador1.ColumnasSepPorComasyParentesis = "NroFactura(1650)/Proveedor(1550)/" + _
            "Fecha(1100)/Importe(1300)"
        tbrBuscador1.CampoEnQueBuscar = "NroFactura/b,Proveedor,Fecha,Pagado/$"
    Else 'es venta
        tbrBuscador1.SqlSinLike = "SELECT Facturas.NroFactura, Clientes.Nombre, " + _
            "Facturas.Fecha, Facturas.Pagado " + _
            "FROM Clientes INNER JOIN Facturas ON Clientes.ID = Facturas.IdCliente " + _
            "WHERE Fecha >= #" + stFechaSQL(DTFecha) + "#"
        
        If cmbProveedores <> "TODOS" Then
            Dim IDC As Long
            
            IDC = DB.GetValInRS("Clientes", "ID", "Nombre = '" + cmbProveedores + "'", False)
            
            tbrBuscador1.SqlSinLike = tbrBuscador1.SqlSinLike + _
                " AND IdCliente = " + CStr(IDC)
        End If
        
        tbrBuscador1.OrderBy = "ORDER BY Facturas.ID DESC"
        tbrBuscador1.ColumnasSepPorComasyParentesis = "NroFactura(1650)/Cliente(1550)" + _
            "/Fecha(1100)/Importe(1300)"
        tbrBuscador1.CampoEnQueBuscar = "NroFactura/b,Nombre,Fecha,Pagado/$"
    End If
    
    tbrBuscador1.Recargar
    If tbrBuscador1.GetLstSel <> "" Then tbrBuscador1_Click
End Sub

Private Sub cmdAdd_Click()
    Dim XX As Long, stIDP As String, xY As Long, Cant As Long, Ix As Long, Esta As Boolean
    If lvFactura.ListItems.Count = 0 Then Exit Sub
        
    Esta = False
        
    '(1) controlo que no haya cargado ya la factura
    For XX = 0 To UBound(ProductosaEntregar)
        If ProductosaEntregar(XX) = tbrBuscador1.GetLstSel Then
            MsgBox "Factura ya cargada para ser entregada", vbInformation, "Atención"
            Exit Sub
            Exit For
        End If
    Next XX
    
    '(2) agrego la factura a la matriz
    XX = UBound(ProductosaEntregar) + 1
    ReDim Preserve ProductosaEntregar(XX)
    ProductosaEntregar(XX) = tbrBuscador1.GetLstSel
    lstFacturasC.AddItem tbrBuscador1.GetLstSel
    
    '(3) agrego los productos incluidos en la factura
    For XX = 1 To lvFactura.ListItems.Count
        stIDP = lvFactura.ListItems(XX).SubItems(2)
        Cant = CLng(lvFactura.ListItems(XX).Text)
        '(3b) busco si ya lo habia cargado
        If lvCamion.ListItems.Count > 0 Then
            For xY = 1 To lvCamion.ListItems.Count
                If lvCamion.ListItems(xY).Text = stIDP Then 'si esta le sumo la cantidad
                    lvCamion.ListItems(xY).SubItems(2) = CStr(CLng(lvCamion. _
                        ListItems(xY).SubItems(2) + Cant))
                    Esta = True
                    Exit For '(del xy)
                End If
            Next xY
            'si no esta lo agrego
            If Esta = False Then
                Ix = lvCamion.ListItems.Count + 1
                lvCamion.ListItems.Add Ix
                lvCamion.ListItems(Ix).Text = stIDP
                lvCamion.ListItems(Ix).SubItems(1) = lvFactura.ListItems(XX).SubItems(1)
                lvCamion.ListItems(Ix).SubItems(2) = CStr(Cant)
            End If
            Esta = False
        Else
            Ix = lvCamion.ListItems.Count + 1
            lvCamion.ListItems.Add Ix
            lvCamion.ListItems(Ix).Text = stIDP
            lvCamion.ListItems(Ix).SubItems(1) = lvFactura.ListItems(XX).SubItems(1)
            lvCamion.ListItems(Ix).SubItems(2) = CStr(Cant)
        End If
        
    Next XX
    
End Sub

Private Sub cmdBorrarFactura_Click()
    'borra las facturas de compra
    If tbrBuscador1.GetLstSel = "" Then Exit Sub
    
    Dim Cuanto As Single, Deuda As Single, IdCF As Long, SP() As String
    Dim stAsD As String, stAsH As String, NoEsCompra As Single
    Dim stAsDM As String, stAsHM As String
    
    stAsD = "": stAsH = "": NoEsCompra = 0
    
    If MsgBox("¿Está seguro de borrar la factura y las deudas relacionadas de la misma?", _
        vbInformation + vbYesNo, "Atención") = vbNo Then Exit Sub
    
    '(1) leo el total de la factura
    Cuanto = DB.GetValInRS("FacturaCompra", "Pagado", "NroFactura = '" + _
        tbrBuscador1.GetLstSel + "'", False)
        
    '(2) sumo las deudas relacionadas con esa factura
    Deuda = DB.SumarValInRS("MovProveedores", "Variacion", "Documento = '" + _
        tbrBuscador1.GetLstSel + "'")
    
    '(3) modifico el stock y veo bien que tiene el compra detalle
    Dim RsT As New ADODB.Recordset, ClsP As New clsProducto
    
    RsT.Open "SELECT IDProducto, PrecioTotal, Cantidad FROM CompraDetalle " + _
        "WHERE NroFactura = '" + tbrBuscador1.GetLstSel + "'", DB.CN, adOpenStatic, adLockReadOnly
    
    If RsT.RecordCount > 0 Then
        RsT.MoveFirst
        Do While Not RsT.EOF
            If CLng(RsT("IDProducto")) >= 0 Then
                ClsP.ModificarStock CLng(RsT("IdProducto")), -CLng(RsT("Cantidad")), _
                    , "Por borrado de factura " + tbrBuscador1.GetLstSel
            Else 'IDProducto<0 veo que es
                Select Case CLng(RsT("IDProducto"))
                    Case -1 'descuento
                        'si es compra supongo que lo puso como costo
                        'igual creo que nunca pasa aca
                    Case -2 'iva
                        stAsH = stAsH + "50/"
                        stAsHM = stAsHM + CStr(RsT("PrecioTotal")) + "/"
                    Case Else 'otros conceptos
                        IdCF = CFG.ExistePropiedad("ConceptoCpra " + _
                            Right(CStr(RsT("IDProducto")), 1))
                        If IdCF = 0 Then
                            'ya no existe ese concepto no hago nada
                            'resto no es compra ya que al final lo suma asi se anula
                            NoEsCompra = NoEsCompra - CSng(RsT("PrecioTotal"))
                        Else
                            SP = Split(CFG.GetInfo(IdCF, 4), "_")
                            stAsH = stAsH + SP(0) + "/"
                            stAsHM = stAsHM + CStr(RsT("PrecioTotal")) + "/"
                        End If
                        
                End Select
                NoEsCompra = NoEsCompra + CSng(RsT("PrecioTotal"))
            End If
            RsT.MoveNext
        Loop
    End If
    RsT.Close
    Set RsT = Nothing
    Set ClsP = Nothing
    
    '(4) Leo las cuentas a usarse
    stAsD = stAsD + "78/"
    stAsDM = stAsDM + CStr(Cuanto - Deuda) + "/"
    
    stAsH = stAsH + "54/"
    stAsHM = stAsHM + CStr(Cuanto - NoEsCompra) + "/"
    
    If Deuda > 0 Then
        stAsD = stAsD + "41/"
        stAsDM = stAsDM + CStr(Deuda) + "/"
    End If
    
    stAsD = Left(stAsD, Len(stAsD) - 1)
    stAsH = Left(stAsH, Len(stAsH) - 1)
    stAsDM = Left(stAsDM, Len(stAsDM) - 1)
    stAsHM = Left(stAsHM, Len(stAsHM) - 1)
    
    '(5)hago el asiento que invierta todo
    PC.Asiento stAsD, stAsDM, stAsH, stAsHM, , "Borrado Factura " + tbrBuscador1.GetLstSel
        
    '(6)borro todo
    DB.EXECUTE "DELETE FROM FacturaCompra WHERE NroFactura = '" + tbrBuscador1.GetLstSel + "'"
    DB.EXECUTE "DELETE FROM MovProveedores WHERE Documento = '" + tbrBuscador1.GetLstSel + "'"
    
    '(7) recargo
    tbrBuscador1.Recargar
End Sub

Private Sub cmdEntregar_Click()
    If lvCamion.ListItems.Count = 0 Then Exit Sub
    '4 pasos
    Dim I As Long, IDp As Long, Cant As Long
    Dim ClsP As New clsProducto
    
    '(1) disminuir stock -------------------------------------------------------------
    For I = 1 To lvCamion.ListItems.Count
        IDp = CLng(lvCamion.ListItems(I).Text)
        Cant = CLng(lvCamion.ListItems(I).SubItems(2))
        
        ClsP.ModificarStock CLng(IDp), -Cant, cmbSucursales, _
            "Por Entrega el " + FormatDateTime(Date, vbGeneralDate)
    Next I
    Set ClsP = Nothing
    
    '(2) marcar como entregado las facturas marcadas
    For I = 0 To lstFacturasC.ListCount - 1
        DB.EXECUTE "UPDATE Facturas SET Entregado = 1 WHERE NroFactura = '" + _
            lstFacturasC.List(I) + "'"
    Next I
    
    '(3) pasar de merc en transito(55) a 'XXXX pendiente
    
    
    '(4)
    cmdLimpiar_Click
    chkEnt_Click
    tbrBuscador1.Recargar
    tbrBuscador1.Text = ""
End Sub

Private Sub cmdImprimir_Click(Index As Integer)
    Dim A As Long, Hta As Long, MiY As Single
    Dim tmP1 As String, tmp2 As String
    
    If tbrBuscador1.GetLstSel = "" Then Exit Sub
    
    If lvFactura.ListItems.Count = 0 And Index = 0 Or _
        lvCamion.ListItems.Count = 0 And Index = 1 Then
        MsgBox "¡No hay nada que imprimir!", vbInformation, "Impresión"
        Exit Sub
    End If
    
    Printer.Font.Name = "Arial"
    Printer.Font.Size = 12
    Printer.Font.Bold = True
        
    TP.tPrint 6000, 400, "Fecha: " + FormatDateTime(Date, vbShortDate)
    Dim clC As New clsCliente
        clC.AbrirDatos -2 'datos de mi empresa
        TP.tPrint 400, 800, clC.Nombre
        
        Printer.FontSize = 10
        TP.tPrint 400, 1200, "Direccion: " + clC.Direccion
        TP.tPrint 400, 1500, "CUIT: " + clC.CUIT
        TP.tPrint 8000, 1500, "Cond.IVA: " + clC.Iva, True
        Printer.DrawWidth = 6
        TP.PrintLINE 400, 1750, 8000, 1750
        
        Printer.FontSize = 12
        
    Set clC = Nothing
    
    If Index = 0 Then 'factura seleccionada (OE)
        If cmdImprimir(0).Caption = "Imprimir Orden de Entrega" Then
            TP.tPrint 400, 400, "Orden de Entrega " + tbrBuscador1.GetLstSel
        Else
            TP.tPrint 400, 400, "Copia de Factura " + tbrBuscador1.GetLstSel
        End If
        
        Hta = lvFactura.ListItems.Count
        
        If optV Then
            tmP1 = "Cliente: "
            tmp2 = DB.GetValInRS("Clientes", "Direccion", _
                "Nombre = '" + tbrBuscador1.GetLstSel(1) + "'")
        Else
            tmP1 = "Proveedor: "
            tmp2 = DB.GetValInRS("Proveedores", "Direccion", _
                "Proveedor = '" + tbrBuscador1.GetLstSel(1) + "'")
        End If
            
        Printer.FontSize = 10
        TP.tPrint 400, 1900, tmP1 + tbrBuscador1.GetLstSel(1)
        TP.tPrint 400, 2200, "Direccion: " + tmp2
        TP.PrintLINE 400, 2400, 8000, 2400
    Else 'pedidos a entregar
        TP.tPrint 400, 400, "Productos a ser entregados"
        Hta = lvCamion.ListItems.Count
        Printer.FontSize = 10
    End If
    
    MiY = 2700
    
    TP.tPrint 400, MiY, "Cant."
    TP.tPrint 1200, MiY, "Producto"
    If Index = 0 Then
        TP.tPrint 5000, MiY, "Precio Un."
        TP.tPrint 7000, MiY, "Precio Total"
    End If
    
    Printer.Font.Bold = False
    Printer.FontSize = 9
    
    MiY = MiY + 100
    
    For A = 1 To Hta
        MiY = MiY + 300
        If Index = 0 Then '(OE)
            TP.tPrint 400, MiY, lvFactura.ListItems(A).Text 'Cant
            TP.tPrint 1200, MiY, lvFactura.ListItems(A).SubItems(1) 'Producto
            TP.tPrint 6000, MiY, FormatCurrency(lvFactura.ListItems(A).SubItems(3)), True 'Pr Unit
            TP.tPrint 8000, MiY, FormatCurrency(CSng(lvFactura.ListItems(A).SubItems(3)) * _
                CLng(lvFactura.ListItems(A).Text)), True 'Precio Total
                
        Else '(PAEnt)
            TP.tPrint 400, MiY, lvCamion.ListItems(A).SubItems(2) 'Cant
            TP.tPrint 1200, MiY, lvCamion.ListItems(A).SubItems(1) 'Producto
        End If
        
    Next A
    
    If Index = 0 Then '(0E)
        Printer.FontBold = True
        MiY = MiY + 1000
        TP.tPrint 5500, MiY, "Otros Conc.: "
        TP.tPrint 8000, MiY, lblOtrosConc, True
        
        TP.tPrint 5500, MiY + 350, "Descuentos: "
        TP.tPrint 8000, MiY + 350, lblDescuentos, True
        
        TP.tPrint 5500, MiY + 700, "IVA: "
        TP.tPrint 8000, MiY + 700, lblIVA, True
        
        Printer.FontSize = 10
        TP.tPrint 5500, MiY + 1100, "Total"
        TP.tPrint 8000, MiY + 1100, lblPesosF, True
    Else 'A entregar pongo que facturas voy a entregar
        Dim Letr As Long, strR As String, H As Long
        
        Letr = 0: strR = ""
        Printer.FontBold = True
        MiY = MiY + 800
        TP.tPrint 400, MiY, "Facturas Incluidas: "
        
        Printer.FontBold = False
        Printer.FontSize = 8
        For H = 0 To lstFacturasC.ListCount - 1
            If Letr + Len(lstFacturasC.List(H)) >= 125 Then
                MiY = MiY + 300
                TP.tPrint 400, MiY, strR
                Letr = Len(lstFacturasC.List(H)) + 2
                strR = lstFacturasC.List(H)
            Else
                Letr = Letr + Len(lstFacturasC.List(H)) + 2
                strR = strR + lstFacturasC.List(H)
            End If
            If H < lstFacturasC.ListCount - 2 Then strR = strR + "; "
        Next H
    End If
    
    If MiY < 5000 Then
        MiY = MiY + 2000
    Else
        MiY = 7000
    End If
    
    TP.PrintLINE 400, MiY + 100, 8000, MiY + 100
    Printer.FontSize = 10
    Printer.FontBold = True
    TP.tPrint 8000, MiY + 300, "tbrSoft Desafios Digitales", True
    TP.tPrint 8000, MiY + 600, "CopyRight 2007 - info@tbrsoft.com", True, , , , , False
    
    TP.EndDocTP
End Sub

Private Sub cmdLimpiar_Click()
    ReDim ProductosaEntregar(0)
    ProductosaEntregar(0) = "Nada"
    lvCamion.ListItems.Clear
    lstFacturasC.Clear
End Sub

Private Sub cmdModifCli_Click()
    If tbrBuscador1.GetLstSel = "" Then Exit Sub
        
    Dim ExNroFac As String
    Dim IDC As Long
    
    ExNroFac = tbrBuscador1.GetLstSel
    
    If optC Then
        frmProveedores.Show 1
    Else
        IDC = DB.GetValInRS("Clientes", "ID", "Nombre = '" + tbrBuscador1.GetLstSel(1) + "'", False)
        frmClientes.AbrirDatos IDC
    End If
    
    tbrBuscador1.Text = ExNroFac
    tbrBuscador1.Recargar
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Public Sub AbrirDatos(IdVta As Long, Optional EsCompra As Boolean = False)
    IDVenta = IdVta
    
    If EsCompra Then
        optC.Value = True
        optC_Click
    Else
        optV.Value = True
        optV_Click
    End If
    
    Me.Show 1
End Sub

Private Sub DTFecha_Change()
    cmbProveedores_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    cmbSucursales.Clear
    cmbSucursales.AddItem "CASA CENTRAL"
    CargarCombo cmbSucursales, "SELECT * FROM Sucursales", "Sucursal", , True
    cmbSucursales.ListIndex = 0 'si no hay sucursales no lo hace
    
    lblIVA = FormatCurrency(0)
    lblDescuentos = FormatCurrency(0)
    lblPesosF = FormatCurrency(0)
    lblOtrosConc = FormatCurrency(0)
    
    tbrBuscador1.Contrasena = Contrasena
    tbrBuscador1.ArchivoMDB = ArchivoMDBPrincipal
    
    ReDim Preserve ProductosaEntregar(0)
    ProductosaEntregar(0) = "Nada"
    DTFecha = Date
    
    chkEnt.Value = 0
    chkEnt_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    tbrBuscador1.CN_CLOSE
End Sub

Private Sub optC_Click()
    If optC Then
        cmbProveedores.Clear
        cmbProveedores.AddItem "TODOS"
        CargarCombo cmbProveedores, "SELECT Proveedor FROM Proveedores", "Proveedor", , True
        
        lblSelecciona = "Seleccione Proveedor"
        
        cmdBorrarFactura.Visible = True
        chkEnt.Value = 0
        chkEnt_Click
        chkEnt.Enabled = False
        cmdImprimir(0).Enabled = False
        
        cmbProveedores_Click
    End If
    tbrBuscador1.Text = ""
End Sub

Private Sub optV_Click()
    If optV Then
        cmbProveedores.Clear
        cmbProveedores.AddItem "TODOS"
        CargarCombo cmbProveedores, "SELECT Nombre FROM Clientes WHERE ID >=0", _
            "Nombre", , True
        
        lblSelecciona = "Seleccione Cliente"
        
        cmdBorrarFactura.Visible = False
        chkEnt.Enabled = True
        cmdImprimir(0).Enabled = True
        
        tbrBuscador1.Recargar
    End If
    tbrBuscador1.Text = ""
End Sub

Private Sub tbrBuscador1_Change()
    If tbrBuscador1.GetLstSel = "" Then
        lvFactura.ListItems.Clear
        Exit Sub
    End If
    
    tbrBuscador1_Click
End Sub

Private Sub tbrBuscador1_Click()
    Dim NroFac As String, Iva As Single, SS As String, Desc As Single, OtrosC As Single
    
    If tbrBuscador1.GetLstSel = "" Then
        lvFactura.ListItems.Clear
        Exit Sub
    End If
    
    NroFac = tbrBuscador1.GetLstSel
    
    If optV Then 'ES VENTA!!!!!!!!!!!!!!!!!!!
        SS = "SELECT Productos.nProducto, Ventas.IdProducto, Ventas.Cantidad, Ventas.Precio " + _
            "FROM Productos INNER JOIN Ventas ON Productos.ID = " + _
            "Ventas.IDproducto WHERE Ventas.IdVenta = '" + NroFac + "' " + _
            " AND Ventas.IdProducto > 0 " + _
            "GROUP BY Productos.nProducto, Ventas.IdProducto, Ventas.Cantidad, Ventas.Precio"
    
        CargarComboLV lvFactura, SS, "Cantidad/n,nProducto,IdProducto/n,Precio/$"
        Iva = DB.GetValInRS("Ventas", "Precio", "IdProducto = -2 AND " + _
            "IdVenta = '" + tbrBuscador1.GetLstSel + "'", False)
        OtrosC = DB.SumarValInRS("Ventas", "Precio", "IdProducto <-10 AND " + _
            "IdProducto >-19 AND IdVenta = '" + tbrBuscador1.GetLstSel + "'")
        Desc = DB.GetValInRS("Ventas", "Precio", "IdProducto = -1 AND " + _
            "IdVenta = '" + tbrBuscador1.GetLstSel + "'", False)
        
        lblFacturadeQue = "Factura de Venta a " + UCase(tbrBuscador1.GetLstSel(1))
        lblDireccion = "Direccion: " + DB.GetValInRS("Clientes", "Direccion", _
                "Nombre = '" + tbrBuscador1.GetLstSel(1) + "'")
        
        lblFueEntr.Visible = True
        If DB.GetValInRS("Facturas", "Entregado", "NroFactura = '" + NroFac + "'", False) = 0 Then
            lblFueEntr.ForeColor = vbYellow
            lblFueEntr = "Factura pendiente de ser entregada"
            cmdImprimir(0).Caption = "Imprimir Orden de Entrega"
        Else
            lblFueEntr.ForeColor = vbWhite
            lblFueEntr = "La mercadería de esta factura ya fue entregada"
            cmdImprimir(0).Caption = "Imprimir Copia Factura"
        End If
    Else 'ES COMPRA!!!!!!!!!!!!!!
        SS = "SELECT Productos.nProducto, CompraDetalle.IdProducto, " + _
            "CompraDetalle.Cantidad, CompraDetalle.PrecioTotal " + _
            "FROM Productos INNER JOIN CompraDetalle ON Productos.ID = " + _
            "CompraDetalle.IDproducto WHERE CompraDetalle.NroFactura = '" + _
            CStr(NroFac) + "' AND CompraDetalle.IdProducto > 0 " + _
            "GROUP BY Productos.nProducto, CompraDetalle.IdProducto, " + _
            "CompraDetalle.Cantidad, CompraDetalle.PrecioTotal"
    
        CargarComboLV lvFactura, SS, "Cantidad/n,nProducto,IdProducto/n,PrecioTotal/$"
        
        Iva = DB.GetValInRS("CompraDetalle", "PrecioTotal", "IdProducto = -2 AND " + _
            "NroFactura = '" + tbrBuscador1.GetLstSel + "'", False)
        OtrosC = DB.SumarValInRS("CompraDetalle", "PrecioTotal", "IdProducto <-10 AND " + _
            "IdProducto >-19 AND NroFactura = '" + tbrBuscador1.GetLstSel + "'")
        Desc = DB.GetValInRS("CompraDetalle", "PrecioTotal", "IdProducto = -1 AND " + _
            "NroFactura = '" + tbrBuscador1.GetLstSel + "'", False)
            
        lblFacturadeQue = "Factura de Compra a " + UCase(tbrBuscador1.GetLstSel(1))
        lblDireccion = "Dirección: " + DB.GetValInRS("Proveedores", "Direccion", _
                "Proveedor = '" + tbrBuscador1.GetLstSel(1) + "'")
        lblFueEntr.Visible = False
    End If
    
    lblPesosF = tbrBuscador1.GetLstSel(3)
    lblOtrosConc = FormatCurrency(OtrosC, , , , vbFalse)
    lblIVA = FormatCurrency(Iva, , , , vbFalse)
    lblDescuentos = FormatCurrency(Desc, , , , vbFalse)
End Sub
