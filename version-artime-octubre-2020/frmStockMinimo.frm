VERSION 5.00
Object = "{A7FBD38D-2930-49E3-B60C-9E0202D84549}#15.0#0"; "tbrControles.ocx"
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStockMinimo 
   BackColor       =   &H004E4E4E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Stock Mínimo"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9285
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStockMinimo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin tbrFaroButton.fBoton cmdStock 
      Height          =   420
      Left            =   5820
      TabIndex        =   7
      Top             =   7065
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   741
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Grabar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.CheckBox chkStockMinimo 
      BackColor       =   &H004E4E4E&
      Caption         =   "Ver solamente los que tengan stock menor al mínimo"
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   1710
      TabIndex        =   6
      Top             =   6510
      Width           =   5985
   End
   Begin VB.ComboBox cmbSucursales 
      Height          =   315
      Left            =   3630
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1200
      Width           =   2895
   End
   Begin tbrControles.MouTextBox txtStock 
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   7110
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   661
      Alignment       =   2
      BackColor       =   16777215
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView lvStock 
      Height          =   4695
      Left            =   1500
      TabIndex        =   4
      Top             =   1665
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   8281
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
         Text            =   "IdP"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Producto"
         Object.Width           =   4762
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Stock"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "St.Min."
         Object.Width           =   2117
      EndProperty
   End
   Begin tbrFaroButton.fBoton command1 
      Height          =   420
      Left            =   7590
      TabIndex        =   8
      Top             =   7065
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   741
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione Sucursal"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1140
      TabIndex        =   5
      Top             =   1230
      Width           =   1995
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Mínimo Predeterminado"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   450
      TabIndex        =   3
      Top             =   7170
      Width           =   3405
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Control de Stock Mínimo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   555
      Left            =   1905
      TabIndex        =   1
      Top             =   315
      Width           =   5475
   End
End
Attribute VB_Name = "frmStockMinimo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkStockMinimo_Click()
    CargarDatos
End Sub

Private Sub cmbSucursales_Click()
    CargarDatos
End Sub

Private Sub cmdStock_Click()
    txtStock = ValidarNumeros(txtStock)
    CFG.ModificarNodo 19, , , , txtStock
    
    CargarDatos
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Dim TmP As String
    
    'cargo sucursales
    cmbSucursales.Clear
    cmbSucursales.AddItem "CASA CENTRAL"
    CargarCombo cmbSucursales, "SELECT * FROM Sucursales", "Sucursal", , True
    cmbSucursales.ListIndex = 0 'si no hay sucursales no lo hace
        
    CargarDatos
End Sub

Private Sub CargarDatos()
    Dim ClsP As New clsProducto, I As Long, IdCF As Long, TmP As String
    Dim ParaBorrar() As String, PB As Long
    
    'Stock Minimo Predeterminado hago que controle
    TmP = CFG.GetInfo(19, 4)
    If TmP = "" Then TmP = "10"
    If Not IsNumeric(TmP) Then TmP = "10"
    txtStock = TmP
    CFG.ModificarNodo 19, , , , TmP
    
    'primero cargo lo de la base de datos
    CargarComboLV lvStock, "SELECT ID, nProducto FROM Productos WHERE ID>=0 " + _
        "ORDER BY ID", "ID/n,nProducto"
    
    'ya con la lista de productos le agrego stock y stockminimo
    For I = 1 To lvStock.ListItems.Count
        lvStock.ListItems(I).SubItems(2) = CStr(ClsP.StockProductoenSucursal( _
            lvStock.ListItems(I).Text, cmbSucursales))
        IdCF = CFG.ExistePropiedad("STM " + lvStock.ListItems(I).Text)
        
        If IdCF = 0 Then 'predeterminada
            lvStock.ListItems(I).SubItems(3) = txtStock
        Else
            lvStock.ListItems(I).SubItems(3) = CFG.GetInfo(IdCF, 4)
        End If
    Next I
    
    Set ClsP = Nothing
    
    'borro los que tengan stock al mínimo si es que eligio el chkbox
    If chkStockMinimo Then
        PB = 0
        ReDim ParaBorrar(PB)
        ParaBorrar(PB) = "Nada"
        
        For I = 1 To lvStock.ListItems.Count
            If CLng(lvStock.ListItems(I).SubItems(2)) > _
                CLng(lvStock.ListItems(I).SubItems(3)) Then
                PB = PB + 1
                ReDim Preserve ParaBorrar(PB)
                ParaBorrar(PB) = CStr(I)
            End If
        Next I
        
        'borro nomas
        'de atras para adelante estan en orden asi que
        If UBound(ParaBorrar) <= 0 Then Exit Sub
        For I = UBound(ParaBorrar) To 1 Step -1
            lvStock.ListItems.Remove CLng(ParaBorrar(I))
        Next I
    End If
    
End Sub
