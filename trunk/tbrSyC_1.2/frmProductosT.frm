VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProductosT 
   BackColor       =   &H004E4E4E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listado de Productos"
   ClientHeight    =   9405
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   11190
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProductosT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   11190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin tbrFaroButton.fBoton cmdTipoProd 
      Height          =   375
      Left            =   6120
      TabIndex        =   9
      Top             =   930
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Nuevo Tipo Producto"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H004E4E4E&
      Caption         =   "Columnas para ver"
      ForeColor       =   &H00FFFFFF&
      Height          =   2325
      Left            =   8550
      TabIndex        =   7
      Top             =   1860
      Width           =   2235
      Begin tbrFaroButton.fBoton cmdStockPor 
         Height          =   435
         Left            =   210
         TabIndex        =   12
         Top             =   1830
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   767
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "ver por sucursal"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5717301
      End
      Begin VB.CheckBox chkIncIVA 
         BackColor       =   &H004E4E4E&
         Caption         =   "Incidencia IVA Costo"
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   240
         TabIndex        =   8
         Top             =   1380
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.CheckBox chkStock 
         BackColor       =   &H004E4E4E&
         Caption         =   "Stock"
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   990
         Width           =   1845
      End
      Begin VB.CheckBox chkCosto 
         BackColor       =   &H004E4E4E&
         Caption         =   "Costo"
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   660
         Width           =   1845
      End
      Begin VB.CheckBox chkPrecio 
         BackColor       =   &H004E4E4E&
         Caption         =   "Precio"
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   390
         Value           =   1  'Checked
         Width           =   1845
      End
   End
   Begin VB.ComboBox cmbTipoS 
      Height          =   315
      Left            =   420
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1440
      Width           =   7875
   End
   Begin MSComctlLib.ListView lvProductos 
      Height          =   7065
      Left            =   390
      TabIndex        =   4
      Top             =   1830
      Width           =   7905
      _ExtentX        =   13944
      _ExtentY        =   12462
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Producto"
         Object.Width           =   3440
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Precio"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Costo"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Stock"
         Object.Width           =   2117
      EndProperty
   End
   Begin tbrFaroButton.fBoton cmdImprimir 
      Height          =   405
      Left            =   8670
      TabIndex        =   10
      Top             =   7890
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Imprimir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdSalir 
      Height          =   405
      Left            =   8670
      TabIndex        =   11
      Top             =   8460
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione Tipo Producto"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   420
      TabIndex        =   6
      Top             =   1140
      Width           =   4155
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Listado de Productos"
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
      Height          =   315
      Left            =   2760
      TabIndex        =   5
      Top             =   180
      Width           =   3855
   End
End
Attribute VB_Name = "frmProductosT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkCosto_Click()
    If lvProductos.ListItems.Count = 0 Then Exit Sub

    Dim UUs As Long
    'veo si el usuario que esta trabajando tiene habilitacion para entrar -------
    UUs = ACC.UltUsuarioIngresado
    
    If ACC.ExisteRelacion(UUs, 25) = 0 Then
        MsgBox ACC.GetNombre("Usuario", "Usuarios", UUs) + " no está habilitado " + _
            "para ingresar." + vbCrLf + _
            "Debe Cambiar Sesión a la de un usuario habilitado", vbExclamation, "Atención"
        Exit Sub
    End If
    
    'registro el movimiento(25:Ver Costos y Otros)
    ACC.RegEvento UUs, 25, "Ver Costo del listado"
    '------------------------------------------------------------------

    If chkCosto.Value = 0 Then
        If chkPrecio.Value = 0 And chkStock.Value = 0 Then
            MsgBox "Debe haber al menos 1 columna", vbInformation, "Atención"
            chkCosto.Value = 1
            Exit Sub
        End If
        chkIncIVA.Visible = False
        Frame1.Height = 2115
        chkStock.Top = 1020
        cmdStockPor.Top = 1530
    Else
        chkIncIVA.Visible = True
        chkIncIVA.Top = 1020
        Frame1.Height = 2430
        chkStock.Top = 1350
        cmdStockPor.Top = 1860
    End If
    
    RecargarLVW
End Sub

Private Sub chkIncIVA_Click()
    If lvProductos.ListItems.Count = 0 Then Exit Sub
    RecargarLVW
End Sub

Private Sub chkPrecio_Click()
    If chkPrecio.Value = 0 Then
        If chkCosto.Value = 0 And chkStock.Value = 0 Then
            MsgBox "Debe haber al menos 1 columna", vbInformation, "Atención"
            chkPrecio.Value = 1
            Exit Sub
        End If
    End If
    RecargarLVW
End Sub

Private Sub chkStock_Click()
    If lvProductos.ListItems.Count = 0 Then Exit Sub

    If chkStock.Value = 0 Then
        cmdStockPor.Enabled = False
        If chkCosto.Value = 0 And chkPrecio.Value = 0 Then
            MsgBox "Debe haber al menos 1 columna", vbInformation, "Atención"
            chkStock.Value = 1
            Exit Sub
        End If
    Else
        cmdStockPor.Enabled = True
    End If
    RecargarLVW
End Sub

Private Sub cmbTipos_Click()
    RecargarLVW
End Sub

Private Sub cmdImprimir_Click()
    Dim Tit() As String

    TP.LineasSeparadoras = True
    
    ReDim Preserve Tit(4)
    Tit(4) = "Listado de Productos"
    'datos de mi empresa!!!!!!!!!!!!!!
    Tit(0) = DB.GetValInRS("Clientes", "Nombre", "ID = -2", True)
    Tit(1) = "Direccion: " + DB.GetValInRS("Clientes", "Direccion", "ID = -2", True)
    Tit(2) = "Teléfono: " + DB.GetValInRS("Clientes", "Telefono", "ID = -2", True)
    Tit(3) = "Mail: " + DB.GetValInRS("Clientes", "Mail", "ID = -2", True)
    
    TP.ImprimirlvW lvProductos, Tit, "IDP|Producto|Precio|Costo|Stock", _
        "Tipo: " + cmbTIPOS
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdStockPor_Click()
    If DB.ContarReg("SELECT * FROM Sucursales") = 0 Then
        MsgBox "No tiene más sucursales que CASA CENTRAL", vbInformation, "Atención"
    Else
        frmStockSucursales.Show 1
    End If
End Sub

Private Sub cmdTipoProd_Click()
    frmTipoProducto.Show 1
End Sub

Private Sub Form_Activate()
    Dim X As Long, ClsP As New clsProducto
    Dim Mat() As String
    
    'si o si tiene mas de un renglon
    Mat = ClsP.GetHijoTipo(0)
    cmbTIPOS.Clear
    cmbTIPOS.AddItem "TODOS"
    
    For X = 1 To UBound(Mat)
        cmbTIPOS.AddItem Mat(X) + " | " + _
            DB.GetValInRS("TipoProductos", "TipoProducto", "ID2 = " + Mat(X))
        AgregarHijos cmbTIPOS.ListCount - 1, 1
    Next X
    
    Set ClsP = Nothing
    cmbTIPOS.ListIndex = 0
End Sub

Private Sub AgregarHijos(Renglon As Long, Nivel As Long)
    Dim Y As Long, Mat() As String, SP() As String
    Dim ClsP As New clsProducto, Reng As Long, Niv As Long
    
    SP = Split(cmbTIPOS.List(Renglon), " | ")
    Mat = ClsP.GetHijoTipo(CLng(Trim(SP(0))))
    Reng = Renglon
    Niv = Nivel + 1
    
    If UBound(Mat) > 0 Then
        For Y = 1 To UBound(Mat)
            Reng = Reng + 1
            cmbTIPOS.AddItem String(Niv * 3, " ") + Mat(Y) + " | " + _
                DB.GetValInRS("TipoProductos", "TipoProducto", "ID2 = " + Mat(Y)), Reng
            AgregarHijos Reng, Niv + 1
        Next Y
    End If
    
    Set ClsP = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    'Me.Width = 11800
    'Frame1.Left = 9200
    'lvProductos.Width = 7400
    'cmdSalir.Left = Frame1.Left + 500
    'cmdImprimir.Left = 4000
End Sub

Private Sub RecargarLVW()
    Dim Ancho As Long, stSQL As String, I As Long
    
    lvProductos.ListItems.Clear
    lvProductos.ColumnHeaders(3).Width = 0
    lvProductos.ColumnHeaders(4).Width = 0
    lvProductos.ColumnHeaders(5).Width = 0
    
    Ancho = 1150
    If chkPrecio Then
        lvProductos.ColumnHeaders(3).Width = 1400
        Ancho = Ancho + 1400
    End If
    
    If chkCosto Then
        lvProductos.ColumnHeaders(4).Width = 1400
        Ancho = Ancho + 1400
    End If
    
    If chkStock Then
        lvProductos.ColumnHeaders(5).Width = 1300
        Ancho = Ancho + 1300
    End If
        
    lvProductos.ColumnHeaders(2).Width = lvProductos.Width - Ancho
    'cargo nomas
    stSQL = "SELECT ID, nProducto, pVenta, pCosto, Stock FROM Productos WHERE ID >0"
    If cmbTIPOS <> "TODOS" Then
        Dim SP() As String
        
        SP = Split(cmbTIPOS, "|")
        stSQL = stSQL + " AND IDTipoProducto = " + SP(0)
    End If
    
    CargarComboLV lvProductos, stSQL, "ID/n,nProducto,pVenta/$,pCosto/$,Stock/n"
    
    If chkIncIVA And chkCosto Then 'multiplico toda la columna
        If lvProductos.ListItems.Count = 0 Then Exit Sub
        Dim AjI As Single, IDC As Long, TmP As String
        
        For I = 1 To lvProductos.ListItems.Count
            'si tiene configuracion particular la pongo si no uso la general
            IDC = CFG.ExistePropiedad("IVA " + lvProductos.ListItems(I).Text)
            
            If IDC = 0 Then
                TmP = CFG.GetInfo(7, 4)
            Else
                TmP = CFG.GetInfo(IDC, 4)
            End If
            
            If Not IsNumeric(TmP) Then TmP = "0"
            
            lvProductos.ListItems(I).SubItems(3) = FormatCurrency(DB.GetValInRS("Productos", _
                "pCosto", "ID = " + CStr(lvProductos.ListItems(I).Text), False) * _
                ((100 + CSng(TmP)) / 100), , , , vbFalse)
        Next I
    End If
End Sub
