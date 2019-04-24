VERSION 5.00
Object = "{A7FBD38D-2930-49E3-B60C-9E0202D84549}#17.0#0"; "tbrControles.ocx"
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStockSucursales 
   BackColor       =   &H00544B45&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stock Por Sucursales"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStockSucursales.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin tbrFaroButton.fBoton cmdModificar 
      Height          =   420
      Left            =   6075
      TabIndex        =   22
      Top             =   6360
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   741
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Modificar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdSalir 
      Height          =   420
      Left            =   9675
      TabIndex        =   21
      Top             =   7800
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   741
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.ComboBox cmbSucursal3 
      Height          =   315
      Left            =   5850
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   5970
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00544B45&
      Caption         =   "Traspasos de Mercaderia"
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   6750
      Width           =   7335
      Begin tbrFaroButton.fBoton cmdTraspaso 
         Height          =   420
         Left            =   5115
         TabIndex        =   20
         Top             =   495
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   741
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Realizar Traspaso"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5717301
      End
      Begin VB.ComboBox cmbSucursal2 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   500
         Width           =   1845
      End
      Begin VB.ComboBox cmbSucursal1 
         Height          =   315
         Left            =   210
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   500
         Width           =   1845
      End
      Begin tbrControles.MouTextBox txtCant 
         Height          =   375
         Left            =   4200
         TabIndex        =   17
         Top             =   510
         Width           =   885
         _ExtentX        =   1561
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
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   4260
         TabIndex        =   13
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Mover Producto De"
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   270
         TabIndex        =   10
         Top             =   240
         Width           =   1665
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "a la sucursal"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   2100
         TabIndex        =   9
         Top             =   240
         Width           =   1785
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00544B45&
      Caption         =   "Filtros"
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   135
      TabIndex        =   2
      Top             =   5580
      Width           =   4755
      Begin tbrFaroButton.fBoton cmdLimpiar 
         Height          =   420
         Left            =   3270
         TabIndex        =   19
         Top             =   480
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   741
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Quitar Filtro"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5717301
      End
      Begin VB.ComboBox cmbTipo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   500
         Width           =   1845
      End
      Begin tbrControles.MouTextBox txtPrecio 
         Height          =   405
         Left            =   2250
         TabIndex        =   18
         Top             =   480
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   714
         Alignment       =   2
         BackColor       =   16777215
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Con Precio mayores a"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   2280
         TabIndex        =   6
         Top             =   240
         Width           =   2325
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Por Tipo de Producto"
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   270
         TabIndex        =   4
         Top             =   240
         Width           =   1875
      End
   End
   Begin MSComctlLib.ListView lvSucursales 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   900
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   7858
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Producto"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "C.Cent."
         Object.Width           =   1499
      EndProperty
   End
   Begin tbrControles.MouTextBox txtStockSucu 
      Height          =   375
      Left            =   5010
      TabIndex        =   16
      Top             =   6390
      Width           =   975
      _ExtentX        =   1720
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
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Stock en"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   4860
      TabIndex        =   12
      Top             =   6030
      Width           =   945
   End
   Begin VB.Label lblProducto 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mesa Grande"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   4950
      TabIndex        =   15
      Top             =   5730
      Width           =   2565
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Producto Seleccionado"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   5040
      TabIndex        =   14
      Top             =   5490
      Width           =   2385
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Productos según Sucursal"
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
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   210
      Width           =   7305
   End
End
Attribute VB_Name = "frmStockSucursales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sucursales() As String

Private Sub cmbSucursal3_Click()
    Dim IDp As Long
    Dim ClsP As New clsProducto
    
    If lvSucursales.ListItems.Count > 0 Then
        IDp = CLng(txtInLvW(lvSucursales, lvSucursales.SelectedItem.Index, 0))
    End If
    
    txtStockSucu = CStr(ClsP.StockProductoenSucursal(IDp, cmbSucursal3))
    
    Set ClsP = Nothing
End Sub

Private Sub cmbTipo_Click()
    CargarProductos
End Sub

Private Sub cmdLimpiar_Click()
    cmbTipo = "TODOS"
    txtPrecio = FormatCurrency(0)
    CargarProductos
End Sub

Private Sub cmdModificar_Click()
    Dim ClsP As New clsProducto
    Dim UUs As Long
    
    txtStockSucu = ValidarNumeros(txtStockSucu)
    
    UUs = ACC.UltUsuarioIngresado
    
    If ACC.ExisteRelacion(UUs, 9) = 0 Then
        MsgBox ACC.GetNombre("Usuario", "Usuarios", UUs) + " no está habilitado " + _
            "para ingresar." + vbCrLf + _
            "Debe Cambiar Sesión a la de un usuario habilitado", vbExclamation, "Atención"
        Exit Sub
    End If
    
    'registro el movimiento(9:Modificar Producto)
    ACC.RegEvento UUs, 9, "Modificar stock en sucursales"

    'registro perdida o ganancia por diferencias de stock
    Dim stockV As Long, StockN As Long, IDp As Long
    
    IDp = CLng(txtInLvW(lvSucursales, lvSucursales.SelectedItem.Index, 0))
    
    stockV = CStr(ClsP.StockProductoenSucursal(IDp, cmbSucursal3))
    StockN = CLng(txtStockSucu)
        
    'recien aca veo que esta todo bien
    Dim DifS As Long, DifPesos As Single
    
    DifS = StockN - stockV
    DifPesos = ClsP.GetCosto(IDp) * DifS
    
    'primero lo facil modifico stock
    ClsP.ModificarStock IDp, CSng(DifS), cmbSucursal3, "Modificación Manual en Sucursal"
    
    'ahora ajusto mercaderias en inventario a Diferencias de stock
    'si sobra si falta se hace al revés por quedar negativo
    PC.Asiento "54", CStr(DifPesos), "35", CStr(DifPesos), , _
        "Modificación Stock, IdProducto: " + CStr(IDp) + " de " + cmbSucursal3
    
    CargarProductos
    
    Set ClsP = Nothing
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdTraspaso_Click()
    If ValidarNumeros(txtCant) = 0 Then Exit Sub
    If cmbSucursal1 = cmbSucursal2 Then Exit Sub
    
    Dim clP As New clsProducto, IDp As Long
    Dim OrigSV As Long, DestSV As Long
    Dim OrigSN As Long, DestSN As Long
    
    IDp = CLng(txtInLvW(lvSucursales, lvSucursales.SelectedItem.Index, 0))
    
    OrigSV = clP.StockProductoenSucursal(IDp, cmbSucursal1)
    DestSV = clP.StockProductoenSucursal(IDp, cmbSucursal2)
    OrigSN = OrigSV - CLng(txtCant)
    DestSN = DestSV + CLng(txtCant)
    
    Dim MsJ As String
    
    MsJ = "¿Está seguro de realizar el traspaso de " + vbCrLf + _
        txtCant + " productos de " + _
        UCase(txtInLvW(lvSucursales, lvSucursales.SelectedItem.Index, 1)) + vbCrLf + _
        "De " + cmbSucursal1 + " a " + cmbSucursal2 + "?"
        
    If OrigSN < 0 Then
        MsJ = MsJ + vbCrLf + "¡Quedará stock negativo en " + cmbSucursal1 + "!"
    End If
    
    If DestSN < 0 Then
        MsJ = MsJ + vbCrLf + "¡Quedará stock negativo en " + cmbSucursal2 + "!"
    End If
    
    If MsgBox(MsJ, vbInformation + vbYesNo, "Atención") = vbNo Then Exit Sub
    
    'gueno hago el cambio (no es necesario asiento contable ya que solo es un traspaso)
    '(1) descuento el stock del origen
    clP.ModificarStock IDp, -CLng(txtCant), cmbSucursal1, "Traspaso de Sucursal " + _
        cmbSucursal1 + " a " + cmbSucursal2
    '(2) aumento el stock en destino
    clP.ModificarStock IDp, CLng(txtCant), cmbSucursal2, "Traspaso de Sucursal " + _
        cmbSucursal1 + " a " + cmbSucursal2
    '(3) cargo el lview con las modificaciones
    CargarProductos
    txtCant = 0
    Set clP = Nothing
End Sub

Private Sub Form_Load()
    txtPrecio = FormatCurrency(0)
    txtCant = "0"
    txtStockSucu = "0"
    
    CargarSucursales 'queda muy joya
    
    cmbTipo.Clear
    cmbTipo.AddItem "TODOS"
    CargarCombo cmbTipo, "SELECT Tipoproducto FROM Tipoproductos " + _
        "WHERE ID2 > 0 ORDER BY Tipoproducto", _
        "Tipoproducto", , True
        
    cmbSucursal1.Clear
    cmbSucursal1.AddItem "CASA CENTRAL"
    cmbSucursal2.Clear
    cmbSucursal2.AddItem "CASA CENTRAL"
    cmbSucursal3.Clear
    cmbSucursal3.AddItem "CASA CENTRAL"
    CargarCombo cmbSucursal1, "SELECT Sucursal FROM Sucursales", "Sucursal", , True
    CargarCombo cmbSucursal2, "SELECT Sucursal FROM Sucursales", "Sucursal", , True
    CargarCombo cmbSucursal3, "SELECT Sucursal FROM Sucursales", "Sucursal", , True
    lvSucursales_Click
End Sub

Private Sub CargarSucursales()
    'el largo total es:
        'como maximo = 11600 (9 sucursales de 840 de ancho c/u+2000 nombre y 800 ID,
            ' 850 casa central y 390 para scroll)
        'como minimo = 7350
        'si son mas voy a tener que hacer
    'acordarse que no esta en la tabla sucursales la casa central
    Dim NSuc As Long, AnchoCU As Single
    Dim Rss As New ADODB.Recordset, TmP As Long
    
    NSuc = DB.ContarReg("SELECT * FROM Sucursales")
    
    If NSuc = 0 Then
        AnchoCU = 1400
    Else
        AnchoCU = 7560 / NSuc
        'que no sea tan grande
        If AnchoCU > 1400 Then AnchoCU = 1400
    End If
    
    If Rss.State = adStateOpen Then Rss.Close
    
    Rss.Open "SELECT Sucursal FROM Sucursales", DB.CN, adOpenStatic, adLockReadOnly
    If Rss.RecordCount > 0 Then
        Rss.MoveFirst
        
        ReDim Sucursales(0)
        Do While Not Rss.EOF
            TmP = UBound(Sucursales) + 1
            ReDim Preserve Sucursales(TmP)
            Sucursales(TmP) = Rss("Sucursal")
            
            lvSucursales.ColumnHeaders.Add , , Rss("Sucursal"), AnchoCU, Center
            Rss.MoveNext
        Loop
        
        '¿cuanto queda de ancho el listview?
        lvSucursales.Width = AnchoCU * NSuc + 4040  ' (2800 ID y Nombre, 850 CC y 390 el scroll)
        If lvSucursales.Width < 7350 Then lvSucursales.Width = 7350
        
        '¿y el formulario?
        Me.Width = lvSucursales.Width + 400
        
        '¿El titulo y los botones?
        Label6.Left = lvSucursales.Width / 4 - 500
        cmdSalir.Left = lvSucursales.Left + lvSucursales.Width - cmdSalir.Width
    End If
    
    Rss.Close
    Set Rss = Nothing
End Sub

Private Sub CargarProductos()
    Dim RsSu As New ADODB.Recordset
    Dim TmP As Long, Y As Long, StockUsado As Long
    
    If RsSu.State = adStateOpen Then RsSu.Close
    RsSu.Open GetstFilt, DB.CN, adOpenStatic, adLockReadOnly
    
    lvSucursales.ListItems.Clear
    If RsSu.RecordCount > 0 Then
        TmP = 1
        RsSu.MoveFirst
        
        Do While Not RsSu.EOF
            StockUsado = 0
            lvSucursales.ListItems.Add TmP
            lvSucursales.ListItems(TmP).Text = CStr(RsSu("ID"))
            lvSucursales.ListItems(TmP).SubItems(1) = RsSu("nProducto")
            
            'ahora vamos horizontalmente para ver las sucursales
            'Casa central al ultimo se calcula por descarte
            For Y = 1 To UBound(Sucursales)
                lvSucursales.ListItems(TmP).SubItems(Y + 2) = CStr( _
                    DB.GetValInRS("StockOtraSuc", _
                    "Stock", "IdProducto = " + _
                    lvSucursales.ListItems(TmP).Text + _
                    " AND Sucursal = '" + Sucursales(Y) + "'", False))
                StockUsado = StockUsado + CLng(lvSucursales.ListItems(TmP).SubItems(Y + 2))
            Next Y
            
            lvSucursales.ListItems(TmP).SubItems(2) = NoNuloN(RsSu("Stock")) - StockUsado
            
            TmP = TmP + 1
            RsSu.MoveNext
        Loop
    End If
   
    RsSu.Close
    Set RsSu = Nothing
End Sub

Private Function GetstFilt() As String
    Dim Resp As String
    
    txtPrecio = FormatCurrency(ValidarNumeros(txtPrecio), , , , vbFalse)
    
    Resp = "SELECT Productos.ID, Productos.nProducto, Productos.Stock " + _
        "FROM TipoProductos INNER JOIN Productos ON TipoProductos.ID2 = " + _
        "Productos.IdTipoProducto " + _
        "WHERE Productos.ID >=0 AND Productos.pVenta > " + _
        Replace(CStr(CSng(txtPrecio)), ",", ".") + " "
        
    If cmbTipo <> "TODOS" Then
        Resp = Resp + "AND TipoProductos.TipoProducto = '" + cmbTipo + "' "
    End If

    Resp = Resp + "GROUP BY Productos.ID, Productos.nProducto, Productos.Stock " + _
        "ORDER BY Productos.nProducto"

    GetstFilt = Resp
End Function

Private Sub lvSucursales_Click()
    If lvSucursales.ListItems.Count = 0 Then Exit Sub
    
    lblProducto = txtInLvW(lvSucursales, lvSucursales.SelectedItem.Index, 1)
    cmbSucursal3_Click
End Sub

Private Sub lvSucursales_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtCant_GotFocus()
    PintarTxt txtCant
End Sub

Private Sub txtCant_LostFocus()
    txtCant = ValidarNumeros(txtCant)
End Sub

Private Sub txtPrecio_GotFocus()
    PintarTxt txtPrecio
End Sub

Private Sub txtPrecio_LostFocus()
    txtPrecio = FormatCurrency(ValidarNumeros(txtPrecio), , , , vbFalse)
    CargarProductos
End Sub

Private Sub txtStockSucu_GotFocus()
    PintarTxt txtStockSucu
End Sub

Private Sub txtStockSucu_LostFocus()
    txtStockSucu = ValidarNumeros(txtStockSucu)
End Sub
