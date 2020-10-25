VERSION 5.00
Object = "{57F2DDA9-75EB-4696-9A07-7D86D576ABB4}#27.0#0"; "tbrControles.ocx"
Begin VB.Form frmCompras 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Factura de Compra"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCompras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin tbrControles.MouTextBox txtCant 
      Height          =   405
      Left            =   930
      TabIndex        =   27
      Top             =   5040
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   714
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Forma de Pago"
      Height          =   885
      Left            =   4560
      TabIndex        =   22
      Top             =   6450
      Width           =   1545
      Begin VB.OptionButton chkContado 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Contado"
         Height          =   255
         Left            =   210
         TabIndex        =   24
         Top             =   270
         Value           =   -1  'True
         Width           =   945
      End
      Begin VB.OptionButton chkACuenta 
         BackColor       =   &H00C0C0C0&
         Caption         =   "A Cuenta"
         Height          =   255
         Left            =   210
         TabIndex        =   23
         Top             =   540
         Width           =   1035
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Agregar Proveedor"
      Height          =   555
      Left            =   780
      TabIndex        =   21
      Top             =   5610
      Width           =   1065
   End
   Begin VB.ListBox lstProductos 
      Height          =   2220
      Left            =   390
      TabIndex        =   19
      Top             =   1560
      Width           =   4065
   End
   Begin VB.TextBox txtNameP 
      Height          =   400
      Left            =   2280
      TabIndex        =   1
      Text            =   "sdasd"
      Top             =   840
      Width           =   2115
   End
   Begin VB.ComboBox cmbProveedores 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   2145
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar Pedido"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   6180
      TabIndex        =   17
      Top             =   7530
      Width           =   1515
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   405
      Left            =   1680
      TabIndex        =   16
      Top             =   3990
      Width           =   825
   End
   Begin VB.CommandButton cmdAgProd 
      Caption         =   "Agregar Producto"
      Height          =   555
      Left            =   2010
      TabIndex        =   14
      Top             =   5610
      Width           =   1095
   End
   Begin VB.ListBox lstPaReg 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      ItemData        =   "frmCompras.frx":030A
      Left            =   720
      List            =   "frmCompras.frx":0311
      TabIndex        =   13
      Top             =   8280
      Width           =   4545
   End
   Begin VB.CommandButton cmdDeMas 
      Caption         =   "Agregar Recargo"
      Height          =   345
      Left            =   6450
      TabIndex        =   12
      Top             =   3870
      Width           =   1395
   End
   Begin VB.CommandButton cmdQuitar 
      Caption         =   "<<<"
      Height          =   345
      Left            =   4680
      TabIndex        =   3
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   7860
      TabIndex        =   9
      Top             =   7530
      Width           =   975
   End
   Begin VB.CommandButton cmdPagar 
      Caption         =   "Pagar Factura"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   4590
      TabIndex        =   7
      Top             =   7530
      Width           =   1425
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   ">>>"
      Height          =   345
      Left            =   4680
      TabIndex        =   2
      Top             =   1830
      Width           =   615
   End
   Begin VB.ListBox lstFactura 
      BackColor       =   &H00E6E4D2&
      ForeColor       =   &H00000000&
      Height          =   2700
      ItemData        =   "frmCompras.frx":031F
      Left            =   5400
      List            =   "frmCompras.frx":0321
      TabIndex        =   4
      Top             =   840
      Width           =   4365
   End
   Begin tbrControles.MouTextBox txtPrecio 
      Height          =   405
      Left            =   2010
      TabIndex        =   28
      Top             =   5040
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   714
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
   Begin tbrControles.MouTextBox txtDeMas 
      Height          =   405
      Left            =   8730
      TabIndex        =   29
      Top             =   3840
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   714
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
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Id / Producto"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   450
      TabIndex        =   26
      Top             =   1290
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Para descuentos ingréselo negativo"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   6210
      TabIndex        =   25
      Top             =   6210
      Width           =   2805
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "BUSCAR Nombre/Codigo"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   420
      TabIndex        =   20
      Top             =   870
      Width           =   1845
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione Proveedor"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   18
      Top             =   240
      Width           =   1905
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Pr.Total"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   15
      Top             =   4740
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Recargo"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   7980
      TabIndex        =   11
      Top             =   3930
      Width           =   915
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "A Pagar"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   7860
      TabIndex        =   10
      Top             =   6510
      Width           =   885
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      BackColor       =   &H00E6E4D2&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "$ 5888,88"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6720
      TabIndex        =   8
      Top             =   6870
      Width           =   2085
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Factura de compra"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C9B44E&
      Height          =   675
      Left            =   5430
      TabIndex        =   6
      Top             =   150
      Width           =   4365
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Cant"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1230
      TabIndex        =   5
      Top             =   4740
      Width           =   465
   End
End
Attribute VB_Name = "frmCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ToTal As Single 'es el total de la factura
Dim Precio As Single 'precio de prods a cargar en factura
Dim DeMas As Single 'Es el recargo de la factura
Dim DtDeMas As String 'el detalle de la recarga
Dim nProveedor As String
Dim Nuevo As Boolean 'cuando pague quiero saber si ya estaba grabado como pedido
Dim CtoTmp As Single

Private Sub cmbProveedores_Click()
    nProveedor = cmbProveedores
    Label3 = "Factura de Compra de " + nProveedor
End Sub

Private Sub cmdAgProd_Click()
    frmProductos.AbrirDatos ""
End Sub

Private Sub Limpiar()
    ToTal = 0 '???? no se porque pero se guarda el valor de la factura vieja
    lstFactura.Clear
    txtDeMas = FormatCurrency(0)
    txtPrecio = FormatCurrency(0)
    lblTotal = FormatCurrency(0)
    
End Sub

Private Function GetID(Index As Integer) As Long
    Dim SP() As String
    SP = Split(lstProductos.List(Index), "/")
    
    GetID = CLng(SP(0))
End Function


Private Function GetNameO(Index As Integer) As String
    Dim SP() As String
    SP = Split(lstProductos.List(Index), "/")
    
    GetNameO = SP(1)
End Function

Private Sub cmdAgregar_Click()
        If lstProductos.ListCount = 0 Then MsgBox "No eligio ningún" + _
                " producto": Exit Sub
        
        Precio = ValidarNumeros(txtPrecio)
        txtPrecio = FormatCurrency(Precio, , , , vbFalse)
        
        Select Case Precio
            Case 0
                If MsgBox("Va a incluir un producto con precio 0, si no es asi presione " + _
                    "Cancelar", vbOKCancel + vbInformation, "Atención") = vbCancel Then Exit Sub
            Case Is < 0
                MsgBox "Ingrese sólo valores positivos", vbInformation, "Atención"
                Exit Sub
        End Select
                
        lstFactura.AddItem txtCant + "\" + GetNameO(lstProductos.ListIndex) + _
            "\" + FormatCurrency(CStr(Precio / CSng(txtCant)), 4, , , vbFalse) + "\" + _
            FormatCurrency(txtPrecio, , , , vbFalse) + "\" + CStr(GetID(lstProductos.ListIndex))
        ToTal = ToTal + Precio
        TotalT = ToTal
        
        lblTotal = FormatCurrency(ToTal, , , , vbFalse)
           
        cmdSel.Default = True
        PintarTxt txtNameP
        
End Sub


Private Sub cmdDeMas_Click()
    If lstFactura.ListCount = 0 Then Exit Sub
    If CSng(txtDeMas) = 0 Then Exit Sub
    DeMas = CSng(DeMas)
    
    If CSng(txtDesc) <> 0 Then 'es que ya puso la opcion de descuento
    
    Else
    
        If MsgBox("Solo seleccione esta opción si no piensa agregar mas productos, " + _
            "Esta por agregar una recarga de " + FormatCurrency(DeMas, , , , vbFalse) + ", ¿Es correcto?" + _
            "si no presione cancelar", vbOKCancel, "Incluir Recarga") = vbCancel Then Exit Sub
    End If
    
    'bloqueo esto para que no toquen mas al pedo
    cmdSel.Enabled = False
    cmdAgregar.Enabled = False
    cmdQuitar.Enabled = False
    txtCant.Enabled = False
    txtPrecio.Enabled = False
        
    DtDeMas = InputBox("Detalle el motivo de la recarga", "Motivo de recarga", DtDeMas)
    
    lstFactura.AddItem "Recarga: " + DtDeMas + "\" + FormatCurrency(DeMas, , , , vbFalse)
    ToTal = ToTal + DeMas
    lblTotal = FormatCurrency(ToTal, , , , vbFalse)
    'ver que hacer con descuento
    
    'tambien bloqueo para que no meta otro recargo
    cmdDeMas.Enabled = False
    txtDeMas.Enabled = False
    
End Sub

Private Sub cmdGrabar_Click() 'muy similar a pagar
    If lstFactura.ListCount <= 0 Then Exit Sub
    
    Dim DifP As Single 'es el porcentaje del precio que queda
    If MsgBox("¿Está seguro de grabar el pedido de " + nProveedor + "?", _
        vbOKCancel, "Atención") = vbCancel Then Exit Sub
    
    If ToTal = 0 Then Exit Sub
    If cmdDeMas.Enabled = True Then 'esto pasa cuando no agrego a la factura la recarga
        DifP = 1
    Else
        DifP = 1 + DeMas / (ToTal - DeMas)
    End If
    
     'ahora cargo el listbox listo para grabar el recordset
    Dim SP() As String
    Dim clsC As New clsCompras
     
    lstPaReg.Clear
    For I = 0 To lstFactura.ListCount - 1
        SP = Split(lstFactura.List(I), "\")
        If IsNumeric(SP(0)) Then
            lstPaReg.AddItem CStr(SP(0)) + "\" + SP(1) + "\" + CStr(SP(2) * DifP) + _
            "\" + CStr(SP(4))
        End If
    Next I
    
    Dim Spl() As String  'para el split
    
    Dim IDCt As Long
    IDCt = GetUltIDCmasUno 'si no hago esto y lo pongo en el bucle registra
                                 'para cada renglon und idcompra distinto
    
    For J = 0 To lstPaReg.ListCount - 1
        Spl = Split(lstPaReg.List(J), "\")
        clsC.CargarCompraDt IDCt, CStr(Spl(3)), CLng(Spl(0)), _
            Spl(0) * Spl(2), nProveedor, False
    Next J
    
    Set clsC = Nothing
    
    Unload Me
    
End Sub

Private Sub cmdPagar_Click()
    If lstFactura.ListCount <= 0 Then Exit Sub
    
    Dim tmP As String
    If chkContado Then
        tmP = "Al Contado"
    Else
        tmP = "A Cuenta"
    End If
    
    If MsgBox("Está por registrar la compra " + tmP + " de " + FormatCurrency(ToTal) + _
    " a " + nProveedor + ", ¿Son correctos los datos?", _
        vbOKCancel + vbInformation, "Atencion") = vbCancel Then Exit Sub
    
    Dim DifP As Single 'es el porcentaje del precio que queda
    
    If cmdDeMas.Enabled = True Then 'esto pasa cuando no agrego a la factura la recarga
        DifP = 1
    Else
        DifP = 1 + DeMas / (ToTal - DeMas)
    End If
    
     'ahora cargo el listbox listo para grabar el recordset
    Dim SP() As String
    Dim Spl() As String
    Dim clsC As New clsCompras
    
     'modifico stock y costo
    lstPaReg.Clear
    For I = 0 To lstFactura.ListCount - 1
        SP = Split(lstFactura.List(I), "\")
        If IsNumeric(SP(0)) Then
            lstPaReg.AddItem CStr(SP(0)) + "\" + SP(1) + "\" + CStr(SP(2) * DifP) + _
            "\" + CStr(SP(4))
            clsC.CargarCompra CLng(SP(4)), CLng(SP(0)), SP(2) * DifP
        End If
    Next I
    
    Dim IDCt As Long
    IDCt = GetUltIDCmasUno 'si no hago esto y lo pongo en el bucle registra
                                 'para cada renglon und idcompra distinto
    
     'cargo el detalle de compra
    For J = 0 To lstPaReg.ListCount - 1
        Spl = Split(lstPaReg.List(J), "\")
        clsC.CargarCompraDt IDCt, CStr(Spl(3)), CLng(Spl(0)), _
            Spl(0) * Spl(2), nProveedor, True
    Next J
    Set clsC = Nothing
    
    'registro la compras
    DB.Execute "INSERT INTO Compras (ID,Fecha,Proveedor,Monto) " + _
        "VALUES (" + IdAutonum("Compras") + ",#" + _
        stFechaSQL(Date) + "#,'" + nProveedor + _
        "'," + Replace(CStr(ToTal), ",", ".") + ")"
    
     'libro diario mercaderia a caja si es a cuenta se invierte en forma de pago
    
    PC.Asiento "54", CStr(ToTal), "78", CStr(ToTal), "LibroSubDiario", _
        "Compra Factura Nº " + CStr(IDCt) + " a " + nProveedor
    
    If chkACuenta Then frmClientesMov.AbrirDatos IDCt, True, cmbProveedores
    
    Unload Me
    
End Sub

Private Sub cmdQuitar_Click()
    Dim Dt() As String  'split para sacar pr y cant
    Dim Dto As Single ' lo que voy a restar del total
    
    If lstFactura.ListIndex = -1 Then Exit Sub
    
    Dt = Split(lstFactura, "\")
    lstFactura.RemoveItem (lstFactura.ListIndex)
    Dto = CSng(Dt(3))
    
    ToTal = ToTal - Dto
    lblTotal = FormatCurrency(ToTal, , , , vbFalse)
    
End Sub

Private Sub Command2_Click()
    frmProveedores.Show 1
End Sub

Private Sub Form_Activate()
    Dim tmpProv As String
    tmpProv = ""
    If Nuevo = False Then
        'cuando viene de pedido y se carga el combo el evento click
        'cambia nproveedor al primero de la lista lo grabo en una variable temp
        tmpProv = nProveedor
        
    End If
    CargarCombo cmbProveedores, "SELECT Proveedor FROM Proveedores " + _
        "ORDER BY Proveedor", "Proveedor"
    
    'hago que se ponga el nombre del proveedor del pedido
    'esto puede joder cuando agregue productos o proveedores pero es lo mejor que se me ocurre
    If tmpProv <> "" Then cmbProveedores = tmpProv
    
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim IDcierre As Long, AjustesStock As Single, RFyT As Single, DifStk As Single
    
    'calculo compras para pagina principal
    AjustesStock = PC.ABSSumarconSubcuentas(35, False)
    RFyT = PC.ABSSumarconSubcuentas(23, False)
    DifStk = PC.UltVarCuenta(54) - AjustesStock - RFyT
    
    CpDia = PC.ABSSumarconSubcuentas(18, False) + DifStk
End Sub

Private Sub lstProductos_Click()
    cmdSel.Default = True
End Sub

Private Sub txtNameP_GotFocus()
    cmdSel.Default = True
End Sub

Private Sub txtNameP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then lstProductos.SetFocus
End Sub

Private Sub cmdSel_Click()
    If lstProductos.ListCount = 0 Then
        MsgBox "No ha eligido ningun producto"
        Exit Sub
    End If
    
    Dim ClsP As New clsProducto 'para poner el ultimo costo ya en el textbox
    Dim Spp() As String
    Spp = Split(lstProductos, "/")
    
    txtCant = CStr(1)
    CtoTmp = ClsP.GetCosto(Spp(1))
    txtPrecio = FormatCurrency(CtoTmp, , , , vbFalse)
     
    txtCant.SetFocus
    txtCant.SelStart = 0
    txtCant.SelLength = Len(txtCant)
    cmdSel.Default = False
    
    Set ClsP = Nothing
        
   
End Sub

Private Sub Command1_Click()
    If lstFactura.ListCount > 0 Then
        If MsgBox("¿Está seguro que desea salir sin registrar la compra o " + _
            "grabar el pedido?." + vbCrLf + " Si presiona ACEPTAR Los datos " + _
            "se perderán definitivamenTE", vbOKCancel + vbInformation, _
            "Atención") = vbCancel Then Exit Sub
    End If
    Unload Me
End Sub

Public Function AbrirDatos(Optional Proveedor As String = "...", _
    Optional IDPedido As Long = -1)
    Dim RSpd As New ADODB.Recordset
    Dim Ss As String, PU As Single, stADD As String
          
    Ss = "SELECT CompraDetalle.Cantidad, Productos.nProducto, " + _
        "CompraDetalle.PrecioTotal,Compradetalle.id, productos.id " + _
        "FROM compradetalle INNER JOIN productos " + _
        "ON compradetalle.idproducto = productos.id WHERE idcompra= " + _
        CStr(IDPedido)

    Nuevo = True 'predeterminado
    ToTal = 0 '???? no se porque pero se guarda el valor de la factura vieja
    
    nProveedor = Proveedor
        
    txtDeMas = FormatCurrency(0)
    txtPrecio = FormatCurrency(0)
    txtDesc = FormatCurrency(0)
    txtDescP = FormatPercent(0)
    lblTotal = FormatCurrency(0)
    
    If IDPedido <> -1 Then 'abre un pedido grabado}
        Nuevo = False
        
        lstFactura.Clear
        
        RSpd.CursorLocation = adUseClient
        'el id del pedido debe ser correcto si o si
        RSpd.Open Ss, DB.CN, adOpenStatic, adLockOptimistic

        If RSpd.RecordCount = 0 Then
            RSpd.Close
            Exit Function
        End If
                
        RSpd.MoveFirst
        Do While Not RSpd.EOF
            PU = RSpd("preciototal") / RSpd("cantidad")
            
            stADD = CStr(RSpd("cantidad")) + "\" + RSpd("nproducto") + _
                    "\" + FormatCurrency(PU, 4, , , vbFalse) + "\" + _
                    FormatCurrency(RSpd("preciototal"), , , , vbFalse) + "\" + _
                    CStr(RSpd("Productos.id"))
            
            lstFactura.AddItem stADD
            
            ToTal = ToTal + RSpd("preciototal")
            lblTotal = FormatCurrency(ToTal, , , , vbFalse)
             'lo borro para que no se duplique informacion usodb.cn porque se me complicaba
             'rs.deleteitem porque el recordset esta abierto con inner join
            
            RSpd.MoveNext
        Loop
        
        RSpd.Close
        Set RSpd = Nothing
        
        'ahora borro todos los registros viejos
       DB.Execute "delete from compradetalle where idcompra=" + CStr(IDPedido)
        
    End If
    Me.Show 1
    
End Function



Private Sub Form_Load()
    txtCant = "1"
    
    Limpiar
    
    txtNameP = ""
    FormatearMouTextBox frmCompras
    
End Sub

Private Sub lblTotal_Change()
    lblTotalT = lblTotal
End Sub

Private Sub OpPesos_Click()
    If OpPesos Then
        DescP = 0
        Desc = 0
        txtDescP = FormatPercent(DescP)
        txtDesc.Enabled = True
        txtDesc = FormatCurrency(Desc, , , , vbFalse)
        txtDescP.Enabled = False
        txtDesc.SetFocus
        txtDesc.SelStart = 0
        txtDesc.SelLength = Len(txtDesc)
    End If
        
End Sub

Private Sub OpPorC_Click()
    If OpPorC Then
        txtDescP.Enabled = True
        Desc = 0
        txtDesc = FormatCurrency(Desc, , , , vbFalse)
        DescP = 0
        txtDescP = FormatPercent(DescP)
        txtDesc.Enabled = False
        txtDescP.SetFocus
        txtDescP.SelStart = 0
        txtDescP.SelLength = Len(txtDescP)
    End If
        
End Sub

Private Sub txtCant_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtPrecio.SetFocus
        txtPrecio.SelStart = 0
        txtPrecio.SelLength = Len(txtPrecio)
    End If

End Sub

Private Sub txtCant_Change()
    If Not IsNumeric(txtPrecio) Or Not IsNumeric(txtCant) Then
        txtPrecio = FormatCurrency(0)
        Exit Sub
    End If
    
    txtPrecio = FormatCurrency(CSng(txtCant) * CtoTmp, , , , vbFalse)
End Sub

Private Sub txtCant_LostFocus()
    Dim caNT As Long 'toma enteros si o si si viene decimal redondea
    caNT = ValidarNumeros(txtCant)
    txtCant = CLng(caNT)

End Sub

Private Sub txtDeMas_LostFocus()
    DeMas = ValidarNumeros(txtDeMas)
    txtDeMas = FormatCurrency(DeMas, , , , vbFalse)
End Sub


Private Sub txtNameP_Change()
    If IsNumeric(txtNameP) Then
        CargarCombo lstProductos, "SELECT nproducto, id FROM Productos " _
            & "WHERE ID LIKE '%" + txtNameP + "%' " + _
            "AND ID>=0", "id,nproducto"
    Else
        CargarCombo lstProductos, "SELECT nproducto,id FROM Productos " _
            & "WHERE nProducto LIKE '%" + txtNameP + "%' " + _
            "AND ID>=0", _
            "id,nproducto"
    End If
        
End Sub

Private Sub txtPrecio_GotFocus()
    txtPrecio.SelStart = 0
    txtPrecio.SelLength = Len(txtPrecio)
    cmdAgregar.Default = True
End Sub

Private Function GetUltIDCmasUno() As Long
    Dim UltIDC As Long, RSUid As New ADODB.Recordset
    
    RSUid.Open "SELECT TOP 1 idcompra from compradetalle ORDER BY idcompra desc", _
        DB.CN, adOpenStatic, adLockReadOnly
    If RSUid.RecordCount > 0 Then
        UltIDC = RSUid("idcompra")
    Else
        UltIDC = 0
    End If
    
    RSUid.Close
    Set RSUid = Nothing
    
    GetUltIDCmasUno = UltIDC + 1
End Function




