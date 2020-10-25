VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{A7FBD38D-2930-49E3-B60C-9E0202D84549}#13.0#0"; "tbrControles.ocx"
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmDevolucion 
   BackColor       =   &H00544B45&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mercaderia a Costo"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9525
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDevolucion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   9525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin tbrFaroButton.fBoton cmdTerminar 
      Height          =   525
      Left            =   2910
      TabIndex        =   26
      Top             =   7950
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   926
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "registrar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdEliminar 
      Height          =   525
      Left            =   1140
      TabIndex        =   25
      Top             =   7950
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   926
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "eliminar registro"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdSel2 
      Height          =   435
      Left            =   3990
      TabIndex        =   24
      Top             =   3990
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   ">>"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdSelProd 
      Height          =   435
      Left            =   4290
      TabIndex        =   23
      Top             =   360
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   ">>"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrControles.MouTextBox txtCant 
      Height          =   405
      Left            =   1350
      TabIndex        =   1
      Top             =   4020
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   714
      Alignment       =   2
      BackColor       =   16777215
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Entero          =   -1  'True
   End
   Begin tbrControles.tbrBuscador tbrBuscadorP 
      Height          =   2445
      Left            =   420
      TabIndex        =   0
      Top             =   810
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4313
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
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox cmbSucursales 
      Height          =   315
      Left            =   6030
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   4920
      Width           =   2415
   End
   Begin VB.ComboBox cmbNombre 
      Height          =   315
      Left            =   6030
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   4020
      Width           =   2415
   End
   Begin VB.TextBox txtPU 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1980
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "88.88"
      Top             =   4020
      Width           =   885
   End
   Begin VB.TextBox txtPT 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "888.88"
      Top             =   4020
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00544B45&
      Caption         =   "A cuenta de"
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   6360
      TabIndex        =   16
      Top             =   1350
      Width           =   1965
      Begin VB.OptionButton chkDevolucion 
         BackColor       =   &H00544B45&
         Caption         =   "Devoluciones"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   210
         TabIndex        =   7
         Top             =   870
         Width           =   1725
      End
      Begin VB.OptionButton chkEmpleado 
         BackColor       =   &H00544B45&
         Caption         =   "Empleados"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   210
         TabIndex        =   6
         Top             =   570
         Width           =   1335
      End
      Begin VB.OptionButton chkSocio 
         BackColor       =   &H00544B45&
         Caption         =   "Socios"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   210
         TabIndex        =   5
         Top             =   270
         Value           =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.TextBox txtSubTotal 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6960
      TabIndex        =   12
      Text            =   "$ 000.00"
      Top             =   5910
      Width           =   1550
   End
   Begin VB.TextBox txtTOTAL 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6960
      TabIndex        =   11
      Text            =   "$ 000.00"
      Top             =   7020
      Width           =   1550
   End
   Begin MSComctlLib.ListView lvTodo 
      Height          =   2445
      Left            =   270
      TabIndex        =   18
      Top             =   5250
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   4313
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
         Text            =   "Id.Prod."
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Cant"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Producto"
         Object.Width           =   3263
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Cto.Unit."
         Object.Width           =   1711
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Cto.Total"
         Object.Width           =   1852
      EndProperty
   End
   Begin tbrControles.MouTextBox txtDesc 
      Height          =   405
      Left            =   6960
      TabIndex        =   10
      Top             =   6480
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   714
      Alignment       =   2
      BackColor       =   16777215
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin tbrFaroButton.fBoton Command1 
      Height          =   525
      Left            =   8010
      TabIndex        =   27
      Top             =   8100
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   926
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.Label lblStockSucu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Buscador por Nombre o ID Producto"
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   630
      TabIndex        =   2
      Top             =   3390
      Width           =   4335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione Sucursal"
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   6090
      TabIndex        =   3
      Top             =   4650
      Width           =   1725
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Buscador por Nombre o ID Producto"
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   810
      TabIndex        =   4
      Top             =   450
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " Cant.       P.U            PT"
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   1320
      TabIndex        =   22
      Top             =   3780
      Width           =   2415
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Factura Cargada"
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
      Height          =   435
      Left            =   1410
      TabIndex        =   21
      Top             =   4680
      Width           =   3105
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione Cuenta"
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   6060
      TabIndex        =   17
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descuento"
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   5820
      TabIndex        =   15
      Top             =   6570
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Subtotal"
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   6000
      TabIndex        =   14
      Top             =   5970
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   6150
      TabIndex        =   13
      Top             =   7140
      Width           =   615
   End
End
Attribute VB_Name = "frmDevolucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Subtotal As Single
Dim ToTal As Single 'Total de la pesos de la factura
Dim Desc As Single
Dim ActuTbrB As Boolean

Private Sub chkDevolucion_Click()
    cmbNombre.Clear
End Sub

Private Sub chkEmpleado_Click()
    If chkEmpleado Then
        Dim Empleados() As String, jj As Long
        
        Empleados = PC.GetCuentas(53)
        'cargo combo pero asi
        cmbNombre.Clear
        For jj = 1 To UBound(Empleados)
            cmbNombre.AddItem PC.GetNameCuenta(CLng(Empleados(jj)))
        Next jj
        If cmbNombre.ListCount > 0 Then cmbNombre.ListIndex = 0
    End If
End Sub

Private Sub chkSocio_Click()
    If chkSocio Then
        Dim Socios() As String, jj As Long
        
        Socios = PC.GetCuentas(52)
        'cargo combo pero asi
        cmbNombre.Clear
        For jj = 1 To UBound(Socios)
            cmbNombre.AddItem PC.GetNameCuenta(CLng(Socios(jj)))
        Next jj
        
        If cmbNombre.ListCount > 0 Then cmbNombre.ListIndex = 0
    End If
End Sub

Private Sub cmbSucursales_Click()
    If tbrBuscadorP.GetLstSel = "" Then
        lblStockSucu = ""
    Else
        lblStockSucu = "Stock en " + UCase(cmbSucursales) + ": " + GetST
    End If
End Sub

Private Sub cmdEliminar_Click()
    Dim Descontar As String
    
    If lvTodo.ListItems.Count = 0 Then Exit Sub
    'si no hay productos en lsttodo bloqueo el descuento
    If lvTodo.ListItems.Count = 1 Then txtDesc.Enabled = False
    
    Descontar = txtInLvW(lvTodo, lvTodo.SelectedItem.Index, 4)
    
    lvTodo.ListItems.Remove lvTodo.SelectedItem.Index
    
    Subtotal = Subtotal - CSng(Descontar)
    txtSubTotal = FormatCurrency(Subtotal, , , , vbFalse)
    
    tbrBuscadorP.Text = "Elegir prod" 'tonto pero solo para que se ejecute el suceso
                                  'change y se actualize el stock
    tbrBuscadorP.SetFocus
    tbrBuscadorP.SelStart = 0
    tbrBuscadorP.SelLength = Len(tbrBuscadorP.Text)
    cmdSelProd.Default = True
End Sub

Private Sub cmdSel2_Click()
    If tbrBuscadorP.GetLstSel = "" Then Exit Sub
    
    If Not IsNumeric(txtCant) Then
        txtCant = "1"
        PintarTxt txtCant
        Exit Sub
        
    Else
        'no permito numeros negativos
        If CSng(txtCant) < 1 Then
            txtCant = "1"
            PintarTxt txtCant
            Exit Sub
        End If
        
        If CSng(txtCant) > GetST Then
            If MsgBox("Va a incluir un producto que tiene stock menor a la " + _
                "cantidad que desea devolver ¿Está seguro del registro? " + _
                "Si selecciona ACEPTAR seguirá adelante pero dejará al producto " + _
                "con STOCK NEGATIVO por lo que es recomendable realizar el ajuste " + _
                "correspondiente, si no presione" + _
                "CANCELAR", vbOKCancel + vbExclamation, "ATENCIÓN") = vbCancel Then
                Exit Sub
            End If
        End If
        
        txtDesc.Enabled = True
        
        Subtotal = Subtotal + CSng(txtPT)
        
        Dim XX As Long
        
        XX = lvTodo.ListItems.Count + 1
        lvTodo.ListItems.Add XX
        
        lvTodo.ListItems(XX).Text = tbrBuscadorP.GetLstSel(0)
        lvTodo.ListItems(XX).SubItems(1) = txtCant
        lvTodo.ListItems(XX).SubItems(2) = GetProd
        lvTodo.ListItems(XX).SubItems(3) = FormatCurrency(txtPU, , , , vbFalse)
        lvTodo.ListItems(XX).SubItems(4) = FormatCurrency(txtPT, , , , vbFalse)
        
        tbrBuscadorP.Text = "Elegir prod" 'tonto pero solo para que se ejecute el suceso
                                      'change y se actualize el stock
        tbrBuscadorP.SetFocus
        tbrBuscadorP.SelStart = 0
        tbrBuscadorP.SelLength = Len(tbrBuscadorP.Text)
        
        cmdSelProd.Default = True
        txtCant = "1"
        txtSubTotal = FormatCurrency(Subtotal, , , , vbFalse)
    
    End If
End Sub

Private Sub cmdSelProd_Click()
    If tbrBuscadorP.GetLstSel = "" Then Exit Sub
    
    'pasarlo a los de texto
    txtPU = FormatCurrency(GetPU, , , , vbFalse)
    txtPT = FormatCurrency(CSng(txtCant) * CSng(txtPU), , , , vbFalse)
    
    PintarTxt txtCant
    
    cmdSel2.Default = True
End Sub

Private Sub cmdTerminar_Click()
    '-------------------    VALIDACIÓN  ----------------------------------------
    
    txtDesc = FormatCurrency(ValidarNumeros(txtDesc), , , , vbFalse)
    'primero que todo veo la habilitación!!!!
    'para socios y empleados!
    If chkDevolucion.Value = False Then 'es porque eligio o socio o empleado
        'el evento es: "A Costo - "+nombre de cuenta
        Dim idEv As Long, UUs As Long
        idEv = ACC.GetID("Evento", "Eventos", "A Costo - " + cmbNombre)
        UUs = ACC.UltUsuarioIngresado
    
        If ACC.ExisteRelacion(UUs, idEv) = 0 Then
            MsgBox ACC.GetNombre("Usuario", "Usuarios", UUs) + _
                " no está habilitado para ingresar." + vbCrLf + _
                "Debe Cambiar Sesión a la de un usuario habilitado", vbExclamation, "Atención"
            Exit Sub
        End If
        
        'registro el movimiento(idev:A costo - NombreCuenta)
        ACC.RegEvento UUs, idEv, "Mercadería a costo " + cmbNombre
        
        End If
    ' --------------------------------------------------------------------------
    'si todo ok sigo
    Dim TmP As String
    
    If lvTodo.ListItems.Count = 0 Then
        MsgBox "¡No realizó registros!", vbInformation, "Atención"
        Exit Sub
    End If
    
    If MsgBox("Está por realizar el registro. Si es correcto presione Aceptar", _
        vbOKCancel + vbInformation, "Atencion") = vbCancel Then Exit Sub
    
    Dim VtaFac As Single
    Dim I As Long
    
    Desc = CSng(txtDesc)
    ToTal = Subtotal - Desc
    txtTOTAL = FormatCurrency(ToTal, , , , vbFalse)
    
    For I = 1 To lvTodo.ListItems.Count
        Dim IDp As String, Cto As Single, Cant As Long
        
        IDp = txtInLvW(lvTodo, I, 0)
        Cto = CSng(DB.GetValInRS("Productos", "pcosto", "ID=" + IDp))
        Cant = CLng(txtInLvW(lvTodo, I, 1))
        
        Dim ClsP As New clsProducto
        ClsP.ModificarStock CLng(IDp), -Cant, cmbSucursales, _
            "Por Devolución de Mercadería o A costo"
        VtaFac = VtaFac + Cant * Cto
        Set ClsP = Nothing
    Next I
        
    'registro aca nomas en el libro diario, si es a credito u otro hago
    'un asiento contrario anulando caja
    Dim Caja As Single
    Dim Cuenta As String 'empleado o socio
    Dim IdCz As Long
    
    IdCz = PC.GetIDCuenta(cmbNombre)
    Caja = VtaFac - CSng(txtDesc)
     
    If chkSocio Then
        Cuenta = "Socio"
    End If
    
    If chkEmpleado Then
        Cuenta = "Empleado"
    End If
    
    If chkDevolucion Then
        'caja/dto a mercaderia, la diferencia por ahora la dejo como ventas
        PC.Asiento "78/81", CStr(Caja) + "/" + txtDesc, "54", CStr(VtaFac), _
            "LibroSubDiario", "Por devolución de Mercadería a proveedores"
    Else
        'es socio o empleado
        DB.EXECUTE "INSERT INTO MovSocyEmp (ID,Fecha, IdNivel3," + _
            "Tipo,Variacion,Detalle) VALUES (" + IdAutonum("MovSocyEmp") + _
            ",#" + stFechaSQL(Date) + _
            "#," + CStr(IdCz) + ",'" + Cuenta + "'," + _
            Replace(CStr(-ToTal), ",", ".") + _
            ",'Mercaderia a costo')"
        
        'asientos       'por ahora la dif va como ventas
        'lo pongo como venta a costo: cuenta a ventas
        PC.Asiento CStr(IdCz) + "/81", CStr(Caja) + "/" + CStr(CSng(txtDesc)), "80", CStr(VtaFac), _
            "LibroSubDiario", "Por Mercadería a costo de " + Cuenta + " " + CStr(IdCz)
        'hago otro de costo de venta a mercaderias
        PC.Asiento "18", CStr(VtaFac), "54", CStr(VtaFac), _
            "LibroSubDiario", "Por Mercadería a costo de " + Cuenta + " " + CStr(IdCz)
    End If
    
    VtaFac = 0 'empiezo de vuelta
    
    Unload Me
End Sub

Private Sub Command1_Click()
    If lvTodo.ListItems.Count > 0 Then
        If MsgBox("¿Está seguro que desea salir?", vbOKCancel, "Salir") = vbCancel Then
            Exit Sub
        End If
    End If
    
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    txtDesc = FormatCurrency(0)
    txtCant = "1"
    
    If UBound(PC.GetCuentas(53)) = 0 Then 'no hay empleados
        chkEmpleado.Enabled = False
    End If
    
    cmbSucursales.Clear
    cmbSucursales.AddItem "CASA CENTRAL"
    CargarCombo cmbSucursales, "SELECT * FROM Sucursales", "Sucursal", , True
    cmbSucursales.ListIndex = 0 'si no hay sucursales no lo hace
    
    tbrBuscadorP.Contrasena = Contrasena
    tbrBuscadorP.ArchivoMDB = ArchivoMDBPrincipal
    tbrBuscadorP.SqlSinLike = "SELECT TOP 50 * FROM Productos WHERE ID>0"
    tbrBuscadorP.CampoEnQueBuscar = "id/n,nproducto/b,pCosto/$,Stock/n"
    tbrBuscadorP.ColumnasSepPorComasyParentesis = "ID(600)/Producto(2150)/" + _
        "Costo(1150)/Stock(800)"
    
    ActuTbrB = False
    
    Subtotal = 0
    txtSubTotal = FormatCurrency(Subtotal, , , , vbFalse)
    
    lblfecha = Date
    txtPU = FormatCurrency(0)
    txtPT = FormatCurrency(0)
    chkSocio_Click
    
    Exit Sub
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        If lvTodo.ListItems.Count = 0 Then Exit Sub
        If MsgBox("¿Está seguro que desea salir?", vbOKCancel, "Salir de Ventas") = vbCancel Then
            Cancel = True
        End If
    End If
    
    tbrBuscadorP.CN_CLOSE
    VtaDia = PC.ABSSumarconSubcuentas(17, False)
    CtoDia = PC.ABSSumarconSubcuentas(18, False)
End Sub

Private Function GetPU() As String
    GetPU = tbrBuscadorP.GetLstSel(2)
    If Not IsNumeric(GetPU) Then GetPU = "0"
End Function

Private Function GetProd() As String
    GetProd = tbrBuscadorP.GetLstSel(1)
End Function

Private Function GetProdTd(Indice As Long) As String
    GetProdTd = txtInLvW(lvTodo, Indice, 2)
End Function

Private Function GetST() As String
    'depende de la sucursal que esté seleccionada
    Dim IDp As Long, Resp As Long, ClsP As New clsProducto, Y As Long
    
    If tbrBuscadorP.GetLstSel = "" Then
        GetST = 0
        Exit Function
    End If
    
    IDp = tbrBuscadorP.GetLstSel(0)
    Resp = ClsP.StockProductoenSucursal(IDp, cmbSucursales)
    
    'ahora le descuento los que ya haya cargado en la factura
    For Y = 1 To lvTodo.ListItems.Count
        If lvTodo.ListItems(Y).Text = CStr(IDp) Then
            Resp = Resp - CLng(lvTodo.ListItems(Y).SubItems(1))
        End If
    Next Y
    
    Set ClsP = Nothing
    
    GetST = CStr(Resp)
End Function

Private Sub tbrBuscadorP_Change()
    If IsNumeric(tbrBuscadorP.Text) Then
        tbrBuscadorP.CampoEnQueBuscar = "id/b,nproducto,pCosto/$,Stock/n"
    Else
        tbrBuscadorP.CampoEnQueBuscar = "id/n,nproducto/b,pCosto/$,Stock/n"
    End If
    VerActu
    tbrBuscadorP_Click
End Sub

Private Sub tbrBuscadorP_Click()
    'lo unico el stock no se actualiza hasta que se haga la venta
    cmdSelProd.Default = True
    txtPU = FormatCurrency(GetPU, , , , vbFalse)
    txtPT = FormatCurrency(CSng(txtCant) * CSng(txtPU), , , , vbFalse)
    
    If tbrBuscadorP.GetLstSel = "" Then
        lblStockSucu = ""
    Else
        lblStockSucu = "Stock en " + UCase(cmbSucursales) + ": " + GetST
    End If
End Sub

Private Sub tbrBuscadorP_GotFocus()
    tbrBuscadorP.SelStart = 0
    tbrBuscadorP.SelLength = Len(tbrBuscadorP.Text)
End Sub

Private Sub txtCant_Change()
    If Not IsNumeric(txtCant) Then
        txtPT = 0
    Else
        txtPT = FormatCurrency(CSng(txtCant) * CSng(txtPU), , , , vbFalse)
    End If
End Sub

Private Sub txtCant_LostFocus()
    If Not IsNumeric(txtCant) Then txtCant = "1"
End Sub

Private Sub txtDesc_Change()
    If Not IsNumeric(txtDesc) Then
        Desc = 0
    Else
        Desc = CSng(txtDesc)
    End If
    ToTal = Subtotal - Desc
    txtTOTAL = FormatCurrency(ToTal, , , , vbFalse)
End Sub

Private Sub txtDesc_GotFocus()
    txtDesc.SelStart = 0
    txtDesc.SelLength = Len(txtDesc)
End Sub

Private Sub txtDesc_Lostfocus()
    Desc = ValidarNumeros(txtDesc)
    txtDesc = FormatCurrency(Desc, , , , vbFalse)
    ToTal = Subtotal - Desc
    txtTOTAL = FormatCurrency(ToTal, , , , vbFalse)
End Sub

Private Sub txtSubTotal_Change()
    ToTal = Subtotal - CSng(txtDesc)
    txtTOTAL = FormatCurrency(ToTal, , , , vbFalse)
End Sub

Private Sub VerActu()
    If ActuTbrB = False Then
       ActuTbrB = True
       tbrBuscadorP.Recargar
    Else
        ActuTbrB = False
    End If
End Sub
