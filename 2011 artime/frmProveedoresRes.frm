VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProveedoresRes 
   BackColor       =   &H00544B45&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resumen Proveedores"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9285
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProveedoresRes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin tbrFaroButton.fBoton cmdPagarF 
      Height          =   495
      Left            =   2520
      TabIndex        =   13
      Top             =   7800
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Pagar Factura"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdPTodo 
      Height          =   495
      Left            =   5400
      TabIndex        =   14
      Top             =   4320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Pagar todo"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdAcomodar 
      Height          =   495
      Left            =   3720
      TabIndex        =   15
      Top             =   4320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Pasar a 1 reg"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdPParcial 
      Height          =   495
      Left            =   2040
      TabIndex        =   16
      Top             =   4320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Pago Parcial"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.ComboBox cmbProveedores 
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
      Left            =   2490
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   270
      Width           =   2505
   End
   Begin MSComctlLib.ListView lvDeudas 
      Height          =   2325
      Left            =   240
      TabIndex        =   10
      Top             =   1770
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   4101
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Fecha"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Monto"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Detalle"
         Object.Width           =   5027
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "IdMov"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Nro.Fac"
         Object.Width           =   3175
      EndProperty
   End
   Begin MSComctlLib.ListView lvFactura 
      Height          =   2055
      Left            =   810
      TabIndex        =   11
      Top             =   5640
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   3625
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
         Text            =   "Pr.Total"
         Object.Width           =   2028
      EndProperty
   End
   Begin tbrFaroButton.fBoton command1 
      Height          =   495
      Left            =   7200
      TabIndex        =   12
      Top             =   7800
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione Proveedor"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   570
      TabIndex        =   9
      Top             =   330
      Width           =   2055
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Aaaaaaaaaaaaa"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   1410
      TabIndex        =   8
      Top             =   1290
      Width           =   4965
   End
   Begin VB.Label lblPesos 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "$ 888,88"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6420
      TabIndex        =   7
      Top             =   1230
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Deudas con:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   210
      TabIndex        =   6
      Top             =   1350
      Width           =   1125
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Factura adeudada"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   1890
      TabIndex        =   5
      Top             =   5130
      Width           =   2625
   End
   Begin VB.Label lblPesosF 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5730
      TabIndex        =   4
      Top             =   6030
      Width           =   1485
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo a Pagar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   5400
      TabIndex        =   3
      Top             =   5700
      Width           =   1695
   End
   Begin VB.Label lblParte 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5820
      TabIndex        =   2
      Top             =   7140
      Width           =   1275
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "% Adeudado del Total Factura"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   555
      Left            =   5760
      TabIndex        =   1
      Top             =   6570
      Width           =   1485
   End
End
Attribute VB_Name = "frmProveedoresRes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Nombre As String

Private Sub ActualizaR()
    If Len(Nombre) > 22 Then
        lblNombre = Left(Nombre, 23) + "...."
    Else
        lblNombre = Nombre
    End If
    
    'ahora con movim de $$ solo esto
    Dim W As String

    W = "SELECT * FROM MovProveedores WHERE Proveedor= '" + Nombre + _
    "' order by fecha desc,id desc"
    
    CargarComboLV lvDeudas, W, "Fecha/f,Variacion/$,Detalle,id/n,documento"

    lblPesos = FormatCurrency(DB.SumarValInRS("MovProveedores", "Variacion", _
        "Proveedor = '" + Nombre + "'"), , , , vbFalse)
    
    If lvDeudas.ListItems.Count = 0 Then
        lvFactura.ListItems.Clear
        lblParte = ""
        lblPesosF = ""
        Exit Sub
    End If
    
End Sub

Private Sub cmbProveedores_Click()
    Nombre = cmbProveedores
    ActualizaR
    lvDeudas_Click
End Sub


Private Sub cmdAcomodar_Click()
    'veo si el usuario que esta trabajando tiene habilitacion para entrar
    Dim UUs As Long
    UUs = ACC.UltUsuarioIngresado
    
    If ACC.ExisteRelacion(UUs, 21) = 0 Then
        MsgBox ACC.GetNombre("Usuario", "Usuarios", UUs) + " no está habilitado " + _
            "para ingresar." + vbCrLf + _
            "Debe Cambiar Sesión a la de un usuario habilitado", vbExclamation, "Atención"
        Exit Sub
    End If
    
    Dim MsJe As String
    
    If MsgBox("Atención, ejecutando esta opción borrará todos los registros " + _
        "de: " + UCase(Nombre) + " Dejando sólo un registro con el saldo " + lblPesos.Caption + _
        " Si esta seguro, presione Aceptar, en caso contrario Cancelar. Es recomendable " + _
        " utilizar esta opción solo con el consentimiento del proveedor", _
        vbInformation + vbOKCancel, "Atención") = vbCancel Then Exit Sub
        
    MsJe = InputBox("Escriba aqui una aclaración del movimiento para que " + _
            "quede en el registro", "Aclaración")
    
    'registro el movimiento(21:Pagar a Proveedores)
    ACC.RegEvento UUs, 21, "Resumir movimientos de proveedor " + cmbProveedores
    
    'borro todo
    DB.EXECUTE "DELETE FROM MovProveedores WHERE Proveedor= '" + Nombre + "'"
    
    'dejo 1 solo registro
    DB.EXECUTE "INSERT INTO MovProveedores (ID,Fecha,Proveedor,variacion," + _
        "Detalle) VALUES (" + IdAutonum("MovProveedores") + ",#" + stFechaSQL(Date) + "#,'" + _
        Nombre + "'," + Replace(CStr(CSng(lblPesos)), ",", ".") + _
        ",'SALDO: " + lblPesos + " " + MsJe + "')"

    ActualizaR
        
End Sub

Private Sub cmdPagarF_Click()
    If lvDeudas.ListItems.Count = 0 Then Exit Sub
        
    Dim NFac As String, FechaD As String, IdMov As Long, APagar As Single, Stp As String
    
    NFac = txtInLvW(lvDeudas, lvDeudas.SelectedItem.Index, 4)
    FechaD = txtInLvW(lvDeudas, lvDeudas.SelectedItem.Index, 0)
    IdMov = CLng(txtInLvW(lvDeudas, lvDeudas.SelectedItem.Index, 3))
    APagar = CSng(txtInLvW(lvDeudas, lvDeudas.SelectedItem.Index, 1))
    
    If EsCero(APagar) = True Then
        MsgBox "No hay nada que pagar", vbInformation, "Atención"
        Exit Sub
    End If
        
    If MsgBox("Está a punto de registrar el pago de " + FormatCurrency(APagar) + _
        ". ¿Los datos son correctos?", vbInformation + vbOKCancel, "Atención") _
        = vbCancel Then
        Exit Sub
    End If
    
    Stp = "Factura N°" + NFac
    'Si es INTERÉS!! tengo que borrarlo de los vencimientos
    If Left(NFac, 4) = "INT." Then
        Dim IdMovi As Long
        IdMovi = CLng(Right(NFac, Len(NFac) - InStrRev(NFac, " ")))
        Stp = NFac
        DB.EXECUTE "UPDATE VencimientoProveedor SET Interes = 0 WHERE " + _
            "IdMov = " + CStr(IdMovi)
    End If
    
    'voy a entrar a ese registro y a cambiar los datos
    Dim rsO As New ADODB.Recordset
    rsO.Open "SELECT * FROM MovProveedores WHERE id =" + CStr(IdMov), DB.CN, adOpenStatic, adLockOptimistic
    
    rsO("Fecha") = Date
    rsO("Variacion") = 0
    
    Dim TmP As String
    If NFac = "NO" Then
        TmP = "Pago " + CStr(APagar)
    Else
        TmP = "Pago " + Stp + " (" + lblPesosF + ")"
    End If
    
    rsO("Detalle") = TmP + " Adeudados desde " + FechaD
    rsO("Documento") = "NO"
    
    rsO.Update
    rsO.Close
    Set rsO = Nothing
    
    'borro los vencimientos
    DB.EXECUTE "DELETE FROM VencimientoProveedor WHERE IdMov = " + CStr(IdMov)
    
    MsgBox "Se registró correctamente " + TmP + " Adeudados desde " + FechaD, _
        vbInformation, "Registro"
    
    'registro en el diario (caja a clientes)
    PC.Asiento "41", lblPesosF, "78", lblPesosF, "LibroSubDiario", _
        "Pagado a Proveedor " + Nombre
    
    If CFG.GetInfo(95, 4) = "Si" Then frmPago.AbrirDatos CSng(lblPesosF), False, "Pagado a Proveedor " + Nombre
    
    ActualizaR
    lvDeudas_Click
End Sub

Private Sub cmdPParcial_Click()
    'veo si el usuario que esta trabajando tiene habilitacion para entrar
    Dim UUs As Long
    UUs = ACC.UltUsuarioIngresado
    
    If ACC.ExisteRelacion(UUs, 21) = 0 Then
        MsgBox ACC.GetNombre("Usuario", "Usuarios", UUs) + " no está habilitado " + _
            "para ingresar." + vbCrLf + _
            "Debe Cambiar Sesión a la de un usuario habilitado", vbExclamation, "Atención"
        Exit Sub
    End If
    
    Dim Desc As String 'monto que se descuenta de la cuenta
    Dim Dt As String   'detalle que se va a cargar en la tabla
    
    Desc = InputBox("Ingrese el Monto", "Pago a: " + Nombre)
    Desc = Replace(Desc, ".", ",")
    
    If Desc = "" Then Exit Sub
    If Not IsNumeric(Desc) Then
        MsgBox "¡Debes cargar un número correcto!", vbExclamation, "Atención"
        Exit Sub
    End If
    
    If CSng(Desc) <= 0 Then MsgBox "No puede ingresar valores no positivos": Exit Sub
    
    Dt = InputBox("Ingrese el detalle del pago", "Detalle pago a: " + Nombre)
    
    DB.EXECUTE "INSERT INTO MovProveedores (ID,Fecha,Proveedor," + _
        "Variacion,Detalle,Documento) VALUES (" + IdAutonum("MovProveedores") + _
        ",#" + stFechaSQL(Date) + _
        "#,'" + Nombre + "'," + Replace(CStr(-CSng(Desc)), ",", ".") + ",'" + _
        "(PP) " + Dt + "','NO')"
    
    'registro el movimiento(21:Pagar a Proveedores)
    ACC.RegEvento UUs, 21, "Pago parcial a proveedor " + cmbProveedores
    
    'registro en el diario (Proveedores a Caja)
    PC.Asiento "41", Desc, "78", Desc, "LibroSubDiario", _
        "Pagado a Proveedor " + Nombre
        
    If CFG.GetInfo(95, 4) = "Si" Then frmPago.AbrirDatos CSng(Desc), False, "Pagado a Proveedor " + Nombre
    
    ActualizaR 'actualizo
End Sub

Private Sub cmdPTodo_Click()
    If EsCero(CSng(lblPesos)) = True Then MsgBox "No hay nada que pagar!": Exit Sub
    
    'veo si el usuario que esta trabajando tiene habilitacion para entrar
    Dim UUs As Long
    UUs = ACC.UltUsuarioIngresado
    
    If ACC.ExisteRelacion(UUs, 21) = 0 Then
        MsgBox ACC.GetNombre("Usuario", "Usuarios", UUs) + " no está habilitado " + _
            "para ingresar." + vbCrLf + _
            "Debe Cambiar Sesión a la de un usuario habilitado", vbExclamation, "Atención"
        Exit Sub
    End If
    
    Dim MsJ As String
    
    If MsgBox("Está a punto de borrar el total de la deuda de: " + _
        UCase(Nombre) + vbCrLf + "Presione Aceptar Sólo si ya fue cobrado " + _
        "en efectivo " + FormatCurrency(CSng(lblPesos)) + vbCrLf + _
        "Si el pago es parcial presione Cancelar", vbInformation + vbOKCancel, _
        "Borrar registros") = vbCancel Then Exit Sub 'si esta seguro y puso la clave listo borro todo $$
        
    MsJ = InputBox("Escriba aqui una aclaración del movimiento para que " + _
        "quede en el registro", "Aclaración")
        
    'borro todo
    DB.EXECUTE "DELETE FROM MovProveedores WHERE Proveedor= '" + Nombre + "'"
    
    'registro el movimiento(21:Pagar proveedores)
    ACC.RegEvento UUs, 21, "Pago total a proveedores " + cmbProveedores
    
    'dejo 1 solo registro
    DB.EXECUTE "INSERT INTO MovProveedores (ID,Fecha,Proveedor,variacion," + _
        "Detalle,Documento) VALUES (" + IdAutonum("MovProveedores") + _
        ",#" + stFechaSQL(Date) + "#,'" + _
        Nombre + "',0,'PAGADO TODO! ( " + lblPesos + ") " + MsJ + "','NO')"
        
    'registro en el diario (Proveedores a Caja)
    PC.Asiento "41", lblPesos, "78", lblPesos, "LibroSubDiario", _
        "Pagado a Proveedor " + Nombre
    
    If CFG.GetInfo(95, 4) = "Si" Then frmPago.AbrirDatos CSng(lblPesos), False, "Pagado a Proveedor " + Nombre
    
    'actualizo
    ActualizaR
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    CargarCombo cmbProveedores, "SELECT Proveedor FROM Proveedores", "Proveedor"
    Nombre = cmbProveedores
    
    ActualizaR
End Sub

Private Sub lvDeudas_Click()
    If lvDeudas.ListItems.Count = 0 Then
        lvFactura.ListItems.Clear
        lblParte = ""
        lblPesosF = ""
        Exit Sub
    End If
    
    'voy a hacer que se carguen las facturas si las tiene
    Dim NroFac As String
    NroFac = txtInLvW(lvDeudas, lvDeudas.SelectedItem.Index, 4)
    
    'si es cero no fue una deuda de venta
    If NroFac = "NO" Or Left(NroFac, 4) = "INT." Then
        lvFactura.ListItems.Clear
        lblParte = ""
        lblPesosF = txtInLvW(lvDeudas, lvDeudas.SelectedItem.Index, 1)
        Exit Sub
    End If
    
    'ahora si hago que se cargue
    Dim SS As String
    SS = "SELECT Productos.nProducto, CompraDetalle.Cantidad, CompraDetalle.PrecioTotal " + _
        "FROM Productos INNER JOIN CompraDetalle ON Productos.ID = " + _
        "CompraDetalle.IDproducto WHERE CompraDetalle.NroFactura = '" + _
        CStr(NroFac) + "' " + _
        "GROUP BY Productos.nProducto, CompraDetalle.Cantidad, CompraDetalle.PrecioTotal"

    CargarComboLV lvFactura, SS, "Cantidad/n,nProducto,PrecioTotal/$"
    
    lblPesosF = txtInLvW(lvDeudas, lvDeudas.SelectedItem.Index, 1)
    
    'ahora veo cuanto era el total de la factura para sacar el porcentaje
    Dim TTF As Single
    TTF = DB.GetValInRS("FacturaCompra", "Pagado", _
        "NroFactura = '" + NroFac + "'", False)
    
    If TTF = 0 Then
        lblParte = FormatPercent(1)
    Else
        lblParte = FormatPercent(CSng(lblPesosF) / TTF)
    End If
End Sub

Private Sub lvDeudas_KeyUp(KeyCode As Integer, Shift As Integer)
    lvDeudas_Click
End Sub
