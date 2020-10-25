VERSION 5.00
Object = "{A7FBD38D-2930-49E3-B60C-9E0202D84549}#16.0#0"; "tbrControles.ocx"
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmClientesMov 
   BackColor       =   &H00544B45&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Movimientos Clientes"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmClientesMov.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin tbrFaroButton.fBoton command1 
      Height          =   405
      Left            =   3660
      TabIndex        =   6
      Top             =   6800
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmOK 
      Height          =   405
      Left            =   2190
      TabIndex        =   5
      Top             =   6800
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Aceptar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton command2 
      Height          =   405
      Left            =   4500
      TabIndex        =   19
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Agregar Cliente"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0049453D&
      Caption         =   "Forma de Pago"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   570
      TabIndex        =   15
      Top             =   3870
      Width           =   3855
      Begin VB.OptionButton chkCC 
         BackColor       =   &H0049453D&
         Caption         =   "Cuenta Corriente"
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   210
         TabIndex        =   17
         Top             =   270
         Value           =   -1  'True
         Width           =   2595
      End
      Begin VB.OptionButton chkFin 
         BackColor       =   &H0049453D&
         Caption         =   "A través de Financiera"
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   210
         TabIndex        =   16
         Top             =   570
         Width           =   3075
      End
   End
   Begin tbrControles.MouTextBox txtVariacion 
      Height          =   405
      Left            =   5340
      TabIndex        =   4
      Top             =   4230
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   714
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
   Begin tbrControles.tbrBuscador tbrBuscador1 
      Height          =   2325
      Left            =   540
      TabIndex        =   0
      Top             =   360
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   4101
      BackColor       =   5524293
      BeginProperty Fontt {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00544B45&
      Caption         =   "Condición de pago"
      ForeColor       =   &H00FFFFFF&
      Height          =   945
      Left            =   630
      TabIndex        =   7
      Top             =   5160
      Width           =   6225
      Begin VB.OptionButton chkCuotas 
         BackColor       =   &H00544B45&
         Caption         =   "Configurar Cuotas"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   4170
         MaskColor       =   &H00544B45&
         TabIndex        =   10
         Top             =   360
         Width           =   1875
      End
      Begin VB.OptionButton chkUnico 
         BackColor       =   &H00544B45&
         Caption         =   "Cuota Única para el día"
         ForeColor       =   &H00E0E0E0&
         Height          =   405
         Left            =   210
         MaskColor       =   &H00544B45&
         TabIndex        =   8
         Top             =   300
         Value           =   -1  'True
         Width           =   2325
      End
      Begin MSComCtl2.DTPicker DTVenc1 
         Height          =   345
         Left            =   2610
         TabIndex        =   9
         Top             =   330
         Width           =   1395
         _ExtentX        =   2461
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
         Format          =   130547713
         CurrentDate     =   39196
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00544B45&
      Caption         =   "Monto Deuda"
      ForeColor       =   &H00FFFFFF&
      Height          =   1005
      Left            =   5070
      TabIndex        =   11
      Top             =   2850
      Width           =   1755
      Begin VB.OptionButton chkTodo 
         BackColor       =   &H00544B45&
         Caption         =   "Total Factura"
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
         Left            =   210
         TabIndex        =   1
         Top             =   300
         Value           =   -1  'True
         Width           =   1395
      End
      Begin VB.OptionButton chkParte 
         BackColor       =   &H00544B45&
         Caption         =   "Parte"
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
         Left            =   210
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.TextBox txtDetalle 
      Height          =   630
      Left            =   570
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   3060
      Width           =   4275
   End
   Begin VB.Label lblSelec 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Producto Seleccionado"
      ForeColor       =   &H00E0E0E0&
      Height          =   1515
      Left            =   4590
      TabIndex        =   18
      Top             =   1050
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Detalle"
      ForeColor       =   &H00E0E0E0&
      Height          =   345
      Left            =   600
      TabIndex        =   14
      Top             =   2820
      Width           =   1035
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Importe"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   4530
      TabIndex        =   13
      Top             =   4320
      Width           =   765
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Buscador Cliente"
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   630
      TabIndex        =   12
      Top             =   90
      Width           =   3045
   End
End
Attribute VB_Name = "frmClientesMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IdVta As String 'para deudas relacionadas con ventas en el nro de factura
Dim AnOtaR As Single
Dim TotalFactura As Single 'tb relacionadas con ventas
Dim EsProveedor As Boolean
Dim EsFinanciera As Boolean
Dim AntClienteEleg As String
Dim mNombre As String, IDC As Long

Private Sub chkCC_Click()
    If EsProveedor = False Then
        If chkCC Then
            tbrBuscador1.SqlSinLike = "SELECT id,Nombre FROM Clientes WHERE ID >=0"
            tbrBuscador1.OrderBy = "ORDER BY Nombre"
            tbrBuscador1.CampoEnQueBuscar = "Nombre"
            tbrBuscador1.ColumnasSepPorComasyParentesis = "Cliente(3250)"
            Label2 = "Buscador Cliente"
            Command2.Caption = "Agregar Cliente"
        Else
            tbrBuscador1.SqlSinLike = "SELECT id,Nombre FROM Clientes WHERE ID <-20"
            tbrBuscador1.OrderBy = "ORDER BY Nombre"
            tbrBuscador1.CampoEnQueBuscar = "Nombre"
            tbrBuscador1.ColumnasSepPorComasyParentesis = "Financiera(3250)"
            Command2.Caption = "Agregar Financiera"
            Label2 = "Buscador Financiera"
        End If
    End If
    
    If IdVta <> "NO" Then txtDetalle = "Por Venta " + IdVta
    
    tbrBuscador1.Text = ""
    tbrBuscador1.Recargar
End Sub

Private Sub chkFin_Click()
    If EsProveedor = False Then
        If chkCC Then
            tbrBuscador1.SqlSinLike = "SELECT id,Nombre FROM Clientes WHERE ID >=0"
            tbrBuscador1.OrderBy = "ORDER BY Nombre"
            tbrBuscador1.CampoEnQueBuscar = "Nombre"
            tbrBuscador1.ColumnasSepPorComasyParentesis = "Cliente(3250)"
            Label2 = "Buscador Cliente"
            Command2.Caption = "Agregar Cliente"
        Else
            tbrBuscador1.SqlSinLike = "SELECT id,Nombre FROM Clientes WHERE ID <-20"
            tbrBuscador1.OrderBy = "ORDER BY Nombre"
            tbrBuscador1.CampoEnQueBuscar = "Nombre"
            tbrBuscador1.ColumnasSepPorComasyParentesis = "Financiera(3250)"
            Command2.Caption = "Agregar Financiera"
            Label2 = "Buscador Financiera"
        End If
    End If
    
    If IdVta <> "NO" Then txtDetalle = "Por Venta " + IdVta
    
    If chkFin Then
        If mNombre <> "" And chkCC.Enabled = True Then txtDetalle = txtDetalle + " (" + mNombre + ")"
    End If
    
    tbrBuscador1.Text = ""
    tbrBuscador1.Recargar
End Sub

Private Sub chkParte_Click()
    If chkParte.Value = True Then
        txtVariacion.Enabled = True
    End If
End Sub

Private Sub chkTodo_Click()
    If chkTodo.Value = True Then
        txtVariacion = FormatCurrency(TotalFactura, , , , vbFalse)
        txtVariacion.Enabled = False
    End If
End Sub

Private Sub cmOK_Click()
    Dim IdMov As Long

    If tbrBuscador1.GetLstSel = "" Then
        MsgBox "Debe seleccionar una cuenta"
        Exit Sub
    End If
    
    If Not IsNumeric(txtVariacion) Then
        MsgBox "Debes cargar un número correcto", vbInformation, "Atencion"
        PintarTxt txtVariacion
        Exit Sub
    End If
    
    If CSng(txtVariacion) <= 0 Then
        MsgBox "Sólo puede cargar números positivos", vbInformation, "Atencion"
        PintarTxt txtVariacion
        Exit Sub
    End If
    
    'ya basta de vueltas se los anoto
    If EsProveedor Then 'PROVEEDORES!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        IdMov = IdAutonum("MovProveedores")
        
        DB.EXECUTE "INSERT INTO MovProveedores (ID, Fecha, Proveedor," + _
            "Variacion,Detalle,Documento) VALUES (" + CStr(IdMov) + _
            ",#" + stFechaSQL(Date) + _
            "#,'" + tbrBuscador1.GetLstSel + "'," + _
            Replace(CStr(CSng(txtVariacion)), ",", ".") + _
            ",'" + txtDetalle + "', '" + IdVta + "')"
            
        'registro el asiento Caja a Proveedores
        PC.Asiento "78", txtVariacion, "41", txtVariacion, "LibroSubDiario", _
            "Adeudado a " + tbrBuscador1.GetLstSel
    Else
        'CLIENTES!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        Dim IdCz As Long, IDClOrig As Long, IDcfg As Long, IDFin As Long
        Dim Clscx As New clsCliente
        IdCz = Clscx.GetID(tbrBuscador1.GetLstSel)
        IdMov = IdAutonum("MovClientes")
        
        DB.EXECUTE "INSERT INTO MovClientes (ID, Fecha, CodCliente," + _
            "Variacion,Detalle,Documento) VALUES (" + CStr(IdMov) + _
            ",#" + stFechaSQL(Date) + _
            "#," + CStr(IdCz) + "," + _
            Replace(CStr(CSng(txtVariacion)), ",", ".") + _
            ",'" + txtDetalle + "', '" + IdVta + "')"
       
        'no viene de nada asi que seria Clientes a caja
        'salvo si es la carga inicial va a capital
        
        PC.Asiento "46", txtVariacion, "78", txtVariacion, _
            "LibroSubdiario", "Anotado a Cliente " + CStr(IdCz)
            
        If chkFin Then 'GRABO CONFIGURACION DE QUE CLIENTE QUE ORIGINO LA FACTURA
                    'PAGO POR FINANCIERA ------------------------------------------
            If mNombre <> "" Then
                IDClOrig = Clscx.GetID(mNombre)
                '(1ro) grabo conf 70 ultimo cliente que pago por financiera
                CFG.ModificarNodo 70, , , , CStr(IDClOrig)
                
                '(2do) grabo conf particular con la fecha que lo hizo
                IDcfg = CFG.ExistePropiedad("UPF " + CStr(IDClOrig))
                IDFin = Clscx.GetID(tbrBuscador1.GetLstSel)
                
                If IDcfg = 0 Then 'agrego
                    CFG.AgregarNodo 70, "UPF " + CStr(IDClOrig), CStr(IDFin), _
                        CStr(Date), 0
                Else 'modifico
                    CFG.ModificarNodo IDcfg, 70, , CStr(IDFin), CStr(Date)
                End If
            End If
        End If '------------------------------------------------------------------
         
        Set Clscx = Nothing
    End If
    
    'aunque este todo anotado voy a ¡CUOTAS! si pidio configurar cuotas
    'si no igual le anoto el vencimiento
    If EsProveedor Then
        DB.EXECUTE "INSERT INTO VencimientoProveedor (IdMov, Cuota, Total, Vencimiento) " + _
            "VALUES (" + CStr(IdMov) + "," + _
            Replace(CStr(CSng(txtVariacion)), ",", ".") + ", " + _
            Replace(CStr(CSng(txtVariacion)), ",", ".") + ", #" + _
            stFechaSQL(DTVenc1) + "#)"
    Else
        DB.EXECUTE "INSERT INTO Vencimientos (IdMov, Cuota, Total, Vencimiento) " + _
            "VALUES (" + CStr(IdMov) + "," + _
            Replace(CStr(CSng(txtVariacion)), ",", ".") + ", " + _
            Replace(CStr(CSng(txtVariacion)), ",", ".") + ", #" + _
            stFechaSQL(DTVenc1) + "#)"
    End If
    
    If chkCuotas Then
        If EsProveedor Then
            frmCuotas.AbrirDatos tbrBuscador1.GetLstSel, IdMov, False
        Else
            frmCuotas.AbrirDatos CStr(IdCz), IdMov
        End If
    End If

    Unload Me
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    If EsProveedor Then
        frmProveedores.Show 1
    Else
        If chkCC Then 'es cliente
            frmClientes.AbrirDatos -1
        Else 'es financiera
            frmClientes.AbrirDatos -20
        End If
    End If
End Sub

Public Sub AbrirDatos(Optional Documento As String = "NO", Optional IsProveedor _
    As Boolean = False, Optional Nombre As String = "")
    'si cliente termina en * es financiera
       
    mNombre = Nombre
    IdVta = Documento
    EsProveedor = IsProveedor
    
    If IdVta = "NO" Then
        chkParte = True
        Frame1.Enabled = False
        chkTodo.Enabled = False
        chkParte.Enabled = False
        TotalFactura = 0 'reinicio
    Else
        If EsProveedor Then
            TotalFactura = DB.GetValInRS("FacturaCompra", "Pagado", _
                "NroFactura= '" + IdVta + "' AND EsPedido <> 1", False)
            txtDetalle = "Por Compra " + IdVta
        Else
            TotalFactura = DB.GetValInRS("Facturas", "Pagado", _
                "NroFactura= '" + IdVta + "'", False)
        End If
        
        txtVariacion = FormatCurrency(TotalFactura, , , , vbFalse)
        txtVariacion.Enabled = False
    End If
    
'    tbrBuscador1.SelStart = 0
'    tbrBuscador1.SelLength = Len(tbrBuscador1.Text)
    
    Me.Show 1
End Sub

Private Sub VerConfCliente()
    Dim SPcc() As String, Vcc As String, IdCF As Long, TmP As Long, TPp As String

    Vcc = CFG.GetInfo(40, 4)
    If Vcc = "FIN" Then
        chkFin.Value = True
        tbrBuscador1.Text = ""
    Else
        chkCC.Value = True
    End If
End Sub

Private Sub Form_Activate()
   tbrBuscador1.Recargar
   VerCl
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmOK_Click
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Dim IdQ As String

    tbrBuscador1.Contrasena = Contrasena
    tbrBuscador1.ArchivoMDB = ArchivoMDBPrincipal
    
    txtVariacion = FormatCurrency(TotalFactura)
    Dim CFG24 As String
    CFG24 = CFG.GetInfo(2, 4)
    If IsNumeric(CFG24) Then
        DTVenc1.Value = Date + CLng(CFG24)
    Else
        Terr.AppendLog "NOFECHA-1231"
    End If
       
    If Right(mNombre, 1) = "*" Then
        EsFinanciera = True
        mNombre = Left(mNombre, Len(mNombre) - 1)
        chkCC.Enabled = False
        chkFin.Value = True
        IdQ = "< -20"
    Else
        EsFinanciera = False
        IdQ = ">=0"
    End If
    
    If EsProveedor Then
        tbrBuscador1.SqlSinLike = "SELECT Proveedor FROM Proveedores"
        tbrBuscador1.OrderBy = "ORDER BY Proveedor"
        tbrBuscador1.CampoEnQueBuscar = "Proveedor"
        Label2 = "Buscador Proveedor"
        Command2.Caption = "Agregar Proveedor"
        frmClientesMov.Caption = "Movimientos Proveedores"
        tbrBuscador1.ColumnasSepPorComasyParentesis = "Proveedor(3250)"
        Frame3.Visible = False
    Else
        If EsFinanciera = False Then VerConfCliente
        tbrBuscador1.SqlSinLike = "SELECT id,Nombre FROM Clientes WHERE ID " + IdQ
        tbrBuscador1.OrderBy = "ORDER BY Nombre"
        tbrBuscador1.CampoEnQueBuscar = "Nombre"
        tbrBuscador1.ColumnasSepPorComasyParentesis = "Cliente(3250)"
    End If
    
    tbrBuscador1.Text = mNombre
    tbrBuscador1.SelStart = 0
    tbrBuscador1.SelLength = Len(tbrBuscador1.Text)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    tbrBuscador1.CN_CLOSE
    IdVta = "NO" 'si no queda grabado en algun lado
End Sub

Private Sub VerCl()
    If tbrBuscador1.GetLstSel = "" Then
        lblSelec = ""
    Else
        If EsProveedor Then
            lblSelec = "Proveedor seleccionado: "
        Else
            lblSelec = "Cliente seleccionado: "
        End If
        lblSelec = lblSelec + UCase(tbrBuscador1.GetLstSel)
    End If
End Sub

Private Sub tbrBuscador1_Change()
    VerCl
End Sub

Private Sub tbrBuscador1_Click()
    VerCl
End Sub

Private Sub txtVariacion_GotFocus()
    PintarTxt txtVariacion
End Sub

Private Sub txtVariacion_LostFocus()
    AnOtaR = ValidarNumeros(txtVariacion)
    txtVariacion = FormatCurrency(AnOtaR, , , , vbFalse)
End Sub
