VERSION 5.00
Object = "{A7FBD38D-2930-49E3-B60C-9E0202D84549}#15.0#0"; "tbrControles.ocx"
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmVencimientos 
   BackColor       =   &H00544B45&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vencimientos"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11325
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVencimientos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   11325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin tbrFaroButton.fBoton cmdPagar 
      Height          =   420
      Left            =   3585
      TabIndex        =   5
      Top             =   6465
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   741
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Pagar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin MSDataGridLib.DataGrid DGVenc 
      Height          =   4275
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   7541
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      BackColor       =   16777215
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   1
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   2
            Format          =   "0,000E+00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin tbrControles.MouTextBox txtIntDesc 
      Height          =   435
      Left            =   5550
      TabIndex        =   1
      Top             =   6450
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   767
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
   Begin MSComCtl2.DTPicker DTFecha 
      Height          =   345
      Left            =   6000
      TabIndex        =   0
      Top             =   1350
      Width           =   1410
      _ExtentX        =   2487
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
      Format          =   20971521
      CurrentDate     =   39197
   End
   Begin tbrFaroButton.fBoton command1 
      Height          =   420
      Left            =   9360
      TabIndex        =   6
      Top             =   6480
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   741
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdDescInt 
      Height          =   420
      Left            =   6915
      TabIndex        =   7
      Top             =   6465
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   741
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Descontar Interés"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mostrar Deudas con Vencimientos anteriores al día"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1110
      TabIndex        =   3
      Top             =   1410
      Width           =   4965
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimientos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2250
      TabIndex        =   2
      Top             =   420
      Width           =   4125
   End
End
Attribute VB_Name = "frmVencimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSVenc As New ADODB.Recordset
Dim IsProveedor As Boolean

Private Sub cmdDescInt_Click()
    Dim AdesC As Single, IdMovi As Long, IdMovCli As Long
    Dim CuantoEra As Single
        
    txtIntDesc = FormatCurrency(ValidarNumeros(txtIntDesc), , , , vbFalse)
    If EsCero(CSng(txtIntDesc)) = True Then Exit Sub
    AdesC = CSng(DGVenc.Columns("Interes"))
    If EsCero(AdesC) = True Then Exit Sub
    
    If CSng(txtIntDesc) > AdesC Then
        MsgBox "No puede descontar más intereses de lo pactado originalmente", vbExclamation, "Atención"
        txtIntDesc = FormatCurrency(AdesC, , , , vbFalse)
        PintarTxt txtIntDesc
        Exit Sub
    End If
    
    IdMovi = CLng(DGVenc.Columns("ID"))
    IdMovCli = CLng(DGVenc.Columns("ID.Mov."))
    'descuento el interes
    If IsProveedor Then
        '(1) Descuento en vencimientos
        DB.EXECUTE "UPDATE VencimientoProveedor SET Interes = " + _
            Replace(CStr(AdesC - CSng(txtIntDesc)), ",", ".") + _
            " WHERE ID = " + CStr(IdMovi)
        '(2) Asiento
        'hago el asiento descontando los intereses (tengo una ganancia)
        'proveedores (41) a intereses perdidos (69)
        PC.Asiento "41", txtIntDesc, "69", txtIntDesc, "LibroSubDiario", _
            "Intereses Perdonados por el Proveedor " + DGVenc.Columns("Proveedor")
        '(3) Descuento en MovClientes
        CuantoEra = DB.GetValInRS("MovProveedores", "Variacion", "ID = " + CStr(IdMovCli + 1), False)
        'siempre el id del interes en movclientes es uno mas del mov que lo genero
        DB.EXECUTE "UPDATE MovProveedores SET Variacion = " + _
            Replace(CStr(CuantoEra - CSng(txtIntDesc)), ",", ".") + _
            " WHERE ID = " + CStr(IdMovCli + 1)
    Else
        '(1) Descuento en vencimientos
        DB.EXECUTE "UPDATE Vencimientos SET Interes = " + _
            Replace(CStr(AdesC - CSng(txtIntDesc)), ",", ".") + _
            " WHERE ID = " + CStr(IdMovi)
        '(2) Asiento
        'hago el asiento descontando los intereses (tengo una perdida)
        'intereses ganados (68) a Clientes (46)
        PC.Asiento "68", txtIntDesc, "46", txtIntDesc, "LibroSubDiario", _
            "Intereses Perdonados por al Cliente Nro. " + DGVenc.Columns("IdCli")
        '(3) Descuento en MovClientes
        CuantoEra = DB.GetValInRS("MovClientes", "Variacion", "ID = " + CStr(IdMovCli + 1), False)
        'siempre el id del interes en movclientes es uno mas del mov que lo genero
        DB.EXECUTE "UPDATE MovClientes SET Variacion = " + _
            Replace(CStr(CuantoEra - CSng(txtIntDesc)), ",", ".") + _
            " WHERE ID = " + CStr(IdMovCli + 1)
    End If
    
    '(4) Aviso al usuario
    MsgBox "Se registró correctamente el descuento de " + txtIntDesc + " de interés", vbInformation, "Registro"
    txtIntDesc = FormatCurrency(0)
    RefreshDG
End Sub

Private Sub cmdPagar_Click()
    If RSVenc.RecordCount = 0 Then Exit Sub
    
    Dim NroAc As Long 'nro de acceso
    If IsProveedor = False Then
        NroAc = 20 'cobrar clientes
    Else
        NroAc = 21 'pagar proveedores
    End If
        
    'veo si el usuario que esta trabajando tiene habilitacion para entrar
    Dim UUs As Long
    UUs = ACC.UltUsuarioIngresado
    
    If ACC.ExisteRelacion(UUs, NroAc) = 0 Then
        MsgBox ACC.GetNombre("Usuario", "Usuarios", UUs) + " no está habilitado " + _
            "para ingresar." + vbCrLf + _
            "Debe Cambiar Sesión a la de un usuario habilitado", vbExclamation, "Atención"
        Exit Sub
    End If
    
    'registro el movimiento(20:Cobrar a Clientes) o (21:Pagar Proveedores)
    ACC.RegEvento UUs, NroAc, "Pago cuota"
    
    Dim NFac As String, FechaD As String, IdMov As Long, Vari As Single
    Dim IdPers As String, IDVV As Long
    
    NFac = DGVenc.Columns("Documento")
    FechaD = DGVenc.Columns("Vencim.")
    Vari = DGVenc.Columns("Total")
    IdPers = DGVenc.Columns(0)
    IDVV = DGVenc.Columns("ID") 'El ID en vencimientos
    
    If IsProveedor = False Then
        If MsgBox("Está a punto de registrar el pago de " + FormatCurrency(Vari) + vbCrLf + _
            "de " + UCase(DGVenc.Columns(1)) + " ¿Los datos son correctos?", _
            vbInformation + vbOKCancel, "Atención") = vbCancel Then
            Exit Sub
        End If
        
        DB.EXECUTE "INSERT INTO MovClientes (ID,Fecha,CodCliente," + _
            "Variacion,Detalle,Documento) VALUES (" + IdAutonum("MovClientes") + _
            ",#" + stFechaSQL(Date) + _
            "#," + IdPers + "," + _
            Replace(CStr(-Vari), ",", ".") + ",'" + _
            "(PC) Pago Para Factura N°" + CStr(NFac) + "','" + NFac + "')"
        'registro en el diario (caja a clientes)
        PC.Asiento "78", CStr(Vari), "46", CStr(Vari), "LibroSubDiario", _
            "Cobrado a Cliente Nro. " + IdPers
        DB.EXECUTE "DELETE FROM Vencimientos WHERE ID = " + CStr(IDVV)
    Else
        If MsgBox("Está a punto de registrar el pago de " + FormatCurrency(Vari) + vbCrLf + _
            "a " + UCase(IdPers) + " ¿Los datos son correctos?", _
            vbInformation + vbOKCancel, "Atención") = vbCancel Then
            Exit Sub
        End If
        
        DB.EXECUTE "INSERT INTO MovProveedores (ID,Fecha,Proveedor," + _
            "Variacion,Detalle,Documento) VALUES (" + IdAutonum("MovProveedores") + _
            ",#" + stFechaSQL(Date) + _
            "#,'" + IdPers + "'," + _
            Replace(CStr(-Vari), ",", ".") + ",'" + _
            "(PC) Pago Para Factura N°" + CStr(NFac) + "','" + NFac + "')"
        'registro en el diario (proveedores a caja)
        PC.Asiento "41", CStr(Vari), "78", CStr(Vari), "LibroSubDiario", _
            "Pagado a Proveedor " + IdPers
        DB.EXECUTE "DELETE FROM VencimientoProveedor WHERE ID = " + CStr(IDVV)
    End If
    
    MsgBox "Se realizó el registro correctamente ", vbInformation, "Registro"
    
    If CFG.GetInfo(95, 4) = "Si" Then frmPago.AbrirDatos CSng(lblPesosF), -IsProveedor, "Pago plan de Cuotas " + IdPers
    RefreshDG
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Public Sub AbrirDatos(EsProveedor As Boolean)
    IsProveedor = EsProveedor
    RSVenc.CursorLocation = adUseClient
    
    RefreshDG
    
    If IsProveedor = False Then
        Me.BackColor = &H968D63
        Label1.ForeColor = vbBlack
        Me.Width = 12300
        DGVenc.Width = 11300
    End If
    Me.Show 1
End Sub

Private Sub RefreshDG()
    Set DGVenc.DataSource = Nothing
    If RSVenc.State = adStateOpen Then RSVenc.Close
    
    If IsProveedor = False Then
        'dejo que esten las financieras tambien
        RSVenc.Open "SELECT Clientes.ID, Clientes.Nombre, Vencimientos.Vencimiento, " + _
            "MovClientes.Fecha, DateDiff('d',[Vencimientos]![Vencimiento],Now()) AS " + _
            "Demora, Vencimientos.Cuota, Vencimientos.Interes, Vencimientos.Total, " + _
            "MovClientes.Documento, Vencimientos.IdMov, Vencimientos.ID " + _
            "FROM (Clientes INNER JOIN MovClientes ON Clientes.ID = " + _
            "MovClientes.CodCliente) INNER JOIN Vencimientos ON MovClientes.ID = " + _
            "Vencimientos.IdMov WHERE Vencimientos.Vencimiento < #" + _
            stFechaSQL(DTFecha) + "# " + _
            "ORDER BY Vencimientos.Vencimiento ASC", DB.CN, adOpenStatic, adLockReadOnly
    Else
        RSVenc.Open "SELECT Proveedores.Proveedor, VencimientoProveedor.Vencimiento, " + _
            "MovProveedores.Fecha, DateDiff('d',[VencimientoProveedor]![Vencimiento]," + _
            "Now()) AS Demora, VencimientoProveedor.Cuota, " + _
            "VencimientoProveedor.Interes, VencimientoProveedor.Total, " + _
            "MovProveedores.Documento, VencimientoProveedor.ID, VencimientoProveedor.IdMov " + _
            "FROM (Proveedores INNER JOIN MovProveedores ON Proveedores.Proveedor = " + _
            "MovProveedores.Proveedor) INNER JOIN VencimientoProveedor ON MovProveedores.ID = " + _
            "VencimientoProveedor.IdMov WHERE VencimientoProveedor.Vencimiento < #" + _
            stFechaSQL(DTFecha) + "# " + _
            "ORDER BY VencimientoProveedor.Vencimiento ASC", DB.CN, adOpenStatic, adLockReadOnly
    End If
    
    Set DGVenc.DataSource = RSVenc
    
    DGVenc.Refresh
    AcomodarDG

End Sub

Private Sub DGVenc_HeadClick(ByVal ColIndex As Integer)
    RSVenc.Sort = DGVenc.Columns(ColIndex).DataField
End Sub

Private Sub DTFecha_Change()
    RefreshDG
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    txtIntDesc = FormatCurrency(0)
    DTFecha = Date + CLng(CFG.GetInfo(2, 4))
End Sub

Private Sub AcomodarDG()
    If IsProveedor = False Then
        DGVenc.Columns("Clientes.ID").Caption = "IdCli"
        DGVenc.Columns("IDCli").Width = 700
        DGVenc.Columns("IDCli").Alignment = dbgRight
        DGVenc.Columns("Nombre").Caption = "Cliente"
        DGVenc.Columns("Cliente").Width = 1500
        DGVenc.Columns("Documento").Width = 1600
        DGVenc.Columns("Vencimientos.ID").Caption = "ID"
    Else
        DGVenc.Columns("Proveedor").Width = 1500
        DGVenc.Columns("Documento").Width = 2000
    End If
    
    DGVenc.Columns("ID").Width = 0
    DGVenc.Columns("IDMov").Width = 800
    DGVenc.Columns("IDMov").Alignment = dbgRight
    DGVenc.Columns("IDMov").Caption = "Id.Mov."
    DGVenc.Columns("Fecha").Width = 1000
    DGVenc.Columns("Fecha").Caption = "FechaDoc"
    DGVenc.Columns("Vencimiento").Width = "1000"
    DGVenc.Columns("Vencimiento").Caption = "Vencim."
    DGVenc.Columns("Demora").Alignment = dbgCenter
    DGVenc.Columns("Demora").Width = 800
    DGVenc.Columns("Cuota").Alignment = dbgCenter
    DGVenc.Columns("Cuota").Width = 1100
    DGVenc.Columns("Cuota").NumberFormat = "$0.00"
    DGVenc.Columns("Interes").Alignment = dbgCenter
    DGVenc.Columns("Interes").Width = 1000
    DGVenc.Columns("Interes").NumberFormat = "$0.00"
    DGVenc.Columns("Total").Alignment = dbgCenter
    DGVenc.Columns("Total").Width = 1200
    DGVenc.Columns("Total").NumberFormat = "$0.00"
    DGVenc.Columns("Documento").Alignment = dbgCenter
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set DGVenc.DataSource = Nothing
    
    If RSVenc.State = adStateOpen Then RSVenc.Close
    Set RSVenc = Nothing
End Sub

Private Sub txtIntDesc_GotFocus()
    PintarTxt txtIntDesc
End Sub

Private Sub txtIntDesc_LostFocus()
    txtIntDesc = FormatCurrency(ValidarNumeros(txtIntDesc), , , , vbFalse)
End Sub
