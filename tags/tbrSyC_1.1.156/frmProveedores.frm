VERSION 5.00
Object = "{A7FBD38D-2930-49E3-B60C-9E0202D84549}#15.0#0"; "tbrControles.ocx"
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProveedores 
   BackColor       =   &H00544B45&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proveedores"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10380
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProveedores.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   10380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin tbrFaroButton.fBoton cmdMostrar 
      Height          =   465
      Left            =   8430
      TabIndex        =   24
      Top             =   6900
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Mostrar Ranking"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdNuevo 
      Height          =   465
      Left            =   8250
      TabIndex        =   23
      Top             =   2850
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "agregar nuevo"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdLimpiar 
      Height          =   465
      Left            =   7110
      TabIndex        =   22
      Top             =   2850
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Limpiar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdModificar 
      Height          =   405
      Left            =   7260
      TabIndex        =   21
      Top             =   570
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Modificar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin MSComctlLib.ListView lvRanking 
      Height          =   1725
      Left            =   7320
      TabIndex        =   20
      Top             =   4530
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   3043
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Proveedor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Importe"
         Object.Width           =   2293
      EndProperty
   End
   Begin MSComctlLib.ListView lvProveedores 
      Height          =   2565
      Left            =   240
      TabIndex        =   17
      Top             =   390
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   4524
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Proveedor"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Telefono"
         Object.Width           =   2734
      EndProperty
   End
   Begin tbrControles.MouTextBox txtNDias 
      Height          =   435
      Left            =   7260
      TabIndex        =   4
      Top             =   6900
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   767
      Alignment       =   2
      BackColor       =   16777215
      Text            =   "30"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00544B45&
      Caption         =   "Ultimas 10 Compras Proveedor"
      ForeColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   60
      TabIndex        =   15
      Top             =   3480
      Width           =   3045
      Begin MSComctlLib.ListView lvUltComprasP 
         Height          =   2985
         Left            =   90
         TabIndex        =   19
         Top             =   270
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   5265
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Importe"
            Object.Width           =   2187
         EndProperty
      End
   End
   Begin VB.TextBox txtNombre 
      Height          =   345
      Left            =   5040
      TabIndex        =   0
      Top             =   330
      Width           =   2055
   End
   Begin VB.Frame frmUltCompras 
      BackColor       =   &H00544B45&
      Caption         =   "Ultimas 10 Compras Generales"
      ForeColor       =   &H00FFFFFF&
      Height          =   3405
      Left            =   3150
      TabIndex        =   8
      Top             =   3480
      Width           =   4065
      Begin MSComctlLib.ListView lvUltCompras 
         Height          =   3045
         Left            =   90
         TabIndex        =   18
         Top             =   240
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   5371
         View            =   3
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
            Text            =   "Fecha"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Proveedor"
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Importe"
            Object.Width           =   2117
         EndProperty
      End
   End
   Begin VB.TextBox txtDetalle 
      Height          =   915
      Left            =   5010
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2490
      Width           =   2025
   End
   Begin VB.TextBox txtDireccion 
      Height          =   585
      Left            =   5040
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1590
      Width           =   2055
   End
   Begin VB.TextBox txtTelefono 
      Height          =   345
      Left            =   5040
      TabIndex        =   1
      Top             =   930
      Width           =   2055
   End
   Begin tbrFaroButton.fBoton command1 
      Height          =   405
      Left            =   9000
      TabIndex        =   25
      Top             =   7980
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.Label lblEstadGrales 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   1005
      Left            =   3420
      TabIndex        =   16
      Top             =   6990
      Width           =   3195
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Para agregar un nuevo Proveedor previamente limpie los casilleros."
      ForeColor       =   &H00E0E0E0&
      Height          =   1455
      Left            =   7230
      TabIndex        =   14
      Top             =   1140
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H00544B45&
      Caption         =   "Días atras"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   7350
      TabIndex        =   13
      Top             =   6630
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H00544B45&
      Caption         =   "Nombre"
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   5070
      TabIndex        =   12
      Top             =   90
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H00544B45&
      Caption         =   "Ranking Compras ult 30 dias"
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   7410
      TabIndex        =   11
      Top             =   4140
      Width           =   2655
   End
   Begin VB.Label lblEstadisticas 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   1005
      Left            =   240
      TabIndex        =   10
      Top             =   6990
      Width           =   2805
   End
   Begin VB.Label Label4 
      BackColor       =   &H00544B45&
      Caption         =   "Proveedores"
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   330
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00544B45&
      Caption         =   "Detalle"
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   5070
      TabIndex        =   7
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00544B45&
      Caption         =   "Dirección"
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   5040
      TabIndex        =   6
      Top             =   1350
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00544B45&
      Caption         =   "Teléfono"
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   5070
      TabIndex        =   5
      Top             =   690
      Width           =   1095
   End
End
Attribute VB_Name = "frmProveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim Ndias As Long 'dias atras para hacer el ranking

Private Sub cmdLimpiar_Click()
    Dim TXT As Control
    For Each TXT In frmProveedores
        If TypeOf TXT Is TextBox Then TXT = ""
    Next
    
    txtNombre.SetFocus
End Sub

Private Sub cmdModificar_Click()
    If lvProveedores.ListItems.Count = 0 Then Exit Sub
    If txtNombre = "" Then MsgBox "Debes Cargar Nombre": Exit Sub
    
    Dim TmP As String
    
    TmP = GetProv(lvProveedores.SelectedItem.Index)
    
        'sacar posibles caracteres no desaedos
    If InStr(txtNombre, "'") > 0 Then
        txtNombre = Replace(txtNombre, "'", "")
        MsgBox "Se encontro el caracter ' en el nombre. Se ha quitado ya que no es valido"
    End If
        
    If InStr(txtNombre, "\") > 0 Then
        txtNombre = Replace(txtNombre, "\", "")
        MsgBox "Se encontro el caracter \ en el nombre. Se ha quitado ya que no es valido"
    End If

    If txtNombre <> TmP Then
        Dim RSp As New ADODB.Recordset
        RSp.Open "select * from proveedores where proveedor = '" + _
            txtNombre + "'", DB.CN, adOpenStatic, adLockOptimistic
        'no deberia haber ningun proveedor con el nombre de txtnombre
        
        If RSp.RecordCount > 0 Then
            MsgBox "Ya tiene registro para ese Nombre"
            txtNombre = TmP
            txtNombre.SetFocus
            PintarTxt txtNombre
        Else
        
           DB.EXECUTE "UPDATE Proveedores SET proveedor = '" + txtNombre + _
                "', Telefono = '" + txtTelefono + "', Direccion= '" + txtDireccion + _
                "', Detalle = ' " + txtDetalle + "' WHERE proveedor = '" + TmP + "'"
        End If
        
        RSp.Close
        Set RSp = Nothing
    Else
        
        Dim rSMd As New ADODB.Recordset
        'no hubo cambio de nombre entonces cambio demas datos nomas
        rSMd.Open "select * from proveedores where proveedor = '" + TmP + _
            "'", DB.CN, adOpenStatic, adLockOptimistic
    
        If NoNuloS(rSMd("telefono")) <> txtTelefono Then rSMd("telefono") = txtTelefono
        If NoNuloS(rSMd("direccion")) <> txtDireccion Then rSMd("direccion") = txtDireccion
        If NoNuloS(rSMd("detalle")) <> txtDetalle Then rSMd("detalle") = txtDetalle
    
    
        rSMd.Update
        rSMd.Close
        Set rSMd = Nothing
    
    End If
    RecargarLST
    
End Sub


Private Sub cmdMostrar_Click()
    MostrarDatos
    
End Sub

Private Sub cmdNuevo_Click()
   If txtNombre = "" Then
        MsgBox "Hay datos sin cargar", vbInformation, "Atención"
        Exit Sub
    End If
    
    'sacar posibles caracteres no desaedos
    If InStr(txtNombre, "'") > 0 Then
        txtNombre = Replace(txtNombre, "'", "")
        MsgBox "Se encontro el caracter ' en el nombre. Se ha quitado ya que no es valido"
    End If
        
    If InStr(txtNombre, "\") > 0 Then
        txtNombre = Replace(txtNombre, "\", "")
        MsgBox "Se encontro el caracter \ en el nombre. Se ha quitado ya que no es valido"
    End If
       
   Dim RSp As New ADODB.Recordset
    RSp.Open "select * from proveedores where proveedor = '" + _
        txtNombre + "'", DB.CN, adOpenStatic, adLockOptimistic
    'no deberia haber ningun proveedor con el nombre de txtnombre
    
    If RSp.RecordCount > 0 Then
        MsgBox "Ya tiene registro para ese Nombre"
        txtNombre = TmP
        txtNombre.SetFocus
        txtNombre.SelStart = 0
        txtNombre.SelLength = Len(TmP)
    Else
    
       DB.EXECUTE "INSERT INTO proveedores (proveedor,telefono,direccion,detalle)" + _
            "VALUES ('" + txtNombre + "','" + txtTelefono + "','" + txtDireccion + _
            "','" + txtDetalle + "')"
        
    End If
    
    RSp.Close
    Set RSp = Nothing

RecargarLST
 
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
   Ndias = 30 'predeterminado
   RecargarLST
   MostrarDatos
End Sub

Private Sub MostrarDatos() 'llenar todos los textbox con los datos del proveedor
    If lvProveedores.ListItems.Count = 0 Then Exit Sub
    
    Dim RsM As New ADODB.Recordset
    Dim ClsPP As New clsProveedores
    
    Label5 = "Top 5 Proveedores ult. " + CStr(Ndias) + " días"
    
    RsM.Open "SELECT direccion, detalle FROM Proveedores where proveedor = '" + _
        GetProv(lvProveedores.SelectedItem.Index) + "'", DB.CN, adOpenStatic, adLockReadOnly
    txtNombre = GetProv(lvProveedores.SelectedItem.Index)
    txtTelefono = GetFono(lvProveedores.SelectedItem.Index)
    txtDireccion = NoNuloS(RsM("direccion"))
    txtDetalle = NoNuloS(RsM("detalle"))
    
    'ESTADISTICAS
        'particulares
    Dim CPv As Single
    CPv = ClsPP.GetComprasTiempo(GetProv(lvProveedores.SelectedItem.Index), 30)
    lblEstadisticas = "Compras a " + txtNombre + " en los ultimos 30 dias: " + _
        FormatCurrency(CPv, , , , vbFalse)
    
        'globales
    Dim CTt As Single
    CTt = ClsPP.GetComprasTiempo("Todos", 30)
    lblEstadGrales = "Compras Generales en los ultimos 30 dias: " + _
        FormatCurrency(CTt, , , , vbFalse)
    
            'muestro el % del total
    Dim stPorC
    If CTt = 0 Then 'no va a dar la division
        stPorC = FormatPercent(1) 'todos los demás son 0 es raro pero le pongo 100%
    Else
        stPorC = FormatPercent(CPv / CTt)
    End If
            'lo agrego en particular
    lblEstadisticas = lblEstadisticas + " (" + stPorC + ") del total"
    
    'ULTIMAS 10 COMPRAS GRALES
    Dim XS As String
    
    XS = "SELECT TOP 10 Fecha,Proveedor,Pagado FROM FacturaCompra " + _
        "WHERE EsPedido = 0 ORDER BY Fecha DESC"
    CargarComboLV lvUltCompras, XS, "Fecha/f,Proveedor,Pagado/$"
    
    'ULTIMAS 10 COMPRAS DEL PROVEEDOR
    Dim Xms As String

    Xms = "SELECT TOP 10 Fecha, Pagado FROM FacturaCompra " + _
        "WHERE Proveedor = '" + txtNombre + _
        "' AND EsPedido = 0 ORDER BY Fecha DESC"
    CargarComboLV lvUltComprasP, Xms, "Fecha/f,Pagado/$"
    
    'RANKING DE LOS 5 MEJORES PROVEEDORES
    Dim Zr As String
    
    Zr = "SELECT TOP 5 Proveedor, Sum(Pagado)as SumaPagado FROM FacturaCompra " + _
        "WHERE Fecha > #" + stFechaSQL(Date - Ndias) + _
        "# GROUP BY Proveedor " + _
        "ORDER BY Sum(Pagado) DESC"
    CargarComboLV lvRanking, Zr, "proveedor,SumaPagado/$"
    
    
    RsM.Close
    Set RsM = Nothing
    Set ClsPP = Nothing
    
End Sub

Private Function GetProv(inD As Long) As String 'saca el nombre del textbox
    Dim Prove As String
    
    If inD < 0 Then GetProv = "": Exit Function
    
    Prove = txtInLvW(lvProveedores, inD, 0)
    
    GetProv = Prove
End Function

Private Function GetFono(inD As Long) As String 'saca el telef del textbox
    Dim Fono As String
    
    If inD < 0 Then GetFono = "": Exit Function
    
    Fono = txtInLvW(lvProveedores, inD, 1)
    
    GetFono = Fono
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    FormatearMouTextBox frmProveedores
    txtNDias = "30"
End Sub

Private Sub lvProveedores_Click()
    MostrarDatos
End Sub

Private Sub RecargarLST()
  
    CargarComboLV lvProveedores, "SELECT * FROM Proveedores ORDER BY proveedor", _
        "proveedor,telefono"

End Sub

Private Sub txtNDias_Change()
    If Not IsNumeric(txtNDias) Then
        Ndias = 0
    Else
        Ndias = CLng(Ndias)
    End If
End Sub

Private Sub txtNDias_LostFocus()
    Ndias = ValidarNumeros(txtNDias)
    txtNDias = CStr(Ndias)
End Sub
