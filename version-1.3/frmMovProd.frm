VERSION 5.00
Object = "{A7FBD38D-2930-49E3-B60C-9E0202D84549}#15.0#0"; "tbrControles.ocx"
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMovProd 
   BackColor       =   &H00544B45&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resumen Movimientos de Cuentas Contables"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11235
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMovProd.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   11235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin tbrFaroButton.fBoton command1 
      Height          =   465
      Left            =   9720
      TabIndex        =   21
      Top             =   7860
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdVer 
      Height          =   375
      Left            =   7710
      TabIndex        =   22
      Top             =   2520
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   661
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "ver resumen"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.CheckBox ChkT 
      BackColor       =   &H00544B45&
      Caption         =   "Ver Todos"
      ForeColor       =   &H00E0E0E0&
      Height          =   345
      Left            =   5880
      TabIndex        =   20
      Top             =   3300
      Width           =   2655
   End
   Begin tbrControles.tbrBuscador tbrBuscadorT 
      Height          =   2235
      Left            =   510
      TabIndex        =   0
      Top             =   1590
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   3942
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
   Begin VB.ListBox lstSub 
      Height          =   2205
      Left            =   6810
      TabIndex        =   14
      Top             =   4560
      Width           =   4065
   End
   Begin VB.TextBox txtHasta 
      Alignment       =   2  'Center
      Height          =   390
      Left            =   6390
      TabIndex        =   3
      Text            =   "26-10-2002"
      Top             =   2730
      Width           =   1185
   End
   Begin VB.TextBox txtDesde 
      Alignment       =   2  'Center
      Height          =   390
      Left            =   6390
      TabIndex        =   2
      Text            =   "26-10-2002"
      Top             =   2220
      Width           =   1185
   End
   Begin MSComctlLib.ListView lvResumen 
      Height          =   2535
      Left            =   570
      TabIndex        =   9
      Top             =   4560
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   4471
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
         Text            =   "Id"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Concepto"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Cantidad"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Importe"
         Object.Width           =   2205
      EndProperty
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Estadísticas de Ventas por Tipo de Producto"
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
      Left            =   930
      TabIndex        =   1
      Top             =   360
      Width           =   9435
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Importe"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   9870
      TabIndex        =   19
      Top             =   6990
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   8760
      TabIndex        =   18
      Top             =   6990
      Width           =   975
   End
   Begin VB.Label lblSaldoConCant 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "dfgsdgsfdg"
      Height          =   405
      Left            =   8640
      TabIndex        =   17
      Top             =   7230
      Width           =   1065
   End
   Begin VB.Label lblSaldosCant 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "dfgsdgsfdg"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3660
      TabIndex        =   16
      Top             =   7170
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo con subcuentas"
      ForeColor       =   &H00E0E0E0&
      Height          =   525
      Left            =   7470
      TabIndex        =   15
      Top             =   7170
      Width           =   1155
   End
   Begin VB.Label lblSaldoCon 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "dfgsdgsfdg"
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   9840
      TabIndex        =   13
      Top             =   7230
      Width           =   1065
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Subcuentas de DIFERENCIAS DE CAJA"
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   6810
      TabIndex        =   12
      Top             =   4020
      Width           =   4035
   End
   Begin VB.Label lblSaldo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Totales"
      ForeColor       =   &H00E0E0E0&
      Height          =   345
      Left            =   1710
      TabIndex        =   11
      Top             =   7260
      Width           =   1725
   End
   Begin VB.Label lblSaldos 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "dfgsdgsfdg"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   5040
      TabIndex        =   10
      Top             =   7170
      Width           =   1335
   End
   Begin VB.Label lblRegDesde 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "dfgsdgsfdg"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   675
      Left            =   5880
      TabIndex        =   8
      Top             =   1350
      Width           =   2925
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Resumen de DIFERENCIAS DE CAJA DESDE EL 26-10-2006 AL 23-03-2007"
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   960
      TabIndex        =   7
      Top             =   3930
      Width           =   4485
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta:"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   5310
      TabIndex        =   6
      Top             =   2790
      Width           =   1035
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Desde:"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   5310
      TabIndex        =   5
      Top             =   2250
      Width           =   1035
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione Tipo Producto (Nombre o ID)"
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   1350
      TabIndex        =   4
      Top             =   1260
      Width           =   4455
   End
End
Attribute VB_Name = "frmMovProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ActuTbrB As Boolean
Dim FechaDesde As Date
Dim FechaHasta As Date 'estos son el 1er y ultimo registro
Dim ClsP As New clsProducto

Dim Desde As Date
Dim Hasta As Date 'estos los que elige el usuario
Dim Saldo As Single, SaldoSub As Single
Dim SaldoCant As Long, SaldoSubCant As Long

Private Sub chkT_Click()
    txtDesde_LostFocus
    txtHasta_LostFocus
    
    VerDaTos Desde, Hasta
End Sub

Private Sub cmdVer_Click()
    txtDesde_LostFocus
    txtHasta_LostFocus
    
    VerDaTos Desde, Hasta
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    ActuTbrB = False
    FechaDesde = CDate(DB.GetTop1Rs("Ventas", "Fecha", "ASC", , True))
    FechaHasta = CDate(DB.GetTop1Rs("Ventas", "Fecha", , , True))
    
    Desde = FechaDesde
    Hasta = FechaHasta
    
    tbrBuscadorT.Contrasena = Contrasena
    tbrBuscadorT.ArchivoMDB = ArchivoMDBPrincipal
    tbrBuscadorT.SqlSinLike = "SELECT TOP 50 Id2, TipoProducto FROM TipoProductos " + _
        "WHERE ID2 > 0"
    tbrBuscadorT.OrderBy = "ORDER BY Id2"
    tbrBuscadorT.CampoEnQueBuscar = "Id2/n,TipoProducto/b"
    tbrBuscadorT.ColumnasSepPorComasyParentesis = "ID(700)/Nombre(2650)"
    tbrBuscadorT.Text = "saf"
    tbrBuscadorT.Text = "saf"
    tbrBuscadorT.Text = "" 'XXXX Negradaza
    
    lblTitulo = ""
    lblSaldos = FormatCurrency(0)
    lblSaldosCant = FormatCurrency(0)
    lblSaldoCon = FormatCurrency(0)
    lblSaldoConCant = FormatCurrency(0)
    
    VerDaTos Desde, Hasta
    
    lblRegDesde.Caption = "Tiene Registros desde: " + CStr(FechaDesde)
    txtDesde = CStr(FechaDesde)
    txtHasta = CStr(FechaHasta)
End Sub
    
Private Sub VerDaTos(Dde As Date, Hta As Date)
    Dim IDC As Long, SP() As String
    
    If ChkT Then
        IDC = -1
    Else
        If tbrBuscadorT.GetLstSel = "" Then
            Label4 = ""
            lblTitulo = ""
            lblSaldos = FormatCurrency(0)
            lblSaldosCant = FormatCurrency(0)
            lblSaldoCon = FormatCurrency(0)
            lblSaldoConCant = FormatCurrency(0)
            Exit Sub
        End If
        IDC = CLng(tbrBuscadorT.GetLstSel)
    End If
    
    SP = Split(ClsP.ListarMovTipoProducto(IDC, Dde, Hta, lvResumen), "|")
    Saldo = CSng(SP(0))
    SaldoCant = CLng(SP(1))
    SaldoSub = 0: SaldoSubCant = 0
    
    lblSaldos = FormatCurrency(Saldo, , , , vbFalse)
    lblSaldosCant = CStr(SaldoCant)
    
    If IDC > 0 Then
        lblTitulo = "Resumen de Ventas de " + UCase(tbrBuscadorT.GetLstSel(1)) + _
            " desde el " + CStr(Dde) + " AL " + CStr(Hta)
    Else
        lblTitulo = "Resumen de Ventas por Tipos de Productos Principales"
    End If
    
    VerSub IDC
End Sub
    
Private Sub VerSub(IDC As Long)
    Dim Ctas() As String, A As Long, IdQ As Long, SP1() As String
    Dim Nivel As Long, Sumo As Single, Canti As Long
    
    Sumo = 0: Canti = 0
        
    If IDC < 0 Then
        'estan los tipos de productos da el indice en negativo!!!!
        IdQ = CLng(txtInLvW(lvResumen, -IDC, 0))
        SP1 = Split(ClsP.ListarMovTipoProducto(IdQ, Desde, Hasta, , True), "|")
        SaldoSub = CSng(SP1(0))
        SaldoSubCant = CLng(SP1(1))
    Else
        IdQ = CLng(tbrBuscadorT.GetLstSel)
        
        SaldoSub = Saldo
        SaldoSubCant = SaldoCant
    End If
    
    lstSub.Clear
    Nivel = 2
    
    Dim Nb As String
    
    Nb = UCase(DB.GetValInRS("TipoProductos", "TipoProducto", "ID2 = " + CStr(IdQ)))
    lstSub.AddItem "." + CStr(IdQ) + "." + Nb
    
    Label4 = "SubTipos de Producto de " + Nb
    
    Ctas = ClsP.GetHijoTipo(IdQ)
    
    If UBound(Ctas) = 0 Then
        lstSub.AddItem "¡No tiene subcuentas!"
        lblSaldoCon = FormatCurrency(SaldoSub, , , , vbFalse)
        lblSaldoConCant = CStr(SaldoSubCant)
        Exit Sub
    End If
    
    For A = 1 To UBound(Ctas)
        SP1 = Split(ClsP.ListarMovTipoProducto(CLng(Ctas(A)), Desde, Hasta, , True), "|")
        Sumo = CSng(SP1(0))
        Canti = CLng(SP1(1))
        
        'ponerle los mismos espacios que tenia mas 3
        lstSub.AddItem String(Nivel * 3, " ") + "." + Ctas(A) + "." + _
            DB.GetValInRS("TipoProductos", "TipoProducto", "Id2 = " + Ctas(A)) + _
            " (" + CStr(Canti) + ") : " + _
            FormatCurrency(Sumo, , , , vbFalse)
        AnadirHijos CLng(Ctas(A)), Nivel
        SaldoSub = SaldoSub + Sumo
        SaldoSubCant = SaldoSubCant + Canti
    Next A
    
    lblSaldoCon = FormatCurrency(SaldoSub, , , , vbFalse)
    lblSaldoConCant = CStr(SaldoSubCant)
End Sub

Private Sub AnadirHijos(IdCta As Long, Nivel As Long)
    Dim Hij() As String, Niv As Long, A As Long, Sumo As Single
    Dim SP1() As String, Canti As Long
    
    Hij = ClsP.GetHijoTipo(IdCta)
    Niv = Nivel + 1
    
    Canti = 0: Sumo = 0
    
    If UBound(Hij) > 0 Then
        For A = 1 To UBound(Hij)
            SP1 = Split(ClsP.ListarMovTipoProducto(CLng(Hij(A)), Desde, Hasta, , True), "|")
            Sumo = CSng(SP1(0))
            Canti = CLng(SP1(1))
            
            'ponerle los mismos espacios que tenia mas 3
            lstSub.AddItem String(Niv * 3, " ") + "." + Hij(A) + "." + _
                DB.GetValInRS("TipoProductos", "TipoProducto", "Id2 = " + Hij(A)) + _
                " (" + CStr(Canti) + ") : " + _
                FormatCurrency(Sumo, , , , vbFalse)
            AnadirHijos CLng(Hij(A)), Niv
            SaldoSub = SaldoSub + Sumo
            SaldoSubCant = SaldoSubCant + Canti
        Next A
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    tbrBuscadorT.CN_CLOSE
    Set ClsP = Nothing
End Sub

Private Sub lvResumen_Click()
    If lvResumen.ListItems.Count = 0 Then Exit Sub
    
    If InStrRev(lblTitulo, "Tipos de Producto") > 0 Then
        'son los tipos de productos listados (va el indice negativo negativo!!!)
        VerSub -lvResumen.SelectedItem.Index
    Else
        'son los productos
        VerSub CLng(txtInLvW(lvResumen, lvResumen.SelectedItem.Index, 0))
    End If
End Sub

Private Sub tbrBuscadorT_Click()
    txtDesde_LostFocus
    txtHasta_LostFocus
    
    VerDaTos Desde, Hasta
End Sub

Private Sub txtDesde_GotFocus()
    txtDesde.SelStart = 0
    txtDesde.SelLength = Len(txtDesde)
End Sub

Private Sub txtDesde_LostFocus()
    If IsDate(txtDesde) Then
        Select Case CDate(txtDesde)
            Case Is < FechaDesde
                Desde = FechaDesde
            Case Is >= CDate(txtHasta)
                Desde = CDate(txtHasta)
                Hasta = Desde
            Case Else
                Desde = CDate(txtDesde)
        End Select
    
    End If
    
    txtHasta = CStr(Hasta)
    txtDesde = CStr(Desde)
End Sub

Private Sub txtHasta_GotFocus()
    txtHasta.SelStart = 0
    txtHasta.SelLength = Len(txtHasta)
End Sub

Private Sub txtHasta_LostFocus()
    If IsDate(txtHasta) Then
        Select Case CDate(txtHasta)
            Case Is > FechaHasta
                Hasta = FechaHasta
            Case Is <= CDate(txtDesde)
                Hasta = CDate(txtHasta)
                Desde = Hasta
            Case Else
                Hasta = CDate(txtHasta)
        End Select
    
    End If
    
    txtHasta = CStr(Hasta)
    txtDesde = CStr(Desde)
End Sub
    
Private Sub tbrBuscadorT_Change()
    ChkT.Value = 0
    If IsNumeric(tbrBuscadorT.Text) Then
        tbrBuscadorT.CampoEnQueBuscar = "ID2/b,TipoProducto"
    Else
        tbrBuscadorT.CampoEnQueBuscar = "ID2/n,TipoProducto/b"
    End If
    
    If tbrBuscadorT.Text <> "" Then VerActu
    tbrBuscadorT_Click
End Sub

Private Sub VerActu()
    If ActuTbrB = False Then
       ActuTbrB = True
       tbrBuscadorT.Recargar
    Else
        ActuTbrB = False
    End If

End Sub

