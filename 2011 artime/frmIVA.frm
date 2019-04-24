VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmIVA 
   BackColor       =   &H004E4E4E&
   Caption         =   "Form1"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10140
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIVA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8265
   ScaleWidth      =   10140
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin tbrFaroButton.fBoton cmdImprimir 
      Height          =   495
      Left            =   3750
      TabIndex        =   6
      Top             =   7260
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   873
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Imprimir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin MSComctlLib.ListView lvIVA 
      Height          =   4785
      Left            =   960
      TabIndex        =   3
      Top             =   1860
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   8440
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Fecha"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Documento"
         Object.Width           =   3440
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Razon Social"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Neto"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "IVA"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Total"
         Object.Width           =   2469
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTDe 
      Height          =   345
      Left            =   2340
      TabIndex        =   0
      Top             =   1230
      Width           =   1425
      _ExtentX        =   2514
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
      Format          =   21102593
      CurrentDate     =   39197
   End
   Begin MSComCtl2.DTPicker DTA 
      Height          =   345
      Left            =   5070
      TabIndex        =   1
      Top             =   1200
      Width           =   1485
      _ExtentX        =   2619
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
      Format          =   21102593
      CurrentDate     =   39197
   End
   Begin tbrFaroButton.fBoton cmdSalir 
      Height          =   495
      Left            =   8370
      TabIndex        =   7
      Top             =   7260
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   873
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
      Caption         =   "Hasta"
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   3930
      TabIndex        =   5
      Top             =   1260
      Width           =   1005
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Desde"
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   1200
      TabIndex        =   4
      Top             =   1290
      Width           =   1005
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "iva"
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
      Height          =   405
      Left            =   3960
      TabIndex        =   2
      Top             =   390
      Width           =   1695
   End
End
Attribute VB_Name = "frmIVA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IsVentas As Boolean
Dim Conceptos() As String

Public Sub AbrirDatos(EsVentas As Boolean)
    IsVentas = EsVentas
    
    If IsVentas Then
        Label1 = "IVA Ventas"
        Me.Caption = "IVA Ventas"
    Else
        Label1 = "IVA Compras"
        Me.Caption = "IVA Compras"
    End If
    
    Me.Show 1
End Sub

Private Sub cmdImprimir_Click()
    Dim TitCol As String, I As Long
    Dim Tit() As String

    TP.LineasSeparadoras = True
    
    ReDim Preserve Tit(4)
    Tit(4) = Label1
    'datos de mi empresa!!!!!!!!!!!!!!
    Tit(0) = DB.GetValInRS("Clientes", "Nombre", "ID = -2", True)
    Tit(1) = "Direccion: " + DB.GetValInRS("Clientes", "Direccion", "ID = -2", True)
    Tit(2) = "Teléfono: " + DB.GetValInRS("Clientes", "Telefono", "ID = -2", True)
    Tit(3) = "Mail: " + DB.GetValInRS("Clientes", "Mail", "ID = -2", True)
    
    TitCol = ""
    For I = 1 To lvIVA.ColumnHeaders.Count
        If Len(lvIVA.ColumnHeaders(I).Text) > 10 Then
            TitCol = TitCol + Left(lvIVA.ColumnHeaders(I).Text, 7) + "..."
        Else
            TitCol = TitCol + lvIVA.ColumnHeaders(I).Text
        End If
        
        If I < lvIVA.ColumnHeaders.Count Then TitCol = TitCol + "|"
    Next I

    TP.ImprimirlvW lvIVA, Tit, TitCol, , , , , CStr(DTDe) + " - " + CStr(DTA)
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub DTA_Change()
    CargarDatos
End Sub

Private Sub DTDe_Change()
    CargarDatos
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    ReDim Conceptos(0)
    Conceptos(0) = "Nada"
    Me.Width = 11000
    lvIVA.Width = 10000
    lvIVA.Left = 500
    'del 1 al 8 estan los conceptos
    VerConceptos
    VerColumnas
    DTDe = Date - Day(Date) + 1
    DTA = ProximoMes(DTDe) - 1
    
    CargarDatos
End Sub

Private Sub VerColumnas()
    Dim I As Long
    
    If UBound(Conceptos) > 0 Then
        For I = 1 To UBound(Conceptos)
            lvIVA.ColumnHeaders.Add I + 5, , Conceptos(I), 1100, vbCenter
        Next I
    End If
    
    AjustarAnchos
End Sub

Private Sub VerConceptos()
    Dim X As Long, SP() As String, Hijos() As String, Ix As Long, Como As String
    
    'del 1 al 8 estan los conceptos
    Hijos = CFG.GetHijos(7)
    
    If IsVentas Then
        Como = "Concepto "
    Else
        Como = "ConceptoC"
    End If
    
    Ix = 0
    For X = 1 To UBound(Hijos)
        If InStrRev(CFG.GetInfo(CLng(Hijos(X)), 2), Como) <> 0 Then
            SP = Split(CFG.GetInfo(CLng(Hijos(X)), 4), "_")
            
            Ix = Ix + 1
            ReDim Preserve Conceptos(Ix)
            Conceptos(Ix) = SP(1)
        End If
    Next X
End Sub

Private Sub AjustarAnchos()
    'con la idea de que me.wid=11000 , lviva.wid=10000
    Dim TotAncho As Single, X As Long, Ajuste As Single
    
    TotAncho = 0
    
    For X = 1 To lvIVA.ColumnHeaders.Count
        TotAncho = TotAncho + lvIVA.ColumnHeaders(X).Width
    Next X
    
    Ajuste = (Me.Width - 2900) / TotAncho
    
    For X = 1 To lvIVA.ColumnHeaders.Count
        lvIVA.ColumnHeaders(X).Width = Ajuste * lvIVA.ColumnHeaders(X).Width
    Next X
End Sub

Public Sub CargarDatos()
    lvIVA.ListItems.Clear
    'veo que facturas usa
    Dim FacturasUsa() As String, strFU As String, Y As Long, Tabla As String
    Dim Esta As Boolean, Columnas As Long, R As Long, Tot As Single, TmP As Single
    Dim NbCli As String, IDC As Long
    
    ReDim FacturasUsa(0)
    FacturasUsa(0) = "Nada"
    
    ' Tiene que existir la propiedades 101 y 102 Facturas IVAC y Facturas IVAV
    If IsVentas Then
        'Ventas tambien van a usar las que no discrimina como a consumidor final
        strFU = CFG.GetInfo(102, 4) + CFG.GetInfo(102, 3)
        Tabla = "Facturas"
    Else
        strFU = CFG.GetInfo(101, 4)
        Tabla = "FacturaCompra"
    End If
    
    If strFU = "" Then strFU = "A"
    
    For Y = 1 To Len(strFU)
        ReDim Preserve FacturasUsa(Y)
        FacturasUsa(Y) = Mid(strFU, Y, 1)
    Next Y
    
    'cargo nomas --------------------------------------------------------------
    Dim RsQ As New ADODB.Recordset
    
    Columnas = lvIVA.ColumnHeaders.Count
    RsQ.Open "SELECT Fecha, NroFactura FROM " + Tabla + _
        " WHERE Fecha BETWEEN #" + stFechaSQL(DTDe) + " 0:00# AND #" + _
        stFechaSQL(DTA) + " 23:59#", DB.CN, adOpenStatic, adLockReadOnly
    
    If RsQ.RecordCount > 0 Then
        RsQ.MoveFirst
        
        Do While Not RsQ.EOF
            Esta = False
            Tot = 0
            For Y = 1 To UBound(FacturasUsa)
                If Left(RsQ("NroFactura"), 1) = FacturasUsa(Y) Then
                    Esta = True
                    Exit For
                End If
            Next Y
            
            If Esta Then
                Y = lvIVA.ListItems.Count + 1
            
                lvIVA.ListItems.Add Y
                
                lvIVA.ListItems(Y).Text = CStr(NoNuloD(RsQ("Fecha")))
                lvIVA.ListItems(Y).SubItems(1) = NoNuloS(RsQ("NroFactura"))
                lvIVA.ListItems(Y).SubItems(Columnas - 1) = FormatCurrency( _
                        DB.GetValInRS(Tabla, "Pagado", "NroFactura = '" + _
                        NoNuloS(RsQ("NroFactura")) + "'", False))
                If IsVentas Then
                    'Razon Social
                    IDC = DB.GetValInRS("Facturas", "IdCliente", "NroFactura = '" + _
                        NoNuloS(RsQ("NroFactura")) + "'", False)
                    If IDC = 0 Then
                        NbCli = "Consumidor Final"
                    Else
                        NbCli = DB.GetValInRS("Clientes", "Nombre", "ID = " + CStr(IDC))
                    End If
                    lvIVA.ListItems(Y).SubItems(2) = NbCli
                    
                    'Neto
                    TmP = DB.SumarProducto("Ventas", "Cantidad", "Precio", "IdVenta = '" + _
                        NoNuloS(RsQ("NroFactura")) + "' AND IDProducto >=-1")
                    lvIVA.ListItems(Y).SubItems(3) = FormatCurrency(TmP)
                    Tot = Tot + TmP
                    'IVA
                    TmP = DB.SumarValInRS("Ventas", "Precio", "IdVenta = '" + _
                        NoNuloS(RsQ("NroFactura")) + "' AND IDProducto =-2")
                    lvIVA.ListItems(Y).SubItems(4) = FormatCurrency(TmP)
                    Tot = Tot + TmP
                    'Conceptos
                    If Columnas > 6 Then
                        For R = 5 To Columnas - 1
                            TmP = DB.SumarValInRS("Ventas", "Precio", "IdVenta = '" + _
                                NoNuloS(RsQ("NroFactura")) + "' AND " + _
                                "IDProducto = " + CStr(-(6 + R)))
                            lvIVA.ListItems(Y).SubItems(R) = FormatCurrency(TmP)
                            Tot = Tot + TmP
                        Next R
                    End If
                    'Total
                    lvIVA.ListItems(Y).SubItems(Columnas - 1) = FormatCurrency(Tot)
                    
                Else 'COMPRAS------------------------------------------------------
                    'Razon Social
                    NbCli = DB.GetValInRS("FacturaCompra", "Proveedor", _
                        "NroFactura = '" + NoNuloS(RsQ("NroFactura")) + "'")
                    lvIVA.ListItems(Y).SubItems(2) = NbCli
                    
                    'Neto
                    TmP = DB.SumarValInRS( _
                        "CompraDetalle", "PrecioTotal", "NroFactura = '" + _
                        NoNuloS(RsQ("NroFactura")) + "' AND IDProducto >=-1")
                    lvIVA.ListItems(Y).SubItems(3) = FormatCurrency(TmP)
                    Tot = Tot + TmP
                    'IVA
                    TmP = DB.SumarValInRS( _
                        "CompraDetalle", "PrecioTotal", "NroFactura = '" + _
                        NoNuloS(RsQ("NroFactura")) + "' AND IDProducto =-2")
                    lvIVA.ListItems(Y).SubItems(4) = FormatCurrency(TmP)
                    Tot = Tot + TmP
                    'Conceptos
                    If Columnas > 6 Then
                        For R = 5 To Columnas - 1
                            TmP = DB.SumarValInRS("CompraDetalle", "PrecioTotal", _
                                "NroFactura = '" + NoNuloS(RsQ("NroFactura")) + _
                                "' AND " + "IDProducto = " + CStr(-(6 + R)))
                            lvIVA.ListItems(Y).SubItems(R) = FormatCurrency(TmP)
                            Tot = Tot + TmP
                        Next R
                    End If
                    'Total
                    lvIVA.ListItems(Y).SubItems(Columnas - 1) = FormatCurrency(Tot)
                End If
            End If
            RsQ.MoveNext
        Loop
    End If
    RsQ.Close
    Set RsQ = Nothing
    
    'AGREGO LOS TOTALES
    Y = lvIVA.ListItems.Count + 1
    
    lvIVA.ListItems.Add Y
    lvIVA.ListItems.Add Y + 1
    lvIVA.ListItems.Add Y + 2
    
    Y = Y + 2
    
    lvIVA.ListItems(Y).Text = CStr(DTA)
    lvIVA.ListItems(Y).SubItems(1) = "TOTALES"
    
    For R = 3 To Columnas - 1
        lvIVA.ListItems(Y).SubItems(R) = FormatCurrency(SumarColumna(R))
    Next R
End Sub

Private Function SumarColumna(Col As Long)
    Dim X As Long, Resp As Single
    
    Resp = 0
    
    For X = 1 To lvIVA.ListItems.Count
        If IsNumeric(lvIVA.ListItems(X).SubItems(Col)) Then
            Resp = Resp + CSng(lvIVA.ListItems(X).SubItems(Col))
        End If
    Next X
    
    SumarColumna = Resp
End Function

Private Sub Form_Resize()
    If Me.Width < 10000 Or Me.Height < 7000 Then Exit Sub
    
    lvIVA.Width = Me.Width - 2500
    lvIVA.Left = 1000
    lvIVA.Height = Me.Height - 3500
    
    cmdImprimir.Top = Me.Height - 1100
    cmdImprimir.Left = lvIVA.Width / 2
    cmdSalir.Top = cmdImprimir.Top
    cmdSalir.Left = lvIVA.Width + 1000 - cmdSalir.Width
    
    AjustarAnchos
End Sub

