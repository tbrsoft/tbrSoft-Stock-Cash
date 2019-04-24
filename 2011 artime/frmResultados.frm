VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmResultados 
   BackColor       =   &H00544B45&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resultados por Tipo de Producto"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10455
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmResultados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin tbrFaroButton.fBoton cmdVer 
      Height          =   390
      Left            =   4005
      TabIndex        =   17
      Top             =   3270
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   688
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "ver resumen"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin MSComctlLib.ListView LVResumen 
      Height          =   2355
      Left            =   450
      TabIndex        =   14
      Top             =   4080
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   4154
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tipo Producto"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Ventas"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Costo"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Gcia.Bruta"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   " % Gcia. Bruta"
         Object.Width           =   1887
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00544B45&
      Caption         =   "Resultados Globales"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2505
      Left            =   6030
      TabIndex        =   5
      Top             =   1350
      Width           =   3585
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ventas"
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   0
         Left            =   390
         TabIndex        =   13
         Top             =   480
         Width           =   1365
      End
      Begin VB.Label lblGlobal 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "111"
         Height          =   375
         Index           =   0
         Left            =   1890
         TabIndex        =   12
         Top             =   450
         Width           =   1365
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Costo de Ventas"
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   1
         Left            =   270
         TabIndex        =   11
         Top             =   990
         Width           =   1515
      End
      Begin VB.Label lblGlobal 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   1
         Left            =   1890
         TabIndex        =   10
         Top             =   930
         Width           =   1365
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ganancia Bruta"
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   2
         Left            =   225
         TabIndex        =   9
         Top             =   1530
         Width           =   1530
      End
      Begin VB.Label lblGlobal 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   2
         Left            =   1890
         TabIndex        =   8
         Top             =   1470
         Width           =   1365
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "% Gan. Bruta"
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   3
         Left            =   450
         TabIndex        =   7
         Top             =   2010
         Width           =   1305
      End
      Begin VB.Label lblGlobal 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1111"
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Index           =   3
         Left            =   1860
         TabIndex        =   6
         Top             =   2010
         Width           =   1365
      End
   End
   Begin VB.TextBox txtDesde 
      Alignment       =   2  'Center
      Height          =   345
      Left            =   2640
      TabIndex        =   0
      Text            =   "26-10-2002"
      Top             =   2820
      Width           =   1110
   End
   Begin VB.TextBox txtHasta 
      Alignment       =   2  'Center
      Height          =   345
      Left            =   2640
      TabIndex        =   1
      Text            =   "26-10-2002"
      Top             =   3330
      Width           =   1110
   End
   Begin tbrFaroButton.fBoton cmdImprimir 
      Height          =   390
      Left            =   4260
      TabIndex        =   18
      Top             =   6660
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   688
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "imprimir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton command1 
      Height          =   390
      Left            =   8910
      TabIndex        =   19
      Top             =   6660
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   688
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione el período que desee conocer los resultados por tipos de producto"
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   960
      TabIndex        =   16
      Top             =   2250
      Width           =   4665
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Resultados por Tipo de Producto"
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
      Height          =   645
      Left            =   1740
      TabIndex        =   15
      Top             =   465
      Width           =   7155
   End
   Begin VB.Label lblRegDesde 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "dfgsdgsfdg"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   990
      TabIndex        =   4
      Top             =   1680
      Width           =   4485
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Desde:"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   2850
      Width           =   1035
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta:"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   3390
      Width           =   1035
   End
End
Attribute VB_Name = "frmResultados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FechaDesde As Date
Dim FechaHasta As Date 'estos son el 1er y ultimo registro

Dim Desde As Date
Dim Hasta As Date 'estos los que elige el usuario

Private Sub cmdImprimir_Click()
    Dim Tit() As String

    TP.LineasSeparadoras = True
    
    ReDim Preserve Tit(4)
    Tit(4) = "Resultados por Tipo de Producto"
    'datos de mi empresa!!!!!!!!!!!!!!
    Tit(0) = DB.GetValInRS("Clientes", "Nombre", "ID = -2", True)
    Tit(1) = "Direccion: " + DB.GetValInRS("Clientes", "Direccion", "ID = -2", True)
    Tit(2) = "Teléfono: " + DB.GetValInRS("Clientes", "Telefono", "ID = -2", True)
    Tit(3) = "Mail: " + DB.GetValInRS("Clientes", "Mail", "ID = -2", True)
    
    TP.ImprimirlvW lvResumen, Tit, _
        "ID|TipoProducto|Ventas|Costo|Gcia Bruta|% GB", _
        "Desde el " + txtDesde + " Hasta el " + txtHasta
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub cmdVer_Click()
    VerDaTos CDate(txtDesde), CDate(txtHasta)
End Sub

Private Sub VerDaTos(Desde As Date, Hasta As Date)
    lvResumen.ListItems.Clear
    
    Dim Ventas As Single, CostoVentas As Single
    Dim RsT As New ADODB.Recordset
    
    Dim stGan As String, SpRes() As String, stPorC As String, YY As Long
    
    Ventas = 0: CostoVentas = 0: YY = 0 'yy es el indice del listview
    
    RsT.Open "SELECT Id2, TipoProducto FROM TipoProductos WHERE ID2 > 0", DB.CN, adOpenStatic, adLockReadOnly
    'extras no por que tienes cosas como iva y descuentos
    If RsT.RecordCount > 0 Then
        RsT.MoveFirst
        Do While Not RsT.EOF
            SpRes = Split(StrResultados(CLng(RsT("ID2")), Desde, Hasta), "/")
            If CSng(SpRes(0)) > 0 Then 'ignoro los tipo de productos sin ventas
                Ventas = Ventas + CSng(SpRes(0))
                CostoVentas = CostoVentas + CSng(SpRes(1))
                
                stGan = FormatCurrency(CSng(SpRes(0)) - CSng(SpRes(1)))
                If SpRes(1) > 0 Then 'no se puede dividir en 0
                    stPorC = FormatPercent(CSng(SpRes(0)) / CSng(SpRes(1)) - 1)
                Else
                    stPorC = "N/A"
                End If
                
                YY = lvResumen.ListItems.Count + 1
                
                lvResumen.ListItems.Add YY
                lvResumen.ListItems(YY).Text = RsT("ID2")
                lvResumen.ListItems(YY).SubItems(1) = RsT("Tipoproducto")
                lvResumen.ListItems(YY).SubItems(2) = FormatCurrency(SpRes(0))
                lvResumen.ListItems(YY).SubItems(3) = FormatCurrency(SpRes(1))
                lvResumen.ListItems(YY).SubItems(4) = stGan
                lvResumen.ListItems(YY).SubItems(5) = stPorC
                
            End If
            
            RsT.MoveNext
        Loop
    End If
    
    Dim tmPC As String
    
    If CostoVentas > 0 Then 'no se puede dividir en 0
        tmPC = FormatPercent(Ventas / CostoVentas - 1)
    Else
        tmPC = "N/A"
    End If

    lblGlobal(0) = FormatCurrency(Ventas)
    lblGlobal(1) = FormatCurrency(CostoVentas)
    lblGlobal(2) = FormatCurrency(Ventas - CostoVentas)
    lblGlobal(3) = tmPC
    
    RsT.Close
    Set RsT = Nothing
End Sub

Private Sub Form_Load()
    Dim rSRD As New ADODB.Recordset, I As Integer
    
    lvResumen.View = lvwReport
    
    rSRD.Open "SELECT Fecha FROM Ventas ORDER BY Fecha", DB.CN, adOpenStatic, adLockReadOnly
    
    If rSRD.RecordCount = 0 Then
        lblRegDesde = "No tiene registros"
        cmdVer.Enabled = False
        FechaDesde = Date
        FechaHasta = Date
        
        rSRD.Close
        Set rSRD = Nothing
            
        lvResumen.ListItems.Clear
        
        For I = 0 To 2
            lblGlobal(I) = FormatCurrency(0)
        Next I
        lblGlobal(I) = FormatPercent(0)
            
        Exit Sub
    End If
    
    rSRD.MoveFirst
    FechaDesde = CDate(rSRD("Fecha"))
    Desde = FechaDesde
    
    rSRD.MoveLast
    FechaHasta = CDate(rSRD("Fecha"))
    Hasta = FechaHasta
    
    VerDaTos Desde, Hasta
    
    lblRegDesde.Caption = "Tiene Registros desde: " + CStr(FechaDesde)
    txtDesde = CStr(FechaDesde)
    txtHasta = CStr(FechaHasta)
        
    rSRD.Close
    Set rSRD = Nothing
    
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

'devuelve $Vta/$Cto
Private Function StrResultados(IdTipo As Long, yDesde As Date, yHasta As Date) As String
    Dim SA As String
    Dim Rss As New ADODB.Recordset
    
    SA = "SELECT Sum([Cantidad]*[Precio]) AS Vt, " + _
        "Sum([Cantidad]*[Costo]) AS Ct " + _
        "FROM TipoProductos INNER JOIN (Productos INNER JOIN Ventas ON " + _
        "Productos.ID = Ventas.IDproducto) ON " + _
        "TipoProductos.ID2 = Productos.IdTipoProducto " + _
        "WHERE (((Ventas.Fecha) BETWEEN #" + stFechaSQL(yDesde) + " 00:00:00 # " + _
        "AND #" + stFechaSQL(yHasta) + " 23:59:59#) AND ((Productos.IdTipoProducto)= " + _
        CStr(IdTipo) + "))"

    Rss.Open SA, DB.CN, adOpenStatic, adLockReadOnly
    If Rss.RecordCount = 0 Then
        StrResultados = "0/0"
    Else
        StrResultados = CStr(NoNuloN(Rss("Vt"))) + "/" + CStr(NoNuloN(Rss("Ct")))
    End If
    
    Rss.Close
    Set Rss = Nothing
End Function

