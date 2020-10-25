VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{A7FBD38D-2930-49E3-B60C-9E0202D84549}#15.0#0"; "tbrControles.ocx"
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmMovCtasCont 
   BackColor       =   &H00544B45&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resumen Movimientos de Cuentas Contables"
   ClientHeight    =   7725
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
   Icon            =   "frmMovCtasCont.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin tbrFaroButton.fBoton cmdVer 
      Height          =   465
      Left            =   6420
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Ver Resumen"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   9
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdImprimir 
      Height          =   465
      Left            =   3060
      TabIndex        =   16
      Top             =   7020
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Imprimir"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   9
      fECol           =   5717301
   End
   Begin tbrControles.tbrBuscador tbrBuscador1 
      Height          =   2115
      Left            =   210
      TabIndex        =   0
      Top             =   510
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   3731
      BackColor       =   5524293
      BeginProperty Fontt {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox lstSub 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2790
      Left            =   6960
      TabIndex        =   15
      Top             =   3180
      Width           =   4725
   End
   Begin VB.TextBox txtHasta 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5100
      TabIndex        =   2
      Text            =   "26-10-2002"
      Top             =   1620
      Width           =   1095
   End
   Begin VB.TextBox txtDesde 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5100
      TabIndex        =   1
      Text            =   "26-10-2002"
      Top             =   1140
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvResumen 
      Height          =   2475
      Left            =   90
      TabIndex        =   10
      Top             =   3210
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4366
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Fecha"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "NºAsiento"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Detalle"
         Object.Width           =   5380
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Variacion"
         Object.Width           =   2540
      EndProperty
   End
   Begin tbrFaroButton.fBoton command1 
      Height          =   465
      Left            =   10260
      TabIndex        =   17
      Top             =   7020
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   9
      fECol           =   5717301
   End
   Begin VB.Label lblComent 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Resumen de "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   3300
      TabIndex        =   21
      Top             =   6240
      Width           =   4455
   End
   Begin VB.Label lblMov 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Resumen de "
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
      TabIndex        =   20
      Top             =   5820
      Width           =   3675
   End
   Begin VB.Label lblmenos 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "dfgsdgsfdg"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   370
      Left            =   5460
      TabIndex        =   19
      Top             =   5760
      Width           =   1350
   End
   Begin VB.Label lblMas 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "dfgsdgsfdg"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   370
      Left            =   3990
      TabIndex        =   18
      Top             =   5760
      Width           =   1350
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo con subcuentas"
      ForeColor       =   &H00E0E0E0&
      Height          =   585
      Left            =   7890
      TabIndex        =   4
      Top             =   6300
      Width           =   1755
   End
   Begin VB.Label lblSaldoCon 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "dfgsdgsfdg"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   9810
      TabIndex        =   14
      Top             =   6300
      Width           =   1845
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Subcuentas de DIFERENCIAS DE CAJA"
      ForeColor       =   &H00E0E0E0&
      Height          =   585
      Left            =   7440
      TabIndex        =   13
      Top             =   2700
      Width           =   3645
   End
   Begin VB.Label lblSaldo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Resumen de DIFERENCIAS DE CAJA DESDE EL 26-10-2006 AL 23-03-2007"
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
      Height          =   525
      Left            =   1200
      TabIndex        =   12
      Top             =   6480
      Width           =   3675
   End
   Begin VB.Label lblSaldos 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "dfgsdgsfdg"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   400
      Left            =   4980
      TabIndex        =   11
      Top             =   6450
      Width           =   1845
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   675
      Left            =   4830
      TabIndex        =   9
      Top             =   270
      Width           =   2925
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Resumen de DIFERENCIAS DE CAJA DESDE EL 26-10-2006 AL 23-03-2007"
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   390
      TabIndex        =   8
      Top             =   2760
      Width           =   5895
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta:"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   4020
      TabIndex        =   7
      Top             =   1710
      Width           =   1035
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Desde:"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   4020
      TabIndex        =   6
      Top             =   1170
      Width           =   1035
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione Cuenta (Por Nombre o Número)"
      ForeColor       =   &H00E0E0E0&
      Height          =   405
      Left            =   240
      TabIndex        =   5
      Top             =   180
      Width           =   4275
   End
End
Attribute VB_Name = "frmMovCtasCont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ActuTbrB As Boolean
Dim FechaDesde As Date
Dim FechaHasta As Date 'estos son el 1er y ultimo registro

Dim Desde As Date
Dim Hasta As Date 'estos los que elige el usuario
Dim Saldo As Single, SaldoSub As Single

Private Sub cmdImprimir_Click()
    If lvResumen.ListItems.Count = 0 Then Exit Sub
    If tbrBuscador1.GetLstSel(1) = "" Then Exit Sub
    
    Dim Tit() As String

    TP.LineasSeparadoras = True
    
    ReDim Preserve Tit(4)
    Tit(4) = "Movimientos " + UCase(tbrBuscador1.GetLstSel(1))
    'datos de mi empresa!!!!!!!!!!!!!!
    Tit(0) = DB.GetValInRS("Clientes", "Nombre", "ID = -2", True)
    Tit(1) = "Direccion: " + DB.GetValInRS("Clientes", "Direccion", "ID = -2", True)
    Tit(2) = "Teléfono: " + DB.GetValInRS("Clientes", "Telefono", "ID = -2", True)
    Tit(3) = "Mail: " + DB.GetValInRS("Clientes", "Mail", "ID = -2", True)
    
    TP.ImprimirlvW lvResumen, Tit, "Fecha|NroAs|Detalle|Importe", _
        "Suman " + lblMas + ", Restan " + lblmenos + " (" + lblComent + ")", , 1.25, _
        "Desde el " + CStr(txtDesde) + " Hasta el " + CStr(txtHasta), _
        "Saldo al " + CStr(Hasta) + ": " + lblSaldos
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
    If KeyCode = vbKeyReturn Then cmdVer_Click
End Sub

Private Sub Form_Load()
    ActuTbrB = False
    
    tbrBuscador1.Contrasena = "zuliani"
    tbrBuscador1.ArchivoMDB = CFGBD.GetInfo(81, 4) + "Ctas.mdb"
    tbrBuscador1.SqlSinLike = "SELECT TOP 50 Id, Nombre FROM tblCuentas"
    tbrBuscador1.CampoEnQueBuscar = "Id/n,Nombre/b"
    tbrBuscador1.ColumnasSepPorComasyParentesis = "ID(700)/Nombre(2650)"
    
    FechaDesde = CDate(PC.GetTop1Rs("LibroDiario", "Fecha", "ASC", , True))
    FechaHasta = CDate(PC.GetTop1Rs("LibroDiario", "Fecha", , , True))
    
    Desde = FechaDesde
    Hasta = FechaHasta
    
    lblSaldo = ""
    lblComent = ""
    lblMov = ""
    lblTitulo = ""
    lblMas = FormatCurrency(0)
    lblmenos = FormatCurrency(0)
    lblSaldos = FormatCurrency(0)
    lblSaldoCon = FormatCurrency(0)
    
    VerDaTos Desde, Hasta
    
    lblRegDesde.Caption = "Tiene Registros desde: " + CStr(FechaDesde)
    txtDesde = CStr(FechaDesde)
    txtHasta = CStr(FechaHasta)
End Sub
    
Private Sub VerDaTos(Dde As Date, Hta As Date)
    If tbrBuscador1.GetLstSel = "" Then Exit Sub
    Dim IDC As Long, Mas As Single, Menos As Single
    Dim TmP As Single, TypC As Long, I As Long
    
    IDC = CLng(tbrBuscador1.GetLstSel(0))
    
    Saldo = PC.ABSValor(IDC, PC.ListarMovCuenta(IDC, Dde, Hta, lvResumen, , , True))
    SaldoSub = 0
    
    'tengo que agregar un renglon con lo del subdiario
    SubD = PC.ABSValor(IDC, PC.GetSaldo(IDC, "LibroSubDiario"))
    tmp2 = lvResumen.ListItems.Count + 1
    Saldo = Saldo + SubD

    lvResumen.ListItems.Add tmp2
    
    lvResumen.ListItems(tmp2).Text = FormatDateTime(Date, vbShortDate)
    lvResumen.ListItems(tmp2).SubItems(1) = ""
    lvResumen.ListItems(tmp2).SubItems(2) = "Saldo a la Fecha SubDiario Compras-Ventas"
    lvResumen.ListItems(tmp2).SubItems(3) = FormatCurrency(SubD, , , , vbFalse)
    
    'totales de columas ---------------------------------------------------------
    Mas = CSng(SumarColumnaLVW(lvResumen, 3, 1))
    Menos = CSng(SumarColumnaLVW(lvResumen, 3, -1))
    TmP = CSng(lvResumen.ListItems(1).SubItems(3)) 'saldo inicial
    
    'resto el saldo inicial
    If TmP < 0 Then 'empezo con saldo negativo lo resto de suma menos
        Menos = Menos - TmP
    Else 'es positivo o cero (si es cero no cambia nada)
        Mas = Mas - TmP
    End If
    
    'si es resultado saco los cierres de resultados
    TypC = PC.TipoCuenta(IDC)
    If TypC = 3 Or TypC = 4 Then
        For I = 1 To lvResumen.ListItems.Count
            If Left(lvResumen.ListItems(I).SubItems(2), 21) = "Cierre Resultados al " Then
                TmP = CSng(lvResumen.ListItems(I).SubItems(3))
                If TmP < 0 Then 'es negativo lo saco de menos
                    Menos = Menos - TmP
                Else
                    Mas = Mas - TmP
                End If
            End If
        Next I
        lblComent = "No se tienen en cuenta los Cierre de Resultados"
    Else
        lblComent = "Sumas sin tener en cuenta el Saldo Inicial"
    End If
            
    lblMas = FormatCurrency(Mas, , , , vbFalse)
    lblmenos = FormatCurrency(Menos, , , , vbFalse)
    ' ----------------------------------------------------------------------------
    
    
    'totales
    lblSaldos = FormatCurrency(Saldo, , , , vbFalse)
    lblSaldo = "Saldo al " + CStr(Hta)
    lblTitulo = "Resumen de " + UCase(PC.GetNameCuenta(IDC)) + _
        " Desde el " + CStr(Dde) + " AL " + CStr(Hta)
    Label4 = "Subcuentas de " + UCase(PC.GetNameCuenta(IDC))
        
    VerSub IDC
End Sub
    
Private Sub VerSub(IDC As Long)
    Dim Ctas() As String, A As Long
    Dim Nivel As Long, Sumo As Single
    
    lstSub.Clear
    Nivel = 2: Sumo = 0
    
    lstSub.AddItem "." + CStr(IDC) + "." + PC.GetNameCuenta(IDC)
    
    Ctas = PC.GetCuentas(IDC)
    
    If UBound(Ctas) = 0 Then
        lstSub.AddItem "¡No tiene subcuentas!"
        lblSaldoCon = lblSaldos
        Exit Sub
    End If
    
    For A = 1 To UBound(Ctas)
        Sumo = PC.ABSValor(CLng(Ctas(A)), PC.GetSaldoHasta(CLng(Ctas(A)), Hasta, , True))
        'ponerle los mismos espacios que tenia mas 3
        lstSub.AddItem String(Nivel * 3, " ") + "." + Ctas(A) + "." + _
            PC.GetNameCuenta(CLng(Ctas(A))) + ": " + _
            FormatCurrency(Sumo, , , , vbFalse)
        AnadirHijos CLng(Ctas(A)), Nivel
        SaldoSub = SaldoSub + Sumo
    Next A
    
    lblSaldoCon = FormatCurrency(CSng(lblSaldos) + SaldoSub, , , , vbFalse)
End Sub

Private Sub AnadirHijos(IdCta As Long, Nivel As Long)
    Dim Hij() As String, Niv As Long, A As Long, Sumo As Single
    
    Hij = PC.GetCuentas(IdCta)
    Niv = Nivel + 1
    
    If UBound(Hij) > 0 Then
        For A = 1 To UBound(Hij)
            Sumo = PC.ABSValor(CLng(Hij(A)), PC.GetSaldoHasta(CLng(Hij(A)), Hasta, , True))
            lstSub.AddItem String(Niv * 3, " ") + "." + Hij(A) + "." + _
                PC.GetNameCuenta(CLng(Hij(A))) + ": " + _
                FormatCurrency(Sumo, , , , vbFalse)
            AnadirHijos CLng(Hij(A)), Niv
            SaldoSub = SaldoSub + Sumo
        Next A
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    tbrBuscador1.CN_CLOSE
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
    
Private Sub tbrBuscador1_Change()
    If IsNumeric(tbrBuscador1.Text) Then
        tbrBuscador1.CampoEnQueBuscar = "ID/b,Nombre"
    Else
        tbrBuscador1.CampoEnQueBuscar = "ID/n,Nombre/b"
    End If
    
    If tbrBuscador1.Text <> "" Then VerActu
End Sub

Private Sub VerActu()
    If ActuTbrB = False Then
       ActuTbrB = True
       tbrBuscador1.Recargar
    Else
        ActuTbrB = False
    End If

End Sub

