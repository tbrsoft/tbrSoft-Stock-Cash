VERSION 5.00
Object = "{A7FBD38D-2930-49E3-B60C-9E0202D84549}#14.0#0"; "tbrControles.ocx"
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmEstRes 
   BackColor       =   &H00544B45&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estado de Resultados"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEstRes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin tbrFaroButton.fBoton cmdAgregarO 
      Height          =   435
      Left            =   7740
      TabIndex        =   5
      Top             =   3840
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   ">>>"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton command1 
      Height          =   435
      Left            =   10020
      TabIndex        =   7
      Top             =   6420
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdOk 
      Height          =   435
      Left            =   6720
      TabIndex        =   6
      Top             =   6030
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Cerrar Resultados"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdImprimir 
      Height          =   435
      Left            =   1770
      TabIndex        =   24
      Top             =   6030
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Imprimir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00544B45&
      Caption         =   "Forma de Pago"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1275
      Left            =   6090
      TabIndex        =   22
      Top             =   1620
      Width           =   2475
      Begin VB.OptionButton chkElige 
         BackColor       =   &H00544B45&
         Caption         =   "Seleccionar Aportantes"
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
         Left            =   120
         TabIndex        =   23
         Top             =   900
         Width           =   2295
      End
      Begin VB.OptionButton chkIgual 
         BackColor       =   &H00544B45&
         Caption         =   "Partes Iguales"
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
         Left            =   120
         TabIndex        =   3
         Top             =   570
         Width           =   2200
      End
      Begin VB.OptionButton chkSPart 
         BackColor       =   &H00544B45&
         Caption         =   "Según Participación"
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
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   2200
      End
   End
   Begin MSComctlLib.ListView lvRNA 
      Height          =   3705
      Left            =   150
      TabIndex        =   19
      Top             =   1410
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   6535
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
         Text            =   "IDCta"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cuenta"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Debe"
         Object.Width           =   2328
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Haber"
         Object.Width           =   2328
      EndProperty
   End
   Begin tbrControles.MouTextBox txtPesos 
      Height          =   405
      Left            =   6150
      TabIndex        =   4
      Top             =   4140
      Width           =   1365
      _ExtentX        =   2408
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
   Begin VB.ComboBox cmbNombre 
      Height          =   330
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   3270
      Width           =   2505
   End
   Begin VB.ComboBox cmbCuenta 
      Height          =   330
      ItemData        =   "frmEstRes.frx":6852
      Left            =   6090
      List            =   "frmEstRes.frx":685C
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1110
      Width           =   2505
   End
   Begin VB.ListBox lstResumen 
      Height          =   2160
      Left            =   8700
      TabIndex        =   12
      Top             =   1680
      Width           =   2955
   End
   Begin tbrFaroButton.fBoton cmdQuitar 
      Height          =   435
      Left            =   7740
      TabIndex        =   25
      Top             =   4320
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "<<<"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.Label lblGB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Resultados Para Distribuir Actuales"
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
      Height          =   315
      Left            =   150
      TabIndex        =   21
      Top             =   6630
      Width           =   3405
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Asiento de Cierre"
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
      Index           =   3
      Left            =   360
      TabIndex        =   20
      Top             =   1140
      Width           =   4905
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
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
      Left            =   6120
      TabIndex        =   18
      Top             =   3000
      Width           =   705
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Resultados pendientes de Distribuir"
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
      Height          =   315
      Index           =   2
      Left            =   6030
      TabIndex        =   17
      Top             =   5310
      Width           =   3405
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Monto"
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
      Height          =   315
      Left            =   6030
      TabIndex        =   16
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta"
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
      Height          =   315
      Left            =   6120
      TabIndex        =   15
      Top             =   870
      Width           =   615
   End
   Begin VB.Label lblSaldo 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "$222,22"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   9630
      TabIndex        =   14
      Top             =   5190
      Width           =   2115
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Resumen"
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
      Left            =   8790
      TabIndex        =   13
      Top             =   1470
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Resultados Para Distribuir Actuales"
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
      Height          =   315
      Index           =   1
      Left            =   180
      TabIndex        =   11
      Top             =   5460
      Width           =   3405
   End
   Begin VB.Label lblRNA 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "121212"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Index           =   1
      Left            =   3690
      TabIndex        =   10
      Top             =   5400
      Width           =   2070
   End
   Begin VB.Label lblRNA 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "121212"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Index           =   0
      Left            =   3750
      TabIndex        =   9
      Top             =   510
      Width           =   2010
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Resultados Anteriores no Distribuidos"
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
      Height          =   315
      Index           =   0
      Left            =   90
      TabIndex        =   8
      Top             =   570
      Width           =   3645
   End
End
Attribute VB_Name = "frmEstRes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ventas As Single, CtoVentas As Single, Otros As Single, ResAcum As Single, _
    UNeta As Single, Saldo As Single
Dim IDcierre As Long, AsientoCierre() As String

Private Sub chkElige_Click()
    Socios = PC.GetCuentas(52)
    For jj = 1 To UBound(Socios)
        cmbNombre.AddItem PC.GetNameCuenta(CLng(Socios(jj)))
    Next jj
    
    cmbNombre.ListIndex = 0
End Sub

Private Sub chkIgual_Click()
    cmbNombre.Clear
End Sub

Private Sub chkSPart_Click()
    cmbNombre.Clear
End Sub

Private Sub cmbCuenta_Click()
    If cmbCuenta = "Distribucion Socios" Then
        Dim Socios() As String, jj As Long
        
        Socios = PC.GetCuentas(52)
        'cargo combo pero asi
        cmbNombre.Clear
        For jj = 1 To UBound(Socios)
            cmbNombre.AddItem PC.GetNameCuenta(CLng(Socios(jj)))
        Next jj
        cmbNombre.ListIndex = 0
        Frame1.Visible = False
    Else
        Frame1.Visible = True
        If chkElige Then
            For jj = 1 To UBound(Socios)
                cmbNombre.AddItem PC.GetNameCuenta(CLng(Socios(jj)))
            Next jj
        Else
            cmbNombre.Clear
        End If
    End If
End Sub

Private Sub cmdAgregarO_Click()
    If CSng(ValidarNumeros(txtPesos)) = 0 Then Exit Sub
    
     'si llego aca es que txt es numerico
    Saldo = Saldo - Round(CSng(txtPesos), 2)
    
    Dim strDet As String
    If cmbNombre = "" Then 'Es Capitalizar (o Partes Iguales (PI) o S/Partic(SP))
        strDet = cmbCuenta
        If chkIgual Then strDet = strDet + " PI"
        If chkSPart Then strDet = strDet + " SP"
    Else
        'Es para cuentas particulares o como APORTE particular
        If Frame1.Visible = True Then
            strDet = cmbCuenta + " - " + cmbNombre
        Else
            strDet = cmbNombre
        End If
    End If
    
    lstResumen.AddItem FormatCurrency(txtPesos, , , , vbFalse) + "\" + strDet
        
    lblSaldo = FormatCurrency(Saldo)
    txtPesos = FormatCurrency(0)
End Sub

Private Sub cmdImprimir_Click()
    Dim Tit() As String

    TP.LineasSeparadoras = True
    
    ReDim Preserve Tit(4)
    Tit(4) = "Asiente Cierre Resultados " + FormatDateTime(Date, vbShortDate)
    'datos de mi empresa!!!!!!!!!!!!!!
    Tit(0) = DB.GetValInRS("Clientes", "Nombre", "ID = -2", True)
    Tit(1) = "Direccion: " + DB.GetValInRS("Clientes", "Direccion", "ID = -2", True)
    Tit(2) = "Teléfono: " + DB.GetValInRS("Clientes", "Telefono", "ID = -2", True)
    Tit(3) = "Mail: " + DB.GetValInRS("Clientes", "Mail", "ID = -2", True)
    
    TP.ImprimirlvW lvRNA, Tit, "IDCta|Cuenta|Debe|Haber", _
        "Resultados Sin Distribuir Actuales: " + lblRNA(1), _
        , , "Resultados Anteriores Sin Distribuir: " + lblRNA(0)
End Sub

Private Sub cmdOK_Click()
    If MsgBox("¿Los datos son correctos?, presione Aceptar para continuar", vbInformation + _
        vbOKCancel, "Atencion") = vbCancel Then Exit Sub
    
    Dim Ctas() As String, Hij() As String, I As Long, TmP As Single
    Dim Socios As Single, Capitalizar As Single, SP() As String, Ktas() As Single
    Socios = 0: Capitalizar = 0: TmP = 0
    
    'ktas() tiene las participaciones, la saco aca para que no este afectada por cada
    'modificacion que vaya haciendo en este procedimiento
    
    Dim jj As Long
    For jj = 0 To lstResumen.ListCount - 1
        Ctas = Split(lstResumen.List(jj), "\")
        If Left(Ctas(1), 11) = "Capitalizar" Then 'es para capitalizar
            Hij = PC.GetCuentas(52)
            Select Case Right(Ctas(1), 2)
                Case "PI"
                    For I = 1 To UBound(Hij)
                        Aporte PC.GetNameCuenta(CLng(Hij(I))), _
                            CSng(Ctas(0)) / UBound(Hij), _
                            "Aporte Distribución Ganancias", 16
                    Next I
                Case "SP"
                    For I = 1 To UBound(Hij)
                        ReDim Preserve Ktas(I)
                        Ktas(I) = GetParticipacion(PC.GetNameCuenta(CLng(Hij(I))))
                    Next I
                    
                    For I = 1 To UBound(Hij)
                        Aporte PC.GetNameCuenta(CLng(Hij(I))), _
                            CSng(Ctas(0)) * Ktas(I), _
                            "Aporte Distribución Ganancias", 16
                        TmP = TmP + Ktas(I) * CSng(Ctas(0))
                    Next I
                    
                    'no deberia pasar pero por las dudas
                    If Abs(TmP - CSng(Ctas(0))) > 1 Then
                        PC.Asiento "16", CStr(CSng(Ctas(0) - TmP)), _
                            "15", CStr(CSng(Ctas(0) - TmP)), , "Ajuste"
                    End If
                    
                Case Else 'depende de quien aporta
                    'esta separado por una barra " - "
                    SP = Split(Ctas(1), " - ")
                    Aporte SP(1), CSng(Ctas(0)), "Aporte Distribución Resultados", 16
            End Select
            
            Capitalizar = Capitalizar + CSng(Ctas(0))
            
        Else 'es para un socio
            PC.Asiento "16", CStr(CSng(Ctas(0))), _
                PC.GetIDCuenta(Ctas(1)), CStr(CSng(Ctas(0)))
            'ademas va en movsocyEmp
            Dim IdCz As Long
            IdCz = PC.GetIDCuenta(Ctas(1))
            
            DB.EXECUTE "INSERT INTO MovSocyEmp (ID,Fecha, IdNivel3," + _
                "Tipo,Variacion,Detalle) VALUES (" + IdAutonum("MovSocyEmp") + _
                ",#" + stFechaSQL(Date) + _
                "#," + CStr(IdCz) + ",'Socio'," + _
                Replace(CStr(CSng(Ctas(0))), ",", ".") + _
                ",'De Distribucion de Ganancias')"
                
            Socios = Socios + CSng(Ctas(0))
        End If
    Next jj
    
    'registro la distribucion
    PC.Distribuir IDcierre, Socios, Capitalizar
    
    'registro todo como cerrado resultados
    'cierro o sea recien aca se hace el asientoRNA
    PC.CerrarResultados
    
    MsgBox "Se realizó el registro correctamente Quedando " + _
        FormatCurrency(Saldo) + " en Resultados Acumulados para el próximo cierre", _
        vbInformation, "Cierre Resultados"
    Unload Me
End Sub

Private Sub cmdQuitar_Click()
    Dim sRest As String
    
    If lstResumen.ListIndex < 0 Then Exit Sub
    sRest = Left$(lstResumen, InStr(1, lstResumen, "\") - 1)
    
    Saldo = Saldo + Round(CSng(sRest), 2)
    lblSaldo = FormatCurrency(Saldo)
    lstResumen.RemoveItem (lstResumen.ListIndex)
End Sub

Private Sub Command1_Click()
    If MsgBox("¿Está seguro que desea salir sin realizar la distribución " + _
        "de ganancias? Quedarán " + _
        FormatCurrency(Saldo) + " en Resultados Acumulados para el próximo cierre", _
        vbInformation + vbOKCancel, "Atencion") = vbCancel Then Exit Sub
    
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()

    On Local Error GoTo errRES

    Terr.AnOtaR "EstRes-001"
    Dim IDcierre As Long
    
    cmbCuenta.ListIndex = 0
    UNeta = 0
    txtPesos = FormatCurrency(0)
    
    Terr.AnOtaR "EstRes-002"
    AsientoCierre = PC.AsientoRNA
    
    Terr.AnOtaR "EstRes-003"
    IDcierre = PC.GetUltIdCierreResultados
    
    Terr.AnOtaR "EstRes-004", IDcierre
    'PC.ListarAsientos IDcierre, IDcierre, lvRNA
    'El asiento no se contabilizó aca va a mostrar como quedaría
    UNeta = LlenarAsiento(AsientoCierre)
    
    Terr.AnOtaR "EstRes-005", UNeta
    lblGB = "Ganancia Bruta: " + _
        FormatCurrency(PC.ABSSumarconSubcuentas(17) - _
            PC.ABSSumarconSubcuentas(18))
    
    Terr.AnOtaR "EstRes-006", lblGB.Caption
    ResAcum = PC.ABSSumarconSubcuentas(16)
    lblRNA(0) = FormatCurrency(ResAcum)
    
    Terr.AnOtaR "EstRes-007", lblRNA(0).Caption
    Saldo = ResAcum + UNeta
    
    Terr.AnOtaR "EstRes-008", Saldo
    lblRNA(1) = FormatCurrency(Saldo)
    lblSaldo = FormatCurrency(Saldo)
    
    Terr.AnOtaR "EstRes-009"
    
    Exit Sub
    
errRES:
    Terr.AppendLog "", Terr.ErrToTXT(Err)
    Resume Next
End Sub

Private Function LlenarAsiento(SpAsiento() As String) As Single
    'da como resultado REj
    Dim I As Long, Asi() As String, SP() As String, Resp As Single
    
    Resp = 0
    Asi = SpAsiento
    
    For I = 1 To UBound(Asi)
        SP = Split(Asi(I), " | ")
        
        lvRNA.ListItems.Add I
        
        lvRNA.ListItems(I).Text = SP(0)
        lvRNA.ListItems(I).SubItems(1) = PC.GetNameCuenta(CLng(SP(0)))
        If CLng(SP(0)) = "16" Then Resp = -CSng(SP(1))
        
        If CLng(SP(1)) < 0 Then
            lvRNA.ListItems(I).SubItems(3) = FormatCurrency(-CSng(SP(1)), , , , vbFalse)
        Else
            lvRNA.ListItems(I).SubItems(2) = FormatCurrency(CSng(SP(1)), , , , vbFalse)
        End If

    Next I
    
    LlenarAsiento = Resp
End Function

Private Sub lblSaldo_Change()
    If Saldo >= 0 Then
        Label1(2) = "Ganancias sin distribuir"
    Else
        Label1(2) = "Perdidas sin distribuir"
    End If
End Sub

Private Sub txtPesos_Change()
    cmdAgregarO.Default = True
End Sub

Private Sub txtPesos_GotFocus()
    PintarTxt txtPesos
End Sub

Private Sub txtPesos_LostFocus()
    txtPesos = FormatCurrency(ValidarNumeros(txtPesos), , , , vbFalse)
End Sub
