VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A7FBD38D-2930-49E3-B60C-9E0202D84549}#15.0#0"; "tbrControles.ocx"
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmEgresos 
   BackColor       =   &H00544B45&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Egresos"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10020
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEgresos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   10020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin tbrFaroButton.fBoton cmdAgregar 
      Height          =   465
      Left            =   8370
      TabIndex        =   15
      Top             =   960
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Agregar"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   9
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdEliminar 
      Height          =   465
      Left            =   8370
      TabIndex        =   16
      Top             =   2250
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Eliminar"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   9
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdModificar 
      Height          =   465
      Left            =   8370
      TabIndex        =   17
      Top             =   1590
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Modificar"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   9
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton Command2 
      Height          =   405
      Left            =   8700
      TabIndex        =   14
      Top             =   7200
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   9
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdPagar 
      Height          =   405
      Left            =   8760
      TabIndex        =   3
      Top             =   5250
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Pagar"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   9
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdGrabarDetalle 
      Height          =   405
      Left            =   7290
      TabIndex        =   13
      Top             =   4320
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Grabar"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   9
      fECol           =   5717301
   End
   Begin VB.TextBox txtDetPago 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6870
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   6000
      Width           =   3015
   End
   Begin tbrControles.MouTextBox txtMonto 
      Height          =   435
      Left            =   7410
      TabIndex        =   1
      Top             =   5220
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   767
      BackColor       =   16777215
      Enabled         =   -1  'True
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
   Begin VB.Frame Frame6 
      BackColor       =   &H00544B45&
      Caption         =   "Egresos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2685
      Left            =   90
      TabIndex        =   7
      Top             =   4920
      Width           =   6735
      Begin tbrFaroButton.fBoton cmdImprimir 
         Height          =   435
         Left            =   5130
         TabIndex        =   12
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   767
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Imprimir"
         fEnabled        =   -1  'True
         fFontN          =   "Arial"
         fFontS          =   9
         fECol           =   5717301
      End
      Begin MSComctlLib.ListView lvEgresos 
         Height          =   1785
         Left            =   120
         TabIndex        =   9
         Top             =   780
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   3149
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
            SubItemIndex    =   1
            Text            =   "IDAs"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cuenta"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Detalle"
            Object.Width           =   3969
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Monto"
            Object.Width           =   2293
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTFecha 
         Height          =   345
         Left            =   1530
         TabIndex        =   11
         Top             =   300
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   17104897
         CurrentDate     =   39197
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Egresos desde "
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Left            =   150
         TabIndex        =   8
         Top             =   360
         Width           =   1485
      End
   End
   Begin VB.TextBox txtDescripcion 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   180
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   4170
      Width           =   6945
   End
   Begin VB.ListBox lstCuentas 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3210
      Left            =   150
      TabIndex        =   0
      Top             =   330
      Width           =   8085
   End
   Begin tbrFaroButton.fBoton cmdStats 
      Height          =   705
      Left            =   8370
      TabIndex        =   18
      Top             =   2850
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1244
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Ver Estadísticas"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   9
      fECol           =   5717301
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Detalle Pago"
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
      Left            =   6930
      TabIndex        =   10
      Top             =   5760
      Width           =   1005
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Monto"
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
      Left            =   6480
      TabIndex        =   6
      Top             =   5340
      Width           =   825
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Detalle"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   180
      TabIndex        =   5
      Top             =   3900
      Width           =   1155
   End
End
Attribute VB_Name = "frmEgresos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAgregar_Click()
    Dim NbCta As String, nCta As String
    
    If lstCuentas.ListIndex = -1 Then Exit Sub
    
    nCta = InputBox("Ingrese el número de la cuenta que desee agregar", _
        "Número de cuenta nueva", CStr(PC.GetUltIDMasUno))
    
    If Not IsNumeric(nCta) Then
        MsgBox "Número de cuenta incorrecto", vbInformation, "Atención"
        Exit Sub
    End If
    
    nCta = CStr(CLng(nCta))
    NbCta = Replace(InputBox("Ingrese el nombre de la cuenta", "Nombre"), ".", " ")
    
    If NbCta = "" Then Exit Sub
    
    'crear una cuenta en el nivel elegido
    Dim SP() As String
    SP = Split(lstCuentas, ".")
    
    Select Case PC.AgregarCuenta(CLng(nCta), CLng(SP(1)), NbCta)
        Case 1
            MsgBox "Existen datos vacíos", vbInformation, "Atención"
            Exit Sub
        Case 2
            MsgBox "Ya existe una cuenta con ese nombre", vbInformation, "Atención"
            Exit Sub
        Case 3
            MsgBox "Ya existe una cuenta con ese número", vbInformation, "Atención"
            Exit Sub
    End Select
    
    'para que se actualice medio negra la cosa
    If Right(lstCuentas, 1) = "*" Then
        lstCuentas_DblClick
        lstCuentas_DblClick
    Else
        lstCuentas_DblClick
    End If
End Sub

Private Sub cmdEliminar_Click()
    If lstCuentas.ListIndex = -1 Then Exit Sub
    
    Dim SP() As String
    
    SP = Split(lstCuentas, ".")
    
    If MsgBox("¿Está seguro de eliminar " + UCase(SP(2)) + vbCrLf + _
        " y sus correspondientes subcuentas?", _
        vbInformation + vbYesNo, "Atención") = vbNo Then Exit Sub
    
    PC.EliminarCuenta CLng(SP(1))
    
    CargarDatos
End Sub

Private Sub cmdGrabarDetalle_Click()
    If lstCuentas.ListIndex = -1 Then Exit Sub
    
    Dim SP() As String
    SP = Split(lstCuentas, ".")

    PC.ModificarCuenta CLng(SP(1)), , , txtDescripcion
    
End Sub

Private Sub cmdImprimir_Click()
    If lvEgresos.ListItems.Count = 0 Then Exit Sub
    
    Dim Tit() As String

    TP.LineasSeparadoras = True
    
    ReDim Preserve Tit(4)
    Tit(4) = "Egresos desde el " + FormatDateTime(DTFecha, vbShortDate)
    'datos de mi empresa!!!!!!!!!!!!!!
    Tit(0) = DB.GetValInRS("Clientes", "Nombre", "ID = -2", True)
    Tit(1) = "Direccion: " + DB.GetValInRS("Clientes", "Direccion", "ID = -2", True)
    Tit(2) = "Teléfono: " + DB.GetValInRS("Clientes", "Telefono", "ID = -2", True)
    Tit(3) = "Mail: " + DB.GetValInRS("Clientes", "Mail", "ID = -2", True)
    
    TP.ImprimirlvW lvEgresos, Tit, "Fecha|IDAs|Cuenta|Detalle|Importe", , , 1.35
End Sub

Private Sub cmdModificar_Click()
    Dim nCta As String, nUcTa As String, tmP As String
    Dim SP() As String
    
    If lstCuentas.ListIndex = -1 Then Exit Sub
    SP = Split(lstCuentas, ".")
    
    If Left(lstCuentas, 1) = "." Then
        MsgBox UCase(SP(2)) + " No puede ser modificada", vbInformation, "Atención"
        Exit Sub
    End If
    
    nUcTa = InputBox("Ingrese el número de la cuenta si lo desea modificar", _
        "Número de cuenta nueva", SP(1))
    
    If Not IsNumeric(nUcTa) Then
        MsgBox "Número de cuenta incorrecto", vbInformation, "Atención"
        Exit Sub
    End If
    
    nUcTa = CStr(CLng(nUcTa))
    
    Dim tmpExNombre As String
    'por los * que le pone a las cuentas abiertas
    If Right(SP(2), 1) = "*" Then
        tmpExNombre = Left(SP(2), Len(SP(2)) - 1)
        tmP = "*"
    Else
        tmpExNombre = SP(2)
        tmP = ""
    End If
    
    nCta = Replace(InputBox("Ingrese el nuevo nombre de la cuenta", "Nombre", tmpExNombre), ".", " ")
    
    If nCta = "" Then Exit Sub
    
    nCta = Replace(nCta, "'", " ")
    nCta = Replace(nCta, "*", " ")
    
    'modificar la cuenta
    Dim Modi As Long
    Modi = PC.ModificarCuenta(CLng(SP(1)), nUcTa, nCta, txtDescripcion)
    If Modi <> 0 Then
        MsgBox CStr(Modi) + " Se registraron errores en la modificación", vbInformation, "Atención"
        Exit Sub
    End If
    
    'actualizo
    lstCuentas.List(lstCuentas.ListIndex) = SP(0) + "." + CStr(nUcTa) + "." + nCta + tmP
End Sub




Private Sub cmdPagar_Click()
    'primero valido
    txtMonto = FormatCurrency(ValidarNumeros(txtMonto), , , , vbFalse)
    If CSng(txtMonto) = 0 Then Exit Sub
    
    If lstCuentas.ListIndex = -1 Then
        MsgBox "No seleccionó ninguna cuenta", vbInformation, "Atención"
        Exit Sub
    End If
    
    'segundo veo el nombre de cta
    Dim nCuenta As String, SP() As String
    
    SP = Split(lstCuentas, ".")
    If Right(SP(2), 1) = "*" Then
        nCuenta = Left(SP(2), Len(SP(2)) - 1)
    Else
        nCuenta = SP(2)
    End If
    
    'le pregunto!
    If MsgBox("Está a punto de registrar el gasto en " + UCase(nCuenta) + _
        " por " + txtMonto + " ¿Son Correctos los datos?", vbInformation + vbOKCancel, _
        "Atención") = vbCancel Then Exit Sub
    
    'ahora si le anoto
    'solo va el asiento (cta) a caja
    PC.Asiento CStr(PC.GetIDCuenta(nCuenta)), txtMonto, "78", txtMonto, _
        , "(EG) " + txtDetPago
    
    If CFG.GetInfo(95, 4) = "Si" Then frmPago.AbrirDatos CSng(txtMonto), False, "(EG) " + txtDetPago, "LibroDiario"
    
    PC.ListarEgresos lvEgresos, 0, DTFecha
    txtMonto = FormatCurrency(0)
    txtDetPago = ""
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub DTFecha_Change()
    PC.ListarEgresos lvEgresos, 0, DTFecha
End Sub

Private Sub fBoton1_Click()

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    DTFecha = Date - 7
    CargarDatos
    
    txtMonto = FormatCurrency(0)
End Sub

Private Sub CargarDatos()
    lstCuentas.Clear
    
    lstCuentas.AddItem ".19.Gastos de Comercializacion"
    lstCuentas.AddItem ".20.Gastos de Administracion"
    lstCuentas.AddItem ".37.Gastos de Financiacion"
    lstCuentas.AddItem ".33.Impuestos"
    lstCuentas.AddItem ".34.Servicios"
    lstCuentas.AddItem ".22.Otros Egresos"
    
    FormatearMouTextBox frmEgresos
    PC.ListarEgresos lvEgresos, 0, DTFecha
End Sub

Private Sub lstCuentas_Click()
    If lstCuentas.ListIndex = -1 Then Exit Sub
    
    Dim SP() As String
    SP = Split(lstCuentas, ".")
    
    txtDescripcion = PC.GetDetalle(CLng(SP(1)))
End Sub

Private Sub lstCuentas_DblClick()
    'mostrar los subniveles del elegido
    Dim tmpIx As Long
    tmpIx = lstCuentas.ListIndex
    
    Dim CTA As Long
    Dim SP() As String
    SP = Split(lstCuentas, ".")
    CTA = CLng(SP(1))
    Dim Ctas() As String
    Ctas = PC.GetCuentas(CTA)
    If UBound(Ctas) = 0 Then
        MsgBox "¡No tiene subcuentas!"
        Exit Sub
    End If
    
    'veo si ya mostro las subcuentas, si es asi las escondo
    If Right(lstCuentas, 1) = "*" Then
        'primero le borro el asterisco
        lstCuentas.List(tmpIx) = Left(lstCuentas.List(tmpIx), _
            Len(lstCuentas.List(tmpIx)) - 1)
        
        'escondo las subcuentas
        Dim Niv As Long, I As Long
        Niv = nNivel(tmpIx)
        
        I = tmpIx + 1
        
        Do While Not Niv >= nNivel(I)
            lstCuentas.RemoveItem I
            
            If I > lstCuentas.ListCount - 1 Then Exit Do
            'se baja por la eliminacion
            'I = I + 1
        Loop
        
        Exit Sub
    End If
    
    Dim A As Long
    For A = 1 To UBound(Ctas)
        'ponerle los mismos espacios que tenia mas 3
        lstCuentas.AddItem "   " + SP(0) + "." + Ctas(A) + "." + _
            PC.GetNameCuenta(CLng(Ctas(A))), lstCuentas.ListIndex + 1
        
    Next A
    
    'marco que ya lo abrio
    lstCuentas.List(tmpIx) = lstCuentas.List(tmpIx) + "*"
End Sub

Private Function nNivel(IndiceLista As Long) As Long
    'veo el nivel de la cuenta seleccionada del listbox

    If IndiceLista = -1 Then
        nNivel = -1
        Exit Function
    End If
    
    Dim Spp() As String
    Spp = Split(lstCuentas.List(IndiceLista), ".")
    'spp(0) tiene los espacios que me van a decir en que nivel esta
    If Len(Spp(0)) = 0 Then  'es el nivel1
        nNivel = 1
    Else 'hago la formula, tiene que dar un numero redondo
        nNivel = Round(Len(Spp(0)) / 3 + 1, 0)
    End If
    
End Function

Private Sub Text1_Change()

End Sub

Private Sub txtMonto_GotFocus()
    PintarTxt txtMonto
End Sub

Private Sub txtMonto_LostFocus()
    txtMonto = FormatCurrency(ValidarNumeros(txtMonto), , , , vbFalse)
End Sub
