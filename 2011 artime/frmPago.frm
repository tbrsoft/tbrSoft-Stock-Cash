VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{A7FBD38D-2930-49E3-B60C-9E0202D84549}#13.0#0"; "tbrControles.ocx"
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmPago 
   BackColor       =   &H00544B45&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Forma de Pago"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPago.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin tbrFaroButton.fBoton cmdAdd 
      Height          =   465
      Left            =   3720
      TabIndex        =   2
      Top             =   2040
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Agregar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrControles.MouTextBox txtMonto 
      Height          =   375
      Left            =   2070
      TabIndex        =   1
      Top             =   2250
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   661
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
   Begin MSComctlLib.ListView lvResumen 
      Height          =   2355
      Left            =   1110
      TabIndex        =   8
      Top             =   2970
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4154
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Forma"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Monto"
         Object.Width           =   2028
      EndProperty
   End
   Begin VB.ComboBox cmbForma 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1110
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1800
      Width           =   2415
   End
   Begin tbrFaroButton.fBoton cmdOK 
      Height          =   465
      Left            =   1560
      TabIndex        =   3
      Top             =   6300
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Aceptar (F1)"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdSalir 
      Height          =   465
      Left            =   3450
      TabIndex        =   4
      Top             =   6300
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Presione Escape en caso de realizarse todo en Efvo"
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
      Left            =   210
      TabIndex        =   11
      Top             =   90
      Width           =   4515
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
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
      Height          =   255
      Left            =   1230
      TabIndex        =   10
      Top             =   2340
      Width           =   825
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione Forma de Pago, el Monto y Presione Agregar"
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
      Left            =   930
      TabIndex        =   9
      Top             =   1470
      Width           =   5355
   End
   Begin VB.Label lblQueQuedo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Monto a Cobrar $ 1.233.333.34"
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
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   5460
      Width           =   4545
   End
   Begin VB.Label lblQue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Monto a Cobrar a Carreofur Argentina $ 1.233.333.34"
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
      Height          =   375
      Left            =   210
      TabIndex        =   6
      Top             =   900
      Width           =   5925
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Forma de Pago"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   1440
      TabIndex        =   5
      Top             =   360
      Width           =   3615
   End
End
Attribute VB_Name = "frmPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cuanto As Single, QueHace As String ' (es "Cobrar" o "Pagar")
Dim DetalleAsiento As String
' o es para cobrar o para pagar aca no puedo anotar
' a clientes ni a proveedores!! que se cague
Dim Formas() As String ' de esta forma "nrocta.NombreCuenta"
Dim Ix As Long, Saldo As Single
Dim mLibro As String
    
Private Sub cmdAdd_Click()
    txtMonto = FormatCurrency(ValidarNumeros(txtMonto), , , , vbFalse)
    
    If CSng(txtMonto) = 0 Then txtMonto.SetFocus: Exit Sub
    
    Dim TmP As Long
    
    TmP = lvResumen.ListItems.Count + 1
    
    lvResumen.ListItems.Add TmP
    
    lvResumen.ListItems(TmP).Text = cmbForma
    lvResumen.ListItems(TmP).SubItems(1) = FormatCurrency(CSng(txtMonto))
    
    cmbForma.SetFocus
    txtMonto = FormatCurrency(0)
    CalcularSumas
End Sub

Private Sub cmdOK_Click()
    If EsCero(Saldo) = False Then
        If MsgBox("¿Desea dejar la diferencia como Caja?", vbInformation + vbYesNo) = vbNo Then Exit Sub
    End If
    
    If Saldo = Cuanto Then 'no hace nada pongo todo como caja
        Unload Me
        Exit Sub
    End If
    
    'ahora si todo lo que agrego va contra caja
    Dim Debe As Single, T As Long, IdCta As Long, TmP As String
    
    For T = 1 To lvResumen.ListItems.Count
        'como predeterminado es que estoy cobrando
        Debe = CSng(txtInLvW(lvResumen, T, 1))
        
        If QueHace <> "Cobrar" Then
            Debe = -Debe
        End If
        
        TmP = txtInLvW(lvResumen, T, 0)
        IdCta = CLng(Left(TmP, InStr(1, TmP, ".") - 1))
        
        If IdCta <> 78 Then 'si es caja no hago nada ya esta en caja
            PC.Asiento CStr(IdCta), CStr(Debe), "78", CStr(Debe), _
                mLibro, DetalleAsiento
        End If
    Next T
    
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Public Sub AbrirDatos(Monto As Single, Optional ParaCobrar As Boolean = True, _
    Optional DetAsi As String = "", Optional Libro As String = "LibroSubDiario")
    
    Cuanto = Monto
    mLibro = Libro
    
    If ParaCobrar Then
        QueHace = "Cobrar"
    Else
        QueHace = "Pagar"
    End If
        
    DetalleAsiento = DetAsi
    Me.Show 1
End Sub

Private Sub fBoton2_Click()

End Sub

Private Sub fBoton1_Click()

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then cmdOK_Click
    If KeyCode = vbKeyEscape Then Unload Me
End Sub


Private Sub Form_Load()
    Limpiar
    lblQue = "Monto a " + QueHace + ": " + FormatCurrency(Cuanto)
    CargarFormas
    CalcularSumas
End Sub

Private Sub Limpiar()
    lvResumen.ListItems.Clear
    txtMonto = FormatCurrency(0)
    
End Sub

Private Sub txtMonto_GotFocus()
    PintarTxt txtMonto
End Sub

Private Sub txtMonto_LostFocus()
    txtMonto = FormatCurrency(ValidarNumeros(txtMonto), , , , vbFalse)
End Sub

Private Sub CargarFormas()
    ' son todas las subcuentas y hijos de Caja y Bancos
    Dim S As Long
    
    cmbForma.Clear
    Ix = 0
    ReDim Preserve Formas(Ix)
    Formas(Ix) = "Nada"
    
    AgregarHijos 6 'es caja y bancos que excluyo por ser el padre
    
    For S = 1 To Ix
        cmbForma.AddItem Formas(S)
    Next S
    
    cmbForma.ListIndex = 0
End Sub

Private Sub AgregarHijos(IdCta As Long)
    'primero que todo agrega la cuenta en cuestion
    Dim Hijos() As String, R As Long
    
    If IdCta <> 6 Then 'es caja y bancos que excluyo por ser el padre
        Ix = Ix + 1
        ReDim Preserve Formas(Ix)
        Formas(Ix) = CStr(IdCta) + "." + PC.GetNameCuenta(IdCta)
    End If
    
    Hijos = PC.GetCuentas(IdCta)
    
    For R = 1 To UBound(Hijos)
        AgregarHijos CLng(Hijos(R))
    Next R
    
End Sub

Private Sub CalcularSumas()
    Dim Q As Long
    
    Saldo = 0
    If lvResumen.ListItems.Count > 0 Then
        For Q = 1 To lvResumen.ListItems.Count
            Saldo = Saldo + CSng(txtInLvW(lvResumen, Q, 1))
        Next Q
    End If
    
    Saldo = Cuanto - Saldo
    
    lblQueQuedo = "Pendiente de " + QueHace + ": " + FormatCurrency(Saldo)
End Sub
