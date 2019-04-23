VERSION 5.00
Object = "{57F2DDA9-75EB-4696-9A07-7D86D576ABB4}#21.0#0"; "tbrControles.ocx"
Begin VB.Form frmCuentas 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Egresos"
   ClientHeight    =   8985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10020
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   10020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin tbrControles.MouTextBox txtMonto 
      Height          =   435
      Left            =   8490
      TabIndex        =   1
      Top             =   3990
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   767
      BackColor       =   16777215
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "EGRESOS"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1995
      Left            =   240
      TabIndex        =   10
      Top             =   6570
      Width           =   5205
      Begin VB.ListBox lstEgresos 
         Height          =   1320
         IntegralHeight  =   0   'False
         ItemData        =   "frmE.frx":030A
         Left            =   150
         List            =   "frmE.frx":031D
         TabIndex        =   11
         Top             =   570
         Width           =   4905
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Ultimos 5 Egresos desde el último cierre de caja"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   150
         TabIndex        =   12
         Top             =   300
         Width           =   3825
      End
   End
   Begin VB.CommandButton cmdPagar 
      Caption         =   "Pagar"
      Height          =   405
      Left            =   8610
      TabIndex        =   2
      Top             =   4500
      Width           =   975
   End
   Begin VB.CommandButton cmdGrabarDetalle 
      Caption         =   "Grabar"
      Height          =   405
      Left            =   7290
      TabIndex        =   8
      Top             =   5550
      Width           =   975
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   645
      Left            =   210
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   5400
      Width           =   6945
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   405
      Left            =   8580
      TabIndex        =   5
      Top             =   8010
      Width           =   975
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "Modificar"
      Height          =   405
      Left            =   8580
      TabIndex        =   4
      Top             =   1260
      Width           =   975
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Agregar"
      Height          =   405
      Left            =   8580
      TabIndex        =   3
      Top             =   750
      Width           =   975
   End
   Begin VB.ListBox lstCuentas 
      Height          =   4620
      Left            =   150
      TabIndex        =   0
      Top             =   330
      Width           =   8085
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Monto"
      Height          =   255
      Left            =   8820
      TabIndex        =   9
      Top             =   3720
      Width           =   825
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   210
      TabIndex        =   7
      Top             =   5130
      Width           =   1155
   End
End
Attribute VB_Name = "frmCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()
    Dim NbCta As String
    
    If lstCuentas.ListIndex = -1 Then Exit Sub
    
    NbCta = InputBox("Ingrese el nombre de la cuenta", "Nombre")
    
    If NbCta = "" Then
        MsgBox "Cargue correctamente los datos", vbExclamation, "Atención"
        Exit Sub
    End If
    
    'crear una cuenta en el nivel elegido
    Dim SP() As String
    SP = Split(lstCuentas, ".")
    
    PC.AgregarCuenta CLng(SP(1)), NbCta
    
    'para que se actualice medio negra la cosa
    If Right(lstCuentas, 1) = "*" Then
        lstCuentas_DblClick
        lstCuentas_DblClick
    Else
        lstCuentas_DblClick
    End If
End Sub

Private Sub cmdGrabarDetalle_Click()
    If lstCuentas.ListIndex = -1 Then Exit Sub
    
    Dim SP() As String
    SP = Split(lstCuentas, ".")

    PC.ModificarCuenta SP(2), SP(2), txtDescripcion
    
End Sub

Private Sub cmdModificar_Click()
    Dim nCta As String
    Dim SP() As String
    
    If lstCuentas.ListIndex = -1 Then Exit Sub
    
    SP = Split(lstCuentas, ".")
    nCta = InputBox("Ingrese el nuevo nombre de la cuenta", "Nombre", SP(2))
    
    If nCta = "" Then
        MsgBox "Cargue correctamente los datos", vbExclamation, "Atención"
        Exit Sub
    End If
    
    'modificar la cuenta
    PC.ModificarCuenta SP(2), nCta
    
    'actualizo
    lstCuentas.List(lstCuentas.ListIndex) = SP(0) + "." + SP(1) + "." + nCta
End Sub




Private Sub cmdPagar_Click()
    'primero valido
    txtMonto = FormatCurrency(ValidarNumeros(txtMonto), , , , vbFalse)
    
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
    PC.Asiento CStr(PC.GetIDCuenta(nCuenta)), txtMonto, "78", txtMonto
    
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CargarDaTos
    
    txtMonto = FormatCurrency(0)
End Sub

Private Sub CargarDaTos()
    lstCuentas.Clear
    
    lstCuentas.AddItem ".19.Gastos de Comercializacion"
    lstCuentas.AddItem ".20.Gastos de Administracion"
    lstCuentas.AddItem ".37.Gastos de Financiacion"
    lstCuentas.AddItem ".33.Impuestos"
    lstCuentas.AddItem ".34.Servicios"
    lstCuentas.AddItem ".22.Otros Egresos"
    
    FormatearMouTextBox frmEgresos
    PC.ListarEgresos lstEgresos, 5
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
    Dim CTAS() As String
    CTAS = PC.GetCuentas(CTA)
    If UBound(CTAS) = 0 Then
        MsgBox "No tiene subcuentas!"
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
    For A = 1 To UBound(CTAS)
        'ponerle los mismos espacios que tenia mas 3
        lstCuentas.AddItem "   " + SP(0) + "." + CTAS(A) + "." + _
            PC.GetNameCuenta(CLng(CTAS(A))), lstCuentas.ListIndex + 1
        
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

Private Sub txtMonto_GotFocus()
    PintarTxt txtMonto
End Sub

Private Sub txtMonto_LostFocus()
    txtMonto = FormatCurrency(ValidarNumeros(txtMonto), , , , vbFalse)
End Sub
