VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmBalance 
   BackColor       =   &H00544B45&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Balance"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10815
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBalance.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstGan 
      Height          =   1035
      Left            =   330
      TabIndex        =   24
      Top             =   7560
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.ListBox lstPer 
      Height          =   1035
      Left            =   5280
      TabIndex        =   23
      Top             =   7560
      Visible         =   0   'False
      Width           =   3990
   End
   Begin VB.ListBox lstPN 
      Height          =   1035
      Left            =   5460
      TabIndex        =   3
      Top             =   4830
      Width           =   4800
   End
   Begin VB.ListBox lstPasivo 
      Height          =   2400
      Left            =   5430
      TabIndex        =   2
      Top             =   1230
      Width           =   4800
   End
   Begin VB.ListBox lstActivo 
      Height          =   4935
      Left            =   210
      TabIndex        =   0
      Top             =   1230
      Width           =   4800
   End
   Begin tbrFaroButton.fBoton cmdSalir 
      Height          =   435
      Left            =   9360
      TabIndex        =   22
      Top             =   7770
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Resultados del Ejercicio"
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   5880
      TabIndex        =   21
      Top             =   6120
      Width           =   2415
   End
   Begin VB.Label lblGan 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "dfgsdgsfdg"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   750
      TabIndex        =   20
      Top             =   180
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Label lblPer 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "dfgsdgsfdg"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2580
      TabIndex        =   19
      Top             =   180
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Label lblSocyEmp 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "dfgsdgsfdg"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   510
      TabIndex        =   18
      Top             =   7110
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Label lblProveedores 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "dfgsdgsfdg"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6300
      TabIndex        =   17
      Top             =   7080
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Label lblClientes 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "dfgsdgsfdg"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4470
      TabIndex        =   16
      Top             =   7080
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Resultados del Ejercicio"
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   7050
      TabIndex        =   15
      Top             =   240
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblResEj 
      Alignment       =   2  'Center
      Caption         =   "Label8"
      Height          =   345
      Left            =   8400
      TabIndex        =   14
      Top             =   6060
      Width           =   1815
   End
   Begin VB.Label lblStock 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "dfgsdgsfdg"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2700
      TabIndex        =   13
      Top             =   7110
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   6540
      TabIndex        =   12
      Top             =   6480
      Width           =   1785
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   6510
      TabIndex        =   11
      Top             =   3900
      Width           =   1785
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   1230
      TabIndex        =   10
      Top             =   6420
      Width           =   1785
   End
   Begin VB.Label lblPN 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "dfgsdgsfdg"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   8370
      TabIndex        =   9
      Top             =   6420
      Width           =   1845
   End
   Begin VB.Label lblPasivo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "dfgsdgsfdg"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   8370
      TabIndex        =   8
      Top             =   3840
      Width           =   1845
   End
   Begin VB.Label lblActivo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "dfgsdgsfdg"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3150
      TabIndex        =   7
      Top             =   6360
      Width           =   1845
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Patrimonio Neto"
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   5430
      TabIndex        =   6
      Top             =   4410
      Width           =   4005
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pasivo"
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   5430
      TabIndex        =   5
      Top             =   870
      Width           =   4005
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Activo"
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   210
      TabIndex        =   4
      Top             =   870
      Width           =   4005
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Balance"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   555
      Left            =   4260
      TabIndex        =   1
      Top             =   180
      Width           =   2805
   End
End
Attribute VB_Name = "frmBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Activo As Single, Pasivo As Single, PN As Single

'CONTROLES A HACER
'   1- MAS IMPORTANTE A=P-PN
'   2- Suma del stock con la cuenta mercaderia
'   3- Clientes
'   4- Proveedores
'   5- Socios y Empleados
'   6- Vales

'CONTROL 1 ---------------------------------------------------------------------------
Private Sub Control1()
    'daaaaaaaaale hacelo
    Dim Gan As Single, Per As Single
    
    Gan = PC.ABSSumarconSubcuentas(3)
    Per = PC.ABSSumarconSubcuentas(4)
    
    lblResEj = FormatCurrency(-Gan + Per, , , , vbFalse)
    
End Sub

'CONTROL 2 ---------------------------------------------------------------------------
Private Sub Control2()
    Dim Stock As Single, Contab As Single
    
    Stock = DB.SumarProducto("Productos", "Stock", "pCosto", "ID>=0")
    'contabilidad de saldos de mercaderia en inventario(54)+en transito(55) pendiente
    Contab = PC.GetSaldoHasta(54, Date + 1, , True)
    
    If Abs(Stock - Contab) > 1 Then  'si hay mas de $1 de diferencia ajusto
        PC.Asiento "54", CStr(Stock - Contab), "23", CStr(Stock - Contab), , _
            "Ajuste por diferencia Stock-Contabilidad"
    End If
    
    lblStock = FormatCurrency(Stock, , , , vbFalse)
End Sub

'CONTROL 3 ---------------------------------------------------------------------------
Private Sub Control3()
    Dim Clientes As Single, Contab As Single
    
    Clientes = DB.SumarValInRS("MovClientes", "Variacion", "ID>=0")
    Contab = PC.GetSaldoHasta(46, Date + 1, , True)
    
    If Abs(Clientes - Contab) > 1 Then  'si hay mas de $1 de diferencia ajusto
        PC.Asiento "46", CStr(Clientes - Contab), "23", CStr(Clientes - Contab), , _
            "Ajuste por diferencia Clientes-Contabilidad"
    End If
    
    lblClientes = FormatCurrency(Clientes, , , , vbFalse)
End Sub
    
'CONTROL 4 ---------------------------------------------------------------------------
Private Sub Control4()
    Dim Prov As Single, Contab As Single
    
    Prov = DB.SumarValInRS("MovProveedores", "Variacion", "ID>=0")
    Contab = -PC.GetSaldoHasta(41, Date + 1, , True)
    
    If Abs(Prov - Contab) > 1 Then 'si hay mas de $1 de diferencia ajusto
        PC.Asiento "23", CStr(Prov - Contab), "41", CStr(Prov - Contab), , _
            "Ajuste por diferencia Proveedores-Contabilidad"
    End If
    
    lblProveedores = FormatCurrency(Prov, , , , vbFalse)
End Sub
    
'CONTROL 5 ---------------------------------------------------------------------------
Private Sub Control5()
    Dim SocyEmp As Single, Contab As Single, I As Long
    'tengo que hacerlo 1 x 1
    Dim Cuenta() As String '53 empleados, 52 socios
    
    'se queda con lo que dice MovSocyEmp BD Principal
    Cuenta = PC.GetCuentas(52) 'SOCIOS -----------------------------------------
    If UBound(Cuenta) > 0 Then
        For I = 1 To UBound(Cuenta)
            SocyEmp = DB.SumarValInRS("MovSocyEmp", "Variacion", _
                "IDNivel3 = " + Cuenta(I) + " AND Detalle <> 'Aporte Participación Socio'")
            Contab = -PC.GetSaldoHasta(CLng(Cuenta(I)), Date + 1, , True)

            If Abs(SocyEmp - Contab) > 1 Then 'si hay mas de $1 de diferencia ajusto
                PC.Asiento "23", CStr(SocyEmp - Contab), Cuenta(I), _
                    CStr(SocyEmp - Contab), , _
                    "Ajuste por diferencia Socios y Empleados-Contabilidad (" + Cuenta(I) + ")"
            End If
        Next I
    End If

    Cuenta = PC.GetCuentas(53) 'EMPLEADOS -----------------------------------------
    If UBound(Cuenta) > 0 Then
        For I = 1 To UBound(Cuenta)
            SocyEmp = DB.SumarValInRS("MovSocyEmp", "Variacion", "IDNivel3 = " + Cuenta(I))
            Contab = -PC.GetSaldoHasta(CLng(Cuenta(I)), Date + 1, , True)

            If Abs(SocyEmp - Contab) > 1 Then 'si hay mas de $1 de diferencia ajusto
                PC.Asiento "23", CStr(SocyEmp - Contab), Cuenta(I), _
                    CStr(SocyEmp - Contab), , _
                    "Ajuste por diferencia Socios y Empleados-Contabilidad (" + Cuenta(I) + ")"
            End If
        Next I
    End If
    
    SocyEmp = DB.SumarValInRS("MovSocyEmp", "Variacion", "ID>=0")
    lblSocyEmp = FormatCurrency(SocyEmp, , , , vbFalse)
End Sub
    
'CONTROL 6 ---------------------------------------------------------------------------
Private Sub Control6()
    Dim Vales As Single, Contab As Single
    
    Vales = DB.SumarValInRS("MovEnvases", "DepositoPorEnvase", "ID>=0")
    Contab = PC.ABSValor(93, PC.GetSaldoHasta(93, Date + 1, , True))
    
    If Abs(Vales - Contab) > 1 Then  'si hay mas de $1 de diferencia ajusto
        PC.Asiento "23", CStr(Vales - Contab), "93", CStr(Vales - Contab), , _
            "Ajuste por diferencia Vales-Contabilidad"
    End If
End Sub
    
Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Control1
    Control2
    Control3
    Control4
    Control5
    Control6
    
    Dim Gan As Single, Per As Single
    
    Per = PC.ABSSumarconSubcuentas(3)
    Gan = PC.ABSSumarconSubcuentas(4)
    
    lblResEj = FormatCurrency(Gan - Per, , , , vbFalse)
    
    CargarDatos
    
    lstActivo.ListIndex = 0
    lstPasivo.ListIndex = 0
    lstPN.ListIndex = 0
    
    'prueba
    Label8 = FormatCurrency(CSng(lblActivo) - CSng(lblPasivo) - CSng(lblPN))
End Sub

Private Sub CargarDatos()
    lstActivo.Clear
    lstPasivo.Clear
    lstPN.Clear
    
    lstActivo.AddItem ".1.Activo"
    lstPasivo.AddItem ".2.Pasivo"
    lstPN.AddItem ".5.Patrimonio Neto"
    
    lstGan.AddItem ".3.Ganancias"
    lstPer.AddItem ".4.Perdidas"
    
    lblActivo = FormatCurrency(VerSub(1, lstActivo), , , , vbFalse)
    lblPasivo = FormatCurrency(VerSub(2, lstPasivo), , , , vbFalse)
    lblPN = FormatCurrency(VerSub(5, lstPN) + CSng(lblResEj), , , , vbFalse)
    
    lblGan = FormatCurrency(VerSub(3, lstGan), , , , vbFalse)
    lblPer = FormatCurrency(VerSub(4, lstPer), , , , vbFalse)
End Sub

Private Function VerSub(IDC As Long, lsT As ListBox) As Single
    Dim Ctas() As String, A As Long
    Dim Nivel As Long, Sumo As Single, Resp As Single
    
    Nivel = 2: Sumo = 0: Resp = 0
    
    Ctas = PC.GetCuentas(IDC)
    
    If UBound(Ctas) = 0 Then
        Exit Function
    End If
    
    For A = 1 To UBound(Ctas)
        Sumo = PC.ABSValor(CLng(Ctas(A)), PC.GetSaldoHasta(CLng(Ctas(A)), Date + 1, , True))
        'ponerle los mismos espacios que tenia mas 3
        lsT.AddItem String(Nivel * 3, " ") + "." + Ctas(A) + "." + _
            PC.GetNameCuenta(CLng(Ctas(A))) + ": " + _
            FormatCurrency(Sumo, , , , vbFalse)
        Resp = Resp + Sumo + AnadirHijos(CLng(Ctas(A)), Nivel, lsT)
    Next A
    
    VerSub = Resp
    'lblSaldoCon = FormatCurrency(CSng(lblSaldos) + SaldoSub, , , , vbFalse)
End Function

Private Function AnadirHijos(IdCta As Long, Nivel As Long, lsT As ListBox) As Single
    Dim Hij() As String, Niv As Long, A As Long, Sumo As Single, Resp As Single
    
    Sumo = 0: Resp = 0
    
    Hij = PC.GetCuentas(IdCta)
        
    Niv = Nivel + 1
    
    If UBound(Hij) > 0 Then
        For A = 1 To UBound(Hij)
            Sumo = PC.ABSValor(CLng(Hij(A)), PC.GetSaldoHasta(CLng(Hij(A)), Date + 1, , True))
            lsT.AddItem String(Niv * 3, " ") + "." + Hij(A) + "." + _
                PC.GetNameCuenta(CLng(Hij(A))) + ": " + _
                FormatCurrency(Sumo, , , , vbFalse)
            Resp = Resp + Sumo + AnadirHijos(CLng(Hij(A)), Niv, lsT)

        Next A
    End If
    
    AnadirHijos = Resp
End Function

