VERSION 5.00
Object = "{A7FBD38D-2930-49E3-B60C-9E0202D84549}#16.0#0"; "tbrControles.ocx"
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCierreCaja 
   BackColor       =   &H00544B45&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cierre Caja"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11415
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCierreCaja.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   11415
   StartUpPosition =   3  'Windows Default
   Begin tbrFaroButton.fBoton Command1 
      Height          =   465
      Left            =   9900
      TabIndex        =   19
      Top             =   7560
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   "Arial Narrow"
      fFontS          =   10
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdCerrar 
      Height          =   465
      Left            =   9840
      TabIndex        =   2
      Top             =   5940
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Cerrar Caja"
      fEnabled        =   -1  'True
      fFontN          =   "Arial Narrow"
      fFontS          =   10
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton Command2 
      Height          =   465
      Left            =   9930
      TabIndex        =   1
      Top             =   3150
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Ver"
      fEnabled        =   -1  'True
      fFontN          =   "Arial Narrow"
      fFontS          =   10
      fECol           =   5717301
   End
   Begin tbrControles.MouTextBox txtEfvo 
      Height          =   435
      Left            =   7860
      TabIndex        =   0
      Top             =   3150
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   767
      Alignment       =   2
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
   Begin MSComctlLib.ListView lvCierre 
      Height          =   2535
      Left            =   210
      TabIndex        =   18
      Top             =   2220
      Width           =   6975
      _ExtentX        =   12303
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
         Name            =   "Arial Narrow"
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
         Object.Width           =   1764
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
   Begin VB.Timer Timer1 
      Left            =   360
      Top             =   360
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cierre de Caja"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   795
      Left            =   3300
      TabIndex        =   3
      Top             =   330
      Width           =   4815
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sin datos"
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
      Height          =   345
      Left            =   5040
      TabIndex        =   17
      Top             =   7050
      Width           =   1335
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Dif. Caja a registrar"
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
      Left            =   2700
      TabIndex        =   16
      Top             =   7110
      Width           =   2175
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
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
      Height          =   315
      Left            =   2700
      TabIndex        =   15
      Top             =   6660
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sin cierre de caja"
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
      Height          =   345
      Left            =   3960
      TabIndex        =   14
      Top             =   6660
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   1725
      Left            =   7440
      TabIndex        =   13
      Top             =   4020
      Width           =   3615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Efectivo recontado"
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
      Height          =   435
      Left            =   7740
      TabIndex        =   12
      Top             =   2820
      Width           =   2175
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   1305
      Left            =   7440
      TabIndex        =   11
      Top             =   1290
      Width           =   3615
   End
   Begin VB.Label lblEF 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5430
      TabIndex        =   10
      Top             =   5430
      Width           =   1155
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Efectivo segun contabilidad (es lo que debería haber en caja)"
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
      Height          =   825
      Left            =   1470
      TabIndex        =   9
      Top             =   5400
      Width           =   3705
   End
   Begin VB.Label lblEI 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5430
      TabIndex        =   8
      Top             =   1320
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Efectivo en el ultimo Cierre"
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
      Left            =   2310
      TabIndex        =   7
      Top             =   1470
      Width           =   2865
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Variaciones de Caja"
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
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Top             =   1830
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Disminucion de Fondos"
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
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   4950
      Width           =   2715
   End
   Begin VB.Label lblVariacion 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5460
      TabIndex        =   4
      Top             =   4890
      Width           =   1125
   End
End
Attribute VB_Name = "frmCierreCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Efvo As Single
Dim CajaSup As Single
Dim Cerrado As Boolean
Dim CY As String
Dim VariacionCaja As Single

Private Sub cmdCerrar_Click()
    Dim tmpEf As Single, tmpDif As Single, VarCaja As Single
    Dim strCierre As String, IDcierre As Long, H As Long
    
    Efvo = ValidarNumeros(txtEfvo)
    txtEfvo = FormatCurrency(Efvo, , vbTrue, , vbTrue)
    VarCaja = PC.UltVarCuenta(78)
    
    If MsgBox("Está a punto de cerrar caja por " + FormatCurrency(CSng(txtEfvo), , vbTrue, , vbFalse) + _
        " ¿Desea continuar?", vbYesNo) = vbNo Then Exit Sub
    
    cmdCerrar.Enabled = False
    
    frmWAIT.Show
    frmWAIT.Refresh
    Timer1.Interval = 100
    
    Cerrado = True
    'primero le cuento al usuario cuanto deberia haber en caja
    'registro la diferencia de caja
    cmdCerrar.Enabled = False
    Command2.Enabled = False
    
    '¿Quién esta haciendo el cierre?
    Dim UltUser As String, UUs As Long
    
    UUs = ACC.UltUsuarioIngresado
    UltUser = ACC.GetNombre("Usuario", "Usuarios", UUs)
    
    CajaSup = PC.CerrarCaja  'Cierro!!!!!!!!!!!!!!!!!!!!
    PC.Asiento "78", CStr(Efvo - CajaSup), "32", CStr(Efvo - CajaSup), _
        "LibroDiario", "Cierre de Caja al " + FormatDateTime(Date, vbShortDate) + _
        " realizado por " + UltUser
        
    'negrada pero necesito volver a cerrar caja para que este asiento quede marcado
    'como cerrocaja=si
    PC.CerrarCaja
    
    Label9 = "Cerrada correctamente"
    Label12 = FormatCurrency(Efvo - CajaSup, , vbTrue, , vbFalse)
    tmpEf = PC.ABSSumarconSubcuentas(78)
    tmpDif = Efvo - CajaSup
    
    txtEfvo.Locked = True
    Label7 = "CAJA CERRADA " + vbCrLf + "Efectivo Registrado al cierre: " + _
        FormatCurrency(tmpEf, , , , vbFalse)
    Label7.FontSize = 14
    Label7.Alignment = vbCenter
    

    'registro el movimiento(2:Cerrar Caja) con descripcion chiche
    ACC.RegEvento UUs, 2, "Caja Cerrada" + _
        ". Efectivo: " + FormatCurrency(tmpEf, , vbTrue, , vbFalse) + _
        ". Dif Caja: " + FormatCurrency(tmpDif, , vbTrue, , vbFalse)
    
    'grabo el cierre (o algo asi)
    IDcierre = PC.GetUltIdCierreCaja
        'como es el string cierre?
    strCierre = ""
    For H = 1 To lvCierre.ListItems.Count
        strCierre = strCierre + txtInLvW(lvCierre, H, 2) + _
            ": " + txtInLvW(lvCierre, H, 3)
        If H < lvCierre.ListItems.Count Then strCierre = strCierre + "\\"
    Next H
    
    PC.GrabarCierre IDcierre, Efvo, strCierre, tmpDif, VarCaja
    
    frmWAIT.Hide
    Timer1.Interval = 0
    Unload frmWAIT
    Unload Me
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    txtEfvo = FormatCurrency(ValidarNumeros(txtEfvo), , vbTrue, , vbFalse)
    
    Dim TmP As String
    If Efvo > CajaSup Then
        TmP = "sobrante"
    Else
        TmP = "faltante"
    End If
    
    Label7 = "Hay un " + TmP + " de " + _
        FormatCurrency(Abs(Efvo - CajaSup), , vbTrue, , vbFalse) + "." + vbCrLf + _
        "Si verifica una diferencia considerable presione Salir para realizar " + _
        "las anotaciones que pudiesen no estar registradas." + vbCrLf + _
        "Si la diferencia es razonable continue con el cierre " + _
        "de caja."
    
    cmdCerrar.Enabled = True
    
End Sub

Public Sub AbrirPre(ValEFVO As String)
    txtEfvo = ValEFVO
    Me.Show
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Cerrado = False
    
    Efvo = 0
    txtEfvo = FormatCurrency(Efvo, , vbTrue, , vbFalse)
    txtEfvo.SelStart = 0
    txtEfvo.SelLength = Len(txtEfvo)
    
    'se borra si no!!!!!!!!
    FormatearMouTextBox frmCierreCaja, 11
    
    Label6 = "Está a punto de realizar el cierre de caja con todos sus registros " + _
        "correspondientes, ¿Está seguro que desea realizarlo?, el último cierre se " + _
        "realizó el dia " + CStr(PC.FechaUltimoCierre)
    
    'primero le muestro cuanto deberia haber en caja para ver si le falta registrar
    'algo importante
    
    'lo que tenia MAS lo que se movio en el ejercicio
    lblEI = FormatCurrency(PC.MovCerrados(78), , vbTrue, , vbFalse)
    CajaSup = PC.GetSaldo(78)
    lblEF = FormatCurrency(CajaSup, , , , vbFalse)
    
    '!!!!!!!!!!!! VER !!!!!!!!!!!!!!!!!!!!!!!!
    'listo los mov de caja
    Dim SubD As Single, tmp2 As Long, TmP As Long
    
    'tengo que agregar un renglon con lo del subdiario
    SubD = PC.GetSaldo(78, "LibroSubDiario")
    tmp2 = lvCierre.ListItems.Count + 1
    
    TmP = PC.GetUltIdCierreCaja
    PC.ListarMovCuentaPorNroAsiento 78, TmP + 1, 0, lvCierre, , False
    
    lvCierre.ListItems.Add tmp2
    
    lvCierre.ListItems(tmp2).Text = FormatDateTime(Date, vbShortDate)
    lvCierre.ListItems(tmp2).SubItems(1) = ""
    lvCierre.ListItems(tmp2).SubItems(2) = "Saldo a la Fecha SubDiario Compras-Ventas"
    lvCierre.ListItems(tmp2).SubItems(3) = FormatCurrency(SubD, , vbTrue, , vbFalse)
    
    VariacionCaja = PC.UltVarCuenta(78)
    
    If VariacionCaja < 0 Then
        Label3 = "Disminución de caja"
    Else
        Label3 = "Aumento en caja"
    End If
    
    lblVariacion = FormatCurrency(Abs(VariacionCaja), , vbTrue, , vbFalse)

    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Cerrado Then MsgBox "Quedá registrado en Caja " + _
        FormatCurrency(PC.ABSSumarconSubcuentas(78), , vbTrue, , vbFalse) + _
        " para el próximo cierre", vbInformation, "Cierre Caja"
        
    If Cerrado = True Then 'reinicio variables
        VtaDia = 0
        CtoDia = 0
        CpDia = 0
    End If
End Sub

Private Sub txtEfvo_Change()
    If vbkey > 0 Then cmdCerrar.Enabled = False
End Sub

Private Sub txtEfvo_GotFocus()
    
    txtEfvo.SelStart = 0
    txtEfvo.SelLength = Len(txtEfvo)
End Sub

Private Sub txtEfvo_LostFocus()
    Efvo = ValidarNumeros(txtEfvo)
    txtEfvo = FormatCurrency(Efvo, , vbTrue, , vbFalse)
End Sub
