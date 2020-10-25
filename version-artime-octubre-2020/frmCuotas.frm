VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A7FBD38D-2930-49E3-B60C-9E0202D84549}#13.0#0"; "tbrControles.ocx"
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmCuotas 
   BackColor       =   &H00544B45&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configurar Cuotas"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9450
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCuotas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   9450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin tbrFaroButton.fBoton cmdAdd 
      Height          =   435
      Left            =   7560
      TabIndex        =   24
      Top             =   5550
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "agregar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdAgregarCuotas 
      Height          =   435
      Left            =   5280
      TabIndex        =   23
      Top             =   3750
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "agregar cuotas"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.ComboBox cmbSistemaInteres 
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
      ItemData        =   "frmCuotas.frx":000C
      Left            =   5880
      List            =   "frmCuotas.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   3240
      Visible         =   0   'False
      Width           =   2505
   End
   Begin VB.ComboBox cmbPlazoInteres 
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
      ItemData        =   "frmCuotas.frx":0036
      Left            =   1320
      List            =   "frmCuotas.frx":0040
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   3240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ComboBox cmbTipoInteres 
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
      ItemData        =   "frmCuotas.frx":0054
      Left            =   3270
      List            =   "frmCuotas.frx":005E
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3240
      Visible         =   0   'False
      Width           =   2505
   End
   Begin tbrControles.MouTextBox txtInteresPorC 
      Height          =   345
      Left            =   4650
      TabIndex        =   1
      Top             =   2550
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   609
      Alignment       =   2
      BackColor       =   16777215
      Text            =   "0"
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
   Begin VB.CheckBox chkInteres 
      BackColor       =   &H00544B45&
      Caption         =   "Tiene Interés"
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
      Left            =   2220
      TabIndex        =   0
      Top             =   2580
      Width           =   1575
   End
   Begin VB.ComboBox cmbNCuotas 
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
      ItemData        =   "frmCuotas.frx":0083
      Left            =   3600
      List            =   "frmCuotas.frx":0085
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3810
      Width           =   1545
   End
   Begin MSComCtl2.DTPicker DTVenc 
      Height          =   375
      Left            =   5940
      TabIndex        =   6
      Top             =   4410
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   661
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
      Format          =   20971521
      CurrentDate     =   39196
   End
   Begin MSComctlLib.ListView lvResumen 
      Height          =   2235
      Left            =   2100
      TabIndex        =   9
      Top             =   4920
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   3942
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
         Text            =   "Vencimi."
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Cuota"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Interes"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Total"
         Object.Width           =   2469
      EndProperty
   End
   Begin tbrControles.MouTextBox txtPesos 
      Height          =   405
      Left            =   2820
      TabIndex        =   5
      Top             =   4440
      Width           =   1275
      _ExtentX        =   2249
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
   Begin tbrFaroButton.fBoton cmdTerminar 
      Height          =   435
      Left            =   4350
      TabIndex        =   25
      Top             =   7740
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "registrar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdLimpiar 
      Height          =   435
      Left            =   7560
      TabIndex        =   26
      Top             =   6570
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "limpiar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdQuitar 
      Height          =   435
      Left            =   7560
      TabIndex        =   27
      Top             =   6060
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "quitar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdSalir 
      Height          =   435
      Left            =   7470
      TabIndex        =   28
      Top             =   7920
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D8D9D7&
      BackStyle       =   0  'Transparent
      Caption         =   "Nro. Documento:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   690
      TabIndex        =   21
      Top             =   1170
      Width           =   3150
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Totales"
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
      Left            =   2130
      TabIndex        =   20
      Top             =   7260
      Width           =   795
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5850
      TabIndex        =   19
      Top             =   7200
      Width           =   1305
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4410
      TabIndex        =   18
      Top             =   7200
      Width           =   1305
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2970
      TabIndex        =   17
      Top             =   7230
      Width           =   1305
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "%"
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
      Height          =   405
      Left            =   5370
      TabIndex        =   16
      Top             =   2580
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Interés"
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
      Left            =   3780
      TabIndex        =   15
      Top             =   2610
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label lblDoc 
      BackColor       =   &H00D8D9D7&
      BackStyle       =   0  'Transparent
      Caption         =   "Factura Nro: A-0012-31232334"
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
      Height          =   465
      Left            =   3990
      TabIndex        =   14
      Top             =   1170
      Width           =   4230
   End
   Begin VB.Label lblPendiente 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pendiente $ 1002,11"
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
      Left            =   1980
      TabIndex        =   13
      Top             =   7830
      Width           =   2205
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Completar con Cuotas Mensuales"
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
      Left            =   390
      TabIndex        =   12
      Top             =   3870
      Width           =   3105
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Importe"
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
      Left            =   1890
      TabIndex        =   11
      Top             =   4470
      Width           =   885
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimiento"
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
      Left            =   4680
      TabIndex        =   10
      Top             =   4470
      Width           =   1125
   End
   Begin VB.Label lblImporte 
      Alignment       =   2  'Center
      BackColor       =   &H00D8D9D7&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Importe: $ 4500,32"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2430
      TabIndex        =   8
      Top             =   1710
      Width           =   4200
   End
   Begin VB.Label lblQue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cuotas a Cobrar de Wedaskfkljdfkljg  dfwerrrgargagrh  hrejklerwklg Gomez"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   765
      Left            =   1110
      TabIndex        =   7
      Top             =   330
      Width           =   7275
   End
End
Attribute VB_Name = "frmCuotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IdPers As String, IsCliente As Boolean, IdMovimiento As Long
Dim Importe As Single, Doc As String, Nombre As String, Pendiente As Single

Private Sub chkInteres_Click()
    If chkInteres Then
        Label2.Visible = True
        Label3.Visible = True
        txtInteresPorC.Visible = True
        txtInteresPorC = CFG.GetInfo(11, 4)
        txtInteresPorC = ValidarNumeros(txtInteresPorC)
        cmbTipoInteres.Visible = True
        cmbPlazoInteres.Visible = True
        cmbSistemaInteres.Visible = True
    Else
        Label2.Visible = False
        Label3.Visible = False
        txtInteresPorC.Visible = False
        cmbTipoInteres.Visible = False
        cmbPlazoInteres.Visible = False
        cmbSistemaInteres.Visible = False
    End If
End Sub

Private Sub cmbPlazoInteres_Click()
    CalcularTodosIntereses
End Sub

Private Sub cmbTipoInteres_Click()
    CalcularTodosIntereses
End Sub

Private Sub cmdAdd_Click()
    Dim Ajuste As Single, TmP As Long, EFh As Long
    
    txtInteresPorC = ValidarNumeros(txtInteresPorC)
    txtPesos = FormatCurrency(ValidarNumeros(txtPesos), , , , vbFalse)
    If EsCero(CSng(txtPesos)) = True Then Exit Sub
    
    If chkInteres Then
        Ajuste = CalcularInteres(DateDiff("d", Date, DTVenc), CSng(txtInteresPorC), _
            cmbPlazoInteres, cmbTipoInteres.ListIndex)
    Else
        Ajuste = 1
    End If
    
    
    EFh = EstaFecha(DTVenc)
        
    If EFh = 0 Then
        TmP = lvResumen.ListItems.Count + 1
        lvResumen.ListItems.Add TmP
        
        lvResumen.ListItems(TmP).Text = CStr(DTVenc)
        lvResumen.ListItems(TmP).SubItems(1) = FormatCurrency(CSng(txtPesos), 4, , , vbFalse)
        lvResumen.ListItems(TmP).SubItems(3) = FormatCurrency(CSng(txtPesos) * Ajuste, 4, , , vbFalse)
        lvResumen.ListItems(TmP).SubItems(2) = FormatCurrency( _
            CSng(lvResumen.ListItems(TmP).SubItems(3)) - CSng(txtPesos), 4, , , vbFalse)
    Else
        lvResumen.ListItems(EFh).SubItems(1) = FormatCurrency(CSng( _
            lvResumen.ListItems(EFh).SubItems(1)) + txtPesos, 4, , , vbFalse)
        lvResumen.ListItems(EFh).SubItems(3) = FormatCurrency(CSng( _
            lvResumen.ListItems(EFh).SubItems(1)) * Ajuste, 4, , , vbFalse)
        lvResumen.ListItems(EFh).SubItems(2) = FormatCurrency(CSng( _
            lvResumen.ListItems(EFh).SubItems(3)) - _
            CSng(lvResumen.ListItems(EFh).SubItems(1)), 4, , , vbFalse)
    End If
    
    CalcularTotales
    
    DTVenc = DTVenc + CLng(CFG.GetInfo(2, 4))
    txtPesos = FormatCurrency(0)
    PintarTxt txtPesos
End Sub

Private Function EstaFecha(Fecha As Date) As Long
    Dim I As Long, Resp As Long
    
    Resp = 0
    
    For I = 1 To lvResumen.ListItems.Count
        If lvResumen.ListItems(I).Text = CStr(Fecha) Then
            Resp = I
            Exit For
        End If
    Next I

    EstaFecha = Resp
End Function

Private Function CalculariMensual(Interes As Single, Optional Plazo As String = "Mensual", _
    Optional Capitaliza As Boolean = True) As Single
    Dim Resp As Single
    
    If Plazo = "Mensual" Then
        Resp = Interes / 100
    Else
        If Capitaliza Then
            Resp = (1 + Interes / 100) ^ (1 / 12)
        Else
            Resp = Interes / 12 / 100
        End If
    End If
    
    CalculariMensual = Resp
End Function

Private Function CalcularInteres(Dias As Long, Interes As Single, _
    Optional Plazo As String = "Mensual", Optional Capitaliza As Boolean = True) As Single
    'capitaliza es mensual nada mas por ahora
    
    Dim Resp As Single, Periodos As Single
    
    'devuelve cuanto debo multiplicar el capital para ajustar
    Select Case Plazo
        Case "Mensual"
            Periodos = Dias / 30
            If Capitaliza Then
                Resp = (1 + Interes / 100) ^ Periodos
            Else
                Resp = (1 + Interes / 100 * Periodos)
            End If
        Case "Anual"
            Periodos = Dias / 360
            If Capitaliza Then
                Resp = (1 + Interes / 100) ^ Periodos
            Else
                Resp = (1 + Interes / 100 * Periodos)
            End If
    End Select

    CalcularInteres = Resp
End Function

Private Sub cmdAgregarCuotas_Click()
    Dim nCuo As Long, APagar As Single, Fecha As Date, Intr As Single
    Dim F As Long, TmP As Long, Ajuste As Single
    Dim AmorPesos As Single, IntPesos As Single
    
    CalcularTotales
    
    If EsCero(Pendiente) Then
        MsgBox "No hay saldos pendientes de configurar. " + vbCrLf + _
            "Presione Limpiar en caso de querer configurarlo nuevamente", vbInformation, "Atención"
        Exit Sub
    End If
    
    nCuo = CLng(cmbNCuotas)
    If lvResumen.ListItems.Count = 0 Then
        Fecha = DTVenc
    Else
        Fecha = CDate(lvResumen.ListItems(lvResumen.ListItems.Count).Text)
    End If
    
    TmP = lvResumen.ListItems.Count + 1
    
    If chkInteres.Value = 0 Then
        For F = TmP To nCuo + TmP - 1
                Fecha = ProximoMes(Fecha)
                AmorPesos = Pendiente / nCuo
                IntPesos = 0
                
                lvResumen.ListItems.Add F
                
                lvResumen.ListItems(F).Text = CStr(Fecha)
                lvResumen.ListItems(F).SubItems(1) = FormatCurrency(AmorPesos, 4, , , vbFalse)
                lvResumen.ListItems(F).SubItems(2) = FormatCurrency(IntPesos, 4, , , vbFalse)
                lvResumen.ListItems(F).SubItems(3) = FormatCurrency(AmorPesos, 4, , , vbFalse)
        Next F
        
        CalcularTotales
        Exit Sub
    End If
    
    Intr = CalculariMensual(CSng(txtInteresPorC), cmbPlazoInteres, cmbTipoInteres.ListIndex)
    Select Case cmbSistemaInteres.ListIndex
        Case 0 'SISTEMA FRANCES ------------------------------------------------------
            APagar = ValorCuota(Intr * 100, Pendiente, nCuo)
                    
            For F = TmP To nCuo + TmP - 1
                Fecha = ProximoMes(Fecha)
                AmorPesos = AmortizacionCuota(Intr * 100, APagar, nCuo, F - TmP + 1)
                IntPesos = InteresCuota(Intr * 100, APagar, nCuo, F - TmP + 1)
                
                lvResumen.ListItems.Add F
                
                lvResumen.ListItems(F).Text = CStr(Fecha)
                lvResumen.ListItems(F).SubItems(1) = FormatCurrency(AmorPesos, 4, , , vbFalse)
                lvResumen.ListItems(F).SubItems(2) = FormatCurrency(IntPesos, 4, , , vbFalse)
                lvResumen.ListItems(F).SubItems(3) = FormatCurrency(APagar, 4, , , vbFalse)
            Next F
        Case 1 'SISTEMA ALEMAN ------------------------------------------------------
            AmorPesos = Pendiente / nCuo 'amortizacion constante, int y cuotas decrec
            
            For F = TmP To nCuo + TmP - 1
                IntPesos = Pendiente * Intr
                APagar = AmorPesos + IntPesos
                Pendiente = Pendiente - AmorPesos
                
                Fecha = ProximoMes(Fecha)
                lvResumen.ListItems.Add F
                
                lvResumen.ListItems(F).Text = CStr(Fecha)
                lvResumen.ListItems(F).SubItems(1) = FormatCurrency(AmorPesos, 4, , , vbFalse)
                lvResumen.ListItems(F).SubItems(2) = FormatCurrency(IntPesos, 4, , , vbFalse)
                lvResumen.ListItems(F).SubItems(3) = FormatCurrency(APagar, 4, , , vbFalse)
            Next F
        
        Case 2 'SISTEMA SIMPLE ------------------------------------------------------
            AmorPesos = Pendiente / nCuo
            IntPesos = Pendiente * Intr
            APagar = AmorPesos + IntPesos
            
            For F = TmP To nCuo + TmP - 1
                Fecha = ProximoMes(Fecha)
                lvResumen.ListItems.Add F
                
                lvResumen.ListItems(F).Text = CStr(Fecha)
                lvResumen.ListItems(F).SubItems(1) = FormatCurrency(AmorPesos, 4, , , vbFalse)
                lvResumen.ListItems(F).SubItems(2) = FormatCurrency(IntPesos, 4, , , vbFalse)
                lvResumen.ListItems(F).SubItems(3) = FormatCurrency(APagar, 4, , , vbFalse)
            Next F
    
    End Select
    
    CalcularTotales
End Sub

Private Sub CalcularTotales()
    Dim Tot(2) As Single, I As Long
    
    Tot(0) = 0: Tot(1) = 0: Tot(2) = 0
    For I = 1 To lvResumen.ListItems.Count
        Tot(0) = Tot(0) + CSng(NoNuloN(lvResumen.ListItems(I).SubItems(1)))
        Tot(1) = Tot(1) + CSng(NoNuloN(lvResumen.ListItems(I).SubItems(2)))
        Tot(2) = Tot(2) + CSng(NoNuloN(lvResumen.ListItems(I).SubItems(3)))
    Next I
    
    Label1 = FormatCurrency(Tot(0), , , , vbFalse)
    Label7 = FormatCurrency(Tot(1), , , , vbFalse)
    Label8 = FormatCurrency(Tot(2), , , , vbFalse)
    
    Pendiente = Importe - Tot(0)
    lblPendiente = "Pendiente: " + FormatCurrency(Pendiente, , , , vbFalse)
End Sub

Private Sub cmdLimpiar_Click()
    lvResumen.ListItems.Clear
    DTVenc = Date + CLng(CFG.GetInfo(2, 4))
    CalcularTodosIntereses
End Sub

Private Sub cmdQuitar_Click()
    If lvResumen.ListItems.Count = 0 Then Exit Sub
    
    lvResumen.ListItems.Remove (lvResumen.SelectedItem.Index)
    
    CalcularTodosIntereses
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Public Sub AbrirDatos(IDPersona As String, IdMov As Long, _
    Optional EsCliente As Boolean = True)
    
    Dim TmP As String
    
    IdPers = IDPersona
    IsCliente = EsCliente
    IdMovimiento = IdMov
    
    If IsCliente = True Then
        Nombre = DB.GetValInRS("Clientes", "Nombre", "ID = " + IDPersona)
        TmP = " Cobrar"
        Documento = DB.GetValInRS("MovClientes", "Documento", "ID = " + CStr(IdMovimiento))
        Importe = CSng(DB.GetValInRS("MovClientes", "Variacion", "ID = " + _
            CStr(IdMovimiento), False))
    Else
        Nombre = IdPers
        TmP = " Pagar"
        Documento = DB.GetValInRS("MovProveedores", "Documento", "ID = " + CStr(IdMovimiento))
        Importe = CSng(DB.GetValInRS("MovProveedores", "Variacion", "ID = " + _
            CStr(IdMovimiento), False))
    End If

    lblQue = "Cuotas a" + TmP + " a " + UCase(Nombre)
    lblImporte = "Importe: " + FormatCurrency(Importe, , , , vbFalse)
    lblDoc = Documento
    
    Me.Show 1
End Sub

Private Sub cmdTerminar_Click()
    Dim Tabla As String, Intr As Single

    'NO dejo terminar si no cancela lo pendiente
    If EsCero(Pendiente) = False Then
        MsgBox "Falta configurar cuotas por " + FormatCurrency(Pendiente), vbInformation, "Atención"
        Exit Sub
    End If
    
    'primero que todo borro el vencimiento que se registro en frmClientesMov
    If IsCliente Then
        Tabla = "Vencimientos"
    Else
        Tabla = "VencimientoProveedor"
    End If
    
    Intr = CSng(Label7)
    
    DB.EXECUTE "DELETE FROM " + Tabla + " WHERE IdMov = " + CStr(IdMovimiento)
    'ahora agrego tal cual se ve en el listview
    Dim H As Long
    
    For H = 1 To lvResumen.ListItems.Count
        DB.EXECUTE "INSERT INTO " + Tabla + " (IdMov, Cuota, Interes, Total, Vencimiento) " + _
            "VALUES (" + CStr(IdMovimiento) + "," + _
            Replace(CStr(CSng(lvResumen.ListItems(H).SubItems(1))), ",", ".") + "," + _
            Replace(CStr(CSng(lvResumen.ListItems(H).SubItems(2))), ",", ".") + "," + _
            Replace(CStr(CSng(lvResumen.ListItems(H).SubItems(3))), ",", ".") + ",#" + _
            stFechaSQL(CDate(lvResumen.ListItems(H).Text)) + "#)"
    Next H
    
    If EsCero(Intr) Then
        Unload Me
        Exit Sub
    End If
    
     'SOLO SI HAY INTERES!!!!!!!!!!!
    'asiento contable
    'Clientes (46) a Intereses Cobrados(68)-o-Intereses Pagados (69) a Proveedores (41)
    'tambien se lo tengo que sumar en su cuenta particular
    
    If IsCliente Then
        PC.Asiento "46", CStr(Intr), "68", CStr(Intr), "LibroSubdiario", _
            "Intereses por plan de cuotas a Cliente Nro " + IdPers
        'anoto en cuenta
        DB.EXECUTE "INSERT INTO MovClientes (ID, Fecha, CodCliente," + _
            "Variacion,Detalle,Documento) VALUES (" + IdAutonum("MovClientes") + _
            ",#" + stFechaSQL(Date) + _
            "#," + IdPers + "," + _
            Replace(CStr(Intr), ",", ".") + _
            ",'Intereses por plan de cuotas'," + _
            "'INT. Mov " + CStr(IdMovimiento) + "')"
    Else
        PC.Asiento "69", CStr(Intr), "41", CStr(Intr), "LibroSubdiario", _
            "Intereses por plan de cuotas a Proveedor " + CStr(IdPers)
        'anoto en cuenta
        DB.EXECUTE "INSERT INTO MovProveedores (ID, Fecha, Proveedor," + _
            "Variacion,Detalle,Documento) VALUES (" + IdAutonum("MovProveedores") + _
            ",#" + stFechaSQL(Date) + _
            "#,'" + IdPers + "'," + _
            Replace(CStr(Intr), ",", ".") + _
            ",'Intereses por plan de cuotas'," + _
            "'INT. Mov " + CStr(IdMovimiento) + "')"
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()
    DTVenc = Date + CLng(CFG.GetInfo(2, 4))
    lblPendiente = ""
    cmbPlazoInteres.ListIndex = 0
    cmbTipoInteres.ListIndex = 0
    cmbSistemaInteres.ListIndex = 0
    
    txtPesos = FormatCurrency(0)
    Label1 = FormatCurrency(0)
    Label7 = FormatCurrency(0)
    Label8 = FormatCurrency(0)
    
    Dim Q As Long, QQQ As Long
    
    QQQ = CLng(CFG.GetInfo(12, 4))
    cmbNCuotas.Clear
    For Q = 1 To QQQ
        cmbNCuotas.AddItem CStr(Q)
    Next Q
    
    cmbNCuotas.ListIndex = 0
End Sub

Private Sub txtInteresPorC_GotFocus()
    PintarTxt txtInteresPorC
End Sub

Private Sub txtInteresPorC_LostFocus()
    txtInteresPorC = ValidarNumeros(txtInteresPorC)
    CalcularTodosIntereses
End Sub

Private Sub CalcularTodosIntereses()
    If lvResumen.ListItems.Count = 0 Then Exit Sub
    
    Dim Im As Single, Y As Long, Ajuste As Single
    
    Im = CSng(txtInteresPorC)
    Ajuste = CalcularInteres(DateDiff("d", Date, DTVenc), CSng(txtInteresPorC), _
            cmbPlazoInteres, cmbTipoInteres.ListIndex)
    
    For Y = 1 To lvResumen.ListItems.Count
        Ajuste = CalcularInteres(DateDiff("d", Date, _
            CDate(lvResumen.ListItems(Y).Text)), Im, _
            cmbPlazoInteres, cmbTipoInteres.ListIndex)
        lvResumen.ListItems(Y).SubItems(3) = FormatCurrency(CSng( _
            lvResumen.ListItems(Y).SubItems(1)) * Ajuste, 4, , , vbFalse)
        lvResumen.ListItems(Y).SubItems(2) = FormatCurrency( _
            CSng(lvResumen.ListItems(Y).SubItems(3)) - _
            CSng(lvResumen.ListItems(Y).SubItems(1)), 4, , , vbFalse)
    Next Y
    
    CalcularTotales
End Sub

Private Sub txtPesos_GotFocus()
    PintarTxt txtPesos
End Sub

Private Sub txtPesos_LostFocus()
    txtPesos = FormatCurrency(ValidarNumeros(txtPesos), , , , vbFalse)
End Sub
