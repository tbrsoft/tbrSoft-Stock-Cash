VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{A7FBD38D-2930-49E3-B60C-9E0202D84549}#15.0#0"; "tbrControles.ocx"
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmCuentas 
   BackColor       =   &H00544B45&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cuentas Contables"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCuentas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin tbrFaroButton.fBoton Command2 
      Height          =   405
      Left            =   10440
      TabIndex        =   31
      Top             =   8130
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdAntAsiento 
      Height          =   405
      Left            =   8760
      TabIndex        =   33
      Top             =   7620
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Anterior"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdModAsiento 
      Height          =   465
      Left            =   3390
      TabIndex        =   23
      Top             =   7560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Modificar asiento"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdMoAsiento 
      Height          =   405
      Left            =   3630
      TabIndex        =   22
      Top             =   5430
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Modif Cta"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrControles.MouTextBox txtNroAsiento 
      Height          =   405
      Left            =   5820
      TabIndex        =   19
      Top             =   7620
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   714
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
   Begin VB.TextBox txtDetAsiento 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   180
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "frmCuentas.frx":000C
      Top             =   7050
      Width           =   2925
   End
   Begin MSComctlLib.ListView lvAsiento 
      Height          =   2055
      Left            =   4950
      TabIndex        =   16
      Top             =   5220
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   3625
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
         Text            =   "Id.Cta."
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Concepto"
         Object.Width           =   3440
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Debe"
         Object.Width           =   1958
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Haber"
         Object.Width           =   1958
      EndProperty
   End
   Begin tbrControles.MouTextBox txtDebe 
      Height          =   405
      Left            =   210
      TabIndex        =   1
      Top             =   6180
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   714
      Alignment       =   2
      BackColor       =   16777215
      Text            =   "$ 0,121"
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
   Begin VB.TextBox txtDescripcion 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   7140
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   990
      Width           =   4035
   End
   Begin VB.ListBox lstCuentas 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   180
      TabIndex        =   0
      Top             =   960
      Width           =   6855
   End
   Begin tbrControles.MouTextBox txtHaber 
      Height          =   405
      Left            =   1770
      TabIndex        =   2
      Top             =   6180
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   714
      Alignment       =   2
      BackColor       =   16777215
      Text            =   "$ 0,121"
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
   Begin tbrFaroButton.fBoton cmdAsentar 
      Height          =   405
      Left            =   4980
      TabIndex        =   24
      Top             =   4770
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   714
      fFColor         =   11919317
      fBColor         =   14737632
      fCapt           =   "Incluir Asiento"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdEliminar 
      Height          =   405
      Left            =   7170
      TabIndex        =   25
      Top             =   3690
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Eliminar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdModificar 
      Height          =   405
      Left            =   7170
      TabIndex        =   26
      Top             =   3180
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Modificar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdAgregar 
      Height          =   405
      Left            =   7170
      TabIndex        =   27
      Top             =   2670
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Agregar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdGrabarDetalle 
      Height          =   405
      Left            =   8670
      TabIndex        =   28
      Top             =   2670
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Grabar Descripción"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdLiAsiento 
      Height          =   405
      Left            =   3630
      TabIndex        =   29
      Top             =   6960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Limpiar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdQuAsiento 
      Height          =   405
      Left            =   3630
      TabIndex        =   30
      Top             =   6450
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Quitar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdAgAsiento 
      Height          =   405
      Left            =   3630
      TabIndex        =   3
      Top             =   5940
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Agregar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdSigAsiento 
      Height          =   405
      Left            =   8760
      TabIndex        =   32
      Top             =   8130
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Siguiente"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdVerUltimo 
      Height          =   405
      Left            =   7020
      TabIndex        =   34
      Top             =   8130
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Último Asiento"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdFind 
      Height          =   405
      Left            =   5310
      TabIndex        =   35
      Top             =   8130
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Buscar Asiento"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdNueAsiento 
      Height          =   405
      Left            =   3360
      TabIndex        =   36
      Top             =   8130
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Nuevo Asiento"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.Label lblCtaSelec 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Asiento Manual"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   240
      TabIndex        =   21
      Top             =   4170
      Width           =   6735
   End
   Begin VB.Label lblAsNro 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Asiento Nro 412131"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6810
      TabIndex        =   20
      Top             =   4650
      Width           =   2805
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Detalle Asiento"
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
      Left            =   960
      TabIndex        =   18
      Top             =   6780
      Width           =   2235
   End
   Begin VB.Label lblfecha 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
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
      Left            =   7320
      TabIndex        =   17
      Top             =   4920
      Width           =   1665
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Diferencia"
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
      Left            =   5700
      TabIndex        =   15
      Top             =   7320
      Width           =   1245
   End
   Begin VB.Label lblDif 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Diferencia $ 133,23"
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
      Left            =   6840
      TabIndex        =   14
      Top             =   7320
      Width           =   1245
   End
   Begin VB.Label lblHaber 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
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
      Left            =   9210
      TabIndex        =   13
      Top             =   7320
      Width           =   1065
   End
   Begin VB.Label lblDebe 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
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
      Left            =   8040
      TabIndex        =   12
      Top             =   7320
      Width           =   1005
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Haber"
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
      Left            =   2040
      TabIndex        =   11
      Top             =   5850
      Width           =   765
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Debe"
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
      Left            =   480
      TabIndex        =   10
      Top             =   5850
      Width           =   765
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione la cuenta arriba, ingrese montos al Debe o al Haber y Presione Agregar"
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
      Height          =   615
      Left            =   330
      TabIndex        =   9
      Top             =   5040
      Width           =   3495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Asiento Manual"
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
      Height          =   285
      Left            =   780
      TabIndex        =   8
      Top             =   4590
      Width           =   2355
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cuentas Contables"
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
      Height          =   615
      Left            =   870
      TabIndex        =   7
      Top             =   180
      Width           =   6675
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción Cuenta"
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
      Left            =   7200
      TabIndex        =   6
      Top             =   720
      Width           =   2115
   End
End
Attribute VB_Name = "frmCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Por ahora no se pueden modificar estas cuentas

' 1. Activo
' 2. Pasivo
' 3. Perdidas
' 4. Ganancias
' 5. P.Neto
' 6. Caja y Bancos
'14. Cuentas Particulares  --> (¡NI HIJOS!)
'15. Capital --> (¡NI HIJOS!)
'16. Resultados No Asignados
'17. Ventas --> (¡NI HIJOS!)
'18. Costo de Ventas
'19. Gastos de Comercializacion
'20. Gastos de Administracion"
'22. Otros Egresos
'23. RFyT
'28. Incobrables
'30. Gastos o descuentos por Compras
'32. Diferencias de Caja
'33. Impuestos
'34. Servicios
'35. Diferencias de Stock
'37. Gastos de Financiacion"
'41. Proveedores
'46. Clientes
'50. Credito Fiscal (para el IVA)
'54. Mercaderias en Inventario
'57. Resultado Eliminacion de cuenta
'68. Intereses Cobrados
'69. Intereses Pagados
'76. Debito Fiscal (para el IVA)
'78. Caja
'93. Vales a Pagar

' ESTAN EN LA MATRIZ NOMODIFICAR()

' Ni hacer asientos con estas cuentas
' 14. Cuentas Particulares
' 15. Capital
' 46. Clientes
' 41. Proveedores
'NI SUS HIJOS!!!!!!!!!!!!!!!!!!!!!!!!! (subcuentas)
' ESTAN EN LA MATRIZ NOASIENTO()

Dim Debe As Single, Haber As Single
Dim NOMODIFICAR() As String, NOASIENTO() As String

Private Sub cmdAgAsiento_Click()
    Dim SP() As String, TmP As Long
    
    SP = Split(lstCuentas, ".")
    
    txtDebe = FormatCurrency(ValidarNumeros(txtDebe), , , , vbFalse)
    txtHaber = FormatCurrency(ValidarNumeros(txtHaber), , , , vbFalse)
    
    If lstCuentas.ListIndex = -1 Then
        MsgBox "Debe seleccionar una cuenta", vbInformation, "Atención"
        Exit Sub
    End If
    
    If Left(lstCuentas, 1) <> " " Then
        MsgBox "No puede seleccionar cuentas principales", vbInformation, "Atención"
        Exit Sub
    End If
    
    If EstaCuentaEnMatriz(NOASIENTO, CLng(SP(1))) = True Then
        MsgBox "No se permite utilizar esta cuenta para asientos manuales", vbInformation, "Atención"
        Exit Sub
    End If
    
    If CSng(txtDebe) = 0 And CSng(txtHaber) = 0 Or _
        CSng(txtDebe) <> 0 And CSng(txtHaber) <> 0 Then
        MsgBox "Ingrese correctamente el Debe o el Haber", vbInformation, "Atención"
        Exit Sub
    End If
    
    'LISTO AGREGO EL RENGLON------------------------------------------------------
    TmP = lvAsiento.ListItems.Count + 1
    lvAsiento.ListItems.Add
    
    lvAsiento.ListItems(TmP).Text = SP(1)
    lvAsiento.ListItems(TmP).SubItems(1) = SP(2)
    
    If CSng(txtDebe) <> 0 Then
        lvAsiento.ListItems(TmP).SubItems(2) = txtDebe
        lvAsiento.ListItems(TmP).SubItems(3) = ""
    Else
        lvAsiento.ListItems(TmP).SubItems(2) = ""
        lvAsiento.ListItems(TmP).SubItems(3) = txtHaber
    End If
    
    CalcularSumas
    txtDebe = FormatCurrency(0)
    txtHaber = FormatCurrency(0)
End Sub

Private Sub CalcularSumas()
    Dim I As Long
    
    Debe = 0: Haber = 0
    
    For I = 1 To lvAsiento.ListItems.Count
        If IsNumeric(lvAsiento.ListItems(I).SubItems(2)) Then
            Debe = Debe + CSng(lvAsiento.ListItems(I).SubItems(2))
        End If
        If IsNumeric(lvAsiento.ListItems(I).SubItems(3)) Then
            Haber = Haber + CSng(lvAsiento.ListItems(I).SubItems(3))
        End If
    Next I
    
    lblDebe = FormatCurrency(Debe, , , , vbFalse)
    lblHaber = FormatCurrency(Haber, , , , vbFalse)
    lblDif = FormatCurrency(Debe - Haber, , , , vbFalse)
    
End Sub

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

Private Sub cmdAntAsiento_Click()
    Dim IDAsien As Long
    
    IDAsien = CLng(Right(lblAsNro, Len(lblAsNro) - InStrRev(lblAsNro, "Nro ") - 3))
    
    If IDAsien <= 1 Then Exit Sub
    
    MostrarAsiento IDAsien - 1, False
End Sub

Private Sub cmdAsentar_Click()
    If lvAsiento.ListItems.Count = 0 Then Exit Sub
    
    If EsCero(CSng(lblDif)) = False Then
        MsgBox "Sumas de Debe y Haber deben coincidir", vbInformation, "Atención"
        Exit Sub
    End If
    
    'hago el asiento nomas
    Dim Debitos As String, MontosD As String, Creditos As String, MontosC As String
    Dim I As Long, IDAsien As Long, UltAs As Long
    
    IDAsien = CLng(Right(lblAsNro, Len(lblAsNro) - InStrRev(lblAsNro, "Nro ") - 3))
    UltAs = PC.GetUltIDAsientoMasUno - 1
    
    If IDAsien > UltAs Then 'esnuevo
        IDAsien = 0
    Else
        If cmdQuAsiento.Enabled = False Then 'no cambio nada el asiento viejo
            'entonces no hago nada
            Exit Sub
        End If
            
    End If
    
    
    Debitos = "": Creditos = "": MontosD = "": MontosC = ""
    
    For I = 1 To lvAsiento.ListItems.Count
        'veo si no tiene nada en el debe SEGURO(?) tiene en el haber
        If lvAsiento.ListItems(I).SubItems(2) <> "" Then 'DEBE!! !!!!!!!!!!!!!!!!
            If Debitos <> "" Then Debitos = Debitos + "/"
            If MontosD <> "" Then MontosD = MontosD + "/"
            Debitos = Debitos + lvAsiento.ListItems(I).Text
            MontosD = MontosD + CStr(CSng(lvAsiento.ListItems(I).SubItems(2)))
        Else 'HABER!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!1
            If Creditos <> "" Then Creditos = Creditos + "/"
            If MontosC <> "" Then MontosC = MontosC + "/"
            Creditos = Creditos + lvAsiento.ListItems(I).Text
            MontosC = MontosC + CStr(CSng(lvAsiento.ListItems(I).SubItems(3)))
        End If
        
    Next I
    
    'veamos que pasa
    If PC.Asiento(Debitos, MontosD, Creditos, MontosC, , txtDetAsiento, IDAsien) = 0 Then
        MsgBox "Asiento contabilizado correctamente", vbInformation, "Atención"
    Else
        MsgBox "Asiento NO contabilizado compruebe los datos", vbInformation, "Atención"
    End If
    
    lvAsiento.ListItems.Clear
    txtDetAsiento = ""
    CalcularSumas
End Sub

Private Sub cmdEliminar_Click()
    If lstCuentas.ListIndex = -1 Then Exit Sub
    
    Dim SP() As String
    
    SP = Split(lstCuentas, ".")
    
    If EstaCuentaEnMatriz(NOMODIFICAR, CLng(SP(1))) = True Then
        MsgBox "No se permite eliminar esta cuenta", vbInformation, "Atención"
        Exit Sub
    End If
    
    If MsgBox("¿Está seguro de eliminar " + UCase(SP(2)) + vbCrLf + _
        " y sus correspondientes subcuentas?", _
        vbInformation + vbYesNo, "Atención") = vbNo Then Exit Sub
    
    PC.EliminarCuenta CLng(SP(1))
    
    CargarDatos
    
End Sub

Private Sub cmdFind_Click()
    Dim QueA As Long
    
    QueA = ValidarNumeros(txtNroAsiento)
    
    MostrarAsiento QueA, False
End Sub

Private Sub cmdGrabarDetalle_Click()
    If lstCuentas.ListIndex = -1 Then Exit Sub
    
    Dim SP() As String
    SP = Split(lstCuentas, ".")

    PC.ModificarCuenta CLng(SP(1)), , , txtDescripcion
    
End Sub

Private Sub cmdLiAsiento_Click()
    lvAsiento.ListItems.Clear
End Sub

Private Sub cmdMoAsiento_Click()
    Dim nroCta As String, nCta As String
    
    If lvAsiento.ListItems.Count = 0 Then
        MsgBox "No hay nada ingresado que modificar", vbInformation, "Atención"
        Exit Sub
    End If
   
   
    nroCta = InputBox("Ingrese el Numero de Cuenta que desea " + vbCrLf + _
        "reemplazar a " + UCase(lvAsiento.SelectedItem.SubItems(1)), "Modificar Cuenta")
    
    nroCta = Replace(nroCta, ".", ",")
    
    If Not IsNumeric(nroCta) Or InStrRev(nroCta, ",") <> 0 Then
        MsgBox "Debe seleccionar un número entero valido", vbInformation, "Atención"
        Exit Sub
    End If
    
    If CLng(nroCta) < 6 Or PC.GetNameCuenta(CLng(nroCta)) = "NO EXISTE" Then
        MsgBox "Número de cuenta inválido", vbInformation, "Atención"
        Exit Sub
    End If
    
    'bueno cambio nomas no me hago rogar mas
    lvAsiento.ListItems(lvAsiento.SelectedItem.Index).Text = nroCta
    lvAsiento.ListItems(lvAsiento.SelectedItem.Index).SubItems(1) = PC.GetNameCuenta(CLng(nroCta))
    
End Sub

Private Sub cmdModAsiento_Click()
    cmdMoAsiento.Enabled = True
    cmdAgAsiento.Enabled = True
    cmdQuAsiento.Enabled = True
    cmdLiAsiento.Enabled = True
    
    cmdModAsiento.Enabled = False
End Sub

Private Sub cmdModificar_Click()
    Dim nCta As String, nUcTa As String, TmP As String
    Dim SP() As String
    
    If lstCuentas.ListIndex = -1 Then Exit Sub
    SP = Split(lstCuentas, ".")
    
    If EstaCuentaEnMatriz(NOMODIFICAR, CLng(SP(1))) = True Then
        MsgBox "No se permite modificar esta cuenta", vbInformation, "Atención"
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
        TmP = "*"
    Else
        tmpExNombre = SP(2)
        TmP = ""
    End If
    
    nCta = Replace(InputBox("Ingrese el nuevo nombre de la cuenta", "Nombre", SP(2)), ".", " ")
    
    If nCta = "" Then Exit Sub
    nCta = Replace(nCta, "'", " ")
    nCta = Replace(nCta, "*", " ")
    
    'modificar la cuenta
    PC.ModificarCuenta CLng(SP(1)), nUcTa, nCta, txtDescripcion
    
    'actualizo
    lstCuentas.List(lstCuentas.ListIndex) = SP(0) + "." + nUcTa + "." + nCta + TmP
End Sub

Private Sub cmdNueAsiento_Click()
    Debe = 0: Haber = 0
        
    lvAsiento.ListItems.Clear
    txtMonto = FormatCurrency(0)
    lblfecha = FormatDateTime(Date, vbShortDate)
    
    lblDebe = FormatCurrency(Debe)
    lblHaber = FormatCurrency(Haber)
    txtDebe = FormatCurrency(0)
    txtHaber = FormatCurrency(0)
    txtDetAsiento = ""
    txtNroAsiento = PC.GetUltIDAsientoMasUno - 1
    lblAsNro = "Asiento Nro " + CStr(PC.GetUltIDAsientoMasUno)
    
    cmdMoAsiento.Enabled = True
    cmdAgAsiento.Enabled = True
    cmdQuAsiento.Enabled = True
    cmdLiAsiento.Enabled = True
    
    cmdModAsiento.Enabled = False
End Sub

Private Sub cmdQuAsiento_Click()
    Dim Cual As Long
    
    If lvAsiento.ListItems.Count = 0 Then
        MsgBox "No hay nada ingresado que quitar", vbInformation, "Atención"
        Exit Sub
    End If
    
    Cual = lvAsiento.SelectedItem.Index
    lvAsiento.ListItems.Remove (Cual)
    
    CalcularSumas
End Sub

Private Sub cmdSigAsiento_Click()
    Dim IDAsien As Long, UltAs As Long
    
    IDAsien = CLng(Right(lblAsNro, Len(lblAsNro) - InStrRev(lblAsNro, "Nro ") - 3))
    UltAs = PC.GetUltIDAsientoMasUno - 1
    
    If IDAsien >= UltAs Then Exit Sub
    
    MostrarAsiento IDAsien + 1
End Sub

Private Sub cmdVerUltimo_Click()
    Dim UltA As Long
    
    UltA = PC.GetUltIDAsientoMasUno - 1
    
    MostrarAsiento UltA
End Sub

Private Sub MostrarAsiento(IdAsiento As Long, Optional Adelante As Boolean = True)
    Dim ExisteAs As Date, UltIdAs As Long
    
    UltIdAs = PC.GetUltIDAsientoMasUno - 1
    
    ExisteAs = PC.ListarAsientos(IdAsiento, IdAsiento, lvAsiento)
    
    If ExisteAs < #1/2/1900# Then 'no existe el asiento
        If Adelante Then
            If IdAsiento > UltIdAs Then
                Exit Sub
            Else
                MostrarAsiento IdAsiento + 1
                Exit Sub
            End If
        Else
            If IdAsiento < 1 Then
                Exit Sub
            Else
                MostrarAsiento IdAsiento - 1, False
                Exit Sub
            End If
        End If
    End If
    
    If IdAsiento <= PC.GetUltIdCierreResultados Then
        cmdModAsiento.Enabled = False
    Else
        cmdModAsiento.Enabled = True
    End If
    
    cmdMoAsiento.Enabled = False
    cmdAgAsiento.Enabled = False
    cmdQuAsiento.Enabled = False
    cmdLiAsiento.Enabled = False
    
    lblfecha = FormatDateTime(ExisteAs, vbShortDate)
    lblAsNro = "Asiento Nro " + CStr(IdAsiento)
    txtDetAsiento = PC.GetTop1Rs("LibroDiario", "Detalle", , "IdAsiento = " + CStr(IdAsiento), True)
        
    CalcularSumas
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub


Private Sub fBoton8_Click()

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    CargarDatos
    
    lstCuentas.ListIndex = 0
    lstCuentas_Click
    
    cmdNueAsiento_Click
    
    VerMatricesProhibidas
End Sub

Private Sub VerMatricesProhibidas()
    'Ver Inicio formulario

    ReDim NOMODIFICAR(28)
    NOMODIFICAR(0) = "1":    NOMODIFICAR(1) = "2"
    NOMODIFICAR(2) = "3":    NOMODIFICAR(3) = "4"
    NOMODIFICAR(4) = "5":    NOMODIFICAR(5) = "6"
    NOMODIFICAR(7) = "15" 'la 14 (ctas part se agregan junto a los hijos)
    NOMODIFICAR(8) = "16": 'la 17 (Ventas se agregan junto a los hijos)
    NOMODIFICAR(10) = "18":    NOMODIFICAR(11) = "19"
    NOMODIFICAR(12) = "20":    NOMODIFICAR(13) = "22"
    NOMODIFICAR(14) = "23":    NOMODIFICAR(15) = "28"
    NOMODIFICAR(15) = "30"
    NOMODIFICAR(16) = "32":    NOMODIFICAR(17) = "33"
    NOMODIFICAR(18) = "34":    NOMODIFICAR(19) = "35"
    NOMODIFICAR(20) = "37":    NOMODIFICAR(21) = "41"
    NOMODIFICAR(22) = "46":    NOMODIFICAR(23) = "50"
    NOMODIFICAR(24) = "54":    NOMODIFICAR(25) = "57"
    NOMODIFICAR(24) = "68":    NOMODIFICAR(25) = "69"
    NOMODIFICAR(26) = "76":    NOMODIFICAR(27) = "78"
    NOMODIFICAR(28) = "93"
    
    AgregarHijosMatriz NOMODIFICAR, 14 'Hijos de "Cuentas Particulares (socios y emp)"
    AgregarHijosMatriz NOMODIFICAR, 15 'Hijos de "Capital" por los Aportes de Socios
    AgregarHijosMatriz NOMODIFICAR, 17 'Hijos de "Ventas"
    
' Ni hacer asientos con estas cuentas
' 14. Cuentas Particulares
' 46. Clientes
' 41. Proveedores
    ReDim Preserve NOASIENTO(0)
    
    AgregarHijosMatriz NOASIENTO, 14
    AgregarHijosMatriz NOASIENTO, 15
    AgregarHijosMatriz NOASIENTO, 41
    AgregarHijosMatriz NOASIENTO, 46
End Sub

Private Function EstaCuentaEnMatriz(Matriz() As String, IdCuenta As Long) As Boolean
    Dim I As Long, Resp As Boolean

    Resp = False
    
    For I = 0 To UBound(Matriz)
        If CStr(IdCuenta) = Matriz(I) Then
            Resp = True
            Exit For
        End If
    Next I
    
    EstaCuentaEnMatriz = Resp
End Function

Private Sub AgregarHijosMatriz(Matriz() As String, IdCuenta As Long)
    Dim I As Long, UB As Long, Hij() As String
    
    UB = UBound(Matriz) + 1
    
    'primero agrego la principal
    ReDim Preserve Matriz(UB)
    Matriz(UB) = CStr(IdCuenta)
    
    Hij = PC.GetCuentas(IdCuenta)
    
    For I = 1 To UBound(Hij)
        AgregarHijosMatriz Matriz, CLng(Hij(I))
    Next I
End Sub

Private Sub CargarDatos()
    lstCuentas.Clear
    
    lstCuentas.AddItem ".1.Activo"
    lstCuentas.AddItem ".2.Pasivo"
    lstCuentas.AddItem ".3.Perdida"
    lstCuentas.AddItem ".4.Ganancia"
    lstCuentas.AddItem ".5.Patrimonio Neto"
    
End Sub

Private Sub lblDebe_Change()
    CalcularDif
End Sub

Private Sub lblHaber_Change()
    CalcularDif
End Sub

Private Sub CalcularDif()
    lblDif = FormatCurrency(CSng(lblDebe) - CSng(lblHaber), , , , vbFalse)
End Sub

Private Sub lstCuentas_Click()
    If lstCuentas.ListIndex = -1 Then Exit Sub
    
    Dim SP() As String
    SP = Split(lstCuentas, ".")
    
    txtDescripcion = PC.GetDetalle(CLng(SP(1)))
    lblCtaSelec = "Cuenta Seleccionada: " + SP(2)
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

Private Sub txtDebe_GotFocus()
    PintarTxt txtDebe
End Sub

Private Sub txtDebe_LostFocus()
    txtDebe = FormatCurrency(ValidarNumeros(txtDebe), , , , vbFalse)
End Sub

Private Sub txtHaber_GotFocus()
    PintarTxt txtHaber
End Sub

Private Sub txtHaber_LostFocus()
    txtHaber = FormatCurrency(ValidarNumeros(txtHaber), , , , vbFalse)
End Sub

Private Sub txtNroAsiento_GotFocus()
    PintarTxt txtNroAsiento
End Sub

Private Sub txtNroAsiento_LostFocus()
    txtNroAsiento = ValidarNumeros(txtNroAsiento)
End Sub
