VERSION 5.00
Object = "{A7FBD38D-2930-49E3-B60C-9E0202D84549}#15.0#0"; "tbrControles.ocx"
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmConfig 
   BackColor       =   &H004E4E4E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuraciones"
   ClientHeight    =   9030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12030
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConfig.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9030
   ScaleWidth      =   12030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkCodProd 
      BackColor       =   &H004E4E4E&
      Caption         =   "Es Configurable la impresión del código del producto"
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
      Height          =   615
      Left            =   9150
      TabIndex        =   57
      Top             =   3090
      Width           =   2835
   End
   Begin VB.CheckBox chkPant 
      BackColor       =   &H004E4E4E&
      Caption         =   "¿Desea ver en datos Ventas - Compras - Clientes en Pantalla Principal?"
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
      Height          =   735
      Left            =   9150
      TabIndex        =   56
      Top             =   1650
      Width           =   2775
   End
   Begin VB.CheckBox chkEnvases 
      BackColor       =   &H004E4E4E&
      Caption         =   "Utiliza Envases"
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
      Height          =   615
      Left            =   9150
      TabIndex        =   32
      Top             =   2430
      Width           =   2685
   End
   Begin VB.CheckBox chkContado 
      BackColor       =   &H004E4E4E&
      Caption         =   "¿Desea ver en pago contado, distintas opciones?"
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
      Height          =   615
      Left            =   9150
      TabIndex        =   31
      Top             =   990
      Width           =   2745
   End
   Begin VB.ComboBox cmbDiasB 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4740
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   7350
      Width           =   1275
   End
   Begin VB.CheckBox chkBuP 
      Alignment       =   1  'Right Justify
      BackColor       =   &H004E4E4E&
      Caption         =   "Realizar BackUp todos los meses el día"
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
      Height          =   345
      Left            =   1320
      TabIndex        =   26
      Top             =   7350
      Width           =   3255
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H004E4E4E&
      Caption         =   "Forma de Pago Predeterminada"
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
      Height          =   645
      Left            =   2340
      TabIndex        =   47
      Top             =   4590
      Width           =   5085
      Begin VB.OptionButton chkFin 
         BackColor       =   &H004E4E4E&
         Caption         =   "A través de Financiera"
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   2490
         TabIndex        =   48
         Top             =   270
         Width           =   2385
      End
      Begin VB.OptionButton chkCC 
         BackColor       =   &H004E4E4E&
         Caption         =   "Cuenta Corriente"
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   210
         TabIndex        =   16
         Top             =   270
         Value           =   -1  'True
         Width           =   1965
      End
   End
   Begin VB.ComboBox cmbVendedor 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4740
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   4170
      Width           =   2685
   End
   Begin VB.TextBox txtTitulo 
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
      Height          =   555
      Left            =   4740
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmConfig.frx":058A
      Top             =   990
      Width           =   2715
   End
   Begin VB.TextBox txtLetFac 
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
      Height          =   405
      Left            =   4740
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1620
      Width           =   645
   End
   Begin tbrControles.MouTextBox txtSucFac 
      Height          =   405
      Left            =   5400
      TabIndex        =   3
      Top             =   1620
      Width           =   615
      _ExtentX        =   1085
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
      Largo           =   4
      Entero          =   -1  'True
   End
   Begin tbrControles.MouTextBox txtNroFac 
      Height          =   405
      Left            =   6030
      TabIndex        =   4
      Top             =   1620
      Width           =   1215
      _ExtentX        =   2143
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
      Largo           =   8
      Entero          =   -1  'True
   End
   Begin tbrControles.MouTextBox txtDiasV 
      Height          =   405
      Left            =   4740
      TabIndex        =   6
      Top             =   2130
      Width           =   885
      _ExtentX        =   1561
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
      Largo           =   4
      Entero          =   -1  'True
   End
   Begin tbrControles.MouTextBox txtPesos 
      Height          =   405
      Left            =   5250
      TabIndex        =   8
      Top             =   2640
      Width           =   885
      _ExtentX        =   1561
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
      Largo           =   4
      Entero          =   -1  'True
   End
   Begin tbrControles.MouTextBox txtInteres 
      Height          =   405
      Left            =   4740
      TabIndex        =   10
      Top             =   3150
      Width           =   885
      _ExtentX        =   1561
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
      Largo           =   6
   End
   Begin tbrControles.MouTextBox txtCuotas 
      Height          =   405
      Left            =   4740
      TabIndex        =   12
      Top             =   3660
      Width           =   885
      _ExtentX        =   1561
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
      Largo           =   3
      Entero          =   -1  'True
   End
   Begin tbrControles.MouTextBox txtMovProd 
      Height          =   405
      Left            =   4725
      TabIndex        =   18
      Top             =   5310
      Width           =   885
      _ExtentX        =   1561
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
      Largo           =   3
      Entero          =   -1  'True
   End
   Begin tbrControles.MouTextBox txtMargenVenta 
      Height          =   375
      Left            =   4740
      TabIndex        =   22
      Top             =   6330
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   661
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
   Begin tbrControles.MouTextBox txtDtoVta 
      Height          =   375
      Left            =   4740
      TabIndex        =   24
      Top             =   6840
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   661
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
   Begin tbrControles.MouTextBox txtMovAccesos 
      Height          =   405
      Left            =   4725
      TabIndex        =   20
      Top             =   5820
      Width           =   885
      _ExtentX        =   1561
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
      Largo           =   3
      Entero          =   -1  'True
   End
   Begin tbrControles.MouTextBox txtStock 
      Height          =   375
      Left            =   4740
      TabIndex        =   29
      Top             =   7830
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   661
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
   Begin tbrFaroButton.fBoton Command1 
      Height          =   465
      Left            =   10440
      TabIndex        =   58
      Top             =   8430
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdTitulo 
      Height          =   465
      Left            =   7710
      TabIndex        =   1
      Top             =   1020
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Grabar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdNFac 
      Height          =   465
      Left            =   7710
      TabIndex        =   5
      Top             =   1560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Grabar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdDiasV 
      Height          =   465
      Left            =   7710
      TabIndex        =   7
      Top             =   2070
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Grabar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdPesosV 
      Height          =   465
      Left            =   7710
      TabIndex        =   9
      Top             =   2580
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Grabar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdInteres 
      Height          =   465
      Left            =   7710
      TabIndex        =   11
      Top             =   3120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Grabar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdCuotas 
      Height          =   465
      Left            =   7710
      TabIndex        =   13
      Top             =   3630
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Grabar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdVendedor 
      Height          =   465
      Left            =   7710
      TabIndex        =   15
      Top             =   4140
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Grabar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdFDP 
      Height          =   465
      Left            =   7710
      TabIndex        =   17
      Top             =   4650
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Grabar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdMovProd 
      Height          =   465
      Left            =   7710
      TabIndex        =   19
      Top             =   5160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Grabar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdMovAccesos 
      Height          =   465
      Left            =   7710
      TabIndex        =   21
      Top             =   5670
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Grabar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdMargen 
      Height          =   465
      Left            =   7710
      TabIndex        =   23
      Top             =   6180
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Grabar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdDtoVta 
      Height          =   465
      Left            =   7710
      TabIndex        =   25
      Top             =   6690
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Grabar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdBup 
      Height          =   465
      Left            =   7710
      TabIndex        =   28
      Top             =   7200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Grabar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdStock 
      Height          =   465
      Left            =   7710
      TabIndex        =   30
      Top             =   7710
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Grabar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Mínimo Predeterminado"
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
      Left            =   1110
      TabIndex        =   55
      Top             =   7890
      Width           =   3405
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cant. Máximo de días Reg. Accesos"
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
      Height          =   240
      Left            =   1410
      TabIndex        =   54
      Top             =   5940
      Width           =   3195
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "días"
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
      Height          =   240
      Left            =   5640
      TabIndex        =   53
      Top             =   5940
      Width           =   795
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dto. Venta Ctado. Predeterminado"
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
      Left            =   1110
      TabIndex        =   52
      Top             =   6900
      Width           =   3405
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Margen de Venta Predeterminado"
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
      Left            =   1110
      TabIndex        =   51
      Top             =   6390
      Width           =   3405
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Registros"
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
      Height          =   240
      Left            =   5700
      TabIndex        =   50
      Top             =   5430
      Width           =   1335
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cant. Máximo de Mov. Productos"
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
      Height          =   240
      Left            =   1410
      TabIndex        =   49
      Top             =   5430
      Width           =   3195
   End
   Begin VB.Label lblCreditos 
      Alignment       =   2  'Center
      BackColor       =   &H00DAD8D8&
      BackStyle       =   0  'Transparent
      Caption         =   "drtyyeryeytreytreyery"
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
      Left            =   8220
      TabIndex        =   46
      Top             =   90
      UseMnemonic     =   0   'False
      Width           =   3555
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedor Predeterminado"
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
      Height          =   240
      Left            =   1395
      TabIndex        =   45
      Top             =   4260
      Width           =   3195
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad Máxima de Cuotas"
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
      Height          =   240
      Left            =   1395
      TabIndex        =   44
      Top             =   3750
      Width           =   3195
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cuotas"
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
      Height          =   240
      Left            =   5670
      TabIndex        =   43
      Top             =   3780
      Width           =   705
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Interés Mensual Predeterminado"
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
      Height          =   240
      Left            =   1395
      TabIndex        =   42
      Top             =   3240
      Width           =   3195
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "%"
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
      Height          =   240
      Left            =   5670
      TabIndex        =   41
      Top             =   3240
      Width           =   585
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "tbrSoft Stock & Cash"
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
      Height          =   375
      Left            =   720
      TabIndex        =   40
      Top             =   1140
      UseMnemonic     =   0   'False
      Width           =   3855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Título del Programa"
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
      Height          =   240
      Left            =   5100
      TabIndex        =   39
      Top             =   690
      Width           =   2205
   End
   Begin VB.Label lblPesos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Días"
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
      Height          =   240
      Left            =   4500
      TabIndex        =   38
      Top             =   2700
      Width           =   705
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Monto Predeterminado para calcular vuelto"
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
      Height          =   405
      Left            =   210
      TabIndex        =   37
      Top             =   2700
      Width           =   4365
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Días"
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
      Height          =   240
      Left            =   5670
      TabIndex        =   36
      Top             =   2220
      Width           =   705
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Plazo Predeterminado Vencimiento"
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
      Height          =   240
      Left            =   1395
      TabIndex        =   35
      Top             =   2220
      Width           =   3195
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Próximo Nro Factura"
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
      Height          =   240
      Left            =   1395
      TabIndex        =   34
      Top             =   1680
      Width           =   3195
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Configurar tbrStock & Cash"
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
      Height          =   435
      Left            =   3690
      TabIndex        =   33
      Top             =   240
      UseMnemonic     =   0   'False
      Width           =   4335
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkBuP_Click()
    If chkBuP Then
        cmbDiasB.Enabled = True
    Else
        cmbDiasB.Enabled = False
    End If
End Sub

Private Sub chkCodProd_Click()
    If chkCodProd.Value Then
        CFG.ModificarNodo 6, , , , "Si"
    Else
        CFG.ModificarNodo 6, , , , "No"
    End If
End Sub

Private Sub chkContado_Click()
    If chkContado.Value Then
        CFG.ModificarNodo 95, , , , "Si"
    Else
        CFG.ModificarNodo 95, , , , "No"
    End If
End Sub

Private Sub chkEnvases_Click()
    If chkEnvases.Value Then
        CFG.ModificarNodo 5, , , , "Si"
    Else
        CFG.ModificarNodo 5, , , , "No"
    End If
End Sub

Private Sub chkPant_Click()
    If chkPant.Value Then
        CFGBD.ModificarNodo 1, , , , "Si"
    Else
        CFGBD.ModificarNodo 1, , , , "No"
    End If
End Sub

Private Sub cmdBup_Click()
    Dim SP() As String
    
    SP = Split(CFG.GetInfo(18, 4), "_")
    
    If chkBuP Then
        CFG.ModificarNodo 18, , , , cmbDiasB + "_" + SP(1)
    Else
        CFG.ModificarNodo 18, , , , "32_" + SP(1)
    End If
End Sub

'Private Sub cmdCBarras_Click()
'    Dim TmP As String, I As Long, TaJoya As Boolean
'
'    TmP = CFG.GetInfo(15, 4)
'    TaJoya = False
'
'    'que tenga entre 2 y 4 caracteres
'    If Len(txtCBarras) < 2 Or Len(txtCBarras) > 4 Then
'        MsgBox "No cumple con las condiciones", vbInformation, "tbrStock & Cash"
'        txtCBarras = TmP
'        Exit Sub
'    End If
'
'    'si no es numeric pero incluye un caracter numeric ta joya
'
'    If IsNumeric(txtCBarras) Then
'        MsgBox "No cumple con las condiciones", vbInformation, "tbrStock & Cash"
'        txtCBarras = TmP
'        Exit Sub
'    Else
'        'no es numeric - tenemos que encontrar un numero entre los caracteres
'        For I = 1 To Len(txtCBarras)
'            If IsNumeric(Mid(txtCBarras, I, 1)) Then
'                TaJoya = True
'                Exit For
'            End If
'        Next I
'    End If
'
'    If TaJoya = False Then
'        'no encontro numeros adentro
'        MsgBox "No cumple con las condiciones", vbInformation, "tbrStock & Cash"
'        txtCBarras = TmP
'        Exit Sub
'    Else
'        CFG.ModificarNodo 15, , , , txtCBarras
'        MsgBox "Grabado correctamente", vbInformation, "tbrStock & Cash"
'    End If
'End Sub

Private Sub cmdCuotas_Click()
    txtCuotas = ValidarNumeros(txtCuotas)
    CFG.ModificarNodo 12, , , , txtCuotas
End Sub

Private Sub cmdDiasV_Click()
    txtDiasV = ValidarNumeros(txtDiasV)
    CFG.ModificarNodo 2, , , , txtDiasV
End Sub

Private Sub cmdDtoVta_Click()
    If txtDtoVta = "" Then txtDtoVta = FormatPercent(0): Exit Sub
    'no se porque pero le tengo que sacar el %
    Dim SP() As String
    SP = Split(txtDtoVta, "%")
    txtDtoVta = SP(0)
    CFG.ModificarNodo 60, , , , CStr(CSng(txtDtoVta))
    txtDtoVta = FormatPercent(ValidarNumeros(txtDtoVta) / 100)
End Sub

Private Sub cmdFDP_Click()
    Dim TmP As String
    
    If chkFin.Value = True Then
        TmP = "FIN"
    Else
        TmP = "CC" 'predeterminada
    End If
    
    CFG.ModificarNodo 40, , , , TmP
End Sub

Private Sub cmdInteres_Click()
    txtInteres = ValidarNumeros(txtInteres)
    CFG.ModificarNodo 11, , , , txtInteres
End Sub

Private Sub cmdMargen_Click()
    If txtMargenVenta = "" Then txtMargenVenta = FormatPercent(0): Exit Sub
    'no se porque pero le tengo que sacar el %
    Dim SP() As String
    SP = Split(txtMargenVenta, "%")
    txtMargenVenta = SP(0)
    CFG.ModificarNodo 50, , , , CStr(CSng(txtMargenVenta))
    txtMargenVenta = FormatPercent(ValidarNumeros(txtMargenVenta) / 100)
End Sub

Private Sub cmdMovAccesos_Click()
    txtMovAccesos = ValidarNumeros(txtMovAccesos)
    If CLng(txtMovAccesos) > 0 Then CFG.ModificarNodo 16, , , , txtMovAccesos
End Sub

Private Sub cmdMovProd_Click()
    txtMovProd = ValidarNumeros(txtMovProd)
    If CLng(txtMovProd) > 0 Then CFG.ModificarNodo 17, , , , txtMovProd
End Sub

Private Sub cmdNFac_Click()
    txtLetFac = UCase(txtLetFac)
    txtSucFac = String(4 - Len(txtSucFac), "0") + txtSucFac
    txtNroFac = String(8 - Len(txtNroFac), "0") + txtNroFac
    
    'valido nro factura
    If DB.ContarReg("SELECT NroFactura FROM Facturas WHERE NroFactura = '" + _
        txtLetFac + "-" + txtSucFac + "-" + txtNroFac + "'") <> 0 Then
        MsgBox "Factura ya ingresada", vbInformation, "Atención"
        txtLetFac.SetFocus
        Exit Sub
    End If
    
    'grabo la ultima general
    CFGBD.ModificarNodo 13, , , , txtLetFac + "-" + txtSucFac + "-" + txtNroFac
    
    'grabo la ultima para la letra en particular
    Dim IdQ As Long
    
    IdQ = CFGBD.ExistePropiedad("Factura " + txtLetFac)
    
    If IdQ = 0 Then
        CFGBD.AgregarNodo "13", "Factura " + txtLetFac, "", _
            txtSucFac + "-" + txtNroFac, 0
    Else
        CFGBD.ModificarNodo IdQ, 13, , , txtSucFac + "-" + txtNroFac
    End If
End Sub

Private Sub cmdPesosV_Click()
    txtPesos = ValidarNumeros(txtPesos)
    CFG.ModificarNodo 10, , , , txtPesos
End Sub

Private Sub cmdStock_Click()
    txtStock = ValidarNumeros(txtStock)
    CFG.ModificarNodo 19, , , , txtStock
End Sub

Private Sub cmdTitulo_Click()
    If txtTitulo = "" Then Exit Sub
    
    'grabo nomas
    CFG.ModificarNodo 1, , , , "tbrStock & Cash - " + txtTitulo
    
    MsgBox "Registro grabado correctamente, este se observará la próxima " + vbCrLf + _
        "vez que ingrese al sistema", vbInformation, "Atención"
End Sub

Private Sub cmdVendedor_Click()
    CFG.ModificarNodo 14, , , , cmbVendedor
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Dim TmP As String, FAC As String, SP() As String, H As Long
    
    lblPesos = Left(FormatCurrency(1), 2)
    'titulo
    TmP = CFG.GetInfo(1, 4)
    txtTitulo = Right(TmP, Len(TmP) - 18)
    
    'factura
    FAC = CFGBD.GetInfo(13, 4)
    SP = Split(FAC, "-")
    If UBound(SP) < 2 Then
        txtLetFac = "A"
        txtSucFac = "1111"
        txtNroFac = "11111111"
        CFGBD.ModificarNodo 13, , , , txtLetFac + "-" + txtSucFac + "-" + txtNroFac
    Else
        txtLetFac = UCase(SP(0))
        txtSucFac = String(4 - Len(SP(1)), "0") + SP(1)
        txtNroFac = String(8 - Len(SP(2)), "0") + SP(2)
    End If
    
    'plazo
    TmP = CFG.GetInfo(2, 4)
    If Not IsNumeric(TmP) Then TmP = "0"
    If CLng(TmP) < 0 Then TmP = "30"
    txtDiasV = TmP
    CFG.ModificarNodo 2, , , , TmP
    
    'vuelto
    TmP = CFG.GetInfo(10, 4)
    If Not IsNumeric(TmP) Then TmP = "0"
    If CLng(TmP) = 0 Then TmP = "50"
    txtPesos = TmP
    CFG.ModificarNodo 10, , , , CStr(CLng(TmP))
    
    'interes
    TmP = CFG.GetInfo(11, 4)
    If Not IsNumeric(TmP) Then TmP = "0"
    If CLng(TmP) = 0 Then TmP = "3"
    txtInteres = TmP
    CFG.ModificarNodo 11, , , , CStr(CLng(TmP))
    
    'cuotas
    TmP = CFG.GetInfo(12, 4)
    If Not IsNumeric(TmP) Then TmP = "0"
    If CLng(TmP) = 0 Then TmP = "36"
    txtCuotas = TmP
    CFG.ModificarNodo 12, , , , CStr(CLng(TmP))
    
    'vendedor
    CargarVendedores
    TmP = CFG.GetInfo(14, 4)
    If TmP <> "" Then
        If PC.ExisteNCuenta(TmP) <> 0 Then
            cmbVendedor = TmP
        End If
    End If
    CFG.ModificarNodo 14, , , , cmbVendedor
    
    'forma de pago
    TmP = CFG.GetInfo(40, 4)
    If TmP = "FIN" Then
        chkFin.Value = True
    Else 'seria CC por descarte
        chkCC.Value = True
        TmP = "CC"
    End If
    CFG.ModificarNodo 40, , , , TmP
    
    'cantidad de movimientos productos
    TmP = CFG.GetInfo(17, 4)
    If TmP = "" Then TmP = "500"
    If Not IsNumeric(TmP) Then TmP = "500"
    txtMovProd = TmP
    CFG.ModificarNodo 17, , , , TmP
    
    'cantidad de movimientos accesos
    TmP = CFG.GetInfo(16, 4)
    If TmP = "" Then TmP = "30"
    If Not IsNumeric(TmP) Then TmP = "30"
    txtMovAccesos = TmP
    CFG.ModificarNodo 16, , , , TmP
    
    'Margen de Venta
    TmP = CFG.GetInfo(50, 4)
    If TmP = "" Then TmP = "30"
    If Not IsNumeric(TmP) Then TmP = "30"
    txtMargenVenta = FormatPercent(CSng(TmP) / 100)
    CFG.ModificarNodo 50, , , , TmP
    
    'Descuento de Venta Contado
    TmP = CFG.GetInfo(60, 4)
    If TmP = "" Then TmP = "10"
    If Not IsNumeric(TmP) Then TmP = "10"
    txtDtoVta = FormatPercent(CSng(TmP) / 100)
    CFG.ModificarNodo 60, , , , TmP
    
    'Fecha Bup
    cmbDiasB.Clear
    For H = 1 To 30
        cmbDiasB.AddItem CStr(H)
    Next H
    cmbDiasB.ListIndex = 0
    
    TmP = CFG.GetInfo(18, 4)
    If TmP = "" Then TmP = "5_05/06/2007"
    If InStrRev(TmP, "_") = 0 Then TmP = "5_05/06/2007"
    SP = Split(TmP, "_")
    If Not IsNumeric(SP(0)) Then SP(0) = "5"
    
    If CLng(SP(0)) < 1 Or CLng(SP(0)) > 30 Then
        chkBuP.Value = 0
        cmbDiasB.Enabled = False
    Else
        cmbDiasB.Enabled = True
        cmbDiasB = SP(0)
        chkBuP.Value = 1
    End If
    CFG.ModificarNodo 18, , , , SP(0) + "_" + SP(1)
    
    'Stock Minimo Predeterminado
    TmP = CFG.GetInfo(19, 4)
    If TmP = "" Then TmP = "10"
    If Not IsNumeric(TmP) Then TmP = "10"
    txtStock = TmP
    CFG.ModificarNodo 19, , , , TmP
    
    'muestra otra forma de pagos contado??
    TmP = CFG.GetInfo(95, 4)
    If TmP = "" Then TmP = "No"
    If TmP = "Si" Then
        chkContado.Value = 1
    Else
        chkContado.Value = 0
        CFG.ModificarNodo 95, , , , "No"
    End If
    
    'muestra siquiere ver datos en pantalla
    TmP = CFGBD.GetInfo(1, 4)
    If TmP = "" Then TmP = "Si"
    If TmP = "Si" Then
        chkPant.Value = 1
    Else
        chkPant.Value = 0
        CFGBD.ModificarNodo 1, , , , "No"
    End If
    
    'muestra si usa envases
    TmP = CFG.GetInfo(5, 4)
    If TmP = "" Then TmP = "No"
    If TmP = "Si" Then
        chkEnvases.Value = 1
    Else
        chkEnvases.Value = 0
        CFG.ModificarNodo 5, , , , "No"
    End If
    
    'muestra si configura el codigo del producto
    TmP = CFG.GetInfo(6, 4)
    If TmP = "" Then TmP = "No"
    If TmP = "Si" Then
        chkCodProd.Value = 1
    Else
        chkCodProd.Value = 0
        CFG.ModificarNodo 6, , , , "No"
    End If
    
    'Carácteres de C.Barras (solo entre 2 y 4 con nros y letras)
    TmP = CFG.GetInfo(15, 4)
    txtCBarras = TmP
    
    'datos de la version
    lblCreditos = "Versión tbrStock & Cash " + CStr(App.Major) + "." + CStr(App.Minor) + "." + _
        String(4 - Len(CStr(App.Revision)), "0") + CStr(App.Revision)

End Sub

'Private Sub txtCBarras_GotFocus()
'    PintarTxt txtCBarras
'End Sub

Private Sub txtCuotas_GotFocus()
    PintarTxt txtCuotas
End Sub

Private Sub txtDiasV_GotFocus()
    PintarTxt txtDiasV
End Sub

Private Sub txtDtoVta_GotFocus()
    PintarTxt txtDtoVta
End Sub

Private Sub txtDtoVta_LostFocus()
    'no se porque pero le tengo que sacar el %
    If txtDtoVta = "" Then txtDtoVta = FormatPercent(0): Exit Sub
    Dim SP() As String
    SP = Split(txtDtoVta, "%")
    txtDtoVta = SP(0)
    txtDtoVta = FormatPercent(ValidarNumeros(txtDtoVta) / 100)
End Sub
Private Sub txtInteres_GotFocus()
    PintarTxt txtInteres
End Sub

Private Sub txtLetFac_Change()
    txtLetFac = UCase(Left(txtLetFac, 1))
End Sub

Private Sub txtLetFac_GotFocus()
    PintarTxt txtLetFac
End Sub

Private Sub txtLetFac_LostFocus()
    Dim SP() As String, IDp As Long, NmPr As String
    'veo en esa letra cual es la última que uso
    IDp = CFG.ExistePropiedad("Factura " + txtLetFac)
    If IDp <> 0 Then
        NmPr = CFG.GetInfo(IDp, 4)
        If InStrRev(NmPr, "-") = 0 Then NmPr = NmPr + "-1"
    Else
        NmPr = DB.GetTop1Rs("Ventas", "IdVenta", , "IdVenta LIKE '%" + _
            txtLetFac + "-" + "%'", True)
        
        If NmPr = "" Then
            NmPr = "1-1"
        Else
            NmPr = Right(NmPr, Len(NmPr) - 2)
        End If
    End If
    
    SP = Split(NmPr, "-")
    txtSucFac = String(4 - Len(SP(0)), "0") + SP(0)
    txtNroFac = String(8 - Len(SP(1)), "0") + SP(1)
End Sub

Private Sub txtMargenVenta_GotFocus()
    PintarTxt txtMargenVenta
End Sub

Private Sub txtMargenVenta_LostFocus()
    If txtMargenVenta = "" Then txtMargenVenta = FormatPercent(0): Exit Sub
    'no se porque pero le tengo que sacar el %
    Dim SP() As String
    SP = Split(txtMargenVenta, "%")
    txtMargenVenta = SP(0)
    txtMargenVenta = FormatPercent(ValidarNumeros(txtMargenVenta) / 100)
End Sub

Private Sub txtnroFac_GotFocus()
    PintarTxt txtNroFac
End Sub

Private Sub txtnroFac_LostFocus()
    txtNroFac = ValidarNumeros(txtNroFac)
    txtNroFac = String(8 - Len(txtNroFac), "0") + txtNroFac
End Sub

Private Sub txtPesos_GotFocus()
    PintarTxt txtPesos
End Sub

Private Sub txtSucFac_GotFocus()
    PintarTxt txtSucFac
End Sub

Private Sub txtSucFac_LostFocus()
    txtSucFac = ValidarNumeros(txtSucFac)
    txtSucFac = String(4 - Len(txtSucFac), "0") + txtSucFac
End Sub

Private Sub txtTitulo_GotFocus()
    PintarTxt txtTitulo
End Sub

Private Sub CargarVendedores()
    Dim IdCuentas() As String
    
    IdCuentas = PC.GetCuentas(53)
    
    cmbVendedor.Clear
    
    Dim I As Long
    
    For I = 1 To UBound(IdCuentas)
        cmbVendedor.AddItem PC.GetNameCuenta(CLng(IdCuentas(I)))
    Next I
    
    If cmbVendedor.ListCount > 0 Then cmbVendedor.ListIndex = 0
End Sub
