VERSION 5.00
Object = "{A7FBD38D-2930-49E3-B60C-9E0202D84549}#16.0#0"; "tbrControles.ocx"
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCompras 
   BackColor       =   &H00544B45&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Compras"
   ClientHeight    =   8490
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
   Icon            =   "frmCompras2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin tbrFaroButton.fBoton cmdSalir 
      Height          =   435
      Left            =   10200
      TabIndex        =   45
      Top             =   7920
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   "Arial Narrow"
      fFontS          =   9
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdGrabar 
      Height          =   435
      Left            =   8610
      TabIndex        =   44
      Top             =   7920
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Grabar Pedido"
      fEnabled        =   -1  'True
      fFontN          =   "Arial Narrow"
      fFontS          =   9
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdPagar 
      Height          =   435
      Left            =   7020
      TabIndex        =   43
      Top             =   7920
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Terminar (F1)"
      fEnabled        =   -1  'True
      fFontN          =   "Arial Narrow"
      fFontS          =   9
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdBorrarConc 
      Height          =   465
      Left            =   10620
      TabIndex        =   42
      Top             =   4440
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Borrar"
      fEnabled        =   -1  'True
      fFontN          =   "Arial Narrow"
      fFontS          =   9
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdConceptos 
      Height          =   495
      Left            =   6780
      TabIndex        =   41
      Top             =   4920
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Otros Conceptos"
      fEnabled        =   -1  'True
      fFontN          =   "Arial Narrow"
      fFontS          =   9
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdSel 
      Height          =   495
      Left            =   1350
      TabIndex        =   40
      Top             =   6660
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   873
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Ok"
      fEnabled        =   -1  'True
      fFontN          =   "Arial Narrow"
      fFontS          =   9
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdAgregar 
      Height          =   405
      Left            =   5130
      TabIndex        =   7
      Top             =   6060
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   ">>>"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdAgProd 
      Height          =   495
      Left            =   3750
      TabIndex        =   39
      Top             =   4440
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   873
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Agregar Producto"
      fEnabled        =   -1  'True
      fFontN          =   "Arial Narrow"
      fFontS          =   9
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdQuitar 
      Height          =   405
      Left            =   5250
      TabIndex        =   38
      Top             =   3030
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "<<<"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   9
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton Command2 
      Height          =   495
      Left            =   3750
      TabIndex        =   37
      Top             =   2100
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   873
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Agregar Proveedor"
      fEnabled        =   -1  'True
      fFontN          =   "Arial Narrow"
      fFontS          =   9
      fECol           =   5717301
   End
   Begin VB.CheckBox chkAjustarPrecio 
      BackColor       =   &H00544B45&
      Caption         =   "Ajustar Precio según Margen de Venta configurada"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   150
      TabIndex        =   35
      Top             =   7650
      Width           =   2805
   End
   Begin tbrControles.MouTextBox txtCant 
      Height          =   465
      Left            =   3870
      TabIndex        =   5
      Top             =   5550
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   820
      Alignment       =   2
      BackColor       =   16777215
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Largo           =   20
      Entero          =   -1  'True
   End
   Begin tbrControles.tbrBuscador tbrBuscador1 
      Height          =   3435
      Left            =   150
      TabIndex        =   4
      Top             =   3090
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   6059
      BackColor       =   5524293
      BeginProperty Fontt {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
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
   Begin VB.ComboBox cmbSucursales 
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
      Left            =   3390
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   7860
      Width           =   2505
   End
   Begin VB.TextBox txtLetFac 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   7590
      TabIndex        =   0
      Top             =   690
      Width           =   435
   End
   Begin VB.CheckBox chkEntregado 
      BackColor       =   &H00544B45&
      Caption         =   "Mercaderia se entregó"
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   3360
      TabIndex        =   25
      Top             =   7500
      Value           =   1  'Checked
      Width           =   3195
   End
   Begin VB.CheckBox chkIVA 
      BackColor       =   &H00544B45&
      Caption         =   "IVA "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   345
      Left            =   6390
      TabIndex        =   11
      Top             =   6270
      Width           =   1035
   End
   Begin VB.CheckBox chkRecargo 
      BackColor       =   &H00544B45&
      Caption         =   "Recargo"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   345
      Left            =   6360
      TabIndex        =   9
      Top             =   5790
      Width           =   1305
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00544B45&
      Caption         =   "¿Cómo lo incluyo?"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   885
      Left            =   9360
      TabIndex        =   14
      Top             =   5070
      Visible         =   0   'False
      Width           =   2085
      Begin VB.OptionButton chkCosto 
         BackColor       =   &H00544B45&
         Caption         =   "Al Costo"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   210
         TabIndex        =   20
         Top             =   270
         Value           =   -1  'True
         Width           =   1545
      End
      Begin VB.OptionButton chkGasto 
         BackColor       =   &H00544B45&
         Caption         =   "Como Gasto"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   210
         TabIndex        =   21
         Top             =   540
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00544B45&
      Caption         =   "Forma de Pago"
      ForeColor       =   &H00FFFFFF&
      Height          =   885
      Left            =   6660
      TabIndex        =   32
      Top             =   6810
      Width           =   1665
      Begin VB.OptionButton chkACuenta 
         BackColor       =   &H00544B45&
         Caption         =   "A Cuenta"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   210
         TabIndex        =   24
         Top             =   540
         Width           =   1305
      End
      Begin VB.OptionButton chkContado 
         BackColor       =   &H00544B45&
         Caption         =   "Contado"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   210
         TabIndex        =   23
         Top             =   270
         Value           =   -1  'True
         Width           =   1305
      End
   End
   Begin MSComctlLib.ListView lvFactura 
      Height          =   3135
      Left            =   6240
      TabIndex        =   31
      Top             =   1290
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   5530
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
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
         Text            =   "Cant"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Producto"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Pr.Unit"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Pr.Total"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.ComboBox cmbProveedores 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   270
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2130
      Width           =   3255
   End
   Begin MSComCtl2.DTPicker DTFecha 
      Height          =   345
      Left            =   1740
      TabIndex        =   8
      Top             =   930
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   124125185
      CurrentDate     =   39197
   End
   Begin tbrControles.MouTextBox txtSucFac 
      Height          =   405
      Left            =   8040
      TabIndex        =   1
      Top             =   690
      Width           =   525
      _ExtentX        =   926
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
      Largo           =   4
      Entero          =   -1  'True
   End
   Begin tbrControles.MouTextBox txtNroFac 
      Height          =   405
      Left            =   8580
      TabIndex        =   2
      Top             =   690
      Width           =   1215
      _ExtentX        =   2143
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
      Largo           =   8
      Entero          =   -1  'True
   End
   Begin tbrControles.MouTextBox txtPrecio 
      Height          =   465
      Left            =   3840
      TabIndex        =   6
      Top             =   6540
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   820
      Alignment       =   2
      BackColor       =   16777215
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Largo           =   20
   End
   Begin tbrControles.MouTextBox txtDeMas 
      Height          =   420
      Left            =   7770
      TabIndex        =   10
      Top             =   5730
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   741
      Alignment       =   2
      BackColor       =   16777215
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Largo           =   20
   End
   Begin tbrControles.MouTextBox txtIVAPorC 
      Height          =   420
      Left            =   7770
      TabIndex        =   12
      Top             =   6210
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   741
      Alignment       =   2
      BackColor       =   16777215
      Text            =   "0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Largo           =   20
   End
   Begin MSComctlLib.ListView lvConc2 
      Height          =   1125
      Left            =   6240
      TabIndex        =   36
      Top             =   3750
      Visible         =   0   'False
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   1984
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
         Text            =   "IdCta"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Concepto"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Importe"
         Object.Width           =   2170
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "NroConcepto"
         Object.Width           =   0
      EndProperty
   End
   Begin tbrControles.MouTextBox txtIVAPesos 
      Height          =   420
      Left            =   9210
      TabIndex        =   13
      Top             =   6210
      Visible         =   0   'False
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   741
      Alignment       =   2
      BackColor       =   16777215
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Largo           =   20
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
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
      Left            =   570
      TabIndex        =   15
      Top             =   960
      Width           =   1005
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nro Factura"
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
      Left            =   6180
      TabIndex        =   16
      Top             =   750
      Width           =   1275
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   8730
      TabIndex        =   17
      Top             =   6330
      Width           =   345
   End
   Begin VB.Label lblSelec 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Producto Seleccionado"
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
      Height          =   675
      Left            =   3630
      TabIndex        =   19
      Top             =   3540
      Width           =   2535
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "$ 5888,88"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9180
      TabIndex        =   22
      Top             =   7050
      Width           =   1875
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A Pagar"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   345
      Left            =   9480
      TabIndex        =   34
      Top             =   6720
      Width           =   1245
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Para descuentos ingréselo negativo"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   6330
      TabIndex        =   33
      Top             =   5490
      Width           =   2865
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   4020
      TabIndex        =   30
      Top             =   5250
      Width           =   1305
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Precio Total"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   0
      Left            =   3840
      TabIndex        =   29
      Top             =   6240
      Width           =   1275
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione Producto (Nombre o Código)"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   150
      TabIndex        =   28
      Top             =   2760
      Width           =   4305
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione Proveedor"
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
      Left            =   300
      TabIndex        =   27
      Top             =   1770
      Width           =   1905
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Factura de Compras a Wenceslao Carreofur Wanchope Gomez."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   270
      TabIndex        =   18
      Top             =   120
      Width           =   10995
   End
End
Attribute VB_Name = "frmCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ToTal As Single 'es el total de la factura
Dim Precio As Single 'precio de prods a cargar en factura
Dim DeMas As Single 'Es el recargo de la factura
Dim IVApesos As Single
Dim DtDeMas As String 'el detalle de la recarga
Dim nProveedor As String
Dim tmpProv2 As String 'para cuando agregue productos no se pierda el prov elegido
Dim Nuevo As Boolean 'cuando pague quiero saber si ya estaba grabado como pedido
Dim CtoTmp As Single
Dim ActuTbrBP As Boolean

Private Sub Label3_Click()

End Sub

Private Sub Limpiar()
    ToTal = 0 '???? no se porque pero se guarda el valor de la factura vieja
    lvFactura.ListItems.Clear
    txtDeMas = FormatCurrency(0)
    txtPrecio = FormatCurrency(0)
    lblTotal = FormatCurrency(0)
    txtCant = 1
    lblSelec = ""
    txtLetFac = CFG.GetInfo(8, 4)
End Sub

Private Sub chkIVA_Click()
    If chkIVA Then
        txtIVAPorC.Enabled = True
        txtIVAPesos.Visible = True
        CalcularIVA
    Else
        txtIVAPorC = "0"
        txtIVAPorC.Enabled = False
        txtIVAPesos.Visible = False
    End If
End Sub

Private Sub CalcularIVA()
    Dim tmpIvaPorC As Single
    
    If IsNumeric(txtIVAPorC) = True Then
        tmpIvaPorC = CSng(txtIVAPorC)
    Else
        tmpIvaPorC = 0
    End If
    
    'subtotal y desc y ivaporc tambien seguro tienen que ser validos
    txtIVAPesos = FormatCurrency((ToTal + DeMas) * CSng(tmpIvaPorC) / 100, , , , vbFalse)
    
    CalcularTotal
End Sub

Private Sub chkRecargo_Click()
    If chkRecargo Then
        Label7.Visible = True
        txtDeMas.Visible = True
        txtDeMas = FormatCurrency(0)
        Frame2.Visible = True
    Else
        Label7.Visible = False
        txtDeMas.Visible = False
        Frame2.Visible = False
    End If
End Sub

Private Sub cmbProveedores_Click()
    nProveedor = cmbProveedores
    lblTitulo = "Factura de Compra de " + nProveedor
End Sub

Private Sub cmdAgProd_Click()
    tmpProv2 = cmbProveedores.Text
    frmProductos.AbrirDatos ""
End Sub

Private Sub cmdAgregar_Click()
    On Local Error GoTo errAddPr
        
    Terr.AnOtaR "caba"
        
    If tbrBuscador1.GetLstSel = "" Then
        MsgBox "No eligio ningún producto"
        Exit Sub
    End If
    
    Terr.AnOtaR "cabb", txtPrecio.Text, txtCant.Text
    
    Precio = ValidarNumeros(txtPrecio)
    txtPrecio = FormatCurrency(Precio, , , , vbFalse)
    
    Terr.AnOtaR "cabc", Precio
    
    Select Case Precio
        Case 0
            If MsgBox("Va a incluir un producto con precio 0, si no es asi presione " + _
                "Cancelar", vbOKCancel + vbInformation, "Atención") = vbCancel Then Exit Sub
        Case Is < 0
            MsgBox "Ingrese sólo valores positivos", vbInformation, "Atención"
            Exit Sub
    End Select
    
    Terr.AnOtaR "cabd", lvFactura.ListItems.Count + 1, txtCant.Text
    
    Dim TmP As Long
            
    TmP = lvFactura.ListItems.Count + 1
    
    lvFactura.ListItems.Add TmP
    
    lvFactura.ListItems(TmP).Text = txtCant
    lvFactura.ListItems(TmP).SubItems(1) = tbrBuscador1.GetLstSel(1)
    lvFactura.ListItems(TmP).SubItems(2) = FormatCurrency(CStr(Precio / CSng(txtCant)), 4, , , vbFalse)
    lvFactura.ListItems(TmP).SubItems(3) = FormatCurrency(Precio, 4, , , vbFalse)
    
    Terr.AnOtaR "cabe"
    
    CalcularIVA
    
    Terr.AnOtaR "cabf"
    
    cmdSel.Default = True
    txtCant = "1"
    txtPrecio = FormatCurrency(0)
    tbrBuscador1.Text = ""
    tbrBuscador1.SetFocus
    
    Terr.AnOtaR "cabg"
    
    Exit Sub
    
errAddPr:
    Terr.AppendLog "ErrAddPrCPR", Terr.ErrToTXT(Err)
    MsgBox "Error al agregar el producto, envie registro a tbrSoft"
    
End Sub

Private Sub CalcularTotal()

    On Local Error GoTo errCALC

    Dim I As Long
    
    ToTal = 0
    
    For I = 1 To lvFactura.ListItems.Count
        Terr.AnOtaR "caak", lvFactura.ListItems(I).SubItems(3)
        ToTal = ToTal + CSng(lvFactura.ListItems(I).SubItems(3))
    Next
    
    Terr.AnOtaR "caal", ToTal
    
    If txtDeMas.Visible = True Then
        If IsNumeric(txtDeMas) Then
            Terr.AnOtaR "caam", txtDeMas.Text
            DeMas = CSng(txtDeMas)
        Else
            DeMas = 0
        End If
    Else
        DeMas = 0
    End If
    
    Terr.AnOtaR "caan", DeMas, txtIVAPesos.Text, txtIVAPesos.Visible
    
    If txtIVAPesos.Visible = True Then
        If IsNumeric(txtIVAPesos) Then
            IVApesos = CSng(txtIVAPesos)
        Else
            IVApesos = 0
        End If
    Else
        IVApesos = 0
    End If
    
    Terr.AnOtaR "caao", IVApesos, lvConc2.Visible
    
    If lvConc2.Visible = True Then
        For I = 1 To lvConc2.ListItems.Count
            Terr.AnOtaR "caap", lvConc2.ListItems(I).SubItems(2), ToTal
            ToTal = ToTal + CSng(lvConc2.ListItems(I).SubItems(2))
        Next I
    End If
    
    Terr.AnOtaR "caaq", ToTal, DeMas, IVApesos
    
    lblTotal = FormatCurrency(ToTal + DeMas + IVApesos, , , , vbFalse)
    
    Terr.AnOtaR "caar", lblTotal.Caption
    
    Exit Sub
errCALC:
    Terr.AppendLog "ErrCalcCPR", Terr.ErrToTXT(Err)
    MsgBox "Error al calcular la compra, envie registro a tbrSoft"
    
End Sub

Private Sub cmdBorrarConc_Click()
    lvConc2.ListItems.Clear
    lvConc2.Visible = False
    cmdBorrarConc.Visible = False
    lvFactura.Height = 3150
    CalcularIVA
End Sub

Private Sub cmdConceptos_Click()
    frmOtrosConc.AbrirDatos False
    
    If lvConc2.ListItems.Count > 0 Then
        lvFactura.Height = 2380
        cmdBorrarConc.Visible = True
        lvConc2.Visible = True
    Else
        cmdBorrarConc_Click
    End If
    
    CalcularTotal
End Sub

Private Sub cmdGrabar_Click()
    'muy parecido a pagar
    If lvFactura.ListItems.Count <= 0 Then Exit Sub
    If ToTal = 0 Then Exit Sub
    
    'no le grabo el recargo que lo ponga de vuelta el otario - NI EL IVA
    If MsgBox("¿Está seguro de grabar el pedido de " + UCase(nProveedor) + "?", _
        vbOKCancel, "Atención") = vbCancel Then Exit Sub
    
    '1ro lo agrego en FacturaCompra tildado el campo EsPedido
    DB.EXECUTE "INSERT INTO FacturaCompra (NroFactura,Fecha,Proveedor,Pagado," + _
        "Entregado, EsPedido) " + "VALUES ('" + GetNroFac + "', #" + _
        stFechaSQL(DTFecha) + "#,'" + nProveedor + _
        "'," + Replace(CStr(ToTal), ",", ".") + ", 0, 1)"
    
    Dim clsC As New clsProducto
    Dim IDp As Long, Cant As Long, PTot As Single
     
    For I = 1 To lvFactura.ListItems.Count
        IDp = DB.GetValInRS("Productos", "ID", "nProducto = '" + _
            lvFactura.ListItems(I).SubItems(1) + "'", False)
        Cant = CLng(lvFactura.ListItems(I).Text)
        PTot = CSng(lvFactura.ListItems(I).SubItems(3))
                
        clsC.CargarCompraDt GetNroFac, DTFecha, nProveedor, IDp, Cant, PTot
    Next I
    
    Set clsC = Nothing
    
    Unload Me
End Sub

Private Sub cmdPagar_Click()
    On Local Error GoTo errOkCPR
    
    Terr.AnOtaR "cabp", lvFactura.ListItems.Count
    
    If lvFactura.ListItems.Count <= 0 Then Exit Sub
    
    Dim NroAsiento As Long
    Dim TmP As String
    Dim DifP As Single 'es el porcentaje del precio que queda
    Dim clsC As New clsProducto
    Dim IDp As Long, PrecioAj As Single, Cant As Long
    Dim stDeb As String, DebP As String
    
    Terr.AnOtaR "cabq", chkContado.Value
    
    If chkContado Then
        TmP = "Al Contado"
    Else
        TmP = "A Cuenta"
    End If
    
    Terr.AnOtaR "cabr"
    NroAsiento = PC.GetUltIDAsientoMasUno("LibroSubDiario")
    DeMas = ValidarNumeros(txtDeMas)
    IVApesos = CSng(txtIVAPesos)
    
    Terr.AnOtaR "cabs", NroAsiento, DeMas, IVApesos
    'valido nro factura
    If DB.ContarReg("SELECT NroFactura FROM FacturaCompra WHERE NroFactura = '" + _
        GetNroFac + "'") <> 0 Then
        
        Terr.AnOtaR "cabu"
        
        MsgBox "Factura ya ingresada", vbInformation, "Atención"
        txtLetFac.SetFocus
        Exit Sub
    End If
    
    Terr.AnOtaR "cabt"
    
    If MsgBox("Está por registrar la compra " + TmP + " de " + _
        FormatCurrency(ToTal + DeMas + IVApesos) + _
        " a " + UCase(nProveedor) + ", ¿Son correctos los datos?", _
        vbOKCancel + vbInformation, "Atencion") = vbCancel Then Exit Sub
    
    Terr.AnOtaR "cabv"
    
    If txtDeMas.Visible = False Or chkCosto = False Then
        'esto pasa cuando no agrego a la factura la recarga
        'o cuando lo agrego como un gasto aparte (chkcosto=false)
        DifP = 1
    Else
        DifP = 1 + DeMas / (ToTal)
    End If
    
    Terr.AnOtaR "cabw", DifP
    
     'ahora vamos con los numerajes!!!!!!!!
    
    If chkGasto = True Then 'eligio mandarlo como gasto aparte
    
        Terr.AnOtaR "cabx"
        
        stDeb = stDeb + "/30"
        DebP = CStr(ToTal) + "/" + CStr(DeMas)
        'registro la compra (el gasto va aparte de la compra)
        DB.EXECUTE "INSERT INTO FacturaCompra (NroFactura,Fecha,Proveedor,Pagado," + _
            "Entregado, EsPedido) " + "VALUES ('" + GetNroFac + "', #" + _
            stFechaSQL(DTFecha) + "#,'" + nProveedor + _
            "'," + Replace(CStr(ToTal + IVApesos), ",", ".") + ", " + _
            CStr(chkEntregado) + ", 0)"
    Else
    
        Terr.AnOtaR "caby"
    
        'registro la compra
        DB.EXECUTE "INSERT INTO FacturaCompra (NroFactura,Fecha,Proveedor,Pagado," + _
            "Entregado, EsPedido) " + "VALUES ('" + GetNroFac + "', #" + _
            stFechaSQL(DTFecha) + "#,'" + nProveedor + _
            "'," + Replace(CStr(ToTal + DeMas + IVApesos), ",", ".") + ", " + _
            CStr(chkEntregado) + ", 0)"
    End If
    
    Terr.AnOtaR "cabz"
    
    For I = 1 To lvFactura.ListItems.Count
    
        Terr.AnOtaR "caca", I, lvFactura.ListItems(I).SubItems(1)
    
        IDp = DB.GetValInRS("Productos", "ID", "nProducto = '" + _
            lvFactura.ListItems(I).SubItems(1) + "'", False)
        PrecioAj = CSng(lvFactura.ListItems(I).SubItems(2)) * DifP
        Cant = CLng(lvFactura.ListItems(I).Text)
        'modifico stock y costo (ademas acomoda stock por sucursal)
        '(modifico por mas que la mercadería no haya sido entregada) XXXX
        clsC.CargarCompra IDp, Cant, PrecioAj, cmbSucursales, _
            "Por Compras Factura Nº " + CStr(GetNroFac), -chkAjustarPrecio.Value
        
        'cargo el detalle de compra
        clsC.CargarCompraDt GetNroFac, DTFecha, nProveedor, IDp, Cant, Cant * PrecioAj
    Next I
    
    Terr.AnOtaR "cacb"
    
    'si tiene iva --------------------------------------------------------------------
    Dim Iva As Single
    If txtIVAPesos.Visible = False Then
        Iva = 0
    Else
        Iva = CSng(txtIVAPesos)
        'agrego renglon para el iva
        DB.EXECUTE "INSERT INTO CompraDetalle (ID, Idproducto,Fecha, Cantidad, " + _
            "PrecioTotal, NroFactura, Proveedor) " + _
            "VALUES (" + IdAutonum("CompraDetalle") + ",-2,#" + stFechaSQL(Date) + _
            "#,1, " + Replace(CStr(Iva), ",", ".") + ",'" + CStr(GetNroFac) + "', '" + _
            cmbProveedores + "')"
    End If
    
    Terr.AnOtaR "cacc", Iva
    
    ' si tiene otros conceptos ------------------------------------------------------
    Dim strAs As String, strAs2 As String, Monto As Single
    strAs = "": strAs2 = "": Monto = 0
    If lvConc2.Visible = True Then
        For I = 1 To lvConc2.ListItems.Count
        
            Terr.AnOtaR "cacj", I, lvConc2.ListItems(I).Text, lvConc2.ListItems(I).SubItems(2)
        
            strAs = strAs + lvConc2.ListItems(I).Text 'con los nros de cuentas
            strAs2 = strAs2 + lvConc2.ListItems(I).SubItems(2) 'con los montos
            Monto = Monto + CSng(lvConc2.ListItems(I).SubItems(2))
            
            clsC.CargarCompraDt GetNroFac, DTFecha, nProveedor, _
                CLng("-1" + lvConc2.ListItems(I).SubItems(3)), _
                0, CSng(lvConc2.ListItems(I).SubItems(2))
            
            If I < lvConc2.ListItems.Count Then
                strAs = strAs + "/"
                strAs2 = strAs2 + "/"
            End If
        Next I
        'hago el asiento nomas
        
        PC.Asiento strAs, strAs2, "78", CStr(Monto), "LibroSubDiario", , NroAsiento
    End If
    
    Terr.AnOtaR "cack"
    '--------------------------------------------------------------------------------
    
    Set clsC = Nothing
    
    'libro diario mercaderia a caja si es a cuenta se invierte en forma de pago
    'si no carga a costo el recargo va como 30: Gastos o Descuentos por Compras
    
    stDeb = "54"
    DebP = CStr(ToTal + DeMas - Monto)
    
    Terr.AnOtaR "cacl", DebP
    
    If txtIVAPesos.Visible = True Then
        stDeb = stDeb + "/50"
        DebP = DebP + "/" + CStr(IVApesos)
    End If
    
    Terr.AnOtaR "cacm"
    PC.Asiento stDeb, DebP, "78", CStr(ToTal + DeMas + IVApesos - Monto), "LibroSubDiario", _
        "Compra Factura Nº " + CStr(GetNroFac), -NroAsiento
    'negativo para que agregue este a un asiento ya abierto
    
    If chkACuenta Then frmClientesMov.AbrirDatos GetNroFac, True, cmbProveedores
    If chkContado Then
        If CFG.GetInfo(95, 4) = "Si" Then frmPago.AbrirDatos ToTal + DeMas + IVApesos, False, "Compras Factura Nº " + CStr(GetNroFac)
    End If
    
    Terr.AnOtaR "cacn"
    Unload Me

    Exit Sub
errOkCPR:
    Terr.AppendLog "Err-OKCPR", Terr.ErrToTXT(Err)
    MsgBox "Error al registrar compra, envie registro a tbrSoft"

End Sub

Private Function GetNroFac() As String
    txtSucFac = String(4 - Len(txtSucFac), "0") + txtSucFac
    txtNroFac = String(8 - Len(txtNroFac), "0") + txtNroFac
    
    GetNroFac = UCase(txtLetFac) + "-" + txtSucFac + "-" + txtNroFac + _
        " " + UCase(cmbProveedores)
End Function


Private Sub cmdQuitar_Click()
    Dim Cual As Long
      
    If lvFactura.ListItems.Count = 0 Then Exit Sub
    
    Cual = lvFactura.SelectedItem.Index
    lvFactura.ListItems.Remove (Cual)
    
    CalcularIVA
    tbrBuscador1.SetFocus
End Sub

Private Sub cmdSalir_Click()
    If lvFactura.ListItems.Count > 0 Then
        If MsgBox("¿Está seguro que desea salir sin registrar la compra o " + _
            "grabar el pedido?." + vbCrLf + " Si presiona ACEPTAR Los datos " + _
            "se perderán definitivamenTE", vbOKCancel + vbInformation, _
            "Atención") = vbCancel Then Exit Sub
    End If
    Unload Me
End Sub

Private Sub cmdSel_Click()
    If tbrBuscador1.GetLstSel = "" Then
        MsgBox "No ha eligido ningun producto"
        Exit Sub
    End If
    
    txtCant = CStr(1)
    CtoTmp = DB.GetValInRS("Productos", "pCosto", "Id = " + tbrBuscador1.GetLstSel(0), False)
    txtPrecio = FormatCurrency(CtoTmp, , , , vbFalse)
     
    txtCant.SetFocus
    txtCant.SelStart = 0
    txtCant.SelLength = Len(txtCant)
    'cmdSel.Default = False
    cmdAgregar.Default = True

End Sub

Private Sub Command2_Click()
    frmProveedores.Show 1
End Sub

Private Sub Form_Activate()
    Dim tmpProv As String
    tmpProv = ""
    
    If Nuevo = False Then
        'cuando viene de pedido y se carga el combo el evento click
        'cambia nproveedor al primero de la lista lo grabo en una variable temp
        tmpProv = nProveedor
    End If
    
    CargarCombo cmbProveedores, "SELECT Proveedor FROM Proveedores " + _
        "ORDER BY Proveedor", "Proveedor"
    
    'hago que se ponga el nombre del proveedor del pedido
    'esto puede joder cuando agregue productos o proveedores pero
    'es lo mejor que se me ocurre
    
    If tmpProv <> "" Then
        cmbProveedores = tmpProv
    Else
        'veo si recien apreto "agregar prod", si no pierdo siempre el proveedor
        If tmpProv2 <> "" Then cmbProveedores = tmpProv2
        tmpProv2 = "" 'para futuros activates
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then cmdPagar_Click
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    tbrBuscador1.Contrasena = Contrasena
    tbrBuscador1.ArchivoMDB = ArchivoMDBPrincipal
    tbrBuscador1.SqlSinLike = "SELECT TOP 50 Id, nProducto FROM Productos WHERE ID >0"
    tbrBuscador1.CampoEnQueBuscar = "id/n,nProducto/b"
    tbrBuscador1.ColumnasSepPorComasyParentesis = "ID(700)/nProducto(2130)"
    
    cmbSucursales.Clear
    cmbSucursales.AddItem "CASA CENTRAL"
    CargarCombo cmbSucursales, "SELECT * FROM Sucursales", "Sucursal", , True
    cmbSucursales.ListIndex = 0 'si no hay sucursales no lo hace
    
    DTFecha = Date
    
'    'Margen de Venta ---------------------------------------------
'    Dim TmP As String
'    TmP = CFG.GetInfo(50, 4)
'    If TmP = "" Then TmP = "0"
'    If Not IsNumeric(TmP) Then TmP = "0"
'    If CLng(TmP) <= 0 Then
'        chkAjustarPrecio.Value = 0
'    Else
'        chkAjustarPrecio.Value = 1
'    End If
'    '--------------------------------------------------------------
    
    ActuTbrB = False
    Limpiar
    CalcularIVA
    
    'Otros Conceptos configurados ---------------------------------------------------
    Dim IdCF As Long
    
    IdCF = CFG.ExistePropiedad("ConceptoCpra 1")
    
    If IdCF <> 0 Then 'con que exista un concepto es suficiente
        'lvConceptos.Visible = True
        cmdConceptos.Visible = True
    End If
    '---------------------------------------------------------------------------------
    
    'negrada!!!! por ahora choreo la letra de la factura de ventas!!
    txtLetFac = Left(CFGBD.GetInfo(13, 4), 1)
    txtSucFac = "0000"
    txtNroFac = "00000000"
End Sub

Public Function AbrirDatos(Optional NroFactura As String = "Nada")
        
    On Local Error GoTo errAbrirCpr
    
    Terr.AnOtaR "caab", NroFactura
    
    If NroFactura = "Nada" Then
        Nuevo = True
        nProveedor = cmbProveedores
    Else
        Nuevo = False
        nProveedor = DB.GetValInRS("FacturaCompra", "Proveedor", "NroFactura = '" + _
            NroFactura + "'", True)
        
        Terr.AnOtaR "caac", nProveedor
        
        Dim FacDatos() As String
        FacDatos = Split(NroFactura, "-")
        
        txtLetFac = FacDatos(0)
        txtSucFac = FacDatos(1)
        txtNroFac = Left(FacDatos(2), 8)
        
        Terr.AnOtaR "caad", txtLetFac.Text, txtSucFac.Text, txtNroFac.Text
        
    End If
    
    ToTal = 0 '???? no se porque pero se guarda el valor de la factura vieja
    
    txtDeMas = FormatCurrency(0)
    txtPrecio = FormatCurrency(0)
    lblTotal = FormatCurrency(0)
    txtIVAPesos = FormatCurrency(0)
    
    If Nuevo = False Then 'abre un pedido grabado}
        Terr.AnOtaR "caae"
        CargarComboLV lvFactura, "SELECT Productos.nProducto, CompraDetalle.Cantidad, " + _
            "CompraDetalle.PrecioTotal " + _
            "FROM Productos INNER JOIN CompraDetalle ON Productos.ID = " + _
            "CompraDetalle.IDproducto WHERE NroFactura = '" + NroFactura + "'", _
            "Cantidad/n,nProducto,PrecioTotal/$"
        
        Terr.AnOtaR "caaf"
        'completo los precios unitarios dividiendo PT/cant
        Dim X As Long
        
        For X = 1 To lvFactura.ListItems.Count
            Terr.AnOtaR "caag", X, lvFactura.ListItems(X).SubItems(2), lvFactura.ListItems(X).SubItems(3), lvFactura.ListItems(X).Text
            
            lvFactura.ListItems(X).SubItems(3) = lvFactura.ListItems(X).SubItems(2)
            lvFactura.ListItems(X).SubItems(2) = CSng(lvFactura.ListItems(X).SubItems(3)) / _
                CLng(lvFactura.ListItems(X).Text)
            
            lvFactura.ListItems(X).SubItems(2) = FormatCurrency( _
                Round(lvFactura.ListItems(X).SubItems(2), 4), , , , vbFalse)
        Next X

        Terr.AnOtaR "caah", NroFactura
        
        'ahora borro todos los registros viejos
        DB.EXECUTE "DELETE FROM FacturaCompra WHERE NroFactura= '" + NroFactura + "'"
        
    End If
    
    Terr.AnOtaR "caai"
    CalcularTotal
    
    Terr.AnOtaR "caaj"
    Me.Show 1
    
    Exit Function
    
errAbrirCpr:
    Terr.AppendLog "ErrAbrirCPR", Terr.ErrToTXT(Err)
    MsgBox "Error al abrir la compra, envie registro a tbrSoft"
    
    
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    tbrBuscador1.CN_CLOSE
    
    Dim IDcierre As Long, AjustesStock As Single, RFyT As Single, DifStk As Single
    
    'calculo compras para pagina principal
    AjustesStock = PC.ABSSumarconSubcuentas(35, False)
    RFyT = PC.ABSSumarconSubcuentas(23, False)
    DifStk = PC.UltVarCuenta(54) - AjustesStock - RFyT
    
    CpDia = PC.ABSSumarconSubcuentas(18, False) + DifStk
End Sub

Private Sub tbrBuscador1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdSel_Click
End Sub

Private Sub txtCant_GotFocus()
    cmdAgregar.Default = True
    
    'por si utiliza codigo de barras lo necesita
    lblSelec = "Producto Seleccionado: " + UCase(tbrBuscador1.GetLstSel(1))
End Sub

Private Sub txtCant_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtPrecio.SetFocus
        txtPrecio.SelStart = 0
        txtPrecio.SelLength = Len(txtPrecio)
    End If

End Sub

Private Sub VerActuP()
    If ActuTbrBP = False Then
       ActuTbrBP = True
       tbrBuscador1.Recargar
    Else
        ActuTbrBP = False
    End If

End Sub

Private Sub tbrBuscador1_Change()
    Dim tbR As String, Nombre As String, I As Long

    If Not IsNumeric(tbrBuscador1.Text) Then
        tbrBuscador1.CampoEnQueBuscar = "id/n,nProducto/b"
    Else
        'si cargo con código de barras (supongo que :
            ' tiene como mínimo tiene 7 dígitos y es NUMÉRICO
            
        '1ro leo lo que escribio
        tbR = tbrBuscador1.Text
        
        'veo si tiene mas de 6 digitos ya lo tomo como CODIGO DE BARRAS
        If Len(tbR) > 6 Then '-> ES CODIGO DE BARRAS
            
            'veo si estan los caracteres de C.Barras
            Nombre = DB.GetValInRS("Productos", "nProducto", "CodigodeBarras = '" + tbR + "'")
                
            If Nombre = "" Then 'no lo encuentro
                'tbrBuscador1.Text = ""
                'Exit Sub 'NO SI NO NO LO ENCUENTRO MAS
            Else 'si lo encontro -> lo agrego directamente
                'pongo el nombre en el tbrBuscador
                tbrBuscador1.Text = Nombre
                
                'veo cuantos filtro
                Select Case tbrBuscador1.ListCount
                    Case 0 'no encontro ninguno (no deberia pasar pero ....)
                            ' anulo todo
                        'tbrBuscador1.Text = ""
                        'Exit Sub 'NO SI NO NO LO ENCUENTRO MAS
                    Case 1 'joya ... dale que va!
                        cmdSel_Click
                        tbrBuscador1.SelLength = (Len(tbrBuscador1.Text))
                        Exit Sub
                        'lo agrega solo
                    Case Is > 1 ' tengo que buscar 1 x 1 hasta que coincida exac.
                                ' puede pasar que si es muy cortito el nombre
                                ' y se filtren mucho, como este muestra 50 nomas
                                ' puede pasar que no lo encuentre ... ojala que no
                                ' tener cuenta en el futuro XXXXX
                        For I = 1 To tbrBuscador1.ListCount
                            If Nombre = tbrBuscador1.GetLstSel(1, I) Then
                                'tiene que quedar elegido -> los otros se borraron
                                cmdSel_Click
                                tbrBuscador1.SelLength = (Len(tbrBuscador1.Text))
                                Exit For
                                Exit Sub
                            Else
                                'tengo que borrar porque no tengo un procedimiento
                                'que diga que renglon quede SELECTED
                                tbrBuscador1.BorrarRenglon I
                            End If
                        Next I
                End Select
                
            End If
                                      
            'Exit Sub 'encuentre o no
        Else
            'busca el codigo normal
            tbrBuscador1.CampoEnQueBuscar = "id/b,nproducto"
        End If
    End If
    VerActuP
    
    lblSelec = "Producto Seleccionado: " + UCase(tbrBuscador1.GetLstSel(1))
End Sub

Private Sub tbrBuscador1_Click()
    cmdSel.Default = True
    lblSelec = "Producto Seleccionado: " + UCase(tbrBuscador1.GetLstSel(1))
End Sub

Private Sub txtCant_Change()
    If Not IsNumeric(txtPrecio) Or Not IsNumeric(txtCant) Or tbrBuscador1.GetLstSel = "" Then
        txtPrecio = FormatCurrency(0)
        Exit Sub
    End If
    
    CtoTmp = DB.GetValInRS("Productos", "pCosto", "Id = " + tbrBuscador1.GetLstSel(0), False)
    txtPrecio = FormatCurrency(CSng(txtCant) * CtoTmp, , , , vbFalse)
End Sub

Private Sub txtCant_LostFocus()
    Dim Cant As Long 'toma enteros si o si si viene decimal redondea
    Cant = ValidarNumeros(txtCant)
    txtCant = CLng(Cant)
End Sub

Private Sub txtDeMas_Change()
    CalcularIVA
End Sub

Private Sub txtDeMas_GotFocus()
    PintarTxt txtDeMas
End Sub

Private Sub txtDeMas_LostFocus()
    DeMas = ValidarNumeros(txtDeMas)
    txtDeMas = FormatCurrency(DeMas, , , , vbFalse)
End Sub

Private Sub txtIVAPesos_Change()
    If IsNumeric(txtIVAPesos) Then
        CalcularTotal
    End If
End Sub

Private Sub txtIVAPesos_GotFocus()
    PintarTxt txtIVAPesos
End Sub

Private Sub txtIVAPesos_LostFocus()
    txtIVAPesos = FormatCurrency(ValidarNumeros(txtIVAPesos), , , , vbFalse)
End Sub

Private Sub txtIVAPorC_Change()
    CalcularIVA
End Sub

Private Sub txtIVAPorC_GotFocus()
    PintarTxt txtIVAPorC
End Sub

Private Sub txtIVAPorC_LostFocus()
    txtIVAPorC = ValidarNumeros(txtIVAPorC)
End Sub

Private Sub txtLetFac_Change()
    txtLetFac = UCase(txtLetFac)
End Sub

Private Sub txtLetFac_GotFocus()
    PintarTxt txtLetFac
End Sub

Private Sub txtLetFac_KeyPress(KeyAscii As Integer)
    If Len(txtLetFac) > 0 Then
        If Len(txtLetFac) <> Len(txtLetFac.SelText) Then
            'solo dejo borrar
            PintarTxt txtLetFac
            
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub txtLetFac_LostFocus()
    If txtLetFac = "" Or IsNumeric(txtLetFac) Then txtLetFac = "A"
End Sub

Private Sub txtPrecio_GotFocus()
    PintarTxt txtPrecio
    cmdAgregar.Default = True
End Sub

Private Sub txtPrecio_LostFocus()
    Precio = ValidarNumeros(txtPrecio)
    txtPrecio = FormatCurrency(Precio, , , , vbFalse)
End Sub

Private Sub txtSucFac_GotFocus()
    PintarTxt txtSucFac
End Sub

Private Sub txtSucFac_LostFocus()
    txtSucFac = ValidarNumeros(txtSucFac)
    txtSucFac = String(4 - Len(txtSucFac), "0") + txtSucFac
End Sub

Private Sub txtnroFac_GotFocus()
    PintarTxt txtNroFac
End Sub

Private Sub txtnroFac_LostFocus()
    If Not IsNumeric(txtNroFac) Then txtNroFac = "0"
    txtNroFac = String(8 - Len(CStr(txtNroFac)), "0") + txtNroFac
End Sub
