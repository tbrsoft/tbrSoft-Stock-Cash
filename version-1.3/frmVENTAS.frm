VERSION 5.00
Object = "{A7FBD38D-2930-49E3-B60C-9E0202D84549}#14.0#0"; "tbrControles.ocx"
Object = "{DCB03D77-0A94-4AE8-9495-515B6968EEFB}#4.0#0"; "tbrFacura.ocx"
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmVENTAS 
   BackColor       =   &H00544B45&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ventas"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11850
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVENTAS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   90
      Top             =   150
   End
   Begin tbrFaroButton.fBoton command1 
      Height          =   420
      Left            =   6975
      TabIndex        =   60
      Top             =   7815
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   741
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdImprimir 
      Height          =   420
      Left            =   5355
      TabIndex        =   59
      Top             =   7815
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   741
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Imprimir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdTerminar 
      Height          =   420
      Left            =   3750
      TabIndex        =   58
      Top             =   7815
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   741
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Terminar (F1)"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdConceptos 
      Height          =   420
      Left            =   9930
      TabIndex        =   57
      Top             =   2670
      Visible         =   0   'False
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   741
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Otros Conceptos"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdEliminar 
      Height          =   420
      Left            =   5730
      TabIndex        =   56
      Top             =   3630
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   741
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Eliminar selección"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdBorrarConc 
      Height          =   420
      Left            =   7110
      TabIndex        =   55
      Top             =   2880
      Visible         =   0   'False
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   741
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Borrar Conceptos"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdSelProd 
      Height          =   420
      Left            =   3675
      TabIndex        =   53
      Top             =   735
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   741
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   ">>"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.CheckBox chkDtoVta 
      BackColor       =   &H00544B45&
      Caption         =   "Descuento por Vta Ctado"
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   8940
      TabIndex        =   50
      Top             =   8040
      Value           =   1  'Checked
      Width           =   2595
   End
   Begin VB.ComboBox cmbVendedor 
      Height          =   315
      Left            =   8970
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   5520
      Width           =   2505
   End
   Begin tbrFacura.Factura FaCC 
      Height          =   465
      Left            =   0
      TabIndex        =   48
      Top             =   7710
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   820
   End
   Begin tbrControles.MouTextBox txtSucFac 
      Height          =   405
      Left            =   6465
      TabIndex        =   3
      Top             =   300
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
   Begin tbrControles.MouTextBox txtIVAPorC 
      Height          =   345
      Left            =   10800
      TabIndex        =   10
      Top             =   1800
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   609
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
   Begin tbrControles.MouTextBox txtDesc 
      Height          =   465
      Left            =   10110
      TabIndex        =   8
      Top             =   1230
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   820
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
   Begin tbrControles.MouTextBox txtCant 
      Height          =   405
      Left            =   660
      TabIndex        =   1
      Top             =   4080
      Width           =   585
      _ExtentX        =   1032
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
      Largo           =   20
      Entero          =   -1  'True
   End
   Begin tbrControles.tbrBuscador tbrBuscadorP 
      Height          =   2655
      Left            =   60
      TabIndex        =   0
      Top             =   1170
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4683
      BackColor       =   5524293
      BeginProperty Fontt {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
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
      Height          =   315
      Left            =   8940
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   6330
      Width           =   2505
   End
   Begin VB.CheckBox chkEntregado 
      BackColor       =   &H00544B45&
      Caption         =   "Mercaderia se entregó"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   9000
      TabIndex        =   26
      Top             =   5940
      Value           =   1  'Checked
      Width           =   2475
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00544B45&
      Caption         =   "Datos del Cliente"
      ForeColor       =   &H8000000E&
      Height          =   2775
      Left            =   360
      TabIndex        =   5
      Top             =   4860
      Visible         =   0   'False
      Width           =   8475
      Begin tbrControles.tbrBuscador tbrBuscadorC 
         Height          =   1605
         Left            =   180
         TabIndex        =   17
         Top             =   540
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   2831
         BackColor       =   5524293
         BeginProperty Fontt {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
      Begin VB.TextBox txtIVa 
         Height          =   360
         Left            =   5280
         TabIndex        =   24
         Top             =   1950
         Width           =   3000
      End
      Begin VB.TextBox txtCUIT 
         Height          =   360
         Left            =   5280
         TabIndex        =   23
         Top             =   1500
         Width           =   3000
      End
      Begin VB.TextBox txtDireccion 
         Height          =   630
         Left            =   5280
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   780
         Width           =   3000
      End
      Begin tbrFaroButton.fBoton cmdAgregarCliente 
         Height          =   420
         Left            =   810
         TabIndex        =   61
         Top             =   2220
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   741
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Agregar - Modificar Cliente"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5717301
      End
      Begin VB.Label lblClSelec 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente: Wenceslao Gomez Bolaños Sarlengo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3960
         TabIndex        =   15
         Top             =   330
         Width           =   4215
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Condición IVA"
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   3630
         TabIndex        =   18
         Top             =   2010
         Width           =   1485
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CUIT / CUIL"
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   4050
         TabIndex        =   19
         Top             =   1590
         Width           =   1095
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Direccion"
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   3990
         TabIndex        =   20
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Buscador por Nombre o ID Cliente"
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Left            =   210
         TabIndex        =   25
         Top             =   300
         Width           =   3645
      End
   End
   Begin VB.CheckBox chkCliente 
      BackColor       =   &H00544B45&
      Caption         =   "Cliente Anónimo"
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   420
      TabIndex        =   16
      Top             =   4530
      Value           =   1  'Checked
      Width           =   1875
   End
   Begin VB.TextBox txtLetFac 
      Alignment       =   2  'Center
      Height          =   390
      Left            =   5895
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   300
      Width           =   555
   End
   Begin MSComCtl2.DTPicker DTFecha 
      Height          =   345
      Left            =   1890
      TabIndex        =   21
      Top             =   360
      Width           =   1365
      _ExtentX        =   2408
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
      Format          =   20905985
      CurrentDate     =   39197
   End
   Begin MSComctlLib.ListView lvTodo 
      Height          =   2205
      Left            =   5190
      TabIndex        =   46
      Top             =   1200
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3889
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
         Text            =   "Id.Prod."
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Cant"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Producto"
         Object.Width           =   3263
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Pr.Unit."
         Object.Width           =   1711
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Pr.Total"
         Object.Width           =   1852
      EndProperty
   End
   Begin VB.CheckBox chkIVA 
      BackColor       =   &H00544B45&
      Caption         =   "IVA "
      ForeColor       =   &H00E0E0E0&
      Height          =   345
      Left            =   10080
      TabIndex        =   9
      Top             =   1800
      Width           =   615
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00544B45&
      Caption         =   "Calcular Vuelto"
      ForeColor       =   &H00FFFFFF&
      Height          =   1185
      Left            =   8940
      TabIndex        =   37
      Top             =   6780
      Width           =   2505
      Begin tbrControles.MouTextBox txtPaga 
         Height          =   405
         Left            =   1170
         TabIndex        =   14
         Top             =   300
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   714
         Alignment       =   2
         BackColor       =   16777215
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dar de vuelto: "
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Left            =   60
         TabIndex        =   40
         Top             =   840
         Width           =   1245
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pago con"
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Left            =   120
         TabIndex        =   39
         Top             =   390
         Width           =   825
      End
      Begin VB.Label lblVuelto 
         BackStyle       =   0  'Transparent
         Caption         =   "25,75%"
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   1470
         TabIndex        =   38
         Top             =   810
         Width           =   795
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00544B45&
      Caption         =   "Forma de Pago"
      ForeColor       =   &H00FFFFFF&
      Height          =   945
      Left            =   9000
      TabIndex        =   36
      Top             =   4080
      Width           =   2475
      Begin VB.OptionButton chkACuenta 
         BackColor       =   &H00544B45&
         Caption         =   "A Cuenta ( F2 )"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   210
         TabIndex        =   13
         Top             =   570
         Width           =   1890
      End
      Begin VB.OptionButton chkContado 
         BackColor       =   &H00544B45&
         Caption         =   "Contado"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   210
         TabIndex        =   12
         Top             =   300
         Value           =   -1  'True
         Width           =   1515
      End
   End
   Begin VB.TextBox txtTOTAL 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   10020
      Locked          =   -1  'True
      TabIndex        =   45
      Text            =   "$ 000.00"
      Top             =   3540
      Width           =   1700
   End
   Begin VB.TextBox txtSubTotal 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   10050
      Locked          =   -1  'True
      TabIndex        =   44
      Text            =   "$ 000.00"
      Top             =   420
      Width           =   1700
   End
   Begin VB.TextBox txtPT 
      Alignment       =   2  'Center
      Height          =   390
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "888.88"
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox txtPU 
      Alignment       =   2  'Center
      Height          =   390
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "88.88"
      Top             =   4080
      Width           =   885
   End
   Begin tbrControles.MouTextBox txtNroFac 
      Height          =   405
      Left            =   6990
      TabIndex        =   4
      Top             =   300
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
   Begin tbrControles.MouTextBox txtIVAPesos 
      Height          =   375
      Left            =   10170
      TabIndex        =   11
      Top             =   2190
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
   Begin MSComctlLib.ListView lvConceptos 
      Height          =   1125
      Left            =   5160
      TabIndex        =   51
      Top             =   3510
      Visible         =   0   'False
      Width           =   3825
      _ExtentX        =   6747
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
         Text            =   "IdCta"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Concepto"
         Object.Width           =   3228
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Importe"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "NroConcepto"
         Object.Width           =   0
      EndProperty
   End
   Begin tbrFaroButton.fBoton cmdSel2 
      Height          =   420
      Left            =   3195
      TabIndex        =   54
      Top             =   4080
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   741
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   ">>"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.Label lblCant 
      Alignment       =   2  'Center
      BackColor       =   &H00DAD8D8&
      BackStyle       =   0  'Transparent
      Caption         =   "drtyyeryeytreytreyery"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   7995
      TabIndex        =   52
      Top             =   885
      UseMnemonic     =   0   'False
      Width           =   1905
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedor"
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   9000
      TabIndex        =   49
      Top             =   5220
      Width           =   1365
   End
   Begin VB.Label lblStockSucu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Stock en sucursal: 10"
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   -75
      TabIndex        =   29
      Top             =   7905
      Width           =   3795
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nro Factura"
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   4170
      TabIndex        =   47
      Top             =   390
      Width           =   1605
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Factura Cargada"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   5010
      TabIndex        =   43
      Top             =   825
      Width           =   3105
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      ForeColor       =   &H00E0E0E0&
      Height          =   330
      Left            =   11460
      TabIndex        =   42
      Top             =   1830
      Width           =   345
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   690
      TabIndex        =   41
      Top             =   450
      Width           =   1005
   End
   Begin VB.Label lblPor 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "25,75%"
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   9030
      TabIndex        =   35
      Top             =   3450
      Width           =   885
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Total a Pagar"
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   10110
      TabIndex        =   34
      Top             =   3180
      Width           =   1395
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Subtotal"
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   10380
      TabIndex        =   33
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Descuento"
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   10320
      TabIndex        =   32
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cant.        PU                 PT"
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   690
      TabIndex        =   31
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Buscador por Nombre o ID Producto"
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   510
      TabIndex        =   30
      Top             =   900
      Width           =   3735
   End
End
Attribute VB_Name = "frmVENTAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ActuTbrB As Boolean, ActuTbrBP As Boolean
Dim RSBUSCAR As New ADODB.Recordset
Dim Subtotal As Single
Dim ToTal As Single 'Total de la pesos de la factura
Dim Desc As Single
Dim DatosFac(26) As String, Mases() As String
Dim DiscriminaIVA As Boolean

Private Sub chkACuenta_Click()
    If chkACuenta Then
        chkDtoVta.Value = 0
    End If
End Sub

Private Sub chkCliente_Click()
    If chkCliente Then
        Frame3.Visible = False
        Frame1.Left = 1000
        Frame1.Top = 5000
        Frame2.Left = 4000
        Frame2.Top = 5000
        chkDtoVta.Left = 1000
        chkDtoVta.Top = 6000
        
        Label15.Top = 4700
        Label15.Left = 7500
        cmbVendedor.Left = Label15.Left
        chkEntregado.Left = Label15.Left
        cmbSucursales.Left = Label15.Left
        cmbVendedor.Top = Label15.Top + 300
        chkEntregado.Top = Label15.Top + 720
        cmbSucursales.Top = Label15.Top + 1110
        
        lblStockSucu.Top = 7000
        cmdTerminar.Top = lblStockSucu.Top
        cmdImprimir.Top = lblStockSucu.Top
        command1.Top = lblStockSucu.Top
        Me.Height = 8000
    Else
        Frame3.Visible = True
        tbrBuscadorC.SetFocus
        tbrBuscadorC.Text = "Seleccione Cliente"
        Frame1.Left = 9000
        Frame2.Left = Frame1.Left
        chkDtoVta.Left = Frame1.Left
        Frame1.Top = 4090
        Frame2.Top = 6780
        chkDtoVta.Top = 8040
        
        Label15.Top = 5220
        Label15.Left = 9000
        cmbVendedor.Left = Label15.Left
        chkEntregado.Left = Label15.Left
        cmbSucursales.Left = Label15.Left
        cmbVendedor.Top = Label15.Top + 300
        chkEntregado.Top = Label15.Top + 720
        cmbSucursales.Top = Label15.Top + 1110
        
        lblStockSucu.Top = 7800
        cmdTerminar.Top = lblStockSucu.Top
        cmdImprimir.Top = lblStockSucu.Top
        command1.Top = lblStockSucu.Top
        Me.Height = 8800
    End If
End Sub

Private Sub chkDtoVta_Click()
    'si tiene descuento por venta contado lo pongo -------------------------------
    If chkDtoVta Then
        CalcularDescuentos
    Else
        txtDesc = FormatCurrency(0)
    End If
    '----------------------------------------------------------------------------
    CalcularTotal
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
    
    CalcularTotal
    If IsNumeric(txtIVAPorC) = True Then
        tmpIvaPorC = CSng(txtIVAPorC)
    Else
        tmpIvaPorC = 0
    End If
    
    'subtotal y desc y ivaporc tambien seguro tienen que ser validos
    txtIVAPesos = FormatCurrency((CSng(txtSubTotal) - NoNuloN(txtDesc)) * _
        CSng(tmpIvaPorC) / 100, , , , vbFalse)
    
    CalcularTotal
End Sub

Private Sub cmbSucursales_Click()
    If tbrBuscadorP.GetLstSel = "" Then
        lblStockSucu = ""
    Else
        lblStockSucu = "Stock en " + UCase(cmbSucursales) + ": " + GetST
    End If
End Sub

Private Sub cmdAgregarCliente_Click()
    If tbrBuscadorC.GetLstSel = "" Then 'no eligio ningun cliente es para agregar nuevos
        frmClientes.AbrirDatos -1
    Else
        If tbrBuscadorC.GetLstSel = "Otros" Then
            MsgBox "No puede modificar Otros", vbInformation, "Atencion"
            Exit Sub
        End If
        
        frmClientes.AbrirDatos tbrBuscadorC.GetLstSel(1)
    End If
End Sub

Private Sub cmdBorrarConc_Click()
    lvConceptos.ListItems.Clear
    lvConceptos.Visible = False
    cmdBorrarConc.Visible = False
    lvTodo.Height = 2205
    cmdEliminar.Top = 3480
    cmdEliminar.Left = 6000
    CalcularTotal
End Sub

Private Sub cmdConceptos_Click()
    frmOtrosConc.AbrirDatos
    If lvConceptos.ListItems.Count > 0 Then
        lvTodo.Height = 1850
        cmdEliminar.Top = 3070
        cmdEliminar.Left = 5400
        cmdBorrarConc.Visible = True
        lvConceptos.Visible = True
    Else
        cmdBorrarConc_Click
    End If
    
    CalcularTotal
End Sub

Private Sub cmdEliminar_Click()
    If lvTodo.ListItems.Count = 0 Then Exit Sub
    
    lvTodo.ListItems.Remove lvTodo.SelectedItem.Index
    
    CalcularTotal
    
    tbrBuscadorP.Text = "Elegir prod" 'tonto pero solo para que se ejecute el suceso
                                  'change y se actualize el stock
    
    tbrBuscadorP.SetFocus
    
    cmdSelProd.Default = True
End Sub

Private Function GetST() As String
    'depende de la sucursal que esté seleccionada
    Dim IDp As Long, Resp As Long, ClsP As New clsProducto, Y As Long
    
    If tbrBuscadorP.GetLstSel = "" Then
        GetST = 0
        Exit Function
    End If
    
    IDp = tbrBuscadorP.GetLstSel(0)
    Resp = ClsP.StockProductoenSucursal(IDp, cmbSucursales)
    
    'ahora le descuento los que ya haya cargado en la factura
    For Y = 1 To lvTodo.ListItems.Count
        If lvTodo.ListItems(Y).Text = CStr(IDp) Then
            Resp = Resp - CLng(lvTodo.ListItems(Y).SubItems(1))
        End If
    Next Y
    
    Set ClsP = Nothing
    
    GetST = CStr(Resp)
End Function

Private Sub cmdImprimir_Click()
    ImprimaNomas
    
    Dim FuenteAca As StdFont, ConfUso As String, Ii As Long
    
    Set FuenteAca = FaCC.LeerFuente(AP)
    ConfUso = FaCC.GetConfenUso
    
    Set FaCC.Fuente = FuenteAca
    FaCC.NombreArchConfSinPath = ConfUso
    
    If chkCliente.Value = 0 Then
        FaCC.Texto(0) = tbrBuscadorC.GetLstSel(1)
        FaCC.Texto(1) = tbrBuscadorC.GetLstSel
        FaCC.Texto(2) = txtDireccion
        FaCC.Texto(3) = txtCUIT
        FaCC.Texto(4) = txtIVa
    Else
        FaCC.Texto(0) = ""
        FaCC.Texto(1) = "Consumidor Final"
    End If
    
    For Ii = 5 To 11
        If IsNull(DatosFac(Ii)) Then
            FaCC.Texto(Ii) = ""
        Else
            FaCC.Texto(Ii) = DatosFac(Ii)
        End If
    Next Ii
    
    FaCC.Imprimir GetMatrizFac
End Sub

Private Function GetMatrizFac() As String()
    Dim Resp() As String, I As Long
    
    ReDim Preserve Resp(0)
    Resp(0) = "NADA"
        
    If lvTodo.ListItems.Count = 0 Then
        GetMatrizFac = Resp
        Exit Function
    End If
    
    For I = 1 To lvTodo.ListItems.Count
        ReDim Preserve Resp(I)
            'no va el id del producto
        Resp(I) = lvTodo.ListItems(I).SubItems(1) + "|" + _
            lvTodo.ListItems(I).SubItems(2) + "|" + _
            lvTodo.ListItems(I).SubItems(3) + "|" + _
            lvTodo.ListItems(I).SubItems(4)
    Next I
    
    GetMatrizFac = Resp
End Function

Private Sub cmdSel2_Click()
    If tbrBuscadorP.GetLstSel = "" Then Exit Sub
    
    If lvTodo.ListItems.Count > 3 Then
        If LIC.GetLic < Licencia1 Then
            MsgBox "Su dispone de licencia para utilizar el programa " + vbCrLf + _
                "se le permitirán 4 productos por factura", vbCritical, "Sin licencia"
            Exit Sub
        End If
    End If
    
    If Not IsNumeric(txtCant) Then
        txtCant = "1"
        PintarTxt txtCant
        txtCant.SetFocus
        Exit Sub
    Else
        'no permito numeros negativos
        If CSng(txtCant) < 1 Then
            txtCant.SetFocus
            Exit Sub
        End If
        
        If CSng(txtCant) > GetST Then
            If MsgBox("Va a incluir un producto que tiene stock menor " + _
                "en sucursal " + UCase(cmbSucursales) + vbCrLf + _
                "a la cantidad que desea vender " + vbCrLf + _
                "¿Está seguro del registro? " + vbCrLf + _
                "Si selecciona ACEPTAR seguirá la venta pero dejará al producto " + _
                "con STOCK NEGATIVO por lo que es recomendable realizar el ajuste " + _
                "correspondiente, si no es el producto que desea vender presione " + _
                "CANCELAR", vbOKCancel + vbExclamation, "ATENCIÓN") = vbCancel Then
                tbrBuscadorP.SetFocus
                Exit Sub
            End If
        End If
        
        Dim TmP As Long
        
        TmP = lvTodo.ListItems.Count + 1
        
        lvTodo.ListItems.Add TmP
                
        lvTodo.ListItems(TmP).Text = tbrBuscadorP.GetLstSel
        lvTodo.ListItems(TmP).SubItems(1) = txtCant
        lvTodo.ListItems(TmP).SubItems(2) = GetProd
        
        'si discrimina iva le saco el iva del precio---------------------------------
        Dim PU As Single, PT As Single
        If DiscriminaIVA Then
            Dim IVva As Single
            IVva = CSng(CFG.GetInfo(7, 4))
            PU = CSng(txtPU) / ((100 + IVva) / 100)
            PT = CLng(txtCant) * PU
        Else
            PU = CSng(txtPU): PT = CSng(txtPT)
        End If
        
        lvTodo.ListItems(TmP).SubItems(3) = FormatCurrency(PU, , , , vbFalse)
        lvTodo.ListItems(TmP).SubItems(4) = FormatCurrency(PT, , , , vbFalse)
        
        
        cmdSelProd.Default = True
        txtCant = "1"
    End If
    
    txtDesc.Enabled = True
    
    CalcularIVA
    
    'si tiene descuento por venta contado lo pongo -------------------------------
    If chkDtoVta Then
        CalcularDescuentos
    End If
    '----------------------------------------------------------------------------
    
    tbrBuscadorP.Text = "Elegir prod" 'tonto pero solo para que se ejecute el suceso
                                      'change y se actualize el stock
    tbrBuscadorP.SetFocus
    
End Sub

Private Sub cmdSelProd_Click()
    If tbrBuscadorP.ListCount = 0 Then Exit Sub
    
    'pasarlo a los de texto
    txtPU = FormatCurrency(GetPU, , , , vbFalse)
    txtPT = FormatCurrency(CSng(txtCant) * CSng(txtPU), , , , vbFalse)
    
    txtCant.SetFocus
    
    cmdSel2.Default = True
End Sub

Private Sub cmdTerminar_Click()
    On Local Error GoTo errVta02
    
    Terr.AnOtaR "abar"
    Dim TmP As String
    Dim VtaFac As Single, tmp2 As String
    Dim I As Long, NroAsiento As Long
    
    If lvTodo.ListItems.Count = 0 Then MsgBox "No se registro ninguna venta": Exit Sub
    Terr.AnOtaR "abas"
    If Not IsNumeric(txtDesc) Then
        MsgBox "Cargue correctamente los montos", vbInformation, "Atención"
        Exit Sub
    End If
    
    Terr.AnOtaR "abat", txtIVAPorC.Text
    txtIVAPorC = ValidarNumeros(txtIVAPorC)
    
    'valido nro factura
    If DB.ContarReg("SELECT NroFactura FROM Facturas WHERE NroFactura = '" + _
        GetNroFac + "'") <> 0 Then
        Terr.AnOtaR "abau"
        MsgBox "Factura ya ingresada", vbInformation, "Atención"
        txtLetFac.SetFocus
        Exit Sub
    End If
    
    
    If chkContado Then
        TmP = "Al Contado"
    Else
        TmP = "A Cuenta"
    End If
    
    Terr.AnOtaR "abav", TmP
    
    NroAsiento = PC.GetUltIDAsientoMasUno("LibroSubdiario")
    Desc = ValidarNumeros(txtDesc)
    txtDesc = FormatCurrency(Desc, , , , vbFalse)
    
    Terr.AnOtaR "abaw", NroAsiento, Desc
    CalcularTotal
    
    Terr.AnOtaR "abax"
    If MsgBox("Está por registrar la venta " + TmP + vbCrLf + _
        "por un monto de " + txtTOTAL + " ¿Desea Continuar?", _
        vbYesNo + vbInformation, "Atencion") = vbNo Then Exit Sub
    
    Terr.AnOtaR "abay"
    '1ro anoto la venta en facturas
    ' ¿A qué cliente?
    Dim IDC As Long
    If chkCliente Then
        IDC = 0
    Else
        If tbrBuscadorC.GetLstSel = "" Then
            IDC = 0
        Else
            IDC = CLng(tbrBuscadorC.GetLstSel(1))
        End If
    End If
    
    Terr.AnOtaR "abaz", IDC
    DB.EXECUTE "INSERT INTO Facturas (NroFactura, Fecha, IdCliente, Pagado, Entregado) " + _
        "VALUES ('" + GetNroFac + "', #" + _
        stFechaSQL(DTFecha) + "#, " + CStr(IDC) + "," + _
        Replace(CStr(CSng(txtTOTAL)), ",", ".") + ", " + _
        CStr(chkEntregado.Value) + ")"
    
    Terr.AnOtaR "abba"
    Dim ClsP As New clsProducto
    Dim X As String, IDVend As Long
    
    If cmbVendedor <> "" Then
        IDVend = PC.GetIDCuenta(cmbVendedor)
    Else
        IDVend = 0
    End If
    
    Terr.AnOtaR "abbb", IDVend
    For I = 1 To lvTodo.ListItems.Count
        Dim IDp As String, PR As Single, Cto As Single
        Dim Cant As Long
        Terr.AnOtaR "abbc", I
        IDp = txtInLvW(lvTodo, I, 0)
        PR = CSng(txtInLvW(lvTodo, I, 3))
'        Pr = CSng(DB.GetValInRS("Productos", "pventa", "ID=" + IDP))
        Cto = CSng(DB.GetValInRS("Productos", "pcosto", "ID=" + IDp))
        Cant = CLng(txtInLvW(lvTodo, I, 1))
        
        Terr.AnOtaR "abbd", IDp, PR, Cto, Cant
        'si no discrimina iva le saco el iva al precio en la tabla ventas
        'mas abajo se va a agregar iva
        If chkIVA.Value = 0 And InStrRev(CFG.GetInfo(102, 3), txtLetFac, , vbTextCompare) <> 0 Then
            tmp2 = CFG.GetInfo(7, 4)
            If Not IsNumeric(tmp2) Then tmp2 = "0"
            If CLng(tmp2) < 0 Then tmp2 = "0"
            CFG.ModificarNodo 7, , , , tmp2
            
            PR = PR / ((100 + CSng(tmp2)) / 100)
            Terr.AnOtaR "abbe", PR
        End If
        
        Terr.AnOtaR "abbf"
        'modifico stock solo si la mercaderia solo si se entrego
        If chkEntregado Then
            'solo modifico "StockOtraSuc" si es distinto de casa central, ya que
            'el stock del mismo se calcula por descarte
            Terr.AnOtaR "abbg"
            
            ClsP.ModificarStock CLng(IDp), -Cant, cmbSucursales, _
                "Por Ventas Factura Nº " + CStr(GetNroFac)
        End If
        
        'VtaFac = VtaFac + Cant * Pr
        
        X = "INSERT INTO Ventas (ID, Idproducto,Fecha, Cantidad,Precio," + _
            "Costo,IDVenta,ID3Vendedor) " + _
            "VALUES (" + IdAutonum("Ventas") + ", " + CStr(IDp) + ", #" + _
            stFechaSQL(Date) + "#," + CStr(Cant) + ", " + _
            Replace(CStr(PR), ",", ".") + _
            "," + Replace(CStr(Cto), ",", ".") + ", '" + _
            GetNroFac + "'," + CStr(IDVend) + ")"
            
        Terr.AnOtaR "abbh", X
        
        DB.EXECUTE X
    Next I
    
    Terr.AnOtaR "abbi"
    '------------------------------------------------------------------------------
    'grabo la configuracion de la proximo numero de factura ------------------------
    Dim HH As Boolean, Busca As Long, tmpBk As String, NoBk As String
    HH = False
    tmpBk = Right(GetNroFac, 8)
    NoBk = Left(GetNroFac, Len(GetNroFac) - 8)
    
    Terr.AnOtaR "abbj", NoBk
    If Not IsNumeric(tmpBk) Then
        Busca = 1
    Else
        Busca = CLng(tmpBk) + 1
    End If
    
    Terr.AnOtaR "abbk", Busca
    Do While Not HH = True
        Terr.AnOtaR "abbl"
        If DB.ContarReg("SELECT NroFactura FROM Facturas WHERE NroFactura = '" + _
            NoBk + String(8 - Len(CStr(Busca)), "0") + CStr(Busca) + "'") = 0 Then
            
            HH = True
            Exit Do

        End If
        Terr.AnOtaR "abbm", Busca
        Busca = Busca + 1
    Loop
    Terr.AnOtaR "abbn"
    CFGBD.ModificarNodo 13, , , , NoBk + CStr(Busca)
        'grabo la ultima para la letra en particular -------------------------------
    Dim IdQ As Long
    
    IdQ = CFGBD.ExistePropiedad("Factura " + txtLetFac)
    Terr.AnOtaR "abbo", IdQ
    If IdQ = 0 Then
        CFGBD.AgregarNodo "13", "Factura " + txtLetFac, "", _
            txtSucFac + "-" + txtNroFac, 0
    Else
        CFGBD.ModificarNodo IdQ, 13, , , txtSucFac + "-" + CStr(Busca)
    End If
    Terr.AnOtaR "abbp"
    '------------------------------------------------------------------------------
        
    Set ClsP = Nothing
    'registro aca nomas en el libro diario, si es a credito u otro hago
    'un asiento contrario anulando caja
    Dim Caja As Single, Iva As Single
    
    If txtIVAPesos.Visible = False Then
        Iva = 0
        Terr.AnOtaR "abbq"
    Else
        Terr.AnOtaR "abbr", txtIVAPesos
        Iva = CSng(txtIVAPesos)
        'agrego renglon para el iva
        DB.EXECUTE "INSERT INTO Ventas (ID, Idproducto,Fecha, Cantidad,Precio, " + _
            "Costo,IDVenta,ID3Vendedor) " + _
            "VALUES (" + IdAutonum("Ventas") + ",-2,#" + stFechaSQL(Date) + _
            "#,1, " + Replace(CStr(Iva), ",", ".") + ",0,'" + CStr(GetNroFac) + _
            "'," + CStr(IDVend) + ")"
        Terr.AnOtaR "abbs"
    End If
    
    Terr.AnOtaR "abbt"
    If EsCero(Desc) = False Then
        Terr.AnOtaR "abbu"
        'ahora si hay descuento lo agrego ------------------------------------------
        DB.EXECUTE "INSERT INTO Ventas (ID, Idproducto,Fecha, Cantidad,Precio," + _
            "Costo,IDVenta,ID3Vendedor) " + _
            "VALUES (" + IdAutonum("Ventas") + ",-1,#" + stFechaSQL(Date) + _
            "#,1, " + Replace(CStr(-CSng(txtDesc)), ",", ".") + ",0,'" + _
            GetNroFac + "'," + CStr(IDVend) + ")"
    End If '-------------------------------------------------------------------------
    
    Terr.AnOtaR "abbv"
    ' si tiene otros conceptos ------------------------------------------------------
    Dim strAs As String, strAs2 As String, Monto As Single
    strAs = "": strAs2 = "": Monto = 0
    If lvConceptos.Visible = True Then
        Terr.AnOtaR "abbw"
        For I = 1 To lvConceptos.ListItems.Count
            strAs = strAs + lvConceptos.ListItems(I).Text 'con los nros de cuentas
            strAs2 = strAs2 + lvConceptos.ListItems(I).SubItems(2) 'con los montos
            Monto = Monto + CSng(lvConceptos.ListItems(I).SubItems(2))
            Terr.AnOtaR "abbx", strAs, strAs2, Monto
            'ademas agrego en Ventas
            DB.EXECUTE "INSERT INTO Ventas (ID, Idproducto,Fecha, Cantidad,Precio," + _
                "Costo,IDVenta,ID3Vendedor) " + _
                "VALUES (" + IdAutonum("Ventas") + _
                ",-1" + lvConceptos.ListItems(I).SubItems(3) + _
                ",#" + stFechaSQL(Date) + _
                "#,1, " + _
                Replace(CStr(CSng(lvConceptos.ListItems(I).SubItems(2))), _
                ",", ".") + ",0,'" + _
                GetNroFac + "'," + CStr(IDVend) + ")"
            
            Terr.AnOtaR "abby"
            If I < lvConceptos.ListItems.Count Then
                Terr.AnOtaR "abbz"
                strAs = strAs + "/"
                strAs2 = strAs2 + "/"
            End If
        Next I
        Terr.AnOtaR "abca"
        'hagoel asiento nomas
        
        PC.Asiento "78", CStr(Monto), strAs, strAs2, "LibroSubDiario", , NroAsiento
        Terr.AnOtaR "abcb"
    End If
    Terr.AnOtaR "abcc"
    '--------------------------------------------------------------------------------
    VtaFac = CSng(txtSubTotal)
    Caja = VtaFac - Desc + Iva
    
    '--------------------------------------------------------------------------------
    'COMPLICATED!
    'si no hizo click en iva Y tiene una letra de factura que NO discrimina IVA
    'pero si se contabiliza hago lo siguiente
    
    If chkIVA.Value = 0 And InStrRev(CFG.GetInfo(102, 3), txtLetFac, , vbTextCompare) <> 0 Then
        Terr.AnOtaR "abcd"
        Dim NoDisc As Single
        '1ro me aseguro que IVA sea numérico
        'el valor tm2 lo saca en el bucle For - Next
        
        'en este caso vtaFac ya esta sin iva
        NoDisc = (Caja) / ((100 + CSng(tmp2)) / 100)
        'Iva = Round((VtaFac - Desc) * (CSng(tmP2) / 100), 4)
        Iva = Caja - NoDisc
        VtaFac = Round(NoDisc, 4)
        
        Terr.AnOtaR "abce", NoDisc, Iva, VtaFac
        'agrego el iva en la tabla compras
        DB.EXECUTE "INSERT INTO Ventas (ID, Idproducto,Fecha, Cantidad,Precio, " + _
            "Costo,IDVenta,ID3Vendedor) " + _
            "VALUES (" + IdAutonum("Ventas") + ",-2,#" + stFechaSQL(Date) + _
            "#,1, " + Replace(CStr(Iva), ",", ".") + ",0,'" + CStr(GetNroFac) + _
            "'," + CStr(IDVend) + ")"
        
        Terr.AnOtaR "abcf"
    End If
    
    Terr.AnOtaR "abcg"
    '--------------------------------------------------------------------------------
    
    'caja-dtos a ventas
    'si descuento es cero no hay drama voy a crear una funcion que periodicamente
    'elimine nos registros con $0, igual con el iva
    'XXXXX hacer que configure a que cuenta quiere mandarlo
    PC.Asiento "78/81", CStr(Caja) + "/" + CStr(Desc), "80/76", CStr(VtaFac) + _
        "/" + CStr(Iva), "LibroSubDiario", "Ventas Factura Nº " + CStr(GetNroFac), -NroAsiento
    'el asiento es negativo si es para añadir ya que puede haber datos por otros
    'conceptos comoretencion de gannancias
    
    Terr.AnOtaR "abch"
    
    'ctovta a mercaderias
    'saco cuanto es el costo de venta
    Dim Costo As Single
    Costo = DB.SumarProducto("Ventas", "Cantidad", "Costo", "IdVenta = '" + GetNroFac + "'")
    
    Terr.AnOtaR "abci", Costo
    
    PC.Asiento "18", CStr(Costo), "54", CStr(Costo), "LibroSubDiario", _
        "Costo Factura Nº " + CStr(GetNroFac)
    
    Terr.AnOtaR "abcj", chkACuenta, chkContado
    If chkACuenta Then frmClientesMov.AbrirDatos GetNroFac, False, tbrBuscadorC.GetLstSel
    If chkContado Then
        If CFG.GetInfo(95, 4) = "Si" Then frmPago.AbrirDatos Caja, , "Ventas Factura Nº " + CStr(GetNroFac)
    End If
    
    Terr.AnOtaR "abck"
    VtaFac = 0 'empiezo de vuelta
    
    Unload Me
    
    Exit Sub
    
errVta02:
    Terr.AppendLog "vta0023", Terr.ErrToTXT(Err)
    Resume Next
End Sub

Private Sub ImprimaNomas()
    'asi son los datos
'    FAC.Labeles(0) = "Id Usuario"
'    FAC.Labeles(1) = "Nombre"
'    FAC.Labeles(2) = "Domicilio"
'    FAC.Labeles(3) = " DNI / CUIT"
'    FAC.Labeles(4) = "Condicion IVA"
'    FAC.Labeles(5) = "Fecha"
'    FAC.Labeles(6) = "Concepto 1"
'    FAC.Labeles(7) = "Descuento"
'    FAC.Labeles(8) = "Neto sin IVA"
'    FAC.Labeles(9) = "IVA %"
'    FAC.Labeles(10) = "IVA $"
'    FAC.Labeles(11) = "A Pagar"

    If IsDate(lblfecha) Then
        DatosFac(5) = FormatDateTime(DTFecha, vbShortDate)
    Else
        DatosFac(5) = FormatDateTime(Date, vbShortDate)
    End If
        
    DatosFac(7) = FormatCurrency(Desc)
    DatosFac(8) = FormatCurrency(CSng(txtSubTotal) - Desc)
    DatosFac(9) = FormatPercent(CSng(txtIVAPorC) / 100)
    If IsNumeric(txtIVAPesos) Then
        DatosFac(10) = FormatCurrency(CSng(txtIVAPesos))
    Else
        DatosFac(10) = FormatCurrency(0)
    End If
    
    DatosFac(11) = FormatCurrency(CSng(txtTOTAL))
End Sub


Private Sub Command1_Click()
    If lvTodo.ListItems.Count = 0 Then Unload Me: Exit Sub
    If MsgBox("¿Está seguro que desea salir?", vbOKCancel, "Salir de Ventas") = vbCancel Then Exit Sub
        
    Unload Me
End Sub

Private Sub Form_Activate()
    Terr.AnOtaR "abaq"
    tbrBuscadorC_Click
    
End Sub

Private Sub Form_Load()
    On Local Error GoTo errVta
    Terr.AnOtaR "abaa"
    Limpiar
    Terr.AnOtaR "abdh"
    DTFecha = Date
    chkIVA_Click
    chkCliente_Click
    cmbSucursales.Clear
    cmbSucursales.AddItem "CASA CENTRAL"
    Terr.AnOtaR "abdi"
    CargarCombo cmbSucursales, "SELECT * FROM Sucursales", "Sucursal", , True
    cmbSucursales.ListIndex = 0 'si no hay sucursales no lo hace
    
    Terr.AnOtaR "abab"
    tbrBuscadorC.Contrasena = Contrasena
    tbrBuscadorC.ArchivoMDB = ArchivoMDBPrincipal
    tbrBuscadorC.SqlSinLike = "SELECT TOP 50 Id, Nombre FROM Clientes WHERE ID >=0"
    tbrBuscadorC.OrderBy = "ORDER BY ID ASC"
    tbrBuscadorC.CampoEnQueBuscar = "nombre/b,ID/n"
    tbrBuscadorC.ColumnasSepPorComasyParentesis = "Nombre(2460)/ID(700)"
    
    Terr.AnOtaR "abac"
    tbrBuscadorP.Contrasena = Contrasena
    tbrBuscadorP.ArchivoMDB = ArchivoMDBPrincipal
    tbrBuscadorP.SqlSinLike = "SELECT TOP 50 ID, nProducto, pVenta, Stock FROM Productos WHERE ID>0"
    tbrBuscadorP.OrderBy = "ORDER BY ID ASC"
    tbrBuscadorP.CampoEnQueBuscar = "id/n,nproducto/b,pVenta/$,Stock/n"
    tbrBuscadorP.ColumnasSepPorComasyParentesis = "ID(600)/Producto(2380)/" + _
        "Precio(1000)/Stock(600)"
    
    Terr.AnOtaR "abad"
    'Cargo los Vendedores -------------------------------------------------
    Dim IdCuentas() As String
    
    IdCuentas = PC.GetCuentas(53)
    Terr.AnOtaR "abae", IdCuentas(0)
    cmbVendedor.Clear
    
    Dim I As Long
    
    For I = 1 To UBound(IdCuentas)
        Terr.AnOtaR "abaf", I, IdCuentas(I)
        cmbVendedor.AddItem PC.GetNameCuenta(CLng(IdCuentas(I)))
    Next I
    
    Terr.AnOtaR "abag"
    If cmbVendedor.ListCount > 0 Then cmbVendedor.ListIndex = 0
    ' --------------------------------------------------------------------
    'elijo el que este configurado
    Dim TmP As String
    TmP = CFG.GetInfo(14, 4)
    Terr.AnOtaR "abah", TmP
    If TmP <> "" Then
        Terr.AnOtaR "abai"
        If PC.ExisteNCuenta(TmP) <> 0 Then
            Terr.AnOtaR "abaj"
            cmbVendedor = TmP
        End If
    End If
    Terr.AnOtaR "abak"
    CFG.ModificarNodo 14, , , , cmbVendedor
    ' -----------------------------------------------------------------------------
    'Descuento por Venta Contado ---------------------------------------------------
    TmP = CFG.GetInfo(60, 4)
    Terr.AnOtaR "abal", TmP
    If TmP = "" Then TmP = "0"
    If Not IsNumeric(TmP) Then TmP = "0"
    If CLng(TmP) <= 0 Then
        chkDtoVta.Value = 0
    Else
        chkDtoVta.Value = 1
    End If
    
    Terr.AnOtaR "abam"
    'Otros Conceptos configurados
    Dim IdCF As Long
    
    IdCF = CFG.ExistePropiedad("Concepto 1")
    Terr.AnOtaR "aban", IdCF
    
    If IdCF <> 0 Then 'con que exista un concepto es suficiente
        'lvConceptos.Visible = True
        cmdConceptos.Visible = True
    End If
    '---------------------------------------------------------------------------------
    Terr.AnOtaR "abao"
    DiscriminaIVA = VerSiDiscriminaIVA
    tbrBuscadorP.Recargar
    '---------------------------------------------------------------------------------
    ActuTbrB = False:    ActuTbrBP = False
    
    Terr.AnOtaR "abap"
    Exit Sub
errVta:
    Terr.AppendLog "vta009", Terr.ErrToTXT(Err)
    Resume Next
End Sub

Private Sub Limpiar()
    On Local Error GoTo errClean
    
    Terr.AnOtaR "abda"
    Dim TmpFac As String, SP() As String
    
    TmpFac = CFGBD.GetInfo(13, 4)
    Terr.AnOtaR "abdb", TmpFac
    
    SP = Split(TmpFac, "-")
    txtLetFac = UCase(SP(0))
    txtSucFac = String(4 - Len(SP(1)), "0") + SP(1)
    txtNroFac = String(8 - Len(SP(2)), "0") + SP(2)
    
    Terr.AnOtaR "abde", txtLetFac, txtSucFac, txtNroFac
    
    txtCant.Text = "1"
    Terr.AnOtaR "abdf1", CFG.GetInfo(10, 4)
    txtPaga.Text = CStr(FormatCurrency(CSng(CFG.GetInfo(10, 4))))
    Terr.AnOtaR "abdf2"
    lvTodo.ListItems.Clear
    Terr.AnOtaR "abdf3"
    txtPU.Text = CStr(FormatCurrency(0))
    Terr.AnOtaR "abdf4"
    txtPT.Text = CStr(FormatCurrency(0))
    Terr.AnOtaR "abdf5"
    lblClSelec.Caption = ""
    
    Terr.AnOtaR "abdf"
    Subtotal = 0: Desc = 0
    txtSubTotal = FormatCurrency(Subtotal, , , , vbFalse)
    txtDesc = FormatCurrency(Desc, , , , vbFalse)
    
    Terr.AnOtaR "abdg"
    
    Exit Sub
    
errClean:
    Terr.AppendLog "errClean8273", Terr.ErrToTXT(Err)
    Resume Next
End Sub
Private Function GetNroFac() As String
    GetNroFac = txtLetFac + "-" + _
        String(4 - Len(txtSucFac), "0") + txtSucFac + "-" + _
        String(8 - Len(txtNroFac), "0") + txtNroFac
End Function

Private Function VerSiDiscriminaIVA() As Boolean
    Dim TmP As String, Tm2 As String, Resp As Boolean
    '---------------------------------------------------------------------------------
    'IVA configurado
    'configuracion IVA predeterminado
    TmP = CFG.GetInfo(7, 4)
    Tm2 = CFG.GetInfo(102, 4)
    Resp = False
    
    If Not IsNumeric(TmP) Then TmP = "0"
    If CLng(TmP) < 0 Then TmP = "0"
    If CLng(TmP) > 0 Then
        'veo si esta letra de factura discrimina iva
        If InStrRev(Tm2, txtLetFac, , vbTextCompare) <> 0 Then
            chkIVA.Value = 1
            txtIVAPorC = TmP
            Resp = True
        Else
            txtIVAPorC = "0"
            chkIVA.Value = 0
        End If
    End If
        
    CFG.ModificarNodo 7, , , , TmP
    '---------------------------------------------------------------------------------
    Dim PU As Single, PT As Single, IVva As Single, R As Long, Neto As Single
    
    'OJO EN REALIDAD NETRO ES BRUTO!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    'si discrimina iva le saco el iva del precio a cada renglon (si hay)
    If lvTodo.ListItems.Count > 0 Then
        For R = 1 To lvTodo.ListItems.Count
            Neto = CSng(DB.GetValInRS("Productos", "pVenta", "ID = " + _
                lvTodo.ListItems(R).Text, False))
            If Resp Then
                IVva = CSng(CFG.GetInfo(7, 4))
                lvTodo.ListItems(R).SubItems(3) = FormatCurrency(Neto / _
                    ((100 + IVva) / 100), , , , vbFalse)
            Else
                lvTodo.ListItems(R).SubItems(3) = FormatCurrency(Neto, , , , vbFalse)
            End If
            
            lvTodo.ListItems(R).SubItems(4) = FormatCurrency(CLng(lvTodo.ListItems(R). _
                SubItems(1)) * CSng(lvTodo.ListItems(R).SubItems(3)), , , , vbFalse)
        Next R
    End If
    
    VerSiDiscriminaIVA = Resp
    CalcularIVA
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            cmdTerminar_Click
        
        Case vbKeyF2
            chkACuenta = True
            cmdTerminar_Click
        Case vbKeyEscape
            Command1_Click
        
        'si apreta Ctrl+Nro se carga ese nro en cantidad (para agilizar)
        'solo de 1 a 10
        Case 49
            If Shift = 2 Then txtCant = "1"
        Case 50
            If Shift = 2 Then txtCant = "2"
        Case 51
            If Shift = 2 Then txtCant = "3"
        Case 52
            If Shift = 2 Then txtCant = "4"
        Case 53
            If Shift = 2 Then txtCant = "5"
        Case 54
            If Shift = 2 Then txtCant = "6"
        Case 55
            If Shift = 2 Then txtCant = "7"
        Case 56
            If Shift = 2 Then txtCant = "8"
        Case 57
            If Shift = 2 Then txtCant = "9"
        
    End Select
End Sub

Private Function GetDtoVta(idProd As Long, Monto As Single) As Single
    'CONFIGURACION DTO VTA CONTADO ------------------------------------------------
    'si tiene configuracion particular la pongo si no uso la general
    Dim IDC As Long, Multip As Single
    IDC = CFG.ExistePropiedad("DVC " + CStr(idProd))
    
    If IDC = 0 Then
        Multip = CSng(CFG.GetInfo(60, 4)) / 100
    Else
        Multip = CSng(CFG.GetInfo(IDC, 4)) / 100
    End If
    ' -------------------------------------------------------------------------------
    
    GetDtoVta = Monto * Multip
End Function

Private Sub CalcularDescuentos()
    If lvTodo.ListItems.Count = 0 Then
        txtDesc = FormatCurrency(0)
        txtDesc.Enabled = False
        Exit Sub
    End If
    
    Dim Dtos As Single, I As Long
    Dtos = 0
    
    For I = 1 To lvTodo.ListItems.Count
        Dtos = Dtos + GetDtoVta(lvTodo.ListItems(I).Text, lvTodo.ListItems(I).SubItems(4))
    Next I
    
    txtDesc = FormatCurrency(Dtos, , , , vbFalse)
    txtDesc.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        If lvTodo.ListItems.Count = 0 Then Exit Sub
        If MsgBox("¿Está seguro que desea salir?", vbOKCancel, "Salir de Ventas") = vbCancel Then
            Cancel = True
        End If
    End If
    
    VtaDia = PC.ABSSumarconSubcuentas(17, False)
    CtoDia = PC.ABSSumarconSubcuentas(18, False)
    tbrBuscadorP.CN_CLOSE
    tbrBuscadorC.CN_CLOSE
End Sub

Private Function GetPU() As String
    GetPU = tbrBuscadorP.GetLstSel(2)
    If Not IsNumeric(GetPU) Then GetPU = "0"
End Function

Private Function GetProd() As String
    GetProd = tbrBuscadorP.GetLstSel(1)
End Function

Private Function GetProdTd(Indice As Long) As String
    GetProdTd = txtInLvW(lvTodo, Indice, 2)
End Function

Private Sub tbrbuscadorC_Change()
    If IsNumeric(tbrBuscadorC.Text) Then
        tbrBuscadorC.CampoEnQueBuscar = "Nombre,ID/b"
    Else
        tbrBuscadorC.CampoEnQueBuscar = "Nombre/b,ID/n"
    End If
    
    If tbrBuscadorC.Text <> "" Then VerActu
    
    If tbrBuscadorC.GetLstSel <> "" Then
        lblClSelec = "Cliente: " + tbrBuscadorC.GetLstSel
        
        Dim Cl As New clsCliente
        Cl.AbrirDatos CLng(tbrBuscadorC.GetLstSel(1))
        txtDireccion = Cl.Direccion
        txtCUIT = Cl.CUIT
        txtIVa = Cl.Iva
    Else
        lblClSelec = ""
        txtDireccion = ""
        txtCUIT = ""
        txtIVa = ""
    End If
End Sub

Private Sub tbrBuscadorC_Click()
    Terr.AnOtaR "abfa"
    If tbrBuscadorC.GetLstSel <> "" Then
        lblClSelec.Caption = "Cliente: " + tbrBuscadorC.GetLstSel
        Terr.AnOtaR "abfb", lblClSelec.Caption
        
        Dim Cl As New clsCliente
        Cl.AbrirDatos CLng(tbrBuscadorC.GetLstSel(1))
        Terr.AnOtaR "abfc", tbrBuscadorC.GetLstSel(1)
        
        txtDireccion.Text = Cl.Direccion
        Terr.AnOtaR "abfd", txtDireccion.Text
        txtCUIT.Text = Cl.CUIT
        txtIVa.Text = Cl.Iva
        
        Terr.AnOtaR "abfe", txtCUIT.Text, txtIVa.Text
    Else
        Terr.AnOtaR "abff"
        lblClSelec.Caption = ""
        txtDireccion.Text = ""
        txtCUIT.Text = ""
        txtIVa.Text = ""
    End If
    
    Terr.AnOtaR "abfg"
End Sub

Private Sub VerActu()
    If ActuTbrB = False Then
       ActuTbrB = True
       tbrBuscadorC.Recargar
    Else
        ActuTbrB = False
    End If

End Sub

Private Sub VerActuP()
    If ActuTbrBP = False Then
       ActuTbrBP = True
       tbrBuscadorP.Recargar
    Else
        ActuTbrBP = False
    End If

End Sub

Private Sub tbrBuscadorC_GotFocus()
    tbrBuscadorC.SelStart = 0
    tbrBuscadorC.SelLength = Len(tbrBuscadorC.Text)
End Sub

Private Sub tbrBuscadorP_Change()
    Dim tbR As String, Nombre As String
    Dim I As Long
    
    If Not IsNumeric(tbrBuscadorP.Text) Then
        'busca el codigo normal
        tbrBuscadorP.CampoEnQueBuscar = "id/n,nproducto/b,pVenta/$,Stock/n"
    Else
        'si cargo con código de barras (supongo que :
        ' tiene como mínimo tiene 7 dígitos y es NUMÉRICO
            
        '1ro leo que escribio
        tbR = tbrBuscadorP.Text
        
        'veo si tiene mas de 6 digitos
        If Len(tbR) > 6 Then
            'busco por codigo de barras
            'le saco esos caracteres a tBr
            tbR = Replace(tbR, TmP, "")
            Nombre = DB.GetValInRS("Productos", "nProducto", "CodigodeBarras = '" + tbR + "'")
                
            If Nombre = "" Then 'no lo encuentro
                'tbrBuscadorP.Text = ""
                'Exit Sub  NO SI NO NO LO ENCUENTRA MAS
            Else 'si lo encontro -> lo agrego directamente
                'pongo el nombre en el tbrBuscador
                tbrBuscadorP.Text = Nombre
                
                'veo cuantos filtro
                Select Case tbrBuscadorP.ListCount
                    Case 0 'no encontro ninguno (no deberia pasar pero ....)
                            ' anulo todo
                        'tbrBuscadorP.Text = ""
                        'Exit Sub  NO SI NO NO LO ENCUENTRA MAS
                    Case 1 'joya ... dale que va!
                        cmdSel2_Click
                        tbrBuscadorP.SelLength = (Len(tbrBuscadorP.Text))
                        Exit Sub
                        'lo agrega solo
                    Case Is > 1 ' tengo que buscar 1 x 1 hasta que coincida exac.
                                ' puede pasar que si es muy cortito el nombre
                                ' y se filtren mucho, como este muestra 50 nomas
                                ' puede pasar que no lo encuentre ... ojala que no
                                ' tener cuenta en el futuro XXXXX
                        For I = 1 To tbrBuscadorP.ListCount
                            If Nombre = tbrBuscadorP.GetLstSel(1, I) Then
                                'tiene que quedar elegido -> los otros se borraron
                                cmdSel2_Click
                                tbrBuscadorP.SelLength = (Len(tbrBuscadorP.Text))
                                Exit For
                                Exit Sub
                            Else
                                'tengo que borrar porque no tengo un procedimiento
                                'que diga que renglon quede SELECTED
                                tbrBuscadorP.BorrarRenglon I
                            End If
                        Next I
                    End Select
                'Exit Sub 'encuentre o no
            End If
        Else
            'busca el codigo normal
            tbrBuscadorP.CampoEnQueBuscar = "id/b,nproducto,pVenta/$,Stock/n"
        End If
    End If
    VerActuP
    tbrBuscadorP_Click
End Sub

Private Sub tbrBuscadorP_Click()
    'lo unico el stock no se actualiza hasta que se haga la venta
    cmdSelProd.Default = True
    txtPU = FormatCurrency(GetPU, , , , vbFalse)
    txtPT = FormatCurrency(CSng(txtCant) * CSng(txtPU), , , , vbFalse)
    
    If tbrBuscadorP.GetLstSel = "" Then
        lblStockSucu = ""
    Else
        lblStockSucu = "Stock en " + UCase(cmbSucursales) + ": " + GetST
    End If
    
End Sub

Private Sub tbrBuscadorP_GotFocus()
    tbrBuscadorP.SelStart = 0
    tbrBuscadorP.SelLength = Len(tbrBuscadorP.Text)
End Sub

Private Sub txtCant_Change()
    If Not IsNumeric(txtCant) Then
        txtPT = 0
    Else
        txtPT = FormatCurrency(CSng(txtCant) * CSng(txtPU), , , , vbFalse)
    End If
End Sub

Private Sub txtCant_GotFocus()
    PintarTxt txtCant
End Sub

Private Sub txtCant_LostFocus()
    If Not IsNumeric(txtCant) Then txtCant = "1"
End Sub

Private Sub txtDesc_Change()
    If Not IsNumeric(txtDesc) Then
        Desc = 0
    Else
        Desc = CSng(txtDesc)
    End If
    
    CalcularIVA
End Sub

Private Sub txtDesc_GotFocus()
    PintarTxt txtDesc
End Sub

Private Sub txtDesc_Lostfocus()
    Desc = ValidarNumeros(txtDesc)
    txtDesc = FormatCurrency(Desc, , , , vbFalse)
    CalcularIVA
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

Private Sub txtIVAPorC_LostFocus()
    txtIVAPorC = ValidarNumeros(txtIVAPorC)
End Sub

Private Sub txtLetFac_Change()
    txtLetFac = UCase(Left(txtLetFac, 1))
End Sub

Private Sub txtLetFac_GotFocus()
    PintarTxt txtLetFac
End Sub

Private Sub LlenarNroFac()
    Dim SP() As String, IDp As Long, NmPr As String
    'veo en esa letra cual es la última que uso
    IDp = CFGBD.ExistePropiedad("Factura " + txtLetFac)
    If IDp <> 0 Then
        NmPr = CFGBD.GetInfo(IDp, 4)
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
    LlenarNroFac
    If txtLetFac = "" Or IsNumeric(txtLetFac) Then txtLetFac = "A"
    DiscriminaIVA = VerSiDiscriminaIVA
End Sub

Private Sub txtPaga_Change()
    Terr.AnOtaR "abdf6"
    lblVuelto.Caption = CStr(FormatCurrency(CalcularVuelto))
End Sub

Private Sub txtPaga_GotFocus()
    PintarTxt txtPaga
End Sub

Private Sub txtPaga_LostFocus()
    txtPaga = FormatCurrency(ValidarNumeros(txtPaga), , , , vbFalse)
End Sub

Private Function CalcularVuelto() As Single
    On Local Error GoTo ErrVTL
    Terr.AnOtaR "abdf7", txtPaga.Text, txtTOTAL
    
    Dim Vuelto As Single
    
    If txtPaga.Text = "" Then txtPaga.Text = "0"
    If txtTOTAL.Text = "" Then txtTOTAL.Text = "0"
    
    If Not IsNumeric(txtPaga.Text) Or Not IsNumeric(txtTOTAL.Text) Then
        Terr.AnOtaR "abdf9"
        Vuelto = 0
    Else
        Vuelto = CSng(txtPaga.Text) - CSng(txtTOTAL.Text)
        Terr.AnOtaR "abdf8", Vuelto
    End If
    
    Terr.AnOtaR "abdf10", Vuelto
    If Vuelto > 0 Then
        CalcularVuelto = Vuelto
    Else
        CalcularVuelto = 0
    End If
    
    Exit Function
ErrVTL:
    Terr.AppendLog "errVTL232", Terr.ErrToTXT(Err)
    Resume Next
    
End Function

Private Sub CalcularTotal()
    Dim I As Long, Canti As Long
    
    Subtotal = 0
    For I = 1 To lvTodo.ListItems.Count
        Subtotal = Subtotal + txtInLvW(lvTodo, I, 4)
    Next I
    
    txtSubTotal = FormatCurrency(Subtotal, , , , vbFalse)
    ToTal = Subtotal - Desc
        
    If txtIVAPesos.Visible = True Then
        ToTal = ToTal + CSng(txtIVAPesos)
    End If
    
    If lvConceptos.Visible = True Then
        For I = 1 To lvConceptos.ListItems.Count
            ToTal = ToTal + CSng(lvConceptos.ListItems(I).SubItems(2))
        Next I
    End If
    
    'ya que esta muestro la cantidad de productos que hay
    If lvTodo.ListItems.Count = 0 Then
        lblCant = ""
    Else
        Canti = 0
        For I = 1 To lvTodo.ListItems.Count
            Canti = Canti + CLng(lvTodo.ListItems(I).SubItems(1))
        Next I
        
        lblCant = CStr(Canti) + " Producto"
        If Canti > 1 Then lblCant = lblCant + "s"
    End If
    
    txtTOTAL = FormatCurrency(ToTal, , , , vbFalse)
End Sub

Private Sub txtSubTotal_Change()
    CalcularTotal
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

Private Sub txtTOTAL_Change()
    'para calcular el porcentaje de la compra
    Dim Pre As Single
    Dim Cto As Single
    Pre = 0
    Cto = 0
    
    If lvTodo.ListItems.Count = 0 Then
        lblPor = FormatPercent(0)
        
    Else
        Dim Ii As Long
        For Ii = 1 To lvTodo.ListItems.Count
            'Dim ClsP As New clsProducto
            Dim CtU As Single, IDp As Long, Cant As Long, Prod As String, PTot As Single
            
            IDp = CLng(txtInLvW(lvTodo, Ii, 0))
            Cant = CLng(txtInLvW(lvTodo, Ii, 1))
            Prod = txtInLvW(lvTodo, Ii, 2)
            PTot = CSng(txtInLvW(lvTodo, Ii, 4))
            
            CtU = CSng(DB.GetValInRS("Productos", "pCosto", "ID =" + CStr(IDp), False))
            
            Cto = Cto + Cant * CtU
            Pre = Pre + PTot
            
        Next Ii
        
        'ahora le resto el descuento al precio
        Pre = Pre - Desc
        'si costo es 0 no hago que divide
        If Cto = 0 Then
            lblPor = "N/A"
        Else
            lblPor = FormatPercent(Pre / Cto - 1)
        End If
    End If
    
    lblVuelto = FormatCurrency(CalcularVuelto, , , , vbFalse)
End Sub

Private Function GetUltIDVmasUno() As Long
    Dim UltIDC As Long, SP() As String
        
    'EL NRO DE FAC ES ASI POR EJ: A-0021-21002023
    '   tengo que clng(sp(1)+sp(2))
    '   pero como se cual es el mas grande, puede ser que se hayan cargados
    '   en distinto orden. Pero no lo quiero complicar mas asi que es del
    '   ultimo que se cargo o sea donde el ID en Ventas sea mayor
    Dim NFac As String
    
    NFac = DB.GetTop1Rs("Ventas", "IdVenta", , , True)
    If NFac = "" Then
        UltIDC = 0
    Else
        SP = Split(NFac, "-")
        UltIDC = CLng(SP(1) & SP(2))
    End If
    
    GetUltIDVmasUno = UltIDC + 1
End Function
