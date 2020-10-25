VERSION 5.00
Object = "{A7FBD38D-2930-49E3-B60C-9E0202D84549}#15.0#0"; "tbrControles.ocx"
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmResClientes 
   BackColor       =   &H00544B45&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resumen Cuenta Cliente"
   ClientHeight    =   8610
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
   Icon            =   "frmResClientes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00544B45&
      Caption         =   "Imprimir"
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   7140
      TabIndex        =   32
      Top             =   3300
      Width           =   2475
      Begin VB.OptionButton chkDias 
         BackColor       =   &H00544B45&
         Caption         =   "Últimos 30 Días"
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
         Left            =   180
         TabIndex        =   35
         Top             =   510
         Width           =   1890
      End
      Begin VB.OptionButton chkTodo 
         BackColor       =   &H00544B45&
         Caption         =   "Todo"
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
         Left            =   180
         TabIndex        =   34
         Top             =   240
         Value           =   -1  'True
         Width           =   1515
      End
      Begin VB.OptionButton chkMov 
         BackColor       =   &H00544B45&
         Caption         =   "Últimos 20 Movimientos"
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
         Left            =   180
         TabIndex        =   33
         Top             =   810
         Width           =   2250
      End
   End
   Begin tbrFaroButton.fBoton cmdPTodo 
      Height          =   375
      Left            =   3855
      TabIndex        =   22
      Top             =   3345
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Pagar Todo"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdAcomodar 
      Height          =   375
      Left            =   2100
      TabIndex        =   23
      Top             =   3345
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Pasar a 1 Reg"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdPParcial 
      Height          =   375
      Left            =   375
      TabIndex        =   24
      Top             =   3345
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Pago Parcial"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdBorrarCom 
      Height          =   375
      Left            =   9555
      TabIndex        =   27
      Top             =   7065
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Borrar Selección"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdNuevoCom 
      Height          =   375
      Left            =   7545
      TabIndex        =   28
      Top             =   7065
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Nuevo Comentario"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdEParcial 
      Height          =   375
      Left            =   2490
      TabIndex        =   20
      Top             =   7485
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Dev Parcial"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin tbrControles.tbrBuscador tbrBuscadorE 
      Height          =   1785
      Left            =   120
      TabIndex        =   18
      Top             =   4785
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   3149
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
   Begin MSComctlLib.ListView lvFactura 
      Height          =   1695
      Left            =   7920
      TabIndex        =   17
      Top             =   630
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   2990
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cant"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Producto"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Precio Un."
         Object.Width           =   2028
      EndProperty
   End
   Begin MSComctlLib.ListView lvDeudas 
      Height          =   2595
      Left            =   60
      TabIndex        =   16
      Top             =   540
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   4577
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Fecha"
         Object.Width           =   1667
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Monto"
         Object.Width           =   2099
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Detalle"
         Object.Width           =   4383
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "IdMov"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Nro.Fac"
         Object.Width           =   2614
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Saldo"
         Object.Width           =   2152
      EndProperty
   End
   Begin VB.ComboBox cmbEnvases 
      BackColor       =   &H00FFFFFF&
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
      Left            =   420
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   7500
      Width           =   1965
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00544B45&
      Caption         =   "Filtro"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   150
      TabIndex        =   10
      Top             =   4170
      Width           =   3015
      Begin VB.OptionButton OpEnvases 
         BackColor       =   &H00544B45&
         Caption         =   "Envases"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1620
         TabIndex        =   12
         Top             =   220
         Width           =   1035
      End
      Begin VB.OptionButton OpDetalle 
         BackColor       =   &H00544B45&
         Caption         =   "Detalle"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   210
         TabIndex        =   11
         Top             =   220
         Value           =   -1  'True
         Width           =   1065
      End
   End
   Begin MSDataGridLib.DataGrid DGCom 
      Height          =   2175
      Left            =   7290
      TabIndex        =   8
      Top             =   4830
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   3836
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   17
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin tbrFaroButton.fBoton cmdImprimir 
      Height          =   375
      Left            =   5580
      TabIndex        =   21
      Top             =   3345
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Imprimir"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton Command1 
      Height          =   375
      Left            =   10275
      TabIndex        =   25
      Top             =   8130
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdBorrarTCom 
      Height          =   375
      Left            =   8760
      TabIndex        =   26
      Top             =   7605
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Borrar Todo"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdETodo 
      Height          =   375
      Left            =   5460
      TabIndex        =   29
      Top             =   7485
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Dev Todo"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdDevSelec 
      Height          =   375
      Left            =   3975
      TabIndex        =   30
      Top             =   7485
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Dev Selección"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdPagarF 
      Height          =   375
      Left            =   9945
      TabIndex        =   31
      Top             =   3645
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Pagar Factura"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin VB.Label lblLimite 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Aaaaaaaaaaaaa"
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
      Height          =   855
      Left            =   4560
      TabIndex        =   19
      Top             =   3810
      Width           =   2355
   End
   Begin VB.Label lblResumenE 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "sdsadsaDSAdsaDSAfdsafdsafdsafdsafdsafdsafdsafdsafdsafdsafdsgffdsgfdgfdsagfdsgfdsgfdshhjjj, retreytyt, trettrwtey, dasfdsafasdfadsf"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   180
      TabIndex        =   15
      Top             =   6660
      Width           =   7065
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Envases"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   3360
      TabIndex        =   14
      Top             =   4320
      Width           =   1185
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Comentarios"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8640
      TabIndex        =   9
      Top             =   4500
      Width           =   2235
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "% Adeudado del Total Factura"
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
      Left            =   7800
      TabIndex        =   7
      Top             =   2850
      Width           =   2745
   End
   Begin VB.Label lblParte 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10500
      TabIndex        =   6
      Top             =   2790
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo a Pagar"
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
      Left            =   8670
      TabIndex        =   5
      Top             =   2460
      Width           =   1665
   End
   Begin VB.Label lblPesosF 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10350
      TabIndex        =   4
      Top             =   2370
      Width           =   1485
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Factura adeudada"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8160
      TabIndex        =   3
      Top             =   210
      Width           =   2685
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Deudas de:"
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
      Height          =   225
      Left            =   270
      TabIndex        =   2
      Top             =   180
      Width           =   1005
   End
   Begin VB.Label lblPesos 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "$ 888,88"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5970
      TabIndex        =   1
      Top             =   90
      Width           =   1785
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Aaaaaaaaaaaaa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   4605
   End
End
Attribute VB_Name = "frmResClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IDC As Long
Dim Nombre As String
Dim rsCOM As New ADODB.Recordset
Dim UltRengImp As Long

Private Sub VerUltReng() 'para ver que eligio si va a imprimir
    If UltRengImp <> 0 Then 'solo cambia el caption cuando no elige Todo
        If UltRengImp < 0 Then
            'eligio dias atras
            chkDias.Caption = "Últimos " + CStr(-UltRengImp) + " días"
        Else
            chkMov.Caption = "Últimos " + CStr(UltRengImp) + " movimientos"
        End If
    End If
End Sub

Private Sub ActualizaR()
    Dim Clscx As New clsCliente
    Dim W As String
    
    Nombre = Clscx.GetCliente(IDC)
    
    If Len(Nombre) > 23 Then
        lblNombre = Left(Nombre, 23) + "..."
    Else
        lblNombre = Nombre
    End If
    
    W = "SELECT * FROM MovClientes WHERE CodCliente= " + CStr(IDC) + " ORDER BY Fecha DESC, ID DESC"
        
    CargarComboLV lvDeudas, W, "Fecha/f,Variacion/$,Detalle,id/n,Documento,id/n"
    'la ultima columna la quiero para saldo pero como si esta vacia da
    'error le pongo cualquier dato como "ID" y despues SI la lleno con saldos
    
    'cargo los saldos
    For J = lvDeudas.ListItems.Count To 1 Step -1
        If J < lvDeudas.ListItems.Count Then
            Res = CSng(lvDeudas.ListItems(J).ListSubItems(1)) + _
                lvDeudas.ListItems(J + 1).ListSubItems(5)
        Else
            Res = CSng(lvDeudas.ListItems(J).ListSubItems(1))
        End If

        lvDeudas.ListItems(J).ListSubItems(5) = FormatCurrency(Res)
    Next J
    '----------------------------------------------------------
    
    lblPesos = FormatCurrency(Clscx.GetDeuda(IDC), , , , vbFalse)
    
    CargarCombo cmbEnvases, "SELECT Envase FROM Envases", "envase"
   
    Dim XX() As String 'array que va a recibir el resumen de los env debidos
    XX = Clscx.GetMovEnvDe(IDC)
    For I = 0 To UBound(XX)
        ResumenE = ResumenE + XX(I)
    Next I
    lblResumenE = ResumenE
     
     'robo el vale del label
    Dim H As String
    H = Right(lblResumenE, Len(lblResumenE) - InStrRev(lblResumenE, ": ", , 1))
    If H = "" Then
        Vales = 0
    Else
        Vales = CSng(H)
    End If
    
    Set DGCom.DataSource = Nothing
    If rsCOM.State = adStateOpen Then rsCOM.Close
    rsCOM.CursorLocation = adUseClient
    rsCOM.Open "SELECT * FROM ComentariosClientes WHERE IDCliente = " + CStr(IDC) + _
        " ORDER BY ID DESC", DB.CN, adOpenStatic, adLockReadOnly
    
    Set DGCom.DataSource = rsCOM
    
    If lvDeudas.ListItems.Count = 0 Then
        lvFactura.ListItems.Clear
        lblParte = ""
        lblPesosF = ""
        Exit Sub
    End If
    
    Set Clscx = Nothing
    
    ' CONTROLO EL LIMITE ------------------------------------------------------------
    ' Esta en configuracion FDP +IdCliente (xx_xx_Límite)
    Dim IdCF As Long, SP() As String, Lim As Single
    
    IdCF = CFG.ExistePropiedad("FDP " + CStr(IDC))
            
    If IdCF <> 0 Then
        SP = Split(CFG.GetInfo(IdCF, 4), "_")
        If UBound(SP) >= 2 Then
            If IsNumeric(SP(2)) Then
                'el único caso que uso el límite
                Lim = CSng(SP(2))
                If Lim < CSng(lblPesos) Then 'se paso
                    lblLimite = "Se supero el límite (" + FormatCurrency(Lim) + ")"
                Else
                    lblLimite = ""
                End If
            Else
                lblLimite = ""
            End If
        End If
    
    Else 'si es 0 no hago nada
        lblLimite = ""
    End If
            
    ' -------------------------------------------------------------------------------
End Sub

Public Sub AbrirDatos(IdcLiente As Long)
    IDC = IdcLiente
    
    If IDC < 0 Then
        Me.Caption = "Resumen Cuenta Financiera"
    End If
    ActualizaR
    
    AcomodarDG
    
    Me.Show 1
End Sub

Private Sub AcomodarDG()
    DGCom.Columns("ID").Width = 0
    DGCom.Columns("Fecha").Width = 1100
    DGCom.Columns("Hora").Width = 500
    DGCom.Columns("idCliente").Width = 0
    DGCom.Columns("Comentario").Width = 1500
    DGCom.Columns("Usuario").Width = 950
    
    DGCom.Columns("Fecha").Alignment = dbgCenter
    DGCom.Columns("Hora").Alignment = dbgCenter
    
    If CFG.GetInfo(5, 4) = "No" Then
        'COMENTARIOS
        DGCom.Left = lvFactura.Left + lvFactura.Width + 400
        DGCom.Width = DGCom.Width + 1200
        DGCom.Columns(4).Width = DGCom.Columns(4).Width + 1200
    End If
    
    Label9.Left = DGCom.Left + DGCom.Width / 3.5
End Sub

Private Sub chkDias_Click()
    Dim TmP As String
    TmP = InputBox("Ingrese hasta cuántos días desea imprimir " + vbCrLf + _
        "en el resumen", "Días Atras", "30")
    If Not IsNumeric(TmP) Then
        chkTodo.Value = True
        UltRengImp = 0
    Else
        UltRengImp = -CLng(TmP)
    End If
    
    VerUltReng
End Sub

Private Sub chkMov_Click()
    Dim TmP As String
    TmP = InputBox("Ingrese hasta cuántos movimientos desea imprimir " + vbCrLf + _
        "en el resumen", "Días Atras", "10")
    If Not IsNumeric(TmP) Then
        chkTodo.Value = True
        UltRengImp = 0
    Else
        UltRengImp = CLng(TmP)
    End If
    
    VerUltReng
End Sub

Private Sub chkTodo_Click()
    UltRengImp = 0
    VerUltReng
End Sub

Private Sub cmdAcomodar_Click()
    'veo si el usuario que esta trabajando tiene habilitacion para entrar
    Dim UUs As Long
    UUs = ACC.UltUsuarioIngresado
    
    If ACC.ExisteRelacion(UUs, 20) = 0 Then
        MsgBox ACC.GetNombre("Usuario", "Usuarios", UUs) + " no está habilitado " + _
            "para ingresar." + vbCrLf + _
            "Debe Cambiar Sesión a la de un usuario habilitado", vbExclamation, "Atención"
        Exit Sub
    End If
    
    'registro el movimiento(20:Cobrar a Clientes)
    ACC.RegEvento UUs, 20, "Resumir movimientos de cliente " + CStr(IDC)
    
    Dim MsJe As String
    If lstDeudas.ListCount = 0 Then Exit Sub
    
    If MsgBox("Atención, ejecutando esta opción borrará todos los registros " + _
        "de: " + UCase(Nombre) + " Dejando sólo un registro con el saldo " + lblPesos.Caption + _
        " Si esta seguro, presione Aceptar, en caso contrario Cancelar. Es recomendable " + _
        " utilizar esta opción solo con el consentimiento del cliente", _
        vbInformation + vbOKCancel, "Atención") = vbCancel Then Exit Sub
        
    MsJe = InputBox("Escriba aqui una aclaración del movimiento para que " + _
            "quede en el registro", "Aclaración")
    
    'borro todo
    DB.EXECUTE "DELETE from movclientes where CodCliente= " + CStr(IDC)
    'dejo 1 solo registro
    DB.EXECUTE "INSERT INTO MovClientes (ID,Fecha,CodCliente,variacion," + _
        "Detalle) VALUES (" + IdAutonum("MovClientes") + _
        ",#" + stFechaSQL(Date) + "#," + _
        CStr(IDC) + "," + Replace(CStr(CSng(lblPesos)), ",", ".") + _
        ",'SALDO: " + lblPesos + " " + MsJe + "')"

    ActualizaR
        
End Sub

Private Sub cmdBorrarCom_Click()
    If DGCom.ApproxCount <= 0 Then Exit Sub
    
    If MsgBox("¿Está seguro de borrar el comentario?", _
        vbExclamation + vbOKCancel, "Borrar Comentario") = vbCancel Then Exit Sub
        
    DB.EXECUTE "DELETE FROM ComentariosClientes WHERE ID = " + _
        DGCom.Columns("id")
    ActualizaR
    AcomodarDG
End Sub

Private Sub cmdBorrarTCom_Click()
    If MsgBox("¿Está seguro de borrar todos los comentarios?", _
        vbExclamation + vbOKCancel, "Borrar Comentarios") = vbCancel Then Exit Sub
        
    DB.EXECUTE "DELETE FROM ComentariosClientes WHERE IdCliente = " + CStr(IDC)
    
    ActualizaR
    AcomodarDG
End Sub

Private Sub cmdImprimir_Click()
    Dim Tit() As String
    Dim Ensh As Single

    TP.LineasSeparadoras = True
    
    ReDim Preserve Tit(4)
    Tit(4) = "Resumen de Cuenta"
    Tit(0) = lblNombre.Caption
    Tit(1) = DB.GetValInRS("Clientes", "Direccion", "Nombre = '" + lblNombre + "'", True)
    Tit(3) = "Mail: " + DB.GetValInRS("Clientes", "Mail", "Nombre = '" + lblNombre + "'", True)
    Tit(2) = DB.GetValInRS("Clientes", "Telefono", "Nombre = '" + lblNombre + "'", True)

    If CFG.GetInfo(5, 4) = "No" Then
        'NO TIENE ENVASES
        Ensh = 1
    Else
        'ensancho (esta chiquito)
        Ensh = 1.25
    End If
    
    If Date + UltRengImp > CDate(lvDeudas.ListItems(1).Text) And UltRengImp < 0 Then
        MsgBox "No hay movimientos desde el " + CStr(Date + UltRengImp), vbInformation, "Atención"
        Exit Sub
    End If
            
    TP.ImprimirlvW lvDeudas, Tit, _
        "Fecha|Monto|Detalle|ID|Nro.Fac.|Saldo", _
        "Total Deuda: " + CStr(lblPesos.Caption), , Ensh, , , , 30, UltRengImp
End Sub

Private Sub cmdNuevoCom_Click()
    Dim UUs As Long, nBUS As String
    Dim Coment As String
    
    Coment = InputBox("Detalle el comentario sobre " + UCase(Nombre), _
        "Comentarios")
    
    If Coment = "" Then Exit Sub
    
    'grabo nomas
    UUs = ACC.UltUsuarioIngresado
    nBUS = ACC.GetNombre("Usuario", "Usuarios", UUs)
    
    DB.EXECUTE "INSERT INTO ComentariosClientes (ID,IdCliente,Fecha,Hora," + _
        "Comentario,Usuario) " + _
        "VALUES (" + IdAutonum("ComentariosClientes") + _
        "," + CStr(IDC) + ",#" + stFechaSQL(Date) + "#," + _
        CStr(Hour(Now)) + ",'" + Coment + "','" + nBUS + "')"
    
    MsgBox "El comentario fue grabado correctamente", vbInformation, "Atención"
    
    ActualizaR
    AcomodarDG
End Sub

Private Sub cmdDevSelec_Click()
    Dim IDmE As Long
    Dim Det As String
    
    If tbrBuscadorE.GetLstSel = "" Then Exit Sub
    Det = tbrBuscadorE.GetLstSel(1) + " env de " + tbrBuscadorE.GetLstSel(2) + _
        " por " + tbrBuscadorE.GetLstSel(3)
    
    If tbrBuscadorE.GetLstSel(2) = "No tiene" Then
        MsgBox "No hay nada que devolver", vbInformation, "Atencion"
        Exit Sub
    End If
    
    If MsgBox("Seleccione aceptar sólo si fueron devueltos " + _
        Det, vbInformation + vbOKCancel, "Devolución envases") = vbCancel Then Exit Sub
    
    IDmE = tbrBuscadorE.GetLstSel(5)
    
    'borro el registro
    DB.EXECUTE "DELETE FROM MovEnvases where Id = " + CStr(IDmE)
    
    'agrego un registro que detalle la devolucion, si es otro no porque recarga base
    'de datos
    If lblNombre <> "Otros" Then
        DB.EXECUTE "INSERT INTO MovEnvases (ID,Fecha,Envases,codcliente,Detalle) " + _
            "VALUES (" + IdAutonum("MovEnvases") + ",'" + _
            CStr(Date) + "','No tiene','" + CStr(IDC) + "','DEVOLVIO: " + Det + "')"
    End If
    
    'hago el asiento "Vales a Pagar" a "Caja"
    PC.Asiento "93", tbrBuscadorE.GetLstSel(3), "78", tbrBuscadorE.GetLstSel(3)
    ActualizaR
    
    tbrBuscadorE.Recargar
End Sub


Private Sub cmdEParcial_Click()
    Dim cDesc As String 'cant que se descuenta de la cuenta
    Dim DtE As String   'detalle que se va a cargar en la tabla
    
    cDesc = InputBox("Ingrese el la cantidad de envases de: " + _
        UCase(cmbEnvases) + " que entrego el cliente", "Devolución de envases", 1)
    If cDesc = "" Or cDesc = "0" Then Exit Sub
        'XXXX negrada para evitar numeros decimales
    cDesc = Replace(cDesc, ".", "AAAA")
    cDesc = Replace(cDesc, ",", "AAAA")
    
    If Not IsNumeric(cDesc) Then MsgBox "Debes cargar un número " + _
        "correcto!", vbExclamation, "Atención": Exit Sub
    
    If CSng(cDesc) <= 0 Then MsgBox "No puede ingresar valores no positivos": Exit Sub
     
    DtE = InputBox("Ingrese el detalle del movimiento", "Detalle devolución " + _
        "envases de: " + Nombre)
    
    DB.EXECUTE "INSERT INTO MovEnvases (ID,Fecha, CodCliente, CantEnv, Envases,Detalle) " + _
        "VALUES (" + IdAutonum("MovEnvases") + _
        ",#" + stFechaSQL(Date) + "#," + CStr(IDC) + "," + CStr(-CLng(cDesc)) + ",'" + _
        cmbEnvases + "','" + "(DP): Entrego " + cDesc + " envases de " + _
        cmbEnvases + " | " + DtE + "')"
    
    ActualizaR
    tbrBuscadorE.Recargar
End Sub

Private Sub cmdETodo_Click()
    Dim MsJ As String
    Dim cVAle As String 'lo que descontaria de vale
    
    If tbrBuscadorE.ListCount < 1 Then
        MsgBox "No hay nada que devolver!"
        rsMe.Close
        Exit Sub
    End If
    
    cVAle = Replace(CStr(Vales), ",", ".")
   
    MsJ = InputBox("Escriba aqui una aclaración del movimiento para que " + _
        "quede en el registro", "Aclaración")
    
    'borro todo
    DB.EXECUTE "DELETE * from movenvases where codcliente= " + CStr(IDC)
    
    'aclaro
    DB.EXECUTE "INSERT INTO MovEnvases (ID,Fecha, Envases, CodCliente, Detalle) " + _
        "VALUES (" + IdAutonum("MovEnvases") + _
        ",#" + stFechaSQL(Date) + "#,'No tiene'," + CStr(IDC) + ",'" + _
        "DEVOLVIO TODOS: " + "(" + lblResumenE + ") " + MsJ + "')"
        
    ActualizaR
    'registro vales a pagar a caja
    PC.Asiento "93", cVAle, "78", cVAle
    
    tbrBuscadorE.Recargar
    
End Sub


Private Sub cmdPagarF_Click()
    If lvDeudas.ListItems.Count = 0 Then Exit Sub
        
    'veo si el usuario que esta trabajando tiene habilitacion para entrar
    Dim UUs As Long
    UUs = ACC.UltUsuarioIngresado
    
    If ACC.ExisteRelacion(UUs, 20) = 0 Then
        MsgBox ACC.GetNombre("Usuario", "Usuarios", UUs) + " no está habilitado " + _
            "para ingresar." + vbCrLf + _
            "Debe Cambiar Sesión a la de un usuario habilitado", vbExclamation, "Atención"
        Exit Sub
    End If
    
    'registro el movimiento(20:Cobrar a Clientes)
    ACC.RegEvento UUs, 20, "Pago de factura cliente " + CStr(IDC)
    
    Dim NFac As String, FechaD As String, IdMov As Long, Vari As Single
    
    NFac = txtInLvW(lvDeudas, lvDeudas.SelectedItem.Index, 4)
    IdMov = txtInLvW(lvDeudas, lvDeudas.SelectedItem.Index, 3)
    FechaD = txtInLvW(lvDeudas, lvDeudas.SelectedItem.Index, 0)
    Vari = CSng(txtInLvW(lvDeudas, lvDeudas.SelectedItem.Index, 1))
    
    If Vari = 0 Then
        MsgBox "No hay nada que pagar", vbInformation, "Atención"
        Exit Sub
    End If
        
    If MsgBox("Está a punto de registrar el pago de " + FormatCurrency(Vari) + ". ¿Los datos son " + _
        "correctos?", vbInformation + vbOKCancel, "Atención") = vbCancel Then
        Exit Sub
    End If
    
    'voy a entrar a ese registro y a cambiar los datos
    Dim rsO As New ADODB.Recordset
    If rsO.State = adStateOpen Then rsO.Close
    rsO.Open "SELECT * FROM MovClientes WHERE id =" + CStr(IdMov), DB.CN, adOpenStatic, adLockOptimistic
    
    rsO("Fecha") = Date
    rsO("Variacion") = 0
    
    Dim TmP As String
    If NFac = "NO" Then
        TmP = "Pago " + FormatCurrency(Vari)
    Else
        TmP = "Pago factura N°" + CStr(NFac) + " (" + lblPesosF + ")"
    End If
    
    rsO("Detalle") = TmP + " Adeudados desde " + FechaD
    rsO("Documento") = "NO"
    
    rsO.Update
    rsO.Close
    Set rsO = Nothing
    
    'borro los vencimientos
    DB.EXECUTE "DELETE FROM Vencimientos WHERE IdMov = " + CStr(IdMov)
    
    MsgBox "Se registró correctamente " + TmP + " Adeudados desde " + FechaD, _
        vbInformation, "Registro"
    
    'registro en el diario (caja a clientes)
    PC.Asiento "78", lblPesosF, "46", lblPesosF, "LibroSubDiario", _
        "Cobrado a Cliente Nro. " + CStr(IDC)
    
    If CFG.GetInfo(95, 4) = "Si" Then frmPago.AbrirDatos CSng(lblPesosF), True, "Cobrado a Cliente Nro. " + CStr(IDC)
    
    ActualizaR
    lvDeudas_Click
End Sub

Private Sub cmdPParcial_Click()
    'veo si el usuario que esta trabajando tiene habilitacion para entrar
    Dim UUs As Long
    UUs = ACC.UltUsuarioIngresado
    
    If ACC.ExisteRelacion(UUs, 20) = 0 Then
        MsgBox ACC.GetNombre("Usuario", "Usuarios", UUs) + " no está habilitado " + _
            "para ingresar." + vbCrLf + _
            "Debe Cambiar Sesión a la de un usuario habilitado", vbExclamation, "Atención"
        Exit Sub
    End If
    
    'registro el movimiento(20:Cobrar a Clientes)
    ACC.RegEvento UUs, 20, "Pago parcial a cliente " + CStr(IDC)
    
    Dim Desc As String 'monto que se descuenta de la cuenta
    Dim Dt As String   'detalle que se va a cargar en la tabla
    
    Desc = InputBox("Ingrese el Monto", "Pago de: " + Nombre)
    Desc = Replace(Desc, ".", ",")
    
    If Desc = "" Then Exit Sub
    If Not IsNumeric(Desc) Then
        MsgBox "¡Debes cargar un número correcto!", vbExclamation, "Atención"
        Exit Sub
    End If
    
    If CSng(Desc) <= 0 Then MsgBox "No puede ingresar valores no positivos": Exit Sub
    
    Dt = InputBox("Ingrese el detalle del pago", "Detalle pago de: " + Nombre)
    
    DB.EXECUTE "INSERT INTO MovClientes (ID,Fecha,CodCliente," + _
        "Variacion,Detalle,Documento) VALUES (" + IdAutonum("MovClientes") + _
        ",#" + stFechaSQL(Date) + _
        "#," + CStr(IDC) + "," + Replace(CStr(-CSng(Desc)), ",", ".") + ",'" + _
        "(PP) " + Dt + "','NO')"
    
    'registro en el diario (caja a clientes)
    PC.Asiento "78", CStr(Desc), "46", CStr(Desc), "LibroSubDiario", _
        "Cobrado a Cliente Nro. " + CStr(IDC)
    
    If CFG.GetInfo(95, 4) = "Si" Then frmPago.AbrirDatos CSng(Desc), True, "Cobrado a Cliente Nro. " + CStr(IDC)
        
    ActualizaR 'actualizo
End Sub

Private Sub cmdPTodo_Click()
    If lblPesos = 0 Then MsgBox "No hay nada que pagar!": Exit Sub
    
    'veo si el usuario que esta trabajando tiene habilitacion para entrar
    Dim UUs As Long
    UUs = ACC.UltUsuarioIngresado
    
    If ACC.ExisteRelacion(UUs, 20) = 0 Then
        MsgBox ACC.GetNombre("Usuario", "Usuarios", UUs) + " no está habilitado " + _
            "para ingresar." + vbCrLf + _
            "Debe Cambiar Sesión a la de un usuario habilitado", vbExclamation, "Atención"
        Exit Sub
    End If
    
    'registro el movimiento(20:Cobrar a Clientes)
    ACC.RegEvento UUs, 20, "Pago total a cliente " + CStr(IDC)
    
    Dim MsJ As String
    
    If MsgBox("Está a punto de borrar el total de la deuda de: " + _
        UCase(Nombre) + vbCrLf + "Presione Aceptar Sólo si ya fue cobrado " + _
        "en efectivo " + lblPesos.Caption + vbCrLf + _
        "Si el pago es parcial presione cancelar", vbInformation + vbOKCancel, _
        "Borrar registros") = vbCancel Then Exit Sub 'si esta seguro y puso la clave listo borro todo $$
    
        
    MsJ = InputBox("Escriba aqui una aclaración del movimiento para que " + _
        "quede en el registro", "Aclaración")
    'borro todo
    DB.EXECUTE "DELETE from MovClientes where CodCliente= " + CStr(IDC)
    'dejo 1 solo registro
    DB.EXECUTE "INSERT INTO MovClientes (ID,Fecha,CodCliente,variacion," + _
        "Detalle,Documento) VALUES (" + IdAutonum("MovClientes") + _
        ",#" + stFechaSQL(Date) + "#," + _
        CStr(IDC) + ",0,'CANCELÓ TODO! ( " + lblPesos + ") " + MsJ + "','NO')"
        
    'registro en el diario (caja a clientes)
    PC.Asiento "78", lblPesos, "46", lblPesos, "LibroSubDiario", _
        "Cobrado a Cliente Nro. " + CStr(IDC)
    
    If CFG.GetInfo(95, 4) = "Si" Then frmPago.AbrirDatos CSng(lblPesos), True, "Cobrado a Cliente Nro. " + CStr(IDC)
        
    'actualizo
    ActualizaR
        
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    tbrBuscadorE.Recargar
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    lblLimite = ""
    
    tbrBuscadorE.ColumnasSepPorComasyParentesis = "Fecha(900)/Cant(600)/" + _
        "Envases(1100)/$$ Vale(800)/Detalle(2600)/Id(600)"
    
    tbrBuscadorE.Contrasena = Contrasena
    tbrBuscadorE.ArchivoMDB = ArchivoMDBPrincipal
    tbrBuscadorE.SqlSinLike = "SELECT * FROM MovEnvases WHERE codcliente= " + CStr(IDC)
    tbrBuscadorE.OrderBy = "ORDER BY fecha desc,id desc"
    tbrBuscadorE.CampoEnQueBuscar = "fecha/f,cantenv/n,envases,Depositoporenvase/$," + _
        "Detalle/b,id/n"
    
    tbrBuscadorE.Recargar

    'tiene envases????????????????????????????????????????????????????????????????????
    If CFG.GetInfo(5, 4) = "No" Then
        'NO TIENE ENVASES------------------------------------------------------------
        tbrBuscadorE.Visible = False
        Label7.Visible = False
        Frame1.Visible = False
        lblResumenE.Visible = False
        cmbEnvases.Visible = False
        cmdEParcial.Visible = False
        cmdDevSelec.Visible = False
        cmdETodo.Visible = False
        
        'FACTURA
        Label3.Left = 1300
        Label3.Top = lvDeudas.Height + 1900
        lvFactura.Left = 300
        lvFactura.Top = lvDeudas.Height + 2500
        lvFactura.Width = lvFactura.Width + 1000
        lvFactura.ColumnHeaders(2).Width = lvFactura.ColumnHeaders(2).Width + 1000
        Label4.Top = lvDeudas.Height + 4300
        Label4.Left = lvFactura.Left + lvFactura.Width - Label4.Width - lblPesosF.Width - 200
        Label5.Top = lvDeudas.Height + 4800
        Label5.Left = lvFactura.Left + lvFactura.Width - Label5.Width - lblParte.Width - 200
        lblPesosF.Top = Label4.Top - 80
        lblPesosF.Left = Label4.Left + Label4.Width + 200
        lblParte.Top = Label5.Top - 80
        lblParte.Left = Label5.Left + Label5.Width + 200
        cmdPagarF.Left = Label3.Left + 1600
        cmdPagarF.Top = Label5.Top + 500
        
        'COMENTARIOS
        'en acomodarDG
        
        'LVDEUDAS
        lvDeudas.Width = Me.Width - 2000
        lvDeudas.Left = 1000
        lvDeudas.ColumnHeaders(1).Width = 1000
        lvDeudas.ColumnHeaders(2).Width = 1400
        lvDeudas.ColumnHeaders(3).Width = 3300
        lvDeudas.ColumnHeaders(4).Width = 0
        lvDeudas.ColumnHeaders(5).Width = 2500
        lvDeudas.ColumnHeaders(6).Width = 1400
        
        lblPesos.Left = 1000 + lvDeudas.Width - lblPesos.Width
        lblNombre.Left = 1000 + Label2.Width + 50
        lblNombre.Width = lblNombre.Width + 1500
        Label2.Left = 1000
        cmdPParcial.Left = 1000
        cmdAcomodar.Left = 3000
        cmdImprimir.Left = 7000
        cmdPTodo.Left = 5000
        Frame3.Left = 8500
    Else
        'SI TIENE ENVASES ------------------------------------------------------------
    End If
    '----------------------------------------------------------------------------------
    
    UltRengImp = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    tbrBuscadorE.CN_CLOSE
    
    Set DGCom.DataSource = Nothing
    rsCOM.Close
    Set rsCOM = Nothing
End Sub

Private Sub lvDeudas_Click()
    If lvDeudas.ListItems.Count = 0 Then
        lvFactura.ListItems.Clear
        lblParte = ""
        lblPesosF = ""
        Exit Sub
    End If
    
    'voy a hacer que se carguen las facturas si las tiene
    Dim NroFac As String, TTF As Single, SS As String
    
    NroFac = txtInLvW(lvDeudas, lvDeudas.SelectedItem.Index, 4)
    
    'si es cero no fue una deuda de venta
    If NroFac = "NO" Or Left(NroFac, 4) = "INT." Then
        lvFactura.ListItems.Clear
        lblParte = ""
        lblPesosF = txtInLvW(lvDeudas, lvDeudas.SelectedItem.Index, 1)
        Exit Sub
    End If
        
    SS = "SELECT Productos.nProducto, Ventas.Cantidad, Ventas.Precio " + _
        "FROM Productos INNER JOIN Ventas ON Productos.ID = " + _
        "Ventas.IDproducto WHERE (((Ventas.IdVenta) = '" + NroFac + "')) " + _
        "GROUP BY Productos.nProducto, Ventas.Cantidad, Ventas.Precio"

    CargarComboLV lvFactura, SS, "Cantidad/n,nProducto,Precio/$"
            
    'ahora veo cuanto era el total de la factura para sacar el porcentaje
    TTF = DB.SumarProducto("Ventas", "Cantidad", "Precio", _
          "idVenta = '" + NroFac + "'")
      
    lblPesosF = txtInLvW(lvDeudas, lvDeudas.SelectedItem.Index, 1)
    If TTF = 0 Then
        lblParte = " - "
    Else
        lblParte = FormatPercent(CSng(lblPesosF) / TTF)
    End If
End Sub

Private Sub lvDeudas_KeyUp(KeyCode As Integer, Shift As Integer)
    lvDeudas_Click
End Sub
