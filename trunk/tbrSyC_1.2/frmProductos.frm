VERSION 5.00
Object = "{A7FBD38D-2930-49E3-B60C-9E0202D84549}#15.0#0"; "tbrControles.ocx"
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmProductos 
   BackColor       =   &H00EFEFEF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Editar Productos"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11940
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProductos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   11940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin tbrFaroButton.fBoton cmdModificar 
      Height          =   375
      Index           =   8
      Left            =   9030
      TabIndex        =   14
      Top             =   2820
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Modificar"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdModificar 
      Height          =   375
      Index           =   9
      Left            =   8370
      TabIndex        =   18
      Top             =   3900
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Modificar"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdModificar 
      Height          =   375
      Index           =   3
      Left            =   10110
      TabIndex        =   16
      Top             =   3330
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Modificar"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdModificar 
      Height          =   375
      Index           =   2
      Left            =   9000
      TabIndex        =   12
      Top             =   2370
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Modificar"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdModificar 
      Height          =   375
      Index           =   1
      Left            =   9000
      TabIndex        =   9
      Top             =   1890
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Modificar"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdModificar 
      Height          =   375
      Index           =   0
      Left            =   8970
      TabIndex        =   7
      Top             =   1410
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Modificar"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdModificar 
      Height          =   375
      Index           =   4
      Left            =   9330
      TabIndex        =   2
      Top             =   480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Modificar"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdModificar 
      Height          =   375
      Index           =   5
      Left            =   10290
      TabIndex        =   5
      Top             =   960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Modificar"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton command1 
      Height          =   405
      Left            =   10650
      TabIndex        =   56
      Top             =   8220
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdImprimirCod 
      Height          =   405
      Left            =   1830
      TabIndex        =   55
      Top             =   8130
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Imprimir Código Producto"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrControles.tbrBuscador tbrBuscadorP 
      Height          =   2385
      Left            =   180
      TabIndex        =   0
      Top             =   480
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4207
      BackColor       =   15724527
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
   Begin VB.Frame FR_Descripcion 
      BackColor       =   &H00EFEFEF&
      Caption         =   "Descripción para el punto de venta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2385
      Left            =   150
      TabIndex        =   46
      Top             =   5640
      Width           =   5745
      Begin tbrFaroButton.fBoton command7 
         Height          =   435
         Left            =   540
         TabIndex        =   48
         Top             =   1890
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   767
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Eliminar Texto"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5717301
      End
      Begin VB.TextBox txtTextoPV 
         Height          =   1605
         Left            =   270
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   47
         Top             =   240
         Width           =   4785
      End
      Begin tbrFaroButton.fBoton command4 
         Height          =   435
         Left            =   2700
         TabIndex        =   49
         Top             =   1890
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   767
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Grabar este Texto"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5717301
      End
   End
   Begin VB.Frame fr_IMAGENES 
      BackColor       =   &H00EFEFEF&
      Caption         =   "Imágenes del producto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2895
      Left            =   6210
      TabIndex        =   45
      Top             =   5220
      Width           =   5535
      Begin tbrFaroButton.fBoton command5 
         Height          =   435
         Left            =   750
         TabIndex        =   57
         Top             =   2380
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   767
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Agregar Imagen"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5717301
      End
      Begin VB.ListBox lstIMGS 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2070
         IntegralHeight  =   0   'False
         ItemData        =   "frmProductos.frx":08CA
         Left            =   120
         List            =   "frmProductos.frx":08D1
         TabIndex        =   30
         Top             =   270
         Width           =   1125
      End
      Begin VB.PictureBox picFONDO 
         BackColor       =   &H00F5F1EB&
         Height          =   2055
         Left            =   1320
         ScaleHeight     =   1995
         ScaleWidth      =   3585
         TabIndex        =   50
         Top             =   270
         Width           =   3645
         Begin VB.Image imgPV 
            Height          =   1245
            Left            =   780
            Top             =   390
            Width           =   1665
         End
      End
      Begin tbrFaroButton.fBoton command6 
         Height          =   435
         Left            =   3000
         TabIndex        =   61
         Top             =   2380
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   767
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Eliminar Imagen"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5717301
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EFEFEF&
      Caption         =   "Modificar Producto"
      ForeColor       =   &H00000000&
      Height          =   4965
      Left            =   6210
      TabIndex        =   36
      Top             =   240
      Width           =   5535
      Begin tbrFaroButton.fBoton cmdVerC 
         Height          =   375
         Left            =   3780
         TabIndex        =   10
         Top             =   1650
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Ver Costo"
         fEnabled        =   -1  'True
         fFontN          =   "Arial"
         fFontS          =   8
         fECol           =   5717301
      End
      Begin tbrFaroButton.fBoton cmdModificar 
         Height          =   375
         Index           =   7
         Left            =   3000
         TabIndex        =   21
         Top             =   4080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Modificar"
         fEnabled        =   -1  'True
         fFontN          =   "Arial"
         fFontS          =   8
         fECol           =   5717301
      End
      Begin tbrFaroButton.fBoton command3 
         Height          =   375
         Index           =   2
         Left            =   4020
         TabIndex        =   24
         Top             =   4500
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Nuevo Env."
         fEnabled        =   -1  'True
         fFontN          =   "Arial"
         fFontS          =   8
         fECol           =   5717301
      End
      Begin tbrControles.MouTextBox txtDtoCtado 
         Height          =   375
         Left            =   1245
         TabIndex        =   13
         Top             =   2580
         Width           =   1515
         _ExtentX        =   2672
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
      Begin VB.TextBox txtCodigoBarras 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1230
         TabIndex        =   20
         Top             =   4080
         Width           =   1725
      End
      Begin VB.ComboBox cmbEnvases 
         Height          =   315
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   4530
         Width           =   1725
      End
      Begin VB.TextBox txtnProductoM 
         Height          =   375
         Left            =   1250
         TabIndex        =   4
         Top             =   720
         Width           =   2805
      End
      Begin VB.ComboBox cmbTipoM 
         Height          =   315
         Left            =   1250
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   270
         Width           =   1845
      End
      Begin VB.TextBox txtDetalle 
         Height          =   525
         Left            =   1260
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   3000
         Width           =   2595
      End
      Begin tbrControles.MouTextBox txtStock 
         Height          =   375
         Left            =   1230
         TabIndex        =   17
         Top             =   3660
         Width           =   885
         _ExtentX        =   1561
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
      Begin tbrControles.MouTextBox txtIVA 
         Height          =   375
         Left            =   4110
         TabIndex        =   19
         Top             =   3630
         Width           =   645
         _ExtentX        =   1138
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
      Begin tbrFaroButton.fBoton command3 
         Height          =   375
         Index           =   1
         Left            =   4140
         TabIndex        =   3
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Nuevo Tipo"
         fEnabled        =   -1  'True
         fFontN          =   "Arial"
         fFontS          =   8
         fECol           =   5717301
      End
      Begin tbrFaroButton.fBoton cmdModificar 
         Height          =   375
         Index           =   6
         Left            =   3000
         TabIndex        =   23
         Top             =   4500
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Modificar"
         fEnabled        =   -1  'True
         fFontN          =   "Arial"
         fFontS          =   8
         fECol           =   5717301
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "C.Barras"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   0
         TabIndex        =   60
         Top             =   4140
         Width           =   1095
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         ForeColor       =   &H00E0E0E0&
         Height          =   330
         Left            =   4710
         TabIndex        =   54
         Top             =   3690
         Width           =   285
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "IVA al comprar"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   2970
         TabIndex        =   53
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Minimo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   0
         TabIndex        =   52
         Top             =   3720
         Width           =   1155
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dto.Vta.Ctado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   0
         TabIndex        =   51
         Top             =   2640
         Width           =   1185
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Envase"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   0
         TabIndex        =   43
         Top             =   4590
         Width           =   1095
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   0
         TabIndex        =   42
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   0
         TabIndex        =   41
         Top             =   780
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Precio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   0
         TabIndex        =   40
         Top             =   1230
         Width           =   1095
      End
      Begin VB.Label lblPrecio 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   1230
         TabIndex        =   6
         Top             =   1140
         Width           =   1500
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Costo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   0
         TabIndex        =   39
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblCosto 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   1245
         TabIndex        =   8
         Top             =   1620
         Visible         =   0   'False
         Width           =   1500
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Stock"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   0
         TabIndex        =   38
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label lblStock 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   1245
         TabIndex        =   11
         Top             =   2100
         Width           =   1500
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   0
         TabIndex        =   37
         Top             =   3090
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E9F3FE&
      Caption         =   "AGREGAR PRODUCTO NUEVO "
      ForeColor       =   &H00000000&
      Height          =   2175
      Left            =   150
      TabIndex        =   32
      Top             =   3420
      Width           =   5745
      Begin tbrFaroButton.fBoton command2 
         Height          =   405
         Left            =   3480
         TabIndex        =   29
         Top             =   900
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   714
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Agregar Producto"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5717301
      End
      Begin tbrFaroButton.fBoton command3 
         Height          =   405
         Index           =   0
         Left            =   2910
         TabIndex        =   26
         Top             =   390
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   714
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Nuevo Tipo Producto"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5717301
      End
      Begin VB.TextBox txtDetalleN 
         Height          =   380
         Left            =   1020
         TabIndex        =   28
         Top             =   1440
         Width           =   3795
      End
      Begin VB.TextBox txtnProducto 
         Height          =   380
         Left            =   1005
         TabIndex        =   27
         Top             =   885
         Width           =   2445
      End
      Begin VB.ComboBox cmbTipo 
         Height          =   315
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   420
         Width           =   1845
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle"
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   -30
         TabIndex        =   35
         Top             =   1440
         Width           =   885
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   30
         TabIndex        =   34
         Top             =   930
         Width           =   885
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   30
         TabIndex        =   33
         Top             =   420
         Width           =   885
      End
   End
   Begin tbrFaroButton.fBoton fBoton1 
      Height          =   975
      Left            =   5280
      TabIndex        =   58
      Top             =   930
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   1720
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Eliminar producto elegido"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.Label lblProdSel 
      Alignment       =   2  'Center
      BackColor       =   &H00EFEFEF&
      Caption         =   "Label17"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   270
      TabIndex        =   59
      Top             =   2940
      Width           =   4875
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Por Nombre o Código"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2220
      TabIndex        =   44
      Top             =   240
      Width           =   2085
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Busqueda Producto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   31
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "frmProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ActuTbrB As Boolean

Dim FSO As New FileSystemObject

Private Sub cmdImprimirCod_Click()
    
    If tbrBuscadorP.GetLstSel = "" Then Exit Sub
    
    On Local Error GoTo ErrPrintCodigo
    
    'bloqueo el boton
    cmdImprimirCod.Enabled = False
    
    Dim TmP As String, SP() As String, R As Long, IdCF As String
    Dim MiX As Single, MiY As Single, Ancho As Single
        
    'Capricho Casa Arias --------------------------------------------------------------
    ' si no esta permitido la configuracion del codigo producto es que es Arias
    ' Se imprime solamente - Titulo (Nombre de la empresa), TipoProd,Prod y Detalle
    If CFG.GetInfo(6, 4) <> "Si" Then
        
        TP.SetFontPrinter "Times New Roman", 26, True, vbRed
        
        MiX = 800: MiY = 500
        Ancho = 5400
        'titulo -----------------------
        TP.tPrint MiX, MiY, DB.GetValInRS("Clientes", "Nombre", "ID = -2"), , True, Ancho
        MiY = MiY + 800
        
        TP.SetFontPrinter "Arial Narrow", 22, True, vbBlue
        
        'TipoProducto ----------------------------------------------------------------
        MiY = MiY + TP.tPrint(MiX, MiY, "Tipo: " + tbrBuscadorP.GetLstSel(1), , True, Ancho) * 600
        'Nombre ----------------------------------------------------------------------
        MiY = MiY + TP.tPrint(MiX, MiY, tbrBuscadorP.GetLstSel(2), , True, Ancho) * 600
        'Detalle --------------------------------------------------------------------
        MiY = MiY + TP.tPrint(MiX, MiY, txtDetalle, , True, Ancho) * 600
        
        TP.PrintCuadrado vbBlack, 24, 350, 350, 5250, 3600 '4250
        
        Printer.EndDoc
        'desbloqueo el boton
        cmdImprimirCod.Enabled = True
        
        Exit Sub
    End If
    '---------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------
    
    TmP = CFG.GetInfo(9, 4)
    If TmP = "" Then TmP = "ID_Nombre_Precio"
    If InStrRev(TmP, "_") = 0 Then TmP = "ID_Nombre_Precio"
    
    SP = Split(TmP, "_")
    
    IdCF = CFG.ExistePropiedadID(9)
    
    If IdCF = "" Then
        CFG.AgregarNodo 0, "Impresion Codigo", "", TmP, 0, 9
    Else
        CFG.ModificarNodo 9, , , , TmP
    End If
    
    'listo imprimiciono -----------------------------------------------------------
    TP.SetFontPrinter "Arial", 18, True, vbRed
     
    MiX = 800: MiY = 500
    Ancho = 4000
    
    'titulo
    If CFG.GetInfo(9, 3) <> "" Then
       TP.tPrint MiX, MiY, CFG.GetInfo(9, 3), , True, Ancho
        MiY = MiY + 600
    End If

    Printer.FontSize = 16
    Printer.ForeColor = vbBlue
    For R = 0 To UBound(SP)
        Select Case SP(R)
            Case "ID"
                MiY = MiY + TP.tPrint(MiX, MiY, "ID: " + tbrBuscadorP.GetLstSel, , _
                    True, Ancho) * 500
            Case "Nombre"
                MiY = MiY + TP.tPrint(MiX, MiY, tbrBuscadorP.GetLstSel(2), , True, Ancho) * 500
            Case "TipoProducto"
                MiY = MiY + TP.tPrint(MiX, MiY, "Tipo: " + tbrBuscadorP.GetLstSel(1), , True, Ancho) * 500
            Case "Precio"
                MiY = MiY + TP.tPrint(MiX, MiY, "Precio: " + lblPrecio, , True, Ancho) * 500
            Case "DtoVtaCtado"
                MiY = MiY + TP.tPrint(MiX, MiY, "Dto Vta Ctado: " + txtDtoCtado, , True, Ancho) * 500
            Case "Detalle"
                MiY = MiY + TP.tPrint(MiX, MiY, txtDetalle, , True, Ancho) * 500
        End Select
    Next R
        
    ' ------------- IMPRIMO CUADRADO -----------------------------------------------
    TP.PrintCuadrado vbBlack, 24, 350, 350, 4450, 350 + MiY + 200
    
    ' -----------------------------------------------------------------------------
    TP.EndDocTP
    'desbloqueo el boton
    cmdImprimirCod.Enabled = True

    '-------------------------------------------------------------------------------
    
    Exit Sub
    
ErrPrintCodigo:
    If Err.Number = 482 Then
        MsgBox "Error en la impresora. Es posible que este apagada o desconectada"
    Else
        MsgBox "Se ha detectado el error:" + CStr(Err.Number) + vbCrLf + Err.Description
    End If
    
End Sub

Private Sub cmdModificar_Click(Index As Integer)
    Dim sMonto As String
    Dim IDp As Long
    Dim IDC As Long
    
    If tbrBuscadorP.GetLstSel = "" Then Exit Sub
    
    IDp = CLng(tbrBuscadorP.GetLstSel(0))
    
    Select Case Index
        Dim UUs As Long
        
        Case 0 'precio
            sMonto = Replace(InputBox("Ingrese nuevo Precio de Venta", "Nuevo Precio"), ".", ",")
            If sMonto = "" Then Exit Sub
            
            If Not IsNumeric(sMonto) Then
                MsgBox "Cargue correctamente los montos", vbExclamation, "Atención"
                Exit Sub
            End If
            
            sMonto = CStr(CSng(sMonto))
            
            DB.EXECUTE "UPDATE Productos SET pVenta = " + Replace(sMonto, ",", ".") + _
                " WHERE ID = " + CStr(IDp)
            tbrBuscadorP.Text = tbrBuscadorP.GetLstSel(2)
            tbrBuscadorP.SetFocus
            tbrBuscadorP.SelLength = Len(tbrBuscadorP.Text)
        Case 1 'costo 'CUIDADO!!!!!!!!!!!!!!!!!!!
            'veo si el usuario que esta trabajando tiene habilitacion para entrar
            UUs = ACC.UltUsuarioIngresado
            
            If ACC.ExisteRelacion(UUs, 9) = 0 Then
                MsgBox ACC.GetNombre("Usuario", "Usuarios", UUs) + " no está habilitado " + _
                    "para ingresar." + vbCrLf + _
                    "Debe Cambiar Sesión a la de un usuario habilitado", vbExclamation, "Atención"
                Exit Sub
            End If
            
            'registro el movimiento(9:Modificar Producto)
            ACC.RegEvento UUs, 9, "Modificación del Costo del Producto " + CStr(IDp)
            
            'registro perdida o ganancia por tenencia
            Dim CtoV As Single, CtoN As String
            
            CtoV = CSng(lblCosto)
            CtoN = Replace(InputBox("Ingrese el Nuevo Costo sólo si se encuentra " + _
                "completamente seguro de realizar este cambio", "Nuevo Costo"), ".", ",")
            If CtoN = "" Then Exit Sub
            
            If Not IsNumeric(CtoN) Then
                MsgBox "Valor Incorrecto", vbExclamation, "Atención"
                Exit Sub
            Else
                'recien aca veo que esta todo bien
                Dim DifC As Single, DifPesoC As Single
                DifC = CSng(CtoN) - CtoV
                DifPesoC = CLng(lblStock) * DifC
                 
                 'primero lo facil modifico el costo en la base de datos
                DB.EXECUTE "UPDATE Productos SET pCosto = " + _
                   Replace(CStr(CSng(CtoN)), ",", ".") + _
                   " WHERE nproducto ='" + tbrBuscadorP.GetLstSel(2) + "'"
                 
                 'ahora ajusto mercaderias en inventario a Res Fin y por tenencia
                 'si sobra si falta se hace al revés por quedar negativo
                 'lo tiro a capital si es la primera carga
                
                PC.Asiento "54", CStr(DifPesoC), "23", CStr(DifPesoC), , "Ajuste manual"
                 
            End If
                                    
            If IsNumeric(CFG.GetInfo(50, 4)) Then
                If CLng(CFG.GetInfo(50, 4)) > 0 Then
                    If MsgBox("Costo Modificado correctamente." + vbCrLf + _
                        "¿Desea Ajustar el precio según el margen de venta?", _
                        vbInformation + vbYesNo, "Modificar Precio") = vbYes Then
                        Dim ClsP As New clsProducto
                            ClsP.PonerPrecioPorMargen CLng(tbrBuscadorP.GetLstSel), _
                                CSng(CtoN)
                        Set ClsP = Nothing
                    End If
                End If
            End If
            tbrBuscadorP.Text = tbrBuscadorP.GetLstSel(2)
            tbrBuscadorP.SetFocus
            tbrBuscadorP.SelLength = Len(tbrBuscadorP.Text)
            
        Case 2 'Stock
            
            UUs = ACC.UltUsuarioIngresado
            
            If ACC.ExisteRelacion(UUs, 9) = 0 Then
                MsgBox ACC.GetNombre("Usuario", "Usuarios", UUs) + " no está habilitado " + _
                    "para ingresar." + vbCrLf + _
                    "Debe Cambiar Sesión a la de un usuario habilitado", vbExclamation, "Atención"
                Exit Sub
            End If
            
            'registro el movimiento(9:Modificar Producto)
            ACC.RegEvento UUs, 9, "Modificación del Stock del producto " + CStr(IDp)
'
            'registro perdida o ganancia por diferencias de stock
            Dim stockV As Long, StockN As String
            
            stockV = CLng(lblStock)
            StockN = Replace(InputBox("Ingrese el Stock controle de haberlo recontado " + _
                "correctamente. Ingrese sólo valores enteros", "Nuevo Stock"), ",", ".")
            If Not IsNumeric(StockN) Then
                MsgBox "Valor Incorrecto", vbExclamation, "Atención"
                Exit Sub
            Else
                If InStrRev(StockN, ".") <> 0 Then 'es decimal
                    MsgBox "Valor Incorrecto", vbExclamation, "Atención"
                    Exit Sub
                Else 'recien aca veo que esta todo bien
                    Dim DifS As Long, DifPesos As Single
                    DifS = CLng(StockN) - stockV
                    DifPesos = CSng(lblCosto) * DifS
                    
                    'primero lo facil modifico stock
                    Dim CPro As New clsProducto, tmPID As Long
                    
                    tmPID = CLng(DB.GetValInRS("Productos", "Id", _
                        "nproducto ='" + tbrBuscadorP.GetLstSel(2) + "'"))
                        
                    CPro.ModificarStock tmPID, CSng(DifS), , "Modificación Manual"
                    Set CPro = Nothing
                    
                    'ahora ajusto mercaderias en inventario a Diferencias de stock
                    'si sobra si falta se hace al revés por quedar negativo

                    PC.Asiento "54", CStr(DifPesos), "35", CStr(DifPesos), , _
                        "Modificación Stock, IdProducto: " + CStr(tmPID) + " de CASA CENTRAL"
                    
                    tbrBuscadorP.Text = tbrBuscadorP.GetLstSel(2)
                    tbrBuscadorP.SetFocus
                    tbrBuscadorP.SelLength = Len(tbrBuscadorP.Text)
                End If
            End If
            
            
        Case 3 'detalle
            DB.EXECUTE "UPDATE Productos SET Observaciones = '" + txtDetalle + "'" + _
                " WHERE ID = " + CStr(IDp)
                    
            Exit Sub
        Case 4
            Dim IdTT As Long
            
            IdTT = DB.GetValInRS("TipoProductos", "ID2", "TipoProducto = '" + _
                cmbTipoM + "'", False)
            DB.EXECUTE "UPDATE Productos SET IdTipoProducto = " + CStr(IdTT) + _
                " WHERE ID = " + CStr(IDp)
            
            Exit Sub
        Case 6 'envases
            DB.EXECUTE "UPDATE Productos SET CodEnvase = '" + cmbEnvases + "'" + _
                " WHERE ID = " + CStr(IDp)
        
            Exit Sub
       
        Case 7 'Codigo de Barras
            DB.EXECUTE "UPDATE Productos SET CodigodeBarras = '" + txtCodigoBarras + "'" + _
                " WHERE ID = " + CStr(IDp)
            
            Exit Sub
        
        Case 5 '--------------- PRODUCTO ------------------------------------
             'veo si cambio algo si no no hago nada
            If txtnProductoM = tbrBuscadorP.GetLstSel(2) Then Exit Sub
            
             'ver que no se repita
            If DB.ContarReg("SELECT nProducto FROM Productos WHERE nProducto= '" + _
                txtnProductoM + "'") > 0 Then
                MsgBox "¡Ya tiene un producto con ese nombre!", vbExclamation, "Atención"
                txtnProductoM = tbrBuscadorP.GetLstSel(2)
                PintarTxt txtnProductoM
                Exit Sub
            End If
            
            DB.EXECUTE "UPDATE Productos SET nProducto = '" + txtnProductoM + "'" + _
                " WHERE ID = " + CStr(IDp)
            
            'negrada pero lo tengo que hacer 2 veces
            DB.EXECUTE "UPDATE Productos SET nProducto = '" + txtnProductoM + "'" + _
                " WHERE ID = " + CStr(IDp)
            
            H = txtnProductoM.Text
            tbrBuscadorP.Text = ""
            tbrBuscadorP.Text = H
            
            
        Case 8 '-------------- CONF DTO VTA CONTADO --------------------------------
            If txtDtoCtado = "" Then txtDtoCtado = "0"
            If Not IsNumeric(txtDtoCtado) Then txtDtoCtado = "0"
            'no se porque pero le tengo que sacar el %
            Dim SP() As String
            SP = Split(txtDtoCtado, "%")
            txtDtoCtado = SP(0)
                
            'si tiene configuracion particular modifico, si no agrego
            IDC = CFG.ExistePropiedad("DVC " + CStr(IDp))
            
            If IDC = 0 Then
                CFG.AgregarNodo 60, "DVC " + CStr(IDp), "", CStr(CSng(txtDtoCtado)), 0
            Else
                CFG.ModificarNodo IDC, , , , CStr(CSng(txtDtoCtado))
            End If
            
            txtDtoCtado = FormatPercent(ValidarNumeros(txtDtoCtado) / 100)
                
            Exit Sub
            
        Case 9 '------------------ CONFIGURACION STOCK MINIMO -----------------------
            If txtStock = "" Then txtDtoCtado = "0"
            If Not IsNumeric(txtStock) Then txtStock = "0"
            'si tiene configuracion particular modifico, si no agrego
            IDC = CFG.ExistePropiedad("STM " + CStr(IDp))
            
            If IDC = 0 Then
                CFG.AgregarNodo 19, "STM " + CStr(IDp), "", CStr(CLng(txtStock)), 0
            Else
                CFG.ModificarNodo IDC, , , , CStr(CLng(txtStock))
            End If
            Exit Sub
    End Select
    
    tbrBuscadorP.Recargar
    tbrBuscadorP.SetFocus
    tbrBuscadorP.SelLength = Len(tbrBuscadorP.Text)
End Sub

Private Sub cmdVerC_Click()
    Dim UUs As Long
    If LCase(cmdVerC.Caption) = "ver costo" Then
        'veo si el usuario que esta trabajando tiene habilitacion para entrar -------
        UUs = ACC.UltUsuarioIngresado
        
        If ACC.ExisteRelacion(UUs, 25) = 0 Then
            MsgBox ACC.GetNombre("Usuario", "Usuarios", UUs) + " no está habilitado " + _
                "para ingresar." + vbCrLf + _
                "Debe Cambiar Sesión a la de un usuario habilitado", vbExclamation, "Atención"
            Exit Sub
        End If
        
        'registro el movimiento(25:Ver Costos y Otros)
        ACC.RegEvento UUs, 25, "Ver Costo del producto" + CStr(IDp)
        '------------------------------------------------------------------
        
        cmdVerC.Caption = "Esconder Cto"
        lblCosto.Visible = True
        cmdModificar(1).Visible = True
    Else
        cmdVerC.Caption = "Ver Costo"
        lblCosto.Visible = False
        cmdModificar(1).Visible = False
    End If
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Dim H As String, TmP As String
    
    If cmbTipo = "" Then
        MsgBox "No tiene cargado ningún tipo de Producto, Agregue el tipo " + _
            "correspondiente", vbInformation, "Atención"
        Exit Sub
    End If
    
    If txtnProducto = "" Then Exit Sub
        
    'reemplazo comillas simples
    txtnProducto = Replace(txtnProducto, "'", " ")
    txtDetalleN = Replace(txtDetalleN, "'", " ")
    
    'ver que no se repita
    If DB.ContarReg("select nproducto from productos where nproducto= '" + _
        txtnProducto + "'") > 0 Then
        MsgBox "¡Ya tiene un producto con ese nombre!", vbExclamation, "Atención"
        PintarTxt txtnProducto
        Exit Sub
    End If
        
    If DB.ContarReg("select nproducto from productos") > 20 Then
        If LIC.GetLic < Licencia1 Then
            MsgBox "No dispone aun de licencia para seguir cargando productos " + vbCrLf + _
                "Contacte con info@tbrsoft.com", vbCritical, "Sin licencia"
            Exit Sub
        End If
    End If
    
    On Local Error GoTo ErrDuplic
    
DeGuelta:
    TmP = IdAutonum("Productos")
    DB.EXECUTE "INSERT INTO Productos (ID,IdTipoProducto,Nproducto,Observaciones,CodEnvase)" + _
        " VALUES (" + TmP + ",'" + _
        CStr(DB.GetValInRS("TipoProductos", "ID2", "TipoProducto = '" + cmbTipo + "'", False)) + _
        "','" + txtnProducto + "','" + txtDetalleN + "','" + _
        "No Tiene')"
        
    'negrada pero lo tengo que hacer 2 veces (actualizo el recien agregado)
    DB.EXECUTE "UPDATE Productos SET nProducto = '" + txtnProducto + "'" + _
        " WHERE ID = " + CStr(TmP)
            
    H = txtnProducto.Text
    tbrBuscadorP.Text = ""
    tbrBuscadorP.Text = H
        
    txtnProducto = ""
    txtDetalleN = ""
    Exit Sub
    
ErrDuplic:
    If Err.Number = -2147467259 Then
        GoTo DeGuelta
    Else
        MsgBox Err.Number + ": " + Err.Description
    End If

End Sub

Public Sub AbrirDatos(nProducto As String)
    CargarDatos
    
    tbrBuscadorP.Contrasena = Contrasena
    tbrBuscadorP.ArchivoMDB = ArchivoMDBPrincipal
    tbrBuscadorP.SqlSinLike = "SELECT TOP 50 Productos.ID, " + _
        "TipoProductos.TipoProducto, Productos.nProducto " + _
        "FROM TipoProductos INNER JOIN Productos ON TipoProductos.ID2 = " + _
        "Productos.IdTipoProducto WHERE Productos.ID >=0"
    tbrBuscadorP.OrderBy = "ORDER BY ID"
    tbrBuscadorP.CampoEnQueBuscar = "Id/n,TipoProducto,nproducto/b"
    tbrBuscadorP.Text = nProducto
    tbrBuscadorP.ColumnasSepPorComasyParentesis = "ID(600)/Tipo(1500)/Producto(2500)"
        
    Me.Show 1

End Sub

Private Sub Command3_Click(Index As Integer)
    If Index = 2 Then
        frmTipo.Iniciar "Envases"
    Else
        frmTipoProducto.Show 1
    End If
End Sub

Private Sub Command4_Click()
    Dim AA As String
    AA = CFGBD.GetInfo(82, 4) + "img\" + CStr(tbrBuscadorP.GetLstSel(0)) + ".txt"
    
    Dim TE As TextStream
    Set TE = FSO.OpenTextFile(AA, ForWriting, True)
        TE.Write txtTextoPV.Text
    TE.Close
    
    MsgBox "Se grabo OK el texto"
End Sub

Private Sub Command5_Click()
    Dim CM As New CommonDialog
    CM.DialogTitle = "Elija una imagen para el producto"
    CM.Filter = "Imagenes(jpg bmp gif tiff)|*.jpg; *.jpeg; *.bmp;*.gif;*.tiff"
    CM.InitDir = AP
    CM.ShowOpen
    
    Dim F As String
    F = CM.FileName
    If F = "" Then Exit Sub
    
    'conseguirle un buen nombre y ubicarla
    Dim J As Long, F2 As String
    For J = 0 To 20
        F2 = CFGBD.GetInfo(82, 4) + "img\" + CStr(tbrBuscadorP.GetLstSel(0)) + "-" + CStr(J) + ".jpg"
        If FSO.FileExists(F2) = False Then
            FSO.CopyFile F, F2, False
            Exit For
        End If
        'si hay 20 se caga y no la copia
    Next J
    
    lstIMGS.AddItem FSO.GetBaseName(F2) + "." + FSO.GetExtensionName(F2)
    'la eligo y por lo tanto la muestro
    lstIMGS.ListIndex = lstIMGS.ListCount - 1
    
End Sub

Private Sub Command6_Click()
    'ver cual esta elegida!
        
    If lstIMGS.ListIndex = -1 Then
        MsgBox "Elija la imagen que desea eliminar!"
        Exit Sub
    End If

    Dim Elegida As Long
    Elegida = lstIMGS.ListIndex
    
    'borrar la imagen del disco ademas de sacarla de aca
    FSO.DeleteFile CFGBD.GetInfo(82, 4) + "img\" + lstIMGS.List(lstIMGS.ListIndex)
    'ahora acomodar todo!
    lstIMGS.RemoveItem lstIMGS.ListIndex
    'dejo elegido el primero
    If lstIMGS.ListCount > 0 Then
        lstIMGS.ListIndex = 0
    Else
        imgPV.Picture = LoadPicture
    End If
End Sub

Private Sub Command7_Click()
    Dim sTmp As String
    sTmp = CFGBD.GetInfo(82, 4) + "img\" + CStr(tbrBuscadorP.GetLstSel(0)) + ".txt"
    If FSO.FileExists(sTmp) Then FSO.DeleteFile sTmp
    
    txtTextoPV.Text = ""
End Sub

Private Sub fBoton1_Click()
    Dim IDPR As Long
    IDPR = tbrBuscadorP.GetLstSel(0)
    
    'ver que no tenga stock!!!
    'si es asi le decimos que lo de de baja
    Dim SS As String
    SS = "SELECT ID, Stock FROM Productos WHERE ID =" + CStr(IDPR)
        
    Dim rS As New ADODB.Recordset
    rS.CursorLocation = adUseClient
    rS.Open SS, DB.CN, adOpenStatic, adLockReadOnly
    
    If rS.RecordCount = 1 Then
        If rS.Fields("stock") <> 0 Then
            MsgBox "ESTE PRODUCTO TIENE STOCK. LLEVELO A CERO PARA PODER ELIMIAR EL PRODUCTO"
            Exit Sub
        End If
    Else
        MsgBox "No se encontro el producto en la tabla!"
        Exit Sub
    End If
    
    On Local Error Resume Next
    
    DB.EXECUTE "DELETE FROM productos WHERE id=" + CStr(IDPR)
    tbrBuscadorP.Text = ""
    tbrBuscadorP.Text = "NADA"
    tbrBuscadorP.Text = ""
    tbrBuscadorP.Recargar
End Sub

Private Sub Form_Activate()
    tbrBuscadorP.Recargar 'cuando nproducto es "" no recarga
    
    CargarCombo cmbTipo, "SELECT Tipoproducto FROM Tipoproductos " + _
        "WHERE ID2 > 0 ORDER BY Tipoproducto", _
        "Tipoproducto"
    CargarCombo cmbTipoM, "SELECT TipoProducto FROM Tipoproductos " + _
        "WHERE ID2 > 0 ORDER BY TipoProducto", _
        "TipoProducto"
    CargarCombo cmbEnvases, "SELECT Envase FROM Envases ORDER BY Envase", _
        "Envase"
    
    '¿Tiene envases? !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    If CFG.GetInfo(5, 4) = "No" Then
        Label11.Visible = False
        cmbEnvases.Visible = False
        cmdModificar(6).Visible = False
        command3(2).Visible = False
    Else
        Label11.Visible = True
        cmbEnvases.Visible = True
        cmdModificar(6).Visible = True
        command3(2).Visible = True
    End If
End Sub

Private Sub CargarDatos()
    If tbrBuscadorP.GetLstSel = "" Then
        lblProdSel = ""
        Exit Sub
    End If
    
    lblPrecio = FormatCurrency(DB.GetValInRS("Productos", "pVenta", "ID= " + _
        CStr(tbrBuscadorP.GetLstSel(0)), False), , , , vbFalse)
    lblCosto = FormatCurrency(DB.GetValInRS("Productos", "pcosto", "ID= " + _
        CStr(tbrBuscadorP.GetLstSel(0)), False), , , , vbFalse)
    lblStock = DB.GetValInRS("Productos", "stock", "ID= " + _
        CStr(tbrBuscadorP.GetLstSel(0)), False)
    txtCodigoBarras = NoNuloS(DB.GetValInRS("Productos", "CodigodeBarras", "ID= " + _
        CStr(tbrBuscadorP.GetLstSel(0))))
    txtDetalle = DB.GetValInRS("Productos", "Observaciones", "ID= " + _
        CStr(tbrBuscadorP.GetLstSel(0)))
    txtnProductoM = DB.GetValInRS("Productos", "nProducto", "ID= " + _
        CStr(tbrBuscadorP.GetLstSel(0)))
    cmbTipoM.Text = tbrBuscadorP.GetLstSel(1)
    cmbTipo.Text = tbrBuscadorP.GetLstSel(1)
    cmbEnvases.Text = DB.GetValInRS("Productos", "CodEnvase", "ID= " + _
        CStr(tbrBuscadorP.GetLstSel(0)))
    lblProdSel = "Seleccionado: " + tbrBuscadorP.GetLstSel(2)
        
    'CONFIGURACION DTO VTA CONTADO ------------------------------------------------
    'si tiene configuracion particular la pongo si no uso la general
    Dim IDC As Long
    IDC = CFG.ExistePropiedad("DVC " + tbrBuscadorP.GetLstSel)
    
    If IDC = 0 Then
        txtDtoCtado = FormatPercent(CSng(CFG.GetInfo(60, 4)) / 100)
    Else
        txtDtoCtado = FormatPercent(CSng(CFG.GetInfo(IDC, 4)) / 100)
    End If
    ' -------------------------------------------------------------------------------
        
    'CONFIGURACION STOCK MINIMO ------------------------------------------------
    'si tiene configuracion particular la pongo si no uso la general
    IDC = CFG.ExistePropiedad("STM " + tbrBuscadorP.GetLstSel)
    
    If IDC = 0 Then
        txtStock = CFG.GetInfo(19, 4)
    Else
        txtStock = CFG.GetInfo(IDC, 4)
    End If
    ' -------------------------------------------------------------------------------
        
    'CONFIGURACION IVA PREDETERMINADO ------------------------------------------------
    'si tiene configuracion particular la pongo si no uso la general
    IDC = CFG.ExistePropiedad("IVA " + tbrBuscadorP.GetLstSel)
    
    If IDC = 0 Then
        txtIVA = CFG.GetInfo(7, 4)
    Else
        txtIVA = CFG.GetInfo(IDC, 4)
    End If
    ' -------------------------------------------------------------------------------
        
    'ANDRES
    FR_Descripcion.Caption = "Descripcion para el punto de venta producto=" + _
        CStr(tbrBuscadorP.GetLstSel(0))
    txtTextoPV.Text = ""
    'si existiera el detalle lo mostramos
    Dim sTmp As String
    sTmp = CFGBD.GetInfo(82, 4) + "IMG\" + CStr(tbrBuscadorP.GetLstSel(0)) + ".txt"
    If FSO.FileExists(sTmp) Then
        Dim TE As TextStream
        Set TE = FSO.OpenTextFile(sTmp, ForReading, False)
            If TE.AtEndOfStream Then
                txtTextoPV.Text = ""
            Else
                txtTextoPV.Text = TE.ReadAll
            End If
        TE.Close
    End If
    
    'IMAGENES
    fr_IMAGENES.Caption = "Imagenes del producto=" + CStr(tbrBuscadorP.GetLstSel(0))
    
    lstIMGS.Clear
    imgPV.Picture = LoadPicture
    Dim J As Long
    Dim sEXTs()
    sEXTs = Array("jpg", "jpeg", "bmp", "gif", "tiff")
    For J = 0 To 20 'mas de 20 imagenes es degenerado
        Dim K As Long
        For K = 0 To UBound(sEXTs)
            sTmp = CFGBD.GetInfo(82, 4) + "IMG\" + CStr(tbrBuscadorP.GetLstSel(0)) + "-" + CStr(J) + "." + sEXTs(K)
            If FSO.FileExists(sTmp) Then
                'encontre una imagen
                lstIMGS.AddItem FSO.GetBaseName(sTmp) + "." + FSO.GetExtensionName(sTmp)
                Exit For 'ya no hay nada que buscar con este numero
            End If
        Next K
    Next J
    'si hay alguna imagen mostrar la primera!
    If lstIMGS.ListCount > 0 Then lstIMGS.ListIndex = 0
End Sub

Private Sub LoadImgPV(iAR As String)
    imgPV.Stretch = False 'para que tome el tamaño que tiene que ser
    imgPV.Picture = LoadPicture(iAR)
    'ver ahi la proporcion
    Dim Prop As Single
    Prop = imgPV.Width / imgPV.Height
    'definbir el final segun corresponda
    Dim Ancho As Single, Alto As Single
    'probar si entraria con ancho maximo ...
    Ancho = PicFondo.Width
    Alto = Ancho / Prop
    If Alto > PicFondo.Height Then
        'cambiar todo! supuestamente si fallo el otro este no falla
        Alto = PicFondo.Height
        Ancho = Alto * Prop
    End If
    
    imgPV.Stretch = True
    imgPV.Width = Ancho
    imgPV.Height = Alto
    
    imgPV.Top = PicFondo.Height / 2 - imgPV.Height / 2
    imgPV.Left = PicFondo.Width / 2 - imgPV.Width / 2
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If tbrBuscadorP.GetLstSel = "" Then Exit Sub
    
    If KeyCode = 123 And Shift = 2 Then 'Ctrl+F12 solo para CASA ARIAS !!!!!!!!!!!!!!
        'modifico el detalle de esta manera
        'Costo+0+precio+vbcrlf+DtoVentaContado
        Dim Que As String
        
        Que = CStr(Round(CSng(lblCosto), 0)) + "0" + CStr(Round(CSng(lblPrecio), 0)) + _
            vbCrLf + CStr(Round(CSng(Replace(txtDtoCtado, "%", "")), 0))
        
        txtDetalle = Que
        DB.EXECUTE "UPDATE Productos SET Observaciones = '" + Que + "'" + _
                " WHERE ID = " + tbrBuscadorP.GetLstSel
    End If
    
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    tbrBuscadorP.SetDelayEventClick 1
    
    CargarCombo cmbTipo, "SELECT Tipoproducto FROM Tipoproductos " + _
        "WHERE ID2 > 0 ORDER BY Tipoproducto", _
        "Tipoproducto"
    CargarCombo cmbTipoM, "SELECT TipoProducto FROM Tipoproductos " + _
        "WHERE ID2 > 0 ORDER BY TipoProducto", _
        "tipoproducto"
    CargarCombo cmbEnvases, "SELECT Envase FROM Envases ORDER BY Envase", _
        "Envase"
    
    ActuTbrB = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    tbrBuscadorP.CN_CLOSE
End Sub

Private Sub lstIMGS_Click()
    Dim sTmp As String
    sTmp = CFGBD.GetInfo(82, 4) + "IMG\" + lstIMGS.List(lstIMGS.ListIndex)
    If FSO.FileExists(sTmp) Then
        LoadImgPV sTmp
    Else
        'la imagen elegida no existe ¿La quitamos de la lista?"
    End If
End Sub

Private Sub tbrBuscadorP_Change()
    Dim tbR As String, Nombre As String
    Dim I As Long
    
    If Not IsNumeric(tbrBuscadorP.Text) Then
        'busca el codigo normal
        tbrBuscadorP.CampoEnQueBuscar = "id/n,Tipoproducto,nproducto/b"
    Else
        'si cargo con código de barras (supongo que :
        ' tiene como mínimo tiene 7 dígitos y es NUMÉRICO
        
        '1ro leo lo QUE escribio
        tbR = tbrBuscadorP.Text
        
        'veo si tiene mas de 6 digitos es CODIGO DE BARRAS
        If Len(tbR) > 6 Then
            Nombre = DB.GetValInRS("Productos", "nProducto", "CodigodeBarras = '" + tbR + "'")
                
            If Nombre = "" Then 'no lo encuentro
                'tbrBuscadorP.Text = ""
                'Exit Sub NO porque si no no lo encuentra mas
            Else 'si lo encontro -> lo agrego directamente
                'pongo el nombre en el tbrBuscador
                tbrBuscadorP.Text = Nombre
                
                'veo cuantos filtro
                Select Case tbrBuscadorP.ListCount
                    Case 0 'no encontro ninguno (no deberia pasar pero ....)
                            ' anulo todo
                        'tbrBuscadorP.Text = "" NO SI NO NO LO ENCUENTRA MAS
                        'Exit Sub
                    Case 1 'joya ... dale que va!
                        tbrBuscadorP.SelStart = 0
                        tbrBuscadorP.SelLength = Len(Nombre)
                    Case Is > 1 ' tengo que buscar 1 x 1 hasta que coincida exac.
                                ' puede pasar que si es muy cortito el nombre
                                ' y se filtren mucho, como este muestra 50 nomas
                                ' puede pasar que no lo encuentre ... ojala que no
                                ' tener cuenta en el futuro XXXXX
                        For I = 1 To tbrBuscadorP.ListCount
                            If Nombre = tbrBuscadorP.GetLstSel(1, I) Then
                                'tiene que quedar elegido -> los otros se borraron
                                tbrBuscadorP.SelStart = 0
                                tbrBuscadorP.SelLength = Len(Nombre)
                                Exit For
                            Else
                                'tengo que borrar porque no tengo un procedimiento
                                'que diga que renglon quede SELECTED
                                tbrBuscadorP.BorrarRenglon I
                            End If
                        Next I
                End Select
                
            End If
            
            'Exit Sub 'encuentre o no
            
        Else
            'busca el codigo normal
            tbrBuscadorP.CampoEnQueBuscar = "id/b,Tipoproducto,nproducto"
        End If
    End If
    
    If tbrBuscadorP.Text <> "" Then VerActu
    
    If tbrBuscadorP.GetLstSel = "" Then
        lblPrecio = ""
        lblCosto = ""
        lblStock = ""
        txtnProductoM = ""
        cmdModificar(0).Enabled = False
        cmdModificar(1).Enabled = False
        cmdModificar(2).Enabled = False
        cmdModificar(3).Enabled = False
        cmdModificar(4).Enabled = False
        cmdModificar(5).Enabled = False
        cmdModificar(6).Enabled = False
        cmdModificar(7).Enabled = False
        cmdModificar(8).Enabled = False
        cmdModificar(9).Enabled = False
        lstIMGS.Clear
        imgPV.Picture = LoadPicture
    Else
        cmdModificar(0).Enabled = True
        cmdModificar(1).Enabled = True
        cmdModificar(2).Enabled = True
        cmdModificar(3).Enabled = True
        cmdModificar(4).Enabled = True
        cmdModificar(5).Enabled = True
        cmdModificar(6).Enabled = True
        cmdModificar(7).Enabled = True
        cmdModificar(8).Enabled = True
        cmdModificar(9).Enabled = True
    End If
      
    CargarDatos
End Sub

Private Sub tbrBuscadorP_Click()
    CargarDatos
End Sub

Private Sub VerActu()
    If ActuTbrB = False Then
       ActuTbrB = True
       tbrBuscadorP.Recargar
    Else
        ActuTbrB = False
    End If

End Sub

Private Sub tbrBuscadorP_GotFocus()
    tbrBuscadorP_Click
End Sub

Private Sub txtIVA_GotFocus()
    PintarTxt txtIVA
End Sub

Private Sub txtIVA_LostFocus()
    Dim IDC As Long
    
    If txtIVA = "" Then Exit Sub
    If tbrBuscadorP.GetLstSel = "" Then Exit Sub
    
    If Not IsNumeric(txtIVA) Then
        txtIVA = CFG.GetInfo(7, 4)
        Exit Sub
    End If
    
    'si tiene configuracion particular modifico, si no agrego
    'distinta a la predeterminado
    If CSng(txtIVA) <> CSng(CFG.GetInfo(7, 4)) Then
        IDC = CFG.ExistePropiedad("IVA " + tbrBuscadorP.GetLstSel)
        
        If IDC = 0 Then
            CFG.AgregarNodo 7, "IVA " + tbrBuscadorP.GetLstSel, "", txtIVA, 0
        Else
            CFG.ModificarNodo IDC, , , , txtIVA
        End If
    End If
End Sub

Private Sub txtnProducto_GotFocus()
    Command2.Default = True
End Sub

Private Sub txtnProductoM_GotFocus()
    cmdModificar(5).Default = True
End Sub
