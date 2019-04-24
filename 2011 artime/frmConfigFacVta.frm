VERSION 5.00
Object = "{A7FBD38D-2930-49E3-B60C-9E0202D84549}#15.0#0"; "tbrControles.ocx"
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmConfigFacVta 
   BackColor       =   &H004E4E4E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configurar Conceptos en Factura de Venta"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10515
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConfigFacVta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin tbrFaroButton.fBoton cmdFacaMos 
      Height          =   465
      Left            =   8880
      TabIndex        =   4
      Top             =   5550
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Grabar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdFacSin 
      Height          =   465
      Left            =   8880
      TabIndex        =   6
      Top             =   6060
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Grabar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdQuitar 
      Height          =   435
      Left            =   7755
      TabIndex        =   2
      Top             =   4710
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Quitar Último Concepto"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdAgregar 
      Height          =   435
      Left            =   5460
      TabIndex        =   10
      Top             =   4710
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Agregar Concepto"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.TextBox txtFacSin 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7800
      TabIndex        =   5
      Top             =   6120
      Width           =   1005
   End
   Begin VB.TextBox txtFacaMos 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7800
      TabIndex        =   3
      Top             =   5610
      Width           =   1005
   End
   Begin tbrControles.tbrBuscador tbrBuscadorC 
      Height          =   1905
      Left            =   120
      TabIndex        =   0
      Top             =   4170
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   3360
      BackColor       =   5131854
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
   Begin VB.TextBox txtConcepto 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5550
      TabIndex        =   1
      Top             =   4170
      Width           =   4275
   End
   Begin tbrControles.MouTextBox txtIVA 
      Height          =   375
      Left            =   4770
      TabIndex        =   7
      Top             =   6900
      Width           =   945
      _ExtentX        =   1667
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
   Begin tbrFaroButton.fBoton cmdIVA 
      Height          =   465
      Left            =   6270
      TabIndex        =   8
      Top             =   6810
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Grabar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdSalir 
      Height          =   465
      Left            =   8760
      TabIndex        =   9
      Top             =   6810
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.Label Label17 
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
      Height          =   330
      Left            =   5790
      TabIndex        =   35
      Top             =   6930
      Width           =   285
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Facturas a mostrarse en IVA Ventas"
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
      Height          =   315
      Left            =   6000
      TabIndex        =   34
      Top             =   5280
      Width           =   3765
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Facturas sin discriminar IVA"
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
      Left            =   4860
      TabIndex        =   33
      Top             =   6150
      Width           =   2805
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sin Espacios y sin comas"
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
      Height          =   345
      Left            =   7110
      TabIndex        =   32
      Top             =   6540
      Width           =   2385
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Facturas discriminado IVA "
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
      Height          =   435
      Left            =   5340
      TabIndex        =   31
      Top             =   5670
      Width           =   2385
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "No necesita ser agregado como concepto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   10
      Left            =   3300
      TabIndex        =   30
      Top             =   7350
      Width           =   3915
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Detalle del concepto"
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
      Index           =   9
      Left            =   5580
      TabIndex        =   29
      Top             =   3900
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione cuenta a imputar por el concepto 3"
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
      Index           =   8
      Left            =   570
      TabIndex        =   28
      Top             =   3900
      Width           =   4665
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Concepto 8"
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
      Index           =   7
      Left            =   5295
      TabIndex        =   27
      Top             =   2940
      Width           =   2475
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Concepto 7"
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
      Index           =   6
      Left            =   5295
      TabIndex        =   26
      Top             =   2100
      Width           =   2475
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Concepto 6"
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
      Index           =   5
      Left            =   5295
      TabIndex        =   25
      Top             =   1230
      Width           =   2475
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Concepto 5"
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
      Index           =   4
      Left            =   5295
      TabIndex        =   24
      Top             =   390
      Width           =   2475
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Concepto 4"
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
      Index           =   3
      Left            =   345
      TabIndex        =   23
      Top             =   2910
      Width           =   2475
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Concepto 3"
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
      Index           =   2
      Left            =   345
      TabIndex        =   22
      Top             =   2070
      Width           =   2475
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Concepto 2"
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
      Index           =   1
      Left            =   345
      TabIndex        =   21
      Top             =   1230
      Width           =   2475
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Concepto 1"
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
      Index           =   0
      Left            =   360
      TabIndex        =   20
      Top             =   360
      Width           =   2475
   End
   Begin VB.Label lblConcepto 
      BackColor       =   &H00D1CFD3&
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
      Height          =   465
      Index           =   7
      Left            =   5295
      TabIndex        =   19
      Top             =   3195
      Width           =   4425
   End
   Begin VB.Label lblConcepto 
      BackColor       =   &H00D1CFD3&
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
      Height          =   465
      Index           =   6
      Left            =   5295
      TabIndex        =   18
      Top             =   2355
      Width           =   4425
   End
   Begin VB.Label lblConcepto 
      BackColor       =   &H00D1CFD3&
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
      Height          =   465
      Index           =   5
      Left            =   5295
      TabIndex        =   17
      Top             =   1500
      Width           =   4425
   End
   Begin VB.Label lblConcepto 
      BackColor       =   &H00D1CFD3&
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
      Height          =   465
      Index           =   4
      Left            =   5295
      TabIndex        =   16
      Top             =   660
      Width           =   4425
   End
   Begin VB.Label lblConcepto 
      BackColor       =   &H00D1CFD3&
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
      Height          =   465
      Index           =   3
      Left            =   345
      TabIndex        =   15
      Top             =   3195
      Width           =   4425
   End
   Begin VB.Label lblConcepto 
      BackColor       =   &H00D1CFD3&
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
      Height          =   465
      Index           =   2
      Left            =   345
      TabIndex        =   14
      Top             =   2355
      Width           =   4425
   End
   Begin VB.Label lblConcepto 
      BackColor       =   &H00D1CFD3&
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
      Height          =   465
      Index           =   1
      Left            =   345
      TabIndex        =   13
      Top             =   1500
      Width           =   4425
   End
   Begin VB.Label lblConcepto 
      BackColor       =   &H00D1CFD3&
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
      Height          =   465
      Index           =   0
      Left            =   345
      TabIndex        =   12
      Top             =   660
      Width           =   4425
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "IVA Predeterminado"
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
      Left            =   2670
      TabIndex        =   11
      Top             =   6960
      Width           =   1905
   End
End
Attribute VB_Name = "frmConfigFacVta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IsVenta As Boolean
Dim StrConc As String

Private Sub cmdAgregar_Click()
    If txtConcepto = "" Then
        MsgBox "Ingrese algún concepto", vbInformation, "Atención"
        Exit Sub
    End If
    
    Dim Cual As Long
    
    If tbrBuscadorC.GetLstSel = "" Then Exit Sub
    
    Cual = CLng(Right(Label1(8), 1))
    
    AgregarOtro (Cual)
    
    txtConcepto = ""
    tbrBuscadorC.Text = ""
    tbrBuscadorC.SetFocus
End Sub

Private Sub AgregarOtro(NroConc As Long)
    '1ro lo agrego como configuracion
    CFG.AgregarNodo 7, StrConc + CStr(NroConc), "", _
        tbrBuscadorC.GetLstSel + "_" + txtConcepto, 0

    AgregarConceptos
End Sub

Private Sub cmdFacaMos_Click()
    Dim cfiG As Long
    If IsVenta Then
        cfiG = 102
    Else
        cfiG = 101
    End If
    
    If txtFacaMos <> "" Then
        txtFacaMos = UCase(txtFacaMos)
        CFG.ModificarNodo cfiG, , , , txtFacaMos
    End If
End Sub

Private Sub cmdFacSin_Click()
    'solo se podra presionar en ventas ya que si no esta invisible
    
    If txtFacSin = "" Then Exit Sub
    Dim TmP As String, FacAM As String, R As Long, Tm2 As String
    
    TmP = txtFacSin
    FacAM = CFG.GetInfo(102, 4)
    txtFacSin = ""
    'comparo letra con letra que no se repita
    For R = 1 To Len(TmP)
        Tm2 = UCase(Mid(TmP, R, 1))
        If InStrRev(FacAM, Tm2, , vbTextCompare) <> 0 Then
            MsgBox Tm2 + " fue borrado ya que ya está configurado como IVA discriminado", vbInformation, "Atención"
        Else
            txtFacSin = txtFacSin + UCase(Tm2)
        End If
    Next R
    
    CFG.ModificarNodo 102, , , txtFacSin
End Sub

Private Sub cmdIVA_Click()
    txtIVA = ValidarNumeros(txtIVA)
    CFG.ModificarNodo 7, , , , txtIVA
End Sub

Private Sub cmdQuitar_Click()
    'quito el concepto segun diga el label1(8) -1
    Dim IdCF As Long
    
    IdCF = CFG.GetID(StrConc + CStr(CLng(Right(Label1(8), 1)) - 1))
    
    CFG.EliminarNodo IdCF
    
    'limpio los label y que los cargue de vuelta
    For IdCF = 0 To 7
        lblConcepto(IdCF) = ""
    Next IdCF
    
    AgregarConceptos
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Public Function AbrirDatos(Optional EsVenta As Boolean = True)
    IsVenta = EsVenta
    
    Me.Show 1
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Dim TmP As String, cfiG As Long
    
    If IsVenta Then
        StrConc = "Concepto "
        cfiG = 102
    Else
        StrConc = "ConceptoCpra "
        Label2 = "Facturas a mostrar en IVA Compras"
        cfiG = 101
        Label3.Visible = False
        txtFacSin.Visible = False
        cmdFacSin.Visible = False
        Label4.Visible = False
        Label5.Top = 6150
    End If
    
    'configuracion para mostrar en Iva Ventas o compras
    TmP = CFG.GetInfo(cfiG, 4)
    If TmP = "" Then TmP = "A"
    txtFacaMos = TmP
    CFG.ModificarNodo cfiG, , , , TmP
    
    'configuracion IVA predeterminado
    TmP = CFG.GetInfo(7, 4)
    If Not IsNumeric(TmP) Then TmP = "0"
    If CLng(TmP) < 0 Then TmP = "0"
    txtIVA = TmP
    CFG.ModificarNodo 7, , , , TmP
    
    If IsVenta Then 'tambien que facturas no se discriminan
        TmP = CFG.GetInfo(cfiG, 3)
        If TmP = "" Then TmP = "B"
        txtFacSin = TmP
        CFG.ModificarNodo cfiG, , , TmP
    End If
    
    AgregarConceptos
    
    tbrBuscadorC.Contrasena = "zuliani"
    tbrBuscadorC.ArchivoMDB = CFGBD.GetInfo(81, 4) + "Ctas.mdb"
    tbrBuscadorC.SqlSinLike = "SELECT TOP 50 Id, Nombre FROM tblCuentas WHERE ID>5"
    tbrBuscadorC.CampoEnQueBuscar = "Id/n,Nombre/b"
    tbrBuscadorC.ColumnasSepPorComasyParentesis = "ID(700)/Nombre(2650)"
    tbrBuscadorC.Recargar
End Sub

Private Sub AgregarConceptos()
    Dim YasTa As Boolean, X As Long, IdCF As Long, SP() As String

    YasTa = False: X = 1
    'agrego los conceptos
    Do While Not YasTa = True
        IdCF = CFG.ExistePropiedad(StrConc + CStr(X))
            'estos conceptos los pongo todos como hijos de 7 (iva predet)
        
        If IdCF = 0 Then 'no hay mas conceptos corto aca
            YasTa = True
            lblConcepto(X) = ""
        Else
            SP = Split(CFG.GetInfo(IdCF, 4), "_") 'tiene IdCta_DetalleConcepto
            
            If PC.GetNameCuenta(CLng(SP(0))) <> "NO EXISTE" Then
                lblConcepto(X - 1) = SP(0) + " " + PC.GetNameCuenta(CLng(SP(0))) + ": " + SP(1)
                X = X + 1
            Else
                CFG.EliminarNodo (IdCF)
            End If
        End If
    Loop
    Label1(8) = "Seleccione cuenta a imputar por el concepto " + CStr(X)
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    tbrBuscadorC.CN_CLOSE
End Sub

