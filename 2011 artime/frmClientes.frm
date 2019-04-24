VERSION 5.00
Object = "{A7FBD38D-2930-49E3-B60C-9E0202D84549}#15.0#0"; "tbrControles.ocx"
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmClientes 
   BackColor       =   &H0049453D&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10935
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmClientes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin tbrControles.MouTextBox txtIngreso 
      Height          =   375
      Left            =   4770
      TabIndex        =   11
      Top             =   5340
      Width           =   1035
      _ExtentX        =   1826
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
   Begin VB.TextBox txtAntiguedad 
      BackColor       =   &H00FFFFFF&
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
      Left            =   2100
      TabIndex        =   10
      Top             =   5370
      Width           =   1245
   End
   Begin VB.CheckBox chkTrabaja 
      BackColor       =   &H00444444&
      Caption         =   "Tiene Trabajo"
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   1530
      TabIndex        =   6
      Top             =   4020
      Width           =   1755
   End
   Begin VB.TextBox txtTrabajo 
      BackColor       =   &H00FFFFFF&
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
      Left            =   2100
      TabIndex        =   8
      Top             =   4440
      Width           =   3700
   End
   Begin VB.TextBox txtTelefono2 
      BackColor       =   &H00FFFFFF&
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
      Left            =   2100
      TabIndex        =   9
      Top             =   4890
      Width           =   3700
   End
   Begin VB.TextBox txtTelefono 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   2100
      TabIndex        =   4
      Top             =   3120
      Width           =   3700
   End
   Begin VB.TextBox txtMail 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   2100
      TabIndex        =   5
      Top             =   3570
      Width           =   3700
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0049453D&
      Caption         =   "Cuotas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   6090
      TabIndex        =   27
      Top             =   5130
      Width           =   3675
      Begin tbrControles.MouTextBox txtDias 
         Height          =   375
         Left            =   2100
         TabIndex        =   17
         Top             =   210
         Width           =   825
         _ExtentX        =   1455
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
         Largo           =   4
         Entero          =   -1  'True
      End
      Begin VB.OptionButton chkUnico 
         BackColor       =   &H0049453D&
         Caption         =   "Cuota Única a los"
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
         Left            =   150
         MaskColor       =   &H00544B45&
         TabIndex        =   15
         Top             =   210
         Value           =   -1  'True
         Width           =   1965
      End
      Begin VB.OptionButton chkCuotas 
         BackColor       =   &H0049453D&
         Caption         =   "En Cuotas"
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
         Left            =   210
         MaskColor       =   &H00544B45&
         TabIndex        =   28
         Top             =   720
         Width           =   1875
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
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
         Height          =   645
         Left            =   3090
         TabIndex        =   29
         Top             =   330
         Width           =   675
      End
   End
   Begin tbrControles.tbrBuscador tbrBuscadorF 
      Height          =   1695
      Left            =   6000
      TabIndex        =   14
      Top             =   2610
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   2990
      BackColor       =   4801853
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
   Begin VB.Frame Frame1 
      BackColor       =   &H0049453D&
      Caption         =   "Forma de Pago"
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
      Height          =   1485
      Left            =   6000
      TabIndex        =   24
      Top             =   810
      Width           =   4275
      Begin VB.OptionButton chkNada 
         BackColor       =   &H0049453D&
         Caption         =   "Al contado"
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   210
         TabIndex        =   13
         Top             =   240
         Value           =   -1  'True
         Width           =   1840
      End
      Begin VB.OptionButton chkFin 
         BackColor       =   &H0049453D&
         Caption         =   "A través de Financiera"
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   210
         TabIndex        =   26
         Top             =   870
         Width           =   2000
      End
      Begin VB.OptionButton chkCC 
         BackColor       =   &H0049453D&
         Caption         =   "Cuenta Corriente"
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   210
         TabIndex        =   25
         Top             =   570
         Width           =   1840
      End
   End
   Begin VB.TextBox txtIVa 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   2100
      TabIndex        =   3
      Top             =   2640
      Width           =   3700
   End
   Begin VB.TextBox txtCUIT 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   2100
      TabIndex        =   2
      Top             =   2190
      Width           =   3700
   End
   Begin VB.TextBox txtDireccion 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   630
      Left            =   2100
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1500
      Width           =   3700
   End
   Begin VB.TextBox txtDetalle 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   5940
      Width           =   3700
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2100
      TabIndex        =   0
      Top             =   930
      Width           =   3700
   End
   Begin tbrControles.MouTextBox txtLimite 
      Height          =   375
      Left            =   7560
      TabIndex        =   16
      Top             =   4500
      Width           =   1425
      _ExtentX        =   2514
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
   Begin MSComCtl2.DTPicker DTNac 
      Height          =   375
      Left            =   4500
      TabIndex        =   7
      Top             =   3990
      Width           =   1335
      _ExtentX        =   2355
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
      Format          =   61407233
      CurrentDate     =   39259
   End
   Begin tbrFaroButton.fBoton Command2 
      Height          =   435
      Left            =   9510
      TabIndex        =   39
      Top             =   7400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdLimpiar 
      Height          =   435
      Left            =   4500
      TabIndex        =   40
      Top             =   7400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Limpiar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdOK 
      Height          =   435
      Left            =   2280
      TabIndex        =   41
      Top             =   7400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Grabar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nacimiento"
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
      Left            =   3330
      TabIndex        =   38
      Top             =   4080
      Width           =   1125
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Límite Crédito"
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
      Left            =   6060
      TabIndex        =   37
      Top             =   4560
      Width           =   1365
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ing.Mensual"
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
      Left            =   3390
      TabIndex        =   36
      Top             =   5400
      Width           =   1365
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Antigüedad"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   150
      TabIndex        =   35
      Top             =   5430
      Width           =   1800
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Detalle Trabajo"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   150
      TabIndex        =   34
      Top             =   4530
      Width           =   1800
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono Trabajo"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   150
      TabIndex        =   33
      Top             =   4950
      Width           =   1800
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   150
      TabIndex        =   32
      Top             =   3210
      Width           =   1800
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mail"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   150
      TabIndex        =   31
      Top             =   3630
      Width           =   1800
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione Financiera"
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
      Left            =   6060
      TabIndex        =   30
      Top             =   2370
      Width           =   3795
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Condición IVA"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   150
      TabIndex        =   23
      Top             =   2700
      Width           =   1800
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CUIT / CUIL / DNI"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   150
      TabIndex        =   22
      Top             =   2280
      Width           =   1800
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Direccion"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   150
      TabIndex        =   21
      Top             =   1530
      Width           =   1800
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cliesdfsdgsfd"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   2010
      TabIndex        =   20
      Top             =   150
      Width           =   6015
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Detalle"
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   150
      TabIndex        =   19
      Top             =   6030
      Width           =   1800
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      ForeColor       =   &H00E0E0E0&
      Height          =   345
      Left            =   150
      TabIndex        =   18
      Top             =   960
      Width           =   1800
   End
End
Attribute VB_Name = "frmClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSBUSCAR As New ADODB.Recordset
Dim EsFinan As Boolean
Dim IDC As Long 'codigo del cliente que estoy trabajando
'si es -1 es nuevo

Private Sub chkCC_Click()
    VerDaTos
End Sub

Private Sub chkFin_Click()
    VerDaTos
End Sub

Private Sub chkNada_Click()
    VerDaTos
End Sub

Private Sub chkTrabaja_Click()
    VerDaTos
End Sub

Private Sub cmdLimpiar_Click()
    Dim XX As Control
    
    For Each XX In frmClientes
        If TypeOf XX Is TextBox Then
            If XX <> txtNombre Then XX = ""
        End If
        
        If TypeOf XX Is tbrControles.MouTextBox Then
            XX = FormatCurrency(0)
        End If
    Next
    
    chkTrabaja.Value = 0
End Sub

Private Sub cmdOK_Click()
    If txtNombre = "" Then MsgBox "No cargó ningún nombre de cliente": Exit Sub
    txtIngreso = FormatCurrency(ValidarNumeros(txtIngreso), , , , vbFalse)
    
    Dim Cl As New clsCliente
    If IDC <> -1 And IDC <> -20 Then Cl.AbrirDatos IDC 'solamente para que mcodigo=idc
    Cl.Nombre = txtNombre
    Cl.Detalle = txtDetalle
    Cl.Direccion = txtDireccion
    Cl.CUIT = txtCUIT
    Cl.Iva = txtIVA
    Cl.Telefono = txtTelefono
    Cl.Mail = txtMail
    Cl.TieneTrabajo = chkTrabaja.Value
    Cl.Nacimiento = DTNac
    
    Dim ESnue As Boolean
    
    If IDC = -1 Or IDC = -20 Then
        ESnue = True
    Else
        ESnue = False
    End If

    If Cl.Grabar(ESnue, EsFinan) = 1 Then
        MsgBox "Ya existe una Cuenta con ese nombre", vbExclamation, "Atención"
        Exit Sub
    End If
    
    Set Cl = Nothing
        
    'si tiene trabajo grabar en configuracion los datos asi -----------------------
    '30|0|DatosTrabajoCliente|0|IDc_DetT_Ant_Fono_Ingr|0 -------------------------
    Dim IdCF As Long
    
    'si es nuevo cambio IDC al nuevo idc
    If IDC = -1 Then IDC = DB.GetValInRS("Clientes", "ID", _
            "Nombre = '" + txtNombre + "'", False)
    
    If chkTrabaja And IDC > 0 Then
        'veo si existe
        IdCF = CFG.ExistePropiedad("DTC " + CStr(IDC))
        
        If IdCF = 0 Then 'agrego
            CFG.AgregarNodo 30, "DTC " + CStr(IDC), "", CStr(IDC) + "_" + _
                txtTrabajo + "_" + txtAntiguedad + "_" + txtTelefono2 + _
                "_" + CStr(CSng(txtIngreso)), 0
        Else 'modifico
            CFG.ModificarNodo IdCF, 30, "DTC " + CStr(IDC), , CStr(IDC) + "_" + _
                txtTrabajo + "_" + txtAntiguedad + "_" + txtTelefono2 + _
                "_" + CStr(CSng(txtIngreso))
        End If
    End If
    '------------------------------------------------------------------------------
    
    'si condicion de pago no es contado grabo como es -----------------------------
    ' 40|0|FormadePago|0|Forma_Dias o NroC o Financiera_Lim|0
    Dim StFDP As String
    
    If IDC <> -2 And chkNada.Value = False Then 'que no sea MiEmpresa todo lo demas se pueda conf
        'veo si existe
        IdCF = CFG.ExistePropiedad("FDP " + CStr(IDC))
        
        If chkCC Then
            StFDP = "CC" 'asi empieza
            If chkUnico Then 'en cuota unica pongo solo los dias
                StFDP = StFDP + "_" + CStr(txtDias)
            Else 'es en cuotas
                StFDP = StFDP + "_CUO"
            End If
        Else
            StFDP = "FN_" + tbrBuscadorF.GetLstSel
        End If
        
        StFDP = StFDP + "_" + CStr(CSng(txtLimite))
        
        If IdCF = 0 Then 'agrego
            CFG.AgregarNodo 40, "FDP " + CStr(IDC), "", StFDP, 0
        Else 'modifico
            CFG.ModificarNodo IdCF, 40, "FDP " + CStr(IDC), , StFDP
        End If
    End If
    
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub CargarDatosPr(IdCl As Long)
    Dim Cl As New clsCliente
    Cl.AbrirDatos IdCl
    txtNombre = Cl.Nombre
    If txtNombre = "Otros" Then
        txtNombre.Enabled = False
    End If
    
    txtDetalle = Cl.Detalle
    txtDireccion = Cl.Direccion
    txtCUIT = Cl.CUIT
    txtIVA = Cl.Iva
    txtTelefono = Cl.Telefono
    txtMail = Cl.Mail
    
    If IdCl > -20 And IdCl <> -2 Then
        chkTrabaja.Value = Cl.TieneTrabajo
        If Year(Cl.Nacimiento) > 1900 Then DTNac = Cl.Nacimiento
    End If
    
    Set Cl = Nothing
End Sub

Public Sub AbrirDatos(IdcLiente As Long)
    Dim IdCF As Long, SP() As String
    
    IDC = IdcLiente
    EsFinan = False
    
    If IDC <> -1 And IDC > -20 Then
        
        frmClientes.Caption = "Modificar Detalle Cliente"
        CargarDatosPr IDC
        
        'veo si tiene configuracion DE TRABAJO
        ' es asi '30|0|DatosTrabajoCliente|0|IDc_DetT_Ant_Fono_Ingr|0
        IdCF = CFG.ExistePropiedad("DTC " + CStr(IDC))
        
        If IdCF <> 0 Then
            SP = Split(CFG.GetInfo(IdCF, 4), "_")
            txtTrabajo = SP(1)
            txtAntiguedad = SP(2)
            txtTelefono2 = SP(3)
            txtIngreso = FormatCurrency(NoNuloN(SP(4)), , , , vbFalse)
        End If
        
        If IDC = -2 Then 'mis datos de la empresa ---------------------------------
            frmClientes.Caption = "Datos de mi empresa"
            Me.BackColor = &HA57A47
            Me.Height = 6350
            txtDetalle.Top = 4100
            Label2.Top = 4300
            cmdOK.Top = 5700
            cmdLimpiar.Top = 5700
            command2.Top = 5700
            command2.Left = 5000
            cmdOK.Left = 1800
            cmdLimpiar.Left = 3300
            Me.Width = 7000
            Me.Height = 7500
            Label3.ForeColor = vbBlack
            Label3.Left = 800
            
            Frame1.Visible = False
            Frame2.Visible = False
            tbrBuscadorF.Visible = False
            Label5.Visible = False
            Label9.Visible = False
            Label10.Visible = False
            Label11.Visible = False
            Label14.Visible = False
            Label15.Visible = False
            Label16.Visible = False
            
            chkTrabaja.Visible = False
            txtTrabajo.Visible = False
            txtTelefono2.Visible = False
            txtIngreso.Visible = False
            txtAntiguedad.Visible = False
            txtLimite.Visible = False
            DTNac.Visible = False
        End If '--------------------------------------------------------------------
        Set Cl = Nothing
    Else
        frmClientes.Caption = "Agregar Cliente"
    End If
    
    FormaDePago IDC
    
    If IDC <= -20 Then 'FINANCIERA !!! (si es nuevo es -20 justo) -------------------
        EsFinan = True
        If IDC = -20 Then 'ES NUEVA FINANCIERA
            frmClientes.Caption = "Agregar Financiera"
        Else
            frmClientes.Caption = "Modificar Datos de Financiera"
            tbrBuscadorF.Text = DB.GetValInRS("Clientes", "Nombre", "ID=" + CStr(IDC))
        End If
        
        Me.BackColor = &H968D63
        Me.Height = 7000
        txtDetalle.Top = 4100
        Label2.Top = 4300
        Label3.ForeColor = vbBlack
        Label3.Left = 50
        cmdOK.Top = 5600
        cmdLimpiar.Top = 5600
        command2.Top = 5600
        command2.Left = 5000
        cmdOK.Left = 1800
        cmdLimpiar.Left = 3300
        'Me.Width = 6900
        Label5.Top = 200
        tbrBuscadorF.Visible = True
        tbrBuscadorF.Top = 400
        tbrBuscadorF.Height = 2000
        tbrBuscadorF.Width = 3500
        tbrBuscadorF.FontT.Size = 10
        tbrBuscadorF.FontT.Bold = True
        tbrBuscadorF.BackColor = Me.BackColor
                
        Frame1.Top = 2700
        Label9.Visible = False
        Label10.Visible = False
        Label11.Visible = False
        Label14.Visible = False
        Label16.Visible = False

        chkFin.Visible = False
        chkTrabaja.Visible = False
        txtTrabajo.Visible = False
        txtTelefono2.Visible = False
        txtIngreso.Visible = False
        txtAntiguedad.Visible = False
        DTNac.Visible = False
    End If '------------------------------------------------------------------------
    
    Label3 = Me.Caption
    VerDaTos
    Me.Show 1
End Sub

Private Sub FormaDePago(IdCl As Long)
    'veo si tiene configuracion DE FORMA DE PAGO
    ' 40|0|FormadePago|0|Forma_Dias o NroC o Financiera_Lim|0
    Dim IdCF As Long
    
    IdCF = CFG.ExistePropiedad("FDP " + CStr(IdCl))
    
    If IdCF <> 0 Then
        SP = Split(CFG.GetInfo(IdCF, 4), "_")
        Select Case SP(0)
            Case "CC"
                chkCC.Value = True
                If SP(1) = "CUO" Then
                    chkCuotas.Value = True
                Else
                    chkUnico.Value = True
                    txtDias = CStr(NoNuloN(SP(1)))
                End If
                
            Case "FN"
                chkFin.Value = True
                tbrBuscadorF.Text = SP(1)
                
            Case Else
                chkNada.Value = True
        End Select
        
        txtLimite = FormatCurrency(NoNuloN(SP(2)), , , , vbFalse)
    End If
End Sub

Private Sub VerDaTos()
    'TRABAJO ----------------------------------------------------------------------
    If EsFinan = False Then
        If chkTrabaja Then
            txtTrabajo.Enabled = True
            txtAntiguedad.Enabled = True
            txtTelefono2.Enabled = True
            txtIngreso.Enabled = True
        Else
            txtTrabajo.Enabled = False
            txtAntiguedad.Enabled = False
            txtTelefono2.Enabled = False
            txtIngreso.Enabled = False
        End If
    End If
    '-----------------------------------------------------------------------------
    
    'COMO FINANCIA ---------------------------------------------------------------
    If chkNada Then
        If EsFinan = False Then
            tbrBuscadorF.Visible = False
            Label5.Visible = False
        End If
        Frame2.Visible = False
        Frame1.Width = 3000
    Else
        If chkCC Then
            Frame2.Visible = True
            If EsFinan = False Then
                tbrBuscadorF.Visible = False
                Label5.Visible = False
                Frame2.Left = 6000
            Else
                Frame2.Left = 7900
                Frame2.Top = Frame1.Top
                Frame2.Width = 2900
                Frame1.Width = 2000
            End If
        Else 'es a traves de financiera
            If EsFinan = False Then
                tbrBuscadorF.Visible = True
                Label5.Visible = True
                Frame2.Visible = False
                tbrBuscadorF.Width = 3500
                tbrBuscadorF.Recargar
            End If
        End If
    End If
    '-----------------------------------------------------------------------------
    tbrBuscadorF_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdOK_Click
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    tbrBuscadorF.Contrasena = Contrasena
    tbrBuscadorF.ArchivoMDB = ArchivoMDBPrincipal
    tbrBuscadorF.SqlSinLike = "SELECT Id,Nombre FROM Clientes WHERE ID <=-20"
    tbrBuscadorF.CampoEnQueBuscar = "Nombre/b,ID/n"
    tbrBuscadorF.ColumnasSepPorComasyParentesis = "Nombre(3100)/ID(0)"
    
    txtIngreso = FormatCurrency(0)
    txtLimite = FormatCurrency(0)
    txtDias = 0
    DTNac = Date - 13000
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    tbrBuscadorF.CN_CLOSE
End Sub

Private Sub tbrBuscadorF_Change()
    tbrBuscadorF_Click
    
    If tbrBuscadorF.GetLstSel = "" And EsFinan Then
        txtNombre = tbrBuscadorF.Text
    End If
End Sub

Private Sub tbrBuscadorF_Click()
    If tbrBuscadorF.GetLstSel = "" And EsFinan Then
        IDC = -20
        Label3 = "Agregar Financiera"
        frmClientes.Caption = Label3
        cmdLimpiar_Click
        Exit Sub
    End If
    
    If EsFinan Then
        IDC = tbrBuscadorF.GetLstSel(1)
        CargarDatosPr IDC
        FormaDePago IDC
        Label3 = "Modificar Datos Financiera"
        frmClientes.Caption = Label3
    End If
End Sub

Private Sub txtDias_Change()
    txtDias = ValidarNumeros(txtDias)
End Sub

Private Sub txtDias_GotFocus()
    PintarTxt txtDias
End Sub

Private Sub txtIngreso_GotFocus()
    PintarTxt txtIngreso
End Sub

Private Sub txtIngreso_LostFocus()
    txtIngreso = FormatCurrency(ValidarNumeros(txtIngreso), , , , vbFalse)
End Sub

Private Sub txtLimite_GotFocus()
    PintarTxt txtLimite
End Sub

Private Sub txtLimite_LostFocus()
    txtLimite = FormatCurrency(ValidarNumeros(txtLimite), , , , vbFalse)
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
