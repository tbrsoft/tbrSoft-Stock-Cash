VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmConfigImprCP 
   BackColor       =   &H004E4E4E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración Impresión Codigo Producto"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7515
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConfigImprCP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin tbrFaroButton.fBoton cmdLimpiar 
      Height          =   435
      Left            =   3330
      TabIndex        =   8
      Top             =   3270
      Width           =   855
      _ExtentX        =   1508
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
      Left            =   3330
      TabIndex        =   9
      Top             =   2640
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "<<<"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdAgregar 
      Height          =   435
      Left            =   3330
      TabIndex        =   6
      Top             =   2040
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   ">>>"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.TextBox txtTitulo 
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
      Left            =   2280
      TabIndex        =   4
      Top             =   870
      Width           =   3700
   End
   Begin VB.CheckBox chkTitulo 
      BackColor       =   &H004E4E4E&
      Caption         =   "Tiene Título"
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
      Left            =   690
      TabIndex        =   3
      Top             =   930
      Width           =   1305
   End
   Begin VB.ListBox lstA 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2790
      Left            =   4560
      TabIndex        =   1
      Top             =   1500
      Width           =   2535
   End
   Begin VB.ListBox lstDE 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2790
      ItemData        =   "frmConfigImprCP.frx":000C
      Left            =   420
      List            =   "frmConfigImprCP.frx":0022
      TabIndex        =   0
      Top             =   1500
      Width           =   2595
   End
   Begin tbrFaroButton.fBoton Command1 
      Height          =   345
      Left            =   2640
      TabIndex        =   7
      Top             =   4950
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   609
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "grabar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdSalir 
      Height          =   345
      Left            =   5220
      TabIndex        =   10
      Top             =   4950
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   609
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(mínimo 3 conceptos)"
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
      Left            =   4560
      TabIndex        =   5
      Top             =   4290
      Width           =   2355
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Configurar Impresión Código Producto"
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
      Height          =   435
      Left            =   60
      TabIndex        =   2
      Top             =   240
      UseMnemonic     =   0   'False
      Width           =   7005
   End
End
Attribute VB_Name = "frmConfigImprCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkTitulo_Click()
    If chkTitulo Then
        txtTitulo.Visible = True
    Else
        txtTitulo = ""
        txtTitulo.Visible = False
    End If
End Sub

Private Sub cmdAgregar_Click()
    If lstDE.ListIndex = -1 Then Exit Sub
    
    lstA.AddItem lstDE
End Sub

Private Sub cmdLimpiar_Click()
    lstA.Clear
End Sub

Private Sub cmdQuitar_Click()
    If lstA.ListIndex = -1 Then Exit Sub
    
    lstA.RemoveItem lstA.ListIndex
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    GrabarConf
End Sub

Private Sub Form_Load()
    Dim TmP As String, SP() As String, R As Long, S As Long
    
    TmP = CFG.GetInfo(9, 4)
    If TmP = "" Then TmP = "ID_Nombre_Precio"
    If InStrRev(TmP, "_") = 0 Then TmP = "ID_Nombre_Precio"
    
    SP = Split(TmP, "_")
    
    lstA.Clear
    
    For R = 0 To UBound(SP)
        'tiene que estar el lstDE
        For S = 0 To lstDE.ListCount - 1
            If SP(R) = lstDE.List(S) Then
                lstA.AddItem SP(R)
                'Exit For
            End If
        Next S
    Next R
    
    If CFG.GetInfo(9, 3) = "" Then
        chkTitulo.Value = 0
        txtTitulo.Visible = False
    Else
        chkTitulo.Value = 1
        txtTitulo.Visible = True
        txtTitulo = CFG.GetInfo(9, 3)
    End If
    
    GrabarConf
End Sub

Private Sub GrabarConf()
    If lstA.ListCount <= 0 Then Exit Sub
    
    Dim R As Long, TmP As String
    
    TmP = ""
    For R = 0 To lstA.ListCount - 1
        TmP = TmP + lstA.List(R)
        If R < lstA.ListCount - 1 Then TmP = TmP + "_"
    Next R
        
    CFG.ModificarNodo 9, , , txtTitulo, TmP
End Sub
