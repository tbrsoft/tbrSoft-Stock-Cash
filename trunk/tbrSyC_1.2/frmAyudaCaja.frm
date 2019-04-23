VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmAyudaCaja 
   BackColor       =   &H00544B45&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ayuda Conteo Caja"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6180
   Icon            =   "frmAyudaCaja.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin tbrFaroButton.fBoton Command2 
      Height          =   450
      Left            =   4230
      TabIndex        =   13
      Top             =   4650
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   794
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton Command1 
      Height          =   645
      Left            =   4050
      TabIndex        =   12
      Top             =   2340
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   1138
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Cerrar Caja con Este Importe"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.TextBox tTOT 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   11
      Left            =   2100
      TabIndex        =   11
      Text            =   "0000.00"
      Top             =   4620
      Width           =   1605
   End
   Begin VB.TextBox tTOT 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   10
      Left            =   2520
      TabIndex        =   35
      Text            =   "0"
      Top             =   4080
      Width           =   1200
   End
   Begin VB.TextBox tTOT 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   9
      Left            =   2520
      TabIndex        =   34
      Text            =   "0"
      Top             =   3690
      Width           =   1200
   End
   Begin VB.TextBox tTOT 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   8
      Left            =   2520
      TabIndex        =   33
      Text            =   "0"
      Top             =   3300
      Width           =   1200
   End
   Begin VB.TextBox tTOT 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   7
      Left            =   2520
      TabIndex        =   32
      Text            =   "0"
      Top             =   2910
      Width           =   1200
   End
   Begin VB.TextBox tTOT 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   6
      Left            =   2520
      TabIndex        =   31
      Text            =   "0"
      Top             =   2520
      Width           =   1200
   End
   Begin VB.TextBox tTOT 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   5
      Left            =   2520
      TabIndex        =   30
      Text            =   "0"
      Top             =   2130
      Width           =   1200
   End
   Begin VB.TextBox tTOT 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   4
      Left            =   2520
      TabIndex        =   29
      Text            =   "0"
      Top             =   1740
      Width           =   1200
   End
   Begin VB.TextBox tTOT 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   3
      Left            =   2520
      TabIndex        =   28
      Text            =   "0"
      Top             =   1350
      Width           =   1200
   End
   Begin VB.TextBox tTOT 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   2
      Left            =   2520
      TabIndex        =   27
      Text            =   "0"
      Top             =   960
      Width           =   1200
   End
   Begin VB.TextBox tTOT 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   1
      Left            =   2520
      TabIndex        =   26
      Text            =   "0"
      Top             =   570
      Width           =   1200
   End
   Begin VB.TextBox tTOT 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   0
      Left            =   2520
      TabIndex        =   25
      Text            =   "0"
      Top             =   180
      Width           =   1200
   End
   Begin VB.TextBox tCANT 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   1740
      TabIndex        =   10
      Text            =   "0"
      Top             =   4080
      Width           =   765
   End
   Begin VB.TextBox tCANT 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   1740
      TabIndex        =   9
      Text            =   "0"
      Top             =   3690
      Width           =   765
   End
   Begin VB.TextBox tCANT 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   1740
      TabIndex        =   8
      Text            =   "0"
      Top             =   3300
      Width           =   765
   End
   Begin VB.TextBox tCANT 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   1740
      TabIndex        =   7
      Text            =   "0"
      Top             =   2910
      Width           =   765
   End
   Begin VB.TextBox tCANT 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   1740
      TabIndex        =   6
      Text            =   "0"
      Top             =   2520
      Width           =   765
   End
   Begin VB.TextBox tCANT 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   1740
      TabIndex        =   5
      Text            =   "0"
      Top             =   2130
      Width           =   765
   End
   Begin VB.TextBox tCANT 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1740
      TabIndex        =   4
      Text            =   "0"
      Top             =   1740
      Width           =   765
   End
   Begin VB.TextBox tCANT 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1740
      TabIndex        =   3
      Text            =   "0"
      Top             =   1350
      Width           =   765
   End
   Begin VB.TextBox tCANT 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1740
      TabIndex        =   2
      Text            =   "0"
      Top             =   960
      Width           =   765
   End
   Begin VB.TextBox tCANT 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1740
      TabIndex        =   1
      Text            =   "0"
      Top             =   570
      Width           =   765
   End
   Begin VB.TextBox tCANT 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1740
      TabIndex        =   0
      Text            =   "0"
      Top             =   180
      Width           =   765
   End
   Begin VB.TextBox tBill 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   960
      TabIndex        =   24
      Text            =   "0,05"
      Top             =   4080
      Width           =   765
   End
   Begin VB.TextBox tBill 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   960
      TabIndex        =   23
      Text            =   "0,10"
      Top             =   3690
      Width           =   765
   End
   Begin VB.TextBox tBill 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   960
      TabIndex        =   22
      Text            =   "0,25"
      Top             =   3300
      Width           =   765
   End
   Begin VB.TextBox tBill 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   960
      TabIndex        =   21
      Text            =   "0,50"
      Top             =   2910
      Width           =   765
   End
   Begin VB.TextBox tBill 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   960
      TabIndex        =   20
      Text            =   "1"
      Top             =   2520
      Width           =   765
   End
   Begin VB.TextBox tBill 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   960
      TabIndex        =   19
      Text            =   "2"
      Top             =   2130
      Width           =   765
   End
   Begin VB.TextBox tBill 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   960
      TabIndex        =   18
      Text            =   "5"
      Top             =   1740
      Width           =   765
   End
   Begin VB.TextBox tBill 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   960
      TabIndex        =   17
      Text            =   "10"
      Top             =   1350
      Width           =   765
   End
   Begin VB.TextBox tBill 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   960
      TabIndex        =   16
      Text            =   "20"
      Top             =   960
      Width           =   765
   End
   Begin VB.TextBox tBill 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   15
      Text            =   "50"
      Top             =   570
      Width           =   765
   End
   Begin VB.TextBox tBill 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   14
      Text            =   "100"
      Top             =   180
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
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
      Height          =   405
      Left            =   600
      TabIndex        =   36
      Top             =   4710
      Width           =   1335
   End
End
Attribute VB_Name = "frmAyudaCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    command1.Enabled = True
    
    If EstaHabilitado(2, "Ingreso Cierre Caja desde Contar Plata") > 0 Then
        frmCierreCaja.AbrirPre tTOT(11).Text
    End If
    
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Dim B As Long
    For B = 0 To 11
        tTOT(B) = FormatCurrency(0)
    Next B
End Sub

Private Sub tCANT_Change(Index As Integer)
    On Error Resume Next
    tTOT(Index) = FormatCurrency(CSng(tBill(Index)) * CSng(tCANT(Index)), , , , vbFalse)
    
    Dim sTOT As Single: sTOT = 0
    Dim A As Long
    For A = 0 To 10
        sTOT = sTOT + CSng(tTOT(A))
    Next A
    tTOT(11) = FormatCurrency(sTOT, , , , vbFalse)
End Sub

Private Sub tCANT_GotFocus(Index As Integer)
    PintarTxt tCANT(Index)
'    If tCANT(Index) = "" Then Exit Sub
'    tCANT(Index).SelStart = 0
'    tCANT(Index).SelLength = Len(tCANT(Index))
End Sub

Private Sub tCANT_LostFocus(Index As Integer)
    tCANT(Index) = ValidarNumeros(tCANT(Index))
End Sub

