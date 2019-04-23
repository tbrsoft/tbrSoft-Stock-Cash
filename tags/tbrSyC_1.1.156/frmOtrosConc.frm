VERSION 5.00
Object = "{A7FBD38D-2930-49E3-B60C-9E0202D84549}#13.0#0"; "tbrControles.ocx"
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmOtrosConc 
   BackColor       =   &H004E4E4E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Otros Conceptos"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7020
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOtrosConc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin tbrFaroButton.fBoton cmdOK 
      Height          =   405
      Left            =   3060
      TabIndex        =   25
      Top             =   4770
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "aceptar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdConfigurar 
      Height          =   405
      Left            =   960
      TabIndex        =   24
      Top             =   4770
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "configurar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrControles.MouTextBox txtConcepto 
      Height          =   405
      Index           =   0
      Left            =   3800
      TabIndex        =   1
      Top             =   390
      Visible         =   0   'False
      Width           =   1485
      _ExtentX        =   2619
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
   End
   Begin tbrControles.MouTextBox txtConcepto 
      Height          =   405
      Index           =   1
      Left            =   3800
      TabIndex        =   3
      Top             =   870
      Visible         =   0   'False
      Width           =   1485
      _ExtentX        =   2619
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
   End
   Begin tbrControles.MouTextBox txtConcepto 
      Height          =   405
      Index           =   2
      Left            =   3800
      TabIndex        =   5
      Top             =   1380
      Visible         =   0   'False
      Width           =   1485
      _ExtentX        =   2619
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
   End
   Begin tbrControles.MouTextBox txtConcepto 
      Height          =   405
      Index           =   3
      Left            =   3800
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   1485
      _ExtentX        =   2619
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
   End
   Begin tbrControles.MouTextBox txtConcepto 
      Height          =   405
      Index           =   4
      Left            =   3800
      TabIndex        =   9
      Top             =   2400
      Visible         =   0   'False
      Width           =   1485
      _ExtentX        =   2619
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
   End
   Begin tbrControles.MouTextBox txtConcepto 
      Height          =   405
      Index           =   5
      Left            =   3800
      TabIndex        =   11
      Top             =   2910
      Visible         =   0   'False
      Width           =   1485
      _ExtentX        =   2619
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
   End
   Begin tbrControles.MouTextBox txtConcepto 
      Height          =   405
      Index           =   6
      Left            =   3800
      TabIndex        =   13
      Top             =   3390
      Visible         =   0   'False
      Width           =   1485
      _ExtentX        =   2619
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
   End
   Begin tbrControles.MouTextBox txtConcepto 
      Height          =   405
      Index           =   7
      Left            =   3800
      TabIndex        =   15
      Top             =   3900
      Visible         =   0   'False
      Width           =   1485
      _ExtentX        =   2619
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
   End
   Begin tbrFaroButton.fBoton cmdSalir 
      Height          =   405
      Left            =   5100
      TabIndex        =   26
      Top             =   4770
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.Label lblCta 
      Caption         =   "Label1"
      Height          =   345
      Index           =   7
      Left            =   5580
      TabIndex        =   23
      Top             =   3960
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label lblCta 
      Caption         =   "Label1"
      Height          =   345
      Index           =   6
      Left            =   5580
      TabIndex        =   22
      Top             =   3450
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label lblCta 
      Caption         =   "Label1"
      Height          =   345
      Index           =   5
      Left            =   5580
      TabIndex        =   21
      Top             =   2940
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label lblCta 
      Caption         =   "Label1"
      Height          =   345
      Index           =   4
      Left            =   5580
      TabIndex        =   20
      Top             =   2460
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label lblCta 
      Caption         =   "Label1"
      Height          =   345
      Index           =   3
      Left            =   5580
      TabIndex        =   19
      Top             =   1950
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label lblCta 
      Caption         =   "Label1"
      Height          =   345
      Index           =   2
      Left            =   5580
      TabIndex        =   18
      Top             =   1410
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label lblCta 
      Caption         =   "Label1"
      Height          =   345
      Index           =   1
      Left            =   5580
      TabIndex        =   17
      Top             =   930
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label lblCta 
      Caption         =   "Label1"
      Height          =   345
      Index           =   0
      Left            =   5580
      TabIndex        =   16
      Top             =   420
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label lblConcepto 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   7
      Left            =   350
      TabIndex        =   14
      Top             =   4020
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label lblConcepto 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   6
      Left            =   350
      TabIndex        =   12
      Top             =   3510
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label lblConcepto 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   5
      Left            =   350
      TabIndex        =   10
      Top             =   3030
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label lblConcepto 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   4
      Left            =   350
      TabIndex        =   8
      Top             =   2520
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label lblConcepto 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   3
      Left            =   350
      TabIndex        =   6
      Top             =   2040
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label lblConcepto 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   2
      Left            =   350
      TabIndex        =   4
      Top             =   1500
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label lblConcepto 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   1
      Left            =   255
      TabIndex        =   2
      Top             =   990
      Visible         =   0   'False
      Width           =   3225
   End
   Begin VB.Label lblConcepto 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   0
      Left            =   300
      TabIndex        =   0
      Top             =   510
      Visible         =   0   'False
      Width           =   3195
   End
End
Attribute VB_Name = "frmOtrosConc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IsVenta As Boolean
Dim StrConc As String

Private Sub cmdConfigurar_Click()
    frmConfigFacVta.AbrirDatos
    
    VerQueConceptos
End Sub

Private Sub cmdOK_Click()
    Dim LvCual As ListView, I As Long, J As Long
    
    If IsVenta Then
        Set LvCual = frmVENTAS.lvConceptos
    Else
        Set LvCual = frmCompras.lvConc2
    End If
    
    LvCual.ListItems.Clear
    For I = 0 To 7
        J = LvCual.ListItems.Count + 1
        If lblConcepto(I).Visible = True Then
            If Not IsNumeric(txtConcepto(I)) Then txtConcepto(I) = FormatCurrency(0)
            If EsCero(CSng(txtConcepto(I))) = False Then
                LvCual.ListItems.Add J
                
                LvCual.ListItems(J).Text = lblCta(I)
                LvCual.ListItems(J).SubItems(1) = lblConcepto(I)
                LvCual.ListItems(J).SubItems(2) = FormatCurrency(CSng(txtConcepto(I)), , , vbFalse)
                LvCual.ListItems(J).SubItems(3) = CStr(I + 1)
            End If
        Else
            Exit For
        End If
    Next I
    
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Public Sub AbrirDatos(Optional EsVenta As Boolean = True)
    IsVenta = EsVenta
    
    Me.Show 1
End Sub

Private Sub Form_Load()
    If IsVenta Then
        StrConc = "Concepto "
    Else
        StrConc = "ConceptoCpra "
    End If
    
    VerQueConceptos
End Sub

Private Sub VerQueConceptos()
    Dim I As Long, IdCF As Long, SP() As String
    
    'primero los pongo todo invisible
    For I = 0 To 7
        lblConcepto(I).Visible = False
        txtConcepto(I).Visible = False
    Next I

    For I = 1 To 8
        IdCF = CFG.ExistePropiedad(StrConc + CStr(I))
    
        If IdCF = 0 Then
            Exit For
        Else
            SP = Split(CFG.GetInfo(IdCF, 4), "_")
            If PC.GetNameCuenta(CLng(SP(0))) <> "NO EXISTE" Then
                lblConcepto(I - 1).Visible = True
                lblConcepto(I - 1) = UCase(SP(1))
                txtConcepto(I - 1).Visible = True
                txtConcepto(I - 1) = FormatCurrency(0)
                lblCta(I - 1) = SP(0)
            Else
                CFG.EliminarNodo (IdCF)
            End If
        End If
    Next I
    
End Sub

Private Sub txtConcepto_GotFocus(Index As Integer)
    PintarTxt txtConcepto(Index)
End Sub

Private Sub txtConcepto_LostFocus(Index As Integer)
    txtConcepto(Index) = FormatCurrency(ValidarNumeros(txtConcepto(Index)), , , , vbFalse)
End Sub
