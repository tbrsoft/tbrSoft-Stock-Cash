VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmParticipacion 
   BackColor       =   &H00544B45&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Participación Socios"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6930
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmParticipacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin tbrFaroButton.fBoton cmdSalir 
      Height          =   435
      Left            =   5190
      TabIndex        =   2
      Top             =   4890
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin MSComctlLib.ListView lvParti 
      Height          =   2355
      Left            =   390
      TabIndex        =   0
      Top             =   1080
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   4154
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Socio"
         Object.Width           =   4498
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Participación"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   " % "
         Object.Width           =   1940
      EndProperty
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Participaciones"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   390
      TabIndex        =   1
      Top             =   690
      Visible         =   0   'False
      Width           =   2025
   End
End
Attribute VB_Name = "frmParticipacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    ReCargarDatos
End Sub

Private Sub ReCargarDatos()
    Dim IdCuentas() As String, I As Long, Nombre As String, TmP As Single
    
    IdCuentas = PC.GetCuentas(52)
    TmP = 0
    
    lvParti.ListItems.Clear
    
    For I = 1 To UBound(IdCuentas)
        Nombre = PC.GetNameCuenta(CLng(IdCuentas(I)))
        
        lvParti.ListItems.Add I
        
        lvParti.ListItems(I).Text = Nombre
        lvParti.ListItems(I).SubItems(1) = FormatCurrency(-PC.GetSaldo( _
            PC.GetIDCuenta("Aporte Participación " + Nombre)), , , , vbFalse)
        lvParti.ListItems(I).SubItems(2) = FormatPercent(GetParticipacion(Nombre))
        
        TmP = TmP + CSng(lvParti.ListItems(I).SubItems(1))
    Next I
    
    lvParti.ListItems.Add I
    lvParti.ListItems.Add I + 1
    lvParti.ListItems.Add I + 2
    
    lvParti.ListItems(I + 2).Text = "TOTAL"
    lvParti.ListItems(I + 2).SubItems(1) = FormatCurrency(TmP, , , , vbFalse)
    lvParti.ListItems(I + 2).SubItems(2) = FormatPercent(TmP / PC.ABSSumarconSubcuentas(15))
End Sub
