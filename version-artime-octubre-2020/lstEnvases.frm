VERSION 5.00
Object = "{A7FBD38D-2930-49E3-B60C-9E0202D84549}#15.0#0"; "tbrControles.ocx"
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form lstEnvases 
   BackColor       =   &H00544B45&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Envases y Vales"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8655
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "lstEnvases.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin tbrFaroButton.fBoton cmdAgregarCliente 
      Height          =   435
      Left            =   2800
      TabIndex        =   16
      Top             =   90
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Agregar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton command5 
      Height          =   345
      Left            =   6345
      TabIndex        =   5
      Top             =   4725
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Cero Todo"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton command1 
      Height          =   435
      Left            =   5595
      TabIndex        =   8
      Top             =   6000
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Aceptar"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin tbrControles.MouTextBox Text1 
      Height          =   405
      Left            =   6480
      TabIndex        =   7
      Top             =   5250
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   714
      Alignment       =   2
      BackColor       =   16777215
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin tbrControles.tbrBuscador tbrBuscador1 
      Height          =   2265
      Left            =   480
      TabIndex        =   0
      Top             =   540
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   3995
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
   Begin VB.ListBox lstEnvases 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3840
      IntegralHeight  =   0   'False
      Left            =   5010
      TabIndex        =   1
      Top             =   540
      Width           =   3435
   End
   Begin VB.TextBox txtSoftE 
      Height          =   1545
      Left            =   750
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3210
      Width           =   2925
   End
   Begin VB.ListBox lstCantEnvases 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3840
      IntegralHeight  =   0   'False
      Left            =   4410
      TabIndex        =   12
      Top             =   540
      Width           =   585
   End
   Begin VB.TextBox txtDetalleE 
      Height          =   555
      Left            =   750
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   5280
      Width           =   2955
   End
   Begin tbrFaroButton.fBoton command3 
      Height          =   345
      Left            =   4590
      TabIndex        =   2
      Top             =   4725
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   609
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "+"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton command4 
      Height          =   345
      Left            =   5175
      TabIndex        =   3
      Top             =   4725
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   609
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "-"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton command6 
      Height          =   345
      Left            =   5760
      TabIndex        =   4
      Top             =   4725
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   609
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "+5"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton command2 
      Height          =   435
      Left            =   6795
      TabIndex        =   9
      Top             =   6000
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Buscador Cliente"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   690
      TabIndex        =   11
      Top             =   210
      Width           =   2625
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Detalle SOFT envases"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   750
      TabIndex        =   17
      Top             =   2910
      Width           =   2445
   End
   Begin VB.Label L 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Envases a anotar"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   1
      Left            =   4440
      TabIndex        =   15
      Top             =   270
      Width           =   2085
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Mi detalle envases"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   840
      TabIndex        =   14
      Top             =   5010
      Width           =   2115
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "$ Por envases"
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
      Left            =   4560
      TabIndex        =   13
      Top             =   5310
      Width           =   1665
   End
End
Attribute VB_Name = "lstEnvases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Y() As String ' contiene los envases+"\"can
Dim IDC As Long
Dim Vales As Single 'el valor cargado en $por envases

Private Function HayEnvases() As Boolean
    HayEnvases = False
    Dim I As Long
     'si no hay nada que lo deje false de 1
     
    If lstCantEnvases.ListCount = 0 Then Exit Function
    
    For I = 0 To lstCantEnvases.ListCount
        If lstCantEnvases <> "0" Then HayEnvases = True
    Next I
End Function

Private Sub cmdAgregarCliente_Click()
    frmClientes.AbrirDatos -1
End Sub

Private Sub Command1_Click()
    If tbrBuscador1.GetLstSel = "" Then
        MsgBox "Debe seleccionar un cliente", vbInformation, "Atención"
        Exit Sub
    End If
    
    Dim clscl As New clsCliente
    
    IDC = clscl.GetID(tbrBuscador1.GetLstSel)
    
    Set clscl = Nothing
    
    If Vales < 0 Then MsgBox "Todos los valores deben ser positivos", vbExclamation, _
        "Atención": Exit Sub
        
        
    If HayEnvases = False Then
        MsgBox "No hay nada que anotar!, registre algun movimiento o cierre " + _
            "la ventana ", vbExclamation, "Atención"
        Exit Sub
    End If
    
    If Vales <> 0 And HayEnvases = False Then
        MsgBox "No puede poner vales cuando no hay elegido envases!!", vbExclamation, "Atención"
        Vales = 0
        Text1 = FormatCurrency(Vales, , , , vbFalse)
        Exit Sub
    End If
    
    'ESTA TODO LISTO REGISTRAR LOS MOVIMIENTOS DE PLATA Y DE ENVASES
    'PUEDE NO HABER VARIACION O NO HABER ENVASES
    If HayEnvases Then
         'PARA QUE ASEGURE QUE SI ES OTRO (ASI USA MOu) ponga en detalle quien es nombre
        If tbrBuscador1.GetLstSel = "Otros" Then
            If txtDetalleE = "" Then
                MsgBox "Ha elegido OTROS y hay envases para anotar sin nombre de cliente " + _
                    "debe ingresar su nombre en " + _
                    "Detalle Productos para ser registrado correctamente", vbInformation, "ATENCIÓN!!"
                Exit Sub
            End If
        End If
        
        
        For I = 0 To lstCantEnvases.ListCount - 1
            'XXXX el vale va solo en el primer producto algun dia poner
            'registros hijos
            If lstCantEnvases.List(I) <> "0" Then
                Dim S2 As String
                S2 = "INSERT INTO Movenvases " + _
                    "(ID,Fecha, cantenv, codcliente, Envases,depositoporenvase," + _
                    "detalle) VALUES (" + IdAutonum("MovEnvases") + _
                    ",'" + CStr(Date) + "', '" + _
                    lstCantEnvases.List(I) + "', '" + CStr(IDC) + "', '" + _
                    lstEnvases.List(I) + "','" + CStr(Vales) + "', '" + txtDetalleE + _
                    " - " + txtSoftE + "')"
                DB.EXECUTE S2
                If Vales > 0 Then
                    'registro vales a pagar a caja
                    PC.Asiento "78", CStr(Vales), "93", CStr(Vales), "LibroSubdiario", _
                        "Ingreso por Vales"
                    Vales = 0 'para que no lo ponga en el prox For
                End If
            End If
            
        Next I
    
    End If
        
    Unload Me
End Sub

Private Sub Command2_Click()
    If Variacion <> 0 Or HayEnvases Then
        If MsgBox("¿Esta seguro que desea salir sin cargar estos datos?", vbOKCancel, _
            "Atención") = vbCancel Then Cancel = True
    End If
    Unload Me
End Sub

Private Sub Command3_Click()
    If lstCantEnvases.ListIndex = -1 Then Exit Sub
    lstCantEnvases.List(lstCantEnvases.ListIndex) = CStr(1 + CLng(lstCantEnvases))
    'devolver la ubicacion a donde va
    
    command1.SetFocus 'por si apreta enter
    If lstEnvases.ListIndex > -1 Then lstCantEnvases.ListIndex = lstEnvases.ListIndex
    GenerarDetalleSoftE
    
End Sub

Private Sub Command4_Click()
    command1.SetFocus 'por si apreta enter
    If lstCantEnvases.ListIndex = -1 Then Exit Sub
    If lstCantEnvases = "0" Then Exit Sub
    lstCantEnvases.List(lstCantEnvases.ListIndex) = CStr(CLng(lstCantEnvases) - 1)
    'devolver la ubicacion a donde va
    If lstEnvases.ListIndex > -1 Then lstCantEnvases.ListIndex = lstEnvases.ListIndex
    GenerarDetalleSoftE
End Sub

Private Sub Command5_Click()
    Dim I As Long
    For I = 0 To lstCantEnvases.ListCount - 1
        lstCantEnvases.List(I) = "0"
    Next I
    'devolver la ubicacion a donde va
    If lstEnvases.ListIndex > -1 Then lstCantEnvases.ListIndex = lstEnvases.ListIndex
    GenerarDetalleSoftE
End Sub

Private Sub Command6_Click()
    If lstCantEnvases.ListIndex = -1 Then Exit Sub
    lstCantEnvases.List(lstCantEnvases.ListIndex) = CStr(5 + CLng(lstCantEnvases))
    'devolver la ubicacion a donde va
    If lstEnvases.ListIndex > -1 Then lstCantEnvases.ListIndex = lstEnvases.ListIndex
    GenerarDetalleSoftE
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    FormatearMouTextBox Me
    
    Vales = 0
    Text1 = FormatCurrency(Vales, , , , vbFalse)
    
    tbrBuscador1.Contrasena = Contrasena
    tbrBuscador1.ArchivoMDB = ArchivoMDBPrincipal
    tbrBuscador1.SqlSinLike = "SELECT * FROM Clientes WHERE ID >=0"
    tbrBuscador1.OrderBy = "ORDER BY Nombre"
    tbrBuscador1.CampoEnQueBuscar = "Nombre"
    tbrBuscador1.ColumnasSepPorComasyParentesis = "Nombre(3000)"
    
    CargarCombo lstEnvases, "select envase from envases", "envase"
    
    Dim I As Long
    For I = 0 To lstEnvases.ListCount - 1
        If lstEnvases.List(I) = "No tiene" Then lstEnvases.RemoveItem (I)
    Next
    
    'poner en cero!!!
    For I = 0 To lstEnvases.ListCount - 1
        lstCantEnvases.AddItem "0"
    Next I
    'elegir el primero para que se eligan en las 2 listas!
    lstEnvases.ListIndex = lstEnvases.ListCount - 1

    GenerarDetalleSoftE
    
    tbrBuscador1.Recargar
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    tbrBuscador1.CN_CLOSE
End Sub

Private Sub lstEnvases_Click()
    'existe el caso de cuando se carga primero envases todos y despues los ceros!
    If lstCantEnvases.ListCount = lstEnvases.ListCount Then
        lstCantEnvases.ListIndex = lstEnvases.ListIndex
    End If
End Sub

Private Sub GenerarDetalleSoftE()
    Dim Detalle2 As String
    Detalle2 = ""
    For I = 0 To lstEnvases.ListCount - 1
        If lstCantEnvases.List(I) > 0 Then
            If Detalle2 = "" Then Detalle2 = "Env:" + vbCrLf
            Detalle2 = Detalle2 + lstCantEnvases.List(I) + " " + lstEnvases.List(I) + vbCrLf
        End If
    Next I
    txtSoftE = Detalle2
End Sub


Private Sub Text1_Change()
    If Not IsNumeric(Text1) Then
        Vales = 0
    Else
        Vales = CSng(Text1)
    End If
    
End Sub

Private Sub Text1_GotFocus()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)
End Sub

Private Sub Text1_LostFocus()
    Vales = ValidarNumeros(Text1)
    Text1 = FormatCurrency(Vales, , , , vbFalse)
    
End Sub
