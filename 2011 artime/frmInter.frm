VERSION 5.00
Object = "{A7FBD38D-2930-49E3-B60C-9E0202D84549}#17.0#0"; "tbrControles.ocx"
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmInter 
   BackColor       =   &H00E6DFD5&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5115
   Icon            =   "frmInter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin tbrFaroButton.fBoton cmdOK 
      Height          =   525
      Left            =   870
      TabIndex        =   1
      Top             =   3210
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   926
      fFColor         =   6553600
      fBColor         =   15130581
      fCapt           =   "Aceptar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   16118251
   End
   Begin tbrControles.tbrBuscador tbrBuscador1 
      Height          =   2115
      Left            =   120
      TabIndex        =   0
      Top             =   210
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   3731
      BackColor       =   15130581
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
   Begin tbrFaroButton.fBoton cmdSalir 
      Height          =   525
      Left            =   2700
      TabIndex        =   2
      Top             =   3210
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   926
      fFColor         =   6553600
      fBColor         =   15130581
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   16118251
   End
   Begin VB.Label lblSelec 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   210
      TabIndex        =   4
      Top             =   2430
      Width           =   4605
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   330
      TabIndex        =   3
      Top             =   3210
      Width           =   4425
   End
End
Attribute VB_Name = "frmInter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sIndice As Integer
Dim ActuTbrB As Boolean
Dim AVacio As Boolean

Private Sub cmdOK_Click()
    If tbrBuscador1.GetLstSel = "" And AVacio = False Then
        MsgBox "No ha realizado una selección", vbInformation, "Atención"
        Exit Sub
    End If

    Select Case sIndice
        Case 0, 2, 3
            RespuestaInter(sIndice) = tbrBuscador1.GetLstSel
        Case 1, 4
            RespuestaInter(sIndice) = tbrBuscador1.GetLstSel(1)
    End Select

    Unload Me
End Sub

Private Sub cmdSalir_Click()
    RespuestaInter(sIndice) = "-1" 'que quede registro de que se fue cancelando
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdSalir_Click
    If KeyCode = vbKeyReturn Then cmdOK_Click
End Sub

Public Sub AbrirDatos(Indice As Integer, Optional sLabel As String = "", _
    Optional AceptaVacio As Boolean = False)
    
    ' Por ahora:
    ' "0*IdCliente"
    ' "1*NombreCliente"
    ' "2*NombreProveedor"
    ' "3*IdProducto"
    ' "4*NombreProducto"
    ' El indice es igual que este numero en la ventana esa
    
    sIndice = Indice
    AVacio = AceptaVacio

    If sLabel = "" Then
        Me.Height = 4500
    Else
        Me.Height = 5200
        cmdOK.Top = Label1.Top + 700
        cmdSalir.Top = cmdOK.Top
        
        Label1 = sLabel
    End If

    Me.Show 1
End Sub

Private Sub Form_Load()
    LimResInter 'limpio la matriz de respuesta
    ActuTbrB = False
    
    tbrBuscador1.Contrasena = Contrasena
    tbrBuscador1.ArchivoMDB = ArchivoMDBPrincipal
    
    Select Case sIndice
        Case 2
            Me.Caption = "Seleccione Proveedor"
            tbrBuscador1.SqlSinLike = "SELECT Proveedor FROM Proveedores"
            tbrBuscador1.OrderBy = "ORDER BY Proveedor"
            tbrBuscador1.CampoEnQueBuscar = "Proveedor"
            Me.Caption = "Seleccionar Proveedor"
            tbrBuscador1.ColumnasSepPorComasyParentesis = "Proveedor(3250)"
        Case 0, 1
            Me.Caption = "Seleccione Cliente"
            tbrBuscador1.SqlSinLike = "SELECT id,Nombre FROM Clientes WHERE ID >= 0"
            tbrBuscador1.OrderBy = "ORDER BY Nombre"
            tbrBuscador1.CampoEnQueBuscar = "ID/n,Nombre/b"
            tbrBuscador1.ColumnasSepPorComasyParentesis = "ID(600)/Cliente(3765)"
        Case 3, 4
            Me.Caption = "Seleccione Producto"
            tbrBuscador1.SqlSinLike = "SELECT TOP 50 Productos.ID, " + _
                "TipoProductos.TipoProducto, Productos.nProducto " + _
                "FROM TipoProductos INNER JOIN Productos ON TipoProductos.ID2 = " + _
                "Productos.IdTipoProducto WHERE Productos.ID >=0"
            tbrBuscador1.OrderBy = "ORDER BY ID"
            tbrBuscador1.CampoEnQueBuscar = "Id/n,nproducto/b"
            tbrBuscador1.ColumnasSepPorComasyParentesis = "ID(600)/Producto(2500)"
    End Select
    
    tbrBuscador1.Recargar
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    tbrBuscador1.CN_CLOSE
End Sub

Private Sub VerActu()
    If ActuTbrB = False Then
       ActuTbrB = True
       tbrBuscador1.Recargar
    Else
        ActuTbrB = False
    End If
End Sub

Private Sub tbrBuscador1_Change()
    If sIndice = 2 Then Exit Sub 'proveedor no tiene ID
    
    If IsNumeric(tbrBuscador1.Text) Then
        Select Case sIndice
            Case 0, 1
                'busca el codigo normal
                tbrBuscador1.CampoEnQueBuscar = "ID/b,Nombre"
            Case 3, 4
                tbrBuscador1.CampoEnQueBuscar = "Id/b,nproducto"
        End Select
    Else
        Select Case sIndice
            Case 0, 1
                tbrBuscador1.CampoEnQueBuscar = "ID/n,Nombre/b"
            Case 3, 4
                tbrBuscador1.CampoEnQueBuscar = "Id/n,nproducto/b"
        End Select
    End If

    If tbrBuscador1.Text <> "" Then VerActu
    Seleccion
End Sub

Private Sub Seleccion()
    If tbrBuscador1.GetLstSel = "" Then
        lblSelec = "Sin Selección"
    Else
        If sIndice = 2 Then
            lblSelec = "Selección: " + UCase(tbrBuscador1.GetLstSel)
        Else
            lblSelec = "Selección: " + UCase(tbrBuscador1.GetLstSel(1))
        End If
    End If
    
End Sub

Private Sub tbrBuscador1_Click()
    Seleccion
End Sub
