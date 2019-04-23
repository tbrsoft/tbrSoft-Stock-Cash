VERSION 5.00
Object = "{A7FBD38D-2930-49E3-B60C-9E0202D84549}#15.0#0"; "tbrControles.ocx"
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmComentarios 
   BackColor       =   &H00544B45&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comentarios Clientes"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9930
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmComentarios.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   9930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin tbrFaroButton.fBoton Command1 
      Height          =   465
      Left            =   5760
      TabIndex        =   7
      Top             =   7000
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdBorrarTCom 
      Height          =   465
      Left            =   3690
      TabIndex        =   6
      Top             =   7000
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Borrar Todo"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdBorrarCom 
      Height          =   465
      Left            =   1590
      TabIndex        =   5
      Top             =   7000
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Borrar Selección"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdNuevoCom 
      Height          =   405
      Left            =   4800
      TabIndex        =   4
      Top             =   2460
      Width           =   2000
      _ExtentX        =   3519
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Nuevo Comentario"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrControles.tbrBuscador tbrBuscadorC 
      Height          =   2205
      Left            =   630
      TabIndex        =   0
      Top             =   1110
      Width           =   4000
      _ExtentX        =   7064
      _ExtentY        =   3889
      BackColor       =   5524293
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
   Begin MSDataGridLib.DataGrid DGCom 
      Height          =   3225
      Left            =   570
      TabIndex        =   1
      Top             =   3510
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   5689
      _Version        =   393216
      BackColor       =   16777215
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione Cliente"
      ForeColor       =   &H00E0E0E0&
      Height          =   345
      Left            =   600
      TabIndex        =   3
      Top             =   810
      Width           =   2025
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Comentarios Clientes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   435
      Left            =   2880
      TabIndex        =   2
      Top             =   210
      Width           =   4335
   End
End
Attribute VB_Name = "frmComentarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rsCOM As New ADODB.Recordset

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
    If MsgBox("¿Está seguro de borrar todos los comentarios de TODOS los clientes?", _
        vbExclamation + vbOKCancel, "Borrar Comentarios") = vbCancel Then Exit Sub
        
    DB.EXECUTE "DELETE FROM ComentariosClientes"
    
    ActualizaR
    AcomodarDG
End Sub

Private Sub cmdNuevoCom_Click()
    Dim IDC As Long, UUs As Long, nBUS As String, Cliente As String
    Dim Coment As String
    
    If tbrBuscadorC.GetLstSel = "" Then
        MsgBox "Debe elegir un cliente para hacer un comentario", vbInformation, "Atención"
        Exit Sub
    End If
    
    Cliente = tbrBuscadorC.GetLstSel
    Coment = InputBox("Detalle el comentario sobre " + UCase(Cliente), _
        "Comentarios")
    
    If Coment = "" Then Exit Sub
    
    'grabo nomas
    IDC = DB.GetValInRS("Clientes", "ID", "Nombre = '" + Cliente + "'", False)
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

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    FormatearMouTextBox frmComentarios

    tbrBuscadorC.Contrasena = Contrasena
    tbrBuscadorC.ArchivoMDB = ArchivoMDBPrincipal
    tbrBuscadorC.SqlSinLike = "SELECT id,Nombre FROM Clientes WHERE ID >=0"
    tbrBuscadorC.OrderBy = "ORDER BY Nombre"
    tbrBuscadorC.CampoEnQueBuscar = "Nombre"
    tbrbuscadorC_Change
    tbrBuscadorC.ColumnasSepPorComasyParentesis = "Nombre(3700)"
    tbrBuscadorC.Recargar
    
    ActualizaR
    AcomodarDG
End Sub

Public Sub AbrirDatos(NombreCliente As String)
    tbrBuscadorC.Text = NombreCliente
    tbrBuscadorC.Text = Right(tbrBuscadorC.Text, Len(tbrBuscadorC.Text) - 1)
    Me.Show 1
End Sub

Private Sub AcomodarDG()
    DGCom.Columns("ID").Width = 0
    DGCom.Columns("Fecha").Width = 1100
    DGCom.Columns("Hora").Width = 500
    DGCom.Columns("Nombre").Width = 1600
    DGCom.Columns("Comentario").Width = 3500
    DGCom.Columns("Usuario").Width = 1200
    
    DGCom.Columns("Fecha").Alignment = dbgCenter
    DGCom.Columns("Hora").Alignment = dbgCenter
    
End Sub

Private Sub ActualizaR()
    Dim S As String, S2 As String
    
    If tbrBuscadorC.Text = "" Then
        Label1 = "Comentarios Generales"
        S2 = ""
    Else
        Label1 = "Comentarios de " + UCase(tbrBuscadorC.GetLstSel)
        S2 = "WHERE Nombre = '" + tbrBuscadorC.GetLstSel + "'"
    End If
    
    Set DGCom.DataSource = Nothing
    If rsCOM.State = adStateOpen Then rsCOM.Close
    rsCOM.CursorLocation = adUseClient
    
    S = "SELECT ComentariosClientes.ID, ComentariosClientes.Fecha, " + _
        "ComentariosClientes.Hora, Clientes.Nombre, " + _
        "ComentariosClientes.Comentario, ComentariosClientes.Usuario " + _
        "FROM Clientes INNER JOIN ComentariosClientes ON Clientes.ID = " + _
        "ComentariosClientes.IdCliente " + S2 + " ORDER BY comentariosClientes.ID DESC"
    
    rsCOM.Open S, DB.CN, adOpenStatic, adLockReadOnly
    
    Set DGCom.DataSource = rsCOM
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    tbrBuscadorC.CN_CLOSE
    
    Set DGCom.DataSource = Nothing
    rsCOM.Close
    Set rsCOM = Nothing
End Sub

Private Sub tbrbuscadorC_Change()
    ActualizaR
    AcomodarDG
End Sub

Private Sub tbrbuscadorc_Click()
    ActualizaR
    AcomodarDG
End Sub
