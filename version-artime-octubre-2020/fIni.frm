VERSION 5.00
Begin VB.Form fIni 
   BackColor       =   &H00E9F3FE&
   Caption         =   "tbrStock n' Cash"
   ClientHeight    =   7815
   ClientLeft      =   225
   ClientTop       =   225
   ClientWidth     =   11760
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fIni.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7815
   ScaleWidth      =   11760
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pFONDO 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   9180
      ScaleHeight     =   555
      ScaleWidth      =   1305
      TabIndex        =   25
      Top             =   5880
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Frame frmSub 
      BackColor       =   &H00404040&
      Caption         =   "Seleccione la opción"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5955
      Left            =   3240
      TabIndex        =   9
      Top             =   270
      Width           =   2900
      Begin VB.Label lblSub 
         BackColor       =   &H00404040&
         Caption         =   "Sistema"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   12
         Left            =   100
         TabIndex        =   22
         Top             =   5220
         Width           =   2680
      End
      Begin VB.Label lblSub 
         BackColor       =   &H00404040&
         Caption         =   "Sistema"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   11
         Left            =   100
         TabIndex        =   21
         Top             =   4820
         Width           =   2680
      End
      Begin VB.Label lblSub 
         BackColor       =   &H00404040&
         Caption         =   "Sistema"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   10
         Left            =   100
         TabIndex        =   20
         Top             =   4420
         Width           =   2680
      End
      Begin VB.Label lblSub 
         BackColor       =   &H00404040&
         Caption         =   "Sistema"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   9
         Left            =   100
         TabIndex        =   19
         Top             =   4020
         Width           =   2680
      End
      Begin VB.Label lblSub 
         BackColor       =   &H00404040&
         Caption         =   "Sistema"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   8
         Left            =   100
         TabIndex        =   18
         Top             =   3620
         Width           =   2680
      End
      Begin VB.Label lblSub 
         BackColor       =   &H00404040&
         Caption         =   "Sistema"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   7
         Left            =   100
         TabIndex        =   17
         Top             =   3220
         Width           =   2680
      End
      Begin VB.Label lblSub 
         BackColor       =   &H00404040&
         Caption         =   "Sistema"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   6
         Left            =   100
         TabIndex        =   16
         Top             =   2820
         Width           =   2680
      End
      Begin VB.Label lblSub 
         BackColor       =   &H00404040&
         Caption         =   "Sistema"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   5
         Left            =   100
         TabIndex        =   15
         Top             =   2420
         Width           =   2680
      End
      Begin VB.Label lblSub 
         BackColor       =   &H00404040&
         Caption         =   "Sistema"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   4
         Left            =   100
         TabIndex        =   14
         Top             =   2020
         Width           =   2680
      End
      Begin VB.Label lblSub 
         BackColor       =   &H00404040&
         Caption         =   "Sistema"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   3
         Left            =   100
         TabIndex        =   13
         Top             =   1620
         Width           =   2680
      End
      Begin VB.Label lblSub 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Sistema"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   2
         Left            =   100
         TabIndex        =   12
         Top             =   1220
         Width           =   2680
      End
      Begin VB.Label lblSub 
         BackColor       =   &H00404040&
         Caption         =   "Sistema"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   100
         TabIndex        =   11
         Top             =   820
         Width           =   2680
      End
      Begin VB.Label lblSub 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Sistema"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   100
         TabIndex        =   10
         Top             =   420
         Width           =   2680
      End
   End
   Begin VB.Frame frMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Menu tbrSoft Stock Cash"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   7005
      Left            =   270
      TabIndex        =   1
      Top             =   210
      Width           =   2955
      Begin VB.Image imgMenu 
         Height          =   375
         Index           =   7
         Left            =   150
         Picture         =   "fIni.frx":169B2
         Stretch         =   -1  'True
         Top             =   6420
         Width           =   375
      End
      Begin VB.Label lblMenu 
         BackColor       =   &H00000000&
         Caption         =   "Números del Día"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   7
         Left            =   735
         TabIndex        =   24
         Top             =   6450
         Width           =   1995
      End
      Begin VB.Label lblMenu 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Configuraciones"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   6
         Left            =   650
         TabIndex        =   8
         Top             =   3390
         Width           =   2000
      End
      Begin VB.Image imgMenu 
         Height          =   375
         Index           =   6
         Left            =   70
         Picture         =   "fIni.frx":1C194
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   380
      End
      Begin VB.Label lblMenu 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Contabilidad"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   5
         Left            =   650
         TabIndex        =   7
         Top             =   2890
         Width           =   2000
      End
      Begin VB.Image imgMenu 
         Height          =   375
         Index           =   5
         Left            =   70
         Picture         =   "fIni.frx":1C71E
         Stretch         =   -1  'True
         Top             =   2860
         Width           =   380
      End
      Begin VB.Label lblMenu 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Socios y Empleados"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   4
         Left            =   650
         TabIndex        =   6
         Top             =   2390
         Width           =   2000
      End
      Begin VB.Image imgMenu 
         Height          =   375
         Index           =   4
         Left            =   70
         Picture         =   "fIni.frx":1CCA8
         Stretch         =   -1  'True
         Top             =   2360
         Width           =   380
      End
      Begin VB.Label lblMenu 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedores"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   650
         TabIndex        =   5
         Top             =   1890
         Width           =   2000
      End
      Begin VB.Image imgMenu 
         Height          =   375
         Index           =   3
         Left            =   70
         Picture         =   "fIni.frx":1D232
         Stretch         =   -1  'True
         Top             =   1860
         Width           =   380
      End
      Begin VB.Label lblMenu 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Clientes"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   650
         TabIndex        =   4
         Top             =   1390
         Width           =   2000
      End
      Begin VB.Image imgMenu 
         Height          =   375
         Index           =   2
         Left            =   70
         Picture         =   "fIni.frx":1DAFC
         Stretch         =   -1  'True
         Top             =   1360
         Width           =   380
      End
      Begin VB.Label lblMenu 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Productos"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   650
         TabIndex        =   3
         Top             =   890
         Width           =   2000
      End
      Begin VB.Image imgMenu 
         Height          =   375
         Index           =   1
         Left            =   70
         Picture         =   "fIni.frx":1E3C6
         Stretch         =   -1  'True
         Top             =   860
         Width           =   380
      End
      Begin VB.Image imgMenu 
         Height          =   375
         Index           =   0
         Left            =   70
         Picture         =   "fIni.frx":21180
         Stretch         =   -1  'True
         Top             =   360
         Width           =   380
      End
      Begin VB.Label lblMenu 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Sistema"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   650
         TabIndex        =   2
         Top             =   390
         Width           =   2000
      End
   End
   Begin VB.Label lblDetalle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   6300
      TabIndex        =   23
      Top             =   360
      Width           =   45
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Actualizado 10/2020 - #JubilamePorFavor"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3690
      TabIndex        =   0
      Top             =   7530
      Width           =   7875
   End
End
Attribute VB_Name = "fIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim IndiceSub As Integer

Private Sub Form_Activate()
    Terr.AnOtaR "abfm"
    'calculo movimientos!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    'LOS VOY CALCULANDO CADA VEZ QUE HAGO CAMBIO NOMAS CON VARIABLES GLOBALES
    AjustesStock = PC.ABSSumarconSubcuentas(35, False)
    RFyT = PC.ABSSumarconSubcuentas(23, False)
    DifStk = PC.UltVarCuenta(54) - AjustesStock - RFyT
    CpDia = CtoDia + DifStk
    
    VtaDia = PC.ABSSumarconSubcuentas(17, False)
    CtoDia = PC.ABSSumarconSubcuentas(18, False)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Terr.AnOtaR "abcp", KeyCode
    Select Case KeyCode
        Case vbKeyEscape
            frmSub.Width = 0
            InvisiblesSub
            IndiceSub = 50
        
        Case vbKeyF1
            Terr.AnOtaR "abcq"
            MenuPrincipal 2, 0
            
        Case vbKeyF2
            MenuPrincipal 3, 0
            
        Case vbKeyF3
            MenuPrincipal 2, 3
            
        Case vbKeyF4
            If Shift = 4 Then 'Poder seguir usando el Alt+F4 para salir
                Unload Me
                Exit Sub
            End If
            
            If IndiceSub = 3 Then 'anotar a proveedores
                MenuPrincipal 3, 3
            
            Else 'anotar a clientes -> Va en todos los casos que no tenga abierto menu PROVEEDORES!!!
                MenuPrincipal 2, 4
            End If
            
        Case vbKeyF5
            If IndiceSub = 3 Then 'Resumen Cta Proveedores
                MenuPrincipal 3, 4
            
            Else 'Resumen Cta Clientes -> Va en todos los casos que no tenga abierto menu PROVEEDORES!!!
                MenuPrincipal 2, 5
            End If
            
        Case vbKeyF7
            MenuPrincipal 1, 1
            
        Case vbKeyF8
            MenuPrincipal 5, 0
            
        Case vbKeyF And Shift = 2
            If IndiceSub = 3 Then 'Facturas de Compra
                MenuPrincipal 3, 5
            
            Else 'Facturas de Venta -> Va en todos los casos que no tenga abierto menu PROVEEDORES!!!
                MenuPrincipal 2, 6
            End If
            
        Case vbKeyP And Shift = 2
            MenuPrincipal 3, 1
            
        Case vbKeyQ And Shift = 2
            MenuPrincipal 2, 1
            
        Case vbKeyS And Shift = 2
            MenuPrincipal 1, 0
            
            
    End Select
End Sub

Private Sub Form_Load()
    IndiceSub = 50 'solo para que no empiece en 0
    lblDetalle.BackColor = Me.BackColor
   
    '----------CONFIGURACIONES -------------------------------------------
    '(1) TITULO DEL PROGRAMA
    Me.Caption = CFG.GetInfo(1, 4)
    
    frmSub.Width = 0
    
    Dim FSO As New Scripting.FileSystemObject
    pFONDO.AutoSize = True
    If FSO.FileExists(AP + "\fondos\f1.jpg") Then pFONDO.Picture = LoadPicture(AP + "\fondos\f1.jpg")
    
End Sub
            



Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmSub.Width = 0
    InvisiblesSub
    IndiceSub = 50
End Sub

Private Sub InvisiblesSub()
    Dim I As Long
    
    For I = 0 To 12
        lblSub(I) = ""
        lblSub(I).Visible = False
    Next I
End Sub

Private Sub LlenarSub(inD As Integer)
    
    IndiceSub = inD
    frmSub.Caption = "Opciones de " + UCase(lblMenu(inD))

    Select Case inD
        Case 0 'SISTEMA ---------------------------------------------------------------------
            SubSepPorAst "Administrar Accesos*Cambiar Usuario*Cambiar Contraseña*" + _
                "Movimientos Usuarios*Limpieza Base Datos*Exportar Datos Pto Venta*" + _
                "Importar Productos*Exportar Productos"
        Case 1 'PRODUCTOS -------------------------------------------------------------------
            SubSepPorAst "Listado de Productos (Ctrl+S)*Info Productos (F7)*Tipos de Producto*" + _
                "Sucursales*Stock Por Sucursales*Control Stock Mínimo*Historial Mov. Productos*" + _
                "Imprimir Ofertas"
        Case 2 'CLIENTES --------------------------------------------------------------------
            SubSepPorAst "Ventas (F1)*Anular Venta (Ctrl+Q)*Listado de Clientes (Ctrl+C)*" + _
                "Detalle de Clientes (F3)*Anotar a Clientes (F4)*Resumen de Cuenta (F5)*" + _
                "Facturas (Ctrl+F)*Estadísticas de Venta*Envases y Vales (Ctrl+V)*" + _
                "Clientes con Deudas*Comentarios*Vencimientos*Editar Financieras"
        Case 3 'PROVEEDORES -----------------------------------------------------------------
            SubSepPorAst "Compras (F2)*Listado de Proveedores (Ctrl+P)*Pedidos*Anotar a Proveedores (F4)*" + _
                "Resumen de Cuenta (F5)*Facturas (Ctrl+F)*Vencimientos"
        Case 4 'SOCIOS Y EMPLEADOS ----------------------------------------------------------
            SubSepPorAst "Socios*Empleados*Comisión por Venta*Mercadería a Costo"
        Case 5 'CONTABILIDAD ----------------------------------------------------------------
            SubSepPorAst "Cerrar Caja (F8)*Cierres Anteriores*Contar Plata*" + _
                "Cerrar Resultados*Resultados por Tipo de Prod.*Resumen de Cuentas Contables*" + _
                "Balance*IVA Compras*IVA Ventas*Egresos*Asientos Contables"
        Case 6 'CONFIGURACIONES -------------------------------------------------------------
            SubSepPorAst "Datos Empresa*Configuraciones Grales*Conc. Factura Venta*" + _
                "Conc. Factura Compra*Formato Factura*Impresión Código Producto"
    End Select
End Sub

Private Sub SubSepPorAst(Cadena As String)
    Dim Resp() As String, I As Long

    Resp = Split(Cadena, "*")
    
    For I = 0 To 12
        If I <= UBound(Resp) Then
            lblSub(I).Visible = True
            lblSub(I) = Resp(I)
            If Len(lblSub(I)) > 20 Then
                lblSub(I).FontSize = 9
            Else
                lblSub(I).FontSize = 10
            End If
        Else 'el resto los limpio
            lblSub(I) = ""
            lblSub(I).Visible = False
        End If
    Next I
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim I As Integer
    For I = 0 To lblSub.Count - 1
        lblSub(I).BackColor = frmSub.BackColor
    Next I
    
    LlenarDetalle 50, 50
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Local Error GoTo errFIN
    Terr.AnOtaR "abfh"
    'solo si estaba grabando a full lo paro
    If Command = "e1" Then Terr.StopGrabaTodo
    
    LimpiarMovProdViejos 'limpia del  historial de mov de productos s/config
    If Not IsNumeric(CFG.GetInfo(16, 4)) Then
        CFG.ModificarNodo 16, , , , "30"
    End If
    
    
    ACC.LimpiarMov CLng(CFG.GetInfo(16, 4))
    
    ACC.RegEvento ACC.UltUsuarioIngresado, 6, "Salida del sistema"
    ACC.Desconectar
    DB.CN_CLOSE
    PC.CN_CLOSE
    
    
    Set CFG = Nothing
    Set CFGBD = Nothing
    Set ACC = Nothing
    Set DB = Nothing
    Set PC = Nothing
    Set TP = Nothing
    
    Exit Sub
errFIN:
    Terr.AppendLog "fn029", Terr.ErrToTXT(Err)
    Resume Next
End Sub

Private Sub Form_Resize()
    If Me.Height < 500 Or Me.Width < 500 Then Exit Sub 'esta minimizando

    If Me.Height < 8000 Or Me.Width < 11000 Then
        If Me.Height < 8000 Then Me.Height = 8000
        If Me.Width < 11000 Then Me.Width = 11000
        Exit Sub
    End If
    
    AjustarTamano
    
    'poner fondo
    On Local Error Resume Next
    Me.AutoRedraw = True
    Me.PaintPicture pFONDO.Picture, 50, 50, Me.Width - 300, Me.Height - 300, 0, 0, pFONDO.Width, pFONDO.Height
    
End Sub

Private Sub AjustarTamano()
    Label9.Top = Me.Height - 1100
    Label9.Left = Me.Width - Label9.Width - 400
    
    
    frMenu.Height = Label9.Top - 700
    lblMenu(7).Top = frMenu.Height + frMenu.Top - 765
    imgMenu(7).Top = lblMenu(7).Top
End Sub

Private Sub frMenu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmSub.Width = 0
    InvisiblesSub
    IndiceSub = 50
End Sub

Private Sub frMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim I As Integer
    For I = 0 To lblSub.Count - 1
        lblSub(I).BackColor = frmSub.BackColor
    Next I
    
    LlenarDetalle 50, 50
End Sub

Private Sub frmSub_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim I As Integer
    For I = 0 To lblSub.Count - 1
        lblSub(I).BackColor = frmSub.BackColor
    Next I
End Sub

Private Sub imgMenu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgMenu(Index).Left = 100
End Sub

Private Sub imgMenu_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim I As Long
    
    imgMenu(Index).Left = 150
    
    If Index = 7 Then
        frmNumerosDelDia.Show 1
        Exit Sub 'Numeros del día
    End If
    
    If IndiceSub = Index Then
        For I = frmSub.Width * 10 To 0 Step -1
            frmSub.Width = I / 10
        Next I
        
        InvisiblesSub
        IndiceSub = 50
    Else
        LlenarSub Index
        
        For I = 0 To 30000 Step 1
            frmSub.Width = I / 10
        Next I
    End If
End Sub

Private Sub lblMenu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMenu(Index).Left = 600
End Sub

Private Sub lblMenu_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim I As Long
    
    lblMenu(Index).Left = 650
    
    If Index = 7 Then
        If EstaHabilitado(25, "Ver Numeros del Día") > 0 Then
            frmNumerosDelDia.Show 1
        End If
        
        Exit Sub 'Numeros del día
    End If
    
    If IndiceSub = Index Then
        For I = frmSub.Width * 10 To 0 Step -1
            frmSub.Width = I / 10
        Next I
        
        InvisiblesSub
        IndiceSub = 50
    Else
        LlenarSub Index
        
        For I = 0 To 30000 Step 1
            frmSub.Width = I / 10
        Next I
    End If
End Sub

Private Sub lblSub_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblSub(Index).Left = 60
    lblSub(Index).BackColor = frMenu.BackColor
    
    MenuPrincipal IndiceSub, Index
    
    lblSub(Index).Left = 100
    lblSub(Index).BackColor = frmSub.BackColor
End Sub

Private Sub lblSub_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblSub(Index).BackColor = frmSub.BackColor - 30
    LlenarDetalle IndiceSub, Index
End Sub

Private Sub LlenarDetalle(Indice As Integer, SubIndice As Integer)
    Dim TT As String
    
    If SubIndice < 20 Then 'cuando uso el procedimiento para limpiar es > 20
        TT = "  " + UCase(lblSub(SubIndice).Caption) + vbCrLf + vbCrLf
        lblDetalle.BackColor = frmSub.BackColor
        'lblDetalle.BorderStyle = 1
    Else
        TT = ""
        lblDetalle.BackColor = Me.BackColor
        'lblDetalle.BorderStyle = 0
    End If
    
    Select Case Indice
        Case 0
            Select Case SubIndice
                Case 0
                    TT = TT + "Se utiliza para Administrar " + vbCrLf + _
                              "permisos para distintas funciones " + vbCrLf + _
                              "del programa"
                Case 1
                    TT = TT + "Cierra secion del usuario actual " + vbCrLf + _
                              "para iniciar otra "
                              
                Case 2
                    TT = TT + "Cambia la contraseña del usuario actual "
                Case 1
                
            End Select
        Case 1
        
        
    End Select

    lblDetalle = TT
End Sub

Private Sub MenuPrincipal(Indice As Integer, SubIndice As Integer)
    
    On Local Error GoTo errMP
    
    Terr.AnOtaR "abcr", Indice, SubIndice
    Dim TmP As Long, sTmp As String, F As String
    Dim Usuario As String, UUs As Long, tmpRp As Long
    Dim IdCli As Long
    
    Dim CM As New CommonDialog
    Dim FSO As New FileSystemObject
    
    Select Case Indice
        Case 0 'SISTEMA ---------------------------------------------------------------------
            Select Case SubIndice
                Case 0 '...........................................................................
                    'veo si el usuario que esta trabajando tiene habilitacion para entrar
                    If EstaHabilitado(1, "Administración de Accesos") > 0 Then
                        ACC.DefinirPermisos
                    End If
                
                Case 1 '...........................................................................
                    Usuario = ACC.GetNombre("Usuario", "Usuarios", ACC.UltUsuarioIngresado)
                    
                    If MsgBox("¿Está seguro que desea cerrar la Sesión " + Usuario + "?", _
                        vbOKCancel, "Cierre Sesión") = vbCancel Then Exit Sub
                    
                    'cierro sesión
                    ACC.RegEvento ACC.UltUsuarioIngresado, 6, "Por cambio de usuario"
                        'hago que abra el nuevo usuario
                    
                    If ACC.ValidarUsuario(5) = -1 Then 'no ingreso ninguno nuevo
                        'entonces dejo las cosas como estaban (inicia sesion otra vez)
                        MsgBox "No pudo cambiarse la sesión"
                        ACC.RegEvento ACC.UltUsuarioIngresado, 5, "No pudo cambiar de usuario"
                    Else
                        Usuario = ACC.GetNombre("Usuario", "Usuarios", ACC.UltUsuarioIngresado)
                        MsgBox "Usuario " + Usuario + " ha iniciado sesión"
                        'mnCambiaUs.Caption = "Cerrar Sesión " + Usuario
                    End If
                    
                Case 2 '...........................................................................
                    UUs = ACC.UltUsuarioIngresado
                    tmpRp = ACC.CambiarContrasena(UUs)
                    
                    Terr.AnOtaR "abcs", UUs, tmpRp
                    Select Case tmpRp
                        Case 0
                            MsgBox "Contraseña Modificada correctamente", vbInformation, "Cambio Contraseña"
                        Case 1
                            MsgBox "Hubo Datos sin Cargar", vbInformation, "Cambio Contraseña"
                        Case 2
                            MsgBox "Clave ingresada es incorrecta", vbInformation, "Cambio Contraseña"
                        Case 3
                            MsgBox "Mal confirada la Contraseña", vbInformation, "Cambio Contraseña"
                    End Select
                Case 3 '...........................................................................
                    Terr.AnOtaR "abct"
                    ACC.MostrarMovimientos
                Case 4 '...........................................................................
                    If EstaHabilitado(18, "Ingreso a Limpieza Base de Datos") > 0 Then
                        Shell AP + "limpiezabasededatos.exe", vbNormalFocus
                        Terr.AnOtaR "abcu"
                        Unload Me
                        End
                    End If
                Case 5 '...........................................................................
                    Terr.AnOtaR "abcv"
                    frmExportPV.Show 1
                Case 6 '...........................................................................
                    frmImportProd.Show 1
                Case 7 '...........................................................................
                    frmExportProd.Show 1
                    
            End Select
            
        Case 1 'PRODUCTOS -------------------------------------------------------------------
            Select Case SubIndice
                Case 0 '..........................................................................
                    frmProductosT.Show 1
                
                Case 1 '..........................................................................
                    frmProductos.AbrirDatos ""

                Case 2 '..........................................................................
                    frmTipoProducto.Show 1
                    
                Case 3 '..........................................................................
                    frmTipo.Iniciar "Sucursales"
                        
                Case 4 '..........................................................................
                
                    If DB.ContarReg("SELECT * FROM Sucursales") = 0 Then
                        MsgBox "No tiene más sucursales que CASA CENTRAL." + vbCrLf + _
                            "Se mostrará la ventana de Stock Mínimo", vbInformation, "Atención"
                        frmStockMinimo.Show 1
                    Else
                        frmStockSucursales.Show 1
                    End If
                
                Case 5 '......................................................................
                    frmStockMinimo.Show 1
                    
                Case 6 '......................................................................
                    frmHistorialMovProd.Show 1
                
                Case 7 '......................................................................
                    frmPrintOfertas.Show 1
            End Select
        
        Case 2 'CLIENTES --------------------------------------------------------------------
            Select Case SubIndice
                Case 0
                    Terr.AnOtaR "abcm"
                    'primero que todo veo si hay productos
                    If DB.ContarReg("select nproducto from Productos") = 0 Then
                        MsgBox "No hay productos cargados en la base de datos, solucione este problema " + _
                            "y luego ingrese las ventas", vbInformation, "Atención"
                        Exit Sub
                    End If
                    Terr.AnOtaR "abcn"
                    frmVENTAS.Show 1
                Case 1
                    frmAnularV.Show 1
                Case 2
                    frmClientes2.Show 1
                Case 3
                    frmInter.AbrirDatos 0, "Seleccione Cliente" + vbCrLf + _
                        "Para agregar Cliente, no seleccione nada", True
                
                    If RespuestaInter(0) = "" Then
                        frmClientes.AbrirDatos -1
                    Else
                        If CLng(RespuestaInter(0)) <> -1 Then
                            frmClientes.AbrirDatos CLng(RespuestaInter(0))
                        End If
                    End If
                        
                Case 4 'Anotar a Clientes ..................................................
                    If EstaHabilitado(8, "Anotar a Clientes") > 0 Then
                        frmClientesMov.AbrirDatos , , ""
                    End If
                Case 5
                    frmInter.AbrirDatos 0 '0 es IdCliente
                    
                    If RespuestaInter(0) = "" Then
                        'MsgBox "No seleccionó ningún cliente", vbInformation, "Atención"
                        Exit Sub
                    Else
                        IdCli = CLng(RespuestaInter(0))
                        frmResClientes.AbrirDatos IdCli
                    End If
                Case 6
                    frmVerFactura.AbrirDatos -1
                Case 7
                    If DB.ContarReg("SELECT * FROM VENTAS") = 0 Then
                        MsgBox "Aún no tiene registrado ventas", vbInformation, "Sin registro Ventas"
                    Else
                        frmMovProd.Show 1
                    End If

                Case 8
                    'tiene envases?????????
                    If CFG.GetInfo(5, 4) <> "No" Then lstEnvases.Show 1
                Case 9
                    frmClientesOrg.Show 1
                Case 10
                    frmComentarios.Show 1
                Case 11
                    frmVencimientos.AbrirDatos False
                Case 12
                    'busco el id menor, si no es menor que -20 no hay financieras
                    Dim IdOp As Long
                    
                    IdOp = DB.GetTop1Rs("Clientes", "ID", "ASC")
                    If IdOp > -20 Then IdOp = -20
                    
                    frmClientes.AbrirDatos IdOp
        
            End Select
        
        Case 3 'PROVEEDORES -----------------------------------------------------------------
            Select Case SubIndice
                Case 0 'COMPRAS ...................................................................
                    'primero que todo veo si hay productos
                    If DB.ContarReg("SELECT nProducto FROM Productos") = 0 Then
                        MsgBox "No hay productos cargados en la base de datos, solucione este problema " + _
                            "y luego ingrese las compras", vbInformation, "Atención"
                        Terr.AppendSinHist "NohayProd"
                        Exit Sub
                    End If
                    
                    ' y si no hay proveedores?
                    If DB.ContarReg("select proveedor from Proveedores") = 0 Then
                        MsgBox "No hay Proveedores cargados en la base de datos, solucione este problema " + _
                            "y luego ingrese las compras", vbInformation, "Atención"
                        Terr.AppendSinHist "NohayProv"
                        Exit Sub
                    End If
                    
                    Terr.AnOtaR "caaa"
                    frmCompras.AbrirDatos
                
                Case 1
                    frmProveedores.Show 1
                Case 2
                    frmPedidos.Show 1
                Case 3
                    frmClientesMov.AbrirDatos , True
                
                Case 4 'RESUMEN DE CUENTA ........................................................
                    ' y si no hay proveedores
                    If DB.ContarReg("select proveedor from Proveedores") = 0 Then
                        MsgBox "No hay Proveedores cargados en la base de datos", vbInformation, "Atención"
                        Exit Sub
                    End If
                    
                    frmProveedoresRes.Show 1
                Case 5
                    frmVerFactura.AbrirDatos -1, True
                Case 6
                    frmVencimientos.AbrirDatos True
            End Select
        
        Case 4 'SOCIOS Y EMPLEADOS ----------------------------------------------------------
            Select Case SubIndice
                
                Case 0
                    'veo si el usuario que esta trabajando tiene habilitacion para entrar
                    If EstaHabilitado(7, "Ingreso a Ventana Socios") > 0 Then
                        frmSocyEmp.AbrirDatos True
                    End If
                Case 1
                    If EstaHabilitado(3, "Ingreso Ventana de Pago Sueldos") > 0 Then
                        frmSocyEmp.AbrirDatos False
                    End If
                
                Case 2 'COMISIONES POR VENTAS.......................................................
                    '1ro que como minimo haya un empleado------------------------------------------
                     Dim IdE() As String
                     
                     IdE = PC.GetCuentas(53)
                     If UBound(IdE) = 0 Then
                         MsgBox "No tiene registrados empleados", vbInformation, "Atención"
                         Exit Sub
                     End If
                     
                     '-----------------------------------------------------------------------------
                     
                     '2do que todo controlo el acceso ----------------------------------------------
                     If EstaHabilitado(3, "Ingreso Ventana Comisiones por Ventas") > 0 Then
                         frmComisiones.Show 1
                     End If

                Case 3 'MERCADERÍA A COSTO .......................................................
                
                    If EstaHabilitado(11, "Ingreso Ventana Devoluciones y Merc. a Costo") > 0 Then
                        frmDevolucion.Show 1
                    End If
                    
            End Select
            
        Case 5 'CONTABILIDAD ----------------------------------------------------------------
            Select Case SubIndice
                Case 0 'CERRAR CAJA .............................................................
                    If EstaHabilitado(2, "Abrir Ventana Cierre de Caja") > 0 Then
                         frmCierreCaja.Show
                    End If
                Case 1
                    If EstaHabilitado(23, "Ingreso Ventana Cierres Anteriores") > 0 Then
                        PC.CierresViejos
                    End If
                Case 2
                    frmAyudaCaja.Show
                    
                Case 3 'CERRAR RESULTADOS ...........................................................
                    'primero obligo que cierre de caja antes de hacer resultados
                    If EstaHabilitado(2, "Ingreso a Ventana Cierre de Resultados") > 0 Then
                        If PC.CerroCaja Then
                            frmEstRes.Show 1
                        Else
                            MsgBox "No puede Cerrar resultados, realice primero el Cierre de Caja ", _
                                vbInformation, "Atención"
                        End If
                    End If
                Case 4
                    frmResultados.Show 1
                Case 5
                    frmMovCtasCont.Show 1
                Case 6
                    frmBalance.Show 1
                Case 7
                    frmIVA.AbrirDatos False
                    
                Case 8
                    frmIVA.AbrirDatos True
                Case 9
                    frmEgresos.Show 1
                
                Case 10 'ASIENTOS .....................................................................
                    'veo si el usuario que esta trabajando tiene habilitacion para entrar
                    If EstaHabilitado(24, "Ingreso Ventana Asiento-Configuración") > 0 Then
                        frmCuentas.Show 1
                    End If
            End Select
        Case 6 'CONFIGURACIONES -------------------------------------------------------------
            Select Case SubIndice
                Case 0
                    frmClientes.AbrirDatos -2
                Case 1
                    'veo si el usuario que esta trabajando tiene habilitacion para entrar
                    If EstaHabilitado(24, "Ingreso Ventana Configuraciones Generales") > 0 Then
                        frmConfig.Show 1
                    End If
                Case 2
                    If EstaHabilitado(24, "Ingreso Ventana Conf Factura Venta") > 0 Then
                        frmConfigFacVta.AbrirDatos
                    End If
            Case 3
                    If EstaHabilitado(24, "Ingreso Ventana Conf Factura Compra") > 0 Then
                        frmConfigFacVta.AbrirDatos False
                    End If
                Case 4
                    frmFactura.Show 1
                Case 5
                    frmConfigImprCP.Show 1
                    
            End Select
    End Select
    
    Set FSO = Nothing
    Set CM = Nothing
    
    Exit Sub
errMP:
    Terr.AppendLog "errMP", Terr.ErrToTXT(Err)
    Resume Next
End Sub

