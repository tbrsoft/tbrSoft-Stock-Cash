VERSION 5.00
Object = "{A7FBD38D-2930-49E3-B60C-9E0202D84549}#15.0#0"; "tbrControles.ocx"
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmHistorialMovProd 
   BackColor       =   &H00544B45&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Historial Movimientos Productos"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11415
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHistorialMovProd.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin tbrFaroButton.fBoton cmdBorrarViejos 
      Height          =   435
      Left            =   390
      TabIndex        =   12
      Top             =   7890
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Borrar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrControles.MouTextBox txtRegVie 
      Height          =   375
      Left            =   1740
      TabIndex        =   11
      Top             =   7920
      Width           =   885
      _ExtentX        =   1561
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
      Entero          =   -1  'True
   End
   Begin tbrControles.tbrBuscador tbrBuscadorP 
      Height          =   2295
      Left            =   180
      TabIndex        =   0
      Top             =   1530
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   4048
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
   Begin VB.CheckBox chkT 
      BackColor       =   &H00544B45&
      Caption         =   "Ver Historial Todos los productos"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   6120
      TabIndex        =   3
      Top             =   3180
      Width           =   4065
   End
   Begin VB.ComboBox cmbSucursal2 
      Height          =   315
      Left            =   6090
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1560
      Width           =   2985
   End
   Begin MSDataGridLib.DataGrid DGHi 
      Height          =   3405
      Left            =   180
      TabIndex        =   6
      Top             =   3960
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   6006
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16777215
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   1
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
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
            Type            =   2
            Format          =   "0,000E+00"
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
         MarqueeStyle    =   3
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin tbrFaroButton.fBoton cmdSalir 
      Height          =   435
      Left            =   9900
      TabIndex        =   13
      Top             =   7890
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "registros más viejos"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      Top             =   7980
      Width           =   1875
   End
   Begin VB.Label lblTotalReg 
      BackStyle       =   0  'Transparent
      Caption         =   "Busqueda Producto"
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   720
      TabIndex        =   4
      Top             =   7620
      Width           =   3465
   End
   Begin VB.Label lblProdSel 
      Alignment       =   2  'Center
      BackColor       =   &H00D8D9D7&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   945
      Left            =   6090
      TabIndex        =   10
      Top             =   2100
      Width           =   3075
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione Sucursal"
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   6120
      TabIndex        =   9
      Top             =   1260
      Width           =   2025
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Busqueda Producto"
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   690
      TabIndex        =   8
      Top             =   1230
      Width           =   1965
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Por Nombre o Código"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   2490
      TabIndex        =   7
      Top             =   1230
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Historial Movimientos Productos"
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
      Height          =   585
      Left            =   2190
      TabIndex        =   5
      Top             =   300
      Width           =   7365
   End
End
Attribute VB_Name = "frmHistorialMovProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsHi As New ADODB.Recordset
Dim ActuTbrB As Boolean

Private Sub chkT_Click()
    RenovarRS
End Sub

Private Sub cmbSucursal2_Click()
    RenovarRS
End Sub

Private Sub cmdBorrarViejos_Click()
    txtRegVie = CStr(CLng(ValidarNumeros(txtRegVie)))
    BorrarViejos CLng(txtRegVie)
    RenovarRS
    MsgBox "Se realizó correctamente el borrado"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    ActuTbrB = False
    
    txtRegVie = "0"
    cmbSucursal2.Clear
    cmbSucursal2.AddItem "TODOS"
    cmbSucursal2.AddItem "CASA CENTRAL"
    CargarCombo cmbSucursal2, "SELECT Sucursal FROM Sucursales", "Sucursal", , True
    cmbSucursal2.ListIndex = 0 'si no hay sucursales no lo hace
    
    tbrBuscadorP.Contrasena = Contrasena
    tbrBuscadorP.ArchivoMDB = ArchivoMDBPrincipal
    tbrBuscadorP.SqlSinLike = "SELECT TOP 50 Productos.ID, " + _
        "TipoProductos.TipoProducto, Productos.nProducto " + _
        "FROM TipoProductos INNER JOIN Productos ON TipoProductos.ID2 = " + _
        "Productos.IdTipoProducto WHERE Productos.ID >=0"
    tbrBuscadorP.OrderBy = "ORDER BY ID"

    tbrBuscadorP.CampoEnQueBuscar = "Id/n,TipoProducto,nproducto/b"
    tbrBuscadorP.ColumnasSepPorComasyParentesis = "ID(600)/Tipo(1500)/Producto(2500)"
    tbrBuscadorP.Text = "saf"
    tbrBuscadorP.Text = "saf"
    tbrBuscadorP.Text = ""
    
    If tbrBuscadorP.GetLstSel <> "" Then
        RenovarRS
    Else
        lblProdSel = ""
    End If
    
    LimpiarMovProdViejos
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    tbrBuscadorP.CN_CLOSE
    
    Set DGHi.DataSource = Nothing
    If RsHi.State = adStateOpen Then RsHi.Close
    Set RsHi = Nothing
End Sub

Private Sub RenovarRS()
    Set DGHi.DataSource = Nothing
    If RsHi.State = adStateOpen Then RsHi.Close
    
    Dim S As String
    
    S = "SELECT MovimientosProductos.Fecha, MovimientosProductos.Hora, " + _
        "Productos.nProducto, MovimientosProductos.Sucursal, " + _
        "MovimientosProductos.Variacion, MovimientosProductos.StockLuego, " + _
        "MovimientosProductos.Usuario, MovimientosProductos.Detalle " + _
        "FROM Productos INNER JOIN MovimientosProductos ON Productos.ID = " + _
        "MovimientosProductos.IDProducto"
    
    If chkT.Value = 0 Then
        If tbrBuscadorP.GetLstSel = "" Then
            lblProdSel = "Ningún Producto seleccionado"
            S = S + " WHERE Productos.ID = -10203"
        Else
            S = S + " WHERE Productos.ID = " + tbrBuscadorP.GetLstSel(0)
            
            If cmbSucursal2 <> "TODOS" Then
                S = S + " AND Sucursal = '" + cmbSucursal2 + "'"
            End If
            
            lblProdSel = "Producto Seleccionado: " + UCase(tbrBuscadorP.GetLstSel(2))
        End If
    Else
        lblProdSel = "Estadísticas de movimientos de todos los productos"
    End If
    
    S = S + " ORDER BY MovimientosProductos.ID DESC"
    
    RsHi.CursorLocation = adUseClient
    RsHi.Open S, DB.CN, adOpenStatic, adLockReadOnly
    
    Set DGHi.DataSource = RsHi
    AcomodarDG
        
    lblTotalReg = "Tiene Registrados " + _
        CStr(DB.ContarReg("SELECT * FROM MovimientosProductos")) + _
        " Movimientos"
End Sub

Private Sub AcomodarDG()
    'DGHi.Columns("ID").Width = 0
    DGHi.Columns("Fecha").Width = 1100
    DGHi.Columns("Hora").Width = 500
    If chkT Then
        DGHi.Columns("nProducto").Width = 1500
        Me.Width = 12000
        DGHi.Width = 11565
    Else
        DGHi.Columns("nProducto").Width = 0
        Me.Width = 10500
        DGHi.Width = 10065
    End If
        
    DGHi.Columns("Sucursal").Width = 1350
    DGHi.Columns("Variacion").Width = 950
    DGHi.Columns("Variacion").Caption = "Var."
    DGHi.Columns("StockLuego").Width = 950
    DGHi.Columns("StockLuego").Caption = "Stock"
    DGHi.Columns("Usuario").Width = 1300
    DGHi.Columns("Detalle").Width = 3300
    
    DGHi.Columns("Fecha").Alignment = dbgCenter
    DGHi.Columns("Hora").Alignment = dbgCenter
    DGHi.Columns("Var.").Alignment = dbgCenter
    DGHi.Columns("Stock").Alignment = dbgRight
    
    DGHi.RowHeight = 280
    
    cmdSalir.Left = DGHi.Left + DGHi.Width - cmdSalir.Width
End Sub

Private Sub Image1_Click()

End Sub

Private Sub tbrBuscadorP_Change()
    If IsNumeric(tbrBuscadorP.Text) Then
        tbrBuscadorP.CampoEnQueBuscar = "id/b,tipoproducto,nproducto"
    Else
        tbrBuscadorP.CampoEnQueBuscar = "id/n,tipoproducto,nproducto/b"
    End If
    
    If tbrBuscadorP.Text <> "" Then VerActu
    RenovarRS
End Sub

Private Sub tbrBuscadorP_Click()
    RenovarRS
End Sub

Private Sub VerActu()
    If ActuTbrB = False Then
       ActuTbrB = True
       tbrBuscadorP.Recargar
    Else
        ActuTbrB = False
    End If

End Sub

Private Sub txtRegVie_GotFocus()
    PintarTxt txtRegVie
End Sub

Private Sub txtRegVie_LostFocus()
    txtRegVie = CStr(CLng(ValidarNumeros(txtRegVie)))
End Sub
