VERSION 5.00
Object = "{A7FBD38D-2930-49E3-B60C-9E0202D84549}#15.0#0"; "tbrControles.ocx"
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmAnularV 
   BackColor       =   &H00544B45&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Anular Ventas"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   Icon            =   "frmAnularV.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin tbrControles.tbrBuscador tbrBuscador1 
      Height          =   4305
      Left            =   480
      TabIndex        =   0
      Top             =   540
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   7594
      BackColor       =   5524293
      BeginProperty Fontt {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
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
   Begin tbrFaroButton.fBoton cmdSalir 
      Height          =   480
      Left            =   3570
      TabIndex        =   2
      Top             =   5150
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   847
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   9
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdEliminar 
      Height          =   480
      Left            =   1470
      TabIndex        =   3
      Top             =   5150
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   847
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Eliminar"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   9
      fECol           =   5717301
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione la venta del día que desea anular."
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
      Left            =   990
      TabIndex        =   1
      Top             =   270
      Width           =   4635
   End
End
Attribute VB_Name = "frmAnularV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ID As Long
Dim Cant As Long
Dim Producto As String
Dim Precio As Single
Dim Costo As Single

Private Sub cmdEliminar_Click()
    If tbrBuscador1.GetLstSel = "" Then Exit Sub
      'le doy valores a la variable con los datos del registro seleccionado
    ID = tbrBuscador1.GetLstSel(4)
    Costo = tbrBuscador1.GetLstSel(3)
    Cant = tbrBuscador1.GetLstSel(0)
    Producto = tbrBuscador1.GetLstSel(1)
    Precio = tbrBuscador1.GetLstSel(2)
    If MsgBox("¿Está seguro que desea eliminar la venta de " + CStr(Cant) + " " + _
        Producto + "?", vbExclamation + vbOKCancel, "Atención") = vbCancel Then Exit Sub
     
     'sigue aca si dio OK
    Dim rSeLiminar As New ADODB.Recordset
     'para descontar vtasup de la tabla caja
    Dim RestarVentas As Single
    Dim restarCosto As Single
    
    RestarVentas = (Cant * Precio)
    restarCosto = (Cant * Costo)
    
    DB.EXECUTE "DELETE FROM Ventas WHERE id=" + CStr(ID)
    Dim clP As New clsProducto
        clP.ModificarStock DB.GetValInRS("Productos", "ID", "nProducto = '" + _
            Producto + "'", False), Cant, , "Por anulación de venta"
    Set clP = Nothing
    
    'anulo venta
    PC.Asiento "80", CStr(RestarVentas), "78", CStr(RestarVentas)
    'anulo costo
    PC.Asiento "54", CStr(restarCosto), "18", CStr(restarCosto)
    
    RestarVentas = 0 'empiezo de vuelta
    restarCosto = 0
    tbrBuscador1.Recargar
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()

    tbrBuscador1.Contrasena = Contrasena
    tbrBuscador1.ArchivoMDB = ArchivoMDBPrincipal
    tbrBuscador1.SqlSinLike = "Select ventas.idproducto, ventas.id, productos.nProducto," + _
        "ventas.cantidad, productos.id,ventas.precio, ventas.costo from Productos INNER JOIN Ventas on " + _
        "Productos.id = Ventas.idproducto " + _
        "where ventas.fecha = #" + stFechaSQL(Date) + "# AND Productos.ID >=0"
    tbrBuscador1.CampoEnQueBuscar = "cantidad/n,nproducto/b,precio/$,costo/$,ventas.id/n"
    tbrBuscador1.OrderBy = "order by ventas.id desc"
    tbrBuscador1.ColumnasSepPorComasyParentesis = "Cant(600)/Producto(2000)/" + _
        "Precio(980)/Costo(980)/IDVta(0)"

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    FormatearMouTextBox frmAnularV, 8
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    tbrBuscador1.CN_CLOSE
    
    VtaDia = PC.ABSSumarconSubcuentas(17, False)
    CtoDia = PC.ABSSumarconSubcuentas(18, False)
End Sub

