VERSION 5.00
Object = "{A7FBD38D-2930-49E3-B60C-9E0202D84549}#15.0#0"; "tbrControles.ocx"
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmTipoProducto 
   BackColor       =   &H00544B45&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipos de Producto"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   12375
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTipoProducto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   12375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin tbrFaroButton.fBoton command6 
      Height          =   435
      Left            =   10425
      TabIndex        =   16
      Top             =   1095
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Eliminar imagen"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton command5 
      Height          =   435
      Left            =   8715
      TabIndex        =   15
      Top             =   1095
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Elegir imagen"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.Frame fr_IMAGENES 
      BackColor       =   &H00544B45&
      Caption         =   "Imagenes del producto"
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   8640
      TabIndex        =   5
      Top             =   900
      Width           =   3540
      Begin VB.PictureBox picFONDO 
         BackColor       =   &H00E6DFD5&
         Height          =   2055
         Left            =   90
         ScaleHeight     =   1995
         ScaleWidth      =   3255
         TabIndex        =   6
         Top             =   720
         Width           =   3315
         Begin VB.Image imgPV 
            Height          =   1245
            Left            =   690
            Top             =   390
            Width           =   1665
         End
      End
   End
   Begin VB.ListBox lstCuentas 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   5925
      IntegralHeight  =   0   'False
      Left            =   150
      TabIndex        =   4
      Top             =   960
      Width           =   8355
   End
   Begin tbrControles.MouTextBox txtMargenVenta 
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   7110
      Width           =   1275
      _ExtentX        =   2249
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
   End
   Begin tbrFaroButton.fBoton command1 
      Height          =   405
      Left            =   10860
      TabIndex        =   12
      Top             =   7905
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdGrabar 
      Height          =   435
      Left            =   3540
      TabIndex        =   9
      Top             =   7080
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Grabar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdAplicarCambios 
      Height          =   435
      Left            =   1485
      TabIndex        =   10
      Top             =   7650
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Aplicar Cambios"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdGrabarTodos 
      Height          =   435
      Left            =   3255
      TabIndex        =   11
      Top             =   7650
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Grabar a todos los demás Tipos de Productos"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdAgregarR 
      Height          =   405
      Left            =   8685
      TabIndex        =   1
      Top             =   5460
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Agregar Principal"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdModificar 
      Height          =   405
      Left            =   8685
      TabIndex        =   2
      Top             =   5910
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Modificar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdAgregar 
      Height          =   400
      Left            =   8685
      TabIndex        =   0
      Top             =   4995
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Agregar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdEliminar 
      Height          =   405
      Left            =   8670
      TabIndex        =   3
      Top             =   6390
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Eliminar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Para próximos productos. Para anteriores Presione Aplicar Cambios."
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   4950
      TabIndex        =   14
      Top             =   7110
      Width           =   3630
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Margen de Venta"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   540
      TabIndex        =   13
      Top             =   7170
      Width           =   1545
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos de Producto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   3210
      TabIndex        =   7
      Top             =   240
      Width           =   4995
   End
End
Attribute VB_Name = "frmTipoProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FSO As New Scripting.FileSystemObject

Private Sub cmdAgregar_Click()
    Dim NbCta As String, nCta As String
    
    If lstCuentas.ListIndex = -1 Then Exit Sub
    
    nCta = InputBox("Ingrese el ID del Tipo de Producto que desee agregar", _
        "Id Tipo Producto", CStr(IdAutonum("TipoProductos", "ID2")))
    
    If Not IsNumeric(nCta) Then
        MsgBox "ID incorrecto", vbInformation, "Atención"
        Exit Sub
    End If
    
    If CLng(nCta) < 0 Then
        MsgBox "ID incorrecto, debe ser un número positivo", vbInformation, "Atención"
        Exit Sub
    End If
    
    nCta = CStr(CLng(nCta))
    NbCta = Replace(InputBox("Ingrese el nombre del tipo de producto", "Nombre"), ".", " ")
    
    If NbCta = "" Then Exit Sub
    
    'crear una cuenta en el nivel elegido
    Dim SP() As String
    SP = Split(lstCuentas, ".")
    
    If DB.ContarReg("SELECT * FROM TipoProductos WHERE TipoProducto = '" + _
        NbCta + "'") > 0 Then
        MsgBox "Ya existe un Tipo de Producto con ese nombre", vbInformation, "Atención"
        Exit Sub
    End If
    
    If DB.ContarReg("SELECT * FROM TipoProductos WHERE ID2 = " + _
        CStr(nCta)) > 0 Then
        MsgBox "Ya existe un Tipo de Producto con ese ID", vbInformation, "Atención"
        Exit Sub
    End If
    
    'si llegó hasta acá agrego el tipo de producto
    DB.EXECUTE "INSERT INTO TipoProductos (ID2,IDAnt,TipoProducto) VALUES (" + _
        CStr(nCta) + "," + CStr(SP(1)) + ",'" + NbCta + "')"
    
    'para que se actualice medio negra la cosa
    If Right(lstCuentas, 1) = "*" Then
        lstCuentas_DblClick
        lstCuentas_DblClick
    Else
        lstCuentas_DblClick
    End If
End Sub

Private Sub cmdAgregarR_Click()
    Dim NbCta As String, nCta As String
    
    nCta = InputBox("Ingrese el ID del Tipo de Producto que desee agregar", _
        "Id Tipo Producto", CStr(IdAutonum("TipoProductos", "ID2")))
    
    If Not IsNumeric(nCta) Then
        MsgBox "ID incorrecto", vbInformation, "Atención"
        Exit Sub
    End If
    
    If CLng(nCta) < 0 Then
        MsgBox "ID incorrecto, debe ser un número positivo", vbInformation, "Atención"
        Exit Sub
    End If
    
    nCta = CStr(CLng(nCta))
    NbCta = Replace(InputBox("Ingrese el nombre del tipo de producto", "Nombre"), ".", " ")
    
    If NbCta = "" Then Exit Sub
    
    If DB.ContarReg("SELECT * FROM TipoProductos WHERE TipoProducto = '" + _
        NbCta + "'") > 0 Then
        MsgBox "Ya existe un Tipo de Producto con ese nombre", vbInformation, "Atención"
        Exit Sub
    End If
    
    If DB.ContarReg("SELECT * FROM TipoProductos WHERE ID2 = " + _
        CStr(nCta)) > 0 Then
        MsgBox "Ya existe un Tipo de Producto con ese ID", vbInformation, "Atención"
        Exit Sub
    End If
    
    'si llegó hasta acá agrego el tipo de producto
    DB.EXECUTE "INSERT INTO TipoProductos (ID2,IDAnt,TipoProducto) VALUES (" + _
        CStr(nCta) + ",0,'" + NbCta + "')"
    
    'para que se actualice
    CargarDatos
End Sub

Private Sub cmdAplicarCambios_Click()
    'primero ya lo dejo grabado
    cmdGrabar_Click
    
    If lstCuentas.ListIndex = -1 Then Exit Sub
    'no se porque pero le tengo que sacar el %
    Dim SP() As String, Mult As Single
    
    SP = Split(txtMargenVenta, "%")
    txtMargenVenta = SP(0)
    If Not IsNumeric(txtMargenVenta) Then
        txtMargenVenta = FormatPercent(0)
        Exit Sub
    End If
    
    Mult = CSng(txtMargenVenta) / 100 + 1
    'vuelvo a ponerlo bien
    txtMargenVenta = FormatPercent(CSng(txtMargenVenta) / 100)
    
    Dim RsT As New ADODB.Recordset
    RsT.Open "SELECT ID,pCosto FROM Productos WHERE IdTipoProducto = " + CStr(GetID), DB.CN, adOpenStatic, adLockReadOnly
    
    If RsT.RecordCount > 0 Then
        RsT.MoveFirst
        Do While Not RsT.EOF
            DB.EXECUTE "UPDATE Productos SET pVenta = " + _
                Replace(Mult * CSng(RsT("pCosto")), ",", ".") + _
                " WHERE ID = " + CStr(RsT("ID"))
            RsT.MoveNext
        Loop
    End If
    
    RsT.Close
    Set RsT = Nothing
    
    MsgBox "Los cambios se realizaron satisfactoriamente", vbInformation, "Cambios Precios s/Margen de Venta"
End Sub

Private Sub cmdEliminar_Click()
    'permitir borrar cuando no hay ningun producto en ninguna de los subgrupos
    ' puede ser que para simplificar permita pasar todos los productos en este tipo
    ' a otro para que quede vacio y simpifique el tramite
    
    'ver el id del tipo elegido
    Dim SP() As String
    SP = Split(lstCuentas, ".")
    'POR QUE PASA ESTO ARTIME ???
    If UBound(SP) = 0 Then
        MsgBox "No se puede saber cual esta elegido!"
        Exit Sub
    End If
    
    'sp(1) es el codigo elegido
    Dim CTA As Long
    CTA = CLng(SP(1))
    Dim tmpIx As Long
    tmpIx = lstCuentas.ListIndex
    
    'LISTAR TODOS LOS CODIGOS EN QUE DEBO BUSCAR PRODUCTOS
    Dim ListaCODs() As Long
    ReDim ListaCODs(0)
    'el primero es el tipo elegido!
    ListaCODs(0) = CTA
    
    'ver cuantos prouctos tienen este tipo
    Dim Ctas() As String
    Ctas = GetHijoTipo(CLng(CTA))
    If UBound(Ctas) = 0 Then
        ' no hay subcuentas
        GoTo CONTARPRODS
    End If
    
    Dim J As Long
    'el cero dice "NADA" lo hizo artime ¿?¿?¿?¿?
    For J = 1 To UBound(Ctas)
        ReDim Preserve ListaCODs(UBound(ListaCODs) + 1)
        ListaCODs(UBound(ListaCODs)) = CLng(Ctas(J))
    Next J
    
    'hacer una rutina recursiva para descubrir todo el arbol de grupos
    
    Dim K As Long
    'el indice cero ya lo revise, es el elegido
    For J = 1 To UBound(ListaCODs)
        Ctas = GetHijoTipo(ListaCODs(J))
        If UBound(Ctas) > 0 Then
            For K = 1 To UBound(Ctas)
                'se agregan y se extiende el for
                ReDim Preserve ListaCODs(UBound(ListaCODs) + 1)
                ListaCODs(UBound(ListaCODs)) = CLng(Ctas(K))
            Next K
        End If
    Next J
    
    'ver cuantos productos en cada subtipo de el elegido
CONTARPRODS:
    
    Dim RSTipo As New ADODB.Recordset
    RSTipo.CursorLocation = adUseClient
    
    Dim totPROD As Long
    totPROD = 0
    
    For J = 0 To UBound(ListaCODs)
        
        RSTipo.Open "Select idtipoproducto from Productos " + _
            "WHERE idtipoproducto=" + CStr(ListaCODs(J)), _
            DB.CN, adOpenStatic, adLockReadOnly
            
        totPROD = totPROD + RSTipo.RecordCount
        RSTipo.Close
    Next J
    
    'eliminar si se puede
    If totPROD = 0 Then
        'borrar todos sus subgrupos
        DB.EXECUTE "DELETE FROM TipoProductos WHERE IDant = " + SP(1)
        'y el grupo
        DB.EXECUTE "DELETE FROM TipoProductos WHERE ID2 = " + SP(1)
        'quitar del listBox!
    Else
        MsgBox "Existen " + CStr(totPROD) + " productos en el tipo elegido" + vbCrLf + _
            "Asígnelos a otra cuenta o elimínelos antes de poder eliminar", vbExclamation, "Atención"
    End If
    
    CargarDatos
End Sub

Private Sub cmdGrabar_Click()
    If lstCuentas.ListIndex = -1 Then Exit Sub
    If txtMargenVenta = "" Then txtMargenVenta = "0"
    'no se porque pero le tengo que sacar el %
    Dim SP() As String
    SP = Split(txtMargenVenta, "%")
    txtMargenVenta = SP(0)
        
    'si tiene configuracion particular modifico, si no agrego
    Dim IDC As Long
    IDC = CFG.ExistePropiedad("MDV " + CStr(GetID))
    
    If IDC = 0 Then
        CFG.AgregarNodo 50, "MDV " + CStr(GetID), "", CStr(CSng(txtMargenVenta)), 0
    Else
        CFG.ModificarNodo IDC, , , , CStr(CSng(txtMargenVenta))
    End If
    
    txtMargenVenta = FormatPercent(ValidarNumeros(txtMargenVenta) / 100)
End Sub

Private Sub cmdGrabarTodos_Click()
    'lo grabo como general, no hace falta que haga uno general por cada uno
    
    'no se porque pero le tengo que sacar el %
    Dim SP() As String
    SP = Split(txtMargenVenta, "%")
    txtMargenVenta = SP(0)
    CFG.ModificarNodo 50, , , , CStr(CSng(txtMargenVenta))
    txtMargenVenta = FormatPercent(ValidarNumeros(txtMargenVenta) / 100)
End Sub

Private Sub cmdModificar_Click()
    Dim nCta As String, nUcTa As String, TmP As String
    Dim SP() As String
    
    If lstCuentas.ListIndex = -1 Then Exit Sub
    SP = Split(lstCuentas, ".")
    
    nUcTa = InputBox("Ingrese el ID del Tipo de Producto si lo desea modificar", _
        "Número de cuenta nueva", SP(1))
    
    If Not IsNumeric(nUcTa) Then
        MsgBox "Número de cuenta incorrecto", vbInformation, "Atención"
        Exit Sub
    End If
    
    If CLng(nUcTa) < 0 Then
        MsgBox "ID incorrecto, debe ser un número positivo", vbInformation, "Atención"
        Exit Sub
    End If
    
    nUcTa = CStr(CLng(nUcTa))
    
    Dim tmpExNombre As String
    'por los * que le pone a las cuentas abiertas
    If Right(SP(2), 1) = "*" Then
        tmpExNombre = Left(SP(2), Len(SP(2)) - 1)
        TmP = "*"
    Else
        tmpExNombre = SP(2)
        TmP = ""
    End If
    nCta = Replace(InputBox("Ingrese el nuevo nombre del Tipo de Producto", _
        "Nombre", tmpExNombre), ".", " ")
    
    If nCta = "" Then Exit Sub
    
    nCta = Replace(nCta, "'", " ")
    nCta = Replace(nCta, "*", " ")
    
    'modificar la cuenta
    If DB.ContarReg("SELECT * FROM TipoProductos WHERE TipoProducto = '" + _
        nCta + "'") > 0 Then
        If nCta <> SP(2) Then
            MsgBox "Ya existe un Tipo de Producto con ese nombre", vbInformation, "Atención"
            Exit Sub
        End If
    End If
    
    If DB.ContarReg("SELECT * FROM TipoProductos WHERE ID2 = " + _
        CStr(nUcTa)) > 0 Then
        If nUcTa <> CLng(SP(1)) Then
            MsgBox "Ya existe un Tipo de Producto con ese ID", vbInformation, "Atención"
            Exit Sub
        End If
    End If
    
    'modifico nomas
    DB.EXECUTE "UPDATE TipoProductos SET ID2 = " + CStr(nUcTa) + ", " + _
        "TipoProducto = '" + nCta + "'" + _
        "WHERE ID2 = " + SP(1)
    
    'actualizo
    lstCuentas.List(lstCuentas.ListIndex) = SP(0) + "." + CStr(nUcTa) + "." + nCta + TmP
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command5_Click()
    If lstCuentas.ListIndex = -1 Then
        MsgBox "No eligio ningun tipo al que aplicar la imagen"
        Exit Sub
    End If
    
    Dim CM As New CommonDialog
    CM.DialogTitle = "Elija una imagen para el TIPO de producto"
    CM.Filter = "Imagenes(jpg bmp gif tiff)|*.jpg; *.jpeg; *.bmp;*.gif;*.tiff"
    CM.InitDir = AP
    CM.ShowOpen
    
    Dim F As String
    F = CM.FileName
    If F = "" Then Exit Sub
    
    'ver si ya existe
    Dim SP() As String, F2 As String
    
    SP = Split(lstCuentas, ".")
    
    F2 = CFGBD.GetInfo(82, 4) + "img\TP" + SP(1) + ".jpg"
    If FSO.FileExists(F2) Then
        If MsgBox("Ya existe una imagen para este TIPO, ¿desea reemplazarla?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        FSO.DeleteFile F2, True
    End If
    
    FSO.CopyFile F, F2
    
    LoadImgPV F2
End Sub

Private Sub LoadImgPV(iAR As String)
    imgPV.Stretch = False 'para que tome el tamaño que tiene que ser
    imgPV.Picture = LoadPicture(iAR)
    'ver ahi la proporcion
    Dim Prop As Single
    Prop = imgPV.Width / imgPV.Height
    'definbir el final segun corresponda
    Dim Ancho As Single, Alto As Single
    'probar si entraria con ancho maximo ...
    Ancho = PicFondo.Width
    Alto = Ancho / Prop
    If Alto > PicFondo.Height Then
        'cambiar todo! supuestamente si fallo el otro este no falla
        Alto = PicFondo.Height
        Ancho = Alto * Prop
    End If
    
    imgPV.Stretch = True
    imgPV.Width = Ancho
    imgPV.Height = Alto
    
    imgPV.Top = PicFondo.Height / 2 - imgPV.Height / 2
    imgPV.Left = PicFondo.Width / 2 - imgPV.Width / 2
End Sub


Private Sub Command6_Click()
    If lstCuentas.ListIndex = -1 Then
        MsgBox "No eligio ningun tipo"
        Exit Sub
    End If
    
    'ver si ya existe
    Dim SP() As String
    SP = Split(lstCuentas, ".")
    
    F2 = CFGBD.GetInfo(82, 4) + "img\TP" + SP(1) + ".jpg"
    If FSO.FileExists(F2) Then FSO.DeleteFile F2, True
    imgPV.Picture = LoadPicture
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    CargarDatos
End Sub

Private Sub CargarDatos()
    Dim Y As Long, IDTp As Long
    
    CargarCombo lstCuentas, "SELECT TipoProducto FROM TipoProductos WHERE IdAnt = 0 " + _
        "AND ID2 > 0", "TipoProducto"
    
    For Y = 0 To lstCuentas.ListCount - 1
        IDTp = DB.GetValInRS("TipoProductos", "ID2", "TipoProducto = '" + _
            lstCuentas.List(Y) + "'", False)
        lstCuentas.List(Y) = "." + CStr(IDTp) + "." + lstCuentas.List(Y)
    Next Y
    
    lstCuentas_Click
End Sub

Private Sub lstCuentas_Click()
    If lstCuentas.ListIndex = -1 Then Exit Sub
    
    Dim SP() As String
    SP = Split(lstCuentas, ".")
    imgPV.Picture = LoadPicture
    'POR QUE PASA ESTO ARTIME ???
    If UBound(SP) = 0 Then Exit Sub
    
    Dim sTmp As String
    sTmp = CFGBD.GetInfo(82, 4) + "IMG\TP" + SP(1) + ".jpg"
    If FSO.FileExists(sTmp) Then LoadImgPV sTmp
    
    'si tiene configuracion particular la pongo si no uso la general
    Dim IDC As Long
    IDC = CFG.ExistePropiedad("MDV " + SP(1))
    
    If IDC = 0 Then
        txtMargenVenta = FormatPercent(CSng(CFG.GetInfo(50, 4)) / 100)
    Else
        txtMargenVenta = FormatPercent(CSng(CFG.GetInfo(IDC, 4)) / 100)
    End If
End Sub

Private Sub lstCuentas_DblClick()
    'mostrar los subniveles del elegido
    Dim tmpIx As Long
    tmpIx = lstCuentas.ListIndex
    
    Dim CTA As Long
    Dim SP() As String
    SP = Split(lstCuentas, ".")
    CTA = CLng(SP(1))
    Dim Ctas() As String
    Ctas = GetHijoTipo(CLng(CTA))
    If UBound(Ctas) = 0 Then
        MsgBox "¡No tiene subcuentas!"
        Exit Sub
    End If
    
    'veo si ya mostro las subcuentas, si es asi las escondo
    If Right(lstCuentas, 1) = "*" Then
        'primero le borro el asterisco
        lstCuentas.List(tmpIx) = Left(lstCuentas.List(tmpIx), _
            Len(lstCuentas.List(tmpIx)) - 1)
        
        'escondo las subcuentas
        Dim Niv As Long, I As Long
        Niv = nNivel(tmpIx)
        
        I = tmpIx + 1
        
        Do While Not Niv >= nNivel(I)
            lstCuentas.RemoveItem I
            
            If I > lstCuentas.ListCount - 1 Then Exit Do
            'se baja por la eliminacion
            'I = I + 1
        Loop
        
        Exit Sub
    End If
    
    Dim A As Long, TipoP As String
    
    For A = 1 To UBound(Ctas)
        TipoP = DB.GetValInRS("TipoProductos", "TipoProducto", "ID2 = " + Ctas(A))
        'ponerle los mismos espacios que tenia mas 3
        lstCuentas.AddItem "   " + SP(0) + "." + Ctas(A) + "." + _
            TipoP, lstCuentas.ListIndex + 1
        
    Next A
    
    'marco que ya lo abrio
    lstCuentas.List(tmpIx) = lstCuentas.List(tmpIx) + "*"
End Sub

Private Function GetHijoTipo(IdTipoProducto As Long) As String()
    Dim Resp() As String, Ix As Long
    Dim RSh As New ADODB.Recordset
    '*************************************************
    'LO PUSE YO ANDRES, EL RECORDCOUNT ANDA MAL SIN ESTO!!!!
    RSh.CursorLocation = adUseClient
    '*************************************************
    Ix = 0
    ReDim Preserve Resp(Ix)
    Resp(Ix) = "NADA"
    
    If RSh.State = adStateOpen Then RSh.Close
    RSh.Open "SELECT ID2 FROM TipoProductos WHERE IdAnt = " + CStr(IdTipoProducto), DB.CN, adOpenStatic, adLockReadOnly
    
    If RSh.RecordCount > 0 Then
        RSh.MoveFirst
        Do While Not RSh.EOF
            Ix = Ix + 1
            ReDim Preserve Resp(Ix)
            Resp(Ix) = CStr(NoNuloN(RSh("ID2")))
            RSh.MoveNext
        Loop
    End If
    
    RSh.Close
    Set RSh = Nothing
    
    GetHijoTipo = Resp
End Function

Private Function nNivel(IndiceLista As Long) As Long
    'veo el nivel de la cuenta seleccionada del listbox

    If IndiceLista = -1 Then
        nNivel = -1
        Exit Function
    End If
    
    Dim Spp() As String
    Spp = Split(lstCuentas.List(IndiceLista), ".")
    'spp(0) tiene los espacios que me van a decir en que nivel esta
    If Len(Spp(0)) = 0 Then  'es el nivel1
        nNivel = 1
    Else 'hago la formula, tiene que dar un numero redondo
        nNivel = Round(Len(Spp(0)) / 3 + 1, 0)
    End If
    
End Function

Private Sub txtMargenVenta_GotFocus()
    PintarTxt txtMargenVenta
End Sub

Private Sub txtMargenVenta_LostFocus()
    'no se porque pero le tengo que sacar el %
    Dim SP() As String
    SP = Split(txtMargenVenta, "%")
    txtMargenVenta = SP(0)
    txtMargenVenta = FormatPercent(ValidarNumeros(txtMargenVenta) / 100)
End Sub

Private Function GetID() As Long
    If lstCuentas.ListIndex = -1 Then Exit Function
    
    Dim SP() As String
    
    SP = Split(lstCuentas, ".")
    
    GetID = CLng(SP(1))
End Function
