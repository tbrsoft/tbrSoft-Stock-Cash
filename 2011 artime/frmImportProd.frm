VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmImportProd 
   BackColor       =   &H00544B45&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importar Productos"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9885
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmImportProd.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin tbrFaroButton.fBoton command1 
      Height          =   465
      Left            =   3810
      TabIndex        =   6
      Top             =   5980
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Aceptar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdDesmar 
      Height          =   435
      Left            =   4500
      TabIndex        =   5
      Top             =   4380
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Desmarcar Todos"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdMAR 
      Height          =   435
      Left            =   2370
      TabIndex        =   4
      Top             =   4380
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Marcar Todos"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdEXPORT 
      Height          =   405
      Left            =   2730
      TabIndex        =   3
      Top             =   360
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Definir Ubicación e Productos a Importar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.ComboBox cmbTIPOS 
      Height          =   315
      Left            =   690
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   5400
      Width           =   8565
   End
   Begin MSComctlLib.ListView lvProd 
      Height          =   3285
      Left            =   180
      TabIndex        =   1
      Top             =   930
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   5794
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "nProducto"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "pCosto"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "pVenta"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "CodEnvase"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Observaciones"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "CodigodeBarras"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "BUSCAR"
         Object.Width           =   2540
      EndProperty
   End
   Begin tbrFaroButton.fBoton Command2 
      Height          =   465
      Left            =   8520
      TabIndex        =   7
      Top             =   5980
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   820
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Importar ahora los marcados al tipo:"
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   750
      TabIndex        =   2
      Top             =   5070
      Width           =   3435
   End
End
Attribute VB_Name = "frmImportProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDesmar_Click()
    Dim I As Long
    
    For I = 1 To lvProd.ListItems.Count
        lvProd.ListItems(I).Checked = False
    Next I
End Sub

Private Sub cmdEXPORT_Click()

    'tener en cuenta el tipo de producto cambiara!!!
    Dim CM As New CommonDialog
    CM.Filter = "Archivos de Exportacion Stock & Cash (*.ESC)|*.ESC"
    
    CM.ShowOpen
    Dim F As String
    F = CM.FileName
    If F = "" Then
        MsgBox "¡No se ha definido la carpeta!", vbInformation, "Atención"
        Exit Sub
    End If
    
    'ver que tenga archivos para importar!
    Dim FS As New Scripting.FileSystemObject
    Dim Fols As Scripting.Folder
    Dim Fils As Scripting.File
    
    Dim ContFIL As Long 'cuantos archivos.FIL hay
    ContFIL = 0
    lvProd.ListItems.Clear
    
    'ir abriendolos con JUSE
    Dim JS As New tbrJUSE.clsJUSE
    JS.ReadFile F
    Dim FolTMP As String
    Randomize
    FolTMP = AP + "exp" + CStr(Int(Rnd * 2000))
    If FS.FolderExists(FolTMP) = False Then FS.CreateFolder FolTMP
    
    Dim H As Long
    For H = 1 To JS.CantArchs
        JS.Extract FolTMP, H
    Next H
    
    Set Fols = FS.GetFolder(FolTMP)
    
    Dim J As Long
    For Each Fils In Fols.Files
        'los mios tiene extencion "FIL"
        If LCase(Right(Fils.path, 3)) = "fil" Then
            Dim EXTR As String
            EXTR = Left(Fils.path, Len(Fils.path) - 4) '+ "\" NO PERMITE BORRAR CON LA BARRA AL FINAL!!!
            If FS.FolderExists(EXTR) Then FS.DeleteFolder EXTR, True
            FS.CreateFolder EXTR
            
            'saco todo lo que hay adentro
            ContFIL = ContFIL + 1
            JS.ReadFile Fils.path
            For J = 1 To JS.CantArchs
                JS.Extract EXTR, J
            Next J
            
            'dentro de cada carpeta hay un archivo con el mismo numero que la
            'carpeta, necesito ese numero
            Dim SP() As String, SP2() As String
            Dim nInterno As String
            SP = Split(EXTR, "_")
            nInterno = SP(UBound(SP) - 1)
            
            'leer el archivo de datos que si o si tiene que estar
            If FS.FileExists(EXTR + "\EX_" + nInterno + ".DB1") Then
                Dim TT As TextStream, TES As String
                Set TT = FS.OpenTextFile(EXTR + "\EX_" + nInterno + ".DB1", ForReading, False)
                    TES = TT.ReadAll
                TT.Close
                
                
                SP = Split(TES, Chr(6)) 'cada uno de los renglones
                'meterlo en el listview
                'quitar el tipo anterior por que todavia no se definio
                lvProd.ListItems.Add ContFIL
                'ID NO VA
                'IdTipoProducto no va
                'nProducto
                SP2 = Split(SP(2), Chr(5))
                lvProd.ListItems(ContFIL).Text = SP2(1)
                'pCosto
                SP2 = Split(SP(3), Chr(5))
                lvProd.ListItems(ContFIL).SubItems(1) = SP2(1)
                'pVenta
                SP2 = Split(SP(4), Chr(5))
                lvProd.ListItems(ContFIL).SubItems(2) = SP2(1)
                'Stock no va
                'CodEnvase
                SP2 = Split(SP(6), Chr(5))
                lvProd.ListItems(ContFIL).SubItems(3) = SP2(1)
                'Observaciones
                SP2 = Split(SP(7), Chr(5))
                lvProd.ListItems(ContFIL).SubItems(4) = SP2(1)
                'CodigodeBarras
                SP2 = Split(SP(8), Chr(5))
                lvProd.ListItems(ContFIL).SubItems(5) = SP2(1)
                'PARA LEER SI IMPORTA Y BUSCAR IMAGENES Y DESCRIPCIONES
                lvProd.ListItems(ContFIL).SubItems(6) = EXTR
            End If
        End If
    Next
    
    Set JS = Nothing
    
    If ContFIL <= 0 Then
        MsgBox "No se han encontrado archivos", vbInformation, "Atención"
    Else
        MsgBox "Se han encontrado " + CStr(ContFIL) + " archivos", vbInformation, "Archivos encontrados"
    End If
End Sub

Private Sub cmdMAR_Click()
    Dim I As Long
    
    For I = 1 To lvProd.ListItems.Count
        lvProd.ListItems(I).Checked = True
    Next I
End Sub

Private Sub Command1_Click()
    If lvProd.ListItems.Count = 0 Then Exit Sub

    'ver el codigo del tipo al que voy
    Dim SP() As String
    SP = Split(cmbTIPOS, "|")
    Dim TIPO As Long
    TIPO = CLng(Trim(SP(0)))
    
    Dim FS As New Scripting.FileSystemObject
    
    'ir definiendo los IDs para poder escribir las imagenes y las descripciones
    Dim J As Long
    Dim NewID As Long
    Dim SSS As String
    For J = 1 To lvProd.ListItems.Count
        If lvProd.ListItems(J).Checked Then
            'ver si ya existe el nombre del producto
            If DB.ContarReg("select nproducto from productos where nproducto= '" + _
                lvProd.ListItems(J).Text + "'") > 0 Then
                MsgBox "¡Ya tiene un producto con el nombre !" + _
                    vbCrLf + lvProd.ListItems(J).Text + vbCrLf + _
                    "No se agregara", _
                    vbExclamation, "Atención"
        
                GoTo SSIIGG
            End If
            
            'obtener un nuevo id
            NewID = IdAutonum("Productos")
            
            'meter en la base de datos como corresponde
            'por que el artime mete como string el idTipoProd
            SSS = "INSERT INTO Productos (ID,IdTipoProducto,Nproducto," + _
                    "pCosto,pVenta,CodEnvase,Observaciones,CodigodeBarras) VALUES (" + _
                    CStr(NewID) + "," + _
                    CStr(TIPO) + ",'" + _
                    lvProd.ListItems(J).Text + "'," + _
                    Replace(lvProd.ListItems(J).SubItems(1), ",", ".") + "," + _
                    Replace(lvProd.ListItems(J).SubItems(2), ",", ".") + ",'" + _
                    lvProd.ListItems(J).SubItems(3) + "','" + _
                    lvProd.ListItems(J).SubItems(4) + "','" + _
                    lvProd.ListItems(J).SubItems(5) + "')"
            
            DB.EXECUTE SSS
            
            'buscar externos
            Dim BUSCAREN As String
            'ya tiene la barra al final
            BUSCAREN = lvProd.ListItems(J).SubItems(6)
            'ver si hay imagenes
            Dim Ars() As String
            Ars = ObtenerArch(BUSCAREN, "*.jpg")
            If UBound(Ars) = 0 Then GoTo DESCs
            'hay al menos una imagen
            Dim H As Long, SSP() As String, Dest As String
            Dim fOrig As String
            For H = 1 To UBound(Ars)
                fOrig = FS.GetParentFolderName(Ars(H))
                Dest = FS.GetBaseName(Ars(H)) 'ya se la extencion!
                SSP = Split(Dest, "-")
                Dest = CStr(NewID) + "-" + SSP(1) + ".jpg" 'respeto el numero que habia atras
                Dest = CFGBD.GetInfo(82, 4) + "img\" + Dest
                If FS.FileExists(Dest) Then FS.DeleteFile Dest, True
                FS.MoveFile Ars(H), Dest
            Next H
            'ver si hay descripcion
DESCs:
            Dim Ars2() As String
            Ars2 = ObtenerArch(BUSCAREN, "*.txt")
            If UBound(Ars2) = 0 Then GoTo SSIIGG
            'hay al menos una imagen
            Dim H2 As Long, SSP2() As String, Dest2 As String
            Dim fOrig2 As String
            For H2 = 1 To UBound(Ars2)
                fOrig2 = FS.GetParentFolderName(Ars2(H2))
                Dest2 = FS.GetBaseName(Ars2(H2)) 'ya se la extencion!
                SSP2 = Split(Dest2, ".")
                Dest2 = CStr(NewID) + ".txt"
                Dest2 = CFGBD.GetInfo(82, 4) + "img\" + Dest2
                If FS.FileExists(Dest2) Then FS.DeleteFile Dest2, True
                FS.MoveFile Ars2(H2), Dest2
            Next H2
            
SSIIGG:
        End If
    Next J
    
    MsgBox "Se completó con éxito la importación", vbInformation, "Importación terminada"
    lvProd.ListItems.Clear
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Dim X As Long, ClsP As New clsProducto
    Dim Mat() As String
    
    'si o si tiene mas de un renglon
    Mat = ClsP.GetHijoTipo(0)
    cmbTIPOS.Clear
    'cmbTIPOS.AddItem "TODOS"
    
    For X = 1 To UBound(Mat)
        cmbTIPOS.AddItem Mat(X) + " | " + _
            DB.GetValInRS("TipoProductos", "TipoProducto", "ID2 = " + Mat(X))
        AgregarHijos cmbTIPOS.ListCount - 1, 1
    Next X
    
    Set ClsP = Nothing
    cmbTIPOS.ListIndex = 0
End Sub

Private Sub AgregarHijos(Renglon As Long, Nivel As Long)
    Dim Y As Long, Mat() As String, SP() As String
    Dim ClsP As New clsProducto, Reng As Long, Niv As Long
    
    SP = Split(cmbTIPOS.List(Renglon), " | ")
    Mat = ClsP.GetHijoTipo(CLng(Trim(SP(0))))
    Reng = Renglon
    Niv = Nivel + 1
    
    If UBound(Mat) > 0 Then
        For Y = 1 To UBound(Mat)
            Reng = Reng + 1
            cmbTIPOS.AddItem String(Niv * 3, " ") + Mat(Y) + " | " + _
                DB.GetValInRS("TipoProductos", "TipoProducto", "ID2 = " + Mat(Y)), Reng
            AgregarHijos Reng, Niv + 1
        Next Y
    End If
    
    Set ClsP = Nothing
End Sub

