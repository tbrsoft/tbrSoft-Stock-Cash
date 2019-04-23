VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExportProd 
   BackColor       =   &H00544B45&
   Caption         =   "Exportar Productos"
   ClientHeight    =   3885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10485
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmExportProd.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3885
   ScaleWidth      =   10485
   StartUpPosition =   2  'CenterScreen
   Begin tbrFaroButton.fBoton cmdMAR 
      Height          =   435
      Left            =   720
      TabIndex        =   4
      Top             =   3180
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Marcar Todos"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.PictureBox pBAR 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   105
      Left            =   7950
      ScaleHeight     =   45
      ScaleWidth      =   1635
      TabIndex        =   3
      Top             =   3630
      Width           =   1695
   End
   Begin VB.ComboBox cmbTIPOS 
      Height          =   315
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   90
      Width           =   9765
   End
   Begin MSComctlLib.ListView lvProd 
      Height          =   1995
      Left            =   210
      TabIndex        =   1
      Top             =   900
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   3519
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Producto"
         Object.Width           =   5292
      EndProperty
   End
   Begin MSComctlLib.ListView lvProd2 
      Height          =   1995
      Left            =   6000
      TabIndex        =   2
      Top             =   900
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   3519
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Producto"
         Object.Width           =   3528
      EndProperty
   End
   Begin tbrFaroButton.fBoton cmdADD 
      Height          =   375
      Left            =   4890
      TabIndex        =   5
      Top             =   1470
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   ">>"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdEXPORT 
      Height          =   435
      Left            =   7860
      TabIndex        =   6
      Top             =   3180
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Exportar Esta Lista"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdLimpiar 
      Height          =   435
      Left            =   5700
      TabIndex        =   7
      Top             =   3180
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Limpiar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdDesmar 
      Height          =   435
      Left            =   2850
      TabIndex        =   8
      Top             =   3180
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Desmarcar Todos"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdQuitar 
      Height          =   375
      Left            =   4890
      TabIndex        =   9
      Top             =   2100
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "<<"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
End
Attribute VB_Name = "frmExportProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbTipos_Click()
    If cmbTIPOS.ListIndex < 0 Then Exit Sub
    
    If cmbTIPOS = "TODOS" Then
        CargarComboLV lvProd, "SELECT ID, nProducto FROM Productos WHERE ID >= 0", "ID/n,nProducto"
    Else
        Dim IdT As String, SP() As String
        
        SP = Split(cmbTIPOS, " | ")
        IdT = Trim(SP(0))
        
        CargarComboLV lvProd, "SELECT ID, nProducto FROM Productos " + _
            "WHERE IDTipoProducto = " + IdT, "ID/n,nProducto"
    End If
End Sub

Private Sub cmdAdd_Click()
    If lvProd.ListItems.Count = 0 Then Exit Sub
    
    Dim X As Long, Y As Long, IDp As String, Esta As Boolean, tmP As Long
    
    Esta = False
    For X = 1 To lvProd.ListItems.Count
        If lvProd.ListItems(X).Checked = True Then
            'veo si ya esta
            IDp = lvProd.ListItems(X).Text
            For Y = 1 To lvProd2.ListItems.Count
                If IDp = lvProd2.ListItems(Y).Text Then
                    Esta = True
                    Exit For
                End If
            Next Y
            If Esta = False Then
                tmP = lvProd2.ListItems.Count + 1
                lvProd2.ListItems.Add tmP
                lvProd2.ListItems(tmP).Text = IDp
                lvProd2.ListItems(tmP).SubItems(1) = lvProd.ListItems(X).SubItems(1)
            Else
                Esta = False
            End If
        End If
    Next X
    
    cmdDesmar_Click
End Sub

Private Sub cmdDesmar_Click()
Dim I As Long
    
    For I = 1 To lvProd.ListItems.Count
        lvProd.ListItems(I).Checked = False
    Next I
End Sub

Private Sub cmdEXPORT_Click()
    If lvProd2.ListItems.Count = 0 Then Exit Sub

    Dim J As Long, H As Long
    Dim IDp As Long
    Dim ProdDB As String
    Dim FS1 As New Scripting.FileSystemObject
    Dim AA As String
    
    'ver donde los va a meter
    Dim CM As New CommonDialog
    CM.InitDir = AP
    CM.ShowFolder
    
    AA = CM.InitDir
    If AA = "" Or AA = AP Then
        MsgBox "Se ha cancelado", vbInformation, "Atención"
        Exit Sub
    End If
    
    Dim IB As String
    IB = InputBox("Defina un nombre para este paquete de productos exportados", , "No lo deje en blanco")
    'por las dudas que ponga caracteresno validos!
    IB = Replace(IB, "?", "")
    IB = Replace(IB, "\", "")
    IB = Replace(IB, "|", "")
    IB = Replace(IB, "/", "")
    IB = Replace(IB, "¡", "")
    IB = Replace(IB, "!", "")
    IB = Replace(IB, "¿", "")
    
    If Right(AA, 1) <> "\" Then AA = AA + "\"
    
    'ARCHIVO COMPLETO CON TODOS LOS PEQUEÑOS JUSE DE CADA PRODUCTO
    Dim JS2 As New tbrJUSE.clsJUSE
    JS2.Archivo = AA + "EXPORT_" + IB + "_.ESC"
        
    'tengo esta carpeta para cada uno de los archivos
    AA = AA + "EXPORT\"
        
    'lo dejo limpito
    If FS1.FolderExists(AA) Then
        If MsgBox("Ya hay una exportacion aqui!" + vbCrLf + _
            "Desea eliminarla ?", vbCritical + vbYesNo, "Atención") = vbNo Then Exit Sub
            
        FS1.DeleteFolder AA, True
        
    End If
    
    If FS1.FolderExists(AA) = False Then FS1.CreateFolder AA
    
    Dim Fol1 As String 'carpeta base de la exportacion
    Fol1 = AA
    
    For J = 1 To lvProd2.ListItems.Count
        pBAR.Width = (J / lvProd2.ListItems.Count) * cmdEXPORT.Width
        'si no lo mato al final se me acumulan todos al final, nunca selimpia el paquete!
        Dim JS As New tbrJUSE.clsJUSE
        
        AA = Fol1 + "EXPORT_" + CStr(J) + "_.FIL" 'para hacer un split por "_"
        
        JS.Archivo = AA
        
        'leer lo de la base de datos
        IDp = lvProd2.ListItems(J).Text
        Dim TTE As TextStream
        AA = Fol1 + "\EX_" + CStr(J) + ".db1"
        Set TTE = FS1.CreateTextFile(AA, True)
            ProdDB = GetTXTProd(IDp)
            TTE.Write ProdDB
        TTE.Close
        
        JS.AddFile AA
        
        'FS1.DeleteFile AA, True
        'leer los archivos de descripcion
        AA = CFGBD.GetInfo(82, 4) + "img\" + CStr(IDp) + ".txt"
        If FS1.FileExists(AA) Then JS.AddFile AA
        
        'leer las imagenes
        For H = 0 To 20
            AA = CFGBD.GetInfo(82, 4) + "img\" + CStr(IDp) + "-" + CStr(H) + ".jpg"
            If FS1.FileExists(AA) Then JS.AddFile AA
            
            AA = CFGBD.GetInfo(82, 4) + "img\" + CStr(IDp) + "-" + CStr(H) + ".gif"
            If FS1.FileExists(AA) Then JS.AddFile AA
            
            AA = CFGBD.GetInfo(82, 4) + "img\" + CStr(IDp) + "-" + CStr(H) + ".bmp"
            If FS1.FileExists(AA) Then JS.AddFile AA
        Next H
        
        'JUNTAR TODO ESTE PRODUCTO
        JS.Unir False
        Set JS = Nothing
        
        'agrego este paquetito al paquete general
        JS2.AddFile Fol1 + "EXPORT_" + CStr(J) + "_.FIL"
            
    Next J
    
    'meter todos ahora en un solo archivo Export
    JS2.Unir False
    
    'borro todo estos paquetes(delefolderno puede tener barra al final
    FS1.DeleteFolder Left(Fol1, Len(Fol1) - 1)
    
    'si llego aca puedo borrar la carpeta con todos los sueltos chiquitos
    
    MsgBox "Se ha exportado con exito" + vbCrLf + _
        "El archivo se encuentra en:" + vbCrLf + _
        JS2.Archivo, vbInformation, "Exportación exitosa"
        
    lvProd2.ListItems.Clear
End Sub

Private Function GetTXTProd(IDDP As Long) As String  'le doy el codigo delproducto solo
    Dim rS As New ADODB.Recordset
    rS.CursorLocation = adUseClient
    rS.Open "Select * from productos where id=" + CStr(IDDP), DB.CN, adOpenStatic, adLockReadOnly
    
    If rS.RecordCount <> 1 Then
        GetTXTProd = ""
        Exit Function
    End If
    
    Dim tmP As String
    Dim F As ADODB.Field
    Dim T As String
    For Each F In rS.Fields
        If IsNull(F.Value) Then
            T = ""
        Else
            T = CStr(F.Value)
        End If
        
        tmP = tmP + F.Name + Chr(5) + T + Chr(6)
    Next
    GetTXTProd = tmP
End Function

Private Sub cmdLimpiar_Click()
    lvProd2.ListItems.Clear
End Sub

Private Sub cmdMAR_Click()
    Dim I As Long
    
    For I = 1 To lvProd.ListItems.Count
        lvProd.ListItems(I).Checked = True
    Next I
End Sub

Private Sub cmdQuitar_Click()
    If lvProd2.ListItems.Count = 0 Then Exit Sub
    Dim H As Long
    
    For H = lvProd2.ListItems.Count To 1 Step -1
        If lvProd2.ListItems(H).Checked Then lvProd2.ListItems.Remove (H)
    Next H
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
    cmbTIPOS.AddItem "TODOS"
    
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

Private Sub Form_Resize()
    If Me.Width < 7000 Or Me.Height < 5000 Then Exit Sub
    
    On Local Error Resume Next
    lvProd.Top = cmbTIPOS.Top + cmbTIPOS.Height + 120
    lvProd2.Top = lvProd.Top
    
    lvProd.Height = Me.Height - lvProd.Top - cmdEXPORT.Height - 600
    lvProd2.Height = lvProd.Height
    
    lvProd.Width = (Me.Width / 2) - cmdADD.Width - 200
    lvProd2.Width = lvProd.Width
    
    lvProd.Left = 60
    lvProd2.Left = Me.Width - lvProd2.Width - 260
    
    cmdMAR.Top = lvProd.Top + lvProd.Height + 60
    cmdDesmar.Top = cmdMAR.Top
    cmdEXPORT.Top = cmdMAR.Top
    cmdLimpiar.Top = cmdMAR.Top
    cmdADD.Top = lvProd.Left + lvProd.Height / 2
    cmdQuitar.Top = cmdADD.Top + 600
    
    cmdMAR.Left = 60
    cmdDesmar.Left = cmdMAR.Left + cmdMAR.Width + 60
    cmdADD.Left = lvProd.Left + lvProd.Width + 60
    cmdQuitar.Left = cmdADD.Left
    cmdEXPORT.Left = Me.Width - cmdEXPORT.Width - 260
    cmdLimpiar.Left = cmdEXPORT.Left - 1500
    
    pBAR.Left = cmdEXPORT.Left
    pBAR.Top = cmdEXPORT.Top + cmdEXPORT.Height + 30
    pBAR.Width = cmdEXPORT.Width
End Sub

