VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrintOfertas 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ofertas a publicar"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   10275
   Icon            =   "frmPrintOfertas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   10275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbOfertas 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6870
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   360
      Width           =   3015
   End
   Begin VB.ComboBox cmbTIPOS 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   810
      Width           =   9525
   End
   Begin tbrFaroButton.fBoton cmdMAR 
      Height          =   435
      Left            =   570
      TabIndex        =   0
      Top             =   3555
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
   Begin MSComctlLib.ListView lvProd 
      Height          =   1995
      Left            =   360
      TabIndex        =   2
      Top             =   1260
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
      ForeColor       =   16777215
      BackColor       =   12632256
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Left            =   6090
      TabIndex        =   3
      Top             =   1260
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
      ForeColor       =   0
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Left            =   4980
      TabIndex        =   4
      Top             =   1800
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
      Left            =   7620
      TabIndex        =   5
      Top             =   3555
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Imprimir Ofertas"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdLimpiar 
      Height          =   435
      Left            =   6630
      TabIndex        =   6
      Top             =   3555
      Width           =   855
      _ExtentX        =   1508
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
      Left            =   2700
      TabIndex        =   7
      Top             =   3555
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
      Left            =   4980
      TabIndex        =   8
      Top             =   2430
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
   Begin tbrFaroButton.fBoton cmdSalir 
      Height          =   435
      Left            =   5100
      TabIndex        =   11
      Top             =   3570
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ofertas preexistentes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   405
      Left            =   5190
      TabIndex        =   10
      Top             =   300
      Width           =   1635
   End
End
Attribute VB_Name = "frmPrintOfertas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FSO As New Scripting.FileSystemObject

Private Sub cmbOfertas_Click()
    lvProd2.ListItems.Clear
    If cmbOfertas = "NUEVA" Then Exit Sub
    
    Dim F As String
    F = AP + cmbOfertas
    
    Dim TE As TextStream, TmP As Long, SP() As String, J As Long
    Set TE = FSO.OpenTextFile(F, ForReading, False)
        
        SP = Split(TE.ReadAll, Chr(5))
        
        For J = 0 To UBound(SP)
            TmP = lvProd2.ListItems.Count + 1
            lvProd2.ListItems.Add TmP
            lvProd2.ListItems(TmP).Text = SP(J)
            lvProd2.ListItems(TmP).SubItems(1) = GetNameProd(CLng(SP(J)))
        Next J
    TE.Close
End Sub

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
    
    cmbOfertas = "NUEVA" 'CAMBIO
    If lvProd.ListItems.Count = 0 Then Exit Sub
    
    Dim X As Long, Y As Long, IDp As String, Esta As Boolean, TmP As Long
    
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
                TmP = lvProd2.ListItems.Count + 1
                lvProd2.ListItems.Add TmP
                lvProd2.ListItems(TmP).Text = IDp
                lvProd2.ListItems(TmP).SubItems(1) = lvProd.ListItems(X).SubItems(1)
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
    If lvProd2.ListItems.Count = 0 Then
        MsgBox "No Hay Productos Seleccionados", vbInformation, "Atención"
        Exit Sub
    End If
    
    Dim OfertaSel As String

    'grabar esta oferta si es nueva
    If cmbOfertas = "NUEVA" Then
    
        Dim TE As TextStream
        
        Dim Nar As String
        Nar = AP + CStr(Year(Date)) + "-" + _
                 CStr(Month(Date)) + "-" + _
                 CStr(Day(Date)) + "-" + _
                 CStr(Hour(Time)) + "-" + _
                 CStr(Minute(Time)) + ".OFERTA"
        
        If FSO.FileExists(Nar) Then FSO.DeleteFile Nar, True
        
        Set TE = FSO.CreateTextFile(Nar, True)
        
        Dim J As Long, IDp As Long
        For J = 1 To lvProd2.ListItems.Count
            'leer lo de la base de datos
            IDp = lvProd2.ListItems(J).Text
            
            TE.Write CStr(IDp)
            'para que al leer no tener uno vacio al ultimo en el split
            If J < lvProd2.ListItems.Count Then TE.Write Chr(5)
            
        Next J
        
        TE.Close
        OfertaSel = Nar
    Else
        OfertaSel = AP + cmbOfertas
    End If
    
    frmPreviewOferta.VER OfertaSel
    frmPreviewOferta.Show 1
    
End Sub

Private Function GetNameProd(idProd As Long) As String
    Dim rS As New ADODB.Recordset
    rS.CursorLocation = adUseClient
    rS.Open "Select * from productos where id=" + CStr(idProd), DB.CN, adOpenStatic, adLockReadOnly
    
    If rS.RecordCount <> 1 Then
        GetNameProd = ""
        Exit Function
    End If
    
    Dim TmP As String
    If IsNull(rS.Fields("nProducto")) Then
        TmP = ""
    Else
        TmP = rS.Fields("nProducto")
    End If
    
    GetNameProd = TmP
End Function

Private Function GetTXTProd(IDDP As Long) As String  'le doy el codigo delproducto solo
    Dim rS As New ADODB.Recordset
    rS.CursorLocation = adUseClient
    rS.Open "Select * from productos where id=" + CStr(IDDP), DB.CN, adOpenStatic, adLockReadOnly
    
    If rS.RecordCount <> 1 Then
        GetTXTProd = ""
        Exit Function
    End If
    
    Dim TmP As String
    Dim F As ADODB.Field
    Dim T As String
    For Each F In rS.Fields
        If IsNull(F.Value) Then
            T = ""
        Else
            T = CStr(F.Value)
        End If
        
        TmP = TmP + F.Name + Chr(5) + T + Chr(6)
    Next
    GetTXTProd = TmP
End Function

Private Sub cmdLimpiar_Click()
    cmbOfertas = "NUEVA" 'CAMBIO
    lvProd2.ListItems.Clear
End Sub

Private Sub cmdMAR_Click()
    Dim I As Long
    
    For I = 1 To lvProd.ListItems.Count
        lvProd.ListItems(I).Checked = True
    Next I
End Sub

Private Sub cmdQuitar_Click()
    cmbOfertas = "NUEVA" 'CAMBIO
    
    If lvProd2.ListItems.Count = 0 Then Exit Sub
    Dim H As Long
    
    For H = lvProd2.ListItems.Count To 1 Step -1
        If lvProd2.ListItems(H).Checked Then lvProd2.ListItems.Remove (H)
    Next H
End Sub

Private Sub cmdSalir_Click()
    Unload Me
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
    
    Dim FI As Scripting.File
    Dim FO As Scripting.Folder
    
    
    cmbOfertas.Clear
    cmbOfertas.AddItem "NUEVA"
    Set FO = FSO.GetFolder(AP)
    For Each FI In FO.Files
        If LCase(FSO.GetExtensionName(FI)) = "oferta" Then cmbOfertas.AddItem FI.Name
    Next
    
    cmbOfertas.ListIndex = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
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



