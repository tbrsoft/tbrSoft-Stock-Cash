VERSION 5.00
Object = "{A7FBD38D-2930-49E3-B60C-9E0202D84549}#15.0#0"; "tbrControles.ocx"
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClientes2 
   BackColor       =   &H0049453D&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de Clientes"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12015
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmClientes2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   12015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H0049453D&
      Caption         =   "Filtros"
      ForeColor       =   &H00FFFFFF&
      Height          =   1665
      Left            =   330
      TabIndex        =   10
      Top             =   180
      Width           =   11595
      Begin VB.CheckBox chkFin 
         BackColor       =   &H0049453D&
         Caption         =   "Compró con Financiera"
         ForeColor       =   &H00E0E0E0&
         Height          =   350
         Left            =   660
         TabIndex        =   8
         Top             =   1200
         Width           =   2415
      End
      Begin VB.CheckBox chkTrabaja 
         BackColor       =   &H0049453D&
         Caption         =   "Tiene Trabajo"
         ForeColor       =   &H00E0E0E0&
         Height          =   350
         Left            =   6570
         TabIndex        =   6
         Top             =   720
         Width           =   1605
      End
      Begin VB.TextBox txtCondIVA 
         Height          =   360
         Left            =   1440
         TabIndex        =   4
         Top             =   730
         Width           =   1815
      End
      Begin VB.TextBox txtDireccion 
         Height          =   350
         Left            =   4140
         TabIndex        =   1
         Top             =   270
         Width           =   2025
      End
      Begin VB.TextBox txtNAME 
         Height          =   360
         Left            =   1440
         TabIndex        =   0
         Top             =   270
         Width           =   1815
      End
      Begin tbrControles.MouTextBox txtEdad1 
         Height          =   350
         Left            =   7170
         TabIndex        =   2
         Top             =   270
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   609
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
         Largo           =   3
         Entero          =   -1  'True
      End
      Begin tbrControles.MouTextBox txtEdad2 
         Height          =   350
         Left            =   8160
         TabIndex        =   3
         Top             =   270
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   609
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
         Largo           =   3
         Entero          =   -1  'True
      End
      Begin tbrControles.MouTextBox txtIngreso 
         Height          =   345
         Left            =   10230
         TabIndex        =   7
         Top             =   930
         Visible         =   0   'False
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   609
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
      Begin tbrControles.MouTextBox txtLimite 
         Height          =   345
         Left            =   5310
         TabIndex        =   5
         Top             =   705
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   609
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
      Begin tbrControles.MouTextBox txtDiasFin 
         Height          =   345
         Left            =   6720
         TabIndex        =   9
         Top             =   1230
         Visible         =   0   'False
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   609
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
         Largo           =   4
         Entero          =   -1  'True
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Días sin comprar por Financiera mayor a"
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   2190
         TabIndex        =   20
         Top             =   1260
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "días"
         ForeColor       =   &H00E0E0E0&
         Height          =   345
         Left            =   7830
         TabIndex        =   21
         Top             =   1290
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Límite Crédito mayor a"
         ForeColor       =   &H00E0E0E0&
         Height          =   345
         Left            =   3210
         TabIndex        =   19
         Top             =   780
         Width           =   2085
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ing. Mensual Mayor a"
         ForeColor       =   &H00E0E0E0&
         Height          =   345
         Left            =   8250
         TabIndex        =   17
         Top             =   960
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Condición IVA"
         ForeColor       =   &H00E0E0E0&
         Height          =   345
         Left            =   150
         TabIndex        =   16
         Top             =   780
         Width           =   1395
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "años"
         ForeColor       =   &H00E0E0E0&
         Height          =   345
         Left            =   8940
         TabIndex        =   15
         Top             =   300
         Width           =   555
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "y"
         ForeColor       =   &H00E0E0E0&
         Height          =   345
         Left            =   7800
         TabIndex        =   14
         Top             =   330
         Width           =   225
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Edad entre"
         ForeColor       =   &H00E0E0E0&
         Height          =   345
         Left            =   5820
         TabIndex        =   13
         Top             =   330
         Width           =   1305
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre/Codigo"
         ForeColor       =   &H00E0E0E0&
         Height          =   345
         Left            =   60
         TabIndex        =   12
         Top             =   330
         Width           =   1485
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección"
         ForeColor       =   &H00E0E0E0&
         Height          =   345
         Left            =   3270
         TabIndex        =   11
         Top             =   330
         Width           =   825
      End
   End
   Begin MSComctlLib.ListView lvCli 
      Height          =   3495
      Left            =   60
      TabIndex        =   18
      Top             =   2100
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   6165
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   3440
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Direccion"
         Object.Width           =   4004
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "CUIT/DNI"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Telefono"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Fecha Nac."
         Object.Width           =   2028
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Ult.Cpra.Fciera."
         Object.Width           =   2028
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Financiera"
         Object.Width           =   2222
      EndProperty
   End
   Begin tbrFaroButton.fBoton cmdADD 
      Height          =   435
      Left            =   1470
      TabIndex        =   22
      Top             =   6030
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Agregar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdMOD 
      Height          =   435
      Left            =   2970
      TabIndex        =   23
      Top             =   6030
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Modificar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdKILL 
      Height          =   435
      Left            =   4470
      TabIndex        =   24
      Top             =   6030
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Borrar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdImprimir 
      Height          =   435
      Left            =   5940
      TabIndex        =   25
      Top             =   6030
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Imprimir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton Command2 
      Height          =   435
      Left            =   7440
      TabIndex        =   26
      Top             =   6030
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Quitar Filtros"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton Command1 
      Height          =   435
      Left            =   8940
      TabIndex        =   27
      Top             =   6030
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
End
Attribute VB_Name = "frmClientes2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UsoTrabaja As Boolean, UsoFinanciera As Boolean, CargaInicial As Boolean

Public Sub CargarDatos(NewSQL As String)
    If EstaTodoBlanco And NewSQL <> "SELECT * FROM Clientes WHERE ID>=0 ORDER BY ID" Then
        Limpiar False
    Else
        'cambia el SQL de los clientes
        '(1) filtro lo de access -------------------------------------------------------
        CargarComboLV lvCli, NewSQL, "ID/n,Nombre,Direccion,CUIT,Telefono,Nacimiento/f"
        
        '(2) Filtrar por configuracion
        FiltrarCFG
        
        '(3) Agregar Columna de Ult.Compra
        AgregarUltimaCompra
    End If
End Sub

Private Function EstaTodoBlanco() As Boolean
    Dim Resp As Boolean
    
    Resp = True
    
    If txtNAME + txtDireccion + txtCondIVA <> "" Then
        Resp = False
    Else
        If IsNumeric(txtLimite) And IsNumeric(txtEdad1) And IsNumeric(txtEdad2) Then
            If CSng(txtLimite) = 0 And CSng(txtEdad1) = 0 And CSng(txtEdad2) = 100 Then
                If txtIngreso.Visible = False And txtDiasFin.Visible = False Then
                    Resp = True
                Else
                    Resp = False
                End If
            Else
                Resp = False
            End If
        Else
            Resp = False
        End If
    End If
    
    EstaTodoBlanco = Resp
End Function

Private Sub FiltrarCFG()
    Dim TmP As String, PB As Long, R As Long, IdCF As Long, SP() As String
    Dim ParaBorrar() As String 'con el index del renglon que no entran en el filtro
    
    PB = 0
    ReDim ParaBorrar(PB)
    ParaBorrar(PB) = "Nada"
    ' LIMITE ------------------------------------------------------------------------
    If Not IsNumeric(txtLimite) Then txtLimite = "0"
    If CLng(txtLimite) > 0 Then 'hacer filtro solo si eligio uno mayor que 0
        For R = 1 To lvCli.ListItems.Count
            IdCF = CFG.ExistePropiedad("FDP " + lvCli.ListItems(R).Text)
            If IdCF <> 0 Then
                SP = Split(CFG.GetInfo(IdCF, 4), "_")
                If UBound(SP) = 2 Then
                    If IsNumeric(SP(2)) Then
                       If CSng(SP(2)) < CSng(txtLimite) Then
                          PB = PB + 1
                          ReDim Preserve ParaBorrar(PB)
                          ParaBorrar(PB) = CStr(R)
                       End If
                    Else
                        CFG.ModificarNodo IdCF, , , , "CC_30_0"
                    End If
                Else
                    CFG.ModificarNodo IdCF, , , , "CC_30_0"
                End If
            End If
        Next R
    End If
    ' -------------------------------------------------------------------------------
    
    ' INGRESO MENSUAL ---------------------------------------------------------------
    If txtIngreso.Visible = False Then txtIngreso = "0"
    If Not IsNumeric(txtIngreso) Then txtIngreso = "0"
    If CLng(txtIngreso) > 0 Then 'hacer filtro solo si eligio uno mayor que 0
        For R = 1 To lvCli.ListItems.Count
            If YaEstaBorrado(R, ParaBorrar) = False Then
                IdCF = CFG.ExistePropiedad("DTC " + lvCli.ListItems(R).Text)
                If IdCF <> 0 Then
                    SP = Split(CFG.GetInfo(IdCF, 4), "_")
                    If UBound(SP) = 2 Then
                        If IsNumeric(SP(2)) Then
                           If CSng(SP(2)) < CSng(txtIngreso) Then
                              PB = PB + 1
                              ReDim Preserve ParaBorrar(PB)
                              ParaBorrar(PB) = CStr(R)
                           End If
                        Else
                            CFG.ModificarNodo IdCF, , , , "1_Empleado Público___1000"
                        End If
                    Else
                        CFG.ModificarNodo IdCF, , , , "1_Empleado Público___1000"
                    End If
                End If
            End If
        Next R
    End If
    ' -------------------------------------------------------------------------------
    
    ' FINANCIERA --------------------------------------------------------------------
    If chkFin.Value <> 0 Then
        If Not IsNumeric(txtDiasFin) Then txtDiasFin = "0"
        For R = 1 To lvCli.ListItems.Count
            If YaEstaBorrado(R, ParaBorrar) = False Then
                IdCF = CFG.ExistePropiedad("UPF " + lvCli.ListItems(R).Text)
                If IdCF <> 0 Then
                    ' DIAS FINANCIERA --------------------------------------------------------------------
                    If CLng(txtDiasFin) > 0 Then
                        TmP = CFG.GetInfo(IdCF, 4)
                        If IsDate(TmP) Then
                            If DateDiff("d", CDate(TmP), Date) < CLng(txtDiasFin) Then
                                PB = PB + 1
                                ReDim Preserve ParaBorrar(PB)
                                ParaBorrar(PB) = CStr(R)
                            End If
                        Else
                            CFG.ModificarNodo IdCF, 70, , , CStr(Date)
                        End If
                    End If
                Else
                    'idcf es 0 lo que significa que no hizo nunca compras con
                    'financiera -> borro de la lista
                    If YaEstaBorrado(R, ParaBorrar) = False Then
                        PB = PB + 1
                        ReDim Preserve ParaBorrar(PB)
                        ParaBorrar(PB) = CStr(R)
                    End If
                End If
            End If
        Next R
        ' -------------------------------------------------------------------------------
    Else
        'si no uso financiera no pasa nada, pero si uso significa que ya hizo click
        'y quiere ver unicamente los que no compraron con financiera
        If UsoFinanciera = True Then
            For R = 1 To lvCli.ListItems.Count
                If YaEstaBorrado(R, ParaBorrar) = False Then
                    IdCF = CFG.ExistePropiedad("UPF " + lvCli.ListItems(R).Text)
                    If IdCF <> 0 Then 'compro alguna vez lo borro
                        PB = PB + 1
                        ReDim Preserve ParaBorrar(PB)
                        ParaBorrar(PB) = CStr(R)
                    End If
                End If
            Next R
        End If
    End If
    ' -------------------------------------------------------------------------------
    
    ' BORRO LOS QUE NO VAN -----------------------------------------------------------
    ' (de atras para adelante para que no haya problemas con el indice)
    If UBound(ParaBorrar) > 0 Then
        TmP = CStr(DB.ContarReg("SELECT ID FROM Clientes"))
        For R = CLng(TmP) To 1 Step -1
            If YaEstaBorrado(R, ParaBorrar) = True Then
                lvCli.ListItems.Remove (R)
            End If
        Next
    End If
End Sub

Private Function YaEstaBorrado(Ix As Long, Matrix() As String) As Boolean
    Dim Resp As Boolean, Ii As Long
    
    If UBound(Matrix) = 0 Then
        YaEstaBorrado = False
        Exit Function
    End If

    Resp = False
    For Ii = 0 To UBound(Matrix)
        If Matrix(Ii) = CStr(Ix) Then
            Resp = True
            Exit For
        End If
    Next Ii
    
    YaEstaBorrado = Resp
End Function

Private Sub AgregarUltimaCompra()
    If lvCli.ListItems.Count = 0 Then Exit Sub
    
    Dim I As Long, IdCF As Long, TmP As String
    
    For I = 1 To lvCli.ListItems.Count
        IdCF = CFG.ExistePropiedad("UPF " + lvCli.ListItems(I).Text)
        
        If IdCF <> 0 Then
            lvCli.ListItems(I).SubItems(6) = CFG.GetInfo(IdCF, 4)
            If lvCli.ListItems(I).SubItems(6) <> "" Then
                TmP = CFG.GetInfo(IdCF, 3)
                If TmP <> "" Then
                    lvCli.ListItems(I).SubItems(7) = DB.GetValInRS("Clientes", _
                        "Nombre", "ID = " + CStr(TmP), True)
                End If
            End If
        End If
    Next I
End Sub

Private Sub chkFin_Click()
    txtDiasFin = 0
    UsoFinanciera = True
    If chkFin Then
        Label8.Visible = True
        Label9.Visible = True
        txtDiasFin.Visible = True
    Else
        Label8.Visible = False
        Label9.Visible = False
        txtDiasFin.Visible = False
    End If
    
    If CargaInicial = False Then CargarDatos GetSQL
End Sub

Private Sub chkTrabaja_Click()
    UsoTrabaja = True
    txtIngreso = FormatCurrency(0)
    If chkTrabaja Then
        Label7.Visible = True
        txtIngreso.Visible = True
    Else
        Label7.Visible = False
        txtIngreso.Visible = False
    End If
    
    If CargaInicial = False Then CargarDatos GetSQL
End Sub

Private Sub cmdAdd_Click()
    frmClientes.AbrirDatos -1
End Sub

Private Sub cmdImprimir_Click()
    Dim Parte1 As String, Parte2 As String, TmP As String, tmPN As Long
    Dim TitCol As String, I As Long
    Dim Tit() As String
    
    TitCol = ""
    For I = 1 To lvCli.ColumnHeaders.Count
        If Len(lvCli.ColumnHeaders(I).Text) > 10 Then
            TitCol = TitCol + Left(lvCli.ColumnHeaders(I).Text, 7) + "..."
        Else
            TitCol = TitCol + lvCli.ColumnHeaders(I).Text
        End If
        
        If I < lvCli.ColumnHeaders.Count Then TitCol = TitCol + "|"
    Next I
    
    TmP = GetSQL: tmPN = Round(Len(GetSQL) / 2, 0)
    Parte1 = Left(TmP, tmPN)
    Parte2 = Right(TmP, Len(TmP) - tmPN)
    
    TP.LineasSeparadoras = True
    
    ReDim Preserve Tit(4)
    Tit(4) = "Listado de Clientes"
        'datos de mi empresa!!!!!!!!!!!!!!
    Tit(0) = DB.GetValInRS("Clientes", "Nombre", "ID = -2", True)
    Tit(1) = "Direccion: " + DB.GetValInRS("Clientes", "Direccion", "ID = -2", True)
    Tit(2) = "Teléfono: " + DB.GetValInRS("Clientes", "Telefono", "ID = -2", True)
    Tit(3) = "Mail: " + DB.GetValInRS("Clientes", "Mail", "ID = -2", True)
    
    TP.ImprimirlvW lvCli, Tit, _
        TitCol, "Filtro: " + Parte2, , , Parte1, _
        "Límite Cred.> " + txtLimite + ", Ingresos > " + txtIngreso + " " + _
        "Dias sin Comp.Finan.: " + txtDiasFin
End Sub

Private Sub cmdKILL_Click()
    If lvCli.ListItems.Count = 0 Then Exit Sub
        
    Dim IdCl As Long, nCli As String, IDp As Long
    Dim Deuda As Single
    
    IdCl = CLng(lvCli.ListItems(lvCli.SelectedItem.Index).Text)
    nCli = lvCli.ListItems(lvCli.SelectedItem.Index).SubItems(1)
    
    If IdCl = 0 Then
        MsgBox "No puede eliminar este cliente", vbInformation, "Atención"
        Exit Sub
    End If
        
    If MsgBox("¿Está seguro de borrar toda la información de " + _
        lvCli.ListItems(lvCli.SelectedItem.Index).SubItems(1) + "?" + vbCrLf + _
        "La información se perderá definitivamente", vbExclamation + vbYesNo, _
        "¡ATENCIÓN!") = vbNo Then Exit Sub
    
    'borro primero las cofiguraciones ----------------------------------------------
    '----- Último pago financiera
    IDp = CFG.ExistePropiedad("UPF " + CStr(IdCl))
    If IDp > 0 Then CFG.EliminarNodo IDp
    
    '----- Otros Detalles
    IDp = CFG.ExistePropiedad("DTC " + CStr(IdCl))
    If IDp > 0 Then CFG.EliminarNodo IDp
    
    'contabilidad anulo la deuda ----------------------------------------------------
    Dim Clscx As New clsCliente
    Deuda = Clscx.GetDeuda(IdCl)
    Set Clscx = Nothing
    
    PC.Asiento "57", CStr(Deuda), "46", CStr(Deuda), , "Eliminación Cliente " + nCli
    
    'ahora baso con base de datos ---------------------------------------------------
    'calculo que por integridad se borra el resto
    DB.EXECUTE "DELETE FROM Clientes WHERE ID = " + CStr(IdCl)
    
    'elimino el renglon
    lvCli.ListItems.Remove lvCli.SelectedItem.Index
        
End Sub

Private Sub cmdMOD_Click()
    frmClientes.AbrirDatos CLng(txtInLvW(lvCli, lvCli.SelectedItem.Index, 0))
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Limpiar(Optional TambienControles As Boolean = True)
    Dim XX As Control

    If TambienControles Then
        For Each XX In frmClientes2
            If TypeOf XX Is TextBox Then
                XX = ""
            End If
            
            If TypeOf XX Is tbrControles.MouTextBox Then
                XX = "0"
            End If
            
            txtEdad2 = "100"
            chkTrabaja.Value = 0
            chkFin.Value = 0
            txtIngreso = FormatCurrency(0)
            txtLimite = txtIngreso
        Next
    End If
    
    UsoFinanciera = False
    UsoTrabaja = False
    CargarDatos "SELECT * FROM Clientes WHERE ID>=0 ORDER BY ID"
End Sub

Private Sub Command2_Click()
    CargaInicial = True
    Limpiar
    CargaInicial = False
End Sub

Private Sub Form_Activate()
    Me.Height = 8000
    Me.Width = 11500
    
    If CargaInicial = False Then CargarDatos GetSQL
End Sub

Public Function GetSQL() As String
    'devuelve el SQLText para la tabla de Clientes
    Dim S As String, TmP As String
    S = ""
    If txtDireccion <> "" Then S = S + " Direccion like '%" + txtDireccion + "%' "
    '------------------------------------------------------------------------------
    If IsNumeric(txtNAME) Then
        S = S + " ID like '%" + txtNAME + "%'"
    Else
        If txtNAME <> "" Then
            If S <> "" Then S = S + " AND "
            S = S + " Nombre like '%" + txtNAME + "%'"
        End If
    End If
    '-------------------------------------------------------------------------------
    If txtCondIVA <> "" Then
        If S <> "" Then S = S + " AND "
        S = S + " IVA like '%" + txtCondIVA + "%' "
    End If
    '-------------------------------------------------------------------------------
    If UsoTrabaja Then 'solo si hizo un click y demostro que queria elegir eso
        If S <> "" Then S = S + " AND "
        Dim xF As Long
        If chkTrabaja Then
            S = S + "TieneTrabajo <> 0"
        Else
            S = S + "TieneTrabajo = 0"
        End If
    End If
    '-------------------------------------------------------------------------------
    If txtEdad1 <> "" And txtEdad2 <> "" Then
        If IsNumeric(txtEdad1) And IsNumeric(txtEdad2) Then
            Dim Nac1 As String, Nac2 As String
            
            If S <> "" Then S = S + " AND "
            If CLng(txtEdad1) > CLng(txtEdad2) Then 'los ordeno
                Nac1 = txtEdad1: txtEdad1 = txtEdad2: txtEdad2 = Nac1
                'entonces siempre txtedad1 es el menor
            End If
            
            Nac1 = CStr(Month(Date)) + "/" + CStr(Day(Date)) + "/" + _
                CStr(Year(Date) - CLng(txtEdad1))
            Nac2 = CStr(Month(Date)) + "/" + CStr(Day(Date) + 1) + "/" + _
                CStr(Year(Date) - CLng(txtEdad2) - 1)
            S = S + " Nacimiento BETWEEN #" + Nac1 + "# AND #" + Nac2 + "#"
        End If
    End If
    '-------------------------------------------------------------------------------
    TmP = "SELECT * FROM Clientes WHERE ID>=0"
    If S <> "" Then TmP = TmP + " AND " + S
    TmP = TmP + " ORDER BY ID"
    
    GetSQL = TmP
End Function

Private Function GetAnos(FechaNac As Date) As Long
    Dim Resp As Long
    
    Resp = Year(Date) - Year(FechaNac)
    If Month(Date) < Month(FechaNac) Then Resp = Resp - 1
    
    GetAnos = Resp
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    'Command2_Click
    UsoFinanciera = False
    UsoTrabaja = False
    CargaInicial = True
    txtEdad1 = "0"
    txtEdad2 = "100"
    txtLimite = FormatCurrency(0)
    CargarDatos GetSQL
    CargaInicial = False
End Sub

Private Sub Form_Resize()
    cmdADD.Top = Me.Height - cmdADD.Height - 650
    cmdMOD.Top = cmdADD.Top
    cmdKILL.Top = cmdADD.Top
    command1.Top = cmdADD.Top
    Command2.Top = cmdADD.Top
    cmdImprimir.Top = cmdADD.Top

    Frame1.Width = Me.Width - Frame1.Left - 150
    lvCli.Width = Frame1.Width
    lvCli.Height = Me.Height - cmdADD.Height - 1000 - lvCli.Top
End Sub

Private Sub txtCondIVA_Change()
    If CargaInicial = False Then CargarDatos GetSQL
End Sub

Private Sub txtDiasFin_Change()
    If CargaInicial = False Then CargarDatos GetSQL
End Sub

Private Sub txtDiasFin_GotFocus()
    PintarTxt txtDiasFin
End Sub

Private Sub txtDiasFin_LostFocus()
    txtDiasFin = ValidarNumeros(txtDiasFin)
End Sub

Private Sub txtDireccion_Change()
    If CargaInicial = False Then CargarDatos GetSQL
End Sub

Private Sub txtEdad1_Change()
    If CargaInicial = False Then CargarDatos GetSQL
End Sub

Private Sub txtEdad1_GotFocus()
    PintarTxt txtEdad1
End Sub

Private Sub txtEdad1_LostFocus()
    txtEdad1 = ValidarNumeros(txtEdad1)
End Sub

Private Sub txtEdad2_Change()
    If CargaInicial = False Then CargarDatos GetSQL
End Sub

Private Sub txtEdad2_GotFocus()
    PintarTxt txtEdad2
End Sub

Private Sub txtEdad2_LostFocus()
    txtEdad2 = ValidarNumeros(txtEdad2, 100)
End Sub

Private Sub txtIngreso_Change()
    If CargaInicial = False Then CargarDatos GetSQL
End Sub

Private Sub txtIngreso_GotFocus()
    PintarTxt txtIngreso
End Sub

Private Sub txtIngreso_LostFocus()
    txtIngreso = FormatCurrency(ValidarNumeros(txtIngreso), , , , vbFalse)
End Sub

Private Sub txtLimite_Change()
    If CargaInicial = False Then CargarDatos GetSQL
End Sub

Private Sub txtLimite_GotFocus()
    PintarTxt txtLimite
End Sub

Private Sub txtLimite_LostFocus()
    txtLimite = FormatCurrency(ValidarNumeros(txtLimite), , , , vbFalse)
End Sub

Private Sub txtNAME_Change()
    If CargaInicial = False Then CargarDatos GetSQL
End Sub

Private Sub txtNAME_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
