VERSION 5.00
Object = "{A7FBD38D-2930-49E3-B60C-9E0202D84549}#15.0#0"; "tbrControles.ocx"
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSocyEmp 
   BackColor       =   &H0049453D&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cuenta Socios"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8745
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSocyEmp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00544B45&
      Caption         =   "Imprimir"
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   5760
      TabIndex        =   29
      Top             =   7020
      Width           =   2475
      Begin VB.OptionButton chkMov 
         BackColor       =   &H00544B45&
         Caption         =   "Últimos 20 Movimientos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   180
         TabIndex        =   32
         Top             =   810
         Width           =   2250
      End
      Begin VB.OptionButton chkTodo 
         BackColor       =   &H00544B45&
         Caption         =   "Todo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   180
         TabIndex        =   31
         Top             =   240
         Value           =   -1  'True
         Width           =   1515
      End
      Begin VB.OptionButton chkDias 
         BackColor       =   &H00544B45&
         Caption         =   "Últimos 30 Días"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   180
         TabIndex        =   30
         Top             =   510
         Width           =   1890
      End
   End
   Begin tbrFaroButton.fBoton cmdModificar 
      Height          =   420
      Left            =   3435
      TabIndex        =   5
      Top             =   1395
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   741
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Modificar"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   9
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdEliminar 
      Height          =   420
      Left            =   3435
      TabIndex        =   6
      Top             =   1875
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   741
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Eliminar"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   9
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdAdd 
      Height          =   420
      Left            =   3435
      TabIndex        =   4
      Top             =   945
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   741
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Agregar"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   9
      fECol           =   5717301
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0049453D&
      Caption         =   "Aportes"
      ForeColor       =   &H00FFFFFF&
      Height          =   1185
      Left            =   4680
      TabIndex        =   18
      Top             =   2805
      Width           =   3585
      Begin tbrControles.MouTextBox txtAporte 
         Height          =   345
         Left            =   420
         TabIndex        =   21
         Top             =   600
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
      Begin tbrFaroButton.fBoton cmdAportes 
         Height          =   420
         Left            =   1830
         TabIndex        =   22
         Top             =   570
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   741
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Registar"
         fEnabled        =   -1  'True
         fFontN          =   "Arial"
         fFontS          =   9
         fECol           =   5717301
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Monto"
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Left            =   420
         TabIndex        =   23
         Top             =   360
         Width           =   705
      End
   End
   Begin MSComctlLib.ListView lvDeudas 
      Height          =   2265
      Left            =   210
      TabIndex        =   14
      Top             =   4680
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   3995
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Fecha"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Variación"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Detalle"
         Object.Width           =   6791
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "IDM"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Saldo"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0049453D&
      Caption         =   "Tipo de Movimiento"
      ForeColor       =   &H00FFFFFF&
      Height          =   2565
      Left            =   4680
      TabIndex        =   10
      Top             =   120
      Width           =   3585
      Begin VB.TextBox txtDetalle 
         Height          =   405
         Left            =   330
         TabIndex        =   2
         Top             =   1830
         Width           =   2925
      End
      Begin VB.OptionButton chkResta 
         BackColor       =   &H0049453D&
         Caption         =   "Extracción"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   210
         TabIndex        =   0
         Top             =   270
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton chkSuma 
         BackColor       =   &H0049453D&
         Caption         =   "Depósito"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   210
         TabIndex        =   8
         Top             =   540
         Width           =   1755
      End
      Begin tbrControles.MouTextBox txtVariacion 
         Height          =   345
         Left            =   360
         TabIndex        =   1
         Top             =   1170
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
      Begin tbrFaroButton.fBoton cmdRegistrar 
         Height          =   420
         Left            =   1680
         TabIndex        =   3
         Top             =   1140
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   741
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Registar"
         fEnabled        =   -1  'True
         fFontN          =   "Arial"
         fFontS          =   9
         fECol           =   5717301
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Monto"
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Left            =   360
         TabIndex        =   20
         Top             =   930
         Width           =   705
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle"
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   330
         TabIndex        =   19
         Top             =   1590
         Width           =   675
      End
   End
   Begin VB.ListBox lstSocyEmp 
      Height          =   2400
      Left            =   360
      TabIndex        =   7
      Top             =   630
      Width           =   2895
   End
   Begin tbrControles.MouTextBox txtComision 
      Height          =   345
      Left            =   1830
      TabIndex        =   15
      Top             =   3150
      Visible         =   0   'False
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   609
      Alignment       =   2
      BackColor       =   16777215
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin tbrFaroButton.fBoton Command1 
      Height          =   420
      Left            =   2280
      TabIndex        =   24
      Top             =   7140
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   741
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   9
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdAcomodar 
      Height          =   420
      Left            =   210
      TabIndex        =   25
      Top             =   7140
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   741
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Pasar a 1 Reg."
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   9
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdParti 
      Height          =   420
      Left            =   180
      TabIndex        =   26
      Top             =   7680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   741
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Participación"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   9
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdComision 
      Height          =   420
      Left            =   780
      TabIndex        =   27
      Top             =   3810
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   741
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Grabar Configuración"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   9
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdImprimir 
      Height          =   435
      Left            =   4170
      TabIndex        =   28
      Top             =   7170
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Imprimir"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   2850
      TabIndex        =   17
      Top             =   3180
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Comisiones"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   570
      TabIndex        =   16
      Top             =   3210
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Ale"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   2730
      TabIndex        =   13
      Top             =   4290
      Width           =   3645
   End
   Begin VB.Label lblPesos 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "111111111"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6690
      TabIndex        =   12
      Top             =   4260
      Width           =   1605
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta Particular de:"
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   810
      TabIndex        =   11
      Top             =   4290
      Width           =   1845
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   450
      TabIndex        =   9
      Top             =   300
      Width           =   2295
   End
End
Attribute VB_Name = "frmSocyEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cuenta As String 'su es socio o empleado
Dim nCuenta As Long
Dim AnOtaR As Single
Dim UltRengImp As Long

Private Sub VerUltReng() 'para ver que eligio si va a imprimir
    If UltRengImp <> 0 Then 'solo cambia el caption cuando no elige Todo
        If UltRengImp < 0 Then
            'eligio dias atras
            chkDias.Caption = "Últimos " + CStr(-UltRengImp) + " días"
        Else
            chkMov.Caption = "Últimos " + CStr(UltRengImp) + " movimientos"
        End If
    End If

End Sub

Public Sub AbrirDatos(EsSocio As Boolean)
    If EsSocio Then
        Cuenta = "Socio"
        nCuenta = 52
    Else 'es empleado
        Cuenta = "Empleado"
        nCuenta = 53
        Frame2.Visible = False
        cmdParti.Visible = False
    End If
    
    Me.Show 1
End Sub

Private Sub ReCargarDatos()
    Dim inD As Long
    Dim IdCuentas() As String
    
    IdCuentas = PC.GetCuentas(nCuenta)
    
    If lstSocyEmp.ListCount > 0 Then inD = lstSocyEmp.ListIndex
    
    lstSocyEmp.Clear
    
    Dim I As Long
    
    For I = 1 To UBound(IdCuentas)
        lstSocyEmp.AddItem PC.GetNameCuenta(CLng(IdCuentas(I)))
    Next I
    
    'dejo el indice donde estaba
    If lstSocyEmp.ListCount > 0 Then
        If inD >= lstSocyEmp.ListCount Then
            lstSocyEmp.ListIndex = 0
        Else
            lstSocyEmp.ListIndex = inD
        End If
    End If
    
    VerSiHay
End Sub

Private Sub ActualizarResumen()
    If lstSocyEmp.ListCount > 0 Then 'hago el resumen
        lblNombre = lstSocyEmp
        
        Dim W As String, IDC As Long, IdCF As Long, J As Long, Res As Single
        
        IDC = PC.GetIDCuenta(lstSocyEmp)
        cmdComision.Enabled = True
        
        W = "SELECT * FROM movSocyEmp WHERE Idnivel3= " + CStr(IDC) + _
            " ORDER BY Fecha DESC, ID DESC"
        CargarComboLV lvDeudas, W, "Fecha/f,Variacion/$,Detalle,id/n,id/n"
        'la ultima columna la quiero para saldo pero como si esta vacia da
        'error le pongo cualquier dato como "ID" y despues SI la lleno con saldos
        
        'cargo los saldos
        For J = lvDeudas.ListItems.Count To 1 Step -1
            If J < lvDeudas.ListItems.Count Then
                Res = CSng(lvDeudas.ListItems(J).ListSubItems(1)) + _
                    lvDeudas.ListItems(J + 1).ListSubItems(4)
            Else
                Res = CSng(lvDeudas.ListItems(J).ListSubItems(1))
            End If

            lvDeudas.ListItems(J).ListSubItems(4) = FormatCurrency(Res)
        Next J
        '----------------------------------------------------------
    
        lblPesos = FormatCurrency(PC.ABSValor(IDC, PC.GetSaldo(IDC)))
        IdCF = CFG.GetID("Comision " + lstSocyEmp)
        txtComision = NoNuloN(CFG.GetInfo(IdCF, 4))
    Else
        lvDeudas.ListItems.Clear
        lblPesos = FormatCurrency(0)
        cmdComision.Enabled = False
    End If
End Sub

Private Sub chkDias_Click()
    Dim TmP As String
    TmP = InputBox("Ingrese hasta cuántos días desea imprimir " + vbCrLf + _
        "en el resumen", "Días Atras", "30")
    If Not IsNumeric(TmP) Then
        chkTodo.Value = True
        UltRengImp = 0
    Else
        UltRengImp = -CLng(TmP)
    End If
    
    VerUltReng
End Sub

Private Sub chkMov_Click()
    Dim TmP As String
    TmP = InputBox("Ingrese hasta cuántos movimientos desea imprimir " + vbCrLf + _
        "en el resumen", "Días Atras", "10")
    If Not IsNumeric(TmP) Then
        chkTodo.Value = True
        UltRengImp = 0
    Else
        UltRengImp = CLng(TmP)
    End If
    
    VerUltReng
End Sub

Private Sub chkTodo_Click()
    UltRengImp = 0
    VerUltReng
End Sub

Private Sub cmdAcomodar_Click()
    Dim MsJe As String
    If lstDeudas.ListCount = 0 Then Exit Sub
    
    If MsgBox("Atención, ejecutando esta opción borrará todos los registros " + _
        "de: " + UCase(Nombre) + " Dejando solo un registro con el saldo " + lblPesos + _
        ". Si está seguro, presione aceptar, en caso contrario cancele. ", vbInformation, _
        "Atención") = vbCancel Then 'si esta seguro y puso la clave listo borro y dejo solo el saldo
        Exit Sub
    End If
    
    Dim IDC As Long
    IDC = PC.GetIDCuenta(lstSocyEmp)
        
    MsJe = InputBox("Escriba aqui una aclaración del movimiento para que " + _
        "quede en el registro", "Aclaración")
    
    'borro todo
    DB.EXECUTE "DELETE FROM MovSocyEmp WHERE IdNivel3=" + CStr(IDC)
    
    'agrego un renglon con el saldo
    DB.EXECUTE "INSERT INTO MovSocyEmp (ID,Fecha,IdNivel3,Detalle,Variacion) " + _
        "VALUES (" + IdAutonum("MovSocyEmp") + _
        ",#" + stFechaSQL(Date) + "#," + CStr(IDC) + ",'Saldo: " + lblPesos + " " + MsJe + "'," + _
        Replace(CStr(CSng(lblPesos)), ",", ".") + ")"
    
    ActualizarResumen
    
End Sub

Private Sub cmdAdd_Click()
    Dim NombreC As String
    
    NombreC = Replace(InputBox("Ingrese el nuevo Nombre de la cuenta", "Modificar"), _
        "'", " ")
    
    'veo que no haya ningun otra cuenta con ese nombre
    If PC.ExisteNCuenta(NombreC) <> 0 Or NombreC = "" Then
        MsgBox "Ingreso incorrecto o Cuenta existente", vbInformation, "Atención"
        Exit Sub
    End If
        
    If Cuenta = "Empleado" Then
        'primero lo agrego en sueldos y salarios(perdida)
        Dim nCta As Long
        
        PC.AgregarCuenta PC.GetUltIDMasUno, 36, "Sueldo " + NombreC
                        
        'y tambien agrego su cuenta particular
        PC.AgregarCuenta PC.GetUltIDMasUno, 53, NombreC
        
        'agrego en configuracion una que sea "comision "+nombre
        CFG.AgregarNodo 20, "Comision " + NombreC, "", 0, 0
        
        
    Else 'es socio
        PC.AgregarCuenta PC.GetUltIDMasUno, 52, NombreC
    End If
    
    'agrego tambien un evento "A COSTO - XXXX" para que el empleado o socio saque
    'a costo unicamente lo suyo (solo se habilita a el) ... bueno lo debe habilitar
    'el administrador para que ese empleado o socio pueda hacerlo, aca solo creo el evento
    ACC.AgregarEvento "A Costo - " + NombreC
        
    ReCargarDatos
End Sub

Private Sub cmdAportes_Click()
    If lstSocyEmp.ListCount = 0 Or lstSocyEmp.ListIndex = -1 Then
        Exit Sub
    End If
    
    If Not IsNumeric(txtAporte) Then
        MsgBox "Debes cargar un número correcto", vbInformation, "Atencion"
        PintarTxt txtVariacion
        Exit Sub
    End If
    
    'dejo que meta importe negativos para corregir
    If CSng(txtAporte) = 0 Then
        PintarTxt txtVariacion
        Exit Sub
    End If
    
    'listo registro
    Aporte lstSocyEmp, CSng(txtAporte), "Ingreso Manual Aporte"
    
    txtAporte = FormatCurrency(0)
    MsgBox "Aporte Registrado correctamente", vbInformation, "Atención"
End Sub

Private Sub cmdComision_Click()
    If lstSocyEmp.ListIndex = -1 Then Exit Sub
        
    txtComision = ValidarNumeros(txtComision, -1)
    If CSng(txtComision) = -1 Then
        lstSocyEmp_Click
        Exit Sub
    End If
    
    Dim IdCF As Long
     
    IdCF = CFG.GetID("Comision " + lstSocyEmp)
    CFG.ModificarNodo IdCF, , , , CStr(CSng(txtComision))
End Sub

Private Sub cmdEliminar_Click()
    Dim TmP As Single
    
    If lstSocyEmp.ListCount = 1 Then
        If Cuenta = "Socio" Then
            MsgBox "Debe al menos un socio cargado en el sistema", vbInformation, "Atención"
            Exit Sub
        End If
    End If
    
    '(a) ---------------------- VALIDACION -------------------------------------------
    Dim UUs As Long
    UUs = ACC.UltUsuarioIngresado
    
    If ACC.ExisteRelacion(UUs, 17) = 0 Then
        MsgBox ACC.GetNombre("Usuario", "Usuarios", UUs) + " no está habilitado " + _
            "para ingresar." + vbCrLf + _
            "Debe Cambiar Sesión a la de un usuario habilitado", vbExclamation, "Atención"
        Exit Sub
    End If
    
    'registro el movimiento(17:Borrar Socios y Empleados)
    ACC.RegEvento UUs, 17, "Borrado de " + Cuenta + " " + lstSocyEmp
    '----------------------------------------------------------------------------
    
    If MsgBox("¿Está seguro que desea eliminar la cuenta?. Los Datos se borrarán " + _
        "definitivamente.", vbExclamation + vbOKCancel, "Borrar Cuenta") = vbCancel Then Exit Sub
    
    'peligroso tener cuidado!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    'cosas a hacer
    '(1) si tiene movimientos avisarle que se van a borrar y van como resultados
    '(2) pasar el saldo como resultado si tuviera (perdida o ganancia) NO
        'lo que voy a hacer en el paso 4 cuando eliminecuenta modifica por 57 (RECta)
    '(3) borrarle todos los movimientos
    '(4) borrarle la cuenta nomas
    '(5) si es empleado tambien borrar la cuenta sueldo + .....
    '(6) eliminar el evento "A Costo - "+(empleado o usuario)
    'creo que no falta nada vamos a ver que pasa
    '(7) Eliminar la configuracion de la comision
    '(8) Elimino los aportes si tiene
    
    '(1)-------- ¿Tiene saldo en su cuenta? ----------------------------------
    Dim Rsdo As Single, IdCz As Long
    Rsdo = CSng(Replace(lblPesos, ".", ""))
    IdCz = PC.GetIDCuenta(lstSocyEmp)
'
'    If EsCero(Rsdo) = False Then
'        '(2) ------------- paso a resultados ----------------------------
'        'asiento: cuentax a 57-Resultado Eliminacion Cuenta (si es egreso (sdo negativo) ya vere)
'        PC.Asiento CStr(IdCz), CStr(Rsdo), "57", CStr(Rsdo), , _
'            "Por Eliminación de cuenta " + CStr(IdCz)
'    End If
    
    '(3) -------- Borro movimientos de este ------------------------------------
    DB.EXECUTE "DELETE FROM MovSocyEmp WHERE IdNivel3 = " + CStr(IdCz)
    
    '(4) --------- Borro las cuentas -------------------------------------------
    PC.EliminarCuenta IdCz
    
    '(5) --------- Si es empleado tambien tiene una cuenta "Sueldo "+nombre -----
    If Cuenta = "Empleado" Then
        Dim iDC2 As Long
        iDC2 = PC.GetIDCuenta("Sueldo " + lstSocyEmp)
        PC.EliminarCuenta iDC2
    End If
    
    '(6) ----------- Borro evento "A Costo - xx" --------------------------------
    ACC.EliminarEvento "A Costo - " + lstSocyEmp
    
    '(7) Elimino nombre de la configuracion de comision de vendedor
    If Cuenta = "Empleado" Then
        IdCz = CFG.GetID("Comision " + lstSocyEmp)
        CFG.EliminarNodo IdCz
    End If
    
    '(8) Elimino Aportes
    iDC2 = PC.GetIDCuenta("Aporte Participación " + lstSocyEmp)
    If iDC2 > -1 Then
        TmP = -PC.GetSaldo(iDC2)
        PC.Asiento CStr(iDC2), CStr(TmP), "15", CStr(TmP), , _
            "Eliminada cuenta participación socio"
        PC.EliminarCuenta iDC2
    End If
    
    lstSocyEmp_Click
    ReCargarDatos
End Sub

Private Sub cmdImprimir_Click()
    Dim Tit() As String

    TP.LineasSeparadoras = True
    
    ReDim Preserve Tit(4)
    Tit(4) = "Resumen de Cuenta"
    Tit(0) = lblNombre.Caption
    Tit(1) = ""
    Tit(3) = ""
    Tit(2) = ""

    If Date + UltRengImp > CDate(lvDeudas.ListItems(1).Text) And UltRengImp < 0 Then
        MsgBox "No hay movimientos desde el " + CStr(Date + UltRengImp), vbInformation, "Atención"
        Exit Sub
    End If

    TP.ImprimirlvW lvDeudas, Tit, _
        "Fecha|Monto|Detalle|ID|Saldo", "Saldo Actual: " + CStr(lblPesos.Caption), _
        , , , , , 30, UltRengImp
End Sub

Private Sub cmdModificar_Click()
    If lstSocyEmp.ListCount = 0 Then
        MsgBox "No ha seleccionado ninguna cuenta", vbInformation, "Atención"
        Exit Sub
    End If
    
    Dim NombreC As String, NombreV As String, IdViejo As Long
    NombreV = lstSocyEmp
    
    NombreC = Replace(InputBox("Ingrese el nuevo Nombre de la cuenta", "Modificar", _
        lstSocyEmp), "'", " ")
    
    'veo que no haya ningun otra cuenta con ese nombre
    If PC.ExisteNCuenta(NombreC) <> 0 Or NombreC = "" Then
        MsgBox "Ingreso incorrecto o Cuenta existente", vbInformation, "Atención"
        Exit Sub
    End If
        
    'sea socio o empleado las cuentas son irrepetibles
    IdViejo = PC.GetIDCuenta(NombreV)
    PC.ModificarCuenta IdViejo, , NombreC
    
    'en empleados esta la cuenta sueldos
    IdViejo = PC.GetIDCuenta("Sueldo " + NombreV)
    PC.ModificarCuenta IdViejo, , "Sueldo " + NombreC
     
    'en socios esta la cuenta de participacion!
    IdViejo = PC.GetIDCuenta("Aporte Participación " + NombreV)
    PC.ModificarCuenta IdViejo, , "Aporte Participación " + NombreC
     
    ACC.ModificarEvento "A Costo - " + NombreV, "A Costo - " + NombreC
    
    'cambio el nombre de la configuracion de comision de vendedor
    If Cuenta = "Empleado" Then
        IdViejo = CFG.GetID("Comision " + NombreV)
        CFG.ModificarNodo IdViejo, , "Comision " + NombreC
    End If
    
    ReCargarDatos
End Sub

Private Sub cmdParti_Click()
    frmParticipacion.Show 1
End Sub

Private Sub cmdRegistrar_Click()
    Dim strQue As String
    
    If lstSocyEmp.ListCount = 0 Then
        MsgBox "No tiene Registros de " + Cuenta + "s", vbInformation, "Atención"
        Exit Sub
    End If
    
    If Not IsNumeric(txtVariacion) Then
        MsgBox "Debes cargar un número correcto", vbInformation, "Atencion"
        PintarTxt txtVariacion
        Exit Sub
    End If
    
    If CSng(txtVariacion) = 0 Then
        PintarTxt txtVariacion
        Exit Sub
    End If
    
    If chkResta Then 'extrae ft
        strQue = chkResta.Caption
        AnOtaR = -CSng(txtVariacion)
    Else 'deposita guita
        AnOtaR = CSng(txtVariacion)
        strQue = chkSuma.Caption
    End If
    
    If MsgBox("¿Está seguro de registrar el " + UCase(strQue) + " por " + _
        txtVariacion + vbCrLf + _
        " a " + UCase(lstSocyEmp) + "?", vbInformation + vbYesNo, "Registro") = vbNo Then Exit Sub
    
    'ya basta de vueltas se los anoto
    IdCz = PC.GetIDCuenta(lstSocyEmp)
    
    'es indiferente socios y empleados estan en el mismo nivel
    DB.EXECUTE "INSERT INTO MovSocyEmp (ID,Fecha, IdNivel3," + _
        "Tipo,Variacion,Detalle) VALUES (" + IdAutonum("MovSocyEmp") + _
        ",#" + stFechaSQL(Date) + _
        "#," + CStr(IdCz) + ",'" + Cuenta + "'," + _
        Replace(CStr(AnOtaR), ",", ".") + _
        ",'" + txtDetalle + "')"
    
    'no viene de nada asi que seria caja a (cta de soc o emp)
    'anotar esta predeterminado positivo el deposito
    If Cuenta = "Socio" Then
        PC.Asiento "78", CStr(AnOtaR), CStr(IdCz), CStr(AnOtaR), , _
            "Ingreso Manual Socios y Empleados"
    Else 'EMPLEADO
        If chkSuma Then 'no hay movimiento de caja solo se le acredito a su cuenta
                           'por su trabajo es sdos(perdida) a cta part socio
            
            PC.Asiento PC.GetIDCuenta("Sueldo " + lstSocyEmp), CStr(AnOtaR), _
                PC.GetIDCuenta(lstSocyEmp), CStr(AnOtaR), , _
                "Ingreso Manual Socios y Empleados"
        
        Else 'es menor que cero cuando se paga o sea hay mov de caja
             'es cta empleado a caja
        
            PC.Asiento PC.GetIDCuenta(lstSocyEmp), CStr(-CSng(AnOtaR)), _
                "78", CStr(-CSng(AnOtaR)), , "Ingreso Manual Socios y Empleados"
        End If
    End If
    
    MsgBox "Se ha realizado el registro correctamente", vbInformation, "Registro Completo"
    
    AnOtaR = 0
    txtVariacion = FormatCurrency(AnOtaR)
    txtDetalle = ""
    
    ReCargarDatos
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    lblPesos = FormatCurrency(0)
    
    AnOtaR = 0
    
    frmSocyEmp.Caption = "Cuenta " + Cuenta + "s"
    Label1 = "Seleccione " + Cuenta
    txtVariacion = FormatCurrency(AnOtaR)
    txtAporte = FormatCurrency(0)
    
    If Cuenta = "Empleado" Then
        chkSuma.Caption = "Acreditar Sueldo"
        chkResta.Caption = "Pago Sueldo"
        Label6.Visible = True
        Label7.Visible = True
        txtComision.Visible = True
        cmdComision.Visible = True
    End If
    
    UltRengImp = 0
    ReCargarDatos
    ActualizarResumen
    VerAportes
End Sub

Private Sub VerAportes()
    Dim TmP As String, Kta As Single, Kapi As Single, I As Long
    Dim Hij() As String, Tm2 As Single, tM3 As Single, tm4 As Long
    
    If Cuenta <> "Socio" Then Exit Sub
    If lstSocyEmp.ListCount = 0 Then Exit Sub
    
    Kapi = PC.ABSSumarconSubcuentas(15)
    Hij = PC.GetCuentas(15)
    
    Kta = 0
    If UBound(Hij) > 0 Then
        For I = 1 To UBound(Hij)
            TmP = PC.GetNameCuenta(CLng(Hij(I)))
            If Left(TmP, 20) = "Aporte Participación" Then
                'veo si el existe el socio
                'resto todo ya que las que se borran(ctas) se pasan automaticamente a K
                Tm2 = -PC.GetSaldo(CLng(Hij(I)))
                Kta = Kta + Tm2
                If ExisteSocio(Right(TmP, Len(TmP) - 21)) Then
                    '(1) controlo que coincida Cont-BD
                        'se queda con lo que dice MovSocyEmp BD Principal
                    tM3 = DB.SumarValInRS("MovSocyEmp", "Variacion", _
                        "IDNivel3 = " + Hij(I) + " AND Detalle = '" + _
                        "Aporte Participación Socio'")
                    If Abs(Tm2 - tM3) > 1 Then 'si hay mas de $1 de diferencia ajusto
                        PC.Asiento "15", CStr(tM3 - Tm2), Hij(I), _
                            CStr(tM3 - Tm2), , _
                            "Ajuste por diferencia Socios - Contabilidad (" + _
                            Hij(I) + ")"
                    End If
                    
                Else 'no existe socio lo borro (paso a capital)
                    PC.Asiento Hij(I), CStr(Tm2), "15", CStr(Tm2), , _
                        "Eliminada cuenta participación socio inexistente"
                    PC.EliminarCuenta CLng(Hij(I))
                End If
            End If
        Next I
    End If
    
    'hay diferencia entre BD y Cont
    If Abs(Kapi - Kta) > 1 Then
        'depende de cuantos socios haya
        tm4 = lstSocyEmp.ListCount
        'cuanto pa cada uno
        tM3 = (Kapi - Kta) / tm4
        For I = 0 To tm4 - 1
            Aporte lstSocyEmp.List(I), tM3, "Ajuste Participación Socio " + _
                lstSocyEmp.List(I), 15
        Next I
    End If
End Sub

Private Function ExisteSocio(nSocio As String) As Boolean
    Dim Resp As Boolean, I As Long
    
    If lstSocyEmp.ListCount = 0 Then
        ExisteSocio = False
        Exit Function
    End If
    
    Resp = False
    
    For I = 0 To lstSocyEmp.ListCount - 1
        If lstSocyEmp.List(I) = nSocio Then
            Resp = True
            Exit For
        End If
    Next I
    
    ExisteSocio = Resp
End Function

Private Sub VerSiHay()
    If lstSocyEmp.ListCount <= 0 Then
        cmdRegistrar.Enabled = False
        Frame1.Enabled = False
        lvDeudas.ListItems.Clear
        cmdComision.Enabled = False
    Else
        cmdRegistrar.Enabled = True
        Frame1.Enabled = True
        cmdComision.Enabled = True
    End If
End Sub

Private Sub lstSocyEmp_Click()
    ActualizarResumen
End Sub

Private Sub txtAporte_GotFocus()
    PintarTxt txtAporte
End Sub

Private Sub txtAporte_LostFocus()
    txtAporte = FormatCurrency(ValidarNumeros(txtAporte))
End Sub

Private Sub txtComision_GotFocus()
    PintarTxt txtComision
End Sub

Private Sub txtComision_LostFocus()
    txtComision = ValidarNumeros(txtComision)
End Sub

Private Sub txtVariacion_GotFocus()
    PintarTxt txtVariacion
End Sub

Private Sub txtVariacion_LostFocus()
    AnOtaR = ValidarNumeros(txtVariacion)
    txtVariacion = FormatCurrency(AnOtaR, , , , vbFalse)
End Sub
