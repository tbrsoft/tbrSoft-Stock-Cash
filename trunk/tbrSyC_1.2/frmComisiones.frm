VERSION 5.00
Object = "{A7FBD38D-2930-49E3-B60C-9E0202D84549}#15.0#0"; "tbrControles.ocx"
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmComisiones 
   BackColor       =   &H0049453D&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comisiones por Vendedor"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9180
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmComisiones.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   9180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin tbrFaroButton.fBoton Command1 
      Height          =   480
      Left            =   7000
      TabIndex        =   18
      Top             =   4900
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   847
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdAcreditar 
      Height          =   480
      Left            =   3150
      TabIndex        =   17
      Top             =   4900
      Width           =   2000
      _ExtentX        =   3519
      _ExtentY        =   847
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Acredtar Comisiones"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdComision 
      Height          =   480
      Left            =   900
      TabIndex        =   16
      Top             =   4900
      Width           =   2000
      _ExtentX        =   3519
      _ExtentY        =   847
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Grabar Configuración"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdVer 
      Height          =   480
      Left            =   7000
      TabIndex        =   15
      Top             =   2190
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   847
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Aceptar"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin MSComCtl2.DTPicker DTPickerDe 
      Height          =   375
      Left            =   3750
      TabIndex        =   1
      Top             =   1650
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   61341697
      CurrentDate     =   39259
   End
   Begin VB.ListBox lstSocyEmp 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2790
      Left            =   450
      TabIndex        =   0
      Top             =   900
      Width           =   3105
   End
   Begin tbrControles.MouTextBox txtComision 
      Height          =   345
      Left            =   1920
      TabIndex        =   3
      Top             =   4020
      Width           =   915
      _ExtentX        =   1614
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
   Begin MSComCtl2.DTPicker DTPickerA 
      Height          =   375
      Left            =   5220
      TabIndex        =   2
      Top             =   1650
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   61341697
      CurrentDate     =   39259
   End
   Begin tbrControles.MouTextBox txtVentas 
      Height          =   345
      Left            =   7200
      TabIndex        =   4
      Top             =   2850
      Width           =   1305
      _ExtentX        =   2302
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
   Begin tbrControles.MouTextBox txtComDebe 
      Height          =   345
      Left            =   7200
      TabIndex        =   5
      Top             =   3420
      Width           =   1305
      _ExtentX        =   2302
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
   Begin VB.Label lblEmpleado 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Obtener Ventas del Vendedor seleccionado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3750
      TabIndex        =   14
      Top             =   870
      Width           =   2715
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Comisiones que corresponderían"
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
      Left            =   4050
      TabIndex        =   13
      Top             =   3480
      Width           =   3075
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ventas del Período"
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
      Height          =   435
      Left            =   5100
      TabIndex        =   12
      Top             =   2910
      Width           =   1995
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta el día"
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
      Height          =   285
      Left            =   5280
      TabIndex        =   11
      Top             =   1380
      Width           =   1065
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Desde el día"
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
      Height          =   285
      Left            =   3780
      TabIndex        =   10
      Top             =   1380
      Width           =   1065
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Obtener Ventas del Vendedor seleccionado"
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
      Height          =   435
      Left            =   4080
      TabIndex        =   9
      Top             =   420
      Width           =   2055
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Comisiones"
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
      Height          =   285
      Left            =   660
      TabIndex        =   8
      Top             =   4080
      Width           =   1065
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   2940
      TabIndex        =   7
      Top             =   4050
      Width           =   285
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Empleados"
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
      Height          =   285
      Left            =   480
      TabIndex        =   6
      Top             =   570
      Width           =   2295
   End
End
Attribute VB_Name = "frmComisiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub VerResultados()
    If lstSocyEmp.ListIndex = -1 Then Exit Sub
    txtComision = ValidarNumeros(txtComision)
    
    Dim Vtas As Single, Comi As Single, IDVend
    
    IDVend = PC.GetIDCuenta(lstSocyEmp)
    Vtas = DB.SumarProducto("Ventas", "Cantidad", "Precio", _
        "Fecha BETWEEN #" + stFechaSQL(DTPickerDe) + " 00:00# AND #" + _
        stFechaSQL(DTPickerA) + " 23:59# AND ID3Vendedor = " + CStr(IDVend))
    Comi = Vtas * CSng(txtComision) / 100
    
    txtVentas = FormatCurrency(Vtas, , , , vbFalse)
    txtComDebe = FormatCurrency(Comi, , , , vbFalse)
End Sub

Private Sub cmdAcreditar_Click()
    Dim IdCz As Long
    If EsCero(txtComDebe) = True Then Exit Sub
    txtComision = ValidarNumeros(txtComision)
    
    'le pregunto
    If MsgBox("¿Desea acreditar " + txtComDebe + " a " + lstSocyEmp + _
        " por comisiones por venta?", vbInformation + vbYesNo) = vbNo Then Exit Sub
        
    'ya basta de vueltas se los anoto
    IdCz = PC.GetIDCuenta(lstSocyEmp)
    
    'es indiferente socios y empleados estan en el mismo nivel
    DB.EXECUTE "INSERT INTO MovSocyEmp (ID,Fecha, IdNivel3," + _
        "Tipo,Variacion,Detalle) VALUES (" + IdAutonum("MovSocyEmp") + _
        ",#" + stFechaSQL(Date) + "#," + CStr(IdCz) + ",'Empleado'," + _
        Replace(CStr(CSng(txtComDebe)), ",", ".") + _
        ",'Comisiones " + FormatPercent(CSng(txtComision) / 100) + _
        " " + CStr(DTPickerDe) + " al " + CStr(DTPickerA) + "')"
    
    PC.Asiento PC.GetIDCuenta("Sueldo " + lstSocyEmp), txtComDebe, _
         PC.GetIDCuenta(lstSocyEmp), CStr(txtComDebe), , _
         "Comisiones Entre el " + CStr(DTPickerDe) + " al " + CStr(DTPickerA)
    
    MsgBox "Se ha realizado el registro correctamente", vbInformation, "Registro Completo"
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

Private Sub cmdVer_Click()
    VerResultados
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Dim IdCuentas() As String, I As Long
    
    IdCuentas = PC.GetCuentas(53)
    
    lstSocyEmp.Clear
    
    For I = 1 To UBound(IdCuentas)
        lstSocyEmp.AddItem PC.GetNameCuenta(CLng(IdCuentas(I)))
    Next I
        
    lstSocyEmp.ListIndex = 0
    
    DTPickerDe = Date - Day(Date) + 1 'primer dia del mes
    DTPickerA = Date
    
    lstSocyEmp_Click
End Sub

Private Sub lstSocyEmp_Click()
    If lstSocyEmp.ListCount > 0 Then 'hago el resumen
        lblEmpleado = lstSocyEmp
        Dim IdCF As Long
        
        cmdComision.Enabled = True
        
        IdCF = CFG.GetID("Comision " + lstSocyEmp)
        txtComision = NoNuloN(CFG.GetInfo(IdCF, 4))
    Else
        cmdComision.Enabled = False
    End If
    
    VerResultados
End Sub

Private Sub txtComision_GotFocus()
    PintarTxt txtComision
End Sub

Private Sub txtComision_LostFocus()
    txtComision = ValidarNumeros(txtComision)
    
    VerResultados
End Sub
