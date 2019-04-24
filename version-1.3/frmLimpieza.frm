VERSION 5.00
Begin VB.Form frmLimpieza 
   BackColor       =   &H00544B45&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Limpieza Base de Datos"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLimpieza.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   11670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCarpImg 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cambiar"
      Height          =   435
      Left            =   9285
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   4170
      Width           =   1500
   End
   Begin VB.CommandButton cmdUbicBDC 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cambiar"
      Height          =   435
      Left            =   9250
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   1860
      Width           =   1500
   End
   Begin VB.CommandButton cmdUbicBDP 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cambiar"
      Height          =   435
      Left            =   9250
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   720
      Width           =   1500
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00D5EDFB&
      Caption         =   "REINICIAR SISTEMAS"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   9180
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   6150
      Width           =   1455
   End
   Begin VB.CommandButton cmdCambiarFld 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cambiar"
      Height          =   435
      Left            =   9250
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3180
      Width           =   1500
   End
   Begin VB.CommandButton cmdBackUp 
      Caption         =   "Hacer Backup"
      Height          =   465
      Index           =   2
      Left            =   7710
      TabIndex        =   7
      Top             =   2580
      Width           =   1395
   End
   Begin VB.CommandButton cmdBackUp 
      Caption         =   "Hacer Backup"
      Height          =   465
      Index           =   1
      Left            =   7680
      TabIndex        =   4
      Top             =   1290
      Width           =   1395
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00544B45&
      Caption         =   "Accesos"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1125
      Left            =   3540
      TabIndex        =   29
      Top             =   6990
      Width           =   5085
      Begin VB.ComboBox cmbDias3 
         Height          =   360
         ItemData        =   "frmLimpieza.frx":030A
         Left            =   1110
         List            =   "frmLimpieza.frx":031D
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   360
         Width           =   1005
      End
      Begin VB.CommandButton cmdAccesos 
         Caption         =   "Borrar Movimientos Usuarios"
         Height          =   435
         Left            =   2280
         TabIndex        =   34
         Top             =   360
         Width           =   2445
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Días Atrás"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   150
         TabIndex        =   36
         Top             =   420
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00544B45&
      Caption         =   "Contabilidad"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1755
      Left            =   3540
      TabIndex        =   28
      Top             =   5130
      Width           =   5115
      Begin VB.CommandButton cmdContabilidad 
         Caption         =   "Resumir Asientos"
         Height          =   405
         Left            =   2370
         TabIndex        =   33
         Top             =   960
         Width           =   2445
      End
      Begin VB.ComboBox cmbDias2 
         Height          =   360
         ItemData        =   "frmLimpieza.frx":0337
         Left            =   1140
         List            =   "frmLimpieza.frx":034A
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   690
         Width           =   1005
      End
      Begin VB.CommandButton cmdCierres 
         Caption         =   "Borrar Registros de Cierres"
         Height          =   405
         Left            =   2370
         TabIndex        =   30
         Top             =   390
         Width           =   2445
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Días Atrás"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   180
         TabIndex        =   32
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00544B45&
      Caption         =   "Principal"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   2955
      Left            =   420
      TabIndex        =   21
      Top             =   5130
      Width           =   2955
      Begin VB.CommandButton cmdSocyEmp 
         Caption         =   "Resumir a Saldos Únicamente Cuentas de Socios y Empleados"
         Height          =   585
         Left            =   180
         TabIndex        =   27
         Top             =   2130
         Width           =   2625
      End
      Begin VB.CommandButton cmdSaldo0 
         Caption         =   "Borrar Detalles de Clientes y Proveedores con Saldo 0"
         Height          =   615
         Left            =   210
         TabIndex        =   26
         Top             =   1320
         Width           =   2625
      End
      Begin VB.ComboBox cmbDias 
         Height          =   360
         ItemData        =   "frmLimpieza.frx":0364
         Left            =   270
         List            =   "frmLimpieza.frx":0377
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   750
         Width           =   1005
      End
      Begin VB.CommandButton cmdVentas 
         Caption         =   "Borrar Detalles de Compras y Ventas "
         Height          =   765
         Left            =   1380
         TabIndex        =   23
         Top             =   420
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Días Atrás"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   270
         TabIndex        =   25
         Top             =   420
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdRestaurar 
      BackColor       =   &H00E8D5D8&
      Caption         =   "Restaurar"
      Height          =   465
      Index           =   2
      Left            =   9250
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2580
      Width           =   1500
   End
   Begin VB.CommandButton cmdRestaurar 
      BackColor       =   &H00E8D5D8&
      Caption         =   "Restaurar"
      Height          =   465
      Index           =   1
      Left            =   9250
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1290
      Width           =   1500
   End
   Begin VB.CommandButton cmdRestaurar 
      BackColor       =   &H00E8D5D8&
      Caption         =   "Restaurar"
      Height          =   465
      Index           =   0
      Left            =   9250
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   180
      Width           =   1500
   End
   Begin VB.CommandButton cmdBackUp 
      Caption         =   "Hacer Backup"
      Height          =   465
      Index           =   0
      Left            =   7680
      TabIndex        =   1
      Top             =   180
      Width           =   1395
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   60
   End
   Begin VB.CommandButton cmdCompactarA 
      Caption         =   "Compactar y Reparar"
      Height          =   495
      Left            =   5790
      TabIndex        =   6
      Top             =   2580
      Width           =   1820
   End
   Begin VB.CommandButton cmdCompactarC 
      Caption         =   "Compactar y Reparar"
      Height          =   495
      Left            =   5760
      TabIndex        =   3
      Top             =   1290
      Width           =   1820
   End
   Begin VB.CommandButton cmdCompactar 
      Caption         =   "Compactar y Reparar"
      Height          =   495
      Left            =   5760
      TabIndex        =   0
      Top             =   180
      Width           =   1820
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   405
      Left            =   9420
      TabIndex        =   10
      Top             =   7650
      Width           =   1125
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ubicación Carpeta de Imágenes"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   450
      TabIndex        =   48
      Top             =   4230
      Width           =   3615
   End
   Begin VB.Label lblCarpImg 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4170
      TabIndex        =   47
      Top             =   4200
      Width           =   4905
   End
   Begin VB.Label lblUbicBDC 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4140
      TabIndex        =   45
      Top             =   1890
      Width           =   4905
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ubicación BD Contabilidad(¡Importante!)"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   420
      TabIndex        =   44
      Top             =   1920
      Width           =   3615
   End
   Begin VB.Label lblUbicBDP 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4140
      TabIndex        =   42
      Top             =   750
      Width           =   4905
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ubicación BD Principal (¡Importante!)"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   420
      TabIndex        =   41
      Top             =   780
      Width           =   3615
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Carpeta a Grabarse/Recuperarse Backups"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   390
      TabIndex        =   38
      Top             =   3240
      Width           =   3615
   End
   Begin VB.Label lblFldBackups 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4110
      TabIndex        =   37
      Top             =   3210
      Width           =   4905
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Herramientas particulares"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   465
      Left            =   2640
      TabIndex        =   22
      Top             =   4680
      Width           =   4425
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha último Backup"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   315
      Left            =   5970
      TabIndex        =   20
      Top             =   3750
      Width           =   1815
   End
   Begin VB.Label lblFecha2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   7920
      TabIndex        =   19
      Top             =   3690
      Width           =   3345
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tamaño de Base de datos Accesos"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   510
      TabIndex        =   18
      Top             =   2670
      Width           =   3435
   End
   Begin VB.Label lblKBsA 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   4140
      TabIndex        =   17
      Top             =   2550
      Width           =   1545
   End
   Begin VB.Label lblKBsC 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   4140
      TabIndex        =   16
      Top             =   1290
      Width           =   1545
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tamaño de Base de datos Contabilidad"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   540
      TabIndex        =   15
      Top             =   1410
      Width           =   3435
   End
   Begin VB.Label lblFecha 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2430
      TabIndex        =   14
      Top             =   3660
      Width           =   3345
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha última limpieza"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   90
      TabIndex        =   13
      Top             =   3750
      Width           =   2205
   End
   Begin VB.Label lblKBs 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   4140
      TabIndex        =   12
      Top             =   150
      Width           =   1545
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tamaño de Base de datos Principal"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   510
      TabIndex        =   11
      Top             =   270
      Width           =   3435
   End
End
Attribute VB_Name = "frmLimpieza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DB As New tbrBasedeDatos.clsDataBase
Dim PC As New tbrPruebaCont.clsPruebaContab
Dim ACC As New tbrAccesos.clsTbrAccesos
Dim CFG As New tbrArbol.clstbrArbol

Dim AP As String, WF As String, PFl As String

Dim FSO As New Scripting.FileSystemObject
Dim archDB As String, archDBCont As String, archDBAcc As String, mArchivo As String

Private Sub cmdAccesos_Click()
    If MsgBox("¿Está seguro de realizar este proceso?", _
        vbInformation + vbOKCancel, "Herramientas Particulares") = vbCancel Then Exit Sub
    
    frmWAIT.Show
    frmWAIT.Refresh
    Timer1.Interval = 500 'medio segundo
    
    ACC.LimpiarMov CLng(cmbDias3)
        
    MsgBox "El proceso se realizó exitosamente", vbInformation, "Proceso finalizado"
    
    frmWAIT.Hide
    Timer1.Interval = 0
    
    MostrarDatos
End Sub

Private Sub cmdBackUp_Click(Index As Integer)
    Dim Resp As Long
    
    If FSO.FolderExists(lblFldBackups) = False Then
        MsgBox "Seleccione una carpeta correcta a donde se grabará el backup", vbInformation, "Atención"
        Exit Sub
    End If
    
    frmWAIT.Show
    frmWAIT.Refresh
    Timer1.Interval = 500 'medio segundo
    
    Select Case Index
        Case 0 'principal
            Resp = DB.Backup(lblUbicBDP + "dmbou.dv", lblFldBackups)
            Resp = DB.Backup(lblUbicBDP + "config.abl", lblFldBackups)
            Resp = DB.Backup(AP + "configBD.abl", lblFldBackups)
        Case 1 'contabilidad
            Resp = DB.Backup(lblUbicBDC + "ctas.mdb", lblFldBackups)
        Case 2 'accesos
            Resp = DB.Backup(AP + "acc.moe", lblFldBackups)
    End Select
    
    If Resp = 0 Then
        MsgBox "La copia se realizó correctamente", vbInformation, "Backup Realizado"
    Else
        MsgBox "Se detectaron fallas en la copia de seguridad." + vbCrLf + _
            "No se realizó la copia de seguridad (" + CStr(Resp) + ")", vbInformation, "Atención"
    End If
    
    Timer1.Interval = 0
    frmWAIT.Hide
    
    GrabarFecha True
    MostrarDatos
End Sub

Private Sub cmdCambiarFld_Click()
    Dim CM As New CommonDialog
    Dim PathArch As String
    
    CM.InitDir = ""
    CM.ShowFolder
    PathArch = CM.InitDir
    
    If PathArch <> "" Then
        If Right(PathArch, 1) <> "\" Then PathArch = PathArch + "\"
        
        If PathArch = WF Then
            MsgBox "No se permite elegir esta carpeta, seleccione otra", vbInformation, "Atención"
            cmdCambiarFld_Click
            Exit Sub
        End If
        
        lblFldBackups = PathArch
    End If
    
    Set CM = Nothing
End Sub

Private Sub cmdCarpImg_Click()
    Dim CM As New CommonDialog, IdCF As String
    Dim PathArch As String, FSO As New FileSystemObject
    
    CM.InitDir = ""
    CM.ShowFolder
    PathArch = CM.InitDir
    
    If PathArch = "" Then Exit Sub
    
    If Right(PathArch, 1) <> "\" Then PathArch = PathArch + "\"
    
    'FSO.CopyFile lblUbicBDP + "dmbou.dv", PathArch + "dmbou.dv", True
    MsgBox "La nueva carpeta fue cambiada De " + lblCarpImg + " a " + PathArch, vbInformation, "Atención"
    
    If FSO.FolderExists(PathArch + "img\") = False Then
        FSO.CreateFolder PathArch + "img\"
    End If
    
    lblCarpImg = PathArch
    
    IdCF = CFG.ExistePropiedadID(82)
    If IdCF = "" Then
        CFG.AgregarNodo 0, "UbicacionCarpImagenes", "", PathArch, 0, 82
    Else
        CFG.ModificarNodo 82, , , , PathArch
    End If
    
    Set CM = Nothing
    Set FSO = Nothing
End Sub

Private Sub cmdCierres_Click()
    If MsgBox("¿Está seguro de realizar este proceso?", _
        vbInformation + vbOKCancel, "Herramientas Particulares") = vbCancel Then Exit Sub
    
    frmWAIT.Show
    frmWAIT.Refresh
    Timer1.Interval = 500 'medio segundo
    
    PC.LimpiarCierres CLng(cmbDias2)
    
    frmWAIT.Hide
    Timer1.Interval = 0
    
    MsgBox "El proceso se realizó exitosamente", vbInformation, "Proceso finalizado"
    MostrarDatos
End Sub

Private Sub cmdCompactar_Click()
    'desconectar
    DB.CN_CLOSE
    Dim Tmp As Long
    
    frmWAIT.Show
    frmWAIT.Refresh
    Timer1.Interval = 100
    
    Tmp = DB.CompactarBASE(archDB, "chapita15")
    
    If Tmp <> 0 Then
        MsgBox "No se pudo compactar. Reintente luego de reiniciar el sistema." + _
            vbCrLf + CStr(Tmp), vbExclamation
    Else
        MsgBox "Se compacto OK", vbInformation
    End If
    
    frmWAIT.Hide
    Timer1.Interval = 0
    
    
    GrabarFecha False
    MostrarDatos
    
    'conectar!!!!!!!!!!!!!!!!!!!!!!!!!!!
    DB.cn_CONECTAR_MDB archDB, "chapita15"
End Sub

Private Sub cmdCompactarA_Click()
    ACC.Desconectar
    Dim Tmp As Long
    
    frmWAIT.Show
    frmWAIT.Refresh
    Timer1.Interval = 100
    
    Tmp = DB.CompactarBASE(archDBAcc, "zuliani")
    
    If Tmp <> 0 Then
        MsgBox "No se pudo compactar. Reintente luego de reiniciar el sistema." + _
            vbCrLf + Tmp, vbExclamation
    Else
        MsgBox "Se compacto OK", vbInformation
    End If
    
    frmWAIT.Hide
    Timer1.Interval = 0
    
    GrabarFecha False
    
    ACC.DBFilename2 = AP + "acc.moe"
    ACC.Conectar
    cmbDias.ListIndex = 0
     
    MostrarDatos
End Sub

Private Sub cmdCompactarC_Click()
    PC.CN_CLOSE
    Dim Tmp As Long
    
    frmWAIT.Show
    frmWAIT.Refresh
    Timer1.Interval = 100
    
    Tmp = DB.CompactarBASE(archDBCont, "zuliani")
    
    If Tmp <> 0 Then
        MsgBox "No se pudo compactar. Reintente luego de reiniciar el sistema." + _
            vbCrLf + CStr(Tmp), vbExclamation
    Else
        MsgBox "Se compacto OK", vbInformation
    End If
    
    frmWAIT.Hide
    Timer1.Interval = 0
    
    
    PC.PSW = "zuliani"
    PC.ArchMDB = lblUbicBDC + "Ctas.mdb"
    PC.Conectar
    
    MostrarDatos
End Sub

Private Sub cmdContabilidad_Click()
    If MsgBox("¿Está seguro de realizar este proceso?", _
        vbInformation + vbOKCancel, "Herramientas Particulares") = vbCancel Then Exit Sub
    
    frmWAIT.Show
    frmWAIT.Refresh
    Timer1.Interval = 500 'medio segundo
    
    PC.ResumirAsientos CLng(cmbDias2)
    PC.ResumirAsientos CLng(cmbDias2), "LibroSubDiario"
        
    frmWAIT.Hide
    Timer1.Interval = 0
    
    MsgBox "El proceso se realizó exitosamente", vbInformation, "Proceso finalizado"
    
    MostrarDatos
End Sub

Private Sub cmdRestaurar_Click(Index As Integer)
    Dim TmpArchivo As String, TmpDest As String, TmP2 As Long
    
    If FSO.FolderExists(lblFldBackups) = False Then
        MsgBox "Seleccione una carpeta correcta de donde se pueda recuperar el backup", vbInformation, "Atención"
        Exit Sub
    End If
    
    frmWAIT.Show
    frmWAIT.Refresh
    Timer1.Interval = 500 'medio segundo
    
    Select Case Index
        Case 0 'general
            TmpArchivo = lblFldBackups + "dmbou.dv"
            TmpDest = lblUbicBDP
        Case 1 'contabilidad
            TmpArchivo = lblFldBackups + "ctas.mdb"
            TmpDest = lblUbicBDC
        Case 2 'accesos
            TmpArchivo = lblFldBackups + "acc.moe"
            TmpDest = AP
    End Select
    
    TmP2 = DB.RestaurarBackup(TmpArchivo, TmpDest)
    
    Select Case TmP2
        Case 0
            MsgBox "La base de datos se ha restaurado exitosamente", vbInformation, "Archivo Restaurado"
        Case 1
            MsgBox "No se ha encontrado la copia de seguridad " + _
                "de la base de datos", vbInformation, "Atención"
        Case 3
            MsgBox "Procedimiento cancelado", vbInformation, "Atención"
        Case Else
            MsgBox "Se detectaron fallas restaurando la copia de seguridad." + vbCrLf + _
            "No se realizó correctamente", vbInformation, "Atención"
    End Select
    
    MostrarDatos
    
    frmWAIT.Hide
    Timer1.Interval = 0
    
    GrabarFecha True
    
End Sub

Private Sub cmdSaldo0_Click()
    
    If MsgBox("¿Está seguro de realizar este proceso?", _
        vbInformation + vbOKCancel, "Herramientas Particulares") = vbCancel Then Exit Sub
    
    frmWAIT.Show
    frmWAIT.Refresh
    Timer1.Interval = 500 'medio segundo
    
    Dim RS0 As New ADODB.Recordset
    Dim RS0a As New ADODB.Recordset
    
    RS0.Open "SELECT CodCliente From MovClientes " + _
        "GROUP BY CodCliente " + _
        "HAVING (((Sum(Variacion))=0))", DB.CN, adOpenStatic, adLockReadOnly
    RS0a.Open "SELECT MovProveedores.Proveedor FROM MovProveedores GROUP BY " + _
        "MovProveedores.Proveedor HAVING (((Sum(MovProveedores.Variacion))=0))", DB.CN, adOpenStatic, adLockReadOnly

    If RS0.RecordCount > 0 Then
        RS0.MoveFirst
        Do While Not RS0.EOF
            DB.Execute "DELETE FROM MovClientes WHERE CodCliente = " + _
                CStr(RS0("CodCliente"))
            RS0.MoveNext
        Loop
    End If
    
    If RS0a.RecordCount > 0 Then
        RS0a.MoveFirst
        Do While Not RS0a.EOF
            DB.Execute "DELETE FROM MovProveedores WHERE Proveedor = '" + _
                RS0a("Proveedor") + "'"
            RS0a.MoveNext
        Loop
    End If
    
    RS0.Close
    Set RS0 = Nothing
    RS0a.Close
    Set RS0a = Nothing
    
    MsgBox "El proceso se realizó exitosamente", vbInformation, "Proceso finalizado"
    
    frmWAIT.Hide
    Timer1.Interval = 0
    
    MostrarDatos
End Sub

Private Sub cmdSocyEmp_Click()
    If MsgBox("¿Está seguro de realizar este proceso?", _
        vbInformation + vbOKCancel, "Herramientas Particulares") = vbCancel Then Exit Sub
    Dim RsS As New ADODB.Recordset, Tipo As String, Saldo As Single
    
    frmWAIT.Show
    frmWAIT.Refresh
    Timer1.Interval = 500 'medio segundo
    
    RsS.Open "SELECT MovSocyEmp.Tipo,MovSocyEmp.IdNivel3, Sum(MovSocyEmp.Variacion) " + _
        "AS SumaDeVariacion FROM MovSocyEmp GROUP BY MovSocyEmp.Tipo, " + _
        "MovSocyEmp.IdNivel3", DB.CN, adOpenStatic, adLockReadOnly
    
    If RsS.RecordCount > 0 Then
        RsS.MoveFirst
        Do While Not RsS.EOF
            'grabo los datos
            Tipo = RsS("Tipo"): Saldo = RsS("SumadeVariacion")
            'borro todos los movimientos
            DB.Execute "DELETE FROM MovSocyEmp WHERE IdNivel3 = " + CStr(RsS("IdNivel3"))
            'agrego renglón con el saldo nomas
            DB.Execute "INSERT INTO MovSocyEmp (Fecha,Tipo,IdNivel3,Variacion,Detalle) " + _
                "VALUES (#" + stFechaSQL(Date) + "#,'" + Tipo + "'," + CStr(RsS("IdNivel3")) + _
                "," + Replace(CStr(Saldo), ",", ".") + ",'Resumen de Cuenta')"
            RsS.MoveNext
        Loop
    End If
        
    MsgBox "El proceso se realizó exitosamente", vbInformation, "Proceso finalizado"
    
    RsS.Close
    Set RsS = Nothing
    
    frmWAIT.Hide
    Timer1.Interval = 0
    
    MostrarDatos
End Sub

Private Sub cmdUbicBDC_Click()
    Dim CM As New CommonDialog, IdCF As String
    Dim PathArch As String, FSO As New FileSystemObject
    
    PC.CN_CLOSE
    
    CM.InitDir = ""
    CM.ShowFolder
    PathArch = CM.InitDir
    
    If PathArch = "" Then Exit Sub
    
    If Right(PathArch, 1) <> "\" Then PathArch = PathArch + "\"
    
    If FSO.FileExists(PathArch + "ctas.mdb") = False Then
        FSO.MoveFile lblUbicBDC + "ctas.mdb", PathArch + "ctas.mdb"
        MsgBox "El archivo fue movido De " + lblUbicBDP + " a " + PathArch, vbInformation, "Atención"
    Else
        MsgBox "Se cambió la configuración de BD contabilidad", vbInformation, "Atención"
    End If
    
    lblUbicBDC = PathArch
    
    IdCF = CFG.ExistePropiedadID(81)
    If IdCF = "" Then
        CFG.AgregarNodo 0, "DireccionBDCont", "", PathArch, 0, 81
    Else
        CFG.ModificarNodo 81, , , , PathArch
    End If
    
    archDB = lblUbicBDP + "dmbou.dv"
    archDBCont = lblUbicBDC + "Ctas.mdb"
    
    'conecto de vuelta
    PC.PSW = "zuliani"
    PC.ArchMDB = lblUbicBDC + "Ctas.mdb"
    PC.Conectar
    
    MostrarDatos
    Set CM = Nothing
    Set FSO = Nothing
End Sub

Private Sub cmdUbicBDP_Click()
    Dim CM As New CommonDialog, IdCF As String
    Dim PathArch As String, FSO As New FileSystemObject
    
    DB.CN_CLOSE
    CM.InitDir = ""
    CM.ShowFolder
    PathArch = CM.InitDir
    
    If PathArch = "" Then Exit Sub
    
    If Right(PathArch, 1) <> "\" Then PathArch = PathArch + "\"
    
    If FSO.FileExists(PathArch + "dmbou.dv") = False Then
        FSO.MoveFile lblUbicBDP + "dmbou.dv", PathArch + "dmbou.dv"
        MsgBox "El archivo fue movido De " + lblUbicBDP + " a " + PathArch, vbInformation, "Atención"
    Else
        MsgBox "Se cambió la configuración de BD Principal por " + vbrclf + _
            "otra existente en esa carpeta", vbInformation, "Atención"
    End If
    
    'la configuracion aparte
    If FSO.FileExists(PathArch + "config.abl") = False Then
        FSO.CopyFile lblUbicBDP + "config.abl", PathArch + "config.abl", True
    End If
    
    lblUbicBDP = PathArch
    IdCF = CFG.ExistePropiedadID(80)
    
    If IdCF = "" Then
        CFG.AgregarNodo 0, "DireccionBD", "", PathArch, 0, 80
    Else
        CFG.ModificarNodo 80, , , , PathArch
    End If
    
    archDB = lblUbicBDP + "dmbou.dv"
    archDBCont = lblUbicBDC + "Ctas.mdb"
    
    MostrarDatos
    
    'conectar DB de vuelta!!!!!!!!!!!!!!!!!!!!!!!!!!!
    DB.cn_CONECTAR_MDB archDB, "chapita15"
    
    Set CM = Nothing
    Set FSO = Nothing
End Sub

Private Sub cmdVentas_Click()
    If MsgBox("¿Está seguro de realizar este proceso?", _
        vbInformation + vbOKCancel, "Herramientas Particulares") = vbCancel Then Exit Sub
    
    frmWAIT.Show
    frmWAIT.Refresh
    Timer1.Interval = 500 'medio segundo
    
    DB.Execute "DELETE FROM Ventas WHERE Fecha < #" + _
        CStr(stFechaSQL(Date - CLng(cmbDias))) + "#"
        
    DB.Execute "DELETE FROM CompraDetalle WHERE Fecha < #" + _
        CStr(stFechaSQL(Date - CLng(cmbDias))) + "#"
        
    MsgBox "El proceso se realizó exitosamente", vbInformation, "Proceso finalizado"
    
    frmWAIT.Hide
    Timer1.Interval = 0
    
    MostrarDatos
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()

    Dim FFS As New Scripting.FileSystemObject
    Dim B1 As String, B2 As String, B3 As String
    
    'PRINCIPAL!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    B1 = lblUbicBDP + "dmbou.dv"
    'CONTABILIDAD!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    B2 = lblUbicBDC + "ctas.mdb"
    'ACCESOS!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    B3 = AP + "acc.moe"
    
    'mete las bases de datos en otro lado y carga las vacias!
    
    Dim RES As VbMsgBoxResult
    RES = MsgBox("¿ Las bases anteriores se borran definitivamente ?" + vbCrLf + _
        "Si coloca no se reiniciará resguardando las bases.", _
        vbCritical + vbQuestion + vbYesNoCancel, "CUIDADO!!!!")
        
    If RES = vbYes Then
        'desconectarlas
        DB.CN_CLOSE
        PC.CN_CLOSE
        ACC.Desconectar
        'borrar las bases
        FFS.DeleteFile B1, True
        FFS.DeleteFile B2, True
        FFS.DeleteFile B3, True
        'copiar las vacias
        FFS.CopyFile AP + "reini\dmbou.dv", lblUbicBDP
        FFS.CopyFile AP + "reini\ctas.mdb", lblUbicBDC
        FFS.CopyFile AP + "reini\acc.moe", AP
    End If
    
    If RES = vbNo Then
        'desconectarlas
        DB.CN_CLOSE
        PC.CN_CLOSE
        ACC.Desconectar
        'se copian a otro lado
        If FFS.FolderExists(WF + "clean\") Then FFS.CreateFolder WF + "clean\"
        
        FFS.MoveFile B1, WF + "clean\b1.x"
        FFS.MoveFile B2, WF + "clean\b2.x"
        FFS.MoveFile B3, WF + "clean\b3.x"
        
        'copiar las vacias
        FFS.CopyFile AP + "reini\dmbou.dv", lblUbicBDP
        FFS.CopyFile AP + "reini\ctas.mdb", lblUbicBDC
        FFS.CopyFile AP + "reini\acc.moe", AP
    End If
    
    If RES = vbCancel Then Exit Sub
    
    MsgBox "Los sistemas se reiniciaron!"
    Unload Me
    End
End Sub

Private Sub Form_Load()
    Dim Tmp As Long, CM As New CommonDialog, PathArch As String, IdCF As String

    AP = App.path
    WF = FSO.GetSpecialFolder(WindowsFolder)
    PFl = FSO.GetSpecialFolder(SystemFolder)
    
    If Right(AP, 1) <> "\" Then AP = AP + "\"
    If Right(WF, 1) <> "\" Then WF = WF + "\"
    
    'conectar!!!!!!!!!!!!!!!!!!!!!!!!!!!
    'ACCESOS!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    ACC.DBFilename2 = AP + "acc.moe"
    ACC.Conectar
    
    'CONFIGURACION !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    CFG.Archivo = AP + "configBD.abl"
    
    lblUbicBDP = CFG.GetInfo(80, 4)
    lblUbicBDC = CFG.GetInfo(81, 4)
    lblCarpImg = CFG.GetInfo(82, 4)
    
    Tmp = 1
    'CONTROLO QUE EXISTA BD PRINCIPAL ------------------------------------------------
    PathArch = lblUbicBDP
    Do While Tmp < 4
        If FSO.FileExists(PathArch + "dmbou.dv") = False Then
            MsgBox "No se encuentra la Base de Datos PRINCIPAL en " + _
                PathArch + vbCrLf + _
                "Seleccione la carpeta correcta", vbExclamation, "Atención"
            
            CM.InitDir = ""
            CM.ShowFolder
            PathArch = CM.InitDir
            
            If PathArch <> "" Then
                If Right(PathArch, 1) <> "\" Then PathArch = PathArch + "\"
                
                If FSO.FileExists(PathArch + "dmbou.dv") = False Then
                    Tmp = Tmp + 1
                Else
                    lblUbicBDP = PathArch
                    IdCF = CFG.ExistePropiedadID(80)
                    If IdCF = "" Then
                        CFG.AgregarNodo 0, "DireccionBD", "", PathArch, 0, 80
                    Else
                        CFG.ModificarNodo 80, , , , PathArch
                    End If
                    Tmp = 100
                End If
            Else
                Tmp = Tmp + 1
            End If
        Else
            Tmp = 100
        End If
    Loop
    If Tmp <> 100 Then
        MsgBox "El sistema se cerrará por no encontrarse BD Principal", vbCritical, "Atención"
        Unload Me
        Exit Sub
    End If
    ' ---------------------------------------------------------------------------------
    Tmp = 1
    'CONTROLO QUE EXISTA BD CONTABILIDAD ---------------------------------------------
    PathArch = lblUbicBDC
    Do While Tmp < 4
        If FSO.FileExists(PathArch + "ctas.mdb") = False Then
            MsgBox "No se encuentra la Base de Datos CONTABILIDAD en " + _
                PathArch + vbCrLf + _
                "Seleccione la carpeta correcta", vbExclamation, "Atención"
            
            CM.InitDir = ""
            CM.ShowFolder
            PathArch = CM.InitDir
            
            If PathArch <> "" Then
                If Right(PathArch, 1) <> "\" Then PathArch = PathArch + "\"
                
                If FSO.FileExists(PathArch + "ctas.mdb") = False Then
                    Tmp = Tmp + 1
                Else
                    lblUbicBDC = PathArch
                    IdCF = CFG.ExistePropiedadID(81)
                    If IdCF = "" Then
                        CFG.AgregarNodo 0, "DireccionBDCont", "", PathArch, 0, 81
                    Else
                        CFG.ModificarNodo 81, , , , PathArch
                    End If
                    Tmp = 100
                End If
            Else
                Tmp = Tmp + 1
            End If
        Else
            Tmp = 100
        End If
    Loop
    If Tmp <> 100 Then
        MsgBox "El sistema se cerrará por no encontrarse BD Contabilidad", vbCritical, "Atención"
        Unload Me
        Exit Sub
    End If
    ' ---------------------------------------------------------------------------------
    Tmp = 1
    'CONTROLO QUE EXISTA LA CARPETA IMAGEN ------------------------------------------------
    PathArch = lblCarpImg
    Do While Tmp < 4
        If FSO.FolderExists(PathArch) = False Then
            MsgBox "Path para carpeta de imagen no existe no existe seleccione otra " + _
                PathArch + vbCrLf + _
                "Seleccione la carpeta correcta", vbExclamation, "Atención"
            
            CM.InitDir = ""
            CM.ShowFolder
            PathArch = CM.InitDir
            
            If PathArch <> "" Then
                If Right(PathArch, 1) <> "\" Then PathArch = PathArch + "\"
                
                If FSO.FolderExists(PathArch) = False Then
                    Tmp = Tmp + 1
                Else
                    If FSO.FolderExists(PathArch + "img\") = False Then
                        FSO.CreateFolder PathArch + "img\"
                    End If
                    
                    lblUbicBDP = PathArch
                    
                    IdCF = CFG.ExistePropiedadID(82)
                    If IdCF = "" Then
                        CFG.AgregarNodo 0, "UbicacionCarpImagenes", "", PathArch, 0, 82
                    Else
                        CFG.ModificarNodo 82, , , , PathArch
                    End If
                    Tmp = 100
                End If
            Else
                Tmp = Tmp + 1
            End If
        Else
            Tmp = 100
        End If
    Loop
    If Tmp <> 100 Then
        MsgBox "El sistema se cerrará por no encontrarse BD", vbCritical, "Atención"
        Unload Me
        Exit Sub
    End If
    ' ---------------------------------------------------------------------------------
    
    cmbDias.ListIndex = 0
    cmbDias2.ListIndex = 0
    cmbDias3.ListIndex = 0
   
    'BASE DE DATOS!!!!!!!!!!!!!!!!!!!!!!!!!!!
    DB.cn_CONECTAR_MDB CFG.GetInfo(80, 4) + "dmbou.dv", "chapita15"
   
    'CONTABILIDAD!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    PC.ArchMDB = CFG.GetInfo(81, 4) + "Ctas.mdb"
    PC.PSW = "zuliani"
    PC.Conectar
           
    archDB = lblUbicBDP + "dmbou.dv"
    archDBCont = lblUbicBDC + "Ctas.mdb"
    archDBAcc = AP + "acc.moe"
    mArchivo = AP + "LimpFecha.lpz"
   
    Set CM = Nothing
    MostrarDatos
End Sub

Private Sub MostrarDatos()
    mArchivo = AP + "LimpFecha.lpz"
    Dim spFh() As String
    
    spFh = Split(LeerFecha, "|")
    
    lblKBs = MostrarTamanoDeArchivo(archDB)
    lblKBsC = MostrarTamanoDeArchivo(archDBCont)
    lblKBsA = MostrarTamanoDeArchivo(archDBAcc)
    lblFecha = spFh(0)
    lblFecha2 = spFh(1)
    
    'algunas versiones no la tienen entonces
    If UBound(spFh) > 1 Then 'si esta
        lblFldBackups = spFh(2)
    Else
        lblFldBackups = PFl 'ya la próxima queda grabada bien
    End If
    
    If FSO.FolderExists(lblFldBackups) = False Then
        lblFldBackups = PFl
    End If
End Sub

Function MostrarTamanoDeArchivo(EspecificacionDeArchivo)
    Dim F, S As String
    
    If FSO.FileExists(EspecificacionDeArchivo) = False Then
      S = "0  bytes"
      MostrarTamanoDeArchivo = S
      Exit Function
    End If
    
    Set F = FSO.GetFile(EspecificacionDeArchivo)

    Select Case F.Size
      Case Is < 1024
          S = F.Size & " bytes"
      Case Is >= 1024, Is < 1048576
          S = F.Size / 1024 & " KBytes"
      Case Is >= 1048576
          S = F.Size / 1048576 & " GBytes"
    End Select
    
    MostrarTamanoDeArchivo = S
    
    Set F = Nothing

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload frmWAIT
    DB.CN_CLOSE
    ACC.Desconectar
    PC.CN_CLOSE
    
    Set PC = Nothing
    Set ACC = Nothing
    Set DB = Nothing
    Set FSO = Nothing
End Sub

Private Sub Timer1_Timer()
    Dim C As Long
    C = C + 1
    If C = 3 Then C = 0
    frmWAIT.Label1 = "Procesando" + String(C, ".")
End Sub

Private Sub GrabarFecha(EsBackup As Boolean)
    'la fecha esta de esta forma fechaLimpieza|fechaBackUp|UltPathBackupUsado
    Dim TE As TextStream
    Dim Hoy As String, Spl() As String, tmpG As String
    
    Hoy = Format(Date, "long date")
       
    If FSO.FileExists(mArchivo) = False Then
        
        Set TE = FSO.CreateTextFile(mArchivo, True)
            TE.WriteLine Hoy + "|" + Hoy
        TE.Close
    End If
    
    Dim tmpS As String
    
    Set TE = FSO.OpenTextFile(mArchivo, ForReading, True)
        tmpS = TE.ReadLine
    TE.Close
    
    Spl = Split(tmpS, "|")
    
    If EsBackup Then
        tmpG = Spl(0) + "|" + Hoy
    Else
        tmpG = Hoy + "|" + Spl(1)
    End If
    
    'le agrego la carpeta de backups siempre
    
    tmpG = tmpG + "|" + lblFldBackups
        
    Set TE = FSO.OpenTextFile(mArchivo, ForWriting, True)
        TE.WriteLine tmpG
    TE.Close
    
    Set TE = Nothing
    
End Sub

Private Function LeerFecha() As String
    Dim TE As TextStream
    Dim FSO As New Scripting.FileSystemObject
    Dim Hoy As String
    
    Hoy = Format(Date, "long date")
    
    If FSO.FileExists(mArchivo) = False Then
        
        Set TE = FSO.CreateTextFile(mArchivo, True)
            TE.WriteLine Hoy + "|" + Hoy + "|" + PFl
        TE.Close
    End If
    
    Set TE = FSO.OpenTextFile(mArchivo, ForReading, True)
        LeerFecha = TE.ReadLine
    TE.Close
    
    Set TE = Nothing
End Function


'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
Private Function stFechaSQL(FECHA As Date) As String
    Dim FechaChota As String    'sql tiene la fecha al reves por eso
    Dim h() As String
    h = Split(CStr(FECHA), "/")
    FechaChota = h(1) + "/" + h(0) + "/" + h(2)
    stFechaSQL = FechaChota
End Function
