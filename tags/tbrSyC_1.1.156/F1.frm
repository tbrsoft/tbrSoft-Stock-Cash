VERSION 5.00
Begin VB.Form F1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   9240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frAll 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3825
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   9165
      Begin VB.Timer Timer1 
         Left            =   330
         Top             =   5310
      End
      Begin VB.PictureBox shBAR 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   75
         Left            =   6120
         ScaleHeight     =   75
         ScaleWidth      =   135
         TabIndex        =   1
         Top             =   2700
         Width           =   135
      End
      Begin VB.Image Image1 
         Height          =   3780
         Left            =   30
         Picture         =   "F1.frx":0000
         Top             =   30
         Width           =   9120
      End
   End
End
Attribute VB_Name = "F1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Salir As Boolean
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Salir = True
End Sub

Private Sub Form_Load()
    On Local Error GoTo ErrLoad
    AP = App.path
    If Right(AP, 1) <> "\" Then AP = AP + "\"
    
    Terr.FileLog = AP + "regSyc.log"
    
    If Command = "e1" Then
        Terr.FileLogGrabaTodo = AP + "REG-SyC.W15"
        Terr.ModoGrabaTodo = True
        Terr.StartGrabaTodo
    End If
    
    Terr.LargoAcumula = 600
    
    Terr.AnOtaR "aaaa"
    
    frAll.Width = Image1.Width
    frAll.Height = Image1.Height
    Image1.Top = 0
    Image1.Left = 0
    frAll.Top = 0
    frAll.Left = 0
    Me.Width = frAll.Width
    Me.Height = frAll.Height
    
    Dim DifStk As Single, RFyT As Single, AjustesStock As Single, IDcierre As Long
    Dim FSO As New FileSystemObject
    Dim MovSB 'lo que cambia por c/proceso el show bar
    
    Terr.AppendSinHist "aaab" + vbCrLf + _
        CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision) + vbCrLf + _
        AP
        
    WF = FSO.GetSpecialFolder(WindowsFolder)
    If Right(WF, 1) <> "\" Then WF = WF + "\"
    
    frAll.Visible = True
    Me.Show
    Me.Refresh
    
    shBAR.Width = 200
    shBAR.Refresh
    
    'ver que licencia tiene
    LIC.GetDatosPC WF + "zaec112aa.tmy"
    LIC.PutArchLic WF + "klesoft.tes"
    'si no existe no importa, no le da bola
    
    Terr.AnOtaR "aaac"
    
    'ACCESOS!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    ACC.DBFilename2 = AP + "acc.moe"
    ACC.Conectar
    
    Terr.AnOtaR "aaad"
    Dim TTT As Long
    
    TTT = ACC.ValidarUsuario(5)
    
    Terr.AnOtaR "aaae", TTT
    
    If TTT < 0 Then
        If TTT = -1 Then MsgBox "Se cerrará el sistema por falta de validación", vbCritical, "Atención"
        Unload Me
        
        Exit Sub
    End If
    
    'CONFIGURACIONES!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    
    shBAR.Width = 500
    shBAR.Refresh
    
    CFGBD.ConfVigentes(1) = "1*VerDatosPantalla*Si"
    CFGBD.ConfVigentes(2) = "80*DireccionBD*" + WF
    CFGBD.ConfVigentes(3) = "81*DireccionBDCont*" + AP
    CFGBD.ConfVigentes(4) = "82*UbicacionCarpImagenes*" + AP
    CFGBD.ConfVigentes(5) = "13*Prox Nro Factura*A-0001-0103"
    CFGBD.Archivo = AP + "configBD.abl"
    
    'CONFIGURACIONES VIGENTES --(Stock n Cash) --------------------------------------------
    CFG.ConfVigentes(1) = "1*Titulo*tbrStock & Cash - Software Administración Pequeñas Empresas"
    CFG.ConfVigentes(2) = "2*Dias Vencimiento*30"
    CFG.ConfVigentes(3) = "3*IdSucursalPredeterminada*1"
    CFG.ConfVigentes(4) = "4*Tipos IVA*0"
    CFG.ConfVigentes(5) = "5*Usa Envases*No"
    CFG.ConfVigentes(6) = "6*Conf Cod Prod*No"
    CFG.ConfVigentes(7) = "7*IVA Pred*0"
    CFG.ConfVigentes(8) = "8*LetraFacturaCompraPred*C"
    CFG.ConfVigentes(9) = "9*Impresion Codigo*ID_Nombre_Precio"
    CFG.ConfVigentes(10) = "10*Calcular Vuelto*50"
    CFG.ConfVigentes(11) = "11*Interes Mensual Pred*3"
    CFG.ConfVigentes(12) = "12*Cantidad Maxima de Cuotas*36)"
    CFG.ConfVigentes(13) = "14*Vendedor Predeterminado*Larry O."
    CFG.ConfVigentes(14) = "15*Caracteres C.Barras*t8r" 'Sin uso lo dejo por las dudas
    CFG.ConfVigentes(15) = "16*Cantidad de Mov Accesos*30"
    CFG.ConfVigentes(16) = "17*Cantidad de Mov Productos*111"
    CFG.ConfVigentes(17) = "18*FechaBup*32_05/06/2007"
    CFG.ConfVigentes(18) = "19*StockMinimo*10"
    CFG.ConfVigentes(19) = "20*Comisiones*0"
    CFG.ConfVigentes(20) = "30*DatosTrabajoCliente*IDc_DetT_Ant_Fono_Ingr"
    CFG.ConfVigentes(21) = "40*FormadePago*CC"
    CFG.ConfVigentes(22) = "50*MgenVta*30"
    CFG.ConfVigentes(23) = "60*Dto Vta Contado*0"
    CFG.ConfVigentes(24) = "70*Ultimo Pago Financiera*1" 'dice el cliente que lo hizo
    CFG.ConfVigentes(25) = "90*Tipo Usuario*Servidor"
    CFG.ConfVigentes(26) = "95*Usa forma Contado*No"
    CFG.ConfVigentes(27) = "100*Depende del Pais*1"
    CFG.ConfVigentes(28) = "101*Facturas IVAC*AC"
    CFG.ConfVigentes(29) = "102*Facturas IVAV*AB"
    CFG.Archivo = CFGBD.GetInfo(80, 4) + "Config.abl"
    '--------------------------------------------------------------------------------
    
'    Terr.AppendLog "aaaf", CFG.ContarRenglones
'    'controlo que no este roto el archivo ---------------------------------------------
'    CFG.Archivo = CFGBD.GetInfo(80, 4) + "Config.abl"
'    If CFG.ContarRenglones < 5 Then 'esta roto
'        CFG.SeguridadRestaura
'    Else
'        CFG.SeguridadCopia
'    End If
'    '---------------------------------------------------------------------------------
    
    Terr.AnOtaR "aaag"
    ControlarBD
    
    shBAR.Width = 700
    shBAR.Refresh
    
    
    'BASE DE DATOS!!!!!!!!!!!!!!!!!!!!!!!!!!!
    ArchivoMDBPrincipal = CFGBD.GetInfo(80, 4) + "dmbou.dv"
    Terr.AnOtaR "aaah", ArchivoMDBPrincipal
    Contrasena = "chapita15"
    
    DB.cn_CONECTAR_MDB ArchivoMDBPrincipal, Contrasena
    
    Terr.AnOtaR "aaai"
    'CONTABILIDAD!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    PC.ArchMDB = CFGBD.GetInfo(81, 4) + "Ctas.mdb"
    PC.PSW = "zuliani"
    
    PC.Conectar
    Terr.AnOtaR "aaaj"
    ' TIPO USUARIO Servidor o Cliente -----------------------------------------------
    Dim TmP As String, IdCF As String
    
    IdCF = CFG.ExistePropiedadID(90)
    Terr.AnOtaR "aaak", IdCF
    
    If IdCF = "" Then
        TmP = "Servidor"
        CFG.AgregarNodo 0, "Tipo Usuario", "", TmP, 0, 90
    Else
        TmP = CFG.GetInfo(90, 4)
        If TmP = "" Or (TmP <> "Servidor" And TmP <> "Cliente") Then
            TmP = "Servidor"
            CFG.ModificarNodo 90, , , , TmP
        End If
    End If
    
    TipoUsuario = TmP
    Terr.AnOtaR "aaal", TmP
    '--------------------------------------------------------------------------------

    shBAR.Width = 900
    shBAR.Refresh
    'DB.Execute "UPDATE Productos SET Stock=50"
    'DB.Execute "UPDATE MovClientes SET IdVta=0"
    
    ' ------------------------------------------------------------------------------
    ' HAGO EL BACK UP? -------------------------------------------------------------
    Dim UltBUp() As String
    Terr.AnOtaR "aaam"
    UltBUp = Split(CFG.GetInfo(18, 4), "_")
    
    If Not IsDate(UltBUp(1)) Then
        UltBUp(1) = Date '"05/07/2007" la fecha de hoy (corregido por andres 1/2/08)
    End If
    
    If Not IsNumeric(UltBUp(0)) Then
        UltBUp(0) = "5"
    End If
    Terr.AnOtaR "aaan"
    If Date > ProximoMes(CDate(UltBUp(1))) Then 'paso mas de 1 mes
        
        If CLng(UltBUp(0)) >= 1 And CLng(UltBUp(0)) <= 30 Then
            Terr.AnOtaR "aaao"
            If MsgBox("Ya ha pasado más de un mes sin realizar Backup." + vbCrLf + _
                "¿Desea realizarlo ahora?", vbYesNo, "Atención") = vbNo Then
                If MsgBox("¿Desea que se le recuerde nuevamente?.", _
                    vbYesNo, "Atención") = vbNo Then
                    'grabo que no quiere hacer mas backups
                    
                    CFG.ModificarNodo 18, , , , "32_" + CStr(Date)
                Else
                    Terr.AnOtaR "aaap"
                    'nada!, le va a volvera a romper las bolas la proxima
                End If
            Else
                Terr.AnOtaR "aaaq"
                HacerCopiaSeguridadEs
                Terr.AnOtaR "aaar"
                CFG.ModificarNodo 18, , , , UltBUp(0) + "_" + CStr(Date)
            End If
        Else
            Terr.AnOtaR "aaas"
            'nada por que tiene elegido no hacer bups
        End If
        
    Else 'no paso 1 mes
        Terr.AnOtaR "aaat"
        'solo pregunto si hoy es el dia que corresponde
        If Day(Date) = CLng(UltBUp(0)) And Date <> CDate(UltBUp(1)) Then
            
            If MsgBox("Hoy le correspondería hacer el Backup." + vbCrLf + _
                "¿Desea realizarlo?", vbInformation + vbYesNo) = vbYes Then
                Terr.AnOtaR "aaau"
                HacerCopiaSeguridadEs
                Terr.AnOtaR "aaav"
                CFG.ModificarNodo 18, , , , UltBUp(0) + "_" + CStr(Date)
            End If
        End If
    End If
    
    ' ------------------------------------------------------------------------------
    shBAR.Width = 1100
    shBAR.Refresh
    
    MovSB = 200
    Dim T As Double
    Do While shBAR.Width < 2000
        T = Timer
        Do While T + 0.05 > Timer
            DoEvents
            If Salir Then GoTo fin
        Loop
        shBAR.Width = shBAR.Width + MovSB / 3
        shBAR.Refresh
    Loop
    Terr.AnOtaR "aaaw"
    If FSO.FolderExists(CFGBD.GetInfo(82, 4) + "IMG\") = False Then
        FSO.CreateFolder CFGBD.GetInfo(82, 4) + "IMG\"
    End If
    Terr.AnOtaR "aaax"
fin:
    Set FSO = Nothing
    Unload Me
    Unload frmWAIT
    fIni.Show
    Terr.AnOtaR "aaay"
    Exit Sub
    
ErrLoad:
    Terr.AppendLog Terr.ErrToTXT(Err), "LOAD INDEX"
    Resume Next
End Sub

Private Sub ControlarBD()
    Dim UbicBDP As String, UbicBDC As String
    Dim PathArch As String, TmP As Long, IdCF As String
    Dim FSO As New FileSystemObject, CM As New CommonDialog

    UbicBDP = CFGBD.GetInfo(80, 4)
    UbicBDC = CFGBD.GetInfo(81, 4)
    
    TmP = 1
    'CONTROLO QUE EXISTA BD PRINCIPAL ------------------------------------------------
    PathArch = UbicBDP
    Do While TmP < 4
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
                    TmP = TmP + 1
                Else
                    IdCF = CFGBD.ExistePropiedadID(80)
                    If IdCF = "" Then
                        CFGBD.AgregarNodo 0, "DireccionBD", "", PathArch, 0, 80
                    Else
                        CFGBD.ModificarNodo 80, , , , PathArch
                    End If
                    TmP = 100
                End If
            Else
                TmP = TmP + 1
            End If
        Else
            TmP = 100
        End If
    Loop
    If TmP <> 100 Then
        MsgBox "El sistema se cerrará por no encontrarse BD Principal", vbCritical, "Atención"
        CerraMe
        Exit Sub
    End If
    ' ---------------------------------------------------------------------------------
    TmP = 1
    'CONTROLO QUE EXISTA BD CONTABILIDAD ---------------------------------------------
    PathArch = UbicBDC
    Do While TmP < 4
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
                    TmP = TmP + 1
                Else
                    IdCF = CFGBD.ExistePropiedadID(81)
                    If IdCF = "" Then
                        CFGBD.AgregarNodo 0, "DireccionBDCont", "", PathArch, 0, 81
                    Else
                        CFGBD.ModificarNodo 81, , , , PathArch
                    End If
                    TmP = 100
                End If
            Else
                TmP = TmP + 1
            End If
        Else
            TmP = 100
        End If
    Loop
    If TmP <> 100 Then
        MsgBox "El sistema se cerrará por no encontrarse BD Contabilidad", vbCritical, "Atención"
        CerraMe
        Exit Sub
    End If
    ' ---------------------------------------------------------------------------------
    
    Set FSO = Nothing
End Sub

Private Sub CerraMe()
    ACC.Desconectar
    DB.CN_CLOSE
    PC.CN_CLOSE
    
    Set CFG = Nothing
    Set CFGBD = Nothing
    Set ACC = Nothing
    Set DB = Nothing
    Set PC = Nothing
    Set TP = Nothing
    Unload Me
    End
End Sub

Private Sub HacerCopiaSeguridadEs()
    Dim SP() As String, FSO As New Scripting.FileSystemObject
    
    SP = Split(LeerFecha, "|")
    
    If FSO.FolderExists(SP(2)) = False Then
        SP(2) = FSO.GetSpecialFolder(SystemFolder)
        Exit Sub
    End If
    
    If Right(SP(2), 1) <> "\" Then SP(2) = SP(2) + "\"
    
    frmWAIT.Show
    frmWAIT.Refresh
    Timer1.Interval = 500 'medio segundo
    
    'principal
    Resp = DB.Backup(CFGBD.GetInfo(80, 4) + "db.dv", SP(2))
    'contabilidad
    Resp = DB.Backup(CFGBD.GetInfo(81, 4) + "ctas.mdb", SP(2))
    'accesos
    Resp = DB.Backup(AP + "acc.moe", SP(2))
    'configuracion tambien
        'original (va siempre en la misma carpeta BDOrig)
    Resp = DB.Backup(CFGBD.GetInfo(80, 4) + "config.abl", SP(2))
        ' la que dice donde estan las bases de datos esta en APP
    Resp = DB.Backup(AP + "configBD.abl", SP(2))
    
    If Resp = 0 Then
        MsgBox "La copia se realizó correctamente en" + vbCrLf + _
            SP(2), vbInformation, "Backup Realizado"
    Else
        MsgBox "Se detectaron fallas en la copia de seguridad." + vbCrLf + _
            "No se realizó la copia de seguridad (" + CStr(Resp) + ")", vbInformation, "Atención"
    End If
    
    Timer1.Interval = 0
    'frmWAIT.Hide
    
    Set FSO = Nothing
End Sub

Private Function LeerFecha() As String
    Dim TE  As TextStream
    Dim FSO As New Scripting.FileSystemObject
    Dim Hoy As String, mArchivo As String
    
    Hoy = Format(Date, "long date")
    mArchivo = AP + "LimpFecha.lpz"
    
    If FSO.FileExists(mArchivo) = False Then
        Set TE = FSO.CreateTextFile(mArchivo, True)
            TE.WriteLine Hoy + "|" + Hoy + FSO.GetSpecialFolder(SystemFolder)
        TE.Close
    End If
    
    Set TE = FSO.OpenTextFile(mArchivo, ForReading, True)
        LeerFecha = TE.ReadLine
    TE.Close
    
    Set FSO = Nothing
    Set TE = Nothing
End Function

