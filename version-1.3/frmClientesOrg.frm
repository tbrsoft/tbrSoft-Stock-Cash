VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClientesOrg 
   BackColor       =   &H00544B45&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clientes con deudas"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8295
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmClientesOrg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin tbrFaroButton.fBoton cmdSalir 
      Height          =   495
      Left            =   6900
      TabIndex        =   8
      Top             =   6060
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   873
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdImprimir 
      Height          =   495
      Left            =   6960
      TabIndex        =   7
      Top             =   3330
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   873
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Imprimir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdIncobrables 
      Height          =   525
      Left            =   3840
      TabIndex        =   6
      Top             =   6030
      Width           =   2500
      _ExtentX        =   4419
      _ExtentY        =   926
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Registrar Como Incobrable"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton Command1 
      Height          =   525
      Left            =   1200
      TabIndex        =   5
      Top             =   6030
      Width           =   2500
      _ExtentX        =   4419
      _ExtentY        =   926
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Ver Resumen Detallado"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.ComboBox cmbDias 
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
      ItemData        =   "frmClientesOrg.frx":058A
      Left            =   3570
      List            =   "frmClientesOrg.frx":059D
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1290
      Width           =   1125
   End
   Begin MSComctlLib.ListView lvClientes 
      Height          =   3585
      Left            =   990
      TabIndex        =   3
      Top             =   1980
      Width           =   5745
      _ExtentX        =   10134
      _ExtentY        =   6324
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cliente"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Sin Mov Desde"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Importe"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Clientes Con Deudas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   900
      TabIndex        =   4
      Top             =   330
      Width           =   5745
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "días"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   4860
      TabIndex        =   2
      Top             =   1350
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ver Clientes sin movimientos hace:"
      ForeColor       =   &H00E0E0E0&
      Height          =   435
      Left            =   1530
      TabIndex        =   1
      Top             =   1200
      Width           =   2025
   End
End
Attribute VB_Name = "frmClientesOrg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UltCliente As String
Dim UltFecha As Date
Dim UltID As Long
Dim FechaMV As Date 'fecha mas vieja
Dim clscl As New clsCliente 'para usar el getDeuda
Dim Deuda As Single 'aca pongo el resultado
    

Private Sub ActualizaR(Tiempo As Long)
    Dim FechaD As Date, TmP As Long
    Dim RsVv As New ADODB.Recordset
    Dim Clientes() As String, Fechas() As Date
        
    FechaD = Date - Tiempo
    lvClientes.ListItems.Clear
    Label1.Caption = "Sin movimientos desde " + CStr(FechaD)
    
    RsVv.CursorLocation = adUseClient
    RsVv.Open "SELECT Clientes.Nombre ,Clientes.id, MovClientes.Fecha " + _
        "FROM Clientes INNER JOIN MovClientes ON Clientes.ID = " + _
        "MovClientes.CodCliente" + _
        " ORDER BY Clientes.nombre, movclientes.fecha", _
       DB.CN, adOpenDynamic, adLockReadOnly
        
        'where movclientes.fecha <= #" + CStr(FechaSQL(FechaD))
   
    If RsVv.RecordCount = 0 Then Exit Sub
    
     'veo la fecha mas vieja para compararla con el filtro
    FechaMV = Date
    RsVv.MoveFirst
    Do While Not RsVv.EOF
        If RsVv("fecha") < FechaMV Then FechaMV = RsVv("fecha")
        RsVv.MoveNext
    Loop
    
    If FechaMV > FechaD Then
        'si no pasa eso en el 1ro es porque el dia mas viejo es menor al
        'filtro pedido por eso salgo
        RsVv.Close
        Set clscl = Nothing
        Exit Sub
    End If
    
     'busco los clientes con la fecha del ult movimiento en ese tiempo
    RsVv.MoveFirst
     'no lo agrego hasta que vea que hay uno mas nuevo
    UltCliente = RsVv("nombre")
    UltFecha = RsVv("fecha")
    UltID = RsVv("Id")
     
    TmP = lvClientes.ListItems.Count + 1
      'por si es el unico registro
    If RsVv.RecordCount = 1 Then
        If UltFecha <= FechaD Then
            If UltCliente <> "Otros" Then
                Deuda = clscl.GetDeuda(UltID)
                'puede ser que deuda sea cero, agrego solo si es distinto de cero
                If Deuda <> 0 Then
                    lvClientes.ListItems.Add TmP
                    
                    lvClientes.ListItems(TmP).Text = UltCliente
                    lvClientes.ListItems(TmP).SubItems(1) = CStr(UltFecha)
                    lvClientes.ListItems(TmP).SubItems(2) = FormatCurrency(Deuda, , , , vbFalse)
                    lvClientes.ListItems(TmP).SubItems(3) = CStr(UltID)
                End If
            End If
        End If
        
        Set RsVv = Nothing
        Set clscl = Nothing
    
        Exit Sub
    End If
    
   
    RsVv.MoveFirst
    Do While Not RsVv.EOF
        If RsVv("nombre") <> UltCliente Then
            'solo si cambia de nombre es que el anterior era el registro mas nuevo
            'veo si fue antes del filtro si es asi lo agrego
            If UltFecha <= FechaD Then
                If UltCliente <> "Otros" Then 'HHHH
                    Deuda = clscl.GetDeuda(UltID)
                    'puede ser que deuda sea cero, agrego solo si es distinto de cero
                    If Deuda <> 0 Then
                        lvClientes.ListItems.Add TmP
                    
                        lvClientes.ListItems(TmP).Text = UltCliente
                        lvClientes.ListItems(TmP).SubItems(1) = CStr(UltFecha)
                        lvClientes.ListItems(TmP).SubItems(2) = FormatCurrency(Deuda, , , , vbFalse)
                        lvClientes.ListItems(TmP).SubItems(3) = CStr(UltID)
                        TmP = TmP + 1
                    End If
                End If
            End If
        End If
        
        UltCliente = RsVv("Nombre")
        UltFecha = RsVv("Fecha")
        UltID = RsVv("Id")
        RsVv.MoveNext
    Loop
    
        'voy al ultimo que no se estaba agregando
    If UltFecha <= FechaD Then
        If UltCliente <> "Otros" Then 'HHHH
            Deuda = clscl.GetDeuda(UltID)
            lvClientes.ListItems.Add TmP
                    
            lvClientes.ListItems(TmP).Text = UltCliente
            lvClientes.ListItems(TmP).SubItems(1) = CStr(UltFecha)
            lvClientes.ListItems(TmP).SubItems(2) = FormatCurrency(Deuda, , , , vbFalse)
            lvClientes.ListItems(TmP).SubItems(3) = CStr(UltID)
        End If
    End If
    
    RsVv.Close
    Set RsVv = Nothing
    Set clscl = Nothing
End Sub

Private Function GetNombre(inD As Long) As String
    GetNombre = txtInLvW(lvClientes, inD, 0)
End Function

Private Function GetFecha(inD As Long) As Date
    GetFecha = CDate(txtInLvW(lvClientes, inD, 1))
End Function

Private Function GetDeuda(inD As Long) As Single
    GetDeuda = CSng(txtInLvW(lvClientes, inD, 2))
End Function

Private Function GetID(inD As Long) As Long
    GetID = CLng(txtInLvW(lvClientes, inD, 3))
End Function

Private Sub cmbDias_Click()
    ActualizaR (CLng(cmbDias))
End Sub

Private Sub cmdImprimir_Click()
    Dim Tit() As String

    TP.LineasSeparadoras = True
    
    ReDim Preserve Tit(4)
    Tit(4) = "Clientes Con Deudas"
    'datos de mi empresa!!!!!!!!!!!!!!
    Tit(0) = DB.GetValInRS("Clientes", "Nombre", "ID = -2", True)
    Tit(1) = "Direccion: " + DB.GetValInRS("Clientes", "Direccion", "ID = -2", True)
    Tit(2) = "Teléfono: " + DB.GetValInRS("Clientes", "Telefono", "ID = -2", True)
    Tit(3) = "Mail: " + DB.GetValInRS("Clientes", "Mail", "ID = -2", True)
    
    TP.ImprimirlvW lvClientes, Tit, _
        "Cliente|Sin Mov Desde|Importe", Label1, , 1.53, , , , 30
End Sub

Private Sub cmdIncobrables_Click()
    If lvClientes.ListItems.Count = 0 Then Exit Sub
    
    Dim UUs As Long, nBUS As String
    UUs = ACC.UltUsuarioIngresado
    nBUS = ACC.GetNombre("Usuario", "Usuarios", UUs)
    
    If ACC.ExisteRelacion(UUs, 8) = 0 Then
        MsgBox ACC.GetNombre("Usuario", "Usuarios", UUs) + " no está habilitado " + _
            "para ingresar." + vbCrLf + _
            "Debe Cambiar Sesión a la de un usuario habilitado", vbExclamation, "Atención"
        Exit Sub
    Else
        'registro el movimiento(8:Anotar a Clientes)
        ACC.RegEvento UUs, 8, "Registrar a " + CStr(lvClientes.ListItems.Count) + _
            " clientes como incobrables"
    End If

    Dim Ii As Long
        
    For Ii = 1 To lvClientes.ListItems.Count
        If GetDeuda(Ii) > 0 Then 'sin tabular si no no tengo mas espacio

        If lvClientes.ListItems(Ii).Checked = True Then

            If MsgBox("¿Está seguro de registrar como incobrable a " + GetNombre(Ii) + _
                " y registrar como perdida " + FormatCurrency(GetDeuda(Ii)), _
                vbExclamation + vbYesNo, "Atención") = vbYes Then

                'borro los registros
                DB.EXECUTE "DELETE FROM MovClientes WHERE CodCliente= " + CStr(GetID(Ii))
         
                'Dejo registros de la deuda incobrable
                DB.EXECUTE "INSERT INTO MovClientes (ID,Fecha,CodCliente,Detalle," + _
                    "Variacion,Documento) VALUES (" + IdAutonum("MovClientes") + _
                    ",#" + stFechaSQL(Date) + "#," + CStr(GetID(Ii)) + _
                    ",'PERDIDA POR INCOBRABLE!! (" + FormatCurrency(GetDeuda(Ii)) + _
                    ") desde " + CStr(GetFecha(Ii)) + "',0,'NO')"
                    
                'hago el asiento "perdida por incobrables" a "Clientes"
                PC.Asiento "28", CStr(GetDeuda(Ii)), "46", CStr(GetDeuda(Ii)), _
                    "LibroSubDiario", _
                    "Pérdida por Incobrable Cliente Nro " + CStr(GetID(Ii))
                
                'registro un comentario en el cliente
                DB.EXECUTE "INSERT INTO ComentariosClientes (ID,IdCliente,Fecha,Hora," + _
                    "Comentario,Usuario) " + _
                    "VALUES (" + IdAutonum("ComentariosClientes") + _
                    "," + CStr(GetID(Ii)) + ",#" + stFechaSQL(Date) + "#," + _
                    CStr(Hour(Now)) + ",'" + _
                    "Registro deudas Incobrable por " + FormatCurrency(GetDeuda(Ii)) + _
                    "','" + nBUS + "')"
                
            End If 'no tabule mas porque no tengo espacio
            End If
        End If
    Next Ii
    
    cmbDias_Click
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub



Private Sub Command1_Click()
    If lvClientes.ListItems.Count = 0 Then Exit Sub
    
    frmResClientes.AbrirDatos GetID(lvClientes.SelectedItem.Index)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    cmbDias.ListIndex = 0
End Sub
