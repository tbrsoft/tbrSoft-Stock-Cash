VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmTipo 
   BackColor       =   &H004E4E4E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sucursales"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTipo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin tbrFaroButton.fBoton cmdEliminar 
      Height          =   435
      Left            =   3570
      TabIndex        =   3
      Top             =   1530
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   16777215
      fCapt           =   "Eliminar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdModificar 
      Height          =   435
      Left            =   3570
      TabIndex        =   4
      Top             =   930
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   16777215
      fCapt           =   "Modificar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdAgregar 
      Height          =   435
      Left            =   3570
      TabIndex        =   1
      Top             =   390
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   16777215
      fCapt           =   "Agregar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.ListBox lstSucu 
      Height          =   3570
      Left            =   300
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   390
      Width           =   3075
   End
   Begin tbrFaroButton.fBoton Command2 
      Height          =   435
      Left            =   3600
      TabIndex        =   2
      Top             =   3720
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   16777215
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
End
Attribute VB_Name = "frmTipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mTabla As String

Private Sub cmdAgregar_Click()
    Dim S As String
    
    S = InputBox(mTabla + " Ingreso nuevo")
    If S = "" Then Exit Sub
        
    If mTabla = "Envases" Then
        If DB.ContarReg("SELECT * FROM Envases WHERE Envase='" + S + "'") > 0 Then
            MsgBox "Ya tiene Envase con ese nombre", vbInformation, "Atención"
        Else
            DB.EXECUTE "INSERT INTO Envases (Envase) " + _
                "VALUES ('" + S + "') "
            CargarDatos
        End If
    Else
        If DB.ContarReg("SELECT * FROM Sucursales WHERE Sucursal='" + S + "'") > 0 Then
            MsgBox "Ese Tipo de producto ya existe", vbInformation, "Atención"
        Else
            DB.EXECUTE "INSERT INTO Sucursales (Sucursal) " + _
                "VALUES ('" + S + "') "
            CargarDatos
        End If
    End If
End Sub

Private Sub cmdEliminar_Click()
    If lstSucu.ListIndex = -1 Then Exit Sub
    
    If mTabla = "Envases" Then
        DB.EXECUTE "UPDATE Productos SET CodEnvase = 'No Tiene' " + _
            "WHERE CodEnvase = '" + lstSucu + "'"
        DB.EXECUTE "DELETE FROM Envases WHERE Envase = '" + lstSucu + "'"
    Else
        If lstSucu = "CASA CENTRAL" Then
            MsgBox "Casa Central no puede ser eliminado", vbInformation, "Atención"
            Exit Sub
        End If
        If MsgBox("¿Está seguro de borrar esta sucursal?." + vbCrLf + _
            "El stock del mismo pasará a CASA CENTRAL", vbInformation + vbYesNo, _
            "Atención") = vbNo Then Exit Sub
        DB.EXECUTE "DELETE FROM Sucursales WHERE Sucursal = '" + lstSucu + "'"
    End If
    '.... todavia no se que hacer
    CargarDatos
End Sub

Private Sub cmdModificar_Click()
    If lstSucu.ListIndex = -1 Then Exit Sub
    Dim S As String, S2 As String
    
    S = lstSucu
    S2 = InputBox(mTabla + " Ingreso nuevo", , S)
    If S2 = "" Then Exit Sub

    If mTabla = "Sucursales" Then
        If DB.ContarReg("SELECT * FROM Sucursales WHERE Sucursal='" + S2 + "'") > 0 Then
            If S2 <> S Then 'si son iguales es que no modifico nada al final
                MsgBox "Ya tiene una Sucursal con ese nombre", vbInformation, "Atención"
                Exit Sub
            End If
        End If
        
        DB.EXECUTE "UPDATE Sucursales SET Sucursal = '" + S2 + "'" + _
            " WHERE Sucursal = '" + S + "'"
        CargarDatos
    Else 'Envases
        If DB.ContarReg("SELECT * FROM Envases WHERE Envase='" + S2 + "'") > 0 Then
            If S2 <> S Then 'si son iguales es que no modifico nada al final
                MsgBox "Ya tiene Envase con ese nombre", vbInformation, "Atención"
                Exit Sub
            End If
        End If
        
        DB.EXECUTE "UPDATE Envases SET Envase = '" + S2 + "'" + _
            " WHERE Envase = '" + S + "'"
        CargarDatos
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    CargarDatos
End Sub

Public Sub Iniciar(Tabla As String)
    mTabla = Tabla
    
    If mTabla = "Envases" Then Me.BackColor = &HE8D5D8
    
    Me.Show 1
End Sub

Private Sub CargarDatos()
    Dim I As Long
    
    If mTabla = "Envases" Then
        CargarCombo lstSucu, "SELECT Envase FROM Envases", "Envase"
        
        For I = 0 To lstSucu.ListCount - 1
            If lstSucu.List(I) = "No tiene" Then lstSucu.RemoveItem (I)
            'Exit For
        Next I
    Else
        lstSucu.Clear
        lstSucu.AddItem "CASA CENTRAL"
        CargarCombo lstSucu, "SELECT * FROM Sucursales", "Sucursal", , True
    End If
End Sub
