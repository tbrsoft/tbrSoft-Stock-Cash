VERSION 5.00
Object = "{DCB03D77-0A94-4AE8-9495-515B6968EEFB}#4.0#0"; "tbrfacura.ocx"
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmFactura 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración Formato Factura"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11820
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFactura.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   11820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin tbrFaroButton.fBoton cmdSalir 
      Height          =   480
      Left            =   10230
      TabIndex        =   5
      Top             =   8010
      Width           =   1250
      _ExtentX        =   2196
      _ExtentY        =   847
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdGrabarComo 
      Height          =   480
      Left            =   9120
      TabIndex        =   4
      Top             =   7410
      Width           =   1250
      _ExtentX        =   2196
      _ExtentY        =   847
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Grabar como"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdGrabar 
      Height          =   480
      Left            =   7710
      TabIndex        =   3
      Top             =   7410
      Width           =   1250
      _ExtentX        =   2196
      _ExtentY        =   847
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Grabar"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdFuente 
      Height          =   480
      Left            =   9120
      TabIndex        =   2
      Top             =   6900
      Width           =   1250
      _ExtentX        =   2196
      _ExtentY        =   847
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Fuente"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin tbrFaroButton.fBoton cmdAbrir 
      Height          =   480
      Left            =   7710
      TabIndex        =   1
      Top             =   6900
      Width           =   1250
      _ExtentX        =   2196
      _ExtentY        =   847
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Abrir"
      fEnabled        =   -1  'True
      fFontN          =   "Arial"
      fFontS          =   8
      fECol           =   5717301
   End
   Begin tbrFacura.Factura FAC 
      Height          =   8535
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   15055
   End
End
Attribute VB_Name = "frmFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FuenteAca As StdFont
Dim CDg As New CommonDialog
Dim FSO As New FileSystemObject

Private Sub cmdAbrir_Click()
    FAC.AbrirConfiguracion
End Sub

Private Sub cmdFuente_Click()
    Dim Nombre As String, Tamano As Long
    
    CDg.ShowFont
    
    Nombre = CDg.FontName
    Tamano = CDg.FontSize
    
    If Nombre <> "" Then
        FuenteAca.Name = Nombre
    Else
        FuenteAca.Name = "Arial"
    End If
    
    If Tamano < 8 Then
        FuenteAca.Size = 8
    Else
        FuenteAca.Size = Tamano
    End If
    
    GrabarFuente FuenteAca.Name, FuenteAca.Size
End Sub

Private Sub cmdGrabar_Click()
    FAC.GrabarConfiguracion False
End Sub

Private Sub cmdGrabarComo_Click()
    FAC.GrabarConfiguracion True
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub LeerFuente()
    Dim tmPF As StdFont, TmP As String, TT As String, SP() As String
    Dim TE As TextStream
    
    TmP = AP + "Fte.fys"
    Set tmPF = Me.Font 'predeterminado
    
    If FSO.FileExists(TmP) = False Then
        FSO.CreateTextFile TmP, True
        Set TE = FSO.OpenTextFile(TmP, ForWriting, True)
            TE.WriteLine "Arial | 9"
        TE.Close
    End If
    
    Set TE = FSO.OpenTextFile(TmP, ForReading)
        TT = TE.ReadLine
    TE.Close
    
    SP = Split(TT, " | ")
        
    tmPF.Name = SP(0)
    tmPF.Size = CLng(SP(1))
        
    Set TE = Nothing
    
    Set FuenteAca = tmPF
End Sub

Private Sub GrabarFuente(Nombre As String, Tamano As Long)
    Dim TmP As String
    
    TmP = AP + "Fte.fys"
    
    If FSO.FileExists(TmP) = False Then
        FSO.CreateTextFile TmP, True
    End If
    
    Set TE = FSO.OpenTextFile(TmP, ForWriting, True)
        TE.WriteLine Nombre + " | " + CStr(Tamano)
    TE.Close

End Sub

Private Sub fBoton1_Click()

End Sub

Private Sub Form_Load()
    LeerFuente
    
    FAC.Labeles(0) = "Id Usuario"
    FAC.Labeles(1) = "Nombre"
    FAC.Labeles(2) = "Domicilio"
    FAC.Labeles(3) = "DNI / CUIT"
    FAC.Labeles(4) = "Condicion IVA"
    FAC.Labeles(5) = "Fecha"
    FAC.Labeles(6) = "Factura"
    FAC.Labeles(7) = "Descuento"
    FAC.Labeles(8) = "Neto sin IVA"
    FAC.Labeles(9) = "IVA %"
    FAC.Labeles(10) = "IVA $"
    FAC.Labeles(11) = "A Pagar"
    FAC.CargarLabels
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set CDg = Nothing
    Set FSO = Nothing
End Sub
