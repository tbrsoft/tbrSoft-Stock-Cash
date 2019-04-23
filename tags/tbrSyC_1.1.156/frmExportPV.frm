VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmExportPV 
   BackColor       =   &H00544B45&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exportar datos para punto de venta"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   4980
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmExportPV.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin tbrFaroButton.fBoton command1 
      Height          =   435
      Left            =   1710
      TabIndex        =   1
      Top             =   2040
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "ok"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5717301
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Defina carpeta de destino y exportar:"
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   30
      TabIndex        =   2
      Top             =   1500
      Width           =   4785
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmExportPV.frx":058A
      ForeColor       =   &H00E0E0E0&
      Height          =   885
      Left            =   90
      TabIndex        =   0
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "frmExportPV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim CM As New CommonDialog
    CM.InitDir = AP
    
    CM.ShowFolder
    
    Dim F As String
    F = CM.InitDir
    
    If F = "" Or F = AP Then
        MsgBox "No se ha definido no se exportara"
        Exit Sub
    End If
    
    If Right(F, 1) <> "\" Then F = F + "\"
    
    'copiar la base de datos y la carpeta IMG
    
    Dim FFS As New Scripting.FileSystemObject
    
    If FFS.FolderExists(F + "PV\") Then FFS.CreateFolder F + "PV\"
    
    'BASE DE DATOS
    If FFS.FileExists(ArchivoMDBPrincipal) = False Then
        MsgBox "¡No se encuentra la base para copiar! No se seguirá"
        Exit Sub
    End If
    
    If FFS.FolderExists(F + "pv\") = False Then
        FFS.CreateFolder F + "pv\"
    End If
    
    FFS.CopyFile ArchivoMDBPrincipal, F + "PV\BASE.X"
    FFS.CopyFile CFGBD.GetInfo(80, 4) + "config.abl", F + "PV\BASE.X2"
    
    'IMAGENES
    If FFS.FolderExists(CFGBD.GetInfo(82, 4) + "IMG\") = False Then
        MsgBox "No se encuentran las imagenes! No se seguirá"
        Exit Sub
    End If
    
    FFS.CopyFolder CFGBD.GetInfo(82, 4) + "IMG", F + "PV\"
    
    MsgBox "Se ha exportado sin problemas !"
    
    Unload Me
    
End Sub
