VERSION 5.00
Begin VB.Form frmWAIT 
   BackColor       =   &H00544B45&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Espere"
   ClientHeight    =   600
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2490
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleWidth      =   2490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Procesando..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   2295
   End
End
Attribute VB_Name = "frmWAIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
