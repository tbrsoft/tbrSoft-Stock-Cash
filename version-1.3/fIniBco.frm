VERSION 5.00
Begin VB.Form fIniBco 
   BackColor       =   &H00F3F4E1&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cheques y Bancos"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8415
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fIniBco.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Emisión"
      Height          =   345
      Left            =   720
      TabIndex        =   4
      Top             =   2640
      Width           =   1425
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Depósito"
      Height          =   345
      Left            =   720
      TabIndex        =   3
      Top             =   1980
      Width           =   1425
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Salida"
      Height          =   345
      Left            =   750
      TabIndex        =   2
      Top             =   1680
      Width           =   1425
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Ingreso"
      Height          =   345
      Left            =   780
      TabIndex        =   1
      Top             =   1320
      Width           =   1425
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cheques"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   780
      TabIndex        =   0
      Top             =   630
      Width           =   1635
   End
End
Attribute VB_Name = "fIniBco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
