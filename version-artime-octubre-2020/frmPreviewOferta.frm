VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmPreviewOferta 
   Caption         =   "Previsualizacion Impresion"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Cambiar diseño"
      Height          =   4035
      Left            =   30
      TabIndex        =   13
      Top             =   2610
      Width           =   1275
      Begin tbrFaroButton.fBoton fBAR 
         Height          =   165
         Left            =   120
         TabIndex        =   25
         Top             =   3600
         Width           =   105
         _ExtentX        =   185
         _ExtentY        =   291
         fFColor         =   6553600
         fBColor         =   192
         fCapt           =   ""
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   12632319
      End
      Begin VB.TextBox txtSEP 
         Height          =   345
         Left            =   30
         TabIndex        =   24
         Text            =   "300"
         Top             =   2940
         Width           =   435
      End
      Begin VB.TextBox txtML 
         Height          =   345
         Left            =   30
         TabIndex        =   20
         Text            =   "300"
         Top             =   2250
         Width           =   435
      End
      Begin VB.TextBox txtMS 
         Height          =   345
         Left            =   30
         TabIndex        =   19
         Text            =   "400"
         Top             =   1590
         Width           =   435
      End
      Begin VB.TextBox txtV 
         Height          =   345
         Left            =   30
         TabIndex        =   16
         Text            =   "5"
         Top             =   990
         Width           =   345
      End
      Begin VB.TextBox txtH 
         Height          =   345
         Left            =   30
         TabIndex        =   15
         Text            =   "3"
         Top             =   420
         Width           =   345
      End
      Begin tbrFaroButton.fBoton fBoton4 
         Height          =   615
         Left            =   90
         TabIndex        =   14
         Top             =   3360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1085
         fFColor         =   6553600
         fBColor         =   16761024
         fCapt           =   "Repintar"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   16777215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Separador"
         Height          =   195
         Index           =   2
         Left            =   30
         TabIndex        =   23
         Top             =   2730
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Margen Lateral"
         Height          =   195
         Index           =   1
         Left            =   30
         TabIndex        =   22
         Top             =   2040
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Margen Sup."
         Height          =   195
         Index           =   1
         Left            =   30
         TabIndex        =   21
         Top             =   1380
         Width           =   915
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vertical"
         Height          =   195
         Index           =   0
         Left            =   30
         TabIndex        =   18
         Top             =   810
         Width           =   525
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Horizontal"
         Height          =   195
         Index           =   0
         Left            =   30
         TabIndex        =   17
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.PictureBox picFINAL 
      Height          =   1215
      Left            =   8610
      ScaleHeight     =   1155
      ScaleWidth      =   945
      TabIndex        =   11
      Top             =   1560
      Visible         =   0   'False
      Width           =   1005
   End
   Begin tbrFaroButton.fBoton fBoton1 
      Height          =   435
      Left            =   90
      TabIndex        =   4
      Top             =   150
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   767
      fFColor         =   6553600
      fBColor         =   16761024
      fCapt           =   "Fondo"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   16777215
   End
   Begin VB.PictureBox PicFondo 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   5235
      Left            =   1770
      ScaleHeight     =   5235
      ScaleWidth      =   6465
      TabIndex        =   0
      Top             =   360
      Width           =   6465
      Begin VB.HScrollBar hSc 
         Height          =   315
         Left            =   60
         TabIndex        =   3
         Top             =   4860
         Width           =   1875
      End
      Begin VB.VScrollBar vSc 
         Height          =   1245
         Left            =   6030
         TabIndex        =   2
         Top             =   3930
         Width           =   375
      End
      Begin VB.PictureBox picPrint 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3915
         Left            =   30
         ScaleHeight     =   3915
         ScaleWidth      =   4935
         TabIndex        =   1
         Top             =   30
         Width           =   4935
         Begin VB.PictureBox PIC2 
            Height          =   465
            Index           =   0
            Left            =   4230
            ScaleHeight     =   405
            ScaleWidth      =   525
            TabIndex        =   10
            Top             =   240
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.Label lbOBS 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "14% Descuento efectivo. 14% Descuento efectivo. 14% Descuento efectivo. 14% Descuento efectivo. 14% Descuento efectivo. "
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   0
            Left            =   480
            TabIndex        =   8
            Top             =   1860
            Visible         =   0   'False
            Width           =   3225
         End
         Begin VB.Label lbDesc 
            BackStyle       =   0  'Transparent
            Caption         =   "14% Descuento efectivo"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Index           =   0
            Left            =   1860
            TabIndex        =   7
            Top             =   1290
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.Label lbPrecio 
            BackStyle       =   0  'Transparent
            Caption         =   "$ 3.999"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   1830
            TabIndex        =   6
            Top             =   840
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label lbNAME 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00404040&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Motorola v300 nombnre largo de prod"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   0
            Left            =   480
            TabIndex        =   5
            Top             =   180
            Visible         =   0   'False
            Width           =   3285
         End
         Begin VB.Image IM 
            Height          =   1305
            Index           =   0
            Left            =   480
            Stretch         =   -1  'True
            Top             =   540
            Visible         =   0   'False
            Width           =   1335
         End
      End
   End
   Begin tbrFaroButton.fBoton fBoton2 
      Height          =   435
      Left            =   120
      TabIndex        =   9
      Top             =   1950
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   767
      fFColor         =   6553600
      fBColor         =   16761024
      fCapt           =   "Imprimir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   16777215
   End
   Begin tbrFaroButton.fBoton fBoton3 
      Height          =   525
      Left            =   90
      TabIndex        =   12
      Top             =   660
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   926
      fFColor         =   6553600
      fBColor         =   16761024
      fCapt           =   "Siguiente pag."
      fEnabled        =   0   'False
      fFontN          =   ""
      fFontS          =   0
      fECol           =   16777215
   End
   Begin VB.Line Line1 
      X1              =   1200
      X2              =   1200
      Y1              =   30
      Y2              =   1410
   End
End
Attribute VB_Name = "frmPreviewOferta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const WM_PAINT = &HF
Private Const WM_PRINT = &H317
Private Const PRF_CLIENT = &H4& ' Draw the window's client area.
Private Const PRF_CHILDREN = &H10& ' Draw all visible child windows.
Private Const PRF_OWNED = &H20& ' Draw all owned windows.

Private Declare Function SendMessage Lib "user32" Alias _
   "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
   ByVal wParam As Long, ByVal lParam As Long) As Long
   
Dim FSO As New Scripting.FileSystemObject

Private OfHoriz As Long 'cantidad de ofertas horizontal por página
Private OfVert As Long 'cantidad de ofertas vertical por página
Private MargenLat As Long
Private MargenSup As Long
Private Separador As Long

Private Type Oferta
    X As Long
    Ancho As Long
    Y As Long
    Alto As Long
    
    Precio As Single
    DescEfvo As Long
    Nombre As String
    idPord As Long
    Observaciones As String
    
    PathImage As String
End Type

Dim FileOferta As String
Dim OF() As Oferta
Dim LoadHastaOferta As Long 'hasta que numero se mostro de los que habia

Private Sub fBoton1_Click()
    Dim CM As New CommonDialog
    CM.Filter = "Imagenes (*.jpg *.jpeg *.gif *.bmp)|*.jpg; *.jpeg; *.gif; *.bmp"
    CM.InitDir = App.path
    CM.ShowOpen
    
    Dim F As String
    F = CM.FileName
    
    If F = "" Then Exit Sub
    
    picPrint.AutoSize = True
    picPrint.Picture = LoadPicture(F)
    
    
    Dim pixPorCM As Single
    pixPorCM = 567 'segun ayuda de vb. Dice centímetro lógico
    
    'picPrint.Width = 21 * pixPorCM
    'picPrint.Height = 29.7 * pixPorCM
    
    Limpiar
    
    VER FileOferta
    
End Sub

Public Sub VER(fOfertas As String, Optional DesdeArt As Long = -1)
    
    FileOferta = fOfertas
    
    Dim TE As TextStream, TmP As Long, SP() As String, J As Long, H As Long
    Set TE = FSO.OpenTextFile(fOfertas, ForReading, False)
        
        SP = Split(TE.ReadAll, Chr(5))
        ReDim OF(0)
        
        Dim PR As Single
        Dim DC As Long
        
        'si estoy viendo otra página ....
        Dim INI_OF As Long
        If DesdeArt = -1 Then
            INI_OF = 0
        Else
            'si llego al final
            If DesdeArt >= UBound(SP) Then
                INI_OF = 0
            Else
                INI_OF = DesdeArt
            End If
        End If
        
        'si hay solo 1 productyo es cero !!! y da error
        If UBound(SP) = 0 Then
            MsgBox "Hay solo 1 producto, utilize al menos 2"
            Exit Sub
        End If
        For J = INI_OF To UBound(SP)
        
            fBAR.Width = ((J - INI_OF) / (UBound(SP) - INI_OF)) * Frame1.Width
            
            H = UBound(OF) + 1
            ReDim Preserve OF(H)
            OF(H).idPord = SP(J)
            OF(H).Nombre = GetDataProd(CLng(SP(J)), "nProducto")
            
            PR = GetDataProd(CLng(SP(J)), "pVenta")
            OF(H).Precio = PR
            
            DC = GetDataProd(CLng(SP(J)), "DescEfvo")
            OF(H).DescEfvo = DC
            
            Dim OB As String
            OB = GetDataProd(CLng(SP(JF)), "Observaciones")
            OF(H).Observaciones = OB
            
            OF(H).PathImage = CFGBD.GetInfo(82, 4) + "IMG\" + SP(J) + "-0.jpg"  'la primera de las posibles imagenes a mostrar
            OF(H).Alto = (picPrint.Height - (MargenLat * 2) - (Separador * (OfVert + 1))) / OfVert
            OF(H).Ancho = (picPrint.Width - (MargenSup * 2) - (Separador * (OfHoriz + 1))) / OfHoriz
        Next J
    TE.Close

    'imprimir en el picturebox
    Dim Cont As Long, PosX As Long, PosY As Long
    PosX = 1: PosY = 1
    
    For J = 1 To UBound(OF)
        
        fBAR.Width = (J / UBound(OF)) * Frame1.Width
        
        Load IM(J)
        Load lbNAME(J)
        Load lbPrecio(J)
        Load lbDesc(J)
        Load lbOBS(J)
        Load PIC2(J)
        
        IM(J).Width = OF(J).Ancho / 2
        IM(J).Height = OF(J).Alto - lbNAME(J).Height - lbOBS(J).Height
        
        lbNAME(J).Caption = OF(J).Nombre
        If OF(J).Precio > 0 Then
            lbPrecio(J).Caption = "$ " + CStr(Round(OF(J).Precio, 2))
        Else
            lbPrecio(J).Caption = ""
        End If
        
        If OF(J).DescEfvo > 0 Then
            lbDesc(J).Caption = "Descuento Efvo: " + CStr(OF(J).DescEfvo) + " %"
        Else
            lbDesc(J).Caption = ""
        End If
            
        lbOBS(J).Caption = OF(J).Observaciones
        
        IM(J).Left = MargenLat + Separador + ((PosX - 1) * (OF(J).Ancho + Separador))
        lbNAME(J).Left = MargenLat + Separador + (PosX - 1) * (OF(J).Ancho + Separador)
        lbPrecio(J).Left = MargenLat + Separador + ((PosX - 1) * (OF(J).Ancho + Separador)) + (OF(J).Ancho / 2)
        lbDesc(J).Left = lbPrecio(J).Left
        lbOBS(J).Left = lbNAME(J).Left
        
        IM(J).Top = MargenSup + Separador + ((PosY - 1) * (OF(J).Alto + Separador)) + lbNAME(J).Height
        lbNAME(J).Top = MargenSup + Separador + ((PosY - 1) * (OF(J).Alto + Separador))
        lbPrecio(J).Top = lbNAME(J).Top + lbNAME(J).Height
        lbDesc(J).Top = lbPrecio(J).Top + lbPrecio(J).Height
        lbOBS(J).Top = IM(J).Top + IM(J).Height + 60
        
        lbNAME(J).Width = OF(J).Ancho
        lbPrecio(J).Width = lbNAME(J).Width / 2
        lbDesc(J).Width = lbNAME(J).Width / 2
        lbOBS(J).Width = OF(J).Ancho
        
        IM(J).Visible = True
        lbNAME(J).Visible = True
        lbPrecio(J).Visible = True
        lbDesc(J).Visible = True
        lbOBS(J).Visible = True
        
        If FSO.FileExists(OF(J).PathImage) Then
            Dim Stp As New stdole.StdPicture
            Set Stp = LoadPicture(OF(J).PathImage)
            
            'calcular segun el tamaño disponible
            Dim Ancho As Long
            Dim Alto As Long
            Dim Prop As Single
            Prop = Stp.Width / Stp.Height
                            
            Ancho = IM(J).Width
            Alto = Ancho / Prop
            
            If Alto > IM(J).Height Then
                Alto = IM(J).Height
                Ancho = IM(J).Width * Prop
            End If
            
            'PIC2(J).PaintPicture STP, IM(J).Width / 2 - Ancho / 2, _
                                             IM(J).Height / 2 - Alto / 2, Ancho / 2, Alto / 2
            
            'IM(J).Picture = PIC2(J).Image
            PIC2(J).BorderStyle = 0
            PIC2(J).Width = Ancho
            PIC2(J).Height = Alto
            
'            PIC2(J).PaintPicture STP, 0, 0, Ancho, Alto
'            PIC2(J).Left = IM(J).Left + (IM(J).Width / 2 - Ancho / 2)
'            PIC2(J).Top = IM(J).Top + (IM(J).Height / 2 - Alto / 2)
'            PIC2(J).Visible = True
'            PIC2(J).ZOrder
            picPrint.PaintPicture Stp, IM(J).Left + (IM(J).Width / 2 - Ancho / 2), _
                IM(J).Top + (IM(J).Height / 2 - Alto / 2), _
                Ancho, Alto
        End If
        
        PosX = PosX + 1
        If PosX > OfHoriz Then
            PosX = 1
            PosY = PosY + 1
            If PosY > OfVert Then
                'fin pagina!
                LoadHastaOferta = J
                'habilitar boton de cambio de página
                fBoton3.Enabled = True
                Exit For
                
            End If
        End If
        
    Next J
    
    LoadHastaOferta = INI_OF + J
End Sub

Private Sub fBoton2_Click()
    'imprimir todo igual que esta en el picturebox
    On Error GoTo ErrPrint
'    Printer.PaintPicture picPrint.Image, 0, 0, picPrint.Width, picPrint.Height, 0, 0, picPrint.Width, picPrint.Height
'
'    Dim J As Long
'    For J = 1 To lbNAME.Count - 1
'        Printer.PaintPicture PIC2(J).Image, PIC2(J).Left, PIC2(J).Top, PIC2(J).Width, PIC2(J).Height
'    Next J
'
'    'poner los textos en sus lugares
'
'    Printer.EndDoc

    DoEvents

    picFINAL.Width = picPrint.Width
    picFINAL.Height = picPrint.Height

    picPrint.SetFocus
    picFINAL.AutoRedraw = True
    rv = SendMessage(picPrint.hwnd, WM_PAINT, picFINAL.hdc, 0)
    rv = SendMessage(picPrint.hwnd, WM_PRINT, picFINAL.hdc, PRF_CHILDREN + PRF_CLIENT + PRF_OWNED)
    picFINAL.Picture = picFINAL.Image
    picFINAL.AutoRedraw = False
    
    Printer.PaintPicture picFINAL.Picture, 0, 0
    Printer.EndDoc

    Exit Sub
ErrPrint:
    MsgBox "Error de impresora" + vbCrLf + CStr(Err.Number) + ": " + Err.Description
End Sub

Private Sub fBoton3_Click()
    'ver en que numero quedo y seguir
    Limpiar
    VER FileOferta, LoadHastaOferta + 1
End Sub

Private Sub fBoton4_Click()
    On Local Error Resume Next
        
    Limpiar
    fBAR.Width = 15
    fBAR.Left = 15
    fBAR.Top = fBoton4.Top + fBoton4.Height - fBAR.Height - 15
    
    'valores predetrerminados
    OfHoriz = CLng(txtH)
    OfVert = CLng(txtV)
    
    MargenSup = CLng(txtMS)
    MargenLat = CLng(txtML)
    Separador = CLng(txtSEP)
    
    VER FileOferta
End Sub

Private Sub Form_Load()
    picPrint.AutoRedraw = True
    picPrint.Picture = LoadPicture
    picPrint.BackColor = vbWhite
    
    Dim pixPorCM As Single
    pixPorCM = 567 'segun ayuda de vb. Dice centímetro lógico
    
    picPrint.Width = 21 * pixPorCM
    picPrint.Height = 29.7 * pixPorCM
    
    PIC2(0).AutoRedraw = True
    PIC2(0).Visible = False
    PIC2(0).AutoSize = True
    
    'valores predetrerminados
    OfHoriz = 3
    OfVert = 5
    
    MargenSup = 400
    MargenLat = 300
    Separador = 600
    
    fBoton4.Left = 0
    fBoton4.Width = Frame1.Width
    fBAR.Width = 15
    fBAR.Left = 15
    fBAR.Top = fBoton4.Top + fBoton4.Height - fBAR.Height - 15
    
End Sub

Private Sub Form_Resize()
    On Local Error Resume Next
    
    PicFondo.Left = Line1.X1 + 60
    PicFondo.Top = 100
    PicFondo.Width = Me.Width - Line1.X1 - 160
    PicFondo.Height = Me.Height - PicFondo.Top - 460
    
    vSc.Left = PicFondo.Width - vSc.Width '- 30
    vSc.Top = 0
    vSc.Height = PicFondo.Height - hSc.Height
    
    hSc.Left = 0
    hSc.Top = PicFondo.Height - hSc.Height '- 30
    hSc.Width = PicFondo.Width - vSc.Width
    
    'ver el picPrint para estaBLECER EL VALOR DE LOS SCROLS
    picPrint.Top = 0
    picPrint.Left = 0
    
    If picPrint.Width > PicFondo.Width Then
        Dim DifW As Long
        DifW = picPrint.Width - PicFondo.Width
        
        hSc.Min = 0
        hSc.Value = 0
        hSc.Max = DifW
        hSc.SmallChange = DifW / 30
        hSc.LargeChange = DifW / 10
        hSc.Enabled = True
    Else
        hSc.Min = 0
        hSc.Value = 0
        hSc.Max = 0
        hSc.Enabled = False
    End If
    
    If picPrint.Height > PicFondo.Height Then
        Dim DifH As Long
        DifH = picPrint.Height - PicFondo.Height
        
        vSc.Min = 0
        vSc.Value = 0
        vSc.Max = DifH
        vSc.SmallChange = DifH / 30
        vSc.LargeChange = DifH / 10
        vSc.Enabled = True
    Else
        vSc.Min = 0
        vSc.Value = 0
        vSc.Max = 0
        vSc.Enabled = False
    End If
    
End Sub

Private Sub hSc_Change()
    picPrint.Left = -hSc.Value
End Sub

Private Sub vSc_Change()
    picPrint.Top = -vSc.Value
End Sub

Private Function GetDataProd(idProd As Long, stField As String) As String
    
    If stField = "DescEfvo" Then 'no esta en la base de datos!
        Dim Desc As Long
        Dim tmp2 As Long
        tmp2 = CFG.ExistePropiedad("DVC " + CStr(idProd))
        If tmp2 = 0 Then
            Desc = CFG.GetInfo(60, 4)
        Else
            Desc = CFG.GetInfo(tmp2, 4)
        End If
        GetDataProd = Desc
    Else
        Dim rS As New ADODB.Recordset
        rS.CursorLocation = adUseClient
        rS.Open "Select * from productos where id=" + CStr(idProd), DB.CN, adOpenStatic, adLockReadOnly
        
        If rS.RecordCount <> 1 Then
            GetNameProd = ""
            Exit Function
        End If
        Dim TmP As String
        If IsNull(rS.Fields(stField)) Then
            TmP = ""
        Else
            TmP = rS.Fields(stField)
            
        End If
        
        GetDataProd = TmP
    End If
    
End Function

Private Sub Limpiar()
    picPrint.Cls
    Dim J As Long
    For J = lbNAME.Count - 1 To 1 Step -1
        Unload IM(J)
        Unload lbNAME(J)
        Unload lbPrecio(J)
        Unload lbDesc(J)
        Unload lbOBS(J)
        Unload PIC2(J)
    Next J
End Sub
