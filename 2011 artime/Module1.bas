Attribute VB_Name = "Module1"
Public AP As String
Public WF As String
Public VtaDia As Single
Public CtoDia As Single
Public CpDia As Single
Public ArchivoMDBPrincipal As String
Public Contrasena As String
Public TipoUsuario As String
Public RespuestaInter(10) As String
    ' Por ahora:
    ' "0*IdCliente"
    ' "1*NombreCliente"
    ' "2*NombreProveedor"
    ' "3*IdProducto"
    ' "4*NombreProducto"
    ' El indice es igual que este numero en la ventana esa

Public DB As New tbrBasedeDatos.clsDataBase
Public PC As New tbrPruebaCont.clsPruebaContab
Public ACC As New tbrAccesos.clsTbrAccesos
Public CFG As New tbrArbol.clstbrArbol
Public CFGBD As New tbrArbol.clstbrArbol
Public TP As New tbrPrintTodo4.clsPrint
Public Terr As New tbrErrores.clsTbrERR
Public LIC As New tbrK

Public Function SumarColumnaLVW(LvW As ListView, Col As Long, _
    Optional Cuales As Long = 0) As Single
    'Cuales:
        ' 0 todos
        ' 1 solo positivos
        ' -1 solo negativos
    'por ahora no suma columna 0 XXXX
    
    Dim Colu As Long, I As Long, Res As Single, TmP As Single
    
    'if col>lvw ver que no sea la columna mayor a la cantidad de columnas XXXX
    
    Colu = Col
    Res = 0
    
    For I = 1 To LvW.ListItems.Count
        If Not IsNumeric(LvW.ListItems(I).SubItems(Colu)) Then
            SumarColumnaLVW = 0
            Exit Function
        End If
        
        TmP = CSng(LvW.ListItems(I).SubItems(Colu))
        
        Select Case Cuales
            Case 0
                Res = Res + TmP
            Case 1
                If TmP > 0 Then Res = Res + TmP
            Case -1
                If TmP < 0 Then Res = Res + TmP
        End Select
    Next I

    SumarColumnaLVW = Res
End Function

Public Function stFechaSQL(Fecha As Date) As String
    stFechaSQL = Format(Fecha, "mm/dd/yyyy")
End Function

Public Function NoNuloN(J) As Single
    If IsNumeric(J) Then
        NoNuloN = J
    Else
        NoNuloN = 0
    End If
End Function

Public Function NoNuloS(S) As String
    If IsNull(S) Then
        NoNuloS = ""
    Else
        NoNuloS = S
    End If
End Function

Public Function NoNuloD(D) As Date
    If IsNull(S) Then
        NoNuloD = #1/1/1899#
    Else
        If IsDate(D) Then
            NoNuloD = D
        Else
            NoNuloD = #1/3/1899#
        End If
    End If
End Function

Public Function ValidarNumeros(TxtDeNroDudoso As Object, _
    Optional NroQueQuedaSiEstaMal As Single = 0) As Single
    
    Dim Nro As String
    
    Nro = TxtDeNroDudoso
    
    If Not IsNumeric(Nro) Then
        MsgBox "Debes cargar un numero correcto!", vbExclamation, "Atención"
        ValidarNumeros = NroQueQuedaSiEstaMal
        Exit Function
    End If
    
    ValidarNumeros = Nro
    
End Function

Public Sub CargarComboLV(Lv As ListView, sqlText As String, _
    CamposSeparadosPorComas As String)
    
    Lv.ListItems.Clear
    
    Dim Campos() As String
    Campos = Split(CamposSeparadosPorComas, ",")
    
    Dim rS As New ADODB.Recordset
    rS.Open sqlText, DB.CN, adOpenStatic, adLockReadOnly
    
    Dim S As String, AA As Long, TmP As Long
    Dim NombreRealCampo As String, Ult2 As String 'ultimos dos caracteres del campo
            
    If rS.RecordCount = 0 Then Exit Sub
    
    TmP = 1
    
    'los titulos (campos) no importan supongo que estan cargados a mano
    
    rS.MoveFirst
    Do While Not rS.EOF
        S = ""
        
        'el primero indice 1 cargo aca
        Lv.ListItems.Add TmP
        
        For AA = 0 To UBound(Campos)
            
            Ult2 = Right(Campos(AA), 2)
    
            Select Case Ult2
                Case "/n"
                    NombreRealCampo = Mid(Campos(AA), 1, Len(Campos(AA)) - 2)
                    S = CStr(NoNuloN(rS(NombreRealCampo)))
                    
                    If AA = 0 Then
                        Lv.ListItems(TmP).Text = S
                    Else
                        Lv.ListItems(TmP).SubItems(AA) = S
                    End If
                Case "/f"
                    NombreRealCampo = Mid(Campos(AA), 1, Len(Campos(AA)) - 2)
                    S = NoNuloD(rS(NombreRealCampo))
                    If AA = 0 Then
                        Lv.ListItems(TmP).Text = S
                    Else
                        Lv.ListItems(TmP).SubItems(AA) = S
                    End If
                Case "/$"
                    NombreRealCampo = Mid(Campos(AA), 1, Len(Campos(AA)) - 2)
                    S = FormatCurrency(rS(NombreRealCampo), , , , vbFalse)
                    If AA = 0 Then
                        Lv.ListItems(TmP).Text = S
                    Else
                        Lv.ListItems(TmP).SubItems(AA) = S
                    End If
                Case Else
                    S = NoNuloS(rS(Campos(AA)))
                    If AA = 0 Then
                        Lv.ListItems(TmP).Text = S
                    Else
                        Lv.ListItems(TmP).SubItems(AA) = S
                    End If
            End Select
            
        Next AA
        
        rS.MoveNext
        
        TmP = TmP + 1
    Loop
    rS.Close
    Set rS = Nothing
    
End Sub

Public Function ProximoMes(Fecha As Date, Optional Dia As Long = 0) As Date
    Dim DD As Long, MM As Long, AA As Long
    Dim D2 As Long, M2 As Long, A2 As Long
    
    If Dia > 31 Then Dia = 0
    DD = Day(Fecha): MM = Month(Fecha): AA = Year(Fecha)
    
    If MM = 12 Then
        M2 = 1
        A2 = AA + 1
    Else
        M2 = MM + 1
        A2 = AA
    End If
    
    If Dia = 0 Then
        D2 = DD
        
        If M2 = 2 And D2 > 28 Then D2 = 28 'por febrero que tiene 28 dias
        If D2 >= 31 And (M2 = 4 Or M2 = 6 Or M2 = 9 Or M2 = 11) Then D2 = 30 'meses de 30 d
    Else
        D2 = Dia
    End If
    
    ProximoMes = CDate(CStr(D2) + "/" + CStr(M2) + "/" + CStr(A2))
End Function

Public Sub CargarCombo(CMB As Object, sqlText As String, _
    CamposSeparadosPorComas As String, Optional Separador As String = "/", _
    Optional Agregado As Boolean = False)
    'ZZZZ pasar al completo
    
    'CamposSeparadosPorComas es una lista separada por comas de los campos. _
        Ademas se le puede agregar al final _
        /n al final para indicar que es numero _
        /f para fechas _
        /$ para currency _
        predeterminado es string
    If Agregado = False Then CMB.Clear
    
    Dim Campos() As String
    Campos = Split(CamposSeparadosPorComas, ",")
        
    Dim rS As New ADODB.Recordset
    rS.Open sqlText, DB.CN, adOpenStatic, adLockReadOnly
    Dim S As String, AA As Long
    
    If rS.RecordCount = 0 Then Exit Sub
    rS.MoveFirst
    Do While Not rS.EOF
        S = ""
        For AA = 0 To UBound(Campos)
            Dim Ult2 As String 'ultimos dos caracteres del campo
            Ult2 = Right(Campos(AA), 2)
            Dim NombreRealCampo As String
            Select Case Ult2
                Case "/n"
                    NombreRealCampo = Mid(Campos(AA), 1, Len(Campos(AA)) - 2)
                    S = S + CStr(NoNuloN(rS(NombreRealCampo)))
                Case "/f"
                    NombreRealCampo = Mid(Campos(AA), 1, Len(Campos(AA)) - 2)
                    S = S + CStr(rS(NombreRealCampo))
                Case "/$"
                    NombreRealCampo = Mid(Campos(AA), 1, Len(Campos(AA)) - 2)
                    S = S + FormatCurrency(rS(NombreRealCampo), , , , vbFalse)
                Case Else
                    S = S + NoNuloS(rS(Campos(AA)))
            End Select
            'si no es el ultimo poner la barra separadora
            If AA < UBound(Campos) Then S = S + Separador
        Next AA
        CMB.AddItem S
        rS.MoveNext
    Loop
    rS.Close
    Set rS = Nothing
    
    CMB.ListIndex = 0
End Sub

Public Function EsCero(Nro As Single) As Boolean
    If Abs(Nro) < 0.01 Then
        EsCero = True
    Else
        EsCero = False
    End If
End Function

Public Sub PintarTxt(TXT As Control)
    Dim xTXX As Control
    Set xTXX = TXT
    xTXX.SetFocus
    xTXX.SelStart = 0
    xTXX.SelLength = Len(xTXX)
    Set xTXX = Nothing
End Sub

Public Sub FormatearMouTextBox(FormName As Form, Optional SizeFont As Long = 9)
    Dim xT As Control
    
    For Each xT In FormName
        If TypeOf xT Is MouTextBox Then
            xT.Font.Name = FormName.Font.Name
            xT.Font.Size = SizeFont
            xT.BackColor = &HFFFFFF
            xT.Enabled = True
            xT.Alignment = vbCenter
        Else        'ya que esta tambien los tbrbuscador
        
            If TypeOf xT Is tbrBuscador Then
                xT.BackColor = FormName.BackColor
                xT.Font.Name = FormName.Font.Name
                xT.Font.Size = SizeFont
                xT.FontT.Name = FormName.Font.Name
                xT.FontT.Size = SizeFont
            End If
        End If
    Next
    
End Sub

Public Function IdAutonum(NombreTabla As String, _
    Optional NombreCampo As String = "ID") As String
    Dim Resp As Long
    
    Resp = DB.GetTop1Rs(NombreTabla, NombreCampo) + 1

    IdAutonum = CStr(Resp)
End Function

Public Function txtInLvW(Lista As ListView, Renglon As Long, Orden As Integer) As String
    'devuelve "OUT LISTA" si se solicita un orden no existente
    Dim Palabra As String
        
    If Orden = 0 Then
        Palabra = Lista.ListItems(Renglon).Text
    Else
        Palabra = Lista.ListItems(Renglon).SubItems(Orden)
    End If
    
    txtInLvW = Palabra
End Function

Public Sub LimpiarMovProdViejos()
    Dim BV As Long, TmP As String, RCont As Long
    
    TmP = CFG.GetInfo(17, 4)
    If Not IsNumeric(TmP) Then
        BV = 750
        CFG.ModificarNodo 17, , , , "750"
    Else
        BV = CLng(TmP)
    End If
    
    RCont = DB.ContarReg("SELECT ID FROM MovimientosProductos")
    If RCont > BV Then BorrarViejos (RCont - BV)
End Sub

Public Sub BorrarViejos(Cuantos As Long)
    If Cuantos <= 0 Then Exit Sub
    'no hay una forma mejor de hacerlo creo XXXX
    Dim RSn As New ADODB.Recordset
    Dim IdQ As Long
    
    RSn.Open "SELECT TOP " + CStr(Cuantos) + " ID FROM MovimientosProductos ORDER BY ID", DB.CN, adOpenStatic, adLockReadOnly
    If RSn.RecordCount > 0 Then
        RSn.MoveLast
        IdQ = CLng(RSn("ID"))
    Else
        IdQ = 0
    End If
    
    If IdQ > 0 Then
        DB.EXECUTE "DELETE FROM MovimientosProductos WHERE ID <= " + CStr(IdQ)
    End If
    
    RSn.Close
    Set RSn = Nothing
End Sub

''''' -------------------CALCULAR CUOTAS EN SISTEMA FRANCES --------------------
' formula matemática
Public Function ValorCuota(Interes As Single, Prestamo As Single, NroCuotas As Long) As Single
    'devuelve del valor de cuotas de un prestamo (Valor Inicial)
    Dim I As Single, Resp As Single, n As Long, V As Single
    
    I = Interes / 100: n = NroCuotas: V = Prestamo
    
    Resp = (I / (1 - (1 + I) ^ -n)) * V
    
    ValorCuota = Resp
End Function

Public Function AmortizacionCuota(Interes As Single, Cuota As Single, _
    NroCuotas As Long, QueCuotaEs As Long) As Single
    Dim I As Single, Resp As Single, n As Long, C As Single, R As Long

    I = Interes / 100: n = NroCuotas: C = Cuota: R = QueCuotaEs
    
    Resp = C * ((1 + I) ^ (R - n - 1))
    
    AmortizacionCuota = Resp
End Function

Public Function InteresCuota(Interes As Single, Cuota As Single, _
    NroCuotas As Long, QueCuotaEs As Long) As Single

    InteresCuota = Cuota - AmortizacionCuota(Interes, Cuota, NroCuotas, QueCuotaEs)
End Function

Public Function RedondeoArriba(Numero As Single) As Long
    Dim Resp As Long

    If Numero - Int(Numero) <> 0 Then 'tenia decimales
        Resp = Int(Numero) + 1 'redondeo arriba
    Else  'ya era un numero entero, lo dejo asi
        Resp = Numero
    End If
    
    RedondeoArriba = Resp
        
End Function

Public Function ObtenerArch(Carpeta As String, sSearch As String) As String()
    
    If Right(Carpeta, 1) <> "\" Then Carpeta = Carpeta + "\"
    Dim TMPmatriz() As String
    ReDim Preserve TMPmatriz(0) 'para que devuelva algo
    
    Dim NombreArchivo As String, ContadorArch As Long
    NombreArchivo = Dir$(Carpeta + sSearch)
    Do While Len(NombreArchivo)
        ContadorArch = ContadorArch + 1
        ReDim Preserve TMPmatriz(ContadorArch)
        TMPmatriz(ContadorArch) = Carpeta + NombreArchivo
        NombreArchivo = Dir$
    Loop
    
    ObtenerArch = TMPmatriz
End Function

Public Sub Aporte(nSocio As String, Importe As Single, _
    Optional DetAsi As String = "", Optional idCtaJ As Long = 78)
    Dim IdCz As Long
    
    IdCz = PC.GetIDCuenta("Aporte Participación " + nSocio)
    If IdCz = -1 Then
        PC.AgregarCuenta PC.GetUltIDMasUno, 15, "Aporte Participación " + nSocio
    End If
    
    IdCz = PC.GetIDCuenta("Aporte Participación " + nSocio)
    
    '1ro agrego MovSoc
    DB.EXECUTE "INSERT INTO MovSocyEmp (ID,Fecha, IdNivel3," + _
        "Tipo,Variacion,Detalle) VALUES (" + IdAutonum("MovSocyEmp") + _
        ",#" + stFechaSQL(Date) + _
        "#," + CStr(IdCz) + ",'Socio'," + _
        Replace(CStr(Importe), ",", ".") + _
        ",'Aporte Participación Socio')"
    
    '2do Contabilidad - no queda en su cuenta particular
                        ' va a estar en una cuenta "Aporte Participación "+ Nombre
    PC.Asiento CStr(idCtaJ), CStr(Importe), CStr(IdCz), CStr(Importe), , DetAsi
    
End Sub

Public Function GetParticipacion(nSocio As String) As Single
    Dim Kapi As Single, Kta As Single
    
    Kapi = PC.ABSSumarconSubcuentas(15)
    Kta = -PC.GetSaldo(PC.GetIDCuenta("Aporte Participación " + nSocio))
    
    GetParticipacion = Kta / Kapi
End Function

Public Function EstaHabilitado(Evento As Long, Optional Registro As String = "") As Long
    '0 No Habilitado
    '1 Si Habilitado
    Dim Resp As Long
    
    'veo si el usuario que esta trabajando tiene habilitacion para entrar
    Dim UUs As Long
    UUs = ACC.UltUsuarioIngresado
    
    If ACC.ExisteRelacion(UUs, Evento) = 0 Then
        MsgBox ACC.GetNombre("Usuario", "Usuarios", UUs) + " no está habilitado " + _
            "para ingresar." + vbCrLf + _
            "Debe Cambiar Sesión a la de un usuario habilitado", vbExclamation, "Atención"
        Resp = 0
    Else
        'registro el movimiento(xx:xxxxxxxxxxxxxxx)
        If Registro <> "" Then ACC.RegEvento UUs, Evento, Registro
        Resp = 1
    End If
    
    EstaHabilitado = Resp
End Function

Public Sub LimResInter()
    Dim I As Long
    
    For I = 0 To 10
        RespuestaInter(I) = ""
    Next I
End Sub
