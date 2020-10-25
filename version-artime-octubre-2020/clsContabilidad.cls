VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsContabilidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'este modulo permite crear y administrar un plan de cuenta
'incluye desde las losta de cuentas hasta los movimientos y balances
Private Enum TipoCuenta
    Activo = 1
    Pasivo = 2
    Perdida = 3
    Ganancia = 4
    PNeto = 5
End Enum

Private mArchMDB As String
Private mCN As New ADODB.Connection
Private mPsW As String

Public Property Let PSW(newPSW As String)
    mPsW = newPSW
End Property

Public Property Get ArchMDB() As String
    ArchMDB = mArchMDB
End Property

Public Property Let ArchMDB(NewArch As String)
    mArchMDB = NewArch
End Property

Public Sub Conectar()
    mCN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + _
        mArchMDB + ";Jet OLEDB:Database Password=" + mPsW
    mCN.Open
End Sub

Public Sub CN_CLOSE()
    If mCN.State = adStateOpen Then mCN.Close
End Sub

Public Sub AgregarCuenta(DependeDeCuenta As Long, Nombre As String, _
    Optional Descripcion As String = "")
    If Nombre = "" Or DependeDeCuenta < 0 Then
        MsgBox "Cargue correctamente los datos", vbExclamation, "Atención"
        Exit Sub
    End If
    
    If ExisteNCuenta(Nombre) > 0 Then
        MsgBox "Ya existe una cuenta con ese nombre", vbExclamation, "Atención"
        Exit Sub
    End If
    
    Dim S As String
    S = "Insert Into tblCuentas (IdAnt, Nombre, Descripcion, Modificable) " + _
        " VALUES (" + CStr(DependeDeCuenta) + ",'" + Nombre + "','" + Descripcion + "',0)"

    
    mCN.Execute S
    
End Sub

Public Sub ModificarCuenta(exNombre As String, newNombre As String, _
    Optional Descripcion As String = "")
    
    If newNombre = "" Then
        MsgBox "Cargue correctamente los datos", vbExclamation, "Atención"
        Exit Sub
    End If
    
    Dim tmpExNombre As String
    'por los * que le pone a las cuentas abiertas
    If Right(exNombre, 1) = "*" Then
        tmpExNombre = Left(exNombre, Len(exNombre) - 1)
    Else
        tmpExNombre = exNombre
    End If
    
    If ExisteNCuenta(newNombre) > 0 Then
        If tmpExNombre = newNombre Then
            'nada escribio lo mismo, solo la descripcion
            mCN.Execute "UPDATE tblCuentas SET descripcion = '" + _
                Descripcion + "' WHERE Nombre = '" + tmpExNombre + "'"
            Exit Sub
        Else
            MsgBox "Ya existe una cuenta con ese nombre"
        End If
        Exit Sub
    End If
    
    Dim S As String
    S = "UPDATE tblCuentas SET Nombre = '" + _
        newNombre + "' WHERE Nombre = '" + tmpExNombre + "'"
    mCN.Execute S
    
End Sub

Public Function EliminarCuenta(IdCuenta As Long)
    'XXXX proximamente hacer que se borren las cuentas hijas
    mCN.Execute "DELETE FROM tblCuentas WHERE ID = " + CStr(IdCuenta)
End Function

Public Function GetCuentas(IDBuscado As Long) As String()
    'devuelve todas las cuentas correspondientes a determinado Codigo de cuenta
    'devuelve los hijos de cada nivel. Con el parametro en cero devuelve los de la raiz
    
    'devuelve una matriz con los valores DESDE 1 en adelante si tiene solo cero
    'es por que no haya nada
    
    Dim Ret() As String
    ReDim Ret(0) 'para que devuelva algo por mas que no haya registros
    
    Dim rS As New ADODB.Recordset
    rS.CursorLocation = adUseClient
    rS.Open "Select * from tblCuentas Where IdAnt = " + CStr(IDBuscado), mCN, adOpenStatic, adLockReadOnly
    If rS.RecordCount = 0 Then
        GetCuentas = Ret
        Exit Function
    End If
    
    rS.MoveFirst
    Dim C As Long
    C = 1
    Do While Not rS.EOF
        ReDim Preserve Ret(C)
        Ret(C) = rS.Fields("id")
        rS.MoveNext
        C = C + 1
    Loop
    rS.Close
    Set rS = Nothing
    GetCuentas = Ret
    
End Function

Public Function GetNameCuenta(ID As Long) As String
    Dim rS As New ADODB.Recordset
    rS.CursorLocation = adUseClient
    rS.Open "Select * from tblCuentas Where Id = " + CStr(ID), mCN, adOpenStatic, adLockReadOnly
    If rS.RecordCount = 0 Then
        GetNameCuenta = "NO EXISTE"
        Exit Function
    End If
    
    GetNameCuenta = rS.Fields("nombre")
    rS.Close
    Set rS = Nothing
    
End Function

Public Function GetIDCuenta(Nombre As String) As Long
    Dim rS As New ADODB.Recordset
    rS.CursorLocation = adUseClient
    rS.Open "Select * from tblCuentas Where Nombre = '" + Nombre + "'", _
        mCN, adOpenStatic, adLockReadOnly
    If rS.RecordCount = 0 Then
        GetIDCuenta = -1
        Exit Function
    End If
    
    GetIDCuenta = rS.Fields("id")
    rS.Close
    Set rS = Nothing
    
End Function

Public Function GetDetalle(IdCuenta As Long) As String
    Dim rS As New ADODB.Recordset
    
    rS.CursorLocation = adUseClient
    rS.Open "Select * from tblCuentas Where Id = " + CStr(IdCuenta), _
        mCN, adOpenStatic, adLockReadOnly
    If rS.RecordCount = 0 Then
        GetDetalle = "NO EXISTE"
        Exit Function
    End If
    
    GetDetalle = NoNuloS(rS.Fields("Descripcion"))
    rS.Close
    Set rS = Nothing

End Function

Public Function EDITAR()
    frmEditarCuentas.Show 1
End Function

Private Sub Class_Terminate()
    If mCN.State = adStateOpen Then mCN.Close
End Sub

Private Sub AgregarMov(IdCta As Long, Var As Single, FECHA As Date)
    mCN.Execute "Insert Into Tblmovimientos (IdCuenta,Variacion,Fecha) VALUES " + _
        "(" + CStr(IdCta) + "," + Replace(CStr(Var), ",", ".") + ",#" + _
        stFechaSQL(FECHA) + "#)"
End Sub

'acepta valores negativos para no tener que hacer
'if para ver si se paga o se cobra por ejemplo
Public Function Asiento(DebitosSepPorBarras As String, MontosD As String, _
    CreditosSepPorBarras As String, MontosC As String) As Long
        '0 JOIIIIIIA
        '1 si alguno de los argumentos es vacio
        '2 si no coinciden las cantidades de montos y de cuentas
        '3 si debitos y creditos no son iguales
    
    Dim Debitos() As String, MDebitos() As String
    Dim Creditos() As String, MCreditos() As String
    
    If DebitosSepPorBarras = "" Or CreditosSepPorBarras = "" _
        Or MontosD = "" Or MontosC = "" Then
        
        Asiento = 1 'algo se cargo mal
        Exit Function
    End If
    
    Debitos = Split(DebitosSepPorBarras, "/")
    Creditos = Split(CreditosSepPorBarras, "/")
    MDebitos = Split(MontosD, "/")
    MCreditos = Split(MontosC, "/")
    
    If UBound(Debitos) <> UBound(MDebitos) Or UBound(Creditos) <> UBound(MCreditos) Then
        Asiento = 2 'falta algun monto para alguna cuenta
        Exit Function
    End If
    
    'controlo que montos de debitos y creditos sean iguales
    If ControlDyC(MDebitos, MCreditos) = False Then
        Asiento = 3 'debitos y creditos no son iguales
        Exit Function
    End If
    
        'registro debitos
    For I = 0 To UBound(Debitos)
        If Abs(CSng(MDebitos(I))) > 0.001 Then
            AgregarMov CLng(Debitos(I)), CSng(MDebitos(I)), Date
        End If
    Next
        
        'registo creditos
    For I = 0 To UBound(Creditos)
        If Abs(CSng(MCreditos(I))) > 0.001 Then
            AgregarMov CLng(Creditos(I)), -CSng(MCreditos(I)), Date
        End If
    Next I
    
End Function

'saca los codigos de los 3 niveles separados por barras por ej 013/024/000
Private Function ControlDyC(Debitos() As String, Creditos() As String) As Boolean
    Dim Ii As Long, jj As Long, Deb As Single, Cred As Single
    Deb = 0
    Cred = 0
    
    ControlDyC = False
    
    For Ii = 0 To UBound(Debitos)
        Deb = Deb + CSng(Debitos(Ii))
    Next
    For jj = 0 To UBound(Creditos)
        Cred = Cred + CSng(Creditos(jj))
    Next

    'puede haber una diferencia como maximo de 50 cents
    If Abs(Deb - Cred) < 0.5 Then ControlDyC = True
    
End Function


Public Sub ResumirAsientos()
    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    'resumo todo lo que haya en tblmovimientos (librodiario) a 1 solo asiento
    
    Dim SS As String
    Dim RsR As New ADODB.Recordset
    
    'XXXX el nuevo programa de ravanillo no tiene idmov VER!!!!!
    'CUANDO IdMov es cero ya tuvo un cierre de caja
    SS = "SELECT tblMovimientos.IdCuenta, Sum(tblMovimientos.Variacion) AS " + _
        "SumaDeVariacion From tblMovimientos GROUP BY tblMovimientos.IdCuenta"
    
    RsR.Open SS, mCN, adOpenStatic, adLockOptimistic
    
    If RsR.RecordCount = 0 Then
        RsR.Close
        Set RsR = Nothing
        Exit Sub
    End If
    
    RsR.MoveFirst
    Do While Not RsR.EOF
        'primero le pongo una fecha rara para identificarlos al toque la borro
        mCN.Execute "INSERT INTO tblMovimientos (IdCuenta,Variacion,Fecha) " + _
            " VALUES (" + CStr(RsR("idcuenta")) + "," + _
            Replace(CStr(RsR("SumaDeVariacion")), ",", ".") + ",#10/14/1492#)"
            
        RsR.MoveNext
    Loop
    
    RsR.Close
    Set RsR = Nothing
    
    'ahora borro todos los asientos individuales entre esas fechas
    mCN.Execute "DELETE FROM tblmovimientos WHERE fecha <> #10/14/1492#"
    
    'ahora si a los que quedaron (los resumidos) les pongo la fecha de hoy
    mCN.Execute "UPDATE tblMovimientos SET Fecha = #" + stFechaSQL(Date) + "# " + _
        "WHERE fecha = #10/14/1492#"
End Sub

Public Function ExisteNCuenta(NombreCuenta As String) As Long
    '0 si no hay ninguna cuenta con ese nombre
    'otro nro si si
    Dim RSBusca As New ADODB.Recordset
    Dim Res As Long
    
    RSBusca.Open "SELECT Nombre FROM tblCuentas WHERE nombre = '" + _
        NombreCuenta + "'", mCN, adOpenStatic, adLockReadOnly
    ExisteNCuenta = RSBusca.RecordCount
    RSBusca.Close
    Set RSBusca = Nothing
    
End Function

Public Function FechaUltimoCierre() As Date
    Dim rsF As New ADODB.Recordset
    rsF.Open "SELECT TOP 1 Fecha FROM BalanceDiario ORDER BY IdCierre DESC", _
        mCN, adOpenStatic, adLockReadOnly
        
    If rsF.RecordCount = 0 Then
        FechaUltimoCierre = #1/1/1990#  'alguna fecha vieja total es solo para resumir todos los
    Else
        rsF.MoveFirst
        FechaUltimoCierre = CDate(rsF("Fecha"))
    End If
    
    rsF.Close
    Set rsF = Nothing
End Function

Public Function CierreCaja() As Long 'va a devolver el id de este cierre
    'primero busco el ultimo cierre de caja que se puede ver donde quedo el ultimo
    'balanceDiario apartir de ahi cancelo todas las cuentas de resultado y las asigno
    'a resultados del ejercicio
    
    'primero resumo todo a un asiento
    Dim ResEje As Single 'va a sumar con las perdidas despues lo invierto
    Dim IdCierre As Long
    
    IdCierre = UltIDCierreMasUno
    CierreCaja = IdCierre
    ResEje = 0 'inicializo
    
    'paso todo a un asiento nomas, quedan con la fecha actual
    ResumirAsientos
    
    'voy buscando todas las cuentas
    Dim RSRes As New ADODB.Recordset 'todas las cuentas
    Dim RSL As New ADODB.Recordset   'con los mov de c/u
    Dim strTabla As String
    
    RSRes.Open "SELECT ID FROM tblcuentas", mCN, adOpenStatic, adLockReadOnly
    'no deberia pasar pero......
    If RSRes.RecordCount = 0 Then Exit Function
    
    RSRes.MoveFirst
    Do While Not RSRes.EOF
        Dim AnOtaR As Single
        AnOtaR = 0
        
        If TipoCuenta(RSRes("id")) = 3 Or TipoCuenta(RSRes("id")) = 4 Then
            strTabla = "ResultadosDiario"
        Else
            strTabla = "BalanceDiario"
        End If
        
        AnOtaR = GetMonto(RSRes("id"), -1, "tblMovimientos")
        'si es patrimonial se acumula con el saldo del cierre anterior
        If strTabla = "BalanceDiario" Then
            AnOtaR = AnOtaR + GetMonto(RSRes("id"), IdCierre - 1)
        End If
        
        If EsCero(AnOtaR) = False Then
           
               mCN.Execute "INSERT INTO " + strTabla + _
                   " (Fecha,IDCierre,IdCuenta,Monto,Estado) " + _
                   "VALUES (#" + stFechaSQL(Date) + "#," + _
                   CStr(IdCierre) + "," + _
                   CStr(RSRes("Id")) + "," + _
                   Replace(AnOtaR, ",", ".") + ",'Abierto')"
               
               If strTabla = "ResultadosDiario" Then ResEje = ResEje + AnOtaR
        End If

        RSRes.MoveNext
    Loop
    
    RSRes.Close
    
    Set RSRes = Nothing
    
    'ahora agrego el resultado NO Asignados que TIENE QUE ser igual a todo lo que
    'cancele de las cuentas de resultado, lo pongo en BalanceDiario
    mCN.Execute "INSERT INTO BalanceDiario (Fecha,IdCierre,IdCuenta,Monto,Estado) " + _
        "VALUES (#" + stFechaSQL(Date) + "#," + CStr(IdCierre) + ",16," + _
        Replace(CStr(ResEje), ",", ".") + ",'Abierto')"
    
    'borro todo del libro diario veamos que pasa
    mCN.Execute "DELETE * FROM tblmovimientos"
    
    'listo deberian quedar todos los libros re re re re joia
    
End Function

Public Function UltIDCierreMasUno(Optional Abierto As Boolean = True) As Long
    Dim rsF As New ADODB.Recordset
    If Abierto Then
        rsF.Open "SELECT TOP 1 IDCierre from BalanceDiario ORDER BY IDCierre desc", _
            mCN, adOpenStatic, adLockReadOnly
    Else 'solo de los cerrados
        rsF.Open "SELECT TOP 1 IDCierre from BalanceDiario WHERE Estado = 'Cerrado' " + _
            "ORDER BY IDCierre desc", _
            mCN, adOpenStatic, adLockReadOnly
    End If
    
    If rsF.RecordCount = 0 Then
        UltIDCierreMasUno = 1
    Else
        UltIDCierreMasUno = CLng(rsF("idcierre")) + 1
    End If
    
    rsF.Close
    Set rsF = Nothing
    
End Function

Public Function TipoCuenta(IdCta As Long) As Long
    'busca el id padre
    '1 activo, 2 pasivo, 3 perdida,4 ganancia, 5 PN
    Dim rsQ As New ADODB.Recordset
    Dim tmp As Long 'graba el idant
    
    If IdCta < 5 Then
        TipoCuenta = IdCta
        Set rsQ = Nothing
        Exit Function
    End If
    
    rsQ.Open "SELECT Id,IdAnt FROM tblCuentas WHERE Id = " + CStr(IdCta), _
        mCN, adOpenStatic, adLockReadOnly
    
    If rsQ.RecordCount = 0 Then
        TipoCuenta = IdCta
    Else
        If rsQ("IdAnt") = 0 Then
            TipoCuenta = IdCta
        Else
            TipoCuenta = TipoCuenta(rsQ("IdAnt"))
        End If
    End If
    
    rsQ.Close
    Set rsQ = Nothing
End Function


've la variacion en comparacion con el ultimo cierre de caja de balance diario
Public Function UltVariacionCuenta(IdCuenta As Long, _
    IdCierre As Long) As Single
    
    UltVariacionCuenta = ABSSumarconSubcuentas(IdCuenta, "BalanceDiario", IdCierre) - _
        ABSSumarconSubcuentas(IdCuenta, "BalanceDiario", IdCierre - 1)
    
End Function

've la variacion en comparacion con el ultimo cierre de caja con el que tengo ahora
Public Function UltVariacionCuentaUlt(IdCuenta As Long, _
    IdCierre As Long) As Single
    
    UltVariacionCuentaUlt = SaldoActual(IdCuenta) - _
        ABSSumarconSubcuentas(IdCuenta, "BalanceDiario", IdCierre)
    
End Function


Public Function SumarConSubcuentas(ByVal IdCta As Long, Tabla As String, _
    Optional IdCierre As Long = -1) As Single
    'si idcierre = -1 busca el ultimo cierre
    'si idcierre = -2 toma todos los cierres abiertos (sirve para resultados que
    ' no acumula
    
    Dim Hijos() As String
    Dim Suma As Single
    'Dim TipoC As Long
    
    Suma = 0
    'TipoC = TipoCuenta(IdCta)
    If IdCierre = -1 Then IdCierre = UltIDCierreMasUno - 1
    
    'sumo el padre y los hijos los hago que se sumen con esta misma funcion
    Suma = GetMonto(IdCta, IdCierre, Tabla)
    
    'ahora busco los hijos que tengan en idant el id de la cuenta
    Hijos = GetCuentas(IdCta)
    Dim jj As Long
    
    If UBound(Hijos) > 0 Then
        For jj = 1 To UBound(Hijos)
            Suma = Suma + SumarConSubcuentas(Hijos(jj), Tabla, IdCierre)
        Next
    End If
    
    SumarConSubcuentas = Suma

End Function

Public Function ABSSumarconSubcuentas(IdCta As Long, Tabla As String, _
    Optional IdCierre As Long = -1) As Single
    Dim TipoC As Long
    Dim Suma As Single
    
    TipoC = TipoCuenta(IdCta)
    Suma = SumarConSubcuentas(IdCta, Tabla, IdCierre)
    
    Select Case TipoC
        Case 1, 3 'activo y perdida deben tener saldos deudores (positivo en librodiario)
            ABSSumarconSubcuentas = Suma 'no hago nada mantengo su signo
        Case 2, 4, 5
            ABSSumarconSubcuentas = -Suma 'cambio el signo de los saldos
    End Select

End Function

Public Function GetMonto(IdCta As Long, IdCierre As Long, _
    Optional Tabla As String = "BalanceDiario") As Single
    
    Dim rS As New ADODB.Recordset
    Dim Campo As String, sWhere As String
    
    If Tabla = "tblMovimientos" Then
        Campo = "Variacion"
        sWhere = "IdCuenta = " + CStr(IdCta)
    Else
        Campo = "Monto"
        If IdCierre = -2 Then 'tomo todos los cierres abiertos
            sWhere = "IdCuenta = " + CStr(IdCta) + " AND Estado = 'Abierto'"
        Else
            sWhere = "IdCuenta = " + CStr(IdCta) + " AND IdCierre = " + CStr(IdCierre)
        End If
    End If
    
    rS.Open "SELECT " + Campo + " FROM " + Tabla + " WHERE " + sWhere, mCN, _
        adOpenStatic, adLockReadOnly
    
    Select Case rS.RecordCount
        Case 0
            GetMonto = 0
        Case 1
            GetMonto = rS(Campo)
        Case Else
            GetMonto = SumarValInRS(Tabla, Campo, sWhere)
    
    End Select
    
    rS.Close
    Set rS = Nothing
End Function

Public Function SaldoActual(IdCuenta As Long) As Single
    'lo que habia en balance (o resultados) MAS lo que se movio en el ejercicio
    
    Dim TipoC As Long, Suma As Single, Tabla As String
    
    TipoC = TipoCuenta(IdCuenta)
    
    Select Case TipoC
        Case 1, 2, 5 'patrimoniales
            Tabla = "BalanceDiario"
        Case 3, 4 'de resultado
            Tabla = "ResultadosDiario"
    End Select
    
    Dim A As Single, b As Single
    A = SumarConSubcuentas(IdCuenta, Tabla, -1)
    b = SumarConSubcuentas(IdCuenta, "tblMovimientos")
    Suma = A + b
    
    'lo dejo en positivo
    Select Case TipoC
        Case 1, 3 'cuentas deudoras
            SaldoActual = Suma
        Case 2, 4, 5 'cuentas acreedoras
            SaldoActual = -Suma
    End Select

End Function

Public Sub CerrarResultados(IdCierre As Long)
    Dim S As String, rsER As New ADODB.Recordset
    '!!!!!!!!!!RESULTADOS !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    
    S = "SELECT idCuenta, Sum(Monto) AS SumaDeMonto " + _
        "FROM ResultadosDiario " + _
        "WHERE Estado = 'Abierto' " + _
        "GROUP BY IdCuenta"
    rsER.Open S, mCN, adOpenStatic, adLockReadOnly
    
    'copio todo el recordset con estado=cerrado
    If rsER.RecordCount > 0 Then
        rsER.MoveFirst
        Do While Not rsER.EOF
            mCN.Execute "INSERT INTO ResultadosDiario " + _
                "(Fecha,IDCierre,IdCuenta,Monto,Estado) " + _
                "VALUES (#" + stFechaSQL(Date) + "#," + _
                CStr(IdCierre) + "," + _
                CStr(rsER("IdCuenta")) + "," + _
                Replace(CStr(rsER("sumadeMonto")), ",", ".") + ",'Cerrado')"
            rsER.MoveNext
        Loop
        
        rsER.Close
        Set rsER = Nothing
        'y borro todos los subperiodos
        mCN.Execute "DELETE FROM ResultadosDiario WHERE Estado = 'Abierto'"
    
    '!!!!!!!!!!!!!!!BALANCE!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    'solo cambio a cerrado el ultimo cierre y  los anteriores
    'le pongo el mismo idcierre que result
        Dim exID As Long
        exID = UltIDCierreMasUno - 1
        
    'el ultimo idcierre se hizo si o si con el cierre de caja lo cambio
    'a este para que coincida con el de resultados
        'primero paso todos a cerrado
        mCN.Execute "UPDATE BalanceDiario SET Estado = 'Cerrado'"
        'ahora le pongo el id
        mCN.Execute "UPDATE BalanceDiario SET IdCierre = " + CStr(IdCierre) + _
            " WHERE IdCierre = " + CStr(exID)
    End If

End Sub

Public Function ListarResultados(IdCierre As Long, Lst As Object) As Single
    'devuelve la utilidad neta
    Dim rsER As New ADODB.Recordset
    Dim UtNeta As Single
    Dim Ventas As Single, CtoVentas As Single, Otros As Single, Alquileres As Single
    
    Ventas = 0: CtoVentas = 0: Otros = 0: UtNeta = 0: Alquileres = 0
    
    Ventas = ABSSumarconSubcuentas(17, "ResultadosDiario", IdCierre)
    CtoVentas = -ABSSumarconSubcuentas(18, "ResultadosDiario", IdCierre)
    Alquileres = ABSSumarconSubcuentas(30, "ResultadosDiario", IdCierre)
    
    rsER.Open "SELECT Id FROM TblCuentas WHERE idAnt = 3 or idAnt=4", _
        mCN, adOpenStatic, adLockReadOnly
    'deberia haber si o si pero... cuidado
    
    Lst.AddItem "ALQUILERES " + FormatCurrency(Alquileres)
    Lst.AddItem "VENTAS " + FormatCurrency(Ventas)
    Lst.AddItem "COSTO DE VENTAS " + FormatCurrency(CtoVentas)
    Lst.AddItem " "
    Lst.AddItem "GANANCIA BRUTA " + FormatCurrency(Alquileres + Ventas + CtoVentas)
    Lst.AddItem " "
    
    'ahora meto los demas
    rsER.MoveFirst
    Do While Not rsER.EOF
        Select Case rsER("id")
            Case 17, 18, 30, 31 'ventas y costo de ventas y alquileres y resultados
                'nada ya esta hecho
            Case Else
                Dim tmpMonto As Single
                tmpMonto = -SumarConSubcuentas(rsER("id"), _
                    "ResultadosDiario", IdCierre)
                
                If EsCero(tmpMonto) = False Then
                    Lst.AddItem GetNameCuenta(rsER("id")) + _
                        " " + FormatCurrency(tmpMonto)
                End If
                'por las dudas lo dejo por ahi son moneditas que hace que quede
                'redondo, igual es raro que pase
                Otros = Otros + tmpMonto
        End Select
        rsER.MoveNext
    Loop
    
    UtNeta = Ventas + CtoVentas + Otros
    
    rsER.Close
    Set rsER = Nothing
    
    Lst.AddItem " "
    Lst.AddItem "GANANCIA NETA " + FormatCurrency(UtNeta)
    
    ListarResultados = UtNeta
End Function

Public Function Distribuir(IdCierre As Long, Socios As Single, Capitalizar As Single)
    'registro la distribucion
    mCN.Execute "INSERT INTO Distribuciones (IdCierre,Socios,Capitalizar) VALUES (" + _
        CStr(IdCierre) + "," + Replace(CStr(Socios), ",", ".") + "," + _
        Replace(CStr(Capitalizar), ",", ".") + ")"

End Function

Public Function VerDistribucion(IdCierre As Long, Cuenta As String) As Single
    'cuenta es socios o capitalizar, por ahora
    
    VerDistribucion = GetValInRS("Distribuciones", Cuenta, _
        "IdCierre = " + CStr(IdCierre), False)

End Function

Public Function ListarCierreCaja(Lst As Object, IdCierre As Long) As String()
    'da una matriz cada indice da:
    '0: da como resultado aumento o disminucion de caja
    '1: el string con el resumen
    Dim tmp(1) As String, tmP2 As Long
    
    'empiezo resumiendo los asientos del libro diario
    ResumirAsientos
    
    'ahora veo la difencia en clientes para ver cuanto de los alquileres y
    'ventas realmente cobradas lo mismo tengo que hacer con proveedores
    'gastos a pagar para ver los movimientos de caja todavia no tiene eso el programa
    'por eso va esto nomas ir mejorandolo a medida que se agregan estos movimientos
    'principales cuentas a analizar (SIC V)-BsCambio,Cred x Vtas,DsComerciales,Gastos
    'a pagar y Deudas Bancarias
    'por ahora veo las 2 primeras usando Mercaderia en Inventario y Clientes
    
    Dim stRes As String 'para guardar el cierre
    Dim VtasCob As Single, CprasPag As Single, SdosPag As Single, DivDist As Single, _
        Otros As Single
    VtasCob = 0: CprasPag = 0: SdosPag = 0: DivDist = 0: Otros = 0
    
    '1ro veo ventas cobradas(usando clientes)
    Dim DifCli As Single, Incob As Single, Alqui As Single
    DifCli = UltVariacionCuentaUlt(46, IdCierre)
     'es ventas+alquil-var clientes , las cuentas de ganancia estan negativos
     'incobrables tengo que sacarlo ya que modifica cuenta clientes pero no por efvo
    Incob = ABSSumarconSubcuentas(28, "tblMovimientos", IdCierre)
    Alqui = ABSSumarconSubcuentas(30, "tblMovimientos", IdCierre)
    
    VtasCob = ABSSumarconSubcuentas(17, "tblMovimientos", IdCierre) + _
        Alqui - DifCli - Incob
        
    '2do veo compras pagadas (usando cto vta, mercaderia y proveedores
    'cuidado con los ajustes de stock los resto
    Dim DifStk As Single, DifPro As Single, AjustesStock As Single
    Dim RFyT As Single, PDanadas As Single, MercACtoEmp As Single
    Dim MercACtoSoc As Single
    
    AjustesStock = ABSSumarconSubcuentas(35, "tblMovimientos", IdCierre)
    PDanadas = ABSSumarconSubcuentas(145, "tblMovimientos", IdCierre)
    RFyT = ABSSumarconSubcuentas(23, "tblMovimientos", IdCierre)
    DifStk = UltVariacionCuentaUlt(54, IdCierre) - AjustesStock - RFyT + PDanadas
    DifPro = UltVariacionCuentaUlt(41, IdCierre)
    
    CprasPag = -(ABSSumarconSubcuentas(18, "tblMovimientos", IdCierre) + _
        DifStk - DifPro)
    
    '3ro veo sueldos pagados(usando las cuentas particulares de empleados)
    Dim DifEmp As Single
    DifEmp = UltVariacionCuentaUlt(53, IdCierre)
     'es sdos - variacion de sus ctas part, las cuentas de perdida estan positivas
    SdosPag = -(ABSSumarconSubcuentas(36, "tblMovimientos", IdCierre) - _
        DifEmp)
        
    '4to veo extracciones de socio (usando sus cuentas particulares)
    Dim DifSoc As Single, GanDist As Single
    DifSoc = UltVariacionCuentaUlt(52, IdCierre)
    
    GanDist = VerDistribucion(IdCierre, "Socios")
    
    DivDist = DifSoc - GanDist
    
    'xxxx pendiente los demas que sean a pagar como gastos, luz y cosas asi
    Dim zz As String, zy As String, zx As String, zw As String
    
    Lst.Clear
    Lst.AddItem "MÁS:"
    zz = "Vtas y Alqu. Cobrados: " + FormatCurrency(VtasCob)
    Lst.AddItem zz
    Lst.AddItem ""
    Lst.AddItem "MENOS:"
    zy = "Compras Pagadas: " + FormatCurrency(CprasPag)
    Lst.AddItem zy
    zx = "Empleados: " + FormatCurrency(SdosPag)
    Lst.AddItem zx
    zw = "Socios: " + FormatCurrency(DivDist)
    Lst.AddItem zw
    Lst.AddItem ""
    Lst.AddItem "OTROS:"
    
    stRes = zz + "\\" + zy + "\\" + zx + "\\" + zw + "\\"
        
    'ahora hago un bucle con todas las cuentas de resultado partiendo de que
    'supuestamente trabajan por lo percibido
    
    Dim rsO As New ADODB.Recordset
    rsO.Open "SELECT idcuenta,Variacion FROM tblMovimientos", mCN, adOpenStatic, adLockReadOnly
    If rsO.RecordCount > 0 Then
        rsO.MoveFirst
        Do While Not rsO.EOF
            tmP2 = TipoCuenta(rsO("idcuenta"))
            If tmP2 = 3 Or tmP2 = 4 Then
                If EsCaja(rsO("IdCuenta")) = 1 Then
                    'nada
                Else
                    Dim sTM As String

                    sTM = GetNameCuenta(rsO("idcuenta")) + _
                        ": " + FormatCurrency(-rsO("Variacion"))
                    If Left(sTM, 6) <> "Sueldo" Then
                        Lst.AddItem sTM
                        Otros = Otros + rsO("Variacion")

                        stRes = stRes + sTM
                        If Not rsO.EOF Then stRes = stRes + "\\"
                    End If
                End If
                
'                Select Case rsO("idcuenta")
'                    Case 17, 18, 23, 28, 31, 35, 36, 81, 82, 83, 94, 95, 96, 97, _
'                        141, 142, 143, 144, 145, 160
'                        'nada
'                    Case Else
'                        Dim sTM As String
'
'                        sTM = GetNameCuenta(rsO("idcuenta")) + _
'                            ": " + FormatCurrency(-rsO("Variacion"))
'                        If Left(sTM, 6) <> "Sueldo" Then
'                            Lst.AddItem sTM
'                            Otros = Otros + rsO("Variacion")
'
'                            stRes = stRes + sTM
'                            If Not rsO.EOF Then stRes = stRes + "\\"
'                        End If
'                End Select
            End If
            rsO.MoveNext
        Loop
    End If
    
    rsO.Close
    Set rsO = Nothing
    
    tmp(0) = CStr(VtasCob + CprasPag - Otros + SdosPag + DivDist)
    tmp(1) = stRes
    
    ListarCierreCaja = tmp
End Function

Public Sub GrabarCierre(IdCierre As Long, Efvo As Single, stResumen As String, _
    Dif As Single, Var As Single)
    'dejo grabado este cierre
    mCN.Execute "INSERT INTO CierresViejos (IDCierre,Fecha,Hora,Efvo,sCierre," + _
        "Diferencia,Variacion) " + _
        "VALUES (" + CStr(IdCierre) + ",#" + stFechaSQL(Date) + "#,'" + _
        CStr(Time) + "'," + _
        Replace(CStr(Efvo), ",", ".") + ",'" + stResumen + "'," + _
        Replace(CStr(Dif), ",", ".") + "," + _
        Replace(CStr(Var), ",", ".") + ")"
End Sub

Public Function GetCierre(IdCierre As Long) As String
    Dim rS As New ADODB.Recordset
    Dim stResumen As String
    
    rS.Open "SELECT sCierre FROM CierresViejos WHERE IdCierre = " + CStr(IdCierre), _
        mCN, adOpenStatic, adLockReadOnly
    
    If rS.RecordCount = 0 Then
        stResumen = "\\"
    Else
        stResumen = rS("sCierre")
    End If
    
    rS.Close
    Set rS = Nothing
    
    GetCierre = stResumen
    
End Function

Public Function CerroCaja() As Boolean
    'primero obligo que cierre de caja antes de hacer resultados
    CerroCaja = True 'predeterminado
    
    Dim rsV As New ADODB.Recordset
    rsV.Open "SELECT idcuenta FROM tblMovimientos ORDER BY id desc", _
        mCN, adOpenStatic, adLockReadOnly
    
    If rsV.RecordCount > 0 Then
        rsV.MoveFirst
        Do While Not rsV.EOF 'veo cuenta por cuenta si no hay de resultados se cerro
            If TipoCuenta(rsV("idcuenta")) = 3 Or TipoCuenta(rsV("idcuenta")) = 4 Then
                CerroCaja = False
                Exit Do
            End If
            rsV.MoveNext
        Loop
    End If
End Function

Public Sub ListarEgresos(Lst As Object, Top As Long)
    Dim RsEg As New ADODB.Recordset
    Dim Ttop As Long, tmp As Long
    
    'va a ir restando a medida que vaya agregando
    Ttop = Top
    
    Lst.Clear
    RsEg.Open "SELECT * FROM tblMovimientos " + _
        "ORDER BY Id desc", mCN, adOpenStatic, adLockReadOnly
        '"WHERE Cod1 = 15 OR " + _
        '"Cod1=16 OR Cod1=18 OR Cod1=33 OR Cod1=34 ORDER BY Id", DB.CN, adOpenStatic, adLockReadOnly
    'codigos de egresos 15:Comercializacion, 16: Administ, 18:Otros, 33:Imp y 34: serv
    If RsEg.RecordCount > 0 Then
        RsEg.MoveFirst
        Do While Not (RsEg.EOF Or Ttop = 0)
            tmp = TipoCuenta(RsEg("IdCuenta"))
            If tmp = 3 Or tmp = 4 Then
                If EsCaja(RsEg("idCuenta")) = 1 Then
                        'nada
                Else
                    Lst.AddItem CStr(RsEg("Fecha")) + "\" + _
                        GetNameCuenta(RsEg("idcuenta")) + "\" + _
                        FormatCurrency(RsEg("Variacion"))
                    Ttop = Ttop - 1
                End If
            End If
            
            RsEg.MoveNext
        Loop
    End If
    
    RsEg.Close
    Set RsEg = Nothing

End Sub

Private Function EsCaja(IdCuenta As Long) As Long
    'para cuentas de resultados, devuelve 1 si no genera caja, 0 si si
    Dim tmp As Long
    Dim rS As New ADODB.Recordset
    
    tmp = 0 ' predeterminado
    
    rS.Open "SELECT Caja FROM tblCuentas WHERE ID = " + CStr(IdCuenta), mCN, adOpenStatic, adLockReadOnly
    
    If rS.RecordCount > 0 Then
        tmp = rS("Caja")
    End If
    
    EsCaja = tmp
    
    rS.Close
    Set rS = Nothing
End Function

Public Function GetRsCierres(Cantidad As Long) As ADODB.Recordset
    Dim rS3 As New ADODB.Recordset
    
    If rS3.State = adStateOpen Then rS3.Close
    
    rS3.CursorLocation = adUseClient
    rS3.Open "SELECT TOP " + CStr(Cantidad) + " * FROM CierresViejos ORDER BY ID DESC", _
        mCN, adOpenStatic, adLockReadOnly
    
    Set GetRsCierres = rS3
    'rS.Close
    'Set rS = Nothing
End Function

Public Sub LimpiarCierres(DiasAtras As Long)
    mCN.Execute "DELETE FROM CierresViejos WHERE Fecha <#" + _
        stFechaSQL(Date - DiasAtras) + "#"
End Sub

Public Sub LimpiarContabilidad(DiasAtras As Long)
    mCN.Execute "DELETE FROM ResultadosDiario WHERE Fecha <#" + _
        stFechaSQL(Date - DiasAtras) + "# AND Estado = 'Cerrado'"
    mCN.Execute "DELETE FROM BalanceDiario WHERE Fecha <#" + _
        stFechaSQL(Date - DiasAtras) + "# AND Estado = 'Cerrado'"
End Sub

'!!!!!!!!!!!!!!!!!!!!!!GLOBALES DE LA DLL !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Public Function EsCero(Nro As Single) As Boolean
    If Abs(Nro) < 0.01 Then
        EsCero = True
    Else
        EsCero = False
    End If
End Function

Public Function GetValInRS(Tabla As String, Campo As String, sWhere As String, _
    Optional EsString As Boolean = True)
    'devuelve el valor de un campo en una tabla en  un campo especificado dada
    'determinada condicion (supongo que es solo 1)
    
    Dim rS As New ADODB.Recordset
    rS.Open "Select " + Campo + " FROM " + Tabla + " WHERE " + sWhere, _
        mCN, adOpenStatic, adLockReadOnly
    If rS.RecordCount = 0 Then
        If EsString Then
            GetValInRS = ""
        Else
            GetValInRS = 0
        End If
    Else
        If EsString Then
            GetValInRS = NoNuloS(rS.Fields(Campo))
        Else
            GetValInRS = NoNuloN(rS.Fields(Campo))
        End If
    End If
    
    rS.Close
    Set rS = Nothing
End Function

Public Function SumarValInRS(Tabla As String, Campo As String, _
    sWhere As String) As Single
    
    'devuelve la suma de un campo en una tabla en  un campo especificado dada
    'determinada condicion
    Dim Suma As Single
    Suma = 0
    
    Dim rS As New ADODB.Recordset
    rS.Open "Select " + Campo + " FROM " + Tabla + " WHERE " + sWhere, _
        mCN, adOpenStatic, adLockReadOnly
    If rS.RecordCount = 0 Then
        
    Else
        rS.MoveFirst
        Do While Not rS.EOF
            Suma = Suma + NoNuloN(rS.Fields(Campo))
            
            rS.MoveNext
        Loop
        
    End If
    
    SumarValInRS = Suma
    rS.Close
    Set rS = Nothing
End Function

Public Function NoNuloS(S) As String
    If IsNull(S) Then
        NoNuloS = ""
    Else
        NoNuloS = S
    End If
End Function

Public Function NoNuloN(J) As Single
    If IsNumeric(J) Then
        NoNuloN = J
    Else
        NoNuloN = 0
    End If
End Function

Public Function stFechaSQL(FECHA As Date) As String
    Dim h() As String
    h = Split(CStr(FECHA), "/")
    stFechaSQL = h(1) + "/" + h(0) + "/" + h(2)
    
End Function

Public Sub CierresViejos()
    frmCierresViejos.AbrirDatos mArchMDB, mPsW
End Sub

Public Sub CargarCombo(CMB As Object, sqlText As String, _
    CamposSeparadosPorComas As String, Optional Separador As String = "\")
    'ZZZZ pasar al completo
    
    'CamposSeparadosPorComas es una lista separada por comas de los campos. _
        Ademas se le puede agregar al final _
        /n al final para indicar que es numero _
        /f para fechas _
        /$ para currency _
        predeterminado es string
    
    Dim Campos() As String
    Campos = Split(CamposSeparadosPorComas, ",")
        
    Dim rS As New ADODB.Recordset
    rS.Open sqlText, mCN, adOpenStatic, adLockReadOnly
    Dim S As String, AA As Long
    CMB.Clear
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


