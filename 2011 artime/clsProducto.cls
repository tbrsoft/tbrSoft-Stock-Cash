VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub ModificarStock(CodProd As Long, Variacion As Long, _
    Optional Sucursal As String = "CASA CENTRAL", _
    Optional Detalle As String = "PorVenta")
    
    Dim rS As New ADODB.Recordset
    rS.CursorLocation = adUseClient
    rS.Open "SELECT ID, Stock FROM Productos WHERE ID=" + CStr(CodProd), DB.CN, adOpenStatic, adLockOptimistic
    If rS.RecordCount >= 0 Then
        rS.Fields("Stock") = rS.Fields("stock") + Variacion
        rS.Update
    End If
    
    rS.Close
    Set rS = Nothing
    
    'solo modifico "StockOtraSuc" si es distinto de casa central, ya que
    'el stock del mismo se calcula por descarte
    If Sucursal <> "CASA CENTRAL" Then
        Dim exSt As Long, StAh As Long
        '(1) cuento cuanto habia de stock en esa sucursal
        exSt = DB.GetValInRS("StockOtraSuc", "Stock", "IdProducto = " + _
            CStr(CodProd) + " AND Sucursal = '" + Sucursal + "'", False)
        '(2) sumo el cambio
        StAh = exSt + Variacion
        '(3) elimino el registro viejo (no modifico ya que puede pasar que no haya
        DB.EXECUTE "DELETE FROM StockOtraSuc WHERE IdProducto = " + _
            CStr(CodProd) + " AND Sucursal = '" + Sucursal + "'"
        '(4) AGREGO el nuevo registro
        DB.EXECUTE "INSERT INTO StockOtraSuc (IdProducto, Sucursal, Stock) VALUES (" + _
            CStr(CodProd) + ",'" + Sucursal + "'," + CStr(StAh) + ")"
    End If
    
    'capricho de Ravanillo
    'registrar que motivo la variacion (Fecha,IDP,Sucu,Var,StockLuego,Usuario,Detalle)
    Dim StockNow As Long, UUs As Long, Usuario As String
    
    StockNow = StockProductoenSucursal(CodProd, Sucursal)
    UUs = ACC.UltUsuarioIngresado
    Usuario = ACC.GetNombre("Usuario", "Usuarios", UUs)
    
    DB.EXECUTE "INSERT INTO MovimientosProductos (Fecha, Hora, IdProducto, Sucursal," + _
        "Variacion, StockLuego, Usuario, Detalle) VALUES (#" + stFechaSQL(Date) + _
        "#," + CStr(Hour(Now)) + "," + CStr(CodProd) + ",'" + Sucursal + "'," + _
        CStr(Variacion) + "," + CStr(StockNow) + ",'" + Usuario + "','" + Detalle + "')"
        
    
End Sub

Public Function GetCosto(IDProducto As Long) As Single
    Dim RsCT As New ADODB.Recordset
    Dim Cto As Single
    
    RsCT.Open "SELECT pCosto FROM Productos WHERE ID = " + CStr(IDProducto), _
       DB.CN, adOpenStatic, adLockReadOnly
    
    If RsCT.RecordCount > 0 Then
        Cto = NoNuloN(RsCT("pcosto"))
    Else
        Cto = 0
    End If
    
    RsCT.Close
    Set RsCT = Nothing
    GetCosto = Cto
End Function

Public Sub PonerPrecioPorMargen(idProd As Long, Optional Cto As Single = -100)
    Dim IDp As Long, Ct As Single, IdTipoProd As Long, Multip As Single, IDC As Long
    
    IDp = idProd
    IdTipoProd = DB.GetValInRS("Productos", "IdTipoProducto", "ID = " + CStr(IDp), False)
    
    If Cto = -100 Then
        Ct = DB.GetValInRS("Productos", "pCosto", "ID = " + CStr(IDp), False)
    Else
        Ct = Cto
    End If
    
    'si tiene configuracion particular la pongo si no uso la general
    IDC = CFG.ExistePropiedad("MDV " + CStr(IDp))
    
    If IDC = 0 Then
        Multip = CSng(CFG.GetInfo(50, 4)) / 100 + 1
    Else
        Multip = CSng(CFG.GetInfo(IDC, 4) / 100) + 1
    End If
    
    DB.EXECUTE "UPDATE Productos SET pVenta = " + _
        Replace(CStr(Multip * Ct), ",", ".") + _
        " WHERE ID = " + CStr(IDp)
End Sub

Public Function ListarMovTipoProducto(IdTipoProd As Long, Desde As Date, Hasta As Date, _
    Optional LvW As Object, Optional SinListar As Boolean = False) As String
    'object tiene 3 columnas: Concepto, Cantidad e Importe
    'da como resultado el Importe de $$ de ventas + "|"+ Cantidad
    'si idTipoProd es -1 van todos
    
    Dim TmP As Long, S As String
    Dim RsL As New ADODB.Recordset
    Dim Canti As Long, Import As Single
    
    If Hasta = 0 Then Hasta = Date
        
    TmP = 0
    Canti = 0: Import = 0
    If SinListar = False Then LvW.ListItems.Clear
    
    'ahora si vamos con los mov
    If IdTipoProd > 0 Then
        S = "SELECT Ventas.IDproducto, Productos.nProducto, Sum(Ventas.Cantidad) " + _
            "AS SumaCant, Sum([Ventas]![Precio]*[Ventas]![Cantidad]) AS Importe " + _
            "FROM Productos INNER JOIN Ventas ON Productos.ID = Ventas.IDproducto " + _
            "WHERE (((Ventas.Fecha) Between #" + stFechaSQL(Desde) + "# AND #" + _
            stFechaSQL(Hasta) + "#) AND ((Productos.IdTipoProducto)= " + _
            CStr(IdTipoProd) + ")) AND Ventas.IdProducto > 0 " + _
            "GROUP BY Ventas.IDproducto, Productos.nProducto"
            
    Else 'SOLO LOS PRODUCTOS PRINCIPALES NO! TODOS!!!!
        S = "SELECT TipoProductos.ID2, TipoProductos.TipoProducto, Sum(Ventas.Cantidad) " + _
            "AS SumaCant, Sum([Ventas]![Precio]*[Ventas]![Cantidad]) AS Importe " + _
            "FROM TipoProductos INNER JOIN (Productos INNER JOIN Ventas ON " + _
            "Productos.ID = Ventas.IDproducto) ON TipoProductos.ID2 = " + _
            "Productos.IdTipoProducto " + _
            "WHERE (((Ventas.Fecha) BETWEEN #" + stFechaSQL(Desde) + _
            "# And #" + stFechaSQL(Hasta) + "#)) AND TipoProductos.ID2 > 0 " + _
            "GROUP BY TipoProductos.ID2, TipoProductos.TipoProducto"
            
    End If
            
    If RsL.State = adStateOpen Then RsL.Close
    RsL.Open S, DB.CN, adOpenStatic, adLockReadOnly
    
    If RsL.RecordCount > 0 Then
        RsL.MoveFirst
        
        Do While Not RsL.EOF
            'solo si quiere que se liste
            If SinListar = False Then
                TmP = TmP + 1
                LvW.ListItems.Add TmP
                
                If IdTipoProd > 0 Then
                    LvW.ListItems(TmP).Text = CStr(RsL("IDProducto"))
                    LvW.ListItems(TmP).SubItems(1) = CStr(RsL("nProducto"))
                Else
                    LvW.ListItems(TmP).Text = CStr(RsL("ID2"))
                    LvW.ListItems(TmP).SubItems(1) = CStr(RsL("TipoProducto"))
                End If
                
                LvW.ListItems(TmP).SubItems(2) = NoNuloN(RsL("SumaCant"))
                LvW.ListItems(TmP).SubItems(3) = FormatCurrency(NoNuloN(RsL("Importe")), , , , vbFalse)
            End If
            'esto si!!!
            Canti = Canti + NoNuloN(RsL("SumaCant"))
            Import = Import + NoNuloN(RsL("Importe"))
            
            RsL.MoveNext
        Loop
    End If
    
    RsL.Close
    Set RsL = Nothing
    
    ListarMovTipoProducto = CStr(Import) + "|" + CStr(Canti)
End Function

Public Function GetHijoTipo(IdTipoProducto As Long) As String()
    Dim Resp() As String, Ix As Long
    Dim RSh As New ADODB.Recordset
    
    Ix = 0
    ReDim Preserve Resp(Ix)
    Resp(Ix) = "NADA"
    
    If RSh.State = adStateOpen Then RSh.Close
    RSh.Open "SELECT ID2 FROM TipoProductos WHERE IdAnt = " + CStr(IdTipoProducto), DB.CN, adOpenStatic, adLockReadOnly
    
    If RSh.RecordCount > 0 Then
        RSh.MoveFirst
        Do While Not RSh.EOF
            Ix = Ix + 1
            ReDim Preserve Resp(Ix)
            Resp(Ix) = CStr(NoNuloN(RSh("ID2")))
            RSh.MoveNext
        Loop
    End If
    
    RSh.Close
    Set RSh = Nothing
    
    GetHijoTipo = Resp
End Function

Public Function StockProductoenSucursal(IDProducto As Long, Sucursal As String) As Long
    Dim Resp As Long, TmP As Long
    
    TmP = DB.GetValInRS("StockOtraSuc", "Stock", "IdProducto = " + _
            CStr(IDProducto) + " AND Sucursal = '" + Sucursal + "'", False)
    
    If Sucursal = "CASA CENTRAL" Then
        'mas complicado tengo que calcular el stock y restarle todos los de la
        'tabla stockotrasuc donde este el idproducto
        Resp = DB.GetValInRS("Productos", "Stock", "ID = " + CStr(IDProducto), False) - _
            DB.SumarValInRS("StockOtraSuc", "Stock", "IdProducto = " + _
            CStr(IDProducto))
    Else
        Resp = DB.GetValInRS("StockOtraSuc", "Stock", "IdProducto = " + _
            CStr(IDProducto) + " AND Sucursal = '" + Sucursal + "'", False)
    End If
    
    StockProductoenSucursal = Resp
End Function

Public Sub CargarCompra(idProd As Long, Cant As Long, Precio As Single, _
    Optional Sucursal As String = "CASA CENTRAL", _
    Optional Detalle As String = "Por Compra", Optional AjustoPrecio As Boolean = False)
    
    Dim RsPr As New ADODB.Recordset
    Dim CantRS As Long   'van a tener el stock y cto de la tabla
    Dim pCostoRs As Single
    Dim TtRs As Single 'es el stock en $$ antes de compra
    Dim pCostoF As Single 'precio que va a quedar registrado
    
    If RsPr.State = adStateOpen Then RsPr.Close
    
    Dim S As String
    S = "SELECT pCosto, Stock FROM Productos WHERE Id = " + _
        CStr(idProd)
    RsPr.Open S, DB.CN, adOpenDynamic, adLockOptimistic
    
    
    If IsNull(RsPr("Stock")) Or RsPr("Stock") = "" Then
        CantRS = 0
    Else
        CantRS = RsPr("Stock")
    End If
    
    If IsNull(RsPr("pCosto")) Or RsPr("pCosto") = "" Then
        pCostoRs = 0
    Else
        pCostoRs = RsPr("pCosto")
    End If
    
    TtRs = pCostoRs * CantRS
    
    If CantRS + Cant = 0 Then
        'va a pasar esto justo cuando la compra sea de un producto con stock
        'negativo en la misma cantidad que compro por lo tanto es un promedio
        'exacto entre los 2 valores
        pCostoF = (Precio + pCostoRs) / 2
    Else
       
        pCostoF = (TtRs + Cant * Precio) / (CantRS + Cant)
    End If
    
    RsPr("stock") = CantRS + Cant
    RsPr("pcosto") = pCostoF
    
    RsPr.Update
    RsPr.Close
    Set RsPr = Nothing
    
    If AjustoPrecio = True Then PonerPrecioPorMargen idProd, pCostoF
    
    'solo modifico "StockOtraSuc" si es distinto de casa central, ya que
    'el stock del mismo se calcula por descarte
    If Sucursal <> "CASA CENTRAL" Then
        Dim exSt As Long, StAh As Long
        '(1) cuento cuanto habia de stock en esa sucursal
        exSt = DB.GetValInRS("StockOtraSuc", "Stock", "IdProducto = " + _
            CStr(idProd) + " AND Sucursal = '" + Sucursal + "'", False)
        '(2) sumo el cambio
        StAh = exSt + Cant
        '(3) elimino el registro viejo (no modifico ya que puede pasar que no haya)
        DB.EXECUTE "DELETE FROM StockOtraSuc WHERE IdProducto = " + _
            CStr(idProd) + " AND Sucursal = '" + Sucursal + "'"
        '(4) AGREGO el nuevo registro
        DB.EXECUTE "INSERT INTO StockOtraSuc (IdProducto, Sucursal, Stock) VALUES (" + _
            CStr(idProd) + ",'" + Sucursal + "'," + CStr(StAh) + ")"
    End If
    
    'capricho de Ravanillo
    'registrar que motivo la variacion (Fecha,IDP,Sucu,Var,StockLuego,Usuario,Detalle)
    Dim StockNow As Long, UUs As Long, Usuario As String
    
    StockNow = StockProductoenSucursal(idProd, Sucursal)
    UUs = ACC.UltUsuarioIngresado
    Usuario = ACC.GetNombre("Usuario", "Usuarios", UUs)
    
    DB.EXECUTE "INSERT INTO MovimientosProductos (Fecha, Hora, IdProducto, Sucursal," + _
        "Variacion, StockLuego, Usuario, Detalle) VALUES (#" + stFechaSQL(Date) + _
        "#," + CStr(Hour(Now)) + "," + CStr(idProd) + ",'" + Sucursal + "'," + _
        CStr(Cant) + "," + CStr(StockNow) + ",'" + Usuario + "','" + Detalle + "')"
End Sub

Public Sub CargarCompraDt(NroFactura As String, Fecha As Date, Proveedor As String, _
    idProd As Long, Cant As Long, PrTotal As Single)
    Dim IDp As Long
    
    IDp = idProd
    
    DB.CN.EXECUTE "INSERT INTO CompraDetalle (ID,NroFactura,Fecha,Proveedor, " + _
        "IdProducto, Cantidad, PrecioTotal) VALUES (" + CStr(IdAutonum("CompraDetalle")) + _
        ",'" + NroFactura + "', #" + stFechaSQL(Fecha) + "#, '" + _
        Proveedor + "'," + CStr(IDp) + "," + CStr(Cant) + "," + _
        Replace(CStr(PrTotal), ",", ".") + ")"
End Sub

