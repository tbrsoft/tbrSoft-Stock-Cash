VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function GetComprasTiempo(NProve As String, UltNDias As Long) As Single
    'Para que sume el total -> NProve="Todos"
    
    Dim DiaD As Date
    Dim RScc As New ADODB.Recordset
    Dim SX As String
    
    DiaD = Date - UltNDias
        
    If NProve = "Todos" Then
        SX = "SELECT Proveedor, Sum(Pagado) AS Ppreciototal FROM FacturaCompra " + _
            "WHERE EsPedido = 0  AND Fecha > #" + stFechaSQL(DiaD) + "# " + _
            "GROUP BY proveedor"
    Else
        SX = "SELECT Proveedor, Sum(Pagado) as Ppreciototal FROM FacturaCompra " + _
            "WHERE EsPedido = 0 AND Fecha > #" + stFechaSQL(DiaD) + "# " + _
            "AND Proveedor = '" + NProve + "' GROUP BY Proveedor"
    End If
    
    RScc.Open SX, DB.CN, adOpenStatic, adLockReadOnly
    Select Case RScc.RecordCount
        Case 0 'no tiene compras en los ult NN dias
            GetComprasTiempo = 0
        Case 1 'joya
            GetComprasTiempo = CSng(RScc("Ppreciototal"))
        Case Else
            Dim tmpS As Single
            tmpS = 0
            
            RScc.MoveFirst
            Do While Not RScc.EOF
                tmpS = tmpS + CSng(RScc("Ppreciototal"))
                RScc.MoveNext
            Loop
            
            GetComprasTiempo = tmpS
        End Select
    
    RScc.Close
    Set RCC = Nothing
    
End Function

