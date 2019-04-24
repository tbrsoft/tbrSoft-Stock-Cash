VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mCodigo As Long
Private mNombre As String
Private mDetalle As String
Private mDireccion As String
Private mCUIT As String
Private mIVA As String
Private mTelefono As String
Private mMail As String
Private mTieneTrabajo As Long
Private mNacimiento As Date

Public Property Get Codigo() As Long
    Codigo = mCodigo
End Property

Public Property Get Nombre() As String
    Nombre = mNombre
End Property

Public Property Let Nombre(newNombre As String)
    If newNombre = "" Then MsgBox "Debes seleccionar un nombre de cliente"
    mNombre = newNombre
End Property

Public Property Get Detalle() As String
    Detalle = mDetalle
End Property

Public Property Let Detalle(newDetalle As String)
    mDetalle = newDetalle
End Property

Public Property Get Direccion() As String
    Direccion = mDireccion
End Property

Public Property Let Direccion(newDireccion As String)
    mDireccion = newDireccion
End Property

Public Property Get CUIT() As String
    CUIT = mCUIT
End Property

Public Property Let CUIT(newCUIT As String)
    mCUIT = newCUIT
End Property

Public Property Get Telefono() As String
    Telefono = mTelefono
End Property

Public Property Let Telefono(newTelefono As String)
    mTelefono = newTelefono
End Property

Public Property Get Mail() As String
    Mail = mMail
End Property

Public Property Let Mail(newMail As String)
    mMail = newMail
End Property

Public Property Get TieneTrabajo() As Long
    TieneTrabajo = mTieneTrabajo
End Property

Public Property Let TieneTrabajo(newTT As Long)
    mTieneTrabajo = newTT
End Property

Public Property Get Nacimiento() As Date
    Nacimiento = mNacimiento
End Property

Public Property Let Nacimiento(newNac As Date)
    mNacimiento = newNac
End Property

Public Property Get Iva() As String
    Iva = mIVA
End Property

Public Property Let Iva(newIVA As String)
    mIVA = newIVA
End Property

Public Function GetCliente(IdcLiente As Long) As String
    Dim rsTT As New ADODB.Recordset
    
    rsTT.Open "SELECT ID,Nombre FROM clientes WHERE ID=" + CStr(IdcLiente), _
        DB.CN, adOpenStatic, adLockReadOnly
    If rsTT.RecordCount <> 1 Then
        GetCliente = "SIN DATOS"
    Else
        GetCliente = rsTT("Nombre")
    End If
    
    rsTT.Close
    Set rsTT = Nothing
End Function

Public Function GetID(Nombre As String) As Long
    Dim rsTT As New ADODB.Recordset
    
    rsTT.Open "SELECT ID,Nombre FROM clientes WHERE Nombre='" + Nombre + "'", _
        DB.CN, adOpenStatic, adLockReadOnly
    If rsTT.RecordCount <> 1 Then
        GetID = -1
    Else
        GetID = rsTT("id")
    End If
    
    rsTT.Close
    Set rsTT = Nothing
End Function

Public Function GetDeuda(IdcLiente As Long) As Single
    Dim Dda As Single
    
    Dim rSGt As New ADODB.Recordset
    If rSGt.State = adStateOpen Then rSGt.Close
    rSGt.CursorLocation = adUseClient
    rSGt.Open "select variacion from movclientes where CodCliente= " + CStr(IdcLiente), _
        DB.CN, adOpenDynamic, adLockReadOnly
    
    If rSGt.RecordCount = 0 Then 'no debe nada
        GetDeuda = 0
        Exit Function
    End If
    
    rSGt.MoveFirst
    Do While Not rSGt.EOF
        Dda = Dda + CSng(rSGt("variacion"))
        rSGt.MoveNext
    Loop
    rSGt.Close
    Set rSGt = Nothing
    
    GetDeuda = Dda
End Function


Public Sub AbrirDatos(IdcLiente As Long)
    'idcliente es el codigo del que se va a modificar
    'si es -1 es uno nuevo
    mCodigo = IdcLiente
    If IdcLiente <> -1 And IdcLiente <> -20 Then
        Dim rS As New ADODB.Recordset
        rS.Open "select * from clientes WHERE ID = " + CStr(IdcLiente), DB.CN, adOpenForwardOnly, adLockReadOnly
        mNombre = rS("Nombre")
        'txtNombre.SelStart = 0     'en el caso que quiere que aparezca seleccionado
        'txtNombre.SelLength = Len(txtNombre)
        mDetalle = NoNuloS(rS("Detalle"))
        mDireccion = NoNuloS(rS("Direccion"))
        mCUIT = NoNuloS(rS("CUIT"))
        mIVA = NoNuloS(rS("IVA"))
        mTelefono = NoNuloS(rS("Telefono"))
        mMail = NoNuloS(rS("Mail"))
        mTieneTrabajo = Abs(NoNuloN(rS("TieneTrabajo")))
        mNacimiento = NoNuloD(rS("Nacimiento"))
        
        rS.Close
        Set rS = Nothing
    End If
End Sub

Public Function Grabar(isNew As Boolean, Optional EsFinanciera As Boolean = False) As Long
    '1 repetido
    '0 todo OK
    On Local Error GoTo ErrDuplic
    
    Dim RCC As Long, IDC As Long

DeGuelta:
    RCC = DB.ContarReg("SELECT ID FROM Clientes WHERE Nombre= '" + mNombre + "'")
    
    If isNew = True Then
        If EsFinanciera = False Then 'es cliente
            IDC = IdAutonum("Clientes")
        Else 'es financiera
            If DB.ContarReg("SELECT ID FROM Clientes WHERE ID<-20") = 0 Then
                IDC = -21
            Else
                IDC = DB.GetTop1Rs("Clientes", "ID", "ASC", , False) - 1
            End If
        End If
    Else
        IDC = DB.GetValInRS("Clientes", "ID", "Nombre = '" + mNombre + "'", False)
    End If
    
    If isNew Then
        If RCC >= 1 Then
            Grabar = 1
            Exit Function
        End If
            
        DB.EXECUTE "INSERT INTO Clientes (ID,Nombre,Direccion,CUIT,IVA," + _
            "Detalle,Telefono,Mail,TieneTrabajo,Nacimiento) " + _
            "VALUES (" + CStr(IDC) + ",'" + mNombre + "','" + _
            NoNuloS(mDireccion) + "','" + NoNuloS(mCUIT) + "','" + _
            NoNuloS(mIVA) + "','" + NoNuloS(mDetalle) + "','" + _
            NoNuloS(mTelefono) + "','" + NoNuloS(mMail) + "'," + _
            CStr(NoNuloN(mTieneTrabajo)) + ",#" + stFechaSQL(mNacimiento) + "#)"
    
    Else 'esta modificando por lo que nombre se puede repetir pero solo con el mismo iD
        If RCC >= 1 Then
            If IDC <> mCodigo Then
                Grabar = 1
                Exit Function
            End If
            
        End If
        
        'SEGUIRAQUI
        'LA BASE DE DATOS ATERIOR AL 26/03/09 el detalle es campo de texto 50
        'ahora la cambie a memo
        'If Len(Detalle) > 49 Then Detalle = Left(Detalle, 49)
        
        DB.EXECUTE "UPDATE Clientes SET Nombre = '" + mNombre + "', " + _
            "Detalle = '" + NoNuloS(mDetalle) + "', " + _
            "Direccion = '" + NoNuloS(mDireccion) + "', " + _
            "CUIT = '" + NoNuloS(mCUIT) + "', " + _
            "IVA = '" + NoNuloS(mIVA) + "', " + _
            "Telefono = '" + NoNuloS(mTelefono) + "', " + _
            "Mail = '" + NoNuloS(mMail) + "', " + _
            "TieneTrabajo = " + CStr(NoNuloN(mTieneTrabajo)) + ", " + _
            "Nacimiento = #" + stFechaSQL(mNacimiento) + "#" + _
            " WHERE ID =" + CStr(mCodigo)
    End If
    
    Exit Function
    
ErrDuplic:
    If Err.Number = -2147467259 Then
        GoTo DeGuelta
    Else
        MsgBox CStr(Err.Number) + ": " + CStr(Err.Description)
        Resume Next
    End If
End Function


Public Function GetMovEnvDe(IdcLiente As Long) As String()
    Dim rS As New ADODB.Recordset
    rS.CursorLocation = adUseClient
    
    rS.Open "Select * from movenvases where codcliente=" + CStr(IdcLiente), _
        DB.CN, adOpenDynamic, adLockReadOnly
    If rS.RecordCount = 0 Then 'entonces no debe envases
        Dim Gt(0) As String
        Gt(0) = ""
        GetMovEnvDe = Gt
        Exit Function
    Else 'si tiene registros aca contamos envases x envases cuantos
         'hay de cada uno
      
            
        Dim I As Long, J As Long
        'variable con valores unicos y su cantidad
        Dim Resultado() As String, CantResultado() As Long
        'poner el primer valor ya a mano para poder inicializar la matriz de resultados!
        ReDim Preserve Resultado(0): ReDim Preserve CantResultado(0)
        rS.MoveFirst
        Resultado(0) = rS("Envases")
        CantResultado(0) = CLng(rS("cantenv"))
        
        Dim Encontrado As Boolean 'para saber si un valor ya se habia encontrado
            'y solo lo suma, si no agrega un nuevo valor
            'va desde uno por que el cero ya va de prepo a mano para inicializar
        
        Do While Not rS.EOF
            rS.MoveNext
            If rS.EOF Then Exit Do
            Encontrado = False 'solo aviso si lo encuentra
            For J = 0 To UBound(Resultado)
                If rS("Envases") = Resultado(J) Then
                    CantResultado(J) = CantResultado(J) + CLng(rS("cantenv"))
                    Encontrado = True
                End If
            Next J
            If Encontrado = False Then 'agregarlo!!
                Dim NewInd As Long
                NewInd = UBound(Resultado) + 1
                ReDim Preserve Resultado(NewInd)
                Resultado(NewInd) = rS("Envases")
                ReDim Preserve CantResultado(NewInd)
                CantResultado(NewInd) = rS("cantenv") 'es el primero PAPA!
            Else
                ' ya lo sumo!, no hacer nada!
            End If
        Loop
        
        'acomodar todo y devolver
        Dim TmP() As String
        For I = 0 To UBound(Resultado)
            ReDim Preserve TmP(I)
              'XXXX esto pasa solo cuando devuelve todo los envases se crea
              'un renglon en la tabla movenvases con el campo envases= " "
              'y 0 en el campo cantenvases, entonces dejo vacio temp(i)
            If CantResultado(I) = 0 Then
                TmP(I) = ""
            Else
                TmP(I) = CStr(CantResultado(I)) + "-" + Resultado(I) + " \\ "
                'ya le meti un separador al final para separarlos cuando
                'los meta todo en un label
            End If
        Next I
        
        
            
            'XXXX negrada artime si no convence hacer algo distinto
            'incluye en el ultimo miembro del array el total de $$
            'de vales
        Dim Vales As Single
        rS.MoveFirst
        Do While Not rS.EOF
            If rS("depositoporenvase") <> "" Then
                Vales = Vales + CSng(rS("depositoporenvase"))
            End If
            rS.MoveNext
        Loop
        
        rS.Close
        Set rS = Nothing
        
        Dim Nw As Long
        Nw = UBound(TmP) + 1
        Set rS = Nothing
        ReDim Preserve TmP(Nw)
        TmP(Nw) = "Total Vales: " + FormatCurrency(Vales, , , , vbFalse)
        GetMovEnvDe = TmP
    End If
End Function


