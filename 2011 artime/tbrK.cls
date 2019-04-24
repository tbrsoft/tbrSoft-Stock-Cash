VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tbrK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'modulo del manejo de la licencia
Dim NF As New tbrDATA.clsTODO

Public Enum nLIC
    SinLeer = -1 'cuando estemodulo no se fijo todavia
    SinCargar = 0
    CargadaMal = 1
    Licencia1 = 2
    Licencia2 = 3
    Licencia3 = 4
End Enum

Private MyLic As nLIC

Public Function GetDatosPC(ArchDest As String)
    'obtener el archivo para madar a tbrSoft
    NF.SetSF "tbrStock"
    NF.DoNow ArchDest
End Function

Public Sub PutArchLic(Arch As String)

    'ver si no existe o no definio el archivo
    Dim FSO As New Scripting.FileSystemObject
    If (Arch = "") Or (FSO.FileExists(Arch)) = False Then
        MyLic = SinCargar
        Exit Sub
    End If

    'poner archivo de licencia
    Dim NR As Long 'numero de la licencia
    Dim DD As String 'valor por referencia que me da el codigo de la PC para que fue hecho
'    If DD = "" Then
'        NR = -1
'    Else
        NR = NF.GetNR(Arch, DD)
        If NF.GetSF = "" Or NF.GetSF = LCase("tbrStock") Then
            'nada
        Else
            MyLic = CargadaMal
        End If
        
'    End If
    If UCase(Trim(DD)) <> UCase(Trim(NF.GetRF)) Then
        MyLic = CargadaMal
    Else
        'ver que tipo de licencia me mandaron
        Select Case NR
            Case 23, 55, 12
                MyLic = Licencia1
            Case 43, 11, 99
                MyLic = Licencia2
            Case 10, 20, 30
                MyLic = Licencia3
            Case Else 'otros numeros no los mande yo
                MyLic = CargadaMal
        End Select
    End If
End Sub

Public Function GetLic() As nLIC
    GetLic = MyLic
End Function

Private Sub Class_Initialize()
    MyLic = SinLeer
End Sub
