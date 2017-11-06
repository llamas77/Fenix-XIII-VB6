Attribute VB_Name = "modAdmins"
'programado por maTih.-
 
Option Explicit
 
Public GMList() As Integer
 
Sub GenerarArray()
 
' @ Un poco feo, pero para ahorrar posibles rt9.
 
ReDim GMList(1 To 1) As Integer
 
End Sub
 
Sub ConectarAdmin(ByVal gmIndex As Integer)
 
' @ Agrega un gm a la lista.
 
'Setea nuevas dimensiones.
ReDim Preserve GMList(1 To (UBound(GMList()) + 1)) As Integer
 
'Agrega a la lista.
GMList(UBound(GMList())) = gmIndex
 
End Sub
 
Sub DesConectarAdmin(ByVal gmIndex As Integer)
 
' @ Saca un gm de la lista.
 
Dim gmPos   As Integer
 
gmPos = IndexEnArray(gmIndex)
 
'Algo raro..
If Not gmPos <> -1 Then Exit Sub
 
'quita de la lista.
GMList(gmPos) = -1
 
'compacta.
CompactarAdmins
 
End Sub
 
Sub CompactarAdmins()
 
' @ Ordena la lista.
 
Dim loopX   As Long
Dim NotSlot As Integer
Dim Temp()  As Integer
 
For loopX = 1 To UBound(GMList())
    'Hya gm,
    If (GMList(loopX) <> -1) And (GMList(loopX) <> 0) Then
        NotSlot = NotSlot + 1
        ReDim Preserve Temp(1 To NotSlot) As Integer
        Temp(NotSlot) = GMList(loopX)
    End If
Next loopX
 
If Not NotSlot <> 0 Then Exit Sub
 
'setea el array.
ReDim GMList(1 To NotSlot) As Integer
 
For loopX = 1 To NotSlot
    GMList(loopX) = Temp(loopX)
Next loopX
 
End Sub
 
Sub EnviarToAdmins(ByRef dataTosend As String)
 
' @ Envia datos a los gms.
 
Dim loopX   As Long
 
For loopX = 1 To UBound(GMList())
    'Si hay un index.
    If (GMList(loopX) <> -1) And (GMList(loopX) <> 0) Then
        'Si está logeado.
        If UserList(GMList(loopX)).ConnID <> -1 Then
            Call EnviarDatosASlot(GMList(loopX), dataTosend)
        End If
    End If
Next loopX
 
End Sub
 
Function IndexEnArray(ByVal gmIndex As Integer) As Integer
 
' @ Devuelve la posición en el array de un gm.
 
Dim loopX   As Long
 
For loopX = 1 To UBound(GMList())
    If GMList(loopX) = gmIndex Then
       IndexEnArray = loopX
       Exit Function
    End If
Next loopX
 
IndexEnArray = -1
 
End Function
