Attribute VB_Name = "TCP"
'Argentum Online 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Sub DarCuerpo(ByVal UserIndex As Integer)
'*************************************************
'Author: Nacho (Integer)
'Last modified: 14/03/2007
'Elije una cabeza para el usuario y le da un body
'*************************************************
Dim NewBody As Integer
Dim UserRaza As Byte
Dim UserGenero As Byte

UserGenero = UserList(UserIndex).Genero
UserRaza = UserList(UserIndex).raza

Select Case UserGenero
   Case eGenero.Hombre
        Select Case UserRaza
            Case eRaza.Humano
                NewBody = 1
            Case eRaza.Elfo
                NewBody = 2
            Case eRaza.Drow
                NewBody = 3
            Case eRaza.Enano
                NewBody = 52
            Case eRaza.Gnomo
                NewBody = 52
        End Select
   Case eGenero.Mujer
        Select Case UserRaza
            Case eRaza.Humano
                NewBody = 1
            Case eRaza.Elfo
                NewBody = 2
            Case eRaza.Drow
                NewBody = 3
            Case eRaza.Gnomo
                NewBody = 52
            Case eRaza.Enano
                NewBody = 52
        End Select
End Select

UserList(UserIndex).Char.body = NewBody
End Sub

Private Function ValidarCabeza(ByVal UserRaza As Byte, ByVal UserGenero As Byte, ByVal Head As Integer) As Boolean

Select Case UserGenero
    Case eGenero.Hombre
        Select Case UserRaza
            Case eRaza.Humano
                ValidarCabeza = (Head >= HUMANO_H_PRIMER_CABEZA And _
                                Head <= HUMANO_H_ULTIMA_CABEZA)
            Case eRaza.Elfo
                ValidarCabeza = (Head >= ELFO_H_PRIMER_CABEZA And _
                                Head <= ELFO_H_ULTIMA_CABEZA)
            Case eRaza.Drow
                ValidarCabeza = (Head >= DROW_H_PRIMER_CABEZA And _
                                Head <= DROW_H_ULTIMA_CABEZA)
            Case eRaza.Enano
                ValidarCabeza = (Head >= ENANO_H_PRIMER_CABEZA And _
                                Head <= ENANO_H_ULTIMA_CABEZA)
            Case eRaza.Gnomo
                ValidarCabeza = (Head >= GNOMO_H_PRIMER_CABEZA And _
                                Head <= GNOMO_H_ULTIMA_CABEZA)
        End Select
    
    Case eGenero.Mujer
        Select Case UserRaza
            Case eRaza.Humano
                ValidarCabeza = (Head >= HUMANO_M_PRIMER_CABEZA And _
                                Head <= HUMANO_M_ULTIMA_CABEZA)
            Case eRaza.Elfo
                ValidarCabeza = (Head >= ELFO_M_PRIMER_CABEZA And _
                                Head <= ELFO_M_ULTIMA_CABEZA)
            Case eRaza.Drow
                ValidarCabeza = (Head >= DROW_M_PRIMER_CABEZA And _
                                Head <= DROW_M_ULTIMA_CABEZA)
            Case eRaza.Enano
                ValidarCabeza = (Head >= ENANO_M_PRIMER_CABEZA And _
                                Head <= ENANO_M_ULTIMA_CABEZA)
            Case eRaza.Gnomo
                ValidarCabeza = (Head >= GNOMO_M_PRIMER_CABEZA And _
                                Head <= GNOMO_M_ULTIMA_CABEZA)
        End Select
End Select
        
End Function

Function AsciiValidos(ByVal cad As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(mid$(cad, i, 1))
    
    If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
        AsciiValidos = False
        Exit Function
    End If
    
Next i

AsciiValidos = True

End Function

Function Numeric(ByVal cad As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(mid$(cad, i, 1))
    
    If (car < 48 Or car > 57) Then
        Numeric = False
        Exit Function
    End If
    
Next i

Numeric = True

End Function


Function NombrePermitido(ByVal Nombre As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim i As Integer

For i = 1 To UBound(ForbidenNames)
    If InStr(Nombre, ForbidenNames(i)) Then
            NombrePermitido = False
            Exit Function
    End If
Next i

NombrePermitido = True

End Function

Function ValidateSkills(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim LoopC As Integer

For LoopC = 1 To NUMSKILLS
    If UserList(UserIndex).Stats.UserSkills(LoopC) < 0 Then
        Exit Function
        If UserList(UserIndex).Stats.UserSkills(LoopC) > 100 Then UserList(UserIndex).Stats.UserSkills(LoopC) = 100
    End If
Next LoopC

ValidateSkills = True
    
End Function

Sub ConnectNewUser(ByVal UserIndex As Integer, ByRef Name As String, ByRef Password As String, ByVal UserRaza As eRaza, ByVal UserSexo As eGenero, _
                    ByRef UserEmail As String, ByVal Hogar As eCiudad, ByVal Head As Integer, Skills() As Byte)
'*************************************************
'Author: Unknown
'Last modified: 3/12/2009
'Conecta un nuevo Usuario
'23/01/2007 Pablo (ToxicWaste) - Agregué ResetFaccion al crear usuario
'24/01/2007 Pablo (ToxicWaste) - Agregué el nuevo mana inicial de los magos.
'12/02/2007 Pablo (ToxicWaste) - Puse + 1 de const al Elfo normal.
'20/04/2007 Pablo (ToxicWaste) - Puse -1 de fuerza al Elfo.
'09/01/2008 Pablo (ToxicWaste) - Ahora los modificadores de Raza se controlan desde Balance.dat
'11/19/2009: Pato - Modifico la maná inicial del bandido.
'03/12/2009: Budi - Optimización del código.
'*************************************************
Dim i As Long

With UserList(UserIndex)

    If Not AsciiValidos(Name) Or LenB(Name) = 0 Then
        Call WriteErrorMsg(UserIndex, "Nombre inválido.")
        Exit Sub
    End If
    
    If UserList(UserIndex).flags.UserLogged Then
        Call LogCheating("El usuario " & UserList(UserIndex).Name & " ha intentado crear a " & Name & " desde la IP " & UserList(UserIndex).ip)
        
        'Kick player ( and leave character inside :D )!
        Call CloseSocketSL(UserIndex)
        Call Cerrar_Usuario(UserIndex)
        
        Exit Sub
    End If
    
    '¿Existe el personaje?
    If FileExist(CharPath & UCase$(Name) & ".chr", vbNormal) = True Then
        Call WriteErrorMsg(UserIndex, "Ya existe el personaje.")
        Exit Sub
    End If
    
    'Tiró los dados antes de llegar acá??
    If .Stats.UserAtributos(eAtributos.Fuerza) = 0 Then
        Call WriteErrorMsg(UserIndex, "Debe tirar los dados antes de poder crear un personaje.")
        Exit Sub
    End If
    
    If Not ValidarCabeza(UserRaza, UserSexo, Head) Then
        Call LogCheating("El usuario " & Name & " ha seleccionado la cabeza " & Head & " desde la IP " & .ip)
        
        Call WriteErrorMsg(UserIndex, "Cabeza inválida, elija una cabeza seleccionable.")
        Exit Sub
    End If
    
    .flags.Muerto = 0
    .flags.Escondido = 0
    
    
    .Name = Trim$(Name)
    .Clase = eClass.Ciudadano
    .raza = UserRaza
    .Genero = UserSexo
    .email = UserEmail
    .Hogar = Hogar
    
    '[Pablo (Toxic Waste) 9/01/08]
    .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) + ModRaza(UserRaza).Fuerza
    .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) + ModRaza(UserRaza).Agilidad
    .Stats.UserAtributos(eAtributos.Inteligencia) = .Stats.UserAtributos(eAtributos.Inteligencia) + ModRaza(UserRaza).Inteligencia
    .Stats.UserAtributos(eAtributos.Carisma) = .Stats.UserAtributos(eAtributos.Carisma) + ModRaza(UserRaza).Carisma
    .Stats.UserAtributos(eAtributos.Constitucion) = .Stats.UserAtributos(eAtributos.Constitucion) + ModRaza(UserRaza).Constitucion
    '[/Pablo (Toxic Waste)]
    
    .Stats.SkillPts = 10
    For i = 1 To NUMSKILLS
        If Skills(i) > 0 Then
            .Stats.SkillPts = .Stats.SkillPts - Skills(i)
        End If
    Next i
    
    If .Stats.SkillPts <> 0 Then
        Call WriteErrorMsg(UserIndex, "Skills inválidos.")
        Exit Sub
    End If
    '.Stats.SkillPts = 0
    
    CopyMemory .Stats.UserSkills(1), Skills(1), NUMSKILLS
    '* lenb(skills(1)) 'podría servir si a alguien se le ocurre usar skills integers
    
    .Char.heading = eHeading.SOUTH
    
    Call DarCuerpo(UserIndex)
    .Char.Head = Head
    
    .OrigChar = .Char
    
    Dim MiInt As Long
    MiInt = RandomNumber(1, .Stats.UserAtributos(eAtributos.Constitucion) \ 3)
    
    .Stats.MaxHp = 15 + MiInt
    .Stats.MinHp = 15 + MiInt
    
    MiInt = RandomNumber(1, .Stats.UserAtributos(eAtributos.Agilidad) \ 6)
    If MiInt = 1 Then MiInt = 2
    
    .Stats.MaxSta = 20 * MiInt
    .Stats.MinSta = 20 * MiInt
    
    
    .Stats.MaxAGU = 100
    .Stats.MinAGU = 100
    
    .Stats.MaxHam = 100
    .Stats.MinHam = 100
    
    .Stats.MaxHIT = 2
    .Stats.MinHIT = 1
    
    .Stats.GLD = 0
    
    .Stats.Exp = 0
    .Stats.ELU = 300
    .Stats.ELV = 1
    
    '???????????????? INVENTARIO ¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿
    Dim Slot As Byte
    .Invent.NroItems = 4
    
    .Invent.Object(1).OBJIndex = ManzanaNewbie
    .Invent.Object(1).Amount = 100
    
    .Invent.Object(2).OBJIndex = 468
    .Invent.Object(2).Amount = 100
    
    .Invent.Object(3).OBJIndex = 460
    .Invent.Object(3).Amount = 1
    .Invent.Object(3).Equipped = 1

    ' Ropa (Newbie)
    Slot = Slot + 1
    Select Case UserRaza
        Case eRaza.Humano
            .Invent.Object(Slot).OBJIndex = 463
        Case eRaza.Elfo
            .Invent.Object(Slot).OBJIndex = 464
        Case eRaza.Drow
            .Invent.Object(Slot).OBJIndex = 465
        Case eRaza.Enano
            .Invent.Object(Slot).OBJIndex = 466
        Case eRaza.Gnomo
            .Invent.Object(Slot).OBJIndex = 466
    End Select
    
    ' Equipo ropa
    .Invent.Object(Slot).Amount = 1
    .Invent.Object(Slot).Equipped = 1
    
    .Invent.ArmourEqpSlot = Slot
    .Invent.ArmourEqpObjIndex = .Invent.Object(Slot).OBJIndex

    'Arma (Newbie)
    Slot = Slot + 1
    .Invent.Object(Slot).OBJIndex = 460
    
    ' Equipo arma
    .Invent.Object(Slot).Amount = 1
    .Invent.Object(Slot).Equipped = 1
    
    .Invent.WeaponEqpObjIndex = .Invent.Object(Slot).OBJIndex
    .Invent.WeaponEqpSlot = Slot
    
    .Char.WeaponAnim = ObjData(.Invent.WeaponEqpObjIndex).WeaponAnim

    ' Manzanas (Newbie)
    Slot = Slot + 1
    .Invent.Object(Slot).OBJIndex = 467
    .Invent.Object(Slot).Amount = 100
    
    ' Jugos (Nwbie)
    Slot = Slot + 1
    .Invent.Object(Slot).OBJIndex = 468
    .Invent.Object(Slot).Amount = 100
    
    ' Sin casco y escudo
    .Char.ShieldAnim = NingunEscudo
    .Char.CascoAnim = NingunCasco
    
    ' Total Items
    .Invent.NroItems = Slot
    
    .GuildID = 0
    .flags.WaitingApprovement = 0
    .flags.IsLeader = 0
    
    #If ConUpTime Then
        .LogOnTime = Now
        .UpTime = 0
    #End If

End With

'Valores Default de facciones al Activar nuevo usuario
Call ResetFacciones(UserIndex)

Password = MD5String(Password)

Call WriteVar(CharPath & UCase$(Name) & ".chr", "INIT", "Password", Password) 'grabamos el password aqui afuera, para no mantenerlo cargado en memoria

Call SaveUser(UserIndex, CharPath & UCase$(Name) & ".chr")
  
'Open User
Call ConnectUser(UserIndex, Name, Password)
  
End Sub

Sub CloseSocket(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo Errhandler
    
    If UserIndex = LastUser Then
        Do Until UserList(LastUser).flags.UserLogged
            LastUser = LastUser - 1
            If LastUser < 1 Then Exit Do
        Loop
    End If
    
    'Call SecurityIp.IpRestarConexion(GetLongIp(UserList(UserIndex).ip))
    
    If (EsGM(UserIndex) Or EsAdmin(UserIndex)) Then Call DesConectarAdmin(UserIndex)
    
    If UserList(UserIndex).Faccion.Bando <> eFaccion.Neutral Then QuitarMiembroFaccion UserIndex
    
    If UserList(UserIndex).ConnID <> -1 Then
        Call CloseSocketSL(UserIndex)
    End If
    
    'Es el mismo user al que está revisando el centinela??
    'IMPORTANTE!!! hacerlo antes de resetear así todavía sabemos el nombre del user
    ' y lo podemos loguear
    If Centinela.RevisandoUserIndex = UserIndex Then _
        Call modCentinela.CentinelaUserLogout
    
    'mato los comercios seguros
    If UserList(UserIndex).ComUsu.DestUsu > 0 Then
        If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.UserLogged Then
            If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                Call WriteConsoleMsg(UserList(UserIndex).ComUsu.DestUsu, "Comercio cancelado por el otro usuario", FontTypeNames.FONTTYPE_TALK)
                Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)
                Call FlushBuffer(UserList(UserIndex).ComUsu.DestUsu)
            End If
        End If
    End If
    
    Call DisconnectMember(UserList(UserIndex).GuildID, UserIndex)
    
    'Empty buffer for reuse
    Call UserList(UserIndex).incomingData.ReadASCIIStringFixed(UserList(UserIndex).incomingData.Length)
    
    If UserList(UserIndex).flags.UserLogged Then
        If NumUsers > 0 Then NumUsers = NumUsers - 1
        Call CloseUser(UserIndex)
        
    Else
        Call ResetUserSlot(UserIndex)
    End If
    
    UserList(UserIndex).ConnID = -1
    UserList(UserIndex).ConnIDValida = False
    
Exit Sub

Errhandler:
    UserList(UserIndex).ConnID = -1
    UserList(UserIndex).ConnIDValida = False
    Call ResetUserSlot(UserIndex)

    Call LogError("CloseSocket - Error = " & Err.Number & " - Descripción = " & Err.description & " - UserIndex = " & UserIndex)
End Sub

'[Alejo-21-5]: Cierra un socket sin limpiar el slot
Sub CloseSocketSL(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

If UserList(UserIndex).ConnID <> -1 And UserList(UserIndex).ConnIDValida Then
    Call BorraSlotSock(UserList(UserIndex).ConnID)
    Call WSApiCloseSocket(UserList(UserIndex).ConnID)
    UserList(UserIndex).ConnIDValida = False
End If

End Sub

''
' Send an string to a Slot
'
' @param userIndex The index of the User
' @param Datos The string that will be send
'CSEH: ErrLog
Public Function EnviarDatosASlot(ByVal UserIndex As Integer, ByRef Datos As String) As Long
    '***************************************************
    'Author: Unknown
    'Last Modification: 01/10/07
    'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
    'Now it uses the clsByteQueue class and don`t make a FIFO Queue of String
    '***************************************************
    '<EhHeader>
    On Error GoTo EnviarDatosASlot_Err
    '</EhHeader>
        Dim ret As Long
    
100     ret = WsApiEnviar(UserIndex, Datos)
    
105     If ret <> 0 And ret <> WSAEWOULDBLOCK Then
            ' Close the socket avoiding any critical error
110         Call CloseSocketSL(UserIndex)
115         Call Cerrar_Usuario(UserIndex)
        End If

    '<EhFooter>
    Exit Function

EnviarDatosASlot_Err:
        Call LogError("Error en EnviarDatosASlot: " & Erl & " - " & Err.description)
    '</EhFooter>
End Function
Function EstaPCarea(index As Integer, Index2 As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim X As Integer, Y As Integer
For Y = UserList(index).Pos.Y - MinYBorder + 1 To UserList(index).Pos.Y + MinYBorder - 1
        For X = UserList(index).Pos.X - MinXBorder + 1 To UserList(index).Pos.X + MinXBorder - 1

            If MapData(UserList(index).Pos.map, X, Y).UserIndex = Index2 Then
                EstaPCarea = True
                Exit Function
            End If
        
        Next X
Next Y
EstaPCarea = False
End Function

Function HayPCarea(Pos As WorldPos) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim X As Integer, Y As Integer
For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
        For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1
            If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                If MapData(Pos.map, X, Y).UserIndex > 0 Then
                    HayPCarea = True
                    Exit Function
                End If
            End If
        Next X
Next Y
HayPCarea = False
End Function

Function HayOBJarea(Pos As WorldPos, OBJIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim X As Integer, Y As Integer
For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
        For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1
            If MapData(Pos.map, X, Y).ObjInfo.OBJIndex = OBJIndex Then
                HayOBJarea = True
                Exit Function
            End If
        
        Next X
Next Y
HayOBJarea = False
End Function
Function ValidateChr(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

ValidateChr = UserList(UserIndex).Char.Head <> 0 _
                And UserList(UserIndex).Char.body <> 0 _
                And ValidateSkills(UserIndex)

End Function

Sub ConnectUser(ByVal UserIndex As Integer, ByRef Name As String, ByRef Password As String)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 3/12/2009 (Budi)
'26/03/2009: ZaMa - Agrego por default que el color de dialogo de los dioses, sea como el de su nick.
'12/06/2009: ZaMa - Agrego chequeo de nivel al loguear
'14/09/2009: ZaMa - Ahora el usuario esta protegido del ataque de npcs al loguear
'11/27/2009: Budi - Se envian los InvStats del personaje y su Fuerza y Agilidad
'03/12/2009: Budi - Optimización del código
'***************************************************
Dim N As Integer
Dim i As Long

With UserList(UserIndex)

    If .flags.UserLogged Then
        Call LogCheating("El usuario " & .Name & " ha intentado loguear a " & Name & " desde la IP " & .ip)
        'Kick player ( and leave character inside :D )!
        Call CloseSocketSL(UserIndex)
        Call Cerrar_Usuario(UserIndex)
        Exit Sub
    End If
    
    'Reseteamos los FLAGS
    .flags.Escondido = 0
    .flags.TargetNPC = 0
    .flags.TargetNpcTipo = eNPCType.Comun
    .flags.TargetObj = 0
    .flags.TargetUser = 0
    .Char.FX = 0
    
    'Controlamos no pasar el maximo de usuarios
    If NumUsers >= MaxUsers Then
        Call WriteErrorMsg(UserIndex, "El servidor ha alcanzado el máximo de usuarios soportado, por favor vuelva a intertarlo más tarde.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
    
    '¿Este IP ya esta conectado?
    If AllowMultiLogins = 0 Then
        If CheckForSameIP(UserIndex, .ip) = True Then
            Call WriteErrorMsg(UserIndex, "No es posible usar más de un personaje al mismo tiempo.")
            Call FlushBuffer(UserIndex)
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
    End If
    
    '¿Existe el personaje?
    If Not FileExist(CharPath & UCase$(Name) & ".chr", vbNormal) Then
        Call WriteErrorMsg(UserIndex, "El personaje no existe.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
    
    '¿Es el passwd valido?
    If UCase$(Password) <> UCase$(GetVar(CharPath & UCase$(Name) & ".chr", "INIT", "Password")) Then
        Call WriteErrorMsg(UserIndex, "Password incorrecto.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
    
    '¿Ya esta conectado el personaje?
    If CheckForSameName(Name) Then
        If UserList(NameIndex(Name)).Counters.Saliendo Then
            Call WriteErrorMsg(UserIndex, "El usuario está saliendo.")
        Else
            Call WriteErrorMsg(UserIndex, "Perdón, un usuario con el mismo nombre se ha logueado.")
        End If
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
    
    'Reseteamos los privilegios
    .flags.Privilegios = 0
    
    'Vemos que clase de user es (se lo usa para setear los privilegios al loguear el PJ)
    If EsAdmin(Name) Then
        .flags.Privilegios = .flags.Privilegios Or PlayerType.Admin
        Call LogGM(Name, "Se conecto con ip:" & .ip)
        Call ConectarAdmin(UserIndex)
    ElseIf EsDios(Name) Then
        .flags.Privilegios = .flags.Privilegios Or PlayerType.Dios
        Call LogGM(Name, "Se conecto con ip:" & .ip)
        Call ConectarAdmin(UserIndex)
    ElseIf EsSemiDios(Name) Then
        .flags.Privilegios = .flags.Privilegios Or PlayerType.SemiDios
        Call LogGM(Name, "Se conecto con ip:" & .ip)
    ElseIf EsConsejero(Name) Then
        .flags.Privilegios = .flags.Privilegios Or PlayerType.Consejero
        Call LogGM(Name, "Se conecto con ip:" & .ip)
    Else
        .flags.Privilegios = .flags.Privilegios Or PlayerType.User
        .flags.AdminPerseguible = True
    End If
    
    'Add RM flag if needed
    If EsRolesMaster(Name) Then
        .flags.Privilegios = .flags.Privilegios Or PlayerType.RoleMaster
    End If
    
    If ServerSoloGMs > 0 Then
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)) = 0 Then
            Call WriteErrorMsg(UserIndex, "Servidor restringido a administradores. Por favor reintente en unos momentos.")
            Call FlushBuffer(UserIndex)
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
    End If
    
    'Cargamos el personaje
    Dim Leer As New clsIniReader
    
    Call Leer.Initialize(CharPath & UCase$(Name) & ".chr")
    
    'Cargamos los datos del personaje
    Call LoadUserInit(UserIndex, Leer)
    
    Call LoadUserStats(UserIndex, Leer)
    
    If Not ValidateChr(UserIndex) Then
        Call WriteErrorMsg(UserIndex, "Error en el personaje.")
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
    
    Set Leer = Nothing
    
    If EnTesting And .Stats.ELV >= 18 Then
        Call WriteErrorMsg(UserIndex, "Servidor en Testing por unos minutos, conectese con PJs de nivel menor a 18. No se conecte con Pjs que puedan resultar importantes por ahora pues pueden arruinarse.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
    
    If .Invent.EscudoEqpSlot = 0 Then .Char.ShieldAnim = NingunEscudo
    If .Invent.CascoEqpSlot = 0 Then .Char.CascoAnim = NingunCasco
    If .Invent.WeaponEqpSlot = 0 Then .Char.WeaponAnim = NingunArma
    
    If .Invent.MochilaEqpSlot > 0 Then
        .CurrentInventorySlots = MAX_NORMAL_INVENTORY_SLOTS + ObjData(.Invent.Object(.Invent.MochilaEqpSlot).OBJIndex).MochilaType * 5
    Else
        .CurrentInventorySlots = MAX_NORMAL_INVENTORY_SLOTS
    End If
    
    Call UpdateUserInv(True, UserIndex, 0)
    Call UpdateUserHechizos(True, UserIndex, 0)
    
    If .flags.Paralizado Then
        Call WriteParalizeOK(UserIndex)
    End If
    
    ''
    'TODO : Feo, esto tiene que ser parche cliente
    If .flags.Estupidez = 0 Then
        Call WriteDumbNoMore(UserIndex)
    End If
    
    .flags.IsLeader = 0
    
    If .GuildID <> 0 Then
        
        If StrComp(UCase$(Guilds(.GuildID).LeaderName), UCase$(Name)) = 0 Then
            .flags.IsLeader = 1
        Else
            For i = 0 To 2
                If StrComp(UCase$(Guilds(.GuildID).RecruiterName(i)), UCase$(Name)) = 0 Then
                    .flags.IsLeader = 2
                    Exit For
                End If
            Next
            
        End If
        
    End If
    
    'Posicion de comienzo
    If .Pos.map = 0 Then
        Select Case .Hogar
            Case eCiudad.cNix
                .Pos = Nix
            Case eCiudad.cUllathorpe
                .Pos = Ullathorpe
            Case eCiudad.cBanderbill
                .Pos = Banderbill
            Case eCiudad.cLindos
                .Pos = Lindos
            Case eCiudad.cArghal
                .Pos = Arghal
            Case Else
                .Hogar = eCiudad.cUllathorpe
                .Pos = Ullathorpe
        End Select
    Else
        If Not MapaValido(.Pos.map) Then
            Call WriteErrorMsg(UserIndex, "El PJ se encuenta en un mapa inválido.")
            Call FlushBuffer(UserIndex)
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
    End If
    
    'Si es gm lo seteo en el mapa de gms
    If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)) Then
        .Pos.map = 86
        .Pos.X = 50
        .Pos.Y = 50
    End If
    
    'Tratamos de evitar en lo posible el "Telefrag". Solo 1 intento de loguear en pos adjacentes.
    'Codigo por Pablo (ToxicWaste) y revisado por Nacho (Integer), corregido para que realmetne ande y no tire el server por Juan Martín Sotuyo Dodero (Maraxus)
    If MapData(.Pos.map, .Pos.X, .Pos.Y).UserIndex <> 0 Or MapData(.Pos.map, .Pos.X, .Pos.Y).NpcIndex <> 0 Then
        Dim FoundPlace As Boolean
        Dim esAgua As Boolean
        Dim tX As Long
        Dim tY As Long
        
        FoundPlace = False
        esAgua = (MapData(.Pos.map, .Pos.X, .Pos.Y).Agua = 1)
        
        For tY = .Pos.Y - 1 To .Pos.Y + 1
            For tX = .Pos.X - 1 To .Pos.X + 1
                If esAgua Then
                    'reviso que sea pos legal en agua, que no haya User ni NPC para poder loguear.
                    If LegalPos(.Pos.map, tX, tY, True, False) Then
                        FoundPlace = True
                        Exit For
                    End If
                Else
                    'reviso que sea pos legal en tierra, que no haya User ni NPC para poder loguear.
                    If LegalPos(.Pos.map, tX, tY, False, True) Then
                        FoundPlace = True
                        Exit For
                    End If
                End If
            Next tX
            
            If FoundPlace Then _
                Exit For
        Next tY
        
        If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
            .Pos.X = tX
            .Pos.Y = tY
        Else
            'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un NPC, lo pisamos.
            If MapData(.Pos.map, .Pos.X, .Pos.Y).UserIndex <> 0 Then
               'Si no encontramos lugar, y abajo teniamos a un usuario, lo pisamos y cerramos su comercio seguro
                If UserList(MapData(.Pos.map, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu > 0 Then
                    'Le avisamos al que estaba comerciando que se tuvo que ir.
                    If UserList(UserList(MapData(.Pos.map, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu).flags.UserLogged Then
                        Call FinComerciarUsu(UserList(MapData(.Pos.map, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu)
                        Call WriteConsoleMsg(UserList(MapData(.Pos.map, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu, "Comercio cancelado. El otro usuario se ha desconectado.", FontTypeNames.FONTTYPE_TALK)
                        Call FlushBuffer(UserList(MapData(.Pos.map, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu)
                    End If
                    'Lo sacamos.
                    If UserList(MapData(.Pos.map, .Pos.X, .Pos.Y).UserIndex).flags.UserLogged Then
                        Call FinComerciarUsu(MapData(.Pos.map, .Pos.X, .Pos.Y).UserIndex)
                        Call WriteErrorMsg(MapData(.Pos.map, .Pos.X, .Pos.Y).UserIndex, "Alguien se ha conectado donde te encontrabas, por favor reconéctate...")
                        Call FlushBuffer(MapData(.Pos.map, .Pos.X, .Pos.Y).UserIndex)
                    End If
                End If
                
                Call CloseSocket(MapData(.Pos.map, .Pos.X, .Pos.Y).UserIndex)
            End If
        End If
    End If
    
    'Nombre de sistema
    .Name = Name
    
    .showName = True 'Por default los nombres son visibles
    
    'If in the water, and has a boat, equip it!
    If .Invent.BarcoObjIndex > 0 And _
            ((MapData(.Pos.map, .Pos.X, .Pos.Y).Agua = 1) Or BodyIsBoat(.Char.body)) Then
        Dim Barco As ObjData
        Barco = ObjData(.Invent.BarcoObjIndex)
        .Char.Head = 0
        If .flags.Muerto = 0 Then
            Call ToogleBoatBody(UserIndex)
        Else
            .Char.body = iFragataFantasmal
            .Char.ShieldAnim = NingunEscudo
            .Char.WeaponAnim = NingunArma
            .Char.CascoAnim = NingunCasco
        End If
        
        .flags.Navegando = 1
    End If
    
    
    'Info
    Call WriteUserIndexInServer(UserIndex) 'Enviamos el User index
    
    
    Call WriteChangeMap(UserIndex, .Pos.map, MapInfo(.Pos.map).MapVersion) 'Carga el mapa
    Call WritePlayMidi(UserIndex, val(ReadField(1, MapInfo(.Pos.map).Music, 45)))
    
    If .flags.Privilegios <> PlayerType.User And .flags.Privilegios <> (PlayerType.User Or PlayerType.ChaosCouncil) And .flags.Privilegios <> (PlayerType.User Or PlayerType.RoyalCouncil) Then
        .flags.ChatColor = RGB(255, 128, 32)
    ElseIf .flags.Privilegios = (PlayerType.User Or PlayerType.RoyalCouncil) Then
        .flags.ChatColor = RGB(0, 255, 255)
    ElseIf .flags.Privilegios = (PlayerType.User Or PlayerType.ChaosCouncil) Then
        .flags.ChatColor = RGB(255, 128, 64)
    Else
        .flags.ChatColor = vbWhite
    End If
    
    
    ''[EL OSO]: TRAIGO ESTO ACA ARRIBA PARA DARLE EL IP!
    #If ConUpTime Then
        .LogOnTime = Now
    #End If
    
    'Crea  el personaje del usuario
    Call MakeUserChar(True, .Pos.map, UserIndex, .Pos.map, .Pos.X, .Pos.Y)
    
    Call WriteUserCharIndexInServer(UserIndex)
    ''[/el oso]
    
    Call DoTileEvents(UserIndex, .Pos.map, .Pos.X, .Pos.Y)
    
    Call CheckUserLevel(UserIndex)
    Call WriteUpdateUserStats(UserIndex)
    
    Call WriteUpdateHungerAndThirst(UserIndex)
    Call WriteUpdateStrenghtAndDexterity(UserIndex)
        
   ' Call SendMOTD(UserIndex)
    
    If haciendoBK Then
        Call WritePauseToggle(UserIndex)
        Call WriteConsoleMsg(UserIndex, "Servidor> Por favor espera algunos segundos, el WorldSave está ejecutándose.", FontTypeNames.FONTTYPE_SERVER)
    End If
    
    If EnPausa Then
        Call WritePauseToggle(UserIndex)
        Call WriteConsoleMsg(UserIndex, "Servidor> Lo sentimos mucho pero el servidor se encuentra actualmente detenido. Intenta ingresar más tarde.", FontTypeNames.FONTTYPE_SERVER)
    End If
    
    If UserList(UserIndex).Faccion.Bando <> eFaccion.Neutral Then AgregarMiembroFaccion UserIndex
    
    'Actualiza el Num de usuarios
    'DE ACA EN ADELANTE GRABA EL CHARFILE, OJO!
    NumUsers = NumUsers + 1
    .flags.UserLogged = True
    
    'usado para borrar Pjs
    Call WriteVar(CharPath & .Name & ".chr", "INIT", "Logged", "1")
        
    MapInfo(.Pos.map).NumUsers = MapInfo(.Pos.map).NumUsers + 1
    
    If .Stats.SkillPts > 0 Then
        Call WriteSendSkills(UserIndex)
        Call WriteLevelUp(UserIndex, .Stats.SkillPts)
    End If
    
    If NumUsers > recordusuarios Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Record de usuarios conectados simultaneamente." & "Hay " & NumUsers & " usuarios.", FontTypeNames.FONTTYPE_INFO))
        recordusuarios = NumUsers
        Call WriteVar(IniPath & "Server.ini", "INIT", "Record", str$(recordusuarios))
    End If
    
    If .NroMascotas > 0 And MapInfo(.Pos.map).Pk Then
        For i = 1 To MAXMASCOTAS
            If .MascotasType(i) > 0 Then
                .MascotasIndex(i) = SpawnNpc(.MascotasType(i), .Pos, True, True)
                
                If .MascotasIndex(i) > 0 Then
                    Npclist(.MascotasIndex(i)).MaestroUser = UserIndex
                    Call FollowAmo(.MascotasIndex(i))
                Else
                    .MascotasIndex(i) = 0
                End If
            End If
        Next i
    End If
    
    If .flags.Navegando = 1 Then
        Call WriteNavigateToggle(UserIndex)
    End If
    
    Call ConnectMember(.GuildID, UserIndex)
    
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FXIDs.FXWARP, 0))
    
    Call WriteLoggedMessage(UserIndex)
    
    If PuedeFaccion(UserIndex) Then Call WriteEligeFaccion(UserIndex, True)
    If PuedeSubirClase(UserIndex) Then Call WriteSubeClase(UserIndex, True)
    If PuedeRecompensa(UserIndex) Then Call WriteEligeRecompensa(UserIndex, True)
    
   ' Call modGuilds.SendGuildNews(UserIndex)
    
    ' Esta protegido del ataque de npcs por 5 segundos, si no realiza ninguna accion
    Call IntervaloPermiteSerAtacado(UserIndex, True)
    
    If Lloviendo Then
        Call WriteRainToggle(UserIndex)
    End If
    
  '  tStr = modGuilds.a_ObtenerRechazoDeChar(.name)
    
   ' If LenB(tStr) <> 0 Then
   '     Call WriteShowMessageBox(UserIndex, "Tu solicitud de ingreso al clan ha sido rechazada. El clan te explica que: " & tStr)
   ' End If

    Call MostrarNumUsers
    
    N = FreeFile
    Open App.path & "\logs\numusers.log" For Output As N
    Print #N, NumUsers
    Close #N
    
    N = FreeFile
    'Log
    Open App.path & "\logs\Connect.log" For Append Shared As #N
    Print #N, .Name & " ha entrado al juego. UserIndex:" & UserIndex & " " & time & " " & Date
    Close #N

End With
End Sub

'Sub SendMOTD(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

'    Dim j As Long
'
'    Call WriteGuildChat(UserIndex, "Mensajes de entrada:")
'    For j = 1 To MaxLines
'        Call WriteGuildChat(UserIndex, MOTD(j).texto)
'    Next j
'End Sub

Sub ResetFacciones(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 23/01/2007
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
'*************************************************
    With UserList(UserIndex).Faccion
        .Ataco(1) = 0
        .Ataco(2) = 0
        .Jerarquia = 0
        .Matados(0) = 0
        .Matados(1) = 0
        .Matados(2) = 0
        .Torneos = 0
    End With
End Sub

Sub ResetContadores(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'05/20/2007 Integer - Agregue todas las variables que faltaban.
'*************************************************
    With UserList(UserIndex).Counters
        .AGUACounter = 0
        .AttackCounter = 0
        .Ceguera = 0
        .COMCounter = 0
        .Estupidez = 0
        .Frio = 0
        .HPCounter = 0
        .IdleCount = 0
        .Invisibilidad = 0
        .Paralisis = 0
        .Pena = 0
        .PiqueteC = 0
        .STACounter = 0
        .Veneno = 0
        .Trabajando = 0
        .Ocultando = 0
        .bPuedeMeditar = False
        .Mimetismo = 0
        .Saliendo = False
        .Salir = 0
        .TiempoOculto = 0
        .TimerMagiaGolpe = 0
        .TimerGolpeMagia = 0
        .TimerLanzarSpell = 0
        .TimerPuedeAtacar = 0
        .TimerPuedeUsarArco = 0
        .TimerPuedeTrabajar = 0
        .TimerUsar = 0
        .goHome = 0
        .AsignedSkills = 0
    End With
End Sub

Sub ResetCharInfo(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(UserIndex).Char
        .body = 0
        .CascoAnim = 0
        .CharIndex = 0
        .FX = 0
        .Head = 0
        .loops = 0
        .heading = 0
        .loops = 0
        .ShieldAnim = 0
        .WeaponAnim = 0
    End With
End Sub

Sub ResetBasicUserInfo(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(UserIndex)
        .Name = vbNullString
        .desc = vbNullString
        .DescRM = vbNullString
        .Pos.map = 0
        .Pos.X = 0
        .Pos.Y = 0
        .ip = vbNullString
        .Clase = 0
        .email = vbNullString
        .Genero = 0
        .Hogar = 0
        .raza = 0
    
        
        With .Stats
            .Banco = 0
            .ELV = 0
            .ELU = 0
            .Exp = 0
            .def = 0
            '.CriminalesMatados = 0
            .NPCsMuertos = 0
            .UsuariosMatados = 0
            .SkillPts = 0
            .GLD = 0
            .UserAtributos(1) = 0
            .UserAtributos(2) = 0
            .UserAtributos(3) = 0
            .UserAtributos(4) = 0
            .UserAtributos(5) = 0
            .UserAtributosBackUP(1) = 0
            .UserAtributosBackUP(2) = 0
            .UserAtributosBackUP(3) = 0
            .UserAtributosBackUP(4) = 0
            .UserAtributosBackUP(5) = 0
        End With
        
    End With
End Sub


Sub ResetUserFlags(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 06/28/2008
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'03/29/2006 Maraxus - Reseteo el CentinelaOK también.
'06/28/2008 NicoNZ - Agrego el flag Inmovilizado
'*************************************************
    With UserList(UserIndex).flags
        .Comerciando = False
        .Ban = 0
        .Escondido = 0
        .DuracionEfecto = 0
        .NpcInv = 0
        .StatsChanged = 0
        .TargetNPC = 0
        .TargetNpcTipo = eNPCType.Comun
        .TargetObj = 0
        .TargetObjMap = 0
        .TargetObjX = 0
        .TargetObjY = 0
        .TargetUser = 0
        .TipoPocion = 0
        .TomoPocion = False
        .Descuento = vbNullString
        .Hambre = 0
        .Sed = 0
        .Descansar = False
        .Vuela = 0
        .Navegando = 0
        .Oculto = 0
        .Envenenado = 0
        .invisible = 0
        .Paralizado = 0
        .Inmovilizado = 0
        .Maldicion = 0
        .Bendicion = 0
        .Meditando = 0
        .Privilegios = 0
        .PuedeMoverse = 0
        .OldBody = 0
        .OldHead = 0
        .AdminInvisible = 0
        .ValCoDe = 0
        .Hechizo = 0
        .TimesWalk = 0
        .StartWalk = 0
        .CountSH = 0
        .Silenciado = 0
        .CentinelaOK = False
        .AdminPerseguible = False
        .lastMap = 0
        .Traveling = 0
        .AtacadoPorNpc = 0
        .AtacadoPorUser = 0
        .NoPuedeSerAtacado = False
        .EnConsulta = False
        .Ignorado = False
    End With
End Sub

Sub ResetUserSpells(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim LoopC As Long
    For LoopC = 1 To MAXUSERHECHIZOS
        UserList(UserIndex).Stats.UserHechizos(LoopC) = 0
    Next LoopC
End Sub

Sub ResetUserPets(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim LoopC As Long
    
    UserList(UserIndex).NroMascotas = 0
        
    For LoopC = 1 To MAXMASCOTAS
        UserList(UserIndex).MascotasIndex(LoopC) = 0
        UserList(UserIndex).MascotasType(LoopC) = 0
    Next LoopC
End Sub

Sub ResetUserBanco(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim LoopC As Long
    
    For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
          UserList(UserIndex).BancoInvent.Object(LoopC).Amount = 0
          UserList(UserIndex).BancoInvent.Object(LoopC).Equipped = 0
          UserList(UserIndex).BancoInvent.Object(LoopC).OBJIndex = 0
    Next LoopC
    
    UserList(UserIndex).BancoInvent.NroItems = 0
End Sub

Public Sub LimpiarComercioSeguro(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With UserList(UserIndex).ComUsu
        If .DestUsu > 0 Then
            Call FinComerciarUsu(.DestUsu)
            Call FinComerciarUsu(UserIndex)
        End If
    End With
End Sub

Sub ResetUserSlot(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim i As Long

UserList(UserIndex).ConnIDValida = False
UserList(UserIndex).ConnID = -1

Call LimpiarComercioSeguro(UserIndex)
Call ResetFacciones(UserIndex)
Call ResetContadores(UserIndex)
Call ResetCharInfo(UserIndex)
Call ResetBasicUserInfo(UserIndex)
Call ResetUserFlags(UserIndex)
Call LimpiarInventario(UserIndex)
Call ResetUserSpells(UserIndex)
Call ResetUserPets(UserIndex)
Call ResetUserBanco(UserIndex)
With UserList(UserIndex).ComUsu
    .Acepto = False
    
    For i = 1 To MAX_OFFER_SLOTS
        .Cant(i) = 0
        .Objeto(i) = 0
    Next i
    
    .GoldAmount = 0
    .DestNick = vbNullString
    .DestUsu = 0
End With
 
End Sub

Sub CloseUser(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo Errhandler

Dim N As Integer
Dim map As Integer
Dim Name As String
Dim i As Integer

Dim aN As Integer

aN = UserList(UserIndex).flags.AtacadoPorNpc
If aN > 0 Then
      Npclist(aN).Movement = Npclist(aN).flags.OldMovement
      Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
      Npclist(aN).flags.AttackedBy = vbNullString
End If

UserList(UserIndex).flags.AtacadoPorNpc = 0
UserList(UserIndex).flags.NPCAtacado = 0

map = UserList(UserIndex).Pos.map
Name = UCase$(UserList(UserIndex).Name)

UserList(UserIndex).Char.FX = 0
UserList(UserIndex).Char.loops = 0
Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, 0, 0))


UserList(UserIndex).flags.UserLogged = False
UserList(UserIndex).Counters.Saliendo = False

'Le devolvemos el body y head originales
If UserList(UserIndex).flags.AdminInvisible = 1 Then Call DoAdminInvisible(UserIndex)

' Grabamos el personaje del usuario
Call SaveUser(UserIndex, CharPath & Name & ".chr")

'usado para borrar Pjs
Call WriteVar(CharPath & UserList(UserIndex).Name & ".chr", "INIT", "Logged", "0")


'Quitar el dialogo
'If MapInfo(Map).NumUsers > 0 Then
'    Call SendToUserArea(UserIndex, "QDL" & UserList(UserIndex).Char.charindex)
'End If

If MapInfo(map).NumUsers > 0 Then
    Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageRemoveCharDialog(UserList(UserIndex).Char.CharIndex))
End If



'Borrar el personaje
If UserList(UserIndex).Char.CharIndex > 0 Then
    Call EraseUserChar(UserIndex, UserList(UserIndex).flags.AdminInvisible = 1)
End If

'Borrar mascotas
For i = 1 To MAXMASCOTAS
    If UserList(UserIndex).MascotasIndex(i) > 0 Then
        If Npclist(UserList(UserIndex).MascotasIndex(i)).flags.NPCActive Then _
            Call QuitarNPC(UserList(UserIndex).MascotasIndex(i))
    End If
Next i

'Update Map Users
MapInfo(map).NumUsers = MapInfo(map).NumUsers - 1

If MapInfo(map).NumUsers < 0 Then
    MapInfo(map).NumUsers = 0
End If

' Si el usuario habia dejado un msg en la gm's queue lo borramos
If Ayuda.Existe(UserList(UserIndex).Name) Then Call Ayuda.Quitar(UserList(UserIndex).Name)

Call ResetUserSlot(UserIndex)

Call MostrarNumUsers

N = FreeFile(1)
Open App.path & "\logs\Connect.log" For Append Shared As #N
Print #N, Name & " ha dejado el juego. " & "User Index:" & UserIndex & " " & time & " " & Date
Close #N

Exit Sub

Errhandler:
Call LogError("Error en CloseUser. Número " & Err.Number & " Descripción: " & Err.description)

End Sub

Sub ReloadSokcet()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo Errhandler

    Call LogApiSock("ReloadSokcet() " & NumUsers & " " & LastUser & " " & MaxUsers)
    
    If NumUsers <= 0 Then
        Call WSApiReiniciarSockets
    Else
'       Call apiclosesocket(SockListen)
'       SockListen = ListenForConnect(Puerto, hWndMsg, "")
    End If


Exit Sub
Errhandler:
    Call LogError("Error en CheckSocketState " & Err.Number & ": " & Err.description)

End Sub

Public Sub EnviarNoche(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Call WriteSendNight(UserIndex, IIf(DeNoche And (MapInfo(UserList(UserIndex).Pos.map).Zona = Campo Or MapInfo(UserList(UserIndex).Pos.map).Zona = Ciudad), True, False))
    Call WriteSendNight(UserIndex, IIf(DeNoche, True, False))
End Sub

Public Sub EcharPjsNoPrivilegiados()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim LoopC As Long
    
    For LoopC = 1 To LastUser
        If UserList(LoopC).flags.UserLogged And UserList(LoopC).ConnID >= 0 And UserList(LoopC).ConnIDValida Then
            If UserList(LoopC).flags.Privilegios And PlayerType.User Then
                Call CloseSocket(LoopC)
            End If
        End If
    Next LoopC

End Sub
