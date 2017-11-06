Attribute VB_Name = "ModAreas"
'--------------------------------------------------------------------------------
'    Component  : ModAreas
'    Project    : ServerFenix13
'    Author     : Lord Fers
'    Description: http://www.gs-zone.org/temas/areas-en-1024x768-render-a-23x17.91637/
'--------------------------------------------------------------------------------

Option Explicit
 
'>>>>>>AREAS>>>>>AREAS>>>>>>>>AREAS>>>>>>>AREAS>>>>>>>>>>
Public Type AreaInfo
    AreaPerteneceX As Integer
    AreaPerteneceY As Integer
 
    AreaReciveX As Integer
    AreaReciveY As Integer
 
    MinX As Integer '-!!!
    MinY As Integer '-!!!
 
    AreaID As Long
End Type
 
Public Type ConnGroup
    CountEntrys As Long
    OptValue As Long
    UserEntrys() As Long
End Type
 
Public Const USER_NUEVO As Byte = 255
 
'Cuidado:
' ���LAS AREAS EST�N HARDCODEADAS!!!
Private CurDay As Byte
Private CurHour As Byte
 
Private AreasInfo(1 To 100, 1 To 100) As Byte
Private PosToArea(1 To 100) As Byte
 
Private AreasRecive(10) As Integer
 
Public ConnGroups() As ConnGroup
 
Public Sub InitAreas()
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'
'**************************************************************
    Dim LoopC As Long
    Dim loopX As Long
 
' Setup areas...
    For LoopC = 0 To 9
        AreasRecive(LoopC) = (2 ^ LoopC) Or IIf(LoopC <> 0, 2 ^ (LoopC - 1), 0) Or IIf(LoopC <> 9, 2 ^ (LoopC + 1), 0)
    Next LoopC
 
    For LoopC = 1 To 100
        PosToArea(LoopC) = LoopC \ 11
    Next LoopC
 
    For LoopC = 1 To 100
        For loopX = 1 To 100
            'Usamos 121 IDs de area para saber si pasasamos de area "m�s r�pido"
            AreasInfo(LoopC, loopX) = (LoopC \ 11 + 1) * (loopX \ 11 + 1)
        Next loopX
    Next LoopC
 
'Setup AutoOptimizacion de areas
    CurDay = IIf(Weekday(Date) > 6, 1, 2) 'A ke tipo de dia pertenece?
    CurHour = Fix(Hour(time) \ 3) 'A ke parte de la hora pertenece
 
    ReDim ConnGroups(1 To NumMaps) As ConnGroup
 
    For LoopC = 1 To NumMaps
        ConnGroups(LoopC).OptValue = val(GetVar(DatPath & "AreasStats.dat", "Mapa" & LoopC, CurDay & "-" & CurHour))
 
        If ConnGroups(LoopC).OptValue = 0 Then ConnGroups(LoopC).OptValue = 1
        ReDim ConnGroups(LoopC).UserEntrys(1 To ConnGroups(LoopC).OptValue) As Long
    Next LoopC
End Sub
 
Public Sub AreasOptimizacion()
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'Es la funci�n de autooptimizacion.... la idea es no mandar redimensionando arrays grandes todo el tiempo
'**************************************************************
    Dim LoopC As Long
    Dim tCurDay As Byte
    Dim tCurHour As Byte
    Dim EntryValue As Long
 
    If (CurDay <> IIf(Weekday(Date) > 6, 1, 2)) Or (CurHour <> Fix(Hour(time) \ 3)) Then
 
        tCurDay = IIf(Weekday(Date) > 6, 1, 2) 'A ke tipo de dia pertenece?
        tCurHour = Fix(Hour(time) \ 3) 'A ke parte de la hora pertenece
 
        For LoopC = 1 To NumMaps
            EntryValue = val(GetVar(DatPath & "AreasStats.dat", "Mapa" & LoopC, CurDay & "-" & CurHour))
            Call WriteVar(DatPath & "AreasStats.dat", "Mapa" & LoopC, CurDay & "-" & CurHour, CInt((EntryValue + ConnGroups(LoopC).OptValue) \ 2))
     
            ConnGroups(LoopC).OptValue = val(GetVar(DatPath & "AreasStats.dat", "Mapa" & LoopC, tCurDay & "-" & tCurHour))
            If ConnGroups(LoopC).OptValue = 0 Then ConnGroups(LoopC).OptValue = 1
            If ConnGroups(LoopC).OptValue >= MapInfo(LoopC).NumUsers Then ReDim Preserve ConnGroups(LoopC).UserEntrys(1 To ConnGroups(LoopC).OptValue) As Long
        Next LoopC
 
        CurDay = tCurDay
        CurHour = tCurHour
    End If
End Sub
 
Public Sub CheckUpdateNeededUser(ByVal UserIndex As Integer, ByVal Head As Byte, Optional ByVal ButIndex As Boolean = False)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'Es la funci�n clave del sistema de areas... Es llamada al mover un user
'**************************************************************
    If UserList(UserIndex).AreasInfo.AreaID = AreasInfo(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y) Then Exit Sub
 
    Dim MinX As Long, MaxX As Long, MinY As Long, MaxY As Long, X As Long, Y As Long
    Dim TempInt As Long, map As Long
 
    With UserList(UserIndex)
        MinX = .AreasInfo.MinX
        MinY = .AreasInfo.MinY
 
        If Head = eHeading.NORTH Then
            MaxY = MinY - 1
            MinY = MinY - 11
            MaxX = MinX + 32
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
 
        ElseIf Head = eHeading.SOUTH Then
            MaxY = MinY + 44
            MinY = MinY + 33
            MaxX = MinX + 32
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY - 22)
 
        ElseIf Head = eHeading.WEST Then
            MaxX = MinX - 1
            MinX = MinX - 11
            MaxY = MinY + 32
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
 
 
        ElseIf Head = eHeading.EAST Then
            MaxX = MinX + 44
            MinX = MinX + 33
            MaxY = MinY + 32
            .AreasInfo.MinX = CInt(MinX - 22)
            .AreasInfo.MinY = CInt(MinY)
 
     
        ElseIf Head = USER_NUEVO Then
            'Esto pasa por cuando cambiamos de mapa o logeamos...
            MinY = ((.Pos.Y \ 11) - 1) * 11
            MaxY = MinY + 32
     
            MinX = ((.Pos.X \ 11) - 1) * 11
            MaxX = MinX + 32
     
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        End If
 
        If MinY < 1 Then MinY = 1
        If MinX < 1 Then MinX = 1
        If MaxY > 100 Then MaxY = 100
        If MaxX > 100 Then MaxX = 100
 
        map = UserList(UserIndex).Pos.map
 
        'Esto es para ke el cliente elimine lo "fuera de area..."
        Call WriteAreaChanged(UserIndex)
 
        'Actualizamos!!!
        For X = MinX To MaxX
            For Y = MinY To MaxY
         
                '<<< User >>>
                If MapData(map, X, Y).UserIndex Then
             
                    TempInt = MapData(map, X, Y).UserIndex
             
                    If UserIndex <> TempInt Then
                        Call MakeUserChar(False, UserIndex, TempInt, map, X, Y)
                        Call MakeUserChar(False, TempInt, UserIndex, .Pos.map, .Pos.X, .Pos.Y)
                 
                        'Si el user estaba invisible le avisamos al nuevo cliente de eso
                        If UserList(TempInt).flags.invisible Or UserList(TempInt).flags.Oculto Then
                            Call WriteSetInvisible(UserIndex, UserList(TempInt).Char.CharIndex, True)
                        End If
                        If UserList(UserIndex).flags.invisible Or UserList(UserIndex).flags.Oculto Then
                            Call WriteSetInvisible(TempInt, UserList(UserIndex).Char.CharIndex, True)
                        End If
                 
                        Call FlushBuffer(TempInt)
             
                    ElseIf Head = USER_NUEVO Then
                        If Not ButIndex Then
                            Call MakeUserChar(False, UserIndex, UserIndex, map, X, Y)
                        End If
                    End If
                End If
         
                '<<< Npc >>>
                If MapData(map, X, Y).NpcIndex Then
                    Call MakeNPCChar(False, UserIndex, MapData(map, X, Y).NpcIndex, map, X, Y)
                 End If
           
                '<<< Item >>>
                If MapData(map, X, Y).ObjInfo.OBJIndex Then
                    TempInt = MapData(map, X, Y).ObjInfo.OBJIndex
                    If Not EsObjetoFijo(ObjData(TempInt).OBJType) Then
                        Call WriteObjectCreate(UserIndex, ObjData(TempInt).GrhIndex, X, Y)
                 
                        If ObjData(TempInt).OBJType = eOBJType.otPuertas Then
                            Call Bloquear(False, UserIndex, X, Y, MapData(map, X, Y).Blocked)
                            Call Bloquear(False, UserIndex, X - 1, Y, MapData(map, X - 1, Y).Blocked)
                        End If
                    End If
                End If
     
            Next Y
        Next X
 
        'Precalculados :P
        TempInt = .Pos.X \ 11
        .AreasInfo.AreaReciveX = AreasRecive(TempInt)
        .AreasInfo.AreaPerteneceX = 2 ^ TempInt
 
        TempInt = .Pos.Y \ 11
        .AreasInfo.AreaReciveY = AreasRecive(TempInt)
        .AreasInfo.AreaPerteneceY = 2 ^ TempInt
 
        .AreasInfo.AreaID = AreasInfo(.Pos.X, .Pos.Y)
    End With
End Sub
 
Public Sub CheckUpdateNeededNpc(ByVal NpcIndex As Integer, ByVal Head As Byte)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
' Se llama cuando se mueve un Npc
'**************************************************************
    If Npclist(NpcIndex).AreasInfo.AreaID = AreasInfo(Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y) Then Exit Sub
 
    Dim MinX As Long, MaxX As Long, MinY As Long, MaxY As Long, X As Long, Y As Long
    Dim TempInt As Long
 
    With Npclist(NpcIndex)
        MinX = .AreasInfo.MinX
        MinY = .AreasInfo.MinY
 
        If Head = eHeading.NORTH Then
            MaxY = MinY - 1
            MinY = MinY - 11
            MaxX = MinX + 32
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
 
        ElseIf Head = eHeading.SOUTH Then
            MaxY = MinY + 44
            MinY = MinY + 33
            MaxX = MinX + 32
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY - 22)
 
        ElseIf Head = eHeading.WEST Then
            MaxX = MinX - 1
            MinX = MinX - 11
            MaxY = MinY + 32
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
 
 
        ElseIf Head = eHeading.EAST Then
            MaxX = MinX + 44
            MinX = MinX + 33
            MaxY = MinY + 32
            .AreasInfo.MinX = CInt(MinX - 22)
            .AreasInfo.MinY = CInt(MinY)
 
     
        ElseIf Head = USER_NUEVO Then
            'Esto pasa por cuando cambiamos de mapa o logeamos...
            MinY = ((.Pos.Y \ 11) - 1) * 11
            MaxY = MinY + 32
     
            MinX = ((.Pos.X \ 11) - 1) * 11
            MaxX = MinX + 32
     
            .AreasInfo.MinX = CInt(MinX)
            .AreasInfo.MinY = CInt(MinY)
        End If
 
        If MinY < 1 Then MinY = 1
        If MinX < 1 Then MinX = 1
        If MaxY > 100 Then MaxY = 100
        If MaxX > 100 Then MaxX = 100
 
 
        'Actualizamos!!!
        If MapInfo(.Pos.map).NumUsers <> 0 Then
            For X = MinX To MaxX
                For Y = MinY To MaxY
                    If MapData(.Pos.map, X, Y).UserIndex Then _
                        Call MakeNPCChar(False, MapData(.Pos.map, X, Y).UserIndex, NpcIndex, .Pos.map, .Pos.X, .Pos.Y)
                Next Y
            Next X
        End If
 
        'Precalculados :P
        TempInt = .Pos.X \ 11
        .AreasInfo.AreaReciveX = AreasRecive(TempInt)
        .AreasInfo.AreaPerteneceX = 2 ^ TempInt
     
        TempInt = .Pos.Y \ 11
        .AreasInfo.AreaReciveY = AreasRecive(TempInt)
        .AreasInfo.AreaPerteneceY = 2 ^ TempInt
 
        .AreasInfo.AreaID = AreasInfo(.Pos.X, .Pos.Y)
    End With
End Sub
 
Public Sub QuitarUser(ByVal UserIndex As Integer, ByVal map As Integer)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'
'**************************************************************
    Dim TempVal As Long
    Dim LoopC As Long
 
    'Search for the user
    For LoopC = 1 To ConnGroups(map).CountEntrys
        If ConnGroups(map).UserEntrys(LoopC) = UserIndex Then Exit For
    Next LoopC
 
    'Char not found
    If LoopC > ConnGroups(map).CountEntrys Then Exit Sub
 
    'Remove from old map
    ConnGroups(map).CountEntrys = ConnGroups(map).CountEntrys - 1
    TempVal = ConnGroups(map).CountEntrys
 
    'Move list back
    For LoopC = LoopC To TempVal
        ConnGroups(map).UserEntrys(LoopC) = ConnGroups(map).UserEntrys(LoopC + 1)
    Next LoopC
 
    If TempVal > ConnGroups(map).OptValue Then 'Nescesito Redim?
        ReDim Preserve ConnGroups(map).UserEntrys(1 To TempVal) As Long
    End If
End Sub
 
Public Sub AgregarUser(ByVal UserIndex As Integer, ByVal map As Integer, Optional ByVal ButIndex As Boolean = False)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: 04/01/2007
'Modified by Juan Mart�n Sotuyo Dodero (Maraxus)
'   - Now the method checks for repetead users instead of trusting parameters.
'   - If the character is new to the map, update it
'**************************************************************
    Dim TempVal As Long
    Dim EsNuevo As Boolean
    Dim i As Long
 
    If Not MapaValido(map) Then Exit Sub
 
    EsNuevo = True
 
    'Prevent adding repeated users
    For i = 1 To ConnGroups(map).CountEntrys
        If ConnGroups(map).UserEntrys(i) = UserIndex Then
            EsNuevo = False
            Exit For
        End If
    Next i
 
    If EsNuevo Then
        'Update map and connection groups data
        ConnGroups(map).CountEntrys = ConnGroups(map).CountEntrys + 1
        TempVal = ConnGroups(map).CountEntrys
 
        If TempVal > ConnGroups(map).OptValue Then 'Nescesito Redim
            ReDim Preserve ConnGroups(map).UserEntrys(1 To TempVal) As Long
        End If
 
        ConnGroups(map).UserEntrys(TempVal) = UserIndex
    End If
 
    'Update user
    UserList(UserIndex).AreasInfo.AreaID = 0
 
    UserList(UserIndex).AreasInfo.AreaPerteneceX = 0
    UserList(UserIndex).AreasInfo.AreaPerteneceY = 0
    UserList(UserIndex).AreasInfo.AreaReciveX = 0
    UserList(UserIndex).AreasInfo.AreaReciveY = 0
 
    Call CheckUpdateNeededUser(UserIndex, USER_NUEVO, ButIndex)
End Sub
 
Public Sub AgregarNpc(ByVal NpcIndex As Integer)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'
'**************************************************************
    Npclist(NpcIndex).AreasInfo.AreaID = 0
 
    Npclist(NpcIndex).AreasInfo.AreaPerteneceX = 0
    Npclist(NpcIndex).AreasInfo.AreaPerteneceY = 0
    Npclist(NpcIndex).AreasInfo.AreaReciveX = 0
    Npclist(NpcIndex).AreasInfo.AreaReciveY = 0
 
    Call CheckUpdateNeededNpc(NpcIndex, USER_NUEVO)
End Sub

