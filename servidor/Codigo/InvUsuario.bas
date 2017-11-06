Attribute VB_Name = "InvUsuario"
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

Public Function TieneObjetosRobables(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

'17/09/02
'Agregue que la función se asegure que el objeto no es un barco

On Error Resume Next

Dim i As Integer
Dim OBJIndex As Integer

For i = 1 To UserList(UserIndex).CurrentInventorySlots
    OBJIndex = UserList(UserIndex).Invent.Object(i).OBJIndex
    If OBJIndex > 0 Then
            If (ObjData(OBJIndex).OBJType <> eOBJType.otLlaves And _
                ObjData(OBJIndex).OBJType <> eOBJType.otBarcos) Then
                  TieneObjetosRobables = True
                  Exit Function
            End If
    
    End If
Next i
End Function

Function ClasePuedeUsarItem(ByVal UserIndex As Integer, ByVal OBJIndex As Integer, Optional ByRef sMotivo As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 14/01/2010 (ZaMa)
'14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
'***************************************************

On Error GoTo manejador
    
    'Admins can use ANYTHING!
    If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
        If ObjData(OBJIndex).ClaseProhibida(1) <> 0 Then
            Dim i As Integer
            For i = 1 To NUMCLASES
                If ObjData(OBJIndex).ClaseProhibida(i) = UserList(UserIndex).Clase Then
                    ClasePuedeUsarItem = False
                    sMotivo = "Tu clase no puede usar este objeto."
                    Exit Function
                End If
            Next i
        End If
    End If
    
    ClasePuedeUsarItem = True

Exit Function

manejador:
    LogError ("Error en ClasePuedeUsarItem")
End Function

Sub QuitarNewbieObj(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim j As Integer

With UserList(UserIndex)
    For j = 1 To UserList(UserIndex).CurrentInventorySlots
        If .Invent.Object(j).OBJIndex > 0 Then
             
             If ObjData(.Invent.Object(j).OBJIndex).Newbie = 1 Then _
                    Call QuitarUserInvItem(UserIndex, j, MAX_INVENTORY_OBJS)
                    Call UpdateUserInv(False, UserIndex, j)
        
        End If
    Next j
    
    '[Barrin 17-12-03] Si el usuario dejó de ser Newbie, y estaba en el Newbie Dungeon
    'es transportado a su hogar de origen ;)
    If MapInfo(.Pos.map).Restringir Then
        
        Dim DeDonde As WorldPos
        
        Select Case .Hogar
            Case eCiudad.cLindos 'Vamos a tener que ir por todo el desierto... uff!
                DeDonde = Lindos
            Case eCiudad.cUllathorpe
                DeDonde = Ullathorpe
            Case eCiudad.cBanderbill
                DeDonde = Banderbill
            Case Else
                DeDonde = Nix
        End Select
        
        Call WarpUserChar(UserIndex, DeDonde.map, DeDonde.X, DeDonde.Y, True)
    
    End If
    '[/Barrin]
End With

End Sub

Sub LimpiarInventario(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim j As Integer

With UserList(UserIndex)
    For j = 1 To .CurrentInventorySlots
        .Invent.Object(j).OBJIndex = 0
        .Invent.Object(j).Amount = 0
        .Invent.Object(j).Equipped = 0
    Next j
    
    .Invent.NroItems = 0
    
    .Invent.ArmourEqpObjIndex = 0
    .Invent.ArmourEqpSlot = 0
    
    .Invent.WeaponEqpObjIndex = 0
    .Invent.WeaponEqpSlot = 0
    
    .Invent.CascoEqpObjIndex = 0
    .Invent.CascoEqpSlot = 0
    
    .Invent.EscudoEqpObjIndex = 0
    .Invent.EscudoEqpSlot = 0

    .Invent.MunicionEqpObjIndex = 0
    .Invent.MunicionEqpSlot = 0
    
    .Invent.BarcoObjIndex = 0
    .Invent.BarcoSlot = 0
    
    .Invent.MochilaEqpObjIndex = 0
    .Invent.MochilaEqpSlot = 0
    
    .Invent.HerramientaEqpObjIndex = 0
    .Invent.HerramientaEqpslot = 0
    
End With

End Sub

Sub TirarOro(ByVal Cantidad As Long, ByVal UserIndex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 23/01/2007
'23/01/2007 -> Pablo (ToxicWaste): Billetera invertida y explotar oro en el agua.
'***************************************************
On Error GoTo ErrHandler

'If Cantidad > 100000 Then Exit Sub

With UserList(UserIndex)
    'SI EL Pjta TIENE ORO LO TIRAMOS
    If (Cantidad > 0) And (Cantidad <= .Stats.GLD) Then
            Dim MiObj As Obj
            'info debug
            Dim loops As Integer
            
            'Seguridad Alkon (guardo el oro tirado si supera los 50k)
            If Cantidad > 50000 Then
                Dim j As Integer
                Dim k As Integer
                Dim M As Integer
                Dim Cercanos As String
                M = .Pos.map
                For j = .Pos.X - 10 To .Pos.X + 10
                    For k = .Pos.Y - 10 To .Pos.Y + 10
                        If InMapBounds(M, j, k) Then
                            If MapData(M, j, k).UserIndex > 0 Then
                                Cercanos = Cercanos & UserList(MapData(M, j, k).UserIndex).Name & ","
                            End If
                        End If
                    Next k
                Next j
            End If
            '/Seguridad
            Dim Extra As Long
            Dim TeniaOro As Long
            TeniaOro = .Stats.GLD
            If Cantidad > 500000 Then 'Para evitar explotar demasiado
                Extra = Cantidad - 500000
                Cantidad = 500000
            End If
            
            Do While (Cantidad > 0)
                
                If Cantidad > MAX_INVENTORY_OBJS And .Stats.GLD > MAX_INVENTORY_OBJS Then
                    MiObj.Amount = MAX_INVENTORY_OBJS
                    Cantidad = Cantidad - MiObj.Amount
                Else
                    MiObj.Amount = Cantidad
                    Cantidad = Cantidad - MiObj.Amount
                End If
    
                MiObj.OBJIndex = iORO
                
                If EsGM(UserIndex) Then Call LogGM(.Name, "Tiró cantidad:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.OBJIndex).Name)
                Dim AuxPos As WorldPos
                
                If .Clase = eClass.Pirata And .Invent.BarcoObjIndex = 476 Then
                    AuxPos = TirarItemAlPiso(.Pos, MiObj, False)
                    If AuxPos.X <> 0 And AuxPos.Y <> 0 Then
                        .Stats.GLD = .Stats.GLD - MiObj.Amount
                    End If
                Else
                    AuxPos = TirarItemAlPiso(.Pos, MiObj, True)
                    If AuxPos.X <> 0 And AuxPos.Y <> 0 Then
                        .Stats.GLD = .Stats.GLD - MiObj.Amount
                    End If
                End If
                
                'info debug
                loops = loops + 1
                If loops > 100 Then
                    LogError ("Error en tiraroro")
                    Exit Sub
                End If
                
            Loop
            If TeniaOro = .Stats.GLD Then Extra = 0
            If Extra > 0 Then
                .Stats.GLD = .Stats.GLD - Extra
            End If
        
    End If
End With

Exit Sub

ErrHandler:
    Call LogError("Error en TirarOro. Error " & Err.Number & " : " & Err.description)
End Sub

Sub QuitarUserInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

    If Slot < 1 Or Slot > UserList(UserIndex).CurrentInventorySlots Then Exit Sub
    
    With UserList(UserIndex).Invent.Object(Slot)
        If .Amount <= Cantidad And .Equipped = 1 Then
            Call Desequipar(UserIndex, Slot)
        End If
        
        'Quita un objeto
        .Amount = .Amount - Cantidad
        '¿Quedan mas?
        If .Amount <= 0 Then
            UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
            .OBJIndex = 0
            .Amount = 0
        End If
    End With

Exit Sub

ErrHandler:
    Call LogError("Error en QuitarUserInvItem. Error " & Err.Number & " : " & Err.description)
    
End Sub

Sub UpdateUserInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

Dim NullObj As UserOBJ
Dim LoopC As Long

With UserList(UserIndex)
    'Actualiza un solo slot
    If Not UpdateAll Then
    
        'Actualiza el inventario
        If .Invent.Object(Slot).OBJIndex > 0 Then
            Call ChangeUserInv(UserIndex, Slot, .Invent.Object(Slot))
        Else
            Call ChangeUserInv(UserIndex, Slot, NullObj)
        End If
    
    Else
    
    'Actualiza todos los slots
        For LoopC = 1 To .CurrentInventorySlots
            'Actualiza el inventario
            If .Invent.Object(LoopC).OBJIndex > 0 Then
                Call ChangeUserInv(UserIndex, LoopC, .Invent.Object(LoopC))
            Else
                Call ChangeUserInv(UserIndex, LoopC, NullObj)
            End If
        Next LoopC
    End If
    
    Exit Sub
End With

ErrHandler:
    Call LogError("Error en UpdateUserInv. Error " & Err.Number & " : " & Err.description)

End Sub

Sub DropObj(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Num As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim Obj As Obj

With UserList(UserIndex)
    If Num > 0 Then
    
        If Num > .Invent.Object(Slot).Amount Then Num = .Invent.Object(Slot).Amount
        
        Obj.OBJIndex = .Invent.Object(Slot).OBJIndex
        Obj.Amount = Num
        
        If (ItemNewbie(Obj.OBJIndex) And (.flags.Privilegios And PlayerType.User)) Then
            Call WriteConsoleMsg(UserIndex, "No puedes tirar objetos newbie.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If ObjData(Obj.OBJIndex).NoComerciable Then
            Call WriteConsoleMsg(UserIndex, "No puedes tirar este objeto.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Check objeto en el suelo
        If MapData(.Pos.map, X, Y).ObjInfo.OBJIndex = 0 Or MapData(.Pos.map, X, Y).ObjInfo.OBJIndex = Obj.OBJIndex Then
            If Num + MapData(.Pos.map, X, Y).ObjInfo.Amount > MAX_INVENTORY_OBJS Then
                Num = MAX_INVENTORY_OBJS - MapData(.Pos.map, X, Y).ObjInfo.Amount
            End If
            
            Call MakeObj(Obj, map, X, Y)
            Call QuitarUserInvItem(UserIndex, Slot, Num)
            Call UpdateUserInv(False, UserIndex, Slot)
            
            If ObjData(Obj.OBJIndex).OBJType = eOBJType.otBarcos Then
                Call WriteConsoleMsg(UserIndex, "¡¡ATENCIÓN!! ¡ACABAS DE TIRAR TU BARCA!", FontTypeNames.FONTTYPE_TALK)
            End If
            
            If Not .flags.Privilegios And PlayerType.User Then Call LogGM(.Name, "Tiró cantidad:" & Num & " Objeto:" & ObjData(Obj.OBJIndex).Name)
            
        Else
            Call WriteConsoleMsg(UserIndex, "No hay espacio en el piso.", FontTypeNames.FONTTYPE_INFO)
        End If
    End If
End With

End Sub

Sub EraseObj(ByVal Num As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

With MapData(map, X, Y)
    .ObjInfo.Amount = .ObjInfo.Amount - Num
    
    If .ObjInfo.Amount <= 0 Then
        .ObjInfo.OBJIndex = 0
        .ObjInfo.Amount = 0
        
        Call modSendData.SendToAreaByPos(map, X, Y, PrepareMessageObjectDelete(X, Y))
    End If
End With

End Sub

Sub MakeObj(ByRef Obj As Obj, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
    
    If Obj.OBJIndex > 0 And Obj.OBJIndex <= UBound(ObjData) Then
    
        With MapData(map, X, Y)
            If .ObjInfo.OBJIndex = Obj.OBJIndex Then
                .ObjInfo.Amount = .ObjInfo.Amount + Obj.Amount
            Else
                .ObjInfo = Obj
                
                Call modSendData.SendToAreaByPos(map, X, Y, PrepareMessageObjectCreate(ObjData(Obj.OBJIndex).GrhIndex, X, Y))
            End If
        End With
    End If

End Sub

Function MeterItemEnInventario(ByVal UserIndex As Integer, ByRef MiObj As Obj) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

    Dim Slot As Byte
    
    With UserList(UserIndex)
        '¿el user ya tiene un objeto del mismo tipo?
        Slot = 1
        Do Until .Invent.Object(Slot).OBJIndex = MiObj.OBJIndex And _
                 .Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS
           Slot = Slot + 1
           If Slot > .CurrentInventorySlots Then
                 Exit Do
           End If
        Loop
            
        'Sino busca un slot vacio
        If Slot > .CurrentInventorySlots Then
           Slot = 1
           Do Until .Invent.Object(Slot).OBJIndex = 0
               Slot = Slot + 1
               If Slot > .CurrentInventorySlots Then
                   Call WriteConsoleMsg(UserIndex, "No puedes cargar más objetos.", FontTypeNames.FONTTYPE_FIGHT)
                   MeterItemEnInventario = False
                   Exit Function
               End If
           Loop
           .Invent.NroItems = .Invent.NroItems + 1
        End If
    
        If Slot > MAX_NORMAL_INVENTORY_SLOTS And Slot < MAX_INVENTORY_SLOTS Then
            If Not ItemSeCae(MiObj.OBJIndex) Then
                Call WriteConsoleMsg(UserIndex, "No puedes contener objetos especiales en tu " & ObjData(.Invent.MochilaEqpObjIndex).Name & ".", FontTypeNames.FONTTYPE_FIGHT)
                MeterItemEnInventario = False
                Exit Function
            End If
        End If
        'Mete el objeto
        If .Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS Then
           'Menor que MAX_INV_OBJS
           .Invent.Object(Slot).OBJIndex = MiObj.OBJIndex
           .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount + MiObj.Amount
        Else
           .Invent.Object(Slot).Amount = MAX_INVENTORY_OBJS
        End If
    End With
    
    MeterItemEnInventario = True
           
    Call UpdateUserInv(False, UserIndex, Slot)
    
    
    Exit Function
ErrHandler:
    Call LogError("Error en MeterItemEnInventario. Error " & Err.Number & " : " & Err.description)
End Function

Sub GetObj(ByVal UserIndex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 18/12/2009
'18/12/2009: ZaMa - Oro directo a la billetera.
'***************************************************

    Dim Obj As ObjData
    Dim MiObj As Obj
    
    With UserList(UserIndex)
        '¿Hay algun obj?
        If MapData(.Pos.map, .Pos.X, .Pos.Y).ObjInfo.OBJIndex > 0 Then
            '¿Esta permitido agarrar este obj?
            If ObjData(MapData(.Pos.map, .Pos.X, .Pos.Y).ObjInfo.OBJIndex).Agarrable <> 1 Then
                Dim X As Integer
                Dim Y As Integer
                
                X = .Pos.X
                Y = .Pos.Y
                
                Obj = ObjData(MapData(.Pos.map, .Pos.X, .Pos.Y).ObjInfo.OBJIndex)
                MiObj.Amount = MapData(.Pos.map, X, Y).ObjInfo.Amount
                MiObj.OBJIndex = MapData(.Pos.map, X, Y).ObjInfo.OBJIndex
                
                ' Oro directo a la billetera!
                If Obj.OBJType = otGuita Then
                    .Stats.GLD = .Stats.GLD + MiObj.Amount
                    'Quitamos el objeto
                    Call EraseObj(MapData(.Pos.map, X, Y).ObjInfo.Amount, .Pos.map, .Pos.X, .Pos.Y)
                        
                    Call WriteUpdateGold(UserIndex)
                Else
                    If MeterItemEnInventario(UserIndex, MiObj) Then
                    
                        'Quitamos el objeto
                        Call EraseObj(MapData(.Pos.map, X, Y).ObjInfo.Amount, .Pos.map, .Pos.X, .Pos.Y)
                        If Not .flags.Privilegios And PlayerType.User Then Call LogGM(.Name, "Agarro:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.OBJIndex).Name)
        
                    End If
                End If
            End If
        Else
            Call WriteConsoleMsg(UserIndex, "No hay nada aquí.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With

End Sub

Sub Desequipar(ByVal UserIndex As Integer, ByVal Slot As Byte)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

    'Desequipa el item slot del inventario
    Dim Obj As ObjData
    
    With UserList(UserIndex)
        With .Invent
            If (Slot < LBound(.Object)) Or (Slot > UBound(.Object)) Then
                Exit Sub
            ElseIf .Object(Slot).OBJIndex = 0 Then
                Exit Sub
            End If
            
            Obj = ObjData(.Object(Slot).OBJIndex)
        End With
        
        Select Case Obj.OBJType
            Case eOBJType.otWeapon
                With .Invent
                    .Object(Slot).Equipped = 0
                    .WeaponEqpObjIndex = 0
                    .WeaponEqpSlot = 0
                End With
                
                If Not .flags.Mimetizado = 1 Then
                    With .Char
                        .WeaponAnim = NingunArma
                        Call ChangeUserChar(UserIndex, .body, .Head, .heading, .WeaponAnim, .ShieldAnim, .CascoAnim)
                    End With
                End If
            
            Case eOBJType.otFlechas
                With .Invent
                    .Object(Slot).Equipped = 0
                    .MunicionEqpObjIndex = 0
                    .MunicionEqpSlot = 0
                End With
            
            Case eOBJType.otArmadura
                With .Invent
                    .Object(Slot).Equipped = 0
                    .ArmourEqpObjIndex = 0
                    .ArmourEqpSlot = 0
                End With
                
                Call DarCuerpoDesnudo(UserIndex, .flags.Mimetizado = 1)
                With .Char
                    Call ChangeUserChar(UserIndex, .body, .Head, .heading, .WeaponAnim, .ShieldAnim, .CascoAnim)
                End With
                 
            Case eOBJType.otCasco
                With .Invent
                    .Object(Slot).Equipped = 0
                    .CascoEqpObjIndex = 0
                    .CascoEqpSlot = 0
                End With
                
                If Not .flags.Mimetizado = 1 Then
                    With .Char
                        .CascoAnim = NingunCasco
                        Call ChangeUserChar(UserIndex, .body, .Head, .heading, .WeaponAnim, .ShieldAnim, .CascoAnim)
                    End With
                End If
            
            Case eOBJType.otEscudo
                With .Invent
                    .Object(Slot).Equipped = 0
                    .EscudoEqpObjIndex = 0
                    .EscudoEqpSlot = 0
                End With
                
                If Not .flags.Mimetizado = 1 Then
                    With .Char
                        .ShieldAnim = NingunEscudo
                        Call ChangeUserChar(UserIndex, .body, .Head, .heading, .WeaponAnim, .ShieldAnim, .CascoAnim)
                    End With
                End If
            
            Case eOBJType.otMochilas
                With .Invent
                    .Object(Slot).Equipped = 0
                    .MochilaEqpObjIndex = 0
                    .MochilaEqpSlot = 0
                End With
                
                Call InvUsuario.TirarTodosLosItemsEnMochila(UserIndex)
                .CurrentInventorySlots = MAX_NORMAL_INVENTORY_SLOTS
        End Select
    End With
    
    Call WriteUpdateUserStats(UserIndex)
    Call UpdateUserInv(False, UserIndex, Slot)
    
    Exit Sub

ErrHandler:
    Call LogError("Error en Desquipar. Error " & Err.Number & " : " & Err.description)

End Sub

Function SexoPuedeUsarItem(ByVal UserIndex As Integer, ByVal OBJIndex As Integer, Optional ByRef sMotivo As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 14/01/2010 (ZaMa)
'14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
'***************************************************

On Error GoTo ErrHandler
    
    If ObjData(OBJIndex).Mujer = 1 Then
        SexoPuedeUsarItem = UserList(UserIndex).Genero <> eGenero.Hombre
    ElseIf ObjData(OBJIndex).Hombre = 1 Then
        SexoPuedeUsarItem = UserList(UserIndex).Genero <> eGenero.Mujer
    Else
        SexoPuedeUsarItem = True
    End If
    
    If Not SexoPuedeUsarItem Then sMotivo = "Tu género no puede usar este objeto."
    
    Exit Function
ErrHandler:
    Call LogError("SexoPuedeUsarItem")
End Function


Function FaccionPuedeUsarItem(ByVal UserIndex As Integer, ByVal OBJIndex As Integer, Optional ByRef sMotivo As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 14/01/2010 (ZaMa)
'14/01/2010: ZaMa - Agrego el motivo por el que no puede equipar/usar el item.
'***************************************************

    If ObjData(OBJIndex).Real = 1 Then
        If Not criminal(UserIndex) Then
            FaccionPuedeUsarItem = EsArmada(UserIndex)
        Else
            FaccionPuedeUsarItem = False
        End If
    ElseIf ObjData(OBJIndex).Caos = 1 Then
        If criminal(UserIndex) Then
            FaccionPuedeUsarItem = EsCaos(UserIndex)
        Else
            FaccionPuedeUsarItem = False
        End If
    Else
        FaccionPuedeUsarItem = True
    End If
    
    If Not FaccionPuedeUsarItem Then sMotivo = "Tu alineación no puede usar este objeto."

End Function

Sub EquiparInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
'*************************************************
'Author: Unknown
'Last modified: 14/01/2010 (ZaMa)
'01/08/2009: ZaMa - Now it's not sent any sound made by an invisible admin
'14/01/2010: ZaMa - Agrego el motivo especifico por el que no puede equipar/usar el item.
'*************************************************

On Error GoTo ErrHandler

    'Equipa un item del inventario
    Dim Obj As ObjData
    Dim OBJIndex As Integer
    Dim sMotivo As String
    
    With UserList(UserIndex)
        OBJIndex = .Invent.Object(Slot).OBJIndex
        Obj = ObjData(OBJIndex)
        
        If Obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
             Call WriteConsoleMsg(UserIndex, "Sólo los newbies pueden usar este objeto.", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
                
        Select Case Obj.OBJType
            Case eOBJType.otWeapon
               If ClasePuedeUsarItem(UserIndex, OBJIndex, sMotivo) And _
                  FaccionPuedeUsarItem(UserIndex, OBJIndex, sMotivo) And _
                  SkillPuedeUsarItem(UserIndex, OBJIndex, sMotivo) Then
                    'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        'Quitamos del inv el item
                        Call Desequipar(UserIndex, Slot)
                        'Animacion por defecto
                        If .flags.Mimetizado = 1 Then
                            .CharMimetizado.WeaponAnim = NingunArma
                        Else
                            .Char.WeaponAnim = NingunArma
                            Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        End If
                        Exit Sub
                    End If
                    
                    'Quitamos el elemento anterior
                    If .Invent.WeaponEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)
                    End If
                    
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.WeaponEqpObjIndex = OBJIndex
                    .Invent.WeaponEqpSlot = Slot
                    
                    'El sonido solo se envia si no lo produce un admin invisible
                    If Not (.flags.AdminInvisible = 1) Then _
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SACARARMA, .Pos.X, .Pos.Y))
                    
                    If .flags.Mimetizado = 1 Then
                        .CharMimetizado.WeaponAnim = ObjData(OBJIndex).WeaponAnim
                    Else
                        .Char.WeaponAnim = ObjData(OBJIndex).WeaponAnim
                        Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                    End If
               Else
                    Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)
               End If
            
            Case eOBJType.otFlechas
               If ClasePuedeUsarItem(UserIndex, OBJIndex, sMotivo) And _
                  FaccionPuedeUsarItem(UserIndex, OBJIndex, sMotivo) And _
                  SkillPuedeUsarItem(UserIndex, OBJIndex, sMotivo) Then
                        
                        'Si esta equipado lo quita
                        If .Invent.Object(Slot).Equipped Then
                            'Quitamos del inv el item
                            Call Desequipar(UserIndex, Slot)
                            Exit Sub
                        End If
                        
                        'Quitamos el elemento anterior
                        If .Invent.MunicionEqpObjIndex > 0 Then
                            Call Desequipar(UserIndex, .Invent.MunicionEqpSlot)
                        End If
                
                        .Invent.Object(Slot).Equipped = 1
                        .Invent.MunicionEqpObjIndex = OBJIndex
                        .Invent.MunicionEqpSlot = Slot
                        
               Else
                    Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)
               End If
            
            Case eOBJType.otArmadura
                If .flags.Navegando = 1 Then Exit Sub
                
                'Nos aseguramos que puede usarla
                If ClasePuedeUsarItem(UserIndex, OBJIndex, sMotivo) And _
                   SexoPuedeUsarItem(UserIndex, OBJIndex, sMotivo) And _
                   CheckRazaUsaRopa(UserIndex, OBJIndex, sMotivo) And _
                   FaccionPuedeUsarItem(UserIndex, OBJIndex, sMotivo) And _
                   SkillPuedeUsarItem(UserIndex, OBJIndex, sMotivo) Then
                   
                   'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        Call Desequipar(UserIndex, Slot)
                        Call DarCuerpoDesnudo(UserIndex, .flags.Mimetizado = 1)
                        If Not .flags.Mimetizado = 1 Then
                            Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        End If
                        Exit Sub
                    End If
            
                    'Quita el anterior
                    If .Invent.ArmourEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.ArmourEqpSlot)
                    End If
            
                    'Lo equipa
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.ArmourEqpObjIndex = OBJIndex
                    .Invent.ArmourEqpSlot = Slot
                        
                    If .flags.Mimetizado = 1 Then
                        .CharMimetizado.body = Obj.Ropaje
                    Else
                        .Char.body = Obj.Ropaje
                        Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                    End If
                    .flags.Desnudo = 0
                Else
                    Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eOBJType.otCasco
                If .flags.Navegando = 1 Then Exit Sub
                If ClasePuedeUsarItem(UserIndex, OBJIndex, sMotivo) And _
                SkillPuedeUsarItem(UserIndex, OBJIndex, sMotivo) Then
                    'Si esta equipado lo quita
                    If .Invent.Object(Slot).Equipped Then
                        Call Desequipar(UserIndex, Slot)
                        If .flags.Mimetizado = 1 Then
                            .CharMimetizado.CascoAnim = NingunCasco
                        Else
                            .Char.CascoAnim = NingunCasco
                            Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        End If
                        Exit Sub
                    End If
            
                    'Quita el anterior
                    If .Invent.CascoEqpObjIndex > 0 Then
                        Call Desequipar(UserIndex, .Invent.CascoEqpSlot)
                    End If
            
                    'Lo equipa
                    
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.CascoEqpObjIndex = OBJIndex
                    .Invent.CascoEqpSlot = Slot
                    If .flags.Mimetizado = 1 Then
                        .CharMimetizado.CascoAnim = Obj.CascoAnim
                    Else
                        .Char.CascoAnim = Obj.CascoAnim
                        Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eOBJType.otEscudo
                If .flags.Navegando = 1 Then Exit Sub
                
                 If ClasePuedeUsarItem(UserIndex, OBJIndex, sMotivo) And _
                     FaccionPuedeUsarItem(UserIndex, OBJIndex, sMotivo) And _
                     SkillPuedeUsarItem(UserIndex, OBJIndex, sMotivo) Then
        
                     'Si esta equipado lo quita
                     If .Invent.Object(Slot).Equipped Then
                         Call Desequipar(UserIndex, Slot)
                         If .flags.Mimetizado = 1 Then
                             .CharMimetizado.ShieldAnim = NingunEscudo
                         Else
                             .Char.ShieldAnim = NingunEscudo
                             Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                         End If
                         Exit Sub
                     End If
             
                     'Quita el anterior
                     If .Invent.EscudoEqpObjIndex > 0 Then
                         Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)
                     End If
             
                     'Lo equipa
                     
                     .Invent.Object(Slot).Equipped = 1
                     .Invent.EscudoEqpObjIndex = OBJIndex
                     .Invent.EscudoEqpSlot = Slot
                     
                     If .flags.Mimetizado = 1 Then
                         .CharMimetizado.ShieldAnim = Obj.ShieldAnim
                     Else
                         .Char.ShieldAnim = Obj.ShieldAnim
                         
                         Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                     End If
                 Else
                     Call WriteConsoleMsg(UserIndex, sMotivo, FontTypeNames.FONTTYPE_INFO)
                 End If
                 
            Case eOBJType.otMochilas
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If .Invent.Object(Slot).Equipped Then
                    Call Desequipar(UserIndex, Slot)
                    Exit Sub
                End If
                If .Invent.MochilaEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, .Invent.MochilaEqpSlot)
                End If
                .Invent.Object(Slot).Equipped = 1
                .Invent.MochilaEqpObjIndex = OBJIndex
                .Invent.MochilaEqpSlot = Slot
                .CurrentInventorySlots = MAX_NORMAL_INVENTORY_SLOTS + Obj.MochilaType * 5
                Call WriteAddSlots(UserIndex, Obj.MochilaType)
        End Select
    End With
    
    'Actualiza
    Call UpdateUserInv(False, UserIndex, Slot)
    
    Exit Sub
    
ErrHandler:
    Call LogError("EquiparInvItem Slot:" & Slot & " - Error: " & Err.Number & " - Error Description : " & Err.description)
End Sub

'CSEH: ErrLog
Private Function CheckRazaUsaRopa(ByVal UserIndex As Integer, ItemIndex As Integer, Optional ByRef sMotivo As String) As Boolean
    '<EhHeader>
    On Error GoTo CheckRazaUsaRopa_Err
    '</EhHeader>
100     With UserList(UserIndex)
105         If .flags.Privilegios Then
110             CheckRazaUsaRopa = True
                Exit Function
            End If

115         If Len(ObjData(ItemIndex).RazaProhibida(1)) > 0 Then
                Dim i As Integer
120             For i = 1 To NUMRAZAS
125                 If (ObjData(ItemIndex).RazaProhibida(i)) = .raza Then
130                     CheckRazaUsaRopa = False
                        Exit Function
                    End If
                Next
135             CheckRazaUsaRopa = True
            Else
140             CheckRazaUsaRopa = True
            End If

        End With
    
145     If Not CheckRazaUsaRopa Then sMotivo = "Tu raza no puede usar este objeto."
    '<EhFooter>
    Exit Function

CheckRazaUsaRopa_Err:
        Call LogError("Error en CheckRazaUsaRopa: " & Erl & " - " & Err.description)
    '</EhFooter>
End Function

Sub UseInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
'*************************************************
'Author: Unknown
'Last modified: 10/12/2009
'Handels the usage of items from inventory box.
'24/01/2007 Pablo (ToxicWaste) - Agrego el Cuerno de la Armada y la Legión.
'24/01/2007 Pablo (ToxicWaste) - Utilización nueva de Barco en lvl 20 por clase Pirata y Pescador.
'01/08/2009: ZaMa - Now it's not sent any sound made by an invisible admin, except to its own client
'17/11/2009: ZaMa - Ahora se envia una orientacion de la posicion hacia donde esta el que uso el cuerno.
'27/11/2009: Budi - Se envia indivualmente cuando se modifica a la Agilidad o la Fuerza del personaje.
'08/12/2009: ZaMa - Agrego el uso de hacha de madera elfica.
'10/12/2009: ZaMa - Arreglos y validaciones en todos las herramientas de trabajo.
'*************************************************

    Dim Obj As ObjData
    Dim OBJIndex As Integer
    Dim TargObj As ObjData
    Dim MiObj As Obj
    
    With UserList(UserIndex)
    
        If .Invent.Object(Slot).Amount = 0 Then Exit Sub
        
        Obj = ObjData(.Invent.Object(Slot).OBJIndex)
        
        If Obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
            Call WriteConsoleMsg(UserIndex, "Sólo los newbies pueden usar estos objetos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Obj.OBJType = eOBJType.otWeapon Then
            If Obj.proyectil = 1 Then
                
                'valido para evitar el flood pero no bloqueo. El bloqueo se hace en WLC con proyectiles.
                If Not IntervaloPermiteUsar(UserIndex, False) Then Exit Sub
            Else
                'dagas
                If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
            End If
        Else
            If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
        End If
        
        OBJIndex = .Invent.Object(Slot).OBJIndex
        .flags.TargetObjInvIndex = OBJIndex
        .flags.TargetObjInvSlot = Slot
        
        Select Case Obj.OBJType
            Case eOBJType.otWarp
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If Not .flags.TargetNpcTipo = eNPCType.Pirata Then
                    Call WriteConsoleMsg(UserIndex, "No puedes usar este objeto con este npc.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                Else
                    If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 4 Then
                        Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos!", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    Else
                        If Obj.WI = .Pos.map Then
                            Call WarpUserChar(UserIndex, Obj.WMapa, Obj.WX, Obj.WY, True)
                            Call QuitarUserInvItem(UserIndex, Slot, 1)
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_WARP, .Pos.X, .Pos.Y))
                            Call UpdateUserInv(False, UserIndex, Slot)
                        Else
                            Call WriteChatOverHead(UserIndex, "Ese pasaje no te lo he vendido yo, lárgate!", Npclist(.flags.TargetNPC).Char.CharIndex, -1)
                            Exit Sub
                        End If
                    End If
                End If
                
            Case eOBJType.otUseOnce
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
        
                'Usa el item
                .Stats.MinHam = .Stats.MinHam + Obj.MinHam
                If .Stats.MinHam > .Stats.MaxHam Then _
                    .Stats.MinHam = .Stats.MaxHam
                .flags.Hambre = 0
                Call WriteUpdateHungerAndThirst(UserIndex)
                'Sonido
                
                If OBJIndex = e_ObjetosCriticos.Manzana Or OBJIndex = e_ObjetosCriticos.Manzana2 Or OBJIndex = e_ObjetosCriticos.ManzanaNewbie Then
                    Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.MORFAR_MANZANA)
                Else
                    Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.SOUND_COMIDA)
                End If
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                
                Call UpdateUserInv(False, UserIndex, Slot)
        
            Case eOBJType.otGuita
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                .Stats.GLD = .Stats.GLD + .Invent.Object(Slot).Amount
                .Invent.Object(Slot).Amount = 0
                .Invent.Object(Slot).OBJIndex = 0
                .Invent.NroItems = .Invent.NroItems - 1
                
                Call UpdateUserInv(False, UserIndex, Slot)
                Call WriteUpdateGold(UserIndex)
                
            Case eOBJType.otWeapon
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If Not .Stats.MinSta > 0 Then
                    Call WriteConsoleMsg(UserIndex, "Estás muy cansad" & _
                                IIf(.Genero = eGenero.Hombre, "o", "a") & ".", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If ObjData(OBJIndex).proyectil = 1 Then
                    If .Invent.Object(Slot).Equipped = 0 Then
                        Call WriteConsoleMsg(UserIndex, "Antes de usar la herramienta deberías equipartela.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, eSkill.Proyectiles)  'Call WriteWorkRequestTarget(UserIndex, Proyectiles)
                ElseIf .flags.TargetObj = Leña Then
                    If .Invent.Object(Slot).OBJIndex = DAGA Then
                        If .Invent.Object(Slot).Equipped = 0 Then
                            Call WriteConsoleMsg(UserIndex, "Antes de usar la herramienta deberías equipartela.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                            
                        Call TratarDeHacerFogata(.flags.TargetObjMap, _
                            .flags.TargetObjX, .flags.TargetObjY, UserIndex)
                    End If
                End If
            
            Case eOBJType.otHerramientas

                Select Case OBJIndex
                        Case CAÑA_PESCA, RED_PESCA
                            If .Invent.HerramientaEqpObjIndex = CAÑA_PESCA Or .Invent.HerramientaEqpObjIndex = RED_PESCA Then
                                Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, eSkill.Pesca)  'Call WriteWorkRequestTarget(UserIndex, eSkill.Pesca)
                            Else
                                 Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)
                            End If
                            
                        Case HACHA_LEÑADOR, HACHA_LEÑA_ELFICA
                            If .Invent.HerramientaEqpObjIndex = HACHA_LEÑADOR Or .Invent.HerramientaEqpObjIndex = HACHA_LEÑA_ELFICA Then
                                Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, eSkill.Talar)
                            Else
                                Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)
                            End If
                            
                        Case PIQUETE_MINERO
                            If .Invent.HerramientaEqpObjIndex = PIQUETE_MINERO Then
                                Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, eSkill.Mineria)
                            Else
                                Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)
                            End If
                            
                        Case MARTILLO_HERRERO
                            If .Invent.HerramientaEqpObjIndex = MARTILLO_HERRERO Then
                                Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, eSkill.Herreria)
                            Else
                                Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)
                            End If
                            
                        Case SERRUCHO_CARPINTERO
                            If .Invent.HerramientaEqpObjIndex = SERRUCHO_CARPINTERO Then
                                Call EnivarObjConstruibles(UserIndex)
                                Call WriteShowCarpenterForm(UserIndex)
                            Else
                                Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)
                            End If
                        
                        Case HILAR_SASTRE
                            If .Invent.HerramientaEqpObjIndex = HILAR_SASTRE Then
                                'callenviarsastreobj(userindex)
                                'call writeshowsastreform(userindex)
                            Else
                                Call WriteConsoleMsg(UserIndex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)
                            End If
                        Case Else ' Las herramientas no se pueden fundir
                            If ObjData(OBJIndex).SkHerreria > 0 Then
                                ' Solo objetos que pueda hacer el herrero
                                Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, FundirMetal) 'Call WriteWorkRequestTarget(UserIndex, FundirMetal)
                            End If
                    End Select
            Case eOBJType.otPociones
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo. ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If Not IntervaloPermiteGolpeUsar(UserIndex, False) Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Debes esperar unos momentos para tomar otra poción!!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                .flags.TomoPocion = True
                .flags.TipoPocion = Obj.TipoPocion
                        
                Select Case .flags.TipoPocion
                
                    Case 1 'Modif la agilidad
                        .flags.DuracionEfecto = Obj.DuracionEfecto
                
                        'Usa el item
                        .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
                        If .Stats.UserAtributos(eAtributos.Agilidad) > MAXATRIBUTOS Then _
                            .Stats.UserAtributos(eAtributos.Agilidad) = MAXATRIBUTOS
                        If .Stats.UserAtributos(eAtributos.Agilidad) > 2 * .Stats.UserAtributosBackUP(Agilidad) Then .Stats.UserAtributos(eAtributos.Agilidad) = 2 * .Stats.UserAtributosBackUP(Agilidad)
                        
                        'Quitamos del inv el item
                        Call QuitarUserInvItem(UserIndex, Slot, 1)
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        End If
                        Call WriteUpdateDexterity(UserIndex)
                        
                    Case 2 'Modif la fuerza
                        .flags.DuracionEfecto = Obj.DuracionEfecto
                
                        'Usa el item
                        .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
                        If .Stats.UserAtributos(eAtributos.Fuerza) > MAXATRIBUTOS Then _
                            .Stats.UserAtributos(eAtributos.Fuerza) = MAXATRIBUTOS
                        If .Stats.UserAtributos(eAtributos.Fuerza) > 2 * .Stats.UserAtributosBackUP(Fuerza) Then .Stats.UserAtributos(eAtributos.Fuerza) = 2 * .Stats.UserAtributosBackUP(Fuerza)
                        
                        
                        'Quitamos del inv el item
                        Call QuitarUserInvItem(UserIndex, Slot, 1)
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        End If
                        Call WriteUpdateStrenght(UserIndex)
                        
                    Case 3 'Pocion roja, restaura HP
                        'Usa el item
                        .Stats.MinHp = .Stats.MinHp + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
                        If .Stats.MinHp > .Stats.MaxHp Then _
                            .Stats.MinHp = .Stats.MaxHp
                        
                        'Quitamos del inv el item
                        Call QuitarUserInvItem(UserIndex, Slot, 1)
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        End If
                    
                    Case 4 'Pocion azul, restaura MANA
                        'Usa el item
                        'nuevo calculo para recargar mana
                        .Stats.MinMAN = .Stats.MinMAN + Porcentaje(.Stats.MaxMAN, 4) + .Stats.ELV \ 2 + 40 / .Stats.ELV
                        If .Stats.MinMAN > .Stats.MaxMAN Then _
                            .Stats.MinMAN = .Stats.MaxMAN
                        
                        'Quitamos del inv el item
                        Call QuitarUserInvItem(UserIndex, Slot, 1)
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        End If
                        
                    Case 5 ' Pocion violeta
                        If .flags.Envenenado = 1 Then
                            .flags.Envenenado = 0
                            Call WriteConsoleMsg(UserIndex, "Te has curado del envenenamiento.", FontTypeNames.FONTTYPE_INFO)
                        End If
                        'Quitamos del inv el item
                        Call QuitarUserInvItem(UserIndex, Slot, 1)
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        End If
                        
                    Case 6  ' Pocion Negra
                        If .flags.Privilegios And PlayerType.User Then
                            Call QuitarUserInvItem(UserIndex, Slot, 1)
                            Call UserDie(UserIndex)
                            Call WriteConsoleMsg(UserIndex, "Sientes un gran mareo y pierdes el conocimiento.", FontTypeNames.FONTTYPE_FIGHT)
                        End If
               End Select
               Call WriteUpdateUserStats(UserIndex)
               Call UpdateUserInv(False, UserIndex, Slot)
        
             Case eOBJType.otBebidas
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                .Stats.MinAGU = .Stats.MinAGU + Obj.MinSed
                If .Stats.MinAGU > .Stats.MaxAGU Then _
                    .Stats.MinAGU = .Stats.MaxAGU
                .flags.Sed = 0
                Call WriteUpdateHungerAndThirst(UserIndex)
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                
                ' Los admin invisibles solo producen sonidos a si mismos
                If .flags.AdminInvisible = 1 Then
                    Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                End If
                
                Call UpdateUserInv(False, UserIndex, Slot)
            
            Case eOBJType.otLlaves
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If .flags.TargetObj = 0 Then Exit Sub
                TargObj = ObjData(.flags.TargetObj)
                '¿El objeto clickeado es una puerta?
                If TargObj.OBJType = eOBJType.otPuertas Then
                    '¿Esta cerrada?
                    If TargObj.Cerrada = 1 Then
                          '¿Cerrada con llave?
                          If TargObj.Llave > 0 Then
                             If TargObj.clave = Obj.clave Then
                 
                                MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.OBJIndex _
                                = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.OBJIndex).IndexCerrada
                                .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.OBJIndex
                                Call WriteConsoleMsg(UserIndex, "Has abierto la puerta.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                             Else
                                Call WriteConsoleMsg(UserIndex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                             End If
                          Else
                             If TargObj.clave = Obj.clave Then
                                MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.OBJIndex _
                                = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.OBJIndex).IndexCerradaLlave
                                Call WriteConsoleMsg(UserIndex, "Has cerrado con llave la puerta.", FontTypeNames.FONTTYPE_INFO)
                                .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.OBJIndex
                                Exit Sub
                             Else
                                Call WriteConsoleMsg(UserIndex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                             End If
                          End If
                    Else
                          Call WriteConsoleMsg(UserIndex, "No está cerrada.", FontTypeNames.FONTTYPE_INFO)
                          Exit Sub
                    End If
                End If
            
            Case eOBJType.otBotellaVacia
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If MapData(.Pos.map, .flags.TargetX, .flags.TargetY).Agua <> 1 Then
                    Call WriteConsoleMsg(UserIndex, "No hay agua allí.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                MiObj.Amount = 1
                MiObj.OBJIndex = ObjData(.Invent.Object(Slot).OBJIndex).IndexAbierta
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                If Not MeterItemEnInventario(UserIndex, MiObj) Then
                    Call TirarItemAlPiso(.Pos, MiObj)
                End If
                
                Call UpdateUserInv(False, UserIndex, Slot)
            
            Case eOBJType.otBotellaLlena
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                .Stats.MinAGU = .Stats.MinAGU + Obj.MinSed
                If .Stats.MinAGU > .Stats.MaxAGU Then _
                    .Stats.MinAGU = .Stats.MaxAGU
                .flags.Sed = 0
                Call WriteUpdateHungerAndThirst(UserIndex)
                MiObj.Amount = 1
                MiObj.OBJIndex = ObjData(.Invent.Object(Slot).OBJIndex).IndexCerrada
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                If Not MeterItemEnInventario(UserIndex, MiObj) Then
                    Call TirarItemAlPiso(.Pos, MiObj)
                End If
                
                Call UpdateUserInv(False, UserIndex, Slot)
            
            Case eOBJType.otPergaminos
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If .Stats.MaxMAN > 0 Then
                    If .flags.Hambre = 0 And _
                        .flags.Sed = 0 Then
                        Call AgregarHechizo(UserIndex, Slot)
                        Call UpdateUserInv(False, UserIndex, Slot)
                    Else
                        Call WriteConsoleMsg(UserIndex, "Estás demasiado hambriento y sediento.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "No tienes conocimientos de las Artes Arcanas.", FontTypeNames.FONTTYPE_INFO)
                End If
            Case eOBJType.otMinerales
                If .flags.Muerto = 1 Then
                     Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                     Exit Sub
                End If
                Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, FundirMetal) 'Call WriteWorkRequestTarget(UserIndex, FundirMetal)
               
            Case eOBJType.otInstrumentos
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo puedes usar ítems cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If Obj.Real Then '¿Es el Cuerno Real?
                    If FaccionPuedeUsarItem(UserIndex, OBJIndex) Then
                        If MapInfo(.Pos.map).Pk = False Then
                            Call WriteConsoleMsg(UserIndex, "No hay peligro aquí. Es zona segura.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y))
                        Else
                            Call AlertarFaccionarios(UserIndex)
                            Call SendData(SendTarget.toMap, .Pos.map, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y))
                        End If
                        
                        Exit Sub
                    Else
                        Call WriteConsoleMsg(UserIndex, "Sólo miembros del ejército real pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                ElseIf Obj.Caos Then '¿Es el Cuerno Legión?
                    If FaccionPuedeUsarItem(UserIndex, OBJIndex) Then
                        If MapInfo(.Pos.map).Pk = False Then
                            Call WriteConsoleMsg(UserIndex, "No hay peligro aquí. Es zona segura.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        ' Los admin invisibles solo producen sonidos a si mismos
                        If .flags.AdminInvisible = 1 Then
                            Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y))
                        Else
                            Call AlertarFaccionarios(UserIndex)
                            Call SendData(SendTarget.toMap, .Pos.map, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y))
                        End If
                        
                        Exit Sub
                    Else
                        Call WriteConsoleMsg(UserIndex, "Sólo miembros de la legión oscura pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                End If
                'Si llega aca es porque es o Laud o Tambor o Flauta
                ' Los admin invisibles solo producen sonidos a si mismos
                If .flags.AdminInvisible = 1 Then
                    Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y))
                End If
               
            Case eOBJType.otBarcos
                'Verifica si esta aproximado al agua antes de permitirle navegar
                If .Stats.ELV < 25 Then
                    If Not esTrabajador(.Clase) And .Clase <> eClass.Pirata Then
                        Call WriteConsoleMsg(UserIndex, "Para recorrer los mares debes ser nivel 25 o superior.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    Else
                        If .Stats.ELV < 20 Then
                            Call WriteConsoleMsg(UserIndex, "Para recorrer los mares debes ser nivel 20 o superior.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                    End If
                End If
                
                If ((LegalPos(.Pos.map, .Pos.X - 1, .Pos.Y, True, False) _
                        Or LegalPos(.Pos.map, .Pos.X, .Pos.Y - 1, True, False) _
                        Or LegalPos(.Pos.map, .Pos.X + 1, .Pos.Y, True, False) _
                        Or LegalPos(.Pos.map, .Pos.X, .Pos.Y + 1, True, False)) _
                        And .flags.Navegando = 0) _
                        Or .flags.Navegando = 1 Then
                    Call DoNavega(UserIndex, Obj, Slot)
                Else
                    Call WriteConsoleMsg(UserIndex, "¡Debes aproximarte al agua para usar el barco!", FontTypeNames.FONTTYPE_INFO)
                End If
                
        End Select
    
    End With

End Sub

Sub EnivarArmasConstruibles(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Call WriteBlacksmithWeapons(UserIndex)
End Sub
 
Sub EnivarObjConstruibles(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Call WriteCarpenterObjects(UserIndex)
End Sub

Sub EnivarArmadurasConstruibles(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Call WriteBlacksmithArmors(UserIndex)
End Sub

Sub TirarTodo(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error Resume Next

    With UserList(UserIndex)
        If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = 6 Then Exit Sub
        
        Call TirarTodosLosItems(UserIndex)
        
        Dim Cantidad As Long
        Cantidad = .Stats.GLD - CLng(.Stats.ELV) * 10000
        
        If Cantidad > 0 Then _
            Call TirarOro(Cantidad, UserIndex)
    End With

End Sub

Public Function ItemSeCae(ByVal index As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With ObjData(index)
        ItemSeCae = (.Real <> 1 Or .NoSeCae = 0) And _
                    (.Caos <> 1 Or .NoSeCae = 0) And _
                    .OBJType <> eOBJType.otLlaves And _
                    .OBJType <> eOBJType.otBarcos And _
                    .NoSeCae = 0
    End With

End Function

Sub TirarTodosLosItems(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 12/01/2010 (ZaMa)
'12/01/2010: ZaMa - Ahora los piratas no explotan items solo si estan entre 20 y 25
'***************************************************

    Dim i As Byte
    Dim NuevaPos As WorldPos
    Dim MiObj As Obj
    Dim ItemIndex As Integer
    Dim DropAgua As Boolean
    Dim PosibilidadesZafa As Integer
    Dim ZafaMinerales As Boolean
    
    With UserList(UserIndex)
        
        If .Clase = eClass.Pirata And .Recompensas(2) = 1 And (RandomNumber(1, 10) <= 1) Then Exit Sub
        
        If .Clase = eClass.Minero Then
            If .Recompensas(1) = 2 Then PosibilidadesZafa = 2
            If .Recompensas(3) = 2 Then PosibilidadesZafa = PosibilidadesZafa + 3
            ZafaMinerales = CInt(RandomNumber(1, 10)) <= PosibilidadesZafa
        End If

        For i = 1 To .CurrentInventorySlots
            ItemIndex = .Invent.Object(i).OBJIndex
            If ItemIndex > 0 Then
                 If ItemSeCae(ItemIndex) And Not (ObjData(ItemIndex).OBJType = eOBJType.otMinerales And ZafaMinerales) Then
                    NuevaPos.X = 0
                    NuevaPos.Y = 0
                    
                    'Creo el Obj
                    MiObj.Amount = .Invent.Object(i).Amount
                    MiObj.OBJIndex = ItemIndex

                    DropAgua = True
                    ' Es pirata?
                    If .Clase = eClass.Pirata Then
                        ' Si tiene galeon equipado
                        If .Invent.BarcoObjIndex = 476 Then
                            ' Limitación por nivel, después dropea normalmente
                            If .Stats.ELV >= 20 And .Stats.ELV <= 25 Then
                                ' No dropea en agua
                                DropAgua = False
                            End If
                        End If
                    End If
                    
                    Call Tilelibre(.Pos, NuevaPos, MiObj, DropAgua, True)
                    
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                        Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.map, NuevaPos.X, NuevaPos.Y)
                    End If
                 End If
            End If
        Next i
    End With
End Sub

Function ItemNewbie(ByVal ItemIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If ItemIndex < 1 Or ItemIndex > UBound(ObjData) Then Exit Function
    
    ItemNewbie = ObjData(ItemIndex).Newbie = 1
End Function

Sub TirarTodosLosItemsNoNewbies(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 23/11/2009
'07/11/09: Pato - Fix bug #2819911
'23/11/2009: ZaMa - Optimizacion de codigo.
'***************************************************
    Dim i As Byte
    Dim NuevaPos As WorldPos
    Dim MiObj As Obj
    Dim ItemIndex As Integer
    
    With UserList(UserIndex)
        If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = 6 Then Exit Sub
        
        For i = 1 To UserList(UserIndex).CurrentInventorySlots
            ItemIndex = .Invent.Object(i).OBJIndex
            If ItemIndex > 0 Then
                If ItemSeCae(ItemIndex) And Not ItemNewbie(ItemIndex) Then
                    NuevaPos.X = 0
                    NuevaPos.Y = 0
                    
                    'Creo MiObj
                    MiObj.Amount = .Invent.Object(i).Amount
                    MiObj.OBJIndex = ItemIndex
                    'Pablo (ToxicWaste) 24/01/2007
                    'Tira los Items no newbies en todos lados.
                    Tilelibre .Pos, NuevaPos, MiObj, True, True
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                        Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.map, NuevaPos.X, NuevaPos.Y)
                    End If
                End If
            End If
        Next i
    End With

End Sub

Sub TirarTodosLosItemsEnMochila(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 12/01/09 (Budi)
'***************************************************
    Dim i As Byte
    Dim NuevaPos As WorldPos
    Dim MiObj As Obj
    Dim ItemIndex As Integer
    
    With UserList(UserIndex)
        If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = 6 Then Exit Sub
        
        For i = MAX_NORMAL_INVENTORY_SLOTS + 1 To .CurrentInventorySlots
            ItemIndex = .Invent.Object(i).OBJIndex
            If ItemIndex > 0 Then
                If ItemSeCae(ItemIndex) Then
                    NuevaPos.X = 0
                    NuevaPos.Y = 0
                    
                    'Creo MiObj
                    MiObj.Amount = .Invent.Object(i).Amount
                    MiObj.OBJIndex = ItemIndex
                    Tilelibre .Pos, NuevaPos, MiObj, True, True
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                        Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.map, NuevaPos.X, NuevaPos.Y)
                    End If
                End If
            End If
        Next i
    End With

End Sub

Public Function getObjType(ByVal OBJIndex As Integer) As eOBJType
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    If OBJIndex > 0 Then
        getObjType = ObjData(OBJIndex).OBJType
    End If
    
End Function

Function SkillPuedeUsarItem(UserIndex As Integer, ByVal OBJIndex As Integer, Motivo As String) As Boolean

With UserList(UserIndex)

    If .flags.Privilegios Then
        SkillPuedeUsarItem = True
        Exit Function
    End If
    
    If ObjData(OBJIndex).SkillCombate > .Stats.UserSkills(eSkill.Armas) Then
        Motivo = "No tienes suficiente habilidad en Combate."
        Exit Function
    End If
    
    If ObjData(OBJIndex).SkillApuñalar > .Stats.UserSkills(eSkill.Apuñalar) Then
        Motivo = "No tienes suficiente habilidad para Apuñalar."
        Exit Function
    End If
    
    If ObjData(OBJIndex).SkillProyectiles > .Stats.UserSkills(eSkill.Proyectiles) Then
        Motivo = "No tienes suficiente habilidad para usar armas de proyectiles."
        Exit Function
    End If
    
    If ObjData(OBJIndex).SkResistencia > .Stats.UserSkills(eSkill.Resis) Then
        Motivo = "No tienes la resistencia necesaria."
        Exit Function
    End If
    
    If ObjData(OBJIndex).SkDefensa > .Stats.UserSkills(eSkill.Defensa) Then
        Motivo = "Necesitas conocer mejor la defensa con escudos."
        Exit Function
    End If
    
    If ObjData(OBJIndex).SkillTacticas > .Stats.UserSkills(eSkill.Tacticas) Then
        Motivo = "Debes mejorar tus tacticas en combate."
        Exit Function
    End If

End With

SkillPuedeUsarItem = True

End Function
