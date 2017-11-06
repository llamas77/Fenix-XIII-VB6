Attribute VB_Name = "Trabajo"
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

Private Const GASTO_ENERGIA_TRABAJADOR As Byte = 2
Private Const GASTO_ENERGIA_NO_TRABAJADOR As Byte = 6


Public Sub DoPermanecerOculto(ByVal UserIndex As Integer)
'********************************************************
'Autor: Nacho (Integer)
'Last Modif: 11/19/2009
'Chequea si ya debe mostrarse
'Pablo (ToxicWaste): Cambie los ordenes de prioridades porque sino no andaba.
'11/19/2009: Pato - Ahora el bandido se oculta la mitad del tiempo de las demás clases.
'13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
'13/01/2010: ZaMa - Arreglo condicional para que el bandido camine oculto.
'********************************************************
On Error GoTo ErrHandler
    With UserList(UserIndex)
        .Counters.TiempoOculto = .Counters.TiempoOculto - 1
        If .Counters.TiempoOculto <= 0 Then
            
            If .Clase = eClass.Bandido Then
                .Counters.TiempoOculto = Int(IntervaloOculto / 2)
            Else
                .Counters.TiempoOculto = IntervaloOculto
            End If
            
            If .Clase = eClass.Cazador And .Stats.UserSkills(eSkill.Ocultarse) > 90 Then
                If .Invent.ArmourEqpObjIndex = 648 Or .Invent.ArmourEqpObjIndex = 360 Then
                    Exit Sub
                End If
            End If
            .Counters.TiempoOculto = 0
            .flags.Oculto = 0
            
            If .flags.Navegando = 1 Then
                If .Clase = eClass.Pirata Then
                    ' Pierde la apariencia de fragata fantasmal
                    Call ToogleBoatBody(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "¡Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, NingunArma, _
                                        NingunEscudo, NingunCasco)
                End If
            Else
                If .flags.invisible = 0 Then
                    Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                    Call SetInvisible(UserIndex, .Char.CharIndex, False)
                End If
            End If
        End If
    End With
    
    Exit Sub

ErrHandler:
    Call LogError("Error en Sub DoPermanecerOculto")


End Sub

Public Sub DoOcultarse(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 13/01/2010 (ZaMa)
'Pablo (ToxicWaste): No olvidar agregar IntervaloOculto=500 al Server.ini.
'Modifique la fórmula y ahora anda bien.
'13/01/2010: ZaMa - El pirata se transforma en galeon fantasmal cuando se oculta en agua.
'***************************************************

On Error GoTo ErrHandler

    Dim Suerte As Double
    Dim res As Integer
    Dim Skill As Integer
    
    With UserList(UserIndex)
        Skill = .Stats.UserSkills(eSkill.Ocultarse)
        
        Suerte = (((0.000002 * Skill - 0.0002) * Skill + 0.0064) * Skill + 0.1124) * 100
        
        res = RandomNumber(1, 100)
        
        If res <= Suerte Then
        
            .flags.Oculto = 1
            Suerte = (-0.000001 * (100 - Skill) ^ 3)
            Suerte = Suerte + (0.00009229 * (100 - Skill) ^ 2)
            Suerte = Suerte + (-0.0088 * (100 - Skill))
            Suerte = Suerte + (0.9571)
            Suerte = Suerte * IntervaloOculto
            .Counters.TiempoOculto = Suerte
            
            ' No es pirata o es uno sin barca
            If .flags.Navegando = 0 Then
                Call SetInvisible(UserIndex, .Char.CharIndex, True)
        
                Call WriteConsoleMsg(UserIndex, "¡Te has escondido entre las sombras!", FontTypeNames.FONTTYPE_INFO)
            ' Es un pirata navegando
            Else
                ' Le cambiamos el body a galeon fantasmal
                .Char.body = iFragataFantasmal
                ' Actualizamos clientes
                Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, NingunArma, _
                                    NingunEscudo, NingunCasco)
            End If
            
            Call SubirSkill(UserIndex, eSkill.Ocultarse, 15)
        Else
            '[CDT 17-02-2004]
            If Not .flags.UltimoMensaje = 4 Then
                Call WriteConsoleMsg(UserIndex, "¡No has logrado esconderte!", FontTypeNames.FONTTYPE_INFO)
                .flags.UltimoMensaje = 4
            End If
            '[/CDT]
            
            Call SubirSkill(UserIndex, eSkill.Ocultarse, 5)
        End If
        
        .Counters.Ocultando = .Counters.Ocultando + 1
    End With
    
    Exit Sub

ErrHandler:
    Call LogError("Error en Sub DoOcultarse")

End Sub

Public Sub DoNavega(ByVal UserIndex As Integer, ByRef Barco As ObjData, ByVal Slot As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 13/01/2010 (ZaMa)
'13/01/2010: ZaMa - El pirata pierde el ocultar si desequipa barca.
'***************************************************

    Dim ModNave As Single
    
    With UserList(UserIndex)
        ModNave = ModNavegacion(.Clase, UserIndex)
        
        If .Stats.UserSkills(eSkill.Navegacion) / ModNave < Barco.MinSkill Then
            Call WriteConsoleMsg(UserIndex, "No tienes suficientes conocimientos para usar este barco.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(UserIndex, "Para usar este barco necesitas " & Barco.MinSkill * ModNave & " puntos en navegacion.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        .Invent.BarcoObjIndex = .Invent.Object(Slot).OBJIndex
        .Invent.BarcoSlot = Slot
        
        ' No estaba navegando
        If .flags.Navegando = 0 Then
            
            .Char.Head = 0
            
            ' No esta muerto
            If .flags.Muerto = 0 Then
            
                Call ToogleBoatBody(UserIndex)
                
                If .Clase = eClass.Pirata Then
                    If .flags.Oculto = 1 Then
                        .flags.Oculto = 0
                        Call SetInvisible(UserIndex, .Char.CharIndex, False)
                        Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
                
            ' Esta muerto
            Else
                .Char.body = iFragataFantasmal
                .Char.ShieldAnim = NingunEscudo
                .Char.WeaponAnim = NingunArma
                .Char.CascoAnim = NingunCasco
            End If
            
            ' Comienza a navegar
            .flags.Navegando = 1
        
        ' Estaba navegando
        Else
            ' No esta muerto
            If .flags.Muerto = 0 Then
                .Char.Head = .OrigChar.Head
                
                If .Clase = eClass.Pirata Then
                    If .flags.Oculto = 1 Then
                        ' Al desequipar barca, perdió el ocultar
                        .flags.Oculto = 0
                        .Counters.Ocultando = 0
                        Call WriteConsoleMsg(UserIndex, "¡Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
                
                If .Invent.ArmourEqpObjIndex > 0 Then
                    .Char.body = ObjData(.Invent.ArmourEqpObjIndex).Ropaje
                Else
                    Call DarCuerpoDesnudo(UserIndex)
                End If
                
                If .Invent.EscudoEqpObjIndex > 0 Then _
                    .Char.ShieldAnim = ObjData(.Invent.EscudoEqpObjIndex).ShieldAnim
                If .Invent.WeaponEqpObjIndex > 0 Then _
                    .Char.WeaponAnim = ObjData(.Invent.WeaponEqpObjIndex).WeaponAnim
                If .Invent.CascoEqpObjIndex > 0 Then _
                    .Char.CascoAnim = ObjData(.Invent.CascoEqpObjIndex).CascoAnim
                    
            ' Esta muerto
            Else
                .Char.body = iCuerpoMuerto
                .Char.Head = iCabezaMuerto
                .Char.ShieldAnim = NingunEscudo
                .Char.WeaponAnim = NingunArma
                .Char.CascoAnim = NingunCasco
            End If
            
            ' Termina de navegar
            .flags.Navegando = 0
        End If
        
        ' Actualizo clientes
        Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
    End With
    
    Call WriteNavigateToggle(UserIndex)

End Sub

Public Sub FundirMineral(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

    With UserList(UserIndex)
        If .flags.TargetObjInvIndex > 0 Then
           
           If ObjData(.flags.TargetObjInvIndex).OBJType = eOBJType.otMinerales And _
                ObjData(.flags.TargetObjInvIndex).MinSkill <= .Stats.UserSkills(eSkill.Mineria) / ModFundicion(.Clase) Then
                Call DoLingotes(UserIndex)
           Else
                Call WriteConsoleMsg(UserIndex, "No tienes conocimientos de minería suficientes para trabajar este mineral.", FontTypeNames.FONTTYPE_INFO)
           End If
        
        End If
    End With

    Exit Sub

ErrHandler:
    Call LogError("Error en FundirMineral. Error " & Err.Number & " : " & Err.description)

End Sub

Public Sub FundirArmas(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler
    With UserList(UserIndex)
        If .flags.TargetObjInvIndex > 0 Then
            If ObjData(.flags.TargetObjInvIndex).OBJType = eOBJType.otWeapon Then
                If ObjData(.flags.TargetObjInvIndex).SkHerreria <= .Stats.UserSkills(eSkill.Herreria) / ModHerreria(.Clase) Then
                    Call DoFundir(UserIndex)
                Else
                    Call WriteConsoleMsg(UserIndex, "No tienes los conocimientos suficientes en herrería para fundir este objeto.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
    End With
    
    Exit Sub
ErrHandler:
    Call LogError("Error en FundirArmas. Error " & Err.Number & " : " & Err.description)
End Sub

Function TieneObjetos(ByVal ItemIndex As Integer, ByVal cant As Integer, ByVal UserIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim i As Integer
    Dim Total As Long
    For i = 1 To UserList(UserIndex).CurrentInventorySlots
        If UserList(UserIndex).Invent.Object(i).OBJIndex = ItemIndex Then
            Total = Total + UserList(UserIndex).Invent.Object(i).Amount
        End If
    Next i
    
    If cant <= Total Then
        TieneObjetos = True
        Exit Function
    End If
        
End Function

Public Sub QuitarObjetos(ByVal ItemIndex As Integer, ByVal cant As Integer, ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 05/08/09
'05/08/09: Pato - Cambie la funcion a procedimiento ya que se usa como procedimiento siempre, y fixie el bug 2788199
'***************************************************

    Dim i As Integer
    For i = 1 To UserList(UserIndex).CurrentInventorySlots
        With UserList(UserIndex).Invent.Object(i)
            If .OBJIndex = ItemIndex Then
                If .Amount <= cant And .Equipped = 1 Then Call Desequipar(UserIndex, i)
                
                .Amount = .Amount - cant
                If .Amount <= 0 Then
                    cant = Abs(.Amount)
                    UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
                    .Amount = 0
                    .OBJIndex = 0
                Else
                    cant = 0
                End If
                
                Call UpdateUserInv(False, UserIndex, i)
                
                If cant = 0 Then Exit Sub
            End If
        End With
    Next i

End Sub

Sub HerreroQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal CantidadItems As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 16/11/2009
'16/11/2009: ZaMa - Ahora considera la cantidad de items a construir
'***************************************************

Dim Descuento As Single
    
    With ObjData(ItemIndex)
        If UserList(UserIndex).Clase = eClass.Herrero Then
            If UserList(UserIndex).Recompensas(1) = 1 And .OBJType <> eOBJType.otCasco And .OBJType <> eOBJType.otEscudo Then
                If CInt(RandomNumber(1, 4)) <= 1 Then Descuento = 0.5
            ElseIf UserList(UserIndex).Recompensas(1) = 2 And .OBJType = eOBJType.otCasco Or .OBJType = eOBJType.otEscudo Then
                Descuento = 0.5
            End If
            
            Descuento = Descuento * (1 - 0.25 * Buleano(UserList(UserIndex).Recompensas(3) = 1 And .OBJType <> eOBJType.otCasco And .OBJType <> eOBJType.otEscudo))
        End If
    
        If .LingH > 0 Then Call QuitarObjetos(LingoteHierro, Descuento * .LingH * CantidadItems, UserIndex)
        If .LingP > 0 Then Call QuitarObjetos(LingotePlata, Descuento * .LingP * CantidadItems, UserIndex)
        If .LingO > 0 Then Call QuitarObjetos(LingoteOro, Descuento * .LingO * CantidadItems, UserIndex)
    End With
End Sub

Sub CarpinteroQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal CantidadItems As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 16/11/2009
'16/11/2009: ZaMa - Ahora quita tambien madera elfica
'***************************************************

Dim Descuento As Single

    With ObjData(ItemIndex)
        If UserList(UserIndex).Clase = eClass.Carpintero And _
            UserList(UserIndex).Recompensas(2) = 2 And _
                .OBJType = eOBJType.otBarcos Then
                    Descuento = 0.8
        Else
            Descuento = 1
        End If
        
        If .Madera > 0 Then Call QuitarObjetos(Leña, Descuento * .Madera * CantidadItems, UserIndex)
        If .MaderaElfica > 0 Then Call QuitarObjetos(LeñaElfica, Descuento * .MaderaElfica * CantidadItems, UserIndex)
    End With
End Sub

Function CarpinteroTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal Cantidad As Integer, Optional ByVal ShowMsg As Boolean = False) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 16/11/2009
'16/11/2009: ZaMa - Agregada validacion a madera elfica.
'16/11/2009: ZaMa - Ahora considera la cantidad de items a construir
'***************************************************

Dim Descuento As Single

    With ObjData(ItemIndex)
        If UserList(UserIndex).Clase = eClass.Carpintero And _
            UserList(UserIndex).Recompensas(2) = 2 And _
                .OBJType = eOBJType.otBarcos Then
                    Descuento = 0.8
        Else
            Descuento = 1
        End If
        
        If .Madera > 0 Then
            If Not TieneObjetos(Leña, Descuento * .Madera * Cantidad, UserIndex) Then
                If ShowMsg Then Call WriteConsoleMsg(UserIndex, "No tienes suficiente madera.", FontTypeNames.FONTTYPE_INFO)
                CarpinteroTieneMateriales = False
                Exit Function
            End If
        End If
        
        If .MaderaElfica > 0 Then
            If Not TieneObjetos(LeñaElfica, Descuento * .MaderaElfica * Cantidad, UserIndex) Then
                If ShowMsg Then Call WriteConsoleMsg(UserIndex, "No tienes suficiente madera élfica.", FontTypeNames.FONTTYPE_INFO)
                CarpinteroTieneMateriales = False
                Exit Function
            End If
        End If
    
    End With
    CarpinteroTieneMateriales = True

End Function

Function Piel(ByVal UserIndex As Integer, ByVal Tipo As Byte, ByVal Obj As Integer) As Integer

    Select Case Tipo
        Case 1
            Piel = ObjData(Obj).PielLobo
            If UserList(UserIndex).Clase = eClass.Sastre And UserList(UserIndex).Stats.ELV >= 18 Then Piel = Piel * 0.8
        Case 2
            Piel = ObjData(Obj).PielOsoPardo
            If UserList(UserIndex).Clase = eClass.Sastre And UserList(UserIndex).Stats.ELV >= 18 Then Piel = Piel * 0.8
        Case 3
            Piel = ObjData(Obj).PielOsoPolar
            If UserList(UserIndex).Clase = eClass.Sastre And UserList(UserIndex).Stats.ELV >= 18 Then Piel = Piel * 0.8
    End Select

End Function
Function SastreTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal cantT As Integer) As Boolean
Dim PielL As Integer, PielO As Integer, PielP As Integer
cantT = MaximoInt(1, cantT)

PielL = ObjData(ItemIndex).PielLobo
PielO = ObjData(ItemIndex).PielOsoPardo
PielP = ObjData(ItemIndex).PielOsoPolar

If UserList(UserIndex).Clase = eClass.Sastre And UserList(UserIndex).Stats.ELV >= 18 Then
    PielL = 0.8 * PielL
    PielO = 0.8 * PielO
    PielP = 0.8 * PielP
End If

If PielL Then
    If Not TieneObjetos(PLobo, CInt(PielL * ModSastre(UserList(UserIndex).Clase)) * cantT, UserIndex) Then
        Call WriteConsoleMsg(UserIndex, "No tienes suficientes pieles.", FontTypeNames.FONTTYPE_INFO)
        SastreTieneMateriales = False
        Exit Function
    End If
End If

If PielO Then
    If Not TieneObjetos(POsoPardo, CInt(PielO * ModSastre(UserList(UserIndex).Clase)) * cantT, UserIndex) Then
        Call WriteConsoleMsg(UserIndex, "No tienes suficientes pieles.", FontTypeNames.FONTTYPE_INFO)
        SastreTieneMateriales = False
        Exit Function
    End If
End If
    
If PielP Then
    If Not TieneObjetos(POsoPolar, CInt(PielP * ModSastre(UserList(UserIndex).Clase)) * cantT, UserIndex) Then
        Call WriteConsoleMsg(UserIndex, "No tienes suficientes pieles.", FontTypeNames.FONTTYPE_INFO)
        SastreTieneMateriales = False
        Exit Function
    End If
End If
    
SastreTieneMateriales = True

End Function
Sub SastreQuitarMateriales(UserIndex As Integer, ItemIndex As Integer, cantT As Integer)
Dim PielL As Integer, PielO As Integer, PielP As Integer

PielL = ObjData(ItemIndex).PielLobo
PielO = ObjData(ItemIndex).PielOsoPardo
PielP = ObjData(ItemIndex).PielOsoPolar

If UserList(UserIndex).Clase = eClass.Sastre And UserList(UserIndex).Stats.ELV >= 18 Then
    PielL = 0.8 * PielL
    PielO = 0.8 * PielO
    PielP = 0.8 * PielP
End If

If PielL Then Call QuitarObjetos(PLobo, CInt(PielL * ModSastre(UserList(UserIndex).Clase)) * cantT, UserIndex)
If PielO Then Call QuitarObjetos(POsoPardo, CInt(PielO * ModSastre(UserList(UserIndex).Clase)) * cantT, UserIndex)
If PielP Then Call QuitarObjetos(POsoPolar, CInt(PielP * ModSastre(UserList(UserIndex).Clase)) * cantT, UserIndex)

End Sub
Public Sub SastreConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal cantT As Integer)

If SastreTieneMateriales(UserIndex, ItemIndex, cantT) And _
   UserList(UserIndex).Stats.UserSkills(Sastreria) / ModRopas(UserList(UserIndex).Clase) >= _
   ObjData(ItemIndex).SkSastreria And _
   PuedeConstruirSastre(ItemIndex, UserIndex) And _
   UserList(UserIndex).Invent.HerramientaEqpObjIndex = HILAR_SASTRE Then
        
    Call SastreQuitarMateriales(UserIndex, ItemIndex, cantT)
    Call WriteConsoleMsg(UserIndex, "¡Has creado un ropaje!", FontTypeNames.FONTTYPE_INFO)
    
    Dim MiObj As Obj
    MiObj.Amount = MaximoInt(1, cantT)
    MiObj.OBJIndex = ItemIndex
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    
    Call CheckUserLevel(UserIndex)

    Call SubirSkill(UserIndex, Sastreria, 5)

Else
    Call WriteConsoleMsg(UserIndex, "No has podido crear el ropaje.", FontTypeNames.FONTTYPE_INFO)

End If

End Sub

Public Function PuedeConstruirSastre(ItemIndex As Integer, UserIndex As Integer) As Boolean
Dim i As Long
Dim N As Integer

N = val(GetVar(DatPath & "ObjSastre.dat", "INIT", "NumObjs"))

For i = 1 To UBound(ObjSastre)
    If ObjSastre(i) = ItemIndex Then
        PuedeConstruirSastre = True
        Exit Function
    End If
Next

PuedeConstruirSastre = False

End Function

Function HerreroTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal CantidadItems As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 16/11/2009
'16/11/2009: ZaMa - Agregada validacion a madera elfica.
'***************************************************
Dim Descuento As Single

    With ObjData(ItemIndex)
        
        If UserList(UserIndex).Clase = eClass.Herrero Then
            If UserList(UserIndex).Recompensas(1) = 1 And .OBJType <> eOBJType.otCasco And .OBJType <> eOBJType.otEscudo Then
                If CInt(RandomNumber(1, 4)) <= 1 Then Descuento = 0.5
            ElseIf UserList(UserIndex).Recompensas(1) = 2 And .OBJType = eOBJType.otCasco Or .OBJType = eOBJType.otEscudo Then
                Descuento = 0.5
            End If
            
            Descuento = Descuento * (1 - 0.25 * Buleano(UserList(UserIndex).Recompensas(3) = 1 And .OBJType <> eOBJType.otCasco And .OBJType <> eOBJType.otEscudo))
        End If
        
        If .LingH > 0 Then
            If Not TieneObjetos(LingoteHierro, Descuento * .LingH * CantidadItems, UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "No tienes suficientes lingotes de hierro.", FontTypeNames.FONTTYPE_INFO)
                HerreroTieneMateriales = False
                Exit Function
            End If
        End If
        If .LingP > 0 Then
            If Not TieneObjetos(LingotePlata, Descuento * .LingP * CantidadItems, UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "No tienes suficientes lingotes de plata.", FontTypeNames.FONTTYPE_INFO)
                HerreroTieneMateriales = False
                Exit Function
            End If
        End If
        If .LingO > 0 Then
            If Not TieneObjetos(LingoteOro, Descuento * .LingO * CantidadItems, UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "No tienes suficientes lingotes de oro.", FontTypeNames.FONTTYPE_INFO)
                HerreroTieneMateriales = False
                Exit Function
            End If
        End If
    End With
    HerreroTieneMateriales = True
End Function

Public Function PuedeConstruir(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal CantidadItems As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 24/08/2009
'24/08/2008: ZaMa - Validates if the player has the required skill
'16/11/2009: ZaMa - Validates if the player has the required amount of materials, depending on the number of items to make
'***************************************************
PuedeConstruir = HerreroTieneMateriales(UserIndex, ItemIndex, CantidadItems) And _
                    Round(UserList(UserIndex).Stats.UserSkills(eSkill.Herreria) / ModHerreria(UserList(UserIndex).Clase), 0) >= ObjData(ItemIndex).SkHerreria
End Function

Public Function PuedeConstruirHerreria(ByVal ItemIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
Dim i As Long

For i = 1 To UBound(ArmasHerrero)
    If ArmasHerrero(i) = ItemIndex Then
        PuedeConstruirHerreria = True
        Exit Function
    End If
Next i
For i = 1 To UBound(ArmadurasHerrero)
    If ArmadurasHerrero(i) = ItemIndex Then
        PuedeConstruirHerreria = True
        Exit Function
    End If
Next i
PuedeConstruirHerreria = False
End Function

Public Sub HerreroConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 16/11/2009
'16/11/2009: ZaMa - Implementado nuevo sistema de construccion de items
'***************************************************
Dim CantidadItems As Integer
Dim TieneMateriales As Boolean

With UserList(UserIndex)
    CantidadItems = .Construir.PorCiclo
    
    If .Construir.Cantidad < CantidadItems Then _
        CantidadItems = .Construir.Cantidad
        
    If .Construir.Cantidad > 0 Then _
        .Construir.Cantidad = .Construir.Cantidad - CantidadItems
        
    If CantidadItems = 0 Then
        Call WriteStopWorking(UserIndex)
        Exit Sub
    End If
    
    If PuedeConstruirHerreria(ItemIndex) Then
        
        While CantidadItems > 0 And Not TieneMateriales
            If PuedeConstruir(UserIndex, ItemIndex, CantidadItems) Then
                TieneMateriales = True
            Else
                CantidadItems = CantidadItems - 1
            End If
        Wend
        
        ' Chequeo si puede hacer al menos 1 item
        If Not TieneMateriales Then
            Call WriteConsoleMsg(UserIndex, "No tienes suficientes materiales.", FontTypeNames.FONTTYPE_INFO)
            Call WriteStopWorking(UserIndex)
            Exit Sub
        End If
        
        'Sacamos energía
        If esTrabajador(.Clase) Then
            'Chequeamos que tenga los puntos antes de sacarselos
            If .Stats.MinSta >= GASTO_ENERGIA_TRABAJADOR Then
                .Stats.MinSta = .Stats.MinSta - GASTO_ENERGIA_TRABAJADOR
                Call WriteUpdateSta(UserIndex)
            Else
                Call WriteConsoleMsg(UserIndex, "No tienes suficiente energía.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        Else
            'Chequeamos que tenga los puntos antes de sacarselos
            If .Stats.MinSta >= GASTO_ENERGIA_NO_TRABAJADOR Then
                .Stats.MinSta = .Stats.MinSta - GASTO_ENERGIA_NO_TRABAJADOR
                Call WriteUpdateSta(UserIndex)
            Else
                Call WriteConsoleMsg(UserIndex, "No tienes suficiente energía.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        
        Call HerreroQuitarMateriales(UserIndex, ItemIndex, CantidadItems)
        ' AGREGAR FX
        If ObjData(ItemIndex).OBJType = eOBJType.otWeapon Then
            Call WriteConsoleMsg(UserIndex, "Has construido " & IIf(CantidadItems > 1, CantidadItems & " armas!", "el arma!"), FontTypeNames.FONTTYPE_INFO)
        ElseIf ObjData(ItemIndex).OBJType = eOBJType.otEscudo Then
            Call WriteConsoleMsg(UserIndex, "Has construido " & IIf(CantidadItems > 1, CantidadItems & " escudos!", "el escudo!"), FontTypeNames.FONTTYPE_INFO)
        ElseIf ObjData(ItemIndex).OBJType = eOBJType.otCasco Then
            Call WriteConsoleMsg(UserIndex, "Has construido " & IIf(CantidadItems > 1, CantidadItems & " cascos!", "el casco!"), FontTypeNames.FONTTYPE_INFO)
        ElseIf ObjData(ItemIndex).OBJType = eOBJType.otArmadura Then
            Call WriteConsoleMsg(UserIndex, "Has construido " & IIf(CantidadItems > 1, CantidadItems & " armaduras", "la armadura!"), FontTypeNames.FONTTYPE_INFO)
        End If
        
        Dim MiObj As Obj
        
        CantidadItems = CantidadItems * (1 + Buleano(RandomNumber(1, 10) = 1 And .Clase = eClass.Herrero And .Recompensas(3) = 2))
        MiObj.Amount = CantidadItems
        MiObj.OBJIndex = ItemIndex
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(.Pos, MiObj)
        End If
        
        Call SubirSkill(UserIndex, eSkill.Herreria, 5)
        Call UpdateUserInv(True, UserIndex, 0)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(MARTILLOHERRERO, .Pos.X, .Pos.Y))

        .Counters.Trabajando = .Counters.Trabajando + 1
    End If
End With
End Sub

Public Function PuedeConstruirCarpintero(ByVal ItemIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
Dim i As Long

For i = 1 To UBound(ObjCarpintero)
    If ObjCarpintero(i) = ItemIndex Then
        PuedeConstruirCarpintero = True
        Exit Function
    End If
Next i
PuedeConstruirCarpintero = False

End Function

Public Sub CarpinteroConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 16/11/2009
'24/08/2008: ZaMa - Validates if the player has the required skill
'16/11/2009: ZaMa - Implementado nuevo sistema de construccion de items
'***************************************************
Dim CantidadItems As Integer
Dim TieneMateriales As Boolean

With UserList(UserIndex)
    CantidadItems = .Construir.PorCiclo
    
    If .Construir.Cantidad < CantidadItems Then _
        CantidadItems = .Construir.Cantidad
        
    If .Construir.Cantidad > 0 Then _
        .Construir.Cantidad = .Construir.Cantidad - CantidadItems
        
    If CantidadItems = 0 Then
        Call WriteStopWorking(UserIndex)
        Exit Sub
    End If

    If Round(.Stats.UserSkills(eSkill.Carpinteria) \ ModCarpinteria(.Clase), 0) >= _
       ObjData(ItemIndex).SkCarpinteria And _
       PuedeConstruirCarpintero(ItemIndex) And _
       .Invent.WeaponEqpObjIndex = SERRUCHO_CARPINTERO Then
       
        ' Calculo cuantos item puede construir
        While CantidadItems > 0 And Not TieneMateriales
            If CarpinteroTieneMateriales(UserIndex, ItemIndex, CantidadItems) Then
                TieneMateriales = True
            Else
                CantidadItems = CantidadItems - 1
            End If
        Wend
        
        ' No tiene los materiales ni para construir 1 item?
        If Not TieneMateriales Then
            ' Para que muestre el mensaje
            Call CarpinteroTieneMateriales(UserIndex, ItemIndex, 1, True)
            Call WriteStopWorking(UserIndex)
            Exit Sub
        End If
       
        'Sacamos energía
        If esTrabajador(.Clase) Then
            'Chequeamos que tenga los puntos antes de sacarselos
            If .Stats.MinSta >= GASTO_ENERGIA_TRABAJADOR Then
                .Stats.MinSta = .Stats.MinSta - GASTO_ENERGIA_TRABAJADOR
                Call WriteUpdateSta(UserIndex)
            Else
                Call WriteConsoleMsg(UserIndex, "No tienes suficiente energía.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        Else
            'Chequeamos que tenga los puntos antes de sacarselos
            If .Stats.MinSta >= GASTO_ENERGIA_NO_TRABAJADOR Then
                .Stats.MinSta = .Stats.MinSta - GASTO_ENERGIA_NO_TRABAJADOR
                Call WriteUpdateSta(UserIndex)
            Else
                Call WriteConsoleMsg(UserIndex, "No tienes suficiente energía.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        
        Call CarpinteroQuitarMateriales(UserIndex, ItemIndex, CantidadItems)
        Call WriteConsoleMsg(UserIndex, "Has construido " & CantidadItems & _
                            IIf(CantidadItems = 1, " objeto!", " objetos!"), FontTypeNames.FONTTYPE_INFO)
        
        Dim MiObj As Obj
        MiObj.Amount = CantidadItems
        MiObj.OBJIndex = ItemIndex
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(.Pos, MiObj)
        End If
        
        
        Call SubirSkill(UserIndex, eSkill.Carpinteria, 5)
        Call UpdateUserInv(True, UserIndex, 0)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(LABUROCARPINTERO, .Pos.X, .Pos.Y))
    
        .Counters.Trabajando = .Counters.Trabajando + 1
    
    ElseIf .Invent.WeaponEqpObjIndex <> SERRUCHO_CARPINTERO Then
        Call WriteConsoleMsg(UserIndex, "Debes tener equipado el serrucho para trabajar.", FontTypeNames.FONTTYPE_INFO)
    End If
End With
End Sub

Private Function MineralesParaLingote(ByVal Lingote As iMinerales) As Integer
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
    Select Case Lingote
        Case iMinerales.HierroCrudo
            MineralesParaLingote = 6
        Case iMinerales.PlataCruda
            MineralesParaLingote = 18
        Case iMinerales.OroCrudo
            MineralesParaLingote = 34
        Case Else
            MineralesParaLingote = 10000
    End Select
End Function


Public Sub DoLingotes(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 16/11/2009
'16/11/2009: ZaMa - Implementado nuevo sistema de construccion de items
'***************************************************
'    Call LogTarea("Sub DoLingotes")
    Dim Slot As Integer
    Dim obji As Integer
    Dim CantidadItems As Integer
    Dim TieneMinerales As Boolean

    With UserList(UserIndex)
        CantidadItems = MaximoInt(1, CInt((.Stats.ELV - 4) / 5))

        Slot = .flags.TargetObjInvSlot
        obji = .Invent.Object(Slot).OBJIndex
        
        While CantidadItems > 0 And Not TieneMinerales
            If .Invent.Object(Slot).Amount >= MineralesParaLingote(obji) * CantidadItems Then
                TieneMinerales = True
            Else
                CantidadItems = CantidadItems - 1
            End If
        Wend
        
        If Not TieneMinerales Or ObjData(obji).OBJType <> eOBJType.otMinerales Then
            Call WriteConsoleMsg(UserIndex, "No tienes suficientes minerales para hacer un lingote.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount - MineralesParaLingote(obji) * CantidadItems
        If .Invent.Object(Slot).Amount < 1 Then
            .Invent.Object(Slot).Amount = 0
            .Invent.Object(Slot).OBJIndex = 0
        End If
        
        Dim MiObj As Obj
        MiObj.Amount = CantidadItems
        MiObj.OBJIndex = ObjData(.flags.TargetObjInvIndex).LingoteIndex
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(.Pos, MiObj)
        End If
        Call UpdateUserInv(False, UserIndex, Slot)
        Call WriteConsoleMsg(UserIndex, "¡Has obtenido " & CantidadItems & " lingote" & _
                            IIf(CantidadItems = 1, "", "s") & "!", FontTypeNames.FONTTYPE_INFO)
    
        .Counters.Trabajando = .Counters.Trabajando + 1
    End With
End Sub

Public Sub DoFundir(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 03/06/2010
'03/06/2010 - Pato: Si es el último ítem a fundir y está equipado lo desequipamos.
'11/03/2010 - ZaMa: Reemplazo división por producto para uan mejor performanse.
'***************************************************
Dim i As Integer
Dim Num As Integer
Dim Slot As Byte
Dim Lingotes(2) As Integer

    With UserList(UserIndex)
        Slot = .flags.TargetObjInvSlot
        
        With .Invent.Object(Slot)
            .Amount = .Amount - 1
            
            If .Amount < 1 Then
                If .Equipped = 1 Then Call Desequipar(UserIndex, Slot)
                
                .Amount = 0
                .OBJIndex = 0
            End If
        End With
        
        Num = RandomNumber(10, 25)
        
        Lingotes(0) = (ObjData(.flags.TargetObjInvIndex).LingH * Num) * 0.01
        Lingotes(1) = (ObjData(.flags.TargetObjInvIndex).LingP * Num) * 0.01
        Lingotes(2) = (ObjData(.flags.TargetObjInvIndex).LingO * Num) * 0.01
    
    Dim MiObj(2) As Obj
    
    For i = 0 To 2
        MiObj(i).Amount = Lingotes(i)
        MiObj(i).OBJIndex = LingoteHierro + i 'Una gran negrada pero práctica
        If MiObj(i).Amount > 0 Then
            If Not MeterItemEnInventario(UserIndex, MiObj(i)) Then
                Call TirarItemAlPiso(.Pos, MiObj(i))
            End If
            Call UpdateUserInv(True, UserIndex, Slot)
        End If
    Next i
    
    Call WriteConsoleMsg(UserIndex, "¡Has obtenido el " & Num & "% de los lingotes utilizados para la construcción del objeto!", FontTypeNames.FONTTYPE_INFO)

    .Counters.Trabajando = .Counters.Trabajando + 1

End With

End Sub

Function ModNavegacion(ByVal Clase As eClass, ByVal UserIndex As Integer) As Single
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 27/11/2009
'27/11/2009: ZaMa - A worker can navigate before only if it's an expert fisher
'12/04/2010: ZaMa - Arreglo modificador de pescador, para que navegue con 60 skills.
'***************************************************
Select Case Clase
    Case eClass.Pirata
        ModNavegacion = 1
    Case Else
        ModNavegacion = 2
End Select

If esTrabajador(Clase) Then
    If UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) = 100 Then
        ModNavegacion = 1.71
    Else
        ModNavegacion = 2
    End If
End If

End Function


Function ModFundicion(ByVal Clase As eClass) As Single
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

If esTrabajador(Clase) Then
    ModFundicion = 1
Else
    ModFundicion = 3
End If

End Function

Function ModCarpinteria(ByVal Clase As eClass) As Integer
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

If esTrabajador(Clase) Then
    ModCarpinteria = 1
Else
    ModCarpinteria = 3
End If


End Function

Function ModSastre(ByVal Clase As eClass) As Double

Select Case (Clase)
    Case eClass.Sastre
        ModSastre = 1
    Case Else
        ModSastre = 3
End Select

End Function

Function ModRopas(Clase As eClass) As Double

Select Case (Clase)
    Case eClass.Sastre
        ModRopas = 1
    Case Else
        ModRopas = 3
End Select

End Function

Function ModHerreria(ByVal Clase As eClass) As Single
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
If esTrabajador(Clase) Then
    ModHerreria = 1
Else
    ModHerreria = 4
End If

End Function

Function ModDomar(ByVal Clase As eClass) As Integer
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
    Select Case Clase
        Case eClass.Druida
            ModDomar = 6
        Case eClass.Cazador
            ModDomar = 6
        Case eClass.Clerigo
            ModDomar = 7
        Case Else
            ModDomar = 10
    End Select
End Function

Function FreeMascotaIndex(ByVal UserIndex As Integer) As Integer
'***************************************************
'Author: Unknown
'Last Modification: 02/03/09
'02/03/09: ZaMa - Busca un indice libre de mascotas, revisando los types y no los indices de los npcs
'***************************************************
    Dim j As Integer
    For j = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasType(j) = 0 Then
            FreeMascotaIndex = j
            Exit Function
        End If
    Next j
End Function

Sub DoDomar(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
'***************************************************
'Author: Nacho (Integer)
'Last Modification: 01/05/2010
'12/15/2008: ZaMa - Limits the number of the same type of pet to 2.
'02/03/2009: ZaMa - Las criaturas domadas en zona segura, esperan afuera (desaparecen).
'01/05/2010: ZaMa - Agrego bonificacion 11% para domar con flauta magica.
'***************************************************

On Error GoTo ErrHandler

    Dim puntosDomar As Integer
    Dim puntosRequeridos As Integer
    Dim CanStay As Boolean
    Dim petType As Integer
    Dim NroPets As Integer
    
    
    If Npclist(NpcIndex).MaestroUser = UserIndex Then
        Call WriteConsoleMsg(UserIndex, "Ya domaste a esa criatura.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If

    With UserList(UserIndex)
        If .NroMascotas < MAXMASCOTAS Then
            
            If Npclist(NpcIndex).MaestroNpc > 0 Or Npclist(NpcIndex).MaestroUser > 0 Then
                Call WriteConsoleMsg(UserIndex, "La criatura ya tiene amo.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            If Not PuedeDomarMascota(UserIndex, NpcIndex) Then
                Call WriteConsoleMsg(UserIndex, "No puedes domar más de dos criaturas del mismo tipo.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            puntosDomar = CInt(.Stats.UserAtributos(eAtributos.Carisma)) * CInt(.Stats.UserSkills(eSkill.Domar))
            
            puntosRequeridos = Npclist(NpcIndex).flags.Domable
            
            If puntosRequeridos <= puntosDomar And RandomNumber(1, 5) = 1 Then
                Dim index As Integer
                .NroMascotas = .NroMascotas + 1
                index = FreeMascotaIndex(UserIndex)
                .MascotasIndex(index) = NpcIndex
                .MascotasType(index) = Npclist(NpcIndex).Numero
                
                Npclist(NpcIndex).MaestroUser = UserIndex
                
                Call FollowAmo(NpcIndex)
                Call ReSpawnNpc(Npclist(NpcIndex))
                
                Call WriteConsoleMsg(UserIndex, "La criatura te ha aceptado como su amo.", FontTypeNames.FONTTYPE_INFO)
                
                ' Es zona segura?
                CanStay = (MapInfo(.Pos.map).Pk = True)
                
                If Not CanStay Then
                    petType = Npclist(NpcIndex).Numero
                    NroPets = .NroMascotas
                    
                    Call QuitarNPC(NpcIndex)
                    
                    .MascotasType(index) = petType
                    .NroMascotas = NroPets
                    
                    Call WriteConsoleMsg(UserIndex, "No se permiten mascotas en zona segura. Éstas te esperarán afuera.", FontTypeNames.FONTTYPE_INFO)
                End If
                
                Call SubirSkill(UserIndex, eSkill.Domar, 15)
        
            Else
                If Not .flags.UltimoMensaje = 5 Then
                    Call WriteConsoleMsg(UserIndex, "No has logrado domar la criatura.", FontTypeNames.FONTTYPE_INFO)
                    .flags.UltimoMensaje = 5
                End If
                
                Call SubirSkill(UserIndex, eSkill.Domar, 5)
            End If
        Else
            Call WriteConsoleMsg(UserIndex, "No puedes controlar más criaturas.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    
    Exit Sub

ErrHandler:
    Call LogError("Error en DoDomar. Error " & Err.Number & " : " & Err.description)

End Sub

''
' Checks if the user can tames a pet.
'
' @param integer userIndex The user id from who wants tame the pet.
' @param integer NPCindex The index of the npc to tome.
' @return boolean True if can, false if not.
Private Function PuedeDomarMascota(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Boolean
'***************************************************
'Author: ZaMa
'This function checks how many NPCs of the same type have
'been tamed by the user.
'Returns True if that amount is less than two.
'***************************************************
    Dim i As Long
    Dim numMascotas As Long
    
    For i = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasType(i) = Npclist(NpcIndex).Numero Then
            numMascotas = numMascotas + 1
        End If
    Next i
    
    If numMascotas <= 1 Then PuedeDomarMascota = True
    
End Function

Sub DoAdminInvisible(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 12/01/2010 (ZaMa)
'Makes an admin invisible o visible.
'13/07/2009: ZaMa - Now invisible admins' chars are erased from all clients, except from themselves.
'12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados cuando pierden el efecto del mimetismo.
'***************************************************
    
    With UserList(UserIndex)
        If .flags.AdminInvisible = 0 Then
            ' Sacamos el mimetizmo
            If .flags.Mimetizado = 1 Then
                .Char.body = .CharMimetizado.body
                .Char.Head = .CharMimetizado.Head
                .Char.CascoAnim = .CharMimetizado.CascoAnim
                .Char.ShieldAnim = .CharMimetizado.ShieldAnim
                .Char.WeaponAnim = .CharMimetizado.WeaponAnim
                .Counters.Mimetismo = 0
                .flags.Mimetizado = 0
                ' Se fue el efecto del mimetismo, puede ser atacado por npcs
                .flags.Ignorado = False
            End If
            
            .flags.AdminInvisible = 1
            .flags.invisible = 1
            .flags.Oculto = 1
            .flags.OldBody = .Char.body
            .flags.OldHead = .Char.Head
            .Char.body = 0
            .Char.Head = 0
            
            ' Solo el admin sabe que se hace invi
            Call EnviarDatosASlot(UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, True))
            'Le mandamos el mensaje para que borre el personaje a los clientes que estén cerca
            Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterRemove(.Char.CharIndex))
        Else
            .flags.AdminInvisible = 0
            .flags.invisible = 0
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            .Char.body = .flags.OldBody
            .Char.Head = .flags.OldHead
            
            ' Solo el admin sabe que se hace visible
            Call EnviarDatosASlot(UserIndex, PrepareMessageCharacterChange(.Char.body, .Char.Head, .Char.heading, _
            .Char.CharIndex, .Char.WeaponAnim, .Char.ShieldAnim, .Char.FX, .Char.loops, .Char.CascoAnim))
            Call EnviarDatosASlot(UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
             
            'Le mandamos el mensaje para crear el personaje a los clientes que estén cerca
            Call MakeUserChar(True, .Pos.map, UserIndex, .Pos.map, .Pos.X, .Pos.Y, True)
        End If
    End With
    
End Sub

Sub TratarDeHacerFogata(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim Suerte As Byte
Dim exito As Byte
Dim Obj As Obj
Dim posMadera As WorldPos

If Not LegalPos(map, X, Y) Then Exit Sub

With posMadera
    .map = map
    .X = X
    .Y = Y
End With

If MapData(map, X, Y).ObjInfo.OBJIndex <> 58 Then
    Call WriteConsoleMsg(UserIndex, "Necesitas clickear sobre leña para hacer ramitas.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If Distancia(posMadera, UserList(UserIndex).Pos) > 2 Then
    Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos para prender la fogata.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If UserList(UserIndex).flags.Muerto = 1 Then
    Call WriteConsoleMsg(UserIndex, "No puedes hacer fogatas estando muerto.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If MapData(map, X, Y).ObjInfo.Amount < 3 Then
    Call WriteConsoleMsg(UserIndex, "Necesitas por lo menos tres troncos para hacer una fogata.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

Dim SupervivenciaSkill As Byte

SupervivenciaSkill = UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia)

If SupervivenciaSkill >= 0 And SupervivenciaSkill < 6 Then
    Suerte = 3
ElseIf SupervivenciaSkill >= 6 And SupervivenciaSkill <= 34 Then
    Suerte = 2
ElseIf SupervivenciaSkill >= 35 Then
    Suerte = 1
End If

exito = RandomNumber(1, Suerte)

If exito = 1 Then
    Obj.OBJIndex = FOGATA_APAG
    Obj.Amount = MapData(map, X, Y).ObjInfo.Amount \ 3
    
    Call WriteConsoleMsg(UserIndex, "Has hecho " & Obj.Amount & " fogatas.", FontTypeNames.FONTTYPE_INFO)
    
    Call MakeObj(Obj, map, X, Y)
    
    'Seteamos la fogata como el nuevo TargetObj del user
    UserList(UserIndex).flags.TargetObj = FOGATA_APAG
    
    Call SubirSkill(UserIndex, eSkill.Supervivencia)
Else
    '[CDT 17-02-2004]
    If Not UserList(UserIndex).flags.UltimoMensaje = 10 Then
        Call WriteConsoleMsg(UserIndex, "No has podido hacer la fogata.", FontTypeNames.FONTTYPE_INFO)
        UserList(UserIndex).flags.UltimoMensaje = 10
    End If
    '[/CDT]
    
    'Call SubirSkill(UserIndex, eSkill.Supervivencia, False)
End If

End Sub

Public Sub DoPescar(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 16/11/2009
'16/11/2009: ZaMa - Implementado nuevo sistema de extraccion.
'***************************************************
On Error GoTo ErrHandler

Dim Suerte As Integer
Dim res As Integer
Dim CantidadItems As Integer

If esTrabajador(UserList(UserIndex).Clase) Then
    Call QuitarSta(UserIndex, EsfuerzoPescarPescador)
Else
    Call QuitarSta(UserIndex, EsfuerzoPescarGeneral)
End If

Dim Skill As Integer
Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Pesca)
Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)

res = RandomNumber(1, Suerte)

If res <= 6 Then
    Dim MiObj As Obj
    
    If esTrabajador(UserList(UserIndex).Clase) Then
        With UserList(UserIndex)
            
            CantidadItems = Fix(4 + ((0.29 + 0.07 * Buleano(.Recompensas(1) = 1)) * Skill))
        End With
        
        MiObj.Amount = RandomNumber(1, CantidadItems)
    Else
        MiObj.Amount = 1
    End If
    MiObj.OBJIndex = Pescado
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    
    Call WriteConsoleMsg(UserIndex, "¡Has pescado un lindo pez!", FontTypeNames.FONTTYPE_INFO)
    
    Call SubirSkill(UserIndex, eSkill.Pesca, 15)
Else
    '[CDT 17-02-2004]
    If Not UserList(UserIndex).flags.UltimoMensaje = 6 Then
      Call WriteConsoleMsg(UserIndex, "¡No has pescado nada!", FontTypeNames.FONTTYPE_INFO)
      UserList(UserIndex).flags.UltimoMensaje = 6
    End If
    '[/CDT]
    
    Call SubirSkill(UserIndex, eSkill.Pesca, 5)
End If

UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

Exit Sub

ErrHandler:
    Call LogError("Error en DoPescar. Error " & Err.Number & " : " & Err.description)
End Sub

Public Sub DoPescarRed(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
On Error GoTo ErrHandler

Dim iSkill As Integer
Dim Suerte As Integer
Dim res As Integer
Dim EsPescador As Boolean

If esTrabajador(UserList(UserIndex).Clase) Then
    Call QuitarSta(UserIndex, EsfuerzoPescarPescador)
    EsPescador = True
Else
    Call QuitarSta(UserIndex, EsfuerzoPescarGeneral)
    EsPescador = False
End If

iSkill = UserList(UserIndex).Stats.UserSkills(eSkill.Pesca)

' m = (60-11)/(1-10)
' y = mx - m*10 + 11

Suerte = Int(-0.00125 * iSkill * iSkill - 0.3 * iSkill + 49)

If Suerte > 0 Then
    res = RandomNumber(1, Suerte)
    
    If res < 6 Then
        Dim MiObj As Obj
        Dim PecesPosibles(1 To 4) As Integer
        
        PecesPosibles(1) = PESCADO1
        PecesPosibles(2) = PESCADO2
        PecesPosibles(3) = PESCADO3
        PecesPosibles(4) = PESCADO4
        
        If EsPescador = True Then
            MiObj.Amount = RandomNumber(1, 5)
        Else
            MiObj.Amount = 1
        End If
        MiObj.OBJIndex = PecesPosibles(RandomNumber(LBound(PecesPosibles), UBound(PecesPosibles)))
        
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
        End If
        
        Call WriteConsoleMsg(UserIndex, "¡Has pescado algunos peces!", FontTypeNames.FONTTYPE_INFO)
        
        Call SubirSkill(UserIndex, eSkill.Pesca, 15)
    Else
        Call WriteConsoleMsg(UserIndex, "¡No has pescado nada!", FontTypeNames.FONTTYPE_INFO)
        Call SubirSkill(UserIndex, eSkill.Pesca, 5)
    End If
End If
        
Exit Sub

ErrHandler:
    Call LogError("Error en DoPescarRed")
End Sub

''
' Try to steal an item / gold to another character
'
' @param LadrOnIndex Specifies reference to user that stoles
' @param VictimaIndex Specifies reference to user that is being stolen

Public Sub DoRobar(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 05/04/2010
'Last Modification By: ZaMa
'24/07/08: Marco - Now it calls to WriteUpdateGold(VictimaIndex and LadrOnIndex) when the thief stoles gold. (MarKoxX)
'27/11/2009: ZaMa - Optimizacion de codigo.
'18/12/2009: ZaMa - Los ladrones ciudas pueden robar a pks.
'01/04/2010: ZaMa - Los ladrones pasan a robar oro acorde a su nivel.
'05/04/2010: ZaMa - Los armadas no pueden robarle a ciudadanos jamas.
'23/04/2010: ZaMa - No se puede robar mas sin energia.
'23/04/2010: ZaMa - El alcance de robo pasa a ser de 1 tile.
'*************************************************

On Error GoTo ErrHandler

    If Not MapInfo(UserList(VictimaIndex).Pos.map).Pk Then Exit Sub
    
    If TriggerZonaPelea(LadrOnIndex, VictimaIndex) <> TRIGGER6_AUSENTE Then Exit Sub
    
    
    With UserList(LadrOnIndex)
    
        ' Caos robando a caos?
        If EsCaos(VictimaIndex) And EsCaos(LadrOnIndex) Then
            Call WriteConsoleMsg(LadrOnIndex, "No puedes robar a otros miembros de la legión oscura.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        
        ' Tiene energia?
        If .Stats.MinSta < 15 Then
            If .Genero = eGenero.Hombre Then
                Call WriteConsoleMsg(LadrOnIndex, "Estás muy cansado para robar.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(LadrOnIndex, "Estás muy cansada para robar.", FontTypeNames.FONTTYPE_INFO)
            End If
            
            Exit Sub
        End If
        
        ' Quito energia
        Call QuitarSta(LadrOnIndex, 15)
        
        Dim GuantesHurto As Boolean
    
        If UserList(VictimaIndex).flags.Privilegios And PlayerType.User Then
            
            Dim Suerte As Integer
            Dim res As Integer
            Dim RobarSkill As Byte
            
            RobarSkill = .Stats.UserSkills(eSkill.Robar)
                
            res = RandomNumber(1, 100)
                
            If res > RobarSkill \ 10 + 25 * Buleano(.Clase = eClass.Ladron) + 5 * Buleano(.Clase = eClass.Ladron And .Recompensas(1) = 2) Then 'Exito robo
                If (.Clase = eClass.Ladron) And (.Recompensas(2) = 2) Then
                    If (RandomNumber(1, 100) <= IIf(.Recompensas(3) = 2, 20, 10)) Then
                        If TieneObjetosRobables(VictimaIndex) Then
                            Call RobarObjeto(LadrOnIndex, VictimaIndex)
                        Else
                            Call WriteConsoleMsg(LadrOnIndex, UserList(VictimaIndex).Name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)
                        End If
                    End If
                Else 'Roba oro
                    If UserList(VictimaIndex).Stats.GLD > 0 Then
                        Dim N As Integer
                        
                        If .Clase = eClass.Ladron Then
                        ' Si no tine puestos los guantes de hurto roba un 50% menos. Pablo (ToxicWaste)
                            If GuantesHurto Then
                                N = RandomNumber(.Stats.ELV * 50, .Stats.ELV * 100)
                            Else
                                N = RandomNumber(.Stats.ELV * 25, .Stats.ELV * 50)
                            End If
                        Else
                            N = RandomNumber(1, 100)
                        End If
                        If N > UserList(VictimaIndex).Stats.GLD Then N = UserList(VictimaIndex).Stats.GLD
                        UserList(VictimaIndex).Stats.GLD = UserList(VictimaIndex).Stats.GLD - N
                        
                        .Stats.GLD = .Stats.GLD + N
                        If .Stats.GLD > MAXORO Then _
                            .Stats.GLD = MAXORO
                        
                        Call WriteConsoleMsg(LadrOnIndex, "Le has robado " & N & " monedas de oro a " & UserList(VictimaIndex).Name, FontTypeNames.FONTTYPE_INFO)
                        Call WriteUpdateGold(LadrOnIndex) 'Le actualizamos la billetera al ladron
                        
                        Call WriteUpdateGold(VictimaIndex) 'Le actualizamos la billetera a la victima
                        Call FlushBuffer(VictimaIndex)
                    Else
                        Call WriteConsoleMsg(LadrOnIndex, UserList(VictimaIndex).Name & " no tiene oro.", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
                
                'Call SubirSkill(LadrOnIndex, eSkill.Robar)
            Else
                Call WriteConsoleMsg(LadrOnIndex, "¡No has logrado robar nada!", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(VictimaIndex, "¡" & .Name & " ha intentado robarte!", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(VictimaIndex, "¡" & .Name & " es un criminal!", FontTypeNames.FONTTYPE_INFO)
                Call FlushBuffer(VictimaIndex)
                
'                Call SubirSkill(LadrOnIndex, eSkill.Robar, False)
            End If
            
            Call SubirSkill(LadrOnIndex, eSkill.Robar)

        End If
    End With

Exit Sub

ErrHandler:
    Call LogError("Error en DoRobar. Error " & Err.Number & " : " & Err.description)

End Sub

''
' Check if one item is stealable
'
' @param VictimaIndex Specifies reference to victim
' @param Slot Specifies reference to victim's inventory slot
' @return If the item is stealable
Public Function ObjEsRobable(ByVal VictimaIndex As Integer, ByVal Slot As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
' Agregué los barcos
' Esta funcion determina qué objetos son robables.
'***************************************************

Dim OI As Integer

OI = UserList(VictimaIndex).Invent.Object(Slot).OBJIndex

ObjEsRobable = _
ObjData(OI).OBJType <> eOBJType.otLlaves And _
UserList(VictimaIndex).Invent.Object(Slot).Equipped = 0 And _
ObjData(OI).Real = 0 And _
ObjData(OI).Caos = 0 And _
ObjData(OI).OBJType <> eOBJType.otBarcos

End Function

''
' Try to steal an item to another character
'
' @param LadrOnIndex Specifies reference to user that stoles
' @param VictimaIndex Specifies reference to user that is being stolen
Public Sub RobarObjeto(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 02/04/2010
'02/04/2010: ZaMa - Modifico la cantidad de items robables por el ladron.
'***************************************************

Dim flag As Boolean
Dim i As Integer
flag = False

If RandomNumber(1, 12) < 6 Then 'Comenzamos por el principio o el final?
    i = 1
    Do While Not flag And i <= UserList(VictimaIndex).CurrentInventorySlots
        'Hay objeto en este slot?
        If UserList(VictimaIndex).Invent.Object(i).OBJIndex > 0 Then
           If ObjEsRobable(VictimaIndex, i) Then
                 If RandomNumber(1, 10) < 4 Then flag = True
           End If
        End If
        If Not flag Then i = i + 1
    Loop
Else
    i = 20
    Do While Not flag And i > 0
      'Hay objeto en este slot?
      If UserList(VictimaIndex).Invent.Object(i).OBJIndex > 0 Then
         If ObjEsRobable(VictimaIndex, i) Then
               If RandomNumber(1, 10) < 4 Then flag = True
         End If
      End If
      If Not flag Then i = i - 1
    Loop
End If

If flag Then
    Dim MiObj As Obj
    Dim Num As Byte
    Dim ObjAmount As Integer
    
    ObjAmount = UserList(VictimaIndex).Invent.Object(i).Amount
    
    'Cantidad al azar entre el 5% y el 10% del total, con minimo 1.
    Num = MaximoInt(1, RandomNumber(ObjAmount * 0.05, ObjAmount * 0.1))
                                
    MiObj.Amount = Num
    MiObj.OBJIndex = UserList(VictimaIndex).Invent.Object(i).OBJIndex
    
    UserList(VictimaIndex).Invent.Object(i).Amount = ObjAmount - Num
                
    If UserList(VictimaIndex).Invent.Object(i).Amount <= 0 Then
          Call QuitarUserInvItem(VictimaIndex, CByte(i), 1)
    End If
            
    Call UpdateUserInv(False, VictimaIndex, CByte(i))
                
    If Not MeterItemEnInventario(LadrOnIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(LadrOnIndex).Pos, MiObj)
    End If
    
    If UserList(LadrOnIndex).Clase = eClass.Ladron Then
        Call WriteConsoleMsg(LadrOnIndex, "Has robado " & MiObj.Amount & " " & ObjData(MiObj.OBJIndex).Name, FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(LadrOnIndex, "Has hurtado " & MiObj.Amount & " " & ObjData(MiObj.OBJIndex).Name, FontTypeNames.FONTTYPE_INFO)
    End If
Else
    Call WriteConsoleMsg(LadrOnIndex, "No has logrado robar ningún objeto.", FontTypeNames.FONTTYPE_INFO)
End If

'If exiting, cancel de quien es robado
Call CancelExit(VictimaIndex)

End Sub

Public Sub DoApuñalar(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer)
'***************************************************
'Autor: Nacho (Integer) & Unknown (orginal version)
'Last Modification: 04/17/08 - (NicoNZ)
'Simplifique la cuenta que hacia para sacar la suerte
'y arregle la cuenta que hacia para sacar el daño
'***************************************************
Dim Suerte As Integer
Dim Skill As Integer

Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar)

Suerte = 20 - 1.2 * Skill \ 10

Select Case UserList(UserIndex).Clase
    Case eClass.Asesino
        Suerte = Suerte - 3 - Buleano(UserList(UserIndex).Recompensas(3) = 2)
    
    Case eClass.Bardo
        Suerte = Suerte - 2 - Buleano(UserList(UserIndex).Recompensas(3) = 1)
End Select


If RandomNumber(1, 100) <= Suerte Then
    If VictimUserIndex <> 0 Then
        If UserList(UserIndex).Clase = eClass.Asesino And UserList(UserIndex).Recompensas(3) = 1 Then
            daño = Round(daño * 1.7, 0)
        Else
            daño = Round(daño * 1.5, 0)
        End If
        
        UserList(VictimUserIndex).Stats.MinHp = UserList(VictimUserIndex).Stats.MinHp - daño
        Call WriteConsoleMsg(UserIndex, "Has apuñalado a " & UserList(VictimUserIndex).Name & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(VictimUserIndex, "Te ha apuñalado " & UserList(UserIndex).Name & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)
        
        Call FlushBuffer(VictimUserIndex)
    Else
        If UserList(UserIndex).Clase = eClass.Asesino Then
            daño = daño * 2
        Else
            daño = daño * 1.5
        End If
        
        Npclist(VictimNpcIndex).Stats.MinHp = Npclist(VictimNpcIndex).Stats.MinHp - Int(daño)
        Call WriteConsoleMsg(UserIndex, "Has apuñalado la criatura por " & Int(daño), FontTypeNames.FONTTYPE_FIGHT)
        '[Alejo]
        Call CalcularDarExp(UserIndex, VictimNpcIndex, daño)
    End If
    
    Call SubirSkill(UserIndex, eSkill.Apuñalar, 5)
Else
    Call WriteConsoleMsg(UserIndex, "¡No has logrado apuñalar a tu enemigo!", FontTypeNames.FONTTYPE_FIGHT)
    'Call SubirSkill(UserIndex, eSkill.Apuñalar, False)
End If

End Sub

Public Sub DoGolpeCritico(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer)
'***************************************************
'Autor: Pablo (ToxicWaste)
'Last Modification: 28/01/2007
'***************************************************
Dim Suerte As Integer
Dim Skill As Integer

If UserList(UserIndex).Clase <> eClass.Bandido Then Exit Sub
If UserList(UserIndex).Invent.WeaponEqpSlot = 0 Then Exit Sub
If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).Name <> "Espada Vikinga" Then Exit Sub


Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling)

Suerte = Int((((0.00000003 * Skill + 0.000006) * Skill + 0.000107) * Skill + 0.0893) * 100)

If RandomNumber(0, 100) < Suerte Then
    daño = Int(daño * 0.75)
    If VictimUserIndex <> 0 Then
        UserList(VictimUserIndex).Stats.MinHp = UserList(VictimUserIndex).Stats.MinHp - daño
        Call WriteConsoleMsg(UserIndex, "Has golpeado críticamente a " & UserList(VictimUserIndex).Name & " por " & daño & ".", FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(VictimUserIndex, UserList(UserIndex).Name & " te ha golpeado críticamente por " & daño & ".", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Npclist(VictimNpcIndex).Stats.MinHp = Npclist(VictimNpcIndex).Stats.MinHp - daño
        Call WriteConsoleMsg(UserIndex, "Has golpeado críticamente a la criatura por " & daño & ".", FontTypeNames.FONTTYPE_FIGHT)
        '[Alejo]
        Call CalcularDarExp(UserIndex, VictimNpcIndex, daño)
    End If
End If

End Sub

Public Sub QuitarSta(ByVal UserIndex As Integer, ByVal Cantidad As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Cantidad
    If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
    Call WriteUpdateSta(UserIndex)
    
Exit Sub

ErrHandler:
    Call LogError("Error en QuitarSta. Error " & Err.Number & " : " & Err.description)
    
End Sub

Public Sub DoTalar(ByVal UserIndex As Integer, Optional ByVal DarMaderaElfica As Boolean = False)
'***************************************************
'Autor: Unknown
'Last Modification: 16/11/2009
'16/11/2009: ZaMa - Ahora Se puede dar madera elfica.
'16/11/2009: ZaMa - Implementado nuevo sistema de extraccion.
'***************************************************
On Error GoTo ErrHandler

Dim Suerte As Integer
Dim res As Integer
Dim CantidadItems As Integer

If esTrabajador(UserList(UserIndex).Clase) Then
    Call QuitarSta(UserIndex, EsfuerzoTalarLeñador)
Else
    Call QuitarSta(UserIndex, EsfuerzoTalarGeneral)
End If

Dim Skill As Integer
Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Talar)
Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)

res = RandomNumber(1, Suerte)

If res <= 6 Then
    Dim MiObj As Obj
    
    If esTrabajador(UserList(UserIndex).Clase) Then
        With UserList(UserIndex)
            If .Clase = eClass.Trabajador Then
                CantidadItems = Fix(4 + ((0.29 + 0.07 * Buleano(.Recompensas(1) = 1)) * Skill)) '1 + MaximoInt(1, CInt((.Stats.ELV - 4) / 5))
            Else
                CantidadItems = 1
            End If
        End With
        
        MiObj.Amount = RandomNumber(1, CantidadItems)
    Else
        MiObj.Amount = 1
    End If
    
    MiObj.OBJIndex = IIf(DarMaderaElfica, LeñaElfica, Leña)
    
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
        
    End If
    
    Call WriteConsoleMsg(UserIndex, "¡Has conseguido algo de leña!", FontTypeNames.FONTTYPE_INFO)
Else
    '[CDT 17-02-2004]
    If Not UserList(UserIndex).flags.UltimoMensaje = 8 Then
        Call WriteConsoleMsg(UserIndex, "¡No has obtenido leña!", FontTypeNames.FONTTYPE_INFO)
        UserList(UserIndex).flags.UltimoMensaje = 8
    End If
    '[/CDT]
End If

Call SubirSkill(UserIndex, eSkill.Talar, 5)
    
UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

Exit Sub

ErrHandler:
    Call LogError("Error en DoTalar")

End Sub

Public Sub DoMineria(ByVal UserIndex As Integer)
'***************************************************
'Autor: Unknown
'Last Modification: 16/11/2009
'16/11/2009: ZaMa - Implementado nuevo sistema de extraccion.
'***************************************************
On Error GoTo ErrHandler

Dim Suerte As Integer
Dim res As Integer
Dim CantidadItems As Integer

With UserList(UserIndex)
    If esTrabajador(.Clase) Then
        Call QuitarSta(UserIndex, EsfuerzoExcavarMinero)
    Else
        Call QuitarSta(UserIndex, EsfuerzoExcavarGeneral)
    End If
    
    Dim Skill As Integer
    Skill = .Stats.UserSkills(eSkill.Mineria)
    Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)
    
    res = RandomNumber(1, Suerte)
    
    If res <= 5 Then
        Dim MiObj As Obj
        
        If .flags.TargetObj = 0 Then Exit Sub
        
        MiObj.OBJIndex = ObjData(.flags.TargetObj).MineralIndex
        
        If esTrabajador(UserList(UserIndex).Clase) Then
            CantidadItems = Fix(4 + ((0.29 + 0.07 * Buleano(.Recompensas(1) = 1 And .Invent.HerramientaEqpObjIndex = PICO_EXPERTO * Skill))))
            
            MiObj.Amount = RandomNumber(1, CantidadItems)
        Else
            MiObj.Amount = 1
        End If
        
        If Not MeterItemEnInventario(UserIndex, MiObj) Then _
            Call TirarItemAlPiso(.Pos, MiObj)
        
        Call WriteConsoleMsg(UserIndex, "¡Has extraido algunos minerales!", FontTypeNames.FONTTYPE_INFO)
        
    Else
        '[CDT 17-02-2004]
        If Not .flags.UltimoMensaje = 9 Then
            Call WriteConsoleMsg(UserIndex, "¡No has conseguido nada!", FontTypeNames.FONTTYPE_INFO)
            .flags.UltimoMensaje = 9
        End If
        '[/CDT]
    End If
    
    Call SubirSkill(UserIndex, eSkill.Mineria, 5)

    .Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1
End With

Exit Sub

ErrHandler:
    Call LogError("Error en Sub DoMineria")

End Sub

Public Sub DoMeditar(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    With UserList(UserIndex)
        .Counters.IdleCount = 0
        
        Dim Suerte As Integer
        Dim res As Integer
        Dim cant As Integer
        Dim MeditarSkill As Byte
    
    
        'Barrin 3/10/03
        'Esperamos a que se termine de concentrar
        Dim TActual As Long
        TActual = GetTickCount() And &H7FFFFFFF
        If TActual - .Counters.tInicioMeditar < TIEMPO_INICIOMEDITAR Then
            Exit Sub
        End If
        
        If .Counters.bPuedeMeditar = False Then
            .Counters.bPuedeMeditar = True
        End If
            
        If .Stats.MinMAN >= .Stats.MaxMAN Then
            Call WriteConsoleMsg(UserIndex, "Has terminado de meditar.", FontTypeNames.FONTTYPE_INFO)
            Call WriteMeditateToggle(UserIndex)
            .flags.Meditando = False
            .Char.FX = 0
            .Char.loops = 0
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
            Exit Sub
        End If
        
        MeditarSkill = .Stats.UserSkills(eSkill.Meditar)
        
        If MeditarSkill <= 10 And MeditarSkill >= -1 Then
            Suerte = 35
        ElseIf MeditarSkill <= 20 And MeditarSkill >= 11 Then
            Suerte = 30
        ElseIf MeditarSkill <= 30 And MeditarSkill >= 21 Then
            Suerte = 28
        ElseIf MeditarSkill <= 40 And MeditarSkill >= 31 Then
            Suerte = 24
        ElseIf MeditarSkill <= 50 And MeditarSkill >= 41 Then
            Suerte = 22
        ElseIf MeditarSkill <= 60 And MeditarSkill >= 51 Then
            Suerte = 20
        ElseIf MeditarSkill <= 70 And MeditarSkill >= 61 Then
            Suerte = 18
        ElseIf MeditarSkill <= 80 And MeditarSkill >= 71 Then
            Suerte = 15
        ElseIf MeditarSkill <= 90 And MeditarSkill >= 81 Then
            Suerte = 10
        ElseIf MeditarSkill < 100 And MeditarSkill >= 91 Then
            Suerte = 7
        ElseIf MeditarSkill = 100 Then
            Suerte = 5
        End If
        res = RandomNumber(1, Suerte)
        
        If res = 1 Then
            
            cant = Porcentaje(.Stats.MaxMAN, PorcentajeRecuperoMana)
            If cant <= 0 Then cant = 1
            .Stats.MinMAN = .Stats.MinMAN + cant
            If .Stats.MinMAN > .Stats.MaxMAN Then _
                .Stats.MinMAN = .Stats.MaxMAN
            
            If Not .flags.UltimoMensaje = 22 Then
                Call WriteConsoleMsg(UserIndex, "¡Has recuperado " & cant & " puntos de maná!", FontTypeNames.FONTTYPE_INFO)
                .flags.UltimoMensaje = 22
            End If
            
            Call WriteUpdateMana(UserIndex)
        End If
        
        Call SubirSkill(UserIndex, eSkill.Meditar)
    End With
End Sub

Public Sub DoDesequipar(ByVal UserIndex As Integer, ByVal VictimIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modif: 15/04/2010
'Unequips either shield, weapon or helmet from target user.
'***************************************************

    Dim Probabilidad As Integer
    Dim Resultado As Integer
    Dim WrestlingSkill As Byte
    Dim AlgoEquipado As Boolean
    
    With UserList(UserIndex)
        
        ' Si no esta solo con manos, no desequipa tampoco.
        If .Invent.WeaponEqpObjIndex > 0 Then Exit Sub
        
        WrestlingSkill = .Stats.UserSkills(eSkill.Wrestling)
        
        Probabilidad = WrestlingSkill * 0.2 + .Stats.ELV * 0.66
   End With
   
   With UserList(VictimIndex)
        ' Si tiene escudo, intenta desequiparlo
        If .Invent.EscudoEqpObjIndex > 0 Then
            
            Resultado = RandomNumber(1, 100)
            
            If Resultado <= Probabilidad Then
                ' Se lo desequipo
                Call Desequipar(VictimIndex, .Invent.EscudoEqpSlot)
                
                Call WriteConsoleMsg(UserIndex, "Has logrado desequipar el escudo de tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                
                If .Stats.ELV < 20 Then
                    Call WriteConsoleMsg(VictimIndex, "¡Tu oponente te ha desequipado el escudo!", FontTypeNames.FONTTYPE_FIGHT)
                End If
                
                Call FlushBuffer(VictimIndex)
                
                Exit Sub
            End If
            
            AlgoEquipado = True
        End If
        
        ' No tiene escudo, o fallo desequiparlo, entonces trata de desequipar arma
        If .Invent.WeaponEqpObjIndex > 0 Then
            
            Resultado = RandomNumber(1, 100)
            
            If Resultado <= Probabilidad Then
                ' Se lo desequipo
                Call Desequipar(VictimIndex, .Invent.WeaponEqpSlot)
                
                Call WriteConsoleMsg(UserIndex, "Has logrado desarmar a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                
                If .Stats.ELV < 20 Then
                    Call WriteConsoleMsg(VictimIndex, "¡Tu oponente te ha desarmado!", FontTypeNames.FONTTYPE_FIGHT)
                End If
                
                Call FlushBuffer(VictimIndex)
                
                Exit Sub
            End If
            
            AlgoEquipado = True
        End If
        
        ' No tiene arma, o fallo desequiparla, entonces trata de desequipar casco
        If .Invent.CascoEqpObjIndex > 0 Then
            
            Resultado = RandomNumber(1, 100)
            
            If Resultado <= Probabilidad Then
                ' Se lo desequipo
                Call Desequipar(VictimIndex, .Invent.CascoEqpSlot)
                
                Call WriteConsoleMsg(UserIndex, "Has logrado desequipar el casco de tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                
                If .Stats.ELV < 20 Then
                    Call WriteConsoleMsg(VictimIndex, "¡Tu oponente te ha desequipado el casco!", FontTypeNames.FONTTYPE_FIGHT)
                End If
                
                Call FlushBuffer(VictimIndex)
                
                Exit Sub
            End If
            
            AlgoEquipado = True
        End If
    
        If AlgoEquipado Then
            Call WriteConsoleMsg(UserIndex, "Tu oponente no tiene equipado items!", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(UserIndex, "No has logrado desequipar ningún item a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
        End If
    
    End With


End Sub

Public Sub DoHurtar(ByVal UserIndex As Integer, ByVal VictimaIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modif: 03/03/2010
'Implements the pick pocket skill of the Bandit :)
'03/03/2010 - Pato: Sólo se puede hurtar si no está en trigger 6 :)
'***************************************************
If TriggerZonaPelea(UserIndex, VictimaIndex) <> TRIGGER6_AUSENTE Then Exit Sub

If UserList(UserIndex).Clase <> eClass.Bandido Then Exit Sub
'Esto es precario y feo, pero por ahora no se me ocurrió nada mejor.
'Uso el slot de los anillos para "equipar" los guantes.
'Y los reconozco porque les puse DefensaMagicaMin y Max = 0

Dim res As Integer
res = RandomNumber(1, 100)
If (res < 20) Then
    If TieneObjetosRobables(VictimaIndex) Then
        Call RobarObjeto(UserIndex, VictimaIndex)
        Call WriteConsoleMsg(VictimaIndex, "¡" & UserList(UserIndex).Name & " es un Bandido!", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(UserIndex, UserList(VictimaIndex).Name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)
    End If
End If

End Sub

Public Sub DoHandInmo(ByVal UserIndex As Integer, ByVal VictimaIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modif: 17/02/2007
'Implements the special Skill of the Thief
'***************************************************
If UserList(VictimaIndex).flags.Paralizado = 1 Then Exit Sub
If UserList(UserIndex).Clase <> eClass.Ladron Then Exit Sub
    
    
Dim res As Integer
res = RandomNumber(0, 100)
If res < (UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) / 4) Then
    UserList(VictimaIndex).flags.Paralizado = 1
    UserList(VictimaIndex).Counters.Paralisis = IntervaloParalizado / 2
    Call WriteParalizeOK(VictimaIndex)
    Call WriteConsoleMsg(UserIndex, "Tu golpe ha dejado inmóvil a tu oponente", FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(VictimaIndex, "¡El golpe te ha dejado inmóvil!", FontTypeNames.FONTTYPE_INFO)
End If

End Sub

Public Sub Desarmar(ByVal UserIndex As Integer, ByVal VictimIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 02/04/2010 (ZaMa)
'02/04/2010: ZaMa - Nueva formula para desarmar.
'***************************************************

    Dim Probabilidad As Integer
    Dim Resultado As Integer
    Dim WrestlingSkill As Byte
    
    With UserList(UserIndex)
        WrestlingSkill = .Stats.UserSkills(eSkill.Wrestling)
        
        Probabilidad = WrestlingSkill * 0.2 + .Stats.ELV * 0.66
        
        Resultado = RandomNumber(1, 100)
        
        If Resultado <= Probabilidad Then
            Call Desequipar(VictimIndex, UserList(VictimIndex).Invent.WeaponEqpSlot)
            Call WriteConsoleMsg(UserIndex, "Has logrado desarmar a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
            If UserList(VictimIndex).Stats.ELV < 20 Then
                Call WriteConsoleMsg(VictimIndex, "¡Tu oponente te ha desarmado!", FontTypeNames.FONTTYPE_FIGHT)
            End If
            Call FlushBuffer(VictimIndex)
        End If
    End With
    
End Sub


Public Function MaxItemsConstruibles(ByVal UserIndex As Integer) As Integer
'***************************************************
'Author: ZaMa
'Last Modification: 29/01/2010
'
'***************************************************
    MaxItemsConstruibles = MaximoInt(1, CInt((UserList(UserIndex).Stats.ELV - 4) / 5))
End Function

Public Function esTrabajador(ByVal Clase As eClass)

esTrabajador = (Clase = eClass.Carpintero Or Clase = eClass.Experto_Madera Or Clase = eClass.Experto_Minerales Or _
                Clase = eClass.Herrero Or Clase = eClass.Minero Or Clase = eClass.Pescador Or Clase = eClass.Sastre Or _
                Clase = eClass.Talador Or Clase = eClass.Trabajador)
End Function
