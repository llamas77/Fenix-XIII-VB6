Attribute VB_Name = "SistemaCombate"

'Argentum Online 0.12.2
'Copyright (C) 2002 M�rquez Pablo Ignacio
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez
'
'Dise�o y correcci�n del modulo de combate por
'Gerardo Saiz, gerardosaiz@yahoo.com
'

'9/01/2008 Pablo (ToxicWaste) - Ahora TODOS los modificadores de Clase se controlan desde Balance.dat


Option Explicit

Public Declare Function PoderAtaqueWrestling Lib "aolib.dll" Alias "PoderAtaqueWresterling" (ByVal Skill As Byte, ByVal Agilidad As Integer, Clase As Byte, ByVal Nivel As Byte) As Integer

Public Const MAXDISTANCIAARCO As Byte = 18

Public Function MinimoInt(ByVal A As Integer, ByVal b As Integer) As Integer
    If A > b Then
        MinimoInt = b
    Else
        MinimoInt = A
    End If
End Function

Public Function MaximoInt(ByVal A As Integer, ByVal b As Integer) As Integer
    If A > b Then
        MaximoInt = A
    Else
        MaximoInt = b
    End If
End Function

Private Function PoderEvasionEscudo(ByVal UserIndex As Integer) As Long
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    PoderEvasionEscudo = (UserList(UserIndex).Stats.UserSkills(eSkill.Defensa) * Mods(eMods.EVAESCUDO, UserList(UserIndex).Clase)) / 2
End Function

Private Function PoderEvasion(ByVal UserIndex As Integer) As Long
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
    With UserList(UserIndex)
        Select Case .Stats.UserSkills(Tacticas)
            Case Is < 31
                PoderEvasion = .Stats.UserSkills(Tacticas) * Mods(eMods.EVASION, .Clase)
            Case Is < 61
                PoderEvasion = (.Stats.UserSkills(Tacticas) + .Stats.UserAtributos(Agilidad)) * Mods(eMods.EVASION, .Clase)
            Case Is < 91
                PoderEvasion = (.Stats.UserSkills(Tacticas) + 2 * .Stats.UserAtributos(Agilidad)) * Mods(eMods.EVASION, .Clase)
            Case Else
                PoderEvasion = (.Stats.UserSkills(Tacticas) + 3 * .Stats.UserAtributos(Agilidad)) * Mods(eMods.EVASION, .Clase)
        End Select
        
        PoderEvasion = PoderEvasion + (2.5 * MaximoInt(.Stats.ELV - 12, 0))
    End With
End Function

Private Function PoderAtaqueArma(ByVal UserIndex As Integer) As Long
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
    With UserList(UserIndex)
        Select Case .Stats.UserSkills(Armas)
            Case Is < 31
                PoderAtaqueArma = .Stats.UserSkills(Armas) * Mods(eMods.CUERPOACUERPO, .Clase)
            Case Is < 61
                PoderAtaqueArma = (.Stats.UserSkills(Armas) + .Stats.UserAtributos(Agilidad)) * Mods(eMods.CUERPOACUERPO, .Clase)
            Case Is < 91
                PoderAtaqueArma = (.Stats.UserSkills(Armas) + 2 * .Stats.UserAtributos(Agilidad)) * Mods(eMods.CUERPOACUERPO, .Clase)
            Case Else
                PoderAtaqueArma = (.Stats.UserSkills(Armas) + 3 * .Stats.UserAtributos(Agilidad)) * Mods(eMods.CUERPOACUERPO, .Clase)
        End Select
        
        PoderAtaqueArma = PoderAtaqueArma + 2.5 * MaximoInt(.Stats.ELV - 12, 0)
    End With
End Function

Private Function PoderAtaqueProyectil(ByVal UserIndex As Integer) As Long
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim PoderAtaqueTemp As Long
    
    With UserList(UserIndex)
        Select Case .Stats.UserSkills(Proyectiles)
            Case Is < 31
                PoderAtaqueProyectil = .Stats.UserSkills(Proyectiles) * Mods(CONARCOS, .Clase)
            Case Is < 61
                PoderAtaqueProyectil = (.Stats.UserSkills(Proyectiles) + .Stats.UserAtributos(Agilidad)) * Mods(eMods.CONARCOS, .Clase)
            Case Is < 91
                PoderAtaqueProyectil = (.Stats.UserSkills(Proyectiles) + 2 * .Stats.UserAtributos(Agilidad)) * Mods(eMods.CONARCOS, .Clase)
            Case Else
                PoderAtaqueProyectil = (.Stats.UserSkills(Proyectiles) + 3 * .Stats.UserAtributos(Agilidad)) * Mods(eMods.CONARCOS, .Clase)
        End Select
        
        PoderAtaqueProyectil = (PoderAtaqueProyectil + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))

    End With
End Function

Public Function UserImpactoNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim PoderAtaque As Long
    Dim Arma As Integer
    Dim Skill As eSkill
    Dim ProbExito As Long
    
    Arma = UserList(UserIndex).Invent.WeaponEqpObjIndex
    
    With UserList(UserIndex)
        If Arma > 0 Then 'Usando un arma
            If ObjData(Arma).proyectil = 1 Then
                PoderAtaque = (1 + 0.05 * Buleano(.Clase = eClass.Arquero And .Recompensas(3) = 1) + _
                0.1 * Buleano(.Recompensas(3) = 1 And (.Clase = eClass.Guerrero Or .Clase = eClass.Cazador) * _
                PoderAtaqueProyectil(UserIndex)))
                
                Skill = eSkill.Proyectiles
            Else
                PoderAtaque = (1 + 0.05 * Buleano(.Clase = eClass.Paladin And .Recompensas(3) = 2)) * PoderAtaqueArma(UserIndex)
                Skill = eSkill.Armas
            End If
        Else 'Peleando con pu�os
            PoderAtaque = PoderAtaqueWrestling(18, UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad), DameClaseFenix(UserList(UserIndex).Clase), UserList(UserIndex).Stats.ELV) \ 4
            Skill = eSkill.Wrestling
        End If
    End With
        
    ' Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + ((PoderAtaque - Npclist(NpcIndex).PoderEvasion) * 0.4)))
    
    UserImpactoNpc = (RandomNumber(1, 100) <= ProbExito)
    
    Call SubirSkill(UserIndex, Skill)
    
End Function

Public Function NpcImpacto(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Revisa si un NPC logra impactar a un user o no
'03/15/2006 Maraxus - Evit� una divisi�n por cero que eliminaba NPCs
'*************************************************
    Dim Rechazo As Boolean
    Dim ProbRechazo As Long
    Dim ProbExito As Long
    Dim UserEvasion As Long
    Dim NpcPoderAtaque As Long
    Dim PoderEvasioEscudo As Long
    Dim SkillTacticas As Long
    Dim SkillDefensa As Long
    
    With UserList(UserIndex)
        UserEvasion = (1 + 0.05 * Buleano(.Recompensas(3) = 2 And (.Clase = eClass.Arquero Or .Clase = eClass.Nigromante) _
        * PoderEvasion(UserIndex)))
        
        NpcPoderAtaque = Npclist(NpcIndex).PoderAtaque
        PoderEvasioEscudo = PoderEvasionEscudo(UserIndex)
        
        SkillTacticas = UserList(UserIndex).Stats.UserSkills(eSkill.Tacticas)
        SkillDefensa = UserList(UserIndex).Stats.UserSkills(eSkill.Defensa)
        
        'Esta usando un escudo ???
        If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then UserEvasion = UserEvasion + PoderEvasioEscudo
        
        ' Chances are rounded
        ProbExito = MaximoInt(10, MinimoInt(90, 50 + ((NpcPoderAtaque - UserEvasion) * 0.4)))
        
        NpcImpacto = (RandomNumber(1, 100) <= ProbExito)
        
        ' el usuario esta usando un escudo ???
        If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
            If Not NpcImpacto Then
                If SkillDefensa + SkillTacticas > 0 Then  'Evitamos divisi�n por cero
                    ' Chances are rounded
                    ProbRechazo = MaximoInt(10, MinimoInt(90, 100 * SkillDefensa / (SkillDefensa + SkillTacticas)))
                    Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
                    
                    If Rechazo Then
                        'Se rechazo el ataque con el escudo
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_ESCUDO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                        Call WriteMultiMessage(UserIndex, eMessages.BlockedWithShieldUser) 'Call WriteBlockedWithShieldUser(UserIndex)
                        Call SubirSkill(UserIndex, eSkill.Defensa, 25)
                    End If
                End If
            End If
        End If
        
    End With
End Function

Public Function CalcularDa�o(ByVal UserIndex As Integer, Optional ByVal NpcIndex As Integer = 0) As Long
'***************************************************
'Author: Unknown
'Last Modification: 01/04/2010 (ZaMa)
'01/04/2010: ZaMa - Modifico el da�o de wrestling.
'01/04/2010: ZaMa - Agrego bonificadores de wrestling para los guantes.
'***************************************************
    Dim Da�oArma As Long
    Dim Da�oUsuario As Long
    Dim Arma As ObjData
    Dim ModifClase As Single
    Dim proyectil As ObjData
    Dim Da�oMaxArma As Long
    Dim Da�oMinArma As Long
    Dim OBJIndex As Integer
    
    ''sacar esto si no queremos q la matadracos mate el Dragon si o si
    Dim matoDragon As Boolean
    matoDragon = False
    
    With UserList(UserIndex)
        If .Invent.WeaponEqpObjIndex > 0 Then
            Arma = ObjData(.Invent.WeaponEqpObjIndex)
            
            ' Ataca a un npc?
            If NpcIndex > 0 Then
                If Arma.proyectil = 1 Then
                    ModifClase = Mods(eMods.Da�oConArcos, .Clase)
                    Da�oArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                    Da�oMaxArma = Arma.MaxHIT
                    
                    If Arma.Municion = 1 Then
                        proyectil = ObjData(.Invent.MunicionEqpObjIndex)
                        Da�oArma = Da�oArma + _
                        RandomNumber(proyectil.MinHIT + 10 * Buleano(.flags.BonusFlecha) + _
                                        5 * Buleano(.Clase = eClass.Arquero And .Recompensas(3) = 2), _
                                    proyectil.MaxHIT + 15 * Buleano(.flags.BonusFlecha) + _
                                        3 * Buleano(eClass.Arquero And .Recompensas(3) = 2))
                        
                        ' For some reason this isn't done...
                        'Da�oMaxArma = Da�oMaxArma + proyectil.MaxHIT
                    End If
                Else
                    ModifClase = Mods(eMods.Da�oCuerpoACuerpo, .Clase)
                    
                    If .Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then ' Usa la mata Dragones?
                        If Npclist(NpcIndex).NPCtype = DRAGON Then 'Ataca Dragon?
                            Da�oArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                            Da�oMaxArma = Arma.MaxHIT
                            matoDragon = True ''sacar esto si no queremos q la matadracos mate el Dragon si o si
                        Else ' Sino es Dragon da�o es 1
                            Da�oArma = 1
                            Da�oMaxArma = 1
                        End If
                    Else
                        Da�oArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                        Da�oMaxArma = Arma.MaxHIT
                    End If
                End If
            Else ' Ataca usuario
                If Arma.proyectil = 1 Then
                    ModifClase = Mods(eMods.Da�oConArcos, .Clase)
                    Da�oArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                    Da�oMaxArma = Arma.MaxHIT
                     
                    If Arma.Municion = 1 Then
                        proyectil = ObjData(.Invent.MunicionEqpObjIndex)
                        Da�oArma = Da�oArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHIT)
                        ' For some reason this isn't done...
                        'Da�oMaxArma = Da�oMaxArma + proyectil.MaxHIT
                    End If
                Else
                    ModifClase = Mods(eMods.Da�oCuerpoACuerpo, .Clase)
                    
                    If .Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then
                        'ModifClase = ModClase(.Clase).Da�oArmas
                        Da�oArma = 1 ' Si usa la espada mataDragones da�o es 1
                        Da�oMaxArma = 1
                    Else
                        Da�oArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                        Da�oMaxArma = Arma.MaxHIT
                    End If
                End If
            End If
            
        End If
        
        Da�oUsuario = RandomNumber(.Stats.MinHIT, .Stats.MaxHIT)
        
        ''sacar esto si no queremos q la matadracos mate el Dragon si o si
        If matoDragon Then
            CalcularDa�o = Npclist(NpcIndex).Stats.MinHp + Npclist(NpcIndex).Stats.def
        Else
            CalcularDa�o = (3 * Da�oArma + ((Da�oMaxArma / 5) * MaximoInt(0, .Stats.UserAtributos(eAtributos.Fuerza) - 15)) + Da�oUsuario) * ModifClase
        End If
    End With
End Function

Public Sub UserDa�oNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 07/04/2010 (ZaMa)
'25/01/2010: ZaMa - Agrego poder acuchillar npcs.
'07/04/2010: ZaMa - Los asesinos apu�alan acorde al da�o base sin descontar la defensa del npc.
'***************************************************

    Dim da�o As Long
    Dim Da�oBase As Long
    
    Da�oBase = CalcularDa�o(UserIndex, NpcIndex)
    
    'esta navegando? si es asi le sumamos el da�o del barco
    If UserList(UserIndex).flags.Navegando = 1 Then
        If UserList(UserIndex).Invent.BarcoObjIndex > 0 Then
            Da�oBase = Da�oBase + RandomNumber(ObjData(UserList(UserIndex).Invent.BarcoObjIndex).MinHIT, _
                                        ObjData(UserList(UserIndex).Invent.BarcoObjIndex).MaxHIT)
        End If
    End If
    
    With Npclist(NpcIndex)
        da�o = Da�oBase - .Stats.def
        
        If da�o < 0 Then da�o = 0
        
        'Call WriteUserHitNPC(UserIndex, da�o)
        Call WriteMultiMessage(UserIndex, eMessages.UserHitNPC, da�o)
        Call CalcularDarExp(UserIndex, NpcIndex, da�o)
        If .Stats.AutoCurar = 0 Then .Stats.MinHp = .Stats.MinHp - da�o
        
        If .Stats.MinHp > 0 Then
            'Trata de apu�alar por la espalda al enemigo
            If PuedeApu�alar(UserIndex) Then
               Call DoApu�alar(UserIndex, NpcIndex, 0, Da�oBase)
            End If
            
            'trata de dar golpe cr�tico
            Call DoGolpeCritico(UserIndex, NpcIndex, 0, da�o)

        End If
        
        
        If .Stats.MinHp <= 0 Then
            ' Si era un Dragon perdemos la espada mataDragones
            If .NPCtype = DRAGON Then
                'Si tiene equipada la matadracos se la sacamos
                If UserList(UserIndex).Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then
                    Call QuitarObjetos(EspadaMataDragonesIndex, 1, UserIndex)
                End If
            End If
            
            ' Para que las mascotas no sigan intentando luchar y
            ' comiencen a seguir al amo
            Dim j As Integer
            For j = 1 To MAXMASCOTAS
                If UserList(UserIndex).MascotasIndex(j) > 0 Then
                    If Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNPC = NpcIndex Then
                        Npclist(UserList(UserIndex).MascotasIndex(j)).TargetNPC = 0
                        Npclist(UserList(UserIndex).MascotasIndex(j)).Movement = TipoAI.SigueAmo
                    End If
                End If
            Next j
            
            Call MuereNpc(NpcIndex, UserIndex)
        End If
    End With
End Sub

Public Sub NpcDa�o(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim da�o As Integer
    Dim Lugar As Integer
    Dim absorbido As Integer
    Dim defbarco As Integer
    Dim Obj As ObjData
    
    da�o = RandomNumber(Npclist(NpcIndex).Stats.MinHIT, Npclist(NpcIndex).Stats.MaxHIT)
    
    With UserList(UserIndex)
        If .flags.Navegando = 1 And .Invent.BarcoObjIndex > 0 Then
            Obj = ObjData(.Invent.BarcoObjIndex)
            defbarco = RandomNumber(Obj.MinDef, Obj.MaxDef)
        End If
        
        Lugar = RandomNumber(PartesCuerpo.bCabeza, PartesCuerpo.bTorso)
        
        Select Case Lugar
            Case PartesCuerpo.bCabeza
                'Si tiene casco absorbe el golpe
                If .Invent.CascoEqpObjIndex > 0 Then
                   Obj = ObjData(.Invent.CascoEqpObjIndex)
                   If Obj.Gorro = 0 Then absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
                End If
          Case Else
                'Si tiene armadura absorbe el golpe
                If .Invent.ArmourEqpObjIndex > 0 Then
                    Dim Obj2 As ObjData
                    Obj = ObjData(.Invent.ArmourEqpObjIndex)
                    If .Invent.EscudoEqpObjIndex Then
                        Obj2 = ObjData(.Invent.EscudoEqpObjIndex)
                        absorbido = RandomNumber(Obj.MinDef + Obj2.MinDef, Obj.MaxDef + Obj2.MaxDef)
                    Else
                        absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
                   End If
                End If
        End Select
        
        absorbido = absorbido + defbarco + 2 * Buleano(.Clase = eClass.Guerrero And .Recompensas(2) = 2)
        
        da�o = da�o - absorbido
        If da�o < 1 Then da�o = 1
        
        Call WriteMultiMessage(UserIndex, eMessages.NPCHitUser, Lugar, da�o)
        'Call WriteNPCHitUser(UserIndex, Lugar, da�o)
        
        If .flags.Privilegios And PlayerType.User Then .Stats.MinHp = .Stats.MinHp - da�o
        
        If .flags.Meditando Then
            If da�o > Fix(.Stats.MinHp / 100 * .Stats.UserAtributos(eAtributos.Inteligencia) * .Stats.UserSkills(eSkill.Meditar) / 100 * 12 / (RandomNumber(0, 5) + 7)) Then
                .flags.Meditando = False
                Call WriteMeditateToggle(UserIndex)
                Call WriteConsoleMsg(UserIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
                .Char.FX = 0
                .Char.loops = 0
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
            End If
        End If
        
        'Muere el usuario
        If .Stats.MinHp <= 0 Then
            Call WriteMultiMessage(UserIndex, eMessages.NPCKillUser) 'Call WriteNPCKillUser(UserIndex) ' Le informamos que ha muerto ;)
            
            If Npclist(NpcIndex).MaestroUser > 0 Then
                Call AllFollowAmo(Npclist(NpcIndex).MaestroUser)
            Else
                'Al matarlo no lo sigue mas
                If Npclist(NpcIndex).Stats.Alineacion = 0 Then
                    Npclist(NpcIndex).Movement = Npclist(NpcIndex).flags.OldMovement
                    Npclist(NpcIndex).Hostile = Npclist(NpcIndex).flags.OldHostil
                    Npclist(NpcIndex).flags.AttackedBy = vbNullString
                End If
            End If
            
            Call UserDie(UserIndex)
        End If
    End With
End Sub

Public Sub CheckPets(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, Optional ByVal CheckElementales As Boolean = True)
'***************************************************
'Author: Unknown
'Last Modification: 15/04/2010
'15/04/2010: ZaMa - Las mascotas no se apropian de npcs.
'***************************************************

    Dim j As Integer
    
    ' Si no tengo mascotas, para que cheaquear lo demas?
    If UserList(UserIndex).NroMascotas = 0 Then Exit Sub
    
    If Not PuedeAtacarNPC(UserIndex, NpcIndex, , True) Then Exit Sub
    
    With UserList(UserIndex)
        For j = 1 To MAXMASCOTAS
            If .MascotasIndex(j) > 0 Then
               If .MascotasIndex(j) <> NpcIndex Then
                If CheckElementales Or (Npclist(.MascotasIndex(j)).Numero <> ELEMENTALFUEGO And Npclist(.MascotasIndex(j)).Numero <> ELEMENTALTIERRA) Then
                    
                    If Npclist(.MascotasIndex(j)).TargetNPC = 0 Then Npclist(.MascotasIndex(j)).TargetNPC = NpcIndex
                    Npclist(.MascotasIndex(j)).Movement = TipoAI.NpcAtacaNpc
                End If
               End If
            End If
        Next j
    End With
End Sub

Public Sub AllFollowAmo(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim j As Integer
    
    For j = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasIndex(j) > 0 Then
            Call FollowAmo(UserList(UserIndex).MascotasIndex(j))
        End If
    Next j
End Sub

Public Function NpcAtacaUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean
'*************************************************
'Author: Unknown
'Last modified: -
'
'*************************************************

    With UserList(UserIndex)
        If .flags.AdminInvisible = 1 Then Exit Function
        If (Not .flags.Privilegios And PlayerType.User) <> 0 And Not .flags.AdminPerseguible Then Exit Function
    End With
    
    With Npclist(NpcIndex)
        If .Stats.AutoCurar = 1 Then Exit Function
        
        ' El npc puede atacar ???
        If .CanAttack = 1 Then
            NpcAtacaUser = True
            Call CheckPets(NpcIndex, UserIndex, False)
            
            If .Target = 0 Then .Target = UserIndex
            
            If UserList(UserIndex).flags.AtacadoPorNpc = 0 And UserList(UserIndex).flags.AtacadoPorUser = 0 Then
                UserList(UserIndex).flags.AtacadoPorNpc = NpcIndex
            End If
        Else
            NpcAtacaUser = False
            Exit Function
        End If
        
        .CanAttack = 0
        
        If .flags.Snd1 > 0 Then
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(.flags.Snd1, .Pos.X, .Pos.Y))
        End If
    End With
    
    If NpcImpacto(NpcIndex, UserIndex) Then
        With UserList(UserIndex)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO, .Pos.X, .Pos.Y))
            
            If .flags.Meditando = False Then
                If .flags.Navegando = 0 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FXSANGRE, 0))
                End If
            End If
            
            Call NpcDa�o(NpcIndex, UserIndex)
            Call WriteUpdateHP(UserIndex)
            
            '�Puede envenenar?
            If Npclist(NpcIndex).Veneno = 1 Then Call NpcEnvenenarUser(UserIndex)
        End With
    Else
        Call WriteMultiMessage(UserIndex, eMessages.NPCSwing)
    End If
    
    Call SubirSkill(UserIndex, eSkill.Tacticas)
    
    'Controla el nivel del usuario
    Call CheckUserLevel(UserIndex)
End Function

Private Function NpcImpactoNpc(ByVal Atacante As Integer, ByVal Victima As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim PoderAtt As Long
    Dim PoderEva As Long
    Dim ProbExito As Long
    
    PoderAtt = Npclist(Atacante).PoderAtaque
    PoderEva = Npclist(Victima).PoderEvasion
    
    ' Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + (PoderAtt - PoderEva) * 0.4))
    NpcImpactoNpc = (RandomNumber(1, 100) <= ProbExito)
End Function

Public Sub NpcDa�oNpc(ByVal Atacante As Integer, ByVal Victima As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim da�o As Integer
    
    With Npclist(Atacante)
        da�o = RandomNumber(.Stats.MinHIT, .Stats.MaxHIT)
        If Npclist(Victima).Stats.AutoCurar = 0 Then Npclist(Victima).Stats.MinHp = Npclist(Victima).Stats.MinHp - da�o
        
        If Npclist(Victima).Stats.MinHp < 1 Then
            .Movement = .flags.OldMovement
            
            If LenB(.flags.AttackedBy) <> 0 Then
                .Hostile = .flags.OldHostil
            End If
            
            If .MaestroUser > 0 Then
                Call FollowAmo(Atacante)
            End If
            
            Call MuereNpc(Victima, .MaestroUser)
        End If
    End With
End Sub

Public Sub NpcAtacaNpc(ByVal Atacante As Integer, ByVal Victima As Integer, Optional ByVal cambiarMOvimiento As Boolean = True)
'*************************************************
'Author: Unknown
'Last modified: 01/03/2009
'01/03/2009: ZaMa - Las mascotas no pueden atacar al rey si quedan pretorianos vivos.
'*************************************************
    
    With Npclist(Atacante)
        
        'Es el Rey Preatoriano?
        If Npclist(Victima).Numero = PRKING_NPC Then
            If pretorianosVivos > 0 Then
                Call WriteConsoleMsg(.MaestroUser, "Debes matar al resto del ej�rcito antes de atacar al rey!", FontTypeNames.FONTTYPE_FIGHT)
                .TargetNPC = 0
                Exit Sub
            End If
        End If
        
        ' El npc puede atacar ???
        If .CanAttack = 1 Then
            .CanAttack = 0
            If cambiarMOvimiento Then
                Npclist(Victima).TargetNPC = Atacante
                Npclist(Victima).Movement = TipoAI.NpcAtacaNpc
            End If
        Else
            Exit Sub
        End If
        
        If .flags.Snd1 > 0 Then
            Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(.flags.Snd1, .Pos.X, .Pos.Y))
        End If
        
        If NpcImpactoNpc(Atacante, Victima) Then
            If Npclist(Victima).flags.Snd2 > 0 Then
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(Npclist(Victima).flags.Snd2, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))
            Else
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO2, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))
            End If
        
            If .MaestroUser > 0 Then
                Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(SND_IMPACTO, .Pos.X, .Pos.Y))
            Else
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))
            End If
            
            Call NpcDa�oNpc(Atacante, Victima)
        Else
            If .MaestroUser > 0 Then
                Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
            Else
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_SWING, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))
            End If
        End If
    End With
End Sub

Public Function UsuarioAtacaNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 14/01/2010 (ZaMa)
'12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados por npcs cuando los atacan.
'14/01/2010: ZaMa - Lo transformo en funci�n, para que no se pierdan municiones al atacar targets inv�lidos.
'***************************************************

On Error GoTo Errhandler

    If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then Exit Function
    
    Call NPCAtacado(NpcIndex, UserIndex)
    
    If UserImpactoNpc(UserIndex, NpcIndex) Then
        If Npclist(NpcIndex).flags.Snd2 > 0 Then
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(Npclist(NpcIndex).flags.Snd2, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y))
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO2, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y))
        End If
        
        Call UserDa�oNpc(UserIndex, NpcIndex)
    Else
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SWING, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
        Call WriteMultiMessage(UserIndex, eMessages.UserSwing)
    End If
    
    ' Revel� su condici�n de usuario al atacar, los npcs lo van a atacar
    UserList(UserIndex).flags.Ignorado = False
    
    UsuarioAtacaNpc = True
    
    Exit Function
    
Errhandler:
    Call LogError("Error en UsuarioAtacaNpc. Error " & Err.Number & " : " & Err.description)
    
End Function

Public Sub UsuarioAtaca(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim index As Integer
    Dim AttackPos As WorldPos
    
    'Check bow's interval
    If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub
    
    'Check Spell-Magic interval
    If Not IntervaloPermiteMagiaGolpe(UserIndex) Then
        'Check Attack interval
        If Not IntervaloPermiteAtacar(UserIndex) Then
            Exit Sub
        End If
    End If
    
    With UserList(UserIndex)
        'Quitamos stamina
        If .Stats.MinSta >= 10 Then
            Call QuitarSta(UserIndex, RandomNumber(1, 10))
        Else
            If .Genero = eGenero.Hombre Then
                Call WriteConsoleMsg(UserIndex, "Est�s muy cansado para luchar.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Est�s muy cansada para luchar.", FontTypeNames.FONTTYPE_INFO)
            End If
            Exit Sub
        End If
        
        AttackPos = .Pos
        Call HeadtoPos(.Char.heading, AttackPos)
        
        'Exit if not legal
        If AttackPos.X < XMinMapSize Or AttackPos.X > XMaxMapSize Or AttackPos.Y <= YMinMapSize Or AttackPos.Y > YMaxMapSize Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
            Exit Sub
        End If
        
        index = MapData(AttackPos.map, AttackPos.X, AttackPos.Y).UserIndex
        
        'Look for user
        If index > 0 Then
            Call UsuarioAtacaUsuario(UserIndex, index)
            Call WriteUpdateUserStats(UserIndex)
            Call WriteUpdateUserStats(index)
            Exit Sub
        End If
        
        index = MapData(AttackPos.map, AttackPos.X, AttackPos.Y).NpcIndex
        
        'Look for NPC
        If index > 0 Then
            If Npclist(index).Attackable Then
                If Npclist(index).MaestroUser > 0 And MapInfo(Npclist(index).Pos.map).Pk = False Then
                    Call WriteConsoleMsg(UserIndex, "No puedes atacar mascotas en zona segura.", FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub
                End If
                
                Call UsuarioAtacaNpc(UserIndex, index)
            Else
                Call WriteConsoleMsg(UserIndex, "No puedes atacar a este NPC.", FontTypeNames.FONTTYPE_FIGHT)
            End If
            
            Call WriteUpdateUserStats(UserIndex)
            
            Exit Sub
        End If
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
        Call WriteUpdateUserStats(UserIndex)
        
        If .Counters.Trabajando Then .Counters.Trabajando = .Counters.Trabajando - 1
            
        If .Counters.Ocultando Then .Counters.Ocultando = .Counters.Ocultando - 1
    End With
End Sub

Public Function UsuarioImpacto(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo Errhandler

    Dim ProbRechazo As Long
    Dim Rechazo As Boolean
    Dim ProbExito As Long
    Dim PoderAtaque As Long
    Dim UserPoderEvasion As Long
    Dim UserPoderEvasionEscudo As Long
    Dim Arma As Integer
    Dim SkillTacticas As Long
    Dim SkillDefensa As Long
    Dim ProbEvadir As Long
    Dim Skill As eSkill
    
    SkillTacticas = UserList(VictimaIndex).Stats.UserSkills(eSkill.Tacticas)
    SkillDefensa = UserList(VictimaIndex).Stats.UserSkills(eSkill.Defensa)
    
    Arma = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
    
    'Calculamos el poder de evasion...
    UserPoderEvasion = (1 + 0.05 * Buleano(UserList(VictimaIndex).Recompensas(3) = 2 And _
        (UserList(VictimaIndex).Clase = eClass.Arquero Or UserList(VictimaIndex).Clase = eClass.Nigromante))) * _
        PoderEvasion(VictimaIndex)
    
    If UserList(VictimaIndex).Invent.EscudoEqpObjIndex > 0 Then
       UserPoderEvasionEscudo = PoderEvasionEscudo(VictimaIndex)
       UserPoderEvasion = UserPoderEvasion + UserPoderEvasionEscudo
    Else
        UserPoderEvasionEscudo = 0
    End If
    
    'todo: check damages
    'Esta usando un arma ???
    If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then
        If ObjData(Arma).proyectil = 1 Then
            PoderAtaque = (1 + 0.05 * Buleano(UserList(AtacanteIndex).Clase = eClass.Arquero And UserList(AtacanteIndex).Recompensas(3) = 1) + _
                0.1 * Buleano(UserList(AtacanteIndex).Recompensas(3) = 1 And (UserList(AtacanteIndex).Clase = eClass.Guerrero Or UserList(AtacanteIndex).Clase = eClass.Cazador))) * _
                PoderAtaqueProyectil(AtacanteIndex)
                
            Skill = eSkill.Proyectiles
        Else
            PoderAtaque = (1 + 0.05 * Buleano(UserList(AtacanteIndex).Clase = eClass.Paladin And UserList(AtacanteIndex).Recompensas(3) = 2)) * _
                PoderAtaqueArma(AtacanteIndex)
                
            Skill = eSkill.Armas
        End If
    Else
        PoderAtaque = PoderAtaqueWrestling(18, UserList(AtacanteIndex).Stats.UserAtributos(eAtributos.Agilidad), DameClaseFenix(UserList(AtacanteIndex).Clase), UserList(AtacanteIndex).Stats.ELV)
        Skill = eSkill.Wrestling
    End If
    
    ' Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + (PoderAtaque - UserPoderEvasion) * 0.4))
    
    ' Se reduce la evasion un 25%
    'If UserList(VictimaIndex).flags.Meditando = True Then
    '    ProbEvadir = (100 - ProbExito) * 0.75
    '    ProbExito = MinimoInt(90, 100 - ProbEvadir)
    'End If
    
    UsuarioImpacto = (RandomNumber(1, 100) <= ProbExito)
    
    ' el usuario esta usando un escudo ???
    If UserList(VictimaIndex).Invent.EscudoEqpObjIndex > 0 Then
        'Fallo ???
        If Not UsuarioImpacto Then
            ' Chances are rounded
            ProbRechazo = MaximoInt(10, MinimoInt(90, 100 * SkillDefensa / (SkillDefensa + SkillTacticas)))
            Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
            If Rechazo Then
                'Se rechazo el ataque con el escudo
                Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessagePlayWave(SND_ESCUDO, UserList(VictimaIndex).Pos.X, UserList(VictimaIndex).Pos.Y))
                  
                Call WriteMultiMessage(AtacanteIndex, eMessages.BlockedWithShieldother)
                Call WriteMultiMessage(VictimaIndex, eMessages.BlockedWithShieldUser)
                
                Call SubirSkill(VictimaIndex, eSkill.Defensa, 25)
            End If
        End If
    End If
    
    If Not UsuarioImpacto Then
        Call SubirSkill(AtacanteIndex, Skill)
    End If
    
    Call FlushBuffer(VictimaIndex)
    
    Exit Function
    
Errhandler:
    Dim AtacanteNick As String
    Dim VictimaNick As String
    
    If AtacanteIndex > 0 Then AtacanteNick = UserList(AtacanteIndex).Name
    If VictimaIndex > 0 Then VictimaNick = UserList(VictimaIndex).Name
    
    Call LogError("Error en UsuarioImpacto. Error " & Err.Number & " : " & Err.description & " AtacanteIndex: " & _
             AtacanteIndex & " Nick: " & AtacanteNick & " VictimaIndex: " & VictimaIndex & " Nick: " & VictimaNick)
End Function

Public Function UsuarioAtacaUsuario(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 14/01/2010 (ZaMa)
'14/01/2010: ZaMa - Lo transformo en funci�n, para que no se pierdan municiones al atacar targets
'                    inv�lidos, y evitar un doble chequeo innecesario
'***************************************************

On Error GoTo Errhandler

    If Not PuedeAtacar(AtacanteIndex, VictimaIndex) Then Exit Function
    
    With UserList(AtacanteIndex)
        If Distancia(.Pos, UserList(VictimaIndex).Pos) > MAXDISTANCIAARCO Then
           Call WriteConsoleMsg(AtacanteIndex, "Est�s muy lejos para disparar.", FontTypeNames.FONTTYPE_FIGHT)
           Exit Function
        End If
        
        Call UsuarioAtacadoPorUsuario(AtacanteIndex, VictimaIndex)
        
        If UsuarioImpacto(AtacanteIndex, VictimaIndex) Then
            Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayWave(SND_IMPACTO, .Pos.X, .Pos.Y))
            
            If UserList(VictimaIndex).flags.Navegando = 0 Then
                Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.CharIndex, FXSANGRE, 0))
            End If
            
            'Pablo (ToxicWaste): Guantes de Hurto del Bandido en acci�n
            If .Clase = eClass.Bandido Then
                Call DoDesequipar(AtacanteIndex, VictimaIndex)
                
            'y ahora, el ladr�n puede llegar a paralizar con el golpe.
            ElseIf .Clase = eClass.Ladron Then
                Call DoHandInmo(AtacanteIndex, VictimaIndex)
            End If
            
            Call SubirSkill(VictimaIndex, eSkill.Tacticas)
            Call UserDa�oUser(AtacanteIndex, VictimaIndex)
        Else
            ' Invisible admins doesn't make sound to other clients except itself
            If .flags.AdminInvisible = 1 Then
                Call EnviarDatosASlot(AtacanteIndex, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
            Else
                Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
            End If
            
            Call WriteMultiMessage(AtacanteIndex, eMessages.UserSwing)
            Call WriteMultiMessage(VictimaIndex, eMessages.UserAttackedSwing, AtacanteIndex)
            Call SubirSkill(VictimaIndex, eSkill.Tacticas)
        End If
        
        If .Clase = eClass.Ladron Then Call Desarmar(AtacanteIndex, VictimaIndex)
    End With
    
    UsuarioAtacaUsuario = True
    
    Exit Function
    
Errhandler:
    Call LogError("Error en UsuarioAtacaUsuario. Error " & Err.Number & " : " & Err.description)
End Function

Public Sub UserDa�oUser(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 12/01/2010 (ZaMa)
'12/01/2010: ZaMa - Implemento armas arrojadizas y probabilidad de acuchillar
'11/03/2010: ZaMa - Ahora no cuenta la muerte si estaba en estado atacable, y no se vuelve criminal
'***************************************************
    
On Error GoTo Errhandler

    Dim da�o As Long
    Dim Lugar As Byte
    Dim absorbido As Long
    Dim defbarco As Integer
    Dim Obj As ObjData
    
    da�o = CalcularDa�o(AtacanteIndex)
    
    Call UserEnvenena(AtacanteIndex, VictimaIndex)
    
    With UserList(AtacanteIndex)
        If .flags.Navegando = 1 And .Invent.BarcoObjIndex > 0 Then
             Obj = ObjData(.Invent.BarcoObjIndex)
             da�o = da�o + RandomNumber(Obj.MinHIT, Obj.MaxHIT)
        End If
        
        If UserList(VictimaIndex).flags.Navegando = 1 And UserList(VictimaIndex).Invent.BarcoObjIndex > 0 Then
             Obj = ObjData(UserList(VictimaIndex).Invent.BarcoObjIndex)
             defbarco = RandomNumber(Obj.MinDef, Obj.MaxDef)
        End If
        
        Lugar = RandomNumber(PartesCuerpo.bCabeza, PartesCuerpo.bTorso)
        
        Select Case Lugar
            Case PartesCuerpo.bCabeza
                'Si tiene casco absorbe el golpe
                If UserList(VictimaIndex).Invent.CascoEqpObjIndex > 0 Then
                    If Not (.Clase = eClass.Arquero And .Recompensas(3) = 2) Then
                        Obj = ObjData(UserList(VictimaIndex).Invent.CascoEqpObjIndex)
                        absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
                    End If
                End If
            
            Case Else
                'Si tiene armadura absorbe el golpe
                If UserList(VictimaIndex).Invent.ArmourEqpObjIndex > 0 Then
                    Obj = ObjData(UserList(VictimaIndex).Invent.ArmourEqpObjIndex)
                    Dim Obj2 As ObjData
                    If UserList(VictimaIndex).Invent.EscudoEqpObjIndex Then
                        Obj2 = ObjData(UserList(VictimaIndex).Invent.EscudoEqpObjIndex)
                        absorbido = RandomNumber(Obj.MinDef + Obj2.MinDef, Obj.MaxDef + Obj2.MaxDef)
                    Else
                        absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
                    End If
                End If
        End Select
        
        absorbido = absorbido + defbarco + 2 * Buleano(UserList(VictimaIndex).Clase = eClass.Guerrero And UserList(VictimaIndex).Recompensas(2) = 2)
        da�o = da�o - absorbido
        If da�o < 0 Then da�o = 1
                    
        Call WriteMultiMessage(AtacanteIndex, eMessages.UserHittedUser, UserList(VictimaIndex).Char.CharIndex, Lugar, da�o)
        Call WriteMultiMessage(VictimaIndex, eMessages.UserHittedByUser, .Char.CharIndex, Lugar, da�o)
        
        UserList(VictimaIndex).Stats.MinHp = UserList(VictimaIndex).Stats.MinHp - da�o
        
        If .flags.Hambre = 0 And .flags.Sed = 0 Then
            'Si usa un arma quizas suba "Combate con armas"
            If .Invent.WeaponEqpObjIndex > 0 Then
                If ObjData(.Invent.WeaponEqpObjIndex).proyectil Then
                    'es un Arco. Sube Armas a Distancia
                    Call SubirSkill(AtacanteIndex, eSkill.Proyectiles)
                Else
                    'Sube combate con armas.
                    Call SubirSkill(AtacanteIndex, eSkill.Armas)
                End If
            Else
                'sino tal vez lucha libre
                Call SubirSkill(AtacanteIndex, eSkill.Wrestling)
            End If
                    
            'Trata de apu�alar por la espalda al enemigo
            If PuedeApu�alar(AtacanteIndex) Then
                Call DoApu�alar(AtacanteIndex, 0, VictimaIndex, da�o)
            End If
            'e intenta dar un golpe cr�tico [Pablo (ToxicWaste)]
            Call DoGolpeCritico(AtacanteIndex, 0, VictimaIndex, da�o)
        End If
        
        If UserList(VictimaIndex).Stats.MinHp <= 0 Then
            
            ' Para que las mascotas no sigan intentando luchar y
            ' comiencen a seguir al amo
            Dim j As Integer
            For j = 1 To MAXMASCOTAS
                If .MascotasIndex(j) > 0 Then
                    If Npclist(.MascotasIndex(j)).Target = VictimaIndex Then
                        Npclist(.MascotasIndex(j)).Target = 0
                        Call FollowAmo(.MascotasIndex(j))
                    End If
                End If
            Next j
            
            Call ActStats(VictimaIndex, AtacanteIndex)
            Call UserDie(VictimaIndex)
        Else
            'Est� vivo - Actualizamos el HP
            Call WriteUpdateHP(VictimaIndex)
        End If
    End With
    
    'Controla el nivel del usuario
    Call CheckUserLevel(AtacanteIndex)
    
    Call FlushBuffer(VictimaIndex)
    
    Exit Sub
    
Errhandler:
    Dim AtacanteNick As String
    Dim VictimaNick As String
    
    If AtacanteIndex > 0 Then AtacanteNick = UserList(AtacanteIndex).Name
    If VictimaIndex > 0 Then VictimaNick = UserList(VictimaIndex).Name
    
    Call LogError("Error en UserDa�oUser. Error " & Err.Number & " : " & Err.description & " AtacanteIndex: " & _
             AtacanteIndex & " Nick: " & AtacanteNick & " VictimaIndex: " & VictimaIndex & " Nick: " & VictimaNick)
End Sub

Sub UsuarioAtacadoPorUsuario(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer)
'***************************************************
'Autor: Unknown
'Last Modification: 05/05/2010
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
'10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
'05/05/2010: ZaMa - Ahora no suma puntos de bandido al atacar a alguien en estado atacable.
'***************************************************

    If TriggerZonaPelea(AttackerIndex, VictimIndex) = TRIGGER6_PERMITE Then Exit Sub
    
    Dim EraCriminal As Boolean
    Dim VictimaEsAtacable As Boolean
    
    With UserList(VictimIndex)
        If .flags.Meditando Then
            .flags.Meditando = False
            Call WriteMeditateToggle(VictimIndex)
            Call WriteConsoleMsg(VictimIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
            .Char.FX = 0
            .Char.loops = 0
            Call SendData(SendTarget.ToPCArea, VictimIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
        End If
    End With
    
    Call AllMascotasAtacanUser(AttackerIndex, VictimIndex)
    Call AllMascotasAtacanUser(VictimIndex, AttackerIndex)
    
    'Si la victima esta saliendo se cancela la salida
    Call CancelExit(VictimIndex)
    Call FlushBuffer(VictimIndex)
End Sub

Sub AllMascotasAtacanUser(ByVal victim As Integer, ByVal Maestro As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
    'Reaccion de las mascotas
    Dim iCount As Integer
    
    For iCount = 1 To MAXMASCOTAS
        If UserList(Maestro).MascotasIndex(iCount) > 0 Then
            Npclist(UserList(Maestro).MascotasIndex(iCount)).flags.AttackedBy = UserList(victim).Name
            Npclist(UserList(Maestro).MascotasIndex(iCount)).Movement = TipoAI.NPCDEFENSA
            Npclist(UserList(Maestro).MascotasIndex(iCount)).Hostile = 1
        End If
    Next iCount
End Sub

Public Function PuedeAtacar(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer) As Boolean
'***************************************************
'Autor: Unknown
'Last Modification: 02/04/2010
'Returns true if the AttackerIndex is allowed to attack the VictimIndex.
'24/01/2007 Pablo (ToxicWaste) - Ordeno todo y agrego situacion de Defensa en ciudad Armada y Caos.
'24/02/2009: ZaMa - Los usuarios pueden atacarse entre si.
'02/04/2010: ZaMa - Los armadas no pueden atacar nunca a los ciudas, salvo que esten atacables.
'***************************************************
On Error GoTo Errhandler

    'MUY importante el orden de estos "IF"...
    
    'Estas muerto no podes atacar
    If UserList(AttackerIndex).flags.Muerto = 1 Then
        Call WriteConsoleMsg(AttackerIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
        PuedeAtacar = False
        Exit Function
    End If
    
    'No podes atacar a alguien muerto
    If UserList(VictimIndex).flags.Muerto = 1 Then
        Call WriteConsoleMsg(AttackerIndex, "No puedes atacar a un esp�ritu.", FontTypeNames.FONTTYPE_INFO)
        PuedeAtacar = False
        Exit Function
    End If
    
    ' No podes atacar si estas en consulta
    If UserList(AttackerIndex).flags.EnConsulta Then
        Call WriteConsoleMsg(AttackerIndex, "No puedes atacar usuarios mientras estas en consulta.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    
    ' No podes atacar si esta en consulta
    If UserList(VictimIndex).flags.EnConsulta Then
        Call WriteConsoleMsg(AttackerIndex, "No puedes atacar usuarios mientras estan en consulta.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If

    'Estamos en una Arena? o un trigger zona segura?
    Select Case TriggerZonaPelea(AttackerIndex, VictimIndex)
        Case eTrigger6.TRIGGER6_PERMITE
            PuedeAtacar = (UserList(VictimIndex).flags.AdminInvisible = 0)
            Exit Function
        
        Case eTrigger6.TRIGGER6_PROHIBE
            PuedeAtacar = False
            Exit Function
        
        Case eTrigger6.TRIGGER6_AUSENTE
            'Si no estamos en el Trigger 6 entonces es imposible atacar un gm
            If (UserList(VictimIndex).flags.Privilegios And PlayerType.User) = 0 Then
                If UserList(VictimIndex).flags.AdminInvisible = 0 Then Call WriteConsoleMsg(AttackerIndex, "El ser es demasiado poderoso.", FontTypeNames.FONTTYPE_WARNING)
                PuedeAtacar = False
                Exit Function
            End If
    End Select
    
    'Ataca un ciudadano?
    If Not Criminal(VictimIndex) Then
        ' El atacante es ciuda?
        If Not Criminal(AttackerIndex) Then
            ' El atacante es armada?
            If EsArmada(AttackerIndex) Then
                ' La victima es armada?
                If EsArmada(VictimIndex) Then
                    ' No puede
                    Call WriteConsoleMsg(AttackerIndex, "Los soldados del ej�rcito real tienen prohibido atacar ciudadanos.", FontTypeNames.FONTTYPE_WARNING)
                    Exit Function
                End If
            End If
            
        End If
    ' Ataca a un criminal
    Else
        'Sos un Caos atacando otro caos?
        If EsCaos(VictimIndex) Then
            If EsCaos(AttackerIndex) Then
                Call WriteConsoleMsg(AttackerIndex, "Los miembros de la legi�n oscura tienen prohibido atacarse entre s�.", FontTypeNames.FONTTYPE_WARNING)
                Exit Function
            End If
        End If
    End If
    
    ' Un ciuda es atacado
    If Not Criminal(VictimIndex) Then
        ' Por un armada sin seguro
        If EsArmada(AttackerIndex) Then
            ' No puede
            Call WriteConsoleMsg(AttackerIndex, "Los soldados del ej�rcito real tienen prohibido atacar ciudadanos.", FontTypeNames.FONTTYPE_WARNING)
            PuedeAtacar = False
            Exit Function
        End If
    End If
    
    'Estas en un Mapa Seguro?
    If MapInfo(UserList(VictimIndex).Pos.map).Pk = False Then
        Call WriteConsoleMsg(AttackerIndex, "Esta es una zona segura, aqu� no puedes atacar a otros usuarios.", FontTypeNames.FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function
    End If
    
    'Estas atacando desde un trigger seguro? o tu victima esta en uno asi?
    If MapData(UserList(VictimIndex).Pos.map, UserList(VictimIndex).Pos.X, UserList(VictimIndex).Pos.Y).trigger = eTrigger.ZONASEGURA Or _
        MapData(UserList(AttackerIndex).Pos.map, UserList(AttackerIndex).Pos.X, UserList(AttackerIndex).Pos.Y).trigger = eTrigger.ZONASEGURA Then
        Call WriteConsoleMsg(AttackerIndex, "No puedes pelear aqu�.", FontTypeNames.FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function
    End If
    
    PuedeAtacar = True
Exit Function

Errhandler:
    Call LogError("Error en PuedeAtacar. Error " & Err.Number & " : " & Err.description)
End Function

Public Function PuedeAtacarNPC(ByVal AttackerIndex As Integer, ByVal NpcIndex As Integer, _
                Optional ByVal Paraliza As Boolean = False, Optional ByVal IsPet As Boolean = False) As Boolean
'***************************************************
'Autor: Unknown Author (Original version)
'Returns True if AttackerIndex can attack the NpcIndex
'Last Modification: 16/11/2009
'24/01/2007 Pablo (ToxicWaste) - Orden y correcci�n de ataque sobre una mascota y guardias
'14/08/2007 Pablo (ToxicWaste) - Reescribo y agrego TODOS los casos posibles cosa de usar
'esta funci�n para todo lo referente a ataque a un NPC. Ya sea Magia, F�sico o a Distancia.
'16/11/2009: ZaMa - Agrego validacion de pertenencia de npc.
'02/04/2010: ZaMa - Los armadas ya no peuden atacar npcs no hotiles.
'***************************************************
    
    Dim OwnerUserIndex As Integer
    
    'Estas muerto?
    If UserList(AttackerIndex).flags.Muerto = 1 Then
        Call WriteConsoleMsg(AttackerIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    
    'Sos consejero?
    If UserList(AttackerIndex).flags.Privilegios And PlayerType.Consejero Then
        'No pueden atacar NPC los Consejeros.
        Exit Function
    End If
    
    ' No podes atacar si estas en consulta
    If UserList(AttackerIndex).flags.EnConsulta Then
        Call WriteConsoleMsg(AttackerIndex, "No puedes atacar npcs mientras estas en consulta.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    
    'Es una criatura atacable?
    If Npclist(NpcIndex).Attackable = 0 Then
        Call WriteConsoleMsg(AttackerIndex, "No puedes atacar esta criatura.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    
    'Es valida la distancia a la cual estamos atacando?
    If Distancia(UserList(AttackerIndex).Pos, Npclist(NpcIndex).Pos) >= MAXDISTANCIAARCO Then
       Call WriteConsoleMsg(AttackerIndex, "Est�s muy lejos para disparar.", FontTypeNames.FONTTYPE_FIGHT)
       Exit Function
    End If
    
    'Es una criatura No-Hostil?
    If Npclist(NpcIndex).Hostile = 0 Then
        'Es Guardia del Caos?
        If Npclist(NpcIndex).NPCtype = eNPCType.Guardiascaos Then
            'Lo quiere atacar un caos?
            If EsCaos(AttackerIndex) Then
                Call WriteConsoleMsg(AttackerIndex, "No puedes atacar Guardias del Caos siendo de la legi�n oscura.", FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If
        'Es guardia Real?
        ElseIf Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
            'Lo quiere atacar un Armada?
            If EsArmada(AttackerIndex) Then
                Call WriteConsoleMsg(AttackerIndex, "No puedes atacar Guardias Reales siendo del ej�rcito real.", FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If
    
        'No era un Guardia, asi que es una criatura No-Hostil com�n.
        'Para asegurarnos que no sea una Mascota:
        ElseIf Npclist(NpcIndex).MaestroUser = 0 Then
            'Si sos ciudadano tenes que quitar el seguro para atacarla.
            If Not Criminal(AttackerIndex) Then
                
                ' Si sos armada no podes atacarlo directamente
                If EsArmada(AttackerIndex) Then
                    Call WriteConsoleMsg(AttackerIndex, "Los miembros del ej�rcito real no pueden atacar npcs no hostiles.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                End If

                PuedeAtacarNPC = True
                Exit Function
            End If
        End If
    End If
    
    'Es el NPC mascota de alguien?
    If Npclist(NpcIndex).MaestroUser > 0 Then
        If Not Criminal(Npclist(NpcIndex).MaestroUser) Then
        
            'Es mascota de un Ciudadano.
            If EsArmada(AttackerIndex) Then
                'El atacante es Armada y esta intentando atacar mascota de un Ciudadano
                Call WriteConsoleMsg(AttackerIndex, "Los miembros del ej�rcito real no pueden atacar mascotas de ciudadanos.", FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If
            
        Else
            'Es mascota de un Criminal.
            If EsCaos(Npclist(NpcIndex).MaestroUser) Then
                'Es Caos el Due�o.
                If EsCaos(AttackerIndex) Then
                    'Un Caos intenta atacar una criatura de un Caos. No puede atacar.
                    Call WriteConsoleMsg(AttackerIndex, "Los miembros de la legi�n oscura no pueden atacar mascotas de otros legionarios. ", FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                End If
            End If
        End If
    End If
    
    'Es el Rey Preatoriano?
    If esPretoriano(NpcIndex) = 4 Then
        If pretorianosVivos > 0 Then
            Call WriteConsoleMsg(AttackerIndex, "Debes matar al resto del ej�rcito antes de atacar al rey.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Function
        End If
    End If
    
    PuedeAtacarNPC = True
End Function


Sub CalcularDarExp(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal ElDa�o As Long)
'***************************************************
'Autor: Nacho (Integer)
'Last Modification: 03/09/06 Nacho
'Reescribi gran parte del Sub
'Ahora, da toda la experiencia del npc mientras este vivo.
'***************************************************
    Dim ExpaDar As Long
    
    '[Nacho] Chekeamos que las variables sean validas para las operaciones
    If ElDa�o <= 0 Then ElDa�o = 0
    If Npclist(NpcIndex).Stats.MaxHp <= 0 Then Exit Sub
    If ElDa�o > Npclist(NpcIndex).Stats.MinHp Then ElDa�o = Npclist(NpcIndex).Stats.MinHp
    
    '[Nacho] La experiencia a dar es la porcion de vida quitada * toda la experiencia
    ExpaDar = CLng(ElDa�o * (Npclist(NpcIndex).GiveEXP / Npclist(NpcIndex).Stats.MaxHp))
    If ExpaDar <= 0 Then Exit Sub
    
    '[Nacho] Vamos contando cuanta experiencia sacamos, porque se da toda la que no se dio al user que mata al NPC
            'Esto es porque cuando un elemental ataca, no se da exp, y tambien porque la cuenta que hicimos antes
            'Podria dar un numero fraccionario, esas fracciones se acumulan hasta formar enteros ;P
    If ExpaDar > Npclist(NpcIndex).flags.ExpCount Then
        ExpaDar = Npclist(NpcIndex).flags.ExpCount
        Npclist(NpcIndex).flags.ExpCount = 0
    Else
        Npclist(NpcIndex).flags.ExpCount = Npclist(NpcIndex).flags.ExpCount - ExpaDar
    End If
    
    '[Nacho] Le damos la exp al user
    If ExpaDar > 0 Then
            UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpaDar
            If UserList(UserIndex).Stats.Exp > MAXEXP Then _
                UserList(UserIndex).Stats.Exp = MAXEXP
            Call WriteConsoleMsg(UserIndex, "Has ganado " & ExpaDar & " puntos de experiencia.", FontTypeNames.FONTTYPE_FIGHT)
        End If
        
        Call CheckUserLevel(UserIndex)

End Sub

Public Function TriggerZonaPelea(ByVal Origen As Integer, ByVal Destino As Integer) As eTrigger6
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

'TODO: Pero que rebuscado!!
'Nigo:  Te lo redise�e, pero no te borro el TODO para que lo revises.
On Error GoTo Errhandler
    Dim tOrg As eTrigger
    Dim tDst As eTrigger
    
    tOrg = MapData(UserList(Origen).Pos.map, UserList(Origen).Pos.X, UserList(Origen).Pos.Y).trigger
    tDst = MapData(UserList(Destino).Pos.map, UserList(Destino).Pos.X, UserList(Destino).Pos.Y).trigger
    
    If tOrg = eTrigger.ZONAPELEA Or tDst = eTrigger.ZONAPELEA Then
        If tOrg = tDst Then
            TriggerZonaPelea = TRIGGER6_PERMITE
        Else
            TriggerZonaPelea = TRIGGER6_PROHIBE
        End If
    Else
        TriggerZonaPelea = TRIGGER6_AUSENTE
    End If

Exit Function
Errhandler:
    TriggerZonaPelea = TRIGGER6_AUSENTE
    LogError ("Error en TriggerZonaPelea - " & Err.description)
End Function

Sub UserEnvenena(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim ObjInd As Integer
    
    ObjInd = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
    
    If ObjInd > 0 Then
        If ObjData(ObjInd).proyectil = 1 Then
            ObjInd = UserList(AtacanteIndex).Invent.MunicionEqpObjIndex
        End If
        
        If ObjInd > 0 Then
            If ObjData(ObjInd).Envenena = 1 Then
                
                If RandomNumber(1, 100) < 60 Then
                    UserList(VictimaIndex).flags.Envenenado = 1
                    Call WriteConsoleMsg(VictimaIndex, "��" & UserList(AtacanteIndex).Name & " te ha envenenado!!", FontTypeNames.FONTTYPE_FIGHT)
                    Call WriteConsoleMsg(AtacanteIndex, "��Has envenenado a " & UserList(VictimaIndex).Name & "!!", FontTypeNames.FONTTYPE_FIGHT)
                End If
            End If
        End If
    End If
    
    Call FlushBuffer(VictimaIndex)
End Sub
