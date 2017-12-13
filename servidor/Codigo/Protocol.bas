Attribute VB_Name = "Protocol"
'**************************************************************
' Protocol.bas - Handles all incoming / outgoing messages for client-server communications.
' Uses a binary protocol designed by myself.
'
' Designed and implemented by Juan Mart�n Sotuyo Dodero (Maraxus)
' (juansotuyo@gmail.com)
'**************************************************************

'**************************************************************************
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
'**************************************************************************

''
'Handles all incoming / outgoing packets for client - server communications
'The binary prtocol here used was designed by Juan Mart�n Sotuyo Dodero.
'This is the first time it's used in Alkon, though the second time it's coded.
'This implementation has several enhacements from the first design.
'
' @author Juan Mart�n Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version 1.0.0
' @date 20060517

Option Explicit

''
'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Private Const SEPARATOR As String * 1 = vbNullChar

''
'Auxiliar ByteQueue used as buffer to generate messages not intended to be sent right away.
'Specially usefull to create a message once and send it over to several clients.
Private auxiliarBuffer As New clsByteQueue

Private Enum eGuildForms
    eGFList = 0
    eGFLeaders
    eGFMembers
End Enum: Private GuildForm As eGuildForms

Private Enum ServerPacketID
    Logged                  ' LOGGED
    RemoveDialogs           ' QTDL
    RemoveCharDialog        ' QDL
    NavigateToggle          ' NAVEG
    Disconnect              ' FINOK
    CommerceEnd             ' FINCOMOK
    BankEnd                 ' FINBANOK
    CommerceInit            ' INITCOM
    BankInit                ' INITBANCO
    UserCommerceInit        ' INITCOMUSU
    UserCommerceEnd         ' FINCOMUSUOK
    UserOfferConfirm
    CommerceChat
    ShowBlacksmithForm      ' SFH
    ShowCarpenterForm       ' SFC
    UpdateSta               ' ASS
    UpdateMana              ' ASM
    UpdateHP                ' ASH
    UpdateGold              ' ASG
    UpdateBankGold
    UpdateExp               ' ASE
    ChangeMap               ' CM
    PosUpdate               ' PU
    ChatOverHead            ' ||
    ConsoleMsg              ' || - Beware!! its the same as above, but it was properly splitted
    ShowMessageBox          ' !!
    UserIndexInServer       ' IU
    UserCharIndexInServer   ' IP
    CharacterCreate         ' CC
    CharacterRemove         ' BP
    CharacterChangeNick
    CharacterMove           ' MP, +, * and _ '
    ForceCharMove
    CharacterChange         ' CP
    ObjectCreate            ' HO
    ObjectDelete            ' BO
    BlockPosition           ' BQ
    PlayMIDI                ' TM
    PlayWave                ' TW
    AreaChanged             ' CA
    PauseToggle             ' BKW
    RainToggle              ' LLU
    CreateFX                ' CFX
    UpdateUserStats         ' EST
    WorkRequestTarget       ' T01
    ChangeInventorySlot     ' CSI
    ChangeBankSlot          ' SBO
    ChangeSpellSlot         ' SHS
    Atributes               ' ATR
    BlacksmithWeapons       ' LAH
    BlacksmithArmors        ' LAR
    CarpenterObjects        ' OBR
    RestOK                  ' DOK
    ErrorMsg                ' ERR
    Blind                   ' CEGU
    Dumb                    ' DUMB
    ShowSignal              ' MCAR
    ChangeNPCInventorySlot  ' NPCI
    UpdateHungerAndThirst   ' EHYS
    Fame                    ' FAMA
    MiniStats               ' MEST
    LevelUp                 ' SUNI
    AddForumMsg             ' FMSG
    ShowForumForm           ' MFOR
    SetInvisible            ' NOVER
    DiceRoll                ' DADOS
    MeditateToggle          ' MEDOK
    BlindNoMore             ' NSEGUE
    DumbNoMore              ' NESTUP
    SendSkills              ' SKILLS
    TrainerCreatureList     ' LSTCRI
    ParalizeOK              ' PARADOK
    ShowUserRequest         ' PETICIO
    TradeOK                 ' TRANSOK
    BankOK                  ' BANCOOK
    ChangeUserTradeSlot     ' COMUSUINV
    SendNight               ' NOC
    Pong
    UpdateTagAndStatus
    SpawnList               ' SPL
    ShowSOSForm             ' MSOS
    ShowMOTDEditionForm     ' ZMOTD
    ShowGMPanelForm         ' ABPANEL
    UserNameList            ' LISTUSU
    UpdateStrenghtAndDexterity
    UpdateStrenght
    UpdateDexterity
    AddSlots
    MultiMessage
    StopWorking
    CancelOfferItem
    SubeClase
    ShowFormClase
    EligeFaccion
    ShowFaccionForm
    EligeRecompensa
    ShowRecompensaForm
    SendGuildForm
    GuildFoundation
End Enum

Private Enum ClientPacketID
    LoginExistingChar       'OLOGIN
    ThrowDices              'TIRDAD
    LoginNewChar            'NLOGIN
    Talk                    ';
    Yell                    '-
    Whisper                 '\
    Walk                    'M
    RequestPositionUpdate   'RPU
    Attack                  'AT
    PickUp                  'AG
    RequestAtributes        'ATR
    RequestFame             'FAMA
    RequestSkills           'ESKI
    RequestMiniStats        'FEST
    CommerceEnd             'FINCOM
    UserCommerceEnd         'FINCOMUSU
    UserCommerceConfirm
    CommerceChat
    BankEnd                 'FINBAN
    UserCommerceOk          'COMUSUOK
    UserCommerceReject      'COMUSUNO
    Drop                    'TI
    CastSpell               'LH
    LeftClick               'LC
    DoubleClick             'RC
    Work                    'UK
    UseSpellMacro           'UMH
    UseItem                 'USA
    CraftBlacksmith         'CNS
    CraftCarpenter          'CNC
    WorkLeftClick           'WLC
    SpellInfo               'INFS
    EquipItem               'EQUI
    ChangeHeading           'CHEA
    ModifySkills            'SKSE
    Train                   'ENTR
    CommerceBuy             'COMP
    BankExtractItem         'RETI
    CommerceSell            'VEND
    BankDeposit             'DEPO
    ForumPost               'DEMSG
    MoveSpell               'DESPHE
    MoveBank
    UserCommerceOffer       'OFRECER
    Online                  '/ONLINE
    Quit                    '/SALIR
    RequestAccountState     '/BALANCE
    PetStand                '/QUIETO
    PetFollow               '/ACOMPA�AR
    ReleasePet              '/LIBERAR
    TrainList               '/ENTRENAR
    Rest                    '/DESCANSAR
    Meditate                '/MEDITAR
    Resucitate              '/RESUCITAR
    Heal                    '/CURAR
    Help                    '/AYUDA
    RequestStats            '/EST
    CommerceStart           '/COMERCIAR
    BankStart               '/BOVEDA
    Enlist                  '/ENLISTAR
    Information             '/INFORMACION
    Reward                  '/RECOMPENSA
    UpTime                  '/UPTIME
    Inquiry                 '/ENCUESTA ( with no params )
    CentinelReport          '/CENTINELA
    CouncilMessage          '/BMSG
    RoleMasterRequest       '/ROL
    GMRequest               '/GM
    bugReport               '/_BUG
    ChangeDescription       '/DESC
    Punishments             '/PENAS
    ChangePassword          '/CONTRASE�A
    Gamble                  '/APOSTAR
    InquiryVote             '/ENCUESTA ( with parameters )
    LeaveFaction            '/RETIRAR ( with no arguments )
    BankExtractGold         '/RETIRAR ( with arguments )
    BankDepositGold         '/DEPOSITAR
    Denounce                '/DENUNCIAR
    Ping                    '/PING
    GMCommands
    InitCrafting
    Home
    Consulta
    RequestClaseForm
    EligioClase
    EligioFaccion
    RequestFaccionForm
    RequestRecompensaForm
    EligioRecompensa
    RequestGuildWindow
    GuildFoundate
    GuildConfirmFoundation
    GuildRequest
    moveItem
End Enum

''
'The last existing client packet id.
Private Const LAST_CLIENT_PACKET_ID As Byte = 128

Public Enum FontTypeNames
    FONTTYPE_TALK
    FONTTYPE_FIGHT
    FONTTYPE_WARNING
    FONTTYPE_INFO
    FONTTYPE_INFOBOLD
    FONTTYPE_EJECUCION
    FONTTYPE_PARTY
    FONTTYPE_VENENO
    FONTTYPE_GUILD
    FONTTYPE_SERVER
    FONTTYPE_GUILDMSG
    FONTTYPE_CONSEJO
    FONTTYPE_CONSEJOCAOS
    FONTTYPE_CONSEJOVesA
    FONTTYPE_CONSEJOCAOSVesA
    FONTTYPE_CENTINELA
    FONTTYPE_GMMSG
    FONTTYPE_GM
    FONTTYPE_CITIZEN
    FONTTYPE_CONSE
    FONTTYPE_DIOS
    FONTTYPE_NEWBIE
    FONTTYPE_NEUTRAL
    FONTTYPE_GUILDWELCOME
    FONTTYPE_GUILDLOGIN
End Enum

Public Enum eEditOptions
    eo_Gold = 1
    eo_Experience
    eo_Body
    eo_Head
    eo_CiticensKilled
    eo_CriminalsKilled
    eo_Level
    eo_Class
    eo_Skills
    eo_SkillPointsLeft
    eo_Nobleza
    eo_Asesino
    eo_Sex
    eo_Raza
    eo_addGold
End Enum

''
' Handles incoming data.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleIncomingData(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 01/09/07
'
'***************************************************
On Error Resume Next
    Dim packetID As Byte
    
    packetID = UserList(UserIndex).incomingData.PeekByte()
    
    'Does the packet requires a logged user??
    If Not (packetID = ClientPacketID.ThrowDices _
      Or packetID = ClientPacketID.LoginExistingChar _
      Or packetID = ClientPacketID.LoginNewChar) Then
        
        'Is the user actually logged?
        If Not UserList(UserIndex).flags.UserLogged Then
            Call CloseSocket(UserIndex)
            Exit Sub
        
        'He is logged. Reset idle counter if id is valid.
        ElseIf packetID <= LAST_CLIENT_PACKET_ID Then
            UserList(UserIndex).Counters.IdleCount = 0
        End If
    ElseIf packetID <= LAST_CLIENT_PACKET_ID Then
        UserList(UserIndex).Counters.IdleCount = 0
        
        'Is the user logged?
        If UserList(UserIndex).flags.UserLogged Then
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
    End If
    
    ' Ante cualquier paquete, pierde la proteccion de ser atacado.
    UserList(UserIndex).flags.NoPuedeSerAtacado = False
    
    Select Case packetID
        Case ClientPacketID.LoginExistingChar       'OLOGIN
            Call HandleLoginExistingChar(UserIndex)
        
        Case ClientPacketID.ThrowDices              'TIRDAD
            Call HandleThrowDices(UserIndex)
        
        Case ClientPacketID.LoginNewChar            'NLOGIN
            Call HandleLoginNewChar(UserIndex)
        
        Case ClientPacketID.Talk                    ';
            Call HandleTalk(UserIndex)
        
        Case ClientPacketID.Yell                    '-
            Call HandleYell(UserIndex)
        
        Case ClientPacketID.Whisper                 '\
            Call HandleWhisper(UserIndex)
        
        Case ClientPacketID.Walk                    'M
            Call HandleWalk(UserIndex)
        
        Case ClientPacketID.RequestPositionUpdate   'RPU
            Call HandleRequestPositionUpdate(UserIndex)
        
        Case ClientPacketID.Attack                  'AT
            Call HandleAttack(UserIndex)
        
        Case ClientPacketID.PickUp                  'AG
            Call HandlePickUp(UserIndex)
                
        Case ClientPacketID.RequestAtributes        'ATR
            Call HandleRequestAtributes(UserIndex)
        
        Case ClientPacketID.RequestFame             'FAMA
            Call HandleRequestFame(UserIndex)
        
        Case ClientPacketID.RequestSkills           'ESKI
            Call HandleRequestSkills(UserIndex)
        
        Case ClientPacketID.RequestMiniStats        'FEST
            Call HandleRequestMiniStats(UserIndex)
        
        Case ClientPacketID.CommerceEnd             'FINCOM
            Call HandleCommerceEnd(UserIndex)
            
        Case ClientPacketID.CommerceChat
            Call HandleCommerceChat(UserIndex)
        
        Case ClientPacketID.UserCommerceEnd         'FINCOMUSU
            Call HandleUserCommerceEnd(UserIndex)
            
        Case ClientPacketID.UserCommerceConfirm
            Call HandleUserCommerceConfirm(UserIndex)
        
        Case ClientPacketID.BankEnd                 'FINBAN
            Call HandleBankEnd(UserIndex)
        
        Case ClientPacketID.UserCommerceOk          'COMUSUOK
            Call HandleUserCommerceOk(UserIndex)
        
        Case ClientPacketID.UserCommerceReject      'COMUSUNO
            Call HandleUserCommerceReject(UserIndex)
        
        Case ClientPacketID.Drop                    'TI
            Call HandleDrop(UserIndex)
        
        Case ClientPacketID.CastSpell               'LH
            Call HandleCastSpell(UserIndex)
        
        Case ClientPacketID.LeftClick               'LC
            Call HandleLeftClick(UserIndex)
        
        Case ClientPacketID.DoubleClick             'RC
            Call HandleDoubleClick(UserIndex)
        
        Case ClientPacketID.Work                    'UK
            Call HandleWork(UserIndex)
        
        Case ClientPacketID.UseSpellMacro           'UMH
            Call HandleUseSpellMacro(UserIndex)
        
        Case ClientPacketID.UseItem                 'USA
            Call HandleUseItem(UserIndex)
        
        Case ClientPacketID.CraftBlacksmith         'CNS
            Call HandleCraftBlacksmith(UserIndex)
        
        Case ClientPacketID.CraftCarpenter          'CNC
            Call HandleCraftCarpenter(UserIndex)
        
        Case ClientPacketID.WorkLeftClick           'WLC
            Call HandleWorkLeftClick(UserIndex)
        
        Case ClientPacketID.SpellInfo               'INFS
            Call HandleSpellInfo(UserIndex)
        
        Case ClientPacketID.EquipItem               'EQUI
            Call HandleEquipItem(UserIndex)
        
        Case ClientPacketID.ChangeHeading           'CHEA
            Call HandleChangeHeading(UserIndex)
        
        Case ClientPacketID.ModifySkills            'SKSE
            Call HandleModifySkills(UserIndex)
        
        Case ClientPacketID.Train                   'ENTR
            Call HandleTrain(UserIndex)
        
        Case ClientPacketID.CommerceBuy             'COMP
            Call HandleCommerceBuy(UserIndex)
        
        Case ClientPacketID.BankExtractItem         'RETI
            Call HandleBankExtractItem(UserIndex)
        
        Case ClientPacketID.CommerceSell            'VEND
            Call HandleCommerceSell(UserIndex)
        
        Case ClientPacketID.BankDeposit             'DEPO
            Call HandleBankDeposit(UserIndex)
        
        Case ClientPacketID.ForumPost               'DEMSG
            Call HandleForumPost(UserIndex)
        
        Case ClientPacketID.MoveSpell               'DESPHE
            Call HandleMoveSpell(UserIndex)
            
        Case ClientPacketID.MoveBank
            Call HandleMoveBank(UserIndex)
        
        Case ClientPacketID.UserCommerceOffer       'OFRECER
            Call HandleUserCommerceOffer(UserIndex)
                  
        Case ClientPacketID.Online                  '/ONLINE
            Call HandleOnline(UserIndex)
        
        Case ClientPacketID.Quit                    '/SALIR
            Call HandleQuit(UserIndex)
        
        Case ClientPacketID.RequestAccountState     '/BALANCE
            Call HandleRequestAccountState(UserIndex)
        
        Case ClientPacketID.PetStand                '/QUIETO
            Call HandlePetStand(UserIndex)
        
        Case ClientPacketID.PetFollow               '/ACOMPA�AR
            Call HandlePetFollow(UserIndex)
            
        Case ClientPacketID.ReleasePet              '/LIBERAR
            Call HandleReleasePet(UserIndex)
        
        Case ClientPacketID.TrainList               '/ENTRENAR
            Call HandleTrainList(UserIndex)
        
        Case ClientPacketID.Rest                    '/DESCANSAR
            Call HandleRest(UserIndex)
        
        Case ClientPacketID.Meditate                '/MEDITAR
            Call HandleMeditate(UserIndex)
        
        Case ClientPacketID.Resucitate              '/RESUCITAR
            Call HandleResucitate(UserIndex)
        
        Case ClientPacketID.Heal                    '/CURAR
            Call HandleHeal(UserIndex)
        
        Case ClientPacketID.Help                    '/AYUDA
            Call HandleHelp(UserIndex)
        
        Case ClientPacketID.RequestStats            '/EST
            Call HandleRequestStats(UserIndex)
        
        Case ClientPacketID.CommerceStart           '/COMERCIAR
            Call HandleCommerceStart(UserIndex)
        
        Case ClientPacketID.BankStart               '/BOVEDA
            Call HandleBankStart(UserIndex)
        
        Case ClientPacketID.Enlist                  '/ENLISTAR
            Call HandleEnlist(UserIndex)
        
        Case ClientPacketID.Information             '/INFORMACION
            Call HandleInformation(UserIndex)
        
        Case ClientPacketID.Reward                  '/RECOMPENSA
            Call HandleReward(UserIndex)
        
        Case ClientPacketID.UpTime                  '/UPTIME
            Call HandleUpTime(UserIndex)
        
        Case ClientPacketID.Inquiry                 '/ENCUESTA ( with no params )
            Call HandleInquiry(UserIndex)
        
        Case ClientPacketID.CentinelReport          '/CENTINELA
            Call HandleCentinelReport(UserIndex)
        
        Case ClientPacketID.CouncilMessage          '/BMSG
            Call HandleCouncilMessage(UserIndex)
        
        Case ClientPacketID.RoleMasterRequest       '/ROL
            Call HandleRoleMasterRequest(UserIndex)
        
        Case ClientPacketID.GMRequest               '/GM
            Call HandleGMRequest(UserIndex)
        
        Case ClientPacketID.bugReport               '/_BUG
            Call HandleBugReport(UserIndex)
        
        Case ClientPacketID.ChangeDescription       '/DESC
            Call HandleChangeDescription(UserIndex)
        
        Case ClientPacketID.Punishments             '/PENAS
            Call HandlePunishments(UserIndex)
        
        Case ClientPacketID.ChangePassword          '/CONTRASE�A
            Call HandleChangePassword(UserIndex)
        
        Case ClientPacketID.Gamble                  '/APOSTAR
            Call HandleGamble(UserIndex)
        
        Case ClientPacketID.InquiryVote             '/ENCUESTA ( with parameters )
            Call HandleInquiryVote(UserIndex)
        
        Case ClientPacketID.LeaveFaction            '/RETIRAR ( with no arguments )
            Call HandleLeaveFaction(UserIndex)
        
        Case ClientPacketID.BankExtractGold         '/RETIRAR ( with arguments )
            Call HandleBankExtractGold(UserIndex)
        
        Case ClientPacketID.BankDepositGold         '/DEPOSITAR
            Call HandleBankDepositGold(UserIndex)
        
        Case ClientPacketID.Denounce                '/DENUNCIAR
            Call HandleDenounce(UserIndex)
        
        Case ClientPacketID.Ping                    '/PING
            Call HandlePing(UserIndex)
          
        Case ClientPacketID.GMCommands              'GM Messages
            Call HandleGMCommands(UserIndex)
            
        Case ClientPacketID.InitCrafting
            Call HandleInitCrafting(UserIndex)
        
        Case ClientPacketID.Home
            Call HandleHome(UserIndex)
            
        Case ClientPacketID.Consulta
            Call HandleConsulta(UserIndex)
        
        Case ClientPacketID.RequestClaseForm
            Call HandleRequestClaseForm(UserIndex)
        
        Case ClientPacketID.EligioClase
            Call HandleEligioClase(UserIndex)
            
        Case ClientPacketID.EligioFaccion
            Call HandleEligioFaccion(UserIndex)
        
        Case ClientPacketID.RequestFaccionForm
            Call HandleRequestFaccionForm(UserIndex)
        
        Case ClientPacketID.RequestRecompensaForm
            Call HandleRequestRecompensaForm(UserIndex)
        
        Case ClientPacketID.EligioRecompensa
            Call HandleEligioRecompensa(UserIndex)
        
        Case ClientPacketID.RequestGuildWindow
            Call HandleRequestGuildWindow(UserIndex)
        
        Case ClientPacketID.GuildFoundate
            Call HandleGuildFoundate(UserIndex)
        
        Case ClientPacketID.GuildConfirmFoundation
            Call HandleGuildConfirmFoundation(UserIndex)
        
        Case ClientPacketID.GuildRequest
            Call HandleGuildRequest(UserIndex)
            
        Case ClientPacketID.moveItem
            Call HandleMoveItem(UserIndex)
            
        Case Else
            'ERROR : Abort!
            Call CloseSocket(UserIndex)
    End Select
    
    'Done with this packet, move on to next one or send everything if no more packets found
    If UserList(UserIndex).incomingData.Length > 0 And Err.Number = 0 Then
        Err.Clear
        Call HandleIncomingData(UserIndex)
    
    ElseIf Err.Number <> 0 And Not Err.Number = UserList(UserIndex).incomingData.NotEnoughDataErrCode Then
        'An error ocurred, log it and kick player.
        Call LogError("Error: " & Err.Number & " [" & Err.description & "] " & " Source: " & Err.Source & _
                        vbTab & " HelpFile: " & Err.HelpFile & vbTab & " HelpContext: " & Err.HelpContext & _
                        vbTab & " LastDllError: " & Err.LastDllError & vbTab & _
                        " - UserIndex: " & UserIndex & " - producido al manejar el paquete: " & CStr(packetID))
        Call CloseSocket(UserIndex)
    
    Else
        'Flush buffer - send everything that has been written
        Call FlushBuffer(UserIndex)
    End If
End Sub

Public Sub WriteMultiMessage(ByVal UserIndex As Integer, ByVal MessageIndex As Integer, Optional ByVal Arg1 As Long, Optional ByVal Arg2 As Long, Optional ByVal Arg3 As Long, Optional ByVal StringArg1 As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.MultiMessage)
        Call .WriteByte(MessageIndex)
        
        Select Case MessageIndex
            Case eMessages.DontSeeAnything, eMessages.NPCSwing, eMessages.NPCKillUser, eMessages.BlockedWithShieldUser, _
                eMessages.BlockedWithShieldother, eMessages.UserSwing, eMessages.NobilityLost, _
                eMessages.CantUseWhileMeditating, eMessages.CancelHome, eMessages.FinishHome
            
            Case eMessages.NPCHitUser
                Call .WriteByte(Arg1) 'Target
                Call .WriteInteger(Arg2) 'damage
                
            Case eMessages.UserHitNPC
                Call .WriteLong(Arg1) 'damage
                
            Case eMessages.UserAttackedSwing
                Call .WriteInteger(UserList(Arg1).Char.CharIndex)
                
            Case eMessages.UserHittedByUser
                Call .WriteInteger(Arg1) 'AttackerIndex
                Call .WriteByte(Arg2) 'Target
                Call .WriteInteger(Arg3) 'damage
                
            Case eMessages.UserHittedUser
                Call .WriteInteger(Arg1) 'AttackerIndex
                Call .WriteByte(Arg2) 'Target
                Call .WriteInteger(Arg3) 'damage
                
            Case eMessages.WorkRequestTarget
                Call .WriteByte(Arg1) 'skill
            
            Case eMessages.HaveKilledUser '"Has matado a " & UserList(VictimIndex).name & "!" "Has ganado " & DaExp & " puntos de experiencia."
                Call .WriteInteger(UserList(Arg1).Char.CharIndex) 'VictimIndex
                Call .WriteLong(Arg2) 'Expe
            
            Case eMessages.UserKill '"�" & .name & " te ha matado!"
                Call .WriteInteger(UserList(Arg1).Char.CharIndex) 'AttackerIndex
            
            Case eMessages.EarnExp
            
            Case eMessages.Home
                Call .WriteByte(CByte(Arg1))
                Call .WriteInteger(CInt(Arg2))
                'El cliente no conoce nada sobre nombre de mapas y hogares, por lo tanto _
                 hasta que no se pasen los dats e .INFs al cliente, esto queda as�.
                Call .WriteASCIIString(StringArg1) 'Call .WriteByte(CByte(Arg2))
                
        End Select
    End With
Exit Sub ''

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Private Sub HandleGMCommands(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo Errhandler

Dim Command As Byte

With UserList(UserIndex)
    Call .incomingData.ReadByte
    
    Command = .incomingData.PeekByte
    
    Select Case Command
        Case eGMCommands.GMMessage                '/GMSG
            Call HandleGMMessage(UserIndex)
        
        Case eGMCommands.showName                '/SHOWNAME
            Call HandleShowName(UserIndex)
        
        Case eGMCommands.OnlineRoyalArmy
            Call HandleOnlineRoyalArmy(UserIndex)
        
        Case eGMCommands.OnlineChaosLegion       '/ONLINECAOS
            Call HandleOnlineChaosLegion(UserIndex)
        
        Case eGMCommands.GoNearby                '/IRCERCA
            Call HandleGoNearby(UserIndex)
        
        Case eGMCommands.comment                 '/REM
            Call HandleComment(UserIndex)
        
        Case eGMCommands.serverTime              '/HORA
            Call HandleServerTime(UserIndex)
        
        Case eGMCommands.Where                   '/DONDE
            Call HandleWhere(UserIndex)
        
        Case eGMCommands.CreaturesInMap          '/NENE
            Call HandleCreaturesInMap(UserIndex)
        
        Case eGMCommands.WarpMeToTarget          '/TELEPLOC
            Call HandleWarpMeToTarget(UserIndex)
        
        Case eGMCommands.WarpChar                '/TELEP
            Call HandleWarpChar(UserIndex)
        
        Case eGMCommands.Silence                 '/SILENCIAR
            Call HandleSilence(UserIndex)
        
        Case eGMCommands.SOSShowList             '/SHOW SOS
            Call HandleSOSShowList(UserIndex)
        
        Case eGMCommands.SOSRemove               'SOSDONE
            Call HandleSOSRemove(UserIndex)
        
        Case eGMCommands.GoToChar                '/IRA
            Call HandleGoToChar(UserIndex)
        
        Case eGMCommands.invisible               '/INVISIBLE
            Call HandleInvisible(UserIndex)
        
        Case eGMCommands.GMPanel                 '/PANELGM
            Call HandleGMPanel(UserIndex)
        
        Case eGMCommands.RequestUserList         'LISTUSU
            Call HandleRequestUserList(UserIndex)
        
        Case eGMCommands.Working                 '/TRABAJANDO
            Call HandleWorking(UserIndex)
        
        Case eGMCommands.Hiding                  '/OCULTANDO
            Call HandleHiding(UserIndex)
        
        Case eGMCommands.Jail                    '/CARCEL
            Call HandleJail(UserIndex)
        
        Case eGMCommands.KillNPC                 '/RMATA
            Call HandleKillNPC(UserIndex)
        
        Case eGMCommands.WarnUser                '/ADVERTENCIA
            Call HandleWarnUser(UserIndex)
        
        Case eGMCommands.EditChar                '/MOD
            Call HandleEditChar(UserIndex)
        
        Case eGMCommands.RequestCharInfo         '/INFO
            Call HandleRequestCharInfo(UserIndex)
        
        Case eGMCommands.RequestCharStats        '/STAT
            Call HandleRequestCharStats(UserIndex)
        
        Case eGMCommands.RequestCharGold         '/BAL
            Call HandleRequestCharGold(UserIndex)
        
        Case eGMCommands.RequestCharInventory    '/INV
            Call HandleRequestCharInventory(UserIndex)
        
        Case eGMCommands.RequestCharBank         '/BOV
            Call HandleRequestCharBank(UserIndex)
        
        Case eGMCommands.RequestCharSkills       '/SKILLS
            Call HandleRequestCharSkills(UserIndex)
        
        Case eGMCommands.ReviveChar              '/REVIVIR
            Call HandleReviveChar(UserIndex)
        
        Case eGMCommands.OnlineGM                '/ONLINEGM
            Call HandleOnlineGM(UserIndex)
        
        Case eGMCommands.OnlineMap               '/ONLINEMAP
            Call HandleOnlineMap(UserIndex)
        
        Case eGMCommands.Kick                    '/ECHAR
            Call HandleKick(UserIndex)
        
        Case eGMCommands.Execute                 '/EJECUTAR
            Call HandleExecute(UserIndex)
        
        Case eGMCommands.BanChar                 '/BAN
            Call HandleBanChar(UserIndex)
        
        Case eGMCommands.UnbanChar               '/UNBAN
            Call HandleUnbanChar(UserIndex)
        
        Case eGMCommands.NPCFollow               '/SEGUIR
            Call HandleNPCFollow(UserIndex)
        
        Case eGMCommands.SummonChar              '/SUM
            Call HandleSummonChar(UserIndex)
        
        Case eGMCommands.SpawnListRequest        '/CC
            Call HandleSpawnListRequest(UserIndex)
        
        Case eGMCommands.SpawnCreature           'SPA
            Call HandleSpawnCreature(UserIndex)
        
        Case eGMCommands.ResetNPCInventory       '/RESETINV
            Call HandleResetNPCInventory(UserIndex)
        
        Case eGMCommands.CleanWorld              '/LIMPIAR
            Call HandleCleanWorld(UserIndex)
        
        Case eGMCommands.ServerMessage           '/RMSG
            Call HandleServerMessage(UserIndex)
        
        Case eGMCommands.NickToIP                '/NICK2IP
            Call HandleNickToIP(UserIndex)
        
        Case eGMCommands.IPToNick                '/IP2NICK
            Call HandleIPToNick(UserIndex)
        
        Case eGMCommands.TeleportCreate          '/CT
            Call HandleTeleportCreate(UserIndex)
        
        Case eGMCommands.TeleportDestroy         '/DT
            Call HandleTeleportDestroy(UserIndex)
        
        Case eGMCommands.RainToggle              '/LLUVIA
            Call HandleRainToggle(UserIndex)
        
        Case eGMCommands.SetCharDescription      '/SETDESC
            Call HandleSetCharDescription(UserIndex)
        
        Case eGMCommands.ForceMIDIToMap          '/FORCEMIDIMAP
            Call HanldeForceMIDIToMap(UserIndex)
        
        Case eGMCommands.ForceWAVEToMap          '/FORCEWAVMAP
            Call HandleForceWAVEToMap(UserIndex)
        
        Case eGMCommands.RoyalArmyMessage        '/REALMSG
            Call HandleRoyalArmyMessage(UserIndex)
        
        Case eGMCommands.ChaosLegionMessage      '/CAOSMSG
            Call HandleChaosLegionMessage(UserIndex)
        
        Case eGMCommands.CitizenMessage          '/CIUMSG
            Call HandleCitizenMessage(UserIndex)
        
        Case eGMCommands.CriminalMessage         '/CRIMSG
            Call HandleCriminalMessage(UserIndex)
        
        Case eGMCommands.TalkAsNPC               '/TALKAS
            Call HandleTalkAsNPC(UserIndex)
        
        Case eGMCommands.DestroyAllItemsInArea   '/MASSDEST
            Call HandleDestroyAllItemsInArea(UserIndex)
        
        Case eGMCommands.AcceptRoyalCouncilMember '/ACEPTCONSE
            Call HandleAcceptRoyalCouncilMember(UserIndex)
        
        Case eGMCommands.AcceptChaosCouncilMember '/ACEPTCONSECAOS
            Call HandleAcceptChaosCouncilMember(UserIndex)
        
        Case eGMCommands.ItemsInTheFloor         '/PISO
            Call HandleItemsInTheFloor(UserIndex)
        
        Case eGMCommands.MakeDumb                '/ESTUPIDO
            Call HandleMakeDumb(UserIndex)
        
        Case eGMCommands.MakeDumbNoMore          '/NOESTUPIDO
            Call HandleMakeDumbNoMore(UserIndex)
        
        Case eGMCommands.DumpIPTables            '/DUMPSECURITY
            Call HandleDumpIPTables(UserIndex)
        
        Case eGMCommands.CouncilKick             '/KICKCONSE
            Call HandleCouncilKick(UserIndex)
        
        Case eGMCommands.SetTrigger              '/TRIGGER
            Call HandleSetTrigger(UserIndex)
        
        Case eGMCommands.AskTrigger              '/TRIGGER with no args
            Call HandleAskTrigger(UserIndex)
        
        Case eGMCommands.BannedIPList            '/BANIPLIST
            Call HandleBannedIPList(UserIndex)
        
        Case eGMCommands.BannedIPReload          '/BANIPRELOAD
            Call HandleBannedIPReload(UserIndex)
        
        Case eGMCommands.BanIP                   '/BANIP
            Call HandleBanIP(UserIndex)
        
        Case eGMCommands.UnbanIP                 '/UNBANIP
            Call HandleUnbanIP(UserIndex)
        
        Case eGMCommands.CreateItem              '/ITEM
            Call HandleCreateItem(UserIndex)
        
        Case eGMCommands.DestroyItems            '/DEST
            Call HandleDestroyItems(UserIndex)
        
        Case eGMCommands.ChaosLegionKick         '/NOCAOS
            Call HandleChaosLegionKick(UserIndex)
        
        Case eGMCommands.RoyalArmyKick           '/NOREAL
            Call HandleRoyalArmyKick(UserIndex)
        
        Case eGMCommands.ForceMIDIAll            '/FORCEMIDI
            Call HandleForceMIDIAll(UserIndex)
        
        Case eGMCommands.ForceWAVEAll            '/FORCEWAV
            Call HandleForceWAVEAll(UserIndex)
        
        Case eGMCommands.RemovePunishment        '/BORRARPENA
            Call HandleRemovePunishment(UserIndex)
        
        Case eGMCommands.TileBlockedToggle       '/BLOQ
            Call HandleTileBlockedToggle(UserIndex)
        
        Case eGMCommands.KillNPCNoRespawn        '/MATA
            Call HandleKillNPCNoRespawn(UserIndex)
        
        Case eGMCommands.KillAllNearbyNPCs       '/MASSKILL
            Call HandleKillAllNearbyNPCs(UserIndex)
        
        Case eGMCommands.LastIP                  '/LASTIP
            Call HandleLastIP(UserIndex)
        
        Case eGMCommands.ChangeMOTD              '/MOTDCAMBIA
            Call HandleChangeMOTD(UserIndex)
        
        Case eGMCommands.SetMOTD                 'ZMOTD
            Call HandleSetMOTD(UserIndex)
        
        Case eGMCommands.SystemMessage           '/SMSG
            Call HandleSystemMessage(UserIndex)
        
        Case eGMCommands.CreateNPC               '/ACC
            Call HandleCreateNPC(UserIndex)
        
        Case eGMCommands.CreateNPCWithRespawn    '/RACC
            Call HandleCreateNPCWithRespawn(UserIndex)
        
        Case eGMCommands.NavigateToggle          '/NAVE
            Call HandleNavigateToggle(UserIndex)
        
        Case eGMCommands.ServerOpenToUsersToggle '/RESTRINGIR
            Call HandleServerOpenToUsersToggle(UserIndex)
        
        Case eGMCommands.TurnOffServer           '/APAGAR
            Call HandleTurnOffServer(UserIndex)
        
        Case eGMCommands.ResetFactions           '/RAJAR
            Call HandleResetFactions(UserIndex)
        
        Case eGMCommands.RequestCharMail         '/LASTEMAIL
            Call HandleRequestCharMail(UserIndex)
        
        Case eGMCommands.AlterPassword           '/APASS
            Call HandleAlterPassword(UserIndex)
        
        Case eGMCommands.AlterMail               '/AEMAIL
            Call HandleAlterMail(UserIndex)
        
        Case eGMCommands.AlterName               '/ANAME
            Call HandleAlterName(UserIndex)
        
        Case eGMCommands.ToggleCentinelActivated '/CENTINELAACTIVADO
            Call HandleToggleCentinelActivated(UserIndex)
        
        Case Declaraciones.eGMCommands.DoBackUp  '/DOBACKUP
            Call HandleDoBackUp(UserIndex)
        
        Case eGMCommands.SaveMap                 '/GUARDAMAPA
            Call HandleSaveMap(UserIndex)
        
        Case eGMCommands.ChangeMapInfoPK         '/MODMAPINFO PK
            Call HandleChangeMapInfoPK(UserIndex)
        
        Case eGMCommands.ChangeMapInfoBackup     '/MODMAPINFO BACKUP
            Call HandleChangeMapInfoBackup(UserIndex)
        
        Case eGMCommands.ChangeMapInfoRestricted '/MODMAPINFO RESTRINGIR
            Call HandleChangeMapInfoRestricted(UserIndex)
        
        Case eGMCommands.ChangeMapInfoNoMagic    '/MODMAPINFO MAGIASINEFECTO
            Call HandleChangeMapInfoNoMagic(UserIndex)
        
        Case eGMCommands.ChangeMapInfoLand       '/MODMAPINFO TERRENO
            Call HandleChangeMapInfoLand(UserIndex)
        
        Case eGMCommands.ChangeMapInfoZone       '/MODMAPINFO ZONA
            Call HandleChangeMapInfoZone(UserIndex)
        
        Case eGMCommands.SaveChars               '/GRABAR
            Call HandleSaveChars(UserIndex)
        
        Case eGMCommands.CleanSOS                '/BORRAR SOS
            Call HandleCleanSOS(UserIndex)
        
        Case eGMCommands.ShowServerForm          '/SHOW INT
            Call HandleShowServerForm(UserIndex)
        
        Case eGMCommands.night                   '/NOCHE
            Call HandleNight(UserIndex)
        
        Case eGMCommands.KickAllChars            '/ECHARTODOSPJS
            Call HandleKickAllChars(UserIndex)
        
        Case eGMCommands.ReloadNPCs              '/RELOADNPCS
            Call HandleReloadNPCs(UserIndex)
        
        Case eGMCommands.ReloadServerIni         '/RELOADSINI
            Call HandleReloadServerIni(UserIndex)
        
        Case eGMCommands.ReloadSpells            '/RELOADHECHIZOS
            Call HandleReloadSpells(UserIndex)
        
        Case eGMCommands.ReloadObjects           '/RELOADOBJ
            Call HandleReloadObjects(UserIndex)
        
        Case eGMCommands.ChatColor               '/CHATCOLOR
            Call HandleChatColor(UserIndex)
        
        Case eGMCommands.Ignored                 '/IGNORADO
            Call HandleIgnored(UserIndex)
        
        Case eGMCommands.CheckSlot               '/SLOT
            Call HandleCheckSlot(UserIndex)
        
        Case eGMCommands.SetIniVar               '/SETINIVAR
            Call HandleSetIniVar(UserIndex)
        
        Case eGMCommands.WarpToMap               '/GO
            Call HandleWarpToMap(UserIndex)
        
        Case eGMCommands.StaffMessage            '/STAFF
            Call HandleStaffMessage(UserIndex)
        
        Case eGMCommands.SearchObjs              '/BUSCAR
            Call HandleSearchObjs(UserIndex)
        
        Case eGMCommands.Countdown               '/CUENTA
            Call HandleCountdown(UserIndex)
        
        Case eGMCommands.WinTournament           '/GANOTORNEO
            Call HandleWinTournament(UserIndex)
        
        Case eGMCommands.LoseTournament          '/PERDIOTORNEO
            Call HandleLoseTournament(UserIndex)
            
        Case eGMCommands.WinQuest                '/GANOQUEST
            Call HandleWinQuest(UserIndex)
            
        Case eGMCommands.LoseQuest               '/PERDIOQUEST
            Call HandleLoseQuest(UserIndex)
    End Select
End With

Exit Sub

Errhandler:
    Call LogError("Error en GmCommands. Error: " & Err.Number & " - " & Err.description & _
                  ". Paquete: " & Command)

End Sub

''
' Handles the "Home" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleHome(ByVal UserIndex As Integer)
'***************************************************
'Author: Budi
'Creation Date: 06/01/2010
'Last Modification: 05/06/10
'Pato - 05/06/10: Add the Ucase$ to prevent problems.
'***************************************************
With UserList(UserIndex)
    Call .incomingData.ReadByte
    If .flags.TargetNpcTipo = eNPCType.Gobernador Then
        Call setHome(UserIndex, Npclist(.flags.TargetNPC).Ciudad, .flags.TargetNPC)
    Else
        If .flags.Muerto = 1 Then
        'TODO: checkear que no este en prisi�n
            If .flags.Traveling = 0 Then
                If Ciudades(.Hogar).Map <> .Pos.Map Then
                    Call goHome(UserIndex)
                Else
                    Call WriteConsoleMsg(UserIndex, "Ya te encuentras en tu hogar.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                Call WriteMultiMessage(UserIndex, eMessages.CancelHome)
                .flags.Traveling = 0
                .Counters.goHome = 0
            End If
        Else
            Call WriteConsoleMsg(UserIndex, "Debes estar muerto para utilizar este comando.", FontTypeNames.FONTTYPE_INFO)
        End If
    End If
End With
End Sub

''
' Handles the "LoginExistingChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLoginExistingChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

On Error GoTo Errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte

    Dim UserName As String
    Dim Password As String
    Dim version As String
    
    UserName = Buffer.ReadASCIIString()
    
    Password = Buffer.ReadASCIIString()

    'Convert version number to string
    version = CStr(Buffer.ReadByte()) & "." & CStr(Buffer.ReadByte()) & "." & CStr(Buffer.ReadByte())
    
    If Not AsciiValidos(UserName) Then
        Call WriteErrorMsg(UserIndex, "Nombre inv�lido.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        
        Exit Sub
    End If
    
    If Not PersonajeExiste(UserName) Then
        Call WriteErrorMsg(UserIndex, "El personaje no existe.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        
        Exit Sub
    End If
        
        
        If BANCheck(UserName) Then
            Call WriteErrorMsg(UserIndex, "Se te ha prohibido la entrada a Argentum Online debido a tu mal comportamiento. Puedes consultar el reglamento y el sistema de soporte desde www.argentumonline.com.ar")
        ElseIf Not VersionOK(version) Then
            Call WriteErrorMsg(UserIndex, "Esta versi�n del juego es obsoleta, la versi�n correcta es la " & ULTIMAVERSION & ". La misma se encuentra disponible en www.argentumonline.com.ar")
        Else
            Call ConnectUser(UserIndex, UserName, MD5String(Password))
        End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(Buffer)
    
Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "ThrowDices" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleThrowDices(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    With UserList(UserIndex).Stats
        .UserAtributos(eAtributos.Fuerza) = MaximoInt(15, 13 + RandomNumber(0, 3) + RandomNumber(0, 2))
        .UserAtributos(eAtributos.Agilidad) = MaximoInt(15, 12 + RandomNumber(0, 3) + RandomNumber(0, 3))
        .UserAtributos(eAtributos.Inteligencia) = MaximoInt(16, 13 + RandomNumber(0, 3) + RandomNumber(0, 2))
        .UserAtributos(eAtributos.Carisma) = MaximoInt(15, 12 + RandomNumber(0, 3) + RandomNumber(0, 3))
        .UserAtributos(eAtributos.Constitucion) = 16 + RandomNumber(0, 1) + RandomNumber(0, 1)
    End With
    
    Call WriteDiceRoll(UserIndex)
End Sub

''
' Handles the "LoginNewChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLoginNewChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 15 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

On Error GoTo Errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte

    Dim UserName As String
    Dim Password As String
    Dim version As String
    Dim race As eRaza
    Dim gender As eGenero
    Dim homeland As eCiudad
    Dim Head As Integer
    Dim mail As String
    Dim Skills(1 To NUMSKILLS) As Byte
    
    If PuedeCrearPersonajes = 0 Then
        Call WriteErrorMsg(UserIndex, "La creaci�n de personajes en este servidor se ha deshabilitado.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        
        Exit Sub
    End If
    
    If ServerSoloGMs <> 0 Then
        Call WriteErrorMsg(UserIndex, "Servidor restringido a administradores. Consulte la p�gina oficial o el foro oficial para m�s informaci�n.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        
        Exit Sub
    End If
    
    If aClon.MaxPersonajes(UserList(UserIndex).ip) Then
        Call WriteErrorMsg(UserIndex, "Has creado demasiados personajes.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        
        Exit Sub
    End If
    
    UserName = Buffer.ReadASCIIString()
    

    Password = Buffer.ReadASCIIString()

    'Convert version number to string
    version = CStr(Buffer.ReadByte()) & "." & CStr(Buffer.ReadByte()) & "." & CStr(Buffer.ReadByte())

    
    race = Buffer.ReadByte()
    gender = Buffer.ReadByte()
    Head = Buffer.ReadInteger
    mail = Buffer.ReadASCIIString()
    homeland = Buffer.ReadByte()
    
    Dim i As Long
    
    For i = 1 To NUMSKILLS
        Skills(i) = Buffer.ReadByte
    Next
    
        If Not VersionOK(version) Then
            Call WriteErrorMsg(UserIndex, "Esta versi�n del juego es obsoleta, la versi�n correcta es la " & ULTIMAVERSION & ". La misma se encuentra disponible en www.argentumonline.com.ar")
        Else
            Call ConnectNewUser(UserIndex, UserName, Password, race, gender, mail, homeland, Head, Skills)
        End If

    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(Buffer)
    
Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "Talk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTalk(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 13/01/2010
'15/07/2009: ZaMa - Now invisible admins talk by console.
'23/09/2009: ZaMa - Now invisible admins can't send empty chat.
'13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
    
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Chat As String
        
        Chat = Buffer.ReadASCIIString()
        
        '[Consejeros & GMs]
        If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
            Call LogGM(.Name, "Dijo: " & Chat)
        End If
        
        'I see you....
        If .flags.Oculto > 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            
            If .flags.Navegando = 1 Then
                If .Clase = eClass.Pirata Then
                    ' Pierde la apariencia de fragata fantasmal
                    Call ToogleBoatBody(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "�Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, NingunArma, _
                                        NingunEscudo, NingunCasco)
                End If
            Else
                If .flags.invisible = 0 Then
                    Call UsUaRiOs.SetInvisible(UserIndex, UserList(UserIndex).Char.CharIndex, False)
                    Call WriteConsoleMsg(UserIndex, "�Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        If LenB(Chat) <> 0 Then

            If Not (.flags.AdminInvisible = 1) Then
                If .flags.Muerto = 1 Then
                    Call SendData(SendTarget.ToDeadArea, UserIndex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, CHAT_COLOR_DEAD_CHAR))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, .flags.ChatColor))
                End If
            Else
                If RTrim$(Chat) <> "" Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageConsoleMsg("Gm> " & Chat, FontTypeNames.FONTTYPE_GM))
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "Yell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleYell(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 13/01/2010 (ZaMa)
'15/07/2009: ZaMa - Now invisible admins yell by console.
'13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
    
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Chat As String
        
        Chat = Buffer.ReadASCIIString()
        

        '[Consejeros & GMs]
        If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
            Call LogGM(.Name, "Grito: " & Chat)
        End If
            
        'I see you....
        If .flags.Oculto > 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            
            If .flags.Navegando = 1 Then
                If .Clase = eClass.Pirata Then
                    ' Pierde la apariencia de fragata fantasmal
                    Call ToogleBoatBody(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "�Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, NingunArma, _
                                        NingunEscudo, NingunCasco)
                End If
            Else
                If .flags.invisible = 0 Then
                    Call UsUaRiOs.SetInvisible(UserIndex, .Char.CharIndex, False)
                    Call WriteConsoleMsg(UserIndex, "�Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
            
        If LenB(Chat) <> 0 Then

            If .flags.Privilegios And PlayerType.User Then
                If UserList(UserIndex).flags.Muerto = 1 Then
                    Call SendData(SendTarget.ToDeadArea, UserIndex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, CHAT_COLOR_DEAD_CHAR))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, vbRed))
                End If
            Else
                If Not (.flags.AdminInvisible = 1) Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, CHAT_COLOR_GM_YELL))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageConsoleMsg("Gm> " & Chat, FontTypeNames.FONTTYPE_GM))
                End If
            End If
        End If
        
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "Whisper" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWhisper(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 15/07/2009
'28/05/2009: ZaMa - Now it doesn't appear any message when private talking to an invisible admin
'15/07/2009: ZaMa - Now invisible admins wisper by console.
'***************************************************
    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Chat As String
        Dim targetCharIndex As Integer
        Dim TargetUserIndex As Integer
        Dim targetPriv As PlayerType
        
        targetCharIndex = Buffer.ReadInteger()
        Chat = Buffer.ReadASCIIString()
        
        TargetUserIndex = CharIndexToUserIndex(targetCharIndex)
        
        If .flags.Muerto Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!! Los muertos no pueden comunicarse con el mundo de los vivos. ", FontTypeNames.FONTTYPE_INFO)
        Else
            If TargetUserIndex = INVALID_INDEX Then
                Call WriteConsoleMsg(UserIndex, "Usuario inexistente.", FontTypeNames.FONTTYPE_INFO)
            Else
                targetPriv = UserList(TargetUserIndex).flags.Privilegios
                'A los dioses y admins no vale susurrarles si no sos uno vos mismo (as� no pueden ver si est�n conectados o no)
                If (targetPriv And (PlayerType.Dios Or PlayerType.Admin)) <> 0 And (.flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) <> 0 Then
                    ' Controlamos que no este invisible
                    If UserList(TargetUserIndex).flags.AdminInvisible <> 1 Then
                        Call WriteConsoleMsg(UserIndex, "No puedes susurrarle a los Dioses y Admins.", FontTypeNames.FONTTYPE_INFO)
                    End If
                'A los Consejeros y SemiDioses no vale susurrarles si sos un PJ com�n.
                ElseIf (.flags.Privilegios And PlayerType.User) <> 0 And (Not targetPriv And PlayerType.User) <> 0 Then
                    ' Controlamos que no este invisible
                    If UserList(TargetUserIndex).flags.AdminInvisible <> 1 Then
                        Call WriteConsoleMsg(UserIndex, "No puedes susurrarle a los GMs.", FontTypeNames.FONTTYPE_INFO)
                    End If
                ElseIf Not EstaPCarea(UserIndex, TargetUserIndex) Then
                    Call WriteConsoleMsg(UserIndex, "Est�s muy lejos del usuario.", FontTypeNames.FONTTYPE_INFO)
                
                Else
                    '[Consejeros & GMs]
                    If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
                        Call LogGM(.Name, "Le dijo a '" & UserList(TargetUserIndex).Name & "' " & Chat)
                    End If
                    
                    If LenB(Chat) <> 0 Then

                        If Not (.flags.AdminInvisible = 1) Then
                            Call WriteChatOverHead(UserIndex, Chat, .Char.CharIndex, vbBlue)
                            Call WriteChatOverHead(TargetUserIndex, Chat, .Char.CharIndex, vbBlue)
                            Call FlushBuffer(TargetUserIndex)
                            
                            '[CDT 17-02-2004]
                            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
                                Call SendData(SendTarget.ToAdminsAreaButConsejeros, UserIndex, PrepareMessageChatOverHead("A " & UserList(TargetUserIndex).Name & "> " & Chat, .Char.CharIndex, vbYellow))
                            End If
                        Else
                            Call WriteConsoleMsg(UserIndex, "Susurraste> " & Chat, FontTypeNames.FONTTYPE_GM)
                            If UserIndex <> TargetUserIndex Then Call WriteConsoleMsg(TargetUserIndex, "Gm susurra> " & Chat, FontTypeNames.FONTTYPE_GM)
                            
                            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
                                Call SendData(SendTarget.ToAdminsAreaButConsejeros, UserIndex, PrepareMessageConsoleMsg("Gm dijo a " & UserList(TargetUserIndex).Name & "> " & Chat, FontTypeNames.FONTTYPE_GM))
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "Walk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWalk(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 13/01/2010 (ZaMa)
'11/19/09 Pato - Now the class bandit can walk hidden.
'13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
'***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim dummy As Long
    Dim TempTick As Long
    Dim heading As eHeading
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        heading = .incomingData.ReadByte()
        
        'Prevent SpeedHack
        If .flags.TimesWalk >= 30 Then
            TempTick = GetTickCount And &H7FFFFFFF
            dummy = (TempTick - .flags.StartWalk)
            
            ' 5800 is actually less than what would be needed in perfect conditions to take 30 steps
            '(it's about 193 ms per step against the over 200 needed in perfect conditions)
            If dummy < 5800 Then
                If TempTick - .flags.CountSH > 30000 Then
                    .flags.CountSH = 0
                End If
                
                If Not .flags.CountSH = 0 Then
                    If dummy <> 0 Then _
                        dummy = 126000 \ dummy
                    
                    Call LogHackAttemp("Tramposo SH: " & .Name & " , " & dummy)
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .Name & " ha sido echado por el servidor por posible uso de SH.", FontTypeNames.FONTTYPE_SERVER))
                    Call CloseSocket(UserIndex)
                    
                    Exit Sub
                Else
                    .flags.CountSH = TempTick
                End If
            End If
            .flags.StartWalk = TempTick
            .flags.TimesWalk = 0
        End If
        
        .flags.TimesWalk = .flags.TimesWalk + 1
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        'TODO: Deber�a decirle por consola que no puede?
        'Esta usando el /HOGAR, no se puede mover
        If .flags.Traveling = 1 Then Exit Sub
        
        If .flags.Paralizado = 0 Then
            If .flags.Meditando Then
                'Stop meditating, next action will start movement.
                .flags.Meditando = False
                .Char.FX = 0
                .Char.loops = 0
                
                Call WriteMeditateToggle(UserIndex)
                Call WriteConsoleMsg(UserIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
                
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
            Else
                'Move user
                Call MoveUserChar(UserIndex, heading)
                
                'Stop resting if needed
                If .flags.Descansar Then
                    .flags.Descansar = False
                    
                    Call WriteRestOK(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "Has dejado de descansar.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        Else    'paralized
            If Not .flags.UltimoMensaje = 1 Then
                .flags.UltimoMensaje = 1
                
                Call WriteConsoleMsg(UserIndex, "No puedes moverte porque est�s paralizado.", FontTypeNames.FONTTYPE_INFO)
            End If
            
            .flags.CountSH = 0
        End If
        
        'Can't move while hidden except he is a thief
        If .flags.Oculto = 1 And .flags.AdminInvisible = 0 Then
            If Not (.Clase = eClass.Ladron And .Recompensas(2) = 1) Then
                .flags.Oculto = 0
                .Counters.TiempoOculto = 0
            
                If .flags.Navegando = 1 Then
                    If .Clase = eClass.Pirata Then
                        ' Pierde la apariencia de fragata fantasmal
                        Call ToogleBoatBody(UserIndex)
                        Call WriteConsoleMsg(UserIndex, "�Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                        Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, NingunArma, _
                                        NingunEscudo, NingunCasco)
                    End If
                Else
                    'If not under a spell effect, show char
                    If .flags.invisible = 0 Then
                        Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                        Call UsUaRiOs.SetInvisible(UserIndex, .Char.CharIndex, False)
                    End If
                End If
            End If
        End If
    End With
End Sub

''
' Handles the "RequestPositionUpdate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestPositionUpdate(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    UserList(UserIndex).incomingData.ReadByte
    
    Call WritePosUpdate(UserIndex)
End Sub

''
' Handles the "Attack" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAttack(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 13/01/2010
'Last Modified By: ZaMa
'10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo.
'13/11/2009: ZaMa - Se cancela el estado no atacable al atcar.
'13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'If dead, can't attack
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'If user meditates, can't attack
        If .flags.Meditando Then
            Exit Sub
        End If
        
        'If equiped weapon is ranged, can't attack this way
        If .Invent.WeaponEqpObjIndex > 0 Then
            If ObjData(.Invent.WeaponEqpObjIndex).proyectil = 1 Then
                Call WriteConsoleMsg(UserIndex, "No puedes usar as� este arma.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        'Attack!
        Call UsuarioAtaca(UserIndex)
        
        'Now you can be atacked
        .flags.NoPuedeSerAtacado = False
        
        'I see you...
        If .flags.Oculto > 0 And .flags.AdminInvisible = 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            
            If .flags.Navegando = 1 Then
                If .Clase = eClass.Pirata Then
                    ' Pierde la apariencia de fragata fantasmal
                    Call ToogleBoatBody(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "�Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, NingunArma, _
                                        NingunEscudo, NingunCasco)
                End If
            Else
                If .flags.invisible = 0 Then
                    Call UsUaRiOs.SetInvisible(UserIndex, .Char.CharIndex, False)
                    Call WriteConsoleMsg(UserIndex, "�Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
    End With
End Sub

''
' Handles the "PickUp" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePickUp(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 07/25/09
'02/26/2006: Marco - Agregu� un checkeo por si el usuario trata de agarrar un item mientras comercia.
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'If dead, it can't pick up objects
        If .flags.Muerto = 1 Then Exit Sub
        
        'If user is trading items and attempts to pickup an item, he's cheating, so we kick him.
        If .flags.Comerciando Then Exit Sub
        
        'Lower rank administrators can't pick up items
        If .flags.Privilegios And PlayerType.Consejero Then
            If Not .flags.Privilegios And PlayerType.RoleMaster Then
                Call WriteConsoleMsg(UserIndex, "No puedes tomar ning�n objeto.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        
        Call GetObj(UserIndex)
    End With
End Sub

''
' Handles the "RequestAtributes" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestAtributes(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call WriteAttributes(UserIndex)
End Sub

''
' Handles the "RequestFame" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestFame(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call EnviarFama(UserIndex)
End Sub

''
' Handles the "RequestSkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestSkills(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call WriteSendSkills(UserIndex)
End Sub

''
' Handles the "RequestMiniStats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestMiniStats(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call WriteMiniStats(UserIndex)
End Sub

''
' Handles the "CommerceEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceEnd(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    'User quits commerce mode
    UserList(UserIndex).flags.Comerciando = False
    Call WriteCommerceEnd(UserIndex)
End Sub

''
' Handles the "UserCommerceEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceEnd(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 11/03/2010
'11/03/2010: ZaMa - Le avisa por consola al que cencela que dejo de comerciar.
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Quits commerce mode with user
        If .ComUsu.DestUsu > 0 Then
            If UserList(.ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                Call WriteConsoleMsg(.ComUsu.DestUsu, .Name & " ha dejado de comerciar con vos.", FontTypeNames.FONTTYPE_TALK)
                Call FinComerciarUsu(.ComUsu.DestUsu)
                
                'Send data in the outgoing buffer of the other user
                Call FlushBuffer(.ComUsu.DestUsu)
            End If
        End If
        
        Call FinComerciarUsu(UserIndex)
        Call WriteConsoleMsg(UserIndex, "Has dejado de comerciar.", FontTypeNames.FONTTYPE_TALK)
    End With
End Sub

''
' Handles the "UserCommerceConfirm" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleUserCommerceConfirm(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'
'***************************************************
    
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte

    'Validate the commerce
    If PuedeSeguirComerciando(UserIndex) Then
        'Tell the other user the confirmation of the offer
        Call WriteUserOfferConfirm(UserList(UserIndex).ComUsu.DestUsu)
        UserList(UserIndex).ComUsu.Confirmo = True
    End If
    
End Sub

Private Sub HandleCommerceChat(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 03/12/2009
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
    
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Chat As String
        
        Chat = Buffer.ReadASCIIString()
        
        If LenB(Chat) <> 0 Then
            If PuedeSeguirComerciando(UserIndex) Then

                Chat = UserList(UserIndex).Name & "> " & Chat
                Call WriteCommerceChat(UserIndex, Chat, FontTypeNames.FONTTYPE_PARTY)
                Call WriteCommerceChat(UserList(UserIndex).ComUsu.DestUsu, Chat, FontTypeNames.FONTTYPE_PARTY)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub


''
' Handles the "BankEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankEnd(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'User exits banking mode
        .flags.Comerciando = False
        Call WriteBankEnd(UserIndex)
    End With
End Sub

''
' Handles the "UserCommerceOk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceOk(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    'Trade accepted
    Call AceptarComercioUsu(UserIndex)
End Sub

''
' Handles the "UserCommerceReject" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceReject(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim otherUser As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        otherUser = .ComUsu.DestUsu
        
        'Offer rejected
        If otherUser > 0 Then
            If UserList(otherUser).flags.UserLogged Then
                Call WriteConsoleMsg(otherUser, .Name & " ha rechazado tu oferta.", FontTypeNames.FONTTYPE_TALK)
                Call FinComerciarUsu(otherUser)
                
                'Send data in the outgoing buffer of the other user
                Call FlushBuffer(otherUser)
            End If
        End If
        
        Call WriteConsoleMsg(UserIndex, "Has rechazado la oferta del otro usuario.", FontTypeNames.FONTTYPE_TALK)
        Call FinComerciarUsu(UserIndex)
    End With
End Sub

''
' Handles the "Drop" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDrop(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 07/25/09
'07/25/09: Marco - Agregu� un checkeo para patear a los usuarios que tiran items mientras comercian.
'***************************************************
    If UserList(UserIndex).incomingData.Length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim Slot As Byte
    Dim Amount As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        

        'low rank admins can't drop item. Neither can the dead nor those sailing.
        If .flags.Navegando = 1 Or _
           .flags.Muerto = 1 Or _
           ((.flags.Privilegios And PlayerType.Consejero) <> 0 And (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0) Then Exit Sub

        'If the user is trading, he can't drop items => He's cheating, we kick him.
        If .flags.Comerciando Then Exit Sub

        'Are we dropping gold or other items??
        If Slot = FLAGORO Then
            If Amount > 10000 Then Exit Sub 'Don't drop too much gold

            Call TirarOro(Amount, UserIndex)
            
            Call WriteUpdateGold(UserIndex)
        Else
            'Only drop valid slots
            If Slot <= MAX_INVENTORY_SLOTS And Slot > 0 Then
                If .Invent.Object(Slot).OBJIndex = 0 Then
                    Exit Sub
                End If
                
                Call DropObj(UserIndex, Slot, Amount, .Pos.Map, .Pos.X, .Pos.Y)
            End If
        End If
    End With
End Sub

''
' Handles the "CastSpell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCastSpell(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'13/11/2009: ZaMa - Ahora los npcs pueden atacar al usuario si quizo castear un hechizo
'***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Spell As Byte
        
        Spell = .incomingData.ReadByte()
        
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Now you can be atacked
        .flags.NoPuedeSerAtacado = False
        
        If Spell < 1 Then
            .flags.Hechizo = 0
            Exit Sub
        ElseIf Spell > MAXUSERHECHIZOS Then
            .flags.Hechizo = 0
            Exit Sub
        End If
        
        .flags.Hechizo = .Stats.UserHechizos(Spell)
    End With
End Sub

''
' Handles the "LeftClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLeftClick(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim X As Byte
        Dim Y As Byte
        
        X = .ReadByte()
        Y = .ReadByte()
        
        Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
    End With
End Sub

''
' Handles the "DoubleClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDoubleClick(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim X As Byte
        Dim Y As Byte
        
        X = .ReadByte()
        Y = .ReadByte()
        
        Call Accion(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
    End With
End Sub

''
' Handles the "Work" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWork(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 13/01/2010 (ZaMa)
'13/01/2010: ZaMa - El pirata se puede ocultar en barca
'***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Skill As eSkill
        
        Skill = .incomingData.ReadByte()
        
        If UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        Select Case Skill
            Case Robar, Magia, Domar
                Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, Skill) 'Call WriteWorkRequestTarget(UserIndex, Skill)
            Case Ocultarse
            
                If .flags.EnConsulta Then
                    Call WriteConsoleMsg(UserIndex, "No puedes ocultarte si est�s en consulta.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            
                If .flags.Navegando = 1 Then
                    If .Clase <> eClass.Pirata Then
                        '[CDT 17-02-2004]
                        If Not .flags.UltimoMensaje = 3 Then
                            Call WriteConsoleMsg(UserIndex, "No puedes ocultarte si est�s navegando.", FontTypeNames.FONTTYPE_INFO)
                            .flags.UltimoMensaje = 3
                        End If
                        '[/CDT]
                        Exit Sub
                    End If
                End If
                
                If .flags.Oculto = 1 Then
                    '[CDT 17-02-2004]
                    If Not .flags.UltimoMensaje = 2 Then
                        Call WriteConsoleMsg(UserIndex, "Ya est�s oculto.", FontTypeNames.FONTTYPE_INFO)
                        .flags.UltimoMensaje = 2
                    End If
                    '[/CDT]
                    Exit Sub
                End If
                
                Call DoOcultarse(UserIndex)
        End Select
    End With
End Sub

''
' Handles the "InitCrafting" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInitCrafting(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 29/01/2010
'
'***************************************************
    Dim TotalItems As Long
    Dim ItemsPorCiclo As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        TotalItems = .incomingData.ReadLong
        ItemsPorCiclo = .incomingData.ReadInteger
        
        If TotalItems > 0 Then
            
            .Construir.Cantidad = TotalItems
            .Construir.PorCiclo = MinimoInt(MaxItemsConstruibles(UserIndex), ItemsPorCiclo)
            
        End If
    End With
End Sub

''
' Handles the "UseSpellMacro" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUseSpellMacro(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Call SendData(SendTarget.ToAdmins, UserIndex, PrepareMessageConsoleMsg(.Name & " fue expulsado por Anti-macro de hechizos.", FontTypeNames.FONTTYPE_VENENO))
        Call WriteErrorMsg(UserIndex, "Has sido expulsado por usar macro de hechizos. Recomendamos leer el reglamento sobre el tema macros.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
    End With
End Sub

''
' Handles the "UseItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUseItem(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot As Byte
        
        Slot = .incomingData.ReadByte()
        
        If Slot <= .CurrentInventorySlots And Slot > 0 Then
            If .Invent.Object(Slot).OBJIndex = 0 Then Exit Sub
        End If
        
        If .flags.Meditando Then
            Exit Sub    'The error message should have been provided by the client.
        End If
        
        Call UseInvItem(UserIndex, Slot)
    End With
End Sub

''
' Handles the "CraftBlacksmith" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCraftBlacksmith(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim Item As Integer
        
        Item = .ReadInteger()
        
        If Item < 1 Then Exit Sub
        
        If ObjData(Item).SkHerreria = 0 Then Exit Sub
        
        Call HerreroConstruirItem(UserIndex, Item)
    End With
End Sub

''
' Handles the "CraftCarpenter" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCraftCarpenter(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim Item As Integer
        
        Item = .ReadInteger()
        
        If Item < 1 Then Exit Sub
        
        If ObjData(Item).SkCarpinteria = 0 Then Exit Sub
        
        Call CarpinteroConstruirItem(UserIndex, Item)
    End With
End Sub

''
' Handles the "WorkLeftClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWorkLeftClick(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 14/01/2010 (ZaMa)
'16/11/2009: ZaMa - Agregada la posibilidad de extraer madera elfica.
'12/01/2010: ZaMa - Ahora se admiten armas arrojadizas (proyectiles sin municiones).
'14/01/2010: ZaMa - Ya no se pierden municiones al atacar npcs con due�o.
'***************************************************
    If UserList(UserIndex).incomingData.Length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim X As Byte
        Dim Y As Byte
        Dim Skill As eSkill
        Dim DummyInt As Integer
        Dim tU As Integer   'Target user
        Dim tN As Integer   'Target NPC
        
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
        
        Skill = .incomingData.ReadByte()
        
        
        If .flags.Muerto = 1 Or .flags.Descansar Or .flags.Meditando _
        Or Not InMapBounds(.Pos.Map, X, Y) Then Exit Sub

        If Not InRangoVision(UserIndex, X, Y) Then
            Call WritePosUpdate(UserIndex)
            Exit Sub
        End If
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        Select Case Skill
            Case eSkill.Proyectiles
            
                'Check attack interval
                If Not IntervaloPermiteAtacar(UserIndex, False) Then Exit Sub
                'Check Magic interval
                If Not IntervaloPermiteLanzarSpell(UserIndex, False) Then Exit Sub
                'Check bow's interval
                If Not IntervaloPermiteUsarArcos(UserIndex) Then Exit Sub
                
                Dim Atacked As Boolean
                Atacked = True
                
                'Make sure the item is valid and there is ammo equipped.
                With .Invent
                    ' Tiene arma equipada?
                    If .WeaponEqpObjIndex = 0 Then
                        DummyInt = 1
                    ' En un slot v�lido?
                    ElseIf .WeaponEqpSlot < 1 Or .WeaponEqpSlot > UserList(UserIndex).CurrentInventorySlots Then
                        DummyInt = 1
                    ' Usa munici�n? (Si no la usa, puede ser un arma arrojadiza)
                    ElseIf ObjData(.WeaponEqpObjIndex).Municion = 1 Then
                        ' La municion esta equipada en un slot valido?
                        If .MunicionEqpSlot < 1 Or .MunicionEqpSlot > UserList(UserIndex).CurrentInventorySlots Then
                            DummyInt = 1
                        ' Tiene munici�n?
                        ElseIf .MunicionEqpObjIndex = 0 Then
                            DummyInt = 1
                        ' Son flechas?
                        ElseIf ObjData(.MunicionEqpObjIndex).OBJType <> eOBJType.otFlechas Then
                            DummyInt = 1
                        ' Tiene suficientes?
                        ElseIf .Object(.MunicionEqpSlot).Amount < 1 Then
                            DummyInt = 1
                        End If
                    ' Es un arma de proyectiles?
                    ElseIf ObjData(.WeaponEqpObjIndex).proyectil <> 1 Then
                        DummyInt = 2
                    End If
                    
                    If DummyInt <> 0 Then
                        If DummyInt = 1 Then
                            Call WriteConsoleMsg(UserIndex, "No tienes municiones.", FontTypeNames.FONTTYPE_INFO)
                            
                            Call Desequipar(UserIndex, .WeaponEqpSlot)
                        End If
                        
                        Call Desequipar(UserIndex, .MunicionEqpSlot)
                        Exit Sub
                    End If
                End With
                
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
                
                Call LookatTile(UserIndex, .Pos.Map, X, Y)
                
                tU = .flags.TargetUser
                tN = .flags.TargetNPC
                
                'Validate target
                If tU > 0 Then
                    'Only allow to atack if the other one can retaliate (can see us)
                    If Abs(UserList(tU).Pos.Y - .Pos.Y) > RANGO_VISION_Y Then
                        Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos para atacar.", FontTypeNames.FONTTYPE_WARNING)
                        Exit Sub
                    End If
                    
                    'Prevent from hitting self
                    If tU = UserIndex Then
                        Call WriteConsoleMsg(UserIndex, "�No puedes atacarte a vos mismo!", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Attack!
                    Atacked = UsuarioAtacaUsuario(UserIndex, tU)
                    
                ElseIf tN > 0 Then
                    'Only allow to atack if the other one can retaliate (can see us)
                    If Abs(Npclist(tN).Pos.Y - .Pos.Y) > RANGO_VISION_Y And Abs(Npclist(tN).Pos.X - .Pos.X) > RANGO_VISION_X Then
                        Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos para atacar.", FontTypeNames.FONTTYPE_WARNING)
                        Exit Sub
                    End If
                    
                    'Is it attackable???
                    If Npclist(tN).Attackable <> 0 Then
                        
                        'Attack!
                        Atacked = UsuarioAtacaNpc(UserIndex, tN)
                    End If
                End If
                
                ' Solo pierde la munici�n si pudo atacar al target, o tiro al aire
                If Atacked Then
                    With .Invent
                        ' Tiene equipado arco y flecha?
                        If ObjData(.WeaponEqpObjIndex).Municion = 1 Then
                            DummyInt = .MunicionEqpSlot
                        
                            
                            'Take 1 arrow away - we do it AFTER hitting, since if Ammo Slot is 0 it gives a rt9 and kicks players
                            Call QuitarUserInvItem(UserIndex, DummyInt, 1)
                            
                            If .Object(DummyInt).Amount > 0 Then
                                'QuitarUserInvItem unequips the ammo, so we equip it again
                                .MunicionEqpSlot = DummyInt
                                .MunicionEqpObjIndex = .Object(DummyInt).OBJIndex
                                .Object(DummyInt).Equipped = 1
                            Else
                                .MunicionEqpSlot = 0
                                .MunicionEqpObjIndex = 0
                            End If
                        ' Tiene equipado un arma arrojadiza
                        Else
                            DummyInt = .WeaponEqpSlot
                            
                            'Take 1 knife away
                            Call QuitarUserInvItem(UserIndex, DummyInt, 1)
                            
                            If .Object(DummyInt).Amount > 0 Then
                                'QuitarUserInvItem unequips the weapon, so we equip it again
                                .WeaponEqpSlot = DummyInt
                                .WeaponEqpObjIndex = .Object(DummyInt).OBJIndex
                                .Object(DummyInt).Equipped = 1
                            Else
                                .WeaponEqpSlot = 0
                                .WeaponEqpObjIndex = 0
                            End If
                            
                        End If
                        
                        Call UpdateUserInv(False, UserIndex, DummyInt)
                    End With
               End If
            
            Case eSkill.Magia
                'Check the map allows spells to be casted.
                If MapInfo(.Pos.Map).MagiaSinEfecto > 0 Then
                    Call WriteConsoleMsg(UserIndex, "Una fuerza oscura te impide canalizar tu energ�a.", FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub
                End If
                
                'Target whatever is in that tile
                Call LookatTile(UserIndex, .Pos.Map, X, Y)
                
                'If it's outside range log it and exit
                If Abs(.Pos.X - X) > RANGO_VISION_X Or Abs(.Pos.Y - Y) > RANGO_VISION_Y Then
                    Call LogCheating("Ataque fuera de rango de " & .Name & "(" & .Pos.Map & "/" & .Pos.X & "/" & .Pos.Y & ") ip: " & .ip & " a la posici�n (" & .Pos.Map & "/" & X & "/" & Y & ")")
                    Exit Sub
                End If
                
                'Check bow's interval
                If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub
                
                
                'Check Spell-Hit interval
                If Not IntervaloPermiteGolpeMagia(UserIndex) Then
                    'Check Magic interval
                    If Not IntervaloPermiteLanzarSpell(UserIndex) Then
                        Exit Sub
                    End If
                End If
                
                
                'Check intervals and cast
                If .flags.Hechizo > 0 Then
                    Call LanzarHechizo(.flags.Hechizo, UserIndex)
                    .flags.Hechizo = 0
                Else
                    Call WriteConsoleMsg(UserIndex, "�Primero selecciona el hechizo que quieres lanzar!", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.Pesca
                DummyInt = .Invent.WeaponEqpObjIndex
                If DummyInt = 0 Then Exit Sub
                
                'Check interval
                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                
                'Basado en la idea de Barrin
                'Comentario por Barrin: jah, "basado", caradura ! ^^
                If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 1 Then
                    Call WriteConsoleMsg(UserIndex, "No puedes pescar desde donde te encuentras.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If MapData(.Pos.Map, X, Y).Agua = 1 Then
                    Select Case DummyInt
                        Case CA�A_PESCA
                            Call DoPescar(UserIndex)
                        
                        Case RED_PESCA
                            If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                                Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos para pescar.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                            
                            Call DoPescarRed(UserIndex)
                        
                        Case Else
                            Exit Sub    'Invalid item!
                    End Select
                    
                    'Play sound!
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PESCAR, .Pos.X, .Pos.Y))
                Else
                    Call WriteConsoleMsg(UserIndex, "No hay agua donde pescar. Busca un lago, r�o o mar.", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.Robar
                'Does the map allow us to steal here?
                If MapInfo(.Pos.Map).Pk Then
                    
                    'Check interval
                    If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                    
                    'Target whatever is in that tile
                    Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                    
                    tU = .flags.TargetUser
                    
                    If tU > 0 And tU <> UserIndex Then
                        'Can't steal administrative players
                        If UserList(tU).flags.Privilegios And PlayerType.User Then
                            If UserList(tU).flags.Muerto = 0 Then
                                 If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > IIf(.Clase = eClass.Ladron And .Recompensas(3) = 1, 4, 1) Then
                                     Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                                     Exit Sub
                                 End If
                                 
                                 '17/09/02
                                 'Check the trigger
                                 If MapData(UserList(tU).Pos.Map, X, Y).trigger = eTrigger.ZONASEGURA Then
                                     Call WriteConsoleMsg(UserIndex, "No puedes robar aqu�.", FontTypeNames.FONTTYPE_WARNING)
                                     Exit Sub
                                 End If
                                 
                                 If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONASEGURA Then
                                     Call WriteConsoleMsg(UserIndex, "No puedes robar aqu�.", FontTypeNames.FONTTYPE_WARNING)
                                     Exit Sub
                                 End If
                                 
                                 Call DoRobar(UserIndex, tU)
                            End If
                        End If
                    Else
                        Call WriteConsoleMsg(UserIndex, "�No hay a quien robarle!", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "�No puedes robar en zonas seguras!", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.Talar
                'Check interval
                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                
                If .Invent.WeaponEqpObjIndex = 0 Then
                    Call WriteConsoleMsg(UserIndex, "Deber�as equiparte el hacha.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If .Invent.WeaponEqpObjIndex <> HACHA_LE�ADOR And _
                    .Invent.WeaponEqpObjIndex <> HACHA_LE�A_ELFICA Then
                    ' Podemos llegar ac� si el user equip� el anillo dsp de la U y antes del click
                    Exit Sub
                End If
                
                DummyInt = MapData(.Pos.Map, X, Y).ObjInfo.OBJIndex
                
                If DummyInt > 0 Then
                    If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                        Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Barrin 29/9/03
                    If .Pos.X = X And .Pos.Y = Y Then
                        Call WriteConsoleMsg(UserIndex, "No puedes talar desde all�.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    '�Hay un arbol donde clickeo?
                    If ObjData(DummyInt).OBJType = eOBJType.otArboles And .Invent.WeaponEqpObjIndex = HACHA_LE�ADOR Then
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_TALAR, .Pos.X, .Pos.Y))
                        Call DoTalar(UserIndex)
                    ElseIf ObjData(DummyInt).OBJType = eOBJType.otArbolElfico And .Invent.WeaponEqpObjIndex = HACHA_LE�A_ELFICA Then
                        If .Invent.WeaponEqpObjIndex = HACHA_LE�A_ELFICA Then
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_TALAR, .Pos.X, .Pos.Y))
                            Call DoTalar(UserIndex, True)
                        Else
                            Call WriteConsoleMsg(UserIndex, "El hacha utilizado no es suficientemente poderosa.", FontTypeNames.FONTTYPE_INFO)
                        End If
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "No hay ning�n �rbol ah�.", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.Mineria
                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                                
                If .Invent.WeaponEqpObjIndex = 0 Then Exit Sub
                
                If .Invent.WeaponEqpObjIndex <> PIQUETE_MINERO Then
                    ' Podemos llegar ac� si el user equip� el anillo dsp de la U y antes del click
                    Exit Sub
                End If
                
                'Target whatever is in the tile
                Call LookatTile(UserIndex, .Pos.Map, X, Y)
                
                DummyInt = MapData(.Pos.Map, X, Y).ObjInfo.OBJIndex
                
                If DummyInt > 0 Then
                    'Check distance
                    If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                        Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    DummyInt = MapData(.Pos.Map, X, Y).ObjInfo.OBJIndex 'CHECK
                    '�Hay un yacimiento donde clickeo?
                    If ObjData(DummyInt).OBJType = eOBJType.otYacimiento Then
                        Call DoMineria(UserIndex)
                    Else
                        Call WriteConsoleMsg(UserIndex, "Ah� no hay ning�n yacimiento.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "Ah� no hay ning�n yacimiento.", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.Domar
                'Modificado 25/11/02
                'Optimizado y solucionado el bug de la doma de
                'criaturas hostiles.
                
                'Target whatever is that tile
                Call LookatTile(UserIndex, .Pos.Map, X, Y)
                tN = .flags.TargetNPC
                
                If tN > 0 Then
                    If Npclist(tN).flags.Domable > 0 Then
                        If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                            Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        If LenB(Npclist(tN).flags.AttackedBy) <> 0 Then
                            Call WriteConsoleMsg(UserIndex, "No puedes domar una criatura que est� luchando con un jugador.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        Call DoDomar(UserIndex, tN)
                    Else
                        Call WriteConsoleMsg(UserIndex, "No puedes domar a esa criatura.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "�No hay ninguna criatura all�!", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case FundirMetal    'UGLY!!! This is a constant, not a skill!!
                'Check interval
                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                
                'Check there is a proper item there
                If .flags.TargetObj > 0 Then
                    If ObjData(.flags.TargetObj).OBJType = eOBJType.otFragua Then
                        'Validate other items
                        If .flags.TargetObjInvSlot < 1 Or .flags.TargetObjInvSlot > .CurrentInventorySlots Then
                            Exit Sub
                        End If
                        
                        ''chequeamos que no se zarpe duplicando oro
                        If .Invent.Object(.flags.TargetObjInvSlot).OBJIndex <> .flags.TargetObjInvIndex Then
                            If .Invent.Object(.flags.TargetObjInvSlot).OBJIndex = 0 Or .Invent.Object(.flags.TargetObjInvSlot).Amount = 0 Then
                                Call WriteConsoleMsg(UserIndex, "No tienes m�s minerales.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                            
                            ''FUISTE
                            Call WriteErrorMsg(UserIndex, "Has sido expulsado por el sistema anti cheats.")
                            Call FlushBuffer(UserIndex)
                            Call CloseSocket(UserIndex)
                            Exit Sub
                        End If
                        If ObjData(.flags.TargetObjInvIndex).OBJType = eOBJType.otMinerales Then
                            Call FundirMineral(UserIndex)
                        ElseIf ObjData(.flags.TargetObjInvIndex).OBJType = eOBJType.otWeapon Then
                            Call FundirArmas(UserIndex)
                        End If
                    Else
                        Call WriteConsoleMsg(UserIndex, "Ah� no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "Ah� no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.Herreria
                'Target wehatever is in that tile
                Call LookatTile(UserIndex, .Pos.Map, X, Y)
                
                If .flags.TargetObj > 0 Then
                    If ObjData(.flags.TargetObj).OBJType = eOBJType.otYunque Then
                        Call EnivarArmasConstruibles(UserIndex)
                        Call EnivarArmadurasConstruibles(UserIndex)
                        Call WriteShowBlacksmithForm(UserIndex)
                    Else
                        Call WriteConsoleMsg(UserIndex, "Ah� no hay ning�n yunque.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "Ah� no hay ning�n yunque.", FontTypeNames.FONTTYPE_INFO)
                End If
        End Select
    End With
End Sub



''
' Handles the "SpellInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpellInfo(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim spellSlot As Byte
        Dim Spell As Integer
        
        spellSlot = .incomingData.ReadByte()
        
        'Validate slot
        If spellSlot < 1 Or spellSlot > MAXUSERHECHIZOS Then
            Call WriteConsoleMsg(UserIndex, "�Primero selecciona el hechizo!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate spell in the slot
        Spell = .Stats.UserHechizos(spellSlot)
        If Spell > 0 And Spell < NumeroHechizos + 1 Then
            With Hechizos(Spell)
                'Send information
                Call WriteConsoleMsg(UserIndex, "%%%%%%%%%%%% INFO DEL HECHIZO %%%%%%%%%%%%" & vbCrLf _
                                               & "Nombre:" & .Nombre & vbCrLf _
                                               & "Descripci�n:" & .desc & vbCrLf _
                                               & "Skill requerido: " & .MinSkill & " de magia." & vbCrLf _
                                               & "Man� necesario: " & .ManaRequerido & vbCrLf _
                                               & "Energ�a necesaria: " & .StaRequerido & vbCrLf _
                                               & "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%", FontTypeNames.FONTTYPE_INFO)
            End With
        End If
    End With
End Sub

''
' Handles the "EquipItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEquipItem(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim itemSlot As Byte
        
        itemSlot = .incomingData.ReadByte()
        
        'Dead users can't equip items
        If .flags.Muerto = 1 Then Exit Sub
        
        'Validate item slot
        If itemSlot > .CurrentInventorySlots Or itemSlot < 1 Then Exit Sub
        
        If .Invent.Object(itemSlot).OBJIndex = 0 Then Exit Sub
        
        Call EquiparInvItem(UserIndex, itemSlot)
    End With
End Sub

''
' Handles the "ChangeHeading" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangeHeading(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 06/28/2008
'Last Modified By: NicoNZ
' 10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
' 06/28/2008: NicoNZ - S�lo se puede cambiar si est� inmovilizado.
'***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim heading As eHeading
        Dim posX As Integer
        Dim posY As Integer
                
        heading = .incomingData.ReadByte()
        
        If .flags.Paralizado = 1 And .flags.Inmovilizado = 0 Then
            Select Case heading
                Case eHeading.NORTH
                    posY = -1
                Case eHeading.EAST
                    posX = 1
                Case eHeading.SOUTH
                    posY = 1
                Case eHeading.WEST
                    posX = -1
            End Select
            
                If LegalPos(.Pos.Map, .Pos.X + posX, .Pos.Y + posY, CBool(.flags.Navegando), Not CBool(.flags.Navegando)) Then
                    Exit Sub
                End If
        End If
        
        'Validate heading (VB won't say invalid cast if not a valid index like .Net languages would do... *sigh*)
        If heading > 0 And heading < 5 Then
            .Char.heading = heading
            Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
        End If
    End With
End Sub

''
' Handles the "ModifySkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleModifySkills(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 11/19/09
'11/19/09: Pato - Adapting to new skills system.
'***************************************************
    If UserList(UserIndex).incomingData.Length < 1 + NUMSKILLS Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim i As Long
        Dim Count As Integer
        Dim points(1 To NUMSKILLS) As Byte
        
        'Codigo para prevenir el hackeo de los skills
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        For i = 1 To NUMSKILLS
            points(i) = .incomingData.ReadByte()
            
            If points(i) < 0 Then
                Call LogHackAttemp(.Name & " IP:" & .ip & " trat� de hackear los skills.")
                .Stats.SkillPts = 0
                Call CloseSocket(UserIndex)
                Exit Sub
            End If
            
            Count = Count + points(i)
        Next i
        
        If Count > .Stats.SkillPts Then
            Call LogHackAttemp(.Name & " IP:" & .ip & " trat� de hackear los skills.")
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        
        .Counters.AsignedSkills = MinimoInt(10, .Counters.AsignedSkills + Count)
        
        With .Stats
            For i = 1 To NUMSKILLS
                If points(i) > 0 Then
                    .SkillPts = .SkillPts - points(i)
                    .UserSkills(i) = .UserSkills(i) + points(i)
                    
                    'Client should prevent this, but just in case...
                    If .UserSkills(i) > 100 Then
                        .SkillPts = .SkillPts + .UserSkills(i) - 100
                        .UserSkills(i) = 100
                    End If
                End If
            Next i
        End With
    End With
End Sub

''
' Handles the "Train" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTrain(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim SpawnedNpc As Integer
        Dim PetIndex As Byte
        
        PetIndex = .incomingData.ReadByte()
        
        If .flags.TargetNPC = 0 Then Exit Sub
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Entrenador Then Exit Sub
        
        If Npclist(.flags.TargetNPC).Mascotas < MAXMASCOTASENTRENADOR Then
            If PetIndex > 0 And PetIndex < Npclist(.flags.TargetNPC).NroCriaturas + 1 Then
                'Create the creature
                SpawnedNpc = SpawnNpc(Npclist(.flags.TargetNPC).Criaturas(PetIndex).NpcIndex, Npclist(.flags.TargetNPC).Pos, True, False)
                
                If SpawnedNpc > 0 Then
                    Npclist(SpawnedNpc).MaestroNpc = .flags.TargetNPC
                    Npclist(.flags.TargetNPC).Mascotas = Npclist(.flags.TargetNPC).Mascotas + 1
                End If
            End If
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No puedo traer m�s criaturas, mata las existentes.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
        End If
    End With
End Sub

''
' Handles the "CommerceBuy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceBuy(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot As Byte
        Dim Amount As Integer
        
        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        '�El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
            
        '�El NPC puede comerciar?
        If Npclist(.flags.TargetNPC).Comercia = 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No tengo ning�n inter�s en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
            Exit Sub
        End If
        
        'Only if in commerce mode....
        If Not .flags.Comerciando Then
            Call WriteConsoleMsg(UserIndex, "No est�s comerciando.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'User compra el item
        Call Comercio(eModoComercio.Compra, UserIndex, .flags.TargetNPC, Slot, Amount)
    End With
End Sub

''
' Handles the "BankExtractItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankExtractItem(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot As Byte
        Dim Amount As Integer
        
        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        '�El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
        
        '�Es el banquero?
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then
            Exit Sub
        End If
        
        'User retira el item del slot
        Call UserRetiraItem(UserIndex, Slot, Amount)
    End With
End Sub

''
' Handles the "CommerceSell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceSell(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot As Byte
        Dim Amount As Integer
        
        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        '�El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
        
        '�El NPC puede comerciar?
        If Npclist(.flags.TargetNPC).Comercia = 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No tengo ning�n inter�s en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
            Exit Sub
        End If
        
        'User compra el item del slot
        Call Comercio(eModoComercio.Venta, UserIndex, .flags.TargetNPC, Slot, Amount)
    End With
End Sub

''
' Handles the "BankDeposit" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankDeposit(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot As Byte
        Dim Amount As Integer
        
        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        '�El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
        
        '�El NPC puede comerciar?
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then
            Exit Sub
        End If
        
        'User deposita el item del slot rdata
        Call UserDepositaItem(UserIndex, Slot, Amount)
    End With
End Sub

''
' Handles the "ForumPost" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForumPost(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 02/01/2010
'02/01/2010: ZaMa - Implemento nuevo sistema de foros
'***************************************************
    If UserList(UserIndex).incomingData.Length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim ForumMsgType As eForumMsgType
        
        Dim Title As String
        Dim Post As String
        Dim ForumIndex As Integer
        Dim ForumType As Byte
                
        ForumMsgType = Buffer.ReadByte()
        
        Title = Buffer.ReadASCIIString()
        Post = Buffer.ReadASCIIString()
        
        If .flags.TargetObj > 0 Then
            ForumType = ForumAlignment(ForumMsgType)
            
            Select Case ForumType
            
                Case eForumType.ieGeneral
                    ForumIndex = GetForumIndex(ObjData(.flags.TargetObj).ForoID)
                    
                Case eForumType.ieREAL
                    ForumIndex = GetForumIndex(FORO_REAL_ID)
                    
                Case eForumType.ieCAOS
                    ForumIndex = GetForumIndex(FORO_CAOS_ID)
                    
            End Select
            
            Call AddPost(ForumIndex, Post, .Name, Title, EsAnuncio(ForumMsgType))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "MoveSpell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMoveSpell(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim dir As Integer
        
        If .ReadBoolean() Then
            dir = 1
        Else
            dir = -1
        End If
        
        Call DesplazarHechizo(UserIndex, dir, .ReadByte())
    End With
End Sub

''
' Handles the "MoveBank" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMoveBank(ByVal UserIndex As Integer)
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 06/14/09
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim dir As Integer
        Dim Slot As Byte
        Dim TempItem As Obj
        
        If .ReadBoolean() Then
            dir = 1
        Else
            dir = -1
        End If
        
        Slot = .ReadByte()
    End With
        
    With UserList(UserIndex)
        TempItem.OBJIndex = .BancoInvent.Object(Slot).OBJIndex
        TempItem.Amount = .BancoInvent.Object(Slot).Amount
        
        If dir = 1 Then 'Mover arriba
            .BancoInvent.Object(Slot) = .BancoInvent.Object(Slot - 1)
            .BancoInvent.Object(Slot - 1).OBJIndex = TempItem.OBJIndex
            .BancoInvent.Object(Slot - 1).Amount = TempItem.Amount
        Else 'mover abajo
            .BancoInvent.Object(Slot) = .BancoInvent.Object(Slot + 1)
            .BancoInvent.Object(Slot + 1).OBJIndex = TempItem.OBJIndex
            .BancoInvent.Object(Slot + 1).Amount = TempItem.Amount
        End If
    End With
    
    Call UpdateBanUserInv(True, UserIndex, 0)
    Call UpdateVentanaBanco(UserIndex)

End Sub



''
' Handles the "UserCommerceOffer" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceOffer(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 24/11/2009
'24/11/2009: ZaMa - Nuevo sistema de comercio
'***************************************************
    If UserList(UserIndex).incomingData.Length < 7 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Amount As Long
        Dim Slot As Byte
        Dim tUser As Integer
        Dim OfferSlot As Byte
        Dim OBJIndex As Integer
        
        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadLong()
        OfferSlot = .incomingData.ReadByte()
        
        'Get the other player
        tUser = .ComUsu.DestUsu
        
        ' If he's already confirmed his offer, but now tries to change it, then he's cheating
        If UserList(UserIndex).ComUsu.Confirmo = True Then
            
            ' Finish the trade
            Call FinComerciarUsu(UserIndex)
        
            If tUser <= 0 Or tUser > MaxUsers Then
                Call FinComerciarUsu(tUser)
                Call Protocol.FlushBuffer(tUser)
            End If
        
            Exit Sub
        End If
        
        'If slot is invalid and it's not gold or it's not 0 (Substracting), then ignore it.
        If ((Slot < 0 Or Slot > UserList(UserIndex).CurrentInventorySlots) And Slot <> FLAGORO) Then Exit Sub
        
        'If OfferSlot is invalid, then ignore it.
        If OfferSlot < 1 Or OfferSlot > MAX_OFFER_SLOTS + 1 Then Exit Sub
        
        ' Can be negative if substracted from the offer, but never 0.
        If Amount = 0 Then Exit Sub
        
        'Has he got enough??
        If Slot = FLAGORO Then
            ' Can't offer more than he has
            If Amount > .Stats.GLD - .ComUsu.GoldAmount Then
                Call WriteCommerceChat(UserIndex, "No tienes esa cantidad de oro para agregar a la oferta.", FontTypeNames.FONTTYPE_TALK)
                Exit Sub
            End If
        Else
            'If modifing a filled offerSlot, we already got the objIndex, then we don't need to know it
            If Slot <> 0 Then OBJIndex = .Invent.Object(Slot).OBJIndex
            ' Can't offer more than he has
            If Not TieneObjetos(OBJIndex, _
                TotalOfferItems(OBJIndex, UserIndex) + Amount, UserIndex) Then
                
                Call WriteCommerceChat(UserIndex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
                Exit Sub
            End If
            
            If ItemNewbie(OBJIndex) Then
                Call WriteCancelOfferItem(UserIndex, OfferSlot)
                Exit Sub
            End If
            
            'Don't allow to sell boats if they are equipped (you can't take them off in the water and causes trouble)
            If .flags.Navegando = 1 Then
                If .Invent.BarcoSlot = Slot Then
                    Call WriteCommerceChat(UserIndex, "No puedes vender tu barco mientras lo est�s usando.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub
                End If
            End If
            
            If .Invent.MochilaEqpSlot > 0 Then
                If .Invent.MochilaEqpSlot = Slot Then
                    Call WriteCommerceChat(UserIndex, "No puedes vender tu mochila mientras la est�s usando.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub
                End If
            End If
        End If
        
        
                
        Call AgregarOferta(UserIndex, OfferSlot, OBJIndex, Amount, Slot = FLAGORO)
        
        Call EnviarOferta(tUser, OfferSlot)
    End With
End Sub


''
' Handles the "Online" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnline(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i As Long
    Dim Count As Long
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        For i = 1 To LastUser
            If LenB(UserList(i).Name) <> 0 Then
                If UserList(i).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then _
                    Count = Count + 1
            End If
        Next i
        
        Call WriteConsoleMsg(UserIndex, "N�mero de usuarios: " & CStr(Count), FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "Quit" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleQuit(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 04/15/2008 (NicoNZ)
'If user is invisible, it automatically becomes
'visible before doing the countdown to exit
'04/15/2008 - No se reseteaban lso contadores de invi ni de ocultar. (NicoNZ)
'***************************************************
    Dim tUser As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Paralizado = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes salir estando paralizado.", FontTypeNames.FONTTYPE_WARNING)
            Exit Sub
        End If
        
        'exit secure commerce
        If .ComUsu.DestUsu > 0 Then
            tUser = .ComUsu.DestUsu
            
            If UserList(tUser).flags.UserLogged Then
                If UserList(tUser).ComUsu.DestUsu = UserIndex Then
                    Call WriteConsoleMsg(tUser, "Comercio cancelado por el otro usuario.", FontTypeNames.FONTTYPE_TALK)
                    Call FinComerciarUsu(tUser)
                End If
            End If
            
            Call WriteConsoleMsg(UserIndex, "Comercio cancelado.", FontTypeNames.FONTTYPE_TALK)
            Call FinComerciarUsu(UserIndex)
        End If
        
        Call Cerrar_Usuario(UserIndex)
    End With
End Sub


''
' Handles the "RequestAccountState" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestAccountState(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim earnings As Integer
    Dim Percentage As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't check their accounts
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre �l.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
            Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Select Case Npclist(.flags.TargetNPC).NPCtype
            Case eNPCType.Banquero
                Call WriteChatOverHead(UserIndex, "Tienes " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            
            Case eNPCType.Timbero
                If Not .flags.Privilegios And PlayerType.User Then
                    earnings = Apuestas.Ganancias - Apuestas.Perdidas
                    
                    If earnings >= 0 And Apuestas.Ganancias <> 0 Then
                        Percentage = Int(earnings * 100 / Apuestas.Ganancias)
                    End If
                    
                    If earnings < 0 And Apuestas.Perdidas <> 0 Then
                        Percentage = Int(earnings * 100 / Apuestas.Perdidas)
                    End If
                    
                    Call WriteConsoleMsg(UserIndex, "Entradas: " & Apuestas.Ganancias & " Salida: " & Apuestas.Perdidas & " Ganancia Neta: " & earnings & " (" & Percentage & "%) Jugadas: " & Apuestas.Jugadas, FontTypeNames.FONTTYPE_INFO)
                End If
        End Select
    End With
End Sub

''
' Handles the "PetStand" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePetStand(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre �l.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's his pet
        If Npclist(.flags.TargetNPC).MaestroUser <> UserIndex Then Exit Sub
        
        'Do it!
        Npclist(.flags.TargetNPC).Movement = TipoAI.ESTATICO
        
        Call Expresar(.flags.TargetNPC, UserIndex)
    End With
End Sub

''
' Handles the "PetFollow" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePetFollow(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre �l.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make usre it's the user's pet
        If Npclist(.flags.TargetNPC).MaestroUser <> UserIndex Then Exit Sub
        
        'Do it
        Call FollowAmo(.flags.TargetNPC)
        
        Call Expresar(.flags.TargetNPC, UserIndex)
    End With
End Sub


''
' Handles the "ReleasePet" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleReleasePet(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 18/11/2009
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre �l.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make usre it's the user's pet
        If Npclist(.flags.TargetNPC).MaestroUser <> UserIndex Then Exit Sub
        
        'Do it
        Call QuitarPet(UserIndex, .flags.TargetNPC)
            
    End With
End Sub

''
' Handles the "TrainList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTrainList(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre �l.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's the trainer
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Entrenador Then Exit Sub
        
        Call WriteTrainerCreatureList(UserIndex, .flags.TargetNPC)
    End With
End Sub

''
' Handles the "Rest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRest(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!! Solo puedes usar �tems cuando est�s vivo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If HayOBJarea(.Pos, FOGATA) Then
            Call WriteRestOK(UserIndex)
            
            If Not .flags.Descansar Then
                Call WriteConsoleMsg(UserIndex, "Te acomod�s junto a la fogata y comienzas a descansar.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Te levantas.", FontTypeNames.FONTTYPE_INFO)
            End If
            
            .flags.Descansar = Not .flags.Descansar
        Else
            If .flags.Descansar Then
                Call WriteRestOK(UserIndex)
                Call WriteConsoleMsg(UserIndex, "Te levantas.", FontTypeNames.FONTTYPE_INFO)
                
                .flags.Descansar = False
                Exit Sub
            End If
            
            Call WriteConsoleMsg(UserIndex, "No hay ninguna fogata junto a la cual descansar.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "Meditate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMeditate(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 04/15/08 (NicoNZ)
'Arregl� un bug que mandaba un index de la meditacion diferente
'al que decia el server.
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!! S�lo puedes meditar cuando est�s vivo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Can he meditate?
        If .Stats.MaxMAN = 0 Then
             Call WriteConsoleMsg(UserIndex, "S�lo las clases m�gicas conocen el arte de la meditaci�n.", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
        
        'Admins don't have to wait :D
        If Not .flags.Privilegios And PlayerType.User Then
            .Stats.MinMAN = .Stats.MaxMAN
            Call WriteConsoleMsg(UserIndex, "Man� restaurado.", FontTypeNames.FONTTYPE_VENENO)
            Call WriteUpdateMana(UserIndex)
            Exit Sub
        End If
        
        Call WriteMeditateToggle(UserIndex)
        
        If .flags.Meditando Then _
           Call WriteConsoleMsg(UserIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
        
        .flags.Meditando = Not .flags.Meditando
        
        'Barrin 3/10/03 Tiempo de inicio al meditar
        If .flags.Meditando Then
            .Counters.tInicioMeditar = GetTickCount() And &H7FFFFFFF
            
            Call WriteConsoleMsg(UserIndex, "Te est�s concentrando. En " & Fix(TIEMPO_INICIOMEDITAR / 1000) & " segundos comenzar�s a meditar.", FontTypeNames.FONTTYPE_INFO)
            
            .Char.loops = INFINITE_LOOPS
            
            'Show proper FX according to level
            If .Stats.ELV < 13 Then
                .Char.FX = FXIDs.FXMEDITARCHICO
            
            ElseIf .Stats.ELV < 25 Then
                .Char.FX = FXIDs.FXMEDITARMEDIANO
            
            ElseIf .Stats.ELV < 35 Then
                .Char.FX = FXIDs.FXMEDITARGRANDE
            
            ElseIf .Stats.ELV < 42 Then
                .Char.FX = FXIDs.FXMEDITARXGRANDE
            
            Else
                .Char.FX = FXIDs.FXMEDITARXXGRANDE
            End If
            
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, .Char.FX, INFINITE_LOOPS))
        Else
            .Counters.bPuedeMeditar = False
            
            .Char.FX = 0
            .Char.loops = 0
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
        End If
    End With
End Sub

''
' Handles the "Resucitate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResucitate(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Se asegura que el target es un npc
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre �l.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate NPC and make sure player is dead
        If (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Revividor _
            And (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.ResucitadorNewbie Or Not EsNewbie(UserIndex))) _
            Or .flags.Muerto = 0 Then Exit Sub
        
        'Make sure it's close enough
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "El sacerdote no puede resucitarte debido a que est�s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Call RevivirUsuario(UserIndex)
        Call WriteConsoleMsg(UserIndex, "��Has sido resucitado!!", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "Consulta" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleConsulta(ByVal UserIndex As String)
'***************************************************
'Author: ZaMa
'Last Modification: 01/05/2010
'Habilita/Deshabilita el modo consulta.
'01/05/2010: ZaMa - Agrego validaciones.
'***************************************************
    
    Dim UserConsulta As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        ' Comando exclusivo para gms
        If Not EsGM(UserIndex) Then Exit Sub
        
        UserConsulta = .flags.TargetUser
        
        'Se asegura que el target es un usuario
        If UserConsulta = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un usuario, haz click izquierdo sobre �l.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        ' No podes ponerte a vos mismo en modo consulta.
        If UserConsulta = UserIndex Then Exit Sub
        
        ' No podes estra en consulta con otro gm
        If EsGM(UserConsulta) Then
            Call WriteConsoleMsg(UserIndex, "No puedes iniciar el modo consulta con otro administrador.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Dim UserName As String
        UserName = UserList(UserConsulta).Name
        
        ' Si ya estaba en consulta, termina la consulta
        If UserList(UserConsulta).flags.EnConsulta Then
            Call WriteConsoleMsg(UserIndex, "Has terminado el modo consulta con " & UserName & ".", FontTypeNames.FONTTYPE_INFOBOLD)
            Call WriteConsoleMsg(UserConsulta, "Has terminado el modo consulta.", FontTypeNames.FONTTYPE_INFOBOLD)
            Call LogGM(.Name, "Termino consulta con " & UserName)
            
            UserList(UserConsulta).flags.EnConsulta = False
        
        ' Sino la inicia
        Else
            Call WriteConsoleMsg(UserIndex, "Has iniciado el modo consulta con " & UserName & ".", FontTypeNames.FONTTYPE_INFOBOLD)
            Call WriteConsoleMsg(UserConsulta, "Has iniciado el modo consulta.", FontTypeNames.FONTTYPE_INFOBOLD)
            Call LogGM(.Name, "Inicio consulta con " & UserName)
            
            With UserList(UserConsulta)
                .flags.EnConsulta = True
                
                ' Pierde invi u ocu
                If .flags.invisible = 1 Or .flags.Oculto = 1 Then
                    .flags.Oculto = 0
                    .flags.invisible = 0
                    .Counters.TiempoOculto = 0
                    .Counters.Invisibilidad = 0
                    
                    Call UsUaRiOs.SetInvisible(UserConsulta, UserList(UserConsulta).Char.CharIndex, False)
                End If
            End With
        End If
        
        Call UsUaRiOs.SetConsulatMode(UserConsulta)
    End With

End Sub

''
' Handles the "Heal" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHeal(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Se asegura que el target es un npc
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre �l.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Revividor _
            And Npclist(.flags.TargetNPC).NPCtype <> eNPCType.ResucitadorNewbie) _
            Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "El sacerdote no puede curarte debido a que est�s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        .Stats.MinHp = .Stats.MaxHp
        
        Call WriteUpdateHP(UserIndex)
        
        Call WriteConsoleMsg(UserIndex, "��Has sido curado!!", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "RequestStats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestStats(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call SendUserStatsTxt(UserIndex, UserIndex)
End Sub

''
' Handles the "Help" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHelp(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call SendHelp(UserIndex)
End Sub

''
' Handles the "CommerceStart" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceStart(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i As Integer
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Is it already in commerce mode??
        If .flags.Comerciando Then
            Call WriteConsoleMsg(UserIndex, "Ya est�s comerciando.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC > 0 Then
            'Does the NPC want to trade??
            If Npclist(.flags.TargetNPC).Comercia = 0 Then
                If LenB(Npclist(.flags.TargetNPC).desc) <> 0 Then
                    Call WriteChatOverHead(UserIndex, "No tengo ning�n inter�s en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                End If
                
                Exit Sub
            End If
            
            If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Start commerce....
            Call IniciarComercioNPC(UserIndex)
        '[Alejo]
        ElseIf .flags.TargetUser > 0 Then
            'User commerce...
            'Can he commerce??
            If .flags.Privilegios And PlayerType.Consejero Then
                Call WriteConsoleMsg(UserIndex, "No puedes vender �tems.", FontTypeNames.FONTTYPE_WARNING)
                Exit Sub
            End If
            
            'Is the other one dead??
            If UserList(.flags.TargetUser).flags.Muerto = 1 Then
                Call WriteConsoleMsg(UserIndex, "��No puedes comerciar con los muertos!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Is it me??
            If .flags.TargetUser = UserIndex Then
                Call WriteConsoleMsg(UserIndex, "��No puedes comerciar con vos mismo!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Check distance
            If Distancia(UserList(.flags.TargetUser).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos del usuario.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Is he already trading?? is it with me or someone else??
            If UserList(.flags.TargetUser).flags.Comerciando = True And _
                UserList(.flags.TargetUser).ComUsu.DestUsu <> UserIndex Then
                Call WriteConsoleMsg(UserIndex, "No puedes comerciar con el usuario en este momento.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'Initialize some variables...
            .ComUsu.DestUsu = .flags.TargetUser
            .ComUsu.DestNick = UserList(.flags.TargetUser).Name
            For i = 1 To MAX_OFFER_SLOTS
                .ComUsu.cant(i) = 0
                .ComUsu.Objeto(i) = 0
            Next i
            .ComUsu.GoldAmount = 0
            
            .ComUsu.Acepto = False
            .ComUsu.Confirmo = False
            
            'Rutina para comerciar con otro usuario
            Call IniciarComercioConUsuario(UserIndex, .flags.TargetUser)
        Else
            Call WriteConsoleMsg(UserIndex, "Primero haz click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "BankStart" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankStart(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .flags.Comerciando Then
            Call WriteConsoleMsg(UserIndex, "Ya est�s comerciando.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC > 0 Then
            If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'If it's the banker....
            If Npclist(.flags.TargetNPC).NPCtype = eNPCType.Banquero Then
                Call IniciarDeposito(UserIndex)
            End If
        Else
            Call WriteConsoleMsg(UserIndex, "Primero haz click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "Enlist" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEnlist(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre �l.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble _
            Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(UserIndex, "Debes acercarte m�s.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If ClaseBase(.Clase) Or ClaseTrabajadora(.Clase) Then Exit Sub
        
        Call Enlistar(UserIndex, Npclist(.flags.TargetNPC).flags.Faccion)
    End With
End Sub

''
' Handles the "Information" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInformation(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim Matados As Integer
    Dim NextRecom As Integer
    Dim Diferencia As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre �l.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble _
                Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        
        Select Case .Faccion.Jerarquia
        
            Case 1
                NextRecom = REQUIERE_MATADOS_SEGUNDA
            Case 2
                NextRecom = REQUIERE_MATADOS_TERCERA
            Case 3
                NextRecom = REQUIERE_MATADOS_CUARTA
            Case 4
                Call WriteMultiMessage(UserIndex, eMessages.LastHierarchy, Npclist(.flags.TargetNPC).Char.CharIndex)
                Exit Sub
        End Select
        
        If Npclist(.flags.TargetNPC).flags.Faccion <> .Faccion.Bando Then
            Call WriteChatOverHead(UserIndex, "��No perteneces a nuestras tropas!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            Exit Sub
        End If
        
        If .Faccion.Bando = eFaccion.Real Then
            Matados = .Faccion.Matados(eFaccion.Caos)
            Diferencia = NextRecom - Matados
            
            If Diferencia > 0 Then
                Call WriteChatOverHead(UserIndex, "Tu deber es combatir criminales, mata " & Diferencia & " criminales m�s y te dar� una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            Else
                Call WriteChatOverHead(UserIndex, "Tu deber es combatir criminales, y ya has matado los suficientes como para merecerte una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            End If
        Else
            Matados = .Faccion.Matados(eFaccion.Real)
            Diferencia = NextRecom - Matados
            
            If Diferencia > 0 Then
                Call WriteChatOverHead(UserIndex, "Tu deber es sembrar el caos y la desesperanza, mata " & Diferencia & " ciudadanos m�s y te dar� una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            Else
                Call WriteChatOverHead(UserIndex, "Tu deber es sembrar el caos y la desesperanza, y creo que est�s en condiciones de merecer una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            End If
        End If
    End With
End Sub

''
' Handles the "Reward" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleReward(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre �l.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble _
            Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).flags.Faccion = .Faccion.Bando Then
            Call Recompensado(UserIndex)
        End If
    End With
End Sub

''
' Handles the "RequestMOTD" message.
'
' @param    userIndex The index of the user sending the message.

'Private Sub HandleRequestMOTD(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
'    Call UserList(UserIndex).incomingData.ReadByte
    
'    Call SendMOTD(UserIndex)
'End Sub

''
' Handles the "UpTime" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUpTime(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 01/10/08
'01/10/2008 - Marcos Martinez (ByVal) - Automatic restart removed from the server along with all their assignments and varibles
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Dim time As Long
    Dim UpTimeStr As String
    
    'Get total time in seconds
    time = ((GetTickCount() And &H7FFFFFFF) - tInicioServer) \ 1000
    
    'Get times in dd:hh:mm:ss format
    UpTimeStr = (time Mod 60) & " segundos."
    time = time \ 60
    
    UpTimeStr = (time Mod 60) & " minutos, " & UpTimeStr
    time = time \ 60
    
    UpTimeStr = (time Mod 24) & " horas, " & UpTimeStr
    time = time \ 24
    
    If time = 1 Then
        UpTimeStr = time & " d�a, " & UpTimeStr
    Else
        UpTimeStr = time & " d�as, " & UpTimeStr
    End If
    
    Call WriteConsoleMsg(UserIndex, "Server Online: " & UpTimeStr, FontTypeNames.FONTTYPE_INFO)
End Sub

''
' Handles the "Inquiry" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInquiry(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    ConsultaPopular.SendInfoEncuesta (UserIndex)
End Sub

''
' Handles the "CentinelReport" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCentinelReport(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Call CentinelaCheckClave(UserIndex, .incomingData.ReadInteger())
    End With
End Sub


''
' Handles the "CouncilMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCouncilMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Chat As String
        
        Chat = Buffer.ReadASCIIString()
        
        If LenB(Chat) <> 0 Then

            If .flags.Privilegios And PlayerType.RoyalCouncil Then
                Call SendData(SendTarget.ToConsejo, UserIndex, PrepareMessageConsoleMsg("(Consejero) " & .Name & "> " & Chat, FontTypeNames.FONTTYPE_CONSEJO))
            ElseIf .flags.Privilegios And PlayerType.ChaosCouncil Then
                Call SendData(SendTarget.ToConsejoCaos, UserIndex, PrepareMessageConsoleMsg("(Consejero) " & .Name & "> " & Chat, FontTypeNames.FONTTYPE_CONSEJOCAOS))
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "RoleMasterRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRoleMasterRequest(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim request As String
        
        request = Buffer.ReadASCIIString()
        
        If LenB(request) <> 0 Then
            Call WriteConsoleMsg(UserIndex, "Su solicitud ha sido enviada.", FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToRolesMasters, 0, PrepareMessageConsoleMsg(.Name & " PREGUNTA ROL: " & request, FontTypeNames.FONTTYPE_GUILDMSG))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "GMRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMRequest(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If Not Ayuda.Existe(.Name) Then
            Call WriteConsoleMsg(UserIndex, "El mensaje ha sido entregado, ahora s�lo debes esperar que se desocupe alg�n GM.", FontTypeNames.FONTTYPE_INFO)
            Call Ayuda.Push(.Name)
        Else
            Call Ayuda.Quitar(.Name)
            Call Ayuda.Push(.Name)
            Call WriteConsoleMsg(UserIndex, "Ya hab�as mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "BugReport" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBugReport(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Dim N As Integer
        
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim bugReport As String
        
        bugReport = Buffer.ReadASCIIString()
        
        N = FreeFile
        Open App.path & "\LOGS\BUGs.log" For Append Shared As N
        Print #N, "Usuario:" & .Name & "  Fecha:" & Date & "    Hora:" & time
        Print #N, "BUG:"
        Print #N, bugReport
        Print #N, "########################################################################"
        Close #N
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "ChangeDescription" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangeDescription(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim description As String
        
        description = Buffer.ReadASCIIString()
        
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes cambiar la descripci�n estando muerto.", FontTypeNames.FONTTYPE_INFO)
        Else
            If Not AsciiValidos(description) Then
                Call WriteConsoleMsg(UserIndex, "La descripci�n tiene caracteres inv�lidos.", FontTypeNames.FONTTYPE_INFO)
            Else
                .desc = Trim$(description)
                Call WriteConsoleMsg(UserIndex, "La descripci�n ha cambiado.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "Punishments" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePunishments(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 25/08/2009
'25/08/2009: ZaMa - Now only admins can see other admins' punishment list
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Name As String
        Dim Count As Integer
        
        Name = Buffer.ReadASCIIString()
        
        If LenB(Name) <> 0 Then
            If (InStrB(Name, "\") <> 0) Then
                Name = Replace(Name, "\", "")
            End If
            If (InStrB(Name, "/") <> 0) Then
                Name = Replace(Name, "/", "")
            End If
            If (InStrB(Name, ":") <> 0) Then
                Name = Replace(Name, ":", "")
            End If
            If (InStrB(Name, "|") <> 0) Then
                Name = Replace(Name, "|", "")
            End If
            
            If (EsAdmin(Name) Or EsDios(Name) Or EsSemiDios(Name) Or EsConsejero(Name) Or EsRolesMaster(Name)) And (UserList(UserIndex).flags.Privilegios And PlayerType.User) Then
                Call WriteConsoleMsg(UserIndex, "No puedes ver las penas de los administradores.", FontTypeNames.FONTTYPE_INFO)
            Else
                If FileExist(CharPath & Name & ".chr", vbNormal) Then
                    Count = val(GetVar(CharPath & Name & ".chr", "PENAS", "Cant"))
                    If Count = 0 Then
                        Call WriteConsoleMsg(UserIndex, "Sin prontuario..", FontTypeNames.FONTTYPE_INFO)
                    Else
                        While Count > 0
                            Call WriteConsoleMsg(UserIndex, Count & " - " & GetVar(CharPath & Name & ".chr", "PENAS", "P" & Count), FontTypeNames.FONTTYPE_INFO)
                            Count = Count - 1
                        Wend
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "Personaje """ & Name & """ inexistente.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "ChangePassword" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangePassword(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Creation Date: 10/10/07
'Last Modified By: Rapsodius
'***************************************************

    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        Dim oldPass As String
        Dim newPass As String
        Dim oldPass2 As String
        
        'Remove packet ID
        Call Buffer.ReadByte
        

        oldPass = UCase$(Buffer.ReadASCIIString())
        newPass = UCase$(Buffer.ReadASCIIString())
     
        If LenB(newPass) = 0 Then
            Call WriteConsoleMsg(UserIndex, "Debes especificar una contrase�a nueva, int�ntalo de nuevo.", FontTypeNames.FONTTYPE_INFO)
        Else
            oldPass2 = UCase$(GetVar(CharPath & UserList(UserIndex).Name & ".chr", "INIT", "Password"))
            
            If oldPass2 <> oldPass Then
                Call WriteConsoleMsg(UserIndex, "La contrase�a actual proporcionada no es correcta. La contrase�a no ha sido cambiada, int�ntalo de nuevo.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteVar(CharPath & UserList(UserIndex).Name & ".chr", "INIT", "Password", newPass)
                Call WriteConsoleMsg(UserIndex, "La contrase�a fue cambiada con �xito.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub


''
' Handles the "Gamble" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGamble(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Amount As Integer
        
        Amount = .incomingData.ReadInteger()
        
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
        ElseIf .flags.TargetNPC = 0 Then
            'Validate target NPC
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre �l.", FontTypeNames.FONTTYPE_INFO)
        ElseIf Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        ElseIf Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Timbero Then
            Call WriteChatOverHead(UserIndex, "No tengo ning�n inter�s en apostar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        ElseIf Amount < 1 Then
            Call WriteChatOverHead(UserIndex, "El m�nimo de apuesta es 1 moneda.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        ElseIf Amount > 5000 Then
            Call WriteChatOverHead(UserIndex, "El m�ximo de apuesta es 5000 monedas.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        ElseIf .Stats.GLD < Amount Then
            Call WriteChatOverHead(UserIndex, "No tienes esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        Else
            If RandomNumber(1, 100) <= 47 Then
                .Stats.GLD = .Stats.GLD + Amount
                Call WriteChatOverHead(UserIndex, "�Felicidades! Has ganado " & CStr(Amount) & " monedas de oro.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                
                Apuestas.Perdidas = Apuestas.Perdidas + Amount
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(Apuestas.Perdidas))
            Else
                .Stats.GLD = .Stats.GLD - Amount
                Call WriteChatOverHead(UserIndex, "Lo siento, has perdido " & CStr(Amount) & " monedas de oro.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                
                Apuestas.Ganancias = Apuestas.Ganancias + Amount
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Ganancias", CStr(Apuestas.Ganancias))
            End If
            
            Apuestas.Jugadas = Apuestas.Jugadas + 1
            
            Call WriteVar(DatPath & "apuestas.dat", "Main", "Jugadas", CStr(Apuestas.Jugadas))
            
            Call WriteUpdateGold(UserIndex)
        End If
    End With
End Sub

''
' Handles the "InquiryVote" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInquiryVote(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim opt As Byte
        
        opt = .incomingData.ReadByte()
        
        Call WriteConsoleMsg(UserIndex, ConsultaPopular.doVotar(UserIndex, opt), FontTypeNames.FONTTYPE_GUILD)
    End With
End Sub

''
' Handles the "BankExtractGold" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankExtractGold(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Amount As Long
        
        Amount = .incomingData.ReadLong()
        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
             Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre �l.", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Amount > 0 And Amount <= .Stats.Banco Then
             .Stats.Banco = .Stats.Banco - Amount
             .Stats.GLD = .Stats.GLD + Amount
             Call WriteChatOverHead(UserIndex, "Ten�s " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        Else
            Call WriteChatOverHead(UserIndex, "No tienes esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        End If
        
        Call WriteUpdateGold(UserIndex)
        Call WriteUpdateBankGold(UserIndex)
    End With
End Sub

''
' Handles the "LeaveFaction" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLeaveFaction(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    Dim TalkToKing As Boolean
    Dim TalkToDemon As Boolean
    Dim NpcIndex As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        ' Chequea si habla con el rey o el demonio. Puede salir sin hacerlo, pero si lo hace le reponden los npcs
        NpcIndex = .flags.TargetNPC
        If NpcIndex <> 0 Then
            ' Es rey o domonio?
            If Npclist(NpcIndex).NPCtype = eNPCType.Noble Then
                'Rey?
                If Npclist(NpcIndex).flags.Faccion = eFaccion.Real Then
                    TalkToKing = True
                ' Demonio
                Else
                    TalkToDemon = True
                End If
            End If
        End If
               
        'Quit the Royal Army?
        If .Faccion.Bando = eFaccion.Real Then
            ' Si le pidio al demonio salir de la armada, este le responde.
            If TalkToDemon Then
                Call WriteChatOverHead(UserIndex, "���Sal de aqu� buf�n!!!", _
                                       Npclist(NpcIndex).Char.CharIndex, vbWhite)
            
            Else
                ' Si le pidio al rey salir de la armada, le responde.
                If TalkToKing Then
                    Call WriteChatOverHead(UserIndex, "Ser�s bienvenido a las fuerzas imperiales si deseas regresar.", _
                                           Npclist(NpcIndex).Char.CharIndex, vbWhite)
                End If
                
                Call Expulsar(UserIndex)
                
            End If
        
        'Quit the Chaos Legion?
        ElseIf .Faccion.Bando = eFaccion.Caos Then
            ' Si le pidio al rey salir del caos, le responde.
            If TalkToKing Then
                Call WriteChatOverHead(UserIndex, "���Sal de aqu� maldito criminal!!!", _
                                       Npclist(NpcIndex).Char.CharIndex, vbWhite)
            Else
                ' Si le pidio al demonio salir del caos, este le responde.
                If TalkToDemon Then
                    Call WriteChatOverHead(UserIndex, "Ya volver�s arrastrandote.", _
                                           Npclist(NpcIndex).Char.CharIndex, vbWhite)
                End If
                
                Call Expulsar(UserIndex)
            End If
        ' No es faccionario
        Else
        
            ' Si le hablaba al rey o demonio, le repsonden ellos
            If NpcIndex > 0 Then
                Call WriteChatOverHead(UserIndex, "�No perteneces a ninguna facci�n!", _
                                       Npclist(NpcIndex).Char.CharIndex, vbWhite)
            Else
                Call WriteConsoleMsg(UserIndex, "�No perteneces a ninguna facci�n!", FontTypeNames.FONTTYPE_FIGHT)
            End If
        
        End If
        
    End With
    
End Sub

''
' Handles the "BankDepositGold" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankDepositGold(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Amount As Long
        
        Amount = .incomingData.ReadLong()
        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "��Est�s muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre �l.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Est�s demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then Exit Sub
        
        If Amount > 0 And Amount <= .Stats.GLD Then
            .Stats.Banco = .Stats.Banco + Amount
            .Stats.GLD = .Stats.GLD - Amount
            Call WriteChatOverHead(UserIndex, "Ten�s " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            
            Call WriteUpdateGold(UserIndex)
            Call WriteUpdateBankGold(UserIndex)
        Else
            Call WriteChatOverHead(UserIndex, "No ten�s esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        End If
    End With
End Sub

''
' Handles the "Denounce" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDenounce(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Text As String
        
        Text = Buffer.ReadASCIIString()
        
        If .flags.Silenciado = 0 Then

            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(LCase$(.Name) & " DENUNCIA: " & Text, FontTypeNames.FONTTYPE_GUILDMSG))
            Call WriteConsoleMsg(UserIndex, "Denuncia enviada, espere..", FontTypeNames.FONTTYPE_INFO)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "GMMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 01/08/07
'Last Modification by: (liquid)
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim message As String
        
        message = Buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.User Then
            Call LogGM(.Name, "Mensaje a Gms:" & message)
        
            If LenB(message) <> 0 Then

                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & "> " & message, FontTypeNames.FONTTYPE_GMMSG))
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "ShowName" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleShowName(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            .showName = Not .showName 'Show / Hide the name
            
            Call RefreshCharStatus(UserIndex)
        End If
    End With
End Sub

''
' Handles the "OnlineRoyalArmy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineRoyalArmy(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
    
        Dim i As Long
        Dim list As String

        For i = 1 To MiembrosAlianza.Count
            If UserList(MiembrosAlianza.Item(i)).ConnID <> -1 Then
                If EsArmada(MiembrosAlianza.Item(i)) Then
                    If UserList(MiembrosAlianza.Item(i)).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Or _
                      .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                        list = list & UserList(i).Name & ", "
                    End If
                End If
            End If
        Next i
    End With
    
    If Len(list) > 0 Then
        Call WriteConsoleMsg(UserIndex, "Reales conectados: " & Left$(list, Len(list) - 2), FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(UserIndex, "No hay reales conectados.", FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

''
' Handles the "OnlineChaosLegion" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineChaosLegion(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
    
        Dim i As Long
        Dim list As String

        For i = 1 To MiembrosCaos.Count
            If UserList(MiembrosCaos.Item(i)).ConnID <> -1 Then
                If EsCaos(MiembrosCaos.Item(i)) Then
                    If UserList(MiembrosCaos.Item(i)).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Or _
                      .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                        list = list & UserList(i).Name & ", "
                    End If
                End If
            End If
        Next i
    End With

    If Len(list) > 0 Then
        Call WriteConsoleMsg(UserIndex, "Caos conectados: " & Left$(list, Len(list) - 2), FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(UserIndex, "No hay Caos conectados.", FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

''
' Handles the "GoNearby" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGoNearby(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 01/10/07
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        
        UserName = Buffer.ReadASCIIString()
        
        Dim tIndex As Integer
        Dim X As Long
        Dim Y As Long
        Dim i As Long
        Dim Found As Boolean
        
        tIndex = NameIndex(UserName)
        
        'Check the user has enough powers
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            'Si es dios o Admins no podemos salvo que nosotros tambi�n lo seamos
            If Not (EsDios(UserName) Or EsAdmin(UserName)) Or (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) Then
                If tIndex <= 0 Then 'existe el usuario destino?
                    Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                Else
                    For i = 2 To 5 'esto for sirve ir cambiando la distancia destino
                        For X = UserList(tIndex).Pos.X - i To UserList(tIndex).Pos.X + i
                            For Y = UserList(tIndex).Pos.Y - i To UserList(tIndex).Pos.Y + i
                                If MapData(UserList(tIndex).Pos.Map, X, Y).UserIndex = 0 Then
                                    If LegalPos(UserList(tIndex).Pos.Map, X, Y, True, True) Then
                                        Call WarpUserChar(UserIndex, UserList(tIndex).Pos.Map, X, Y, True)
                                        Call LogGM(.Name, "/IRCERCA " & UserName & " Mapa:" & UserList(tIndex).Pos.Map & " X:" & UserList(tIndex).Pos.X & " Y:" & UserList(tIndex).Pos.Y)
                                        Found = True
                                        Exit For
                                    End If
                                End If
                            Next Y
                            
                            If Found Then Exit For  ' Feo, pero hay que abortar 3 fors sin usar GoTo
                        Next X
                        
                        If Found Then Exit For  ' Feo, pero hay que abortar 3 fors sin usar GoTo
                    Next i
                    
                    'No space found??
                    If Not Found Then
                        Call WriteConsoleMsg(UserIndex, "Todos los lugares est�n ocupados.", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "Comment" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleComment(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim comment As String
        comment = Buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.User Then
            Call LogGM(.Name, "Comentario: " & comment)
            Call WriteConsoleMsg(UserIndex, "Comentario salvado...", FontTypeNames.FONTTYPE_INFO)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "ServerTime" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleServerTime(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 01/08/07
'Last Modification by: (liquid)
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
    
        If .flags.Privilegios And PlayerType.User Then Exit Sub
    
        Call LogGM(.Name, "Hora.")
    End With
    
    Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Hora: " & time & " " & Date, FontTypeNames.FONTTYPE_INFO))
End Sub

''
' Handles the "Where" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWhere(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.User Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                If (UserList(tUser).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) <> 0 Or ((UserList(tUser).flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) <> 0) And (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0) Then
                    Call WriteConsoleMsg(UserIndex, "Ubicaci�n  " & UserName & ": " & UserList(tUser).Pos.Map & ", " & UserList(tUser).Pos.X & ", " & UserList(tUser).Pos.Y & ".", FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.Name, "/Donde " & UserName)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "CreaturesInMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreaturesInMap(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 30/07/06
'Pablo (ToxicWaste): modificaciones generales para simplificar la visualizaci�n.
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Map As Integer
        Dim i, j As Long
        Dim NPCcount1, NPCcount2 As Integer
        Dim NPCcant1() As Integer
        Dim NPCcant2() As Integer
        Dim List1() As String
        Dim List2() As String
        
        Map = .incomingData.ReadInteger()
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        If MapaValido(Map) Then
            For i = 1 To LastNPC
                'VB isn't lazzy, so we put more restrictive condition first to speed up the process
                If Npclist(i).Pos.Map = Map Then
                    '�esta vivo?
                    If Npclist(i).flags.NPCActive And Npclist(i).Hostile = 1 And Npclist(i).Stats.Alineacion = 2 Then
                        If NPCcount1 = 0 Then
                            ReDim List1(0) As String
                            ReDim NPCcant1(0) As Integer
                            NPCcount1 = 1
                            List1(0) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                            NPCcant1(0) = 1
                        Else
                            For j = 0 To NPCcount1 - 1
                                If Left$(List1(j), Len(Npclist(i).Name)) = Npclist(i).Name Then
                                    List1(j) = List1(j) & ", (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                                    NPCcant1(j) = NPCcant1(j) + 1
                                    Exit For
                                End If
                            Next j
                            If j = NPCcount1 Then
                                ReDim Preserve List1(0 To NPCcount1) As String
                                ReDim Preserve NPCcant1(0 To NPCcount1) As Integer
                                NPCcount1 = NPCcount1 + 1
                                List1(j) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                                NPCcant1(j) = 1
                            End If
                        End If
                    Else
                        If NPCcount2 = 0 Then
                            ReDim List2(0) As String
                            ReDim NPCcant2(0) As Integer
                            NPCcount2 = 1
                            List2(0) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                            NPCcant2(0) = 1
                        Else
                            For j = 0 To NPCcount2 - 1
                                If Left$(List2(j), Len(Npclist(i).Name)) = Npclist(i).Name Then
                                    List2(j) = List2(j) & ", (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                                    NPCcant2(j) = NPCcant2(j) + 1
                                    Exit For
                                End If
                            Next j
                            If j = NPCcount2 Then
                                ReDim Preserve List2(0 To NPCcount2) As String
                                ReDim Preserve NPCcant2(0 To NPCcount2) As Integer
                                NPCcount2 = NPCcount2 + 1
                                List2(j) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                                NPCcant2(j) = 1
                            End If
                        End If
                    End If
                End If
            Next i
            
            Call WriteConsoleMsg(UserIndex, "Npcs Hostiles en mapa: ", FontTypeNames.FONTTYPE_WARNING)
            If NPCcount1 = 0 Then
                Call WriteConsoleMsg(UserIndex, "No hay NPCS Hostiles.", FontTypeNames.FONTTYPE_INFO)
            Else
                For j = 0 To NPCcount1 - 1
                    Call WriteConsoleMsg(UserIndex, NPCcant1(j) & " " & List1(j), FontTypeNames.FONTTYPE_INFO)
                Next j
            End If
            Call WriteConsoleMsg(UserIndex, "Otros Npcs en mapa: ", FontTypeNames.FONTTYPE_WARNING)
            If NPCcount2 = 0 Then
                Call WriteConsoleMsg(UserIndex, "No hay m�s NPCS.", FontTypeNames.FONTTYPE_INFO)
            Else
                For j = 0 To NPCcount2 - 1
                    Call WriteConsoleMsg(UserIndex, NPCcant2(j) & " " & List2(j), FontTypeNames.FONTTYPE_INFO)
                Next j
            End If
            Call LogGM(.Name, "Numero enemigos en mapa " & Map)
        End If
    End With
End Sub

''
' Handles the "WarpMeToTarget" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarpMeToTarget(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 26/03/09
'26/03/06: ZaMa - Chequeo que no se teletransporte donde haya un char o npc
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim X As Integer
        Dim Y As Integer
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        X = .flags.TargetX
        Y = .flags.TargetY
        
        Call FindLegalPos(UserIndex, .flags.TargetMap, X, Y)
        Call WarpUserChar(UserIndex, .flags.TargetMap, X, Y, True)
        Call LogGM(.Name, "/TELEPLOC a x:" & .flags.TargetX & " Y:" & .flags.TargetY & " Map:" & .Pos.Map)
    End With
End Sub

''
' Handles the "WarpChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarpChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 26/03/2009
'26/03/2009: ZaMa -  Chequeo que no se teletransporte a un tile donde haya un char o npc.
'***************************************************
    If UserList(UserIndex).incomingData.Length < 7 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim Map As Integer
        Dim X As Integer
        Dim Y As Integer
        Dim tUser As Integer
        
        UserName = Buffer.ReadASCIIString()
        Map = Buffer.ReadInteger()
        X = Buffer.ReadByte()
        Y = Buffer.ReadByte()
        
        If Not .flags.Privilegios And PlayerType.User Then
            If MapaValido(Map) And LenB(UserName) <> 0 Then
                If UCase$(UserName) <> "YO" Then
                    If Not .flags.Privilegios And PlayerType.Consejero Then
                        tUser = NameIndex(UserName)
                    End If
                Else
                    tUser = UserIndex
                End If
            
                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                ElseIf InMapBounds(Map, X, Y) Then
                    Call FindLegalPos(tUser, Map, X, Y)
                    Call WarpUserChar(tUser, Map, X, Y, True, True)
                    Call WriteConsoleMsg(UserIndex, UserList(tUser).Name & " transportado.", FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.Name, "Transport� a " & UserList(tUser).Name & " hacia " & "Mapa" & Map & " X:" & X & " Y:" & Y)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
        
    Call LogError("Handle WarpChar: " & Err.description)
End Sub

''
' Handles the "Silence" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSilence(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.User Then
            tUser = NameIndex(UserName)
        
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                If UserList(tUser).flags.Silenciado = 0 Then
                    UserList(tUser).flags.Silenciado = 1
                    Call WriteConsoleMsg(UserIndex, "Usuario silenciado.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteShowMessageBox(tUser, "Estimado usuario, ud. ha sido silenciado por los administradores. Sus denuncias ser�n ignoradas por el servidor de aqu� en m�s. Utilice /GM para contactar un administrador.")
                    Call LogGM(.Name, "/silenciar " & UserList(tUser).Name)
                
                    'Flush the other user's buffer
                    Call FlushBuffer(tUser)
                Else
                    UserList(tUser).flags.Silenciado = 0
                    Call WriteConsoleMsg(UserIndex, "Usuario des silenciado.", FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.Name, "/DESsilenciar " & UserList(tUser).Name)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "SOSShowList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSOSShowList(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        Call WriteShowSOSForm(UserIndex)
    End With
End Sub

''
' Handles the "SOSRemove" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSOSRemove(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        UserName = Buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.User Then _
            Call Ayuda.Quitar(UserName)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "GoToChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGoToChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 26/03/2009
'26/03/2009: ZaMa -  Chequeo que no se teletransporte a un tile donde haya un char o npc.
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim X As Integer
        Dim Y As Integer
        
        UserName = Buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            'Si es dios o Admins no podemos salvo que nosotros tambi�n lo seamos
            If Not (EsDios(UserName) Or EsAdmin(UserName)) Or (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Then
                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                Else
                    X = UserList(tUser).Pos.X
                    Y = UserList(tUser).Pos.Y + 1
                    Call FindLegalPos(UserIndex, UserList(tUser).Pos.Map, X, Y)
                    
                    Call WarpUserChar(UserIndex, UserList(tUser).Pos.Map, X, Y, True)
                    
                    If .flags.AdminInvisible = 0 Then
                        Call WriteConsoleMsg(tUser, .Name & " se ha trasportado hacia donde te encuentras.", FontTypeNames.FONTTYPE_INFO)
                        Call FlushBuffer(tUser)
                    End If
                    
                    Call LogGM(.Name, "/IRA " & UserName & " Mapa:" & UserList(tUser).Pos.Map & " X:" & UserList(tUser).Pos.X & " Y:" & UserList(tUser).Pos.Y)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "Invisible" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInvisible(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        Call DoAdminInvisible(UserIndex)
        Call LogGM(.Name, "/INVISIBLE")
    End With
End Sub

''
' Handles the "GMPanel" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMPanel(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        Call WriteShowGMPanelForm(UserIndex)
    End With
End Sub

''
' Handles the "GMPanel" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestUserList(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 01/09/07
'Last modified by: Lucas Tavolaro Ortiz (Tavo)
'I haven`t found a solution to split, so i make an array of names
'***************************************************
    Dim i As Long
    Dim names() As String
    Dim Count As Long
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub
        
        ReDim names(1 To LastUser) As String
        Count = 1
        
        For i = 1 To LastUser
            If (LenB(UserList(i).Name) <> 0) Then
                If UserList(i).flags.Privilegios And PlayerType.User Then
                    names(Count) = UserList(i).Name
                    Count = Count + 1
                End If
            End If
        Next i
        
        If Count > 1 Then Call WriteUserNameList(UserIndex, names(), Count - 1)
    End With
End Sub

''
' Handles the "Working" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWorking(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i As Long
    Dim users As String
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub
        
        For i = 1 To LastUser
            If UserList(i).flags.UserLogged And UserList(i).Counters.Trabajando > 0 Then
                users = users & ", " & UserList(i).Name
                
                ' Display the user being checked by the centinel
                If modCentinela.Centinela.RevisandoUserIndex = i Then _
                    users = users & " (*)"
            End If
        Next i
        
        If LenB(users) <> 0 Then
            users = Right$(users, Len(users) - 2)
            Call WriteConsoleMsg(UserIndex, "Usuarios trabajando: " & users, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "No hay usuarios trabajando.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "Hiding" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHiding(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i As Long
    Dim users As String
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub
        
        For i = 1 To LastUser
            If (LenB(UserList(i).Name) <> 0) And UserList(i).Counters.Ocultando > 0 Then
                users = users & UserList(i).Name & ", "
            End If
        Next i
        
        If LenB(users) <> 0 Then
            users = Left$(users, Len(users) - 2)
            Call WriteConsoleMsg(UserIndex, "Usuarios ocultandose: " & users, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "No hay usuarios ocultandose.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "Jail" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleJail(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim reason As String
        Dim jailTime As Byte
        Dim Count As Byte
        Dim tUser As Integer
        
        UserName = Buffer.ReadASCIIString()
        reason = Buffer.ReadASCIIString()
        jailTime = Buffer.ReadByte()
        
        If InStr(1, UserName, "+") Then
            UserName = Replace(UserName, "+", " ")
        End If
        
        '/carcel nick@motivo@<tiempo>
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (Not .flags.Privilegios And PlayerType.User) <> 0 Then
            If LenB(UserName) = 0 Or LenB(reason) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Utilice /carcel nick@motivo@tiempo", FontTypeNames.FONTTYPE_INFO)
            Else
                tUser = NameIndex(UserName)
                
                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "El usuario no est� online.", FontTypeNames.FONTTYPE_INFO)
                Else
                    If Not UserList(tUser).flags.Privilegios And PlayerType.User Then
                        Call WriteConsoleMsg(UserIndex, "No puedes encarcelar a administradores.", FontTypeNames.FONTTYPE_INFO)
                    ElseIf jailTime > 60 Then
                        Call WriteConsoleMsg(UserIndex, "No pued�s encarcelar por m�s de 60 minutos.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        If (InStrB(UserName, "\") <> 0) Then
                            UserName = Replace(UserName, "\", "")
                        End If
                        If (InStrB(UserName, "/") <> 0) Then
                            UserName = Replace(UserName, "/", "")
                        End If
                        
                        If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                            Count = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                            Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", Count + 1)
                            Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & Count + 1, LCase$(.Name) & ": CARCEL " & jailTime & "m, MOTIVO: " & LCase$(reason) & " " & Date & " " & time)
                        End If
                        
                        Call Encarcelar(tUser, jailTime, .Name)
                        Call LogGM(.Name, " encarcel� a " & UserName)
                    End If
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "KillNPC" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillNPC(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 04/22/08 (NicoNZ)
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        Dim tNPC As Integer
        Dim auxNPC As npc
        
        'Los consejeros no pueden RMATAr a nada en el mapa pretoriano
        If .flags.Privilegios And PlayerType.Consejero Then
            If .Pos.Map = MAPA_PRETORIANO Then
                Call WriteConsoleMsg(UserIndex, "Los consejeros no pueden usar este comando en el mapa pretoriano.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        
        tNPC = .flags.TargetNPC
        
        If tNPC > 0 Then
            Call WriteConsoleMsg(UserIndex, "RMatas (con posible respawn) a: " & Npclist(tNPC).Name, FontTypeNames.FONTTYPE_INFO)
            
            auxNPC = Npclist(tNPC)
            Call QuitarNPC(tNPC)
            Call ReSpawnNpc(auxNPC)
            
            .flags.TargetNPC = 0
        Else
            Call WriteConsoleMsg(UserIndex, "Antes debes hacer click sobre el NPC.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "WarnUser" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarnUser(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/26/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim reason As String
        Dim privs As PlayerType
        Dim Count As Byte
        
        UserName = Buffer.ReadASCIIString()
        reason = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (Not .flags.Privilegios And PlayerType.User) <> 0 Then
            If LenB(UserName) = 0 Or LenB(reason) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Utilice /advertencia nick@motivo", FontTypeNames.FONTTYPE_INFO)
            Else
                privs = UserDarPrivilegioLevel(UserName)
                
                If Not privs And PlayerType.User Then
                    Call WriteConsoleMsg(UserIndex, "No puedes advertir a administradores.", FontTypeNames.FONTTYPE_INFO)
                Else
                    If (InStrB(UserName, "\") <> 0) Then
                            UserName = Replace(UserName, "\", "")
                    End If
                    If (InStrB(UserName, "/") <> 0) Then
                            UserName = Replace(UserName, "/", "")
                    End If
                    
                    If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                        Count = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", Count + 1)
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & Count + 1, LCase$(.Name) & ": ADVERTENCIA por: " & LCase$(reason) & " " & Date & " " & time)
                        
                        Call WriteConsoleMsg(UserIndex, "Has advertido a " & UCase$(UserName) & ".", FontTypeNames.FONTTYPE_INFO)
                        Call LogGM(.Name, " advirtio a " & UserName)
                    End If
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "EditChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEditChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 11/06/2009
'02/03/2009: ZaMa - Cuando editas nivel, chequea si el pj puede permanecer en clan faccionario
'11/06/2009: ZaMa - Todos los comandos se pueden usar aunque el pj este offline
'***************************************************
    If UserList(UserIndex).incomingData.Length < 8 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim opcion As Byte
        Dim Arg1 As String
        Dim Arg2 As String
        Dim valido As Boolean
        Dim LoopC As Byte
        Dim CommandString As String
        Dim UserCharPath As String
        Dim Var As Long
        
        
        UserName = Replace(Buffer.ReadASCIIString(), "+", " ")
        
        If UCase$(UserName) = "YO" Then
            tUser = UserIndex
        Else
            tUser = NameIndex(UserName)
        End If
        
        opcion = Buffer.ReadByte()
        Arg1 = Buffer.ReadASCIIString()
        Arg2 = Buffer.ReadASCIIString()
        
        If .flags.Privilegios And PlayerType.RoleMaster Then
            Select Case .flags.Privilegios And (PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)
                Case PlayerType.Consejero
                    ' Los RMs consejeros s�lo se pueden editar su head, body y level
                    valido = tUser = UserIndex And _
                            (opcion = eEditOptions.eo_Body Or opcion = eEditOptions.eo_Head Or opcion = eEditOptions.eo_Level)
                
                Case PlayerType.SemiDios
                    ' Los RMs s�lo se pueden editar su level y el head y body de cualquiera
                    valido = (opcion = eEditOptions.eo_Level And tUser = UserIndex) _
                            Or opcion = eEditOptions.eo_Body Or opcion = eEditOptions.eo_Head
                
                Case PlayerType.Dios
                    ' Los DRMs pueden aplicar los siguientes comandos sobre cualquiera
                    ' pero si quiere modificar el level s�lo lo puede hacer sobre s� mismo
                    valido = (opcion = eEditOptions.eo_Level And tUser = UserIndex) Or _
                            opcion = eEditOptions.eo_Body Or _
                            opcion = eEditOptions.eo_Head Or _
                            opcion = eEditOptions.eo_CiticensKilled Or _
                            opcion = eEditOptions.eo_CriminalsKilled Or _
                            opcion = eEditOptions.eo_Class Or _
                            opcion = eEditOptions.eo_Skills Or _
                            opcion = eEditOptions.eo_addGold
            End Select
            
        ElseIf .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then   'Si no es RM debe ser dios para poder usar este comando
            valido = True
        End If

        If valido Then
            UserCharPath = CharPath & UserName & ".chr"
            If tUser <= 0 And Not FileExist(UserCharPath) Then
                Call WriteConsoleMsg(UserIndex, "Est�s intentando editar un usuario inexistente.", FontTypeNames.FONTTYPE_INFO)
                Call LogGM(.Name, "Intent� editar un usuario inexistente.")
            Else
                'For making the Log
                CommandString = "/MOD "
                
                Select Case opcion
                    Case eEditOptions.eo_Gold
                        If val(Arg1) <= MAX_ORO_EDIT Then
                            If tUser <= 0 Then ' Esta offline?
                                Call WriteVar(UserCharPath, "STATS", "GLD", val(Arg1))
                                Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                            Else ' Online
                                UserList(tUser).Stats.GLD = val(Arg1)
                                Call WriteUpdateGold(tUser)
                            End If
                        Else
                            Call WriteConsoleMsg(UserIndex, "No est� permitido utilizar valores mayores a " & MAX_ORO_EDIT & ". Su comando ha quedado en los logs del juego.", FontTypeNames.FONTTYPE_INFO)
                        End If
                    
                        ' Log it
                        CommandString = CommandString & "ORO "
                
                    Case eEditOptions.eo_Experience
                        If val(Arg1) > 20000000 Then
                                Arg1 = 20000000
                        End If
                        
                        If tUser <= 0 Then ' Offline
                            Var = GetVar(UserCharPath, "STATS", "EXP")
                            Call WriteVar(UserCharPath, "STATS", "EXP", Var + val(Arg1))
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else ' Online
                            UserList(tUser).Stats.Exp = UserList(tUser).Stats.Exp + val(Arg1)
                            Call CheckUserLevel(tUser)
                            Call WriteUpdateExp(tUser)
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "EXP "
                    
                    Case eEditOptions.eo_Body
                        If tUser <= 0 Then
                            Call WriteVar(UserCharPath, "INIT", "Body", Arg1)
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call ChangeUserChar(tUser, val(Arg1), UserList(tUser).Char.Head, UserList(tUser).Char.heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim)
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "BODY "
                    
                    Case eEditOptions.eo_Head
                        If tUser <= 0 Then
                            Call WriteVar(UserCharPath, "INIT", "Head", Arg1)
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call ChangeUserChar(tUser, UserList(tUser).Char.body, val(Arg1), UserList(tUser).Char.heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim)
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "HEAD "
                    
                    Case eEditOptions.eo_CriminalsKilled
                        Var = IIf(val(Arg1) > MAXUSERMATADOS, MAXUSERMATADOS, val(Arg1))
                        
                        If tUser <= 0 Then ' Offline
                            Call WriteVar(UserCharPath, "FACCIONES", "CrimMatados", Var)
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else ' Online
                            UserList(tUser).Faccion.Matados(eFaccion.Caos) = Var
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "CRI "
                    
                    Case eEditOptions.eo_CiticensKilled
                        Var = IIf(val(Arg1) > MAXUSERMATADOS, MAXUSERMATADOS, val(Arg1))
                        
                        If tUser <= 0 Then ' Offline
                            Call WriteVar(UserCharPath, "FACCIONES", "CiudMatados", Var)
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else ' Online
                            UserList(tUser).Faccion.Matados(eFaccion.Real) = Var
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "CIU "
                    
                    Case eEditOptions.eo_Level
                        If val(Arg1) > STAT_MAXELV Then
                            Arg1 = CStr(STAT_MAXELV)
                            Call WriteConsoleMsg(UserIndex, "No puedes tener un nivel superior a " & STAT_MAXELV & ".", FONTTYPE_INFO)
                        End If
                        
                        ' Chequeamos si puede permanecer en el clan
                       ' If val(Arg1) >= 25 Then
                            
                           ' Dim GI As Integer
                           ' If tUser <= 0 Then
                           '     GI = GetVar(UserCharPath, "GUILD", "GUILDINDEX")
                          '  Else
                           '     GI = UserList(tUser).GuildIndex
                         '   End If
                            
                        '    If GI > 0 Then
                        '        If modGuilds.GuildAlignment(GI) = "Del Mal" Or modGuilds.GuildAlignment(GI) = "Real" Then
                       '             'We get here, so guild has factionary alignment, we have to expulse the user
                       '             Call modGuilds.m_EcharMiembroDeClan(-1, UserName)
                       '
                       '             Call SendData(SendTarget.ToGuildMembers, GI, PrepareMessageConsoleMsg(UserName & " deja el clan.", FontTypeNames.FONTTYPE_GUILD))
                       '             ' Si esta online le avisamos
                       '             If tUser > 0 Then _
                       '                 Call WriteConsoleMsg(tUser, "�Ya tienes la madurez suficiente como para decidir bajo que estandarte pelear�s! Por esta raz�n, hasta tanto no te enlistes en la facci�n bajo la cual tu clan est� alineado, estar�s exclu�do del mismo.", FontTypeNames.FONTTYPE_GUILD)
                       '         End If
                       '     End If
                       'End If
                        
                        If tUser <= 0 Then ' Offline
                            Call WriteVar(UserCharPath, "STATS", "ELV", val(Arg1))
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else ' Online
                            UserList(tUser).Stats.ELV = val(Arg1)
                            If PuedeSubirClase(tUser) Then Call WriteSubeClase(tUser, True) Else Call WriteSubeClase(tUser, False)
                            Call WriteUpdateUserStats(tUser)
                        End If
                    
                        ' Log it
                        CommandString = CommandString & "LEVEL "
                    
                    Case eEditOptions.eo_Class
                        For LoopC = 1 To NUMCLASES
                            If UCase$(ListaClases(LoopC)) = UCase$(Arg1) Then Exit For
                        Next LoopC
                            
                        If LoopC > NUMCLASES Then
                            Call WriteConsoleMsg(UserIndex, "Clase desconocida. Intente nuevamente.", FontTypeNames.FONTTYPE_INFO)
                        Else
                            If tUser <= 0 Then ' Offline
                                Call WriteVar(UserCharPath, "INIT", "Clase", LoopC)
                                Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                            Else ' Online
                                UserList(tUser).Clase = LoopC
                                
                                If PuedeSubirClase(tUser) Then Call WriteSubeClase(tUser, True) Else Call WriteSubeClase(tUser, False)
                            End If
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "CLASE "
                        
                    Case eEditOptions.eo_Skills
                        For LoopC = 1 To NUMSKILLS
                            If UCase$(Replace$(SkillsNames(LoopC), " ", "+")) = UCase$(Arg1) Then Exit For
                        Next LoopC
                        
                        If LoopC > NUMSKILLS Then
                            Call WriteConsoleMsg(UserIndex, "Skill Inexistente!", FontTypeNames.FONTTYPE_INFO)
                        Else
                            If tUser <= 0 Then ' Offline
                                Call WriteVar(UserCharPath, "Skills", "SK" & LoopC, Arg2)
                                
                                Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                            Else ' Online
                                UserList(tUser).Stats.UserSkills(LoopC) = val(Arg2)
                            End If
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "SKILLS "
                    
                    Case eEditOptions.eo_SkillPointsLeft
                        If tUser <= 0 Then ' Offline
                            Call WriteVar(UserCharPath, "STATS", "SkillPtsLibres", Arg1)
                            Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        Else ' Online
                            UserList(tUser).Stats.SkillPts = val(Arg1)
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "SKILLSLIBRES "
                    
                    Case eEditOptions.eo_Sex
                        Dim Sex As Byte
                        Sex = IIf(UCase$(Arg1) = "MUJER", eGenero.Mujer, 0) ' Mujer?
                        Sex = IIf(UCase$(Arg1) = "HOMBRE", eGenero.Hombre, Sex) ' Hombre?
                        
                        If Sex <> 0 Then ' Es Hombre o mujer?
                            If tUser <= 0 Then ' OffLine
                                Call WriteVar(UserCharPath, "INIT", "Genero", Sex)
                                Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                            Else ' Online
                                UserList(tUser).Genero = Sex
                            End If
                        Else
                            Call WriteConsoleMsg(UserIndex, "Genero desconocido. Intente nuevamente.", FontTypeNames.FONTTYPE_INFO)
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "SEX "
                    
                    Case eEditOptions.eo_Raza
                        Dim raza As Byte
                        
                        Arg1 = UCase$(Arg1)
                        Select Case Arg1
                            Case "HUMANO"
                                raza = eRaza.Humano
                            Case "ELFO"
                                raza = eRaza.Elfo
                            Case "DROW"
                                raza = eRaza.Drow
                            Case "ENANO"
                                raza = eRaza.Enano
                            Case "GNOMO"
                                raza = eRaza.Gnomo
                            Case Else
                                raza = 0
                        End Select
                        
                            
                        If raza = 0 Then
                            Call WriteConsoleMsg(UserIndex, "Raza desconocida. Intente nuevamente.", FontTypeNames.FONTTYPE_INFO)
                        Else
                            If tUser <= 0 Then
                                Call WriteVar(UserCharPath, "INIT", "Raza", raza)
                                Call WriteConsoleMsg(UserIndex, "Charfile Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                            Else
                                UserList(tUser).raza = raza
                            End If
                        End If
                            
                        ' Log it
                        CommandString = CommandString & "RAZA "
                        
                    Case eEditOptions.eo_addGold
                    
                        Dim bankGold As Long
                        
                        If Abs(Arg1) > MAX_ORO_EDIT Then
                            Call WriteConsoleMsg(UserIndex, "No est� permitido utilizar valores mayores a " & MAX_ORO_EDIT & ".", FontTypeNames.FONTTYPE_INFO)
                        Else
                            If tUser <= 0 Then
                                bankGold = GetVar(CharPath & UserName & ".chr", "STATS", "BANCO")
                                Call WriteVar(UserCharPath, "STATS", "BANCO", IIf(bankGold + val(Arg1) <= 0, 0, bankGold + val(Arg1)))
                                Call WriteConsoleMsg(UserIndex, "Se le ha agregado " & Arg1 & " monedas de oro a " & UserName & ".", FONTTYPE_TALK)
                            Else
                                UserList(tUser).Stats.Banco = IIf(UserList(tUser).Stats.Banco + val(Arg1) <= 0, 0, UserList(tUser).Stats.Banco + val(Arg1))
                                Call WriteConsoleMsg(tUser, STANDARD_BOUNTY_HUNTER_MESSAGE, FONTTYPE_TALK)
                            End If
                        End If
                        
                        ' Log it
                        CommandString = CommandString & "AGREGAR "
                        
                    Case Else
                        Call WriteConsoleMsg(UserIndex, "Comando no permitido.", FontTypeNames.FONTTYPE_INFO)
                        CommandString = CommandString & "UNKOWN "
                        
                End Select
                
                CommandString = CommandString & Arg1 & " " & Arg2
                Call LogGM(.Name, CommandString & " " & UserName)
                
            End If
        End If
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub


''
' Handles the "RequestCharInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharInfo(ByVal UserIndex As Integer)
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/08/07
'Last Modification by: (liquid).. alto bug zapallo..
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
                
        Dim targetName As String
        Dim TargetIndex As Integer
        
        targetName = Replace$(Buffer.ReadASCIIString(), "+", " ")
        TargetIndex = NameIndex(targetName)
        
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then
            'is the player offline?
            If TargetIndex <= 0 Then
                'don't allow to retrieve administrator's info
                If Not (EsDios(targetName) Or EsAdmin(targetName)) Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline, buscando en charfile.", FontTypeNames.FONTTYPE_INFO)
                    Call SendUserStatsTxtOFF(UserIndex, targetName)
                End If
            Else
                'don't allow to retrieve administrator's info
                If UserList(TargetIndex).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then
                    Call SendUserStatsTxt(UserIndex, TargetIndex)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "RequestCharStats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharStats(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        UserName = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            Call LogGM(.Name, "/STAT " & UserName)
            
            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline. Leyendo charfile... ", FontTypeNames.FONTTYPE_INFO)
                
                Call SendUserMiniStatsTxtFromChar(UserIndex, UserName)
            Else
                Call SendUserMiniStatsTxt(UserIndex, tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "RequestCharGold" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharGold(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.Name, "/BAL " & UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline. Leyendo charfile... ", FontTypeNames.FONTTYPE_TALK)
                
                Call SendUserOROTxtFromChar(UserIndex, UserName)
            Else
                Call WriteConsoleMsg(UserIndex, "El usuario " & UserName & " tiene " & UserList(tUser).Stats.Banco & " en el banco.", FontTypeNames.FONTTYPE_TALK)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "RequestCharInventory" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharInventory(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.Name, "/INV " & UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline. Leyendo del charfile...", FontTypeNames.FONTTYPE_TALK)
                
                Call SendUserInvTxtFromChar(UserIndex, UserName)
            Else
                Call SendUserInvTxt(UserIndex, tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "RequestCharBank" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharBank(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.Name, "/BOV " & UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline. Leyendo charfile... ", FontTypeNames.FONTTYPE_TALK)
                
                Call SendUserBovedaTxtFromChar(UserIndex, UserName)
            Else
                Call SendUserBovedaTxt(UserIndex, tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "RequestCharSkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharSkills(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim LoopC As Long
        Dim message As String
        
        UserName = Buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.Name, "/STATS " & UserName)
            
            If tUser <= 0 Then
                If (InStrB(UserName, "\") <> 0) Then
                        UserName = Replace(UserName, "\", "")
                End If
                If (InStrB(UserName, "/") <> 0) Then
                        UserName = Replace(UserName, "/", "")
                End If
                
                For LoopC = 1 To NUMSKILLS
                    message = message & "CHAR>" & SkillsNames(LoopC) & " = " & GetVar(CharPath & UserName & ".chr", "SKILLS", "SK" & LoopC) & vbCrLf
                Next LoopC
                
                Call WriteConsoleMsg(UserIndex, message & "CHAR> Libres:" & GetVar(CharPath & UserName & ".chr", "STATS", "SKILLPTSLIBRES"), FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendUserSkillsTxt(UserIndex, tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "ReviveChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleReviveChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 11/03/2010
'11/03/2010: ZaMa - Al revivir con el comando, si esta navegando le da cuerpo e barca.
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Buffer.ReadASCIIString()
        
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If UCase$(UserName) <> "YO" Then
                tUser = NameIndex(UserName)
            Else
                tUser = UserIndex
            End If
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                With UserList(tUser)
                    'If dead, show him alive (naked).
                    If .flags.Muerto = 1 Then
                        .flags.Muerto = 0
                        
                        If .flags.Navegando = 1 Then
                            Call ToogleBoatBody(UserIndex)
                        Else
                            Call DarCuerpoDesnudo(tUser)
                        End If
                        
                        If .flags.Traveling = 1 Then
                            .flags.Traveling = 0
                            .Counters.goHome = 0
                            Call WriteMultiMessage(tUser, eMessages.CancelHome)
                        End If
                        
                        Call ChangeUserChar(tUser, .Char.body, .OrigChar.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        
                        Call WriteConsoleMsg(tUser, UserList(UserIndex).Name & " te ha resucitado.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(tUser, UserList(UserIndex).Name & " te ha curado.", FontTypeNames.FONTTYPE_INFO)
                    End If
                    
                    .Stats.MinHp = .Stats.MaxHp
                    
                    If .flags.Traveling = 1 Then
                        .Counters.goHome = 0
                        .flags.Traveling = 0
                        Call WriteMultiMessage(tUser, eMessages.CancelHome)
                    End If
                    
                End With
                
                Call WriteUpdateHP(tUser)
                
                Call FlushBuffer(tUser)
                
                Call LogGM(.Name, "Resucito a " & UserName)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "OnlineGM" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineGM(ByVal UserIndex As Integer)
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 12/28/06
'
'***************************************************
    Dim i As Long
    Dim list As String
    Dim priv As PlayerType
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub

        priv = PlayerType.Consejero Or PlayerType.SemiDios
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then priv = priv Or PlayerType.Dios Or PlayerType.Admin
        
        For i = 1 To LastUser
            If UserList(i).flags.UserLogged Then
                If UserList(i).flags.Privilegios And priv Then _
                    list = list & UserList(i).Name & ", "
            End If
        Next i
        
        If LenB(list) <> 0 Then
            list = Left$(list, Len(list) - 2)
            Call WriteConsoleMsg(UserIndex, list & ".", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "No hay GMs Online.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "OnlineMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineMap(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 23/03/2009
'23/03/2009: ZaMa - Ahora no requiere estar en el mapa, sino que por defecto se toma en el que esta, pero se puede especificar otro
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Map As Integer
        Map = .incomingData.ReadInteger
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        Dim LoopC As Long
        Dim list As String
        Dim priv As PlayerType
        
        priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then priv = priv + (PlayerType.Dios Or PlayerType.Admin)
        
        For LoopC = 1 To LastUser
            If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).Pos.Map = Map Then
                If UserList(LoopC).flags.Privilegios And priv Then _
                    list = list & UserList(LoopC).Name & ", "
            End If
        Next LoopC
        
        If Len(list) > 2 Then list = Left$(list, Len(list) - 2)
        
        Call WriteConsoleMsg(UserIndex, "Usuarios en el mapa: " & list, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "Kick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKick(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim rank As Integer
        
        rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        
        UserName = Buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "El usuario no est� online.", FontTypeNames.FONTTYPE_INFO)
            Else
                If (UserList(tUser).flags.Privilegios And rank) > (.flags.Privilegios And rank) Then
                    Call WriteConsoleMsg(UserIndex, "No puedes echar a alguien con jerarqu�a mayor a la tuya.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " ech� a " & UserName & ".", FontTypeNames.FONTTYPE_INFO))
                    Call CloseSocket(tUser)
                    Call LogGM(.Name, "Ech� a " & UserName)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "Execute" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleExecute(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                If Not UserList(tUser).flags.Privilegios And PlayerType.User Then
                    Call WriteConsoleMsg(UserIndex, "��Est�s loco?? ��C�mo vas a pi�atear un gm?? :@", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call UserDie(tUser)
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " ha ejecutado a " & UserName & ".", FontTypeNames.FONTTYPE_EJECUCION))
                    Call LogGM(.Name, " ejecuto a " & UserName)
                End If
            Else
                Call WriteConsoleMsg(UserIndex, "No est� online.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "BanChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBanChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim reason As String
        
        UserName = Buffer.ReadASCIIString()
        reason = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            Call BanCharacter(UserIndex, UserName, reason)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "UnbanChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUnbanChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim cantPenas As Byte
        
        UserName = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If
            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If
            
            If Not FileExist(CharPath & UserName & ".chr", vbNormal) Then
                Call WriteConsoleMsg(UserIndex, "Charfile inexistente (no use +).", FontTypeNames.FONTTYPE_INFO)
            Else
                If (val(GetVar(CharPath & UserName & ".chr", "FLAGS", "Ban")) = 1) Then
                    Call UnBan(UserName)
                
                    'penas
                    cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                    Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
                    Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, LCase$(.Name) & ": UNBAN. " & Date & " " & time)
                
                    Call LogGM(.Name, "/UNBAN a " & UserName)
                    Call WriteConsoleMsg(UserIndex, UserName & " unbanned.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, UserName & " no est� baneado. Imposible unbanear.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "NPCFollow" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleNPCFollow(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        If .flags.TargetNPC > 0 Then
            Call DoFollow(.flags.TargetNPC, .Name)
            Npclist(.flags.TargetNPC).flags.Inmovilizado = 0
            Npclist(.flags.TargetNPC).flags.Paralizado = 0
            Npclist(.flags.TargetNPC).Contadores.Paralisis = 0
        End If
    End With
End Sub

''
' Handles the "SummonChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSummonChar(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 26/03/2009
'26/03/2009: ZaMa - Chequeo que no se teletransporte donde haya un char o npc
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim X As Integer
        Dim Y As Integer
        
        UserName = Buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "El jugador no est� online.", FontTypeNames.FONTTYPE_INFO)
            Else
                If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Or _
                  (UserList(tUser).flags.Privilegios And (PlayerType.Consejero Or PlayerType.User)) <> 0 Then
                    Call WriteConsoleMsg(tUser, .Name & " te ha trasportado.", FontTypeNames.FONTTYPE_INFO)
                    X = .Pos.X
                    Y = .Pos.Y + 1
                    Call FindLegalPos(tUser, .Pos.Map, X, Y)
                    Call WarpUserChar(tUser, .Pos.Map, X, Y, True, True)
                    Call LogGM(.Name, "/SUM " & UserName & " Map:" & .Pos.Map & " X:" & .Pos.X & " Y:" & .Pos.Y)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes invocar a dioses y admins.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "SpawnListRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpawnListRequest(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        Call EnviarSpawnList(UserIndex)
    End With
End Sub

''
' Handles the "SpawnCreature" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpawnCreature(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim npc As Integer
        npc = .incomingData.ReadInteger()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If npc > 0 And npc <= UBound(Declaraciones.SpawnList()) Then _
              Call SpawnNpc(Declaraciones.SpawnList(npc).NpcIndex, .Pos, True, False)
            
            Call LogGM(.Name, "Sumoneo " & Declaraciones.SpawnList(npc).NpcName)
        End If
    End With
End Sub

''
' Handles the "ResetNPCInventory" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResetNPCInventory(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        If .flags.TargetNPC = 0 Then Exit Sub
        
        Call ResetNpcInv(.flags.TargetNPC)
        Call LogGM(.Name, "/RESETINV " & Npclist(.flags.TargetNPC).Name)
    End With
End Sub

''
' Handles the "CleanWorld" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCleanWorld(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LimpiarMundo
    End With
End Sub

''
' Handles the "ServerMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleServerMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim message As String
        message = Buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If LenB(message) <> 0 Then
                Call LogGM(.Name, "Mensaje Broadcast:" & message)
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).Name & ": " & message, FontTypeNames.FONTTYPE_VENENO))
                ''''''''''''''''SOLO PARA EL TESTEO'''''''
                ''''''''''SE USA PARA COMUNICARSE CON EL SERVER'''''''''''
                frmMain.txtChat.Text = frmMain.txtChat.Text & vbNewLine & UserList(UserIndex).Name & ": " & message
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "NickToIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleNickToIP(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 24/07/07
'Pablo (ToxicWaste): Agrego para uqe el /nick2ip tambien diga los nicks en esa ip por pedido de la DGM.
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim priv As PlayerType
        
        UserName = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tUser = NameIndex(UserName)
            Call LogGM(.Name, "NICK2IP Solicito la IP de " & UserName)

            If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin
            Else
                priv = PlayerType.User
            End If
            
            If tUser > 0 Then
                If UserList(tUser).flags.Privilegios And priv Then
                    Call WriteConsoleMsg(UserIndex, "El ip de " & UserName & " es " & UserList(tUser).ip, FontTypeNames.FONTTYPE_INFO)
                    Dim ip As String
                    Dim lista As String
                    Dim LoopC As Long
                    ip = UserList(tUser).ip
                    For LoopC = 1 To LastUser
                        If UserList(LoopC).ip = ip Then
                            If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).flags.UserLogged Then
                                If UserList(LoopC).flags.Privilegios And priv Then
                                    lista = lista & UserList(LoopC).Name & ", "
                                End If
                            End If
                        End If
                    Next LoopC
                    If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
                    Call WriteConsoleMsg(UserIndex, "Los personajes con ip " & ip & " son: " & lista, FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                Call WriteConsoleMsg(UserIndex, "No hay ning�n personaje con ese nick.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "IPToNick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleIPToNick(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim ip As String
        Dim LoopC As Long
        Dim lista As String
        Dim priv As PlayerType
        
        ip = .incomingData.ReadByte() & "."
        ip = ip & .incomingData.ReadByte() & "."
        ip = ip & .incomingData.ReadByte() & "."
        ip = ip & .incomingData.ReadByte()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, "IP2NICK Solicito los Nicks de IP " & ip)
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
            priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin
        Else
            priv = PlayerType.User
        End If

        For LoopC = 1 To LastUser
            If UserList(LoopC).ip = ip Then
                If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).flags.UserLogged Then
                    If UserList(LoopC).flags.Privilegios And priv Then
                        lista = lista & UserList(LoopC).Name & ", "
                    End If
                End If
            End If
        Next LoopC
        
        If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
        Call WriteConsoleMsg(UserIndex, "Los personajes con ip " & ip & " son: " & lista, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "TeleportCreate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTeleportCreate(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 22/03/2010
'15/11/2009: ZaMa - Ahora se crea un teleport con un radio especificado.
'22/03/2010: ZaMa - Harcodeo los teleps y radios en el dat, para evitar mapas bugueados.
'***************************************************
    If UserList(UserIndex).incomingData.Length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim mapa As Integer
        Dim X As Byte
        Dim Y As Byte
        Dim Radio As Byte
        
        mapa = .incomingData.ReadInteger()
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
        Radio = .incomingData.ReadByte()
        
        Radio = MinimoInt(Radio, 6)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call LogGM(.Name, "/CT " & mapa & "," & X & "," & Y & "," & Radio)
        
        If Not MapaValido(mapa) Or Not InMapBounds(mapa, X, Y) Then _
            Exit Sub
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y - 1).ObjInfo.OBJIndex > 0 Then _
            Exit Sub
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y - 1).TileExit.Map > 0 Then _
            Exit Sub
        
        If MapData(mapa, X, Y).ObjInfo.OBJIndex > 0 Then
            Call WriteConsoleMsg(UserIndex, "Hay un objeto en el piso en ese lugar.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If MapData(mapa, X, Y).TileExit.Map > 0 Then
            Call WriteConsoleMsg(UserIndex, "No puedes crear un teleport que apunte a la entrada de otro.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Dim ET As Obj
        ET.Amount = 1
        ' Es el numero en el dat. El indice es el comienzo + el radio, todo harcodeado :(.
        ET.OBJIndex = TELEP_OBJ_INDEX + Radio
        
        With MapData(.Pos.Map, .Pos.X, .Pos.Y - 1)
            .TileExit.Map = mapa
            .TileExit.X = X
            .TileExit.Y = Y
        End With
        
        Call MakeObj(ET, .Pos.Map, .Pos.X, .Pos.Y - 1)
    End With
End Sub

''
' Handles the "TeleportDestroy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTeleportDestroy(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(UserIndex)
        Dim mapa As Integer
        Dim X As Byte
        Dim Y As Byte
        
        'Remove packet ID
        Call .incomingData.ReadByte
        
        '/dt
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        mapa = .flags.TargetMap
        X = .flags.TargetX
        Y = .flags.TargetY
        
        If Not InMapBounds(mapa, X, Y) Then Exit Sub
        
        With MapData(mapa, X, Y)
            If .ObjInfo.OBJIndex = 0 Then Exit Sub
            
            If ObjData(.ObjInfo.OBJIndex).OBJType = eOBJType.otTeleport And .TileExit.Map > 0 Then
                Call LogGM(UserList(UserIndex).Name, "/DT: " & mapa & "," & X & "," & Y)
                
                Call EraseObj(.ObjInfo.Amount, mapa, X, Y)
                
                If MapData(.TileExit.Map, .TileExit.X, .TileExit.Y).ObjInfo.OBJIndex = 651 Then
                    Call EraseObj(1, .TileExit.Map, .TileExit.X, .TileExit.Y)
                End If
                
                .TileExit.Map = 0
                .TileExit.X = 0
                .TileExit.Y = 0
            End If
        End With
    End With
End Sub

''
' Handles the "RainToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRainToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        Call LogGM(.Name, "/LLUVIA")
        Lloviendo = Not Lloviendo
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
    End With
End Sub

''
' Handles the "SetCharDescription" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSetCharDescription(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim tUser As Integer
        Dim desc As String
        
        desc = Buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Or (.flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
            tUser = .flags.TargetUser
            If tUser > 0 Then
                UserList(tUser).DescRM = desc
            Else
                Call WriteConsoleMsg(UserIndex, "Haz click sobre un personaje antes.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "ForceMIDIToMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HanldeForceMIDIToMap(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim midiID As Byte
        Dim mapa As Integer
        
        midiID = .incomingData.ReadByte
        mapa = .incomingData.ReadInteger
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            'Si el mapa no fue enviado tomo el actual
            If Not InMapBounds(mapa, 50, 50) Then
                mapa = .Pos.Map
            End If
        
            If midiID = 0 Then
                'Ponemos el default del mapa
                Call SendData(SendTarget.toMap, mapa, PrepareMessagePlayMidi(MapInfo(.Pos.Map).Music))
            Else
                'Ponemos el pedido por el GM
                Call SendData(SendTarget.toMap, mapa, PrepareMessagePlayMidi(midiID))
            End If
        End If
    End With
End Sub

''
' Handles the "ForceWAVEToMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceWAVEToMap(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim waveID As Byte
        Dim mapa As Integer
        Dim X As Byte
        Dim Y As Byte
        
        waveID = .incomingData.ReadByte()
        mapa = .incomingData.ReadInteger()
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
        'Si el mapa no fue enviado tomo el actual
            If Not InMapBounds(mapa, X, Y) Then
                mapa = .Pos.Map
                X = .Pos.X
                Y = .Pos.Y
            End If
            
            'Ponemos el pedido por el GM
            Call SendData(SendTarget.toMap, mapa, PrepareMessagePlayWave(waveID, X, Y))
        End If
    End With
End Sub

''
' Handles the "RoyalArmyMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRoyalArmyMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim message As String
        message = Buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToReal, 0, PrepareMessageConsoleMsg("EJ�RCITO REAL> " & message, FontTypeNames.FONTTYPE_TALK))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "ChaosLegionMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChaosLegionMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim message As String
        message = Buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToCaos, 0, PrepareMessageConsoleMsg("FUERZAS DEL CAOS> " & message, FontTypeNames.FONTTYPE_TALK))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "CitizenMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCitizenMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim message As String
        message = Buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToCiudadanos, 0, PrepareMessageConsoleMsg("CIUDADANOS> " & message, FontTypeNames.FONTTYPE_TALK))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "CriminalMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCriminalMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim message As String
        message = Buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToCriminales, 0, PrepareMessageConsoleMsg("CRIMINALES> " & message, FontTypeNames.FONTTYPE_TALK))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "TalkAsNPC" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTalkAsNPC(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/29/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim message As String
        message = Buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            'Asegurarse haya un NPC seleccionado
            If .flags.TargetNPC > 0 Then
                Call SendData(SendTarget.ToNPCArea, .flags.TargetNPC, PrepareMessageChatOverHead(message, Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
            Else
                Call WriteConsoleMsg(UserIndex, "Debes seleccionar el NPC por el que quieres hablar antes de usar este comando.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "DestroyAllItemsInArea" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDestroyAllItemsInArea(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Dim X As Long
        Dim Y As Long
        Dim bIsExit As Boolean
        
        For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
            For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1
                If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                    If MapData(.Pos.Map, X, Y).ObjInfo.OBJIndex > 0 Then
                        bIsExit = MapData(.Pos.Map, X, Y).TileExit.Map > 0
                        If ItemNoEsDeMapa(MapData(.Pos.Map, X, Y).ObjInfo.OBJIndex, bIsExit) Then
                            Call EraseObj(MAX_INVENTORY_OBJS, .Pos.Map, X, Y)
                        End If
                    End If
                End If
            Next X
        Next Y
        
        Call LogGM(UserList(UserIndex).Name, "/MASSDEST")
    End With
End Sub

''
' Handles the "AcceptRoyalCouncilMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAcceptRoyalCouncilMember(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue aceptado en el honorable Consejo Real de Banderbill.", FontTypeNames.FONTTYPE_CONSEJO))
                With UserList(tUser)
                    If .flags.Privilegios And PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios - PlayerType.ChaosCouncil
                    If Not .flags.Privilegios And PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios + PlayerType.RoyalCouncil
                    
                    Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)
                End With
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "ChaosCouncilMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAcceptChaosCouncilMember(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue aceptado en el Concilio de las Sombras.", FontTypeNames.FONTTYPE_CONSEJO))
                
                With UserList(tUser)
                    If .flags.Privilegios And PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios - PlayerType.RoyalCouncil
                    If Not .flags.Privilegios And PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios + PlayerType.ChaosCouncil

                    Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)
                End With
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "ItemsInTheFloor" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleItemsInTheFloor(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Dim tObj As Integer
        Dim X As Long
        Dim Y As Long
        
        For X = 5 To 95
            For Y = 5 To 95
                tObj = MapData(.Pos.Map, X, Y).ObjInfo.OBJIndex
                If tObj > 0 Then
                    If ObjData(tObj).OBJType <> eOBJType.otArboles Then
                        Call WriteConsoleMsg(UserIndex, "(" & X & "," & Y & ") " & ObjData(tObj).Name, FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            Next Y
        Next X
    End With
End Sub

''
' Handles the "MakeDumb" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMakeDumb(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Buffer.ReadASCIIString()
        
        If ((.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Or ((.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) = (PlayerType.SemiDios Or PlayerType.RoleMaster))) Then
            tUser = NameIndex(UserName)
            'para deteccion de aoice
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteDumb(tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "MakeDumbNoMore" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMakeDumbNoMore(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Buffer.ReadASCIIString()
        
        If ((.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Or ((.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) = (PlayerType.SemiDios Or PlayerType.RoleMaster))) Then
            tUser = NameIndex(UserName)
            'para deteccion de aoice
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteDumbNoMore(tUser)
                Call FlushBuffer(tUser)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "DumpIPTables" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDumpIPTables(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call SecurityIp.DumpTables
    End With
End Sub

''
' Handles the "CouncilKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCouncilKick(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            tUser = NameIndex(UserName)
            If tUser <= 0 Then
                If FileExist(CharPath & UserName & ".chr") Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline, echando de los consejos.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteVar(CharPath & UserName & ".chr", "CONSEJO", "PERTENECE", 0)
                    Call WriteVar(CharPath & UserName & ".chr", "CONSEJO", "PERTENECECAOS", 0)
                Else
                    Call WriteConsoleMsg(UserIndex, "No se encuentra el charfile " & CharPath & UserName & ".chr", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                With UserList(tUser)
                    If .flags.Privilegios And PlayerType.RoyalCouncil Then
                        Call WriteConsoleMsg(tUser, "Has sido echado del consejo de Banderbill.", FontTypeNames.FONTTYPE_TALK)
                        .flags.Privilegios = .flags.Privilegios - PlayerType.RoyalCouncil
                        
                        Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue expulsado del consejo de Banderbill.", FontTypeNames.FONTTYPE_CONSEJO))
                    End If
                    
                    If .flags.Privilegios And PlayerType.ChaosCouncil Then
                        Call WriteConsoleMsg(tUser, "Has sido echado del Concilio de las Sombras.", FontTypeNames.FONTTYPE_TALK)
                        .flags.Privilegios = .flags.Privilegios - PlayerType.ChaosCouncil
                        
                        Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue expulsado del Concilio de las Sombras.", FontTypeNames.FONTTYPE_CONSEJO))
                    End If
                End With
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "SetTrigger" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSetTrigger(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim tTrigger As Byte
        Dim tLog As String
        
        tTrigger = .incomingData.ReadByte()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        If tTrigger >= 0 Then
            MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = tTrigger
            tLog = "Trigger " & tTrigger & " en mapa " & .Pos.Map & " " & .Pos.X & "," & .Pos.Y
            
            Call LogGM(.Name, tLog)
            Call WriteConsoleMsg(UserIndex, tLog, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "AskTrigger" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAskTrigger(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 04/13/07
'
'***************************************************
    Dim tTrigger As Byte
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        tTrigger = MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger
        
        Call LogGM(.Name, "Miro el trigger en " & .Pos.Map & "," & .Pos.X & "," & .Pos.Y & ". Era " & tTrigger)
        
        Call WriteConsoleMsg(UserIndex, _
            "Trigger " & tTrigger & " en mapa " & .Pos.Map & " " & .Pos.X & ", " & .Pos.Y _
            , FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "BannedIPList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBannedIPList(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Dim lista As String
        Dim LoopC As Long
        
        Call LogGM(.Name, "/BANIPLIST")
        
        For LoopC = 1 To BanIps.Count
            lista = lista & BanIps.Item(LoopC) & ", "
        Next LoopC
        
        If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
        
        Call WriteConsoleMsg(UserIndex, lista, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the "BannedIPReload" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBannedIPReload(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call BanIpGuardar
        Call BanIpCargar
    End With
End Sub


''
' Handles the "BanIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBanIP(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 07/02/09
'Agregado un CopyBuffer porque se producia un bucle
'inifito al intentar banear una ip ya baneada. (NicoNZ)
'07/02/09 Pato - Ahora no es posible saber si un gm est� o no online.
'***************************************************
    If UserList(UserIndex).incomingData.Length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim bannedIP As String
        Dim tUser As Integer
        Dim reason As String
        Dim i As Long
        
        ' Is it by ip??
        If Buffer.ReadBoolean() Then
            bannedIP = Buffer.ReadByte() & "."
            bannedIP = bannedIP & Buffer.ReadByte() & "."
            bannedIP = bannedIP & Buffer.ReadByte() & "."
            bannedIP = bannedIP & Buffer.ReadByte()
        Else
            tUser = NameIndex(Buffer.ReadASCIIString())
            
            If tUser > 0 Then bannedIP = UserList(tUser).ip
        End If
        
        reason = Buffer.ReadASCIIString()
        
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            If LenB(bannedIP) > 0 Then
                Call LogGM(.Name, "/BanIP " & bannedIP & " por " & reason)
                
                If BanIpBuscar(bannedIP) > 0 Then
                    Call WriteConsoleMsg(UserIndex, "La IP " & bannedIP & " ya se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call BanIpAgrega(bannedIP)
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " bane� la IP " & bannedIP & " por " & reason, FontTypeNames.FONTTYPE_FIGHT))
                    
                    'Find every player with that ip and ban him!
                    For i = 1 To LastUser
                        If UserList(i).ConnIDValida Then
                            If UserList(i).ip = bannedIP Then
                                Call BanCharacter(UserIndex, UserList(i).Name, "IP POR " & reason)
                            End If
                        End If
                    Next i
                End If
            ElseIf tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "El personaje no est� online.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "UnbanIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUnbanIP(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim bannedIP As String
        
        bannedIP = .incomingData.ReadByte() & "."
        bannedIP = bannedIP & .incomingData.ReadByte() & "."
        bannedIP = bannedIP & .incomingData.ReadByte() & "."
        bannedIP = bannedIP & .incomingData.ReadByte()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        If BanIpQuita(bannedIP) Then
            Call WriteConsoleMsg(UserIndex, "La IP """ & bannedIP & """ se ha quitado de la lista de bans.", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "La IP """ & bannedIP & """ NO se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handles the "CreateItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreateItem(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 07/01/17
'- 07/01/17 Rhynne: Agregue el parametro cantidad de manera opcional y ahora el item se pone en el inventario
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim tObj As Integer
        Dim tCant As Integer
        tObj = .incomingData.ReadInteger()
        tCant = .incomingData.ReadInteger()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
            
        Call LogGM(.Name, "/ITEM: " & tObj & "(" & tCant & ")")
        
       ' If MapData(.Pos.Map, .Pos.X, .Pos.Y - 1).ObjInfo.OBJIndex > 0 Then _
            Exit Sub
        
        'If MapData(.Pos.Map, .Pos.X, .Pos.Y - 1).TileExit.Map > 0 Then _
            Exit Sub
        
        If tObj < 1 Or tObj > NumObjDatas Then Exit Sub
        
        'Is the object not null?
        If LenB(ObjData(tObj).Name) = 0 Then Exit Sub
        
        Dim Objeto As Obj
        Call WriteConsoleMsg(UserIndex, "Haz creado el �tem: " & ObjData(tObj).Name & " (x" & tCant & ")", FontTypeNames.FONTTYPE_INFO)
        
        Objeto.Amount = tCant
        Objeto.OBJIndex = tObj
        Call MeterItemEnInventario(UserIndex, Objeto)
    End With
End Sub

''
' Handles the "DestroyItems" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDestroyItems(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.OBJIndex = 0 Then Exit Sub
        
        Call LogGM(.Name, "/DEST")
        
        If ObjData(MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.OBJIndex).OBJType = eOBJType.otTeleport And _
            MapData(.Pos.Map, .Pos.X, .Pos.Y).TileExit.Map > 0 Then
            
            Call WriteConsoleMsg(UserIndex, "No puede destruir teleports as�. Utilice /DT.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Call EraseObj(10000, .Pos.Map, .Pos.X, .Pos.Y)
    End With
End Sub

''
' Handles the "ChaosLegionKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChaosLegionKick(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If
            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If
            tUser = NameIndex(UserName)
            
            Call LogGM(.Name, "ECHO DEL CAOS A: " & UserName)
    
            If tUser > 0 Then
                Call Expulsar(tUser) 'todo, se puede reenlistar?
                Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas del caos y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(tUser, .Name & " te ha expulsado en forma definitiva de las fuerzas del caos.", FontTypeNames.FONTTYPE_FIGHT)
                Call FlushBuffer(tUser)
            Else
                If FileExist(CharPath & UserName & ".chr") Then
                    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "Bando", eFaccion.Neutral)
                    'Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "Extra", "Expulsado por " & .Name)
                    Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas del caos y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, UserName & ".chr inexistente.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "RoyalArmyKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRoyalArmyKick(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        
        UserName = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If
            If (InStrB(UserName, "/") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If
            tUser = NameIndex(UserName)
            
            Call LogGM(.Name, "ECH� DE LA REAL A: " & UserName)
            
            If tUser > 0 Then
                Call Expulsar(tUser) 'todo se puede reenlistar?
                Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas reales y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(tUser, .Name & " te ha expulsado en forma definitiva de las fuerzas reales.", FontTypeNames.FONTTYPE_FIGHT)
                Call FlushBuffer(tUser)
            Else
                If FileExist(CharPath & UserName & ".chr") Then
                    Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "Bando", eFaccion.Neutral)
                    'Call WriteVar(CharPath & UserName & ".chr", "FACCIONES", "Extra", "Expulsado por " & .Name)
                    Call WriteConsoleMsg(UserIndex, UserName & " expulsado de las fuerzas reales y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, UserName & ".chr inexistente.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "ForceMIDIAll" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceMIDIAll(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim midiID As Byte
        midiID = .incomingData.ReadByte()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " broadcast m�sica: " & midiID, FontTypeNames.FONTTYPE_SERVER))
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayMidi(midiID))
    End With
End Sub

''
' Handles the "ForceWAVEAll" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceWAVEAll(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim waveID As Byte
        waveID = .incomingData.ReadByte()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(waveID, NO_3D_SOUND, NO_3D_SOUND))
    End With
End Sub

''
' Handles the "RemovePunishment" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRemovePunishment(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 1/05/07
'Pablo (ToxicWaste): 1/05/07, You can now edit the punishment.
'***************************************************
    If UserList(UserIndex).incomingData.Length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim punishment As Byte
        Dim NewText As String
        
        UserName = Buffer.ReadASCIIString()
        punishment = Buffer.ReadByte
        NewText = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If LenB(UserName) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Utilice /borrarpena Nick@NumeroDePena@NuevaPena", FontTypeNames.FONTTYPE_INFO)
            Else
                If (InStrB(UserName, "\") <> 0) Then
                        UserName = Replace(UserName, "\", "")
                End If
                If (InStrB(UserName, "/") <> 0) Then
                        UserName = Replace(UserName, "/", "")
                End If
                
                If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                    Call LogGM(.Name, " borro la pena: " & punishment & "-" & _
                      GetVar(CharPath & UserName & ".chr", "PENAS", "P" & punishment) _
                      & " de " & UserName & " y la cambi� por: " & NewText)
                    
                    Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & punishment, LCase$(.Name) & ": <" & NewText & "> " & Date & " " & time)
                    
                    Call WriteConsoleMsg(UserIndex, "Pena modificada.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "TileBlockedToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTileBlockedToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

        Call LogGM(.Name, "/BLOQ")
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 0 Then
            MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 1
        Else
            MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 0
        End If
        
        Call Bloquear(True, .Pos.Map, .Pos.X, .Pos.Y, MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked)
    End With
End Sub

''
' Handles the "KillNPCNoRespawn" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillNPCNoRespawn(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        If .flags.TargetNPC = 0 Then Exit Sub
        
        Call QuitarNPC(.flags.TargetNPC)
        Call LogGM(.Name, "/MATA " & Npclist(.flags.TargetNPC).Name)
    End With
End Sub

''
' Handles the "KillAllNearbyNPCs" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillAllNearbyNPCs(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Dim X As Long
        Dim Y As Long
        
        For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
            For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1
                If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                    If MapData(.Pos.Map, X, Y).NpcIndex > 0 Then Call QuitarNPC(MapData(.Pos.Map, X, Y).NpcIndex)
                End If
            Next X
        Next Y
        Call LogGM(.Name, "/MASSKILL")
    End With
End Sub

''
' Handles the "LastIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLastIP(ByVal UserIndex As Integer)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 12/30/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim lista As String
        Dim LoopC As Byte
        Dim priv As Integer
        Dim validCheck As Boolean
        
        priv = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        UserName = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            'Handle special chars
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "\", "")
            End If
            If (InStrB(UserName, "\") <> 0) Then
                UserName = Replace(UserName, "/", "")
            End If
            If (InStrB(UserName, "+") <> 0) Then
                UserName = Replace(UserName, "+", " ")
            End If
            
            'Only Gods and Admins can see the ips of adminsitrative characters. All others can be seen by every adminsitrative char.
            If NameIndex(UserName) > 0 Then
                validCheck = (UserList(NameIndex(UserName)).flags.Privilegios And priv) = 0 Or (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
            Else
                validCheck = (UserDarPrivilegioLevel(UserName) And priv) = 0 Or (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
            End If
            
            If validCheck Then
                Call LogGM(.Name, "/LASTIP " & UserName)
                
                If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                    lista = "Las ultimas IPs con las que " & UserName & " se conect� son:"
                    For LoopC = 1 To 5
                        lista = lista & vbCrLf & LoopC & " - " & GetVar(CharPath & UserName & ".chr", "INIT", "LastIP" & LoopC)
                    Next LoopC
                    Call WriteConsoleMsg(UserIndex, lista, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "Charfile """ & UserName & """ inexistente.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                Call WriteConsoleMsg(UserIndex, UserName & " es de mayor jerarqu�a que vos.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "ChatColor" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleChatColor(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Last modified by: Juan Mart�n Sotuyo Dodero (Maraxus)
'Change the user`s chat color
'***************************************************
    If UserList(UserIndex).incomingData.Length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim color As Long
        
        color = RGB(.incomingData.ReadByte(), .incomingData.ReadByte(), .incomingData.ReadByte())
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
            .flags.ChatColor = color
        End If
    End With
End Sub

''
' Handles the "Ignored" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleIgnored(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Ignore the user
'***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            .flags.AdminPerseguible = Not .flags.AdminPerseguible
        End If
    End With
End Sub

''
' Handles the "CheckSlot" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleCheckSlot(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 09/09/2008 (NicoNZ)
'Check one Users Slot in Particular from Inventory
'***************************************************
    If UserList(UserIndex).incomingData.Length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        'Reads the UserName and Slot Packets
        Dim UserName As String
        Dim Slot As Byte
        Dim tIndex As Integer
        
        UserName = Buffer.ReadASCIIString() 'Que UserName?
        Slot = Buffer.ReadByte() 'Que Slot?
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.Dios) Then
            tIndex = NameIndex(UserName)  'Que user index?
            
            Call LogGM(.Name, .Name & " Checke� el slot " & Slot & " de " & UserName)
               
            If tIndex > 0 Then
                If Slot > 0 And Slot <= UserList(tIndex).CurrentInventorySlots Then
                    If UserList(tIndex).Invent.Object(Slot).OBJIndex > 0 Then
                        Call WriteConsoleMsg(UserIndex, " Objeto " & Slot & ") " & ObjData(UserList(tIndex).Invent.Object(Slot).OBJIndex).Name & " Cantidad:" & UserList(tIndex).Invent.Object(Slot).Amount, FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(UserIndex, "No hay ning�n objeto en slot seleccionado.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(UserIndex, "Slot Inv�lido.", FontTypeNames.FONTTYPE_TALK)
                End If
            Else
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_TALK)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "ReloadObjects" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleReloadObjects(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reload the objects
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha recargado los objetos.")
        
        Call LoadOBJData
    End With
End Sub

''
' Handles the "ReloadSpells" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleReloadSpells(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reload the spells
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha recargado los hechizos.")
        
        Call CargarHechizos
    End With
End Sub

''
' Handle the "ReloadServerIni" message.
'
' @param userIndex The index of the user sending the message

Public Sub HandleReloadServerIni(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reload the Server`s INI
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha recargado los INITs.")
        
        Call LoadSini
    End With
End Sub

''
' Handle the "ReloadNPCs" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleReloadNPCs(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Reload the Server`s NPC
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
         
        Call LogGM(.Name, .Name & " ha recargado los NPCs.")
    
        Call CargaNpcsDat
    
        Call WriteConsoleMsg(UserIndex, "Npcs.dat recargado.", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handle the "KickAllChars" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleKickAllChars(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Kick all the chars that are online
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha echado a todos los personajes.")
        
        Call EcharPjsNoPrivilegiados
    End With
End Sub

''
' Handle the "Night" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleNight(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Last modified by: Juan Mart�n Sotuyo Dodero (Maraxus)
'
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        If UCase$(.Name) <> "MARAXUS" Then Exit Sub
        
        DeNoche = Not DeNoche
        
        Dim i As Long
        
        For i = 1 To NumUsers
            If UserList(i).flags.UserLogged And UserList(i).ConnID > -1 Then
                Call EnviarNoche(i)
            End If
        Next i
    End With
End Sub

''
' Handle the "ShowServerForm" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleShowServerForm(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Show the server form
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha solicitado mostrar el formulario del servidor.")
        Call frmMain.mnuMostrar_Click
    End With
End Sub

''
' Handle the "CleanSOS" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCleanSOS(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Clean the SOS
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha borrado los SOS.")
        
        Call Ayuda.Reset
    End With
End Sub

''
' Handle the "SaveChars" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSaveChars(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/23/06
'Save the characters
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha guardado todos los chars.")
        
        Call GuardarUsuarios
    End With
End Sub

''
' Handle the "ChangeMapInfoBackup" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoBackup(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Last modified by: Juan Mart�n Sotuyo Dodero (Maraxus)
'Change the backup`s info of the map
'***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim doTheBackUp As Boolean
        
        doTheBackUp = .incomingData.ReadBoolean()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0 Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha cambiado la informaci�n sobre el BackUp.")
        
        'Change the boolean to byte in a fast way
        If doTheBackUp Then
            MapInfo(.Pos.Map).BackUp = 1
        Else
            MapInfo(.Pos.Map).BackUp = 0
        End If
        
        'Change the boolean to string in a fast way
        Call WriteVar(App.path & MapPath & "mapa" & .Pos.Map & ".dat", "Mapa" & .Pos.Map, "backup", MapInfo(.Pos.Map).BackUp)
        
        Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Backup: " & MapInfo(.Pos.Map).BackUp, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handle the "ChangeMapInfoPK" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoPK(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Last modified by: Juan Mart�n Sotuyo Dodero (Maraxus)
'Change the pk`s info of the  map
'***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim isMapPk As Boolean
        
        isMapPk = .incomingData.ReadBoolean()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0 Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha cambiado la informaci�n sobre si es PK el mapa.")
        
        MapInfo(.Pos.Map).Pk = isMapPk
        
        'Change the boolean to string in a fast way
        Call WriteVar(App.path & MapPath & "mapa" & .Pos.Map & ".dat", "Mapa" & .Pos.Map, "Pk", IIf(isMapPk, "1", "0"))

        Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " PK: " & MapInfo(.Pos.Map).Pk, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handle the "ChangeMapInfoRestricted" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoRestricted(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Restringido -> Options: "NEWBIE", "NO", "ARMADA", "CAOS", "FACCION".
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    Dim tStr As String
    
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove Packet ID
        Call Buffer.ReadByte
        
        tStr = Buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "NEWBIE" Then
                Call LogGM(.Name, .Name & " ha cambiado la informaci�n sobre si es restringido el mapa.")
                MapInfo(UserList(UserIndex).Pos.Map).Restringir = True
                Call WriteVar(App.path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "Restringir", 1)
                Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Restringido: " & MapInfo(.Pos.Map).Restringir, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Opciones para restringir: 'NEWBIE'", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handle the "ChangeMapInfoNoMagic" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoNoMagic(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'MagiaSinEfecto -> Options: "1" , "0".
'***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim nomagic As Boolean
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        nomagic = .incomingData.ReadBoolean
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.Name, .Name & " ha cambiado la informaci�n sobre si est� permitido usar la magia el mapa.")
            MapInfo(UserList(UserIndex).Pos.Map).MagiaSinEfecto = nomagic
            Call WriteVar(App.path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "MagiaSinEfecto", nomagic)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " MagiaSinEfecto: " & MapInfo(.Pos.Map).MagiaSinEfecto, FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

''
' Handle the "ChangeMapInfoLand" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoLand(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Terreno -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    Dim tStr As String
    
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove Packet ID
        Call Buffer.ReadByte
        
        tStr = Buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
                Call LogGM(.Name, .Name & " ha cambiado la informaci�n del terreno del mapa.")
                MapInfo(UserList(UserIndex).Pos.Map).Terreno = tStr
                Call WriteVar(App.path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "Terreno", tStr)
                Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Terreno: " & MapInfo(.Pos.Map).Terreno, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "Igualmente, el �nico �til es 'NIEVE' ya que al ingresarlo, la gente muere de fr�o en el mapa.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handle the "ChangeMapInfoZone" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMapInfoZone(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Zona -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    Dim tStr As String
    
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove Packet ID
        Call Buffer.ReadByte
        
        tStr = Buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
                Call LogGM(.Name, .Name & " ha cambiado la informaci�n de la zona del mapa.")
                MapInfo(UserList(UserIndex).Pos.Map).Zona = tStr
                Call WriteVar(App.path & MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "Zona", tStr)
                Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Zona: " & MapInfo(.Pos.Map).Zona, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "Igualmente, el �nico �til es 'DUNGEON' ya que al ingresarlo, NO se sentir� el efecto de la lluvia en este mapa.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handle the "SaveMap" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSaveMap(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Saves the map
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha guardado el mapa " & CStr(.Pos.Map))
        
        Call GrabarMapa(.Pos.Map, App.path & "\WorldBackUp\Mapa" & .Pos.Map)
        
        Call WriteConsoleMsg(UserIndex, "Mapa Guardado.", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handle the "DoBackUp" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleDoBackUp(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Show guilds messages
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha hecho un backup.")
        
        Call ES.DoBackUp 'Sino lo confunde con la id del paquete
    End With
End Sub

''
' Handle the "ToggleCentinelActivated" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleToggleCentinelActivated(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/26/06
'Last modified by: Juan Mart�n Sotuyo Dodero (Maraxus)
'Activate or desactivate the Centinel
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        centinelaActivado = Not centinelaActivado
        
        With Centinela
            .RevisandoUserIndex = 0
            .clave = 0
            .TiempoRestante = 0
        End With
    
        If CentinelaNPCIndex Then
            Call QuitarNPC(CentinelaNPCIndex)
            CentinelaNPCIndex = 0
        End If
        
        If centinelaActivado Then
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("El centinela ha sido activado.", FontTypeNames.FONTTYPE_SERVER))
        Else
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("El centinela ha sido desactivado.", FontTypeNames.FONTTYPE_SERVER))
        End If
    End With
End Sub

''
' Handle the "AlterName" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleAlterName(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'Change user name
'***************************************************
    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        'Reads the userName and newUser Packets
        Dim UserName As String
        Dim newName As String
        Dim changeNameUI As Integer
      '  Dim GuildIndex As Integer
        
        UserName = Buffer.ReadASCIIString()
        newName = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If LenB(UserName) = 0 Or LenB(newName) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Usar: /ANAME origen@destino", FontTypeNames.FONTTYPE_INFO)
            Else
                changeNameUI = NameIndex(UserName)
                
                If changeNameUI > 0 Then
                    Call WriteConsoleMsg(UserIndex, "El Pj est� online, debe salir para hacer el cambio.", FontTypeNames.FONTTYPE_WARNING)
                Else
                    If Not FileExist(CharPath & UserName & ".chr") Then
                        Call WriteConsoleMsg(UserIndex, "El pj " & UserName & " es inexistente.", FontTypeNames.FONTTYPE_INFO)
                    Else
                    '    GuildIndex = val(GetVar(CharPath & UserName & ".chr", "GUILD", "GUILDINDEX"))
                     '
                      '  If GuildIndex > 0 Then
                    '        Call WriteConsoleMsg(UserIndex, "El pj " & UserName & " pertenece a un clan, debe salir del mismo con /salirclan para ser transferido.", FontTypeNames.FONTTYPE_INFO)
                    '    Else
                            If Not FileExist(CharPath & newName & ".chr") Then
                                Call FileCopy(CharPath & UserName & ".chr", CharPath & UCase$(newName) & ".chr")
                    
                                Call WriteConsoleMsg(UserIndex, "Transferencia exitosa.", FontTypeNames.FONTTYPE_INFO)
                    
                                Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Ban", "1")
                    
                                Dim cantPenas As Byte
                    
                                cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                    
                                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", CStr(cantPenas + 1))
                    
                                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & CStr(cantPenas + 1), LCase$(.Name) & ": BAN POR Cambio de nick a " & UCase$(newName) & " " & Date & " " & time)
                    
                                Call LogGM(.Name, "Ha cambiado de nombre al usuario " & UserName & ". Ahora se llama " & newName)
                            Else
                                Call WriteConsoleMsg(UserIndex, "El nick solicitado ya existe.", FontTypeNames.FONTTYPE_INFO)
                            End If
                      ' End If
                    End If
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handle the "AlterName" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleAlterMail(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'Change user password
'***************************************************
    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim newMail As String
        
        UserName = Buffer.ReadASCIIString()
        newMail = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If LenB(UserName) = 0 Or LenB(newMail) = 0 Then
                Call WriteConsoleMsg(UserIndex, "usar /AEMAIL <pj>-<nuevomail>", FontTypeNames.FONTTYPE_INFO)
            Else
                If Not FileExist(CharPath & UserName & ".chr") Then
                    Call WriteConsoleMsg(UserIndex, "No existe el charfile " & UserName & ".chr", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteVar(CharPath & UserName & ".chr", "CONTACTO", "Email", newMail)
                    Call WriteConsoleMsg(UserIndex, "Email de " & UserName & " cambiado a: " & newMail, FontTypeNames.FONTTYPE_INFO)
                End If
                
                Call LogGM(.Name, "Le ha cambiado el mail a " & UserName)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handle the "AlterPassword" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleAlterPassword(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'Change user password
'***************************************************
    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim copyFrom As String
        Dim Password As String
        
        UserName = Replace(Buffer.ReadASCIIString(), "+", " ")
        copyFrom = Replace(Buffer.ReadASCIIString(), "+", " ")
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "Ha alterado la contrase�a de " & UserName)
            
            If LenB(UserName) = 0 Or LenB(copyFrom) = 0 Then
                Call WriteConsoleMsg(UserIndex, "usar /APASS <pjsinpass>@<pjconpass>", FontTypeNames.FONTTYPE_INFO)
            Else
                If Not FileExist(CharPath & UserName & ".chr") Or Not FileExist(CharPath & copyFrom & ".chr") Then
                    Call WriteConsoleMsg(UserIndex, "Alguno de los PJs no existe " & UserName & "@" & copyFrom, FontTypeNames.FONTTYPE_INFO)
                Else
                    Password = GetVar(CharPath & copyFrom & ".chr", "INIT", "Password")
                    Call WriteVar(CharPath & UserName & ".chr", "INIT", "Password", Password)
                    
                    Call WriteConsoleMsg(UserIndex, "Password de " & UserName & " ha cambiado por la de " & copyFrom, FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handle the "HandleCreateNPC" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCreateNPC(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim NpcIndex As Integer
        
        NpcIndex = .incomingData.ReadInteger()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        NpcIndex = SpawnNpc(NpcIndex, .Pos, True, False)
        
        If NpcIndex <> 0 Then
            Call LogGM(.Name, "Sumone� a " & Npclist(NpcIndex).Name & " en mapa " & .Pos.Map)
        End If
    End With
End Sub


''
' Handle the "CreateNPCWithRespawn" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCreateNPCWithRespawn(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim NpcIndex As Integer
        
        NpcIndex = .incomingData.ReadInteger()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        NpcIndex = SpawnNpc(NpcIndex, .Pos, True, True)
        
        If NpcIndex <> 0 Then
            Call LogGM(.Name, "Sumone� con respawn " & Npclist(NpcIndex).Name & " en mapa " & .Pos.Map)
        End If
    End With
End Sub

''
' Handle the "NavigateToggle" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleNavigateToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 01/12/07
'
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        If .flags.Navegando = 1 Then
            .flags.Navegando = 0
        Else
            .flags.Navegando = 1
        End If
        
        'Tell the client that we are navigating.
        Call WriteNavigateToggle(UserIndex)
    End With
End Sub

''
' Handle the "ServerOpenToUsersToggle" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleServerOpenToUsersToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        If ServerSoloGMs > 0 Then
            Call WriteConsoleMsg(UserIndex, "Servidor habilitado para todos.", FontTypeNames.FONTTYPE_INFO)
            ServerSoloGMs = 0
        Else
            Call WriteConsoleMsg(UserIndex, "Servidor restringido a administradores.", FontTypeNames.FONTTYPE_INFO)
            ServerSoloGMs = 1
        End If
    End With
End Sub

''
' Handle the "TurnOffServer" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleTurnOffServer(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 12/24/06
'Turns off the server
'***************************************************
    Dim handle As Integer
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, "/APAGAR")
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("���" & .Name & " VA A APAGAR EL SERVIDOR!!!", FontTypeNames.FONTTYPE_FIGHT))
        
        'Log
        handle = FreeFile
        Open App.path & "\logs\Main.log" For Append Shared As #handle
        
        Print #handle, Date & " " & time & " server apagado por " & .Name & ". "
        
        Close #handle
        
        Unload frmMain
    End With
End Sub

''
' Handle the "ResetFactions" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleResetFactions(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 06/09/09
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim tUser As Integer
        Dim Char As String
        
        UserName = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "/RAJAR " & UserName)
            
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                Call ResetFacciones(tUser)
            Else
                Char = CharPath & UserName & ".chr"
                
                If FileExist(Char, vbNormal) Then
                    Call WriteVar(Char, "FACCIONES", "EjercitoReal", 0)
                    Call WriteVar(Char, "FACCIONES", "CiudMatados", 0)
                    Call WriteVar(Char, "FACCIONES", "CrimMatados", 0)
                    Call WriteVar(Char, "FACCIONES", "EjercitoCaos", 0)
                    Call WriteVar(Char, "FACCIONES", "FechaIngreso", "No ingres� a ninguna Facci�n")
                    Call WriteVar(Char, "FACCIONES", "rArCaos", 0)
                    Call WriteVar(Char, "FACCIONES", "rArReal", 0)
                    Call WriteVar(Char, "FACCIONES", "rExCaos", 0)
                    Call WriteVar(Char, "FACCIONES", "rExReal", 0)
                    Call WriteVar(Char, "FACCIONES", "recCaos", 0)
                    Call WriteVar(Char, "FACCIONES", "recReal", 0)
                    Call WriteVar(Char, "FACCIONES", "Reenlistadas", 0)
                    Call WriteVar(Char, "FACCIONES", "NivelIngreso", 0)
                    Call WriteVar(Char, "FACCIONES", "MatadosIngreso", 0)
                    Call WriteVar(Char, "FACCIONES", "NextRecompensa", 0)
                Else
                    Call WriteConsoleMsg(UserIndex, "El personaje " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handle the "RequestCharMail" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleRequestCharMail(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 12/26/06
'Request user mail
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim UserName As String
        Dim mail As String
        
        UserName = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If FileExist(CharPath & UserName & ".chr") Then
                mail = GetVar(CharPath & UserName & ".chr", "CONTACTO", "email")
                
                Call WriteConsoleMsg(UserIndex, "Last email de " & UserName & ":" & mail, FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handle the "SystemMessage" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSystemMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/29/06
'Send a message to all the users
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim message As String
        message = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "Mensaje de sistema:" & message)
            
            Call SendData(SendTarget.ToAll, 0, PrepareMessageShowMessageBox(message))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handle the "SetMOTD" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSetMOTD(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 03/31/07
'Set the MOTD
'Modified by: Juan Mart�n Sotuyo Dodero (Maraxus)
'   - Fixed a bug that prevented from properly setting the new number of lines.
'   - Fixed a bug that caused the player to be kicked.
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim newMOTD As String
        Dim auxiliaryString() As String
        Dim LoopC As Long
        
        newMOTD = Buffer.ReadASCIIString()
        auxiliaryString = Split(newMOTD, vbCrLf)
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "Ha fijado un nuevo MOTD")
            
            MaxLines = UBound(auxiliaryString()) + 1
            
            ReDim MOTD(1 To MaxLines)
            
            Call WriteVar(App.path & "\Dat\Motd.ini", "INIT", "NumLines", CStr(MaxLines))
            
            For LoopC = 1 To MaxLines
                Call WriteVar(App.path & "\Dat\Motd.ini", "Motd", "Line" & CStr(LoopC), auxiliaryString(LoopC - 1))
                
                MOTD(LoopC).texto = auxiliaryString(LoopC - 1)
            Next LoopC
            
            Call WriteConsoleMsg(UserIndex, "Se ha cambiado el MOTD con �xito.", FontTypeNames.FONTTYPE_INFO)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handle the "ChangeMOTD" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMOTD(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n sotuyo Dodero (Maraxus)
'Last Modification: 12/29/06
'Change the MOTD
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If (.flags.Privilegios And (PlayerType.RoleMaster Or PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) Then
            Exit Sub
        End If
        
        Dim auxiliaryString As String
        Dim LoopC As Long
        
        For LoopC = LBound(MOTD()) To UBound(MOTD())
            auxiliaryString = auxiliaryString & MOTD(LoopC).texto & vbCrLf
        Next LoopC
        
        If Len(auxiliaryString) >= 2 Then
            If Right$(auxiliaryString, 2) = vbCrLf Then
                auxiliaryString = Left$(auxiliaryString, Len(auxiliaryString) - 2)
            End If
        End If
        
        Call WriteShowMOTDEditionForm(UserIndex, auxiliaryString)
    End With
End Sub

''
' Handle the "Ping" message
'
' @param userIndex The index of the user sending the message

Public Sub HandlePing(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Tavolaro Ortiz (Tavo)
'Last Modification: 12/24/06
'Show guilds messages
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Call WritePong(UserIndex)
    End With
End Sub

''
' Handle the "SetIniVar" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSetIniVar(ByVal UserIndex As Integer)
'***************************************************
'Author: Brian Chaia (BrianPr)
'Last Modification: 01/23/10 (Marco)
'Modify server.ini
'***************************************************
    If UserList(UserIndex).incomingData.Length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

On Error GoTo Errhandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        
        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim sLlave As String
        Dim sClave As String
        Dim sValor As String

        'Obtengo los par�metros
        sLlave = Buffer.ReadASCIIString()
        sClave = Buffer.ReadASCIIString()
        sValor = Buffer.ReadASCIIString()

        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            Dim sTmp As String

            'No podemos modificar [INIT]Dioses ni [Dioses]*
            If (UCase$(sLlave) = "INIT" And UCase$(sClave) = "DIOSES") Or UCase$(sLlave) = "DIOSES" Then
                Call WriteConsoleMsg(UserIndex, "�No puedes modificar esa informaci�n desde aqu�!", FontTypeNames.FONTTYPE_INFO)
            Else
                'Obtengo el valor seg�n llave y clave
                sTmp = GetVar(IniPath & "Server.ini", sLlave, sClave)

                'Si obtengo un valor escribo en el server.ini
                If LenB(sTmp) Then
                    Call WriteVar(IniPath & "Server.ini", sLlave, sClave, sValor)
                    Call LogGM(.Name, "Modific� en server.ini (" & sLlave & " " & sClave & ") el valor " & sTmp & " por " & sValor)
                    Call WriteConsoleMsg(UserIndex, "Modific� " & sLlave & " " & sClave & " a " & sValor & ". Valor anterior " & sTmp, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "No existe la llave y/o clave", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long

    Error = Err.Number

On Error GoTo 0
    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "WarpToMap" message.
'
' @param map The index of the user sending the message.

Private Sub HandleWarpToMap(ByVal UserIndex As Integer)
'***************************************************
'Author: Lorenzo Rivero (Rhynne)
'Last Modification: 06/01/2017
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim Map As Integer

        Map = .incomingData.ReadInteger()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub
        
        Call WarpUserChar(UserIndex, Map, 50, 50, True)
        Call WriteConsoleMsg(UserIndex, UserList(UserIndex).Name & " transportado.", FontTypeNames.FONTTYPE_INFO)
        Call LogGM(.Name, "Transport� a " & UserList(UserIndex).Name & " hacia " & "Mapa" & Map & " X:" & 50 & " Y:" & 50)
    End With
    
Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "StaffMessage" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleStaffMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Lorenzo Rivero (Rhynne)
'Last Modification: 06/01/2017
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim message As String
        message = Buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)) Then
            If LenB(message) <> 0 Then
                Call LogGM(.Name, "Mensaje a Gms:" & message)
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(UserList(UserIndex).Name & "> " & message, FontTypeNames.FONTTYPE_TALK))
                frmMain.txtChat.Text = frmMain.txtChat.Text & vbNewLine & UserList(UserIndex).Name & "> " & message
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "SearchObjs" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleSearchObjs(ByVal UserIndex As Integer)
'***************************************************
'Author: Lorenzo Rivero (Rhynne)
'Last Modification: 06/01/2017
'***************************************************
    If UserList(UserIndex).incomingData.Length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Obj As String
        Dim N As Integer
        Dim i As Integer
        
        Obj = UCase$(Buffer.ReadASCIIString())
        
        If Len(Obj) > 1 Then
            If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
                For i = 1 To UBound(ObjData)
                    If InStr(1, UCase$(ObjData(i).Name), Obj) Then
                         Call WriteConsoleMsg(UserIndex, i & " - " & ObjData(i).Name, FontTypeNames.FONTTYPE_INFO)
                    N = N + 1
                    End If
                Next
        
                If N = 0 Then
                    Call WriteConsoleMsg(UserIndex, "No hubo resultados de la b�squeda: " & Obj & ".", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "Hubo " & N & " resultados de la b�squeda: " & Obj & ".", FontTypeNames.FONTTYPE_INFO)
                End If

            End If
        Else
            Call WriteConsoleMsg(UserIndex, "Debe usar al menos dos o m�s car�cteres.", FontTypeNames.FONTTYPE_INFO)
            
        End If
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the "Countdown" message.
'
' @param map The index of the user sending the message.

Private Sub HandleCountdown(ByVal UserIndex As Integer)
'***************************************************
'Author: Lorenzo Rivero (Rhynne)
'Last Modification: 06/01/2017
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo Errhandler
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim Count As Byte

        Count = .incomingData.ReadByte()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster And PlayerType.Consejero) Then Exit Sub
        
       ' count = GetTickCount
        
       ' Do While ' �prgrun?
       '     If GetTickCount - count = 1000 Then
                Call SendData(SendTarget.toMap, UserIndex, PrepareMessageConsoleMsg("Cuenta regresiva", FontTypeNames.FONTTYPE_INFO))
      '          GetTickCount = count
       '     End If
       ' Loop
        
    End With
    
Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Writes the "Logged" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoggedMessage(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Logged" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Logged)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "RemoveDialogs" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveAllDialogs(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemoveDialogs" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.RemoveDialogs)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "RemoveCharDialog" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character whose dialog will be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveCharDialog(ByVal UserIndex As Integer, ByVal CharIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemoveCharDialog" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageRemoveCharDialog(CharIndex))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "NavigateToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNavigateToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "NavigateToggle" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.NavigateToggle)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "Disconnect" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDisconnect(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Disconnect" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Disconnect)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UserOfferConfirm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserOfferConfirm(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'Writes the "UserOfferConfirm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.UserOfferConfirm)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub


''
' Writes the "CommerceEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceEnd(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceEnd" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.CommerceEnd)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "BankEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankEnd(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankEnd" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BankEnd)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CommerceInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceInit(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CommerceInit" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.CommerceInit)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "BankInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankInit(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankInit" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BankInit)
    Call UserList(UserIndex).outgoingData.WriteLong(UserList(UserIndex).Stats.Banco)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UserCommerceInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceInit(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserCommerceInit" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.UserCommerceInit)
    Call UserList(UserIndex).outgoingData.WriteASCIIString(UserList(UserIndex).ComUsu.DestNick)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UserCommerceEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceEnd(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserCommerceEnd" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.UserCommerceEnd)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowBlacksmithForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowBlacksmithForm(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowBlacksmithForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowBlacksmithForm)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowCarpenterForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowCarpenterForm(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowCarpenterForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowCarpenterForm)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateSta" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateSta(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateMana" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateSta)
        Call .WriteInteger(UserList(UserIndex).Stats.MinSta)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateMana" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateMana(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateMana" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateMana)
        Call .WriteInteger(UserList(UserIndex).Stats.MinMAN)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateHP" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHP(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateMana" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateHP)
        Call .WriteInteger(UserList(UserIndex).Stats.MinHp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateGold" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateGold(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateGold" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateGold)
        Call .WriteLong(UserList(UserIndex).Stats.GLD)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateBankGold" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateBankGold(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'Writes the "UpdateBankGold" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateBankGold)
        Call .WriteLong(UserList(UserIndex).Stats.Banco)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub


''
' Writes the "UpdateExp" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateExp(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateExp" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateExp)
        Call .WriteLong(UserList(UserIndex).Stats.Exp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateStrenghtAndDexterity(ByVal UserIndex As Integer)
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateStrenghtAndDexterity)
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

' Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateDexterity(ByVal UserIndex As Integer)
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateDexterity)
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

' Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateStrenght(ByVal UserIndex As Integer)
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateStrenght)
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ChangeMap" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    map The new map to load.
' @param    version The version of the map in the server to check if client is properly updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMap(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal version As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeMap" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeMap)
        Call .WriteInteger(Map)
        Call .WriteInteger(version)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "PosUpdate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePosUpdate(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PosUpdate" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.PosUpdate)
        Call .WriteByte(UserList(UserIndex).Pos.X)
        Call .WriteByte(UserList(UserIndex).Pos.Y)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ChatOverHead" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    CharIndex The character uppon which the chat will be displayed.
' @param    Color The color to be used when displaying the chat.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChatOverHead(ByVal UserIndex As Integer, ByVal Chat As String, ByVal CharIndex As Integer, ByVal color As Long)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChatOverHead" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageChatOverHead(Chat, CharIndex, color))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ConsoleMsg" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteConsoleMsg(ByVal UserIndex As Integer, ByVal Chat As String, ByVal FontIndex As FontTypeNames)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ConsoleMsg" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageConsoleMsg(Chat, FontIndex))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteCommerceChat(ByVal UserIndex As Integer, ByVal Chat As String, ByVal FontIndex As FontTypeNames)
'***************************************************
'Author: ZaMa
'Last Modification: 05/17/06
'Writes the "ConsoleMsg" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareCommerceConsoleMsg(Chat, FontIndex))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub
            
''
' Writes the "GuildChat" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @remarks  The data is not actually sent until the buffer is properly flushed.

'Public Sub WriteGuildChat(ByVal UserIndex As Integer, ByVal Chat As String)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "GuildChat" message to the given user's outgoing data buffer
'***************************************************
'On Error GoTo Errhandler
'    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageGuildChat(Chat))
'Exit Sub

'Errhandler:
'    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
''        Call FlushBuffer(UserIndex)
'        Resume
'    End If
'End Sub

''
' Writes the "ShowMessageBox" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Message Text to be displayed in the message box.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowMessageBox(ByVal UserIndex As Integer, ByVal message As String)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowMessageBox" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowMessageBox)
        Call .WriteASCIIString(message)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UserIndexInServer" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserIndexInServer(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserIndexInServer" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserIndexInServer)
        Call .WriteInteger(UserIndex)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UserCharIndexInServer" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCharIndexInServer(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UserIndexInServer" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserCharIndexInServer)
        Call .WriteInteger(UserList(UserIndex).Char.CharIndex)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CharacterCreate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    X X coord of the new character's position.
' @param    Y Y coord of the new character's position.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @param    name Name of the new character.
' @param    criminal Determines if the character is a criminal or not.
' @param    privileges Sets if the character is a normal one or any kind of administrative character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterCreate(ByVal UserIndex As Integer, ByVal body As Integer, ByVal Head As Integer, ByVal heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal weapon As Integer, ByVal shield As Integer, _
                                ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer, ByVal Name As String, ByVal Guild As String, _
                                ByVal NickColor As Byte, ByVal Privileges As Byte)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterCreate" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterCreate(body, Head, heading, CharIndex, X, Y, weapon, shield, FX, FXLoops, _
                                                            helmet, Name, Guild, NickColor, Privileges))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CharacterRemove" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterRemove(ByVal UserIndex As Integer, ByVal CharIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterRemove" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterRemove(CharIndex))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CharacterMove" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterMove(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterMove" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterMove(CharIndex, X, Y))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteForceCharMove(ByVal UserIndex, ByVal Direccion As eHeading)
'***************************************************
'Author: ZaMa
'Last Modification: 26/03/2009
'Writes the "ForceCharMove" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageForceCharMove(Direccion))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CharacterChange" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterChange(ByVal UserIndex As Integer, ByVal body As Integer, ByVal Head As Integer, ByVal heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal weapon As Integer, ByVal shield As Integer, _
                                ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CharacterChange" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterChange(body, Head, heading, CharIndex, weapon, shield, FX, FXLoops, helmet))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ObjectCreate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    GrhIndex Grh of the object.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteObjectCreate(ByVal UserIndex As Integer, ByVal GrhIndex As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ObjectCreate" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageObjectCreate(GrhIndex, X, Y))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ObjectDelete" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteObjectDelete(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ObjectDelete" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageObjectDelete(X, Y))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "BlockPosition" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @param    Blocked True if the position is blocked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlockPosition(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal Blocked As Boolean)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BlockPosition" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.BlockPosition)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteBoolean(Blocked)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "PlayMidi" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePlayMidi(ByVal UserIndex As Integer, ByVal midi As Byte, Optional ByVal loops As Integer = -1)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PlayMidi" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayMidi(midi, loops))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "PlayWave" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    wave The wave to be played.
' @param    X The X position in map coordinates from where the sound comes.
' @param    Y The Y position in map coordinates from where the sound comes.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePlayWave(ByVal UserIndex As Integer, ByVal wave As Byte, ByVal X As Byte, ByVal Y As Byte)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 08/08/07
'Last Modified by: Rapsodius
'Added X and Y positions for 3D Sounds
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(wave, X, Y))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub


''
' Writes the "AreaChanged" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAreaChanged(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "AreaChanged" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AreaChanged)
        Call .WriteByte(UserList(UserIndex).Pos.X)
        Call .WriteByte(UserList(UserIndex).Pos.Y)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "PauseToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePauseToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "PauseToggle" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePauseToggle())
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "RainToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRainToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RainToggle" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageRainToggle())
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CreateFX" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character upon which the FX will be created.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateFX(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal FX As Integer, ByVal FXLoops As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CreateFX" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCreateFX(CharIndex, FX, FXLoops))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateUserStats" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateUserStats(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateUserStats" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateUserStats)
        Call .WriteInteger(UserList(UserIndex).Stats.MaxHp)
        Call .WriteInteger(UserList(UserIndex).Stats.MinHp)
        Call .WriteInteger(UserList(UserIndex).Stats.MaxMAN)
        Call .WriteInteger(UserList(UserIndex).Stats.MinMAN)
        Call .WriteInteger(UserList(UserIndex).Stats.MaxSta)
        Call .WriteInteger(UserList(UserIndex).Stats.MinSta)
        Call .WriteLong(UserList(UserIndex).Stats.GLD)
        Call .WriteByte(UserList(UserIndex).Stats.ELV)
        Call .WriteLong(UserList(UserIndex).Stats.ELU)
        Call .WriteLong(UserList(UserIndex).Stats.Exp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "WorkRequestTarget" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Skill The skill for which we request a target.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorkRequestTarget(ByVal UserIndex As Integer, ByVal Skill As eSkill)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "WorkRequestTarget" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.WorkRequestTarget)
        Call .WriteByte(Skill)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ChangeInventorySlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Inventory slot which needs to be updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeInventorySlot(ByVal UserIndex As Integer, ByVal Slot As Byte)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 3/12/09
'Writes the "ChangeInventorySlot" message to the given user's outgoing data buffer
'3/12/09: Budi - Ahora se envia MaxDef y MinDef en lugar de Def
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeInventorySlot)
        Call .WriteByte(Slot)
        
        Dim OBJIndex As Integer
        Dim obData As ObjData
        
        OBJIndex = UserList(UserIndex).Invent.Object(Slot).OBJIndex
        
        If OBJIndex > 0 Then
            obData = ObjData(OBJIndex)
        End If
        
        Call .WriteInteger(OBJIndex)
        Call .WriteASCIIString(obData.Name)
        Call .WriteInteger(UserList(UserIndex).Invent.Object(Slot).Amount)
        Call .WriteBoolean(UserList(UserIndex).Invent.Object(Slot).Equipped)
        Call .WriteInteger(obData.GrhIndex)
        Call .WriteByte(obData.OBJType)
        Call .WriteInteger(obData.MaxHIT)
        Call .WriteInteger(obData.MinHIT)
        Call .WriteInteger(obData.MaxDef)
        Call .WriteInteger(obData.MinDef)
        Call .WriteSingle(SalePrice(OBJIndex))
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteAddSlots(ByVal UserIndex As Integer, ByVal Mochila As eMochilas)
'***************************************************
'Author: Budi
'Last Modification: 01/12/09
'Writes the "AddSlots" message to the given user's outgoing data buffer
'***************************************************
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AddSlots)
        Call .WriteByte(Mochila)
    End With
End Sub


''
' Writes the "ChangeBankSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Inventory slot which needs to be updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeBankSlot(ByVal UserIndex As Integer, ByVal Slot As Byte)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 12/03/09
'Writes the "ChangeBankSlot" message to the given user's outgoing data buffer
'12/03/09: Budi - Ahora se envia MaxDef y MinDef en lugar de s�lo Def
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeBankSlot)
        Call .WriteByte(Slot)
        
        Dim OBJIndex As Integer
        Dim obData As ObjData
        
        OBJIndex = UserList(UserIndex).BancoInvent.Object(Slot).OBJIndex
        
        Call .WriteInteger(OBJIndex)
        
        If OBJIndex > 0 Then
            obData = ObjData(OBJIndex)
        End If
        
        Call .WriteASCIIString(obData.Name)
        Call .WriteInteger(UserList(UserIndex).BancoInvent.Object(Slot).Amount)
        Call .WriteInteger(obData.GrhIndex)
        Call .WriteByte(obData.OBJType)
        Call .WriteInteger(obData.MaxHIT)
        Call .WriteInteger(obData.MinHIT)
        Call .WriteInteger(obData.MaxDef)
        Call .WriteInteger(obData.MinDef)
        Call .WriteLong(obData.Valor)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ChangeSpellSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Spell slot to update.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeSpellSlot(ByVal UserIndex As Integer, ByVal Slot As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeSpellSlot" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeSpellSlot)
        Call .WriteByte(Slot)
        Call .WriteInteger(UserList(UserIndex).Stats.UserHechizos(Slot))
        
        If UserList(UserIndex).Stats.UserHechizos(Slot) > 0 Then
            Call .WriteASCIIString(Hechizos(UserList(UserIndex).Stats.UserHechizos(Slot)).Nombre)
        Else
            Call .WriteASCIIString("(None)")
        End If
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "Atributes" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAttributes(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Atributes" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.Atributes)
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion))
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "BlacksmithWeapons" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlacksmithWeapons(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 04/15/2008 (NicoNZ) Habia un error al fijarse los skills del personaje
'Writes the "BlacksmithWeapons" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim Obj As ObjData
    Dim validIndexes() As Integer
    Dim Count As Integer
    
    ReDim validIndexes(1 To UBound(ArmasHerrero()))
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.BlacksmithWeapons)
        
        For i = 1 To UBound(ArmasHerrero())
            ' Can the user create this object? If so add it to the list....
            If ObjData(ArmasHerrero(i)).SkHerreria <= Round(UserList(UserIndex).Stats.UserSkills(eSkill.Herreria) / ModHerreria(UserList(UserIndex).Clase), 0) Then
                Count = Count + 1
                validIndexes(Count) = i
            End If
        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(Count)
        
        ' Write the needed data of each object
        For i = 1 To Count
            Obj = ObjData(ArmasHerrero(validIndexes(i)))
            Call .WriteASCIIString(Obj.Name)
            Call .WriteInteger(Obj.GrhIndex)
            Call .WriteInteger(Obj.LingH)
            Call .WriteInteger(Obj.LingP)
            Call .WriteInteger(Obj.LingO)
            Call .WriteInteger(ArmasHerrero(validIndexes(i)))
        Next i
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "BlacksmithArmors" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlacksmithArmors(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 04/15/2008 (NicoNZ) Habia un error al fijarse los skills del personaje
'Writes the "BlacksmithArmors" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim Obj As ObjData
    Dim validIndexes() As Integer
    Dim Count As Integer
    
    ReDim validIndexes(1 To UBound(ArmadurasHerrero()))
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.BlacksmithArmors)
        
        For i = 1 To UBound(ArmadurasHerrero())
            ' Can the user create this object? If so add it to the list....
            If ObjData(ArmadurasHerrero(i)).SkHerreria <= Round(UserList(UserIndex).Stats.UserSkills(eSkill.Herreria) / ModHerreria(UserList(UserIndex).Clase), 0) Then
                Count = Count + 1
                validIndexes(Count) = i
            End If
        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(Count)
        
        ' Write the needed data of each object
        For i = 1 To Count
            Obj = ObjData(ArmadurasHerrero(validIndexes(i)))
            Call .WriteASCIIString(Obj.Name)
            Call .WriteInteger(Obj.GrhIndex)
            Call .WriteInteger(Obj.LingH)
            Call .WriteInteger(Obj.LingP)
            Call .WriteInteger(Obj.LingO)
            Call .WriteInteger(ArmadurasHerrero(validIndexes(i)))
        Next i
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CarpenterObjects" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCarpenterObjects(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "CarpenterObjects" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim Obj As ObjData
    Dim validIndexes() As Integer
    Dim Count As Integer
    
    ReDim validIndexes(1 To UBound(ObjCarpintero()))
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.CarpenterObjects)
        
        For i = 1 To UBound(ObjCarpintero())
            ' Can the user create this object? If so add it to the list....
            If ObjData(ObjCarpintero(i)).SkCarpinteria <= UserList(UserIndex).Stats.UserSkills(eSkill.Carpinteria) \ ModCarpinteria(UserList(UserIndex).Clase) Then
                Count = Count + 1
                validIndexes(Count) = i
            End If
        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(Count)
        
        ' Write the needed data of each object
        For i = 1 To Count
            Obj = ObjData(ObjCarpintero(validIndexes(i)))
            Call .WriteASCIIString(Obj.Name)
            Call .WriteInteger(Obj.GrhIndex)
            Call .WriteInteger(Obj.Madera)
            Call .WriteInteger(Obj.MaderaElfica)
            Call .WriteInteger(ObjCarpintero(validIndexes(i)))
        Next i
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "RestOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRestOK(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RestOK" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.RestOK)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ErrorMsg" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteErrorMsg(ByVal UserIndex As Integer, ByVal message As String)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ErrorMsg" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageErrorMsg(message))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "Blind" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlind(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Blind" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Blind)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "Dumb" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumb(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Dumb" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Dumb)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowSignal" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    objIndex Index of the signal to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowSignal(ByVal UserIndex As Integer, ByVal OBJIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowSignal" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowSignal)
        Call .WriteASCIIString(ObjData(OBJIndex).texto)
        Call .WriteInteger(ObjData(OBJIndex).GrhSecundario)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ChangeNPCInventorySlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex   User to which the message is intended.
' @param    slot        The inventory slot in which this item is to be placed.
' @param    obj         The object to be set in the NPC's inventory window.
' @param    price       The value the NPC asks for the object.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeNPCInventorySlot(ByVal UserIndex As Integer, ByVal Slot As Byte, ByRef Obj As Obj, ByVal price As Single)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 12/03/09
'Last Modified by: Budi
'Writes the "ChangeNPCInventorySlot" message to the given user's outgoing data buffer
'12/03/09: Budi - Ahora se envia MaxDef y MinDef en lugar de s�lo Def
'***************************************************
On Error GoTo Errhandler
    Dim ObjInfo As ObjData
    
    If Obj.OBJIndex >= LBound(ObjData()) And Obj.OBJIndex <= UBound(ObjData()) Then
        ObjInfo = ObjData(Obj.OBJIndex)
    End If
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeNPCInventorySlot)
        Call .WriteByte(Slot)
        Call .WriteASCIIString(ObjInfo.Name)
        Call .WriteInteger(Obj.Amount)
        Call .WriteSingle(price)
        Call .WriteInteger(ObjInfo.GrhIndex)
        Call .WriteInteger(Obj.OBJIndex)
        Call .WriteByte(ObjInfo.OBJType)
        Call .WriteInteger(ObjInfo.MaxHIT)
        Call .WriteInteger(ObjInfo.MinHIT)
        Call .WriteInteger(ObjInfo.MaxDef)
        Call .WriteInteger(ObjInfo.MinDef)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UpdateHungerAndThirst" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHungerAndThirst(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "UpdateHungerAndThirst" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateHungerAndThirst)
        Call .WriteByte(UserList(UserIndex).Stats.MaxAGU)
        Call .WriteByte(UserList(UserIndex).Stats.MinAGU)
        Call .WriteByte(UserList(UserIndex).Stats.MaxHam)
        Call .WriteByte(UserList(UserIndex).Stats.MinHam)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "Fame" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteFame(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Fame" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.Fame)

    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "MiniStats" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMiniStats(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "MiniStats" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.MiniStats)
        
        Call .WriteLong(UserList(UserIndex).Faccion.Matados(eFaccion.Neutral))
        Call .WriteLong(UserList(UserIndex).Faccion.Matados(eFaccion.Caos))
        Call .WriteLong(UserList(UserIndex).Faccion.Matados(eFaccion.Real))
        
'TODO : Este valor es calculable, no deber�a NI EXISTIR, ya sea en el servidor ni en el cliente!!!
        Call .WriteLong(UserList(UserIndex).Stats.UsuariosMatados)
        
        Call .WriteInteger(UserList(UserIndex).Stats.NPCsMuertos)
        
        Call .WriteByte(UserList(UserIndex).Clase)
        Call .WriteLong(UserList(UserIndex).Counters.Pena)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "LevelUp" message to the given user's outgoing data buffer.
'
' @param    skillPoints The number of free skill points the player has.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLevelUp(ByVal UserIndex As Integer, ByVal skillPoints As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "LevelUp" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.LevelUp)
        Call .WriteInteger(skillPoints)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "AddForumMsg" message to the given user's outgoing data buffer.
'
' @param    title The title of the message to display.
' @param    message The message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAddForumMsg(ByVal UserIndex As Integer, ByVal ForumType As eForumType, _
                    ByRef Title As String, ByRef Author As String, ByRef message As String)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 02/01/2010
'Writes the "AddForumMsg" message to the given user's outgoing data buffer
'02/01/2010: ZaMa - Now sends Author and forum type
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AddForumMsg)
        Call .WriteByte(ForumType)
        Call .WriteASCIIString(Title)
        Call .WriteASCIIString(Author)
        Call .WriteASCIIString(message)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowForumForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowForumForm(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowForumForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler

    Dim Visibilidad As Byte
    Dim CanMakeSticky As Byte
    
    With UserList(UserIndex)
        Call .outgoingData.WriteByte(ServerPacketID.ShowForumForm)
        
        Visibilidad = eForumVisibility.ieGENERAL_MEMBER
        
        If EsCaos(UserIndex) Or EsGM(UserIndex) Then
            Visibilidad = Visibilidad Or eForumVisibility.ieCAOS_MEMBER
        End If
        
        If EsArmada(UserIndex) Or EsGM(UserIndex) Then
            Visibilidad = Visibilidad Or eForumVisibility.ieREAL_MEMBER
        End If
        
        Call .outgoingData.WriteByte(Visibilidad)
        
        ' Pueden mandar sticky los gms o los del consejo de armada/caos
        If EsGM(UserIndex) Then
            CanMakeSticky = 2
        ElseIf (.flags.Privilegios And PlayerType.ChaosCouncil) <> 0 Then
            CanMakeSticky = 1
        ElseIf (.flags.Privilegios And PlayerType.RoyalCouncil) <> 0 Then
            CanMakeSticky = 1
        End If
        
        Call .outgoingData.WriteByte(CanMakeSticky)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "SetInvisible" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetInvisible(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal invisible As Boolean)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SetInvisible" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageSetInvisible(CharIndex, invisible))
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "DiceRoll" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDiceRoll(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DiceRoll" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.DiceRoll)
        
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion))
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "MeditateToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMeditateToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "MeditateToggle" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.MeditateToggle)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "BlindNoMore" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlindNoMore(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BlindNoMore" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BlindNoMore)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "DumbNoMore" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumbNoMore(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "DumbNoMore" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.DumbNoMore)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "SendSkills" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSendSkills(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 11/19/09
'Writes the "SendSkills" message to the given user's outgoing data buffer
'11/19/09: Pato - Now send the percentage of progress of the skills.
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    
    With UserList(UserIndex)
        Call .outgoingData.WriteByte(ServerPacketID.SendSkills)
        Call .outgoingData.WriteByte(.Clase)
        
        For i = 1 To NUMSKILLS
            Call .outgoingData.WriteByte(UserList(UserIndex).Stats.UserSkills(i))
        Next i
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "TrainerCreatureList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    npcIndex The index of the requested trainer.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrainerCreatureList(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TrainerCreatureList" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim str As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.TrainerCreatureList)
        
        For i = 1 To Npclist(NpcIndex).NroCriaturas
            str = str & Npclist(NpcIndex).Criaturas(i).NpcName & SEPARATOR
        Next i
        
        If LenB(str) > 0 Then _
            str = Left$(str, Len(str) - 1)
        
        Call .WriteASCIIString(str)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ParalizeOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteParalizeOK(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 08/12/07
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
'Writes the "ParalizeOK" message to the given user's outgoing data buffer
'And updates user position
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ParalizeOK)
    Call WritePosUpdate(UserIndex)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowUserRequest" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    details DEtails of the char's request.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowUserRequest(ByVal UserIndex As Integer, ByVal details As String)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowUserRequest" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowUserRequest)
        
        Call .WriteASCIIString(details)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "TradeOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTradeOK(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "TradeOK" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.TradeOK)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "BankOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankOK(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "BankOK" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BankOK)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ChangeUserTradeSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    ObjIndex The object's index.
' @param    amount The number of objects offered.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeUserTradeSlot(ByVal UserIndex As Integer, ByVal OfferSlot As Byte, ByVal OBJIndex As Integer, ByVal Amount As Long)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 12/03/09
'Writes the "ChangeUserTradeSlot" message to the given user's outgoing data buffer
'25/11/2009: ZaMa - Now sends the specific offer slot to be modified.
'12/03/09: Budi - Ahora se envia MaxDef y MinDef en lugar de s�lo Def
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeUserTradeSlot)
        
        Call .WriteByte(OfferSlot)
        Call .WriteInteger(OBJIndex)
        Call .WriteLong(Amount)
        
        If OBJIndex > 0 Then
            Call .WriteInteger(ObjData(OBJIndex).GrhIndex)
            Call .WriteByte(ObjData(OBJIndex).OBJType)
            Call .WriteInteger(ObjData(OBJIndex).MaxHIT)
            Call .WriteInteger(ObjData(OBJIndex).MinHIT)
            Call .WriteInteger(ObjData(OBJIndex).MaxDef)
            Call .WriteInteger(ObjData(OBJIndex).MinDef)
            Call .WriteLong(SalePrice(OBJIndex))
            Call .WriteASCIIString(ObjData(OBJIndex).Name)
        Else ' Borra el item
            Call .WriteInteger(0)
            Call .WriteByte(0)
            Call .WriteInteger(0)
            Call .WriteInteger(0)
            Call .WriteInteger(0)
            Call .WriteInteger(0)
            Call .WriteLong(0)
            Call .WriteASCIIString("")
        End If
    End With
Exit Sub


Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "SendNight" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSendNight(ByVal UserIndex As Integer, ByVal night As Boolean)
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/08/07
'Writes the "SendNight" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.SendNight)
        Call .WriteBoolean(night)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "SpawnList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    npcNames The names of the creatures that can be spawned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnList(ByVal UserIndex As Integer, ByRef npcNames() As String)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "SpawnList" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.SpawnList)
        
        For i = LBound(npcNames()) To UBound(npcNames())
            Tmp = Tmp & npcNames(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowSOSForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowSOSForm(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowSOSForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowSOSForm)
        
        For i = 1 To Ayuda.Longitud
            Tmp = Tmp & Ayuda.VerElemento(i) & SEPARATOR
        Next i
        
        If LenB(Tmp) <> 0 Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub


''
' Writes the "ShowMOTDEditionForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    currentMOTD The current Message Of The Day.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowMOTDEditionForm(ByVal UserIndex As Integer, ByVal currentMOTD As String)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowMOTDEditionForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowMOTDEditionForm)
        
        Call .WriteASCIIString(currentMOTD)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "ShowGMPanelForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGMPanelForm(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowGMPanelForm" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowGMPanelForm)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "UserNameList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    userNameList List of user names.
' @param    Cant Number of names to send.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserNameList(ByVal UserIndex As Integer, ByRef userNamesList() As String, ByVal cant As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06 NIGO:
'Writes the "UserNameList" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Dim i As Long
    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserNameList)
        
        ' Prepare user's names list
        For i = 1 To cant
            Tmp = Tmp & userNamesList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then _
            Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "Pong" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePong(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "Pong" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo Errhandler
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Pong)
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Flushes the outgoing data buffer of the user.
'
' @param    UserIndex User whose outgoing data buffer will be flushed.

Public Sub FlushBuffer(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Sends all data existing in the buffer
'***************************************************
    Dim sndData As String
    
    With UserList(UserIndex).outgoingData
        If .Length = 0 Then _
            Exit Sub

        sndData = .ReadASCIIStringFixed(.Length)
        
        Call EnviarDatosASlot(UserIndex, sndData)
    End With
End Sub

''
' Prepares the "SetInvisible" message and returns it.
'
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The message is written to no outgoing buffer, but only prepared in a single string to be easily sent to several clients.

Public Function PrepareMessageSetInvisible(ByVal CharIndex As Integer, ByVal invisible As Boolean) As String
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "SetInvisible" message and returns it.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.SetInvisible)
        
        Call .WriteInteger(CharIndex)
        Call .WriteBoolean(invisible)
        
        PrepareMessageSetInvisible = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareMessageCharacterChangeNick(ByVal CharIndex As Integer, ByVal newNick As String) As String
'***************************************************
'Author: Budi
'Last Modification: 07/23/09
'Prepares the "Change Nick" message and returns it.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterChangeNick)
        
        Call .WriteInteger(CharIndex)
        Call .WriteASCIIString(newNick)
        
        PrepareMessageCharacterChangeNick = .ReadASCIIStringFixed(.Length)
    End With
End Function

''
' Prepares the "ChatOverHead" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    CharIndex The character uppon which the chat will be displayed.
' @param    Color The color to be used when displaying the chat.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The message is written to no outgoing buffer, but only prepared in a single string to be easily sent to several clients.

Public Function PrepareMessageChatOverHead(ByVal Chat As String, ByVal CharIndex As Integer, ByVal color As Long) As String
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "ChatOverHead" message and returns it.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ChatOverHead)
        Call .WriteASCIIString(Chat)
        Call .WriteInteger(CharIndex)
        
        ' Write rgb channels and save one byte from long :D
        Call .WriteByte(color And &HFF)
        Call .WriteByte((color And &HFF00&) \ &H100&)
        Call .WriteByte((color And &HFF0000) \ &H10000)
        
        PrepareMessageChatOverHead = .ReadASCIIStringFixed(.Length)
    End With
End Function

''
' Prepares the "ConsoleMsg" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageConsoleMsg(ByVal Chat As String, ByVal FontIndex As FontTypeNames) As String
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "ConsoleMsg" message and returns it.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ConsoleMsg)
        Call .WriteASCIIString(Chat)
        Call .WriteByte(FontIndex)
        
        PrepareMessageConsoleMsg = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareCommerceConsoleMsg(ByRef Chat As String, ByVal FontIndex As FontTypeNames) As String
'***************************************************
'Author: ZaMa
'Last Modification: 03/12/2009
'Prepares the "CommerceConsoleMsg" message and returns it.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CommerceChat)
        Call .WriteASCIIString(Chat)
        Call .WriteByte(FontIndex)
        
        PrepareCommerceConsoleMsg = .ReadASCIIStringFixed(.Length)
    End With
End Function

''
' Prepares the "CreateFX" message and returns it.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character upon which the FX will be created.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCreateFX(ByVal CharIndex As Integer, ByVal FX As Integer, ByVal FXLoops As Integer) As String
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CreateFX" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CreateFX)
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)
        
        PrepareMessageCreateFX = .ReadASCIIStringFixed(.Length)
    End With
End Function

''
' Prepares the "PlayWave" message and returns it.
'
' @param    wave The wave to be played.
' @param    X The X position in map coordinates from where the sound comes.
' @param    Y The Y position in map coordinates from where the sound comes.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePlayWave(ByVal wave As Byte, ByVal X As Byte, ByVal Y As Byte) As String
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 08/08/07
'Last Modified by: Rapsodius
'Added X and Y positions for 3D Sounds
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PlayWave)
        Call .WriteByte(wave)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        
        PrepareMessagePlayWave = .ReadASCIIStringFixed(.Length)
    End With
End Function


''
' Prepares the "ShowMessageBox" message and returns it.
'
' @param    Message Text to be displayed in the message box.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageShowMessageBox(ByVal Chat As String) As String
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/08/07
'Prepares the "ShowMessageBox" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ShowMessageBox)
        Call .WriteASCIIString(Chat)
        
        PrepareMessageShowMessageBox = .ReadASCIIStringFixed(.Length)
    End With
End Function


''
' Prepares the "PlayMidi" message and returns it.
'
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePlayMidi(ByVal midi As Byte, Optional ByVal loops As Integer = -1) As String
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "GuildChat" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PlayMIDI)
        Call .WriteByte(midi)
        Call .WriteInteger(loops)
        
        PrepareMessagePlayMidi = .ReadASCIIStringFixed(.Length)
    End With
End Function

''
' Prepares the "PauseToggle" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePauseToggle() As String
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "PauseToggle" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PauseToggle)
        PrepareMessagePauseToggle = .ReadASCIIStringFixed(.Length)
    End With
End Function

''
' Prepares the "RainToggle" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageRainToggle() As String
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "RainToggle" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.RainToggle)
        
        PrepareMessageRainToggle = .ReadASCIIStringFixed(.Length)
    End With
End Function

''
' Prepares the "ObjectDelete" message and returns it.
'
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageObjectDelete(ByVal X As Byte, ByVal Y As Byte) As String
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "ObjectDelete" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ObjectDelete)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        
        PrepareMessageObjectDelete = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareMessageMultiMessage(ByVal MessageIndex As Integer, Optional ByVal Arg1 As Long, Optional ByVal Arg2 As Long, Optional ByVal Arg3 As Long, Optional ByVal StringArg1 As String)

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.MultiMessage)
        Call .WriteByte(MessageIndex)
        
        Select Case MessageIndex
            Case eMessages.DontSeeAnything, eMessages.NPCSwing, eMessages.NPCKillUser, eMessages.BlockedWithShieldUser, _
                eMessages.BlockedWithShieldother, eMessages.UserSwing, eMessages.NobilityLost, _
                eMessages.CantUseWhileMeditating, eMessages.CancelHome, eMessages.FinishHome
            
            Case eMessages.NPCHitUser
                Call .WriteByte(Arg1) 'Target
                Call .WriteInteger(Arg2) 'damage
                
            Case eMessages.UserHitNPC
                Call .WriteLong(Arg1) 'damage
                
            Case eMessages.UserAttackedSwing
                Call .WriteInteger(UserList(Arg1).Char.CharIndex)
                
            Case eMessages.UserHittedByUser
                Call .WriteInteger(Arg1) 'AttackerIndex
                Call .WriteByte(Arg2) 'Target
                Call .WriteInteger(Arg3) 'damage
                
            Case eMessages.UserHittedUser
                Call .WriteInteger(Arg1) 'AttackerIndex
                Call .WriteByte(Arg2) 'Target
                Call .WriteInteger(Arg3) 'damage
                
            Case eMessages.WorkRequestTarget
                Call .WriteByte(Arg1) 'skill
            
            Case eMessages.HaveKilledUser '"Has matado a " & UserList(VictimIndex).name & "!" "Has ganado " & DaExp & " puntos de experiencia."
                Call .WriteInteger(UserList(Arg1).Char.CharIndex) 'VictimIndex
                Call .WriteLong(Arg2) 'Expe
            
            Case eMessages.UserKill '"�" & .name & " te ha matado!"
                Call .WriteInteger(UserList(Arg1).Char.CharIndex) 'AttackerIndex
            
            Case eMessages.EarnExp
            
            Case eMessages.Home
                Call .WriteByte(CByte(Arg1))
                Call .WriteInteger(CInt(Arg2))
                'El cliente no conoce nada sobre nombre de mapas y hogares, por lo tanto _
                 hasta que no se pasen los dats e .INFs al cliente, esto queda as�.
                Call .WriteASCIIString(StringArg1) 'Call .WriteByte(CByte(Arg2))
            
            Case eMessages.GuildCreated
                Call .WriteASCIIString(StringArg1)
                
        End Select
        
        PrepareMessageMultiMessage = .ReadASCIIStringFixed(.Length)
    End With
End Function

''
' Prepares the "BlockPosition" message and returns it.
'
' @param    X X coord of the tile to block/unblock.
' @param    Y Y coord of the tile to block/unblock.
' @param    Blocked Blocked status of the tile
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageBlockPosition(ByVal X As Byte, ByVal Y As Byte, ByVal Blocked As Boolean) As String
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/08/07
'Prepares the "BlockPosition" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.BlockPosition)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteBoolean(Blocked)
        
        PrepareMessageBlockPosition = .ReadASCIIStringFixed(.Length)
    End With
    
End Function

''
' Prepares the "ObjectCreate" message and returns it.
'
' @param    GrhIndex Grh of the object.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageObjectCreate(ByVal GrhIndex As Integer, ByVal X As Byte, ByVal Y As Byte) As String
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'prepares the "ObjectCreate" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ObjectCreate)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteInteger(GrhIndex)
        
        PrepareMessageObjectCreate = .ReadASCIIStringFixed(.Length)
    End With
End Function

''
' Prepares the "CharacterRemove" message and returns it.
'
' @param    CharIndex Character to be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterRemove(ByVal CharIndex As Integer) As String
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CharacterRemove" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterRemove)
        Call .WriteInteger(CharIndex)
        
        PrepareMessageCharacterRemove = .ReadASCIIStringFixed(.Length)
    End With
End Function

''
' Prepares the "RemoveCharDialog" message and returns it.
'
' @param    CharIndex Character whose dialog will be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageRemoveCharDialog(ByVal CharIndex As Integer) As String
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "RemoveCharDialog" message to the given user's outgoing data buffer
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.RemoveCharDialog)
        Call .WriteInteger(CharIndex)
        
        PrepareMessageRemoveCharDialog = .ReadASCIIStringFixed(.Length)
    End With
End Function

''
' Writes the "CharacterCreate" message to the given user's outgoing data buffer.
'
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    X X coord of the new character's position.
' @param    Y Y coord of the new character's position.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @param    name Name of the new character.
' @param    NickColor Determines if the character is a criminal or not, and if can be atacked by someone
' @param    privileges Sets if the character is a normal one or any kind of administrative character.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterCreate(ByVal body As Integer, ByVal Head As Integer, ByVal heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal weapon As Integer, ByVal shield As Integer, _
                                ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer, ByVal Name As String, ByVal Guild As String, _
                                ByVal NickColor As Byte, ByVal Privileges As Byte) As String
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CharacterCreate" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterCreate)
        
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(body)
        Call .WriteInteger(Head)
        Call .WriteByte(heading)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteInteger(weapon)
        Call .WriteInteger(shield)
        Call .WriteInteger(helmet)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)
        Call .WriteASCIIString(Name)
        Call .WriteASCIIString(Guild)
        Call .WriteByte(NickColor)
        Call .WriteByte(Privileges)
        
        PrepareMessageCharacterCreate = .ReadASCIIStringFixed(.Length)
    End With
End Function

''
' Prepares the "CharacterChange" message and returns it.
'
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterChange(ByVal body As Integer, ByVal Head As Integer, ByVal heading As eHeading, _
                                ByVal CharIndex As Integer, ByVal weapon As Integer, ByVal shield As Integer, _
                                ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer) As String
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CharacterChange" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterChange)
        
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(body)
        Call .WriteInteger(Head)
        Call .WriteByte(heading)
        Call .WriteInteger(weapon)
        Call .WriteInteger(shield)
        Call .WriteInteger(helmet)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)
        
        PrepareMessageCharacterChange = .ReadASCIIStringFixed(.Length)
    End With
End Function

''
' Prepares the "CharacterMove" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterMove(ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte) As String
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "CharacterMove" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterMove)
        Call .WriteInteger(CharIndex)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        
        PrepareMessageCharacterMove = .ReadASCIIStringFixed(.Length)
    End With
End Function

Public Function PrepareMessageForceCharMove(ByVal Direccion As eHeading) As String
'***************************************************
'Author: ZaMa
'Last Modification: 26/03/2009
'Prepares the "ForceCharMove" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ForceCharMove)
        Call .WriteByte(Direccion)
        
        PrepareMessageForceCharMove = .ReadASCIIStringFixed(.Length)
    End With
End Function

''
' Prepares the "UpdateTagAndStatus" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageUpdateTagAndStatus(ByVal UserIndex As Integer, ByVal NickColor As Byte, _
                                                ByRef Tag As String) As String
'***************************************************
'Author: Alejandro Salvo (Salvito)
'Last Modification: 04/07/07
'Last Modified By: Juan Mart�n Sotuyo Dodero (Maraxus)
'Prepares the "UpdateTagAndStatus" message and returns it
'15/01/2010: ZaMa - Now sends the nick color instead of the status.
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.UpdateTagAndStatus)
        
        Call .WriteInteger(UserList(UserIndex).Char.CharIndex)
        Call .WriteByte(NickColor)
        Call .WriteASCIIString(Tag)
        
        PrepareMessageUpdateTagAndStatus = .ReadASCIIStringFixed(.Length)
    End With
End Function

''
' Prepares the "ErrorMsg" message and returns it.
'
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageErrorMsg(ByVal message As String) As String
'***************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Prepares the "ErrorMsg" message and returns it
'***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ErrorMsg)
        Call .WriteASCIIString(message)
        
        PrepareMessageErrorMsg = .ReadASCIIStringFixed(.Length)
    End With
End Function

''
' Writes the "StopWorking" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.

Public Sub WriteStopWorking(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 21/02/2010
'
'***************************************************
On Error GoTo Errhandler
    
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.StopWorking)
        
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "CancelOfferItem" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Slot      The slot to cancel.

Public Sub WriteCancelOfferItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 05/03/2010
'
'***************************************************
On Error GoTo Errhandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.CancelOfferItem)
        Call .WriteByte(Slot)
    End With
    
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteEligeFaccion(ByVal UserIndex As Integer, ByVal Show As Boolean)
On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        
        Call .WriteByte(ServerPacketID.EligeFaccion)
        Call .WriteBoolean(Show)
        
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
    
End Sub

Public Sub WriteEligeRecompensa(ByVal UserIndex As Integer, ByVal Show As Boolean)
On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        
        Call .WriteByte(ServerPacketID.EligeRecompensa)
        Call .WriteBoolean(Show)
        
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
    
End Sub

Public Sub WriteSubeClase(ByVal UserIndex As Integer, ByVal Show As Boolean)
On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        
        Call .WriteByte(ServerPacketID.SubeClase)
        Call .WriteBoolean(Show)
    
    End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If

End Sub

Public Sub WriteShowClaseForm(ByVal UserIndex As Integer, ByVal Clase As Byte)
On Error GoTo Errhandler

        With UserList(UserIndex).outgoingData
        
            Call .WriteByte(ServerPacketID.ShowFormClase)
            Call .WriteByte(Clase)
            
        End With
Exit Sub

Errhandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteShowFaccionForm(ByVal UserIndex As Integer)
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowFaccionForm)
End Sub

Public Sub WriteShowRecompensaForm(ByVal UserIndex As Integer, ByVal Clase As Byte, ByVal Recom As Integer)
    With UserList(UserIndex).outgoingData
    
        .WriteByte ServerPacketID.ShowRecompensaForm
        .WriteByte Clase
        .WriteInteger Recom
        
    End With
End Sub
Private Sub HandleEligioClase(ByVal UserIndex As Integer)
    
    With UserList(UserIndex).incomingData
        
        Call .ReadByte
        
        Call RecibirSubClase(UserIndex, .ReadByte)
    
    End With
End Sub

Private Sub HandleRequestClaseForm(ByVal UserIndex As Integer)
    
    UserList(UserIndex).incomingData.ReadByte
    
    Call EnviarSubClase(UserIndex)
End Sub

Private Sub HandleRequestFaccionForm(ByVal UserIndex As Integer)
    UserList(UserIndex).incomingData.ReadByte
    
    Call EnviarFaccion(UserIndex)
End Sub

Private Sub HandleRequestRecompensaForm(ByVal UserIndex As Integer)
    UserList(UserIndex).incomingData.ReadByte
    
    Call EnviarRecompensa(UserIndex)
End Sub
Private Sub HandleWinTournament(ByVal UserIndex As Integer)
    
    On Error GoTo Errhandler
    
    With UserList(UserIndex)
        
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        Buffer.ReadByte
        
        Dim UI As Integer
        
        UI = NameIndex(Buffer.ReadASCIIString())
        
        If UI > 0 Then
            UserList(UI).Events.Torneos = UserList(UI).Events.Torneos + 1
            UserList(UI).Faccion.Torneos = UserList(UI).Faccion.Torneos + 1
        End If 'todo: deslogeado
        
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error

End Sub

Private Sub HandleLoseTournament(ByVal UserIndex As Integer)

    On Error GoTo Errhandler
    
    With UserList(UserIndex)
        
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        Buffer.ReadByte
        
        Dim UI As Integer
        
        UI = NameIndex(Buffer.ReadASCIIString())
        
        'que se supone que hace este comando?
        
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleWinQuest(ByVal UserIndex As Integer)
    
    On Error GoTo Errhandler
    
    With UserList(UserIndex)
        
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        Buffer.ReadByte
        
        Dim UI As Integer
        
        UI = NameIndex(Buffer.ReadASCIIString())
        
        If UI > 0 Then
            UserList(UI).Events.Torneos = UserList(UI).Events.Quests + 1
            'UserList(UI).Faccion.Torneos = UserList(UI).Faccion.Quests + 1 'todo
        End If 'todo: deslogeado
        
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error

End Sub

Private Sub HandleLoseQuest(ByVal UserIndex As Integer)

    On Error GoTo Errhandler
    
    With UserList(UserIndex)
        
        Dim Buffer As New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        Buffer.ReadByte
        
        Dim UI As Integer
        
        UI = NameIndex(Buffer.ReadASCIIString())
        
        'que se supone que hace este comando?
        
        Call .incomingData.CopyBuffer(Buffer)
    End With

Errhandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleEligioFaccion(ByVal UserIndex As Integer)
    Dim Faccion As eFaccion
    
    With UserList(UserIndex)
        .incomingData.ReadByte
        Faccion = .incomingData.ReadByte
        
        If Not PuedeFaccion(UserIndex) Then Exit Sub
        
        If .Faccion.BandoOriginal > 0 Then Exit Sub
        
        'todo
        If Faccion = eFaccion.Neutral Then
            If .Faccion.Bando <> eFaccion.Neutral Then
                'No pod�s declararte neutral perteneciendo ya a una facci�n. 32 51 223 N C
            Else
                '�Has decidido seguir siendo neutral! Pod�s jurar fidelidad cuando lo desees. 190 190 190 N
            End If
            Exit Sub
        End If
        
        If .Faccion.Matados(Faccion) > .Faccion.Matados(Enemigo(Faccion)) Then
            'mensaje 9
            Exit Sub
        End If
        
        'mensaje 10
        .Faccion.BandoOriginal = Faccion
        .Faccion.Bando = Faccion
        .Faccion.Ataco(Faccion) = 0
        If Not PuedeFaccion(UserIndex) Then Call WriteEligeFaccion(UserIndex, False)
        
        'updateuserchar
    End With
    
End Sub

Private Sub HandleEligioRecompensa(ByVal UserIndex As Integer)
    With UserList(UserIndex).incomingData
    
        .ReadByte
        
        Call RecibirRecompensa(UserIndex, .ReadByte)
    End With
End Sub

Private Sub HandleRequestGuildWindow(ByVal UserIndex As Integer)
    
    With UserList(UserIndex)
    
    
        .incomingData.ReadByte
        
        If .GuildID = 0 Then
            Call WriteSendGuildForm(UserIndex, eGFList)
        Else
            
            If .flags.IsLeader > 0 Then
                Call WriteSendGuildForm(UserIndex, eGFList)
            Else
                Call WriteSendGuildForm(UserIndex, eGFMembers)
            End If
        End If
        
    End With
    
End Sub

Private Sub WriteSendGuildForm(ByVal UserIndex As Integer, ByVal GForm As eGuildForms)
    
    With UserList(UserIndex).outgoingData
        
        .WriteByte ServerPacketID.SendGuildForm
        
        .WriteByte GForm
        
        Dim i As Long
        
        Select Case GForm
        
            Case eGuildForms.eGFList
            
            .WriteLong LastGuild
            
            For i = 1 To LastGuild
                If Guilds(i).Deleted = 0 Then
                    .WriteASCIIString Guilds(i).GuildName
                    .WriteByte Guilds(i).Faction
                End If
            Next
            
            Case eGuildForms.eGFLeaders
            
            
        End Select
    End With
End Sub

Private Sub HandleGuildFoundate(ByVal UserIndex As Integer)
    
    With UserList(UserIndex).incomingData
        
        .ReadByte

        Call WriteSendGuildFoundateWindow(UserIndex)

    End With
End Sub

Private Sub WriteSendGuildFoundateWindow(ByVal UserIndex As Integer)
    
    With UserList(UserIndex).outgoingData
        
        .WriteByte ServerPacketID.GuildFoundation
        
    End With
End Sub

Private Sub HandleGuildConfirmFoundation(ByVal UserIndex As Integer)
    
    With UserList(UserIndex).incomingData
    
        .ReadByte
        
        Call CreateGuild(.ReadASCIIString, UserIndex, .ReadByte, .ReadByte, .ReadByte)
    
    End With
End Sub

Private Sub HandleGuildRequest(ByVal UserIndex As Integer)
    
    With UserList(UserIndex).incomingData
        
        .ReadByte
        
        Dim Guild As String
        Guild = .ReadASCIIString
        
        Call SendRequest(UserIndex, GuildIndex(Guild))
    
    End With
End Sub

Public Sub HandleMoveItem(ByVal UserIndex As Integer)
'***************************************************
'Author: Ignacio Mariano Tirabasso (Budi)
'Last Modification: 01/01/2011
'
'***************************************************

With UserList(UserIndex)
 
    Dim originalSlot As Byte
    Dim newSlot As Byte
   
    Call .incomingData.ReadByte
   
    originalSlot = .incomingData.ReadByte
    newSlot = .incomingData.ReadByte
    
    ' Tipo (INVENTARIO, COMERCIO, BOVEDA, ETC).. (no se usa actualmente)
    Call .incomingData.ReadByte
    
    ' TODO: manejar el movimiento en los otros inventarios (comercio / boveda) si es necesario...
    Call InvUsuario.moveItem(UserIndex, originalSlot, newSlot)
   
End With
 
End Sub
