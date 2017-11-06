Attribute VB_Name = "modGuilds"
Option Explicit

Private Const FX_GUILD_CREATED As Byte = 44
Private Const FX_GUILD_ACEPTEDMEM As Byte = 43
Private Const FX_GUILD_DECLAREWAR As Byte = 45

Private Const MIN_LEVEL_CREATEGUILD As Byte = 30

'Private Enum eGuildRequirement
'        eNone = &H0
'        eLevel = &H1
'End Enum

Private Enum eGuildType
        ePublic = 0
        ePrivate = 1
End Enum

Private Type TYPE_GUILD
        Deleted                 As Byte
        
        GuildID                 As Long
        GuildName               As String
        
        FounderName             As String
        FoundationDate          As String
        
        LeaderName              As String
        
        RecruiterName(0 To 2)  As String
        
        Faction                 As Byte
        
        Entrance                As eGuildType
        
        'Requirement             As eGuildRequirement
        
        MinLevel                As Byte
        
        NumBlocked              As Integer
        
        BlockedUsers            As Collection
        
        Members                 As Collection
        
        Requests                As Collection
        
        OnlineMembers           As Collection
End Type

Public Guilds() As TYPE_GUILD

Public LastGuild As Long

'CSEH: ErrLog
Public Sub LoadGuilds()
    '<EhHeader>
    On Error GoTo LoadGuilds_Err
    '</EhHeader>
    
        Dim handle As Integer
        Dim buf As New CsBuffer
        Dim data() As Byte
        
100     handle = FreeFile()
                
105     Open GuildPath & "Guilds.fnx" For Binary Access Read As handle
            
            If LOF(handle) = 0 Then GoTo NoCargar
110         Get handle, , LastGuild
        
115         ReDim data(0 To LOF(handle) - 1)
            'ReDim Guilds(1 To LastGuild) As TYPE_GUILD
        
120         Get handle, , data
            'Get handle, , Guilds
125     Close handle
    
    
130     Call buf.Wrap(data)
    
135     ReDim Guilds(1 To LastGuild)
    
        Dim i As Long
        Dim j As Long
        Dim Count As Long
    
140     For i = 1 To LastGuild
145         With Guilds(i)
            
150             .Deleted = buf.ReadByte
            
155             .GuildID = buf.ReadLong
160             .GuildName = buf.ReadString
            
165             .FounderName = buf.ReadString
170             .FoundationDate = buf.ReadString
175             .LeaderName = buf.ReadString
            
180             For j = 0 To 2
185                 .RecruiterName(j) = buf.ReadString
                Next
            
190             .Faction = buf.ReadByte
            
195             .Entrance = buf.ReadByte
            
200             .MinLevel = buf.ReadByte
            
205             Count = buf.ReadLong
            
210             Set .BlockedUsers = New Collection
            
215             If Count > 0 Then
                
220                 For j = 1 To Count
                
225                     Call .BlockedUsers.Add(buf.ReadString)
                    
230                 Next j
                
                End If
            
235             Count = buf.ReadLong
                
                Set .Members = New Collection
                
240             For j = 1 To Count
245                 Call .Members.Add(buf.ReadString)
                Next
            
250             Count = buf.ReadLong
            
255             If Count > 0 Then
                
260                 Set .Requests = New Collection
                
265                 For j = 1 To Count
                        Dim r As New CsRequest
                    
270                     r.ToGuild = buf.ReadLong
275                     r.RequesterName = buf.ReadString
280                     r.RequestDate = buf.ReadString
                    
285                     Call .Requests.Add(r)
                    Next
                
                End If
            
290             Set .OnlineMembers = New Collection
            
            End With
        Next
    
    '<EhFooter>
    Exit Sub

NoCargar:
    Close handle
    
    Exit Sub
    
LoadGuilds_Err:
        Call LogError("Error en LoadGuilds: " & Erl & " - " & Err.description)
    '</EhFooter>
End Sub

'CSEH: ErrLog
Public Sub DumpGuilds(Optional ByVal ShutDown As Boolean = False)
    
    'There's nothing to write
    If LastGuild = 0 Then Exit Sub
    
    Dim handle As Integer
    Dim i As Long
    Dim j As Long
    Dim buf As New CsBuffer
        
    
    Call buf.WriteLong(LastGuild)
    
    For i = 1 To LastGuild
        Dim c As TYPE_GUILD
        c = Guilds(i)
        
    'TODO: los guilds marcados como -deleted- no se deberian guardar, y al borrar debería hacerse una reorganización de IDS
        Call buf.WriteByte(c.Deleted)
        Call buf.WriteLong(c.GuildID)
        Call buf.WriteString(c.GuildName)
        
        Call buf.WriteString(c.FounderName)
        Call buf.WriteString(c.FoundationDate)
        Call buf.WriteString(c.LeaderName)
        
        Call buf.WriteString(c.RecruiterName(0))
        Call buf.WriteString(c.RecruiterName(1))
        Call buf.WriteString(c.RecruiterName(2))
        
        Call buf.WriteByte(c.Faction)
        Call buf.WriteByte(c.Entrance)
        Call buf.WriteByte(c.MinLevel)
        
        Call buf.WriteLong(c.BlockedUsers.Count)
        
        For j = 1 To c.BlockedUsers.Count
            
            Call buf.WriteString(c.BlockedUsers.Item(j))
            
        Next
        
        Call buf.WriteLong(c.Members.Count)
        
        For j = 1 To c.Members.Count
            Call buf.WriteString(c.Members.Item(j))
        Next
        
        If c.Requests Is Nothing Then
        
            Call buf.WriteLong(0)
        
        Else
        
            Call buf.WriteLong(c.Requests.Count)
            
            For j = 1 To c.Requests.Count
                Dim r As New CsRequest
                
                Set r = c.Requests.Item(i)
                
                Call buf.WriteLong(r.ToGuild)
                Call buf.WriteString(r.RequesterName)
                Call buf.WriteString(r.RequestDate)
                
            Next
            
        End If
        
    Next
    
    handle = FreeFile()
    
    Open GuildPath & "Guilds.fnx" For Binary Access Write As handle
        Put handle, , buf.Buffer
    Close handle
    
    Set buf = Nothing
    
    If ShutDown Then
        Erase Guilds
        'Set Requests = Nothing
    End If
    
End Sub

'CSEH: ErrLog
Public Sub SendToGuild(ByVal GuildID As Long, ByVal sData As String)
    '<EhHeader>
    On Error GoTo SendToGuild_Err
    '</EhHeader>
        Dim i As Long
    
100     With Guilds(GuildID)
    
105         For i = 1 To .OnlineMembers.Count
            
110             If UserList(CInt(.OnlineMembers.Item(i))).ConnIDValida Then
                
115                 Call EnviarDatosASlot(CInt(.OnlineMembers.Item(i)), sData)
                
                End If
            
            Next
        
        End With
    '<EhFooter>
    Exit Sub

SendToGuild_Err:
        Call LogError("Error en SendToGuild: " & Erl & " - " & Err.description)
    '</EhFooter>
End Sub

'CSEH: ErrLog
Public Sub CreateGuild(ByVal GuildName As String, ByVal UserIndex As Integer, _
                        ByVal Faction As Byte, ByVal Entrance As Byte, ByVal MinLevel As Byte)
    '<EhHeader>
    On Error GoTo CreateGuild_Err
    '</EhHeader>
        Dim errstr As String
    
100     GuildName = Trim$(GuildName)
    
105     If CanFoundate(UserIndex, GuildName, Faction, MinLevel, errstr) = False Then
110         Call WriteErrorMsg(UserIndex, errstr)
            Exit Sub
        End If
    
115     LastGuild = LastGuild + 1
        
        ReDim Preserve Guilds(1 To LastGuild) As TYPE_GUILD
        
120     With Guilds(LastGuild)
    
125         .GuildName = GuildName
130         .GuildID = LastGuild
            
            .FounderName = UserList(UserIndex).Name
135         .LeaderName = .FounderName
        
140         .Entrance = Entrance
        
145         .MinLevel = MinLevel
        
150         .Faction = Faction
            
            Set .Members = New Collection
            Set .BlockedUsers = New Collection
            Set .OnlineMembers = New Collection
            Set .Requests = New Collection
            
            UserList(UserIndex).GuildID = LastGuild
            UserList(UserIndex).flags.WaitingApprovement = 0
            
155         Call AddMember(LastGuild, UserList(UserIndex).Name)
160         Call ConnectMember(LastGuild, UserIndex)
        
165         .FoundationDate = Format$(Now, "yy/m/d")
        End With
    
170     Call SendData(SendTarget.ToAll, 0, PrepareMessageMultiMessage(eMessages.GuildCreated, , , , UserList(UserIndex).Name & "," & Guilds(LastGuild).GuildName))
    '<EhFooter>
    Exit Sub

CreateGuild_Err:
        Call LogError("Error en CreateGuild: " & Erl & " - " & Err.description)
    '</EhFooter>
End Sub

'CSEH: ErrLog
Public Sub SendRequest(ByVal UserIndex As Integer, ByVal ToGuildID As Integer)
    
    With UserList(UserIndex)
    
        If .GuildID <> 0 Then
            'Call WriteConsoleMsg(UserIndex, "Ya te encuentras en un clan, debes salir.", FontTypeNames.FONTTYPE_WARNING)
            Call WriteMultiMessage(UserIndex, eMessages.AlreadyInGuild)
            Exit Sub
        End If
        
        If .flags.WaitingApprovement > 0 Then
            'Call WriteConsoleMsg(UserIndex, "Se eliminará la petición al clan anterior.", FontTypeNames.FONTTYPE_INFOBOLD)
            Call WriteMultiMessage(UserIndex, eMessages.PreviousRequest)
            Call DeleteRequest(.flags.WaitingApprovement, .Name)
        End If
        
        If Guilds(ToGuildID).Faction = eFaccion.Caos Then
            
            If .Faccion.Bando = eFaccion.Real Then
                'Call WriteConsoleMsg(UserIndex, "No puedes enviar solicitud a un clan de alineación enemiga.", FontTypeNames.FONTTYPE_WARNING)
                Call WriteMultiMessage(UserIndex, eMessages.EnemyGuild)
                Exit Sub
            End If
            
        ElseIf Guilds(ToGuildID).Faction = eFaccion.Real Then
            
            If .Faccion.Bando = eFaccion.Caos Then
                'Call WriteConsoleMsg(UserIndex, "No puedes enviar solicitud a un clan de alineación enemiga.", FontTypeNames.FONTTYPE_WARNING)
                Call WriteMultiMessage(UserIndex, eMessages.EnemyGuild)
                Exit Sub
            End If
            
        End If
        
        If ToGuildID > LastGuild Then
            Call WriteConsoleMsg(UserIndex, "Se ha encontrado un error, refresca la lista de clanes, por favor.", FontTypeNames.FONTTYPE_WARNING)
            Exit Sub
        End If
        
        Dim i As Long
        
        For i = 1 To Guilds(ToGuildID).BlockedUsers.Count
            If StrComp(UCase$(.Name), UCase$(Guilds(ToGuildID).BlockedUsers.Item(i))) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Tienes prohibido el ingreso a este clan.", FontTypeNames.FONTTYPE_WARNING)
            End If
        Next
        
        'todo capacity
            
        If Guilds(ToGuildID).Entrance = eGuildType.ePublic Then
            
            UserList(UserIndex).flags.WaitingApprovement = 0
            UserList(UserIndex).GuildID = ToGuildID
            Call AddMember(ToGuildID, UserList(UserIndex).Name)
            Call ConnectMember(ToGuildID, UserIndex)
            
        Else
        
            Dim r As New CsRequest
            
            r.RequesterName = UserList(UserIndex).Name
            r.RequestDate = Format$(Now, "yy/m/d")
            r.ToGuild = ToGuildID
            
            Call Guilds(ToGuildID).Requests.Add(r)
            
            Set r = Nothing
        End If
        
    End With
    
End Sub

'CSEH: ErrLog
Public Sub AcceptRequest(ByVal UserIndex As Integer, ByVal RequesterName As Integer)

With UserList(UserIndex)
    
    If .flags.IsLeader = 0 Then 'This should never happen
        Call WriteConsoleMsg(UserIndex, "No tienes permisos para aceptar una solicitud.", FontTypeNames.FONTTYPE_WARNING)
        Exit Sub
    End If
    
    Dim uID As Integer
    
    uID = RequesterToID(.GuildID, RequesterName)
    
    If uID = 0 Then
        Call WriteConsoleMsg(UserIndex, "No existe la petición.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    Dim r As New CsRequest
    
    Set r = Guilds(.GuildID).Requests.Item(uID)
    With r
    
        If .ToGuild = UserList(UserIndex).GuildID Then
            
            Dim i As Integer
            
            i = NameIndex(.RequesterName)
            
            If i <> 0 Then
                
                If UserList(i).GuildID <> 0 Then
                    Call WriteConsoleMsg(UserIndex, "El usuario ya se encuentra en otro clan.", FontTypeNames.FONTTYPE_INFO)
                    Call Guilds(UserList(UserIndex).GuildID).Requests.Remove(uID)
                    Exit Sub
                End If
                
                UserList(i).flags.WaitingApprovement = 0
                UserList(i).GuildID = UserList(UserIndex).GuildID
                Call AddMember(.ToGuild, .RequesterName)
                Call ConnectMember(.ToGuild, i)
                
            Else
                
                If FileExist(CharPath & .RequesterName & ".chr", vbNormal) Then
                    
                    If val(GetVar(CharPath & .RequesterName & ".chr", "Guild", "GuildID")) <> 0 Then
                        Call WriteConsoleMsg(UserIndex, "El usuario ya se encuentra en otro clan.", FontTypeNames.FONTTYPE_INFO)
                        Call Guilds(UserList(UserIndex).GuildID).Requests.Remove(uID)
                        Exit Sub
                    End If
                    
                    Call WriteVar(CharPath & .RequesterName & ".chr", "Guild", "RequestedTo", 0)
                    Call WriteVar(CharPath & .RequesterName & ".chr", "Guild", "GuildID", UserList(UserIndex).GuildID)
                    Call AddMember(.ToGuild, .RequesterName)
                    
                Else
                    Call LogError("*******PELIGRO*******: Se aprobó un ingreso de un char inexistente.")
                End If
                
            End If
            
            Call Guilds(UserList(UserIndex).GuildID).Requests.Remove(uID)
            'todo: refresh acceptance window
        End If
        
    End With
    
    Set r = Nothing
End With

End Sub

'CSEH: ErrLog
Public Sub PruneRequests(ByVal UserIndex As Integer, ByVal GuildID As Long)
    '<EhHeader>
    On Error GoTo PruneRequests_Err
    '</EhHeader>

100     If GuildID <= LastGuild Then
        
105         If UserList(UserIndex).flags.IsLeader = 1 And UserList(UserIndex).GuildID = GuildID Then
            
110             Set Guilds(GuildID).Requests = New Collection
            
                'todo: refresh
            End If
        
            
        End If
    
    '<EhFooter>
    Exit Sub

PruneRequests_Err:
        Call LogError("Error en PruneRequests: " & Erl & " - " & Err.description)
    '</EhFooter>
End Sub

'CSEH: ErrLog
Private Sub DeleteRequest(ByVal GuildID As Long, ByVal UserName As String)
    '<EhHeader>
    On Error GoTo DeleteRequest_Err
    '</EhHeader>
    
100     If GuildID <= LastGuild Then
        
            Dim i As Long
            Dim r As New CsRequest
        
105         With Guilds(GuildID)
        
110             For i = 1 To .Requests.Count
                
115                 Set r = .Requests.Item(i)
                
120                 If StrComp(UCase$(UserName), UCase$(r.RequesterName)) = 0 Then
125                     Call .Requests.Remove(i)
                        Exit Sub
                    End If
                
                Next
        
            End With
        
        Else
130         Call LogError("DeleteRequest: invalid guild ID.")
        End If
    '<EhFooter>
    Exit Sub

DeleteRequest_Err:
        Call LogError("Error en DeleteRequest: " & Erl & " - " & Err.description)
    '</EhFooter>
End Sub

'CSEH: ErrLog
Private Function MemberToID(ByVal GuildID As Long, ByVal UserName As String)
    '<EhHeader>
    On Error GoTo MemberToID_Err
    '</EhHeader>
        Dim i As Long
    
    
100     With Guilds(GuildID)
    
105         For i = 1 To .Members.Count
            
110             If StrComp(UCase$(UserName), UCase$(.Members.Item(i))) = 0 Then
115                 MemberToID = i
                    Exit Function
                End If
            
            Next
    
        End With
    
120     MemberToID = 0
    
    '<EhFooter>
    Exit Function

MemberToID_Err:
        Call LogError("Error en MemberToID: " & Erl & " - " & Err.description)
    '</EhFooter>
End Function

Private Function RequesterToID(ByVal GuildID As Long, ByVal RequesterName As String)
    
    Dim i As Long
    Dim r As CsRequest
    
    For i = 1 To Guilds(GuildID).Requests.Count
            
        Set r = Guilds(GuildID).Requests.Item(i)
        
        If StrComp(UCase$(RequesterName), UCase$(r.RequesterName)) = 0 Then
            RequesterToID = i
            Exit Function
        End If
        
    Next
    
    RequesterToID = 0
    
End Function

'CSEH: ErrLog
Private Sub AddMember(ByVal GuildID As Long, ByVal UserName As String)
    '<EhHeader>
    On Error GoTo AddMember_Err
    '</EhHeader>
    
100     With Guilds(GuildID)
        
105         Call .Members.Add(UserName)
            
            Call WriteVar(CharPath & UserName & ".chr", "Guild", "GuildID", GuildID)
            
            Call SendData(SendTarget.ToGuildMembers, GuildID, PrepareMessageMultiMessage(eMessages.GuildAccepted, , , , UserName))
        End With
    
    '<EhFooter>
    Exit Sub

AddMember_Err:
        Call LogError("Error en AddMember: " & Erl & " - " & Err.description)
    '</EhFooter>
End Sub

'CSEH: ErrLog
Public Sub RemoveMember(ByVal GuildID As Long, ByVal UserName As String)
    '<EhHeader>
    On Error GoTo RemoveMember_Err
    '</EhHeader>
    
100     If GuildID <= LastGuild Then

105         With Guilds(GuildID)
                Dim uID As Integer
            
110             uID = NameIndex(UserName)
            
115             If uID <> 0 Then
            
120                 If UserList(uID).flags.IsLeader = 1 Then 'Check if the leaving is the leader
125                     Call RemovingLeader(GuildID, UserName)
                    
130                     UserList(uID).flags.IsLeader = 0
                    End If
                
135                 UserList(uID).GuildID = 0
                
                Else
                    Dim leader As Byte
140                 leader = val(GetVar(CharPath & UserName & ".chr", "FLAGS", "IsLeader"))
                
145                 If leader = 1 Then
150                     Call RemovingLeader(GuildID, UserName)
                    
155                     Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "IsLeader", CStr(0))
                    End If
                
160                 Call WriteVar(CharPath & UserName & ".chr", "GUILD", "GUILDID", CStr(0))
                End If
        
            End With
        
        End If
    '<EhFooter>
    Exit Sub

RemoveMember_Err:
        Call LogError("Error en RemoveMember: " & Erl & " - " & Err.description)
    '</EhFooter>
End Sub

'CSEH: ErrLog
Private Sub RemovingLeader(ByVal GuildID As Long, ByVal UserName As String)
    '<EhHeader>
    On Error GoTo RemovingLeader_Err
    '</EhHeader>
        Dim substitute As String
        Dim subsID As Integer
        Dim i As Long
    
100     substitute = GetNextRecruiter(GuildID, True)
    
105     With Guilds(GuildID)
    
110         If LenB(substitute) <> 0 Then 'Exist a Recruiter that could take the leaders position
            
115             subsID = NameIndex(substitute)
            
120             .LeaderName = substitute
            
125             If subsID > 0 Then ' Its logged
            
130                 UserList(subsID).flags.IsLeader = 1
                
                Else ' Its Offline
            
135                 Call WriteVar(CharPath & substitute & ".chr", "FLAGS", "IsLeader", 1)
                
                End If
            
140             Call SendData(SendTarget.ToGuildMembers, GuildID, PrepareMessageConsoleMsg(.LeaderName & " es ahora el líder del clan.", FontTypeNames.FONTTYPE_SERVER))
        
            Else 'There's no Recruiters, search for normal members
            
145             If .Members.Count > 1 Then
                
150                 For i = 1 To .Members.Count
155                     If StrComp(UCase$(UserName), UCase$(.Members.Item(i))) <> 0 Then
                        
160                         .LeaderName = .Members.Item(i)
                        
165                         subsID = NameIndex(.LeaderName)
                        
170                         If subsID > 0 Then
                        
175                             UserList(subsID).flags.IsLeader = 1
                            
                            Else
                        
180                             Call WriteVar(CharPath & .LeaderName & ".chr", "FLAGS", "IsLeader", 1)
                        
                            End If
                        
185                         Call SendData(SendTarget.ToGuildMembers, GuildID, PrepareMessageConsoleMsg(.LeaderName & " es ahora el líder del clan.", FontTypeNames.FONTTYPE_SERVER))
                        End If
                    Next
                
                Else 'No Recruiters, and not even more members
190                 .Deleted = 1
                End If
            
            End If
        End With
    '<EhFooter>
    Exit Sub

RemovingLeader_Err:
        Call LogError("Error en RemovingLeader: " & Erl & " - " & Err.description)
    '</EhFooter>
End Sub

'CSEH: ErrLog
Public Sub ConnectMember(ByVal GuildID As Long, ByVal UserIndex As Integer)
    '<EhHeader>
    On Error GoTo ConnectMember_Err
    '</EhHeader>
    
100     If GuildID > 0 Then
    
105         Call Guilds(GuildID).OnlineMembers.Add(UserIndex)
        
110         Call SendData(SendTarget.ToGuildMembers, GuildID, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " se ha conectado.", FontTypeNames.FONTTYPE_GUILDLOGIN))
        End If
    
    '<EhFooter>
    Exit Sub

ConnectMember_Err:
        Call LogError("Error en ConnectMember: " & Erl & " - " & Err.description)
    '</EhFooter>
End Sub

'CSEH: ErrLog
Public Sub DisconnectMember(ByVal GuildID As Long, ByVal UserIndex As Integer)
    '<EhHeader>
    On Error GoTo DisconnectMember_Err
    '</EhHeader>
        
100     If GuildID > 0 Then
105         Call Guilds(GuildID).OnlineMembers.Remove(UserIndex)
        
110         If UserList(UserIndex).flags.IsLeader > 0 Then
115             Call SendData(SendTarget.ToGuildMembers, GuildID, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " se ha desconectado.", FontTypeNames.FONTTYPE_GUILDLOGIN))
            End If
        End If
    
    '<EhFooter>
    Exit Sub

DisconnectMember_Err:
        Call LogError("Error en DisconnectMember: " & Erl & " - " & Err.description)
    '</EhFooter>
End Sub

'CSEH: ErrLog
Private Function GetNextRecruiter(ByVal GuildID As Long, Optional ByVal RemoveRecruiter As Boolean = False) As String
    '<EhHeader>
    On Error GoTo GetNextRecruiter_Err
    '</EhHeader>
    
        Dim i As Long
    
100     With Guilds(GuildID)
    
105         For i = 0 To 2
        
110             If LenB(.RecruiterName(i)) <> 0 Then
115                 GetNextRecruiter = .RecruiterName(i)
                
120                 If RemoveRecruiter Then
                    
125                     If i = 0 Then 'If the promoted its in the first index
                        
130                         .RecruiterName(0) = .RecruiterName(1)
135                         .RecruiterName(1) = .RecruiterName(2)
140                         .RecruiterName(2) = vbNullString
                        
145                     ElseIf i = 1 Then 'If its in the index 1, its 'cause theres nothing in the first index
                        
150                         .RecruiterName(0) = .RecruiterName(2)
155                         .RecruiterName(1) = vbNullString
160                         .RecruiterName(2) = vbNullString
                    
                        Else 'The Recruiter its in the last index (there is only one Recruiter).
                        
165                         .RecruiterName(0) = vbNullString
170                         .RecruiterName(1) = vbNullString
175                         .RecruiterName(2) = vbNullString
                        
                        End If
                    End If
                
                    Exit Function
                End If
            
            Next
        
        End With
    
180     GetNextRecruiter = vbNullString
    
    '<EhFooter>
    Exit Function

GetNextRecruiter_Err:
        Call LogError("Error en GetNextRecruiter: " & Erl & " - " & Err.description)
    '</EhFooter>
End Function

'CSEH: ErrLog
Private Function CanFoundate(ByVal UserIndex As Integer, ByVal GuildName As String, ByVal Faction As Byte, ByVal MinLevel As Byte, ByRef ErrorMessage As String)
    '<EhHeader>
    On Error GoTo CanFoundate_Err
    '</EhHeader>
    
100     With UserList(UserIndex)
105         If .flags.Muerto = 1 Then
110             CanFoundate = False
115             ErrorMessage = "Los muertos no pueden fundar clanes."
                Exit Function
            End If
        
120         If .Stats.ELV < MIN_LEVEL_CREATEGUILD Then
125             CanFoundate = False
130             ErrorMessage = "Debes ser al menos nivel " & MIN_LEVEL_CREATEGUILD & " para poder fundar."
                Exit Function
        
135         ElseIf .Stats.ELV < MinLevel Then
140             CanFoundate = False
145             ErrorMessage = "El nivel mínimo que has elegido es mayor a tu nivel."
                Exit Function
            
            End If
        
            'todo debería tener carisma?
            'todo deberia tener un item?
        
150         If .GuildID <> 0 Then
155             CanFoundate = False
160             ErrorMessage = "No puedes fundar un clan si ya perteneces a otro."
                Exit Function
            End If
        
165         If .Faccion.Bando = eFaccion.Caos Then
            
170             If Faction = eFaccion.Real Then
175                 CanFoundate = False
180                 ErrorMessage = "No puedes elegir una alineación opositora."
                    Exit Function
                End If
            
185         ElseIf .Faccion.Bando = eFaccion.Real Then
            
190             If Faction = eFaccion.Caos Then
195                 CanFoundate = False
200                 ErrorMessage = "No puedes elegir una alineación opositora."
                    Exit Function
                End If
            
            End If
        
            'Delete a previous request
205         If .flags.WaitingApprovement > 0 Then
210             Call DeleteRequest(.flags.WaitingApprovement, .Name)
            End If
        
            Dim i As Long
        
215         For i = 1 To LastGuild
220             If StrComp(UCase$(Guilds(i).GuildName), UCase$(GuildName)) = 0 Then
225                 CanFoundate = False
230                 ErrorMessage = "Ya existe un clan con el mismo nombre."
                    Exit Function
                End If
            Next
        
235         CanFoundate = True
        End With
    '<EhFooter>
    Exit Function

CanFoundate_Err:
        Call LogError("Error en CanFoundate: " & Erl & " - " & Err.description)
    '</EhFooter>
End Function

Public Function GuildIndex(ByVal GuildName As String) As Long
    
    Dim i As Long
    
    For i = 1 To LastGuild
        
        If StrComp(UCase$(GuildName), UCase$(Guilds(i).GuildName)) = 0 Then
            GuildIndex = i
            Exit Function
        End If
        
    Next
    
End Function

Public Function SameGuild(ByVal UserA As Integer, ByVal UserB As Integer) As Boolean
    SameGuild = (UserList(UserA).GuildID = UserList(UserB).GuildID)
End Function
