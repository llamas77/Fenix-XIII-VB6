Attribute VB_Name = "ModFacciones"
Option Explicit

Public MiembrosCaos                         As Collection
Public MiembrosAlianza                      As Collection

Private Const NIVEL_MINIMO_INGRESAR         As Byte = 25

Public Const REQUIERE_MATADOS_PRIMERA       As Integer = 100
Public Const REQUIERE_MATADOS_SEGUNDA       As Integer = 500
Public Const REQUIERE_MATADOS_TERCERA       As Integer = 1000
Public Const REQUIERE_MATADOS_CUARTA        As Integer = 1500

Private Const REQUIERE_TORNEOS_SEGUNDA      As Integer = 1
Private Const REQUIERE_TORNEOS_TERCERA      As Integer = 5
Private Const REQUIERE_TORNEOS_CUARTA       As Integer = 10

Public Mensajes(1 To 2, 1 To 23) As String
Public Armaduras(1 To 2, 1 To 3, 1 To 4, 1 To 2) As Integer

Public Sub EnviarFaccion(ByVal UserIndex As Integer)
    
    If UserList(UserIndex).flags.Muerto = 0 Then 'es necesario que este vivo para hacerlo??
        Call WriteShowFaccionForm(UserIndex)
    End If
End Sub

Public Function Enemigo(ByVal Bando As eFaccion) As Byte

Select Case Bando
    Case eFaccion.Neutral
        Enemigo = 3
    Case eFaccion.Real
        Enemigo = Caos
    Case eFaccion.Caos
        Enemigo = Real
End Select

End Function

Public Sub InitFacciones()

    Set MiembrosCaos = New Collection
    Set MiembrosAlianza = New Collection
    
End Sub

Public Sub AgregarMiembroFaccion(ByVal UserIndex As Integer)
    
    Dim Faccion As eFaccion
    
    Faccion = UserList(UserIndex).Faccion.Bando
    
    Select Case Faccion
    
        Case eFaccion.Caos
            MiembrosCaos.Add UserIndex
            Exit Sub
            
        Case eFaccion.Real
            MiembrosAlianza.Add UserIndex
            Exit Sub
    End Select
End Sub

Public Sub QuitarMiembroFaccion(ByVal UserIndex As Integer)
    
    Dim Faccion As eFaccion
    
    Faccion = UserList(UserIndex).Faccion.Bando
    
    Dim i As Long
    
    Select Case Faccion
    
        Case eFaccion.Caos
            For i = 1 To MiembrosCaos.count
                If MiembrosCaos.Item(i) = UserIndex Then
                    MiembrosCaos.Remove i
                    Exit For
                End If
            Next
            Exit Sub
        Case eFaccion.Real
            For i = 1 To MiembrosAlianza.count
                If MiembrosAlianza.Item(i) = UserIndex Then
                    MiembrosAlianza.Remove i
                    Exit For
                End If
            Next
            Exit Sub
            
    End Select
    
End Sub

Public Sub EnviarDatosJerarquia(ByVal Faccion As eFaccion, ByVal data As String)
Dim i As Long
    
    Select Case Faccion
    
        Case eFaccion.Caos
            
            For i = 1 To MiembrosCaos.count
                If EsCaos(MiembrosCaos.Item(i)) Then
                    Call EnviarDatosASlot(MiembrosCaos.Item(i), data)
                End If
            Next
            Exit Sub
        
        Case eFaccion.Real
            For i = 1 To MiembrosAlianza.count
                If EsArmada(MiembrosAlianza.items(i)) Then
                    Call EnviarDatosASlot(MiembrosAlianza.Item(i), data)
                End If
            Next
            Exit Sub

    End Select
End Sub

Public Sub EnviarDatosFaccion(ByVal Faccion As eFaccion, ByVal data As String)
Dim i As Long
    
    Select Case Faccion
    
        Case eFaccion.Caos
            
            For i = 1 To MiembrosCaos.count
                Call EnviarDatosASlot(MiembrosCaos.Item(i), data)
            Next
            Exit Sub
        
        Case eFaccion.Real
            For i = 1 To MiembrosAlianza.count
                Call EnviarDatosASlot(MiembrosAlianza.Item(i), data)
            Next
            Exit Sub

    End Select
End Sub

Public Function ClaseTrabajadora(ByVal Clase As eClass) As Boolean

ClaseTrabajadora = (Clase > eClass.Ciudadano And Clase < eClass.Luchador)

End Function

Public Sub Recompensado(ByVal UserIndex As Integer)
Dim Fuerzas As Byte
Dim MiObj As Obj

With UserList(UserIndex)
    Fuerzas = .Faccion.Bando
    
    
    If .Faccion.Jerarquia = 0 Then
        Call WriteMultiMessage(UserIndex, eMessages.WrongFaction)
        Exit Sub
    End If
    
    If .Faccion.Jerarquia = 1 Then
        If .Faccion.Matados(Enemigo(Fuerzas)) < REQUIERE_MATADOS_SEGUNDA Then
            Call WriteMultiMessage(UserIndex, eMessages.NeedToKill, REQUIERE_MATADOS_SEGUNDA, .Faccion.Matados(Enemigo(Fuerzas)))
            Exit Sub
        End If
        
        If .Faccion.Torneos < REQUIERE_TORNEOS_SEGUNDA Then
            Call WriteMultiMessage(UserIndex, eMessages.NeedTournaments, REQUIERE_TORNEOS_SEGUNDA, .Faccion.Torneos)
            Exit Sub
        End If
        
        'todo quest
        'If .Faccion.Quests < 1 Then
        '    Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 14) & 1)
        '    Exit Sub
        'End If
        
        .Faccion.Jerarquia = 2
        'Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 15) & Titulo(UserIndex))
        Call WriteMultiMessage(UserIndex, eMessages.HierarchyUpgradre, Titulo(UserIndex))
    ElseIf .Faccion.Jerarquia = 2 Then
        If .Faccion.Matados(Enemigo(Fuerzas)) < REQUIERE_MATADOS_TERCERA Then
            Call WriteMultiMessage(UserIndex, eMessages.NeedToKill, REQUIERE_MATADOS_TERCERA, .Faccion.Matados(Enemigo(Fuerzas)))
            Exit Sub
        End If

        If .Faccion.Torneos < REQUIERE_TORNEOS_TERCERA Then
            Call WriteMultiMessage(UserIndex, eMessages.NeedTournaments, REQUIERE_TORNEOS_TERCERA, .Faccion.Torneos)
            Exit Sub
        End If
        
        'If .Faccion.Quests < 2 Then
        '    Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 14) & 2)
        '    Exit Sub
        'End If
        
        .Faccion.Jerarquia = 3
        Call WriteMultiMessage(UserIndex, eMessages.HierarchyUpgradre, Titulo(UserIndex))
    ElseIf .Faccion.Jerarquia = 3 Then
        If .Faccion.Matados(Enemigo(Fuerzas)) < REQUIERE_MATADOS_CUARTA Then
            Call WriteMultiMessage(UserIndex, eMessages.NeedToKill, REQUIERE_MATADOS_CUARTA, .Faccion.Matados(Enemigo(Fuerzas)))
            Exit Sub
        End If
        
        If .Faccion.Torneos < REQUIERE_TORNEOS_CUARTA Then
            Call WriteMultiMessage(UserIndex, eMessages.NeedTournaments, REQUIERE_TORNEOS_CUARTA, .Faccion.Torneos)
            Exit Sub
        End If
        
        'If .Faccion.Quests < 5 Then
        '    Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 14) & 5)
        '    Exit Sub
        'End If
        
        .Faccion.Jerarquia = 4
        Call WriteMultiMessage(UserIndex, eMessages.HierarchyUpgradre, Titulo(UserIndex))
    End If
    
    
    If .Faccion.Jerarquia < 4 Then
        MiObj.Amount = 1
        MiObj.OBJIndex = Armaduras(Fuerzas, .Faccion.Jerarquia, TipoClase(UserIndex), TipoRaza(UserIndex))
        If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(.Pos, MiObj)
        'todo, nunca tirar item al piso, es estúpido
    Else
        'Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 22) & str(Npclist(.flags.TargetNPC).Char.CharIndex))
        Call WriteMultiMessage(UserIndex, eMessages.LastHierarchy, Npclist(.flags.TargetNPC).Char.CharIndex)
    End If
    
End With

End Sub
Public Sub Expulsar(ByVal UserIndex As Integer)

Call WriteMultiMessage(UserIndex, eMessages.HierarchyExpelled)
UserList(UserIndex).Faccion.Bando = eFaccion.Neutral
UserList(UserIndex).Faccion.Jerarquia = 0
Call RefreshCharStatus(UserIndex)

End Sub
Public Sub Enlistar(ByVal UserIndex As Integer, ByVal Fuerzas As Byte)
Dim MiObj As Obj

With UserList(UserIndex)
    If .Faccion.Bando = eFaccion.Neutral Then
        'Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 1) & str(Npclist(.flags.TargetNPC).Char.CharIndex))
        Call WriteMultiMessage(UserIndex, eMessages.Neutral, Npclist(.flags.TargetNPC).Char.CharIndex)
        Exit Sub
    End If
    
    If .Faccion.Bando = Enemigo(Fuerzas) Then
        'Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 2) & str(Npclist(.flags.TargetNPC).Char.CharIndex))
        Call WriteMultiMessage(UserIndex, eMessages.OppositeSide, Npclist(.flags.TargetNPC).Char.CharIndex)
        Exit Sub
    End If
    
    'todo guilds
    'Dim oGuild As cGuild
    
    'Set oGuild = FetchGuild(.GuildInfo.GuildName)
    
    'If Len(.GuildInfo.GuildName) > 0 Then
    '    If oGuild.Bando <> Fuerzas Then
    '        Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 3) & str(Npclist(.flags.TargetNPC).Char.CharIndex))
    '        Exit Sub
    '    End If
    'End If
    
    If .Faccion.Jerarquia Then
        'Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 4) & str(Npclist(.flags.TargetNPC).Char.CharIndex))
        Call WriteMultiMessage(UserIndex, eMessages.AlreadyBelong, Npclist(.flags.TargetNPC).Char.CharIndex)
        Exit Sub
    End If
    
    If .Faccion.Matados(Enemigo(Fuerzas)) < REQUIERE_MATADOS_PRIMERA Then
        'Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 5) & .Faccion.Matados(enemigo(Fuerzas)) & "!°" & str(Npclist(.flags.TargetNPC).Char.CharIndex))
        Call WriteMultiMessage(UserIndex, eMessages.KillToJoin, REQUIERE_MATADOS_PRIMERA, .Faccion.Matados(Enemigo(Fuerzas)), Npclist(.flags.TargetNPC).Char.CharIndex)
        Exit Sub
    End If
    
    If .Stats.ELV < NIVEL_MINIMO_INGRESAR Then
        'Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 6) & str(Npclist(.flags.TargetNPC).Char.CharIndex))
        Call WriteMultiMessage(UserIndex, eMessages.LevelRequired, NIVEL_MINIMO_INGRESAR, Npclist(.flags.TargetNPC).Char.CharIndex)
        Exit Sub
    End If
    
    'Call SendData(ToIndex, UserIndex, 0, Mensajes(Fuerzas, 7) & str(Npclist(.flags.TargetNPC).Char.CharIndex))
    
    Call WriteMultiMessage(UserIndex, eMessages.FactionWelcome, Npclist(.flags.TargetNPC).Char.CharIndex)
    
    .Faccion.Jerarquia = 1
    
    MiObj.Amount = 1
    MiObj.OBJIndex = Armaduras(Fuerzas, .Faccion.Jerarquia, TipoClase(UserIndex), TipoRaza(UserIndex))
    If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(.Pos, MiObj)
    'todo nunca tirar items
End With

End Sub
Public Function Titulo(UserIndex As Integer) As String

Select Case UserList(UserIndex).Faccion.Bando
    Case Real
        Select Case UserList(UserIndex).Faccion.Jerarquia
            Case 0
                Titulo = "Fiel al Rey"
            Case 1
                Titulo = "Soldado Real"
            Case 2
                Titulo = "General Real"
            Case 3
                Titulo = "Elite Real"
            Case 4
                Titulo = "Héroe Real"
        End Select
    Case Caos
        Select Case UserList(UserIndex).Faccion.Jerarquia
            Case 0
                Titulo = "Fiel a Lord Thek"
            Case 1
                Titulo = "Acólito"
            Case 2
                Titulo = "Jefe de Tropas"
            Case 3
                Titulo = "Elite del Mal"
            Case 4
                Titulo = "Héroe del Mal"
        End Select
End Select

End Function

