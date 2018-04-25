Attribute VB_Name = "modSubClase"
Option Explicit

Public Sub EnviarSubClase(ByVal UserIndex As Integer)

    If UserList(UserIndex).flags.Muerto = 0 Then
        Call WriteShowClaseForm(UserIndex, UserList(UserIndex).Clase)
    End If
            
End Sub

'CSEH: ErrLog
Public Sub RecibirFaccion(ByVal UserIndex As Integer, ByVal Faccion As Byte)
    '<EhHeader>
    On Error GoTo RecibirFaccion_Err
    '</EhHeader>
        If Not PuedeFaccion(UserIndex) Then Exit Sub
        
        With UserList(UserIndex)
            If .Faccion.BandoOriginal > 0 Then
            Call WriteConsoleMsg(UserIndex, "Ya eres fiel a un bando.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            End If
        
            If .Faccion.Matados(Faccion) > .Faccion.Matados(Enemigo(Faccion)) Then
                Call WriteConsoleMsg(UserIndex, "La cantidad de matados de tu facción es mayor a la de la facción enemiga.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
     
            Select Case Faccion
                Case 0
                    Call WriteConsoleMsg(UserIndex, "¡Has decidido seguir siendo neutral! Podés jurar fidelidad cuando lo desees.", FontTypeNames.FONTTYPE_INFO)
                Case 1
                    .Faccion.Bando = 1
                    .Faccion.BandoOriginal = 1
                    .Faccion.Ataco(1) = 0
                    Call WarpUserChar(UserIndex, .Pos.Map, .Pos.X, .Pos.Y, False)
                    Call WriteConsoleMsg(UserIndex, "¡Has jurado fidelidad al Rey!", FontTypeNames.FONTTYPE_CONSEJO)
                Case 2
                    .Faccion.Bando = 2
                    .Faccion.BandoOriginal = 2
                    .Faccion.Ataco(2) = 0
                    Call WarpUserChar(UserIndex, .Pos.Map, .Pos.X, .Pos.Y, False)
                    Call WriteConsoleMsg(UserIndex, "¡Has jurado fidelidad a Lord Thek!", FontTypeNames.FONTTYPE_CONSEJOCAOS)
            End Select
        End With

    '<EhFooter>
    Exit Sub

RecibirFaccion_Err:
        Call LogError("Error en RecibirFaccion: " & Erl & " - " & Err.description)
    '</EhFooter>
End Sub
Public Sub HacerMercenario(ByVal UserIndex As Integer, ByVal Faccion As Byte)
    On Error GoTo HacerMercenario_Err

        With UserList(UserIndex)
            If .Faccion.BandoOriginal <> 0 And .Faccion.Bando = .Faccion.BandoOriginal Then
                Call WriteConsoleMsg(UserIndex, "No necesitas aliarte a un bando si ya eres fiel.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            If .Faccion.Bando <> 0 Then
                Call WriteConsoleMsg(UserIndex, "Ya te aliaste a un bando.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            If Not (.flags.Privilegios And PlayerType.User) Then
                Call WriteConsoleMsg(UserIndex, "No puedes aliarte a mortales.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            Select Case Faccion
                Case 1
                    .Faccion.Bando = 1
                    .Faccion.Ataco(1) = 0
                    Call WarpUserChar(UserIndex, .Pos.Map, .Pos.X, .Pos.Y, False)
                    Call WriteConsoleMsg(UserIndex, "¡Decidiste aliarte al Rey!", FontTypeNames.FONTTYPE_CONSEJO)
                Case 2
                    .Faccion.Bando = 2
                    .Faccion.Ataco(2) = 0
                    Call WarpUserChar(UserIndex, .Pos.Map, .Pos.X, .Pos.Y, False)
                    Call WriteConsoleMsg(UserIndex, "¡Decidiste aliarte a Lord Thek!", FontTypeNames.FONTTYPE_CONSEJOCAOS)
            End Select
        End With
    Exit Sub

HacerMercenario_Err:
        Call LogError("Error en HacerMercenario: " & Erl & " - " & Err.description)

End Sub
Public Sub RecibirSubClase(ByVal UserIndex As Integer, ByVal Clase As Byte)
    '<EhHeader>
    On Error GoTo RecibirSubClase_Err
    '</EhHeader>

     If Not PuedeSubirClase(UserIndex) Then Exit Sub

        With UserList(UserIndex)
            Select Case .Clase
            Case eClass.Ciudadano
                If Clase = 1 Then
                    .Clase = eClass.Trabajador
                Else
                    .Clase = eClass.Luchador
                End If
            
             Case eClass.Trabajador
                Select Case Clase
                    Case 1: .Clase = eClass.Experto_Minerales
                    Case 2: .Clase = eClass.Experto_Madera
                    Case 3: .Clase = eClass.Pescador
                    Case 4: .Clase = eClass.Sastre
                End Select
            
            Case eClass.Experto_Minerales
                 If Clase = 1 Then
                    .Clase = eClass.Minero
                 Else
                    .Clase = eClass.Herrero
                 End If
            
            Case eClass.Experto_Madera
                 If Clase = 1 Then
                    .Clase = eClass.Talador
                 Else
                    .Clase = eClass.Carpintero
                 End If
            
            Case eClass.Luchador
                If Clase = 1 Then
                    .Clase = eClass.Con_Mana
                    Call AprenderHechizo(UserIndex, 2)
                    .Stats.MinMAN = 100
                    .Stats.MaxMAN = 100
                    Call WriteUpdateUserStats(UserIndex)
                    If Not PuedeSubirClase(UserIndex) Then Call WriteSubeClase(UserIndex, False)
                Else
                    .Clase = eClass.Sin_Mana
                End If
            
            Case eClass.Con_Mana
                 Select Case Clase
                    Case 1: .Clase = eClass.Hechicero
                    Case 2: .Clase = eClass.Orden_Sagrada
                    Case 3: .Clase = eClass.Naturalista
                    Case 4: .Clase = eClass.Sigiloso
                 End Select
            
            Case eClass.Hechicero
                 If Clase = 1 Then
                    .Clase = eClass.Mago
                 Else
                    .Clase = eClass.Nigromante
                 End If
            
            Case eClass.Orden_Sagrada
                 If Clase = 1 Then
                    .Clase = eClass.Paladin
                 Else
                    .Clase = eClass.Clerigo
                 End If
                 
            Case eClass.Naturalista
                If Clase = 1 Then
                    .Clase = eClass.Bardo
                Else
                    .Clase = eClass.Druida
                End If
            
            Case eClass.Sigiloso
                 If Clase = 1 Then
                    .Clase = eClass.Asesino
                 Else
                    .Clase = eClass.Cazador
                 End If
            
            Case eClass.Sin_Mana
                 If Clase = 1 Then
                    .Clase = eClass.Bandido
                 Else
                    .Clase = eClass.Caballero
                 End If
            
            Case eClass.Bandido
                 If Clase = 1 Then
                    .Clase = eClass.Pirata
                 Else
                    .Clase = eClass.Ladron
                 End If
            
            Case eClass.Caballero
                 If Clase = 1 Then
                    .Clase = eClass.Guerrero
                 Else
                    .Clase = eClass.Arquero
                 End If
            End Select
        End With

 Call CalcularValores(UserIndex)
 If Not PuedeSubirClase(UserIndex) Then Call WriteSubeClase(UserIndex, False)

    '<EhFooter>
    Exit Sub

RecibirSubClase_Err:
        Call LogError("Error en RecibirSubClase: " & Erl & " - " & Err.description)
    '</EhFooter>
End Sub

'CSEH: ErrLog
Public Sub RecibirRecompensa(ByVal UserIndex As Integer, ByVal Eleccion As Byte)
    '<EhHeader>
    On Error GoTo RecibirRecompensa_Err
    '</EhHeader>
    Dim Recompensa As Byte
    Dim i As Integer

 Recompensa = PuedeRecompensa(UserIndex)

 If Recompensa = 0 Then Exit Sub

 UserList(UserIndex).Recompensas(Recompensa) = Eleccion

 If Recompensas(UserList(UserIndex).Clase, Recompensa, Eleccion).SubeHP Then
     Call AddtoVar(UserList(UserIndex).Stats.MaxHp, Recompensas(UserList(UserIndex).Clase, Recompensa, Eleccion).SubeHP, STAT_MAXHP)
        'Call WriteUpdateHP(UserIndex)
     Call WriteUpdateUserStats(UserIndex)
 End If

 If Recompensas(UserList(UserIndex).Clase, Recompensa, Eleccion).SubeMP Then
     Call AddtoVar(UserList(UserIndex).Stats.MaxMAN, Recompensas(UserList(UserIndex).Clase, Recompensa, Eleccion).SubeMP, 2000 + 200 * Buleano(UserList(UserIndex).Clase = Mago) * 200 + 300 * Buleano(UserList(UserIndex).Clase = Mago And UserList(UserIndex).Recompensas(2) = 2))
        'Call WriteUpdateMana(UserIndex)
     Call WriteUpdateUserStats(UserIndex)
 End If

 For i = 1 To 2
     If Recompensas(UserList(UserIndex).Clase, Recompensa, Eleccion).Obj(i).OBJIndex Then
         If Not MeterItemEnInventario(UserIndex, Recompensas(UserList(UserIndex).Clase, Recompensa, Eleccion).Obj(i)) Then Call TirarItemAlPiso(UserList(UserIndex).Pos, Recompensas(UserList(UserIndex).Clase, Recompensa, Eleccion).Obj(i))
    End If
 Next i

If PuedeRecompensa(UserIndex) = 0 Then Call WriteEligeRecompensa(UserIndex, False)

    '<EhFooter>
    Exit Sub

RecibirRecompensa_Err:
        Call LogError("Error en RecibirRecompensa: " & Erl & " - " & Err.description)
    '</EhFooter>
End Sub

