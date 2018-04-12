Attribute VB_Name = "modSubClase"
Option Explicit

Public Sub EnviarSubClase(ByVal UserIndex As Integer)

    If UserList(UserIndex).flags.Muerto = 0 Then
        Call WriteShowClaseForm(UserIndex, UserList(UserIndex).Clase)
    End If
            
End Sub

'CSEH: ErrLog
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

