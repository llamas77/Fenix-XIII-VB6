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

100     If Not PuedeSubirClase(UserIndex) Then Exit Sub

105     With UserList(UserIndex)

110         Select Case .Clase

                Case eClass.Ciudadano
115                 If Clase = 1 Then
120                     .Clase = eClass.Trabajador
                    Else
125                     .Clase = eClass.Luchador
                    End If
            
130             Case eClass.Trabajador
135                 Select Case Clase
                
                        Case 1: .Clase = eClass.Experto_Minerales
140                     Case 2: .Clase = eClass.Experto_Madera
145                     Case 3: .Clase = eClass.Pescador
150                     Case 4: .Clase = eClass.Sastre
                
                    End Select
            
155             Case eClass.Experto_Minerales
160                 If Clase = 1 Then
165                     .Clase = eClass.Minero
                    Else
170                     .Clase = eClass.Herrero
                    End If
            
175             Case eClass.Experto_Madera
180                 If Clase = 1 Then
185                     .Clase = eClass.Talador
                    Else
190                     .Clase = eClass.Carpintero
                    End If
            
195             Case eClass.Luchador
200                 If Clase = 1 Then
205                     .Clase = eClass.Con_Mana
210                     Call AprenderHechizo(UserIndex, 2)

215                     .Stats.MaxMAN = 100

225                     If Not PuedeSubirClase(UserIndex) Then Call WriteSubeClase(UserIndex, False)
                    Else
230                     .Clase = eClass.Sin_Mana
                    End If
            
235             Case eClass.Con_Mana
240                 Select Case Clase
                
                        Case 1: .Clase = eClass.Hechicero
245                     Case 2: .Clase = eClass.Orden_Sagrada
250                     Case 3: .Clase = eClass.Naturalista
255                     Case 4: .Clase = eClass.Sigiloso
                    
                
                    End Select
            
260             Case eClass.Hechicero
265                 If Clase = 1 Then
270                     .Clase = eClass.Mago
                    Else
275                     .Clase = eClass.Nigromante
                    End If
            
280             Case eClass.Orden_Sagrada
285                 If Clase = 1 Then
290                     .Clase = eClass.Paladin
                    Else
295                     .Clase = eClass.Clerigo
                    End If
300             Case eClass.Naturalista
305                 If Clase = 1 Then
310                     .Clase = eClass.Bardo
                    Else
315                     .Clase = eClass.Druida
                    End If
            
320             Case eClass.Sigiloso
325                 If Clase = 1 Then
330                     .Clase = eClass.Asesino
                    Else
335                     .Clase = eClass.Cazador
                    End If
            
340             Case eClass.Sin_Mana
345                 If Clase = 1 Then
350                     .Clase = eClass.Bandido
                    Else
355                     .Clase = eClass.Caballero
                    End If
            
360             Case eClass.Bandido
365                 If Clase = 1 Then
370                     .Clase = eClass.Pirata
                    Else
375                     .Clase = eClass.Ladron
                    End If
            
380             Case eClass.Caballero
385                 If Clase = 1 Then
390                     .Clase = eClass.Guerrero
                    Else
395                     .Clase = eClass.Arquero
                    End If
            End Select

        End With

400 Call CalcularValores(UserIndex)
405 If Not PuedeSubirClase(UserIndex) Then Call WriteSubeClase(UserIndex, False)

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

100 Recompensa = PuedeRecompensa(UserIndex)

105 If Recompensa = 0 Then Exit Sub

110 UserList(UserIndex).Recompensas(Recompensa) = Eleccion

115 If Recompensas(UserList(UserIndex).Clase, Recompensa, Eleccion).SubeHP Then
120     Call AddtoVar(UserList(UserIndex).Stats.MaxHp, Recompensas(UserList(UserIndex).Clase, Recompensa, Eleccion).SubeHP, STAT_MAXHP)
        'Call WriteUpdateHP(UserIndex)
125     Call WriteUpdateUserStats(UserIndex)
    End If

130 If Recompensas(UserList(UserIndex).Clase, Recompensa, Eleccion).SubeMP Then
135     Call AddtoVar(UserList(UserIndex).Stats.MaxMAN, Recompensas(UserList(UserIndex).Clase, Recompensa, Eleccion).SubeMP, 2000 + 200 * Buleano(UserList(UserIndex).Clase = Mago) * 200 + 300 * Buleano(UserList(UserIndex).Clase = Mago And UserList(UserIndex).Recompensas(2) = 2))
        'Call WriteUpdateMana(UserIndex)
140     Call WriteUpdateUserStats(UserIndex)
    End If

145 For i = 1 To 2
150     If Recompensas(UserList(UserIndex).Clase, Recompensa, Eleccion).Obj(i).OBJIndex Then
155         If Not MeterItemEnInventario(UserIndex, Recompensas(UserList(UserIndex).Clase, Recompensa, Eleccion).Obj(i)) Then Call TirarItemAlPiso(UserList(UserIndex).Pos, Recompensas(UserList(UserIndex).Clase, Recompensa, Eleccion).Obj(i))
        End If
    Next

160 If PuedeRecompensa(UserIndex) = 0 Then Call WriteEligeRecompensa(UserIndex, False)

    '<EhFooter>
    Exit Sub

RecibirRecompensa_Err:
        Call LogError("Error en RecibirRecompensa: " & Erl & " - " & Err.description)
    '</EhFooter>
End Sub

