Attribute VB_Name = "modClases"
Option Explicit

Public Resta(1 To NUMCLASES) As Single
Public Recompensas() As Recompensa
Public AumentoHit(1 To NUMCLASES) As Byte

Sub EstablecerRestas()

Resta(eClass.Ciudadano) = 3
AumentoHit(eClass.Ciudadano) = 3

Resta(eClass.Trabajador) = 2.5
AumentoHit(eClass.Trabajador) = 3

Resta(eClass.Experto_Minerales) = 2.5
AumentoHit(eClass.Experto_Minerales) = 3

Resta(eClass.Minero) = 2.5
AumentoHit(eClass.Minero) = 2

Resta(eClass.Herrero) = 2.5
AumentoHit(eClass.Herrero) = 2

Resta(eClass.Experto_Madera) = 2.5
AumentoHit(eClass.Experto_Madera) = 3

Resta(eClass.Talador) = 2.5
AumentoHit(eClass.Talador) = 2

Resta(eClass.Carpintero) = 2.5
AumentoHit(eClass.Carpintero) = 2

Resta(eClass.Pescador) = 2.5
AumentoHit(eClass.Pescador) = 1

Resta(eClass.Sastre) = 2.5
AumentoHit(eClass.Sastre) = 2

Resta(eClass.Alquimista) = 2.5
AumentoHit(eClass.Alquimista) = 2

Resta(eClass.Luchador) = 3
AumentoHit(eClass.Luchador) = 3

Resta(eClass.Con_Mana) = 3
AumentoHit(eClass.Con_Mana) = 3

Resta(eClass.Hechicero) = 3
AumentoHit(eClass.Hechicero) = 3

Resta(eClass.Mago) = 3
AumentoHit(eClass.Mago) = 1

Resta(eClass.Nigromante) = 3
AumentoHit(eClass.Nigromante) = 1

Resta(eClass.Orden_Sagrada) = 1.5
AumentoHit(eClass.Orden_Sagrada) = 3

Resta(eClass.Paladin) = 0.5
AumentoHit(eClass.Paladin) = 3

Resta(eClass.Clerigo) = 1.5
AumentoHit(eClass.Clerigo) = 2

Resta(eClass.Naturalista) = 2.5
AumentoHit(eClass.Naturalista) = 3

Resta(eClass.Bardo) = 1.5
AumentoHit(eClass.Bardo) = 2

Resta(eClass.Druida) = 3
AumentoHit(eClass.Druida) = 2

Resta(eClass.Sigiloso) = 1.5
AumentoHit(eClass.Sigiloso) = 3

Resta(eClass.Asesino) = 1.5
AumentoHit(eClass.Asesino) = 3

Resta(eClass.Cazador) = 0.5
AumentoHit(eClass.Cazador) = 3

Resta(eClass.Sin_Mana) = 2
AumentoHit(eClass.Sin_Mana) = 2

AumentoHit(eClass.Arquero) = 3

AumentoHit(eClass.Guerrero) = 3

AumentoHit(eClass.Caballero) = 3

AumentoHit(eClass.Bandido) = 2

Resta(eClass.Pirata) = 1.5
AumentoHit(eClass.Pirata) = 2

Resta(eClass.Ladron) = 2.5
AumentoHit(eClass.Ladron) = 2

End Sub

Public Function ClaseBase(ByVal Clase As eClass) As Boolean

ClaseBase = (Clase = eClass.Ciudadano Or Clase = eClass.Trabajador Or Clase = eClass.Experto_Minerales Or _
            Clase = eClass.Experto_Madera Or Clase = eClass.Luchador Or Clase = eClass.Con_Mana Or _
            Clase = Hechicero Or Clase = Orden_Sagrada Or Clase = Naturalista Or _
            Clase = eClass.Sigiloso Or Clase = eClass.Sin_Mana Or Clase = eClass.Bandido Or _
            Clase = eClass.Caballero)

End Function

Public Function PuedeFaccion(ByVal UserIndex As Integer)
    
    With UserList(UserIndex)
    
        PuedeFaccion = Not EsNewbie(UserIndex) And _
        (.Faccion.BandoOriginal = eFaccion.Neutral) And _
        (.flags.Privilegios And PlayerType.User) 'and (.guilindex > 0)
    
    End With
End Function
Public Function PuedeSubirClase(ByVal UserIndex As Integer) As Boolean

    With UserList(UserIndex)
        PuedeSubirClase = (.Stats.ELV >= 3 And .Clase = eClass.Ciudadano) Or _
                    (.Stats.ELV >= 6 And (.Clase = eClass.Luchador Or .Clase = eClass.Trabajador)) Or _
                    (.Stats.ELV >= 9 And (.Clase = eClass.Experto_Minerales Or .Clase = eClass.Experto_Madera Or .Clase = eClass.Con_Mana Or .Clase = eClass.Sin_Mana)) Or _
                    (.Stats.ELV >= 12 And (.Clase = eClass.Caballero Or .Clase = eClass.Bandido Or .Clase = eClass.Hechicero Or .Clase = eClass.Naturalista Or .Clase = eClass.Orden_Sagrada Or .Clase = eClass.Sigiloso))
    
    End With
    
End Function

Function TipoClase(UserIndex As Integer) As Byte

Select Case UserList(UserIndex).Clase
    Case eClass.Paladin, eClass.Asesino, eClass.Cazador
        TipoClase = 2
    Case eClass.Clerigo, eClass.Bardo, eClass.Ladron
        TipoClase = 3
    Case eClass.Mago, eClass.Nigromante, eClass.Druida
        TipoClase = 4
    Case Else
        TipoClase = 1
End Select

End Function
Public Function TipoRaza(UserIndex As Integer) As Byte

If UserList(UserIndex).raza = eRaza.Enano Or UserList(UserIndex).raza = eRaza.Gnomo Then
    TipoRaza = 2
Else: TipoRaza = 1
End If

End Function

Public Function DameClaseFenix(ByVal Clase As eClass) As Byte
'GoDKeR
'Te odio fenix
'Esto debería quitarse cuando se pueda dejar de usar la aolib para calcular el daño del wrestling (en cualquier momento creo yo la función)
'El objetivo de esta "cosa" es dar una clase equivalente a la que hay en fénix ya que la numeración es distinta.

    Select Case CByte(Clase)
        Case 0 To 4
            DameClaseFenix = CByte(Clase)
        Case 5
            DameClaseFenix = 8
        Case 6
            DameClaseFenix = 13
        Case 7
            DameClaseFenix = 14
        Case 8
            DameClaseFenix = 18
        Case 9
            DameClaseFenix = 23
        Case 10
            DameClaseFenix = 27
        Case 11
            DameClaseFenix = 31
        Case 12 To 30
            DameClaseFenix = (CByte(Clase) + 23)
        Case 31, 32
            DameClaseFenix = (CByte(Clase) + 24)
    End Select
    
End Function

Public Sub EnviarRecompensa(ByVal UserIndex As Integer)
    
    With UserList(UserIndex)
        Dim Recom As Integer
        Recom = PuedeRecompensa(UserIndex)
        
        If (Not .flags.Muerto) And Recom Then
            Call WriteShowRecompensaForm(UserIndex, .Clase, Recom)
        End If
    End With
End Sub
Public Sub EstablecerRecompensas()

Recompensas(eClass.Minero, 1, 1).SubeHP = 120

Recompensas(eClass.Mago, 1, 1).Obj(1).OBJIndex = PocionAzulNoCae
Recompensas(eClass.Mago, 1, 1).Obj(1).Amount = 1000
Recompensas(eClass.Mago, 1, 2).Obj(1).OBJIndex = PocionRojaNoCae
Recompensas(eClass.Mago, 1, 2).Obj(1).Amount = 1000
Recompensas(eClass.Mago, 2, 1).SubeHP = 10

Recompensas(eClass.Nigromante, 1, 1).Obj(1).OBJIndex = PocionAzulNoCae
Recompensas(eClass.Nigromante, 1, 1).Obj(1).Amount = 1000
Recompensas(eClass.Nigromante, 1, 2).Obj(1).OBJIndex = PocionRojaNoCae
Recompensas(eClass.Nigromante, 1, 2).Obj(1).Amount = 1000
Recompensas(eClass.Nigromante, 2, 1).SubeHP = 15
Recompensas(eClass.Nigromante, 2, 2).SubeMP = 40

Recompensas(eClass.Paladin, 1, 1).Obj(1).OBJIndex = PocionAzulNoCae
Recompensas(eClass.Paladin, 1, 1).Obj(1).Amount = 1000
Recompensas(eClass.Paladin, 1, 2).Obj(1).OBJIndex = PocionRojaNoCae
Recompensas(eClass.Paladin, 1, 2).Obj(1).Amount = 1000
Recompensas(eClass.Paladin, 2, 1).SubeHP = 5
Recompensas(eClass.Paladin, 2, 1).SubeMP = 10
Recompensas(eClass.Paladin, 2, 2).SubeMP = 30

Recompensas(eClass.Clerigo, 1, 1).Obj(1).OBJIndex = PocionAzulNoCae
Recompensas(eClass.Clerigo, 1, 1).Obj(1).Amount = 1000
Recompensas(eClass.Clerigo, 1, 2).Obj(1).OBJIndex = PocionRojaNoCae
Recompensas(eClass.Clerigo, 1, 2).Obj(1).Amount = 1000
Recompensas(eClass.Clerigo, 2, 1).SubeHP = 10
Recompensas(eClass.Clerigo, 2, 2).SubeMP = 50

Recompensas(eClass.Bardo, 1, 1).Obj(1).OBJIndex = PocionAzulNoCae
Recompensas(eClass.Bardo, 1, 1).Obj(1).Amount = 1000
Recompensas(eClass.Bardo, 1, 2).Obj(1).OBJIndex = PocionRojaNoCae
Recompensas(eClass.Bardo, 1, 2).Obj(1).Amount = 1000
Recompensas(eClass.Bardo, 2, 1).SubeHP = 10
Recompensas(eClass.Bardo, 2, 2).SubeMP = 50

Recompensas(eClass.Druida, 1, 1).Obj(1).OBJIndex = PocionAzulNoCae
Recompensas(eClass.Druida, 1, 1).Obj(1).Amount = 1000
Recompensas(eClass.Druida, 1, 2).Obj(1).OBJIndex = PocionRojaNoCae
Recompensas(eClass.Druida, 1, 2).Obj(1).Amount = 1000
Recompensas(eClass.Druida, 2, 1).SubeHP = 15
Recompensas(eClass.Druida, 2, 2).SubeMP = 40

Recompensas(eClass.Asesino, 1, 1).Obj(1).OBJIndex = PocionAzulNoCae
Recompensas(eClass.Asesino, 1, 1).Obj(1).Amount = 1000
Recompensas(eClass.Asesino, 1, 2).Obj(1).OBJIndex = PocionRojaNoCae
Recompensas(eClass.Asesino, 1, 2).Obj(1).Amount = 1000
Recompensas(eClass.Asesino, 2, 1).SubeHP = 10
Recompensas(eClass.Asesino, 2, 2).SubeMP = 30

Recompensas(eClass.Cazador, 1, 1).Obj(1).OBJIndex = PocionAzulNoCae
Recompensas(eClass.Cazador, 1, 1).Obj(1).Amount = 1000
Recompensas(eClass.Cazador, 1, 2).Obj(1).OBJIndex = PocionRojaNoCae
Recompensas(eClass.Cazador, 1, 2).Obj(1).Amount = 1000
Recompensas(eClass.Cazador, 2, 1).SubeHP = 10
Recompensas(eClass.Cazador, 2, 2).SubeMP = 50

Recompensas(eClass.Arquero, 1, 1).Obj(1).OBJIndex = Flecha
Recompensas(eClass.Arquero, 1, 1).Obj(1).Amount = 1500
Recompensas(eClass.Arquero, 1, 2).Obj(1).OBJIndex = PocionRojaNoCae
Recompensas(eClass.Arquero, 1, 2).Obj(1).Amount = 1000
Recompensas(eClass.Arquero, 2, 1).SubeHP = 10

Recompensas(eClass.Guerrero, 1, 1).Obj(1).OBJIndex = PocionVerdeNoCae
Recompensas(eClass.Guerrero, 1, 1).Obj(1).Amount = 80
Recompensas(eClass.Guerrero, 1, 1).Obj(2).OBJIndex = PocionAmarillaNoCae
Recompensas(eClass.Guerrero, 1, 1).Obj(2).Amount = 100
Recompensas(eClass.Guerrero, 1, 2).Obj(1).OBJIndex = PocionRojaNoCae
Recompensas(eClass.Guerrero, 1, 2).Obj(1).Amount = 1000
Recompensas(eClass.Guerrero, 2, 1).SubeHP = 5

Recompensas(eClass.Pirata, 1, 1).SubeHP = 20
Recompensas(eClass.Pirata, 2, 2).SubeHP = 40
End Sub
