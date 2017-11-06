Attribute VB_Name = "modGameIni"
Option Explicit

Public Type tCabecera 'Cabecera de los con
    Desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public Type tGameIni
    Puerto As Long
    Musica As Byte
    fX As Byte
    tip As Byte
    Password As String
    Name As String
    DirGraficos As String
    DirSonidos As String
    DirMusica As String
    DirMapas As String
    NumeroDeBMPs As Long
    NumeroMapas As Integer
End Type

Public MiCabecera As tCabecera
Public Config_Inicio As tGameIni

Public Sub IniciarCabecera(ByRef Cabecera As tCabecera)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
Cabecera.Desc = "Argentum Online by Noland Studios. Copyright Noland-Studios 2001, pablomarquez@noland-studios.com.ar"
Cabecera.CRC = Rnd * 100
Cabecera.MagicWord = Rnd * 10
End Sub

Public Function LeerGameIni() As tGameIni
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
Dim N As Integer
Dim GameIni As tGameIni
N = FreeFile
Open DirIndex & "Inicio.con" For Binary As #N
Get #N, , MiCabecera

Get #N, , GameIni

Close #N
LeerGameIni = GameIni
End Function

Public Sub EscribirGameIni(ByRef GameIniConfiguration As tGameIni)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
Dim N As Integer
N = FreeFile
Open DirIndex & "Inicio.con" For Binary As #N
Put #N, , MiCabecera
GameIniConfiguration.Password = "DAMMLAMERS!"
Put #N, , GameIniConfiguration
Close #N
End Sub

