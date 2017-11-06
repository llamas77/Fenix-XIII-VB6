Attribute VB_Name = "modMapIO"
Option Explicit

Private MapTitulo As String     ' GS > Almacena el titulo del mapa para el .dat

''
' Obtener el tamaño de un archivo
'
' @param FileName Especifica el path del archivo
' @return   Nos devuelve el tamaño

Public Function FileSize(ByVal FileName As String) As Long

    On Error GoTo FalloFile
    Dim nFileNum As Integer
    Dim lFileSize As Long
    
    nFileNum = FreeFile
    Open FileName For Input As nFileNum
    lFileSize = LOF(nFileNum)
    Close nFileNum
    FileSize = lFileSize
    
    Exit Function
FalloFile:
    FileSize = -1
End Function

''
' Nos dice si existe el archivo/directorio
'
' @param file Especifica el path
' @param FileType Especifica el tipo de archivo/directorio
' @return   Nos devuelve verdadero o falso

Public Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 26/05/06
'*************************************************
If LenB(Dir(file, FileType)) = 0 Then
    FileExist = False
Else
    FileExist = True
End If

End Function

''
' Guarda el Mapa
'
' @param Path Especifica el path del mapa

Public Sub GuardarMapa(Optional Path As String)
'*************************************************
'Author: Unkwown
'Last modified: 07/01/10 - Deut
'*************************************************
Dim loopc As Integer
Dim Y As Integer
Dim X As Integer
Dim InfoTile As Byte

frmMain.Dialog.CancelError = True

If LenB(Path) = 0 Then
    frmMain.ObtenerNombreArchivo True
    Path = frmMain.Dialog.FileName
    If LenB(Path) = 0 Then Exit Sub
End If

'Borramos el mcl anterior si existe
If FileExist(Path, vbNormal) = True Then
    If MsgBox("¿Desea sobrescribir " & Path & "?", vbCritical + vbYesNo) = vbNo Then
        Exit Sub
    Else
        Kill Path
    End If
End If

frmMain.MousePointer = 11

'Borramos el .msv tambien
If FileExist(Left$(Path, Len(Path) - 4) & ".msv", vbNormal) = True Then
    Kill Left$(Path, Len(Path) - 4) & ".msv"
End If

'Escribimos el nuevo mapa..
Open Path For Binary As #1
    Seek #1, 1
    Put #1, , MapInfo.MapVersion 'aca guardo un integer (2 bytes)
    
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
    
            InfoTile = 0
            
            If MapData(X, Y).Blocked Then InfoTile = InfoTile Or 1
            If MapData(X, Y).Graphic(2).GrhIndex Then InfoTile = InfoTile Or 2
            If MapData(X, Y).Graphic(3).GrhIndex Then InfoTile = InfoTile Or 4
            If MapData(X, Y).Graphic(4).GrhIndex Then InfoTile = InfoTile Or 8
            
            If MapData(X, Y).Trigger <> 0 Then
                Rem Nueva forma de guardar los triggers - Por mí
'                InfoTile = CByte(MapData(X, Y).Trigger * 16)
                
                Rem forma vieja de guardar los triggers - anda bien hasta el trigger 7 inclusive
                If MapData(X, Y).Trigger = (MapData(X, Y).Trigger Or 1) Then InfoTile = InfoTile Or 16
                If MapData(X, Y).Trigger = (MapData(X, Y).Trigger Or 2) Then InfoTile = InfoTile Or 32
                If MapData(X, Y).Trigger = (MapData(X, Y).Trigger Or 4) Then InfoTile = InfoTile Or 64

            End If
            
            Put #1, , InfoTile 'aca guardo un byte
            Put #1, , MapData(X, Y).Graphic(1).GrhIndex 'aca guardo un integer (2 bytes)
            
            For loopc = 2 To 4
                If MapData(X, Y).Graphic(loopc).GrhIndex Then
                    Put #1, , MapData(X, Y).Graphic(loopc).GrhIndex 'aca guardo entre 0 y 3 integers (entre 0 y 6 bytes)
                End If
            Next loopc
                
        Next X
    Next Y
Close #1


Path = Left$(Path, Len(Path) - 3)
Path = Path & "msv"

Open Path For Binary As #2
    Seek #2, 1
    Put #2, , MapInfo.MapVersion
    
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            InfoTile = 0
            
            If MapData(X, Y).Blocked Then InfoTile = InfoTile Or 1
            If HayAgua(X, Y) Then InfoTile = InfoTile Or 2
            
            If MapData(X, Y).Trigger <> 0 Then
                Rem forma nueva
'                InfoTile = CByte(MapData(X, Y).Trigger * 16) MAL!
                
                Rem forma vieja
                If MapData(X, Y).Trigger = (MapData(X, Y).Trigger Or 1) Then InfoTile = InfoTile Or 4
                If MapData(X, Y).Trigger = (MapData(X, Y).Trigger Or 2) Then InfoTile = InfoTile Or 8
                If MapData(X, Y).Trigger = (MapData(X, Y).Trigger Or 4) Then InfoTile = InfoTile Or 16
            End If
                      
            If MapData(X, Y).NpcIndex Then InfoTile = InfoTile Or 32
            If MapData(X, Y).OBJInfo.objindex Then InfoTile = InfoTile Or 64
            If MapData(X, Y).TileExit.Map Then InfoTile = InfoTile Or 128
    
            Put #2, , InfoTile
            
            If MapData(X, Y).NpcIndex Then
                Put #2, , CInt(MapData(X, Y).NpcIndex)
            End If
                        
            If MapData(X, Y).OBJInfo.objindex Then
                Put #2, , MapData(X, Y).OBJInfo.objindex
                Put #2, , MapData(X, Y).OBJInfo.Amount
            End If
            
            If MapData(X, Y).TileExit.Map Then
                Put #2, , MapData(X, Y).TileExit.Map
                Put #2, , MapData(X, Y).TileExit.X
                Put #2, , MapData(X, Y).TileExit.Y
            End If
        Next X
    Next Y
Close #2

Call Pestañas(Path)
Call MapInfo_Guardar(Path)
frmMain.MousePointer = 0
MapInfo.Changed = 0

End Sub

''
' Nos pregunta donde guardar el mapa en caso de modificarlo
'
' @param Path Especifica si existiera un path donde guardar el mapa

Public Sub DeseaGuardarMapa(Optional Path As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

If MapInfo.Changed = 1 Then
    If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then
        Call GuardarMapa(Path)
    End If
End If
End Sub

''
' Limpia todo el mapa a uno nuevo
'

Public Sub NuevoMapa()
'*************************************************
'Author: ^[GS]^
'Last modified: 21/05/06
'*************************************************

On Error Resume Next

Dim loopc As Integer
Dim Y As Integer
Dim X As Integer

bAutoGuardarMapaCount = 0

frmMain.mnuReAbrirMapa.Enabled = False
frmMain.TimAutoGuardarMapa.Enabled = False
frmMain.lblMapVersion.Caption = 0

MapaCargado = False

For loopc = 0 To frmMain.MapPest.Count - 1
    frmMain.MapPest(loopc).Enabled = False
Next

frmMain.MousePointer = 11

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
    
        ' Capa 1
        MapData(X, Y).Graphic(1).GrhIndex = 1
        
        ' Bloqueos
        MapData(X, Y).Blocked = 0

        ' Capas 2, 3 y 4
        MapData(X, Y).Graphic(2).GrhIndex = 0
        MapData(X, Y).Graphic(3).GrhIndex = 0
        MapData(X, Y).Graphic(4).GrhIndex = 0

        ' NPCs
        If MapData(X, Y).NpcIndex > 0 Then
            EraseChar MapData(X, Y).CharIndex
            MapData(X, Y).NpcIndex = 0
        End If

        ' OBJs
        MapData(X, Y).OBJInfo.objindex = 0
        MapData(X, Y).OBJInfo.Amount = 0
        MapData(X, Y).ObjGrh.GrhIndex = 0

        ' Translados
        MapData(X, Y).TileExit.Map = 0
        MapData(X, Y).TileExit.X = 0
        MapData(X, Y).TileExit.Y = 0
        
        ' Triggers
        MapData(X, Y).Trigger = 0
        
        InitGrh MapData(X, Y).Graphic(1), 1
    Next X
Next Y

'Map Info
MapInfo.MapVersion = 0
MapInfo.Name = "Nuevo Mapa"
MapInfo.Music = 0
MapInfo.Pk = True
MapInfo.NoMagia = 0
MapInfo.Terreno = "BOSQUE"
MapInfo.Zona = "CAMPO"
MapInfo.Restringir = False
MapInfo.Nivel = vbNull
MapInfo.LeftPunto = 0
MapInfo.TopPunto = 0

Call MapInfo_Actualizar

bRefreshRadar = True ' Radar

'Set changed flag
MapInfo.Changed = 0
frmMain.MousePointer = 0

' Vacio deshacer
modEdicion.Deshacer_Clear

MapaCargado = True
EngineRun = True

frmMain.SetFocus

End Sub

''
' Abrir Mapa
'
' @param Map Especifica el Path del mapa

Public Sub CargarMapa(ByVal Map As String)

On Error Resume Next
    Dim loopc As Integer
    Dim Y As Integer
    Dim X As Integer
    Dim InfoTile As Byte
    Dim i As Byte
    Dim Body As Integer
    Dim Head As Integer
    Dim Heading As Byte
    DoEvents
    
    frmMain.MousePointer = 11
    
    'Cargamos el .MCL
    Open Map For Binary As #1

    Seek #1, 1
    Get #1, , MapInfo.MapVersion

    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            Get #1, , InfoTile
            
'            If InfoTile <> 0 And InfoTile <> 1 And InfoTile <> 3 And InfoTile <> 5 Then
'                InfoTile = InfoTile
'            End If
            
            
            Rem forma vieja de cargar un bloqueo
            Rem lo que hace AND:
'            0 and 0 = 0
'            0 and 1 = 0
'            1 and 0 = 0
'            1 and 1 = 1
            MapData(X, Y).Blocked = (InfoTile And 1) ' osea lo que está haciendo acá es diciendo que si infotile vale 1 hay un bloqueo en ese tile
            
            Rem forma nueva de cargar un bloqueo
'            If InfoTile = 1 Then MapData(X, Y).Blocked = 1
            
            Get #1, , MapData(X, Y).Graphic(1).GrhIndex
            InitGrh MapData(X, Y).Graphic(1), MapData(X, Y).Graphic(1).GrhIndex
            
            
            
            For i = 2 To 4
                If InfoTile And (2 ^ (i - 1)) Then
                    Get #1, , MapData(X, Y).Graphic(i).GrhIndex
                    InitGrh MapData(X, Y).Graphic(i), MapData(X, Y).Graphic(i).GrhIndex
                Else
                    MapData(X, Y).Graphic(i).GrhIndex = 0
                End If
            Next
            
            MapData(X, Y).Trigger = 0
            
            
            
            Rem forma vieja de cargar los triggers - funciona hasta el trigger 7 inclusive, pero más no
            For i = 4 To 6
                If (InfoTile And 2 ^ i) Then MapData(X, Y).Trigger = MapData(X, Y).Trigger Or 2 ^ (i - 4)
            Next
            
            
            
        Next X
    Next Y
    Close #1
    
    'Cargamos el .MSV
    Map = Left$(Map, Len(Map) - 3)
    Map = Map & "msv"
    
    Open Map For Binary As #2

    Seek #2, 1
    Get #2, , MapInfo.MapVersion

    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            Get #2, , InfoTile
            
            MapData(X, Y).Blocked = (InfoTile And 1)
            MapData(X, Y).Agua = Buleano(InfoTile And 2)
            MapData(X, Y).Trigger = 0
            
            For loopc = 2 To 4
                If (InfoTile And 2 ^ loopc) Then MapData(X, Y).Trigger = MapData(X, Y).Trigger Or 2 ^ (loopc - 2)
            Next
            
            If InfoTile And 32 Then
                Get #2, , MapData(X, Y).NpcIndex
    
                If MapData(X, Y).NpcIndex < 0 Then
                    MapData(X, Y).NpcIndex = 0
                Else
                    Body = NpcData(MapData(X, Y).NpcIndex).Body
                    Head = NpcData(MapData(X, Y).NpcIndex).Head
                    Heading = NpcData(MapData(X, Y).NpcIndex).Heading
                    Call MakeChar(NextOpenChar(), Body, Head, Heading, X, Y)
                End If
            End If

            If InfoTile And 64 Then
                Get #2, , MapData(X, Y).OBJInfo.objindex
                Get #2, , MapData(X, Y).OBJInfo.Amount
                If MapData(X, Y).OBJInfo.objindex > 0 Then InitGrh MapData(X, Y).ObjGrh, ObjData(MapData(X, Y).OBJInfo.objindex).GrhIndex
            End If
            
            If InfoTile And 128 Then
                Get #2, , MapData(X, Y).TileExit.Map
                Get #2, , MapData(X, Y).TileExit.X
                Get #2, , MapData(X, Y).TileExit.Y
            End If

        Next X
    Next Y
    
    Close #2

    
    Call Pestañas(Map)
    
    bRefreshRadar = True ' Radar
    
    Call MapInfo_Cargar(Map)
    frmMain.lblMapVersion.Caption = MapInfo.MapVersion
    
    'Set changed flag
    MapInfo.Changed = 0
    
    ' Vacia el Deshacer
    modEdicion.Deshacer_Clear
    
    'Change mouse icon
    frmMain.MousePointer = 0
    MapaCargado = True
    
End Sub

' *****************************************************************************
' MAPINFO *********************************************************************
' *****************************************************************************

''
' Guardar Informacion del Mapa (Info.dat)
'
' @param Archivo Especifica el Path del archivo .mcl

Public Sub MapInfo_Guardar(ByVal Archivo As String)

Dim InfoDatPath As String

    If LenB(MapTitulo) = 0 Then
        MapTitulo = NameMap_Save
    End If

    InfoDatPath = Left$(Archivo, Len(Archivo) - (Len(MapTitulo) + 4))
    
    Call WriteVar(InfoDatPath & "Info.dat", MapTitulo, "Name", MapInfo.Name)
    Call WriteVar(InfoDatPath & "Info.dat", MapTitulo, "MusicNum", MapInfo.Music)
    Call WriteVar(InfoDatPath & "Info.dat", MapTitulo, "NoMagia", Val(MapInfo.NoMagia))
    Call WriteVar(InfoDatPath & "Info.dat", MapTitulo, "Nivel", Val(MapInfo.Nivel))
    Call WriteVar(InfoDatPath & "Info.dat", MapTitulo, "Terreno", MapInfo.Terreno)
    Call WriteVar(InfoDatPath & "Info.dat", MapTitulo, "Zona", MapInfo.Zona)
    Call WriteVar(InfoDatPath & "Info.dat", MapTitulo, "Restringir", Val(MapInfo.Restringir))
    Call WriteVar(InfoDatPath & "Info.dat", MapTitulo, "BackUp", Val(MapInfo.BackUp))

    If MapInfo.Pk Then
        Call WriteVar(InfoDatPath & "Info.dat", MapTitulo, "Pk", "0")
    Else
        Call WriteVar(InfoDatPath & "Info.dat", MapTitulo, "Pk", "1")
    End If
    
End Sub

''
' Abrir Informacion del Mapa (.dat)
'
' @param Archivo Especifica el Path del archivo .mcl

Public Sub MapInfo_Cargar(ByVal Archivo As String)

Dim MapName As String, InfoDatPath As String

    MapName = Dir$(Archivo)
    InfoDatPath = Left$(Archivo, Len(Archivo) - (Len(MapName)))
    MapName = Left$(MapName, Len(MapName) - 4)

    MapInfo.Name = GetVar(InfoDatPath & "Info.dat", MapName, "Name")
    MapInfo.Music = Val(GetVar(InfoDatPath & "Info.dat", MapName, "MusicNum"))
    
    If Val(GetVar(InfoDatPath & "Info.dat", MapName, "Pk")) = 0 Then
        MapInfo.Pk = True
    Else
        MapInfo.Pk = False
    End If
    
    MapInfo.NoMagia = Val(InfoDatPath & GetVar("Info.dat", MapName, "NoMagia"))
    MapInfo.Terreno = GetVar(InfoDatPath & "Info.dat", MapName, "Terreno")
    MapInfo.Zona = GetVar(InfoDatPath & "Info.dat", MapName, "Zona")
    MapInfo.Restringir = Val(GetVar(InfoDatPath & "Info.dat", MapName, "Restringir"))
    MapInfo.BackUp = Val(GetVar(InfoDatPath & "Info.dat", MapName, "BackUp"))
    
    MapInfo_Actualizar
    
End Sub

''
' Actualiza el formulario de MapInfo
'

Public Sub MapInfo_Actualizar()

On Error Resume Next
    ' Mostrar en Formularios
    frmMapInfo.txtMapNombre.Text = MapInfo.Name
    frmMapInfo.txtMapMusica.Text = MapInfo.Music
    frmMapInfo.txtMapTerreno.Text = MapInfo.Terreno
    frmMapInfo.txtMapZona.Text = MapInfo.Zona
    frmMapInfo.chkRestringir.value = MapInfo.Restringir
    frmMapInfo.chkMapBackup.value = MapInfo.BackUp
    frmMapInfo.chkMapMagiaSinEfecto.value = MapInfo.NoMagia
    frmMapInfo.chkMapPK.value = IIf(MapInfo.Pk = True, 1, 0)
    frmMapInfo.txtMapVersion = MapInfo.MapVersion
    frmMapInfo.txtMapLevel = MapInfo.Nivel
    frmMain.lblMapNombre = MapInfo.Name
    frmMain.lblMapMusica = MapInfo.Music

End Sub

''
' Calcula la orden de Pestañas
'
' @param Map Especifica path del mapa

Public Sub Pestañas(ByVal Map As String)

On Error Resume Next
Dim loopc As Integer

For loopc = Len(Map) To 1 Step -1
    If mid(Map, loopc, 1) = "\" Then
        PATH_Save = Left(Map, loopc)
        Exit For
    End If
Next
Map = Right(Map, Len(Map) - (Len(PATH_Save)))
For loopc = Len(Left(Map, Len(Map) - 4)) To 1 Step -1
    If IsNumeric(mid(Left(Map, Len(Map) - 4), loopc, 1)) = False Then
        NumMap_Save = Right(Left(Map, Len(Map) - 4), Len(Left(Map, Len(Map) - 4)) - loopc)
        NameMap_Save = Left(Map, loopc)
        Exit For
    End If
Next
For loopc = (NumMap_Save - 4) To (NumMap_Save + 8)
        If FileExist(PATH_Save & NameMap_Save & loopc & ".mcl", vbArchive) = True Then
            frmMain.MapPest(loopc - NumMap_Save + 4).Visible = True
            frmMain.MapPest(loopc - NumMap_Save + 4).Enabled = True
            frmMain.MapPest(loopc - NumMap_Save + 4).Caption = NameMap_Save & loopc
        Else
            frmMain.MapPest(loopc - NumMap_Save + 4).Visible = False
        End If
Next
End Sub
