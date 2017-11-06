Attribute VB_Name = "Mod_TileEngine"
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Public Enum eRenderState 'Ojo con cambiar esto, tener en cuenta que esta hardcodeado en el mod_Components
            eLogin = 0
            eNewCharInfo
            eNewCharDetails
            eNewCharAttrib
            eNewCharSkills
End Enum

Private FadingRState As eRenderState
Private RenderState As eRenderState
Private BodyExample As Grh
Private FadeOff As Boolean
Private FadeOn As Boolean
Private ConnectAlpha As Integer

Private Type tHelpWindow
            Active As Boolean
            Text() As String 'lineeees
End Type: Public HelpWindow As tHelpWindow

Private Type D3DXIMAGE_INFO_A
    Width As Long
    Height As Long
    Depth As Long
    MipLevels As Long
    Format As CONST_D3DFORMAT
    ResourceType As CONST_D3DRESOURCETYPE
    ImageFileFormat As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type
        
Private Type CharVA
    X As Integer
    Y As Integer
    W As Integer
    H As Integer
    
    Tx1 As Single
    Tx2 As Single
    Ty1 As Single
    Ty2 As Single
End Type

Private Type VFH
    BitmapWidth As Long         'Size of the bitmap itself
    BitmapHeight As Long
    CellWidth As Long           'Size of the cells (area for each character)
    CellHeight As Long
    BaseCharOffset As Byte      'The character we start from
    CharWidth(0 To 255) As Byte 'The actual factual width of each character
    CharVA(0 To 255) As CharVA
End Type

Private Type CustomFont
    HeaderInfo As VFH           'Holds the header information
    Texture As Direct3DTexture8 'Holds the texture of the text
    RowPitch As Integer         'Number of characters per row
    RowFactor As Single         'Percentage of the texture width each character takes
    ColFactor As Single         'Percentage of the texture height each character takes
    CharHeight As Byte          'Height to use for the text - easiest to start with CellHeight value, and keep lowering until you get a good value
    TextureSize As POINTAPI     'Size of the texture
End Type

Public cfonts(1 To 2) As CustomFont ' _Default2 As CustomFont

Private LastInvRender As Long

Private SurfaceDB As clsSurfaceDB
Private SpriteBatch As clsBatch

Public DirectX As DirectX8
Public DirectD3D8 As D3DX8
Public DirectD3D As Direct3D8
Public DirectDevice As Direct3DDevice8

Private MainUITex As Direct3DTexture8
Private LoginTex  As Direct3DTexture8

Private Viewport As D3DVIEWPORT8
Private Projection As D3DMATRIX
Private View As D3DMATRIX
'Private Translation As D3DMATRIX
    
Private OffsetCounterX As Single
Private OffsetCounterY As Single

Private MainViewRect As D3DRECT
Private ConnectRect As D3DRECT

'Map sizes in tiles
Public Const XMaxMapSize As Byte = 100
Public Const XMinMapSize As Byte = 1
Public Const YMaxMapSize As Byte = 100
Public Const YMinMapSize As Byte = 1

Private Const GrhFogata As Integer = 1521

''
'Sets a Grh animation to loop indefinitely.
Private Const INFINITE_LOOPS As Integer = -1

'Posicion en un mapa
Public Type Position
    X As Long
    Y As Long
End Type

'Posicion en el Mundo
Public Type WorldPos
    Map As Integer
    X As Integer
    Y As Integer
End Type

'Contiene info acerca de donde se puede encontrar un grh tamaño y animacion
Public Type GrhData
    sX As Integer
    sY As Integer
    
    FileNum As Long
    
    pixelWidth As Integer
    pixelHeight As Integer
    
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames(1 To 25) As Integer 'todo: remover to 25
    
    Speed As Single
End Type

'apunta a una estructura grhdata y mantiene la animacion
Public Type Grh
    GrhIndex As Long
    FrameCounter As Single
    Speed As Single
    Started As Byte
    Loops As Integer
End Type

'Lista de cuerpos
Public Type BodyData
    Walk(E_Heading.NORTH To E_Heading.WEST) As Grh
    HeadOffset As Position
End Type

'Lista de cabezas
Public Type HeadData
    Head(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'Lista de las animaciones de las armas
Type WeaponAnimData
    WeaponWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData
    ShieldWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type


'Apariencia del personaje
Public Type Char
    Active As Byte
    Heading As E_Heading
    Pos As Position
    
    iHead As Integer
    iBody As Integer
    Body As BodyData
    Head As HeadData
    Casco As HeadData
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    UsandoArma As Boolean
    
    fX As Grh
    FxIndex As Integer
    
    Criminal As Byte
    
    Nombre As String
    NombreOffset As Integer
    
    GuildName As String
    GuildOffset As Integer
    
    ScrollDirectionX As Integer
    ScrollDirectionY As Integer
    
    Moving As Byte
    MoveOffsetX As Single
    MoveOffsetY As Single
    
    pie As Boolean
    muerto As Boolean
    invisible As Boolean
    priv As Byte
End Type

'Info de un objeto
Public Type Obj
    OBJIndex As Integer
    amount As Integer
End Type

'Tipo de las celdas del mapa
Public Type MapBlock
    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjGrh As Grh
    
    NPCIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    
    Trigger As Integer
End Type

'Info de cada mapa
Public Type MapInfo
    Music As String
    Name As String
    StartPos As WorldPos
    MapVersion As Integer
End Type

Public Const DegreeToRadian As Single = 0.01745329251994 'Pi / 180

Public AmbientColor(3) As Long

Public IniPath As String

'Bordes del mapa
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

'Status del user
Public CurMap As Integer 'Mapa actual
Public UserIndex As Integer
Public UserMoving As Byte
Public UserBody As Integer
Public UserHead As Integer
Public UserCharIndex As Integer

Public EngineRun As Boolean

Public FPS As Long
Public FramesPerSecCounter As Long
Private fpsLastCheck As Long

'Tamaño del la vista en Tiles
Private WindowTileWidth As Integer
Private WindowTileHeight As Integer

Private HalfWindowTileWidth As Integer
Private HalfWindowTileHeight As Integer

Public ViewPositionX As Integer
Public ViewPositionY As Integer

Private ScrollDirectionX As Integer
Private ScrollDirectionY As Integer

'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tamaño muy grande puede
'volver el engine muy lento
Public TileBufferSize As Integer

'Private TileBufferPixelOffsetX As Integer
'Private TileBufferPixelOffsetY As Integer

'Tamaño de los tiles en pixels
Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Public ScrollPixelsPerFrameX As Integer
Public ScrollPixelsPerFrameY As Integer

Private timerElapsedTime As Single
Private timerTicksPerFrame As Single
Private animTicksPerFrame As Single
Private engineBaseSpeed As Single
Private engineAnimSpeed As Single

Public NumChars As Integer
Public LastChar As Integer
Public NumWeaponAnims As Integer

Private MainViewWidth   As Integer
Private MainViewHeight  As Integer

Private MouseTileX As Byte
Private MouseTileY As Byte

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Graficos¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public GrhData() As GrhData 'Guarda todos los grh
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As tIndiceFx
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Mapa?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public MapData() As MapBlock ' Mapa
Public MapInfo As MapInfo ' Info acerca del mapa en uso
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Public bRain        As Boolean 'está raineando?
Public bTecho       As Boolean 'hay techo?

Public charlist(1 To 10000) As Char

' Used by GetTextExtentPoint32
Private Type size
    cx As Long
    cy As Long
End Type

Public Enum PlayLoop
    plNone = 0
    plLluviain = 1
    plLluviaout = 2
End Enum

'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Sub CargarCabezas()
    Dim N As Integer
    Dim i As Long
    Dim Numheads As Integer
    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open App.path & "\init\Cabezas.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , Numheads
    
    'Resize array
    ReDim HeadData(0 To Numheads) As HeadData
    ReDim Miscabezas(0 To Numheads) As tIndiceCabeza
    
    For i = 1 To Numheads
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(HeadData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(HeadData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(HeadData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(HeadData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #N

End Sub

Sub CargarCascos()
    Dim N As Integer
    Dim i As Long
    Dim NumCascos As Integer

    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open App.path & "\init\Cascos.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCascos
    
    'Resize array
    ReDim CascoAnimData(0 To NumCascos) As HeadData
    ReDim Miscabezas(0 To NumCascos) As tIndiceCabeza
    
    For i = 1 To NumCascos
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #N
End Sub

Sub CargarCuerpos()
    Dim N As Integer
    Dim i As Long
    Dim NumCuerpos As Integer
    Dim MisCuerpos() As tIndiceCuerpo
    
    N = FreeFile()
    Open App.path & "\init\Personajes.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCuerpos
    
    'Resize array
    ReDim BodyData(0 To NumCuerpos) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
    
    For i = 1 To NumCuerpos
        Get #N, , MisCuerpos(i)
        
        If MisCuerpos(i).Body(1) Then
            InitGrh BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
            InitGrh BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
            InitGrh BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
            InitGrh BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0
            
            BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
        End If
    Next i
    
    Close #N
End Sub

Sub CargarFxs()
    Dim N As Integer
    Dim i As Long
    Dim NumFxs As Integer
    
    N = FreeFile()
    Open App.path & "\init\Fxs.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumFxs
    
    'Resize array
    ReDim FxData(1 To NumFxs) As tIndiceFx
    
    For i = 1 To NumFxs
        Get #N, , FxData(i)
    Next i
    
    Close #N
End Sub

Sub CargarTips()
    Dim N As Integer
    Dim i As Long
    Dim NumTips As Integer
    
    N = FreeFile
    Open App.path & "\init\Tips.ayu" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumTips
    
    'Resize array
    ReDim Tips(1 To NumTips) As String * 255
    
    For i = 1 To NumTips
        Get #N, , Tips(i)
    Next i
    
    Close #N
End Sub

Sub CargarArrayLluvia()
    Dim N As Integer
    Dim i As Long
    Dim Nu As Integer
    
    N = FreeFile()
    Open App.path & "\init\fk.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , Nu
    
    'Resize array
    ReDim bLluvia(1 To Nu) As Byte
    
    For i = 1 To Nu
        Get #N, , bLluvia(i)
    Next i
    
    Close #N
End Sub

Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef TX As Byte, ByRef TY As Byte)
'******************************************
'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
'******************************************
    TX = ViewPositionX + viewPortX \ TilePixelWidth - WindowTileWidth \ 2
    TY = ViewPositionY + viewPortY \ TilePixelHeight - WindowTileHeight \ 2
End Sub

Sub MakeChar(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal X As Integer, ByVal Y As Integer, ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)
On Error Resume Next
    'Apuntamos al ultimo Char
    If CharIndex > LastChar Then LastChar = CharIndex
    
    With charlist(CharIndex)
        'If the char wasn't allready active (we are rewritting it) don't increase char count
        If .Active = 0 Then _
            NumChars = NumChars + 1
        
        If Arma = 0 Then Arma = 2
        If Escudo = 0 Then Escudo = 2
        If Casco = 0 Then Casco = 2
        
        .iHead = Head
        .iBody = Body
        .Head = HeadData(Head)
        .Body = BodyData(Body)
                
        .Arma = WeaponAnimData(Arma)
        
        .Escudo = ShieldAnimData(Escudo)
        .Casco = CascoAnimData(Casco)
        
        .Heading = Heading
        
        'Reset moving stats
        .Moving = 0
        .MoveOffsetX = 0
        .MoveOffsetY = 0
        
        'Update position
        .Pos.X = X
        .Pos.Y = Y
        
        'Make active
        .Active = 1
    End With
    
    'Plot on map
    MapData(X, Y).CharIndex = CharIndex
End Sub

Sub ResetCharInfo(ByVal CharIndex As Integer)
    With charlist(CharIndex)
        .Active = 0
        .Criminal = 0
        .FxIndex = 0
        .invisible = False
        
        .Moving = 0
        .muerto = False
        .Nombre = vbNullString
        .NombreOffset = 0
        
        .GuildName = vbNullString
        .GuildOffset = 0
        
        .pie = False
        .Pos.X = 0
        .Pos.Y = 0
        .UsandoArma = False
    End With
End Sub

Sub EraseChar(ByVal CharIndex As Integer)
'*****************************************************************
'Erases a character from CharList and map
'*****************************************************************
On Error Resume Next
    charlist(CharIndex).Active = 0
    
    'Update lastchar
    If CharIndex = LastChar Then
        Do Until charlist(LastChar).Active = 1
            LastChar = LastChar - 1
            If LastChar = 0 Then Exit Do
        Loop
    End If
    
    MapData(charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y).CharIndex = 0
    
    'Remove char's dialog
    Call Dialogos.RemoveDialog(CharIndex)
    
    Call ResetCharInfo(CharIndex)
    
    'Update NumChars
    NumChars = NumChars - 1
End Sub

'CSEH: ErrLog
Public Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional ByVal Started As Byte = 2)
    '*****************************************************************
    'Sets up a grh. MUST be done before rendering
    '*****************************************************************
    '<EhHeader>
    On Error GoTo InitGrh_Err
    '</EhHeader>
        If GrhIndex = 0 Then Exit Sub

100     Grh.GrhIndex = GrhIndex
    
105     If Started = 2 Then
110         If GrhData(Grh.GrhIndex).NumFrames > 1 Then
115             Grh.Started = 1
            Else
120             Grh.Started = 0
            End If
        Else
            'Make sure the graphic can be started
125         If GrhData(Grh.GrhIndex).NumFrames = 1 Then Started = 0
130         Grh.Started = Started
        End If
    
    
135     If Grh.Started Then
140         Grh.Loops = INFINITE_LOOPS
        Else
145         Grh.Loops = 0
        End If
    
150     Grh.FrameCounter = 1
155     Grh.Speed = GrhData(Grh.GrhIndex).Speed / 1000
    '<EhFooter>
    Exit Sub

InitGrh_Err:
        Call LogError("Error en InitGrh: " & Erl & " - " & Err.Description)
    '</EhFooter>
End Sub

Sub MoveCharbyHead(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)
'*****************************************************************
'Starts the movement of a character in nHeading direction
'*****************************************************************
    Dim addX As Integer
    Dim addY As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim nX As Integer
    Dim nY As Integer
    
    With charlist(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        
        'Figure out which way to move
        Select Case nHeading
            Case E_Heading.NORTH
                addY = -1
        
            Case E_Heading.EAST
                addX = 1
        
            Case E_Heading.SOUTH
                addY = 1
            
            Case E_Heading.WEST
                addX = -1
        End Select
        
        nX = X + addX
        nY = Y + addY
        
        MapData(nX, nY).CharIndex = CharIndex
        .Pos.X = nX
        .Pos.Y = nY
        MapData(X, Y).CharIndex = 0
        
        .MoveOffsetX = -1 * (TilePixelWidth * addX)
        .MoveOffsetY = -1 * (TilePixelHeight * addY)
        
        .Moving = 1
        .Heading = nHeading
        
        .ScrollDirectionX = addX
        .ScrollDirectionY = addY
    End With
    
    If UserEstado = 0 Then Call DoPasosFx(CharIndex)
    
    'areas viejos
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        If CharIndex <> UserCharIndex Then
            Call EraseChar(CharIndex)
        End If
    End If
End Sub

Public Sub DoFogataFx()
    Dim location As Position
    
    If bFogata Then
        bFogata = HayFogata(location)
        If Not bFogata Then
            Call Audio.StopWave(FogataBufferIndex)
            FogataBufferIndex = 0
        End If
    Else
        bFogata = HayFogata(location)
        If bFogata And FogataBufferIndex = 0 Then FogataBufferIndex = Audio.PlayWave("fuego.wav", location.X, location.Y, LoopStyle.Enabled)
    End If
End Sub

Private Function EstaPCarea(ByVal CharIndex As Integer) As Boolean
    With charlist(CharIndex).Pos
        EstaPCarea = .X > ViewPositionX - MinXBorder And .X < ViewPositionX + MinXBorder And .Y > ViewPositionY - MinYBorder And .Y < ViewPositionY + MinYBorder
    End With
End Function

Sub DoPasosFx(ByVal CharIndex As Integer)
    If Not UserNavegando Then
        With charlist(CharIndex)
            If Not .muerto And EstaPCarea(CharIndex) And (.priv = 0 Or .priv > 5) Then
                .pie = Not .pie
                
                If .pie Then
                    Call Audio.PlayWave(SND_PASOS1, .Pos.X, .Pos.Y)
                Else
                    Call Audio.PlayWave(SND_PASOS2, .Pos.X, .Pos.Y)
                End If
            End If
        End With
    Else
' TODO : Actually we would have to check if the CharIndex char is in the water or not....
        Call Audio.PlayWave(SND_NAVEGANDO, charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y)
    End If
End Sub

Sub MoveCharbyPos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)
    Dim X As Integer
    Dim Y As Integer
    Dim addX As Integer
    Dim addY As Integer
    Dim nHeading As E_Heading
    
    With charlist(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        
        MapData(X, Y).CharIndex = 0
        
        addX = nX - X
        addY = nY - Y
        
        If Sgn(addX) = 1 Then
            nHeading = E_Heading.EAST
        ElseIf Sgn(addX) = -1 Then
            nHeading = E_Heading.WEST
        ElseIf Sgn(addY) = -1 Then
            nHeading = E_Heading.NORTH
        ElseIf Sgn(addY) = 1 Then
            nHeading = E_Heading.SOUTH
        End If
        
        MapData(nX, nY).CharIndex = CharIndex
        
        .Pos.X = nX
        .Pos.Y = nY
        
        .MoveOffsetX = -1 * (TilePixelWidth * addX)
        .MoveOffsetY = -1 * (TilePixelHeight * addY)
        
        .Moving = 1
        .Heading = nHeading
        
        .ScrollDirectionX = Sgn(addX)
        .ScrollDirectionY = Sgn(addY)
        
        'parche para que no medite cuando camina
        If .FxIndex = FxMeditar.CHICO Or .FxIndex = FxMeditar.GRANDE Or .FxIndex = FxMeditar.MEDIANO Or .FxIndex = FxMeditar.XGRANDE Or .FxIndex = FxMeditar.XXGRANDE Then
            .FxIndex = 0
        End If
    End With
    
    If Not EstaPCarea(CharIndex) Then Call Dialogos.RemoveDialog(CharIndex)
    
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        Call EraseChar(CharIndex)
    End If
End Sub

Sub MoveScreen(ByVal nHeading As E_Heading)
'******************************************
'Starts the screen moving in a direction
'******************************************
    Dim X As Integer
    Dim Y As Integer
    Dim TX As Integer
    Dim TY As Integer

    Dim addX As Integer
    Dim addY As Integer

    'Figure out which way to move
    Select Case nHeading
        Case E_Heading.NORTH
            addY = -1
        Case E_Heading.EAST
            addX = 1
        Case E_Heading.SOUTH
            addY = 1
        Case E_Heading.WEST
            addX = -1
    End Select

    X = ViewPositionX
    Y = ViewPositionY

    'Fill temp pos
    TX = X + addX
    TY = Y + addY

    'Check to see if its out of bounds
    If InMapBounds(TX, TY) Then

        ViewPositionX = TX
        ViewPositionY = TY

        OffsetCounterX = -(TilePixelWidth * addX)
        OffsetCounterY = -(TilePixelHeight * addY)

        ScrollDirectionX = addX
        ScrollDirectionY = addY

        UserMoving = 1 'is scrolling

        bTecho = IIf(MapData(ViewPositionX, ViewPositionY).Trigger = 1 Or _
                MapData(ViewPositionX, ViewPositionY).Trigger = 2 Or _
                MapData(ViewPositionX, ViewPositionY).Trigger = 4, True, False)
    Else
        ScrollDirectionX = 0
        ScrollDirectionY = 0
    End If
    
        
    'Call D3DXMatrixTranslation(Translation, addX * 32, addY * 32, 0)
    
    'Call D3DXMatrixMultiply(View, View, Translation)
    'Call DirectDevice.SetTransform(D3DTS_VIEW, View)
End Sub

Private Function HayFogata(ByRef location As Position) As Boolean
    Dim J As Long
    Dim k As Long
    
    For J = ViewPositionX - 8 To ViewPositionX + 8
        For k = ViewPositionY - 6 To ViewPositionY + 6
            If InMapBounds(J, k) Then
                If MapData(J, k).ObjGrh.GrhIndex = GrhFogata Then
                    location.X = J
                    location.Y = k
                    
                    HayFogata = True
                    Exit Function
                End If
            End If
        Next k
    Next J
End Function

Function NextOpenChar() As Integer
'*****************************************************************
'Finds next open char slot in CharList
'*****************************************************************
    Dim loopC As Long
    Dim Dale As Boolean
    
    loopC = 1
    Do While charlist(loopC).Active And Dale
        loopC = loopC + 1
        Dale = (loopC <= UBound(charlist))
    Loop
    
    NextOpenChar = loopC
End Function

''
' Loads grh data using the new file format.
'
' @return   True if the load was successfull, False otherwise.

Private Function LoadGrhData() As Boolean
On Error GoTo errorHandler
    Dim Grh As Long
    Dim Frame As Long
    Dim grhCount As Long
    Dim handle As Integer
    Dim fileVersion As Long
    
    'Open files
    handle = FreeFile()
    
    Open IniPath & "Graficos.ind" For Binary Access Read As handle
    Seek handle, 1
    
    'Get file version
    Get handle, , fileVersion
    
    'Get number of grhs
    Get handle, , grhCount
    
    'Resize arrays
    ReDim GrhData(1 To grhCount) As GrhData
    
    Get handle, , Grh
    
    Do Until Grh <= 0
        
        With GrhData(Grh)
            
            'Get number of frames
            Get handle, , .NumFrames
            If .NumFrames <= 0 Then GoTo errorHandler
            
            'ReDim .Frames(1 To GrhData(Grh).NumFrames)
            
            If .NumFrames > 1 Then
                'Read a animation GRH set
                For Frame = 1 To .NumFrames
                    Get handle, , .Frames(Frame)
                    If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then
                        GoTo errorHandler
                    End If
                Next Frame
                
                Get handle, , .Speed
                
                If .Speed <= 0 Then GoTo errorHandler
                
                'Compute width and height
                .pixelHeight = GrhData(.Frames(1)).pixelHeight
                If .pixelHeight <= 0 Then GoTo errorHandler
                
                .pixelWidth = GrhData(.Frames(1)).pixelWidth
                If .pixelWidth <= 0 Then GoTo errorHandler
                
                .TileWidth = GrhData(.Frames(1)).TileWidth
                If .TileWidth <= 0 Then GoTo errorHandler
                
                .TileHeight = GrhData(.Frames(1)).TileHeight
                If .TileHeight <= 0 Then GoTo errorHandler
            Else
                'Read in normal GRH data
                Get handle, , .FileNum
                If .FileNum <= 0 Then GoTo errorHandler
                
                Get handle, , GrhData(Grh).sX
                If .sX < 0 Then GoTo errorHandler
                
                Get handle, , .sY
                If .sY < 0 Then GoTo errorHandler
                
                Get handle, , .pixelWidth
                If .pixelWidth <= 0 Then GoTo errorHandler
                
                Get handle, , .pixelHeight
                If .pixelHeight <= 0 Then GoTo errorHandler
                
                'Compute width and height
                .TileWidth = .pixelWidth / TilePixelHeight
                .TileHeight = .pixelHeight / TilePixelWidth
                
                .Frames(1) = Grh
            End If
        End With
        
        Get handle, , Grh
    Loop
    
    Close handle
    
    LoadGrhData = True
Exit Function

errorHandler:
    LoadGrhData = False
End Function

Function LegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is legal
'*****************************************************************
    'Limites del mapa
    If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        Exit Function
    End If
    
    'Tile Bloqueado?
    If MapData(X, Y).Blocked = 1 Then
        Exit Function
    End If
    
    '¿Hay un personaje?
    If MapData(X, Y).CharIndex > 0 Then
        Exit Function
    End If
   
    If UserNavegando <> HayAgua(X, Y) Then
        Exit Function
    End If
    
    LegalPos = True
End Function

Function MoveToLegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Author: ZaMa
'Last Modify Date: 01/08/2009
'Checks to see if a tile position is legal, including if there is a casper in the tile
'10/05/2009: ZaMa - Now you can't change position with a casper which is in the shore.
'01/08/2009: ZaMa - Now invisible admins can't change position with caspers.
'*****************************************************************
    Dim CharIndex As Integer
    
    'Limites del mapa
    If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        Exit Function
    End If
    
    'Tile Bloqueado?
    If MapData(X, Y).Blocked = 1 Then
        Exit Function
    End If
    
    CharIndex = MapData(X, Y).CharIndex
    '¿Hay un personaje?
    If CharIndex > 0 Then
    
        If MapData(ViewPositionX, ViewPositionY).Blocked = 1 Then
            Exit Function
        End If
        
        With charlist(CharIndex)
            ' Si no es casper, no puede pasar
            If .iHead <> CASPER_HEAD And .iBody <> FRAGATA_FANTASMAL Then
                Exit Function
            Else
                ' No puedo intercambiar con un casper que este en la orilla (Lado tierra)
                If HayAgua(ViewPositionX, ViewPositionY) Then
                    If Not HayAgua(X, Y) Then Exit Function
                Else
                    ' No puedo intercambiar con un casper que este en la orilla (Lado agua)
                    If HayAgua(X, Y) Then Exit Function
                End If
                
                ' Los admins no pueden intercambiar pos con caspers cuando estan invisibles
                If charlist(UserCharIndex).priv > 0 And charlist(UserCharIndex).priv < 6 Then
                    If charlist(UserCharIndex).invisible = True Then Exit Function
                End If
            End If
        End With
    End If
   
    If UserNavegando <> HayAgua(X, Y) Then
        Exit Function
    End If
    
    MoveToLegalPos = True
End Function

Function InMapBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
        Exit Function
    End If
    
    InMapBounds = True
End Function

Sub Draw_GrhIndex(ByVal GrhIndex As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, Color() As Long)

    With GrhData(GrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
        
        'Draw
        Device_Textured_Render X, Y, .pixelWidth, .pixelHeight, .sX, .sY, .FileNum, Color
    End With
End Sub

Sub Draw_Grh(ByRef Grh As Grh, ByVal X As Single, ByVal Y As Single, ByVal Center As Byte, ByVal Animate As Byte, Color() As Long)
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
    Dim CurrentGrhIndex As Integer
    
'On Error GoTo error
    
    If Animate Then
        If Grh.Started Then
            Grh.FrameCounter = Grh.FrameCounter + animTicksPerFrame * Grh.Speed '/ 1000
            
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
        
        'Draw
        Device_Textured_Render X, Y, .pixelWidth, .pixelHeight, .sX, .sY, .FileNum, Color
    End With
Exit Sub

error:
    If Err.Number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        MsgBox "Ocurrió un error inesperado, por favor comuniquelo a los administradores del juego." & vbCrLf & "Descripción del error: " & _
        vbCrLf & Err.Description, vbExclamation, "[ " & Err.Number & " ] Error"
        End
    End If
End Sub

Sub Render_Screen()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 8/14/2007
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Renders everything to the viewport
'**************************************************************
    Dim Y           As Long     'Keeps track of where on map we are
    Dim X           As Long     'Keeps track of where on map we are
    Dim ScreenMinY  As Integer  'Start Y pos on current screen
    Dim ScreenMaxY  As Integer  'End Y pos on current screen
    Dim ScreenMinX  As Integer  'Start X pos on current screen
    Dim ScreenMaxX  As Integer  'End X pos on current screen
    Dim minY        As Integer  'Start Y pos on current map
    Dim maxY        As Integer  'End Y pos on current map
    Dim minX        As Integer  'Start X pos on current map
    Dim maxX        As Integer  'End X pos on current map
    Dim ScreenX     As Integer  'Keeps track of where to place tile on screen
    Dim ScreenY     As Integer  'Keeps track of where to place tile on screen
    Dim minXOffset  As Integer
    Dim minYOffset  As Integer
    Dim PixelOffsetXTemp As Single 'For centering grhs
    Dim PixelOffsetYTemp As Single 'For centering grhs
    
    If UserMoving = 1 Then
        If ScrollDirectionX <> 0 Then
            OffsetCounterX = OffsetCounterX + ScrollPixelsPerFrameX * timerTicksPerFrame * ScrollDirectionX
            If Sgn(OffsetCounterX) = ScrollDirectionX Then
                OffsetCounterX = 0
                ScrollDirectionX = 0
            End If
        End If

        If ScrollDirectionY <> 0 Then
            OffsetCounterY = OffsetCounterY + ScrollPixelsPerFrameY * timerTicksPerFrame * ScrollDirectionY
            If Sgn(OffsetCounterY) = ScrollDirectionY Then
                OffsetCounterY = 0
                ScrollDirectionY = 0
            End If
        End If
        
        If ScrollDirectionX = 0 And ScrollDirectionY = 0 Then
            UserMoving = 0
        End If
    End If
    
    'Figure out Ends and Starts of screen
    ScreenMinY = ViewPositionY - HalfWindowTileHeight
    ScreenMaxY = ViewPositionY + HalfWindowTileHeight
    ScreenMinX = ViewPositionX - HalfWindowTileWidth
    ScreenMaxX = ViewPositionX + HalfWindowTileWidth
    
    minY = ScreenMinY - TileBufferSize
    maxY = ScreenMaxY + TileBufferSize
    minX = ScreenMinX - TileBufferSize
    maxX = ScreenMaxX + TileBufferSize
   
    'Make sure mins and maxs are always in map bounds
    If minY < XMinMapSize Then
        minYOffset = YMinMapSize - minY
        minY = YMinMapSize
    End If
    
    If maxY > YMaxMapSize Then maxY = YMaxMapSize
    
    If minX < XMinMapSize Then
        minXOffset = XMinMapSize - minX
        minX = XMinMapSize
    End If
    
    If maxX > XMaxMapSize Then maxX = XMaxMapSize
    
    'If we can, we render around the view area to make it smoother
    If ScreenMinY > YMinMapSize Then
        ScreenMinY = ScreenMinY - 1
    Else
        ScreenMinY = 1
        ScreenY = 1
    End If
    
    If ScreenMaxY < YMaxMapSize Then ScreenMaxY = ScreenMaxY + 1
    
    If ScreenMinX > XMinMapSize Then
        ScreenMinX = ScreenMinX - 1
    Else
        ScreenMinX = 1
        ScreenX = 1
    End If
    
    If ScreenMaxX < XMaxMapSize Then ScreenMaxX = ScreenMaxX + 1

    ScreenX = ScreenX - 1
    ScreenY = ScreenY - 1
    
    PixelOffsetYTemp = ScreenY * TilePixelHeight - OffsetCounterY
    'Draw floor layer
    For Y = ScreenMinY To ScreenMaxY
        For X = ScreenMinX To ScreenMaxX
            
            PixelOffsetXTemp = ScreenX * TilePixelWidth - OffsetCounterX
            
            'Layer 1 **********************************
            Call Draw_Grh(MapData(X, Y).Graphic(1), _
                PixelOffsetXTemp, _
                PixelOffsetYTemp, _
                0, 1, AmbientColor)
            '******************************************

            'Layer 2 **********************************
            If MapData(X, Y).Graphic(2).GrhIndex <> 0 Then
                Call Draw_Grh(MapData(X, Y).Graphic(2), _
                        PixelOffsetXTemp, _
                        PixelOffsetYTemp, _
                        1, 1, AmbientColor)
            End If
            '******************************************
            
            ScreenX = ScreenX + 1
        Next
    
        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - X + ScreenMinX
        ScreenY = ScreenY + 1
        PixelOffsetYTemp = ScreenY * TilePixelHeight - OffsetCounterY
    Next
    
    'Draw Transparent Layers
    ScreenY = (minYOffset - TileBufferSize)
    ScreenX = (minXOffset - TileBufferSize)
   
    PixelOffsetYTemp = ScreenY * TilePixelHeight - OffsetCounterY
   
    For Y = minY To maxY
        For X = minX To maxX
        
            PixelOffsetXTemp = ScreenX * TilePixelWidth - OffsetCounterX
    
            With MapData(X, Y)
                'Object Layer **********************************
                If .ObjGrh.GrhIndex <> 0 Then
                    Call Draw_Grh(.ObjGrh, _
                            PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, AmbientColor)
                End If
                '***********************************************
    
    
                'Char layer ************************************
                If .CharIndex <> 0 Then
                    Call CharRender(.CharIndex, PixelOffsetXTemp, PixelOffsetYTemp)
                End If
                '*************************************************
    
                
                'Layer 3 *****************************************
                If .Graphic(3).GrhIndex <> 0 Then
                    'Draw
                    Call Draw_Grh(.Graphic(3), _
                            PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, AmbientColor)
                End If
                '************************************************
                
            End With
    
            ScreenX = ScreenX + 1
        Next
        
        ScreenX = ScreenX - X + minX
        ScreenY = ScreenY + 1
        PixelOffsetYTemp = ScreenY * TilePixelHeight - OffsetCounterY
    Next
    
    If Not bTecho Then
        
        ScreenY = (minYOffset - TileBufferSize)
        ScreenX = (minXOffset - TileBufferSize)
        
        PixelOffsetYTemp = ScreenY * TilePixelHeight - OffsetCounterY
        
        For Y = minY To maxY
            For X = minX To maxX
        
               PixelOffsetXTemp = ScreenX * TilePixelWidth - OffsetCounterX
               
                'Layer 4 **********************************
                If MapData(X, Y).Graphic(4).GrhIndex Then
                    'Draw
                    Call Draw_Grh(MapData(X, Y).Graphic(4), _
                        PixelOffsetXTemp, _
                        PixelOffsetYTemp, _
                        1, 1, AmbientColor)
                End If
                '**********************************
               
                ScreenX = ScreenX + 1
            Next
            ScreenX = ScreenX - X + minX
            ScreenY = ScreenY + 1
            PixelOffsetYTemp = ScreenY * TilePixelHeight - OffsetCounterY
        
        Next
    End If
    
End Sub

Public Function RenderSounds()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/30/2008
'Actualiza todos los sonidos del mapa.
'**************************************************************
    If bLluvia(UserMap) = 1 Then
        If bRain Then
            If bTecho Then
                If frmMain.IsPlaying <> PlayLoop.plLluviain Then
                    If RainBufferIndex Then _
                        Call Audio.StopWave(RainBufferIndex)
                    RainBufferIndex = Audio.PlayWave("lluviain.wav", 0, 0, LoopStyle.Enabled)
                    frmMain.IsPlaying = PlayLoop.plLluviain
                End If
            Else
                If frmMain.IsPlaying <> PlayLoop.plLluviaout Then
                    If RainBufferIndex Then _
                        Call Audio.StopWave(RainBufferIndex)
                    RainBufferIndex = Audio.PlayWave("lluviaout.wav", 0, 0, LoopStyle.Enabled)
                    frmMain.IsPlaying = PlayLoop.plLluviaout
                End If
            End If
        End If
    End If
    
    DoFogataFx
End Function

Function HayUserAbajo(ByVal X As Integer, ByVal Y As Integer, ByVal GrhIndex As Integer) As Boolean
    If GrhIndex > 0 Then
        HayUserAbajo = _
            charlist(UserCharIndex).Pos.X >= X - (GrhData(GrhIndex).TileWidth \ 2) _
                And charlist(UserCharIndex).Pos.X <= X + (GrhData(GrhIndex).TileWidth \ 2) _
                And charlist(UserCharIndex).Pos.Y >= Y - (GrhData(GrhIndex).TileHeight - 1) _
                And charlist(UserCharIndex).Pos.Y <= Y
    End If
End Function

Public Function InitTileEngine(ByVal setDisplayFormhWnd As Long, ByVal setTilePixelHeight As Integer, ByVal setTilePixelWidth As Integer, ByVal setWindowTileHeight As Integer, ByVal setWindowTileWidth As Integer, ByVal setTileBufferSize As Integer, ByVal pixelsToScrollPerFrameX As Integer, pixelsToScrollPerFrameY As Integer, ByVal engineSpeed As Single, ByVal animSpeed As Single) As Boolean
'***************************************************
'Author: Aaron Perkins
'Last Modification: 08/14/07
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Creates all DX objects and configures the engine to start running.
'***************************************************
    'Fill startup variables
    TilePixelWidth = setTilePixelWidth
    TilePixelHeight = setTilePixelHeight
    WindowTileHeight = setWindowTileHeight
    WindowTileWidth = setWindowTileWidth
    TileBufferSize = setTileBufferSize
    
    HalfWindowTileHeight = setWindowTileHeight \ 2
    HalfWindowTileWidth = setWindowTileWidth \ 2
    
    'Compute offset in pixels when rendering tile buffer.
    'We diminish by one to get the top-left corner of the tile for rendering.
    'TileBufferPixelOffsetX = ((TileBufferSize - 1) * TilePixelWidth)
    'TileBufferPixelOffsetY = ((TileBufferSize - 1) * TilePixelHeight)
    
    engineBaseSpeed = engineSpeed
    engineAnimSpeed = animSpeed
    
    'Set FPS value to 60 for startup
    FPS = 60
    FramesPerSecCounter = 60
    
    MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
    MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
    MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
    MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)
    
    MainViewWidth = TilePixelWidth * WindowTileWidth
    MainViewHeight = TilePixelHeight * WindowTileHeight
    
    'Resize mapdata array
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
    'Set intial user position
    'viewpositionx = MinXBorder
    'viewpositiony = MinYBorder
    
    ViewPositionX = MinXBorder
    ViewPositionY = MinYBorder
    
    'Set scroll pixels per frame
    ScrollPixelsPerFrameX = pixelsToScrollPerFrameX
    ScrollPixelsPerFrameY = pixelsToScrollPerFrameY
    
    Call DirectX_Init
    
    Call LoadGrhData
    Call CargarCuerpos
    Call CargarCabezas
    Call CargarCascos
    Call CargarFxs
        
    AmbientColor(0) = -1
    AmbientColor(1) = -1
    AmbientColor(2) = -1
    AmbientColor(3) = -1
    
    RenderState = eRenderState.eLogin
    FadingRState = eRenderState.eLogin
    
    FadeOn = True
    ConnectAlpha = 0
    InitTileEngine = True
End Function

Private Sub DirectX_Init()

    Dim DispMode As D3DDISPLAYMODE
    Dim PresentationParameters As D3DPRESENT_PARAMETERS
    
    Set DirectX = New DirectX8
    Set DirectD3D = DirectX.Direct3DCreate
    Set DirectD3D8 = New D3DX8
        
    DirectD3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    
    With PresentationParameters
    
        .Windowed = True
        .SwapEffect = D3DSWAPEFFECT_DISCARD
        
        .BackBufferWidth = 1024
        .BackBufferHeight = 768
        .BackBufferFormat = DispMode.Format
        
    End With

    Set DirectDevice = DirectD3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.MainViewPic.hwnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, PresentationParameters)
    
    Call D3DXMatrixOrthoOffCenterLH(Projection, 0, 1024, 768, 0, -1#, 1#)
    Call D3DXMatrixIdentity(View)
    
    'default camera values
    
    'Call D3DXMatrixTranslation(Translation, 0, 0, 0)
    
    'Call D3DXMatrixMultiply(View, View, Translation)
    
    Call DirectDevice.SetTransform(D3DTS_PROJECTION, Projection)
    Call DirectDevice.SetTransform(D3DTS_VIEW, View)
    
    Engine_Init_RenderStates
    
    Set SurfaceDB = New clsSurfaceDB
    Set SpriteBatch = New clsBatch
    
    Call SurfaceDB.Initialize(DirectD3D8, DirGraficos, 90)
    Call SpriteBatch.Initialise(2000)
    
    Call Engine_Init_FontSettings
    Call Engine_Init_FontTextures
    
    Call InitComponents
    
    Call Input_Init
    
    With MainViewRect
        .X2 = frmMain.MainViewPic.ScaleWidth
        .Y2 = frmMain.MainViewPic.ScaleHeight
    End With
    
    With ConnectRect
    
        .X2 = frmConnect.Render.ScaleWidth
        .Y2 = frmConnect.Render.ScaleHeight
        
    End With
    
    If DirectDevice Is Nothing Then
        MsgBox "No se puede inicializar DirectX. Por favor asegúrese de tener la última versión correctamente instalada."
        Exit Sub
    End If
        
End Sub

Public Sub DirectX_EndScene(ByRef Rect As D3DRECT, ByVal hwnd As Long)
    
    Call SpriteBatch.Flush
    
    Call DirectDevice.EndScene
    Call DirectDevice.Present(Rect, ByVal 0, hwnd, ByVal 0)
End Sub

Private Sub Engine_Init_RenderStates()

    DirectDevice.SetVertexShader (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)
        
    'Set the render states
    With DirectDevice
    
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        .SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        .SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
        
    End With
    
End Sub

Public Sub DeinitTileEngine()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 08/14/07
    'Destroys all DX objects
    '***************************************************

    'Set no texture in the device to avoid memory leaks
    If Not DirectDevice Is Nothing Then
        DirectDevice.SetTexture 0, Nothing
    End If
        
    '// Destroy Textures
    Set SurfaceDB = Nothing
        
    Call Input_Release
        
    Set DirectX = Nothing
    Set DirectD3D = Nothing
    Set DirectDevice = Nothing
    Set DirectD3D8 = Nothing
        
    Set SpriteBatch = Nothing
        
    'Clear arrays
    Erase GrhData
    Erase BodyData
    Erase HeadData
    Erase FxData
    Erase WeaponAnimData
    Erase ShieldAnimData
    Erase CascoAnimData
    Erase MapData
    Erase charlist
        
End Sub

Public Sub Render()
    If EngineRun Then
            
        Call DirectDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)

        Call SpriteBatch.Begin
            
        If frmConnect.Visible Then
            Call Render_Connect
        Else
            'Sólo dibujamos si la ventana no está minimizada
            If frmMain.WindowState <> 1 And frmMain.Visible Then
                
                Call ShowNextFrame(frmMain.MouseX, frmMain.MouseY)
                
                'Play ambient sounds
                Call RenderSounds
                    
                Call CheckKeys
                    
            End If
        End If
        
        'FPS update
        If fpsLastCheck + 1000 < GetTickCount() Then
            FPS = FramesPerSecCounter
            FramesPerSecCounter = 1
            fpsLastCheck = GetTickCount()
        Else
            FramesPerSecCounter = FramesPerSecCounter + 1
        End If
        
        Call SpriteBatch.Finish
    
    End If
    
    Call Time_Update
    Call GameTime_Update
    Call AnimTime_Update
End Sub

Private Sub Render_Connect()
    
    Dim Color(3) As Long
    
    If FadeOff Then
        If ConnectAlpha > 20 Then
            ConnectAlpha = ConnectAlpha - (30 * timerTicksPerFrame)
        Else
            ConnectAlpha = 0
            FadeOff = False
        End If
        
        If ConnectAlpha < 0 Then ConnectAlpha = 0
    End If
    
    If FadeOn Then
        If ConnectAlpha < 200 Then
            ConnectAlpha = ConnectAlpha + (30 * timerTicksPerFrame)
        Else
            ConnectAlpha = 255
            FadeOn = False
        End If
        
        If ConnectAlpha > 255 Then ConnectAlpha = 255
    End If
    
    
    Color(0) = D3DColorARGB(ConnectAlpha, 255, 255, 255)
    Color(1) = Color(0)
    Color(2) = Color(0)
    Color(3) = Color(0)
    
    Call DirectDevice.BeginScene
    
    Call Render_Connect_Background
    
    Call Device_Textured_Render(0, 0, 1024, 768, 0, 0, 999999, White)
    
     Select Case FadingRState
        
        Case eRenderState.eLogin
            Call Device_Textured_Render(384, 306, 256, 205, 0, 0, 1000000, Color)

        Case eRenderState.eNewCharInfo
            Call Device_Textured_Render(384, 189, 256, 512, 256, 0, 1000001, Color)
        
        Case eRenderState.eNewCharDetails
            Call Device_Textured_Render(384, 189, 256, 512, 0, 0, 1000001, Color)
            Call Draw_Head_Selector
            
        Case eRenderState.eNewCharAttrib
            Call Device_Textured_Render(384, 189, 256, 512, 0, 0, 1000002, Color)
        
        Case eRenderState.eNewCharSkills
            Call Device_Textured_Render(384, 189, 256, 512, 256, 0, 1000002, Color)
            
            
    End Select
    
    If (Not FadeOff And ConnectAlpha = 0) Then
        'FadeOff = False
        FadingRState = RenderState
        FadeOn = True
    End If

    Call Render_Help_Window
    
    'todo: apply the correct alpha to each component
    If ConnectAlpha = 255 Then _
        Call RenderComponents(255)
    
    Call Text_Draw(5, 5, FPS, White)
    Call DirectX_EndScene(ConnectRect, frmConnect.Render.hwnd)
End Sub

Private Sub Render_Connect_Background()

Dim X As Long, Y As Long
Dim ScreenX As Long, ScreenY As Long

For X = 34 To 64
    
    For Y = 38 To 62
        
        With MapData(X, Y)
            Call Draw_Grh(.Graphic(1), 16 + ScreenX * 32, 15 + ScreenY * 32, 0, 1, White)
            
            
            If .Graphic(2).GrhIndex <> 0 Then _
                Call Draw_Grh(.Graphic(2), 16 + ScreenX * 32, 15 + ScreenY * 32, 0, 0, White)

        End With
        
        ScreenY = ScreenY + 1
    Next
    
    ScreenY = 0
    ScreenX = ScreenX + 1
    
Next

ScreenY = 0
ScreenX = 0

For X = 34 To 64
    
    For Y = 38 To 62
        
        With MapData(X, Y)
        
            If .Graphic(3).GrhIndex <> 0 Then _
                Call Draw_Grh(.Graphic(3), 16 + ScreenX * 32, 15 + ScreenY * 32, 1, 0, White)
                
        End With
        
        ScreenY = ScreenY + 1
    Next
    
    ScreenY = 0
    ScreenX = ScreenX + 1
    
Next
End Sub

Private Sub Render_Help_Window()
    
    If HelpWindow.Active = False Then Exit Sub
    
    Call Device_Textured_Render(640, 306, 256, 205, 260, 0, 1000000, White)
    
    Dim i As Long
    Dim yOffset As Integer
    
    With HelpWindow
    
        For i = 0 To UBound(.Text)
    
            Call Text_Draw(653, 320 + yOffset, .Text(i), White)
            
            yOffset = yOffset + 15
        Next
        
    End With
    
End Sub

Private Sub Draw_Head_Selector()
    
    Dim Raza As Integer, Sexo As Integer
    Raza = GetSelectedIndex(frmConnect.cmbRaza)
    Sexo = GetSelectedIndex(frmConnect.cmbSexo)
    
    Dim X As Integer
    X = 451
    
    Dim i As Long
    If Raza <> 0 And Sexo <> 0 Then
            
        For i = 0 To UBound(HeadSlider)
            Draw_GrhIndex HeadData(HeadSlider(i)).Head(3).GrhIndex, X, 439, 0, White
            
            X = X + 27
        Next
    End If
    
    If BodyExample.GrhIndex <> 0 Then
        Call Draw_GrhIndex(HeadData(UserHead).Head(3).GrhIndex, 499 + BodyData(UserBody).HeadOffset.X, 486 + BodyData(UserBody).HeadOffset.Y, 1, White)
        Call Draw_Grh(BodyExample, 499, 486, 1, 1, White)
    End If
End Sub

'CSEH: ErrLog
Private Sub ShowNextFrame(ByVal MouseViewX As Integer, ByVal MouseViewY As Integer)
'***************************************************
'Author: Arron Perkins
'Last Modification: 08/14/07
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Updates the game's model and renders everything.
'***************************************************

    'Update mouse position within view area
    Call ConvertCPtoTP(MouseViewX, MouseViewY, MouseTileX, MouseTileY)
     
    Call DirectDevice.BeginScene

     '****** Update screen ******
    If UserCiego Then
        Call CleanViewPort
    Else
        Call Render_Screen
    End If
    
    Call Dialogos.Render
    Call DibujarCartel
    
    ' Call DialogosClanes.Draw
    Call DirectX_EndScene(MainViewRect, 0)
    
    If GetTickCount - LastInvRender > 56 Then
     
        LastInvRender = GetTickCount
     
        Call Inventario.DrawInventory
    End If
          
End Sub

Private Function GetElapsedTime() As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the time that past since the last call
'**************************************************************
    Dim start_time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq
    End If
    
    'Get current time
    Call QueryPerformanceCounter(start_time)
    
    'Calculate elapsed time
    GetElapsedTime = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)
End Function

Private Sub CharRender(ByVal CharIndex As Long, ByVal OffsetCounterX As Integer, ByVal OffsetCounterY As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Draw char's to screen without offcentering them
'***************************************************
    Dim Moved As Boolean
    Dim Pos As Integer
    Dim line As String
    Dim Color(3) As Long
    
With charlist(CharIndex)
        If .Moving Then
            'If needed, move left and right
            If .ScrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.ScrollDirectionX) * timerTicksPerFrame
                
                'Start animations
'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then _
                    .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                Moved = True
                
                'Check if we already got there
                If (Sgn(.ScrollDirectionX) = 1 And .MoveOffsetX >= 0) Or (Sgn(.ScrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .ScrollDirectionX = 0
                End If
            End If
            
            'If needed, move up and down
            If .ScrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.ScrollDirectionY) * timerTicksPerFrame
                
                'Start animations
'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then _
                    .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                Moved = True
                
                'Check if we already got there
                If (Sgn(.ScrollDirectionY) = 1 And .MoveOffsetY >= 0) Or (Sgn(.ScrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .ScrollDirectionY = 0
                End If
            End If
        End If
        
        'If done moving stop animation
        If Not Moved Then
            'Stop animations
            .Body.Walk(.Heading).Started = 0
            .Body.Walk(.Heading).FrameCounter = 1
            
            .Arma.WeaponWalk(.Heading).Started = 0
            .Arma.WeaponWalk(.Heading).FrameCounter = 1
            
            .Escudo.ShieldWalk(.Heading).Started = 0
            .Escudo.ShieldWalk(.Heading).FrameCounter = 1
            
            .Moving = False
        End If
        
        OffsetCounterX = OffsetCounterX + .MoveOffsetX
        OffsetCounterY = OffsetCounterY + .MoveOffsetY
        
        If .Head.Head(.Heading).GrhIndex Then
            If Not .invisible Then
                'Draw Body
                If .Body.Walk(.Heading).GrhIndex Then Call Draw_Grh(.Body.Walk(.Heading), OffsetCounterX, OffsetCounterY, 1, 1, AmbientColor)
            
                'Draw Head
                If .Head.Head(.Heading).GrhIndex Then Call Draw_Grh(.Head.Head(.Heading), OffsetCounterX + .Body.HeadOffset.X, OffsetCounterY + .Body.HeadOffset.Y, 1, 0, AmbientColor)
                    
                'Draw Helmet
                If .Casco.Head(.Heading).GrhIndex Then Call Draw_Grh(.Casco.Head(.Heading), OffsetCounterX + .Body.HeadOffset.X, OffsetCounterY + .Body.HeadOffset.Y, 1, 0, AmbientColor)
                    ' Call Draw_Grh(.Casco.Head(.Heading), OffsetCounterX + .Body.HeadOffset.X, Offsetcountery + .Body.HeadOffset.Y + OFFSET_HEAD, 1, 0)
                
                'Draw Weapon
                If .Arma.WeaponWalk(.Heading).GrhIndex Then Call Draw_Grh(.Arma.WeaponWalk(.Heading), OffsetCounterX, OffsetCounterY, 1, 1, AmbientColor)
                    
                'Draw Shield
                If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call Draw_Grh(.Escudo.ShieldWalk(.Heading), OffsetCounterX, OffsetCounterY, 1, 1, AmbientColor)
                
                If LenB(.Nombre) > 0 Then
                    If Nombres Then
                        Pos = getTagPosition(.Nombre)
                            'Pos = InStr(.Nombre, "<")
                            'If Pos = 0 Then Pos = Len(.Nombre) + 2

                        If .priv = 0 Then
                            Select Case .Criminal
                            
                                Case 2 'ciuda
                                    Color(0) = ColoresPJ(49)
                                    Color(1) = ColoresPJ(49)
                                    Color(2) = ColoresPJ(49)
                                    Color(3) = ColoresPJ(49)
                                Case 3
                                    Color(0) = ColoresPJ(50)
                                    Color(1) = ColoresPJ(50)
                                    Color(2) = ColoresPJ(50)
                                    Color(3) = ColoresPJ(50)
                                Case 4
                                    Color(0) = ColoresPJ(47)
                                    Color(1) = ColoresPJ(47)
                                    Color(2) = ColoresPJ(47)
                                    Color(3) = ColoresPJ(47)
                                Case 5
                                    Color(0) = ColoresPJ(48)
                                    Color(1) = ColoresPJ(48)
                                    Color(2) = ColoresPJ(48)
                                    Color(3) = ColoresPJ(48)
                            End Select
                        Else
                            Color(0) = ColoresPJ(.priv)
                            Color(1) = ColoresPJ(.priv)
                            Color(2) = ColoresPJ(.priv)
                            Color(3) = ColoresPJ(.priv)
                        End If
                            
                        'Nick
                        Call Text_Draw(OffsetCounterX - .NombreOffset, OffsetCounterY + 30, .Nombre, Color)
                               
                        'Guild
                        If LenB(.GuildName) > 0 Then _
                            Call Text_Draw(OffsetCounterX - .GuildOffset, OffsetCounterY + 45, .GuildName, Color)
                           
                    End If
                End If
            End If
        Else
            'Draw Body
            If .Body.Walk(.Heading).GrhIndex Then _
                Call Draw_Grh(.Body.Walk(.Heading), OffsetCounterX, OffsetCounterY, 1, 1, AmbientColor)
        End If
        
        'Update dialogs
        'Call Dialogos.UpdateDialogPos(OffsetCounterX + .Body.HeadOffset.X, Offsetcountery + .Body.HeadOffset.Y + OFFSET_HEAD, CharIndex) '34 son los pixeles del grh de la cabeza que quedan superpuestos al cuerpo
        Call Dialogos.UpdateDialogPos(OffsetCounterX + .Body.HeadOffset.X, OffsetCounterY + .Body.HeadOffset.Y, CharIndex)
        
        'Draw FX
        If .FxIndex <> 0 Then
            
            Call Draw_Grh(.fX, OffsetCounterX + FxData(.FxIndex).OffsetX, OffsetCounterY + FxData(.FxIndex).OffsetY, 1, 1, AmbientColor)
            'Check if animation is over
            If .fX.Started = 0 Then .FxIndex = 0
        End If
        
End With

End Sub

Public Sub SetCharacterFx(ByVal CharIndex As Integer, ByVal fX As Integer, ByVal Loops As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Sets an FX to the character.
'***************************************************
    With charlist(CharIndex)
        .FxIndex = fX
        
        If .FxIndex > 0 Then
            Call InitGrh(.fX, FxData(fX).Animacion)
        
            .fX.Loops = Loops
        End If
    End With
End Sub

Private Sub CleanViewPort()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Fills the viewport with black.
'***************************************************
    'todo: check if inventory still rendering after this
    DirectDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 0, 1#, 0
End Sub

Public Sub Device_Textured_Render(ByVal X As Single, ByVal Y As Single, _
                                  ByVal Width As Integer, ByVal Height As Integer, _
                                  ByVal sX As Integer, ByVal sY As Integer, _
                                  ByVal tex As Long, _
                                  ByRef Color() As Long)

        Dim Texture As Direct3DTexture8
        Dim TexWidth As Integer, TexHeight As Integer
        
        Set Texture = SurfaceDB.Surface(tex, TexWidth, TexHeight)
        
        With SpriteBatch
                '// Seteamos la textura
                Call .SetTexture(Texture)
                
                If TexWidth <> 0 And TexHeight <> 0 Then
                    Call .Draw(X, Y, Width, Height, Color, sX / TexWidth, sY / TexHeight, (sX + Width) / TexWidth, (sY + Height) / TexHeight)
                Else
                    Call .Draw(X, Y, TexWidth, TexHeight, Color)
                End If
        End With
End Sub

Public Sub CheckKeys()
'*****************************************************************
'Checks keys and respond
'*****************************************************************
    'No input allowed while Argentum is not the active window
    'If Not Application.IsAppActive() Then Exit Sub
    
    'No walking when in commerce or banking.
    If Comerciando Then Exit Sub
    
    'No walking while writting in the forum.
    If MirandoForo Then Exit Sub
    
    'If game is paused, abort movement.
    If pausa Then Exit Sub
    
    'TODO: Debería informarle por consola?
    If Traveling Then Exit Sub

    'Control movement interval (this enforces the 1 step loss when meditating / resting client-side)
    'If GetTickCount - LastMovement > 56 Then
    '    LastMovement = GetTickCount
    'Else
    '    Exit Sub
    'End If
    
    'Don't allow any these keys during movement..
    If UserMoving = 0 Then
        If Not UserEstupido Then
            'Move Up
            If Input_Key_Get(CustomKeys.BindedKey(eKeyType.mKeyUp)) Then
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                Call MoveTo(NORTH)
                'todo
                frmMain.Coord.Caption = UserMap & " X: " & ViewPositionX & " Y: " & ViewPositionY
                Exit Sub
            End If
            
            'Move Right
            If Input_Key_Get(CustomKeys.BindedKey(eKeyType.mKeyRight)) Then
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                Call MoveTo(EAST)
                frmMain.Coord.Caption = "(" & UserMap & "," & ViewPositionX & "," & ViewPositionY & ")"
                frmMain.Coord.Caption = UserMap & " X: " & ViewPositionX & " Y: " & ViewPositionY
                Exit Sub
            End If
        
            'Move down
            If Input_Key_Get(CustomKeys.BindedKey(eKeyType.mKeyDown)) Then
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                Call MoveTo(SOUTH)
                frmMain.Coord.Caption = UserMap & " X: " & ViewPositionX & " Y: " & ViewPositionY
                Exit Sub
            End If
        
            'Move left
            If Input_Key_Get(CustomKeys.BindedKey(eKeyType.mKeyLeft)) Then
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                Call MoveTo(WEST)
                frmMain.Coord.Caption = UserMap & " X: " & ViewPositionX & " Y: " & ViewPositionY
                Exit Sub
            End If
            
            ' We haven't moved - Update 3D sounds!
            Call Audio.MoveListener(ViewPositionX, ViewPositionY)
        Else
            Dim kp As Boolean
            kp = (Input_Key_Get(CustomKeys.BindedKey(eKeyType.mKeyUp))) Or _
                Input_Key_Get(CustomKeys.BindedKey(eKeyType.mKeyRight)) Or _
                Input_Key_Get(CustomKeys.BindedKey(eKeyType.mKeyDown)) Or _
                Input_Key_Get(CustomKeys.BindedKey(eKeyType.mKeyLeft))
            
            If kp Then
                Call RandomMove
            Else
                ' We haven't moved - Update 3D sounds!
                Call Audio.MoveListener(ViewPositionX, ViewPositionY)
            End If
            
            If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
            frmMain.Coord.Caption = "(" & ViewPositionX & "," & ViewPositionY & ")"
            frmMain.Coord.Caption = "X: " & ViewPositionX & " Y: " & ViewPositionY
        End If
    End If
End Sub

Public Sub Text_Draw(ByVal Left As Long, ByVal Top As Long, ByVal Text As String, Color() As Long, Optional ByVal Center As Boolean = False)

    Engine_Render_Text SpriteBatch, cfonts(1), Text, Left, Top, Color

End Sub

Private Sub Engine_Render_Text(ByRef Batch As clsBatch, ByRef UseFont As CustomFont, ByVal Text As String, ByVal X As Long, ByVal Y As Long, Color() As Long)
'*****************************************************************
'Render text with a custom font
'*****************************************************************
    Dim TempVA As CharVA
    Dim tempstr() As String
    Dim Count As Integer
    Dim ascii() As Byte
    Dim i As Long
    Dim J As Long

    Dim yOffset As Single
    
    'Check if we have the device
    If DirectDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub

    'Check for valid text to render
    If LenB(Text) = 0 Then Exit Sub
    
    'Get the text into arrays (split by vbCrLf)
    tempstr = Split(Text, vbCrLf)

    'Set the texture
    Batch.SetTexture UseFont.Texture
    
    'Loop through each line if there are line breaks (vbCrLf)
    For i = 0 To UBound(tempstr)
        If Len(tempstr(i)) > 0 Then
            yOffset = i * UseFont.CharHeight
            Count = 0
        
            'Convert the characters to the ascii value
            ascii() = StrConv(tempstr(i), vbFromUnicode)
        
            'Loop through the characters
            For J = 1 To Len(tempstr(i))

                CopyMemory TempVA, UseFont.HeaderInfo.CharVA(ascii(J - 1)), 24 'this number represents the size of "CharVA" struct
                
                TempVA.X = X + Count
                TempVA.Y = Y + yOffset
            
                Batch.Draw TempVA.X, TempVA.Y, TempVA.W, TempVA.H, Color, _
                            TempVA.Tx1, TempVA.Ty1, TempVA.Tx2, TempVA.Ty2

                'Shift over the the position to render the next character
                Count = Count + UseFont.HeaderInfo.CharWidth(ascii(J - 1))
                
            Next J
            
        End If
    Next i

End Sub

Public Function Text_GetWidth(ByRef UseFont As CustomFont, ByVal Text As String) As Integer
'***************************************************
'Returns the width of text
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_GetTextWidth
'***************************************************
Dim i As Integer

    'Make sure we have text
    If LenB(Text) = 0 Then Exit Function
    
    'Loop through the text
    For i = 1 To Len(Text)
        
        'Add up the stored character widths
        Text_GetWidth = Text_GetWidth + UseFont.HeaderInfo.CharWidth(Asc(mid$(Text, i, 1)))
        
    Next i

End Function

Sub Engine_Init_FontTextures()
'*****************************************************************
'Init the custom font textures
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_FontTextures
'*****************************************************************
    Dim TexInfo As D3DXIMAGE_INFO_A

    'Check if we have the device
    If DirectDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub

    '*** Default font ***
    
    'Set the texture
    Set cfonts(1).Texture = DirectD3D8.CreateTextureFromFileEx(DirectDevice, DirGraficos & "Font.png", _
                                                                D3DX_DEFAULT, D3DX_DEFAULT, 0, 0, _
                                                                D3DFMT_UNKNOWN, D3DPOOL_MANAGED, _
                                                                D3DX_FILTER_NONE, D3DX_FILTER_NONE, 0, TexInfo, ByVal 0)
    
    'Store the size of the texture
    cfonts(1).TextureSize.X = TexInfo.Width
    cfonts(1).TextureSize.Y = TexInfo.Height
    
End Sub

Sub Engine_Init_FontSettings()
    '*****************************************************************
    'Init the custom font settings
    'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_FontSettings
    '*****************************************************************
    Dim FileNum  As Byte
    Dim LoopChar As Long
    Dim Row      As Single
    Dim u        As Single
    Dim v        As Single

    '*** Default font ***

    'Load the header information
    FileNum = FreeFile
    Open IniPath & "Font.dat" For Binary As #FileNum
    Get #FileNum, , cfonts(1).HeaderInfo
    Close #FileNum
    
    'Calculate some common values
    cfonts(1).CharHeight = cfonts(1).HeaderInfo.CellHeight - 4
    cfonts(1).RowPitch = cfonts(1).HeaderInfo.BitmapWidth \ cfonts(1).HeaderInfo.CellWidth
    cfonts(1).ColFactor = cfonts(1).HeaderInfo.CellWidth / cfonts(1).HeaderInfo.BitmapWidth
    cfonts(1).RowFactor = cfonts(1).HeaderInfo.CellHeight / cfonts(1).HeaderInfo.BitmapHeight
    
    'Cache the verticies used to draw the character (only requires setting the color and adding to the X/Y values)
    For LoopChar = 0 To 255
        
        'tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
        Row = (LoopChar - cfonts(1).HeaderInfo.BaseCharOffset) \ cfonts(1).RowPitch
        u = ((LoopChar - cfonts(1).HeaderInfo.BaseCharOffset) - (Row * cfonts(1).RowPitch)) * cfonts(1).ColFactor
        v = Row * cfonts(1).RowFactor

        'Set the verticies
        With cfonts(1).HeaderInfo.CharVA(LoopChar)
            .X = 0
            .Y = 0
            .W = cfonts(1).HeaderInfo.CellWidth
            .H = cfonts(1).HeaderInfo.CellHeight
            .Tx1 = u
            .Ty1 = v
            .Tx2 = u + cfonts(1).ColFactor
            .Ty2 = v + cfonts(1).RowFactor
        End With
        
    Next LoopChar
    
End Sub

Public Sub Draw_Box(ByVal X As Integer, ByVal Y As Integer, ByVal W As Integer, ByVal H As Integer, BackgroundColor() As Long)
    
    Call SpriteBatch.SetTexture(Nothing)
    Call SpriteBatch.Draw(X, Y, W, H, BackgroundColor)
End Sub

Public Sub ChangeRenderState(ByVal State As eRenderState)
    FadingRState = RenderState
    RenderState = State
    
    FadeOn = False
    FadeOff = True
    
    With frmConnect
        
        Select Case State
        
            Case eRenderState.eNewCharInfo
                Call ShowComponents(.txtNick, .txtMail, .txtPass, .txtRepPass)
                Call DisableComponents(.btnCrearPj, .btnLogin, .btnHeadDer, .btnHeadIzq)
                Call EnableComponents(.btnSiguiente, .btnAtras)
                Call HideComponents(.txtNombre, .txtPassword, .lblAgilidad, .lblCarisma, .lblConstitucion, .lblFuerza, _
                                    .lblInteligencia, .cmbHogar, .cmbRaza, .cmbSexo, .lstSkills)
                                    
            Case eRenderState.eNewCharDetails
                Call ShowComponents(.cmbHogar, .cmbRaza, .cmbSexo)
                Call EnableComponents(.btnHeadDer, .btnHeadIzq)
                Call HideComponents(.txtNick, .txtMail, .txtPass, .txtRepPass, .lblAgilidad, .lblCarisma, .lblConstitucion, _
                                    .lblFuerza, .lblInteligencia)
            
            Case eRenderState.eNewCharAttrib
                Call DisableComponents(.btnHeadDer, .btnHeadIzq)
                Call ShowComponents(.lblAgilidad, .lblCarisma, .lblConstitucion, .lblFuerza, .lblInteligencia)
                Call HideComponents(.cmbHogar, .cmbRaza, .cmbSexo, .lstSkills)
            
            Case eRenderState.eNewCharSkills
                Call ShowComponents(.lstSkills)
                Call HideComponents(.lblAgilidad, .lblCarisma, .lblConstitucion, .lblFuerza, .lblInteligencia)
                
                With HelpWindow
                    
                    ReDim .Text(0 To 5) As String
                    
                    .Text(0) = "Utiliza la rueda para ver mas skills."
                    .Text(1) = "Click: izquierdo suma, derecho resta."
                    .Text(3) = "Mientras asignas, mantén:"
                    .Text(4) = "    Ctrl: Para asignar 3 puntos."
                    .Text(5) = "    Shift: Para asignarlos todos."
                    
                    .Active = True
                End With
        
        End Select
        
    End With

End Sub

Public Function GetRenderState() As eRenderState
    GetRenderState = RenderState
End Function

Public Sub SetBodyExample(ByVal UserBody As Integer)
    
    BodyExample = BodyData(UserBody).Walk(3)
    BodyExample.Started = 1
    BodyExample.Loops = INFINITE_LOOPS
End Sub

Public Sub Time_Update()
    timerElapsedTime = GetElapsedTime()
End Sub

Private Sub GameTime_Update()
    'Get timing info
    timerTicksPerFrame = timerElapsedTime * engineBaseSpeed
End Sub

Private Sub AnimTime_Update()
    animTicksPerFrame = timerElapsedTime * engineAnimSpeed
End Sub

'CSEH: ErrLog
Public Sub SwitchMap(ByVal Map As Integer)
    Dim Y As Long
    Dim X As Long
    Dim handle As Integer
    Dim Reader As New CsBuffer
    Dim data() As Byte
    Dim ByFlags As Byte
    
    handle = FreeFile()
    
    Open DirMapas & "Mapa" & Map & ".mcl" For Binary As handle
        Seek handle, 1
        ReDim data(0 To LOF(handle) - 1) As Byte
        
        Get handle, , data
    Close handle
    
    Call Reader.Wrap(data)
    
    'map :poop: Header
    MapInfo.MapVersion = Reader.ReadInteger
    
    Dim i As Long
    
    'Load arrays
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
        
            With MapData(X, Y)
                ByFlags = Reader.ReadByte
                    
                .Blocked = ByFlags And 1
                    
                .Graphic(1).GrhIndex = Reader.ReadInteger
                Call InitGrh(.Graphic(1), .Graphic(1).GrhIndex)
                
                For i = 2 To 4
                    If ByFlags And (2 ^ (i - 1)) Then
                        .Graphic(i).GrhIndex = Reader.ReadInteger
                        Call InitGrh(.Graphic(i), .Graphic(i).GrhIndex)
                        
                    Else
                        .Graphic(i).GrhIndex = 0
                    End If
                Next
                
                For i = 4 To 6
                    If (ByFlags And 2 ^ i) Then .Trigger = .Trigger Or 2 ^ (i - 4)
                Next
                
                'Erase NPCs
                If MapData(X, Y).CharIndex > 0 Then
                    Call EraseChar(MapData(X, Y).CharIndex)
                End If
                
                'Erase OBJs
                MapData(X, Y).ObjGrh.GrhIndex = 0
                
            End With
        Next X
    Next Y

    Set Reader = Nothing
    
    
    MapInfo.Name = ""
    MapInfo.Music = ""
    
    CurMap = Map
End Sub

Private Function Input_Key_Get(ByVal key_code As Long) As Boolean
        '**************************************************************
        'Author: Aaron Perkins
        'Last Modify Date: 10/07/2002
        '
        '**************************************************************

        If GetAsyncKeyState(key_code) < 0 Then
                Input_Key_Get = True
        End If

End Function

Public Function Fade_Render_Off()
    
    FadeOff = True
    
End Function

Public Sub InjectAlphaToColor(ByVal Alpha As Byte, ByRef Color() As Long)
    
    Dim i As Long
    
    For i = 0 To 3
        Color(i) = ARGB(Color(i), Alpha)
    Next
End Sub

Private Function ARGB(ByVal lngColor As Long, ByVal Alpha As Byte) As Long
    If Alpha > 127 Then ' handle high bit and prevent overflow
        ARGB = lngColor And ((Alpha And Not &H80) * &H1000000) Or &H80000000
    Else
        ARGB = lngColor And (Alpha * &H1000000)
    End If
End Function

