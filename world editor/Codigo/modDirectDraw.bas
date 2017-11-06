Attribute VB_Name = "modDirectDraw"
Option Explicit

Function LoadWavetoDSBuffer(DS As DirectSound, DSB As DirectSoundBuffer, sFile As String) As Boolean
    
    Dim bufferDesc As DSBUFFERDESC
    Dim waveFormat As WAVEFORMATEX
    
    bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
    
    waveFormat.nFormatTag = WAVE_FORMAT_PCM
    waveFormat.nChannels = 2
    waveFormat.lSamplesPerSec = 22050
    waveFormat.nBitsPerSample = 16
    waveFormat.nBlockAlign = waveFormat.nBitsPerSample / 8 * waveFormat.nChannels
    waveFormat.lAvgBytesPerSec = waveFormat.lSamplesPerSec * waveFormat.nBlockAlign
    Set DSB = DS.CreateSoundBufferFromFile(sFile, bufferDesc, waveFormat)
    
    If Err.Number <> 0 Then
        Exit Function
    End If
    
    LoadWavetoDSBuffer = True
    
End Function

Sub ConvertCPtoTP(StartPixelLeft As Integer, StartPixelTop As Integer, ByVal CX As Single, ByVal CY As Single, tX As Integer, tY As Integer)

Dim HWindowX As Integer
Dim HWindowY As Integer

CX = CX - StartPixelLeft
CY = CY - StartPixelTop

HWindowX = (WindowTileWidth \ 2)
HWindowY = (WindowTileHeight \ 2)

'Figure out X and Y tiles
CX = (CX \ TilePixelWidth)
CY = (CY \ TilePixelHeight)

If CX > HWindowX Then
    CX = (CX - HWindowX)

Else
    If CX < HWindowX Then
        CX = (0 - (HWindowX - CX))
    Else
        CX = 0
    End If
End If

If CY > HWindowY Then
    CY = (0 - (HWindowY - CY))
Else
    If CY < HWindowY Then
        CY = (CY - HWindowY)
    Else
        CY = 0
    End If
End If

tX = UserPos.X + CX
tY = UserPos.Y + CY

End Sub

Function DeInitTileEngine() As Boolean

Dim loopc As Integer

EngineRun = False

'****** Clear DirectX objects ******
Set PrimarySurface = Nothing
Set PrimaryClipper = Nothing
Set BackBufferSurface = Nothing

Set SurfaceDB = Nothing

Set DirectDraw = Nothing

'Reset any channels that are done
For loopc = 1 To NumSoundBuffers
    Set DSBuffers(loopc) = Nothing
Next loopc

Set DirectSound = Nothing

Set DirectX = Nothing

DeInitTileEngine = True

End Function

Sub MakeChar(CharIndex As Integer, Body As Integer, Head As Integer, Heading As Byte, X As Integer, Y As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 by GS
'*************************************************
On Error Resume Next

'Update LastChar
If CharIndex > LastChar Then LastChar = CharIndex
NumChars = NumChars + 1

'Update head, body, ect.
CharList(CharIndex).Body = BodyData(Body)
CharList(CharIndex).Head = HeadData(Head)
CharList(CharIndex).Heading = Heading

'Reset moving stats
CharList(CharIndex).Moving = 0
CharList(CharIndex).MoveOffset.X = 0
CharList(CharIndex).MoveOffset.Y = 0

'Update position
CharList(CharIndex).Pos.X = X
CharList(CharIndex).Pos.Y = Y

'Make active
CharList(CharIndex).Active = 1

'Plot on map
MapData(X, Y).CharIndex = CharIndex

bRefreshRadar = True ' GS

End Sub

Sub EraseChar(CharIndex As Integer)

If CharIndex = 0 Then Exit Sub
'Make un-active
CharList(CharIndex).Active = 0

'Update lastchar
If CharIndex = LastChar Then
    Do Until CharList(LastChar).Active = 1
        LastChar = LastChar - 1
        If LastChar = 0 Then Exit Do
    Loop
End If

MapData(CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y).CharIndex = 0

'Update NumChars
NumChars = NumChars - 1

bRefreshRadar = True ' GS

End Sub

Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional Started As Byte = 2)
'*************************************************
'Author: Unkwown
'Last modified: 31/05/06 - GS
'*************************************************
On Error Resume Next
Grh.GrhIndex = GrhIndex
If Grh.GrhIndex <> 0 Then ' 31/05/2006
    If Started = 2 Then
        If GrhData(Grh.GrhIndex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0
        End If
    Else
        Grh.Started = Started
    End If
    Grh.FrameCounter = 1
    Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
Else
    Grh.FrameCounter = 1
    Grh.Started = 0
    Grh.SpeedCounter = 0
End If

End Sub

Sub MoveCharbyHead(CharIndex As Integer, nHeading As Byte)

Dim addX As Integer
Dim addY As Integer
Dim X As Integer
Dim Y As Integer
Dim nX As Integer
Dim nY As Integer

X = CharList(CharIndex).Pos.X
Y = CharList(CharIndex).Pos.Y

'Figure out which way to move
Select Case nHeading

    Case NORTH
        addY = -1

    Case EAST
        addX = 1

    Case SOUTH
        addY = 1
    
    Case WEST
        addX = -1
        
End Select

nX = X + addX
nY = Y + addY

MapData(nX, nY).CharIndex = CharIndex
CharList(CharIndex).Pos.X = nX
CharList(CharIndex).Pos.Y = nY
MapData(X, Y).CharIndex = 0

CharList(CharIndex).MoveOffset.X = -1 * (TilePixelWidth * addX)
CharList(CharIndex).MoveOffset.Y = -1 * (TilePixelHeight * addY)

CharList(CharIndex).Moving = 1
CharList(CharIndex).Heading = nHeading

End Sub

Sub MoveCharbyPos(CharIndex As Integer, nX As Integer, nY As Integer)

Dim X As Integer
Dim Y As Integer
Dim addX As Integer
Dim addY As Integer
Dim nHeading As Byte

X = CharList(CharIndex).Pos.X
Y = CharList(CharIndex).Pos.Y

addX = nX - X
addY = nY - Y

If Sgn(addX) = 1 Then
    nHeading = EAST
End If

If Sgn(addX) = -1 Then
    nHeading = WEST
End If

If Sgn(addY) = -1 Then
    nHeading = NORTH
End If

If Sgn(addY) = 1 Then
    nHeading = SOUTH
End If

MapData(nX, nY).CharIndex = CharIndex
CharList(CharIndex).Pos.X = nX
CharList(CharIndex).Pos.Y = nY
MapData(X, Y).CharIndex = 0

CharList(CharIndex).MoveOffset.X = -1 * (TilePixelWidth * addX)
CharList(CharIndex).MoveOffset.Y = -1 * (TilePixelHeight * addY)

CharList(CharIndex).Moving = 1
CharList(CharIndex).Heading = nHeading

bRefreshRadar = True ' GS

End Sub

Function NextOpenChar() As Integer

Dim loopc As Integer

loopc = 1
Do While CharList(loopc).Active
    loopc = loopc + 1
Loop

NextOpenChar = loopc

End Function

Function LegalPos(X As Integer, Y As Integer) As Boolean

LegalPos = True

'Check to see if its out of bounds
If X - 8 < 1 Or X - 8 > 100 Or Y - 6 < 1 Or Y - 6 > 100 Then
    LegalPos = False
    Exit Function
End If

'Check to see if its blocked
If MapData(X, Y).Blocked = 1 Then
    LegalPos = False
    Exit Function
End If

'Check for character
If MapData(X, Y).CharIndex > 0 Then
    LegalPos = False
    Exit Function
End If

End Function

Function InMapLegalBounds(X As Integer, Y As Integer) As Boolean

If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    InMapLegalBounds = False
    Exit Function
End If

InMapLegalBounds = True

End Function

Function InMapBounds(X As Integer, Y As Integer) As Boolean

If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
    InMapBounds = False
    Exit Function
End If

InMapBounds = True

End Function

Sub DDrawTransGrhtoSurface(ByRef Surface As DirectDrawSurface7, Grh As Grh, ByVal X As Integer, ByVal Y As Integer, Center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0)
'*************************************************
'Author: Unknow
'Last modified: 09/12/09 - Deut
'*************************************************
If MapaCargado = False Then Exit Sub

Dim iGrhIndex As Integer
Dim SourceRect As RECT
Dim QuitarAnimacion As Boolean

If Grh.GrhIndex = 0 Then Exit Sub

'Figure out what frame to draw (always 1 if not animated)
iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
If iGrhIndex = 0 Then Exit Sub
'Center Grh over X,Y pos
If Center Then
    If GrhData(iGrhIndex).TileWidth <> 1 Then
        Rem Scena * cambio esto para que la capa2 se vea siempre correcta y no con ese rectangulito abajo a la derecha en negro
        Rem si se remplazan los 16 por 32 se arregla pero las cosas que están en la capa2 no quedan centradas.
        X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16 'hard coded for speed
    End If
    If GrhData(iGrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32 'hard coded for speed
    End If
End If

With SourceRect
    .Left = GrhData(iGrhIndex).sX
    .Top = GrhData(iGrhIndex).sY
    .Right = .Left + GrhData(iGrhIndex).pixelWidth
    .Bottom = .Top + GrhData(iGrhIndex).pixelHeight
End With '
Surface.BltFast X, Y, SurfaceDB.Surface(GrhData(iGrhIndex).FileNum), SourceRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY

End Sub
Sub DDrawSpecialToSurface(ByRef Surface As DirectDrawSurface7, ByVal SpecialNumber As Byte, ByVal X As Integer, ByVal Y As Integer)
'*************************************************
'Author: Deut
'Last modified: 11/12/09
'*************************************************
If MapaCargado = False Then Exit Sub

'Dim iGrhIndex As Integer
Dim SourceRect As RECT
'Dim QuitarAnimacion As Boolean

'If Grh.GrhIndex = 0 Then Exit Sub
If SpecialNumber = 0 Then Exit Sub


'Figure out what frame to draw (always 1 if not animated)
'iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
'If iGrhIndex = 0 Then Exit Sub
'Center Grh over X,Y pos
'If Center Then
'    If GrhData(iGrhIndex).TileWidth <> 1 Then
'        Rem Scena * cambio esto para que la capa2 se vea siempre correcta y no con ese rectangulito abajo a la derecha en negro
'        Rem si se remplazan los 16 por 32 se arregla pero las cosas que están en la capa2 no quedan centradas.
'        X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16 'hard coded for speed
'    End If
'    If GrhData(iGrhIndex).TileHeight <> 1 Then
'        Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32 'hard coded for speed
'    End If
'End If

With SourceRect
    .Left = (SpecialNumber - 1) * 32
    .Top = 0
    .Right = .Left + 32
    .Bottom = 32
End With '
Surface.BltFast X, Y, SurfaceDB.Surface(0), SourceRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY

End Sub
'Sub DibujarGrhIndex(ByRef Surface As DirectDrawSurface7, iGrhIndex As Integer, ByVal X As Integer, ByVal Y As Integer, Center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0, Optional Alpha As Byte = 200)
'
'If MapaCargado = False Then Exit Sub
'
'Dim SourceRect As RECT
'Dim QuitarAnimacion As Boolean
'
''Figure out what frame to draw (always 1 if not animated)
'
'If iGrhIndex = 0 Then Exit Sub
''Center Grh over X,Y pos
'If Center Then
'    If GrhData(iGrhIndex).TileWidth <> 1 Then
'        X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16 'hard coded for speed
'    End If
'    If GrhData(iGrhIndex).TileHeight <> 1 Then
'        Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32 'hard coded for speed
'    End If
'End If
'
'With SourceRect
'    .Left = GrhData(iGrhIndex).sX
'    .Top = GrhData(iGrhIndex).sY
'    .Right = .Left + GrhData(iGrhIndex).pixelWidth
'    .Bottom = .Top + GrhData(iGrhIndex).pixelHeight
'End With '
''MotorDeEfectos.DBAlpha SurfaceDB.Surface(GrhData(iGrhIndex).FileNum), SourceRect, False, X, Y, Alpha
'Surface.BltFast X, Y, SurfaceDB.Surface(GrhData(iGrhIndex).FileNum), SourceRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
'
'End Sub

Sub DrawBackBufferSurface()
PrimarySurface.Blt MainViewRect, BackBufferSurface, MainDestRect, DDBLT_WAIT
End Sub

Sub DrawGrhtoHdc(hWnd As Long, hdc As Long, ByVal Grh As Long, SourceRect As RECT, destRect As RECT)

On Error Resume Next
If Grh <= 0 Then Exit Sub
Dim aux As Integer
aux = GrhData(Grh).FileNum
If aux = 0 Then Exit Sub
SecundaryClipper.SetHWnd hWnd
SurfaceDB.Surface(aux).BltToDC hdc, SourceRect, destRect
End Sub

Sub PlayWaveDS(file As String)
    
    'Cylce through avaiable sound buffers
    LastSoundBufferUsed = LastSoundBufferUsed + 1
    If LastSoundBufferUsed > NumSoundBuffers Then
        LastSoundBufferUsed = 1
    End If
    
    If LoadWavetoDSBuffer(DirectSound, DSBuffers(LastSoundBufferUsed), file) Then
        DSBuffers(LastSoundBufferUsed).Play DSBPLAY_DEFAULT
    End If

End Sub

Public Sub DePegar()
    
    Dim X As Integer
    Dim Y As Integer

    For X = 0 To DeSeleccionAncho - 1
        For Y = 0 To DeSeleccionAlto - 1
             MapData(X + DeSeleccionOX, Y + DeSeleccionOY) = DeSeleccionMap(X, Y)
        Next
    Next
End Sub

Public Sub PegarSeleccion() '(mx As Integer, my As Integer)
    
    'podria usar copy mem , pero por las dudas no XD
    Static UltimoX As Integer
    Static UltimoY As Integer
    
    If UltimoX = SobreX And UltimoY = SobreY Then Exit Sub
    
    UltimoX = SobreX
    UltimoY = SobreY
    
    Dim X As Integer
    Dim Y As Integer
    
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SobreX
    DeSeleccionOY = SobreY
    
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To DeSeleccionAncho - 1
        For Y = 0 To DeSeleccionAlto - 1
            DeSeleccionMap(X, Y) = MapData(X + SobreX, Y + SobreY)
        Next
    Next
    
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
             MapData(X + SobreX, Y + SobreY) = SeleccionMap(X, Y)
        Next
    Next
    
    Seleccionando = False

End Sub

Public Sub AccionSeleccion()
    
    Dim X As Integer
    Dim Y As Integer
    
    SeleccionAncho = Abs(SeleccionIX - SeleccionFX) + 1
    SeleccionAlto = Abs(SeleccionIY - SeleccionFY) + 1
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SeleccionIX
    DeSeleccionOY = SeleccionIY
    
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
            DeSeleccionMap(X, Y) = MapData(X + MinInt(SeleccionIX, SeleccionFX), Y + MinInt(SeleccionIY, SeleccionFY))
        Next
    Next
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
           ClickEdit vbLeftButton, MinInt(SeleccionIX, SeleccionFX) + X, MinInt(SeleccionIY, SeleccionFY) + Y
        Next
    Next
    Seleccionando = False
End Sub

Public Sub BlockearSeleccion()
    
    Dim X As Integer
    Dim Y As Integer
    Dim Vacio As MapBlock
    SeleccionAncho = Abs(SeleccionIX - SeleccionFX) + 1
    SeleccionAlto = Abs(SeleccionIY - SeleccionFY) + 1
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SeleccionIX
    DeSeleccionOY = SeleccionIY
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
            DeSeleccionMap(X, Y) = MapData(X + SeleccionIX, Y + SeleccionIY)
        Next
    Next
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
             If MapData(X + SeleccionIX, Y + SeleccionIY).Blocked = 1 Then
                MapData(X + SeleccionIX, Y + SeleccionIY).Blocked = 0
             Else
                MapData(X + SeleccionIX, Y + SeleccionIY).Blocked = 1
            End If
        Next
    Next
    Seleccionando = False
End Sub

Public Sub CortarSeleccion()
    
    CopiarSeleccion
    Dim X As Integer
    Dim Y As Integer
    Dim Vacio As MapBlock
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SeleccionIX
    DeSeleccionOY = SeleccionIY
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
            DeSeleccionMap(X, Y) = MapData(X + SeleccionIX, Y + SeleccionIY)
        Next
    Next
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
             MapData(X + SeleccionIX, Y + SeleccionIY) = Vacio
        Next
    Next
    Seleccionando = False
End Sub
Public Sub CopiarSeleccion()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    'podria usar copy mem , pero por las dudas no XD
    Dim X As Integer
    Dim Y As Integer
    Seleccionando = False
    SeleccionAncho = Abs(SeleccionIX - SeleccionFX) + 1
    SeleccionAlto = Abs(SeleccionIY - SeleccionFY) + 1
    ReDim SeleccionMap(SeleccionAncho, SeleccionAlto) As MapBlock
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
            SeleccionMap(X, Y) = MapData(X + SeleccionIX, Y + SeleccionIY)
        Next
    Next
End Sub
Public Sub GenerarVista()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
   ' hacer una llamada a un seter o geter , es mas lento q una variable
   ' con esto hacemos q no este preguntando a el objeto cadavez
   ' q dibuja , Render mas rapido ;)
'    VerBlockeados = frmMain.cVerBloqueos.value
'    VerTriggers = frmMain.cVerTriggers.value
    VerCapa1 = frmMain.mnuVerCapa1.Checked
    VerCapa2 = frmMain.mnuVerCapa2.Checked
    VerCapa3 = frmMain.mnuVerCapa3.Checked
    VerCapa4 = frmMain.mnuVerCapa4.Checked
    VerTranslados = frmMain.mnuVerTranslados.Checked
    VerObjetos = frmMain.mnuVerObjetos.Checked
    VerNpcs = frmMain.mnuVerNPCs.Checked
    VerBloqueos = frmMain.mnuVerBloqueos.Checked
    VerAgua = frmMain.mnuVerAgua.Checked
    VerTriggers = frmMain.mnuVerTriggers.Checked
    
End Sub

Public Sub RenderScreen(TileX As Integer, TileY As Integer, PixelOffsetX As Integer, PixelOffsetY As Integer)
'*************************************************
'Author: Unknow
'Last modified: 18/12/09 - Deut
'*************************************************
On Error Resume Next

Dim Y As Integer, X As Integer
Dim minY As Integer, minX As Integer, maxY As Integer, maxX   As Integer
Dim iPPx As Integer, iPPy As Integer
Dim ScreenX As Integer, ScreenY As Integer
Dim PixelOffsetXTemp As Integer, PixelOffsetYTemp As Integer
Dim Sobre As Integer
Dim Moved As Byte
Dim Grh As Grh
Dim bCapa As Byte
Dim SelRect As RECT
Dim rSourceRect As RECT, r As RECT
Dim iGrhIndex As Integer
Dim TempChar As Char

BackBufferSurface.BltColorFill r, 0

minY = (TileY - (WindowTileHeight \ 2)) - TileBufferSize
maxY = (TileY + (WindowTileHeight \ 2)) + TileBufferSize
minX = (TileX - (WindowTileWidth \ 2)) - TileBufferSize
maxX = (TileX + (WindowTileWidth \ 2)) + TileBufferSize

If Val(frmMain.cCapas.Text) >= 1 And (frmMain.cCapas.Text) <= 4 Then
    bCapa = Val(frmMain.cCapas.Text)
Else
    bCapa = 1
End If

Call GenerarVista

ScreenY = 8
For Y = (minY + 8) To (maxY - 8)
    ScreenX = 8
    For X = (minX + 8) To (maxX - 8)
        If InMapBounds(X, Y) Then
            If X > 100 Or Y < 1 Then Exit For
            
            'Layer 1 **********************************
            Sobre = -1
            Dim aux As Integer
            Dim dy As Integer
            Dim dx As Integer
            If SobreX = X And SobreY = Y Then
                If frmMain.cSeleccionarSuperficie.value = True Then
                    Sobre = MapData(X, Y).Graphic(bCapa).GrhIndex
                    
                    If frmConfigSup.MOSAICO.value = vbChecked Then
'                        Dim aux As Integer
'                        Dim dy As Integer
'                        Dim dx As Integer
                        
                        If frmConfigSup.DespMosaic.value = vbChecked Then
                            dy = Val(frmConfigSup.DMLargo.Text)
                            dx = Val(frmConfigSup.DMAncho.Text)
                        Else
                            dy = 0
                            dx = 0
                        End If
                        
                        If frmMain.mnuAutoCompletarSuperficies.Checked = False Then
                            aux = Val(frmMain.cGrh.Text) + (((Y + dy) Mod frmConfigSup.mLargo.Text) * frmConfigSup.mAncho.Text) + ((X + dx) Mod frmConfigSup.mAncho.Text)
                            If MapData(X, Y).Graphic(bCapa).GrhIndex <> aux Then
                                MapData(X, Y).Graphic(bCapa).GrhIndex = aux
                                InitGrh MapData(X, Y).Graphic(bCapa), aux
                            End If
                        Else
                            aux = Val(frmMain.cGrh.Text) + (((Y + dy) Mod frmConfigSup.mLargo.Text) * frmConfigSup.mAncho.Text) + ((X + dx) Mod frmConfigSup.mAncho.Text)
                            If MapData(X, Y).Graphic(bCapa).GrhIndex <> aux Then
                                MapData(X, Y).Graphic(bCapa).GrhIndex = aux
                                InitGrh MapData(X, Y).Graphic(bCapa), aux
                            End If
                        End If
                    Else
                        If MapData(X, Y).Graphic(bCapa).GrhIndex <> Val(frmMain.cGrh.Text) Then
                            MapData(X, Y).Graphic(bCapa).GrhIndex = Val(frmMain.cGrh.Text)
                            InitGrh MapData(X, Y).Graphic(bCapa), Val(frmMain.cGrh.Text)
                        End If
                    End If
                End If
'            Else
'                If frmMain.cSeleccionarSuperficie.value = True Then
'                    If frmConfigSup.MOSAICO.value = vbChecked Then
'                        If frmConfigSup.DespMosaic.value = vbChecked Then
'                            dy = Val(frmConfigSup.DMLargo.Text)
'                            dx = Val(frmConfigSup.DMAncho.Text)
'                        Else
'                            dy = 0
'                            dx = 0
'                        End If
'                        Dim HWy As Byte
'                        Dim HWx As Byte
'                        HWx = (SobreX + dx) Mod frmConfigSup.mAncho.Text
'                        HWy = (SobreY + dy) Mod frmConfigSup.mLargo.Text
'                        If (Y >= SobreY - HWy) And (Y < SobreY - HWy + frmConfigSup.mLargo.Text) And (X >= SobreX - HWx) And (X < SobreX - HWx + frmConfigSup.mAncho.Text) Then
'                            Sobre = MapData(X, Y).Graphic(bCapa).GrhIndex
'                            aux = Val(frmMain.cGrh.Text) + (((Y + dy) Mod frmConfigSup.mLargo.Text) * frmConfigSup.mAncho.Text) + ((X + dx) Mod frmConfigSup.mAncho.Text)
'                            If MapData(X, Y).Graphic(bCapa).GrhIndex <> aux Then
'                                MapData(X, Y).Graphic(bCapa).GrhIndex = aux
'                                InitGrh MapData(X, Y).Graphic(bCapa), aux
'                            End If
'                        End If
'                    End If
'                End If
            End If
            
            If VerCapa1 Then
                With MapData(X, Y).Graphic(1)
                    If (.GrhIndex <> 0) Then
                        If Grh.Started = 1 Then
                            If (.SpeedCounter > 0) Then
                                .SpeedCounter = .SpeedCounter - 1
                                If (.SpeedCounter = 0) Then
                                    .SpeedCounter = GrhData(.GrhIndex).Speed
                                    .FrameCounter = .FrameCounter + 1
                                    If (.FrameCounter > GrhData(.GrhIndex).NumFrames) Then _
                                        .FrameCounter = 1
                                End If
                            End If
                        End If
                        iGrhIndex = GrhData(.GrhIndex).Frames(.FrameCounter)
                    End If
                End With
                
                If iGrhIndex <> 0 Then
                    rSourceRect.Left = GrhData(iGrhIndex).sX
                    rSourceRect.Top = GrhData(iGrhIndex).sY
                    rSourceRect.Right = rSourceRect.Left + GrhData(iGrhIndex).pixelWidth
                    rSourceRect.Bottom = rSourceRect.Top + GrhData(iGrhIndex).pixelHeight
                    Call BackBufferSurface.BltFast(((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, SurfaceDB.Surface(GrhData(iGrhIndex).FileNum), rSourceRect, DDBLTFAST_WAIT)
                End If
            End If
            
            'Layer 2 **********************************
            If MapData(X, Y).Graphic(2).GrhIndex <> 0 And VerCapa2 Then
                Call DDrawTransGrhtoSurface(BackBufferSurface, MapData(X, Y).Graphic(2), ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, 1, 1)
            End If
            
            If Sobre >= 0 Then
                If MapData(X, Y).Graphic(bCapa).GrhIndex <> Sobre Then
                    MapData(X, Y).Graphic(bCapa).GrhIndex = Sobre
                    InitGrh MapData(X, Y).Graphic(bCapa), Sobre
                End If
            End If
        End If
        ScreenX = ScreenX + 1
    Next X
    ScreenY = ScreenY + 1
    If Y > 100 Then Exit For
Next Y

ScreenY = 8
For Y = (minY + 8) To (maxY - 1)
    ScreenX = 5
    For X = (minX + 5) To (maxX - 5)
        If InMapBounds(X, Y) Then
            If X > 100 Or X < -3 Then Exit For
            iPPx = ((32 * ScreenX) - 32) + PixelOffsetX
            iPPy = ((32 * ScreenY) - 32) + PixelOffsetY
             
            'Object Layer **********************************
            If MapData(X, Y).OBJInfo.objindex <> 0 And VerObjetos Then
                Call DDrawTransGrhtoSurface(BackBufferSurface, MapData(X, Y).ObjGrh, iPPx, iPPy, 1, 1)
            End If
            
            'Char layer **********************************
            If MapData(X, Y).CharIndex <> 0 And VerNpcs Then
                TempChar = CharList(MapData(X, Y).CharIndex)
                PixelOffsetXTemp = PixelOffsetX
                PixelOffsetYTemp = PixelOffsetY
               
                If TempChar.Head.Head(TempChar.Heading).GrhIndex <> 0 Then
                     Call DDrawTransGrhtoSurface(BackBufferSurface, TempChar.Body.Walk(TempChar.Heading), (PixelPos(ScreenX) + PixelOffsetXTemp), PixelPos(ScreenY) + PixelOffsetYTemp, 1, 1)
                     Call DDrawTransGrhtoSurface(BackBufferSurface, TempChar.Head.Head(TempChar.Heading), (PixelPos(ScreenX) + PixelOffsetXTemp) + TempChar.Body.HeadOffset.X, PixelPos(ScreenY) + PixelOffsetYTemp + TempChar.Body.HeadOffset.Y, 1, 0)
                Else
                    Call DDrawTransGrhtoSurface(BackBufferSurface, TempChar.Body.Walk(TempChar.Heading), (PixelPos(ScreenX) + PixelOffsetXTemp), PixelPos(ScreenY) + PixelOffsetYTemp, 1, 1)
                End If
            End If
             
            'Layer 3 *****************************************
            If MapData(X, Y).Graphic(3).GrhIndex <> 0 And VerCapa3 Then
               Call DDrawTransGrhtoSurface(BackBufferSurface, MapData(X, Y).Graphic(3), iPPx, iPPy, 1, 1)
            End If
        End If
        ScreenX = ScreenX + 1
    Next X
    ScreenY = ScreenY + 1
Next Y

'Tiles blokeadas, techos, triggers , seleccion
ScreenY = 5
For Y = (minY + 5) To (maxY - 1)
    ScreenX = 5
    For X = (minX + 5) To (maxX)
        If X < 101 And X > 0 And Y < 101 And Y > 0 Then
            iPPx = ((32 * ScreenX) - 32) + PixelOffsetX
            iPPy = ((32 * ScreenY) - 32) + PixelOffsetY
            
            'Show layer 4
            If MapData(X, Y).Graphic(4).GrhIndex <> 0 And VerCapa4 Then
                Call DDrawTransGrhtoSurface(BackBufferSurface, MapData(X, Y).Graphic(4), iPPx, iPPy, 1, 1)
            End If
            
            'Show Tile Exits
            If MapData(X, Y).TileExit.Map <> 0 And VerTranslados Then
'                Grh.GrhIndex = 3
'                Grh.FrameCounter = 1
'                Grh.Started = 0
'                Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, iPPx, iPPy, 1, 1)
                Call DDrawSpecialToSurface(BackBufferSurface, 1, iPPx, iPPy)
            End If
            
            'Show water tiles
            If VerAgua Then
                If HayAgua(X, Y) And (Not MapData(X, Y).Blocked Or (MapData(X, Y).Blocked And Not VerBloqueos)) Then
    '                BackBufferSurface.SetForeColor vbWhite
    '                BackBufferSurface.SetFillColor vbBlue
    '                BackBufferSurface.SetFillStyle 0
    '                Call BackBufferSurface.DrawBox(iPPx + 16, iPPy + 16, iPPx + 21, iPPy + 21)
    '                Grh.FrameCounter = 1
    '                Grh.Started = 0
    '                Grh.GrhIndex = 2 'el Grh del indicador de agua
    '                Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, iPPx, iPPy, 1, 1)
                    Call DDrawSpecialToSurface(BackBufferSurface, 2, iPPx, iPPy)
                End If
            End If
                        
            'Show blocked tiles
            If VerBloqueos = True And MapData(X, Y).Blocked = 1 Then
'                BackBufferSurface.SetForeColor vbWhite
'                BackBufferSurface.SetFillColor vbRed
'                BackBufferSurface.SetFillStyle 0
'                Call BackBufferSurface.DrawBox(iPPx + 16, iPPy + 16, iPPx + 21, iPPy + 21)
'                Grh.FrameCounter = 1
'                Grh.Started = 0
'                Grh.GrhIndex = 4 'el Grh del gráfico del bloqueo
'                Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, iPPx, iPPy, 1, 1)
                Call DDrawSpecialToSurface(BackBufferSurface, 4, iPPx, iPPy)
            End If
            
            'Show triggers
            If VerTriggers Then
                If MapData(X, Y).Trigger <> 0 Then
'                    Grh.FrameCounter = 1
'                    Grh.Started = 0
'                    Grh.GrhIndex = 16939
                    If MapData(X, Y).Trigger > 0 And MapData(X, Y).Trigger < 10 Then
                        Call DDrawSpecialToSurface(BackBufferSurface, MapData(X, Y).Trigger + 6, iPPx, iPPy)
                    Else
                        Call DDrawSpecialToSurface(BackBufferSurface, 16, iPPx, iPPy)
                    End If
'                    Select Case MapData(X, Y).Trigger
'                        Case Is = 1
'                            Grh.GrhIndex = Grh.GrhIndex + 1
'                            Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, iPPx, iPPy, 1, 1)
                            
'                        Case Is = 2
'                            Grh.GrhIndex = Grh.GrhIndex + 2
'                            Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, iPPx, iPPy, 1, 1)
'                        Case Is = 3
'                            Grh.GrhIndex = Grh.GrhIndex + 3
'                            Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, iPPx, iPPy, 1, 1)
'                        Case Is = 4
'                            Grh.GrhIndex = Grh.GrhIndex + 4
'                            Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, iPPx, iPPy, 1, 1)
'                        Case Is = 5
'                            Grh.GrhIndex = Grh.GrhIndex + 5
'                            Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, iPPx, iPPy, 1, 1)
'                        Case Is = 6
'                            Grh.GrhIndex = Grh.GrhIndex + 6
'                            Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, iPPx, iPPy, 1, 1)
'                        Case Is = 7
'                            Grh.GrhIndex = Grh.GrhIndex + 7
'                            Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, iPPx, iPPy, 1, 1)
'                        Case Is = 8
'                            Grh.GrhIndex = Grh.GrhIndex + 8
'                            Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, iPPx, iPPy, 1, 1)
'                        Case Is = 9
'                            Grh.GrhIndex = Grh.GrhIndex + 9
'                            Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, iPPx, iPPy, 1, 1)
'                        Case Else
'                            Grh.GrhIndex = Grh.GrhIndex + 10
'                            Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, iPPx, iPPy, 1, 1)
'                    End Select
                End If
            End If
            
            Rem Cuadrícula de mapa - down
            Rem special 17
            If frmMain.CBVerCuadricula.value = True Then
                Call DDrawSpecialToSurface(BackBufferSurface, 17, iPPx, iPPy)
            End If
            Rem up
            
            
            If Seleccionando Then
                If X >= MinInt(SeleccionIX, SeleccionFX) And X <= MaxInt(SeleccionIX, SeleccionFX) And Y >= MinInt(SeleccionIY, SeleccionFY) And Y <= MaxInt(SeleccionIY, SeleccionFY) Then
'                    If X <= SeleccionFX And Y <= SeleccionFY Then
'                        BackBufferSurface.SetForeColor vbMagenta
'                        BackBufferSurface.SetFillColor vbMagenta
'                        BackBufferSurface.SetFillStyle 4
'                        BackBufferSurface.DrawBox iPPx, iPPy, iPPx + 32, iPPy + 32
'                        Grh.FrameCounter = 1
'                        Grh.Started = 0
'                        Grh.GrhIndex = 4131 ' el Grh del rectangulo rayado para la selecciones de frames
'                        Call DDrawTransGrhtoSurface(BackBufferSurface, Grh, iPPx, iPPy, 1, 1)
                        Call DDrawSpecialToSurface(BackBufferSurface, 6, iPPx, iPPy)
'                    End If
                End If
            End If
            
            
            
            Rem para mostrar los rectangulos que va a ocupar la superficie ***********************
                    Rem v1 - Funciona bien - down
'                If frmMain.cSeleccionarSuperficie.value = True Then
'                    If frmConfigSup.MOSAICO.value = vbChecked Then
'                        If frmConfigSup.DespMosaic.value = vbChecked Then
'                            dy = Val(frmConfigSup.DMLargo.Text)
'                            dx = Val(frmConfigSup.DMAncho.Text)
'                        Else
'                            dy = 0
'                            dx = 0
'                        End If
'                        Dim HWy As Byte
'                        Dim HWx As Byte
'                        HWx = (SobreX + dx) Mod frmConfigSup.mAncho.Text
'                        HWy = (SobreY + dy) Mod frmConfigSup.mLargo.Text
'                        If (Y >= SobreY - HWy) And (Y < SobreY - HWy + frmConfigSup.mLargo.Text) And (X >= SobreX - HWx) And (X < SobreX - HWx + frmConfigSup.mAncho.Text) Then
'                            Call DDrawSpecialToSurface(BackBufferSurface, 5, iPPx, iPPy)
'                        End If
'                    End If
'                End If
                    Rem v1 - Funciona bien - up
                If frmMain.cSeleccionarSuperficie.value = True Then
                    If frmConfigSup.MOSAICO.value = vbChecked Then
                        If frmConfigSup.DespMosaic.value = vbChecked Then
                            dy = Val(frmConfigSup.DMLargo.Text)
                            dx = Val(frmConfigSup.DMAncho.Text)
                        Else
                            dy = 0
                            dx = 0
                        End If
                        Dim LimiteInferiorXDeRender As Byte
                        Dim LimiteInferiorYDeRender As Byte
                        Dim AnchoTile As Byte
                        Dim LargoTile As Byte
                        AnchoTile = frmConfigSup.mAncho.Text
                        LargoTile = frmConfigSup.mLargo.Text
                        LimiteInferiorXDeRender = (Int(Val(1 + (SobreX + dx) / AnchoTile) - 1)) * AnchoTile - dx
                        LimiteInferiorYDeRender = (Int(Val(1 + (SobreY + dy) / LargoTile) - 1)) * LargoTile - dy
                        If X >= LimiteInferiorXDeRender And X < LimiteInferiorXDeRender + AnchoTile And Y >= LimiteInferiorYDeRender And Y < LimiteInferiorYDeRender + LargoTile Then
                            Call DDrawSpecialToSurface(BackBufferSurface, 5, iPPx, iPPy)
                        End If
                    End If
                End If
                    
'            rem ******************************
        End If
        ScreenX = ScreenX + 1
    Next X
    ScreenY = ScreenY + 1
Next Y

End Sub

Public Sub DrawText(lngXPos As Integer, lngYPos As Integer, strText As String, lngColor As Long)
   
   If LenB(strText) <> 0 Then
        BackBufferSurface.SetFontTransparency True                           'Set the transparency flag to true
        BackBufferSurface.SetForeColor lngColor                              'Set the color of the text to the color passed to the sub
        BackBufferSurface.SetFont frmMain.Font                               'Set the font used to the font on the form
        BackBufferSurface.DrawText lngXPos, lngYPos, strText, False          'Draw the text on to the screen, in the coordinates specified
   End If

End Sub

'Function HayUserAbajo(X As Integer, Y As Integer, GrhIndex) As Boolean
'
'HayUserAbajo = CharList(UserCharIndex).Pos.X >= X - (GrhData(GrhIndex).TileWidth \ 2) _
'    And CharList(UserCharIndex).Pos.X <= X + (GrhData(GrhIndex).TileWidth \ 2) _
'    And CharList(UserCharIndex).Pos.Y >= Y - (GrhData(GrhIndex).TileHeight - 1) _
'    And CharList(UserCharIndex).Pos.Y <= Y
'
'End Function

Function PixelPos(X As Integer) As Integer

PixelPos = (TilePixelWidth * X) - TilePixelWidth

End Function

Function InitTileEngine(ByRef setDisplayFormhWnd As Long, setMainViewTop As Integer, setMainViewLeft As Integer, setTilePixelHeight As Integer, setTilePixelWidth As Integer, setWindowTileHeight As Integer, setWindowTileWidth As Integer, setTileBufferSize As Integer) As Boolean

Dim SurfaceDesc As DDSURFACEDESC2
Dim ddck As DDCOLORKEY

'Fill startup variables
DisplayFormhWnd = setDisplayFormhWnd
MainViewTop = setMainViewTop
MainViewLeft = setMainViewLeft
TilePixelWidth = setTilePixelWidth
TilePixelHeight = setTilePixelHeight
WindowTileHeight = setWindowTileHeight
WindowTileWidth = setWindowTileWidth
TileBufferSize = setTileBufferSize

MinXBorder = XMinMapSize + (ClienteWidth \ 2)
MaxXBorder = XMaxMapSize - (ClienteWidth \ 2)
MinYBorder = YMinMapSize + (ClienteHeight \ 2)
MaxYBorder = YMaxMapSize - (ClienteHeight \ 2)

MainViewWidth = (TilePixelWidth * WindowTileWidth)
MainViewHeight = (TilePixelHeight * WindowTileHeight)

'Resize mapdata array
ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock

'****** INIT DirectDraw ******
Set DirectX = New DirectX7
Set DirectDraw = DirectX.DirectDrawCreate("")

DirectDraw.SetCooperativeLevel DisplayFormhWnd, DDSCL_NORMAL

'Primary Surface
With SurfaceDesc
    .lFlags = DDSD_CAPS
    .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
End With

Set PrimarySurface = DirectDraw.CreateSurface(SurfaceDesc)
Set PrimaryClipper = DirectDraw.CreateClipper(0)

PrimaryClipper.SetHWnd frmMain.hWnd
PrimarySurface.SetClipper PrimaryClipper

Set SecundaryClipper = DirectDraw.CreateClipper(0)

With BackBufferRect
    .Left = 0
    .Top = 0
    .Right = TilePixelWidth * (WindowTileWidth + (2 * TileBufferSize))
    .Bottom = TilePixelHeight * (WindowTileHeight + (2 * TileBufferSize))
End With

With SurfaceDesc
    .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    .lHeight = BackBufferRect.Bottom
    .lWidth = BackBufferRect.Right
End With

Set BackBufferSurface = DirectDraw.CreateSurface(SurfaceDesc)

'Set color key
ddck.low = 0
ddck.high = 0
BackBufferSurface.SetColorKey DDCKEY_SRCBLT, ddck

'Load graphic data into memory
modIndices.CargarIndicesDeGraficos

If LenB(Dir(DirGraficos & "*.bmp", vbArchive)) = 0 Then
    MsgBox "La carpeta de Graficos esta vacia o incompleta!", vbCritical
    End
Else
    frmCargando.X.Caption = "Iniciando Control de Superficies..."
    Call SurfaceDB.Initialize(DirectDraw, DirGraficos)
End If

'Wave Sound
Set DirectSound = DirectX.DirectSoundCreate("")
DirectSound.SetCooperativeLevel DisplayFormhWnd, DSSCL_PRIORITY
LastSoundBufferUsed = 1

InitTileEngine = True
EngineRun = True
DoEvents

End Function

