Attribute VB_Name = "modFunctions"
Option Explicit

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
    
End Sub

Function HayUserAbajo(X As Integer, Y As Integer, GrhIndex) As Boolean

HayUserAbajo = CharList(UserCharIndex).Pos.X >= X - (GrhData(GrhIndex).TileWidth \ 2) _
    And CharList(UserCharIndex).Pos.X <= X + (GrhData(GrhIndex).TileWidth \ 2) _
    And CharList(UserCharIndex).Pos.Y >= Y - (GrhData(GrhIndex).TileHeight - 1) _
    And CharList(UserCharIndex).Pos.Y <= Y

End Function

Sub MakeChar(CharIndex As Integer, Body As Integer, Head As Integer, Heading As Byte, ByVal X As Integer, ByVal Y As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 by GS
'*************************************************

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


