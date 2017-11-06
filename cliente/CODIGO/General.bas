Attribute VB_Name = "Mod_General"
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

Public iplst As String

Public bFogata As Boolean

Public bLluvia() As Byte ' Array para determinar si
'debemos mostrar la animacion de la lluvia

Private lFrameTimer As Long

Private Type TYPE_LONG_BYTES
        H As Integer
        L As Integer
End Type

Private Type TYPE_LONG
        Value As Long
End Type

'http://stackoverflow.com/questions/6861733/vb6-integer-to-two-bytes-c-short-to-send-over-serial
Public Function IntegersToLong(ByVal H As Integer, ByVal L As Integer) As Long
    Dim TempTL As TYPE_LONG
    Dim TempBL As TYPE_LONG_BYTES
    
    TempBL.H = H
    TempBL.L = L
    
    LSet TempTL = TempBL
    
    IntegersToLong = TempTL.Value
End Function

Public Sub LongToIntegers(ByVal Value As Long, ByRef H As Integer, ByRef L As Integer)
    Dim TempTL As TYPE_LONG
    Dim TempBL As TYPE_LONG_BYTES
    
    TempTL.Value = Value
    
    LSet TempBL = TempTL
    
    H = TempBL.H
    L = TempBL.L
    
End Sub

Public Function DirGraficos() As String
    DirGraficos = App.path & "\" & Config_Inicio.DirGraficos & "\"
End Function

Public Function DirSound() As String
    DirSound = App.path & "\" & Config_Inicio.DirSonidos & "\"
End Function

Public Function DirMidi() As String
    DirMidi = App.path & "\" & Config_Inicio.DirMusica & "\"
End Function

Public Function DirMapas() As String
    DirMapas = App.path & "\" & Config_Inicio.DirMapas & "\"
End Function

Public Function DirExtras() As String
    DirExtras = App.path & "\EXTRAS\"
End Function

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'Initialize randomizer
    Randomize Timer
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
End Function

Public Function GetRawName(ByRef sName As String) As String
'***************************************************
'Author: ZaMa
'Last Modify Date: 13/01/2010
'Last Modified By: -
'Returns the char name without the clan name (if it has it).
'***************************************************

    Dim Pos As Integer
    
    Pos = InStr(1, sName, "<")
    
    If Pos > 0 Then
        GetRawName = Trim(Left(sName, Pos - 1))
    Else
        GetRawName = sName
    End If

End Function

Sub CargarAnimArmas()
On Error Resume Next

    Dim loopC As Long
    Dim arch As String
    
    arch = App.path & "\init\" & "armas.dat"
    
    NumWeaponAnims = Val(GetVar(arch, "INIT", "NumArmas"))
    
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    
    For loopC = 1 To NumWeaponAnims
        InitGrh WeaponAnimData(loopC).WeaponWalk(1), Val(GetVar(arch, "ARMA" & loopC, "Dir1")), 0
        InitGrh WeaponAnimData(loopC).WeaponWalk(2), Val(GetVar(arch, "ARMA" & loopC, "Dir2")), 0
        InitGrh WeaponAnimData(loopC).WeaponWalk(3), Val(GetVar(arch, "ARMA" & loopC, "Dir3")), 0
        InitGrh WeaponAnimData(loopC).WeaponWalk(4), Val(GetVar(arch, "ARMA" & loopC, "Dir4")), 0
    Next loopC
End Sub

Sub CargarColores()
On Error Resume Next
    Dim archivoC As String
    
    archivoC = App.path & "\init\colores.dat"
    
    If Not FileExist(archivoC, vbArchive) Then
'TODO : Si hay que reinstalar, porque no cierra???
        Call MsgBox("ERROR: no se ha podido cargar los colores. Falta el archivo colores.dat, reinstale el juego", vbCritical + vbOKOnly)
        Exit Sub
    End If
    
    Dim i As Long
    
    For i = 0 To 46 '49 y 50 reservados para ciudadano y criminal
        ColoresPJ(i) = D3DColorXRGB(CInt(GetVar(archivoC, CStr(i), "R")), CInt(GetVar(archivoC, CStr(i), "G")), CInt(GetVar(archivoC, CStr(i), "B")))
    Next i
    
    ' Crimi
    ColoresPJ(50) = D3DColorXRGB(CInt(GetVar(archivoC, "CR", "R")), _
    CInt(GetVar(archivoC, "CR", "G")), _
    CInt(GetVar(archivoC, "CR", "B")))
    
    ' Ciuda
    ColoresPJ(49) = D3DColorXRGB(CInt(GetVar(archivoC, "CI", "R")), _
    CInt(GetVar(archivoC, "CI", "G")), _
    CInt(GetVar(archivoC, "CI", "B")))
    
    ' Neutral
    ColoresPJ(48) = D3DColorXRGB(CInt(GetVar(archivoC, "NE", "R")), _
    CInt(GetVar(archivoC, "NE", "G")), _
    CInt(GetVar(archivoC, "NE", "B")))
    
    ColoresPJ(47) = D3DColorXRGB(CInt(GetVar(archivoC, "NW", "R")), _
    CInt(GetVar(archivoC, "NW", "G")), _
    CInt(GetVar(archivoC, "NW", "B")))
    
End Sub

Sub CargarAnimEscudos()
On Error Resume Next

    Dim loopC As Long
    Dim arch As String
    
    arch = App.path & "\init\" & "escudos.dat"
    
    NumEscudosAnims = Val(GetVar(arch, "INIT", "NumEscudos"))
    
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    
    For loopC = 1 To NumEscudosAnims
        InitGrh ShieldAnimData(loopC).ShieldWalk(1), Val(GetVar(arch, "ESC" & loopC, "Dir1")), 0
        InitGrh ShieldAnimData(loopC).ShieldWalk(2), Val(GetVar(arch, "ESC" & loopC, "Dir2")), 0
        InitGrh ShieldAnimData(loopC).ShieldWalk(3), Val(GetVar(arch, "ESC" & loopC, "Dir3")), 0
        InitGrh ShieldAnimData(loopC).ShieldWalk(4), Val(GetVar(arch, "ESC" & loopC, "Dir4")), 0
    Next loopC
End Sub

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal Red As Integer = -1, Optional ByVal Green As Integer, Optional ByVal blue As Integer, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, Optional ByVal bCrLf As Boolean = True)
'******************************************
'Adds text to a Richtext box at the bottom.
'Automatically scrolls to new text.
'Text box MUST be multiline and have a 3D
'apperance!
'Pablo (ToxicWaste) 01/26/2007 : Now the list refeshes properly.
'Juan Martín Sotuyo Dodero (Maraxus) 03/29/2007 : Replaced ToxicWaste's code for extra performance.
'******************************************r
    With RichTextBox
        If Len(.Text) > 1000 Then
            'Get rid of first line
            .SelStart = InStr(1, .Text, vbCrLf) + 1
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF
        End If
        
        .SelStart = Len(.Text)
        .SelLength = 0
        .SelBold = bold
        .SelItalic = italic
        
        If Not Red = -1 Then .SelColor = RGB(Red, Green, blue)
        
        If bCrLf And Len(.Text) > 0 Then Text = vbCrLf & Text
        .SelText = Text
        
        RichTextBox.Refresh
    End With
End Sub

'TODO : Never was sure this is really necessary....
'TODO : 08/03/2006 - (AlejoLp) Esto hay que volarlo...
Public Sub RefreshAllChars()
'*****************************************************************
'Goes through the charlist and replots all the characters on the map
'Used to make sure everyone is visible
'*****************************************************************
    Dim loopC As Long
    
    For loopC = 1 To LastChar
        If charlist(loopC).Active = 1 Then
            MapData(charlist(loopC).Pos.X, charlist(loopC).Pos.Y).CharIndex = loopC
        End If
    Next loopC
End Sub

Sub SaveGameini()
    'Grabamos los datos del usuario en el Game.ini
    Config_Inicio.Name = "BetaTester"
    Config_Inicio.Password = "DammLamers"
    Config_Inicio.Puerto = UserPort
    
    Call EscribirGameIni(Config_Inicio)
End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
    Dim car As Byte
    Dim i As Long
    
    cad = LCase$(cad)
    
    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
        
        If ((car < 97 Or car > 122) Or car = Asc("º")) And (car <> 255) And (car <> 32) Then
            Exit Function
        End If
    Next i
    
    AsciiValidos = True
End Function

Function CheckUserData(ByVal checkemail As Boolean) As Boolean
    'Validamos los datos del user
    Dim loopC As Long
    Dim CharAscii As Integer
    
    If checkemail And UserEmail = "" Then
        MsgBox ("Dirección de email invalida")
        Exit Function
    End If
    
    If UserPassword = "" Then
        MsgBox ("Ingrese un password.")
        Exit Function
    End If
    
    For loopC = 1 To Len(UserPassword)
        CharAscii = Asc(mid$(UserPassword, loopC, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Password inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next loopC
    
    If UserName = "" Then
        MsgBox ("Ingrese un nombre de personaje.")
        Exit Function
    End If
    
    If Len(UserName) > 30 Then
        MsgBox ("El nombre debe tener menos de 30 letras.")
        Exit Function
    End If
    
    For loopC = 1 To Len(UserName)
        CharAscii = Asc(mid$(UserName, loopC, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Nombre inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next loopC
    
    CheckUserData = True
End Function

Sub UnloadAllForms()
On Error Resume Next

    Dim mifrm As Form
    
    For Each mifrm In Forms
        Unload mifrm
    Next
End Sub

Function LegalCharacter(ByVal KeyAscii As Integer) As Boolean
'*****************************************************************
'Only allow characters that are Win 95 filename compatible
'*****************************************************************
    'if backspace allow
    If KeyAscii = 8 Then
        LegalCharacter = True
        Exit Function
    End If
    
    'Only allow space, numbers, letters and special characters
    If KeyAscii < 32 Or KeyAscii = 44 Then
        Exit Function
    End If
    
    If KeyAscii > 126 Then
        Exit Function
    End If
    
    'Check for bad special characters in between
    If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
        Exit Function
    End If
    
    'else everything is cool
    LegalCharacter = True
End Function

Sub SetConnected()
'*****************************************************************
'Sets the client to "Connect" mode
'*****************************************************************
    'Set Connected
    Connected = True
    
    Call SaveGameini
    
    'Unload the connect form
    Unload frmConnect
    
    frmMain.lblName = UserName
    
    Call mod_Components.ClearComponents

    frmMain.Visible = True
        
    FPSFLAG = True
End Sub

Sub CargarTip()
    Dim N As Integer
    N = RandomNumber(1, UBound(Tips))
    
    frmtip.tip.Caption = Tips(N)
End Sub

Sub MoveTo(ByVal Direccion As E_Heading)
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/28/2008
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
' 06/03/2006: AlejoLp - Elimine las funciones Move[NSWE] y las converti a esta
' 12/08/2007: Tavo    - Si el usuario esta paralizado no se puede mover.
' 06/28/2008: NicoNZ - Saqué lo que impedía que si el usuario estaba paralizado se ejecute el sub.
'***************************************************
    Dim LegalOk As Boolean
    
    If Cartel Then Cartel = False
    
    Select Case Direccion
        Case E_Heading.NORTH
            LegalOk = MoveToLegalPos(ViewPositionX, ViewPositionY - 1)
        Case E_Heading.EAST
            LegalOk = MoveToLegalPos(ViewPositionX + 1, ViewPositionY)
        Case E_Heading.SOUTH
            LegalOk = MoveToLegalPos(ViewPositionX, ViewPositionY + 1)
        Case E_Heading.WEST
            LegalOk = MoveToLegalPos(ViewPositionX - 1, ViewPositionY)
    End Select
    
    If LegalOk And Not UserParalizado Then
        Call WriteWalk(Direccion)
        If Not UserDescansar And Not UserMeditar Then
            MoveCharbyHead UserCharIndex, Direccion
            MoveScreen Direccion
        End If
    Else
        If charlist(UserCharIndex).Heading <> Direccion Then
            Call WriteChangeHeading(Direccion)
        End If
    End If
    
    If frmMain.macrotrabajo.Enabled Then Call frmMain.DesactivarMacroTrabajo
    
    ' Update 3D sounds!
    Call Audio.MoveListener(ViewPositionX, ViewPositionY)
End Sub

Sub RandomMove()
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/03/2006
' 06/03/2006: AlejoLp - Ahora utiliza la funcion MoveTo
'***************************************************
    Call MoveTo(RandomNumber(NORTH, WEST))
End Sub

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
'*****************************************************************
'Gets a field from a delimited string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/15/2004
'*****************************************************************
    Dim i As Long
    Dim lastPos As Long
    Dim CurrentPos As Long
    Dim delimiter As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        lastPos = CurrentPos
        CurrentPos = InStr(lastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        ReadField = mid$(Text, lastPos + 1, Len(Text) - lastPos)
    Else
        ReadField = mid$(Text, lastPos + 1, CurrentPos - lastPos - 1)
    End If
End Function

Function FieldCount(ByRef Text As String, ByVal SepASCII As Byte) As Long
'*****************************************************************
'Gets the number of fields in a delimited string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 07/29/2007
'*****************************************************************
    Dim Count As Long
    Dim curPos As Long
    Dim delimiter As String * 1
    
    If LenB(Text) = 0 Then Exit Function
    
    delimiter = Chr$(SepASCII)
    
    curPos = 0
    
    Do
        curPos = InStr(curPos + 1, Text, delimiter)
        Count = Count + 1
    Loop While curPos <> 0
    
    FieldCount = Count
End Function

Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(file, FileType) <> "")
End Function

Sub WriteClientVer()
    Dim hFile As Integer
        
    hFile = FreeFile()
    Open App.path & "\init\Ver.bin" For Binary Access Write Lock Read As #hFile
    Put #hFile, , CLng(777)
    Put #hFile, , CLng(777)
    Put #hFile, , CLng(777)
    
    Put #hFile, , CInt(App.Major)
    Put #hFile, , CInt(App.Minor)
    Put #hFile, , CInt(App.Revision)
    
    Close #hFile
End Sub

Sub Main()
    Call WriteClientVer
    Call InitColours
        
    IniPath = App.path & "\Init\"
    
    'Load config file
    If FileExist(IniPath & "Inicio.con", vbNormal) Then
        Config_Inicio = LeerGameIni()
    End If
    
    CurServerIP = "127.0.0.1"
    CurServerPort = 7666
    
    'Load ao.dat config file
    Call LoadClientSetup

    
    'Read command line. Do it AFTER config file is loaded to prevent this from
    'canceling the effects of "/nores" option.
    Call LeerLineaComandos
    
    'usaremos esto para ayudar en los parches
    Call SaveSetting("ArgentumOnlineCliente", "Init", "Path", App.path & "\")
    
    ChDrive App.path
    ChDir App.path

    MD5HushYo = "0123456789abcdef"  'We aren't using a real MD5
    
    tipf = Config_Inicio.tip
    
    'Set resolution BEFORE the loading form is displayed, therefore it will be centered.
    'Call Resolution.SetResolution
    
    ' Mouse Pointer (Loaded before opening any form with buttons in it)
    If FileExist(DirExtras & "Hand.ico", vbArchive) Then _
        Set picMouseIcon = LoadPicture(DirExtras & "Hand.ico")
    
    frmCargando.Show
    frmCargando.Refresh
    
    'frmConnect.version = "v" & App.Major & "." & App.Minor & " Build: " & App.Revision

    Call AddtoRichTextBox(frmCargando.Status, "Iniciando constantes... ", 255, 255, 255, True, False, True)
    
    Call InicializarNombres
    
    ' Initialize FONTTYPES
    Call Protocol.InitFonts

    Call EstablecerRecompensas
    
    Dim i As Long
    Dim SearchVar As String
    
    For i = 1 To NUMRAZAS
        With ModRaza(i)
            SearchVar = Replace(ListaRazas(i), " ", "")
        
            .Fuerza = Val(GetVar(IniPath & "CharInfo.dat", "MODRAZA", SearchVar + "Fuerza"))
            .Agilidad = Val(GetVar(IniPath & "CharInfo.dat", "MODRAZA", SearchVar + "Agilidad"))
            .Inteligencia = Val(GetVar(IniPath & "CharInfo.dat", "MODRAZA", SearchVar + "Inteligencia"))
            .Carisma = Val(GetVar(IniPath & "CharInfo.dat", "MODRAZA", SearchVar + "Carisma"))
            .Constitucion = Val(GetVar(IniPath & "CharInfo.dat", "MODRAZA", SearchVar + "Constitucion"))
        End With
    Next i
    
    Call AddtoRichTextBox(frmCargando.Status, "Hecho", 255, 0, 0, True, False, False)
    
    Call AddtoRichTextBox(frmCargando.Status, "Iniciando motor gráfico... ", 255, 255, 255, True, False, True)
    
    prgRun = True
    
    If Not InitTileEngine(frmMain.hwnd, 32, 32, 17, 23, 7, 8, 8, 0.018, 0.03) Then
        Call CloseClient
    End If
    
    'Call AddtoRichTextBox(frmCargando.Status, "Hecho", 255, 0, 0, True, False, False)
    
    Call AddtoRichTextBox(frmCargando.Status, "Creando animaciones extra... ", 255, 255, 255, True, False, True)
    
    Call CargarTips
    
UserMap = 1
    
    Call CargarArrayLluvia
    Call CargarAnimArmas
    Call CargarAnimEscudos
    Call CargarColores
    
    Call AddtoRichTextBox(frmCargando.Status, "Hecho", 255, 0, 0, True, False, False)
    
    Call AddtoRichTextBox(frmCargando.Status, "Iniciando DirectSound... ", 255, 255, 255, True, False, True)
    
    'Inicializamos el sonido
    Call Audio.Initialize(DirectX, frmMain.hwnd, App.path & "\" & Config_Inicio.DirSonidos & "\", App.path & "\" & Config_Inicio.DirMusica & "\")
    'Enable / Disable audio
    Audio.MusicActivated = Not ClientSetup.bNoMusic
    Audio.SoundActivated = Not ClientSetup.bNoSound
    Audio.SoundEffectsActivated = Not ClientSetup.bNoSoundEffects
    
    'Inicializamos el inventario gráfico
    Call Inventario.Initialize(frmMain.picInv, MAX_INVENTORY_SLOTS)
    
    Call Audio.MusicMP3Play(App.path & "\MP3\" & MP3_Inicio & ".mp3")
    
    Call AddtoRichTextBox(frmCargando.Status, "Hecho", 255, 0, 0, True, False, False)
    
    Call AddtoRichTextBox(frmCargando.Status, "                    ¡Bienvenido a Argentum Online!", 255, 255, 255, True, False, True)
    
    'Give the user enough time to read the welcome text
    Call Sleep(500)
    
    Unload frmCargando
        
    frmMain.Socket1.Startup

    frmConnect.MousePointer = vbDefault
    frmConnect.Visible = True
    
    'Inicialización de variables globales
    PrimeraVez = True
    pausa = False
    
    'Set the intervals of timers
    Call MainTimer.SetInterval(TimersIndex.Attack, INT_ATTACK)
    Call MainTimer.SetInterval(TimersIndex.Work, INT_WORK)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithU, INT_USEITEMU)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithDblClick, INT_USEITEMDCK)
    Call MainTimer.SetInterval(TimersIndex.SendRPU, INT_SENTRPU)
    Call MainTimer.SetInterval(TimersIndex.CastSpell, INT_CAST_SPELL)
    Call MainTimer.SetInterval(TimersIndex.Arrows, INT_ARROWS)
    Call MainTimer.SetInterval(TimersIndex.CastAttack, INT_CAST_ATTACK)
    
    frmMain.macrotrabajo.Interval = INT_MACRO_TRABAJO
    frmMain.macrotrabajo.Enabled = False
    
   'Init timers
    Call MainTimer.Start(TimersIndex.Attack)
    Call MainTimer.Start(TimersIndex.Work)
    Call MainTimer.Start(TimersIndex.UseItemWithU)
    Call MainTimer.Start(TimersIndex.UseItemWithDblClick)
    Call MainTimer.Start(TimersIndex.SendRPU)
    Call MainTimer.Start(TimersIndex.CastSpell)
    Call MainTimer.Start(TimersIndex.Arrows)
    Call MainTimer.Start(TimersIndex.CastAttack)
    
    'Set the dialog's font
    Dialogos.Font = frmMain.Font
    'DialogosClanes.font = frmMain.font
    
    lFrameTimer = GetTickCount
    
    ' Load the form for screenshots
    Call Load(frmScreenshots)
    
    Call SwitchMap(1)
    
    Call Time_Update
    
    Do While prgRun
    
        Call Render

        'FPS Counter - mostramos las FPS
        If GetTickCount - lFrameTimer >= 1000 Then
            If FPSFLAG Then frmMain.lblFPS.Caption = Mod_TileEngine.FPS
            
            lFrameTimer = GetTickCount
        End If
        
        ' If there is anything to be sent, we send it
        Call FlushBuffer
        
        DoEvents
    Loop
    
    Call CloseClient
End Sub

Sub WriteVar(ByVal file As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, Var, Value, file
End Sub

Function GetVar(ByVal file As String, ByVal Main As String, ByVal Var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************
    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(500) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    getprivateprofilestring Main, Var, vbNullString, sSpaces, Len(sSpaces), file
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

'[CODE 002]:MatuX
'
'  Función para chequear el email
'
'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba y evitar un chequeo innecesario
Public Function CheckMailString(ByVal sString As String) As Boolean
On Error GoTo errHnd
    Dim lPos  As Long
    Dim lX    As Long
    Dim iAsc  As Integer
    
    '1er test: Busca un simbolo @
    lPos = InStr(sString, "@")
    If (lPos <> 0) Then
        '2do test: Busca un simbolo . después de @ + 1
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then _
            Exit Function
        
        '3er test: Recorre todos los caracteres y los valída
        For lX = 0 To Len(sString) - 1
            If Not (lX = (lPos - 1)) Then   'No chequeamos la '@'
                iAsc = Asc(mid$(sString, (lX + 1), 1))
                If Not CMSValidateChar_(iAsc) Then _
                    Exit Function
            End If
        Next lX
        
        'Finale
        CheckMailString = True
    End If
errHnd:
End Function

'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or _
                        (iAsc >= 65 And iAsc <= 90) Or _
                        (iAsc >= 97 And iAsc <= 122) Or _
                        (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)
End Function

'TODO : como todo lo relativo a mapas, no tiene nada que hacer acá....
Function HayAgua(ByVal X As Integer, ByVal Y As Integer) As Boolean
    HayAgua = ((MapData(X, Y).Graphic(1).GrhIndex >= 1505 And MapData(X, Y).Graphic(1).GrhIndex <= 1520) Or _
            (MapData(X, Y).Graphic(1).GrhIndex >= 5665 And MapData(X, Y).Graphic(1).GrhIndex <= 5680) Or _
            (MapData(X, Y).Graphic(1).GrhIndex >= 13547 And MapData(X, Y).Graphic(1).GrhIndex <= 13562)) And _
                MapData(X, Y).Graphic(2).GrhIndex = 0
                
End Function

Public Sub ShowSendTxt()
    If Not frmCantidad.Visible Then
        'todo
        frmMain.SendTxt.Visible = True
        frmMain.SendTxt.SetFocus
    End If
End Sub

Public Sub ShowSendCMSGTxt()
    If Not frmCantidad.Visible Then
    'todo
        'frmMain.SendCMSTXT.Visible = True
        'frmMain.SendCMSTXT.SetFocus
    End If
End Sub

''
' Checks the command line parameters, if you are running Ao with /nores command and checks the AoUpdate parameters
'
'

Public Sub LeerLineaComandos()
'*************************************************
'Author: Unknown
'Last modified: 25/11/2008 (BrianPr)
'
'*************************************************
    Dim T() As String
    Dim i As Long
    
    Dim UpToDate As Boolean

    'Parseo los comandos
    T = Split(Command, " ")
    For i = LBound(T) To UBound(T)
        Select Case UCase$(T(i))
            Case "/NORES" 'no cambiar la resolucion
                NoRes = True
            Case "/UPTODATE"
                UpToDate = True
        End Select
    Next i
    
    'Call AoUpdate(UpToDate, NoRes) ' www.gs-zone.org
End Sub

Private Sub LoadClientSetup()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/19/09
'11/19/09: Pato - Is optional show the frmGuildNews form
'**************************************************************
    Dim fHandle As Integer
    
    If FileExist(App.path & "\init\ao.dat", vbArchive) Then
        fHandle = FreeFile
        
        Open App.path & "\init\ao.dat" For Binary Access Read Lock Write As fHandle
            Get fHandle, , ClientSetup
        Close fHandle
    Else
        'Use dynamic by default
        ClientSetup.bDinamic = True
    End If
    
    NoRes = ClientSetup.bNoRes
    
    If InStr(1, ClientSetup.sGraficos, "Graficos") Then
        GraphicsFile = ClientSetup.sGraficos
    Else
        GraphicsFile = "Graficos3.ind"
    End If
    
 '   ClientSetup.bGuildNews = Not ClientSetup.bGuildNews
  '  DialogosClanes.Activo = Not ClientSetup.bGldMsgConsole
  '  DialogosClanes.CantidadDialogos = ClientSetup.bCantMsgs
End Sub

Private Sub SaveClientSetup()
'**************************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 03/11/10
'
'**************************************************************
    Dim fHandle As Integer
    
    fHandle = FreeFile
    
    'ClientSetup.bNoMusic = Not audio.MusicActivated
    'ClientSetup.bNoSound = Not audio.SoundActivated
    'ClientSetup.bNoSoundEffects = Not audio.SoundEffectsActivated
   ' ClientSetup.bGuildNews = Not ClientSetup.bGuildNews
  '  ClientSetup.bGldMsgConsole = Not DialogosClanes.Activo
    'ClientSetup.bCantMsgs = DialogosClanes.CantidadDialogos
    
    Open App.path & "\init\ao.dat" For Binary As fHandle
        Put fHandle, , ClientSetup
    Close fHandle
End Sub

Private Sub InicializarNombres()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Inicializa los nombres de razas, ciudades, clases, skills, atributos, etc.
'**************************************************************
    Ciudades(eCiudad.cUllathorpe) = "Ullathorpe"
    Ciudades(eCiudad.cNix) = "Nix"
    Ciudades(eCiudad.cBanderbill) = "Banderbill"
    Ciudades(eCiudad.cLindos) = "Lindos"
    Ciudades(eCiudad.cArghal) = "Arghâl"
    
    ListaRazas(eRaza.Humano) = "Humano"
    ListaRazas(eRaza.Elfo) = "Elfo"
    ListaRazas(eRaza.ElfoOscuro) = "Elfo Oscuro"
    ListaRazas(eRaza.Gnomo) = "Gnomo"
    ListaRazas(eRaza.Enano) = "Enano"

    ListaClases(eClass.Ciudadano) = "Ciudadano"
    ListaClases(eClass.Trabajador) = "Trabajador"
    ListaClases(eClass.Experto_Minerales) = "Experto en minerales"
    ListaClases(eClass.MINERO) = "Minero"
    ListaClases(eClass.HERRERO) = "Herrero"
    ListaClases(eClass.Experto_Madera) = "Experto en uso de madera"
    ListaClases(eClass.TALADOR) = "Leñador"
    ListaClases(eClass.CARPINTERO) = "Carpintero"
    ListaClases(eClass.PESCADOR) = "Pescador"
    ListaClases(eClass.Sastre) = "Sastre"
    ListaClases(eClass.Alquimista) = "Alquimista"
    ListaClases(eClass.Luchador) = "Luchador"
    ListaClases(eClass.Con_Mana) = "Con uso de mana"
    ListaClases(eClass.Hechicero) = "Hechicero"
    ListaClases(eClass.MAGO) = "Mago"
    ListaClases(eClass.NIGROMANTE) = "Nigromante"
    ListaClases(eClass.Orden_Sagrada) = "Orden sagrada"
    ListaClases(eClass.PALADIN) = "Paladin"
    ListaClases(eClass.CLERIGO) = "Clerigo"
    ListaClases(eClass.Naturalista) = "Naturalista"
    ListaClases(eClass.BARDO) = "Bardo"
    ListaClases(eClass.DRUIDA) = "Druida"
    ListaClases(eClass.Sigiloso) = "Sigiloso"
    ListaClases(eClass.ASESINO) = "Asesino"
    ListaClases(eClass.CAZADOR) = "Cazador"
    ListaClases(eClass.Sin_Mana) = "Sin uso de mana"
    ListaClases(eClass.ARQUERO) = "Arquero"
    ListaClases(eClass.GUERRERO) = "Guerrero"
    ListaClases(eClass.Caballero) = "Caballero"
    ListaClases(eClass.Bandido) = "Bandido"
    ListaClases(eClass.PIRATA) = "Pirata"
    ListaClases(eClass.LADRON) = "Ladron"
    
    SkillsNames(eSkill.Magia) = "Magia"
    SkillsNames(eSkill.Robar) = "Robar"
    SkillsNames(eSkill.Tacticas) = "Evasión en combate"
    SkillsNames(eSkill.Armas) = "Combate cuerpo a cuerpo"
    SkillsNames(eSkill.Meditar) = "Meditar"
    SkillsNames(eSkill.Apuñalar) = "Apuñalar"
    SkillsNames(eSkill.Ocultarse) = "Ocultarse"
    SkillsNames(eSkill.Supervivencia) = "Supervivencia"
    SkillsNames(eSkill.Talar) = "Talar árboles"
    SkillsNames(eSkill.Comerciar) = "Comercio"
    SkillsNames(eSkill.Defensa) = "Defensa con escudos"
    SkillsNames(eSkill.Pesca) = "Pesca"
    SkillsNames(eSkill.Mineria) = "Mineria"
    SkillsNames(eSkill.Carpinteria) = "Carpinteria"
    SkillsNames(eSkill.Herreria) = "Herreria"
    SkillsNames(eSkill.Liderazgo) = "Liderazgo"
    SkillsNames(eSkill.Domar) = "Domar animales"
    SkillsNames(eSkill.Proyectiles) = "Combate a distancia"
    SkillsNames(eSkill.Wrestling) = "Combate sin armas"
    SkillsNames(eSkill.Navegacion) = "Navegacion"
    SkillsNames(eSkill.Sastreria) = "Sastrería"
    SkillsNames(eSkill.Resis) = "Resistencia Mágica"
    
    AtributosNames(eAtributos.Fuerza) = "Fuerza"
    AtributosNames(eAtributos.Agilidad) = "Agilidad"
    AtributosNames(eAtributos.Inteligencia) = "Inteligencia"
    AtributosNames(eAtributos.Carisma) = "Carisma"
    AtributosNames(eAtributos.Constitucion) = "Constitucion"
End Sub

''
' Removes all text from the console and dialogs

Public Sub CleanDialogs()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Removes all text from the console and dialogs
'**************************************************************
    'Clean console and dialogs
    
    frmMain.RecTxt.Text = vbNullString
    
  '  Call DialogosClanes.RemoveDialogs
    
    Call Dialogos.RemoveAllDialogs
End Sub

Public Sub CloseClient()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 8/14/2007
'Frees all used resources, cleans up and leaves
'**************************************************************
    ' Allow new instances of the client to be opened
    Call PrevInstance.ReleaseInstance
    
    EngineRun = False
    frmCargando.Show
    Call AddtoRichTextBox(frmCargando.Status, "Liberando recursos...", 0, 0, 0, 0, 0, 0)
    
    'Stop tile engine
    Call DeinitTileEngine
    
    Call SaveClientSetup
    
    'Destruimos los objetos públicos creados
    Set CustomMessages = Nothing
    Set CustomKeys = Nothing

    Set Dialogos = Nothing
  '  Set DialogosClanes = Nothing
    'Set Audio = Nothing
    Set Inventario = Nothing
    Set MainTimer = Nothing
    Set incomingData = Nothing
    Set outgoingData = Nothing
    
    Call UnloadAllForms
    
    'Actualizar tip
    Config_Inicio.tip = tipf
    Call EscribirGameIni(Config_Inicio)
    End
End Sub

Public Function esGM(CharIndex As Integer) As Boolean
esGM = False
If charlist(CharIndex).priv >= 1 And charlist(CharIndex).priv <= 5 Or charlist(CharIndex).priv = 25 Then _
    esGM = True

End Function

Public Function getTagPosition(ByVal Nick As String) As Integer
Dim buf As Integer
buf = InStr(Nick, "<")
If buf > 0 Then
    getTagPosition = buf
    Exit Function
End If
buf = InStr(Nick, "[")
If buf > 0 Then
    getTagPosition = buf
    Exit Function
End If
getTagPosition = Len(Nick) + 2
End Function

Public Function getStrenghtColor() As Long
Dim M As Long
M = 255 / MAXATRIBUTOS
getStrenghtColor = RGB(255 - (M * UserFuerza), (M * UserFuerza), 0)
End Function
Public Function getDexterityColor() As Long
Dim M As Long
M = 255 / MAXATRIBUTOS
getDexterityColor = RGB(255, M * UserAgilidad, 0)
End Function

Public Function getCharIndexByName(ByVal Name As String) As Integer
Dim i As Long
For i = 1 To LastChar
    If charlist(i).Nombre = Name Then
        getCharIndexByName = i
        Exit Function
    End If
Next i
End Function


Public Sub LogError(Desc As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.path & "\errores.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & Desc
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub

Public Function ReadFile(FileName As String, Optional size As Long = -1) As Byte()

    Dim wFile As Integer

    wFile = FreeFile
    Open FileName For Binary Access Read As wFile

    If LOF(wFile) > 0 Then
        
        size = LOF(wFile)
        ReDim ReadFile(0 To LOF(wFile) - 1)
        Get wFile, , ReadFile

    End If

    Close #wFile

End Function

Public Sub EstablecerRecompensas()

ReDim Recompensas(1 To NUMCLASES, 1 To 3, 1 To 2) As tRecompensa

Recompensas(eClass.MINERO, 1, 1).Name = "Fortaleza del Trabajador"
Recompensas(eClass.MINERO, 1, 1).Descripcion = "Aumenta la vida en 120 puntos."

Recompensas(eClass.MINERO, 1, 2).Name = "Suerte de Novato"
Recompensas(eClass.MINERO, 1, 2).Descripcion = "Al morir hay 20% de probabilidad de no perder los minerales."

Recompensas(eClass.MINERO, 2, 1).Name = "Destrucción Mágica"
Recompensas(eClass.MINERO, 2, 1).Descripcion = "Inmunidad al paralisis lanzado por otros usuarios."

Recompensas(eClass.MINERO, 2, 2).Name = "Pica Fuerte"
Recompensas(eClass.MINERO, 2, 2).Descripcion = "Permite minar 20% más cantidad de hierro y la plata."

Recompensas(eClass.MINERO, 3, 1).Name = "Gremio del Trabajador"
Recompensas(eClass.MINERO, 3, 1).Descripcion = "Permite minar 20% más cantidad de oro."

Recompensas(eClass.MINERO, 3, 2).Name = "Pico de la Suerte"
Recompensas(eClass.MINERO, 3, 2).Descripcion = "Al morir hay 30% de probabilidad de que no perder los minerales (acumulativo con Suerte de Novato.)"


Recompensas(eClass.HERRERO, 1, 1).Name = "Yunque Rojizo"
Recompensas(eClass.HERRERO, 1, 1).Descripcion = "25% de probabilidad de gastar la mitad de lingotes en la creación de objetos (Solo aplicable a armas y armaduras)."

Recompensas(eClass.HERRERO, 1, 2).Name = "Maestro de la Forja"
Recompensas(eClass.HERRERO, 1, 2).Descripcion = "Reduce los costos de cascos y escudos a un 50%."

Recompensas(eClass.HERRERO, 2, 1).Name = "Experto en Filos"
Recompensas(eClass.HERRERO, 2, 1).Descripcion = "Permite crear las mejores armas (Espada Neithan, Espada Neithan + 1, Espada de Plata + 1 y Daga Infernal)."

Recompensas(eClass.HERRERO, 2, 2).Name = "Experto en Corazas"
Recompensas(eClass.HERRERO, 2, 2).Descripcion = "Permite crear las mejores armaduras (Armaduras de las Tinieblas, Armadura Legendaria y Armaduras del Dragón)."

Recompensas(eClass.HERRERO, 3, 1).Name = "Fundir Metal"
Recompensas(eClass.HERRERO, 3, 1).Descripcion = "Reduce a un 50% la cantidad de lingotes utilizados en fabricación de Armas y Armaduras (acumulable con Yunque Rojizo)."

Recompensas(eClass.HERRERO, 3, 2).Name = "Trabajo en Serie"
Recompensas(eClass.HERRERO, 3, 2).Descripcion = "10% de probabilidad de crear el doble de objetos de los asignados con la misma cantidad de lingotes."


Recompensas(eClass.TALADOR, 1, 1).Name = "Músculos Fornidos"
Recompensas(eClass.TALADOR, 1, 1).Descripcion = "Permite talar 20% más cantidad de madera."

Recompensas(eClass.TALADOR, 1, 2).Name = "Tiempos de Calma"
Recompensas(eClass.TALADOR, 1, 2).Descripcion = "Evita tener hambre y sed."


Recompensas(eClass.CARPINTERO, 1, 1).Name = "Experto en Arcos"
Recompensas(eClass.CARPINTERO, 1, 1).Descripcion = "Permite la creación de los mejores arcos (Élfico y de las Tinieblas)."

Recompensas(eClass.CARPINTERO, 1, 2).Name = "Experto de Varas"
Recompensas(eClass.CARPINTERO, 1, 2).Descripcion = "Permite la creación de las mejores varas (Engarzadas)."

Recompensas(eClass.CARPINTERO, 2, 1).Name = "Fila de Leña"
Recompensas(eClass.CARPINTERO, 2, 1).Descripcion = "Aumenta la creación de flechas a 20 por vez."

Recompensas(eClass.CARPINTERO, 2, 2).Name = "Espíritu de Navegante"
Recompensas(eClass.CARPINTERO, 2, 2).Descripcion = "Reduce en un 20% el coste de madera de las barcas."


Recompensas(eClass.PESCADOR, 1, 1).Name = "Favor de los Dioses"
Recompensas(eClass.PESCADOR, 1, 1).Descripcion = "Pescar 20% más cantidad de pescados."

Recompensas(eClass.PESCADOR, 1, 2).Name = "Pesca en Alta Mar"
Recompensas(eClass.PESCADOR, 1, 2).Descripcion = "Al pescar en barca hay 10% de probabilidad de obtener pescados más caros."


Recompensas(eClass.MAGO, 1, 1).Name = "Pociones de Espíritu"
Recompensas(eClass.MAGO, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

Recompensas(eClass.MAGO, 1, 2).Name = "Pociones de Vida"
Recompensas(eClass.MAGO, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(eClass.MAGO, 2, 1).Name = "Vitalidad"
Recompensas(eClass.MAGO, 2, 1).Descripcion = "Aumenta la vida en 10 puntos."

Recompensas(eClass.MAGO, 2, 2).Name = "Fortaleza Mental"
Recompensas(eClass.MAGO, 2, 2).Descripcion = "Libera el limite de mana máximo."

Recompensas(eClass.MAGO, 3, 1).Name = "Furia del Relámpago"
Recompensas(eClass.MAGO, 3, 1).Descripcion = "Aumenta el daño base máximo de la Descarga Eléctrica en 10 puntos."

Recompensas(eClass.MAGO, 3, 2).Name = "Destrucción"
Recompensas(eClass.MAGO, 3, 2).Descripcion = "Aumenta el daño base mínimo del Apocalipsis en 10 puntos."

Recompensas(eClass.NIGROMANTE, 1, 1).Name = "Pociones de Espíritu"
Recompensas(eClass.NIGROMANTE, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

Recompensas(eClass.NIGROMANTE, 1, 2).Name = "Pociones de Vida"
Recompensas(eClass.NIGROMANTE, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(eClass.NIGROMANTE, 2, 1).Name = "Vida del Invocador"
Recompensas(eClass.NIGROMANTE, 2, 1).Descripcion = "Aumenta la vida en 15 puntos."

Recompensas(eClass.NIGROMANTE, 2, 2).Name = "Alma del Invocador"
Recompensas(eClass.NIGROMANTE, 2, 2).Descripcion = "Aumenta el mana en 40 puntos."

Recompensas(eClass.NIGROMANTE, 3, 1).Name = "Semillas de las Almas"
Recompensas(eClass.NIGROMANTE, 3, 1).Descripcion = "Aumenta el daño base mínimo de la magia en 10 puntos."

Recompensas(eClass.NIGROMANTE, 3, 2).Name = "Bloqueo de las Almas"
Recompensas(eClass.NIGROMANTE, 3, 2).Descripcion = "Aumenta la evasión en un 5%."


Recompensas(eClass.PALADIN, 1, 1).Name = "Pociones de Espíritu"
Recompensas(eClass.PALADIN, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

Recompensas(eClass.PALADIN, 1, 2).Name = "Pociones de Vida"
Recompensas(eClass.PALADIN, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(eClass.PALADIN, 2, 1).Name = "Aura de Vitalidad"
Recompensas(eClass.PALADIN, 2, 1).Descripcion = "Aumenta la vida en 5 puntos y el mana en 10 puntos."

Recompensas(eClass.PALADIN, 2, 2).Name = "Aura de Espíritu"
Recompensas(eClass.PALADIN, 2, 2).Descripcion = "Aumenta el mana en 30 puntos."

Recompensas(eClass.PALADIN, 3, 1).Name = "Gracia Divina"
Recompensas(eClass.PALADIN, 3, 1).Descripcion = "Reduce el coste de mana de Remover Paralisis a 250 puntos."

Recompensas(eClass.PALADIN, 3, 2).Name = "Favor de los Enanos"
Recompensas(eClass.PALADIN, 3, 2).Descripcion = "Aumenta en 5% la posibilidad de golpear al enemigo con armas cuerpo a cuerpo."

Recompensas(eClass.CLERIGO, 1, 1).Name = "Pociones de Espíritu"
Recompensas(eClass.CLERIGO, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

Recompensas(eClass.CLERIGO, 1, 2).Name = "Pociones de Vida"
Recompensas(eClass.CLERIGO, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(eClass.CLERIGO, 2, 1).Name = "Signo Vital"
Recompensas(eClass.CLERIGO, 2, 1).Descripcion = "Aumenta la vida en 10 puntos."

Recompensas(eClass.CLERIGO, 2, 2).Name = "Espíritu de Sacerdote"
Recompensas(eClass.CLERIGO, 2, 2).Descripcion = "Aumenta el mana en 50 puntos."

Recompensas(eClass.CLERIGO, 3, 1).Name = "Sacerdote Experto"
Recompensas(eClass.CLERIGO, 3, 1).Descripcion = "Aumenta la cura base de Curar Heridas Graves en 20 puntos."

Recompensas(eClass.CLERIGO, 3, 2).Name = "Alzamientos de Almas"
Recompensas(eClass.CLERIGO, 3, 2).Descripcion = "El hechizo de Resucitar cura a las personas con su mana, energía, hambre y sed llenas y cuesta 1.100 de mana."

Recompensas(eClass.BARDO, 1, 1).Name = "Pociones de Espíritu"
Recompensas(eClass.BARDO, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

Recompensas(eClass.BARDO, 1, 2).Name = "Pociones de Vida"
Recompensas(eClass.BARDO, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(eClass.BARDO, 2, 1).Name = "Melodía Vital"
Recompensas(eClass.BARDO, 2, 1).Descripcion = "Aumenta la vida en 10 puntos."

Recompensas(eClass.BARDO, 2, 2).Name = "Melodía de la Meditación"
Recompensas(eClass.BARDO, 2, 2).Descripcion = "Aumenta el mana en 50 puntos."

Recompensas(eClass.BARDO, 3, 1).Name = "Concentración"
Recompensas(eClass.BARDO, 3, 1).Descripcion = "Aumenta la probabilidad de Apuñalar a un 20% (con 100 skill)."

Recompensas(eClass.BARDO, 3, 2).Name = "Melodía Caótica"
Recompensas(eClass.BARDO, 3, 2).Descripcion = "Aumenta el daño base del Apocalipsis y la Descarga Electrica en 5 puntos."


Recompensas(eClass.DRUIDA, 1, 1).Name = "Pociones de Espíritu"
Recompensas(eClass.DRUIDA, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

Recompensas(eClass.DRUIDA, 1, 2).Name = "Pociones de Vida"
Recompensas(eClass.DRUIDA, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(eClass.DRUIDA, 2, 1).Name = "Grifo de la Vida"
Recompensas(eClass.DRUIDA, 2, 1).Descripcion = "Aumenta la vida en 15 puntos."

Recompensas(eClass.DRUIDA, 2, 2).Name = "Poder del Alma"
Recompensas(eClass.DRUIDA, 2, 2).Descripcion = "Aumenta el mana en 40 puntos."

Recompensas(eClass.DRUIDA, 3, 1).Name = "Raíces de la Naturaleza"
Recompensas(eClass.DRUIDA, 3, 1).Descripcion = "Reduce el coste de mana de Inmovilizar a 250 puntos."

Recompensas(eClass.DRUIDA, 3, 2).Name = "Fortaleza Natural"
Recompensas(eClass.DRUIDA, 3, 2).Descripcion = "Aumenta la vida de los elementales invocados en 75 puntos."


Recompensas(eClass.ASESINO, 1, 1).Name = "Pociones de Espíritu"
Recompensas(eClass.ASESINO, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

Recompensas(eClass.ASESINO, 1, 2).Name = "Pociones de Vida"
Recompensas(eClass.ASESINO, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(eClass.ASESINO, 2, 1).Name = "Sombra de Vida"
Recompensas(eClass.ASESINO, 2, 1).Descripcion = "Aumenta la vida en 10 puntos."

Recompensas(eClass.ASESINO, 2, 2).Name = "Sombra Mágica"
Recompensas(eClass.ASESINO, 2, 2).Descripcion = "Aumenta el mana en 30 puntos."

Recompensas(eClass.ASESINO, 3, 1).Name = "Daga Mortal"
Recompensas(eClass.ASESINO, 3, 1).Descripcion = "Aumenta el daño de Apuñalar a un 70% más que el golpe."

Recompensas(eClass.ASESINO, 3, 2).Name = "Punteria mortal"
Recompensas(eClass.ASESINO, 3, 2).Descripcion = "Las chances de apuñalar suben a 25% (Con 100 skills)."


Recompensas(eClass.CAZADOR, 1, 1).Name = "Pociones de Espíritu"
Recompensas(eClass.CAZADOR, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

Recompensas(eClass.CAZADOR, 1, 2).Name = "Pociones de Vida"
Recompensas(eClass.CAZADOR, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(eClass.CAZADOR, 2, 1).Name = "Fortaleza del Oso"
Recompensas(eClass.CAZADOR, 2, 1).Descripcion = "Aumenta la vida en 10 puntos."

Recompensas(eClass.CAZADOR, 2, 2).Name = "Fortaleza del Leviatán"
Recompensas(eClass.CAZADOR, 2, 2).Descripcion = "Aumenta el mana en 50 puntos."

Recompensas(eClass.CAZADOR, 3, 1).Name = "Precisión"
Recompensas(eClass.CAZADOR, 3, 1).Descripcion = "Aumenta la puntería con arco en un 10%."

Recompensas(eClass.CAZADOR, 3, 2).Name = "Tiro Preciso"
Recompensas(eClass.CAZADOR, 3, 2).Descripcion = "Las flechas que golpeen la cabeza ignoran la defensa del casco."


Recompensas(eClass.ARQUERO, 1, 1).Name = "Flechas Mortales"
Recompensas(eClass.ARQUERO, 1, 1).Descripcion = "1.500 flechas que caen al morir."

Recompensas(eClass.ARQUERO, 1, 2).Name = "Pociones de Vida"
Recompensas(eClass.ARQUERO, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(eClass.ARQUERO, 2, 1).Name = "Vitalidad Élfica"
Recompensas(eClass.ARQUERO, 2, 1).Descripcion = "Aumenta la vida en 10 puntos."

Recompensas(eClass.ARQUERO, 2, 2).Name = "Paso Élfico"
Recompensas(eClass.ARQUERO, 2, 2).Descripcion = "Aumenta la evasión en un 5%."

Recompensas(eClass.ARQUERO, 3, 1).Name = "Ojo del Águila"
Recompensas(eClass.ARQUERO, 3, 1).Descripcion = "Aumenta la puntería con arco en un 5%."

Recompensas(eClass.ARQUERO, 3, 2).Name = "Disparo Élfico"
Recompensas(eClass.ARQUERO, 3, 2).Descripcion = "Aumenta el daño base mínimo de las flechas en 5 puntos y el máximo en 3 puntos."


Recompensas(eClass.GUERRERO, 1, 1).Name = "Pociones de Poder"
Recompensas(eClass.GUERRERO, 1, 1).Descripcion = "80 pociones verdes y 100 amarillas que no caen al morir."

Recompensas(eClass.GUERRERO, 1, 2).Name = "Pociones de Vida"
Recompensas(eClass.GUERRERO, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(eClass.GUERRERO, 2, 1).Name = "Vida del Mamut"
Recompensas(eClass.GUERRERO, 2, 1).Descripcion = "Aumenta la vida en 5 puntos."

Recompensas(eClass.GUERRERO, 2, 2).Name = "Piel de Piedra"
Recompensas(eClass.GUERRERO, 2, 2).Descripcion = "Aumenta la defensa permanentemente en 2 puntos."

Recompensas(eClass.GUERRERO, 3, 1).Name = "Cuerda Tensa"
Recompensas(eClass.GUERRERO, 3, 1).Descripcion = "Aumenta la puntería con arco en un 10%."

Recompensas(eClass.GUERRERO, 3, 2).Name = "Resistencia Mágica"
Recompensas(eClass.GUERRERO, 3, 2).Descripcion = "Reduce la duración de la parálisis de un minuto a 45 segundos."


Recompensas(eClass.PIRATA, 1, 1).Name = "Marejada Vital"
Recompensas(eClass.PIRATA, 1, 1).Descripcion = "Aumenta la vida en 20 puntos."

Recompensas(eClass.PIRATA, 1, 2).Name = "Aventurero Arriesgado"
Recompensas(eClass.PIRATA, 1, 2).Descripcion = "Permite entrar a los dungeons independientemente del nivel."

Recompensas(eClass.PIRATA, 2, 1).Name = "Riqueza"
Recompensas(eClass.PIRATA, 2, 1).Descripcion = "10% de probabilidad de no perder los objetos al morir."

Recompensas(eClass.PIRATA, 2, 2).Name = "Escamas del Dragón"
Recompensas(eClass.PIRATA, 2, 2).Descripcion = "Aumenta la vida en 40 puntos."

Recompensas(eClass.PIRATA, 3, 1).Name = "Magia Tabú"
Recompensas(eClass.PIRATA, 3, 1).Descripcion = "Inmunidad a la paralisis."

Recompensas(eClass.PIRATA, 3, 2).Name = "Cuerda de Escape"
Recompensas(eClass.PIRATA, 3, 2).Descripcion = "Permite salir del juego en solo dos segundos."


Recompensas(eClass.LADRON, 1, 1).Name = "Codicia"
Recompensas(eClass.LADRON, 1, 1).Descripcion = "Aumenta en 10% la cantidad de oro robado."

Recompensas(eClass.LADRON, 1, 2).Name = "Manos Sigilosas"
Recompensas(eClass.LADRON, 1, 2).Descripcion = "Aumenta en 5% la probabilidad de robar exitosamente."

Recompensas(eClass.LADRON, 2, 1).Name = "Pies sigilosos"
Recompensas(eClass.LADRON, 2, 1).Descripcion = "Permite moverse mientrás se está oculto."

Recompensas(eClass.LADRON, 2, 2).Name = "Ladrón Experto"
Recompensas(eClass.LADRON, 2, 2).Descripcion = "Permite el robo de objetos (10% de probabilidad)."

Recompensas(eClass.LADRON, 3, 1).Name = "Robo Lejano"
Recompensas(eClass.LADRON, 3, 1).Descripcion = "Permite robar a una distancia de hasta 4 tiles."

Recompensas(eClass.LADRON, 3, 2).Name = "Fundido de Sombra"
Recompensas(eClass.LADRON, 3, 2).Descripcion = "Aumenta en 10% la probabilidad de robar objetos."

End Sub

Public Function Ceiling(ByVal X As Double) As Long
   Ceiling = -Int(X * (-1))
End Function
