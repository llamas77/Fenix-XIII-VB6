Attribute VB_Name = "ES"
'Argentum Online 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
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

Public Sub CargarSpawnList()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim N As Integer, LoopC As Integer
    N = val(GetVar(App.path & "\Dat\Invokar.dat", "INIT", "NumNPCs"))
    ReDim SpawnList(N) As tCriaturasEntrenador
    For LoopC = 1 To N
        SpawnList(LoopC).NpcIndex = val(GetVar(App.path & "\Dat\Invokar.dat", "LIST", "NI" & LoopC))
        SpawnList(LoopC).NpcName = GetVar(App.path & "\Dat\Invokar.dat", "LIST", "NN" & LoopC)
    Next LoopC
    
End Sub

Function EsAdmin(ByVal Name As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim NumWizs As Integer
    Dim WizNum As Integer
    Dim NomB As String
    
    NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "Admines"))
    
    For WizNum = 1 To NumWizs
        NomB = UCase$(GetVar(IniPath & "Server.ini", "Admines", "Admin" & WizNum))
        
        If Left$(NomB, 1) = "*" Or Left$(NomB, 1) = "+" Then NomB = Right$(NomB, Len(NomB) - 1)
        If UCase$(Name) = NomB Then
            EsAdmin = True
            Exit Function
        End If
    Next WizNum
    EsAdmin = False

End Function

Function EsDios(ByVal Name As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim NumWizs As Integer
    Dim WizNum As Integer
    Dim NomB As String
    
    NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "Dioses"))
    For WizNum = 1 To NumWizs
        NomB = UCase$(GetVar(IniPath & "Server.ini", "Dioses", "Dios" & WizNum))
        
        If Left$(NomB, 1) = "*" Or Left$(NomB, 1) = "+" Then NomB = Right$(NomB, Len(NomB) - 1)
        If UCase$(Name) = NomB Then
            EsDios = True
            Exit Function
        End If
    Next WizNum
    EsDios = False
End Function

Function EsSemiDios(ByVal Name As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim NumWizs As Integer
    Dim WizNum As Integer
    Dim NomB As String
    
    NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "SemiDioses"))
    For WizNum = 1 To NumWizs
        NomB = UCase$(GetVar(IniPath & "Server.ini", "SemiDioses", "SemiDios" & WizNum))
        
        If Left$(NomB, 1) = "*" Or Left$(NomB, 1) = "+" Then NomB = Right$(NomB, Len(NomB) - 1)
        If UCase$(Name) = NomB Then
            EsSemiDios = True
            Exit Function
        End If
    Next WizNum
    EsSemiDios = False

End Function

Function EsConsejero(ByVal Name As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim NumWizs As Integer
    Dim WizNum As Integer
    Dim NomB As String
    
    NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "Consejeros"))
    For WizNum = 1 To NumWizs
        NomB = UCase$(GetVar(IniPath & "Server.ini", "Consejeros", "Consejero" & WizNum))
        
        If Left$(NomB, 1) = "*" Or Left$(NomB, 1) = "+" Then NomB = Right$(NomB, Len(NomB) - 1)
        If UCase$(Name) = NomB Then
            EsConsejero = True
            Exit Function
        End If
    Next WizNum
    EsConsejero = False
End Function

Function EsRolesMaster(ByVal Name As String) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim NumWizs As Integer
    Dim WizNum As Integer
    Dim NomB As String
    
    NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "RolesMasters"))
    For WizNum = 1 To NumWizs
        NomB = UCase$(GetVar(IniPath & "Server.ini", "RolesMasters", "RM" & WizNum))
        
        If Left$(NomB, 1) = "*" Or Left$(NomB, 1) = "+" Then NomB = Right$(NomB, Len(NomB) - 1)
        If UCase$(Name) = NomB Then
            EsRolesMaster = True
            Exit Function
        End If
    Next WizNum
    EsRolesMaster = False
End Function


Public Function TxtDimension(ByVal Name As String) As Long
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim N As Integer, cad As String, Tam As Long
    N = FreeFile(1)
    Open Name For Input As #N
    Tam = 0
    Do While Not EOF(N)
        Tam = Tam + 1
        Line Input #N, cad
    Loop
    Close N
    TxtDimension = Tam
End Function

Public Sub CargarForbidenWords()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    ReDim ForbidenNames(1 To TxtDimension(DatPath & "NombresInvalidos.txt"))
    Dim N As Integer, i As Integer
    N = FreeFile(1)
    Open DatPath & "NombresInvalidos.txt" For Input As #N
    
    For i = 1 To UBound(ForbidenNames)
        Line Input #N, ForbidenNames(i)
    Next i
    
    Close N

End Sub

Public Sub CargarHechizos()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

'###################################################
'#               ATENCION PELIGRO                  #
'###################################################
'
'  ¡¡¡¡ NO USAR GetVar PARA LEER Hechizos.dat !!!!
'
'El que ose desafiar esta LEY, se las tendrá que ver
'con migo. Para leer Hechizos.dat se deberá usar
'la nueva clase clsLeerInis.
'
'Alejo
'
'###################################################

On Error GoTo Errhandler

    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando Hechizos."
    
    Dim Hechizo As Integer
    Dim Leer As New clsIniReader
    
    Call Leer.Initialize(DatPath & "Hechizos.dat")
    
    'obtiene el numero de hechizos
    NumeroHechizos = val(Leer.GetValue("INIT", "NumeroHechizos"))
    
    ReDim Hechizos(1 To NumeroHechizos) As tHechizo
    
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumeroHechizos
    frmCargando.cargar.Value = 0
    
    'Llena la lista
    For Hechizo = 1 To NumeroHechizos
        With Hechizos(Hechizo)
            .Nombre = Leer.GetValue("Hechizo" & Hechizo, "Nombre")
            .desc = Leer.GetValue("Hechizo" & Hechizo, "Desc")
            .PalabrasMagicas = Leer.GetValue("Hechizo" & Hechizo, "PalabrasMagicas")
            
            .HechizeroMsg = Leer.GetValue("Hechizo" & Hechizo, "HechizeroMsg")
            .TargetMsg = Leer.GetValue("Hechizo" & Hechizo, "TargetMsg")
            .PropioMsg = Leer.GetValue("Hechizo" & Hechizo, "PropioMsg")
            
            .Tipo = val(Leer.GetValue("Hechizo" & Hechizo, "Tipo"))
            .WAV = val(Leer.GetValue("Hechizo" & Hechizo, "WAV"))
            .FXgrh = val(Leer.GetValue("Hechizo" & Hechizo, "Fxgrh"))
            
            .loops = val(Leer.GetValue("Hechizo" & Hechizo, "Loops"))
            
        '    .Resis = val(Leer.GetValue("Hechizo" & Hechizo, "Resis"))
            
            .SubeHP = val(Leer.GetValue("Hechizo" & Hechizo, "SubeHP"))
            .MinHp = val(Leer.GetValue("Hechizo" & Hechizo, "MinHP"))
            .MaxHp = val(Leer.GetValue("Hechizo" & Hechizo, "MaxHP"))
            
            .SubeMana = val(Leer.GetValue("Hechizo" & Hechizo, "SubeMana"))
            .MiMana = val(Leer.GetValue("Hechizo" & Hechizo, "MinMana"))
            .MaMana = val(Leer.GetValue("Hechizo" & Hechizo, "MaxMana"))
            
            .SubeSta = val(Leer.GetValue("Hechizo" & Hechizo, "SubeSta"))
            .MinSta = val(Leer.GetValue("Hechizo" & Hechizo, "MinSta"))
            .MaxSta = val(Leer.GetValue("Hechizo" & Hechizo, "MaxSta"))
            
            .SubeHam = val(Leer.GetValue("Hechizo" & Hechizo, "SubeHam"))
            .MinHam = val(Leer.GetValue("Hechizo" & Hechizo, "MinHam"))
            .MaxHam = val(Leer.GetValue("Hechizo" & Hechizo, "MaxHam"))
            
            .SubeSed = val(Leer.GetValue("Hechizo" & Hechizo, "SubeSed"))
            .MinSed = val(Leer.GetValue("Hechizo" & Hechizo, "MinSed"))
            .MaxSed = val(Leer.GetValue("Hechizo" & Hechizo, "MaxSed"))
            
            .SubeAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "SubeAG"))
            .MinAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "MinAG"))
            .MaxAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "MaxAG"))
            
            .SubeFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "SubeFU"))
            .MinFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "MinFU"))
            .MaxFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "MaxFU"))
            
            .SubeCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "SubeCA"))
            .MinCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "MinCA"))
            .MaxCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "MaxCA"))
            
            
            .Invisibilidad = val(Leer.GetValue("Hechizo" & Hechizo, "Invisibilidad"))
            .Paraliza = val(Leer.GetValue("Hechizo" & Hechizo, "Paraliza"))
            .Inmoviliza = val(Leer.GetValue("Hechizo" & Hechizo, "Inmoviliza"))
            .RemoverParalisis = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverParalisis"))
            .RemoverEstupidez = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverEstupidez"))
            .RemueveInvisibilidadParcial = val(Leer.GetValue("Hechizo" & Hechizo, "RemueveInvisibilidadParcial"))
            
            
            .CuraVeneno = val(Leer.GetValue("Hechizo" & Hechizo, "CuraVeneno"))
            .Envenena = val(Leer.GetValue("Hechizo" & Hechizo, "Envenena"))
            .Maldicion = val(Leer.GetValue("Hechizo" & Hechizo, "Maldicion"))
            .RemoverMaldicion = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverMaldicion"))
            .Bendicion = val(Leer.GetValue("Hechizo" & Hechizo, "Bendicion"))
            .Revivir = val(Leer.GetValue("Hechizo" & Hechizo, "Revivir"))
            
            .Ceguera = val(Leer.GetValue("Hechizo" & Hechizo, "Ceguera"))
            .Estupidez = val(Leer.GetValue("Hechizo" & Hechizo, "Estupidez"))
            
            .Warp = val(Leer.GetValue("Hechizo" & Hechizo, "Warp"))
            
            .Invoca = val(Leer.GetValue("Hechizo" & Hechizo, "Invoca"))
            .NumNpc = val(Leer.GetValue("Hechizo" & Hechizo, "NumNpc"))
            .Cant = val(Leer.GetValue("Hechizo" & Hechizo, "Cant"))
            .Mimetiza = val(Leer.GetValue("hechizo" & Hechizo, "Mimetiza"))
            
        '    .Materializa = val(Leer.GetValue("Hechizo" & Hechizo, "Materializa"))
        '    .ItemIndex = val(Leer.GetValue("Hechizo" & Hechizo, "ItemIndex"))
            
            .MinSkill = val(Leer.GetValue("Hechizo" & Hechizo, "MinSkill"))
            .ManaRequerido = val(Leer.GetValue("Hechizo" & Hechizo, "ManaRequerido"))
            
            'Barrin 30/9/03
            .StaRequerido = val(Leer.GetValue("Hechizo" & Hechizo, "StaRequerido"))
            
            .Baculo = val(Leer.GetValue("Hechizo" & Hechizo, "Baculo"))
            .Nivel = val(Leer.GetValue("Hechizo" & Hechizo, "Nivel"))
            
            .Target = val(Leer.GetValue("Hechizo" & Hechizo, "Target"))
            .Flecha = val(Leer.GetValue("Hechizo" & Hechizo, "Flecha"))
            
            frmCargando.cargar.Value = frmCargando.cargar.Value + 1
            
        End With
    Next Hechizo
    
    Set Leer = Nothing
    
    Exit Sub

Errhandler:
    MsgBox "Error cargando hechizos.dat " & Err.Number & ": " & Err.description
 
End Sub

Sub LoadMotd()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim i As Integer
    
    MaxLines = val(GetVar(App.path & "\Dat\Motd.ini", "INIT", "NumLines"))
    
    ReDim MOTD(1 To MaxLines)
    For i = 1 To MaxLines
        MOTD(i).texto = GetVar(App.path & "\Dat\Motd.ini", "Motd", "Line" & i)
        MOTD(i).Formato = vbNullString
    Next i

End Sub

Public Sub DoBackUp()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    haciendoBK = True
    Dim i As Integer
    
    
    
    ' Lo saco porque elimina elementales y mascotas - Maraxus
    ''''''''''''''lo pongo aca x sugernecia del yind
    'For i = 1 To LastNPC
    '    If Npclist(i).flags.NPCActive Then
    '        If Npclist(i).Contadores.TiempoExistencia > 0 Then
    '            Call MuereNpc(i, 0)
    '        End If
    '    End If
    'Next i
    '''''''''''/'lo pongo aca x sugernecia del yind
    
    
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    
    
    Call LimpiarMundo
    Call WorldSave
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Guardando Clanes.", FontTypeNames.FONTTYPE_SERVER))
    
    Call DumpGuilds
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Clanes Guardados.", FontTypeNames.FONTTYPE_SERVER))
    
  '  Call modGuilds.v_RutinaElecciones
    Call ResetCentinelaInfo     'Reseteamos al centinela
    
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    
    haciendoBK = False
    
    'Log
    On Error Resume Next
    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.path & "\logs\BackUps.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time
    Close #nfile
End Sub

'CSEH: ErrLog
Public Sub GrabarMapa(ByVal map As Long, ByVal MAPFILE As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo GrabarMapa_Err
    '</EhHeader>
        Dim FreeFileMap As Long
        Dim Y As Long
        Dim X As Long
        Dim writer As New clsByteBuffer
    
100     If FileExist(MAPFILE & ".bkp", vbNormal) Then
105         Kill MAPFILE & ".bkp"
        End If

        'Open .map file
110     FreeFileMap = FreeFile
115     Open MAPFILE & ".bkp" For Binary As FreeFileMap
120     Seek FreeFileMap, 1
    
    
125     Call writer.initializeWriter(FreeFileMap)
    
        'Write .bkp file
130     For Y = YMinMapSize To YMaxMapSize
135         For X = XMinMapSize To XMaxMapSize
140             With MapData(map, X, Y)
145                 If .ObjInfo.OBJIndex Then
150                     If Not ObjData(.ObjInfo.OBJIndex).OBJType = eOBJType.otFogata Then
155                         If Not ItemEsDeMapa(map, X, Y) Then
160                             Call writer.putByte(X)
165                             Call writer.putByte(Y)
170                             Call writer.putInteger(.ObjInfo.OBJIndex)
175                             Call writer.putInteger(.ObjInfo.Amount)
                            End If
                        End If
                    
                    End If
                End With
180         Next X
185     Next Y
    
190     Call writer.putByte(100)
    
195     Call writer.saveBuffer
    
        'Close .map file
200     Close FreeFileMap

    '<EhFooter>
    Exit Sub

GrabarMapa_Err:
        Call LogError("Error en GrabarMapa: " & Erl & " - " & Err.description)
    '</EhFooter>
End Sub
Sub LoadArmasHerreria()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim N As Integer, lc As Integer
    
    N = val(GetVar(DatPath & "ArmasHerrero.dat", "INIT", "NumArmas"))
    
    ReDim Preserve ArmasHerrero(1 To N) As Integer
    
    For lc = 1 To N
        ArmasHerrero(lc) = val(GetVar(DatPath & "ArmasHerrero.dat", "Arma" & lc, "Index"))
    Next lc

End Sub

Sub LoadArmadurasHerreria()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim N As Integer, lc As Integer
    
    N = val(GetVar(DatPath & "ArmadurasHerrero.dat", "INIT", "NumArmaduras"))
    
    ReDim Preserve ArmadurasHerrero(1 To N) As Integer
    
    For lc = 1 To N
        ArmadurasHerrero(lc) = val(GetVar(DatPath & "ArmadurasHerrero.dat", "Armadura" & lc, "Index"))
    Next lc

End Sub

Sub LoadBalance()
'***************************************************
'Author: Unknown
'Last Modification: 15/04/2010
'15/04/2010: ZaMa - Agrego recompensas faccionarias.
'***************************************************

    Dim i As Long
    Dim j As Long
    
    For i = 1 To NUMCLASES
        If Len(ListaClases(i)) > 0 Then
            For j = 1 To 6
                Mods(j, i) = Int(GetVar(DatPath & "Balance.dat", ListaClases(i), "Mod" & j)) / 100
            Next
        End If
    Next
    
    'Modificadores de Raza
    For i = 1 To NUMRAZAS
        With ModRaza(i)
            .Fuerza = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Fuerza"))
            .Agilidad = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Agilidad"))
            .Inteligencia = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Inteligencia"))
            .Carisma = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Carisma"))
            .Constitucion = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Constitucion"))
        End With
    Next i
    
    'Modificadores de Vida
    For i = 1 To NUMCLASES
        ModVida(i) = val(GetVar(DatPath & "Balance.dat", "MODVIDA", ListaClases(i)))
    Next i
    
    'Extra
    PorcentajeRecuperoMana = val(GetVar(DatPath & "Balance.dat", "EXTRA", "PorcentajeRecuperoMana"))
    
End Sub

Sub LoadObjSastre()

    Dim N As Integer, lc As Integer
    
    N = val(GetVar(DatPath & "ObjSastre.dat", "INIT", "NumObjs"))
    
    ReDim Preserve ObjSastre(1 To N) As Integer
    
    For lc = 1 To N
        ObjSastre(lc) = val(GetVar(DatPath & "ObjSastre.dat", "Obj" & lc, "Index"))
    Next

End Sub

Sub LoadObjCarpintero()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim N As Integer, lc As Integer
    
    N = val(GetVar(DatPath & "ObjCarpintero.dat", "INIT", "NumObjs"))
    
    ReDim Preserve ObjCarpintero(1 To N) As Integer
    
    For lc = 1 To N
        ObjCarpintero(lc) = val(GetVar(DatPath & "ObjCarpintero.dat", "Obj" & lc, "Index"))
    Next lc

End Sub



Sub LoadOBJData()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

'###################################################
'#               ATENCION PELIGRO                  #
'###################################################
'
'¡¡¡¡ NO USAR GetVar PARA LEER DESDE EL OBJ.DAT !!!!
'
'El que ose desafiar esta LEY, se las tendrá que ver
'con migo. Para leer desde el OBJ.DAT se deberá usar
'la nueva clase clsLeerInis.
'
'Alejo
'
'###################################################

'Call LogTarea("Sub LoadOBJData")

    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando base de datos de los objetos."
    
    '*****************************************************************
    'Carga la lista de objetos
    '*****************************************************************
    Dim Object As Integer
    Dim Leer As New clsIniReader
    
    Call Leer.Initialize(DatPath & "Obj.dat")
    
    'obtiene el numero de obj
    NumObjDatas = val(Leer.GetValue("INIT", "NumObjs"))
    
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumObjDatas
    frmCargando.cargar.Value = 0
    
    
    ReDim Preserve ObjData(1 To NumObjDatas) As ObjData
    
    
    'Llena la lista
    For Object = 1 To NumObjDatas
        With ObjData(Object)
            .Name = Leer.GetValue("OBJ" & Object, "Name")
            
            .NoComerciable = val(Leer.GetValue("OBJ" & Object, "NoComerciable"))
            
            .GrhIndex = val(Leer.GetValue("OBJ" & Object, "GrhIndex"))
            If .GrhIndex = 0 Then
                .GrhIndex = .GrhIndex
            End If
            
            .OBJType = val(Leer.GetValue("OBJ" & Object, "ObjType"))
            
            .Newbie = val(Leer.GetValue("OBJ" & Object, "Newbie"))
            
            Select Case .OBJType
                Case eOBJType.otArmadura
                    .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                    .LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                    .LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                    .LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                    .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                    .Jerarquia = val(Leer.GetValue("OBJ" & Object, "Jerarquia"))
                    .PielLobo = val(Leer.GetValue("OBJ" & Object, "PielLobo"))
                    .PielOsoPardo = val(Leer.GetValue("OBJ" & Object, "PielOsoPardo"))
                    .PielOsoPolar = val(Leer.GetValue("OBJ" & Object, "PielOsoPolar"))
                    .SkSastreria = val(Leer.GetValue("OBJ" & Object, "SkSastreria"))
                    
                Case eOBJType.otEscudo
                    .ShieldAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
                    .LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                    .LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                    .LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                    .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                    .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                
                Case eOBJType.otCasco
                    .CascoAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
                    .LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                    .LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                    .LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                    .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                    .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                    .Gorro = val(Leer.GetValue("OBJ" & Object, "Gorro"))
                    
                Case eOBJType.otWeapon
                    .WeaponAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
                    .Apuñala = val(Leer.GetValue("OBJ" & Object, "Apuñala"))
                    .Envenena = val(Leer.GetValue("OBJ" & Object, "Envenena"))
                    .MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                    .MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                    .proyectil = val(Leer.GetValue("OBJ" & Object, "Proyectil"))
                    .Municion = val(Leer.GetValue("OBJ" & Object, "Municiones"))
                    .Baculo = val(Leer.GetValue("OBJ" & Object, "Baculo"))
                    
                    .LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                    .LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                    .LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                    .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                    .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                    
                
                Case eOBJType.otInstrumentos
                    .Snd1 = val(Leer.GetValue("OBJ" & Object, "SND1"))
                    .Snd2 = val(Leer.GetValue("OBJ" & Object, "SND2"))
                    .Snd3 = val(Leer.GetValue("OBJ" & Object, "SND3"))
                    'Pablo (ToxicWaste)
                    .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                
                Case eOBJType.otMinerales
                    .MinSkill = val(Leer.GetValue("OBJ" & Object, "MinSkill"))
                
                Case eOBJType.otPuertas, eOBJType.otBotellaVacia, eOBJType.otBotellaLlena
                    .IndexAbierta = val(Leer.GetValue("OBJ" & Object, "IndexAbierta"))
                    .IndexCerrada = val(Leer.GetValue("OBJ" & Object, "IndexCerrada"))
                    .IndexCerradaLlave = val(Leer.GetValue("OBJ" & Object, "IndexCerradaLlave"))
                
                Case otPociones
                    .TipoPocion = val(Leer.GetValue("OBJ" & Object, "TipoPocion"))
                    .MaxModificador = val(Leer.GetValue("OBJ" & Object, "MaxModificador"))
                    .MinModificador = val(Leer.GetValue("OBJ" & Object, "MinModificador"))
                    .DuracionEfecto = val(Leer.GetValue("OBJ" & Object, "DuracionEfecto"))
                
                Case eOBJType.otBarcos
                    .MinSkill = val(Leer.GetValue("OBJ" & Object, "MinSkill"))
                    .MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                    .MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                
                Case eOBJType.otFlechas
                    .MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                    .MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                    .Envenena = val(Leer.GetValue("OBJ" & Object, "Envenena"))

                    
                Case eOBJType.otTeleport
                    .Radio = val(Leer.GetValue("OBJ" & Object, "Radio"))
                    
                Case eOBJType.otMochilas
                    .MochilaType = val(Leer.GetValue("OBJ" & Object, "MochilaType"))
                    
                Case eOBJType.otForos
                    Call AddForum(Leer.GetValue("OBJ" & Object, "ID"))
                    
            End Select
            
            .Ropaje = val(Leer.GetValue("OBJ" & Object, "NumRopaje"))
            .HechizoIndex = val(Leer.GetValue("OBJ" & Object, "HechizoIndex"))
            
            .LingoteIndex = val(Leer.GetValue("OBJ" & Object, "LingoteIndex"))
            
            .MineralIndex = val(Leer.GetValue("OBJ" & Object, "MineralIndex"))
            
            .MaxHp = val(Leer.GetValue("OBJ" & Object, "MaxHP"))
            .MinHp = val(Leer.GetValue("OBJ" & Object, "MinHP"))
            
            .Mujer = val(Leer.GetValue("OBJ" & Object, "Mujer"))
            .Hombre = val(Leer.GetValue("OBJ" & Object, "Hombre"))
            
            .MinHam = val(Leer.GetValue("OBJ" & Object, "MinHam"))
            .MinSed = val(Leer.GetValue("OBJ" & Object, "MinAgu"))
            
            .MinDef = val(Leer.GetValue("OBJ" & Object, "MINDEF"))
            .MaxDef = val(Leer.GetValue("OBJ" & Object, "MAXDEF"))
            .def = (.MinDef + .MaxDef) / 2
                        
            .Valor = val(Leer.GetValue("OBJ" & Object, "Valor"))
            
            .Crucial = val(Leer.GetValue("OBJ" & Object, "Crucial"))
            
            .Cerrada = val(Leer.GetValue("OBJ" & Object, "abierta"))
            If .Cerrada = 1 Then
                .Llave = val(Leer.GetValue("OBJ" & Object, "Llave"))
                .clave = val(Leer.GetValue("OBJ" & Object, "Clave"))
            End If
            
            If .OBJType = eOBJType.otWarp Then
                .WMapa = val(Leer.GetValue("OBJ" & Object, "WMapa"))
                .WX = val(Leer.GetValue("OBJ" & Object, "WX"))
                .WY = val(Leer.GetValue("OBJ" & Object, "WY"))
                .WI = val(Leer.GetValue("OBJ" & Object, "WI"))
            End If
            
            'Puertas y llaves
            .clave = val(Leer.GetValue("OBJ" & Object, "Clave"))
            
            .texto = Leer.GetValue("OBJ" & Object, "Texto")
            .GrhSecundario = val(Leer.GetValue("OBJ" & Object, "VGrande"))
            
            .Agarrable = val(Leer.GetValue("OBJ" & Object, "Agarrable"))
            .ForoID = Leer.GetValue("OBJ" & Object, "ID")
            
            Dim Num As Integer
            
            Num = val(Leer.GetValue("OBJ" & Object, "NumClases"))
            
            Dim i As Integer
            For i = 1 To Num
                .ClaseProhibida(i) = val(Leer.GetValue("OBJ" & Object, "CP" & i))
            Next
            
            Num = val(Leer.GetValue("OBJ" & Object, "NumRazas"))
             
            For i = 1 To Num
                .RazaProhibida(i) = val(Leer.GetValue("OBJ" & Object, "RP" & i))
            Next
            
            
            .SkCarpinteria = val(Leer.GetValue("OBJ" & Object, "SkCarpinteria"))
            
            If .SkCarpinteria > 0 Then _
                .Madera = val(Leer.GetValue("OBJ" & Object, "Madera"))
                .MaderaElfica = val(Leer.GetValue("OBJ" & Object, "MaderaElfica"))
            
            'Bebidas
            .MinSta = val(Leer.GetValue("OBJ" & Object, "MinST"))
            
            .NoSeCae = val(Leer.GetValue("OBJ" & Object, "NoSeCae"))
            
            frmCargando.cargar.Value = frmCargando.cargar.Value + 1
        End With
    Next Object
    
    
    Set Leer = Nothing
    
    ' Inicializo los foros faccionarios
    Call AddForum(FORO_CAOS_ID)
    Call AddForum(FORO_REAL_ID)
    
    Exit Sub

Errhandler:
    MsgBox "error cargando objetos " & Err.Number & ": " & Err.description


End Sub

Sub LoadUserStats(ByVal UserIndex As Integer, ByRef UserFile As clsIniReader)
'*************************************************
'Author: Unknown
'Last modified: 11/19/2009
'*************************************************
Dim LoopC As Long

With UserList(UserIndex)
    With .Stats
        For LoopC = 1 To NUMATRIBUTOS
            .UserAtributos(LoopC) = CInt(UserFile.GetValue("ATRIBUTOS", "AT" & LoopC))
            .UserAtributosBackUP(LoopC) = .UserAtributos(LoopC)
        Next LoopC
        
        For LoopC = 1 To NUMSKILLS
            .UserSkills(LoopC) = CInt(UserFile.GetValue("SKILLS", "SK" & LoopC))
        Next LoopC
        
        For LoopC = 1 To MAXUSERHECHIZOS
            .UserHechizos(LoopC) = CInt(UserFile.GetValue("Hechizos", "H" & LoopC))
        Next LoopC
        
        .GLD = CLng(UserFile.GetValue("STATS", "GLD"))
        .Banco = CLng(UserFile.GetValue("STATS", "BANCO"))
        
        .MaxHp = CInt(UserFile.GetValue("STATS", "MaxHP"))
        .MinHp = CInt(UserFile.GetValue("STATS", "MinHP"))
        
        .MinSta = CInt(UserFile.GetValue("STATS", "MinSTA"))
        .MaxSta = CInt(UserFile.GetValue("STATS", "MaxSTA"))
        
        .MaxMAN = CInt(UserFile.GetValue("STATS", "MaxMAN"))
        .MinMAN = CInt(UserFile.GetValue("STATS", "MinMAN"))
        
        .MaxHIT = CInt(UserFile.GetValue("STATS", "MaxHIT"))
        .MinHIT = CInt(UserFile.GetValue("STATS", "MinHIT"))
        
        .MaxAGU = CByte(UserFile.GetValue("STATS", "MaxAGU"))
        .MinAGU = CByte(UserFile.GetValue("STATS", "MinAGU"))
        
        .MaxHam = CByte(UserFile.GetValue("STATS", "MaxHAM"))
        .MinHam = CByte(UserFile.GetValue("STATS", "MinHAM"))
        
        .SkillPts = CInt(UserFile.GetValue("STATS", "SkillPtsLibres"))
        
        .Exp = CDbl(UserFile.GetValue("STATS", "EXP"))
        .ELU = CLng(UserFile.GetValue("STATS", "ELU"))
        .ELV = CByte(UserFile.GetValue("STATS", "ELV"))
        
        
        .UsuariosMatados = CLng(UserFile.GetValue("MUERTES", "UserMuertes"))
        .NPCsMuertos = CInt(UserFile.GetValue("MUERTES", "NpcsMuertes"))
    End With
    
    With .flags
        If CByte(UserFile.GetValue("CONSEJO", "PERTENECE")) Then _
            .Privilegios = .Privilegios Or PlayerType.RoyalCouncil
        
        If CByte(UserFile.GetValue("CONSEJO", "PERTENECECAOS")) Then _
            .Privilegios = .Privilegios Or PlayerType.ChaosCouncil
    End With
End With
End Sub

Sub LoadUserInit(ByVal UserIndex As Integer, ByRef UserFile As clsIniReader)
'*************************************************
'Author: Unknown
'Last modified: 19/11/2006
'Loads the Users records
'23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
'23/01/2007 Pablo (ToxicWaste) - Quito CriminalesMatados de Stats porque era redundante.
'*************************************************
    Dim LoopC As Long
    Dim ln As String
    
    With UserList(UserIndex)
        With .Faccion
            .Bando = val(UserFile.GetValue("FACCIONES", "Bando"))
            .BandoOriginal = val(UserFile.GetValue("FACCIONES", "BandoOriginal"))
            .Matados(0) = val(UserFile.GetValue("FACCIONES", "Matados0"))
            .Matados(1) = val(UserFile.GetValue("FACCIONES", "Matados1"))
            .Matados(2) = val(UserFile.GetValue("FACCIONES", "Matados2"))
            .Jerarquia = val(UserFile.GetValue("FACCIONES", "Jerarquia"))
            .Ataco(1) = val(UserFile.GetValue("FACCIONES", "Ataco1"))
            .Ataco(2) = val(UserFile.GetValue("FACCIONES", "Ataco2"))
            '.Quests = val(UserFile.GetValue("FACCIONES", "Quests"))
            .Torneos = val(UserFile.GetValue("FACCIONES", "Torneos"))
        End With
        
        With .flags
            .Muerto = CByte(UserFile.GetValue("FLAGS", "Muerto"))
            .Escondido = CByte(UserFile.GetValue("FLAGS", "Escondido"))
            .Hambre = CByte(UserFile.GetValue("FLAGS", "Hambre"))
            .Sed = CByte(UserFile.GetValue("FLAGS", "Sed"))
            .Desnudo = CByte(UserFile.GetValue("FLAGS", "Desnudo"))
            .Navegando = CByte(UserFile.GetValue("FLAGS", "Navegando"))
            .Envenenado = CByte(UserFile.GetValue("FLAGS", "Envenenado"))
            .Paralizado = CByte(UserFile.GetValue("FLAGS", "Paralizado"))
            .IsLeader = CByte(UserFile.GetValue("FLAGS", "IsLeader"))
            
            'Matrix
            .lastMap = CInt(UserFile.GetValue("FLAGS", "LastMap"))
        End With
        
        With .Events
            .Quests = CByte(UserFile.GetValue("EVENTOS", "Quests"))
            .Torneos = CByte(UserFile.GetValue("EVENTOS", "Torneos"))
        End With
        
        If .flags.Paralizado = 1 Then
            .Counters.Paralisis = IntervaloParalizado
        End If
        
        
        .Counters.Pena = CLng(UserFile.GetValue("COUNTERS", "Pena"))
        .Counters.AsignedSkills = CByte(val(UserFile.GetValue("COUNTERS", "SkillsAsignados")))
        
        .email = UserFile.GetValue("CONTACTO", "Email")
        
        .Genero = UserFile.GetValue("INIT", "Genero")
        .Clase = UserFile.GetValue("INIT", "Clase")
        .raza = UserFile.GetValue("INIT", "Raza")
        .Hogar = UserFile.GetValue("INIT", "Hogar")
        .Char.heading = CInt(UserFile.GetValue("INIT", "Heading"))
        
        
        With .OrigChar
            .Head = CInt(UserFile.GetValue("INIT", "Head"))
            .body = CInt(UserFile.GetValue("INIT", "Body"))
            .WeaponAnim = CInt(UserFile.GetValue("INIT", "Arma"))
            .ShieldAnim = CInt(UserFile.GetValue("INIT", "Escudo"))
            .CascoAnim = CInt(UserFile.GetValue("INIT", "Casco"))
            
            .heading = eHeading.SOUTH
        End With
        
        #If ConUpTime Then
            .UpTime = CLng(UserFile.GetValue("INIT", "UpTime"))
        #End If
        
        If .flags.Muerto = 0 Then
            .Char = .OrigChar
        Else
            .Char.body = iCuerpoMuerto
            .Char.Head = iCabezaMuerto
            .Char.WeaponAnim = NingunArma
            .Char.ShieldAnim = NingunEscudo
            .Char.CascoAnim = NingunCasco
        End If
        
        
        .desc = UserFile.GetValue("INIT", "Desc")
        
        .Pos.map = CInt(ReadField(1, UserFile.GetValue("INIT", "Position"), 45))
        .Pos.X = CInt(ReadField(2, UserFile.GetValue("INIT", "Position"), 45))
        .Pos.Y = CInt(ReadField(3, UserFile.GetValue("INIT", "Position"), 45))
        
        .Invent.NroItems = CInt(UserFile.GetValue("Inventory", "CantidadItems"))
        
        '[KEVIN]--------------------------------------------------------------------
        '***********************************************************************************
        .BancoInvent.NroItems = CInt(UserFile.GetValue("BancoInventory", "CantidadItems"))
        'Lista de objetos del banco
        For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
            ln = UserFile.GetValue("BancoInventory", "Obj" & LoopC)
            .BancoInvent.Object(LoopC).OBJIndex = CInt(ReadField(1, ln, 45))
            .BancoInvent.Object(LoopC).Amount = CInt(ReadField(2, ln, 45))
        Next LoopC
        '------------------------------------------------------------------------------------
        '[/KEVIN]*****************************************************************************
        
        
        'Lista de objetos
        For LoopC = 1 To MAX_INVENTORY_SLOTS
            ln = UserFile.GetValue("Inventory", "Obj" & LoopC)
            .Invent.Object(LoopC).OBJIndex = CInt(ReadField(1, ln, 45))
            .Invent.Object(LoopC).Amount = CInt(ReadField(2, ln, 45))
            .Invent.Object(LoopC).Equipped = CByte(ReadField(3, ln, 45))
        Next LoopC
        
        'Obtiene el indice-objeto del arma
        .Invent.WeaponEqpSlot = CByte(UserFile.GetValue("Inventory", "WeaponEqpSlot"))
        If .Invent.WeaponEqpSlot > 0 Then
            .Invent.WeaponEqpObjIndex = .Invent.Object(.Invent.WeaponEqpSlot).OBJIndex
        End If
        
        'Obtiene el indice-objeto del armadura
        .Invent.ArmourEqpSlot = CByte(UserFile.GetValue("Inventory", "ArmourEqpSlot"))
        If .Invent.ArmourEqpSlot > 0 Then
            .Invent.ArmourEqpObjIndex = .Invent.Object(.Invent.ArmourEqpSlot).OBJIndex
            .flags.Desnudo = 0
        Else
            .flags.Desnudo = 1
        End If
        
        'Obtiene el indice-objeto del escudo
        .Invent.EscudoEqpSlot = CByte(UserFile.GetValue("Inventory", "EscudoEqpSlot"))
        If .Invent.EscudoEqpSlot > 0 Then
            .Invent.EscudoEqpObjIndex = .Invent.Object(.Invent.EscudoEqpSlot).OBJIndex
        End If
        
        'Obtiene el indice-objeto del casco
        .Invent.CascoEqpSlot = CByte(UserFile.GetValue("Inventory", "CascoEqpSlot"))
        If .Invent.CascoEqpSlot > 0 Then
            .Invent.CascoEqpObjIndex = .Invent.Object(.Invent.CascoEqpSlot).OBJIndex
        End If
        
        'Obtiene el indice-objeto barco
        .Invent.BarcoSlot = CByte(UserFile.GetValue("Inventory", "BarcoSlot"))
        If .Invent.BarcoSlot > 0 Then
            .Invent.BarcoObjIndex = .Invent.Object(.Invent.BarcoSlot).OBJIndex
        End If
        
        'Obtiene el indice-objeto municion
        .Invent.MunicionEqpSlot = CByte(UserFile.GetValue("Inventory", "MunicionSlot"))
        If .Invent.MunicionEqpSlot > 0 Then
            .Invent.MunicionEqpObjIndex = .Invent.Object(.Invent.MunicionEqpSlot).OBJIndex
        End If
        
        .Invent.MochilaEqpSlot = CByte(UserFile.GetValue("Inventory", "MochilaSlot"))
        If .Invent.MochilaEqpSlot > 0 Then
            .Invent.MochilaEqpObjIndex = .Invent.Object(.Invent.MochilaEqpSlot).OBJIndex
        End If
        
        .Invent.HerramientaEqpslot = CByte(UserFile.GetValue("Inventory", "HerramientaSlot"))
        If .Invent.HerramientaEqpslot > 0 Then
            .Invent.HerramientaEqpObjIndex = .Invent.Object(.Invent.HerramientaEqpslot).OBJIndex
        End If
        
        .NroMascotas = CInt(UserFile.GetValue("MASCOTAS", "NroMascotas"))
        For LoopC = 1 To MAXMASCOTAS
            .MascotasType(LoopC) = val(UserFile.GetValue("MASCOTAS", "MAS" & LoopC))
        Next LoopC
        
        
        For LoopC = 1 To 3
            .Recompensas(LoopC) = val(UserFile.GetValue("RECOMPENSAS", "Recompensa" & LoopC))
        Next
        
        ln = UserFile.GetValue("Guild", "GUILDID")
        
        If IsNumeric(ln) Then
            If CLng(ln) > LastGuild Then
                .GuildID = 0
                Call LogError("Usuario: " & .Name & " con guild id incorrecto.")
            Else
                .GuildID = CLng(ln)
            End If
        Else
            .GuildID = 0
        End If
        
        .flags.WaitingApprovement = CLng(UserFile.GetValue("Guild", "RequestedTo"))
        
      '  ln = UserFile.GetValue("Guild", "GUILDINDEX")
       ' If IsNumeric(ln) Then
       '     .GuildIndex = CInt(ln)
       ' Else
       '     .GuildIndex = 0
       ' End If
    End With

End Sub

Function GetVar(ByVal File As String, ByVal Main As String, ByVal Var As String, Optional EmptySpaces As Long = 1024) As String
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim sSpaces As String ' This will hold the input that the program will retrieve
    Dim szReturn As String ' This will be the defaul value if the string is not found
      
    szReturn = vbNullString
      
    sSpaces = Space$(EmptySpaces) ' This tells the computer how long the longest string can be
      
      
    GetPrivateProfileString Main, Var, szReturn, sSpaces, EmptySpaces, File
      
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
  
End Function

'CSEH: ErrLog
Sub CargarBackUp()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo CargarBackUp_Err
    '</EhHeader>

100     If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando backup."
    
        Dim map As Integer
        Dim tFileName As String
        
105         NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
110         Call InitAreas
        
115         frmCargando.cargar.min = 0
120         frmCargando.cargar.max = NumMaps
125         frmCargando.cargar.Value = 0
        
130         MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")
        
        
135         ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
140         ReDim MapInfo(1 To NumMaps) As MapInfo
        
145         Call CargarInfoMapas(MapPath)
        
150         For map = 1 To NumMaps
155             If val(GetVar(App.path & MapPath & "Mapa" & map & ".Dat", "Mapa" & map, "BackUp")) <> 0 Then
160                 tFileName = App.path & "\WorldBackUp\Mapa" & map
                
165                 If Not FileExist(tFileName & ".*") Then 'Miramos que exista al menos uno de los 3 archivos, sino lo cargamos de la carpeta de los mapas
170                     tFileName = App.path & MapPath & "Mapa" & map
                    End If
                Else
175                 tFileName = App.path & MapPath & "Mapa" & map
                End If
            
180             Call CargarMapa(map, tFileName)
            
185             frmCargando.cargar.Value = frmCargando.cargar.Value + 1
190             DoEvents
195         Next map
    '<EhFooter>
    Exit Sub

CargarBackUp_Err:
        Call LogError("Error en CargarBackUp: " & Erl & " - " & Err.description)
    '</EhFooter>
End Sub

'CSEH: ErrLog
Sub LoadMapData()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo LoadMapData_Err
    '</EhHeader>

100     If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando mapas..."
    
        Dim map As Integer
        Dim tFileName As String
    
105         NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
110         Call InitAreas
        
115         frmCargando.cargar.min = 0
120         frmCargando.cargar.max = NumMaps
125         frmCargando.cargar.Value = 0
        
130         MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")
        
        
135         ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
140         ReDim MapInfo(1 To NumMaps) As MapInfo
        
145         Call CargarInfoMapas(MapPath)
        
150         For map = 1 To NumMaps
            
155             tFileName = App.path & MapPath & "Mapa" & map
160             Call CargarMapa(map, tFileName)
            
165             frmCargando.cargar.Value = frmCargando.cargar.Value + 1
170             DoEvents
175         Next map

    '<EhFooter>
    Exit Sub

LoadMapData_Err:
        Call LogError("Error en LoadMapData: " & Erl & " - " & Err.description)
    '</EhFooter>
End Sub

'CSEH: ErrLog
Public Sub CargarInfoMapas(ByVal path As String)
    '<EhHeader>
    On Error GoTo CargarInfoMapas_Err
    '</EhHeader>
    Dim reader As New clsIniManager
    Dim i As Long

100 reader.Initialize (App.path & path & "Info.dat")

105 For i = 1 To NumMaps
110     With MapInfo(i)
    
115         If reader.KeyExists("Mapa" & i) Then
120             .Name = reader.GetValue("Mapa" & i, "Name")
125             .Music = reader.GetValue("Mapa" & i, "MusicNum")
            
                'no voy a adaptar todo el drama de "toppunto" y "leftpunto" solo por que el we de fenix sea un desastre
130             .Pk = val(reader.GetValue("Mapa" & i, "Pk")) = 0
135             .MagiaSinEfecto = val(reader.GetValue("Mapa" & i, "NoMagia"))
            
140             .Terreno = reader.GetValue("Mapa" & i, "Terreno")
145             .Zona = reader.GetValue("Mapa" & i, "Zona")
                
150             .Restringir = val(reader.GetValue("Mapa" & i, "Restringir")) = 1
155             .Nivel = val(reader.GetValue("Mapa" & i, "Nivel"))
160             .BackUp = val(reader.GetValue("Mapa" & i, "Backup"))
            
            End If
        End With
    Next

    '<EhFooter>
    Exit Sub

CargarInfoMapas_Err:
        Call LogError("Error en CargarInfoMapas: " & Erl & " - " & Err.description)
    '</EhFooter>
End Sub
'CSEH: ErrLog
Public Sub CargarMapa(ByVal map As Long, ByVal MAPFl As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo CargarMapa_Err
    '</EhHeader>
        Dim FreeFileMap As Long
        Dim Y As Long
        Dim X As Long
        Dim ByFlags As Byte
        Dim npcfile As String
        Dim i As Long
    
        'array to store map data
        Dim data() As Byte
        Dim mapReader As New clsByteBuffer

100     FreeFileMap = FreeFile

105     Open MAPFl & ".msv" For Binary As #FreeFileMap
110         Seek FreeFileMap, 1
        
115         ReDim data(0 To LOF(FreeFileMap) - 1) As Byte
        
120         Get #FreeFileMap, , data

125     Close #FreeFileMap
    
130     Call mapReader.initializeReader(data)
    
135     MapInfo(map).MapVersion = mapReader.getInteger
    
140     For Y = YMinMapSize To YMaxMapSize
145         For X = XMinMapSize To XMaxMapSize
150             With MapData(map, X, Y)

                
155                 ByFlags = mapReader.getByte

160                 If ByFlags And 1 Then .Blocked = 1
165                 If ByFlags And 2 Then .Agua = 1

175                 For i = 2 To 4
180                     If (ByFlags And 2 ^ i) Then .trigger = .trigger Or 2 ^ (i - 2)
                    Next
                
185                 If ByFlags And 32 Then
190                     .NpcIndex = mapReader.getInteger
                                               
195                     If .NpcIndex > 0 Then
200                         npcfile = DatPath & "NPCs.dat"

                            'Si el npc debe hacer respawn en la pos
                            'original la guardamos
205                         If val(GetVar(npcfile, "NPC" & .NpcIndex, "PosOrig")) = 1 Then
210                             .NpcIndex = OpenNPC(.NpcIndex)
215                             Npclist(.NpcIndex).Orig.map = map
220                             Npclist(.NpcIndex).Orig.X = X
225                             Npclist(.NpcIndex).Orig.Y = Y
                            Else
230                             .NpcIndex = OpenNPC(.NpcIndex)
                            End If
                        
235                         Npclist(.NpcIndex).Pos.map = map
240                         Npclist(.NpcIndex).Pos.X = X
245                         Npclist(.NpcIndex).Pos.Y = Y
                        
250                         Call MakeNPCChar(True, 0, .NpcIndex, map, X, Y)
                        End If
                    End If
                    
255                 If ByFlags And 64 Then
260                     .ObjInfo.OBJIndex = mapReader.getInteger
265                     .ObjInfo.Amount = mapReader.getInteger
                    End If
                
                    'no se porqué fenix tiene estas cosas raras
270                 If .ObjInfo.OBJIndex > UBound(ObjData) Then
275                     .ObjInfo.OBJIndex = 0
280                     .ObjInfo.Amount = 0
                    End If
                
285                 If ByFlags And 128 Then
290                     .TileExit.map = mapReader.getInteger
295                     .TileExit.X = mapReader.getInteger
300                     .TileExit.Y = mapReader.getInteger
                    End If
                
                End With
305         Next X
310     Next Y

    '<EhFooter>
    Exit Sub

CargarMapa_Err:
        Call LogError("Error en CargarMapa " & map & ": " & X & ": " & Y & "** " & Erl & " - " & Err.description)
    '</EhFooter>
End Sub

Sub LoadSini()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim Temporal As Long
    
    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando info de inicio del server."
    
    BootDelBackUp = val(GetVar(IniPath & "Server.ini", "INIT", "IniciarDesdeBackUp"))
    
    
    Puerto = val(GetVar(IniPath & "Server.ini", "INIT", "StartPort"))
    HideMe = val(GetVar(IniPath & "Server.ini", "INIT", "Hide"))
    AllowMultiLogins = val(GetVar(IniPath & "Server.ini", "INIT", "AllowMultiLogins"))
    IdleLimit = val(GetVar(IniPath & "Server.ini", "INIT", "IdleLimit"))
    'Lee la version correcta del cliente
    ULTIMAVERSION = GetVar(IniPath & "Server.ini", "INIT", "Version")
    
    PuedeCrearPersonajes = val(GetVar(IniPath & "Server.ini", "INIT", "PuedeCrearPersonajes"))
    ServerSoloGMs = val(GetVar(IniPath & "Server.ini", "init", "ServerSoloGMs"))
    
    MAPA_PRETORIANO = val(GetVar(IniPath & "Server.ini", "INIT", "MapaPretoriano"))
    
    EnTesting = val(GetVar(IniPath & "Server.ini", "INIT", "Testing"))
    
    'Intervalos
    SanaIntervaloSinDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloSinDescansar"))
    FrmInterv.txtSanaIntervaloSinDescansar.Text = SanaIntervaloSinDescansar
    
    StaminaIntervaloSinDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloSinDescansar"))
    FrmInterv.txtStaminaIntervaloSinDescansar.Text = StaminaIntervaloSinDescansar
    
    SanaIntervaloDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloDescansar"))
    FrmInterv.txtSanaIntervaloDescansar.Text = SanaIntervaloDescansar
    
    StaminaIntervaloDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloDescansar"))
    FrmInterv.txtStaminaIntervaloDescansar.Text = StaminaIntervaloDescansar
    
    IntervaloSed = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloSed"))
    FrmInterv.txtIntervaloSed.Text = IntervaloSed
    
    IntervaloHambre = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloHambre"))
    FrmInterv.txtIntervaloHambre.Text = IntervaloHambre
    
    IntervaloVeneno = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloVeneno"))
    FrmInterv.txtIntervaloVeneno.Text = IntervaloVeneno
    
    IntervaloParalizado = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParalizado"))
    FrmInterv.txtIntervaloParalizado.Text = IntervaloParalizado
    
    IntervaloInvisible = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvisible"))
    FrmInterv.txtIntervaloInvisible.Text = IntervaloInvisible
    
    IntervaloFrio = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFrio"))
    FrmInterv.txtIntervaloFrio.Text = IntervaloFrio
    
    IntervaloWavFx = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWAVFX"))
    FrmInterv.txtIntervaloWAVFX.Text = IntervaloWavFx
    
    IntervaloInvocacion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvocacion"))
    FrmInterv.txtInvocacion.Text = IntervaloInvocacion
    
    IntervaloParaConexion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParaConexion"))
    FrmInterv.txtIntervaloParaConexion.Text = IntervaloParaConexion
    
    '&&&&&&&&&&&&&&&&&&&&& TIMERS &&&&&&&&&&&&&&&&&&&&&&&
    
    IntervaloPuedeSerAtacado = 5000 ' Cargar desde balance.dat
    IntervaloAtacable = 60000 ' Cargar desde balance.dat
    
    IntervaloUserPuedeCastear = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloLanzaHechizo"))
    FrmInterv.txtIntervaloLanzaHechizo.Text = IntervaloUserPuedeCastear
    
    frmMain.TIMER_AI.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloNpcAI"))
    FrmInterv.txtAI.Text = frmMain.TIMER_AI.Interval
    
    frmMain.npcataca.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloNpcPuedeAtacar"))
    FrmInterv.txtNPCPuedeAtacar.Text = frmMain.npcataca.Interval
    
    IntervaloUserPuedeTrabajar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloTrabajo"))
    FrmInterv.txtTrabajo.Text = IntervaloUserPuedeTrabajar
    
    IntervaloUserPuedeAtacar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeAtacar"))
    FrmInterv.txtPuedeAtacar.Text = IntervaloUserPuedeAtacar
    
    'TODO : Agregar estos intervalos al form!!!
    IntervaloMagiaGolpe = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloMagiaGolpe"))
    IntervaloGolpeMagia = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloGolpeMagia"))
    IntervaloGolpeUsar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloGolpeUsar"))
    
    frmMain.tLluvia.Interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloPerdidaStaminaLluvia"))
    FrmInterv.txtIntervaloPerdidaStaminaLluvia.Text = frmMain.tLluvia.Interval
    
    MinutosWs = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWS"))
    If MinutosWs < 60 Then MinutosWs = 180
    
    IntervaloCerrarConexion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloCerrarConexion"))
    IntervaloUserPuedeUsar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeUsar"))
    IntervaloFlechasCazadores = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFlechasCazadores"))
    
    IntervaloOculto = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloOculto"))
    
    '&&&&&&&&&&&&&&&&&&&&& FIN TIMERS &&&&&&&&&&&&&&&&&&&&&&&
      
    recordusuarios = val(GetVar(IniPath & "Server.ini", "INIT", "Record"))
      
    'Max users
    Temporal = val(GetVar(IniPath & "Server.ini", "INIT", "MaxUsers"))
    If MaxUsers = 0 Then
        MaxUsers = Temporal
        ReDim UserList(1 To MaxUsers) As User
    End If
    
    '&&&&&&&&&&&&&&&&&&&&& BALANCE &&&&&&&&&&&&&&&&&&&&&&&
    'Se agregó en LoadBalance y en el Balance.dat
    'PorcentajeRecuperoMana = val(GetVar(IniPath & "Server.ini", "BALANCE", "PorcentajeRecuperoMana"))
    
    ''&&&&&&&&&&&&&&&&&&&&& FIN BALANCE &&&&&&&&&&&&&&&&&&&&&&&
    
    Dim i As Long, j As Long, k As Long, L As Long
    
    For i = 1 To UBound(Armaduras, 1)
        For j = 1 To UBound(Armaduras, 2)
            For k = 1 To UBound(Armaduras, 3)
                For L = 1 To UBound(Armaduras, 4)
                    Armaduras(i, j, k, L) = val(GetVar(IniPath & "Server.ini", "INIT", "Armadura" & i & j & k & L))
                Next
            Next
        Next
    Next

    Ullathorpe.map = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Mapa")
    Ullathorpe.X = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "X")
    Ullathorpe.Y = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Y")
    
    Nix.map = GetVar(DatPath & "Ciudades.dat", "Nix", "Mapa")
    Nix.X = GetVar(DatPath & "Ciudades.dat", "Nix", "X")
    Nix.Y = GetVar(DatPath & "Ciudades.dat", "Nix", "Y")
    
    Banderbill.map = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Mapa")
    Banderbill.X = GetVar(DatPath & "Ciudades.dat", "Banderbill", "X")
    Banderbill.Y = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Y")
    
    Lindos.map = GetVar(DatPath & "Ciudades.dat", "Lindos", "Mapa")
    Lindos.X = GetVar(DatPath & "Ciudades.dat", "Lindos", "X")
    Lindos.Y = GetVar(DatPath & "Ciudades.dat", "Lindos", "Y")
    
    Arghal.map = GetVar(DatPath & "Ciudades.dat", "Arghal", "Mapa")
    Arghal.X = GetVar(DatPath & "Ciudades.dat", "Arghal", "X")
    Arghal.Y = GetVar(DatPath & "Ciudades.dat", "Arghal", "Y")
    
    Ciudades(eCiudad.cUllathorpe) = Ullathorpe
    Ciudades(eCiudad.cNix) = Nix
    Ciudades(eCiudad.cBanderbill) = Banderbill
    Ciudades(eCiudad.cLindos) = Lindos
    Ciudades(eCiudad.cArghal) = Arghal
    
    Call MD5sCarga
    
    Call ConsultaPopular.LoadData

End Sub

Sub WriteVar(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'Escribe VAR en un archivo
'***************************************************

writeprivateprofilestring Main, Var, Value, File
    
End Sub

Sub SaveUser(ByVal UserIndex As Integer, ByVal UserFile As String)
'*************************************************
'Author: Unknown
'Last modified: 12/01/2010 (ZaMa)
'Saves the Users records
'23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
'12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados cuando pierden el efecto del mimetismo.
'*************************************************

On Error GoTo Errhandler

Dim OldUserHead As Long

With UserList(UserIndex)

    'ESTO TIENE QUE EVITAR ESE BUGAZO QUE NO SE POR QUE GRABA USUARIOS NULOS
    'clase=0 es el error, porq el enum empieza de 1!!
    If .Clase = 0 Or .Stats.ELV = 0 Then
        Call LogCriticEvent("Estoy intentantdo guardar un usuario nulo de nombre: " & .Name)
        Exit Sub
    End If
    
    If .flags.Mimetizado = 1 Then
        .Char.body = .CharMimetizado.body
        .Char.Head = .CharMimetizado.Head
        .Char.CascoAnim = .CharMimetizado.CascoAnim
        .Char.ShieldAnim = .CharMimetizado.ShieldAnim
        .Char.WeaponAnim = .CharMimetizado.WeaponAnim
        .Counters.Mimetismo = 0
        .flags.Mimetizado = 0
        ' Se fue el efecto del mimetismo, puede ser atacado por npcs
        .flags.Ignorado = False
    End If
    
    If FileExist(UserFile, vbNormal) Then
        If .flags.Muerto = 1 Then
            OldUserHead = .Char.Head
            .Char.Head = GetVar(UserFile, "INIT", "Head")
        End If
    '       Kill UserFile
    End If
    
    Dim LoopC As Integer
    
    
    Call WriteVar(UserFile, "FLAGS", "Muerto", CStr(.flags.Muerto))
    Call WriteVar(UserFile, "FLAGS", "Escondido", CStr(.flags.Escondido))
    Call WriteVar(UserFile, "FLAGS", "Hambre", CStr(.flags.Hambre))
    Call WriteVar(UserFile, "FLAGS", "Sed", CStr(.flags.Sed))
    Call WriteVar(UserFile, "FLAGS", "Desnudo", CStr(.flags.Desnudo))
    Call WriteVar(UserFile, "FLAGS", "Ban", CStr(.flags.Ban))
    Call WriteVar(UserFile, "FLAGS", "Navegando", CStr(.flags.Navegando))
    Call WriteVar(UserFile, "FLAGS", "Envenenado", CStr(.flags.Envenenado))
    Call WriteVar(UserFile, "FLAGS", "Paralizado", CStr(.flags.Paralizado))
    Call WriteVar(UserFile, "FLAGS", "IsLeader", CStr(.flags.IsLeader))
    
    'Matrix
    Call WriteVar(UserFile, "FLAGS", "LastMap", CStr(.flags.lastMap))
    
    'Eventos
    Call WriteVar(UserFile, "EVENTOS", "Quests", CStr(.Events.Quests))
    Call WriteVar(UserFile, "EVENTOS", "Torneos", CStr(.Events.Torneos))
    
    Call WriteVar(UserFile, "CONSEJO", "PERTENECE", IIf(.flags.Privilegios And PlayerType.RoyalCouncil, "1", "0"))
    Call WriteVar(UserFile, "CONSEJO", "PERTENECECAOS", IIf(.flags.Privilegios And PlayerType.ChaosCouncil, "1", "0"))
    
    
    Call WriteVar(UserFile, "COUNTERS", "Pena", CStr(.Counters.Pena))
    Call WriteVar(UserFile, "COUNTERS", "SkillsAsignados", CStr(.Counters.AsignedSkills))
    
    Call WriteVar(UserFile, "FACCIONES", "Bando", val(.Faccion.Bando))
    Call WriteVar(UserFile, "FACCIONES", "BandoOriginal", val(.Faccion.BandoOriginal))
    Call WriteVar(UserFile, "FACCIONES", "Matados0", val(.Faccion.Matados(0)))
    Call WriteVar(UserFile, "FACCIONES", "Matados1", val(.Faccion.Matados(1)))
    Call WriteVar(UserFile, "FACCIONES", "Matados2", val(.Faccion.Matados(2)))
    
    Call WriteVar(UserFile, "FACCIONES", "Jerarquia", val(.Faccion.Jerarquia))
    Call WriteVar(UserFile, "FACCIONES", "Ataco1", (.Faccion.Ataco(1) = 1))
    Call WriteVar(UserFile, "FACCIONES", "Ataco2", (.Faccion.Ataco(2) = 1))
    
    'Call WriteVar(UserFile, "FACCIONES", "Quests", val(.Faccion.Quests))
    Call WriteVar(UserFile, "FACCIONES", "Torneos", val(.Faccion.Torneos))

    
    '¿Fueron modificados los atributos del usuario?
    If Not .flags.TomoPocion Then
        For LoopC = 1 To UBound(.Stats.UserAtributos)
            Call WriteVar(UserFile, "ATRIBUTOS", "AT" & LoopC, CStr(.Stats.UserAtributos(LoopC)))
        Next LoopC
    Else
        For LoopC = 1 To UBound(.Stats.UserAtributos)
            '.Stats.UserAtributos(LoopC) = .Stats.UserAtributosBackUP(LoopC)
            Call WriteVar(UserFile, "ATRIBUTOS", "AT" & LoopC, CStr(.Stats.UserAtributosBackUP(LoopC)))
        Next LoopC
    End If
    
    For LoopC = 1 To UBound(.Stats.UserSkills)
        Call WriteVar(UserFile, "SKILLS", "SK" & LoopC, CStr(.Stats.UserSkills(LoopC)))
    Next LoopC
    
    
    Call WriteVar(UserFile, "CONTACTO", "Email", .email)
    
    Call WriteVar(UserFile, "INIT", "Genero", .Genero)
    Call WriteVar(UserFile, "INIT", "Raza", .raza)
    Call WriteVar(UserFile, "INIT", "Hogar", .Hogar)
    Call WriteVar(UserFile, "INIT", "Clase", .Clase)
    Call WriteVar(UserFile, "INIT", "Desc", .desc)
    
    Call WriteVar(UserFile, "INIT", "Heading", CStr(.Char.heading))
    
    Call WriteVar(UserFile, "INIT", "Head", CStr(.OrigChar.Head))
    
    If .flags.Muerto = 0 Then
        Call WriteVar(UserFile, "INIT", "Body", CStr(.Char.body))
    End If
    
    Call WriteVar(UserFile, "INIT", "Arma", CStr(.Char.WeaponAnim))
    Call WriteVar(UserFile, "INIT", "Escudo", CStr(.Char.ShieldAnim))
    Call WriteVar(UserFile, "INIT", "Casco", CStr(.Char.CascoAnim))
    
    #If ConUpTime Then
        Dim TempDate As Date
        TempDate = Now - .LogOnTime
        .LogOnTime = Now
        .UpTime = .UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + Hour(TempDate) * 3600 + Minute(TempDate) * 60 + Second(TempDate)
        .UpTime = .UpTime
        Call WriteVar(UserFile, "INIT", "UpTime", .UpTime)
    #End If
    
    'First time around?
    If GetVar(UserFile, "INIT", "LastIP1") = vbNullString Then
        Call WriteVar(UserFile, "INIT", "LastIP1", .ip & " - " & Date & ":" & time)
    'Is it a different ip from last time?
    ElseIf .ip <> Left$(GetVar(UserFile, "INIT", "LastIP1"), InStr(1, GetVar(UserFile, "INIT", "LastIP1"), " ") - 1) Then
        Dim i As Integer
        For i = 5 To 2 Step -1
            Call WriteVar(UserFile, "INIT", "LastIP" & i, GetVar(UserFile, "INIT", "LastIP" & CStr(i - 1)))
        Next i
        Call WriteVar(UserFile, "INIT", "LastIP1", .ip & " - " & Date & ":" & time)
    'Same ip, just update the date
    Else
        Call WriteVar(UserFile, "INIT", "LastIP1", .ip & " - " & Date & ":" & time)
    End If
    
    
    
    Call WriteVar(UserFile, "INIT", "Position", .Pos.map & "-" & .Pos.X & "-" & .Pos.Y)
    
    
    Call WriteVar(UserFile, "STATS", "GLD", CStr(.Stats.GLD))
    Call WriteVar(UserFile, "STATS", "BANCO", CStr(.Stats.Banco))
    
    Call WriteVar(UserFile, "STATS", "MaxHP", CStr(.Stats.MaxHp))
    Call WriteVar(UserFile, "STATS", "MinHP", CStr(.Stats.MinHp))
    
    Call WriteVar(UserFile, "STATS", "MaxSTA", CStr(.Stats.MaxSta))
    Call WriteVar(UserFile, "STATS", "MinSTA", CStr(.Stats.MinSta))
    
    Call WriteVar(UserFile, "STATS", "MaxMAN", CStr(.Stats.MaxMAN))
    Call WriteVar(UserFile, "STATS", "MinMAN", CStr(.Stats.MinMAN))
    
    Call WriteVar(UserFile, "STATS", "MaxHIT", CStr(.Stats.MaxHIT))
    Call WriteVar(UserFile, "STATS", "MinHIT", CStr(.Stats.MinHIT))
    
    Call WriteVar(UserFile, "STATS", "MaxAGU", CStr(.Stats.MaxAGU))
    Call WriteVar(UserFile, "STATS", "MinAGU", CStr(.Stats.MinAGU))
    
    Call WriteVar(UserFile, "STATS", "MaxHAM", CStr(.Stats.MaxHam))
    Call WriteVar(UserFile, "STATS", "MinHAM", CStr(.Stats.MinHam))
    
    Call WriteVar(UserFile, "STATS", "SkillPtsLibres", CStr(.Stats.SkillPts))
      
    Call WriteVar(UserFile, "STATS", "EXP", CStr(.Stats.Exp))
    Call WriteVar(UserFile, "STATS", "ELV", CStr(.Stats.ELV))
    
    
    Call WriteVar(UserFile, "STATS", "ELU", CStr(.Stats.ELU))
    Call WriteVar(UserFile, "MUERTES", "UserMuertes", CStr(.Stats.UsuariosMatados))
    'Call WriteVar(UserFile, "MUERTES", "CrimMuertes", Cstr$(.Stats.CriminalesMatados))
    Call WriteVar(UserFile, "MUERTES", "NpcsMuertes", CStr(.Stats.NPCsMuertos))
      
    '[KEVIN]----------------------------------------------------------------------------
    '*******************************************************************************************
    Call WriteVar(UserFile, "BancoInventory", "CantidadItems", val(.BancoInvent.NroItems))
    Dim loopd As Integer
    For loopd = 1 To MAX_BANCOINVENTORY_SLOTS
        Call WriteVar(UserFile, "BancoInventory", "Obj" & loopd, .BancoInvent.Object(loopd).OBJIndex & "-" & .BancoInvent.Object(loopd).Amount)
    Next loopd
    '*******************************************************************************************
    '[/KEVIN]-----------
      
    'Save Inv
    Call WriteVar(UserFile, "Inventory", "CantidadItems", val(.Invent.NroItems))
    
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        Call WriteVar(UserFile, "Inventory", "Obj" & LoopC, .Invent.Object(LoopC).OBJIndex & "-" & .Invent.Object(LoopC).Amount & "-" & .Invent.Object(LoopC).Equipped)
    Next LoopC
    
    Call WriteVar(UserFile, "Inventory", "WeaponEqpSlot", CStr(.Invent.WeaponEqpSlot))
    Call WriteVar(UserFile, "Inventory", "ArmourEqpSlot", CStr(.Invent.ArmourEqpSlot))
    Call WriteVar(UserFile, "Inventory", "CascoEqpSlot", CStr(.Invent.CascoEqpSlot))
    Call WriteVar(UserFile, "Inventory", "EscudoEqpSlot", CStr(.Invent.EscudoEqpSlot))
    Call WriteVar(UserFile, "Inventory", "BarcoSlot", CStr(.Invent.BarcoSlot))
    Call WriteVar(UserFile, "Inventory", "MunicionSlot", CStr(.Invent.MunicionEqpSlot))
    Call WriteVar(UserFile, "Inventory", "MochilaSlot", CStr(.Invent.MochilaEqpSlot))
    Call WriteVar(UserFile, "Inventory", "HerramientaSlot", CStr(.Invent.HerramientaEqpslot))
    '/Nacho
    
    Dim cad As String
    
    For LoopC = 1 To MAXUSERHECHIZOS
        cad = .Stats.UserHechizos(LoopC)
        Call WriteVar(UserFile, "HECHIZOS", "H" & LoopC, cad)
    Next
    
    Dim NroMascotas As Long
    NroMascotas = .NroMascotas
    
    For LoopC = 1 To MAXMASCOTAS
        ' Mascota valida?
        If .MascotasIndex(LoopC) > 0 Then
            ' Nos aseguramos que la criatura no fue invocada
            If Npclist(.MascotasIndex(LoopC)).Contadores.TiempoExistencia = 0 Then
                cad = .MascotasType(LoopC)
            Else 'Si fue invocada no la guardamos
                cad = "0"
                NroMascotas = NroMascotas - 1
            End If
            Call WriteVar(UserFile, "MASCOTAS", "MAS" & LoopC, cad)
        Else
            cad = .MascotasType(LoopC)
            Call WriteVar(UserFile, "MASCOTAS", "MAS" & LoopC, cad)
        End If
    
    Next
    
    Call WriteVar(UserFile, "MASCOTAS", "NroMascotas", CStr(NroMascotas))
    
    For LoopC = 1 To 3
        Call WriteVar(UserFile, "RECOMPENSAS", "Recompensa" & LoopC, val(.Recompensas(LoopC)))
    Next LoopC
    
    Call WriteVar(UserFile, "Guild", "GUILDID", .GuildID)
    Call WriteVar(UserFile, "Guild", "RequestedTo", .flags.WaitingApprovement)
    
    'Devuelve el head de muerto
    If .flags.Muerto = 1 Then
        .Char.Head = iCabezaMuerto
    End If
End With

Exit Sub

Errhandler:
Call LogError("Error en SaveUser")

End Sub

Function Criminal(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 10/1/2017 - GoDKeR
'
'***************************************************

    Criminal = (UserList(UserIndex).Faccion.Bando = eFaccion.Caos)

End Function

Function Neutro(ByVal UserIndex As Integer) As Boolean
    Neutro = (UserList(UserIndex).Faccion.Bando = eFaccion.Neutral)
End Function

Sub BackUPnPc(NpcIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim NpcNumero As Integer
    Dim npcfile As String
    Dim LoopC As Integer
    
    
    NpcNumero = Npclist(NpcIndex).Numero
    
    'If NpcNumero > 499 Then
    '    npcfile = DatPath & "bkNPCs-HOSTILES.dat"
    'Else
        npcfile = DatPath & "bkNPCs.dat"
    'End If
    
    With Npclist(NpcIndex)
        'General
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Name", .Name)
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Desc", .desc)
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Head", val(.Char.Head))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Body", val(.Char.body))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Heading", val(.Char.heading))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Movement", val(.Movement))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Attackable", val(.Attackable))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Comercia", val(.Comercia))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "TipoItems", val(.TipoItems))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", val(.Hostile))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveEXP", val(.GiveEXP))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveGLD", val(.GiveGLD))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", val(.Hostile))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "InvReSpawn", val(.InvReSpawn))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "NpcType", val(.NPCtype))
        
        
        'Stats
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Alineacion", val(.Stats.Alineacion))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "DEF", val(.Stats.def))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHit", val(.Stats.MaxHIT))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHp", val(.Stats.MaxHp))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHit", val(.Stats.MinHIT))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHp", val(.Stats.MinHp))
        
        
        
        
        'Flags
        Call WriteVar(npcfile, "NPC" & NpcNumero, "ReSpawn", val(.flags.Respawn))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "BackUp", val(.flags.BackUp))
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Domable", val(.flags.Domable))
        
        'Inventario
        Call WriteVar(npcfile, "NPC" & NpcNumero, "NroItems", val(.Invent.NroItems))
        If .Invent.NroItems > 0 Then
           For LoopC = 1 To MAX_INVENTORY_SLOTS
                Call WriteVar(npcfile, "NPC" & NpcNumero, "Obj" & LoopC, .Invent.Object(LoopC).OBJIndex & "-" & .Invent.Object(LoopC).Amount)
           Next LoopC
        End If
    End With

End Sub

Sub CargarNpcBackUp(NpcIndex As Integer, ByVal NpcNumber As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    'Status
    If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando backup Npc"
    
    Dim npcfile As String
    
    'If NpcNumber > 499 Then
    '    npcfile = DatPath & "bkNPCs-HOSTILES.dat"
    'Else
        npcfile = DatPath & "bkNPCs.dat"
    'End If
    
    With Npclist(NpcIndex)
    
        .Numero = NpcNumber
        .Name = GetVar(npcfile, "NPC" & NpcNumber, "Name")
        .desc = GetVar(npcfile, "NPC" & NpcNumber, "Desc")
        .Movement = val(GetVar(npcfile, "NPC" & NpcNumber, "Movement"))
        .NPCtype = val(GetVar(npcfile, "NPC" & NpcNumber, "NpcType"))
        
        .Char.body = val(GetVar(npcfile, "NPC" & NpcNumber, "Body"))
        .Char.Head = val(GetVar(npcfile, "NPC" & NpcNumber, "Head"))
        .Char.heading = val(GetVar(npcfile, "NPC" & NpcNumber, "Heading"))
        
        .Attackable = val(GetVar(npcfile, "NPC" & NpcNumber, "Attackable"))
        .Comercia = val(GetVar(npcfile, "NPC" & NpcNumber, "Comercia"))
        .Hostile = val(GetVar(npcfile, "NPC" & NpcNumber, "Hostile"))
        .GiveEXP = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveEXP"))
        
        
        .GiveGLD = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveGLD"))
        
        .InvReSpawn = val(GetVar(npcfile, "NPC" & NpcNumber, "InvReSpawn"))
        
        .Stats.MaxHp = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHP"))
        .Stats.MinHp = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHP"))
        .Stats.MaxHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHIT"))
        .Stats.MinHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHIT"))
        .Stats.def = val(GetVar(npcfile, "NPC" & NpcNumber, "DEF"))
        .Stats.Alineacion = val(GetVar(npcfile, "NPC" & NpcNumber, "Alineacion"))
        
        
        
        Dim LoopC As Integer
        Dim ln As String
        .Invent.NroItems = val(GetVar(npcfile, "NPC" & NpcNumber, "NROITEMS"))
        If .Invent.NroItems > 0 Then
            For LoopC = 1 To MAX_INVENTORY_SLOTS
                ln = GetVar(npcfile, "NPC" & NpcNumber, "Obj" & LoopC)
                .Invent.Object(LoopC).OBJIndex = val(ReadField(1, ln, 45))
                .Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
               
            Next LoopC
        Else
            For LoopC = 1 To MAX_INVENTORY_SLOTS
                .Invent.Object(LoopC).OBJIndex = 0
                .Invent.Object(LoopC).Amount = 0
            Next LoopC
        End If
        
        .flags.NPCActive = True
        .flags.Respawn = val(GetVar(npcfile, "NPC" & NpcNumber, "ReSpawn"))
        .flags.BackUp = val(GetVar(npcfile, "NPC" & NpcNumber, "BackUp"))
        .flags.Domable = val(GetVar(npcfile, "NPC" & NpcNumber, "Domable"))
        .flags.RespawnOrigPos = val(GetVar(npcfile, "NPC" & NpcNumber, "OrigPos"))
        
        'Tipo de items con los que comercia
        .TipoItems = val(GetVar(npcfile, "NPC" & NpcNumber, "TipoItems"))
    End With

End Sub


Sub LogBan(ByVal BannedIndex As Integer, ByVal UserIndex As Integer, ByVal Motivo As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Call WriteVar(App.path & "\logs\" & "BanDetail.log", UserList(BannedIndex).Name, "BannedBy", UserList(UserIndex).Name)
    Call WriteVar(App.path & "\logs\" & "BanDetail.log", UserList(BannedIndex).Name, "Reason", Motivo)
    
    'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
    Dim mifile As Integer
    mifile = FreeFile
    Open App.path & "\logs\GenteBanned.log" For Append Shared As #mifile
    Print #mifile, UserList(BannedIndex).Name
    Close #mifile

End Sub


Sub LogBanFromName(ByVal BannedName As String, ByVal UserIndex As Integer, ByVal Motivo As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Call WriteVar(App.path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", UserList(UserIndex).Name)
    Call WriteVar(App.path & "\logs\" & "BanDetail.dat", BannedName, "Reason", Motivo)
    
    'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
    Dim mifile As Integer
    mifile = FreeFile
    Open App.path & "\logs\GenteBanned.log" For Append Shared As #mifile
    Print #mifile, BannedName
    Close #mifile

End Sub


Sub Ban(ByVal BannedName As String, ByVal Baneador As String, ByVal Motivo As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Call WriteVar(App.path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", Baneador)
    Call WriteVar(App.path & "\logs\" & "BanDetail.dat", BannedName, "Reason", Motivo)
    
    
    'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
    Dim mifile As Integer
    mifile = FreeFile
    Open App.path & "\logs\GenteBanned.log" For Append Shared As #mifile
    Print #mifile, BannedName
    Close #mifile

End Sub

Public Sub CargaApuestas()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Apuestas.Ganancias = val(GetVar(DatPath & "apuestas.dat", "Main", "Ganancias"))
    Apuestas.Perdidas = val(GetVar(DatPath & "apuestas.dat", "Main", "Perdidas"))
    Apuestas.Jugadas = val(GetVar(DatPath & "apuestas.dat", "Main", "Jugadas"))

End Sub

Public Sub generateMatrix()
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim j As Long
Dim i As Long

ReDim distanceToCities(1 To NumMaps) As HomeDistance

For j = 1 To NUMCIUDADES
    For i = 1 To NumMaps
        distanceToCities(i).distanceToCity(j) = -1
    Next i
Next j

For j = 1 To NUMCIUDADES
    For i = 1 To 4
        Select Case i
            Case eHeading.NORTH
                Call setDistance(getLimit(Ciudades(j).map, eHeading.NORTH), j, i, 0, 1)
            Case eHeading.EAST
                Call setDistance(getLimit(Ciudades(j).map, eHeading.EAST), j, i, 1, 0)
            Case eHeading.SOUTH
                Call setDistance(getLimit(Ciudades(j).map, eHeading.SOUTH), j, i, 0, 1)
            Case eHeading.WEST
                Call setDistance(getLimit(Ciudades(j).map, eHeading.WEST), j, i, -1, 0)
        End Select
    Next i
Next j

End Sub

Public Sub setDistance(ByVal mapa As Integer, ByVal city As Byte, ByVal side As Integer, Optional ByVal X As Integer = 0, Optional ByVal Y As Integer = 0)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim i As Integer
Dim lim As Integer

If mapa <= 0 Or mapa > NumMaps Then Exit Sub

If distanceToCities(mapa).distanceToCity(city) >= 0 Then Exit Sub

If mapa = Ciudades(city).map Then
    distanceToCities(mapa).distanceToCity(city) = 0
Else
    distanceToCities(mapa).distanceToCity(city) = Abs(X) + Abs(Y)
End If

For i = 1 To 4
    lim = getLimit(mapa, i)
    If lim > 0 Then
        Select Case i
            Case eHeading.NORTH
                Call setDistance(lim, city, i, X, Y + 1)
            Case eHeading.EAST
                Call setDistance(lim, city, i, X + 1, Y)
            Case eHeading.SOUTH
                Call setDistance(lim, city, i, X, Y - 1)
            Case eHeading.WEST
                Call setDistance(lim, city, i, X - 1, Y)
        End Select
    End If
Next i
End Sub

Public Function getLimit(ByVal mapa As Integer, ByVal side As Byte) As Integer
'***************************************************
'Author: Budi
'Last Modification: 31/01/2010
'Retrieves the limit in the given side in the given map.
'TODO: This should be set in the .inf map file.
'***************************************************
Dim i, X, Y As Integer

If mapa <= 0 Then Exit Function

For X = 15 To 87
    For Y = 0 To 3
        Select Case side
            Case eHeading.NORTH
                getLimit = MapData(mapa, X, 7 + Y).TileExit.map
            Case eHeading.EAST
                getLimit = MapData(mapa, 92 - Y, X).TileExit.map
            Case eHeading.SOUTH
                getLimit = MapData(mapa, X, 94 - Y).TileExit.map
            Case eHeading.WEST
                getLimit = MapData(mapa, 9 + Y, X).TileExit.map
        End Select
        If getLimit > 0 Then Exit Function
    Next Y
Next X
End Function
