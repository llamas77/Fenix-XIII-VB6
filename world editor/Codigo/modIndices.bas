Attribute VB_Name = "modIndices"
Option Explicit

''
' Carga los indices de Graficos
'

Public Sub CargarIndicesDeGraficos()

On Error GoTo ErrorHandler

    Dim Grh As Long
    Dim Frame As Long
    Dim grhCount As Long
    Dim handle As Integer
    Dim fileVersion As Long
    
    If FileExist(DirIndex & "Graficos.ind", vbArchive) = False Then
        MsgBox "Falta el archivo 'graficos.ind' en " & DirIndex, vbCritical
        End
    End If
    
    'Open files
    handle = FreeFile()
    
    Open DirIndex & "Graficos.ind" For Binary Access Read As handle
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
            If .NumFrames <= 0 Then GoTo ErrorHandler
            
            'ReDim .Frames(1 To GrhData(Grh).NumFrames)
            
            If .NumFrames > 1 Then
                'Read a animation GRH set
                For Frame = 1 To .NumFrames
                    Get handle, , .Frames(Frame)
                    If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then
                        GoTo ErrorHandler
                    End If
                Next Frame
                
                Get handle, , .Speed
                
                If .Speed <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .pixelHeight = GrhData(.Frames(1)).pixelHeight
                If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                .pixelWidth = GrhData(.Frames(1)).pixelWidth
                If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                .TileWidth = GrhData(.Frames(1)).TileWidth
                If .TileWidth <= 0 Then GoTo ErrorHandler
                
                .TileHeight = GrhData(.Frames(1)).TileHeight
                If .TileHeight <= 0 Then GoTo ErrorHandler
            Else
                'Read in normal GRH data
                Get handle, , .FileNum
                If .FileNum <= 0 Then GoTo ErrorHandler
                
                Get handle, , GrhData(Grh).sX
                If .sX < 0 Then GoTo ErrorHandler
                
                Get handle, , .sY
                If .sY < 0 Then GoTo ErrorHandler
                
                Get handle, , .pixelWidth
                If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                Get handle, , .pixelHeight
                If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .TileWidth = .pixelWidth / TilePixelHeight
                .TileHeight = .pixelHeight / TilePixelWidth
                
                .Frames(1) = Grh
            End If
        End With
        
        Get handle, , Grh
    Loop
    
    Close handle
Exit Sub

ErrorHandler:
Close handle
    MsgBox "Error al intentar cargar el Grh número " & Grh & " de graficos.ind en " & DirIndex & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub


Public Sub CargarIndicesDeAgua()
'*************************************************
'Author: Deut
'Last modified: 12/12/09
'*************************************************
   On Error GoTo CargarIndicesDeAgua_Error
If FileExist(DirInterno & "\AGUAS.dat", vbArchive) = False Then
    MsgBox "Falta el archivo 'AGUAS.dat' en " & DirDats, vbCritical
    End
End If
Dim Leer As New clsIniReader
Dim i As Integer
Dim NroAguas As Integer
Leer.Initialize (DirInterno & "\AGUAS.dat")
NroAguas = Val(Leer.GetValue("INIT", "NroAguas"))
ReDim REFAguasArr(1 To NroAguas)
For i = 1 To NroAguas
    REFAguasArr(i) = Val(Leer.GetValue("INIT", "RefAgua" & i))
Next


   On Error GoTo 0
   Exit Sub
CargarIndicesDeAgua_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CargarIndicesDeAgua of Módulo modIndices"
End Sub

''
' Carga los indices de Superficie
'

Public Sub CargarIndicesSuperficie()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************

On Error GoTo Fallo
    If FileExist(IniPath & "GrhIndex\indices.ini", vbArchive) = False Then
        MsgBox "Falta el archivo 'GrhIndex\indices.ini'", vbCritical
        End
    End If
    Dim Leer As New clsIniReader
    Dim i As Integer
    Leer.Initialize IniPath & "GrhIndex\indices.ini"
    MaxSup = Leer.GetValue("INIT", "Referencias")
    ReDim SupData(MaxSup) As SupData
    frmMain.lListado(0).Clear
    For i = 0 To MaxSup
        SupData(i).Name = Leer.GetValue("REFERENCIA" & i, "Nombre")
        SupData(i).Grh = Val(Leer.GetValue("REFERENCIA" & i, "GrhIndice"))
        SupData(i).Width = Val(Leer.GetValue("REFERENCIA" & i, "Ancho"))
        SupData(i).Height = Val(Leer.GetValue("REFERENCIA" & i, "Alto"))
        SupData(i).Block = IIf(Val(Leer.GetValue("REFERENCIA" & i, "Bloquear")) = 1, True, False)
        SupData(i).Capa = Val(Leer.GetValue("REFERENCIA" & i, "Capa"))
        frmMain.lListado(0).AddItem SupData(i).Name & " - #" & i
    Next
    DoEvents
    Exit Sub
Fallo:
    MsgBox "Error al intentar cargar el indice " & i & " de GrhIndex\indices.ini" & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly
End Sub

''
' Carga los indices de Objetos
'

Public Sub CargarIndicesOBJ()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

On Error GoTo Fallo
    If FileExist(DirDats & "\OBJ.dat", vbArchive) = False Then
        MsgBox "Falta el archivo 'OBJ.dat' en " & DirDats, vbCritical
        End
    End If
    Dim Obj As Integer
    Dim Leer As New clsIniReader
    Call Leer.Initialize(DirDats & "\OBJ.dat")
    frmMain.lListado(3).Clear
    NumOBJs = Val(Leer.GetValue("INIT", "NumOBJs"))
    ReDim ObjData(1 To NumOBJs) As ObjData
    For Obj = 1 To NumOBJs
        frmCargando.X.Caption = "Cargando Datos de Objetos..." & Obj & "/" & NumOBJs
        DoEvents
        ObjData(Obj).Name = Leer.GetValue("OBJ" & Obj, "Name")
        ObjData(Obj).GrhIndex = Val(Leer.GetValue("OBJ" & Obj, "GrhIndex"))
        ObjData(Obj).ObjType = Val(Leer.GetValue("OBJ" & Obj, "ObjType"))
        ObjData(Obj).Ropaje = Val(Leer.GetValue("OBJ" & Obj, "NumRopaje"))
        ObjData(Obj).Info = Leer.GetValue("OBJ" & Obj, "Info")
        ObjData(Obj).WeaponAnim = Val(Leer.GetValue("OBJ" & Obj, "Anim"))
        ObjData(Obj).Texto = Leer.GetValue("OBJ" & Obj, "Texto")
        ObjData(Obj).GrhSecundario = Val(Leer.GetValue("OBJ" & Obj, "GrhSec"))
        frmMain.lListado(3).AddItem ObjData(Obj).Name & " - #" & Obj
    Next Obj
    Exit Sub
Fallo:
MsgBox "Error al intentar cargar el Objteto " & Obj & " de OBJ.dat en " & DirDats & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub

''
' Carga los indices de Triggers
'

Public Sub CargarIndicesTriggers()
'*************************************************
'Author: ^[GS]^
'Last modified: 15/12/09 - Deut
'*************************************************

On Error GoTo Fallo
    If FileExist(DirInterno & "Triggers.ini", vbArchive) = False Then
        MsgBox "Falta el archivo 'Triggers.ini' en " & DirIndex, vbCritical
        End
    End If
    Dim NumT As Byte
    Dim T As Byte
    Dim Leer As New clsIniReader
    Call Leer.Initialize(DirInterno & "Triggers.ini")
    frmMain.lListado(4).Clear
    NumT = Val(Leer.GetValue("INIT", "NumTriggers"))
    For T = 1 To NumT
         frmMain.lListado(4).AddItem Leer.GetValue("Trig" & T, "Name") & " - #" & (T)
    Next T

Exit Sub
Fallo:
    MsgBox "Error al intentar cargar el Trigger " & T & " de Triggers.ini en " & DirIndex & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub

''
' Carga los indices de Cuerpos
'

Public Sub CargarIndicesDeCuerpos()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************

On Error GoTo Fallo
    If FileExist(DirIndex & "Personajes.ind", vbArchive) = False Then
        MsgBox "Falta el archivo 'Personajes.ind' en " & DirIndex, vbCritical
        End
    End If
    Dim N As Integer
    Dim i As Integer
    N = FreeFile
    Open DirIndex & "Personajes.ind" For Binary Access Read As #N
    'cabecera
    Get #N, , MiCabecera
    'num de cabezas
    Get #N, , NumBodies
    'Resize array
    ReDim BodyData(0 To NumBodies + 1) As tBodyData
    ReDim MisCuerpos(0 To NumBodies + 1) As tIndiceCuerpo
    For i = 1 To NumBodies
        Get #N, , MisCuerpos(i)
        InitGrh BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
        InitGrh BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
        InitGrh BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
        InitGrh BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0
        BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
        BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
    Next i
    Close #N
Exit Sub
Fallo:
    MsgBox "Error al intentar cargar el Cuerpo " & i & " de Personajes.ind en " & DirIndex & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub

''
' Carga los indices de Cabezas
'

Public Sub CargarIndicesDeCabezas()
On Error GoTo Fallo
    If FileExist(DirIndex & "Cabezas.ind", vbArchive) = False Then
        MsgBox "Falta el archivo 'Cabezas.ind' en " & DirIndex, vbCritical
        End
    End If
    Dim N As Integer
    Dim i As Integer
    Dim MisCabezas() As tIndiceCabeza
    N = FreeFile
    Open DirIndex & "Cabezas.ind" For Binary Access Read As #N
    'cabecera
    Get #N, , MiCabecera
    'num de cabezas
    Get #N, , Numheads
    'Resize array
    ReDim HeadData(0 To Numheads + 1) As tHeadData
    ReDim MisCabezas(0 To Numheads + 1) As tIndiceCabeza
    For i = 1 To Numheads
        Get #N, , MisCabezas(i)
        InitGrh HeadData(i).Head(1), MisCabezas(i).Head(1), 0
        InitGrh HeadData(i).Head(2), MisCabezas(i).Head(2), 0
        InitGrh HeadData(i).Head(3), MisCabezas(i).Head(3), 0
        InitGrh HeadData(i).Head(4), MisCabezas(i).Head(4), 0
    Next i
    Close #N
Exit Sub
Fallo:
    MsgBox "Error al intentar cargar la Cabeza " & i & " de Cabezas.ind en " & DirIndex & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub

''
' Carga los indices de NPCs
'

Public Sub CargarIndicesNPC()
'*************************************************
'Author: ^[GS]^
'Last modified: 26/05/06
'*************************************************

On Error GoTo Fallo
    If FileExist(DirDats & "\NPCs.dat", vbArchive) = False Then
        MsgBox "Falta el archivo 'NPCs.dat' en " & DirDats, vbCritical
        End
    End If
    
    Dim Trabajando As String
    Dim NPC As Integer
    Dim Leer As New clsIniReader
    frmMain.lListado(1).Clear
    frmMain.lListado(2).Clear
    Call Leer.Initialize(DirDats & "\NPCs.dat")
    NumNPCs = Val(Leer.GetValue("INIT", "NumNPCs"))
    'Call Leer.Initialize(DirDats & "\NPCs-HOSTILES.dat")
    'NumNPCsHOST = Val(Leer.GetValue("INIT", "NumNPCs"))
    ReDim NpcData(1 To NumNPCs) As NpcData
    Trabajando = "Dats\NPCs.dat"
    Call Leer.Initialize(DirDats & "\NPCs.dat")
    For NPC = 1 To NumNPCs
        NpcData(NPC).Name = Leer.GetValue("NPC" & NPC, "Name")
        NpcData(NPC).Body = Val(Leer.GetValue("NPC" & NPC, "Body"))
        NpcData(NPC).Head = Val(Leer.GetValue("NPC" & NPC, "Head"))
        NpcData(NPC).Heading = Val(Leer.GetValue("NPC" & NPC, "Heading"))
        If LenB(NpcData(NPC).Name) <> 0 Then frmMain.lListado(1).AddItem NpcData(NPC).Name & " - #" & NPC
    Next NPC

    Exit Sub
Fallo:
    MsgBox "Error al intentar cargar el NPC " & NPC & " de " & Trabajando & " en " & DirDats & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub
