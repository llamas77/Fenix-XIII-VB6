VERSION 5.00
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Argentum Online"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox Render 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   11520
      Left            =   0
      ScaleHeight     =   768
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1024
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   15360
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements DirectXEvent8

Public btnLogin As Integer
Public btnCrearPj As Integer
Public txtNombre As Integer
Public txtPassword As Integer

'//details//
Public cmbHogar As Integer
Public cmbRaza As Integer
Public cmbSexo As Integer
Public btnHeadDer As Integer
Public btnHeadIzq As Integer
'//details//

'//info//
Public txtNick As Integer
Public txtMail As Integer
Public txtPass As Integer
Public txtRepPass As Integer
'//info//

'//attrib//
Public btnDados As Integer
Public lblFuerza As Integer
Public lblAgilidad As Integer
Public lblConstitucion As Integer
Public lblInteligencia As Integer
Public lblCarisma As Integer
'//attrib

'//skills
Public lstSkills As Integer
Public lblSkillLibres As Integer
'//skills

Public btnSiguiente As Integer
Public btnAtras As Integer

Public MouseX As Integer
Public MouseY As Integer

Public Loaded As Boolean

Public SkillPts As Integer
Private uSkills(1 To NUMSKILLS) As Byte

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

'CSEH: ErrLog
Public Sub DirectXEvent8_DXCallback(ByVal EventID As Long)
    '*****************************************************************
    'Handles mouse device events (movement, clicking, mouse wheel scrolling, etc)
    'More info: http://www.vbgore.com/GameClient.frmMain.DirectXEvent8_DXCallback
    '*****************************************************************
    '<EhHeader>
    On Error GoTo DirectXEvent8_DXCallback_Err
    '</EhHeader>
    Dim DevData(1 To 50) As DIDEVICEOBJECTDATA
    Dim NumEvents As Long
    Dim loopC As Long
    Dim Moved As Byte
    Dim OldMousePos As Position
    
        'Check if message is for us
100     If EventID <> MouseEvent Then Exit Sub
        If Me.WindowState <> 0 Then Exit Sub

        'Retrieve data
105     NumEvents = DIDevice.GetDeviceData(DevData, DIGDD_DEFAULT)

        'Loop through data
110     For loopC = 1 To NumEvents
115         Select Case DevData(loopC).lOfs

            'Mouse wheel is scrolled
            Case DIMOFS_Z
                Dim c As Integer
            
120             c = Collision(MouseX, MouseY)
            
                'Scroll the chat buffer if the cursor is over the chat buffer window
125             If c <> -1 Then
130                 If DevData(loopC).lData > 0 Then
135                     Call mod_Components.Execute(c, eComponentEvent.MouseScrollUp)
140                 ElseIf DevData(loopC).lData < 0 Then
145                     Call mod_Components.Execute(c, eComponentEvent.MouseScrollDown)
                    End If

150                 GoTo NextLoopC
                End If

            End Select
        
NextLoopC:

155     Next loopC

    '<EhFooter>
    Exit Sub

DirectXEvent8_DXCallback_Err:
        Call LogError("Error en DirectXEvent8_DXCallback: " & Erl & " - " & Err.Description)
    '</EhFooter>
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        prgRun = False
    End If
End Sub

Private Sub Form_Load()
    '[CODE 002]:MatuX
    EngineRun = True
    
    LoadComponents
End Sub

Public Sub LoadComponents()
    
    If Loaded Then Exit Sub
    
    txtNombre = AddTextBox(429, 357, 170, 19, Black, White)
    txtPassword = AddTextBox(429, 407, 170, 19, Black, White, , True)
    btnLogin = AddRect(406, 438, 100, 37)
    btnCrearPj = AddRect(519, 438, 100, 37)
    
    Call SetEvents(btnLogin, Callback(AddressOf btnLogin_EventHandler))
    Call SetEvents(btnCrearPj, Callback(AddressOf btnNewCharacter_EventHandler))
    
    'details
    cmbHogar = AddComboBox(424, 289, 176, 22, Black)
    cmbRaza = AddComboBox(424, 339, 176, 22, Black)
    cmbSexo = AddComboBox(424, 389, 176, 22, Black)
    btnHeadDer = AddRect(586, 436, 30, 23)
    btnHeadIzq = AddRect(416, 436, 30, 23)
    
    Call SetEvents(btnHeadDer, Callback(AddressOf btnHeadDer_EventHandler))
    Call SetEvents(btnHeadIzq, Callback(AddressOf btnHeadIzq_EventHandler))
    Call SetEvents(cmbRaza, Callback(AddressOf cmbRaza_EventHandler))
    Call SetEvents(cmbSexo, Callback(AddressOf cmbSexo_EventHandler))
    
    'info
    txtNick = AddTextBox(425, 289, 176, 22, Black, White)
    txtMail = AddTextBox(425, 339, 176, 22, Black, White)
    txtPass = AddTextBox(425, 389, 176, 22, Black, White, , True)
    txtRepPass = AddTextBox(425, 439, 176, 22, Black, White, , True)
    
    'attribs
    lblFuerza = AddLabel("0", 485, 335, White)
    lblAgilidad = AddLabel("0", 485, 385, White)
    lblConstitucion = AddLabel("0", 485, 435, White)
    lblInteligencia = AddLabel("0", 485, 485, White)
    lblCarisma = AddLabel("0", 485, 535, White)
    btnDados = AddRect(496, 276, 32, 32)
    
    SkillPts = 10
    
    'sks
    lstSkills = AddFillableListBox(425, 289, 152, 244, Transparent, 10)
    lblSkillLibres = AddLabel("Puntos disponibles: " & SkillPts, 5, 45, White)
    
    Call SetEvents(lstSkills, Callback(AddressOf lstSkill_EventHandler))
    
    btnSiguiente = AddRect(520, 578, 100, 37)
    btnAtras = AddRect(408, 578, 100, 37)
    
    Call SetChild(txtPass, txtRepPass)
    
    Call SetEvents(btnSiguiente, Callback(AddressOf btnSiguiente_EventHandler))
    Call SetEvents(btnAtras, Callback(AddressOf btnAtras_EventHandler))
    Call SetEvents(txtRepPass, Callback(AddressOf txtRepPass_EventHandler))
    Call SetEvents(btnDados, Callback(AddressOf btnDados_EventHandler))
    
    Dim i As Long
    
    For i = 1 To NUMCIUDADES
        Call InsertText(cmbHogar, Ciudades(i), White) 'todo color de faccion por ciudad ;)
    Next
    
    For i = 1 To NUMRAZAS
        Call InsertText(cmbRaza, ListaRazas(i), White)
    Next
    
    Call InsertText(cmbSexo, "Hombre", White)
    Call InsertText(cmbSexo, "Mujer", White)
    
    For i = 1 To NUMSKILLS
        Call InsertText(lstSkills, SkillsNames(i), White)
    Next

    Call DisableComponents(btnAtras, btnSiguiente, btnHeadDer, btnHeadIzq)
    
    Call HideComponents(txtNick, txtPass, txtMail, txtRepPass, lblFuerza, lblAgilidad, lblInteligencia, lblConstitucion, _
                        lblCarisma, cmbHogar, cmbSexo, cmbRaza, lstSkills, lblSkillLibres)
                        
    Loaded = True
End Sub
Private Sub Render_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyTab Then
        Call mod_Components.TabComponent
    ElseIf KeyAscii = vbKeyReturn Then
        Call LoginUser
    Else
    
        Dim c As Integer
        
        c = mod_Components.GetFocused()
        
        If c <> -1 Then
            Call mod_Components.Execute(c, eComponentEvent.KeyPress, KeyAscii, IIf(Components(c).PasswChr, 42, 0))
        End If
    End If
End Sub

Private Sub Render_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim c As Integer
    
    c = Collision(X, Y)
    
    If c <> -1 Then
        Call mod_Components.Execute(c, eComponentEvent.MouseDown, IntegersToLong(X, Y), IntegersToLong(Button, Shift))
    End If
End Sub

Private Sub Render_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseX = X
MouseY = Y
End Sub

Private Sub txtPasswd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call LoginUser
End Sub

Private Sub Render_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim c As Integer
    
    c = Collision(X, Y)
    
    If c <> -1 Then
        Call mod_Components.Execute(c, eComponentEvent.MouseUp, IntegersToLong(X, Y))
    End If
End Sub

Public Sub LoginNewChar()
    Dim i As Integer
    Dim CharAscii As Byte
    
    UserName = GetComponentText(txtNick)
            
    'If Right$(UserName, 1) = " " Then
    '    UserName = RTrim$(UserName)
    '    MsgBox "Nombre invalido, se han removido los espacios al final del nombre"
    'End If
    
    UserRaza = Components(cmbRaza).SelIndex
    UserSexo = Components(cmbSexo).SelIndex
    
    
    'UserAtributos(eAtributos.Fuerza) = Val(GetComponentText(lblFuerza))
    'UserAtributos(eAtributos.Agilidad) = Val(GetComponentText(lblAgilidad))
    'UserAtributos(eAtributos.Constitucion) = Val(GetComponentText(lblConstitucion))
    'UserAtributos(eAtributos.Inteligencia) = Val(GetComponentText(lblInteligencia))
    'UserAtributos(eAtributos.Carisma) = Val(GetComponentText(lblCarisma))
    
    UserHogar = Components(cmbHogar).SelIndex
    
    If Not CheckData Then Exit Sub

    UserPassword = GetComponentText(txtPass)
    
    For i = 1 To Len(UserPassword)
        CharAscii = Asc(mid$(UserPassword, i, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Password inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Sub
        End If
    Next i
    
    UserEmail = GetComponentText(txtMail)
    
    If SkillPts <> 0 Then
        MsgBox "Debes asignar todos los puntos disponibles en los skills."
        Exit Sub
    End If
    
    'Call CopyMemory(UserSkills(1), uSkills(1), NUMSKILLS)
        
    For i = 1 To NUMSKILLS
        UserSkills(i) = GetComponentValue(lstSkills, i)
    Next
    
    EstadoLogin = E_MODO.CrearNuevoPj
    
    If Not frmMain.Socket1.Connected Then

        'MsgBox "Error: Se ha perdido la conexion con el server."
        
        frmMain.Socket1.HostName = CurServerIP
        frmMain.Socket1.RemotePort = CurServerPort
    
        frmMain.Socket1.Connect
        
        Call Login
    Else
        
        Call Login
        
    End If
    
    bShowTutorial = True
End Sub

Public Sub CloseNewChar()

frmMain.Socket1.Disconnect

Call HideComponents(txtNick, txtPass, txtRepPass, txtMail)
Call EnableComponents(btnLogin, btnCrearPj)
Call ShowComponents(txtNombre, txtPassword)
Call ChangeRenderState(eRenderState.eLogin)

End Sub

Public Sub LoginUser()
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If
        
    'update user info
    UserName = GetComponentText(frmConnect.txtNombre)
    
    Dim aux As String
    aux = GetComponentText(frmConnect.txtPassword)

    UserPassword = aux

    If CheckUserData(False) = True Then
        EstadoLogin = Normal
        
        frmMain.Socket1.HostName = CurServerIP
        frmMain.Socket1.RemotePort = CurServerPort
        frmMain.Socket1.Connect

    End If
End Sub

Function CheckData() As Boolean
    If GetComponentText(txtPass) <> GetComponentText(txtRepPass) Then
        MsgBox "Los passwords que tipeo no coinciden, por favor vuelva a ingresarlos."
        Exit Function
    End If
    
    If Not CheckMailString(GetComponentText(txtMail)) Then
        MsgBox "Direccion de mail invalida."
        Exit Function
    End If

    If UserRaza = 0 Then
        MsgBox "Seleccione la raza del personaje."
        Exit Function
    End If
    
    If UserSexo = 0 Then
        MsgBox "Seleccione el sexo del personaje."
        Exit Function
    End If
    
    If UserHogar = 0 Then
        MsgBox "Seleccione el hogar del personaje."
        Exit Function
    End If
    
    Dim i As Integer
    For i = 1 To NUMATRIBUTOS
        If UserAtributos(i) = 0 Then
            MsgBox "Los atributos del personaje son invalidos."
            Exit Function
        End If
    Next i
    
    If Len(UserName) > 30 Then
        MsgBox ("El nombre debe tener menos de 30 letras.")
        Exit Function
    End If
    
    CheckData = True

End Function

