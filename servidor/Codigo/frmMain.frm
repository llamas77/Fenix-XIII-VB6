VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Argentum Online"
   ClientHeight    =   4845
   ClientLeft      =   1950
   ClientTop       =   1815
   ClientWidth     =   5190
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000004&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4845
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.TextBox txtChat 
      Height          =   2775
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   1920
      Width           =   4935
   End
   Begin VB.Timer tPiqueteC 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   120
      Top             =   600
   End
   Begin VB.Timer packetResend 
      Interval        =   10
      Left            =   120
      Top             =   120
   End
   Begin VB.CheckBox SUPERLOG 
      Caption         =   "log"
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton CMDDUMP 
      Caption         =   "dump"
      Height          =   255
      Left            =   3720
      TabIndex        =   8
      Top             =   480
      Width           =   1215
   End
   Begin VB.Timer FX 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   1440
      Top             =   540
   End
   Begin VB.Timer Auditoria 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1080
      Top             =   1140
   End
   Begin VB.Timer GameTimer 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   1440
      Top             =   60
   End
   Begin VB.Timer tLluviaEvent 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   600
      Top             =   1140
   End
   Begin VB.Timer tLluvia 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   600
      Top             =   600
   End
   Begin VB.Timer AutoSave 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   120
      Top             =   1080
   End
   Begin VB.Timer npcataca 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   1560
      Top             =   1140
   End
   Begin VB.Timer KillLog 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   1920
      Top             =   60
   End
   Begin VB.Timer TIMER_AI 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1935
      Top             =   540
   End
   Begin VB.Frame Frame1 
      Caption         =   "BroadCast"
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4935
      Begin VB.CommandButton Command2 
         Caption         =   "Broadcast consola"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   6
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Broadcast clientes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox BroadMsg 
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "Mensaje"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label Escuch 
      Caption         =   "Label2"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label CantUsuarios 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de usuarios:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1725
   End
   Begin VB.Label txStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   5520
      Width           =   45
   End
   Begin VB.Menu mnuControles 
      Caption         =   "Argentum"
      Begin VB.Menu mnuServidor 
         Caption         =   "Configuracion"
      End
      Begin VB.Menu mnuSystray 
         Caption         =   "Systray Servidor"
      End
      Begin VB.Menu mnuCerrar 
         Caption         =   "Cerrar Servidor"
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuMostrar 
         Caption         =   "&Mostrar"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
   
Const NIM_ADD = 0
Const NIM_DELETE = 2
Const NIF_MESSAGE = 1
Const NIF_ICON = 2
Const NIF_TIP = 4

Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDBLCLK = &H203
Const WM_RBUTTONUP = &H205

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

Private Function setNOTIFYICONDATA(hWnd As Long, ID As Long, flags As Long, CallbackMessage As Long, Icon As Long, Tip As String) As NOTIFYICONDATA
    Dim nidTemp As NOTIFYICONDATA

    nidTemp.cbSize = Len(nidTemp)
    nidTemp.hWnd = hWnd
    nidTemp.uID = ID
    nidTemp.uFlags = flags
    nidTemp.uCallbackMessage = CallbackMessage
    nidTemp.hIcon = Icon
    nidTemp.szTip = Tip & Chr$(0)

    setNOTIFYICONDATA = nidTemp
End Function

Sub CheckIdleUser()
    Dim iUserIndex As Long
    
    For iUserIndex = 1 To MaxUsers
        With UserList(iUserIndex)
            'Conexion activa? y es un usuario loggeado?
            If .ConnID <> -1 And .flags.UserLogged Then
                'Actualiza el contador de inactividad
                If .flags.Traveling = 0 Then
                    .Counters.IdleCount = .Counters.IdleCount + 1
                End If
                
                If .Counters.IdleCount >= IdleLimit Then
                    Call WriteShowMessageBox(iUserIndex, "Demasiado tiempo inactivo. Has sido desconectado.")
                    'mato los comercios seguros
                    If .ComUsu.DestUsu > 0 Then
                        If UserList(.ComUsu.DestUsu).flags.UserLogged Then
                            If UserList(.ComUsu.DestUsu).ComUsu.DestUsu = iUserIndex Then
                                Call WriteConsoleMsg(.ComUsu.DestUsu, "Comercio cancelado por el otro usuario.", FontTypeNames.FONTTYPE_TALK)
                                Call FinComerciarUsu(.ComUsu.DestUsu)
                                Call FlushBuffer(.ComUsu.DestUsu) 'flush the buffer to send the message right away
                            End If
                        End If
                        Call FinComerciarUsu(iUserIndex)
                    End If
                    Call Cerrar_Usuario(iUserIndex)
                End If
            End If
        End With
    Next iUserIndex
End Sub

Private Sub Auditoria_Timer()
On Error GoTo errhand
Static centinelSecs As Byte

centinelSecs = centinelSecs + 1

If centinelSecs = 5 Then
    'Every 5 seconds, we try to call the player's attention so it will report the code.
    Call modCentinela.CallUserAttention
    
    centinelSecs = 0
End If

Call PasarSegundo 'sistema de desconexion de 10 segs

Exit Sub

errhand:

Call LogError("Error en Timer Auditoria. Err: " & Err.description & " - " & Err.Number)
Resume Next

End Sub

Private Sub AutoSave_Timer()

On Error GoTo Errhandler
'fired every minute
Static Minutos As Long
Static MinutosLatsClean As Long

Minutos = Minutos + 1

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
Call ModAreas.AreasOptimizacion
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

'Actualizamos el centinela
Call modCentinela.PasarMinutoCentinela

If Minutos = MinutosWs - 1 Then
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Worldsave en 1 minuto ...", FontTypeNames.FONTTYPE_VENENO))
End If

If Minutos >= MinutosWs Then
    Call ES.DoBackUp
    Call aClon.VaciarColeccion
    Minutos = 0
End If

If MinutosLatsClean >= 15 Then
    MinutosLatsClean = 0
    Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales
    Call LimpiarMundo
Else
    MinutosLatsClean = MinutosLatsClean + 1
End If

Call PurgarPenas
Call CheckIdleUser

'<<<<<-------- Log the number of users online ------>>>
Dim N As Integer
N = FreeFile()
Open App.path & "\logs\numusers.log" For Output Shared As N
Print #N, NumUsers
Close #N
'<<<<<-------- Log the number of users online ------>>>

Exit Sub
Errhandler:
    Call LogError("Error en TimerAutoSave " & Err.Number & ": " & Err.description)
    Resume Next
End Sub

Private Sub CMDDUMP_Click()
On Error Resume Next

Dim i As Integer
For i = 1 To MaxUsers
    Call LogCriticEvent(i & ") ConnID: " & UserList(i).ConnID & ". ConnidValida: " & UserList(i).ConnIDValida & " Name: " & UserList(i).Name & " UserLogged: " & UserList(i).flags.UserLogged)
Next i

Call LogCriticEvent("Lastuser: " & LastUser & " NextOpenUser: " & NextOpenUser)

End Sub

Private Sub Command1_Click()
Call SendData(SendTarget.ToAll, 0, PrepareMessageShowMessageBox(BroadMsg.Text))
''''''''''''''''SOLO PARA EL TESTEO'''''''
''''''''''SE USA PARA COMUNICARSE CON EL SERVER'''''''''''
txtChat.Text = txtChat.Text & vbNewLine & "Servidor> " & BroadMsg.Text
End Sub

Public Sub InitMain(ByVal f As Byte)

If f = 1 Then
    Call mnuSystray_Click
Else
    frmMain.Show
End If

End Sub

Private Sub Command2_Click()
Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> " & BroadMsg.Text, FontTypeNames.FONTTYPE_SERVER))
''''''''''''''''SOLO PARA EL TESTEO'''''''
''''''''''SE USA PARA COMUNICARSE CON EL SERVER'''''''''''
txtChat.Text = txtChat.Text & vbNewLine & "Servidor> " & BroadMsg.Text
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
   
   If Not Visible Then
        Select Case X \ Screen.TwipsPerPixelX
                
            Case WM_LBUTTONDBLCLK
                WindowState = vbNormal
                Visible = True
                Dim hProcess As Long
                GetWindowThreadProcessId hWnd, hProcess
                AppActivate hProcess
            Case WM_RBUTTONUP
                hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf AppHook, App.hInstance, App.ThreadID)
                PopupMenu mnuPopUp
                If hHook Then UnhookWindowsHookEx hHook: hHook = 0
        End Select
   End If
   
End Sub

Private Sub QuitarIconoSystray()
On Error Resume Next

'Borramos el icono del systray
Dim i As Integer
Dim nid As NOTIFYICONDATA

nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, vbNull, frmMain.Icon, "")

i = Shell_NotifyIconA(NIM_DELETE, nid)
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

Call QuitarIconoSystray

Call LimpiaWsApi

Dim LoopC As Integer

For LoopC = 1 To MaxUsers
    If UserList(LoopC).ConnID <> -1 Then Call CloseSocket(LoopC)
Next

'Log
Dim N As Integer
N = FreeFile
Open App.path & "\logs\Main.log" For Append Shared As #N
Print #N, Date & " " & time & " server cerrado."
Close #N


Set SonidosMapas = Nothing

Call DumpGuilds(True)

End

End Sub

'CSEH: ErrLog
Private Sub FX_Timer()
    '<EhHeader>
    On Error GoTo FX_Timer_Err
    '</EhHeader>
100 Call SonidosMapas.ReproducirSonidosDeMapas

    '<EhFooter>
    Exit Sub

FX_Timer_Err:
        Call LogError("Error en FX_Timer: " & Erl & " - " & Err.description)
    '</EhFooter>
End Sub

'CSEH: ErrLog
Private Sub GameTimer_Timer()
    '********************************************************
    'Author: Unknown
    'Last Modify Date: -
    '********************************************************
    '<EhHeader>
    On Error GoTo GameTimer_Timer_Err
    '</EhHeader>
        Dim iUserIndex As Long
        Dim bEnviarStats As Boolean
        Dim bEnviarAyS As Boolean
        
        '<<<<<< Procesa eventos de los usuarios >>>>>>
100     For iUserIndex = 1 To MaxUsers 'LastUser
105         With UserList(iUserIndex)
               'Conexion activa?
110            If .ConnID <> -1 Then
                    '¿User valido?
                
115                 If .ConnIDValida And .flags.UserLogged Then
                    
                        '[Alejo-18-5]
120                     bEnviarStats = False
125                     bEnviarAyS = False
                    
                    
130                     If .flags.Paralizado = 1 Then Call EfectoParalisisUser(iUserIndex)
135                     If .flags.Ceguera = 1 Or .flags.Estupidez Then Call EfectoCegueEstu(iUserIndex)
                    
                    
140                     If .flags.Muerto = 0 Then
                        
145                         If .flags.Desnudo <> 0 And (.flags.Privilegios And PlayerType.User) <> 0 Then Call EfectoFrio(iUserIndex)
                        
150                         If .flags.Meditando Then Call DoMeditar(iUserIndex)
                        
155                         If .flags.Envenenado <> 0 And (.flags.Privilegios And PlayerType.User) <> 0 Then Call EfectoVeneno(iUserIndex)
                        
160                         If .flags.AdminInvisible <> 1 Then
165                             If .flags.invisible = 1 Then Call EfectoInvisibilidad(iUserIndex)
170                             If .flags.Oculto = 1 Then Call DoPermanecerOculto(iUserIndex)
                            End If
                        
175                         If .flags.Mimetizado = 1 Then Call EfectoMimetismo(iUserIndex)
                        
                        
185                         If .flags.BonusFlecha Then Call EfectoBonusFlecha(iUserIndex)
                        
190                         Call DuracionPociones(iUserIndex)
                        
195                         Call HambreYSed(iUserIndex, bEnviarAyS)
                        
200                         If .flags.Hambre = 0 And .flags.Sed = 0 Then
205                             If Lloviendo Then
210                                 If Not Intemperie(iUserIndex) Then
215                                     If Not .flags.Descansar Then
                                        'No esta descansando
220                                         Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)
225                                         If bEnviarStats Then
230                                             Call WriteUpdateHP(iUserIndex)
235                                             bEnviarStats = False
                                            End If
240                                         Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar)
245                                         If bEnviarStats Then
250                                             Call WriteUpdateSta(iUserIndex)
255                                             bEnviarStats = False
                                            End If
                                        Else
                                        'esta descansando
260                                         Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloDescansar)
265                                         If bEnviarStats Then
270                                             Call WriteUpdateHP(iUserIndex)
275                                             bEnviarStats = False
                                            End If
280                                         Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)
285                                         If bEnviarStats Then
290                                             Call WriteUpdateSta(iUserIndex)
295                                             bEnviarStats = False
                                            End If
                                            'termina de descansar automaticamente
300                                         If .Stats.MaxHp = .Stats.MinHp And .Stats.MaxSta = .Stats.MinSta Then
305                                             Call WriteRestOK(iUserIndex)
310                                             Call WriteConsoleMsg(iUserIndex, "Has terminado de descansar.", FontTypeNames.FONTTYPE_INFO)
315                                             .flags.Descansar = False
                                            End If
                                        
                                        End If
                                    End If
                                Else
320                                 If Not .flags.Descansar Then
                                    'No esta descansando
                                    
325                                     Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloSinDescansar)
330                                     If bEnviarStats Then
335                                         Call WriteUpdateHP(iUserIndex)
340                                         bEnviarStats = False
                                        End If
345                                     Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloSinDescansar)
350                                     If bEnviarStats Then
355                                         Call WriteUpdateSta(iUserIndex)
360                                         bEnviarStats = False
                                        End If
                                    
                                    Else
                                    'esta descansando
                                    
365                                     Call Sanar(iUserIndex, bEnviarStats, SanaIntervaloDescansar)
370                                     If bEnviarStats Then
375                                         Call WriteUpdateHP(iUserIndex)
380                                         bEnviarStats = False
                                        End If
385                                     Call RecStamina(iUserIndex, bEnviarStats, StaminaIntervaloDescansar)
390                                     If bEnviarStats Then
395                                         Call WriteUpdateSta(iUserIndex)
400                                         bEnviarStats = False
                                        End If
                                        'termina de descansar automaticamente
405                                     If .Stats.MaxHp = .Stats.MinHp And .Stats.MaxSta = .Stats.MinSta Then
410                                         Call WriteRestOK(iUserIndex)
415                                         Call WriteConsoleMsg(iUserIndex, "Has terminado de descansar.", FontTypeNames.FONTTYPE_INFO)
420                                         .flags.Descansar = False
                                        End If
                                    
                                    End If
                                End If
                            End If
                        
425                         If bEnviarAyS Then Call WriteUpdateHungerAndThirst(iUserIndex)
                        
430                         If .NroMascotas > 0 Then Call TiempoInvocacion(iUserIndex)
                        End If 'Muerto
                    Else 'no esta logeado?
                        'Inactive players will be removed!
435                     .Counters.IdleCount = .Counters.IdleCount + 1
440                     If .Counters.IdleCount > IntervaloParaConexion Then
445                         .Counters.IdleCount = 0
450                         Call CloseSocket(iUserIndex)
                        End If
                    End If 'UserLogged
                
                    'If there is anything to be sent, we send it
455                 Call FlushBuffer(iUserIndex)
                End If
            End With
460     Next iUserIndex
    '<EhFooter>
    Exit Sub

GameTimer_Timer_Err:
        Call LogError("Error en GameTimer_Timer: " & Erl & " - " & Err.description)
    '</EhFooter>
End Sub

Private Sub mnuCerrar_Click()


If MsgBox("¡¡Atencion!! Si cierra el servidor puede provocar la perdida de datos. ¿Desea hacerlo de todas maneras?", vbYesNo) = vbYes Then
    Dim f
    For Each f In Forms
        Unload f
    Next
End If

End Sub

Private Sub mnusalir_Click()
    Call mnuCerrar_Click
End Sub

Public Sub mnuMostrar_Click()
On Error Resume Next
    WindowState = vbNormal
    Form_MouseMove 0, 0, 7725, 0
End Sub

Private Sub KillLog_Timer()
On Error Resume Next
If FileExist(App.path & "\logs\connect.log", vbNormal) Then Kill App.path & "\logs\connect.log"
If FileExist(App.path & "\logs\haciendo.log", vbNormal) Then Kill App.path & "\logs\haciendo.log"
If FileExist(App.path & "\logs\stats.log", vbNormal) Then Kill App.path & "\logs\stats.log"
If FileExist(App.path & "\logs\Asesinatos.log", vbNormal) Then Kill App.path & "\logs\Asesinatos.log"
If FileExist(App.path & "\logs\HackAttemps.log", vbNormal) Then Kill App.path & "\logs\HackAttemps.log"
If Not FileExist(App.path & "\logs\nokillwsapi.txt") Then
    If FileExist(App.path & "\logs\wsapi.log", vbNormal) Then Kill App.path & "\logs\wsapi.log"
End If

End Sub

Private Sub mnuServidor_Click()
frmServidor.Visible = True
End Sub

Private Sub mnuSystray_Click()

Dim i As Integer
Dim S As String
Dim nid As NOTIFYICONDATA

S = "ARGENTUM-ONLINE"
nid = setNOTIFYICONDATA(frmMain.hWnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, frmMain.Icon, S)
i = Shell_NotifyIconA(NIM_ADD, nid)
    
If WindowState <> vbMinimized Then WindowState = vbMinimized
Visible = False

End Sub

'CSEH: ErrLog
Private Sub npcataca_Timer()
    '<EhHeader>
    On Error GoTo npcataca_Timer_Err
    '</EhHeader>

    Dim npc As Long

100 For npc = 1 To LastNPC
105     Npclist(npc).CanAttack = 1
110 Next npc

    '<EhFooter>
    Exit Sub

npcataca_Timer_Err:
        Call LogError("Error en npcataca_Timer: " & Erl & " - " & Err.description)
    '</EhFooter>
End Sub

Private Sub packetResend_Timer()
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/01/07
'Attempts to resend to the user all data that may be enqueued.
'***************************************************
On Error GoTo Errhandler:
    Dim i As Long
    
    For i = 1 To MaxUsers
        If UserList(i).ConnIDValida Then
            If UserList(i).outgoingData.Length > 0 Then
                Call EnviarDatosASlot(i, UserList(i).outgoingData.ReadASCIIStringFixed(UserList(i).outgoingData.Length))
            End If
        End If
    Next i

Exit Sub

Errhandler:
    LogError ("Error en packetResend - Error: " & Err.Number & " - Desc: " & Err.description)
    Resume Next
End Sub

Private Sub TIMER_AI_Timer()

On Error GoTo ErrorHandler
Dim NpcIndex As Long
Dim mapa As Integer
Dim e_p As Integer

'Barrin 29/9/03
If Not haciendoBK And Not EnPausa Then
    'Update NPCs
    For NpcIndex = 1 To LastNPC
        
        With Npclist(NpcIndex)
            If .flags.NPCActive Then 'Nos aseguramos que sea INTELIGENTE!
            
                If .flags.Paralizado = 1 Then
                    Call EfectoParalisisNpc(NpcIndex)
                Else
                    e_p = esPretoriano(NpcIndex)
                    If e_p > 0 Then
                        Select Case e_p
                            Case 1  ''clerigo
                                Call PRCLER_AI(NpcIndex)
                            Case 2  ''mago
                                Call PRMAGO_AI(NpcIndex)
                            Case 3  ''cazador
                                Call PRCAZA_AI(NpcIndex)
                            Case 4  ''rey
                                Call PRREY_AI(NpcIndex)
                            Case 5  ''guerre
                                Call PRGUER_AI(NpcIndex)
                        End Select
                    Else
                        'Usamos AI si hay algun user en el mapa
                        If .flags.Inmovilizado = 1 Then
                           Call EfectoParalisisNpc(NpcIndex)
                        End If
                        
                        mapa = .Pos.map
                        
                        If mapa > 0 Then
                            If MapInfo(mapa).NumUsers > 0 Then
                                If .Movement <> TipoAI.ESTATICO Then
                                    Call NPCAI(NpcIndex)
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End With
    Next NpcIndex
End If

Exit Sub

ErrorHandler:
    Call LogError("Error en TIMER_AI_Timer " & Npclist(NpcIndex).Name & " mapa:" & Npclist(NpcIndex).Pos.map)
    Call MuereNpc(NpcIndex, 0)
End Sub

Private Sub tLluvia_Timer()
On Error GoTo Errhandler

Dim iCount As Long
If Lloviendo Then
   For iCount = 1 To LastUser
        Call EfectoLluvia(iCount)
   Next iCount
End If

Exit Sub
Errhandler:
Call LogError("tLluvia " & Err.Number & ": " & Err.description)
End Sub

Private Sub tLluviaEvent_Timer()

On Error GoTo ErrorHandler
Static MinutosLloviendo As Long
Static MinutosSinLluvia As Long

If Not Lloviendo Then
    MinutosSinLluvia = MinutosSinLluvia + 1
    If MinutosSinLluvia >= 15 And MinutosSinLluvia < 1440 Then
        If RandomNumber(1, 100) <= 2 Then
            Lloviendo = True
            MinutosSinLluvia = 0
            Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
        End If
    ElseIf MinutosSinLluvia >= 1440 Then
        Lloviendo = True
        MinutosSinLluvia = 0
        Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
    End If
Else
    MinutosLloviendo = MinutosLloviendo + 1
    If MinutosLloviendo >= 5 Then
        Lloviendo = False
        Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
        MinutosLloviendo = 0
    Else
        If RandomNumber(1, 100) <= 2 Then
            Lloviendo = False
            MinutosLloviendo = 0
            Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
        End If
    End If
End If

Exit Sub
ErrorHandler:
Call LogError("Error tLluviaTimer")

End Sub

Private Sub tPiqueteC_Timer()
   ' Dim NuevoL As Boolean
    
    Dim i As Long
    
On Error GoTo Errhandler
    For i = 1 To LastUser
        With UserList(i)
            If .flags.UserLogged Then
                If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = eTrigger.ANTIPIQUETE Then
                    .Counters.PiqueteC = .Counters.PiqueteC + 1
                    Call WriteConsoleMsg(i, "¡¡¡Estás obstruyendo la vía pública, muévete o serás encarcelado!!!", FontTypeNames.FONTTYPE_INFO)
                    
                    If .Counters.PiqueteC > 23 Then
                        .Counters.PiqueteC = 0
                        Call Encarcelar(i, TIEMPO_CARCEL_PIQUETE)
                    End If
                Else
                    .Counters.PiqueteC = 0
                End If
                
                If .flags.Muerto = 1 Then
                    If .flags.Traveling = 1 Then
                        If .Counters.goHome <= 0 Then
                            Call FindLegalPos(i, Ciudades(.Hogar).map, Ciudades(.Hogar).X, Ciudades(.Hogar).Y)
                            Call WarpUserChar(i, Ciudades(.Hogar).map, Ciudades(.Hogar).X, Ciudades(.Hogar).Y, True)
                            Call WriteMultiMessage(i, eMessages.FinishHome)
                            .flags.Traveling = 0
                        Else
                            .Counters.goHome = .Counters.goHome - 1
                        End If
                    End If
                End If
                
                'ustedes se preguntaran que hace esto aca?
                'bueno la respuesta es simple: el codigo de AO es una mierda y encontrar
                'todos los puntos en los cuales la alineacion puede cambiar es un dolor de
                'huevos, asi que lo controlo aca, cada 6 segundos, lo cual es razonable
        
                
                
                Call FlushBuffer(i)
            End If
        End With
    Next i
Exit Sub

Errhandler:
    Call LogError("Error en tPiqueteC_Timer " & Err.Number & ": " & Err.description)
End Sub
