Attribute VB_Name = "Mod_Input"
Option Explicit
Public DI As DirectInput8
Public DIDevice As DirectInputDevice8
Public MouseEvent As Long

'CSEH: ErrLog
Public Sub Input_Init()
    '*****************************************************************
    'Init the input devices (keyboard and mouse)
    'More info: http://www.vbgore.com/GameClient.Input.Input_Init
    '*****************************************************************
    '<EhHeader>
    On Error GoTo Input_Init_Err
    '</EhHeader>
    Dim diProp As DIPROPLONG

        'Create the device
100     Set DI = DirectX.DirectInputCreate
105     Set DIDevice = DI.CreateDevice("guid_SysMouse")
    
110     If DIDevice Is Nothing Then GoTo Input_Init_Err

115     Call DIDevice.SetCommonDataFormat(DIFORMAT_MOUSE)
    
        'If in windowed mode, free the mouse from the screen
120     Call DIDevice.SetCooperativeLevel(frmConnect.hwnd, DISCL_NONEXCLUSIVE Or DISCL_BACKGROUND)
    
125     diProp.lHow = DIPH_DEVICE
130     diProp.lObj = 0
135     diProp.lData = 50
140     Call DIDevice.SetProperty("DIPROP_BUFFERSIZE", diProp)
145     MouseEvent = DirectX.CreateEvent(frmConnect)
    
150     DIDevice.SetEventNotification MouseEvent
    
            
155     Call DIDevice.Acquire
    
    '<EhFooter>
    Exit Sub

Input_Init_Err:
        Call LogError("Error en Input_Init: " & Erl & " - " & Err.Description)
    '</EhFooter>
End Sub

Public Sub Input_Release()
    
    Call DIDevice.Unacquire
    
    Set DIDevice = Nothing
    Set DI = Nothing
End Sub
