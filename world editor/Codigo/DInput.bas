Attribute VB_Name = "DInput"
'**************************************************************
'
' THIS WORK, INCLUDING THE SOURCE CODE, DOCUMENTATION
' AND RELATED MEDIA AND DATA, IS PLACED INTO THE PUBLIC DOMAIN.
'
' THE ORIGINAL AUTHOR IS RYAN CLARK.
'
' THIS SOFTWARE IS PROVIDED AS-IS WITHOUT WARRANTY
' OF ANY KIND, NOT EVEN THE IMPLIED WARRANTY OF
' MERCHANTABILITY. THE AUTHOR OF THIS SOFTWARE,
' ASSUMES _NO_ RESPONSIBILITY FOR ANY CONSEQUENCE
' RESULTING FROM THE USE, MODIFICATION, OR
' REDISTRIBUTION OF THIS SOFTWARE.
'
'**************************************************************
'
' This file was downloaded from The Game Programming Wiki.
' Come and visit us at http://gpwiki.org
'
'**************************************************************
'
'IMPORTANT TO NOTE:
'
'When using resource files and the CreateSurfaceFromResource
'command in DirectX 7.0, you will not be able to run your program
'by simply pressing F5 or selecting the run menu item, you will
'get an error. dX7 will only recognize your resource file if you
'compile your program and RUN THE COMPILED VERSION.
'
'**************************************************************

'dX Variables
Dim dx As New DirectX7
Dim di As DirectInput
Dim diDEV As DirectInputDevice
Dim diState As DIKEYBOARDSTATE

'Loop counter
Dim i As Integer

'Public array showing which keys are active
Public aKeys(211) As Boolean

'Keycode constants
Global Const DIK_ESCAPE = 1
Global Const DIK_1 = 2
Global Const DIK_2 = 3
Global Const DIK_3 = 4
Global Const DIK_4 = 5
Global Const DIK_5 = 6
Global Const DIK_6 = 7
Global Const DIK_7 = 8
Global Const DIK_8 = 9
Global Const DIK_9 = 10
Global Const DIK_0 = 11
Global Const DIK_MINUS = 12
Global Const DIK_EQUALS = 13
Global Const DIK_BACKSPACE = 14
Global Const DIK_TAB = 15
Global Const DIK_Q = 16
Global Const DIK_W = 17
Global Const DIK_E = 18
Global Const DIK_R = 19
Global Const DIK_T = 20
Global Const DIK_Y = 21
Global Const DIK_U = 22
Global Const DIK_I = 23
Global Const DIK_O = 24
Global Const DIK_P = 25
Global Const DIK_LBRACKET = 26
Global Const DIK_RBRACKET = 27
Global Const DIK_RETURN = 28
Global Const DIK_LCONTROL = 29
Global Const DIK_A = 30
Global Const DIK_S = 31
Global Const DIK_D = 32
Global Const DIK_F = 33
Global Const DIK_G = 34
Global Const DIK_H = 35
Global Const DIK_J = 36
Global Const DIK_K = 37
Global Const DIK_L = 38
Global Const DIK_SEMICOLON = 39
Global Const DIK_APOSTROPHE = 40
Global Const DIK_GRAVE = 41
Global Const DIK_LSHIFT = 42
Global Const DIK_BACKSLASH = 43
Global Const DIK_Z = 44
Global Const DIK_X = 45
Global Const DIK_C = 46
Global Const DIK_V = 47
Global Const DIK_B = 48
Global Const DIK_N = 49
Global Const DIK_M = 50
Global Const DIK_COMMA = 51
Global Const DIK_PERIOD = 52
Global Const DIK_SLASH = 53
Global Const DIK_RSHIFT = 54
Global Const DIK_MULTIPLY = 55
Global Const DIK_LALT = 56
Global Const DIK_SPACE = 57
Global Const DIK_CAPSLOCK = 58
Global Const DIK_F1 = 59
Global Const DIK_F2 = 60
Global Const DIK_F3 = 61
Global Const DIK_F4 = 62
Global Const DIK_F5 = 63
Global Const DIK_F6 = 64
Global Const DIK_F7 = 65
Global Const DIK_F8 = 66
Global Const DIK_F9 = 67
Global Const DIK_F10 = 68
Global Const DIK_NUMLOCK = 69
Global Const DIK_SCROLL = 70
Global Const DIK_NUMPAD7 = 71
Global Const DIK_NUMPAD8 = 72
Global Const DIK_NUMPAD9 = 73
Global Const DIK_SUBTRACT = 74
Global Const DIK_NUMPAD4 = 75
Global Const DIK_NUMPAD5 = 76
Global Const DIK_NUMPAD6 = 77
Global Const DIK_ADD = 78
Global Const DIK_NUMPAD1 = 79
Global Const DIK_NUMPAD2 = 80
Global Const DIK_NUMPAD3 = 81
Global Const DIK_NUMPAD0 = 82
Global Const DIK_DECIMAL = 83
Global Const DIK_F11 = 87
Global Const DIK_F12 = 88
Global Const DIK_NUMPADENTER = 156
Global Const DIK_RCONTROL = 157
Global Const DIK_DIVIDE = 181
Global Const DIK_RALT = 184
Global Const DIK_HOME = 199
Global Const DIK_UP = 200
Global Const DIK_PAGEUP = 201
Global Const DIK_LEFT = 203
Global Const DIK_RIGHT = 205
Global Const DIK_END = 207
Global Const DIK_DOWN = 208
Global Const DIK_PAGEDOWN = 209
Global Const DIK_INSERT = 210
Global Const DIK_DELETE = 211

Public Sub Initialize()

    'Create the direct input object
    Set di = dx.DirectInputCreate()
        
    'Aquire the keyboard as the device
    Set diDEV = di.CreateDevice("GUID_SysKeyboard")
    
    'Get input nonexclusively, only when in foreground mode
    diDEV.SetCommonDataFormat DIFORMAT_KEYBOARD
    diDEV.SetCooperativeLevel frmMain.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
    diDEV.Acquire
    
End Sub

Public Sub CheckKeysDX()
    
    'Get the current state of the keyboard
    diDEV.GetDeviceStateKeyboard diState
    
    'Scan through all the keys to check which are depressed
    For i = 1 To 211
        If diState.Key(i) <> 0 Then
            aKeys(i) = True             'If the key is pressed, set the appropriate array index to true
        Else
            aKeys(i) = False            'If the key is not pressed, set the appropriate array index to false
        End If
    Next
    
End Sub

Public Sub Terminate()
    
    'Unaquire the keyboard when we quit
    diDEV.Unacquire
    
End Sub
