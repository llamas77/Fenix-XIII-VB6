Attribute VB_Name = "modMouseWheel"
Option Explicit

' Declaraciones api
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
    ByVal lpPrevWndFunc As Long, _
    ByVal hWnd As Long, _
    ByVal msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
    ByVal hWnd1 As Long, _
    ByVal hWnd2 As Long, _
    ByVal lpsz1 As String, _
    ByVal lpsz2 As String) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hWnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long


' Constantes para los mensajes del mouse
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  
Private Const WM_MOUSEWHEEL = &H20A

Private Const WHEEL_DELTA = 120
Private Const WHEEL_PAGESCROLL = &HFFFFFFFF

Private lpPrevWndProc As Long
  
Public Sub Hook_Main(hWnd As Long)
    Dim HWND_EDIT As Long
      
    HWND_EDIT = FindWindowEx(hWnd, 0, "EDIT", vbNullString)
      
    ' Inicia el hook
    lpPrevWndProc = SetWindowLong(HWND_EDIT, -4, AddressOf WndProc)
End Sub
  
' Finaliza el hook
Public Sub UnHook_Main(hWnd As Long)
      
    Dim HWND_EDIT As Long
    HWND_EDIT = FindWindowEx(hWnd, 0, "EDIT", vbNullString)
    SetWindowLong HWND_EDIT, -4, lpPrevWndProc
  
End Sub

Public Function WndProc(ByVal hWnd As Long, _
                        ByVal uMsg As Long, _
                        ByVal wParam As Long, _
                        ByVal lParam As Long) As Long

    On Error GoTo errorHandler

    If uMsg = WM_MOUSEWHEEL Then
        'If the flexGrid is the active control then
        'If TypeOf frmCrearPersonaje.ActiveControl Is PictureBox Then
        ' ##### Scroll direction #####
        
        Dim ID As Integer
            
        ID = mod_Components.Collision(frmMain.FormMouseX, frmMain.FormMouseY)
            
        If (HiWord(wParam) / WHEEL_DELTA) < 0 Then

            'Scrolling down
            If ID <> -1 Then
                If Components(ID).Component = eComponentType.TextArea Then
                    Call mod_Components.Execute(ID, eComponentEvent.MouseScrollDown)
                End If
            End If
        Else

            'Scrolling up
            If ID <> -1 Then
                If Components(ID).Component = eComponentType.TextArea Then
                    Call mod_Components.Execute(ID, eComponentEvent.MouseScrollUp)
                End If
            End If
        End If
    
        'Pass the message to default window procedure and then onto the parent
        'DefWindowProc hWnd, uMsg, wParam, lParam
    Else
        'No messages handled, call original window procedure
        WndProc = CallWindowProc(lpPrevWndProc, hWnd, uMsg, wParam, lParam)
    End If

    Exit Function
errorHandler:
    Debug.Print Err.Number & " " & Err.Description

End Function



Public Function HiWord(dw As Long) As Integer

If dw And &H80000000 Then
    HiWord = (dw \ 65535) - 1
Else
    HiWord = dw \ 65535
End If

End Function
