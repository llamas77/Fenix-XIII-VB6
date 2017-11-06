Attribute VB_Name = "mod_Components"
Option Explicit

Private Const MAX_COMBOLIST_LINES As Byte = 5
Private Const MAX_CONSOLE_LINES As Byte = 100
 
Public Enum eComponentEvent
        None = 0
        MouseMove = 1
        MouseDown = 2
        KeyUp = 3
        KeyPress = 4
        MouseScrollUp = 5
        MouseScrollDown = 6
        MouseUp = 7
        MouseDblClick = 8
        ChildClicked = 9
End Enum

Public Enum eComponentType
        Label = 0
        TextBox = 1
        Shape = 2
        TextArea = 3
        Rect = 4
        ListBox = 5
        ComboBox = 6
        FilleableListbox = 7
End Enum

Private Type TYPE_CONSOLE_LINE
        Text As String
        Color(3) As Long
End Type

Private Type tComponent 'todo: rehacer, es terrible, OOP where are u?
        X           As Integer
        Y           As Integer
        W           As Integer
        H           As Integer
        
        Component   As eComponentType
        
        Enable      As Boolean
        Visible     As Boolean
        IsFocusable As Boolean
        ShowOnFocus As Boolean 'Only showed when its focused
        Color(3)    As Long
        
        Text        As String
        TextBuffer  As String 'Buffer
        
        ForeColor(3) As Long
        
        EventsPtr   As Long
        HasEvents   As Boolean
        
        'TextArea
        Lines()     As TYPE_CONSOLE_LINE
        LastLine    As Byte
        
        'first and last line to render in console
        FirstRender As Byte
        LastRender  As Byte
        
        SelIndex    As Integer
        
        Expanded    As Boolean 'combobox
        ListID      As Integer 'combobox
        ChildOf     As Integer
        
        PasswChr    As Byte
        
        Values()    As Integer 'Fillable
        Limit       As Integer
End Type

Private CharHeight      As Integer
Private Focused         As Integer
Private LastComponent   As Integer

Public Components()     As tComponent
Public HeadSlider(0 To 4) As Integer

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpvDest As Long, ByVal lpvSource As Long, ByVal cbCopy As Long)
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long

Public Sub InitComponents()
    
    Focused = -1
    
    CharHeight = cfonts(1).CharHeight
    
    If CharHeight = 0 Then
        MsgBox "InitComponents debe colocarse despúes de inicializar los Textos.", vbCritical
        End
    End If
End Sub

Public Sub ClearComponents()

    Erase Components
    
    Focused = -1
    LastComponent = 0
    
End Sub

Public Function AddListBox(ByVal X As Integer, ByVal Y As Integer, _
                           ByVal W As Integer, ByVal H As Integer, _
                           ByRef BackgroundColor() As Long, Optional ByVal DoRedim As Boolean = True, _
                           Optional ByVal Visible As Boolean = True, Optional ByVal ChildOf As Integer = 0) As Integer
    
    If DoRedim Then
        LastComponent = LastComponent + 1
    
        ReDim Preserve Components(1 To LastComponent) As tComponent
    End If
    
    With Components(LastComponent)
    
        .X = X: .W = W
        .Y = Y: .H = H
        
        .Component = eComponentType.ListBox
        
        .Color(0) = BackgroundColor(0): .Color(1) = BackgroundColor(1)
        .Color(2) = BackgroundColor(2): .Color(3) = BackgroundColor(3)
        
        .IsFocusable = False
        
        .Visible = Visible
        
        .Enable = True
        
        .SelIndex = -1
        .ChildOf = ChildOf
    End With
    
    Call SetEvents(LastComponent, Callback(AddressOf ListBox_EventHandler))
    
    AddListBox = LastComponent
    
End Function

Public Function AddFillableListBox(ByVal X As Integer, ByVal Y As Integer, _
                           ByVal W As Integer, ByVal H As Integer, _
                           ByRef BackgroundColor() As Long, ByVal Limit As Integer) As Integer
    
    LastComponent = LastComponent + 1

    ReDim Preserve Components(1 To LastComponent) As tComponent

    With Components(LastComponent)
    
        .X = X: .W = W
        .Y = Y: .H = H
        
        .Component = eComponentType.FilleableListbox
        
        .Color(0) = BackgroundColor(0): .Color(1) = BackgroundColor(1)
        .Color(2) = BackgroundColor(2): .Color(3) = BackgroundColor(3)
        
        .IsFocusable = False
        
        .Visible = True
        
        .Enable = True
        
        .SelIndex = -1
        
        .Limit = Limit
    End With
    
    'Call SetEvents(LastComponent, Callback(AddressOf FillableListBox_EventHandler))
    
    AddFillableListBox = LastComponent
    
End Function


Public Function AddComboBox(ByVal X As Integer, ByVal Y As Integer, _
                            ByVal W As Integer, ByVal H As Integer, _
                            ByRef BackgroundColor() As Long) As Integer
                            
    LastComponent = LastComponent + 2
    
    ReDim Preserve Components(1 To LastComponent) As tComponent
    
    With Components(LastComponent - 1)
    
        .X = X: .W = W
        .Y = Y: .H = H
        
        .Component = eComponentType.ComboBox
        
        .Color(0) = BackgroundColor(0): .Color(1) = BackgroundColor(1)
        .Color(2) = BackgroundColor(2): .Color(3) = BackgroundColor(3)
            
        .Enable = True
        .Visible = True
        
        .ListID = AddListBox(X + W, Y, W, 0, BackgroundColor, False, False, LastComponent - 1)
    End With
    
    Call SetEvents(LastComponent - 1, Callback(AddressOf ComboBox_EventHandler))
    AddComboBox = (LastComponent - 1)
    
End Function

Public Function AddRect(ByVal X As Integer, ByVal Y As Integer, _
                        ByVal W As Integer, ByVal H As Integer) As Integer
                            
                
    LastComponent = LastComponent + 1
    ReDim Preserve Components(1 To LastComponent) As tComponent
    
    With Components(LastComponent)
    
        .X = X: .W = W
        .Y = Y: .H = H
        
        .Component = eComponentType.Rect
        
        .Enable = True
        .Visible = True
    End With
    
    AddRect = LastComponent
    
End Function

Public Function AddTextArea(ByVal X As Integer, ByVal Y As Integer, _
                            ByVal W As Integer, ByVal H As Integer, _
                            Color() As Long) As Integer
    
    LastComponent = LastComponent + 1
    
    ReDim Preserve Components(1 To LastComponent) As tComponent
        
    With Components(LastComponent)
    
        .X = X: .W = W
        .Y = Y: .H = H
        
        .Component = eComponentType.TextArea
        
        .Color(0) = Color(0): .Color(1) = Color(1)
        .Color(2) = Color(2): .Color(3) = Color(3)
        
        .Enable = True
        .Visible = True
    End With
    
    AddTextArea = LastComponent
    
End Function

Public Function AddLabel(Text As String, ByVal X As Integer, ByVal Y As Integer, Color() As Long) As Integer

    LastComponent = LastComponent + 1
    
    ReDim Preserve Components(1 To LastComponent) As tComponent
        
    With Components(LastComponent)
        
        .X = X
        .Y = Y
        .Component = eComponentType.Label
        
        .Color(0) = Color(0): .Color(1) = Color(1)
        .Color(2) = Color(2): .Color(3) = Color(3)
        
        .Text = Text
        
        .Visible = True
    End With
    
    AddLabel = LastComponent
    
End Function

Public Function AddShape(ByVal X As Integer, ByVal Y As Integer, _
                            ByVal W As Integer, ByVal H As Integer, _
                            ByRef Color() As Long) As Integer
    
    LastComponent = LastComponent + 1
    
    ReDim Preserve Components(1 To LastComponent) As tComponent
    
    
    With Components(LastComponent)
        
        .X = X
        .Y = Y
        .W = W
        .H = H
        .Component = eComponentType.Shape
        
        .Color(0) = Color(0): .Color(1) = Color(1)
        .Color(2) = Color(2): .Color(3) = Color(3)
        
        .Enable = True
        .Visible = True
    End With
    
    AddShape = LastComponent
    
End Function

Public Function AddTextBox(ByVal X As Integer, ByVal Y As Integer, _
                            ByVal W As Integer, ByVal H As Integer, _
                            ByRef Color() As Long, ByRef ForeColor() As Long, _
                            Optional ByVal ShowOnFocus As Boolean = False, Optional ByVal PasswChr As Boolean = False) As Integer
    
    LastComponent = LastComponent + 1
    
    ReDim Preserve Components(1 To LastComponent) As tComponent
    
    With Components(LastComponent)
        
        .X = X
        .Y = Y
        .W = W
        .H = H
        .Component = eComponentType.TextBox
        
        .Color(0) = Color(0): .Color(1) = Color(1)
        .Color(2) = Color(2): .Color(3) = Color(3)
        
        .ForeColor(0) = ForeColor(0): .ForeColor(1) = ForeColor(1)
        .ForeColor(2) = ForeColor(2): .ForeColor(3) = ForeColor(3)
        
        .IsFocusable = True
        .ShowOnFocus = ShowOnFocus
        
        .PasswChr = PasswChr
        
        .Enable = True
        .Visible = True
    End With
    
    Call SetEvents(LastComponent, Callback(AddressOf TextBox_EventHandler))
    
    AddTextBox = LastComponent
    
End Function

Public Sub TabComponent()
    
    Dim i As Long
    Dim startID As Long
    
    If LastComponent <> 0 Then
        
        Focused = FindNextFocusable
        
    End If
End Sub

Private Function FindNextFocusable() As Integer
    
    Dim i As Long
    Dim startID As Long

    If Focused <> -1 Then
        i = Focused
        startID = i
    End If
    
    i = i + 1
    startID = i
    
    Do While (Components(i).IsFocusable = False Or Components(i).Visible = False)
        
        If LastComponent = i Then
            i = 0
        End If
        
        i = i + 1
        
        If startID = i Then
            GoTo s
        End If
    Loop
s:
    If Components(i).IsFocusable Then
        FindNextFocusable = i
    Else
        FindNextFocusable = -1
    End If
End Function

'Listbox and combos
Public Sub InsertText(ByVal ID As Integer, Text As String, TextColor() As Long)
    
    If Not (Components(ID).Component = eComponentType.ComboBox Or _
            Components(ID).Component = eComponentType.ListBox Or _
            Components(ID).Component = eComponentType.FilleableListbox) Then Exit Sub
    
    Dim i As Integer
    
    If Components(ID).Component = eComponentType.ComboBox Then i = Components(ID).ListID Else i = ID
    
    With Components(i)
        
        .LastLine = .LastLine + 1
    
        If .LastLine - 1 = 0 Then
            ReDim .Lines(1 To .LastLine) As TYPE_CONSOLE_LINE 'reused :^)
        Else
            ReDim Preserve .Lines(1 To .LastLine) As TYPE_CONSOLE_LINE
        End If
        
        .Lines(.LastLine).Text = Text
        .Lines(.LastLine).Color(0) = TextColor(0)
        .Lines(.LastLine).Color(1) = TextColor(1)
        .Lines(.LastLine).Color(2) = TextColor(2)
        .Lines(.LastLine).Color(3) = TextColor(3)
        
        If .Component = eComponentType.FilleableListbox Then
            ReDim .Values(0 To .LastLine - 1) As Integer
        End If
        
        Dim LastDrawableLine As Integer
        
        If Components(ID).Component = eComponentType.ComboBox Then .H = .H + CharHeight + 1
        
        LastDrawableLine = Fix(.H / CharHeight)
        
        If .LastLine = 1 Then
            .FirstRender = 1
            .LastRender = 1
            'If Components(ID).Component = eComponentType.ComboBox Then Components(ID).Text = Text
        Else
            .LastRender = .LastLine
            
            If .LastLine >= LastDrawableLine Then
                .FirstRender = .LastLine - (LastDrawableLine - 1)
            Else
                .FirstRender = 1
            End If
        End If
        
    End With
End Sub

Public Sub AppendLine(ByVal ID As Integer, Text As String, TextColor() As Long)
    
    If Not Components(ID).Component = eComponentType.TextArea Then Exit Sub
    
    With Components(ID)
        
        If .LastLine >= MAX_CONSOLE_LINES Then
            .LastLine = 0
        End If
        
        .LastLine = .LastLine + 1
        
        If .LastLine - 1 = 0 Then
            ReDim .Lines(1 To .LastLine) As TYPE_CONSOLE_LINE
        Else
            ReDim Preserve .Lines(1 To .LastLine) As TYPE_CONSOLE_LINE
        End If
        
        .Lines(.LastLine).Text = Text
        .Lines(.LastLine).Color(0) = TextColor(0)
        .Lines(.LastLine).Color(1) = TextColor(1)
        .Lines(.LastLine).Color(2) = TextColor(2)
        .Lines(.LastLine).Color(3) = TextColor(3)
        
        If .LastLine = 1 Then
            .FirstRender = 1
            .LastRender = 1
        Else
            .LastRender = .LastLine
            
            If .LastLine >= 9 Then
                .FirstRender = .LastLine - 8
            Else
                .FirstRender = 1
            End If
            
        End If
    End With
End Sub

Public Sub AppendLineCC(ByVal ID As Integer, Text As String, _
                        Optional ByVal Red As Integer = 1, Optional ByVal Green As Integer = 1, Optional ByVal blue As Integer = 1, _
                        Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, _
                        Optional ByVal NewLine As Boolean = True)
                        
    Dim Color(3) As Long
    
    Color(0) = RGB(Red, Green, blue)
    Color(1) = Color(0)
    Color(2) = Color(0)
    Color(3) = Color(0)
    
    Call AppendLine(ID, Text, Color)
End Sub

Public Sub ClearTextArea(ByVal ID As Integer, Optional ByVal Forced As Boolean = False)

    If Not Components(ID).Component = eComponentType.TextArea Then Exit Sub
    
    With Components(ID)
        
        If (.LastLine >= MAX_CONSOLE_LINES Or Forced) Then
            .LastLine = 0
            .FirstRender = 0
            .LastRender = 0
            
            ReDim .Lines(1) As TYPE_CONSOLE_LINE
            
        End If
    End With
End Sub

Public Function GetComponentValues(ByVal ID As Integer) As Integer()
    If Components(ID).Component <> eComponentType.FilleableListbox Then Exit Function
    
    GetComponentValues = Components(ID).Values
End Function

Public Function GetComponentValue(ByVal ID As Integer, ByVal Index As Integer) As Integer

    If Components(ID).Component <> eComponentType.FilleableListbox Then Exit Function
    
    If Index <= UBound(Components(ID).Values) Then
        GetComponentValue = Components(ID).Values(Index - 1)
    End If
    
End Function


Public Function GetComboText(ByVal ID As Integer) As String
    GetComboText = Components(ID).Text
End Function

Public Function GetSelectedValue(ByVal ID As Integer) As String
    If Components(ID).SelIndex <> 0 Then _
        GetSelectedValue = Components(ID).Lines(Components(ID).SelIndex).Text
End Function

Public Function GetSelectedIndex(ByVal ID As Integer) As Integer
    GetSelectedIndex = Components(ID).SelIndex
End Function

Public Sub EditLabel(ByVal ID As Integer, Text As String, Color() As Long, Optional ByVal X As Integer = -1, Optional ByVal Y As Integer = -1)

    
    With Components(ID)
        
        If .Component <> eComponentType.Label Then Exit Sub
        
        If X <> -1 Then .X = X
        If Y <> -1 Then .Y = Y
        
        .Color(0) = Color(0): .Color(1) = Color(1)
        .Color(2) = Color(2): .Color(3) = Color(3)
        
        .Text = Text
    End With
    
    
End Sub

Public Sub EditShape(ByVal ID As Integer, Color() As Long, _
                        Optional ByVal X As Integer = -1, Optional ByVal Y As Integer = -1, _
                        Optional ByVal W As Integer = -1, Optional ByVal H As Integer = -1)
    
    With Components(ID)
        
        If .Component <> eComponentType.Shape Then Exit Sub
        
        If X <> -1 Then .X = X
        If Y <> -1 Then .Y = Y
        If W <> -1 Then .W = W
        If H <> -1 Then .H = H
        
        .Color(0) = Color(0): .Color(1) = Color(1)
        .Color(2) = Color(2): .Color(3) = Color(3)
    End With
    
End Sub

Public Sub EditTextBox(ByVal ID As Integer, Color() As Long, ForeColor() As Long, _
                        Optional ByVal X As Integer = -1, Optional ByVal Y As Integer = -1, _
                        Optional ByVal W As Integer = -1, Optional ByVal H As Integer = -1, _
                        Optional ByVal ShowOnFocus As Boolean = False)
    
    With Components(ID)
        
        If .Component <> eComponentType.TextBox Then Exit Sub
        
        If X <> -1 Then .X = X
        If Y <> -1 Then .Y = Y
        If W <> -1 Then .W = W
        If H <> -1 Then .H = H
        
        .Color(0) = Color(0): .Color(1) = Color(1)
        .Color(2) = Color(2): .Color(3) = Color(3)
        
        .ForeColor(0) = ForeColor(0): .ForeColor(1) = ForeColor(1)
        .ForeColor(2) = ForeColor(2): .ForeColor(3) = ForeColor(3)
        
        .IsFocusable = True
        .ShowOnFocus = ShowOnFocus
        
    End With
    
End Sub

Public Function SetComponentFocus(ByVal ID As Integer) As Integer
    
    If Focused <> ID Then
        If Components(ID).IsFocusable Then
            Focused = ID
            SetComponentFocus = ID
        End If
    Else
        SetComponentFocus = ID
    End If
    
End Function

Public Sub RenderComponents(ByVal Alpha As Byte)
    
    Dim i As Long
    Dim Component As tComponent
    
    For i = 1 To LastComponent
        Component = Components(i)
        
        With Component
            
            If .Visible = False Then GoTo NextLoop
            
            If Alpha <> 255 Then _
                InjectAlphaToColor Alpha, .Color
            
            Select Case .Component
            
                Case eComponentType.Label
                    Call Text_Draw(.X, .Y, .Text, .Color)
                
                Case eComponentType.Shape
                    Call Draw_Box(.X, .Y, .W, .H, .Color)
                    
                Case eComponentType.TextBox
                    If .ShowOnFocus Then
                        If Focused = i Then
                            Call Draw_Box(.X, .Y, .W, .H, .Color)
                            Call UpdateTextBoxBuffer(i)
                        End If
                    Else
                        Call Draw_Box(.X, .Y, .W, .H, .Color)
                        Call UpdateTextBoxBuffer(i)
                    End If
                
                Case eComponentType.TextArea
                    Call Draw_Box(.X, .Y, .W, .H, .Color)
                    Call UpdateTextArea(i)
                
                Case eComponentType.ComboBox
                    Call Draw_Box(.X, .Y, .W, .H, .Color)
                    Call Text_Draw(.X + 3, .Y + (.H \ 2) - (CharHeight \ 2) - 1, .Text, White)
                    'Call Draw_Box(.X + .W - 10, .Y, .H, .H, Gray)
                    
                    If .Expanded Then
                        'Call Text_Draw(.X + .W - 8, .Y - 1, "<", Black)
                        Call Device_Textured_Render(.X + .W - 27, .Y - 1, 27, 23, 0, 0, 1000003, White)
                    Else
                        'Call Text_Draw(.X + .W - 8, .Y - 1, ">", Black)
                        Call Device_Textured_Render(.X + .W - 27, .Y - 1, 27, 23, 27, 0, 1000003, White)
                    End If
                    
                Case eComponentType.ListBox
                    Call DrawListBox(i)
                
                Case eComponentType.FilleableListbox
                    Call DrawFilleableListBox(i)
            End Select
            
        End With
        
NextLoop:
    Next
    
End Sub

Private Sub DrawListBox(ByVal ID As Integer)
    
    With Components(ID)
    
        Dim i As Long
        Dim yOffset As Integer
        
        Call Draw_Box(.X, .Y, .W, .H, .Color)
        
        If .FirstRender = 0 Then Exit Sub
        
        For i = .FirstRender To .LastRender
            If i = .SelIndex Then
                Call Draw_Box(.X, .Y + 1 + yOffset, .W, CharHeight, Gray)
                Call Text_Draw(.X + 3, .Y + 1 + yOffset, .Lines(i).Text, .Lines(i).Color)
            Else
                Call Text_Draw(.X + 3, .Y + 1 + yOffset, .Lines(i).Text, .Lines(i).Color)
            End If
            yOffset = yOffset + CharHeight
        Next
        
    End With
End Sub

Private Sub DrawFilleableListBox(ByVal ID As Integer)
    
    With Components(ID)
    
        Dim i As Long
        Dim yOffset As Integer
        
        Call Draw_Box(.X, .Y, .W, .H, .Color)
        
        If .FirstRender = 0 Then Exit Sub
        
        For i = .FirstRender To .LastRender
            
            Call Draw_Box(.X, .Y + 1 + yOffset, (.Values(i - 1) * (.W / .Limit)), CharHeight, Gray)
            If .Values(i - 1) <> 0 Then
                Call Text_Draw(.X + 3, .Y + 1 + yOffset, .Lines(i).Text & " (" & .Values(i - 1) & ")", .Lines(i).Color)
            Else
                Call Text_Draw(.X + 3, .Y + 1 + yOffset, .Lines(i).Text, .Lines(i).Color)
            End If
            
            yOffset = yOffset + CharHeight
        Next
        
    End With
End Sub

Private Sub UpdateTextBoxBuffer(ByVal ID As Integer)
    
    'If UserWriting Then
        With Components(ID)
            
            If Not StrComp(.TextBuffer, vbNullString) = 0 Then
                
                Dim renderstr As String
                If .PasswChr Then
                    renderstr = String$(Len(.TextBuffer), "*")
                Else
                    renderstr = .TextBuffer
                End If
                
                If Focused = ID Then
                    Call Text_Draw(.X + 3, .Y + 3, renderstr + "|", .ForeColor)
                Else
                    Call Text_Draw(.X + 3, .Y + 3, renderstr, .ForeColor)
                End If
            Else
                If Focused = ID Then
                    Call Text_Draw(.X + 3, .Y + 3, "|", .ForeColor)
                End If
                
            End If
            
        End With
    'End If
    
End Sub

Private Sub ScrollListUp(ByVal ID As Integer)
    If Not (Components(ID).Component = eComponentType.ListBox Or _
            Components(ID).Component = eComponentType.FilleableListbox) Then Exit Sub
            
    If Components(ID).LastLine = 0 Then Exit Sub
    
    With Components(ID)
        
        If .FirstRender = 1 Then Exit Sub
        
        .FirstRender = .FirstRender - 1
        .LastRender = .LastRender - 1
        
    End With
    
End Sub

Private Sub ScrollListDown(ByVal ID As Integer)
    If Not (Components(ID).Component = eComponentType.ListBox Or _
            Components(ID).Component = eComponentType.FilleableListbox) Then Exit Sub
            
    If Components(ID).LastLine = 0 Then Exit Sub
    
    With Components(ID)
        
        Dim LastDrawableLine As Integer
        
        LastDrawableLine = Fix(.H / CharHeight)
        
        If .LastLine = (LastDrawableLine + .FirstRender) - 1 Then Exit Sub
        
        .FirstRender = .FirstRender + 1
        .LastRender = .LastRender + 1
        
    End With
    
End Sub

Private Sub ScrollConsoleUp(ByVal ID As Integer)
    If Components(ID).Component <> eComponentType.TextArea Then Exit Sub
    If Components(ID).LastLine = 0 Then Exit Sub
    
    With Components(ID)
        
        If .FirstRender = 1 Then Exit Sub
        
        .FirstRender = .FirstRender - 1
        .LastRender = .LastRender - 1
        
    End With
    
End Sub

Private Sub ScrollConsoleDown(ByVal ID As Integer)
    If Components(ID).Component <> eComponentType.TextArea Then Exit Sub
    If Components(ID).LastLine = 0 Then Exit Sub
    
    With Components(ID)
        
        If .FirstRender = .LastRender - 8 Then Exit Sub
        
        .FirstRender = .FirstRender + 1
        .LastRender = .LastRender + 1
        
    End With
    
End Sub

Private Sub UpdateTextArea(ByVal ID As Integer)
    
    With Components(ID)
    
        Dim i As Long
        Dim yOffset As Integer
            
        For i = .FirstRender To .LastRender
            Text_Draw .X + 3, .Y + 2 + yOffset, .Lines(i).Text, .Lines(i).Color
            yOffset = yOffset + 12
        Next
        
    End With
End Sub

Public Sub SetEvents(ByVal ID As Integer, Events As Long)

With Components(ID)

    .HasEvents = True
    
    .EventsPtr = Events
    
End With

End Sub

Public Function GetComponentText(ByVal ID As Integer) As String
        
    GetComponentText = Components(ID).TextBuffer
    
End Function

'@Rezniaq
Public Function Collision(ByVal X As Integer, ByVal Y As Integer) As Integer
 
Dim i                                   As Long
 
'buscamos un objeto que colisione
For i = 1 To LastComponent
    With Components(i)
        'comprobamos X e Y
        If X > .X And X < .X + .W Then
            If Y > .Y And Y < .Y + .H Then
                If .Visible And .Enable Then
                    Collision = i
                    Exit Function
                End If
            End If
        End If
    End With
Next i
 
'no hay colisión
Collision = -1
 
End Function
 
'@Rezniaq
Public Sub Execute(ByVal ID As Integer, ByVal eventIndex As eComponentEvent, Optional ByVal param3 As Long = 0, Optional ByVal param4 As Long = 0)
 
With Components(ID)
    'si el objeto tiene eventos
    If .Enable Then
        If .HasEvents = True Then
            'si el objeto tiene ESTE evento
            If .EventsPtr <> 0 Then
                'llamamos al sub (un parámetro obligatorio es
                'objectIndex, independientemente de que si el sub
                'lo necesita o no, debe poseerlo como parámetro)
                CallWindowProc .EventsPtr, ID, eventIndex, param3, param4
            End If
        End If
    End If
End With
 
End Sub

Public Sub SetFocus(ByVal ID As Integer)
    
    If ID = -1 Then
        Focused = ID
    Else
        If Components(ID).IsFocusable Then Focused = ID
    End If
End Sub

Public Function GetFocused() As Integer

    GetFocused = Focused
    
End Function

Public Function Callback(ByVal param As Long) As Long

        Callback = param
        
End Function

Public Sub HideComponents(ParamArray Comps() As Variant)
    Dim i As Long
    
    For i = 0 To UBound(Comps)
        Components(Comps(i)).Visible = False
    Next
    
End Sub

Public Sub ShowComponents(ParamArray Comps() As Variant)
    Dim i As Long
    
    For i = 0 To UBound(Comps)
        Components(Comps(i)).Visible = True
    Next
    
End Sub

Public Sub DisableComponents(ParamArray Comps() As Variant)
    Dim i As Long
    
    For i = 0 To UBound(Comps)
        Components(Comps(i)).Enable = False
    Next
End Sub

Public Sub EnableComponents(ParamArray Comps() As Variant)
    Dim i As Long
    
    For i = 0 To UBound(Comps)
        Components(Comps(i)).Enable = True
    Next
End Sub

Public Sub SetChild(ByVal FatherID As Integer, ByVal ChildID As Integer) 'no siblings yet
    Components(ChildID).ChildOf = FatherID
End Sub

Private Function Validate(ByVal State As eRenderState) As Boolean
                                    
    Dim i As Long
    Dim CharAscii As Byte
    
    With frmConnect
    
        Select Case State
        
            Case eRenderState.eNewCharInfo
                
                If StrComp(GetComponentText(.txtNick), vbNullString) = 0 Then
                    MsgBox "Elige un nombre para tu personaje."
                    Validate = False
                    Exit Function
                Else
                    UserName = Trim$(GetComponentText(.txtNick))
                End If
                
                If Len(UserName) > 30 Then
                    MsgBox ("El nombre debe tener menos de 30 letras.")
                    Validate = False
                    Exit Function
                End If
                
                If Not CheckMailString(GetComponentText(.txtMail)) Then
                    MsgBox "Escribe un correo electrónico válido."
                    Validate = False
                    Exit Function
                Else
                    UserEmail = GetComponentText(.txtMail)
                End If
                
                If (StrComp(GetComponentText(.txtPass), vbNullString) = 0) Or (StrComp(GetComponentText(.txtRepPass), vbNullString) = 0) Or _
                (StrComp(GetComponentText(.txtPass), GetComponentText(.txtRepPass)) <> 0) Then
                    MsgBox "Contraseña inválida."
                    Validate = False
                    Exit Function
                Else
                    UserPassword = GetComponentText(.txtPass)
                End If
                
                For i = 1 To Len(UserPassword)
                    CharAscii = Asc(mid$(UserPassword, i, 1))
                    If Not LegalCharacter(CharAscii) Then
                        MsgBox ("Password inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
                        Validate = False
                        Exit Function
                    End If
                Next i
                
            Case eRenderState.eNewCharDetails
                
                If GetSelectedIndex(.cmbHogar) = 0 Then
                    MsgBox "Confirma tu lugar de origen."
                    Validate = False
                    Exit Function
                Else
                    UserHogar = GetSelectedIndex(.cmbHogar)
                End If
                
                If GetSelectedIndex(.cmbRaza) = 0 Then
                    MsgBox "Confirma tu raza."
                    Validate = False
                    Exit Function
                End If
                
                If GetSelectedIndex(.cmbSexo) = 0 Then
                    MsgBox "Confirma tu sexo."
                    Validate = False
                    Exit Function
                End If
                
            Case eRenderState.eNewCharAttrib
            
                For i = 1 To NUMATRIBUTOS
                    If UserAtributos(i) = 0 Then
                        MsgBox "Los atributos del personaje son invalidos."
                        Validate = False
                        Exit Function
                    End If
                Next i
        End Select
    
    End With
    
    Validate = True
    
End Function

Private Sub DarCuerpoYCabeza()

    Select Case UserSexo
        Case eGenero.Hombre
            Select Case UserRaza
                Case eRaza.Humano
                    UserHead = HUMANO_H_PRIMER_CABEZA
                    UserBody = HUMANO_H_CUERPO_DESNUDO
                    
                Case eRaza.Elfo
                    UserHead = ELFO_H_PRIMER_CABEZA
                    UserBody = ELFO_H_CUERPO_DESNUDO
                    
                Case eRaza.ElfoOscuro
                    UserHead = DROW_H_PRIMER_CABEZA
                    UserBody = DROW_H_CUERPO_DESNUDO
                    
                Case eRaza.Enano
                    UserHead = ENANO_H_PRIMER_CABEZA
                    UserBody = ENANO_H_CUERPO_DESNUDO
                    
                Case eRaza.Gnomo
                    UserHead = GNOMO_H_PRIMER_CABEZA
                    UserBody = GNOMO_H_CUERPO_DESNUDO
                    
                Case Else
                    UserHead = 0
                    UserBody = 0
            End Select
            
        Case eGenero.Mujer
            Select Case UserRaza
                Case eRaza.Humano
                    UserHead = HUMANO_M_PRIMER_CABEZA
                    UserBody = HUMANO_M_CUERPO_DESNUDO
                    
                Case eRaza.Elfo
                    UserHead = ELFO_M_PRIMER_CABEZA
                    UserBody = ELFO_M_CUERPO_DESNUDO
                    
                Case eRaza.ElfoOscuro
                    UserHead = DROW_M_PRIMER_CABEZA
                    UserBody = DROW_M_CUERPO_DESNUDO
                    
                Case eRaza.Enano
                    UserHead = ENANO_M_PRIMER_CABEZA
                    UserBody = ENANO_M_CUERPO_DESNUDO
                    
                Case eRaza.Gnomo
                    UserHead = GNOMO_M_PRIMER_CABEZA
                    UserBody = GNOMO_M_CUERPO_DESNUDO
            End Select
    End Select
    
    If UserSexo <> 0 And UserRaza <> 0 Then
        Call SetBodyExample(UserBody)
    End If
     
End Sub
Private Sub UpdateHeadSelection()
    
    HeadSlider(0) = CheckCabeza(UserHead - 2)
    HeadSlider(1) = CheckCabeza(UserHead - 1)
    HeadSlider(2) = UserHead
    HeadSlider(3) = CheckCabeza(UserHead + 1)
    HeadSlider(4) = CheckCabeza(UserHead + 2)
    
End Sub

'*********************************************************************************************************************************
'*********************************************************************************************************************************
'*********************************************************************************************************************************
'*****************************************************EVENTS HANDLERS*************************************************************

Public Sub btnHeadDer_EventHandler(ByVal hwnd As Long, _
                                   ByVal msg As Long, _
                                   ByVal param3 As Long, _
                                   ByVal param4 As Long)

    If msg = eComponentEvent.MouseUp Then
        UserHead = CheckCabeza(UserHead - 1)
        Call UpdateHeadSelection
    End If
    
End Sub

Public Sub btnHeadIzq_EventHandler(ByVal hwnd As Long, _
                                   ByVal msg As Long, _
                                   ByVal param3 As Long, _
                                   ByVal param4 As Long)

    If msg = eComponentEvent.MouseUp Then
        UserHead = CheckCabeza(UserHead + 1)
        Call UpdateHeadSelection
    End If
    
End Sub

'This Override the primitive method TextBox_EventHandler
Public Sub txtRepPass_EventHandler(ByVal hwnd As Long, _
                                   ByVal msg As Long, _
                                   ByVal param3 As Long, _
                                   ByVal param4 As Long)
    Dim i As Long
    Dim tempstr As String
    Dim Buffer As String
    
    Buffer = Components(hwnd).TextBuffer
    
    With Components(hwnd)
    
        Select Case msg
            
            Case eComponentEvent.MouseUp
                Call SetFocus(hwnd)
                
            Case eComponentEvent.KeyPress
                If Not (param3 = vbKeyBack) And Not (param3 >= vbKeySpace And param3 <= 250) Then param3 = 0
                
    
                Buffer = Buffer + ChrW$(param3)
                
                'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
                For i = 1 To Len(Buffer)
                    param3 = Asc(mid$(Buffer, i, 1))
    
                    If param3 >= vbKeySpace And param3 <= 250 Then
                        tempstr = tempstr & ChrW$(param3)
                    End If
    
                    If param3 = vbKeyBack And Len(tempstr) > 0 Then
                        tempstr = Left$(tempstr, Len(tempstr) - 1)
                    End If
                Next i
    
                If tempstr <> Buffer Then
                    'We only set it if it's different, otherwise the event will be raised
                    'constantly and the client will crush
                    Buffer = tempstr
                End If
    
                Components(hwnd).TextBuffer = Buffer

                If StrComp(.TextBuffer, Components(.ChildOf).TextBuffer) = 0 Then
                
                    .ForeColor(0) = Green(0): .ForeColor(1) = Green(1)
                    .ForeColor(2) = Green(2): .ForeColor(3) = Green(3)
                    
                Else
                    .ForeColor(0) = Red(0): .ForeColor(1) = Red(1)
                    .ForeColor(2) = Red(2): .ForeColor(3) = Red(3)
                End If
        End Select
        
    End With
End Sub


Public Sub btnLogin_EventHandler(ByVal hwnd As Long, _
                                ByVal msg As Long, _
                                ByVal param3 As Long, _
                                ByVal param4 As Long)
    
    Select Case msg
    
        Case eComponentEvent.MouseUp
            frmConnect.LoginUser
    End Select
End Sub

Public Sub btnDados_EventHandler(ByVal hwnd As Long, _
                                ByVal msg As Long, _
                                ByVal param3 As Long, _
                                ByVal param4 As Long)

    If msg = eComponentEvent.MouseUp Then
        Call WriteThrowDices
        Call FlushBuffer
    End If
End Sub


Public Sub btnNewCharacter_EventHandler(ByVal hwnd As Long, _
                                ByVal msg As Long, _
                                ByVal param3 As Long, _
                                ByVal param4 As Long)
    
    If msg = eComponentEvent.MouseUp Then
        
        UserClase = eClass.Ciudadano
        UserSexo = 0
        UserRaza = 0
        UserHogar = 0
        UserEmail = ""
        UserHead = 0
    
        EstadoLogin = E_MODO.Dados
            
        If frmMain.Socket1.Connected Then
            frmMain.Socket1.Disconnect
            frmMain.Socket1.Cleanup
            DoEvents
        End If
                
        frmMain.Socket1.HostName = CurServerIP
        frmMain.Socket1.RemotePort = CurServerPort
        frmMain.Socket1.Connect
    End If
End Sub

Public Sub cmbRaza_EventHandler(ByVal hwnd As Long, _
                                 ByVal msg As Long, _
                                 ByVal param3 As Long, _
                                 ByVal param4 As Long)
    
    If msg <> 0 Then
        
        Select Case msg
            
            Case eComponentEvent.MouseDown
                Components(hwnd).Expanded = Not Components(hwnd).Expanded
                Components(Components(hwnd).ListID).Visible = Components(hwnd).Expanded
            
            Case eComponentEvent.ChildClicked
                If Components(Components(hwnd).ListID).SelIndex <> 0 Then
                    UserRaza = Components(Components(hwnd).ListID).SelIndex
                    Call DarCuerpoYCabeza
                    Call UpdateHeadSelection
                    
                    With HelpWindow
                    
                        ReDim .Text(0 To 5) As String
                        
                        .Text(0) = "Esta raza te otorga:"
                        .Text(1) = "Fuerza: " & ModRaza(UserRaza).Fuerza
                        .Text(2) = "Agilidad: " & ModRaza(UserRaza).Agilidad
                        .Text(3) = "Constitución: " & ModRaza(UserRaza).Constitucion
                        .Text(4) = "Inteligencia: " & ModRaza(UserRaza).Inteligencia
                        .Text(5) = "Carisma: " & ModRaza(UserRaza).Carisma
                    End With
                    
                    HelpWindow.Active = True
                End If
                
        End Select
    End If
End Sub

Public Sub cmbSexo_EventHandler(ByVal hwnd As Long, _
                                 ByVal msg As Long, _
                                 ByVal param3 As Long, _
                                 ByVal param4 As Long)
    
    If msg <> 0 Then
        
        Select Case msg
            
            Case eComponentEvent.MouseDown
                Components(hwnd).Expanded = Not Components(hwnd).Expanded
                Components(Components(hwnd).ListID).Visible = Components(hwnd).Expanded
            
            Case eComponentEvent.ChildClicked
                If Components(Components(hwnd).ListID).SelIndex <> 0 Then
                    UserSexo = Components(Components(hwnd).ListID).SelIndex
                    Call DarCuerpoYCabeza
                    Call UpdateHeadSelection
                End If

        End Select
    End If
End Sub

Private Sub ComboBox_EventHandler(ByVal hwnd As Long, _
                                 ByVal msg As Long, _
                                 ByVal param3 As Long, _
                                 ByVal param4 As Long)
    
    If msg <> 0 Then
        
        Select Case msg
            
            Case eComponentEvent.MouseDown
                Components(hwnd).Expanded = Not Components(hwnd).Expanded
                Components(Components(hwnd).ListID).Visible = Components(hwnd).Expanded
            
            Case eComponentEvent.ChildClicked
                If Components(Components(hwnd).ListID).SelIndex <> 0 Then
                    Components(hwnd).SelIndex = Components(Components(hwnd).ListID).SelIndex
                End If
            'Case eComponentEvent.MouseScrollUp: If Components(hwnd).Expanded Then Call ScrollListUp(hwnd)
            'Case eComponentEvent.MouseScrollDown: If Components(hwnd).Expanded Then Call ScrollListDown(hwnd)
            
            
        End Select
    End If
End Sub

Public Sub btnSiguiente_EventHandler(ByVal hwnd As Long, _
                                     ByVal msg As Long, _
                                     ByVal param3 As Long, _
                                     ByVal param4 As Long)
    
    Select Case msg
        
        Case eComponentEvent.MouseUp
            Dim State As eRenderState
            State = GetRenderState()
            
            If State = eRenderState.eLogin Then Exit Sub
            
            With frmConnect
            
                If State = eRenderState.eNewCharSkills Then
                    Call frmConnect.LoginNewChar
                Else
                
                    If Validate(State) Then
                        HelpWindow.Active = False
                        Call ChangeRenderState(State + 1)
                    End If
                End If
            
            End With
            
    End Select
    
End Sub

Public Sub btnAtras_EventHandler(ByVal hwnd As Long, _
                                     ByVal msg As Long, _
                                     ByVal param3 As Long, _
                                     ByVal param4 As Long)
    
    Select Case msg
        
        Case eComponentEvent.MouseUp
            
            If GetRenderState() = eRenderState.eLogin Then Exit Sub
            
            If GetRenderState = eRenderState.eNewCharInfo Then
                Call frmConnect.CloseNewChar
            Else
                Call ChangeRenderState(GetRenderState() - 1)
                HelpWindow.Active = False
            End If
    End Select
    
End Sub

Public Sub lstSkill_EventHandler(ByVal hwnd As Long, _
                                 ByVal msg As Long, _
                                 ByVal param3 As Long, _
                                 ByVal param4 As Long)

    Dim X As Integer, Y As Integer
    Dim Button As Integer, Shift As Integer
    
    Dim LastDrawableLine As Integer
        
        
    Call LongToIntegers(param3, X, Y)
    Call LongToIntegers(param4, Button, Shift)
    
    Dim skl As Integer
    skl = frmConnect.SkillPts
    
    'todo: barra scroll
    If msg <> 0 Then
        
        Y = Y - (Components(hwnd).Y - CharHeight / 2)
        Y = (Y + 2) \ CharHeight
        
        With Components(hwnd)
        
            Select Case msg
                
                Case eComponentEvent.MouseDown
                    LastDrawableLine = Fix(.H / CharHeight)
                    
                    If Y > 0 And Y <= LastDrawableLine Then
                        .SelIndex = (.FirstRender - 1) + Y
                        
                        If Button = vbLeftButton Then
                            
                            If skl = 0 Then Exit Sub
                            
                            If Shift And 2 Then
                                If skl >= 3 Then
                                    .Values(.SelIndex - 1) = .Values(.SelIndex - 1) + 3
                                    skl = skl - 3
                                ElseIf skl >= 2 Then
                                    .Values(.SelIndex - 1) = .Values(.SelIndex - 1) + 2
                                    skl = skl - 2
                                Else
                                    .Values(.SelIndex - 1) = .Values(.SelIndex - 1) + 1
                                    skl = skl - 1
                                End If
                            
                            ElseIf Shift And 3 Then
                                .Values(.SelIndex - 1) = .Values(.SelIndex - 1) + skl
                                skl = 0
                            Else
                                .Values(.SelIndex - 1) = .Values(.SelIndex - 1) + 1
                                skl = skl - 1
                            End If
                            
                        ElseIf Button = vbRightButton Then
                            
                            If skl = .Limit Then Exit Sub
                            
                            If Shift And 2 Then
                                If .Values(.SelIndex - 1) >= 3 Then
                                    .Values(.SelIndex - 1) = .Values(.SelIndex - 1) - 3
                                    skl = skl + 3
                                    
                                ElseIf .Values(.SelIndex - 1) >= 2 Then
                                    .Values(.SelIndex - 1) = .Values(.SelIndex - 1) - 2
                                    skl = skl + 2
                                ElseIf .Values(.SelIndex - 1) = 1 Then
                                    .Values(.SelIndex - 1) = 0
                                    skl = skl + 1
                                End If
                            
                            ElseIf Shift And 3 Then
                                skl = skl + .Values(.SelIndex - 1)
                                .Values(.SelIndex - 1) = 0
                            Else
                                If .Values(.SelIndex - 1) >= 1 Then
                                    .Values(.SelIndex - 1) = .Values(.SelIndex - 1) - 1
                                    skl = skl + 1
                                End If
                                
                            End If
                        End If

                    End If
                    
                Case eComponentEvent.MouseScrollUp: Call ScrollListUp(hwnd)
                Case eComponentEvent.MouseScrollDown: Call ScrollListDown(hwnd)
    
            End Select
        End With
    End If
    
    frmConnect.SkillPts = skl
    Call EditLabel(frmConnect.lblSkillLibres, CStr(skl), White)
End Sub

'Primitive Events

Private Sub ListBox_EventHandler(ByVal hwnd As Long, _
                                 ByVal msg As Long, _
                                 ByVal param3 As Long, _
                                 ByVal param4 As Long)

    Dim X As Integer, Y As Integer
    Dim LastDrawableLine As Integer
        
        
    Call LongToIntegers(param3, X, Y)
    
    If msg = eComponentEvent.MouseDown Then
        
        Y = Y - (Components(hwnd).Y - CharHeight \ 2)
        Y = (Y + 2) \ CharHeight
        
        With Components(hwnd)
        
            Select Case msg
                
                Case eComponentEvent.MouseDown
                    LastDrawableLine = Fix(.H / CharHeight)
                    
                    If Y > 0 And Y <= LastDrawableLine Then
                        .SelIndex = (.FirstRender - 1) + Y
                        
                        Dim cho As Integer
                        cho = .ChildOf
                        
                        If cho <> 0 Then
                            
                            Components(cho).Text = .Lines(.SelIndex).Text
                            Components(cho).SelIndex = .SelIndex
                            
                            Call Execute(cho, ChildClicked)
                            
                            .Visible = False
                            Components(cho).Expanded = False
                            
                        End If
                        
                    End If
                    
                Case eComponentEvent.MouseScrollUp: Call ScrollListUp(hwnd)
                Case eComponentEvent.MouseScrollDown: Call ScrollListDown(hwnd)
    
            End Select
        End With
    End If
    
End Sub

Public Sub TextBox_EventHandler(ByVal hwnd As Long, _
                                ByVal msg As Long, _
                                ByVal param3 As Long, _
                                ByVal param4 As Long)

    Dim i As Long
    Dim tempstr As String
    Dim Buffer As String
    
    Buffer = Components(hwnd).TextBuffer
    
    Select Case msg
        
        Case eComponentEvent.MouseUp
            Call SetFocus(hwnd)
            
        Case eComponentEvent.KeyPress
            If Not (param3 = vbKeyBack) And Not (param3 >= vbKeySpace And param3 <= 250) Then param3 = 0
            

            Buffer = Buffer + ChrW$(param3)
            
            'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
            For i = 1 To Len(Buffer)
                param3 = Asc(mid$(Buffer, i, 1))

                If param3 >= vbKeySpace And param3 <= 250 Then
                    tempstr = tempstr & ChrW$(param3)
                End If

                If param3 = vbKeyBack And Len(tempstr) > 0 Then
                    tempstr = Left$(tempstr, Len(tempstr) - 1)
                End If
            Next i

            If tempstr <> Buffer Then
                'We only set it if it's different, otherwise the event will be raised
                'constantly and the client will crush
                Buffer = tempstr
            End If

            Components(hwnd).TextBuffer = Buffer
    End Select
    
End Sub

'CSEH: ErrLog
Private Function CheckCabeza(ByVal Head As Integer) As Integer
        '<EhHeader>
        On Error GoTo CheckCabeza_Err
        '</EhHeader>
        
        If UserSexo = 0 Or UserRaza = 0 Then Exit Function
        
100     Select Case UserSexo

            Case eGenero.Hombre

105             Select Case UserRaza

                    Case eRaza.Humano

110                     If Head > HUMANO_H_ULTIMA_CABEZA Then
115                         CheckCabeza = HUMANO_H_PRIMER_CABEZA + (Head - HUMANO_H_ULTIMA_CABEZA) - 1
120                     ElseIf Head < HUMANO_H_PRIMER_CABEZA Then
125                         CheckCabeza = HUMANO_H_ULTIMA_CABEZA - (HUMANO_H_PRIMER_CABEZA - Head) + 1
                        Else
130                         CheckCabeza = Head
                        End If

135                 Case eRaza.Elfo

140                     If Head > ELFO_H_ULTIMA_CABEZA Then
145                         CheckCabeza = ELFO_H_PRIMER_CABEZA + (Head - ELFO_H_ULTIMA_CABEZA) - 1
150                     ElseIf Head < ELFO_H_PRIMER_CABEZA Then
155                         CheckCabeza = ELFO_H_ULTIMA_CABEZA - (ELFO_H_PRIMER_CABEZA - Head) + 1
                        Else
160                         CheckCabeza = Head
                        End If

165                 Case eRaza.ElfoOscuro

170                     If Head > DROW_H_ULTIMA_CABEZA Then
175                         CheckCabeza = DROW_H_PRIMER_CABEZA + (Head - DROW_H_ULTIMA_CABEZA) - 1
180                     ElseIf Head < DROW_H_PRIMER_CABEZA Then
185                         CheckCabeza = DROW_H_ULTIMA_CABEZA - (DROW_H_PRIMER_CABEZA - Head) + 1
                        Else
190                         CheckCabeza = Head
                        End If

195                 Case eRaza.Enano

200                     If Head > ENANO_H_ULTIMA_CABEZA Then
205                         CheckCabeza = ENANO_H_PRIMER_CABEZA + (Head - ENANO_H_ULTIMA_CABEZA) - 1
210                     ElseIf Head < ENANO_H_PRIMER_CABEZA Then
215                         CheckCabeza = ENANO_H_ULTIMA_CABEZA - (ENANO_H_PRIMER_CABEZA - Head) + 1
                        Else
220                         CheckCabeza = Head
                        End If

225                 Case eRaza.Gnomo

230                     If Head > GNOMO_H_ULTIMA_CABEZA Then
235                         CheckCabeza = GNOMO_H_PRIMER_CABEZA + (Head - GNOMO_H_ULTIMA_CABEZA) - 1
240                     ElseIf Head < GNOMO_H_PRIMER_CABEZA Then
245                         CheckCabeza = GNOMO_H_ULTIMA_CABEZA - (GNOMO_H_PRIMER_CABEZA - Head) + 1
                        Else
250                         CheckCabeza = Head
                        End If
                End Select

255         Case eGenero.Mujer

260             Select Case UserRaza

                    Case eRaza.Humano

265                     If Head > HUMANO_M_ULTIMA_CABEZA Then
270                         CheckCabeza = HUMANO_M_PRIMER_CABEZA + (Head - HUMANO_M_ULTIMA_CABEZA) - 1
275                     ElseIf Head < HUMANO_M_PRIMER_CABEZA Then
280                         CheckCabeza = HUMANO_M_ULTIMA_CABEZA - (HUMANO_M_PRIMER_CABEZA - Head) + 1
                        Else
285                         CheckCabeza = Head
                        End If

290                 Case eRaza.Elfo

295                     If Head > ELFO_M_ULTIMA_CABEZA Then
300                         CheckCabeza = ELFO_M_PRIMER_CABEZA + (Head - ELFO_M_ULTIMA_CABEZA) - 1
305                     ElseIf Head < ELFO_M_PRIMER_CABEZA Then
310                         CheckCabeza = ELFO_M_ULTIMA_CABEZA - (ELFO_M_PRIMER_CABEZA - Head) + 1
                        Else
315                         CheckCabeza = Head
                        End If

320                 Case eRaza.ElfoOscuro

325                     If Head > DROW_M_ULTIMA_CABEZA Then
330                         CheckCabeza = DROW_M_PRIMER_CABEZA + (Head - DROW_M_ULTIMA_CABEZA) - 1
335                     ElseIf Head < DROW_M_PRIMER_CABEZA Then
340                         CheckCabeza = DROW_M_ULTIMA_CABEZA - (DROW_M_PRIMER_CABEZA - Head) + 1
                        Else
345                         CheckCabeza = Head
                        End If

350                 Case eRaza.Enano

355                     If Head > ENANO_M_ULTIMA_CABEZA Then
360                         CheckCabeza = ENANO_M_PRIMER_CABEZA + (Head - ENANO_M_ULTIMA_CABEZA) - 1
365                     ElseIf Head < ENANO_M_PRIMER_CABEZA Then
370                         CheckCabeza = ENANO_M_ULTIMA_CABEZA - (ENANO_M_PRIMER_CABEZA - Head) + 1
                        Else
375                         CheckCabeza = Head
                        End If

380                 Case eRaza.Gnomo

385                     If Head > GNOMO_M_ULTIMA_CABEZA Then
390                         CheckCabeza = GNOMO_M_PRIMER_CABEZA + (Head - GNOMO_M_ULTIMA_CABEZA) - 1
395                     ElseIf Head < GNOMO_M_PRIMER_CABEZA Then
400                         CheckCabeza = GNOMO_M_ULTIMA_CABEZA - (GNOMO_M_PRIMER_CABEZA - Head) + 1
                        Else
405                         CheckCabeza = Head
                        End If
                End Select
        End Select
        '<EhFooter>
        Exit Function

CheckCabeza_Err:
        Call LogError("Error en CheckCabeza: " & Erl & " - " & Err.Description)
        '</EhFooter>
End Function

