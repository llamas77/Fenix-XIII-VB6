VERSION 5.00
Begin VB.Form frmOptimizar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Optimizar Mapa"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3600
   Icon            =   "frmOptimizar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   3600
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkBloquearArbolesEtc 
      Caption         =   "Bloquear Arboles, Carteles, Foros y Yacimientos"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin VB.CheckBox chkMapearArbolesEtc 
      Caption         =   "Mapear Arboles, Carteles, Foros y Yacimientos que no esten en la 3ra Capa"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   3375
   End
   Begin VB.CheckBox chkQuitarTodoBordes 
      Caption         =   "Quitar NPCs, Objetos y Translados en los Bordes Exteriores"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   3375
   End
   Begin VB.CheckBox chkQuitarTrigTrans 
      Caption         =   "Quitar Trigger's en Translados"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   3375
   End
   Begin VB.CheckBox chkQuitarTrigBloq 
      Caption         =   "Quitar Trigger's Bloqueados"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3375
   End
   Begin VB.CheckBox chkQuitarTrans 
      Caption         =   "Quitar Translados Bloqueados"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin FXIIIWorldEditor.lvButtons_H cOptimizar 
      Default         =   -1  'True
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   1815
      _extentx        =   3201
      _extenty        =   1296
      caption         =   "&Optimizar"
      capalign        =   2
      backstyle       =   2
      cgradient       =   0
      font            =   "frmOptimizar.frx":78BA
      mode            =   0
      value           =   0
      cback           =   12648384
   End
   Begin FXIIIWorldEditor.lvButtons_H cCancelar 
      Height          =   735
      Left            =   1920
      TabIndex        =   6
      Top             =   2640
      Width           =   1575
      _extentx        =   2778
      _extenty        =   1296
      caption         =   "&Cancelar"
      capalign        =   2
      backstyle       =   2
      cgradient       =   0
      font            =   "frmOptimizar.frx":78DE
      mode            =   1
      value           =   0
      cback           =   -2147483633
   End
End
Attribute VB_Name = "frmOptimizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Optimizar()
'*************************************************
'Author: ^[GS]^
'Last modified: 16/10/06
'*************************************************
Dim Y As Integer
Dim X As Integer

If Not MapaCargado Then
    Exit Sub
End If

' Quita Translados Bloqueados
' Quita Trigger's Bloqueados
' Quita Trigger's en Translados
' Quita NPCs, Objetos y Translados en los Bordes Exteriores
' Mapea Arboles, Carteles, Foros y Yacimientos que no esten en la 3ra Capa

modEdicion.Deshacer_Add "Aplicar Optimizacion del Mapa" ' Hago deshacer

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        ' ** Quitar NPCs, Objetos y Translados en los Bordes Exteriores
        If (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) And chkQuitarTodoBordes.value = 1 Then
             'Quitar NPCs
            If MapData(X, Y).NpcIndex > 0 Then
                EraseChar MapData(X, Y).CharIndex
                MapData(X, Y).NpcIndex = 0
            End If
            ' Quitar Objetos
            MapData(X, Y).OBJInfo.objindex = 0
            MapData(X, Y).OBJInfo.Amount = 0
            MapData(X, Y).ObjGrh.GrhIndex = 0
            ' Quitar Translados
            MapData(X, Y).TileExit.Map = 0
            MapData(X, Y).TileExit.X = 0
            MapData(X, Y).TileExit.Y = 0
            ' Quitar Triggers
            MapData(X, Y).Trigger = 0
        End If
        ' ** Quitar Translados y Triggers en Bloqueo
        If MapData(X, Y).Blocked = 1 Then
            If MapData(X, Y).TileExit.Map > 0 And chkQuitarTrans.value = 1 Then ' Quita Translado Bloqueado
                MapData(X, Y).TileExit.Map = 0
                MapData(X, Y).TileExit.Y = 0
                MapData(X, Y).TileExit.X = 0
            ElseIf MapData(X, Y).Trigger > 0 And chkQuitarTrigBloq.value = 1 Then ' Quita Trigger Bloqueado
                MapData(X, Y).Trigger = 0
            End If
        End If
        ' ** Quitar Triggers en Translado
        If MapData(X, Y).TileExit.Map > 0 And chkQuitarTrigTrans.value = 1 Then
            If MapData(X, Y).Trigger > 0 Then ' Quita Trigger en Translado
                MapData(X, Y).Trigger = 0
            End If
        End If
        ' ** Mapea Arboles, Carteles, Foros y Yacimientos que no esten en la 3ra Capa
        If MapData(X, Y).OBJInfo.objindex > 0 And (chkMapearArbolesEtc.value = 1 Or chkBloquearArbolesEtc.value = 1) Then
            Select Case ObjData(MapData(X, Y).OBJInfo.objindex).ObjType
                Case 4, 8, 10, 22 ' Arboles, Carteles, Foros, Yacimientos
                    If MapData(X, Y).Graphic(3).GrhIndex <> MapData(X, Y).ObjGrh.GrhIndex And chkMapearArbolesEtc.value = 1 Then MapData(X, Y).Graphic(3) = MapData(X, Y).ObjGrh
                    If chkBloquearArbolesEtc.value = 1 And MapData(X, Y).Blocked = 0 Then MapData(X, Y).Blocked = 1
            End Select
        End If
    Next X
Next Y

'Set changed flag
MapInfo.Changed = 1

End Sub

Private Sub cCancelar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 22/09/06
'*************************************************
Unload Me
End Sub

Private Sub cOptimizar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 22/09/06
'*************************************************
Call Optimizar
End Sub


