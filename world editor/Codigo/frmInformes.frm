VERSION 5.00
Begin VB.Form frmInformes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Informes"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4590
   Icon            =   "frmInformes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FXIIIWorldEditor.lvButtons_H cmdNPCs 
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   4200
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      Caption         =   "NPCs"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin FXIIIWorldEditor.lvButtons_H cmdTranslados 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   4200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "Traslados"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin FXIIIWorldEditor.lvButtons_H cmdObjetos 
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   4200
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Caption         =   "Objetos"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      Height          =   3855
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmInformes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************
Option Explicit


Private Sub ActalizarObjetos()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next
Dim Y As Integer
Dim X As Integer

If Not MapaCargado Then
    Exit Sub
End If

txtInfo.Text = "Informe de Objetos (X,Y)"

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        If MapData(X, Y).OBJInfo.objindex > 0 Then
            txtInfo.Text = txtInfo.Text & vbCrLf & X & "," & Y & " tiene " & MapData(X, Y).OBJInfo.Amount & " del Objeto " & MapData(X, Y).OBJInfo.objindex & " - " & ObjData(MapData(X, Y).OBJInfo.objindex).Name
        End If
    Next X
Next Y

End Sub

''
'   Genera el informe de NPCs
'

Private Sub ActalizarNPCs()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next
Dim Y As Integer
Dim X As Integer

If Not MapaCargado Then
    Exit Sub
End If

txtInfo.Text = "Informe de NPCs/Hostiles (X,Y)"

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        If MapData(X, Y).NpcIndex > 0 Then
            If MapData(X, Y).NpcIndex >= 500 Then
                txtInfo.Text = txtInfo.Text & vbCrLf & X & "," & Y & " tiene " & NpcData(MapData(X, Y).NpcIndex).Name & " (Hostil)"
            Else
                txtInfo.Text = txtInfo.Text & vbCrLf & X & "," & Y & " tiene " & NpcData(MapData(X, Y).NpcIndex).Name
            End If
        End If
    Next X
Next Y

End Sub

''
'   Genera el informe de Translados
'

Private Sub ActalizarTranslados()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next
Dim Y As Integer
Dim X As Integer

If Not MapaCargado Then
    Exit Sub
End If

txtInfo.Text = "Informe de Translados (X,Y)"

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
            If MapData(X, Y).TileExit.Map > 0 Then
                txtInfo.Text = txtInfo.Text & vbCrLf & X & "," & Y & " nos traslada a la posición " & MapData(X, Y).TileExit.X & "," & MapData(X, Y).TileExit.Y & " del Mapa " & MapData(X, Y).TileExit.Map
                If ((X < 20 And MapData(X, Y).TileExit.X < 20) Or (X > 80 And MapData(X, Y).TileExit.X > 80)) And (X <> MapData(X, Y).TileExit.X) Then
                    txtInfo.Text = txtInfo.Text & " (X sospechoso)"
                End If
                If ((Y < 20 And MapData(X, Y).TileExit.Y < 20) Or (Y > 80 And MapData(X, Y).TileExit.Y > 80)) And (Y <> MapData(X, Y).TileExit.Y) Then
                    txtInfo.Text = txtInfo.Text & " (Y sospechoso)"
                End If
            End If
    Next X
Next Y

End Sub


Private Sub cmdNPCs_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call ActalizarNPCs
End Sub

Private Sub cmdObjetos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call ActalizarObjetos
End Sub

Private Sub cmdTranslados_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call ActalizarTranslados
End Sub
