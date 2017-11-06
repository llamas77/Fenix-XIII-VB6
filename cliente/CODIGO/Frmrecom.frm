VERSION 5.00
Begin VB.Form frmRecompensa 
   BackColor       =   &H80000002&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Height          =   375
      Left            =   0
      MousePointer    =   99  'Custom
      Top             =   4320
      Width           =   735
   End
   Begin VB.Image Eleccion 
      Height          =   375
      Index           =   2
      Left            =   4800
      MousePointer    =   99  'Custom
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Image Eleccion 
      Height          =   375
      Index           =   1
      Left            =   1560
      MousePointer    =   99  'Custom
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Descripcion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Index           =   2
      Left            =   3710
      TabIndex        =   4
      Top             =   2150
      Width           =   3015
   End
   Begin VB.Label Descripcion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Index           =   1
      Left            =   490
      TabIndex        =   3
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Width           =   6225
   End
   Begin VB.Label Nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   4070
      TabIndex        =   1
      Top             =   1755
      Width           =   2415
   End
   Begin VB.Label Nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   860
      TabIndex        =   0
      Top             =   1755
      Width           =   2535
   End
End
Attribute VB_Name = "frmRecompensa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FénixAO 1.0
'
'Based on Argentum Online 0.99z
'Copyright (C) 2002 Márquez Pablo Ignacio
'
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
'You should have received a copy of the Affero General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'You can contact the original creator of Argentum Online at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'You can contact me at:
'elpresi@fenixao.com.ar
'www.fenixao.com.ar

Private Sub Eleccion_Click(Index As Integer)

Call WriteEligioRecompensa(Index)
'Call AddtoRichTextBox(frmMain.RecTxt, "¡Has elegido la recompensa " & Nombre(Index) & "!", 255, 250, 55, 1, 0)
Unload Me

End Sub
Private Sub Form_Load()

Me.Picture = LoadPicture(DirGraficos & "Suclases2op.gif")

End Sub
Private Sub Image1_Click()

Unload Me

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If bmoving = False And Button = vbLeftButton Then
    dX = X
    dy = Y
    bmoving = True
End If

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If bmoving And (X <> dX Or Y <> dy) Then Move Left + (X - dX), Top + (Y - dy)

End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then bmoving = False

End Sub
