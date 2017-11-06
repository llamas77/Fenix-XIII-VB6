VERSION 5.00
Begin VB.Form frmElegirCamino 
   BackColor       =   &H80000002&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      Caption         =   "Más información"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4700
      MouseIcon       =   "frmElegirCamino.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      Caption         =   "Más información"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3120
      MouseIcon       =   "frmElegirCamino.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "Más información"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1530
      MouseIcon       =   "frmElegirCamino.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   3650
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   0
      MouseIcon       =   "frmElegirCamino.frx":091E
      MousePointer    =   99  'Custom
      Top             =   7080
      Width           =   735
   End
   Begin VB.Image Fidelidad 
      Height          =   255
      Index           =   2
      Left            =   4800
      MouseIcon       =   "frmElegirCamino.frx":0C28
      MousePointer    =   99  'Custom
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Image Fidelidad 
      Height          =   255
      Index           =   1
      Left            =   1560
      MouseIcon       =   "frmElegirCamino.frx":0F32
      MousePointer    =   99  'Custom
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Image command3 
      Height          =   375
      Left            =   3120
      MouseIcon       =   "frmElegirCamino.frx":123C
      MousePointer    =   99  'Custom
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BackStyle       =   0  'Transparent
      Caption         =   "Mantenerse neutral"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   930
      TabIndex        =   6
      Top             =   4610
      Width           =   5415
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmElegirCamino.frx":1546
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   900
      TabIndex        =   5
      Top             =   4950
      Width           =   5445
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmElegirCamino.frx":16FC
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   3810
      TabIndex        =   4
      Top             =   2040
      Width           =   2805
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmElegirCamino.frx":1802
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   600
      TabIndex        =   3
      Top             =   2100
      Width           =   2880
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmElegirCamino.frx":18FC
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1080
      TabIndex        =   2
      Top             =   840
      Width           =   5415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ser fiel a Lord Thek"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4180
      TabIndex        =   1
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ser fiel al Rey"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   1800
      Width           =   2295
   End
End
Attribute VB_Name = "frmElegirCamino"
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

Private Sub Command3_Click()
Call WriteEligioFaccion(0)
Unload Me
End Sub
Private Sub Fidelidad_Click(Index As Integer)

Unload frmFidelidad

Call frmFidelidad.SetFide(Index)

frmFidelidad.Show

End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture(DirGraficos & "Suclases3op.gif")
End Sub

Private Sub Image1_Click()

Unload Me

End Sub
Private Sub Label10_Click()
Ayuda = 1
SubAyuda = 2
FrmAyuda.Show
End Sub

Private Sub Label8_Click()
Ayuda = 1
SubAyuda = 1
FrmAyuda.Show
End Sub

Private Sub Label9_Click()
Ayuda = 1
SubAyuda = 3
FrmAyuda.Show
End Sub
