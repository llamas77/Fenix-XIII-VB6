VERSION 5.00
Begin VB.Form frmFidelidad 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3750
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1950
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image2 
      Height          =   495
      Left            =   2040
      MouseIcon       =   "frmFidelidad.frx":0000
      MousePointer    =   99  'Custom
      Top             =   1200
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   840
      MouseIcon       =   "frmFidelidad.frx":030A
      MousePointer    =   99  'Custom
      Top             =   1200
      Width           =   735
   End
End
Attribute VB_Name = "frmFidelidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'F�nixAO 1.0
'
'Based on Argentum Online 0.99z
'Copyright (C) 2002 M�rquez Pablo Ignacio
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'You can contact me at:
'elpresi@fenixao.com.ar
'www.fenixao.com.ar

Private Fide As Byte

Public Sub SetFide(ByVal Faccion As eFaccion)
    Fide = Faccion
End Sub
Private Sub Form_Load()

If Fide = 1 Then
Me.Picture = LoadPicture(DirGraficos & "fidelidadrey.gif")
ElseIf Fide = 2 Then
Me.Picture = LoadPicture(DirGraficos & "fidelidadthek.gif")
Else
Unload Me
End If

End Sub

Private Sub Image1_Click()
Call WriteEligioFaccion(Fide)
Unload FrmElegirCamino
Unload Me
End Sub

Private Sub Image2_Click()
Unload Me
End Sub
