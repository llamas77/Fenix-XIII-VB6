VERSION 5.00
Begin VB.Form frmMapInfo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Información del Mapa"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4005
   Icon            =   "frmMapInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   4005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMapLevel 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1260
      MaxLength       =   2
      TabIndex        =   17
      Top             =   1170
      Width           =   2625
   End
   Begin VB.CheckBox chkRestringir 
      Caption         =   "Solo Newbies"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   300
      TabIndex        =   16
      Top             =   2820
      Width           =   1575
   End
   Begin VB.TextBox txtMapVersion 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1260
      TabIndex        =   14
      Text            =   "0"
      Top             =   450
      Width           =   2655
   End
   Begin FXIIIWorldEditor.lvButtons_H cmdMusica 
      Height          =   315
      Left            =   3180
      TabIndex        =   13
      Top             =   810
      Width           =   735
      _extentx        =   1296
      _extenty        =   556
      caption         =   "&Más"
      capalign        =   2
      backstyle       =   2
      cgradient       =   0
      font            =   "frmMapInfo.frx":78BA
      mode            =   0
      value           =   0
      cback           =   -2147483633
   End
   Begin FXIIIWorldEditor.lvButtons_H cmdCerrar 
      Height          =   435
      Left            =   90
      TabIndex        =   12
      Top             =   3330
      Width           =   3855
      _extentx        =   6641
      _extenty        =   767
      caption         =   "&Aceptar"
      capalign        =   2
      backstyle       =   2
      cgradient       =   0
      font            =   "frmMapInfo.frx":78E6
      mode            =   0
      value           =   0
      cback           =   -2147483633
   End
   Begin VB.CheckBox chkMapPK 
      Caption         =   "Zona Insegura"
      BeginProperty DataFormat 
         Type            =   4
         Format          =   "0%"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   8
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   2790
      Width           =   1575
   End
   Begin VB.ComboBox txtMapTerreno 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmMapInfo.frx":7912
      Left            =   1260
      List            =   "frmMapInfo.frx":791F
      TabIndex        =   9
      Top             =   1920
      Width           =   2655
   End
   Begin VB.ComboBox txtMapZona 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmMapInfo.frx":793C
      Left            =   1260
      List            =   "frmMapInfo.frx":7949
      TabIndex        =   8
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox txtMapMusica 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1260
      TabIndex        =   7
      Text            =   "0"
      Top             =   810
      Width           =   1815
   End
   Begin VB.TextBox txtMapNombre 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1260
      TabIndex        =   6
      Text            =   "Nuevo Mapa"
      Top             =   90
      Width           =   2655
   End
   Begin VB.CheckBox chkMapBackup 
      Caption         =   "Backup"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   300
      TabIndex        =   3
      Top             =   2490
      Width           =   1575
   End
   Begin VB.CheckBox chkMapMagiaSinEfecto 
      Caption         =   "No permitir magia"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   2490
      Width           =   1575
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Versión:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   480
      Width           =   675
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   135
      X2              =   3870
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   450
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Terreno:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1950
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Zona:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1590
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Musica:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   705
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   3870
      Y1              =   3240
      Y2              =   3240
   End
End
Attribute VB_Name = "frmMapInfo"
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

Private Sub chkMapBackup_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
MapInfo.BackUp = chkMapBackup.value
MapInfo.Changed = 1
End Sub

Private Sub chkMapMagiaSinEfecto_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
MapInfo.NoMagia = chkMapMagiaSinEfecto.value
MapInfo.Changed = 1
End Sub

Private Sub chkMapPK_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
MapInfo.Pk = chkMapPK.value
MapInfo.Changed = 1
End Sub

Private Sub cmdCerrar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Me.Hide
End Sub

Private Sub cmdMusica_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
frmMusica.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If UnloadMode = vbFormControlMenu Then
    Cancel = True
    Me.Hide
End If
End Sub

Private Sub txtMapMusica_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
MapInfo.Music = txtMapMusica.Text
frmMain.lblMapMusica.Caption = MapInfo.Music
MapInfo.Changed = 1
End Sub

Private Sub txtMapVersion_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
MapInfo.MapVersion = txtMapVersion.Text
frmMain.lblMapVersion.Caption = MapInfo.MapVersion
MapInfo.Changed = 1
End Sub

Private Sub txtMapNombre_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
MapInfo.Name = txtMapNombre.Text
frmMain.lblMapNombre.Caption = MapInfo.Name
MapInfo.Changed = 1
End Sub

Private Sub txtMapRestringir_KeyPress(KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
KeyAscii = 0
End Sub

Private Sub txtMapRestringir_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
MapInfo.Restringir = Me.chkRestringir.value
MapInfo.Changed = 1
End Sub

Private Sub txtMapTerreno_KeyPress(KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
KeyAscii = 0
End Sub

Private Sub txtMapTerreno_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
MapInfo.Terreno = txtMapTerreno.Text
MapInfo.Changed = 1
End Sub

Private Sub txtMapZona_KeyPress(KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
KeyAscii = 0
End Sub

Private Sub txtMapZona_LostFocus()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
MapInfo.Zona = txtMapZona.Text
MapInfo.Changed = 1
End Sub
