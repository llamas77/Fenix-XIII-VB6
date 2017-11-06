VERSION 5.00
Begin VB.Form frmGuildFoundation 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCrear 
      Caption         =   "Crear"
      Height          =   360
      Left            =   1560
      TabIndex        =   9
      Top             =   2640
      Width           =   990
   End
   Begin VB.OptionButton optEntrance 
      Caption         =   "Aprobación"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   8
      Top             =   2160
      Width           =   1455
   End
   Begin VB.OptionButton optEntrance 
      Caption         =   "Libre"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   2160
      Width           =   1455
   End
   Begin VB.ComboBox cmbFaction 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox txtReqLevel 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label lblEntrada 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entrada"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   570
   End
   Begin VB.Label lblFacción 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Facción"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   540
   End
   Begin VB.Label lblNivelRequerido 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel Requerido"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1125
   End
   Begin VB.Label lblNombreDel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del clan"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1140
   End
End
Attribute VB_Name = "frmGuildFoundation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCrear_Click()
    Call WriteGuildConfirmFoundation(txtName.Text, Val(txtReqLevel.Text), cmbFaction.Text, IIf(optEntrance(0).Value = True, 0, 1))
    Unload Me
End Sub

Private Sub Form_Load()
    
    cmbFaction.Clear
    
    Select Case UserFaccion
    
        Case eFaccion.Real
        
            cmbFaction.AddItem "Real"
            cmbFaction.AddItem "Neutral"
            
        Case eFaccion.Caos
        
            cmbFaction.AddItem "Caos"
            cmbFaction.AddItem "Neutral"
        Case eFaccion.Neutral
        
            cmbFaction.AddItem "Caos"
            cmbFaction.AddItem "Neutral"
            cmbFaction.AddItem "Real"
            
    End Select
    
End Sub

Private Sub txtReqLevel_Change()
    If Not IsNumeric(txtReqLevel.Text) Then Exit Sub
    
    If Val(txtReqLevel.Text) > 50 Then
        textreqlevel.Text = CStr(50)
    End If
    
End Sub
