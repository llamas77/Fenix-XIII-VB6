VERSION 5.00
Begin VB.Form frmCargando 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Argentum"
   ClientHeight    =   3495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6060
   Icon            =   "frmCargando.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Height          =   2340
      Left            =   0
      Picture         =   "frmCargando.frx":78BA
      ScaleHeight     =   2280
      ScaleWidth      =   6000
      TabIndex        =   0
      Top             =   0
      Width           =   6060
      Begin VB.Label verX 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "v?.?.?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   120
         TabIndex        =   2
         Top             =   90
         Width           =   555
      End
   End
   Begin VB.Image P6 
      Height          =   480
      Left            =   5430
      Picture         =   "frmCargando.frx":3417E
      ToolTipText     =   "Función de Trigger"
      Top             =   2670
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label L 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Triggers"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   195
      Index           =   5
      Left            =   5400
      TabIndex        =   8
      Top             =   3180
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label L 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Objetos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   195
      Index           =   4
      Left            =   4500
      TabIndex        =   7
      Top             =   3180
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label L 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NPC's"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   195
      Index           =   3
      Left            =   3690
      TabIndex        =   6
      Top             =   3180
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label L 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cabezas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   195
      Index           =   2
      Left            =   2220
      TabIndex        =   5
      Top             =   3180
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label L 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cuerpos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   195
      Index           =   1
      Left            =   1380
      TabIndex        =   4
      Top             =   3180
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label L 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Superficies"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   3180
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image P5 
      Height          =   480
      Left            =   4560
      Picture         =   "frmCargando.frx":34DC0
      ToolTipText     =   "Objetos"
      Top             =   2700
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image P1 
      Height          =   480
      Left            =   480
      Picture         =   "frmCargando.frx":35604
      ToolTipText     =   "Base de Datos"
      Top             =   2640
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image P3 
      Height          =   480
      Left            =   2280
      Picture         =   "frmCargando.frx":35E48
      ToolTipText     =   "Cabezas"
      Top             =   2700
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image P4 
      Height          =   480
      Left            =   3630
      Picture         =   "frmCargando.frx":3668C
      ToolTipText     =   "NPC's"
      Top             =   2760
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image P2 
      Height          =   480
      Left            =   1440
      Picture         =   "frmCargando.frx":372CE
      ToolTipText     =   "Cuerpos"
      Top             =   2670
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label X 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   5655
   End
End
Attribute VB_Name = "frmCargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
