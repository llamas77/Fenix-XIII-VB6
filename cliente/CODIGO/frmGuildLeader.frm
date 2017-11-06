VERSION 5.00
Begin VB.Form frmGuildLeader 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8655
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
   ScaleHeight     =   5385
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraOpciones 
      Caption         =   "Opciones"
      Height          =   1215
      Left            =   3840
      TabIndex        =   26
      Top             =   3960
      Width           =   4455
      Begin VB.OptionButton OptAbierto 
         Caption         =   "Cerrado"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   29
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton OptAbierto 
         Caption         =   "Abierto"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblIngreso 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ingreso:"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame FraReclutadores2 
      Caption         =   "Reclutadores"
      Height          =   1215
      Left            =   240
      TabIndex        =   23
      Top             =   3960
      Width           =   3495
      Begin VB.CommandButton cmdDegradar 
         Caption         =   "Degradar"
         Height          =   360
         Left            =   2160
         TabIndex        =   25
         Top             =   240
         Width           =   990
      End
      Begin VB.ListBox lstReclutadores 
         Height          =   645
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame FraBloqueados 
      Caption         =   "Bloqueados"
      Height          =   1935
      Left            =   3840
      TabIndex        =   18
      Top             =   2040
      Width           =   4455
      Begin VB.CommandButton cmdDesbloquearTodos 
         Caption         =   "Desbloquear todos"
         Height          =   480
         Left            =   2400
         TabIndex        =   22
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdDesbloquear 
         Caption         =   "Desbloquear"
         Height          =   360
         Left            =   2400
         TabIndex        =   21
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdBloquearNuevo 
         Caption         =   "Bloquear a..."
         Height          =   360
         Left            =   3120
         TabIndex        =   20
         Top             =   1440
         Width           =   1215
      End
      Begin VB.ListBox lstBloqueados 
         Height          =   1035
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame FraSolicitudes 
      Caption         =   "Solicitudes"
      Height          =   1935
      Left            =   240
      TabIndex        =   12
      Top             =   2040
      Width           =   3495
      Begin VB.CommandButton cmdRechazarTodas 
         Caption         =   "Rechazar todas"
         Height          =   360
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton cmdRechazar 
         Caption         =   "Rechazar"
         Height          =   360
         Left            =   2280
         TabIndex        =   16
         Top             =   1080
         Width           =   990
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   360
         Left            =   2280
         TabIndex        =   15
         Top             =   720
         Width           =   990
      End
      Begin VB.CommandButton cmdVerInfo 
         Caption         =   "Ver Info"
         Height          =   360
         Left            =   2280
         TabIndex        =   14
         Top             =   360
         Width           =   990
      End
      Begin VB.ListBox lstSolicitudes 
         Height          =   1035
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame FraMiembros 
      Caption         =   "Miembros"
      Height          =   1575
      Left            =   4440
      TabIndex        =   8
      Top             =   240
      Width           =   3855
      Begin VB.CommandButton cmdAscenderA 
         Caption         =   "Ascender a Reclutador"
         Height          =   480
         Left            =   2640
         TabIndex        =   11
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdEchar 
         Caption         =   "Echar"
         Height          =   360
         Left            =   2640
         TabIndex        =   10
         Top             =   360
         Width           =   990
      End
      Begin VB.ListBox lslMiembros 
         Height          =   840
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame FraInfo 
      Caption         =   "Info"
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4095
      Begin VB.Frame FraReclutadores 
         Caption         =   "Reclutadores"
         Height          =   975
         Left            =   2400
         TabIndex        =   4
         Top             =   360
         Width           =   1575
         Begin VB.Label lblRec3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rec3"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   720
            Width           =   360
         End
         Begin VB.Label lblRec2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rec2"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   480
            Width           =   360
         End
         Begin VB.Label lblRec1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rec1"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   360
         End
      End
      Begin VB.Label lblLíder 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Líder:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   405
      End
      Begin VB.Label lblFundador 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fundador:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   750
      End
      Begin VB.Label lblCantidadDe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad de miembros: 0"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1785
      End
   End
End
Attribute VB_Name = "frmGuildLeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
