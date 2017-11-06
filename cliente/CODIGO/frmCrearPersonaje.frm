VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11985
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   799
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picSkills 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   2400
      Left            =   1080
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   190
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   4800
      Width           =   2850
      Begin VB.Label lblSkill1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skill1"
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
         Height          =   195
         Index           =   7
         Left            =   75
         TabIndex        =   40
         Top             =   2175
         Width           =   480
      End
      Begin VB.Label lblSkill1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skill1"
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
         Height          =   195
         Index           =   6
         Left            =   75
         TabIndex        =   39
         Top             =   1875
         Width           =   480
      End
      Begin VB.Label lblSkill1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skill1"
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
         Height          =   195
         Index           =   5
         Left            =   75
         TabIndex        =   38
         Top             =   1575
         Width           =   480
      End
      Begin VB.Label lblSkill1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skill1"
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
         Height          =   195
         Index           =   4
         Left            =   75
         TabIndex        =   35
         Top             =   1275
         Width           =   480
      End
      Begin VB.Label lblSkill1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skill1"
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
         Height          =   195
         Index           =   3
         Left            =   75
         TabIndex        =   34
         Top             =   975
         Width           =   480
      End
      Begin VB.Label lblSkill1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skill1"
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
         Height          =   195
         Index           =   2
         Left            =   75
         TabIndex        =   33
         Top             =   675
         Width           =   480
      End
      Begin VB.Label lblSkill1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skill1"
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
         Height          =   195
         Index           =   1
         Left            =   75
         TabIndex        =   32
         Top             =   375
         Width           =   480
      End
      Begin VB.Label lblSkill1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skill1"
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
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   31
         Top             =   75
         Width           =   480
      End
   End
   Begin VB.TextBox txtMail 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
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
      Height          =   225
      Left            =   240
      TabIndex        =   3
      Top             =   2760
      Width           =   4575
   End
   Begin VB.TextBox txtConfirmPasswd 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   3960
      Width           =   2415
   End
   Begin VB.TextBox txtPasswd 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Timer tAnimacion 
      Left            =   120
      Top             =   0
   End
   Begin VB.ComboBox lstGenero 
      BackColor       =   &H00000000&
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
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":0000
      Left            =   5160
      List            =   "frmCrearPersonaje.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1950
      Width           =   2625
   End
   Begin VB.ComboBox lstRaza 
      BackColor       =   &H00000000&
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
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":001D
      Left            =   5160
      List            =   "frmCrearPersonaje.frx":001F
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1320
      Width           =   2625
   End
   Begin VB.ComboBox lstHogar 
      BackColor       =   &H00000000&
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
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":0021
      Left            =   5160
      List            =   "frmCrearPersonaje.frx":0023
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2640
      Width           =   2625
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
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
      Height          =   225
      Left            =   1560
      MaxLength       =   30
      TabIndex        =   0
      Top             =   840
      Width           =   5055
   End
   Begin VB.PictureBox picPJ 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   10320
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   12
      Top             =   1320
      Width           =   615
   End
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   10320
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   37
      TabIndex        =   13
      Top             =   1320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   1
      Left            =   10035
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   26
      Top             =   840
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   2
      Left            =   10440
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   27
      Top             =   840
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   3
      Left            =   10845
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   28
      Top             =   840
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   4
      Left            =   11280
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   29
      Top             =   840
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   0
      Left            =   9630
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   25
      Top             =   840
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picSkillsTemp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   1080
      ScaleHeight     =   161
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   190
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   4800
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   3960
      Picture         =   "frmCrearPersonaje.frx":0025
      Top             =   7080
      Width           =   225
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   3960
      Picture         =   "frmCrearPersonaje.frx":0367
      Top             =   4800
      Width           =   225
   End
   Begin VB.Label lblPuntosRestantes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Puntos restantes:"
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
      Height          =   195
      Left            =   1080
      TabIndex        =   37
      Top             =   4560
      Width           =   1500
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      Index           =   3
      Visible         =   0   'False
      X1              =   695
      X2              =   721
      Y1              =   81
      Y2              =   81
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      Index           =   2
      Visible         =   0   'False
      X1              =   695
      X2              =   721
      Y1              =   55
      Y2              =   55
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      Index           =   1
      Visible         =   0   'False
      X1              =   721
      X2              =   721
      Y1              =   56
      Y2              =   80
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      Index           =   0
      Visible         =   0   'False
      X1              =   695
      X2              =   695
      Y1              =   56
      Y2              =   80
   End
   Begin VB.Label lblAtributoFinal 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
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
      Height          =   195
      Index           =   5
      Left            =   6765
      TabIndex        =   24
      Top             =   6030
      Width           =   225
   End
   Begin VB.Label lblAtributoFinal 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
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
      Height          =   195
      Index           =   4
      Left            =   6885
      TabIndex        =   23
      Top             =   7830
      Width           =   225
   End
   Begin VB.Label lblAtributoFinal 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
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
      Height          =   195
      Index           =   3
      Left            =   6765
      TabIndex        =   22
      Top             =   6885
      Width           =   225
   End
   Begin VB.Label lblAtributoFinal 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
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
      Height          =   195
      Index           =   2
      Left            =   6765
      TabIndex        =   21
      Top             =   4980
      Width           =   225
   End
   Begin VB.Label lblAtributoFinal 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
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
      Height          =   195
      Index           =   1
      Left            =   6885
      TabIndex        =   20
      Top             =   4050
      Width           =   225
   End
   Begin VB.Label lblModRaza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
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
      Height          =   195
      Index           =   5
      Left            =   6270
      TabIndex        =   19
      Top             =   6030
      Width           =   225
   End
   Begin VB.Label lblModRaza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
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
      Height          =   195
      Index           =   4
      Left            =   6390
      TabIndex        =   18
      Top             =   7830
      Width           =   225
   End
   Begin VB.Label lblModRaza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
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
      Height          =   195
      Index           =   3
      Left            =   6270
      TabIndex        =   17
      Top             =   6885
      Width           =   225
   End
   Begin VB.Label lblModRaza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
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
      Height          =   195
      Index           =   2
      Left            =   6270
      TabIndex        =   16
      Top             =   4980
      Width           =   225
   End
   Begin VB.Label lblModRaza 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "+0"
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
      Height          =   195
      Index           =   1
      Left            =   6390
      TabIndex        =   15
      Top             =   4050
      Width           =   225
   End
   Begin VB.Image imgAtributos 
      Height          =   270
      Left            =   3960
      Top             =   2745
      Width           =   975
   End
   Begin VB.Label lblHelp 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   8400
      TabIndex        =   14
      Top             =   3000
      Width           =   3375
   End
   Begin VB.Image imgVolver 
      Height          =   450
      Left            =   120
      Top             =   8520
      Width           =   1290
   End
   Begin VB.Image imgCrear 
      Height          =   435
      Left            =   1080
      Top             =   7800
      Width           =   3330
   End
   Begin VB.Image imgGenero 
      Height          =   240
      Left            =   5160
      Top             =   1680
      Width           =   705
   End
   Begin VB.Image imgRaza 
      Height          =   255
      Left            =   5160
      Top             =   1080
      Width           =   570
   End
   Begin VB.Image imgPuebloOrigen 
      Height          =   225
      Left            =   5160
      Top             =   2400
      Width           =   1425
   End
   Begin VB.Image imgConstitucion 
      Height          =   255
      Left            =   5640
      Top             =   5520
      Width           =   1920
   End
   Begin VB.Image imgCarisma 
      Height          =   240
      Left            =   6120
      Top             =   7440
      Width           =   765
   End
   Begin VB.Image imgInteligencia 
      Height          =   240
      Left            =   5760
      Top             =   6480
      Width           =   1725
   End
   Begin VB.Image imgAgilidad 
      Height          =   240
      Left            =   5880
      Top             =   4440
      Width           =   735
   End
   Begin VB.Image imgFuerza 
      Height          =   360
      Left            =   5880
      Top             =   3480
      Width           =   1155
   End
   Begin VB.Image imgF 
      Height          =   270
      Left            =   6735
      Top             =   3195
      Width           =   270
   End
   Begin VB.Image imgM 
      Height          =   270
      Left            =   6270
      Top             =   3195
      Width           =   270
   End
   Begin VB.Image imgD 
      Height          =   270
      Left            =   5805
      Top             =   3210
      Width           =   270
   End
   Begin VB.Image imgConfirmPasswd 
      Height          =   255
      Left            =   360
      Top             =   3600
      Width           =   1440
   End
   Begin VB.Image imgPasswd 
      Height          =   255
      Left            =   240
      Top             =   3000
      Width           =   2970
   End
   Begin VB.Image imgNombre 
      Height          =   240
      Left            =   360
      Top             =   480
      Width           =   2595
   End
   Begin VB.Image imgMail 
      Height          =   240
      Left            =   360
      Top             =   2400
      Width           =   1395
   End
   Begin VB.Image imgTirarDados 
      Height          =   1485
      Left            =   10080
      Top             =   6240
      Width           =   1200
   End
   Begin VB.Image DirPJ 
      Height          =   225
      Index           =   1
      Left            =   10695
      Picture         =   "frmCrearPersonaje.frx":06A9
      Top             =   2280
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image DirPJ 
      Height          =   225
      Index           =   0
      Left            =   10320
      Picture         =   "frmCrearPersonaje.frx":09BB
      Top             =   2280
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image HeadPJ 
      Height          =   225
      Index           =   1
      Left            =   11700
      Picture         =   "frmCrearPersonaje.frx":0CCD
      Top             =   885
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image HeadPJ 
      Height          =   225
      Index           =   0
      Left            =   9315
      Picture         =   "frmCrearPersonaje.frx":0FDF
      Top             =   885
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   3120
      Left            =   8880
      Stretch         =   -1  'True
      Top             =   9120
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.Image imgDados 
      Height          =   885
      Left            =   8640
      MouseIcon       =   "frmCrearPersonaje.frx":12F1
      MousePointer    =   99  'Custom
      Top             =   6480
      Width           =   900
   End
   Begin VB.Image imgHogar 
      Height          =   2850
      Left            =   5640
      Picture         =   "frmCrearPersonaje.frx":1443
      Top             =   9120
      Visible         =   0   'False
      Width           =   2985
   End
   Begin VB.Label lblAtributos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "18"
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
      Height          =   195
      Index           =   4
      Left            =   5940
      TabIndex        =   11
      Top             =   7830
      Width           =   225
   End
   Begin VB.Label lblAtributos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "18"
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
      Height          =   195
      Index           =   3
      Left            =   5820
      TabIndex        =   10
      Top             =   6885
      Width           =   225
   End
   Begin VB.Label lblAtributos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "18"
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
      Height          =   195
      Index           =   5
      Left            =   5820
      TabIndex        =   9
      Top             =   6030
      Width           =   225
   End
   Begin VB.Label lblAtributos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "18"
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
      Height          =   195
      Index           =   2
      Left            =   5820
      TabIndex        =   8
      Top             =   4980
      Width           =   225
   End
   Begin VB.Label lblAtributos 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "18"
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
      Height          =   195
      Index           =   1
      Left            =   5940
      TabIndex        =   7
      Top             =   4050
      Width           =   225
   End
End
Attribute VB_Name = "frmCrearPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Private cBotonPasswd As clsGraphicalButton
Private cBotonTirarDados As clsGraphicalButton
Private cBotonMail As clsGraphicalButton
Private cBotonNombre As clsGraphicalButton
Private cBotonConfirmPasswd As clsGraphicalButton
Private cBotonAtributos As clsGraphicalButton
Private cBotonD As clsGraphicalButton
Private cBotonM As clsGraphicalButton
Private cBotonF As clsGraphicalButton
Private cBotonFuerza As clsGraphicalButton
Private cBotonAgilidad As clsGraphicalButton
Private cBotonInteligencia As clsGraphicalButton
Private cBotonCarisma As clsGraphicalButton
Private cBotonConstitucion As clsGraphicalButton
Private cBotonEvasion As clsGraphicalButton
Private cBotonMagia As clsGraphicalButton
Private cBotonVida As clsGraphicalButton
Private cBotonEscudos As clsGraphicalButton
Private cBotonArmas As clsGraphicalButton
Private cBotonArcos As clsGraphicalButton
Private cBotonEspecialidad As clsGraphicalButton
Private cBotonPuebloOrigen As clsGraphicalButton
Private cBotonRaza As clsGraphicalButton
Private cBotonClase As clsGraphicalButton
Private cBotonGenero As clsGraphicalButton
Private cBotonVolver As clsGraphicalButton
Private cBotonCrear As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Private picFullStar As Picture
Private picHalfStar As Picture
Private picGlowStar As Picture

Private Enum eHelp
    iePasswd
    ieTirarDados
    ieMail
    ieNombre
    ieConfirmPasswd
    ieAtributos
    ieD
    ieM
    ieF
    ieFuerza
    ieAgilidad
    ieInteligencia
    ieCarisma
    ieConstitucion
    ieEvasion
    ieMagia
    ieVida
    ieEscudos
    ieArmas
    ieArcos
    ieEspecialidad
    iePuebloOrigen
    ieRaza
    ieClase
    ieGenero
End Enum

Private vHelp(25) As String

Private Type tModRaza
    Fuerza As Single
    Agilidad As Single
    Inteligencia As Single
    Carisma As Single
    Constitucion As Single
End Type

Private ModRaza() As tModRaza

Private NroRazas As Integer

Private Cargando As Boolean

Private currentGrh As Long
Private Dir As E_Heading

Private TopList As Integer
Private SkillPts As Integer
Private YTemp As Integer
Private uSkills(1 To NUMSKILLS) As Byte
Private MouseButton As Integer
Private MouseShift As Integer

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)


Private Sub command1_Click()


End Sub

Private Sub Command2_Click()


End Sub

Private Sub Form_Load()
    Me.Picture = LoadPicture(DirGraficos & "VentanaCrearPersonaje.jpg")

    Cargando = True
    Call LoadCharInfo
    Call CargarEspecialidades
    
    Call IniciarGraficos
    Call CargarCombos
    
    Call LoadHelp
    
    'Call DrawImageInPicture(picPJ, Me.Picture, 0, 0, , , picPJ.Left, picPJ.Top)
    Dir = SOUTH
    
    Call TirarDados
    
    Cargando = False
    
    UserClase = eClass.Ciudadano
    UserSexo = 0
    UserRaza = 0
    UserHogar = 0
    UserEmail = ""
    UserHead = 0
    
    Dim i As Long

    For i = 0 To lblSkill1.UBound
        lblSkill1(i).Caption = SkillsNames(i + 1)
    Next
    
    'clear
    Erase uSkills
    'ReDim uSkills(1 To NUMSKILLS) As Byte
    
    SkillPts = 10
    lblPuntosRestantes.Caption = "Puntos restantes: " & SkillPts
    DrawSklPt TopList + YTemp + 1
    
End Sub

Private Sub CargarEspecialidades()

'    ReDim vEspecialidades(1 To NroClases)
    
 '   vEspecialidades(eClass.Cazador) = "Ocultarse"
 '   vEspecialidades(eClass.Ladron) = "Robar y Ocultarse"
 '   vEspecialidades(eClass.Asesino) = "Apuñalar"
 '   vEspecialidades(eClass.Bandido) = "Combate Sin Armas"
 '   vEspecialidades(eClass.Druida) = "Domar"
 '   vEspecialidades(eClass.Pirata) = "Navegar"
 '   vEspecialidades(eClass.Trabajador) = "Extracción y Construcción"
End Sub

Private Sub IniciarGraficos()

    Dim GrhPath As String
    GrhPath = DirGraficos
    
    Set cBotonPasswd = New clsGraphicalButton
    Set cBotonTirarDados = New clsGraphicalButton
    Set cBotonMail = New clsGraphicalButton
    Set cBotonNombre = New clsGraphicalButton
    Set cBotonConfirmPasswd = New clsGraphicalButton
    Set cBotonAtributos = New clsGraphicalButton
    Set cBotonD = New clsGraphicalButton
    Set cBotonM = New clsGraphicalButton
    Set cBotonF = New clsGraphicalButton
    Set cBotonFuerza = New clsGraphicalButton
    Set cBotonAgilidad = New clsGraphicalButton
    Set cBotonInteligencia = New clsGraphicalButton
    Set cBotonCarisma = New clsGraphicalButton
    Set cBotonConstitucion = New clsGraphicalButton
    Set cBotonEvasion = New clsGraphicalButton
    Set cBotonMagia = New clsGraphicalButton
    Set cBotonVida = New clsGraphicalButton
    Set cBotonEscudos = New clsGraphicalButton
    Set cBotonArmas = New clsGraphicalButton
    Set cBotonArcos = New clsGraphicalButton
    Set cBotonEspecialidad = New clsGraphicalButton
    Set cBotonPuebloOrigen = New clsGraphicalButton
    Set cBotonRaza = New clsGraphicalButton
    Set cBotonClase = New clsGraphicalButton
    Set cBotonGenero = New clsGraphicalButton
    Set cBotonVolver = New clsGraphicalButton
    Set cBotonCrear = New clsGraphicalButton
    
    Set LastPressed = New clsGraphicalButton
    
    
    Call cBotonPasswd.Initialize(imgPasswd, "", GrhPath & "BotonContraseña.jpg", _
                                    GrhPath & "BotonContraseña.jpg", Me, , , False, False)
                                    
    Call cBotonTirarDados.Initialize(imgTirarDados, "", GrhPath & "BotonTirarDados.jpg", _
                                    GrhPath & "BotonTirarDados.jpg", Me, , , False, False)
                                    
    Call cBotonMail.Initialize(imgMail, "", GrhPath & "BotonMailPj.jpg", _
                                    GrhPath & "BotonMailPj.jpg", Me, , , False, False)
                                    
    Call cBotonNombre.Initialize(imgNombre, "", GrhPath & "BotonNombrePJ.jpg", _
                                    GrhPath & "BotonNombrePJ.jpg", Me, , , False, False)
                                    
    Call cBotonConfirmPasswd.Initialize(imgConfirmPasswd, "", GrhPath & "BotonRepetirContraseña.jpg", _
                                    GrhPath & "BotonRepetirContraseña.jpg", Me, , , False, False)
                                    
    Call cBotonAtributos.Initialize(imgAtributos, "", GrhPath & "BotonAtributos.jpg", _
                                    GrhPath & "BotonAtributos.jpg", Me, , , False, False)
                                    
    Call cBotonD.Initialize(imgD, "", GrhPath & "BotonD.jpg", _
                                    GrhPath & "BotonD.jpg", Me, , , False, False)
                                    
    Call cBotonM.Initialize(imgM, "", GrhPath & "BotonM.jpg", _
                                    GrhPath & "BotonM.jpg", Me, , , False, False)
                                    
    Call cBotonF.Initialize(imgF, "", GrhPath & "BotonF.jpg", _
                                    GrhPath & "BotonF.jpg", Me, , , False, False)
                                    
    Call cBotonFuerza.Initialize(imgFuerza, "", GrhPath & "BotonFuerza.jpg", _
                                    GrhPath & "BotonFuerza.jpg", Me, , , False, False)
                                    
    Call cBotonAgilidad.Initialize(imgAgilidad, "", GrhPath & "BotonAgilidad.jpg", _
                                    GrhPath & "BotonAgilidad.jpg", Me, , , False, False)
                                    
    Call cBotonInteligencia.Initialize(imgInteligencia, "", GrhPath & "BotonInteligencia.jpg", _
                                    GrhPath & "BotonInteligencia.jpg", Me, , , False, False)
                                    
    Call cBotonCarisma.Initialize(imgCarisma, "", GrhPath & "BotonCarisma.jpg", _
                                    GrhPath & "BotonCarisma.jpg", Me, , , False, False)
                                    
    Call cBotonConstitucion.Initialize(imgConstitucion, "", GrhPath & "BotonConstitucion.jpg", _
                                    GrhPath & "BotonConstitucion.jpg", Me, , , False, False)
                                                                                                                                                                                    
    Call cBotonPuebloOrigen.Initialize(imgPuebloOrigen, "", GrhPath & "BotonPuebloOrigen.jpg", _
                                    GrhPath & "BotonPuebloOrigen.jpg", Me, , , False, False)
                                    
    Call cBotonRaza.Initialize(imgRaza, "", GrhPath & "BotonRaza.jpg", _
                                    GrhPath & "BotonRaza.jpg", Me, , , False, False)
                                                                        
    Call cBotonGenero.Initialize(imgGenero, "", GrhPath & "BotonGenero.jpg", _
                                    GrhPath & "BotonGenero.jpg", Me, , , False, False)
                                                                        
    Call cBotonVolver.Initialize(imgVolver, "", GrhPath & "BotonVolverRollover.jpg", _
                                    GrhPath & "BotonVolverClick.jpg", Me)
                                    
    Call cBotonCrear.Initialize(imgCrear, "", GrhPath & "BotonCrearPersonajeRollover.jpg", _
                                    GrhPath & "BotonCrearPersonajeClick.jpg", Me)

    Set picFullStar = LoadPicture(GrhPath & "EstrellaSimple.jpg")
    Set picHalfStar = LoadPicture(GrhPath & "EstrellaMitad.jpg")
    Set picGlowStar = LoadPicture(GrhPath & "EstrellaBrillante.jpg")

End Sub

Private Sub CargarCombos()
    Dim i As Integer
    
    
    lstHogar.Clear
    
    For i = LBound(Ciudades()) To UBound(Ciudades())
        lstHogar.AddItem Ciudades(i)
    Next i
    
    lstRaza.Clear
    
    For i = LBound(ListaRazas()) To NroRazas
        lstRaza.AddItem ListaRazas(i)
    Next i
    
End Sub

Function CheckData() As Boolean
    If txtPasswd.Text <> txtConfirmPasswd.Text Then
        MsgBox "Los passwords que tipeo no coinciden, por favor vuelva a ingresarlos."
        Exit Function
    End If
    
    If Not CheckMailString(txtMail.Text) Then
        MsgBox "Direccion de mail invalida."
        Exit Function
    End If

    If UserRaza = 0 Then
        MsgBox "Seleccione la raza del personaje."
        Exit Function
    End If
    
    If UserSexo = 0 Then
        MsgBox "Seleccione el sexo del personaje."
        Exit Function
    End If
    
    If UserHogar = 0 Then
        MsgBox "Seleccione el hogar del personaje."
        Exit Function
    End If
    
    Dim i As Integer
    For i = 1 To NUMATRIBUTOS
        If UserAtributos(i) = 0 Then
            MsgBox "Los atributos del personaje son invalidos."
            Exit Function
        End If
    Next i
    
    If Len(UserName) > 30 Then
        MsgBox ("El nombre debe tener menos de 30 letras.")
        Exit Function
    End If
    
    CheckData = True

End Function

Private Sub TirarDados()
    Call WriteThrowDices
    Call FlushBuffer
End Sub

Private Sub DirPJ_Click(Index As Integer)
    Select Case Index
        Case 0
            Dir = CheckDir(Dir + 1)
        Case 1
            Dir = CheckDir(Dir - 1)
    End Select
    
    Call UpdateHeadSelection
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClearLabel
End Sub

Private Sub HeadPJ_Click(Index As Integer)
    Select Case Index
        Case 0
            UserHead = CheckCabeza(UserHead + 1)
        Case 1
            UserHead = CheckCabeza(UserHead - 1)
    End Select
    
    Call UpdateHeadSelection
    
End Sub

Private Sub UpdateHeadSelection()
    Dim Head As Integer
    
    Head = UserHead
    Call DrawHead(Head, 2)
    
    Head = Head + 1
    Call DrawHead(CheckCabeza(Head), 3)
    
    Head = Head + 1
    Call DrawHead(CheckCabeza(Head), 4)
    
    Head = UserHead
    
    Head = Head - 1
    Call DrawHead(CheckCabeza(Head), 1)
    
    Head = Head - 1
    Call DrawHead(CheckCabeza(Head), 0)
End Sub

Public Sub ScrollUp()
'//--[ScrollUp]---------------------------//
'
'  called from the MainModule WndProc sub
'  when a up-scrolling mouse message is
'  received


Dim i As Long
If TopList < 1 Then
    TopList = 0
Else
    TopList = TopList - 1
End If

For i = 0 To lblSkill1.UBound
    lblSkill1(i).Caption = IIf(uSkills(i + 1 + TopList) > 0, SkillsNames(i + 1 + TopList) & " (" & uSkills(i + 1 + TopList) & ")", SkillsNames(i + 1 + TopList))
Next

DrawSklPt TopList + YTemp + 1

End Sub

Public Sub ScrollDown()
'//--[ScrollDown]---------------------------//
'
'  called from the MainModule WndProc sub
'  when a down-scrolling mouse message is
'  received
'


Dim i As Long
If TopList >= NUMSKILLS - lblSkill1.Count Then
    TopList = NUMSKILLS - lblSkill1.Count
Else
    TopList = TopList + 1
End If

For i = 0 To lblSkill1.UBound
    lblSkill1(i).Caption = IIf(uSkills(i + 1 + TopList) > 0, SkillsNames(i + 1 + TopList) & " (" & uSkills(i + 1 + TopList) & ")", SkillsNames(i + 1 + TopList))
Next

DrawSklPt TopList + YTemp + 1


End Sub


Private Sub Image2_Click()
Dim i As Long
If TopList < 1 Then
    TopList = 0
Else
    TopList = TopList - 1
End If

For i = 0 To lblSkill1.UBound
    lblSkill1(i).Caption = IIf(uSkills(i + 1 + TopList) > 0, SkillsNames(i + 1 + TopList) & " (" & uSkills(i + 1 + TopList) & ")", SkillsNames(i + 1 + TopList))
Next

DrawSklPt TopList + YTemp + 1
End Sub

Private Sub Image3_Click()
'Lo hago de modo que no tengan que tocar mucho código si quieren agrandar el picture
Dim i As Long
If TopList >= NUMSKILLS - lblSkill1.Count Then
    TopList = NUMSKILLS - lblSkill1.Count
Else
    TopList = TopList + 1
End If

For i = 0 To lblSkill1.UBound
    lblSkill1(i).Caption = IIf(uSkills(i + 1 + TopList) > 0, SkillsNames(i + 1 + TopList) & " (" & uSkills(i + 1 + TopList) & ")", SkillsNames(i + 1 + TopList))
Next

DrawSklPt TopList + YTemp + 1
End Sub

Private Sub imgCrear_Click()

    Dim i As Integer
    Dim CharAscii As Byte
    
    UserName = txtNombre.Text
            
    If Right$(UserName, 1) = " " Then
        UserName = RTrim$(UserName)
        MsgBox "Nombre invalido, se han removido los espacios al final del nombre"
    End If
    
    UserRaza = lstRaza.ListIndex + 1
    UserSexo = lstGenero.ListIndex + 1
    
    For i = 1 To NUMATRIBUTES
        UserAtributos(i) = Val(lblAtributos(i).Caption)
    Next i
    
    UserHogar = lstHogar.ListIndex + 1
    
    If Not CheckData Then Exit Sub

    UserPassword = txtPasswd.Text
    
    For i = 1 To Len(UserPassword)
        CharAscii = Asc(mid$(UserPassword, i, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Password inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Sub
        End If
    Next i
    
    UserEmail = txtMail.Text
    
    If SkillPts <> 0 Then
        MsgBox "Debes asignar todos los puntos disponibles en los skills."
        Exit Sub
    End If
    
    Call CopyMemory(UserSkills(1), uSkills(1), NUMSKILLS)
    
    frmMain.Socket1.HostName = CurServerIP
    frmMain.Socket1.RemotePort = CurServerPort

    EstadoLogin = E_MODO.CrearNuevoPj
    
    If Not frmMain.Socket1.Connected Then

        MsgBox "Error: Se ha perdido la conexion con el server."
        Unload Me
        
    Else
        
        Call Login
        
    End If
    
    bShowTutorial = True
End Sub

Private Sub imgDados_Click()
    Call Audio.PlayWave(SND_DICE)
            Call TirarDados
End Sub

Private Sub imgEspecialidad_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieEspecialidad)
End Sub

Private Sub imgNombre_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieNombre)
End Sub

Private Sub imgPasswd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.iePasswd)
End Sub

Private Sub imgConfirmPasswd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieConfirmPasswd)
End Sub

Private Sub imgAtributos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieAtributos)
End Sub

Private Sub imgD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieD)
End Sub

Private Sub imgM_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieM)
End Sub

Private Sub imgF_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieF)
End Sub

Private Sub imgFuerza_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieFuerza)
End Sub

Private Sub imgAgilidad_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieAgilidad)
End Sub

Private Sub imgInteligencia_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieInteligencia)
End Sub

Private Sub imgCarisma_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieCarisma)
End Sub

Private Sub imgConstitucion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieConstitucion)
End Sub

Private Sub imgArcos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieArcos)
End Sub

Private Sub imgArmas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieArmas)
End Sub

Private Sub imgEscudos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieEscudos)
End Sub

Private Sub imgEvasion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieEvasion)
End Sub

Private Sub imgMagia_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieMagia)
End Sub

Private Sub imgMail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieMail)
End Sub

Private Sub imgVida_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieVida)
End Sub

Private Sub imgTirarDados_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieTirarDados)
End Sub

Private Sub imgPuebloOrigen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.iePuebloOrigen)
End Sub

Private Sub imgRaza_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieRaza)
End Sub

Private Sub imgClase_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieClase)
End Sub

Private Sub imgGenero_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieGenero)
End Sub

Private Sub imgVolver_Click()
    Call Audio.PlayMIDI("2.mid")
    
    bShowTutorial = False
    
    Unload Me
End Sub

Private Sub lblSkill1_Click(Index As Integer)
Call picSkills_Click
End Sub

Private Sub lblSkill1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseButton = Button
End Sub

Private Sub lstGenero_Click()
    UserSexo = lstGenero.ListIndex + 1
    Call DarCuerpoYCabeza
End Sub

Private Sub lstRaza_Click()
    UserRaza = lstRaza.ListIndex + 1
    Call DarCuerpoYCabeza
    
    Call UpdateStats
End Sub

Private Sub picHead_Click(Index As Integer)
    ' No se mueve si clickea al medio
    If Index = 2 Then Exit Sub
    
    Dim Counter As Integer
    Dim Head As Integer
    
    Head = UserHead
    
    If Index > 2 Then
        For Counter = Index - 2 To 1 Step -1
            Head = CheckCabeza(Head + 1)
        Next Counter
    Else
        For Counter = 2 - Index To 1 Step -1
            Head = CheckCabeza(Head - 1)
        Next Counter
    End If
    
    UserHead = Head
    
    Call UpdateHeadSelection
    
End Sub

Private Sub picSkills_Click()

If YTemp < 0 Or YTemp > lblSkill1.UBound Then Exit Sub

If MouseButton = vbLeftButton Then
    
    If SkillPts = 0 Then Exit Sub
    
    If MouseShift And 2 Then
        If SkillPts >= 3 Then
            uSkills(TopList + YTemp + 1) = uSkills(TopList + YTemp + 1) + 3
            SkillPts = SkillPts - 3
        ElseIf SkillPts >= 2 Then
            uSkills(TopList + YTemp + 1) = uSkills(TopList + YTemp + 1) + 2
            SkillPts = SkillPts - 2
        Else
            uSkills(TopList + YTemp + 1) = uSkills(TopList + YTemp + 1) + 1
            SkillPts = SkillPts - 1
        End If
    ElseIf MouseShift And 3 Then
        uSkills(TopList + YTemp + 1) = uSkills(TopList + YTemp + 1) + SkillPts
        SkillPts = 0
    Else
        uSkills(TopList + YTemp + 1) = uSkills(TopList + YTemp + 1) + 1
        SkillPts = SkillPts - 1
    End If
ElseIf MouseButton = vbRightButton Then
    If SkillPts = 10 Then Exit Sub
    
    If MouseShift And 2 Then
        If uSkills(TopList + YTemp + 1) >= 3 Then
            uSkills(TopList + YTemp + 1) = uSkills(TopList + YTemp + 1) - 3
            SkillPts = SkillPts + 3
        ElseIf uSkills(TopList + YTemp + 1) >= 2 Then
            uSkills(TopList + YTemp + 1) = uSkills(TopList + YTemp + 1) - 2
            SkillPts = SkillPts + 2
        Else
            uSkills(TopList + YTemp + 1) = uSkills(TopList + YTemp + 1) - 1
            SkillPts = SkillPts + 1
        End If
    ElseIf MouseShift And 3 Then
        SkillPts = SkillPts + uSkills(TopList + YTemp + 1)
        uSkills(TopList + YTemp + 1) = uSkills(TopList + YTemp + 1) - SkillPts
    Else
        uSkills(TopList + YTemp + 1) = uSkills(TopList + YTemp + 1) - 1
        SkillPts = SkillPts + 1
    End If
End If

lblPuntosRestantes.Caption = "Puntos restantes: " & SkillPts
If uSkills(TopList + YTemp + 1) > 0 Then
    lblSkill1(YTemp).Caption = SkillsNames(TopList + YTemp + 1) & " (" & uSkills(TopList + YTemp + 1) & ")"
End If
DrawSklPt TopList + YTemp + 1

End Sub

Private Sub picSkills_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseButton = Button
MouseShift = Shift
End Sub

Private Sub picSkills_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

lblHelp.Caption = "Asigna puntos a las habilidades que creas conveniente." & vbNewLine & _
                "Click izq suma, Click derecho resta" & vbNewLine & _
                "Presiona Ctrl para asignar de a 3 puntos" & vbNewLine & _
                "Presiona Shift para asignar todo." & vbNewLine & _
                "Puedes usar la rueda del mouse para bajar o subir"
                
YTemp = (Y - 1) \ 20

Dim i As Long

For i = 0 To lblSkill1.UBound
    If YTemp = i Then
        lblSkill1(i).ForeColor = vbRed
    Else
        lblSkill1(i).ForeColor = vbWhite
    End If
Next

End Sub

Private Sub tAnimacion_Timer()
    'Dim SR As RECT
    'Dim DR As RECT
    Dim Grh As Long
    Static Frame As Byte
    
    If currentGrh = 0 Then Exit Sub
    UserHead = CheckCabeza(UserHead)
    
    Frame = Frame + 1
    If Frame >= GrhData(currentGrh).NumFrames Then Frame = 1
    'Call DrawImageInPicture(picPJ, Me.Picture, 0, 0, , , picPJ.Left, picPJ.Top)
    
    Grh = GrhData(currentGrh).Frames(Frame)
    
    With GrhData(Grh)
        'SR.Left = .sX
        'SR.Top = .sY
        'SR.Right = SR.Left + .pixelWidth
        'SR.bottom = SR.Top + .pixelHeight
        
        'DR.Left = (picPJ.Width - .pixelWidth) \ 2 - 2
        'DR.Top = (picPJ.Height - .pixelHeight) \ 2 - 2
        'DR.Right = DR.Left + .pixelWidth
        'DR.bottom = DR.Top + .pixelHeight
        
        picTemp.BackColor = picTemp.BackColor
        
        'Call DrawGrhtoHdc(picTemp.hdc, Grh, SR, DR)
        'Call DrawTransparentGrhtoHdc(picPJ.hdc, picTemp.hdc, DR, DR, vbBlack)
    End With
    
    Grh = HeadData(UserHead).Head(Dir).GrhIndex
    
    With GrhData(Grh)
        'SR.Left = .sX
        'SR.Top = .sY
        'SR.Right = SR.Left + .pixelWidth
        'SR.bottom = SR.Top + .pixelHeight
        
        'DR.Left = (picPJ.Width - .pixelWidth) \ 2 - 2
        'DR.Top = DR.bottom + BodyData(UserBody).HeadOffset.Y - .pixelHeight
        'DR.Right = DR.Left + .pixelWidth
        'DR.bottom = DR.Top + .pixelHeight
        
        picTemp.BackColor = picTemp.BackColor
        
        'Call DrawGrhtoHdc(picTemp.hdc, Grh, SR, DR)
        'Call DrawTransparentGrhtoHdc(picPJ.hdc, picTemp.hdc, DR, DR, vbBlack)
    End With
End Sub

Private Sub DrawSklPt(ByVal Skill As Integer)
    'Dim SR As RECT
    'Dim DR As RECT
    'Dim Grh As Long
    
    'Call DrawImageInPicture(picSkills, Me.Picture, 0, 0, , , picSkills.Left, picSkills.Top)
    
    Dim i As Long
    
        With GrhData(14622)
        
            'SR.Left = .sX
            'SR.Top = .sY
            'SR.Right = SR.Left + .pixelWidth
            'SR.bottom = SR.Top + 20
            
            For i = 0 To lblSkill1.UBound
                'DR.Left = 0
                'DR.Top = I * 20
                'DR.Right = uSkills(I + 1 + TopList) * 17
                'DR.bottom = DR.Top + 20 '.pixelHeight
                
                picSkillsTemp.BackColor = picSkillsTemp.BackColor
                
                'Call DrawGrhtoHdc(picSkillsTemp.hdc, 14622, SR, DR)
                'Call DrawTransparentGrhtoHdc(picSkills.hdc, picSkillsTemp.hdc, DR, DR, vbBlack)
            Next
            
        End With
End Sub

Private Sub DrawHead(ByVal Head As Integer, ByVal PicIndex As Integer)

   'Dim SR As RECT
    'Dim DR As RECT
    Dim Grh As Long

    'Call DrawImageInPicture(picHead(PicIndex), Me.Picture, 0, 0, , , picHead(PicIndex).Left, picHead(PicIndex).Top)
    
    Grh = HeadData(Head).Head(Dir).GrhIndex

    With GrhData(Grh)
        'SR.Left = .sX
        'SR.Top = .sY
        'SR.Right = SR.Left + .pixelWidth
        'SR.bottom = SR.Top + .pixelHeight
        
        'DR.Left = (picHead(0).Width - .pixelWidth) \ 2 + 1
        'DR.Top = 0
        'DR.Right = DR.Left + .pixelWidth
        'DR.bottom = DR.Top + .pixelHeight
        
        picTemp.BackColor = picTemp.BackColor
        
        'Call DrawGrhtoHdc(picTemp.hdc, Grh, SR, DR)
       ' Call DrawTransparentGrhtoHdc(picHead(PicIndex).hdc, picTemp.hdc, DR, DR, vbBlack)
    End With
    
End Sub

Private Sub txtConfirmPasswd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieConfirmPasswd)
End Sub

Private Sub txtMail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieMail)
End Sub

Private Sub txtNombre_Change()
    txtNombre.Text = LTrim(txtNombre.Text)
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub DarCuerpoYCabeza()

    Dim bVisible As Boolean
    Dim PicIndex As Integer
    Dim LineIndex As Integer
    
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
                    
                Case Else
                    UserHead = 0
                    UserBody = 0
            End Select
        Case Else
            UserHead = 0
            UserBody = 0
    End Select
    
    bVisible = UserHead <> 0 And UserBody <> 0
    
    HeadPJ(0).Visible = bVisible
    HeadPJ(1).Visible = bVisible
    DirPJ(0).Visible = bVisible
    DirPJ(1).Visible = bVisible
    
    For PicIndex = 0 To 4
        picHead(PicIndex).Visible = bVisible
    Next PicIndex
    
    For LineIndex = 0 To 3
        Line1(LineIndex).Visible = bVisible
    Next LineIndex
    
    If bVisible Then Call UpdateHeadSelection
    
    currentGrh = BodyData(UserBody).Walk(Dir).GrhIndex
    If currentGrh > 0 Then _
        tAnimacion.Interval = 56 'Round(GrhData(currentGrh).Speed / GrhData(currentGrh).NumFrames) 'el speed de fenix es una garcha
End Sub

Private Function CheckCabeza(ByVal Head As Integer) As Integer

Select Case UserSexo
    Case eGenero.Hombre
        Select Case UserRaza
            Case eRaza.Humano
                If Head > HUMANO_H_ULTIMA_CABEZA Then
                    CheckCabeza = HUMANO_H_PRIMER_CABEZA + (Head - HUMANO_H_ULTIMA_CABEZA) - 1
                ElseIf Head < HUMANO_H_PRIMER_CABEZA Then
                    CheckCabeza = HUMANO_H_ULTIMA_CABEZA - (HUMANO_H_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case eRaza.Elfo
                If Head > ELFO_H_ULTIMA_CABEZA Then
                    CheckCabeza = ELFO_H_PRIMER_CABEZA + (Head - ELFO_H_ULTIMA_CABEZA) - 1
                ElseIf Head < ELFO_H_PRIMER_CABEZA Then
                    CheckCabeza = ELFO_H_ULTIMA_CABEZA - (ELFO_H_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case eRaza.ElfoOscuro
                If Head > DROW_H_ULTIMA_CABEZA Then
                    CheckCabeza = DROW_H_PRIMER_CABEZA + (Head - DROW_H_ULTIMA_CABEZA) - 1
                ElseIf Head < DROW_H_PRIMER_CABEZA Then
                    CheckCabeza = DROW_H_ULTIMA_CABEZA - (DROW_H_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case eRaza.Enano
                If Head > ENANO_H_ULTIMA_CABEZA Then
                    CheckCabeza = ENANO_H_PRIMER_CABEZA + (Head - ENANO_H_ULTIMA_CABEZA) - 1
                ElseIf Head < ENANO_H_PRIMER_CABEZA Then
                    CheckCabeza = ENANO_H_ULTIMA_CABEZA - (ENANO_H_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case eRaza.Gnomo
                If Head > GNOMO_H_ULTIMA_CABEZA Then
                    CheckCabeza = GNOMO_H_PRIMER_CABEZA + (Head - GNOMO_H_ULTIMA_CABEZA) - 1
                ElseIf Head < GNOMO_H_PRIMER_CABEZA Then
                    CheckCabeza = GNOMO_H_ULTIMA_CABEZA - (GNOMO_H_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case Else
                UserRaza = lstRaza.ListIndex + 1
                CheckCabeza = CheckCabeza(Head)
        End Select
        
    Case eGenero.Mujer
        Select Case UserRaza
            Case eRaza.Humano
                If Head > HUMANO_M_ULTIMA_CABEZA Then
                    CheckCabeza = HUMANO_M_PRIMER_CABEZA + (Head - HUMANO_M_ULTIMA_CABEZA) - 1
                ElseIf Head < HUMANO_M_PRIMER_CABEZA Then
                    CheckCabeza = HUMANO_M_ULTIMA_CABEZA - (HUMANO_M_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case eRaza.Elfo
                If Head > ELFO_M_ULTIMA_CABEZA Then
                    CheckCabeza = ELFO_M_PRIMER_CABEZA + (Head - ELFO_M_ULTIMA_CABEZA) - 1
                ElseIf Head < ELFO_M_PRIMER_CABEZA Then
                    CheckCabeza = ELFO_M_ULTIMA_CABEZA - (ELFO_M_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case eRaza.ElfoOscuro
                If Head > DROW_M_ULTIMA_CABEZA Then
                    CheckCabeza = DROW_M_PRIMER_CABEZA + (Head - DROW_M_ULTIMA_CABEZA) - 1
                ElseIf Head < DROW_M_PRIMER_CABEZA Then
                    CheckCabeza = DROW_M_ULTIMA_CABEZA - (DROW_M_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case eRaza.Enano
                If Head > ENANO_M_ULTIMA_CABEZA Then
                    CheckCabeza = ENANO_M_PRIMER_CABEZA + (Head - ENANO_M_ULTIMA_CABEZA) - 1
                ElseIf Head < ENANO_M_PRIMER_CABEZA Then
                    CheckCabeza = ENANO_M_ULTIMA_CABEZA - (ENANO_M_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case eRaza.Gnomo
                If Head > GNOMO_M_ULTIMA_CABEZA Then
                    CheckCabeza = GNOMO_M_PRIMER_CABEZA + (Head - GNOMO_M_ULTIMA_CABEZA) - 1
                ElseIf Head < GNOMO_M_PRIMER_CABEZA Then
                    CheckCabeza = GNOMO_M_ULTIMA_CABEZA - (GNOMO_M_PRIMER_CABEZA - Head) + 1
                Else
                    CheckCabeza = Head
                End If
                
            Case Else
                UserRaza = lstRaza.ListIndex + 1
                CheckCabeza = CheckCabeza(Head)
        End Select
    Case Else
        UserSexo = lstGenero.ListIndex + 1
        CheckCabeza = CheckCabeza(Head)
End Select
End Function

Private Function CheckDir(ByRef Dir As E_Heading) As E_Heading

    If Dir > E_Heading.WEST Then Dir = E_Heading.NORTH
    If Dir < E_Heading.NORTH Then Dir = E_Heading.WEST
    
    CheckDir = Dir
    
    currentGrh = BodyData(UserBody).Walk(Dir).GrhIndex
    If currentGrh > 0 Then _
        tAnimacion.Interval = Round(GrhData(currentGrh).Speed / GrhData(currentGrh).NumFrames)

End Function

Private Sub LoadHelp()
    vHelp(eHelp.iePasswd) = "La contraseña que utilizarás para conectar tu personaje al juego."
    vHelp(eHelp.ieTirarDados) = "Presionando sobre la Esfera Roja, se modificarán al azar los atributos de tu personaje, de esta manera puedes elegir los que más te parezcan para definir a tu personaje."
    vHelp(eHelp.ieMail) = "Es sumamente importante que ingreses una dirección de correo electrónico válida, ya que en el caso de perder la contraseña de tu personaje, se te enviará cuando lo requieras, a esa dirección."
    vHelp(eHelp.ieNombre) = "Sé cuidadoso al seleccionar el nombre de tu personaje. Argentum es un juego de rol, un mundo mágico y fantástico, y si seleccionás un nombre obsceno o con connotación política, los administradores borrarán tu personaje y no habrá ninguna posibilidad de recuperarlo."
    vHelp(eHelp.ieConfirmPasswd) = "La contraseña que utilizarás para conectar tu personaje al juego."
    vHelp(eHelp.ieAtributos) = "Son las cualidades que definen tu personaje. Generalmente se los llama ""Dados"". (Ver Tirar Dados)"
    vHelp(eHelp.ieD) = "Son los atributos que obtuviste al azar. Presioná la esfera roja para volver a tirarlos."
    vHelp(eHelp.ieM) = "Son los modificadores por raza que influyen en los atributos de tu personaje."
    vHelp(eHelp.ieF) = "Los atributos finales de tu personaje, de acuerdo a la raza que elegiste."
    vHelp(eHelp.ieFuerza) = "De ella dependerá qué tan potentes serán tus golpes, tanto con armas de cuerpo a cuerpo, a distancia o sin armas."
    vHelp(eHelp.ieAgilidad) = "Este atributo intervendrá en qué tan bueno seas, tanto evadiendo como acertando golpes, respecto de otros personajes como de las criaturas a las q te enfrentes."
    vHelp(eHelp.ieInteligencia) = "Influirá de manera directa en cuánto maná ganarás por nivel."
    vHelp(eHelp.ieCarisma) = "Será necesario tanto para la relación con otros personajes (entrenamiento en parties) como con las criaturas (domar animales)."
    vHelp(eHelp.ieConstitucion) = "Afectará a la cantidad de vida que podrás ganar por nivel."
    vHelp(eHelp.ieEvasion) = "Evalúa la habilidad esquivando ataques físicos."
    vHelp(eHelp.ieMagia) = "Puntúa la cantidad de maná que se tendrá."
    vHelp(eHelp.ieVida) = "Valora la cantidad de salud que se podrá llegar a tener."
    vHelp(eHelp.ieEscudos) = "Estima la habilidad para rechazar golpes con escudos."
    vHelp(eHelp.ieArmas) = "Evalúa la habilidad en el combate cuerpo a cuerpo con armas."
    vHelp(eHelp.ieArcos) = "Evalúa la habilidad en el combate a distancia con arcos. "
    vHelp(eHelp.ieEspecialidad) = ""
    vHelp(eHelp.iePuebloOrigen) = "Define el hogar de tu personaje. Sin embargo, el personaje nacerá en Nemahuak, la ciudad de los novatos."
    vHelp(eHelp.ieRaza) = "De la raza que elijas dependerá cómo se modifiquen los dados que saques. Podés cambiar de raza para poder visualizar cómo se modifican los distintos atributos."
    vHelp(eHelp.ieClase) = "La clase influirá en las características principales que tenga tu personaje, asi como en las magias e items que podrá utilizar. Las estrellas que ves abajo te mostrarán en qué habilidades se destaca la misma."
    vHelp(eHelp.ieGenero) = "Indica si el personaje será masculino o femenino. Esto influye en los items que podrá equipar."
End Sub

Private Sub ClearLabel()
    LastPressed.ToggleToNormal
    lblHelp = ""
End Sub

Private Sub txtNombre_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.ieNombre)
End Sub

Private Sub txtPasswd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = vHelp(eHelp.iePasswd)
End Sub

Public Sub UpdateStats()
    
    Call UpdateRazaMod
End Sub

Private Sub UpdateRazaMod()
    Dim SelRaza As Integer
    Dim i As Integer
    
    
    If lstRaza.ListIndex > -1 Then
    
        SelRaza = lstRaza.ListIndex + 1
        
        With ModRaza(SelRaza)
            lblModRaza(eAtributos.Fuerza).Caption = IIf(.Fuerza >= 0, "+", "") & .Fuerza
            lblModRaza(eAtributos.Agilidad).Caption = IIf(.Agilidad >= 0, "+", "") & .Agilidad
            lblModRaza(eAtributos.Inteligencia).Caption = IIf(.Inteligencia >= 0, "+", "") & .Inteligencia
            lblModRaza(eAtributos.Carisma).Caption = IIf(.Carisma >= 0, "+", "") & .Carisma
            lblModRaza(eAtributos.Constitucion).Caption = IIf(.Constitucion >= 0, "+", "") & .Constitucion
        End With
    End If
    
    ' Atributo total
    For i = 1 To NUMATRIBUTES
        lblAtributoFinal(i).Caption = Val(lblAtributos(i).Caption) + Val(lblModRaza(i))
    Next i
    
End Sub

Private Sub LoadCharInfo()
    Dim SearchVar As String
    Dim i As Integer
    
    NroRazas = UBound(ListaRazas())
    'NroClases = UBound(ListaClases())

    ReDim ModRaza(1 To NroRazas)
    'ReDim ModClase(1 To NroClases)
    
    'Modificadores de Clase
    'For i = 1 To NroClases
    '    With ModClase(i)
    '        SearchVar = ListaClases(i)
    '
    '        .Evasion = Val(GetVar(IniPath & "CharInfo.dat", "MODEVASION", SearchVar))
    '        .AtaqueArmas = Val(GetVar(IniPath & "CharInfo.dat", "MODATAQUEARMAS", SearchVar))
    '        .AtaqueProyectiles = Val(GetVar(IniPath & "CharInfo.dat", "MODATAQUEPROYECTILES", SearchVar))
    '        .DañoArmas = Val(GetVar(IniPath & "CharInfo.dat", "MODDAÑOARMAS", SearchVar))
    '        .DañoProyectiles = Val(GetVar(IniPath & "CharInfo.dat", "MODDAÑOPROYECTILES", SearchVar))
    '        .Escudo = Val(GetVar(IniPath & "CharInfo.dat", "MODESCUDO", SearchVar))
    '        .Hit = Val(GetVar(IniPath & "CharInfo.dat", "HIT", SearchVar))
    '        .Magia = Val(GetVar(IniPath & "CharInfo.dat", "MODMAGIA", SearchVar))
    '        .Vida = Val(GetVar(IniPath & "CharInfo.dat", "MODVIDA", SearchVar))
    '    End With
    'Next i
    
    'Modificadores de Raza


End Sub


