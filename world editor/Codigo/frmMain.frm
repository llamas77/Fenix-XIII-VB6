VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   10635
   ClientLeft      =   405
   ClientTop       =   -1530
   ClientWidth     =   15975
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   709
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1065
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin FXIIIWorldEditor.lvButtons_H CBVerCapa2 
      Height          =   495
      Left            =   15360
      TabIndex        =   143
      Top             =   3600
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   "2"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Frame FrameDespMosaico 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Desp. de Mosaico:"
      Height          =   1095
      Left            =   13200
      TabIndex        =   128
      Top             =   45
      Width           =   2295
      Begin VB.CheckBox CheckDespMosaico 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Activado"
         Height          =   255
         Left            =   120
         TabIndex        =   133
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CommandHorMenos 
         BackColor       =   &H00E0E0E0&
         Caption         =   "-"
         Height          =   255
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   132
         TabStop         =   0   'False
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton CommandHorMas 
         BackColor       =   &H00E0E0E0&
         Caption         =   "+"
         Height          =   255
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   131
         TabStop         =   0   'False
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton CommandVerMas 
         BackColor       =   &H00E0E0E0&
         Caption         =   "+"
         Height          =   255
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   130
         TabStop         =   0   'False
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton CommandVerMenos 
         BackColor       =   &H00E0E0E0&
         Caption         =   "-"
         Height          =   255
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   129
         TabStop         =   0   'False
         Top             =   720
         Width           =   375
      End
      Begin VB.Label LabelHorizontal 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Hori.:"
         Height          =   195
         Left            =   120
         TabIndex        =   142
         Top             =   480
         Width           =   375
      End
      Begin VB.Label LabelVertical 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Vert.:"
         Height          =   195
         Left            =   120
         TabIndex        =   141
         Top             =   720
         Width           =   375
      End
      Begin VB.Label TextVertical 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   140
         Top             =   720
         Width           =   495
      End
      Begin VB.Label TextHorizontal 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   139
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label99 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "(NumPad5)"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   138
         Top             =   260
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "(4)"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   137
         Top             =   510
         Width           =   255
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "(6)"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   136
         Top             =   510
         Width           =   255
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "(8)"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   135
         Top             =   750
         Width           =   255
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "(2)"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   134
         Top             =   750
         Width           =   255
      End
   End
   Begin FXIIIWorldEditor.lvButtons_H CBVerObjetos 
      Height          =   495
      Left            =   15360
      TabIndex        =   127
      Top             =   7920
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
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
      Mode            =   1
      Value           =   0   'False
      Image           =   "frmMain.frx":78BA
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin FXIIIWorldEditor.lvButtons_H CBVerNpcs 
      Height          =   495
      Left            =   15360
      TabIndex        =   126
      Top             =   7440
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
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
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "frmMain.frx":850E
      ImgSize         =   40
      cBack           =   -2147483633
   End
   Begin FXIIIWorldEditor.lvButtons_H CBVerCuadricula 
      Height          =   495
      Left            =   15360
      TabIndex        =   121
      TabStop         =   0   'False
      Top             =   5040
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
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
      cBhover         =   -2147483633
      LockHover       =   1
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "frmMain.frx":9162
      ImgSize         =   48
      cBack           =   14737632
   End
   Begin VB.PictureBox picRadar 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   1590
      Left            =   120
      ScaleHeight     =   106
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   107
      TabIndex        =   96
      Top             =   120
      Width           =   1605
      Begin VB.Label FPS 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FPS: ??"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   150
         Left            =   1065
         TabIndex        =   98
         Top             =   1455
         Width           =   450
      End
      Begin VB.Label POSX 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X: ?? - Y:??"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   150
         Left            =   120
         TabIndex        =   97
         Top             =   1455
         Width           =   675
      End
      Begin VB.Shape ApuntadorRadar 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H000000FF&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   600
         Top             =   600
         Width           =   330
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         Height          =   1365
         Left            =   120
         Top             =   120
         Width           =   1365
      End
   End
   Begin FXIIIWorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   5
      Left            =   11325
      TabIndex        =   47
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1826
      Caption         =   "Tri&gger's (F12)"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   -2147483633
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":92D0
      ImgSize         =   24
      cBack           =   12252311
   End
   Begin FXIIIWorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   4
      Left            =   9810
      TabIndex        =   46
      Top             =   120
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   1826
      Caption         =   "&Objetos (F11)"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   -2147483633
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      Image           =   "frmMain.frx":9896
      ImgSize         =   24
      cBack           =   12252311
   End
   Begin FXIIIWorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   3
      Left            =   8445
      TabIndex        =   45
      Top             =   120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1826
      Caption         =   "&NPC's (F8)"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   -2147483633
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":9D97
      ImgSize         =   24
      cBack           =   12252311
   End
   Begin FXIIIWorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   2
      Left            =   6945
      TabIndex        =   44
      Top             =   120
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   1826
      Caption         =   "&Bloqueos (F7)"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   -2147483633
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":A14B
      ImgSize         =   24
      cBack           =   12252311
   End
   Begin FXIIIWorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   1
      Left            =   5445
      TabIndex        =   43
      Top             =   120
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   1826
      Caption         =   "&Traslados (F6)"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   -2147483633
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   5
      Image           =   "frmMain.frx":A4CC
      ImgSize         =   24
      cBack           =   12252311
   End
   Begin FXIIIWorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   0
      Left            =   4620
      TabIndex        =   42
      Top             =   120
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   1826
      Caption         =   "&Superficie (F5)"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   0
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   5
      Image           =   "frmMain.frx":DB2C
      ImgSize         =   24
      cBack           =   12252311
   End
   Begin FXIIIWorldEditor.lvButtons_H cmdQuitarFunciones 
      Height          =   315
      Left            =   1800
      TabIndex        =   41
      ToolTipText     =   "Quitar Todas las Funciones Activadas"
      Top             =   1410
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   556
      Caption         =   "&Quitar Funciones (F4)"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   0
      cBhover         =   255
      LockHover       =   1
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   8421631
   End
   Begin VB.Timer TimAutoGuardarMapa 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   3960
      Top             =   1920
   End
   Begin VB.TextBox StatTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3885
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "frmMain.frx":11072
      Top             =   6270
      Width           =   4350
   End
   Begin VB.PictureBox pPaneles 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4395
      Left            =   120
      Picture         =   "frmMain.frx":110B2
      ScaleHeight     =   4365
      ScaleWidth      =   4365
      TabIndex        =   8
      Top             =   1800
      Width           =   4395
      Begin VB.TextBox tTY 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Left            =   1200
         TabIndex        =   86
         Text            =   "1"
         Top             =   960
         Visible         =   0   'False
         Width           =   2900
      End
      Begin VB.TextBox tTX 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Left            =   1200
         TabIndex        =   85
         Text            =   "1"
         Top             =   600
         Visible         =   0   'False
         Width           =   2900
      End
      Begin VB.TextBox tTMapa 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Left            =   1200
         TabIndex        =   84
         Text            =   "1"
         Top             =   240
         Visible         =   0   'False
         Width           =   2900
      End
      Begin FXIIIWorldEditor.lvButtons_H cInsertarTrans 
         Height          =   375
         Left            =   240
         TabIndex        =   87
         Top             =   1320
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "&Insertar Traslado"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin FXIIIWorldEditor.lvButtons_H cInsertarTransOBJ 
         Height          =   375
         Left            =   240
         TabIndex        =   88
         Top             =   1680
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "Colocar automaticamente &Objeto"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin FXIIIWorldEditor.lvButtons_H cUnionManual 
         Height          =   375
         Left            =   240
         TabIndex        =   89
         Top             =   2160
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "&Union con Mapa Adyacente (manual)"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin FXIIIWorldEditor.lvButtons_H cUnionAuto 
         Height          =   375
         Left            =   240
         TabIndex        =   90
         Top             =   2520
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "Union con Mapas &Adyacentes (auto)"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin FXIIIWorldEditor.lvButtons_H cQuitarTrans 
         Height          =   375
         Left            =   240
         TabIndex        =   91
         Top             =   3000
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "&Quitar Traslados"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.ComboBox cCapas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         ItemData        =   "frmMain.frx":23716
         Left            =   1080
         List            =   "frmMain.frx":23726
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cGrh 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Left            =   2880
         TabIndex        =   73
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cFiltro 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   0
         Left            =   600
         TabIndex        =   72
         Top             =   2760
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2580
         Index           =   0
         ItemData        =   "frmMain.frx":23736
         Left            =   120
         List            =   "frmMain.frx":23738
         Sorted          =   -1  'True
         TabIndex        =   71
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin FXIIIWorldEditor.lvButtons_H cQuitarEnTodasLasCapas 
         Height          =   375
         Left            =   120
         TabIndex        =   74
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "Quitar en &Capas 2 y 3"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin FXIIIWorldEditor.lvButtons_H cQuitarEnEstaCapa 
         Height          =   375
         Left            =   120
         TabIndex        =   75
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar en esta Capa"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin FXIIIWorldEditor.lvButtons_H cSeleccionarSuperficie 
         Height          =   735
         Left            =   2400
         TabIndex        =   76
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Caption         =   "&Insertar Superficie"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.ComboBox cNumFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   2
         ItemData        =   "frmMain.frx":2373A
         Left            =   3360
         List            =   "frmMain.frx":2373C
         TabIndex        =   67
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cCantFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   2
         ItemData        =   "frmMain.frx":2373E
         Left            =   840
         List            =   "frmMain.frx":23740
         TabIndex        =   0
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2580
         Index           =   3
         ItemData        =   "frmMain.frx":23742
         Left            =   120
         List            =   "frmMain.frx":23744
         TabIndex        =   66
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.ComboBox cFiltro 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   3
         Left            =   600
         TabIndex        =   65
         Top             =   2760
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ComboBox cCantFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   0
         ItemData        =   "frmMain.frx":23746
         Left            =   840
         List            =   "frmMain.frx":23748
         TabIndex        =   58
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cNumFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   0
         ItemData        =   "frmMain.frx":2374A
         Left            =   3360
         List            =   "frmMain.frx":2374C
         TabIndex        =   57
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cFiltro 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   1
         Left            =   600
         TabIndex        =   56
         Top             =   2760
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2580
         Index           =   1
         ItemData        =   "frmMain.frx":2374E
         Left            =   120
         List            =   "frmMain.frx":23750
         TabIndex        =   55
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   3210
         Index           =   4
         ItemData        =   "frmMain.frx":23752
         Left            =   120
         List            =   "frmMain.frx":23754
         TabIndex        =   54
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.PictureBox Picture5 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   10
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture6 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   11
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture7 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   12
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture8 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   13
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture9 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   14
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture11 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   50
         Top             =   0
         Width           =   0
      End
      Begin FXIIIWorldEditor.lvButtons_H cQuitarTrigger 
         Height          =   735
         Left            =   120
         TabIndex        =   51
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1296
         Caption         =   "&Quitar Trigger's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin FXIIIWorldEditor.lvButtons_H cVerTriggers 
         Height          =   375
         Left            =   120
         TabIndex        =   52
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Mostrar Trigger's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin FXIIIWorldEditor.lvButtons_H cInsertarTrigger 
         Height          =   735
         Left            =   2400
         TabIndex        =   53
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Caption         =   "&Insertar Trigger"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin FXIIIWorldEditor.lvButtons_H cAgregarFuncalAzar 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   59
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "Insetar NPC's al &Azar"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin FXIIIWorldEditor.lvButtons_H cQuitarFunc 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   60
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar NPC's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin FXIIIWorldEditor.lvButtons_H cInsertarFunc 
         Height          =   735
         Index           =   0
         Left            =   2400
         TabIndex        =   61
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Caption         =   "&Insertar NPC's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin FXIIIWorldEditor.lvButtons_H cVerBloqueos 
         Height          =   495
         Left            =   120
         TabIndex        =   62
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   873
         Caption         =   "&Mostrar Bloqueos"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin FXIIIWorldEditor.lvButtons_H cInsertarBloqueo 
         Height          =   735
         Left            =   120
         TabIndex        =   63
         Top             =   720
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   1296
         Caption         =   "&Insertar Bloqueos"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin FXIIIWorldEditor.lvButtons_H cQuitarBloqueo 
         Height          =   735
         Left            =   120
         TabIndex        =   64
         Top             =   1560
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   1296
         Caption         =   "&Quitar Bloqueos"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin FXIIIWorldEditor.lvButtons_H cAgregarFuncalAzar 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   68
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "Insetar OBJ's al &Azar"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin FXIIIWorldEditor.lvButtons_H cQuitarFunc 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   69
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar OBJ's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin FXIIIWorldEditor.lvButtons_H cInsertarFunc 
         Height          =   735
         Index           =   2
         Left            =   2400
         TabIndex        =   70
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Caption         =   "&Insertar Objetos"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin FXIIIWorldEditor.lvButtons_H cInsertarFunc 
         Height          =   735
         Index           =   1
         Left            =   2400
         TabIndex        =   83
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Caption         =   "&Insertar NPC's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin FXIIIWorldEditor.lvButtons_H cQuitarFunc 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   82
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar NPC's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin FXIIIWorldEditor.lvButtons_H cAgregarFuncalAzar 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   81
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "Insetar NPC's al &Azar"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.ComboBox cCantFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   1
         ItemData        =   "frmMain.frx":23756
         Left            =   840
         List            =   "frmMain.frx":23758
         TabIndex        =   77
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cFiltro 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   2
         Left            =   600
         TabIndex        =   78
         Top             =   2760
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2580
         Index           =   2
         ItemData        =   "frmMain.frx":2375A
         Left            =   120
         List            =   "frmMain.frx":2375C
         TabIndex        =   79
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.ComboBox cNumFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   1
         ItemData        =   "frmMain.frx":2375E
         Left            =   3360
         List            =   "frmMain.frx":23760
         TabIndex        =   80
         Text            =   "500"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lYver 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Y vertical:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   240
         TabIndex        =   94
         Top             =   1005
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lXhor 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "X horizontal:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   240
         TabIndex        =   93
         Top             =   645
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label lMapN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Mapa:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   240
         TabIndex        =   92
         Top             =   285
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label lbFiltrar 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   2820
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lbCapas 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Capa Actual:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   120
         TabIndex        =   25
         Top             =   3195
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label lbGrh 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Sup Actual:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   2040
         TabIndex        =   24
         Top             =   3195
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lNumFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Numero de NPC:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   1
         Left            =   2160
         TabIndex        =   23
         Top             =   3195
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lCantFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   3195
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lbFiltrar 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   21
         Top             =   2820
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lNumFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Numero de OBJ:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   2
         Left            =   2160
         TabIndex        =   20
         Top             =   3195
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lCantFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   3195
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lbFiltrar 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   18
         Top             =   2820
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lCantFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   3195
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lNumFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Numero de NPC:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   0
         Left            =   2160
         TabIndex        =   16
         Top             =   3195
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lbFiltrar 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   2820
         Visible         =   0   'False
         Width           =   450
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   5460
      Left            =   60
      ScaleHeight     =   364
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   297
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6240
      Width           =   4455
      Begin VB.PictureBox PreviewGrh 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         FillColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   3825
         Left            =   30
         ScaleHeight     =   3795
         ScaleWidth      =   4335
         TabIndex        =   5
         Top             =   30
         Visible         =   0   'False
         Width           =   4365
      End
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   2565
      Top             =   2025
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Caption         =   "Informacion de mapa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   2715
      Begin FXIIIWorldEditor.lvButtons_H cmdInformacionDelMapa 
         Height          =   315
         Left            =   60
         TabIndex        =   95
         Top             =   860
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         Caption         =   "&Información del Mapa"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         cBack           =   8438015
      End
      Begin VB.Label lblMapVersion 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1170
         TabIndex        =   29
         Top             =   620
         Width           =   90
      End
      Begin VB.Label lblMapMusica 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1170
         TabIndex        =   28
         Top             =   360
         Width           =   90
      End
      Begin VB.Label lblMapNombre 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Nuevo Mapa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1170
         TabIndex        =   27
         Top             =   90
         Width           =   900
      End
      Begin VB.Label lblFMusica 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Musica:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   60
         TabIndex        =   7
         Top             =   330
         Width           =   2595
      End
      Begin VB.Label lblFVersion 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Versión:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   60
         TabIndex        =   6
         Top             =   600
         Width           =   2595
      End
      Begin VB.Label lblFNombreMapa 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   60
         TabIndex        =   3
         Top             =   60
         Width           =   2595
      End
   End
   Begin VB.Frame FrameComposicionTile 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Características del Tile de 32x32: -Click Derecho-"
      Height          =   1575
      Left            =   4620
      TabIndex        =   99
      Top             =   10275
      Width           =   10575
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Objeto:"
         Height          =   1215
         Left            =   8250
         TabIndex        =   116
         Top             =   240
         Width           =   2175
         Begin VB.Label LabelOBJCantidad 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cantidad="
            Height          =   255
            Left            =   120
            TabIndex        =   119
            Top             =   690
            Width           =   1935
         End
         Begin VB.Label LabelOBJName 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Name="
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            MousePointer    =   2  'Cross
            TabIndex        =   118
            Top             =   450
            Width           =   1935
         End
         Begin VB.Label LabelOBJIndex 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Index="
            Height          =   255
            Left            =   120
            TabIndex        =   117
            Top             =   210
            Width           =   1935
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "NPC:"
         Height          =   1215
         Left            =   5870
         TabIndex        =   112
         Top             =   240
         Width           =   2415
         Begin VB.Label LabelNPCName 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Name="
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            MousePointer    =   2  'Cross
            TabIndex        =   115
            Top             =   690
            Width           =   2175
         End
         Begin VB.Label LabelNPCIndex 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Index="
            Height          =   255
            Left            =   120
            TabIndex        =   114
            Top             =   450
            Width           =   2175
         End
         Begin VB.Label LabelTipoNPC 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Caption         =   "Hostil"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   113
            Top             =   210
            Width           =   2175
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Traslado:"
         Height          =   1215
         Left            =   4805
         TabIndex        =   108
         Top             =   240
         Width           =   1095
         Begin VB.Label LabelTrasMapa 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Mapa="
            Height          =   255
            Left            =   120
            TabIndex        =   111
            Top             =   210
            Width           =   855
         End
         Begin VB.Label LabelTrasX 
            BackColor       =   &H00E0E0E0&
            Caption         =   "X="
            Height          =   255
            Left            =   120
            TabIndex        =   110
            Top             =   450
            Width           =   855
         End
         Begin VB.Label LabelTrasY 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Y="
            Height          =   255
            Left            =   120
            TabIndex        =   109
            Top             =   690
            Width           =   855
         End
      End
      Begin VB.Frame FramePosicion 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tile:"
         Height          =   1215
         Left            =   120
         TabIndex        =   100
         Top             =   240
         Width           =   1095
         Begin VB.Label LabelBloqueado 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Caption         =   "Bloqueado"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   120
            TabIndex        =   120
            Top             =   720
            Width           =   855
         End
         Begin VB.Label LabelPosY 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Y="
            Height          =   255
            Left            =   120
            TabIndex        =   102
            Top             =   480
            Width           =   855
         End
         Begin VB.Label LabelPosX 
            BackColor       =   &H00E0E0E0&
            Caption         =   "X="
            Height          =   255
            Left            =   120
            TabIndex        =   101
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Superficies de Capas:"
         Height          =   1215
         Left            =   1155
         TabIndex        =   103
         Top             =   240
         Width           =   3675
         Begin VB.Label LabelCapa 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Capa3="
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   120
            MousePointer    =   2  'Cross
            TabIndex        =   107
            Top             =   720
            Width           =   3495
         End
         Begin VB.Label LabelCapa 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Capa4="
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   4
            Left            =   120
            MousePointer    =   2  'Cross
            TabIndex        =   106
            Top             =   960
            Width           =   3495
         End
         Begin VB.Label LabelCapa 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Capa1="
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            MousePointer    =   2  'Cross
            TabIndex        =   105
            Top             =   240
            Width           =   3495
         End
         Begin VB.Label LabelCapa 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Capa2="
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   120
            MousePointer    =   2  'Cross
            TabIndex        =   104
            Top             =   480
            Width           =   3495
         End
      End
   End
   Begin FXIIIWorldEditor.lvButtons_H CBVerBloqueosT 
      Height          =   495
      Left            =   15360
      TabIndex        =   122
      TabStop         =   0   'False
      Top             =   5520
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
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
      cBhover         =   -2147483633
      LockHover       =   1
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "frmMain.frx":23762
      ImgSize         =   24
      cBack           =   14737632
   End
   Begin FXIIIWorldEditor.lvButtons_H CBVerBloqueosA 
      Height          =   495
      Left            =   15360
      TabIndex        =   123
      TabStop         =   0   'False
      Top             =   6000
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
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
      cBhover         =   -2147483633
      LockHover       =   1
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "frmMain.frx":238C8
      ImgSize         =   24
      cBack           =   14737632
   End
   Begin FXIIIWorldEditor.lvButtons_H CBVerTriggers 
      Height          =   495
      Left            =   15360
      TabIndex        =   124
      TabStop         =   0   'False
      Top             =   6480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
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
      cBhover         =   -2147483633
      LockHover       =   1
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "frmMain.frx":23A2E
      ImgSize         =   48
      cBack           =   14737632
   End
   Begin FXIIIWorldEditor.lvButtons_H CBVerTraslados 
      Height          =   495
      Left            =   15360
      TabIndex        =   125
      TabStop         =   0   'False
      Top             =   6960
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
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
      cBhover         =   -2147483633
      LockHover       =   1
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "frmMain.frx":23BD0
      ImgSize         =   48
      cBack           =   14737632
   End
   Begin FXIIIWorldEditor.lvButtons_H CBVerCapa3 
      Height          =   495
      Left            =   15360
      TabIndex        =   144
      Top             =   4080
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   "3"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin FXIIIWorldEditor.lvButtons_H CBVerCapa4 
      Height          =   495
      Left            =   15360
      TabIndex        =   145
      Top             =   4560
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   "4"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin FXIIIWorldEditor.lvButtons_H CBVerCapa1 
      Height          =   495
      Left            =   15360
      TabIndex        =   146
      Top             =   3120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   "1"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   15360
      Picture         =   "frmMain.frx":23D72
      Top             =   2760
      Width           =   465
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   14340
      TabIndex        =   49
      Top             =   1230
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   13575
      TabIndex        =   48
      Top             =   1230
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   5925
      TabIndex        =   40
      Top             =   1230
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   6690
      TabIndex        =   39
      Top             =   1230
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   7455
      TabIndex        =   38
      Top             =   1230
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   8220
      TabIndex        =   37
      Top             =   1230
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   8985
      TabIndex        =   36
      Top             =   1230
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   9750
      TabIndex        =   35
      Top             =   1230
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   10515
      TabIndex        =   34
      Top             =   1230
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   11280
      TabIndex        =   33
      Top             =   1230
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   12045
      TabIndex        =   32
      Top             =   1230
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   5160
      TabIndex        =   31
      Top             =   1230
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   12810
      TabIndex        =   30
      Top             =   1230
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Shape MainViewShp 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00000000&
      FillColor       =   &H00C0E0FF&
      Height          =   8565
      Left            =   4620
      Top             =   1590
      Visible         =   0   'False
      Width           =   10575
   End
   Begin VB.Menu FileMnu 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuArchivoLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNuevoMapa 
         Caption         =   "&Nuevo Mapa"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuAbrirMapa 
         Caption         =   "&Abrir Mapa"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuArchivoLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReAbrirMapa 
         Caption         =   "&Re-Abrir Mapa"
      End
      Begin VB.Menu mnuArchivoLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGuardarMapa 
         Caption         =   "&Guardar Mapa"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuGuardarMapaComo 
         Caption         =   "Guardar Mapa &como..."
      End
      Begin VB.Menu mnuArchivoLine5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGuardarcomoBMP 
         Caption         =   "Guardar vista previa como &BMP"
      End
      Begin VB.Menu mnuGuardarcomoJPG 
         Caption         =   "Guardar vista previa como &JPG"
      End
      Begin VB.Menu mnuTodoMiniMap 
         Caption         =   "Guardar vista previa de TODOS los mapas (JPG)"
      End
      Begin VB.Menu mnuArchivoLine7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
      Begin VB.Menu mnuArchivoLine6 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuEdicion 
      Caption         =   "&Edición"
      Begin VB.Menu mnuDeshacer 
         Caption         =   "&Deshacer"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuUtilizarDeshacer 
         Caption         =   "&Utilizar Deshacer"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuInfoMap 
         Caption         =   "&Información del Mapa"
      End
      Begin VB.Menu mnuLineEdicion1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertar 
         Caption         =   "&Insertar"
         Begin VB.Menu mnuInsertarTransladosAdyasentes 
            Caption         =   "&Translados a Mapas Adyasentes"
         End
         Begin VB.Menu mnuInsertarSuperficieAlAzar 
            Caption         =   "Superficie al &Azar"
         End
         Begin VB.Menu mnuInsertarSuperficieEnBordes 
            Caption         =   "Superficie en los &Bordes del Mapa"
         End
         Begin VB.Menu mnuInsertarSuperficieEnTodo 
            Caption         =   "Superficie en Todo el Mapa"
         End
         Begin VB.Menu mnuBloquearBordes 
            Caption         =   "Bloqueo en &Bordes del Mapa"
         End
         Begin VB.Menu mnuBloquearMapa 
            Caption         =   "Bloqueo en &Todo el Mapa"
         End
      End
      Begin VB.Menu mnuQuitar 
         Caption         =   "&Quitar"
         Begin VB.Menu mnuQuitarTranslados 
            Caption         =   "Todos los &Translados"
         End
         Begin VB.Menu mnuQuitarBloqueos 
            Caption         =   "Todos los &Bloqueos"
         End
         Begin VB.Menu mnuQuitarNPCs 
            Caption         =   "Todos los &NPC's"
         End
         Begin VB.Menu mnuQuitarNPCsHostiles 
            Caption         =   "Todos los NPC's &Hostiles"
         End
         Begin VB.Menu mnuQuitarObjetos 
            Caption         =   "Todos los &Objetos"
         End
         Begin VB.Menu mnuQuitarTriggers 
            Caption         =   "Todos los Tri&gger's"
         End
         Begin VB.Menu mnuQuitarSuperficieBordes 
            Caption         =   "Superficie de los B&ordes"
         End
         Begin VB.Menu mnuQuitarSuperficieDeCapa 
            Caption         =   "Superficie de la &Capa Seleccionada"
         End
         Begin VB.Menu mnuLineEdicion2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuQuitarTODO 
            Caption         =   "TODO"
         End
      End
      Begin VB.Menu mnuLineEdicion3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFunciones 
         Caption         =   "&Funciones"
         Begin VB.Menu mnuQuitarFunciones 
            Caption         =   "&Quitar Funciones"
            Shortcut        =   {F4}
         End
         Begin VB.Menu mnuAutoQuitarFunciones 
            Caption         =   "Auto-&Quitar Funciones"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuConfigAvanzada 
         Caption         =   "Configuracion A&vanzada de Superficie"
      End
      Begin VB.Menu mnuLineEdicion4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutoCompletarSuperficies 
         Caption         =   "Auto-Completar &Superficies"
      End
      Begin VB.Menu mnuAutoCapturarSuperficie 
         Caption         =   "Auto-C&apturar información de la Superficie"
      End
      Begin VB.Menu mnuAutoCapturarTranslados 
         Caption         =   "Auto-&Capturar información de los Translados"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuAutoGuardarMapas 
         Caption         =   "Configuración de Auto-&Guardar Mapas"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "&Ver"
      Visible         =   0   'False
      Begin VB.Menu mnuCapas 
         Caption         =   "...&Capas"
         Begin VB.Menu mnuVerCapa1 
            Caption         =   "Capa &1 (Piso)"
         End
         Begin VB.Menu mnuVerCapa2 
            Caption         =   "Capa &2 (costas, etc)"
         End
         Begin VB.Menu mnuVerCapa3 
            Caption         =   "Capa &3 (arboles, etc)"
         End
         Begin VB.Menu mnuVerCapa4 
            Caption         =   "Capa &4 (techos, etc)"
         End
      End
      Begin VB.Menu mnuVerTranslados 
         Caption         =   "...&Translados"
      End
      Begin VB.Menu mnuVerBloqueos 
         Caption         =   "...&Bloqueos"
      End
      Begin VB.Menu mnuVerNPCs 
         Caption         =   "...&NPC's"
      End
      Begin VB.Menu mnuVerObjetos 
         Caption         =   "...&Objetos"
      End
      Begin VB.Menu mnuVerTriggers 
         Caption         =   "...Tri&gger's"
      End
      Begin VB.Menu mnuVerAgua 
         Caption         =   "...&Agua"
      End
      Begin VB.Menu mnuVerCuadricula 
         Caption         =   "...&Cuadrícula"
      End
      Begin VB.Menu mnuVerAutomatico 
         Caption         =   "Control &Automaticamente"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuPaneles 
      Caption         =   "&Paneles"
      Begin VB.Menu mnuSuperficie 
         Caption         =   "&Superficie"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuTranslados 
         Caption         =   "&Translados"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuBloquear 
         Caption         =   "&Bloquear"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuNPCs 
         Caption         =   "&NPC's"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuObjetos 
         Caption         =   "&Objetos"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuTriggers 
         Caption         =   "Tri&gger's"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuPanelesLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQSuperficie 
         Caption         =   "Ocultar Superficie"
         Shortcut        =   +{F5}
      End
      Begin VB.Menu mnuQTranslados 
         Caption         =   "Ocultar Translados"
         Shortcut        =   +{F6}
      End
      Begin VB.Menu mnuQBloquear 
         Caption         =   "Ocultar Bloquear"
         Shortcut        =   +{F7}
      End
      Begin VB.Menu mnuQNPCs 
         Caption         =   "Ocultar NPC's"
         Shortcut        =   +{F8}
      End
      Begin VB.Menu mnuQObjetos 
         Caption         =   "Ocultar Objetos"
         Shortcut        =   +{F11}
      End
      Begin VB.Menu mnuQTriggers 
         Caption         =   "Ocultar Trigger's"
         Shortcut        =   +{F12}
      End
      Begin VB.Menu mnuFuncionesLine1 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnuInformes 
         Caption         =   "&Informes"
      End
      Begin VB.Menu mnuActualizarIndices 
         Caption         =   "&Actualizar Indices de..."
         Begin VB.Menu mnuActualizarSuperficies 
            Caption         =   "&Superficies"
         End
         Begin VB.Menu mnuActualizarNPCs 
            Caption         =   "&NPC's"
         End
         Begin VB.Menu mnuActualizarObjs 
            Caption         =   "&Objetos"
         End
         Begin VB.Menu mnuActualizarTriggers 
            Caption         =   "&Trigger's"
         End
         Begin VB.Menu mnuActualizarCabezas 
            Caption         =   "C&abezas"
         End
         Begin VB.Menu mnuActualizarCuerpos 
            Caption         =   "C&uerpos"
         End
         Begin VB.Menu mnuActualizarGraficos 
            Caption         =   "&Graficos"
         End
      End
      Begin VB.Menu mnuModoCaminata 
         Caption         =   "Modalidad &Caminata"
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptimizar 
         Caption         =   "Optimi&zar Mapa"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGuardarUltimaConfig 
         Caption         =   "&Guardar Ultima Configuración"
      End
   End
   Begin VB.Menu mnuAcercaDe 
      Caption         =   "&Acerca de"
   End
   Begin VB.Menu mnuObjSc 
      Caption         =   "mnuObjSc"
      Visible         =   0   'False
      Begin VB.Menu mnuConfigObjTrans 
         Caption         =   "&Utilizar como Objeto de Translados"
      End
   End
   Begin VB.Menu MenuPantalla 
      Caption         =   "&Pantalla"
      Visible         =   0   'False
      Begin VB.Menu MenuTHorizontal 
         Caption         =   "&Tamaño Horizontal"
         Begin VB.Menu MenuTamañoH 
            Caption         =   "&Chico"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu MenuTamañoH 
            Caption         =   "&Mediano"
            Index           =   2
         End
         Begin VB.Menu MenuTamañoH 
            Caption         =   "&Grande"
            Index           =   3
         End
      End
      Begin VB.Menu menuTVertical 
         Caption         =   "&Tamaño Vertical"
         Begin VB.Menu MenuTamañoV 
            Caption         =   "&Chico"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu MenuTamañoV 
            Caption         =   "&Mediano"
            Index           =   2
         End
         Begin VB.Menu MenuTamañoV 
            Caption         =   "&Grande"
            Index           =   3
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub PonerAlAzar(ByVal N As Integer, T As Byte)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06 by GS
'*************************************************
Dim objindex As Long
Dim NpcIndex As Long
Dim X, Y, i
Dim Head As Integer
Dim Body As Integer
Dim Heading As Byte
Dim Leer As New clsIniReader
i = N

modEdicion.Deshacer_Add "Aplicar " & IIf(T = 0, "Objetos", "NPCs") & " al Azar" ' Hago deshacer

Do While i > 0
    X = CInt(RandomNumber(XMinMapSize, XMaxMapSize - 1))
    Y = CInt(RandomNumber(YMinMapSize, YMaxMapSize - 1))
    
    Select Case T
        Case 0
            If MapData(X, Y).OBJInfo.objindex = 0 Then
                  i = i - 1
                  If cInsertarBloqueo.value = True Then
                    MapData(X, Y).Blocked = 1
                  Else
                    MapData(X, Y).Blocked = 0
                  End If
                  If cNumFunc(2).Text > 0 Then
                      objindex = cNumFunc(2).Text
                      InitGrh MapData(X, Y).ObjGrh, ObjData(objindex).GrhIndex
                      MapData(X, Y).OBJInfo.objindex = objindex
                      MapData(X, Y).OBJInfo.Amount = Val(cCantFunc(2).Text)
                      Select Case ObjData(objindex).ObjType ' GS
                            Case 4, 8, 10, 22 ' Arboles, Carteles, Foros, Yacimientos
                                MapData(X, Y).Graphic(3) = MapData(X, Y).ObjGrh
                      End Select
                  End If
            End If
        Case 1
           If MapData(X, Y).Blocked = 0 Then
                  i = i - 1
                  If cNumFunc(T - 1).Text > 0 Then
                        NpcIndex = cNumFunc(T - 1).Text
                        Body = NpcData(NpcIndex).Body
                        Head = NpcData(NpcIndex).Head
                        Heading = NpcData(NpcIndex).Heading
                        Call MakeChar(NextOpenChar(), Body, Head, Heading, CInt(X), CInt(Y))
                        MapData(X, Y).NpcIndex = NpcIndex
                  End If
            End If
        Case 2
           If MapData(X, Y).Blocked = 0 Then
                  i = i - 1
                  If cNumFunc(T - 1).Text >= 0 Then
                        NpcIndex = cNumFunc(T - 1).Text
                        Body = NpcData(NpcIndex).Body
                        Head = NpcData(NpcIndex).Head
                        Heading = NpcData(NpcIndex).Heading
                        Call MakeChar(NextOpenChar(), Body, Head, Heading, CInt(X), CInt(Y))
                        MapData(X, Y).NpcIndex = NpcIndex
                  End If
           End If
        End Select
        DoEvents
Loop
End Sub

Private Sub cAgregarFuncalAzar_Click(Index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next
If IsNumeric(cCantFunc(Index).Text) = False Or cCantFunc(Index).Text > 200 Then
    MsgBox "El Valor de Cantidad introducido no es soportado!" & vbCrLf & "El valor maximo es 200.", vbCritical
    Exit Sub
End If
cAgregarFuncalAzar(Index).Enabled = False
Call PonerAlAzar(CInt(cCantFunc(Index).Text), 1 + (IIf(Index = 2, -1, Index)))
cAgregarFuncalAzar(Index).Enabled = True
End Sub

Private Sub CBVerBloqueosA_Click()
Call mnuVerAgua_Click
End Sub

Private Sub CBVerBloqueosT_Click()
Call mnuVerBloqueos_Click
End Sub

Private Sub CBVerCapa1_Click()
Call mnuVerCapa1_Click
End Sub

Private Sub CBVerCapa2_Click()
Call mnuVerCapa2_Click
End Sub

Private Sub CBVerCapa3_Click()
Call mnuVerCapa3_Click
End Sub

Private Sub CBVerCapa4_Click()
Call mnuVerCapa4_Click
End Sub

Private Sub CBVerCuadricula_Click()
Call mnuVerCuadricula_Click
End Sub

Private Sub CBVerNpcs_Click()
Call mnuVerNPCs_Click
End Sub

Private Sub CBVerObjetos_Click()
Call mnuVerObjetos_Click
End Sub

Private Sub CBVerTraslados_Click()
Call mnuVerTranslados_Click
End Sub

Private Sub CBVerTriggers_Click()
Call mnuVerTriggers_Click
End Sub

Private Sub cCantFunc_Change(Index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next
    If Val(cCantFunc(Index)) < 1 Then
      cCantFunc(Index).Text = 1
    End If
    If Val(cCantFunc(Index)) > 10000 Then
      cCantFunc(Index).Text = 10000
    End If
End Sub

Private Sub cCapas_Change()
'*************************************************
'Author: ^[GS]^
'Last modified: 31/05/06
'*************************************************
    If Val(cCapas.Text) < 1 Then
      cCapas.Text = 1
    End If
    If Val(cCapas.Text) > 4 Then
      cCapas.Text = 4
    End If
    cCapas.Tag = vbNullString
End Sub

Private Sub cCapas_KeyPress(KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If IsNumeric(Chr(KeyAscii)) = False Then KeyAscii = 0
End Sub

Private Sub cFiltro_GotFocus(Index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
HotKeysAllow = False
End Sub

Private Sub cFiltro_KeyPress(Index As Integer, KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If KeyAscii = 13 Then
    Call Filtrar(Index)
End If
End Sub

Private Sub cFiltro_LostFocus(Index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
HotKeysAllow = True
End Sub

Private Sub cGrh_KeyPress(KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

On Error GoTo Fallo
If KeyAscii = 13 Then
    Call fPreviewGrh(cGrh.Text)
    If frmMain.cGrh.ListCount > 5 Then
        frmMain.cGrh.RemoveItem 0
    End If
    frmMain.cGrh.AddItem frmMain.cGrh.Text
End If
Exit Sub
Fallo:
    cGrh.Text = 1

End Sub

Private Sub CheckDespMosaico_Click()
'*************************************************
'Author: Deut
'Last modified: 09/12/09
'*************************************************
If CheckDespMosaico.value = 1 Then
    TextHorizontal.Caption = -Val(frmConfigSup.DMAncho.Text)
    TextVertical.Caption = -Val(frmConfigSup.DMLargo.Text)
    frmConfigSup.DespMosaic.value = 1
Else
    TextHorizontal.Caption = ""
    TextVertical.Caption = ""
    frmConfigSup.DespMosaic.value = 0
End If
End Sub

Private Sub cInsertarFunc_Click(Index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If cInsertarFunc(Index).value = True Then
    cQuitarFunc(Index).Enabled = False
    cAgregarFuncalAzar(Index).Enabled = False
    If Index <> 2 Then cCantFunc(Index).Enabled = False
    Call modPaneles.EstSelectPanel((Index) + 3, True)
Else
    cQuitarFunc(Index).Enabled = True
    cAgregarFuncalAzar(Index).Enabled = True
    If Index <> 2 Then cCantFunc(Index).Enabled = True
    Call modPaneles.EstSelectPanel((Index) + 3, False)
End If
End Sub

Private Sub cInsertarTrans_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 22/05/06
'*************************************************
If cInsertarTrans.value = True Then
    cQuitarTrans.Enabled = False
    Call modPaneles.EstSelectPanel(1, True)
Else
    cQuitarTrans.Enabled = True
    Call modPaneles.EstSelectPanel(1, False)
End If
End Sub

Private Sub cInsertarTrigger_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If cInsertarTrigger.value = True Then
    cQuitarTrigger.Enabled = False
    Call modPaneles.EstSelectPanel(5, True)
Else
    cQuitarTrigger.Enabled = True
    Call modPaneles.EstSelectPanel(5, False)
End If
End Sub


Private Sub cmdCommand1_Click()
Dim i As Integer
Dim cantmaps As Integer
cantmaps = CInt(Val(InputBox("Introduzca la cantidad de mapas:", "Cantidad de mapas", "191")))
For i = 1 To cantmaps
    Call CargarMapa(Dialog.FileName)
    
 'Dim Ratio As Integer
 '   Ratio = CInt(Val(InputBox("En que escala queres Renderizar? Entre 1 y 20.", "Elegi Escala", "1")))
 '   If Ratio < 1 Then Ratio = 1
 '   If Ratio >= 1 And Ratio <= 20 Then
 '       RenderToPicture Ratio, True
 '   End If
Next i
End Sub

Private Sub cmdInformacionDelMapa_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
frmMapInfo.Show
frmMapInfo.Visible = True
End Sub

Private Sub cmdQuitarFunciones_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call mnuQuitarFunciones_Click
End Sub

Private Sub CommandHorMas_Click()
'*************************************************
'Author: Deut
'Last modified: 09/12/09
'*************************************************
If frmConfigSup.DespMosaic.value = 1 Then
    TextHorizontal.Caption = Val(TextHorizontal.Caption) + 1
    frmConfigSup.DMAncho.Text = -Val(TextHorizontal.Caption)
End If
End Sub

Private Sub CommandHorMenos_Click()
'*************************************************
'Author: Deut
'Last modified: 09/12/09
'*************************************************
If frmConfigSup.DespMosaic.value = 1 Then
    TextHorizontal.Caption = Val(TextHorizontal.Caption) - 1
    frmConfigSup.DMAncho.Text = -Val(TextHorizontal.Caption)
End If
End Sub

Private Sub CommandVerMas_Click()
'*************************************************
'Author: Deut
'Last modified: 09/12/09
'*************************************************
If frmConfigSup.DespMosaic.value = 1 Then
    TextVertical.Caption = Val(TextVertical.Caption) + 1
    frmConfigSup.DMLargo.Text = -Val(TextVertical.Caption)
End If
End Sub

Private Sub CommandVerMenos_Click()
'*************************************************
'Author: Deut
'Last modified: 09/12/09
'*************************************************
If frmConfigSup.DespMosaic.value = 1 Then
    TextVertical.Caption = Val(TextVertical.Caption) - 1
    frmConfigSup.DMLargo.Text = -Val(TextVertical.Caption)
End If
End Sub

Private Sub cUnionManual_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
cInsertarTrans.value = (cUnionManual.value = True)
Call cInsertarTrans_Click
End Sub

'Private Sub cverBloqueos_Click()
'*************************************************
'Author: Deut
'Last modified: 19/12/09
'*************************************************
'mnuVerBloqueos.Checked = cVerBloqueos.value
'frmMain.CBVerBloqueosT.value = cVerBloqueos.value
'Call mnuVerBloqueos_Click
'End Sub

'Private Sub cverTriggers_Click()
'*************************************************
'Author: Deut
'Last modified: 18/12/09
'*************************************************
'mnuVerTriggers.Checked = cVerTriggers.value
'frmMain.CBVerTriggers.value = cVerTriggers.value
'Call mnuVerTriggers_Click
'End Sub

Private Sub cNumFunc_KeyPress(Index As Integer, KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next

If KeyAscii = 13 Then
    Dim Cont As String
    Cont = frmMain.cNumFunc(Index).Text
    Call cNumFunc_LostFocus(Index)
    If Cont <> frmMain.cNumFunc(Index).Text Then Exit Sub
    If frmMain.cNumFunc(Index).ListCount > 5 Then
        frmMain.cNumFunc(Index).RemoveItem 0
    End If
    frmMain.cNumFunc(Index).AddItem frmMain.cNumFunc(Index).Text
    Exit Sub
ElseIf KeyAscii = 8 Then
    
ElseIf IsNumeric(Chr(KeyAscii)) = False Then
    KeyAscii = 0
    Exit Sub
End If

End Sub

Private Sub cNumFunc_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next
If cNumFunc(Index).Text = vbNullString Then
    frmMain.cNumFunc(Index).Text = IIf(Index = 1, 500, 1)
End If
End Sub

Private Sub cNumFunc_LostFocus(Index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next
    If Index = 0 Then
        If frmMain.cNumFunc(Index).Text > 499 Or frmMain.cNumFunc(Index).Text < 1 Then
            frmMain.cNumFunc(Index).Text = 1
        End If
    ElseIf Index = 1 Then
        If frmMain.cNumFunc(Index).Text < 500 Or frmMain.cNumFunc(Index).Text > 32000 Then
            frmMain.cNumFunc(Index).Text = 500
        End If
    ElseIf Index = 2 Then
        If frmMain.cNumFunc(Index).Text < 1 Or frmMain.cNumFunc(Index).Text > 32000 Then
            frmMain.cNumFunc(Index).Text = 1
        End If
    End If
End Sub

Private Sub cInsertarBloqueo_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
cInsertarBloqueo.Tag = vbNullString
If cInsertarBloqueo.value = True Then
    cQuitarBloqueo.Enabled = False
    Call modPaneles.EstSelectPanel(2, True)
Else
    cQuitarBloqueo.Enabled = True
    Call modPaneles.EstSelectPanel(2, False)
End If
End Sub

Private Sub cQuitarBloqueo_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
cInsertarBloqueo.Tag = vbNullString
If cQuitarBloqueo.value = True Then
    cInsertarBloqueo.Enabled = False
    Call modPaneles.EstSelectPanel(2, True)
Else
    cInsertarBloqueo.Enabled = True
    Call modPaneles.EstSelectPanel(2, False)
End If
End Sub

Private Sub cQuitarEnEstaCapa_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If cQuitarEnEstaCapa.value = True Then
    lListado(0).Enabled = False
    cFiltro(0).Enabled = False
    cGrh.Enabled = False
    cSeleccionarSuperficie.Enabled = False
    cQuitarEnTodasLasCapas.Enabled = False
    Call modPaneles.EstSelectPanel(0, True)
Else
    lListado(0).Enabled = True
    cFiltro(0).Enabled = True
    cGrh.Enabled = True
    cSeleccionarSuperficie.Enabled = True
    cQuitarEnTodasLasCapas.Enabled = True
    Call modPaneles.EstSelectPanel(0, False)
End If
End Sub

Private Sub cQuitarEnTodasLasCapas_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If cQuitarEnTodasLasCapas.value = True Then
    cCapas.Enabled = False
    lListado(0).Enabled = False
    cFiltro(0).Enabled = False
    cGrh.Enabled = False
    cSeleccionarSuperficie.Enabled = False
    cQuitarEnEstaCapa.Enabled = False
    Call modPaneles.EstSelectPanel(0, True)
Else
    cCapas.Enabled = True
    lListado(0).Enabled = True
    cFiltro(0).Enabled = True
    cGrh.Enabled = True
    cSeleccionarSuperficie.Enabled = True
    cQuitarEnEstaCapa.Enabled = True
    Call modPaneles.EstSelectPanel(0, False)
End If
End Sub


Private Sub cQuitarFunc_Click(Index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If cQuitarFunc(Index).value = True Then
    cInsertarFunc(Index).Enabled = False
    cAgregarFuncalAzar(Index).Enabled = False
    cCantFunc(Index).Enabled = False
    cNumFunc(Index).Enabled = False
    cFiltro((Index) + 1).Enabled = False
    lListado((Index) + 1).Enabled = False
    Call modPaneles.EstSelectPanel((Index) + 3, True)
Else
    cInsertarFunc(Index).Enabled = True
    cAgregarFuncalAzar(Index).Enabled = True
    cCantFunc(Index).Enabled = True
    cNumFunc(Index).Enabled = True
    cFiltro((Index) + 1).Enabled = True
    lListado((Index) + 1).Enabled = True
    Call modPaneles.EstSelectPanel((Index) + 3, False)
End If
End Sub

Private Sub cQuitarTrans_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If cQuitarTrans.value = True Then
    cInsertarTransOBJ.Enabled = False
    cInsertarTrans.Enabled = False
    cUnionManual.Enabled = False
    cUnionAuto.Enabled = False
    tTMapa.Enabled = False
    tTX.Enabled = False
    tTY.Enabled = False
    mnuInsertarTransladosAdyasentes.Enabled = False
    Call modPaneles.EstSelectPanel(1, True)
Else
    tTMapa.Enabled = True
    tTX.Enabled = True
    tTY.Enabled = True
    cUnionAuto.Enabled = True
    cUnionManual.Enabled = True
    cInsertarTrans.Enabled = True
    cInsertarTransOBJ.Enabled = True
    mnuInsertarTransladosAdyasentes.Enabled = True
    Call modPaneles.EstSelectPanel(1, False)
End If
End Sub

Private Sub cQuitarTrigger_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If cQuitarTrigger.value = True Then
    lListado(4).Enabled = False
    cInsertarTrigger.Enabled = False
    Call modPaneles.EstSelectPanel(5, True)
Else
    lListado(4).Enabled = True
    cInsertarTrigger.Enabled = True
    Call modPaneles.EstSelectPanel(5, False)
End If
End Sub

Private Sub cSeleccionarSuperficie_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If cSeleccionarSuperficie.value = True Then
    cQuitarEnTodasLasCapas.Enabled = False
    cQuitarEnEstaCapa.Enabled = False
    Call modPaneles.EstSelectPanel(0, True)
Else
    cQuitarEnTodasLasCapas.Enabled = True
    cQuitarEnEstaCapa.Enabled = True
    Call modPaneles.EstSelectPanel(0, False)
End If
End Sub

Private Sub cUnionAuto_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
frmUnionAdyacente.Show
End Sub

Private Sub Form_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Me.SetFocus
End Sub

Private Sub Form_DblClick()
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************
Dim tX As Integer
Dim tY As Integer

If Not MapaCargado Then Exit Sub

If SobreX > 0 And SobreY > 0 Then
    DobleClick Val(SobreX), Val(SobreY)
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^, Deut
'Last modified: 09/12/09
'*************************************************
' HotKeys


If HotKeysAllow = False Then Exit Sub

Select Case KeyAscii
    Case Is = 53 'NumPad5
        If frmConfigSup.DespMosaic.value = 1 Then
            TextHorizontal.Caption = ""
            TextVertical.Caption = ""
            frmConfigSup.DespMosaic.value = 0
            frmMain.CheckDespMosaico.value = 0
        Else
            TextHorizontal.Caption = -Val(frmConfigSup.DMAncho.Text)
            TextVertical.Caption = -Val(frmConfigSup.DMLargo.Text)
            frmConfigSup.DespMosaic.value = 1
            frmMain.CheckDespMosaico.value = 1
        End If
    Case Is = 56 'NumPad8
        CommandVerMenos_Click
    Case Is = 54 'NumPad6
        CommandHorMas_Click
    Case Is = 50 'NumPad2
        CommandVerMas_Click
    Case Is = 52 'NumPad4
        CommandHorMenos_Click
End Select


'Select Case UCase(Chr(KeyAscii))
'    Case "S" ' Activa/Desactiva Insertar Superficie
'        cSeleccionarSuperficie.value = (cSeleccionarSuperficie.value = False)
'        Call cSeleccionarSuperficie_Click
'    Case "T" ' Activa/Desactiva Insertar Translados
'        cInsertarTrans.value = (cInsertarTrans.value = False)
'        Call cInsertarTrans_Click
'    Case "B" ' Activa/Desactiva Insertar Bloqueos
'        cInsertarBloqueo.value = (cInsertarBloqueo.value = False)
'        Call cInsertarBloqueo_Click
'    Case "N" ' Activa/Desactiva Insertar NPCs
'        cInsertarFunc(0).value = (cInsertarFunc(0).value = False)
'        Call cInsertarFunc_Click(0)
'    Case "H" ' Activa/Desactiva Insertar NPCs Hostiles
'        cInsertarFunc(1).value = (cInsertarFunc(1).value = False)
'        Call cInsertarFunc_Click(1)
'    Case "O" ' Activa/Desactiva Insertar Objetos
'        cInsertarFunc(2).value = (cInsertarFunc(2).value = False)
'        Call cInsertarFunc_Click(2)
'    Case "G" ' Activa/Desactiva Insertar Triggers
'        cInsertarTrigger.value = (cInsertarTrigger.value = False)
'        Call cInsertarTrigger_Click
'    Case "Q" ' Quitar Funciones
'        Call mnuQuitarFunciones_Click
'End Select
End Sub

Private Sub Form_Load()
'*************************************************
'Author: Deut
'Last modified: 09/12/09
'*************************************************
DInput.Initialize
End Sub


Private Sub LabelCapa_Click(Index As Integer)
'*************************************************
'Author: Deut
'Last modified: 09/12/09
'*************************************************
cFiltro(0).Text = Right(LabelCapa(Index).Caption, Len(LabelCapa(Index).Caption) - 7)
Call Filtrar(0)
Dim i As Byte
For i = 1 To 5
    SelectPanel(i).value = False
    Call VerFuncion(i, False)
Next
SelectPanel(0).value = True
Call VerFuncion(0, True)
'lListado(0).listIndex = 0
'Call lListado_Click(0)
End Sub

Private Sub LabelNPCName_Click()
'*************************************************
'Author: Deut
'Last modified: 09/12/09
'*************************************************
Dim tPanel As Byte
If LabelTipoNPC.Caption = "Pacífico" Then
    tPanel = 3
Else
    tPanel = 4
End If
cFiltro(tPanel - 2).Text = Right(LabelNPCName.Caption, Len(LabelNPCName.Caption) - 7)
Call Filtrar(tPanel - 2)
Dim i As Byte
For i = 0 To 5
    If i <> tPanel Then
        SelectPanel(i).value = False
        Call VerFuncion(i, False)
    End If
Next
SelectPanel(tPanel).value = True
Call VerFuncion(tPanel, True)
'lListado(tPanel - 2).listIndex = 0
'lListado_Click (tPanel - 2)


End Sub

Private Sub LabelOBJName_Click()
'*************************************************
'Author: Deut
'Last modified: 09/12/09
'*************************************************
cFiltro(3).Text = Right(LabelOBJName.Caption, Len(LabelOBJName.Caption) - 7)
Call Filtrar(3)
Dim i As Byte
For i = 0 To 5
    If i <> 5 Then
        SelectPanel(i).value = False
        Call VerFuncion(i, False)
    End If
Next
SelectPanel(4).value = True
Call VerFuncion(4, True)
'lListado(3).listIndex = 0
'lListado_Click (3)
End Sub

Private Sub lListado_Click(Index As Integer)
'*************************************************
'Author: ^[GS]^,Deut
'Last modified: 09/12/09
'*************************************************
On Error Resume Next
If HotKeysAllow = False Then
    lListado(Index).Tag = lListado(Index).listIndex
    Select Case Index
        Case 0
            cGrh.Text = DameGrhIndex(ReadField(2, lListado(Index).List(lListado(Index).listIndex), Asc("#")))
            If SupData(ReadField(2, lListado(Index).Text, Asc("#"))).Capa <> 0 Then
                If LenB(ReadField(2, lListado(Index).Text, Asc("#"))) = 0 Then cCapas.Tag = cCapas.Text
                cCapas.Text = SupData(ReadField(2, lListado(Index).Text, Asc("#"))).Capa
            Else
                If LenB(cCapas.Tag) <> 0 Then
                    cCapas.Text = cCapas.Tag
                    cCapas.Tag = vbNullString
                End If
            End If
            If SupData(ReadField(2, lListado(Index).Text, Asc("#"))).Block = True Then
                If LenB(cInsertarBloqueo.Tag) = 0 Then cInsertarBloqueo.Tag = IIf(cInsertarBloqueo.value = True, 1, 0)
                cInsertarBloqueo.value = True
                Call cInsertarBloqueo_Click
            Else
                If LenB(cInsertarBloqueo.Tag) <> 0 Then
                    cInsertarBloqueo.value = IIf(Val(cInsertarBloqueo.Tag) = 1, True, False)
                    cInsertarBloqueo.Tag = vbNullString
                    Call cInsertarBloqueo_Click
                End If
            End If
            Call fPreviewGrh(cGrh.Text)
        Case 1
            cNumFunc(0).Text = ReadField(2, lListado(Index).Text, Asc("#"))
        Case 2
            cNumFunc(1).Text = ReadField(2, lListado(Index).Text, Asc("#"))
        Case 3
            cNumFunc(2).Text = ReadField(2, lListado(Index).Text, Asc("#"))
    End Select
Else
    lListado(Index).listIndex = lListado(Index).Tag
End If

End Sub

'Private Sub lListado_GotFocus(Index As Integer)
'FocoEnLista = True
'End Sub
'
'Private Sub lListado_LostFocus(Index As Integer)
'FocoEnLista = False
'End Sub

Private Sub lListado_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
If Index = 3 And Button = 2 Then
    If lListado(3).listIndex > -1 Then Me.PopupMenu mnuObjSc
End If
End Sub

Private Sub lListado_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: ^[GS]^
'Last modified: 22/05/06
'*************************************************
On Error Resume Next
HotKeysAllow = False
FocoEnLista = True
End Sub

Private Sub MapPest_Click(Index As Integer)
'*************************************************
'Author: Deut
'Last modified: 09/12/09
'*************************************************

If (Index + NumMap_Save - 4) <> NumMap_Save Then
    Dialog.CancelError = True
    On Error GoTo ErrHandler
    Dialog.FileName = PATH_Save & NameMap_Save & (Index + NumMap_Save - 4) & ".mcl"
    Call DeseaGuardarMapa(Dialog.FileName)
    Call NuevoMapa
    DoEvents
    Call CargarMapa(Dialog.FileName)
    EngineRun = True

Exit Sub
ErrHandler:
    MsgBox Err.Description
    End If
End Sub

Private Sub MenuTamañoH_Click(Index As Integer)
'*************************************************
'Author: Deut
'Last modified: 09/12/09
'*************************************************
Dim i As Byte
For i = 1 To 3
    If Index = i Then
        MenuTamañoH(i).Checked = True
        With Me.MainViewShp
            .Width = Int(.Width * (1.2) ^ (Index))
        End With
        
    Else
        MenuTamañoH(i).Checked = False
    End If
Next
End Sub


Private Sub mnuActualizarCabezas_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
Call modIndices.CargarIndicesDeCabezas
End Sub

Private Sub mnuActualizarCuerpos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
Call modIndices.CargarIndicesDeCuerpos
End Sub

Private Sub mnuActualizarGraficos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
Call modIndices.CargarIndicesDeGraficos
End Sub

Private Sub mnuActualizarSuperficies_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modIndices.CargarIndicesSuperficie
End Sub

Private Sub mnuAbrirMapa_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Dialog.CancelError = True
On Error GoTo ErrHandler

Call DeseaGuardarMapa(Dialog.FileName)
Call ObtenerNombreArchivo(False)

If Len(Dialog.FileName) < 3 Then Exit Sub

    If WalkMode = True Then
        Call modGeneral.ToggleWalkMode
    End If
    
    Call modMapIO.NuevoMapa
    modMapIO.CargarMapa Dialog.FileName
    DoEvents
    mnuReAbrirMapa.Enabled = True
    EngineRun = True
    
Exit Sub
ErrHandler:
End Sub

Private Sub mnuacercade_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
frmAbout.Show
End Sub

Private Sub mnuActualizarNPCs_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modIndices.CargarIndicesNPC
End Sub

Private Sub mnuActualizarObjs_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modIndices.CargarIndicesOBJ
End Sub

Private Sub mnuActualizarTriggers_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modIndices.CargarIndicesTriggers
End Sub

Private Sub mnuArchiveLine7_Click()

End Sub

Private Sub mnuAutoCapturarTranslados_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************
mnuAutoCapturarTranslados.Checked = (mnuAutoCapturarTranslados.Checked = False)
End Sub

Private Sub mnuAutoCapturarSuperficie_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************
mnuAutoCapturarSuperficie.Checked = (mnuAutoCapturarSuperficie.Checked = False)

End Sub

Private Sub mnuAutoCompletarSuperficies_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
mnuAutoCompletarSuperficies.Checked = (mnuAutoCompletarSuperficies.Checked = False)

End Sub

Private Sub mnuAutoGuardarMapas_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
frmAutoGuardarMapa.Show
End Sub

Private Sub mnuAutoQuitarFunciones_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
mnuAutoQuitarFunciones.Checked = (mnuAutoQuitarFunciones.Checked = False)

End Sub


Private Sub mnuBloquear_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Dim i As Byte
For i = 0 To 5
    If i <> 2 Then
        frmMain.SelectPanel(i).value = False
        Call VerFuncion(i, False)
    End If
Next

modPaneles.VerFuncion 2, True
End Sub

Private Sub mnuBloquearBordes_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Bloquear_Bordes
End Sub

Private Sub mnuBloquearMapa_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Bloqueo_Todo(1)
End Sub

Private Sub mnuConfigAvanzada_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
frmConfigSup.Show
End Sub

Private Sub mnuConfigObjTrans_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
Cfg_TrOBJ = cNumFunc(2).Text
End Sub

Private Sub mnuDeshacer_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 15/10/06
'*************************************************
Call modEdicion.Deshacer_Recover
End Sub

Private Sub mnuGuardarcomoBMP_Click()
'*************************************************
'Author: Salvito
'Last modified: 01/05/2008 - ^[GS]^
'*************************************************
    Dim Ratio As Integer
    Ratio = CInt(Val(InputBox("En que escala queres Renderizar? Entre 1 y 20.", "Elegi Escala", "1")))
    If Ratio < 1 Then Ratio = 1
    If Ratio >= 1 And Ratio <= 20 Then
        RenderToPicture Ratio, True
    End If
End Sub

Private Sub mnuGuardarcomoJPG_Click()
'*************************************************
'Author: Salvito
'Last modified: 01/05/2008 - ^[GS]^
'*************************************************
    Dim Ratio As Integer
    Ratio = CInt(Val(InputBox("En que escala queres Renderizar? Entre 1 y 20.", "Elegi Escala", "1")))
    If Ratio < 1 Then Ratio = 1
    If Ratio >= 1 And Ratio <= 20 Then
        RenderToPicture Ratio, False
    End If
End Sub

Private Sub mnuGuardarMapa_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
modMapIO.GuardarMapa Dialog.FileName
End Sub

Private Sub mnuGuardarMapaComo_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
modMapIO.GuardarMapa
End Sub

Private Sub mnuGuardarUltimaConfig_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 23/05/06
'*************************************************
mnuGuardarUltimaConfig.Checked = (mnuGuardarUltimaConfig.Checked = False)
End Sub

Private Sub mnuInfoMap_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
frmMapInfo.Show
frmMapInfo.Visible = True
End Sub

Private Sub mnuInformes_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
frmInformes.Show
End Sub

Private Sub mnuInsertarSuperficieAlAzar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Superficie_Azar
End Sub

Private Sub mnuInsertarSuperficieEnBordes_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Superficie_Bordes
End Sub

Private Sub mnuInsertarSuperficieEnTodo_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Superficie_Todo
End Sub

Private Sub mnuInsertarTransladosAdyasentes_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
frmUnionAdyacente.Show
End Sub

Private Sub mnuModoCaminata_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************
ToggleWalkMode
End Sub

Private Sub mnuNPCs_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Dim i As Byte
For i = 0 To 5
    If i <> 3 Then
        frmMain.SelectPanel(i).value = False
        Call VerFuncion(i, False)
    End If
Next
modPaneles.VerFuncion 3, True
End Sub



Private Sub mnuNuevoMapa_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next
Dim loopc As Integer

DeseaGuardarMapa Dialog.FileName

For loopc = 0 To frmMain.MapPest.Count
    frmMain.MapPest(loopc).Visible = False
Next

frmMain.Dialog.FileName = Empty

If WalkMode = True Then
    Call modGeneral.ToggleWalkMode
End If

Call modMapIO.NuevoMapa

Call cmdInformacionDelMapa_Click

End Sub

Private Sub mnuObjetos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Dim i As Byte
For i = 0 To 5
    If i <> 4 Then
        frmMain.SelectPanel(i).value = False
        Call VerFuncion(i, False)
    End If
Next
modPaneles.VerFuncion 4, True
End Sub


Private Sub mnuOptimizar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 22/09/06
'*************************************************
frmOptimizar.Show
End Sub

Private Sub mnuQBloquear_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
modPaneles.VerFuncion 2, False
End Sub

Private Sub mnuQNPCs_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
modPaneles.VerFuncion 3, False
End Sub


Private Sub mnuQObjetos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
modPaneles.VerFuncion 4, False
End Sub

Private Sub mnuQSuperficie_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
modPaneles.VerFuncion 0, False
End Sub

Private Sub mnuQTranslados_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
modPaneles.VerFuncion 1, False
End Sub

Private Sub mnuQTriggers_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
modPaneles.VerFuncion 5, False
End Sub


Private Sub mnuQuitarBloqueos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Bloqueo_Todo(0)
End Sub

Private Sub mnuQuitarFunciones_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
' Superficies
cSeleccionarSuperficie.value = False
Call cSeleccionarSuperficie_Click
cQuitarEnEstaCapa.value = False
Call cQuitarEnEstaCapa_Click
cQuitarEnTodasLasCapas.value = False
Call cQuitarEnTodasLasCapas_Click
' Translados
cQuitarTrans.value = False
Call cQuitarTrans_Click
cInsertarTrans.value = False
Call cInsertarTrans_Click
' Bloqueos
cQuitarBloqueo.value = False
Call cQuitarBloqueo_Click
cInsertarBloqueo.value = False
Call cInsertarBloqueo_Click
' Otras funciones
cInsertarFunc(0).value = False
Call cInsertarFunc_Click(0)
cInsertarFunc(1).value = False
Call cInsertarFunc_Click(1)
cInsertarFunc(2).value = False
Call cInsertarFunc_Click(2)
cQuitarFunc(0).value = False
Call cQuitarFunc_Click(0)
cQuitarFunc(1).value = False
Call cQuitarFunc_Click(1)
cQuitarFunc(2).value = False
Call cQuitarFunc_Click(2)
' Triggers
cInsertarTrigger.value = False
Call cInsertarTrigger_Click
cQuitarTrigger.value = False
Call cQuitarTrigger_Click
End Sub

Private Sub mnuQuitarNPCs_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Quitar_NPCs(False)
End Sub

Private Sub mnuQuitarNPCsHostiles_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Quitar_NPCs(True)
End Sub

Private Sub mnuQuitarObjetos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Quitar_Objetos
End Sub

Private Sub mnuQuitarSuperficieBordes_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Quitar_Bordes
End Sub

Private Sub mnuQuitarSuperficieDeCapa_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Quitar_Capa(cCapas.Text)
End Sub

Private Sub mnuQuitarTODO_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Borrar_Mapa
End Sub

Private Sub mnuQuitarTranslados_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 16/10/06
'*************************************************
Call modEdicion.Quitar_Translados
End Sub

Private Sub mnuQuitarTriggers_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Quitar_Triggers
End Sub

Private Sub mnuReAbrirMapa_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error GoTo ErrHandler
    If FileExist(Dialog.FileName, vbArchive) = False Then Exit Sub
    If MapInfo.Changed = 1 Then
        If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then
            modMapIO.GuardarMapa Dialog.FileName
        End If
    End If
    Call modMapIO.NuevoMapa
    modMapIO.CargarMapa Dialog.FileName
    DoEvents
    mnuReAbrirMapa.Enabled = True
    EngineRun = True
Exit Sub
ErrHandler:
End Sub

Private Sub mnuSalir_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Unload Me
End Sub

Private Sub mnuSuperficie_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Dim i As Byte
For i = 0 To 5
    If i <> 0 Then
        frmMain.SelectPanel(i).value = False
        Call VerFuncion(i, False)
    End If
Next
modPaneles.VerFuncion 0, True
End Sub

Private Sub mnuTodoMiniMap_Click()
    Dim Ratio As Integer
    Dim Mapas As String
    Dim Cantidad As Long
    Dim total As Long
   
    Cantidad = CInt(Val(InputBox("Hasta que mapa desea renderizar?", "Cantidad")))
   
   
    For total = 1 To Cantidad
        Mapas = (App.Path & "/Mapas/Mapa" & total & ".mcl")
        If FileExist(Mapas, vbArchive) Then
            Call modMapIO.NuevoMapa
            modMapIO.CargarMapa Mapas
            DoEvents
            EngineRun = True
            Call RenderScreen(UserPos.X, UserPos.Y, 0, 0)
            Call DrawBackBufferSurface
            Ratio = 1
            NombreImagen = total
            RenderToPicture Ratio, False
        End If
    Next total
    MsgBox "Terminado"
End Sub

Private Sub mnuTranslados_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Dim i As Byte
For i = 0 To 5
    If i <> 1 Then
        frmMain.SelectPanel(i).value = False
        Call VerFuncion(i, False)
    End If
Next
modPaneles.VerFuncion 1, True
End Sub

Private Sub mnuTriggers_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Dim i As Byte
For i = 0 To 5
    If i <> 5 Then
        frmMain.SelectPanel(i).value = False
        Call VerFuncion(i, False)
    End If
Next
modPaneles.VerFuncion 5, True
End Sub

Private Sub mnuUtilizarDeshacer_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 16/10/06
'*************************************************
mnuUtilizarDeshacer.Checked = (mnuUtilizarDeshacer.Checked = False)
End Sub

Private Sub mnuVerAgua_Click()
mnuVerAgua.Checked = (mnuVerAgua.Checked = False)
CBVerBloqueosA.value = mnuVerAgua.Checked
End Sub

Private Sub mnuVerAutomatico_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
mnuVerAutomatico.Checked = (mnuVerAutomatico.Checked = False)
End Sub

Private Sub mnuVerBloqueos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
mnuVerBloqueos.Checked = (mnuVerBloqueos.Checked = False)
CBVerBloqueosT.value = mnuVerBloqueos.Checked

End Sub

Private Sub mnuVerCapa1_Click()
mnuVerCapa1.Checked = (mnuVerCapa1.Checked = False)
CBVerCapa1.value = mnuVerCapa1.Checked
End Sub

Private Sub mnuVerCapa2_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
mnuVerCapa2.Checked = (mnuVerCapa2.Checked = False)
CBVerCapa2.value = mnuVerCapa2.Checked
End Sub

Private Sub mnuVerCapa3_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
mnuVerCapa3.Checked = (mnuVerCapa3.Checked = False)
CBVerCapa3.value = mnuVerCapa3.Checked
End Sub

Private Sub mnuVerCapa4_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
mnuVerCapa4.Checked = (mnuVerCapa4.Checked = False)
CBVerCapa4.value = mnuVerCapa4.Checked
End Sub

Private Sub mnuVerCuadricula_Click()
mnuVerCuadricula.Checked = (mnuVerCuadricula.Checked = False)
CBVerCuadricula.value = mnuVerCuadricula.Checked
End Sub

Private Sub mnuVerNPCs_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 26/05/06
'*************************************************
mnuVerNPCs.Checked = (mnuVerNPCs.Checked = False)
CBVerNpcs.value = mnuVerNPCs.Checked
End Sub

Private Sub mnuVerObjetos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 26/05/06
'*************************************************
mnuVerObjetos.Checked = (mnuVerObjetos.Checked = False)
CBVerObjetos.value = mnuVerObjetos.Checked
End Sub

Private Sub mnuVerTranslados_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 26/05/06
'*************************************************
mnuVerTranslados.Checked = (mnuVerTranslados.Checked = False)
CBVerTraslados.value = mnuVerTranslados.Checked

End Sub

Private Sub mnuVerTriggers_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
'cVerTriggers.value = (cVerTriggers.value = False)
mnuVerTriggers.Checked = (mnuVerTriggers.Checked = False)
CBVerTriggers.value = mnuVerTriggers.Checked
'frmMain.cVerTriggers.value = mnuVerTriggers.Checked
End Sub

Private Sub picRadar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
If X < 11 Then X = 11
If X > 89 Then X = 89
If Y < 10 Then Y = 10
If Y > 92 Then Y = 92
UserPos.X = X
UserPos.Y = Y
bRefreshRadar = True
End Sub

Private Sub picRadar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************
MiRadarX = X
MiRadarY = Y
End Sub



Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06 - GS
'Last modified: 20/11/07 - Loopzer
'*************************************************

Dim tX As Integer
Dim tY As Integer

If Not MapaCargado Then Exit Sub

If X <= MainViewShp.Left Or X >= MainViewShp.Left + MainViewWidth Or Y <= MainViewShp.Top Or Y >= MainViewShp.Top + MainViewHeight Then
    Exit Sub
End If

ConvertCPtoTP MainViewShp.Left, MainViewShp.Top, X, Y, tX, tY

'If Shift = 1 And Button = 2 Then PegarSeleccion tX, tY: Exit Sub
If Shift = 1 And Button = 1 Then
    Seleccionando = True
    SeleccionIX = tX '+ UserPos.X
    SeleccionIY = tY '+ UserPos.Y
Else
    ClickEdit Button, tX, tY
End If

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06 - GS
'*************************************************

FocoEnLista = False

Dim tX As Integer
Dim tY As Integer

'Make sure map is loaded
If Not MapaCargado Then Exit Sub
HotKeysAllow = True

'Make sure click is in view window
If X <= MainViewShp.Left Or X >= MainViewShp.Left + MainViewWidth Or Y <= MainViewShp.Top Or Y >= MainViewShp.Top + MainViewHeight Then
    Exit Sub
End If

ConvertCPtoTP MainViewShp.Left, MainViewShp.Top, X, Y, tX, tY

Rem yo - down
PosMouseX = tX
PosMouseY = tY
Rem up

POSX.Caption = "X: " & tX & " - Y: " & tY
If tX < 10 Or tY < 10 Or tX > 90 Or tY > 90 Then
    POSX.ForeColor = vbRed
Else
    POSX.ForeColor = vbWhite
End If
 If Shift = 1 And Button = 1 Then
    Seleccionando = True
    SeleccionFX = tX '+ TileX
    SeleccionFY = tY '+ TileY
Else
    ClickEdit Button, tX, tY
End If
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 16/10/06
'*************************************************

' Guardar configuración
Call WriteVar(IniPath & "WorldEditor.ini", "CONFIGURACION", "GuardarConfig", IIf(frmMain.mnuGuardarUltimaConfig.Checked, "1", "0"))
If frmMain.mnuGuardarUltimaConfig.Checked = True Then
    Call WriteVar(IniPath & "WorldEditor.ini", "CONFIGURACION", "ObjTranslado", Val(Cfg_TrOBJ))
    Call WriteVar(IniPath & "WorldEditor.ini", "CONFIGURACION", "UtilizarDeshacer", IIf(frmMain.mnuUtilizarDeshacer.Checked, "1", "0"))
    Call WriteVar(IniPath & "WorldEditor.ini", "CONFIGURACION", "AutoCapturarTrans", IIf(frmMain.mnuAutoCapturarTranslados.Checked, "1", "0"))
    Call WriteVar(IniPath & "WorldEditor.ini", "CONFIGURACION", "AutoCapturarSup", IIf(frmMain.mnuAutoCapturarSuperficie.Checked, "1", "0"))
    Call WriteVar(IniPath & "WorldEditor.ini", "PATH", "UltimoMapa", Dialog.FileName)
    Call WriteVar(IniPath & "WorldEditor.ini", "MOSTRAR", "LastPos", UserPos.X & "-" & UserPos.Y)
    Call WriteVar(IniPath & "WorldEditor.ini", "MOSTRAR", "ControlAutomatico", IIf(frmMain.mnuVerAutomatico.Checked, "1", "0"))
    Call WriteVar(IniPath & "WorldEditor.ini", "MOSTRAR", "Capa1", IIf(frmMain.mnuVerCapa1.Checked, "1", "0"))
    Call WriteVar(IniPath & "WorldEditor.ini", "MOSTRAR", "Capa2", IIf(frmMain.mnuVerCapa2.Checked, "1", "0"))
    Call WriteVar(IniPath & "WorldEditor.ini", "MOSTRAR", "Capa3", IIf(frmMain.mnuVerCapa3.Checked, "1", "0"))
    Call WriteVar(IniPath & "WorldEditor.ini", "MOSTRAR", "Capa4", IIf(frmMain.mnuVerCapa4.Checked, "1", "0"))
    Call WriteVar(IniPath & "WorldEditor.ini", "MOSTRAR", "Traslados", IIf(frmMain.mnuVerTranslados.Checked, "1", "0"))
    Call WriteVar(IniPath & "WorldEditor.ini", "MOSTRAR", "Bloqueos", IIf(frmMain.mnuVerBloqueos.Checked, "1", "0"))
    Call WriteVar(IniPath & "WorldEditor.ini", "MOSTRAR", "NPCs", IIf(frmMain.mnuVerNPCs.Checked, "1", "0"))
    Call WriteVar(IniPath & "WorldEditor.ini", "MOSTRAR", "Objetos", IIf(frmMain.mnuVerObjetos.Checked, "1", "0"))
    Call WriteVar(IniPath & "WorldEditor.ini", "MOSTRAR", "Triggers", IIf(frmMain.mnuVerTriggers.Checked, "1", "0"))
    Call WriteVar(IniPath & "WorldEditor.ini", "MOSTRAR", "Agua", IIf(frmMain.mnuVerAgua.Checked, "1", "0"))
    Call WriteVar(IniPath & "WorldEditor.ini", "MOSTRAR", "Cuadricula", IIf(frmMain.mnuVerCuadricula.Checked, "1", "0"))
End If

'Allow MainLoop to close program
If prgRun = True Then
    prgRun = False
    Cancel = 1
End If

End Sub





Private Sub pPaneles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*************************************************
'Author: Deut
'Last modified: 09/12/09
'*************************************************
FocoEnLista = False
End Sub

Private Sub SelectPanel_Click(Index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Dim i As Byte
For i = 0 To 5
    If i <> Index Then
        SelectPanel(i).value = False
        Call VerFuncion(i, False)
    End If
Next
If mnuAutoQuitarFunciones.Checked = True Then Call mnuQuitarFunciones_Click
Call VerFuncion(Index, SelectPanel(Index).value)
End Sub

Private Sub TimAutoGuardarMapa_Timer()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If mnuAutoGuardarMapas.Checked = True Then
    bAutoGuardarMapaCount = bAutoGuardarMapaCount + 1
    If bAutoGuardarMapaCount >= bAutoGuardarMapa Then
        If MapInfo.Changed = 1 Then ' Solo guardo si el mapa esta modificado
            modMapIO.GuardarMapa Dialog.FileName
        End If
        bAutoGuardarMapaCount = 0
    End If
End If
End Sub

Public Sub ObtenerNombreArchivo(ByVal Guardar As Boolean)
'*************************************************
'Author: Deut
'Last modified: 09/12/09
'*************************************************
With Dialog
    .Filter = "Mapas de FenixAO (*.mcl)|*.mcl"
    If Guardar Then
        .DialogTitle = "Guardar"
        .DefaultExt = ".txt"
        .FileName = vbNullString
        .flags = cdlOFNPathMustExist
        .ShowSave
    Else
        .DialogTitle = "Abrir"
        .FileName = vbNullString
        .flags = cdlOFNFileMustExist
        .ShowOpen
    End If
End With

End Sub
