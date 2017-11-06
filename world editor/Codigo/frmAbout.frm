VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Acerca de FXIIIWE"
   ClientHeight    =   4365
   ClientLeft      =   2340
   ClientTop       =   1815
   ClientWidth     =   4365
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":78BA
   ScaleHeight     =   3012.801
   ScaleMode       =   0  'User
   ScaleWidth      =   4098.96
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   300
      Left            =   2040
      TabIndex        =   3
      Top             =   720
      Width           =   930
   End
   Begin VB.Label lbltitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   435
      Left            =   2040
      TabIndex        =   2
      Top             =   240
      Width           =   810
   End
   Begin VB.Label lblCred 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Adaptado a FXII por GoDKeR y Rhynne"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   480
      Index           =   5
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   3075
   End
   Begin VB.Label lblCred 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":1462A
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
      Height          =   1125
      Index           =   3
      Left            =   1080
      TabIndex        =   0
      Top             =   1920
      Width           =   2565
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      DrawMode        =   9  'Not Mask Pen
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      Height          =   2025
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   4095
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
    Me.Caption = "Acerca de " & App.Title
    lblVersion.Caption = "v" & App.Major & "." & App.Minor & "." & App.Revision
    lbltitle.Caption = App.Title
End Sub

