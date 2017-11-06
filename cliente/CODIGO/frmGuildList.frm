VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmGuildList 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3420
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
   ScaleHeight     =   3420
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFundar 
      Caption         =   "Fundar"
      Height          =   360
      Left            =   240
      TabIndex        =   3
      Top             =   2880
      Width           =   990
   End
   Begin VB.CommandButton cmdSolicitud 
      Caption         =   "Solicitud"
      Height          =   360
      Left            =   3360
      TabIndex        =   2
      Top             =   2880
      Width           =   990
   End
   Begin MSComctlLib.ListView lstGuilds 
      Height          =   2055
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   3625
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label lblListaDe 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lista de clanes"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1050
   End
End
Attribute VB_Name = "frmGuildList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFundar_Click()
    Call WriteGuildFoundate
End Sub

Private Sub cmdSolicitud_Click()
    Call WriteGuildRequest(lstGuilds.SelectedItem.Text)
End Sub

Private Sub Form_Load()

Dim col As ColumnHeader
Set col = lstGuilds.ColumnHeaders.Add(, , "Nombre", lstGuilds.Width / 2)
Set col = lstGuilds.ColumnHeaders.Add(, , "Facción", lstGuilds.Width / 2)
End Sub
