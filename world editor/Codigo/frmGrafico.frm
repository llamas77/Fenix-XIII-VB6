VERSION 5.00
Begin VB.Form frmGrafico 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Grafico"
   ClientHeight    =   5385
   ClientLeft      =   600
   ClientTop       =   7590
   ClientWidth     =   6870
   Icon            =   "frmGrafico.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   5355
      Left            =   0
      ScaleHeight     =   355
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   455
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6855
      Begin VB.PictureBox ShowPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   5355
         Left            =   0
         ScaleHeight     =   355
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   455
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   6855
      End
   End
End
Attribute VB_Name = "frmGrafico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Deactivate()
Me.Visible = False
End Sub


