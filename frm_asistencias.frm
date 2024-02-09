VERSION 5.00
Begin VB.Form frm_asistencias 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asistencias"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9690
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   9690
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_volver 
      BackColor       =   &H00808080&
      Height          =   615
      Left            =   8640
      Picture         =   "frm_asistencias.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
      Width           =   735
   End
End
Attribute VB_Name = "frm_asistencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmd_volver_Click()
    frm_asistencias.Hide
    Call volver
End Sub
