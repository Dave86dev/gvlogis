VERSION 5.00
Begin VB.Form frm_list 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listados "
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9705
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   9705
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_imp2 
      BackColor       =   &H00808080&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton cmd_imp1 
      BackColor       =   &H00808080&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton cmd_list4 
      BackColor       =   &H00808080&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton cmd_list3 
      BackColor       =   &H00808080&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton cmd_list2 
      BackColor       =   &H00808080&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton cmd_list1 
      BackColor       =   &H00808080&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton cmd_volver 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   7920
      Picture         =   "frm_list.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label lbl6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Impresión de los transportes realizados en una fecha concreta."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   13
      Top             =   3000
      Width           =   6255
   End
   Begin VB.Label lbl5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Impresión de las empresas existentes en la aplicación (nombre,domicilio...tipo..)."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   12
      Top             =   2520
      Width           =   6255
   End
   Begin VB.Label lbltip4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dada una fecha, que transportes se han realizado (medio de transporte, carga e itinerario)."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   9
      Top             =   2040
      Width           =   6135
   End
   Begin VB.Label lbllist3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dado un medio de transporte, que itinerarios y/o cargas ha realizado."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   1680
      Width           =   7215
   End
   Begin VB.Label lbllist2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Todas las cargas con un lugar de procedencia y/o un destino."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   1200
      Width           =   7215
   End
   Begin VB.Label lbllist1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dado un itinerario, que medios de transporte lo han realizado y en que fechas."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   720
      Width           =   7335
   End
   Begin VB.Line Line5 
      X1              =   360
      X2              =   1440
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label lbllistadosdatos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listados de datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.Line Line4 
      X1              =   9000
      X2              =   9000
      Y1              =   3720
      Y2              =   360
   End
   Begin VB.Line Line3 
      X1              =   9000
      X2              =   3480
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line2 
      X1              =   9000
      X2              =   360
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   360
      Y1              =   360
      Y2              =   3720
   End
End
Attribute VB_Name = "frm_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_imp1_Click()
    'impresión 1..
    DataReport1.Show
End Sub

Private Sub cmd_imp2_Click()
    frm_listado2.Show
    frm_list.Enabled = False
End Sub

Private Sub cmd_list1_Click()
    frm_listadosej.Caption = "Listado Nº1"
    frm_list.Hide
    frm_listadosej.Show
    frm_listadosej.lblenunciado.Caption = "Dado un itinerario, que medios de transporte lo han realizado y en que fechas."
    
    frm_listadosej.DataCombo1.Visible = True
    frm_listadosej.DataGrid1.Visible = False
    frm_listadosej.Adodc2.RecordSource = "Select codigo,(punt_part+'  -   '+punt_dest) as itinerario from itinerario"
    frm_listadosej.Adodc2.Refresh
    frm_listadosej.DataCombo1.ListField = "itinerario"
    frm_listadosej.DataCombo1.BoundColumn = "codigo"
End Sub

Private Sub cmd_list2_Click()
    frm_listadosej.Caption = "Listado Nº2"
    frm_list.Hide
    frm_listadosej.Show
    frm_listadosej.lblenunciado.Caption = "Todas las cargas con un lugar de procedencia y/o un destino."
    
    frm_listadosej.Adodc1.RecordSource = "Select * from carga"
    frm_listadosej.Adodc1.Refresh
    
    frm_listadosej.DataGrid1.Visible = True
    frm_listadosej.DataGrid1.Refresh
    frm_listadosej.DataCombo1.Visible = False
    
    frm_listadosej.DataGrid1.Columns(0).Caption = "Codigo carga"
    frm_listadosej.DataGrid1.Columns(1).Caption = "Descripcion"
    frm_listadosej.DataGrid1.Columns(2).Caption = "Procedencia"
    frm_listadosej.DataGrid1.Columns(3).Caption = "Destino"
    frm_listadosej.DataGrid1.Columns(4).Caption = "Valor en Euros"
    frm_listadosej.DataGrid1.Columns(5).Visible = False
    
    frm_listadosej.DataGrid1.Columns(0).Width = 1500
    frm_listadosej.DataGrid1.Columns(4).Width = 1900
    
End Sub

Private Sub cmd_list3_Click()
    frm_listadosej.Caption = "Listado Nº3"
    frm_list.Hide
    frm_listadosej.Show
    frm_listadosej.lblenunciado.Caption = "Dado un medio de transporte, que itinerarios y/o cargas ha realizado."
    
    
    frm_listadosej.DataCombo1.Visible = True
    frm_listadosej.DataGrid1.Visible = False
    frm_listadosej.Adodc2.RecordSource = "Select * from m_trans"
    frm_listadosej.Adodc2.Refresh
    frm_listadosej.DataCombo1.ListField = "marca"
    frm_listadosej.DataCombo1.BoundColumn = "codigo"
    
End Sub

Private Sub cmd_list4_Click()
    frm_listadosej.Caption = "Listado Nº4"
    frm_list.Hide
    frm_listadosej.Show
    frm_listadosej.lblenunciado.Caption = "Dada una fecha, que transportes se han realizado (medio de transporte, carga e itinerario)."
    
    frm_listadosej.cmd_confirmar.Visible = True
    frm_listadosej.Calendar1.Visible = True
    frm_listadosej.DataGrid1.Visible = False
    frm_listadosej.DataCombo1.Visible = False
    
End Sub

Private Sub cmd_volver_Click()
    Call volver
End Sub

Private Sub Form_Load()

End Sub
