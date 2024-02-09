VERSION 5.00
Begin VB.Form frmlistconsul 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consultas rápidas"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6960
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   6960
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdvolver 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      Picture         =   "frmlistconsul.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cmdconsulta 
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
      Index           =   5
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton cmdconsulta 
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
      Index           =   4
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2400
      Width           =   375
   End
   Begin VB.CommandButton cmdconsulta 
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
      Index           =   3
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton cmdconsulta 
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
      Index           =   2
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton cmdconsulta 
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
      Index           =   1
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cmdconsulta 
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
      Index           =   0
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   375
   End
   Begin VB.Line Line5 
      X1              =   2280
      X2              =   6600
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listados:"
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
      Left            =   1200
      TabIndex        =   12
      Top             =   240
      Width           =   975
   End
   Begin VB.Line Line4 
      X1              =   6600
      X2              =   6600
      Y1              =   360
      Y2              =   3360
   End
   Begin VB.Line Line3 
      X1              =   6600
      X2              =   360
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line2 
      X1              =   360
      X2              =   360
      Y1              =   3360
      Y2              =   360
   End
   Begin VB.Line Line1 
      X1              =   1080
      X2              =   360
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Talleres concertados de una empresa."
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
      Left            =   1080
      TabIndex        =   11
      Top             =   3000
      Width           =   3495
   End
   Begin VB.Label lblconsul5 
      BackColor       =   &H00E0E0E0&
      Caption         =   " Medios de transporte de una empresa."
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
      Left            =   1080
      TabIndex        =   10
      Top             =   2520
      Width           =   3375
   End
   Begin VB.Label lblconsul4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Empleados de una empresa."
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
      Left            =   1080
      TabIndex        =   9
      Top             =   2040
      Width           =   3135
   End
   Begin VB.Label lblconsul3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " Empresas q trabajan en Valencia."
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
      Left            =   1080
      TabIndex        =   8
      Top             =   1560
      Width           =   3495
   End
   Begin VB.Label lblconsul2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Empresas de tipo ""local""."
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
      Left            =   1080
      TabIndex        =   7
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label lblconsul1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Empresas con mas de 60000 euros de volumen de negocio."
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
      Left            =   1080
      TabIndex        =   6
      Top             =   600
      Width           =   5295
   End
End
Attribute VB_Name = "frmlistconsul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdconsulta_Click(Index As Integer)
    Select Case Index
    
        Case 0
            frmlistconsul.Hide
            frmconsultas.Show
            frmconsultas.Caption = "Consulta rápida Nº1"
            frmconsultas.lblenunciado.Caption = "Listado de empresas con mas de 60.000 euros de volumen de negocio."
            frmconsultas.DComboempresa.Visible = False
            frmconsultas.Adodc1.RecordSource = "Select nom_emp,domicilio,nom_dir,tip_empres,num_emple from empresa WHERE vol_neg = '>60000' ORDER BY nom_emp"
            
            frmconsultas.Adodc1.Refresh
            frmconsultas.DataGrid1.Visible = True
            
            frmconsultas.DataGrid1.Columns(0).Caption = "Empresa"
            frmconsultas.DataGrid1.Columns(1).Caption = "Dirección"
            frmconsultas.DataGrid1.Columns(2).Caption = "Nombre del director"
            frmconsultas.DataGrid1.Columns(3).Caption = "Tipo"
            frmconsultas.DataGrid1.Columns(4).Caption = "NºEmpleados"
            
            frmconsultas.DataGrid1.Columns(0).Width = 1600
            frmconsultas.DataGrid1.Columns(1).Width = 2200
            frmconsultas.DataGrid1.Columns(2).Width = 2200
            frmconsultas.DataGrid1.Columns(3).Width = 1100
            frmconsultas.DataGrid1.Columns(4).Width = 1600
            
            For i = 0 To 4
                frmconsultas.DataGrid1.Columns(i).AllowSizing = False
            Next
            
        Case 1
            frmlistconsul.Hide
            frmconsultas.Show
            frmconsultas.Caption = "Consulta rápida Nº2"
            frmconsultas.lblenunciado.Caption = "Listado de empresas de tipo 'local'."
            frmconsultas.DComboempresa.Visible = False
            frmconsultas.Adodc1.RecordSource = "Select nom_emp,domicilio,nom_dir,tip_empres,num_emple from empresa WHERE tip_empres= 'local' ORDER BY nom_emp"
            
            frmconsultas.Adodc1.Refresh
            frmconsultas.DataGrid1.Visible = True
            
            frmconsultas.DataGrid1.Columns(0).Caption = "Empresa"
            frmconsultas.DataGrid1.Columns(1).Caption = "Dirección"
            frmconsultas.DataGrid1.Columns(2).Caption = "Nombre del director"
            frmconsultas.DataGrid1.Columns(3).Caption = "Tipo"
            frmconsultas.DataGrid1.Columns(4).Caption = "NºEmpleados"
            
            frmconsultas.DataGrid1.Columns(0).Width = 1600
            frmconsultas.DataGrid1.Columns(1).Width = 2200
            frmconsultas.DataGrid1.Columns(2).Width = 2200
            frmconsultas.DataGrid1.Columns(3).Width = 1100
            frmconsultas.DataGrid1.Columns(4).Width = 1600
            
            For i = 0 To 4
                frmconsultas.DataGrid1.Columns(i).AllowSizing = False
            Next
            
        Case 2
            frmlistconsul.Hide
            frmconsultas.Show
            frmconsultas.Caption = "Consulta rápida Nº3"
            frmconsultas.lblenunciado.Caption = "Listado de empresas que trabajan en Valencia."
            frmconsultas.DComboempresa.Visible = False
            frmconsultas.Adodc1.RecordSource = "Select nom_emp,domicilio,nom_dir,tip_empres,num_emple from empresa INNER JOIN(m_trans INNER JOIN carga ON carga.cod_mtrans = m_trans.codigo) ON m_trans.cod_emp = empresa.codigo WHERE carga.destino = 'Valencia' or carga.proced = 'Valencia';"
            
            frmconsultas.Adodc1.Refresh
            frmconsultas.DataGrid1.Visible = True
            
            frmconsultas.DataGrid1.Columns(0).Caption = "Empresa"
            frmconsultas.DataGrid1.Columns(1).Caption = "Dirección"
            frmconsultas.DataGrid1.Columns(2).Caption = "Nombre del director"
            frmconsultas.DataGrid1.Columns(3).Caption = "Tipo"
            frmconsultas.DataGrid1.Columns(4).Caption = "NºEmpleados"
            
            frmconsultas.DataGrid1.Columns(0).Width = 1600
            frmconsultas.DataGrid1.Columns(1).Width = 2200
            frmconsultas.DataGrid1.Columns(2).Width = 2200
            frmconsultas.DataGrid1.Columns(3).Width = 1100
            frmconsultas.DataGrid1.Columns(4).Width = 1600
            
            For i = 0 To 4
                frmconsultas.DataGrid1.Columns(i).AllowSizing = False
            Next
            
        Case 3
            frmlistconsul.Hide
            frmconsultas.Show
            frmconsultas.Caption = "Consulta rápida Nº4"
            frmconsultas.lblenunciado.Caption = "Listado de empleados de una empresa."
            frmconsultas.lblselec.Visible = True
            frmconsultas.DComboempresa.Visible = True
            frmconsultas.DataGrid1.Visible = False
            
            frmconsultas.Adodc1.Refresh
            frmconsultas.DComboempresa.Text = ""
            frmconsultas.fra_leyenda.Visible = True
            
        Case 4
            frmlistconsul.Hide
            frmconsultas.Show
            frmconsultas.Caption = "Consulta rápida Nº5"
            frmconsultas.lblenunciado.Caption = "Listado de medios de transporte de una empresa."
            frmconsultas.lblselec.Visible = True
            frmconsultas.DComboempresa.Visible = True
            frmconsultas.Adodc1.RecordSource = "Select *  from m_trans inner join empresa ON m_trans.cod_emp = empresa.codigo Where m_trans.cod_emp = '" & frmconsultas.DComboempresa.BoundText & "'"
            frmconsultas.DataGrid1.Visible = False
            
            frmconsultas.DComboempresa.Text = ""
            
        Case 5
            frmlistconsul.Hide
            frmconsultas.Show
            frmconsultas.Caption = "Consulta rápida Nº6"
            frmconsultas.lblenunciado.Caption = "Listado de talleres concertados por una empresa."
            frmconsultas.lblselec.Visible = True
            frmconsultas.DComboempresa.Visible = True
            frmconsultas.DataGrid1.Visible = False
            frmconsultas.Adodc1.RecordSource = "Select * from c_mante inner join empresa ON empresa.codigo = c_mante.cod_emp ORDER BY c_mante.nombre"
            
            frmconsultas.DComboempresa.Text = ""
    End Select
End Sub

Private Sub cmdconsulta_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
    Case 0
        cmdconsulta(Index).BackColor = &HEEEEEE
    Case 1
        cmdconsulta(Index).BackColor = &HEEEEEE
    Case 2
        cmdconsulta(Index).BackColor = &HEEEEEE
    Case 3
        cmdconsulta(Index).BackColor = &HEEEEEE
    Case 4
        cmdconsulta(Index).BackColor = &HEEEEEE
    Case 5
        cmdconsulta(Index).BackColor = &HEEEEEE
End Select
End Sub

Private Sub cmdconsulta_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
    Case 0
        cmdconsulta(Index).BackColor = &H808080
    Case 1
        cmdconsulta(Index).BackColor = &H808080
    Case 2
        cmdconsulta(Index).BackColor = &H808080
    Case 3
        cmdconsulta(Index).BackColor = &H808080
    Case 4
        cmdconsulta(Index).BackColor = &H808080
    Case 5
        cmdconsulta(Index).BackColor = &H808080
End Select
End Sub

Private Sub cmdvolver_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmlistconsul.Hide
    Call volver
    cmdvolver.BackColor = &HEEEEEE
End Sub

Private Sub cmdvolver_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdvolver.BackColor = &H8000000F
End Sub

