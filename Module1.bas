Attribute VB_Name = "Module1"
Option Explicit

Public i As Integer
'declaraci�n de variable i para los distintos bucles for

Public cn As adodb.Connection
'establecemos la conexion con ADODB de manera p�blica en el m�dulo

Public stringbusca
'string en el cual recojeremos los distintos criterios de busqueda de la aplicacion

Public Sub conexion()
'funci�n que realiza la conexi�n con el connection string
 'casa
 cn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=proyecto1;Data Source=."
 'clase
 'cn.ConnectionString = "Provider=SQLOLEDB.1;Password=da4;Persist Security Info=True;User ID=alumno;Initial Catalog=proyecto1;Data Source=A21PROFE"
    cn.Open
End Sub
    
Public Sub volver()
'funcion que vuelve a habilitar el menu / formulario principal y cierra el formulario abierto en ese momento
    frm_ppal.Enabled = True
    frm_distribu.Enabled = True
    frm_ppal.Show
End Sub

Public Sub controlesoriginal()

    frm_ppal.Image1 = LoadPicture(".\imagenes\empresa.jpg")
    frm_ppal.Image3 = LoadPicture(".\imagenes\servicios.jpg")
    frm_ppal.Image2 = LoadPicture(".\imagenes\mantenimiento.jpg")
    
    frm_ppal.Image1.Visible = True
    frm_ppal.Image2.Visible = True
    frm_ppal.Image3.Visible = True

    frm_ppal.fra_1.Visible = False
    frm_ppal.fra_2.Visible = False
    frm_ppal.fra_3.Visible = False
    
End Sub



