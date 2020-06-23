VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clConnSistemas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public Function oAbreConexion(ByVal strTipoBase, ByVal strServidor, ByVal strNombreBase, ByVal strUsuario, ByVal strClave) As ADODB.Connection
Dim objConexion As New ADODB.Connection
Dim sConn As String
Select Case CStr(strTipoBase)
   Case "3"
      sConn = "DSN=" & strNombreBase & ";UID=" & strUsuario & ";PWD=" & strClave
   Case "4"
      sConn = "DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & strServidor & ";DATABASE=" & strNombreBase & ";UID=" & strUsuario & ";PWD=" & strClave & "; "
   Case "5"
      sConn = "DRIVER={MySQL ODBC 5.1 Driver};SERVER=" & strServidor & ";DATABASE=" & strNombreBase & ";UID=" & strUsuario & ";PWD=" & strClave & "; "
      
End Select
objConexion.Open sConn
Set oAbreConexion = objConexion
Set objConexion = Nothing
End Function
