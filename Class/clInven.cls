VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clInven"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Dim mObj As clConnSistemas
Private DBConn As New ADODB.Connection
Private cmdTemp As New ADODB.Command
Dim strSql1 As String
Dim Rs1 As New ADODB.Recordset
 
 
Private Type parametroSql
    Nombre        As String
    Tipo          As DataTypeEnum
    Direccion     As ParameterDirectionEnum
    Tamanio       As ADO_LONGPTR
    Valor         As Variant
End Type


Dim miArray() As String
Private Sub Class_Initialize()
   Set mObj = New clConnSistemas
   DBConn.ConnectionTimeout = 500
   Set DBConn = mObj.oAbreConexion("5", mIPServer, "Inventario", "ssvv", "ssvv")
   'Set DBConn = mObj.oAbreConexion("4", "192.168.2.155", "regnov", "ssvv", "ssvv")
   'Set DBConn = mObj.oAbreConexion("4", "10.10.20.11", "regnov", "root", "123456")
End Sub

Private Sub Class_Terminate()
   DBConn.Close
   Set DBConn = Nothing
   Set cmdTemp = Nothing
   Set Rs1 = Nothing
   Set mObj = Nothing
End Sub

Private Sub sAsignCmd()
   cmdTemp.CommandText = strSql1
   cmdTemp.CommandType = 1
   cmdTemp.CommandTimeout = 36000
   Set cmdTemp.ActiveConnection = DBConn
End Sub

Private Sub sAsignCmd_SP(ByVal cmd As ADODB.Command)
   cmd.CommandText = strSql1
   cmd.CommandType = adCmdStoredProc
   cmd.CommandTimeout = 36000
   Set cmd.ActiveConnection = DBConn
End Sub

Public Function oTabla(ByVal pNameTabla As String, ByVal pOrden As String) As ADODB.Recordset
   strSql1 = "SELECT * FROM " & pNameTabla & " " & pOrden
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oTabla = Rs1
End Function
Public Function oStoreProcedure(ByVal pNameStore As String) As ADODB.Recordset
   Dim cmd As New ADODB.Command
   strSql1 = pNameStore
   sAsignCmd_SP cmd
   Set Rs1 = cmd.Execute
   Set cmd = Nothing
   Set oStoreProcedure = Rs1
End Function

Friend Function oStoreProcedureWithParam(ByVal pNameStore As String, ByRef pArrayParamSql() As parametroSql) As ADODB.Recordset
   
   Dim cmd As New ADODB.Command
   Dim i As Integer
   strSql1 = pNameStore
   sAsignCmd_SP cmd
   
   For i = 0 To UBound(pArrayParamSql)
      cmd.Parameters.Append cmd.CreateParameter(pArrayParamSql(i).Nombre, _
                                                pArrayParamSql(i).Tipo, _
                                                pArrayParamSql(i).Direccion, _
                                                pArrayParamSql(i).Tamanio, _
                                                pArrayParamSql(i).Valor)
   Next
   
   
   
   Set Rs1 = cmd.Execute
   Set cmd = Nothing
   Set oStoreProcedureWithParam = Rs1
End Function

Public Function xStoreProcedure(ByVal pNameStore As String) As ADODB.Recordset
   Dim cmd As New ADODB.Command
   strSql1 = pNameStore
   sAsignCmd_SP cmd
   Set Rs1 = cmd.Execute
   Set cmd = Nothing
   Set xStoreProcedure = Rs1
End Function

Friend Sub xStoreProcedureWithParam(ByVal pNameStore As String, ByRef pArrayParamSql() As parametroSql)
   Dim cmd As New ADODB.Command
   Dim i As Integer
   strSql1 = pNameStore
   sAsignCmd_SP cmd
   For i = 0 To UBound(pArrayParamSql)
      cmd.Parameters.Append cmd.CreateParameter(pArrayParamSql(i).Nombre, _
                                                pArrayParamSql(i).Tipo, _
                                                pArrayParamSql(i).Direccion, _
                                                pArrayParamSql(i).Tamanio, _
                                                pArrayParamSql(i).Valor)
   Next
   'On Error Resume Next
   cmd.Execute
   Set cmd = Nothing
End Sub



Friend Sub xStoreProcedureWithParam_Transaccional(ByVal pNameStore As String, ByRef pArrayParamSql() As parametroSql, ByRef cmd As ADODB.Command)
  'Dim cmd As New ADODB.Command
   Dim i As Integer
   strSql1 = pNameStore
   cmd.CommandText = strSql1
   'sAsignCmd_SP cmd
   
   
     'For mI = FlexProduct.Rows To 3 Step -1
   
   
   For i = cmd.Parameters.Count - 1 To 0 Step -1
      cmd.Parameters.Delete (i)
   Next
   
   For i = 0 To UBound(pArrayParamSql)
      cmd.Parameters.Append cmd.CreateParameter(pArrayParamSql(i).Nombre, _
                                                pArrayParamSql(i).Tipo, _
                                                pArrayParamSql(i).Direccion, _
                                                pArrayParamSql(i).Tamanio, _
                                                pArrayParamSql(i).Valor)
   Next
   'On Error Resume Next
   'cmd.Execute
   'Set cmd = Nothing
End Sub


Public Function oTablaNull(ByVal mTable As String) As ADODB.Recordset
   strSql1 = "SELECT * FROM " & mTable & " WHERE Fecha_Baja IS NULL"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oTablaNull = Rs1
End Function

Public Function oTablaDina(ByVal pNameTabla As String, ByVal pOrden As String) As ADODB.Recordset
   Dim mRs2 As New ADODB.Recordset
   mRs2.CursorType = adOpenDynamic
   strSql1 = "SELECT * FROM " & pNameTabla & " " & pOrden
   sAsignCmd
   mRs2.Open strSql1, DBConn
   Set oTablaDina = mRs2
End Function

Public Function bExistDatoTabla(ByVal pTabla As String, ByVal pWhere As String) As Boolean
   strSql1 = "select * from " & pTabla & " where " & pWhere
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   bExistDatoTabla = Not Rs1.EOF
End Function

Public Function sTablaDescr(ByVal pNameTabla As String, ByVal pWhereCodigo As String, ByVal pCampo As Integer) As String
   strSql1 = "select * from " & pNameTabla & " where " & pWhereCodigo
   sAsignCmd
   sTablaDescr = ""
   Set Rs1 = cmdTemp.Execute
   If Not Rs1.EOF Then
      sTablaDescr = Rs1.Fields(pCampo)
   End If
End Function


















Public Function iCountTabla(ByVal pTabla As String, ByVal pFecha1 As String, ByVal pFecha2 As String, _
                            ByVal pFrom As String, ByVal pWhere As String) As Integer
   iCountTabla = 0
   strSql1 = "select count(" & pFrom & ") as total from " & pTabla & " where " _
      & "fecha between '" & Format(pFecha1 & " " & "00:00:00", "yyyy-mm-dd hh:mm:ss") & "' and '" & Format(pFecha2 & " " & "23:59:59", "yyyy-mm-dd hh:mm:ss") & "' "
   If pWhere <> "" Then
      strSql1 = strSql1 & " " & pWhere
   End If
   
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   If Not Rs1.EOF Then
      iCountTabla = Rs1!Total
   End If
End Function



Public Function oMaxTabla(ByVal pTabla As String, ByVal pCampo As String, ByVal pCondicion As String) As ADODB.Recordset
   strSql1 = "select max(" & pCampo & ") as total from " & pTabla
   If pCondicion <> "" Then
      strSql1 = strSql1 & " where " & pCondicion
   End If
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oMaxTabla = Rs1
End Function



Public Function xInsTabla(ByVal pTabla As String, ByVal pCampos As String, ByVal pValues As String)
   strSql1 = "insert into " & pTabla & " " & pCampos & " values " & pValues
   sAsignCmd
   cmdTemp.Execute
End Function

Public Function xUpTablaDescr(ByVal pTabla As String, ByVal pSets As String, ByVal pWhere As String)
   strSql1 = "update " & pTabla & " set " & pSets & " where " & pWhere
   sAsignCmd
   cmdTemp.Execute
End Function




Public Function sFechaMySQL() As String
   strSql1 = "select current_timestamp as fecha from dual "
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   If Not Rs1.EOF Then
      sFechaMySQL = Rs1!Fecha
   End If
End Function


Public Function xUpWazeDateUp(ByVal pCodAlfa As String)
      
   strSql1 = "update d_waze set fecha_upd=date_add(current_timestamp, interval 1 hour) where codalfa='" & pCodAlfa & "'"
   sAsignCmd
   On Error Resume Next
   cmdTemp.Execute
   If Err.Description <> "" Then
      MsgBox Err.Description
   End If
   
End Function



'===========================================================================================
'=======               Llamadas a Store Procedures MySQL: Schema Inventario         ========
'===========================================================================================

'---------------------------------------------NUEVO



Public Function xInsBodega(ByVal pCodigo As String, ByVal pDescr As String, ByVal pCodAlmacen As String)
   strSql1 = "insert into Bodegas (Codigo, Descripcion, CodAlmacen) values ('" & pCodigo & "','" & pDescr & "','" & pCodAlmacen & "')"
   sAsignCmd
   cmdTemp.Execute
End Function



Public Function xInsUbicacion(ByVal pCodigo As String, ByVal pDescr As String, ByVal pCodBodega As String)
   strSql1 = "insert into Ubicaciones (Codigo, Descripcion, CodBodega) values ('" & pCodigo & "','" & pDescr & "','" & pCodBodega & "')"
   sAsignCmd
   cmdTemp.Execute
End Function



Public Function xUpdBodega(ByVal pCodigo As String, ByVal pDescr As String, ByVal pCodAlmacen As String)
   
   'pStock = Replace(pStock, ",", ".")
   'pStockMinimo = Replace(pStockMinimo, ",", ".")
   
   strSql1 = "Update Bodegas set Descripcion='" & pDescr & "', " & _
               " CodAlmacen ='" & pCodAlmacen & "' " & _
             " where Codigo='" & pCodigo & "'"
   
   sAsignCmd
   cmdTemp.Execute
End Function


Public Function xUpdUbicacion(ByVal pCodigo As String, ByVal pDescr As String, ByVal pCodBodega As String)

   strSql1 = "Update Ubicaciones set Descripcion='" & pDescr & "', " & _
               " CodBodega ='" & pCodBodega & "' " & _
             " where Codigo='" & pCodigo & "'"
   
   sAsignCmd
   cmdTemp.Execute
End Function

Public Function xUpdBaJaBodega(ByVal pCodigo As String)
   
   strSql1 = "Update Bodegas set Fecha_Baja= NOW()" & _
   " where Codigo='" & pCodigo & "'"
   
   sAsignCmd
   cmdTemp.Execute
End Function

Public Function xUpdBaJaUbicacion(ByVal pCodigo As String)
   
   strSql1 = "Update Ubicaciones set Fecha_Baja= NOW()" & _
   " where Codigo='" & pCodigo & "'"
   
   sAsignCmd
   cmdTemp.Execute
End Function


Public Function oBodegas() As ADODB.Recordset
  strSql1 = "SELECT B.Codigo, B.Descripcion, A.Descripcion AS Almacen, B.Fecha_Baja " _
   & " From " _
   & " Bodegas B  " _
   & " Inner Join  " _
   & " Almacenes A ON B.CodAlmacen = A.Codigo  " _
   & " ORDER BY A.Descripcion, B.Descripcion "

   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oBodegas = Rs1
End Function

Public Function oUbicaciones() As ADODB.Recordset
   
 strSql1 = "  SELECT U.Codigo, U.Descripcion, A.Descripcion As Almacen, B.Descripcion as Bodega, U.Fecha_Baja " _
  & " FROM Ubicaciones U " _
  & " Inner Join " _
  & " Bodegas B ON U.CodBodega = B.Codigo " _
  & " Inner Join " _
  & " Almacenes A ON B.CodAlmacen = A.Codigo " _
  & " ORDER BY A.Descripcion, B.Descripcion, U.Descripcion; "

   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oUbicaciones = Rs1
End Function


Public Function xInsProducto(ByVal pCodigo As String, ByVal pDescr As String, ByVal pCodigoSap As String, ByVal pCodUnidadMedida As String)
   strSql1 = "insert into Producto (Codigo, Descripcion, CodigoSAP, CodUnidadMedida) values ('" & pCodigo & "','" & pDescr & "','" & pCodigoSap & "','" & pCodUnidadMedida & "')"
   sAsignCmd
   cmdTemp.Execute
End Function

Public Function xUpdProducto(ByVal pCodigo As String, ByVal pDescr As String, ByVal pCodigoSap As String, ByVal pCodUnidadMedida As String)
 
   strSql1 = "Update Producto set Descripcion='" & pDescr & "', " & _
         " CodigoSap ='" & pCodigoSap & "', " & _
         " CodUnidadMedida ='" & pCodUnidadMedida & "' " & _
   " where Codigo='" & pCodigo & "'"
   
   sAsignCmd
   cmdTemp.Execute
End Function

Public Function xUpdBaJaProducto(ByVal pCodigo As String)
   
   strSql1 = "Update Producto set Fecha_Baja= NOW()" & _
   " where Codigo='" & pCodigo & "'"
   
   sAsignCmd
   cmdTemp.Execute
End Function


Public Function oProductos() As ADODB.Recordset
  strSql1 = "SELECT P.Codigo, P.CodigoSap, P.Descripcion, U.Descripcion AS UnidadMedida, P.Fecha_Baja  " _
   & " From " _
   & " Producto P  " _
   & " Inner Join  " _
   & " UnidadMedida U ON P.CodUnidadMedida = U.Codigo  " _
   & " ORDER BY Codigo "

   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oProductos = Rs1
End Function


Public Function xInsStockMinimo(ByVal pCodProducto As String, ByVal pCodBodega As String, ByVal pStockMinimo As String)
   pStockMinimo = Replace(pStockMinimo, ",", ".")
   
   strSql1 = "insert into StocksMinimos (CodProducto, CodBodega, Stock_Min) values ('" & pCodProducto & "','" & pCodBodega & "'," & pStockMinimo & ")"
   sAsignCmd
   cmdTemp.Execute
End Function


Public Function xUpdStockMinimo(ByVal pCodProducto As String, ByVal pCodBodega As String, ByVal pStockMinimo As String)
 
   pStockMinimo = Replace(pStockMinimo, ",", ".")
   
   strSql1 = "UPDATE StocksMinimos " & _
               " SET Stock_Min =" & pStockMinimo & _
               " WHERE CodProducto='" & pCodProducto & "' " & _
               " AND CodBodega='" & pCodBodega & "' "
   
   sAsignCmd
   cmdTemp.Execute
End Function

Public Function xDelStockMinimo(ByVal pCodProducto As String, ByVal pCodBodega As String)
   
   strSql1 = "DELETE FROM StocksMinimos " & _
               " WHERE CodProducto='" & pCodProducto & "' " & _
               " AND CodBodega='" & pCodBodega & "' "
   
   sAsignCmd
   cmdTemp.Execute
End Function

Public Function oEjecutarSelectDina(ByVal pSentencia As String) As ADODB.Recordset
   Dim mRs2 As New ADODB.Recordset
   mRs2.CursorType = adOpenDynamic
   strSql1 = pSentencia
   sAsignCmd
   mRs2.Open strSql1, DBConn
   Set oEjecutarSelectDina = mRs2
End Function

Public Function oEjecutarSelect(ByVal pSentencia As String) As ADODB.Recordset
   strSql1 = pSentencia
   sAsignCmd
   Set oEjecutarSelect = cmdTemp.Execute
End Function

Public Function sEjecutarSelectDescr(ByVal pSentencia As String, ByVal pCampo As Integer) As String
   strSql1 = pSentencia
   sAsignCmd
   sEjecutarSelectDescr = ""
   Set Rs1 = cmdTemp.Execute
   If Not Rs1.EOF Then
      sEjecutarSelectDescr = Rs1.Fields(pCampo)
   Else
      sEjecutarSelectDescr = ""
   End If
End Function

'---------------------------------------
'Sp: getStockXBodegaConFiltroProducto
'---------------------------------------
Public Function getStockXBodegaConFiltroProducto(ByVal CodBodega As String, ByVal textoFiltroProducto As String) As ADODB.Recordset
   Dim paramCodAlfa As parametroSql
   Dim aParamSql(1) As parametroSql

   paramCodAlfa.Nombre = "vCodBodega"
   paramCodAlfa.Tipo = adVarChar
   paramCodAlfa.Direccion = adParamInput
   paramCodAlfa.Tamanio = 4
   paramCodAlfa.Valor = CodBodega
   aParamSql(0) = paramCodAlfa
  
  
   paramCodAlfa.Nombre = "textoFiltroProducto"
   paramCodAlfa.Tipo = adChar
   paramCodAlfa.Direccion = adParamInput
   paramCodAlfa.Tamanio = 100
   paramCodAlfa.Valor = textoFiltroProducto
   aParamSql(1) = paramCodAlfa
  
   Set getStockXBodegaConFiltroProducto = oStoreProcedureWithParam("getStockXBodegaConFiltroProducto", aParamSql)
End Function


'---------------------------------------
'Sp: getStockXUbicacionConFiltroProducto
'---------------------------------------
Public Function getStockXUbicacionConFiltroProducto(ByVal CodUbicacion As String, ByVal textoFiltroProducto As String) As ADODB.Recordset
   Dim paramCodAlfa As parametroSql
   Dim aParamSql(1) As parametroSql

   paramCodAlfa.Nombre = "vCodUbicacion"
   paramCodAlfa.Tipo = adVarChar
   paramCodAlfa.Direccion = adParamInput
   paramCodAlfa.Tamanio = 4
   paramCodAlfa.Valor = CodUbicacion
   aParamSql(0) = paramCodAlfa
  
   paramCodAlfa.Nombre = "textoFiltroProducto"
   paramCodAlfa.Tipo = adChar
   paramCodAlfa.Direccion = adParamInput
   paramCodAlfa.Tamanio = 100
   paramCodAlfa.Valor = textoFiltroProducto
   aParamSql(1) = paramCodAlfa
  
   Set getStockXUbicacionConFiltroProducto = oStoreProcedureWithParam("getStockXUbicacionConFiltroProducto", aParamSql)
End Function


'---------------------------------------
'Sp: insEgreso
'---------------------------------------
Public Sub insEgreso(ByVal CodProducto As String, ByVal CodUbicacion As String, ByVal Cantidad As Double, ByVal NroVale As String, ByVal CodTipoVale As String, ByVal CodUsuario As String, ByRef cmd As ADODB.Command)
   Dim parametroSql As parametroSql
   Dim aParamSql(5) As parametroSql
   
   parametroSql.Nombre = "vCodProducto"
   parametroSql.Tipo = adChar
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 6
   parametroSql.Valor = CodProducto
   aParamSql(0) = parametroSql
   
   parametroSql.Nombre = "vCodUbicacion"
   parametroSql.Tipo = adChar
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 4
   parametroSql.Valor = CodUbicacion
   aParamSql(1) = parametroSql
   
   parametroSql.Nombre = "vCantidad"
   parametroSql.Tipo = adDouble
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 8
   parametroSql.Valor = Cantidad
   aParamSql(2) = parametroSql
   
   parametroSql.Nombre = "vNroVale"
   parametroSql.Tipo = adVarChar
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 9
   parametroSql.Valor = NroVale
   aParamSql(3) = parametroSql
   
   parametroSql.Nombre = "vCodTipoVale"
   parametroSql.Tipo = adChar
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 1
   parametroSql.Valor = CodTipoVale
   aParamSql(4) = parametroSql
   
   parametroSql.Nombre = "vCodUsuario"
   parametroSql.Tipo = adVarChar
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 25
   parametroSql.Valor = CodUsuario
   aParamSql(5) = parametroSql
   
   xStoreProcedureWithParam_Transaccional "insEgreso", aParamSql, cmd
 End Sub


Public Sub xInsEgreso(ByRef pvEgresosCodProducto() As String, ByRef pvEgresosCodUbicacion() As String, ByRef pvEgresosCantidad() As Double, ByVal pNroVale As String, ByVal pCodTipoVale As String, ByVal pCodBodega As String, ByVal pCodUsuarioRetira As String, ByVal pCodUsuarioAutoriza As String, ByVal pCodUsuario As String, ByRef bRet As Boolean)
On Error GoTo ErrorHandler
Dim strSql1 As String
Dim mi As Integer
'Dim prueba As Integer
   
   cmdTemp.CommandType = 1
   cmdTemp.CommandTimeout = 36000
   Set cmdTemp.ActiveConnection = DBConn
   cmdTemp.ActiveConnection.BeginTrans

   
   strSql1 = "INSERT INTO Consumos_H (NroVale,CodTipoVale,CodBodega,CodUsuarioRetira,CodUsuarioAutoriza,Fecha,CodUsuario)  VALUES ('" & pNroVale & "','" & pCodTipoVale & "','" & pCodBodega & "','" & pCodUsuarioRetira & "', '" & pCodUsuarioAutoriza & "', NOW(),'" & pCodUsuario & "');"

   cmdTemp.CommandText = strSql1
   cmdTemp.Execute

   cmdTemp.CommandType = adCmdStoredProc

   For mi = LBound(pvEgresosCodProducto) To UBound(pvEgresosCodProducto)
      
'     Genero un error(division por cero) ara probar si funcion el rollBackTrans
'      If mI = 1 Then
'         prueba = 1 / 0
'      End If

      insEgreso pvEgresosCodProducto(mi), pvEgresosCodUbicacion(mi), pvEgresosCantidad(mi), pNroVale, pCodTipoVale, pCodUsuario, cmdTemp
      cmdTemp.Execute
   Next

   cmdTemp.ActiveConnection.CommitTrans
   cmdTemp.CommandType = 1
   Exit Sub
ErrorHandler:

   MsgBox "Se produjo un Error al realizar el consumo. Comunicarse con el Administrador" & vbCrLf & "Detalle del error: " & Err.Number & " " & Err.Description, vbCritical, "Error"
   cmdTemp.ActiveConnection.RollbackTrans
   cmdTemp.CommandType = 1
   bRet = False

End Sub

'---------------------------------------
'Sp: insIngreso
'---------------------------------------
Public Sub insIngreso(ByVal CodProducto As String, ByVal CodUbicacion As String, ByVal Cantidad As Double, ByVal NroOC As String, ByVal CodUsuario As String, ByRef cmd As ADODB.Command)
   Dim parametroSql As parametroSql
   Dim aParamSql(4) As parametroSql
   
   parametroSql.Nombre = "vCodProducto"
   parametroSql.Tipo = adChar
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 6
   parametroSql.Valor = CodProducto
   aParamSql(0) = parametroSql
   
   parametroSql.Nombre = "vCodUbicacion"
   parametroSql.Tipo = adChar
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 4
   parametroSql.Valor = CodUbicacion
   aParamSql(1) = parametroSql
   
   parametroSql.Nombre = "vCantidad"
   parametroSql.Tipo = adDouble
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 8
   parametroSql.Valor = Cantidad
   aParamSql(2) = parametroSql
   
   parametroSql.Nombre = "vNroOC"
   parametroSql.Tipo = adVarChar
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 10
   parametroSql.Valor = NroOC
   aParamSql(3) = parametroSql
   
   parametroSql.Nombre = "vCodUsuario"
   parametroSql.Tipo = adVarChar
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 25
   parametroSql.Valor = CodUsuario
   aParamSql(4) = parametroSql
   
   xStoreProcedureWithParam_Transaccional "insIngreso", aParamSql, cmd
 End Sub

Public Sub xInsIngreso(ByRef pvIngresosCodProducto() As String, ByRef pvIngresosCodUbicacion() As String, ByRef pvIngresosCantidad() As Double, ByVal pNroOC As String, ByVal pCodBodega As String, ByVal pCodUsuario As String, ByRef bRet As Boolean)
On Error GoTo ErrorHandler
Dim strSql1 As String
Dim mi As Integer
'Dim prueba As Integer
   
   cmdTemp.CommandType = 1
   cmdTemp.CommandTimeout = 36000
   Set cmdTemp.ActiveConnection = DBConn
   cmdTemp.ActiveConnection.BeginTrans

   'TODO: Hacer el INSERT en Ingresos_H
   strSql1 = "INSERT INTO Ingresos_H (NroOC,CodBodega,Fecha,CodUsuario)  VALUES ('" & pNroOC & "','" & pCodBodega & "', NOW(),'" & pCodUsuario & "');"

   cmdTemp.CommandText = strSql1
   cmdTemp.Execute

   cmdTemp.CommandType = adCmdStoredProc

   For mi = LBound(pvIngresosCodProducto) To UBound(pvIngresosCodProducto)
      
'     Genero un error(division por cero) ara probar si funcion el rollBackTrans
'      If mI = 1 Then
'         prueba = 1 / 0
'      End If

      insIngreso pvIngresosCodProducto(mi), pvIngresosCodUbicacion(mi), pvIngresosCantidad(mi), pNroOC, pCodUsuario, cmdTemp
      cmdTemp.Execute
   Next

   cmdTemp.ActiveConnection.CommitTrans
   cmdTemp.CommandType = 1
   Exit Sub
ErrorHandler:

   MsgBox "Se produjo un Error al realizar el Ingreso. Comunicarse con el Administrador" & vbCrLf & "Detalle del error: " & Err.Number & " " & Err.Description, vbCritical, "Error"
   cmdTemp.ActiveConnection.RollbackTrans
   cmdTemp.CommandType = 1
   bRet = False
End Sub


'---------------------------------------
'Sp: insAjuste
'---------------------------------------
Public Sub insAjuste(ByVal CodProducto As String, ByVal CodUbicacion As String, ByVal Cantidad As Double, ByVal TipoAjuste As String, ByVal MotivoDesc As String, ByVal CodUsuario As String, ByRef cmd As ADODB.Command)
   Dim parametroSql As parametroSql
   Dim aParamSql(5) As parametroSql
   
   parametroSql.Nombre = "vCodProducto"
   parametroSql.Tipo = adChar
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 6
   parametroSql.Valor = CodProducto
   aParamSql(0) = parametroSql
   
   parametroSql.Nombre = "vCodUbicacion"
   parametroSql.Tipo = adChar
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 4
   parametroSql.Valor = CodUbicacion
   aParamSql(1) = parametroSql
   
   parametroSql.Nombre = "vCantidad"
   parametroSql.Tipo = adDouble
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 8
   parametroSql.Valor = Cantidad
   aParamSql(2) = parametroSql
   
   parametroSql.Nombre = "vTipoAjuste"
   parametroSql.Tipo = adChar
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 1
   parametroSql.Valor = TipoAjuste
   aParamSql(3) = parametroSql
   
   parametroSql.Nombre = "vMotivoDesc"
   parametroSql.Tipo = adVarChar
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 150
   parametroSql.Valor = MotivoDesc
   aParamSql(4) = parametroSql
   
   parametroSql.Nombre = "vCodUsuario"
   parametroSql.Tipo = adVarChar
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 25
   parametroSql.Valor = CodUsuario
   aParamSql(5) = parametroSql
   
   xStoreProcedureWithParam_Transaccional "insAjuste", aParamSql, cmd
 End Sub

Public Sub xInsAjuste(ByRef pvAjustesCodProducto() As String, ByRef pvAjustesCodUbicacion() As String, ByRef pvAjustesCantidad() As Double, ByRef pvAjustesTipoAjuste() As String, ByVal pMotivoDesc As String, ByVal pCodUsuario As String, ByRef bRet As Boolean)
On Error GoTo ErrorHandler
Dim strSql1 As String
Dim mi As Integer
   
   cmdTemp.CommandType = 1
   cmdTemp.CommandTimeout = 36000
   Set cmdTemp.ActiveConnection = DBConn
   cmdTemp.ActiveConnection.BeginTrans

   cmdTemp.CommandType = adCmdStoredProc

   For mi = LBound(pvAjustesCodProducto) To UBound(pvAjustesCodProducto)
      insAjuste pvAjustesCodProducto(mi), pvAjustesCodUbicacion(mi), pvAjustesCantidad(mi), pvAjustesTipoAjuste(mi), pMotivoDesc, pCodUsuario, cmdTemp
      cmdTemp.Execute
   Next

   cmdTemp.ActiveConnection.CommitTrans
   cmdTemp.CommandType = 1
   Exit Sub
ErrorHandler:

   MsgBox "Se produjo un Error al realizar el Ajuste. Comunicarse con el Administrador" & vbCrLf & "Detalle del error: " & Err.Number & " " & Err.Description, vbCritical, "Error"
   cmdTemp.ActiveConnection.RollbackTrans
   cmdTemp.CommandType = 1
   bRet = False
End Sub



Public Sub xInsAddItemsOC(ByRef pvIngresosCodProducto() As String, ByRef pvIngresosCodUbicacion() As String, ByRef pvIngresosCantidad() As Double, ByVal pNroOC As String, ByVal pCodUsuario As String, ByRef bRet As Boolean)
On Error GoTo ErrorHandler
Dim mi As Integer

'Dim prueba As Integer
   
'   cmdTemp.CommandType = 1
   cmdTemp.CommandTimeout = 36000
   Set cmdTemp.ActiveConnection = DBConn
   cmdTemp.ActiveConnection.BeginTrans

   cmdTemp.CommandType = adCmdStoredProc

   For mi = LBound(pvIngresosCodProducto) To UBound(pvIngresosCodProducto)
      
'     Genero un error(division por cero) ara probar si funcion el rollBackTrans
'      If mI = 1 Then
'         prueba = 1 / 0
'      End If

      insIngreso pvIngresosCodProducto(mi), pvIngresosCodUbicacion(mi), pvIngresosCantidad(mi), pNroOC, pCodUsuario, cmdTemp
      cmdTemp.Execute
   Next

   cmdTemp.ActiveConnection.CommitTrans
   cmdTemp.CommandType = 1
   Exit Sub
ErrorHandler:


   MsgBox "Se produjo un Error al realizar el Ingreso. Comunicarse con el Administrador" & vbCrLf & "Detalle del error: " & Err.Number & " " & Err.Description, vbCritical, "Error"
   cmdTemp.ActiveConnection.RollbackTrans
   cmdTemp.CommandType = 1
   bRet = False
   
End Sub

Public Function xInsStockMinimo_Notificaciones(ByVal pCodProducto As String, ByVal pCodBodega As String)
   strSql1 = "INSERT INTO StockMinimo_Notificaciones(CodProducto,CodBodega,Fecha_Notificacion) VALUES ('" & pCodProducto & "','" & pCodBodega & "',NOW());"
   sAsignCmd
   cmdTemp.Execute
End Function

Public Function xDelStockMinimo_Notificaciones(ByVal pCodProducto As String, ByVal pCodBodega As String)
   strSql1 = "DELETE FROM StockMinimo_Notificaciones WHERE CodProducto = '" & pCodProducto & "' AND CodBodega ='" & pCodBodega & "';"
   sAsignCmd
   cmdTemp.Execute
End Function


Public Function xInsMovimiento(ByVal pCodProducto As String, ByVal pCodTipoMovimiento As String, ByVal pCantidad As String, ByVal pCodMotivo As String, ByVal pCodUsuario As String, ByVal pObservaciones As String)
   Dim FechaActual As String
   Dim CodUsuario As String
   
   pCantidad = Replace(pCantidad, ",", ".")
   FechaActual = Format(DateTime.Now(), "yyyy-mm-dd hh:mm:ss")
   CodUsuario = pCodUsuario
    
   'INSERTO MOVIMIENTO
   strSql1 = "insert into Movimientos (CodProducto, CodTipoMovimiento, Cantidad, CodMotivo, Fecha, CodUsuario, Observaciones) " & _
               " values " & _
              " ('" & pCodProducto & "','" & pCodTipoMovimiento & "'," & pCantidad & ",'" & pCodMotivo & "','" & FechaActual & "','" & CodUsuario & "','" & pObservaciones & "')"
   sAsignCmd
   cmdTemp.Execute
   
   
   'ACTUALIZO SALDO DE STOCK
   If pCodTipoMovimiento = "I" Then
         strSql1 = "UPDATE Producto " & _
                     " Set Stock = Stock + " & pCantidad & _
                  " WHERE Codigo = '" & pCodProducto & "';"
   Else
         strSql1 = "UPDATE Producto " & _
                     " Set Stock = Stock - " & pCantidad & _
                  " WHERE Codigo = '" & pCodProducto & "';"
   End If
   
   sAsignCmd
   cmdTemp.Execute
End Function











