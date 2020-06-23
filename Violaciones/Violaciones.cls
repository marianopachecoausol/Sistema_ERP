VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clViolaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Private DBConn As New ADODB.Connection
Private cmdTemp As New ADODB.Command
Dim mObj As New clConnSistemas
Dim Rs1 As New ADODB.Recordset
Dim strSql1 As String

Private Sub Class_Initialize()
DBConn.ConnectionTimeout = 500 ' Original en 15
DBConn.ConnectionTimeout = 500 ' Original en 90
Set DBConn = mObj.oAbreConexion("4", mIPServer, "Violaciones", "ssvv", "ssvv") 'oAbreConexion esta en el modulo1
Rs1.CursorType = adOpenDynamic
End Sub

Private Sub class_terminate()
DBConn.Close
Set DBConn = Nothing
Set mObj = Nothing
Set Rs1 = Nothing
End Sub

Private Sub sAsignCmd()
cmdTemp.CommandText = strSql1
cmdTemp.CommandType = 1
cmdTemp.CommandTimeout = 36000
Set cmdTemp.ActiveConnection = DBConn
End Sub

Public Function oTabla(ByVal pNameTabla As String, ByVal pOrden As String) As ADODB.Recordset
strSql1 = "SELECT * FROM " & pNameTabla & " " & pOrden
sAsignCmd
Set Rs1 = cmdTemp.Execute
Set oTabla = Rs1
End Function

Public Function oEjecutarSelect(ByVal pSentencia As String) As ADODB.Recordset
strSql1 = pSentencia
sAsignCmd
Set oEjecutarSelect = cmdTemp.Execute
End Function

Public Function oTablaDina(ByVal pNameTabla As String, ByVal pOrden As String) As ADODB.Recordset
Dim mRs2 As New ADODB.Recordset
mRs2.CursorType = adOpenDynamic
strSql1 = "SELECT * FROM " & pNameTabla & " " & pOrden
sAsignCmd
mRs2.Open strSql1, DBConn
Set oTablaDina = mRs2
End Function

Public Function oTablaNotNull(ByVal pNameTabla As String, ByVal pOrden As String) As ADODB.Recordset
strSql1 = "SELECT * FROM " & pNameTabla & " WHERE baja is null " & pOrden
sAsignCmd
Set Rs1 = cmdTemp.Execute
Set oTablaNotNull = Rs1
End Function

Public Function oTablaCodigo(ByVal pNameTabla As String, ByVal pCodigo As String) As ADODB.Recordset
strSql1 = "SELECT * FROM " & pNameTabla & " where " & pCodigo
sAsignCmd
Set Rs1 = cmdTemp.Execute
Set oTablaCodigo = Rs1
End Function

Public Function sCampoDescrip(ByVal pTabla As String, ByVal pWhereCodigo As String, ByVal pCampoDescr As Integer) As String
strSql1 = "select * from " & pTabla & " where " & pWhereCodigo
sAsignCmd
Set Rs1 = cmdTemp.Execute
sCampoDescrip = ""
If Not Rs1.EOF Then
   sCampoDescrip = Rs1.Fields(pCampoDescr)
End If
End Function

Public Function oViolFechasPatente(ByVal pFecha1 As String, ByVal pFecha2 As String, ByVal pPatente As String, ByVal pTipo As String) As ADODB.Recordset
strSql1 = "select distinct fecha, hora, estacion, via from Registros where fecha between '" & Format(pFecha1, "yyyy-mm-dd") & "' and " _
        & " '" & Format(pFecha2, "yyyy-mm-dd") & "' And TIPO = '" & pTipo & "' And patente='" & pPatente & "' and (pago<>'S' or pago is null) order by 1,2"
sAsignCmd
Set Rs1 = cmdTemp.Execute
Set oViolFechasPatente = Rs1
End Function

Public Function oViolFechasPatEst(ByVal pPatente As String, ByVal pFecha1 As String, _
                               ByVal pFecha2 As String, ByVal pIntervEst As String, Optional pTipo As String) As ADODB.Recordset
If pTipo <> "" Then
   strSql1 = "SELECT * FROM Registros WHERE PATENTE LIKE '" & pPatente & "' And FECHA BETWEEN '" & Format(pFecha1, "yyyy-mm-dd hh:mm:ss") & "' " _
           & " And '" & Format(pFecha2, "yyyy-mm-dd hh:mm:ss") & "' And TIPO = '" & pTipo & "'"
Else
   strSql1 = "SELECT * FROM Registros WHERE PATENTE LIKE '" & pPatente & "' And FECHA BETWEEN '" & Format(pFecha1, "yyyy-mm-dd hh:mm:ss") & "' " _
           & " And '" & Format(pFecha2, "yyyy-mm-dd hh:mm:ss") & "'"
End If
If Trim(pIntervEst) <> "" Then
   strSql1 = strSql1 & " and estacion in (" & pIntervEst & ") "
End If
strSql1 = strSql1 & " order by 1,3,2"
sAsignCmd
Set Rs1 = cmdTemp.Execute
Set oViolFechasPatEst = Rs1
End Function

Public Function oDistPatenteTabla(ByVal pTabla As String, ByVal pPatente As String, ByVal PDireccion As String, ByVal pName As String) As ADODB.Recordset
   Dim mWhere As String
   mWhere = ""
   If Trim(pPatente) <> "" Then
      mWhere = " Patente LIKE '" & Trim(pPatente) & "' "
   End If
   If Trim(pName) <> "" Then
      If mWhere <> "" Then
         mWhere = mWhere & " AND "
      End If
      mWhere = mWhere & "nombre LIKE '%" & Trim(pName) & "%' "
   End If
   If PDireccion <> "" Then
      If mWhere <> "" Then
         mWhere = " AND "
      End If
      mWhere = mWhere & "domicilio LIKE '%" & Trim(PDireccion) & "%' "
   End If
   If mWhere <> "" Then
      mWhere = " where " & mWhere
   End If
   strSql1 = "select distinct patente from " & pTabla & mWhere & " order by patente"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oDistPatenteTabla = Rs1
End Function

Public Function oDistPatenteDirEnv(ByVal pPatente As String, ByVal PDireccion As String, ByVal pName As String) As ADODB.Recordset
   Dim mWhere As String
   mWhere = ""
   If Trim(pPatente) <> "" Then
      mWhere = "and A.Patente LIKE '" & Trim(pPatente) & "' "
   End If
   If Trim(pName) <> "" Then
      mWhere = mWhere & " and nombre LIKE '%" & Trim(pName) & "%' "
   End If
   If Trim(PDireccion) <> "" Then
      mWhere = mWhere & " and domicilio LIKE '%" & Trim(PDireccion) & "%' "
   End If
   strSql1 = "select distinct A.patente from direcciones A, envios B where A.patente=B.patente " & mWhere & " order by A.patente"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oDistPatenteDirEnv = Rs1
End Function

Public Function oRegistrosPatente(ByVal pPatente As String) As ADODB.Recordset
strSql1 = "SELECT * FROM Registros WHERE patente = '" & pPatente & "' ORDER BY 1, 2"
sAsignCmd
Set Rs1 = cmdTemp.Execute
Set oRegistrosPatente = Rs1
End Function

Public Sub UpdPagos(ByVal pPago As String, pFecha As String, ByVal pHora As String, ByVal pPatente As String)
strSql1 = "UPDATE Registros SET pago = '" & pPago & "' WHERE fecha = '" & Format(pFecha, "yyyy-mm-dd") & "' and hora = '" & pHora & "' and patente = '" & pPatente & "'"
sAsignCmd
cmdTemp.Execute
End Sub

Public Sub InsLogPagos(ByVal pCodUsuario As String, pDateTime As Date, ByVal pPatente As String, ByVal pPago As String, ByVal pFecha As String, ByVal pHora As String, ByVal pEst As String)
strSql1 = "INSERT INTO LogPagos VALUES ('" & pCodUsuario & "','" & Format(pDateTime, "yyyy-mm-dd hh:mm:ss") & "','" & pPatente & "','" & pPago & "','" & Format(pFecha, "yyyy-mm-dd") & "','" & pHora & "','" & pEst & "')"
sAsignCmd
cmdTemp.Execute
End Sub

Public Function oDatosPatente(ByVal pPatente As String) As ADODB.Recordset
'strSql1 = "select distinct A.nombre, A.domicilio, A.codpostal, B.descripcion as localidad, C.descripcion as Prov FROM direcciones A, postal B, provincias C " _
        & " where A.patente='" & pPatente & "' AND (A.codpostal=B.codigo AND A.codpcia=B.codpcia) AND A.codpcia=C.codigo"
   strSql1 = "select * FROM direcciones where patente='" & pPatente & "' "
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oDatosPatente = Rs1
End Function

Public Function oCDporPatente(ByVal pPatente As String, Optional pTipo As String) As ADODB.Recordset
If pTipo <> "" Then
   strSql1 = "select distinct a.*, b.descripcion FROM envios a, entregas b WHERE patente='" & pPatente & "' AND a.codentrega=b.codigo And a.tipo = '" & pTipo & "' ORDER BY FECHA DESC"
Else
   strSql1 = "select distinct a.*, b.descripcion FROM envios a, entregas b WHERE patente='" & pPatente & "' AND a.codentrega=b.codigo ORDER BY FECHA DESC"
End If
sAsignCmd
Set Rs1 = cmdTemp.Execute
Set oCDporPatente = Rs1
End Function

Public Function oMarcas() As ADODB.Recordset
 Dim Rs1, strSql1
 Set Rs1 = New ADODB.Recordset
 strSql1 = "SELECT * FROM marcas"
 cmdTemp.CommandText = strSql1
 cmdTemp.CommandType = 1
 cmdTemp.CommandTimeout = 36000
 Set cmdTemp.ActiveConnection = DBConn
 Set Rs1 = cmdTemp.Execute
 Set oMarcas = Rs1
End Function

Public Function oColores() As ADODB.Recordset
 Dim Rs1, strSql1
 Set Rs1 = New ADODB.Recordset
 strSql1 = "SELECT * FROM colores ORDER BY descripcion"
 cmdTemp.CommandText = strSql1
 cmdTemp.CommandType = 1
 cmdTemp.CommandTimeout = 36000
 Set cmdTemp.ActiveConnection = DBConn
 Set Rs1 = cmdTemp.Execute
 Set oColores = Rs1
End Function

Public Function fUltCodigo(ByVal pTabla As String, ByVal pId As String) As String 'último código
 Dim Rs1 As ADODB.Recordset
 Dim strSql1
 Dim mCod As Integer
 Set Rs1 = New ADODB.Recordset
 If pId = "" Then
    strSql1 = "SELECT * FROM " & LCase(pTabla) & " ORDER BY 1 DESC"
 Else
    strSql1 = "SELECT * FROM " & LCase(pTabla) & " WHERE codmarca='" & pId & "' ORDER BY 2 DESC"
 End If
 cmdTemp.CommandText = strSql1
 cmdTemp.CommandType = 1
 cmdTemp.CommandTimeout = 36000
 Set cmdTemp.ActiveConnection = DBConn
 Set Rs1 = cmdTemp.Execute
 mCod = Rs1!Codigo
 mCod = mCod + 1
 fUltCodigo = Format(mCod, "00")
End Function

Public Function fUltEnvio(ByVal pPatente As String, ByVal pTipo As String) As String
Dim Rs1 As ADODB.Recordset
Dim strSql1
Set Rs1 = New ADODB.Recordset
strSql1 = "SELECT nrocarta FROM envios WHERE patente='" & pPatente & "' And tipo = '" & pTipo & "' ORDER BY fecha DESC limit 1"
cmdTemp.CommandText = strSql1
cmdTemp.CommandType = 1
cmdTemp.CommandTimeout = 36000
Set cmdTemp.ActiveConnection = DBConn
Set Rs1 = cmdTemp.Execute
If Not Rs1.EOF Then
   fUltEnvio = Rs1!NROCARTA
Else
   fUltEnvio = ""
End If
End Function

Public Function fUltFechaEnvio(ByVal pPatente As String, ByVal pTipo As String) As String
Dim Rs1 As ADODB.Recordset
Dim strSql1 As String
Set Rs1 = New ADODB.Recordset
strSql1 = "SELECT fecha FROM envios WHERE patente='" & pPatente & "' And TIPO = '" & pTipo & "' ORDER BY fecha DESC limit 1"
cmdTemp.CommandText = strSql1
cmdTemp.CommandType = 1
cmdTemp.CommandTimeout = 36000
Set cmdTemp.ActiveConnection = DBConn
Set Rs1 = cmdTemp.Execute
If Not Rs1.EOF Then
   fUltFechaEnvio = Rs1!Fecha
Else
   fUltFechaEnvio = ""
End If
End Function

Public Function oCountEnviosCD(ByVal pCant As String) As ADODB.Recordset
 Dim Rs1, strSql1
 Dim pWhere As String
 Set Rs1 = New ADODB.Recordset
 pWhere = ""
 If pCant <> "" And pCant <> "0" Then
   pWhere = "HAVING total >= " & pCant
 End If
 strSql1 = "SELECT a.patente, COUNT(a.patente) as total FROM envios a, regpagos b  where a.patente <> b.patente GROUP BY patente  " & pWhere & " ORDER BY 2"
 cmdTemp.CommandText = strSql1
 cmdTemp.CommandType = 1
 cmdTemp.CommandTimeout = 36000
 Set cmdTemp.ActiveConnection = DBConn
 Set Rs1 = cmdTemp.Execute
 Set oCountEnviosCD = Rs1
End Function

Public Function oDetalleEnviosCD(ByVal pPatente As String) As ADODB.Recordset
 Dim Rs1, strSql1
 Set Rs1 = New ADODB.Recordset
 strSql1 = "SELECT distinct a.*, b.nombre FROM envios a, direcciones b  WHERE a.patente='" & pPatente & "' AND a.patente=b.patente ORDER BY fecha"
 cmdTemp.CommandText = strSql1
 cmdTemp.CommandType = 1
 cmdTemp.CommandTimeout = 36000
 Set cmdTemp.ActiveConnection = DBConn
 Set Rs1 = cmdTemp.Execute
 Set oDetalleEnviosCD = Rs1
End Function

Public Function oCodEntrega() As ADODB.Recordset
 Dim Rs1, strSql1
 Set Rs1 = New ADODB.Recordset
 strSql1 = "SELECT * FROM entregas ORDER BY 1"
 cmdTemp.CommandText = strSql1
 cmdTemp.CommandType = 1
 cmdTemp.CommandTimeout = 36000
 Set cmdTemp.ActiveConnection = DBConn
 Set Rs1 = cmdTemp.Execute
 Set oCodEntrega = Rs1
End Function

Public Function oTotalEnviosEntrega(ByVal pFecha1 As String, ByVal pFecha2 As String) As ADODB.Recordset
   Dim Rs1, strSql1
   Set Rs1 = New ADODB.Recordset
   strSql1 = "SELECT a.codentrega, b.descripcion, count(a.nrocarta) as total FROM envios a, entregas b " _
           & "WHERE fecha BETWEEN '" & Format(pFecha1, "yyyy-mm-dd") & "' AND '" & Format(pFecha2, "yyyy-mm-dd") & "'and a.codentrega=b.codigo " _
           & "GROUP BY a.codentrega, b.descripcion"
   cmdTemp.CommandText = strSql1
   cmdTemp.CommandType = 1
   cmdTemp.CommandTimeout = 36000
   Set cmdTemp.ActiveConnection = DBConn
   Set Rs1 = cmdTemp.Execute
   Set oTotalEnviosEntrega = Rs1
End Function

Public Function oCountViolDspFecha(ByVal pDate1 As String, ByVal pDate2 As String) As ADODB.Recordset
   Dim Rs1 As New ADODB.Recordset
   Dim strSql
   strSql = "SELECT a.patente, count(*) as total FROM Registros a, envios b WHERE a.patente = b.patente " _
      & "AND a.fecha > '" & Format(pDate2, "yyyy-mm-dd") & "' AND b.fecha BETWEEN '" & Format(pDate1, "yyyy-mm-dd") & "'" _
      & "AND '" & Format(pDate2, "yyyy-mm-dd") & "' GROUP BY a.patente ORDER BY 1"
   cmdTemp.CommandText = strSql
   cmdTemp.CommandType = 1
   cmdTemp.CommandTimeout = 36000
   Set cmdTemp.ActiveConnection = DBConn
   Set Rs1 = cmdTemp.Execute
   Set oCountViolDspFecha = Rs1
End Function

Public Function iCountViolFechaPatente2(ByVal pPatente As String, ByVal pDate As String, Optional pTipo As String) As Integer
Dim Rs1 As New ADODB.Recordset
Dim strSql
If pTipo <> "" Then
   strSql = "SELECT count(*) as total FROM Registros WHERE patente = '" & pPatente & "' And tipo = '" & pTipo & "' AND fecha >= '" & Format(pDate, "yyyy-mm-dd") & "'"
Else
   strSql = "SELECT count(*) as total FROM Registros WHERE patente = '" & pPatente & "' AND fecha >= '" & Format(pDate, "yyyy-mm-dd") & "'"
End If
cmdTemp.CommandText = strSql
cmdTemp.CommandType = 1
cmdTemp.CommandTimeout = 36000
Set cmdTemp.ActiveConnection = DBConn
Set Rs1 = cmdTemp.Execute
If Not Rs1.EOF Then
   iCountViolFechaPatente2 = Rs1!Total
Else
   iCountViolFechaPatente2 = 0
End If
End Function

Public Function oCartasEnPeriodo(ByVal pDate1 As String, ByVal pDate2 As String) As ADODB.Recordset
   Dim Rs1 As New ADODB.Recordset
   Dim strSql
   strSql = "SELECT patente, nrocarta, fecha FROM envios WHERE fecha BETWEEN '" & Format(pDate1, "yyyy-mm-dd") & "' and '" & Format(pDate2, "yyyy-mm-dd") & "' ORDER BY 1,3 desc"
   cmdTemp.CommandText = strSql
   cmdTemp.CommandType = 1
   cmdTemp.CommandTimeout = 36000
   Set cmdTemp.ActiveConnection = DBConn
   Set Rs1 = cmdTemp.Execute
   Set oCartasEnPeriodo = Rs1
End Function

Public Function oMesAnioEnvios() As ADODB.Recordset
Dim Rs1 As New ADODB.Recordset
   Dim strSql
   strSql = "SELECT DISTINCTROW year(fecha) as anio, month(fecha) as mes FROM envios ORDER BY 1 desc, 2 desc"
   cmdTemp.CommandText = strSql
   cmdTemp.CommandType = 1
   cmdTemp.CommandTimeout = 36000
   Set cmdTemp.ActiveConnection = DBConn
   Set Rs1 = cmdTemp.Execute
   Set oMesAnioEnvios = Rs1
End Function

Public Function oViolxPatenteDate(ByVal pDate1 As String, ByVal pDate2 As String, ByVal pCant As String, ByVal pWhere As String, ByVal pTipo As String) As ADODB.Recordset
Dim Rs1 As New ADODB.Recordset
Dim strSql
strSql = "SELECT distinct Patente, COUNT(distinct fecha, hora) as Total From Registros WHERE tipo = '" & pTipo & "' And Fecha BETWEEN '" & Format(pDate1, "yyyy-mm-dd") & "' AND '" & Format(pDate2, "yyyy-mm-dd") & "' And pago <> 'S' " _
       & pWhere & " GROUP BY Patente HAVING Total >= " & pCant & " ORDER BY Patente"
cmdTemp.CommandText = strSql
cmdTemp.CommandType = 1
cmdTemp.CommandTimeout = 36000
Set cmdTemp.ActiveConnection = DBConn
Set Rs1 = cmdTemp.Execute
Set oViolxPatenteDate = Rs1
End Function

Public Function oViolxPatenteNoEntrega(ByVal pDate1 As String, ByVal pDate2 As String, ByVal pCant As String, ByVal pWhere As String, ByVal pTipo As String) As ADODB.Recordset
Dim Rs1 As New ADODB.Recordset
Dim strSql As String
strSql = "SELECT distinct A.Patente, COUNT(distinct A.fecha, A.hora) as Total From Registros A, envios B WHERE A.Fecha BETWEEN '" & Format(pDate1, "yyyy-mm-dd") & "' AND '" & Format(pDate2, "yyyy-mm-dd") & "' " _
       & " AND A.TIPO = '" & pTipo & "' And B.TIPO = '" & pTipo & "' And B.patente=A.patente AND B.codentrega <> '00' " & pWhere & " GROUP BY A.Patente HAVING Total >= " & pCant & " ORDER BY A.Patente"
cmdTemp.CommandText = strSql
cmdTemp.CommandType = 1
cmdTemp.CommandTimeout = 36000
Set cmdTemp.ActiveConnection = DBConn
Set Rs1 = cmdTemp.Execute
Set oViolxPatenteNoEntrega = Rs1
End Function

Public Function bCDRechazadas(ByVal pPatente As String, ByVal pFecha1 As String, ByVal pFecha2 As String, ByVal pTipo As String) As Boolean
Dim Rs1 As New ADODB.Recordset
Dim strSql As String
Dim mFlag As Boolean
strSql = "SELECT * FROM envios WHERE patente='" & pPatente & "' And tipo = '" & pTipo & "' AND codentrega='00' and fecha between '" & Format(pFecha1, "yyyy-mm-dd") & "' and '" & Format(pFecha2, "yyyy-mm-dd") & "'"
cmdTemp.CommandText = strSql
cmdTemp.CommandType = 1
cmdTemp.CommandTimeout = 36000
Set cmdTemp.ActiveConnection = DBConn
Set Rs1 = cmdTemp.Execute
bCDRechazadas = Rs1.EOF
End Function

Public Function oCDEnvioRecib(ByVal pDate1 As String, ByVal pDate2 As String, ByVal pCant As String, ByVal pWhere As String, ByVal pTipo As String) As ADODB.Recordset
Dim Rs1 As New ADODB.Recordset
Dim strSql As String
Dim mFlag As Boolean
strSql = "SELECT distinct A.Patente, COUNT(distinct A.fecha, A.hora) as Total, count(distinct B.nrocarta) as total2 " _
       & " From Registros A, envios B WHERE A.Fecha BETWEEN '" & Format(pDate1, "yyyy-mm-dd") & "' and '" & Format(pDate2, "yyyy-mm-dd") & "' AND B.patente=A.patente " _
       & " AND B.codentrega = '00' And A.TIPO = '" & pTipo & "' And B.TIPO = '" & pTipo & "' And B.fecha >='2007-01-01' " & pWhere & " GROUP BY Patente Having Total >= " & pCant & " And total2 = 1 ORDER BY Patente"
cmdTemp.CommandText = strSql
cmdTemp.CommandType = 1
cmdTemp.CommandTimeout = 36000
Set cmdTemp.ActiveConnection = DBConn
Set Rs1 = cmdTemp.Execute
Set oCDEnvioRecib = Rs1
End Function

Public Function oCDRecibidas(ByVal pDate1 As String, ByVal pDate2 As String, ByVal pTipo As String) As ADODB.Recordset
Dim Rs1 As New ADODB.Recordset
Dim strSql As String
Dim mFlag As Boolean
strSql = "select patente, fecha, count(nrocarta) as total From envios where codentrega='00' And TIPO = '" & pTipo & "' and fecha between '" & Format(pDate1, "yyyy-mm-dd") & "' and '" & Format(pDate2, "yyyy-mm-dd") & "' group by patente order by 1"
cmdTemp.CommandText = strSql
cmdTemp.CommandType = 1
cmdTemp.CommandTimeout = 36000
Set cmdTemp.ActiveConnection = DBConn
Set Rs1 = cmdTemp.Execute
Set oCDRecibidas = Rs1
End Function

Public Function oEnviosxPatente(ByVal pPatente As String, ByVal pTipo As String) As ADODB.Recordset
Dim Rs1 As New ADODB.Recordset
Dim strSql
Dim mFlag As Boolean
strSql = "SELECT * FROM envios WHERE patente='" & pPatente & "' And tipo = '" & pTipo & "' order by 3 desc"
cmdTemp.CommandText = strSql
cmdTemp.CommandType = 1
cmdTemp.CommandTimeout = 36000
Set cmdTemp.ActiveConnection = DBConn
Set Rs1 = cmdTemp.Execute
Set oEnviosxPatente = Rs1
End Function

Public Function oPatentesReg(ByVal pPatente As String) As ADODB.Recordset
Dim Rs1 As New ADODB.Recordset
Dim strSql As String
Dim mFlag As Boolean
strSql = "SELECT DISTINCT PATENTE FROM Registros WHERE PATENTE LIKE '" & pPatente & "' ORDER BY 1"
cmdTemp.CommandText = strSql
cmdTemp.CommandType = 1
cmdTemp.CommandTimeout = 36000
Set cmdTemp.ActiveConnection = DBConn
Set Rs1 = cmdTemp.Execute
Set oPatentesReg = Rs1
End Function

Public Function oPatentesStBy(ByVal pPatente As String, ByVal pTipo As String) As ADODB.Recordset
Dim Rs1 As New ADODB.Recordset
Dim strSql
Dim mFlag As Boolean
strSql = "select * from regpagos where patente = '" & pPatente & "' And tipo = '" & pTipo & "' order by fecha desc limit 1 "
cmdTemp.CommandText = strSql
cmdTemp.CommandType = 1
cmdTemp.CommandTimeout = 36000
Set cmdTemp.ActiveConnection = DBConn
Set Rs1 = cmdTemp.Execute
Set oPatentesStBy = Rs1
End Function

Public Function oAllLocalidad(ByVal pCodPcia As String) As ADODB.Recordset
   Dim Rs1 As New ADODB.Recordset
   Dim strSql
   Dim mFlag As Boolean
   strSql = "select distinct descripcion, codigo,codpcia from postal where codpcia in ('" & pCodPcia & "') order by 1"
   cmdTemp.CommandText = strSql
   cmdTemp.CommandType = 1
   cmdTemp.CommandTimeout = 36000
   Set cmdTemp.ActiveConnection = DBConn
   Set Rs1 = cmdTemp.Execute
   Set oAllLocalidad = Rs1
End Function

Public Function oCP_Pcia(ByVal pCodPcia As String) As ADODB.Recordset
   Dim Rs1 As New ADODB.Recordset
   Dim strSql
   Dim mFlag As Boolean
   strSql = "SELECT * FROM postal WHERE codpcia='" & pCodPcia & "' ORDER BY 1"
   cmdTemp.CommandText = strSql
   cmdTemp.CommandType = 1
   cmdTemp.CommandTimeout = 36000
   Set cmdTemp.ActiveConnection = DBConn
   Set Rs1 = cmdTemp.Execute
   Set oCP_Pcia = Rs1
End Function

Public Function oCPxLocalidad(ByVal pLocalidad As String) As ADODB.Recordset
   Dim Rs1 As New ADODB.Recordset
   Dim strSql
   Dim mFlag As Boolean
   strSql = "select distinct codigo, codpcia from postal where descripcion='" & pLocalidad & "' order by 1"
   cmdTemp.CommandText = strSql
   cmdTemp.CommandType = 1
   cmdTemp.CommandTimeout = 36000
   Set cmdTemp.ActiveConnection = DBConn
   Set Rs1 = cmdTemp.Execute
   Set oCPxLocalidad = Rs1
End Function

Public Function oLocalidadxCP_PCIA(ByVal pCP As String, ByVal pCodPcia As String) As ADODB.Recordset
   Dim Rs1 As New ADODB.Recordset
   Dim strSql
   Dim mFlag As Boolean
   strSql = "select * from postal where codigo='" & pCP & "' and codpcia='" & pCodPcia & "'"
   cmdTemp.CommandText = strSql
   cmdTemp.CommandType = 1
   cmdTemp.CommandTimeout = 36000
   Set cmdTemp.ActiveConnection = DBConn
   Set Rs1 = cmdTemp.Execute
   Set oLocalidadxCP_PCIA = Rs1
End Function

Public Sub xInsRegistros(ByVal pFecha As String, ByVal pHora As String, ByVal pEst As String, _
               ByVal pVia As String, ByVal pSent As String, ByVal pPatente As String, ByVal pModelo As String, _
               ByVal pColor As String, ByVal pObs As String, ByVal pPago As String, ByVal pUserAud As String, _
               ByVal pFechaAud As String, ByVal pHoraAud As String, ByVal pCodMarca As String, ByVal pFechaDelay As String, ByVal pTipo As String)
strSql1 = "INSERT INTO Registros VALUES ('" & Format(pFecha, "yyyy-mm-dd") & "','" & pHora & "','" & pEst & "'," _
        & "'" & pVia & pSent & "','" & pSent & "','" & pPatente & "','" & pModelo & "','" & pColor & "','" & pObs & "'," _
        & "'" & pPago & "','" & pUserAud & "','" & Format(pFechaAud, "yyyy-mm-dd") & "','" & pHoraAud & "'," _
        & "'" & pCodMarca & "','" & Format(pFechaDelay, "yyyy-mm-dd") & "', '" & pTipo & "')"
sAsignCmd
cmdTemp.Execute
End Sub
     
Public Function xUpdRegistros(ByVal pFecha As String, ByVal pHora As String, ByVal pEst As String, _
               ByVal pVia As String, ByVal pSent As String, ByVal pPatente As String, ByVal pModelo As String, _
               ByVal pColor As String, ByVal pObs As String, ByVal pCodMarca As String, ByVal pxFecha As String, _
               ByVal pxHora As String, ByVal pxEst As String, ByVal pxPatente As String) As Boolean
strSql1 = "update Registros set fecha='" & Format(pFecha, "yyyy-mm-dd") & "',hora='" & pHora & "'," _
        & "estacion='" & pEst & "',via='" & pVia & "',sentido='" & pSent & "',patente='" & pPatente & "'," _
        & "modelo='" & pModelo & "',color='" & pColor & "',obs='" & pObs & "',codmarca='" & pCodMarca & "' " _
        & "where fecha='" & Format(pxFecha, "yyyy-mm-dd") & "' and hora='" & pxHora & "' and estacion='" & pxEst & "' " _
        & "and patente='" & pxPatente & "'"
sAsignCmd
cmdTemp.Execute
xUpdRegistros = True
End Function
               
Public Function xInsertCP(ByVal pCodigo As String, ByVal pCodPcia As String, ByVal pDescripcion As String) As Boolean
   Dim Rs1 As New ADODB.Recordset
   Dim strSql
   Dim mFlag As Boolean
   strSql = "select * from postal where codigo = '" & pCodigo & "' and codpcia = '" & pCodPcia & "'"
   cmdTemp.CommandText = strSql
   cmdTemp.CommandType = 1
   cmdTemp.CommandTimeout = 36000
   Set cmdTemp.ActiveConnection = DBConn
   Set Rs1 = cmdTemp.Execute
   If Rs1.EOF Then
      strSql = "insert into postal values ('" & pCodigo & "','" & pCodPcia & "','" & pDescripcion & "')"
      cmdTemp.CommandText = strSql
      cmdTemp.Execute
      xInsertCP = True
   Else
      xInsertCP = False
   End If
End Function

Public Function xUpdateCP(ByVal pDescripcion As String, ByVal pCodPcia As String, ByVal pCodigo As String, ByVal pCodPciaOrig As String) As Boolean
   Dim strSql
   strSql = "update postal set descripcion='" & pDescripcion & "', codpcia='" & pCodPcia & "' where codigo='" & pCodigo & "' and codpcia='" & pCodPciaOrig & "'"
   cmdTemp.CommandText = strSql
   cmdTemp.CommandType = 1
   cmdTemp.CommandTimeout = 36000
   Set cmdTemp.ActiveConnection = DBConn
   cmdTemp.Execute
   xUpdateCP = True
End Function

Public Function xUpDirecciones(ByVal pName As String, ByVal pDomicilio As String, ByVal pCodPcia As String, ByVal pCodPostal As String, ByVal pPatente As String) As Boolean
   strSql1 = "update direcciones set nombre='" & pName & "', domicilio='" & pDomicilio & "',codpcia='" & pCodPcia & "',codpostal='" & pCodPostal & "' WHERE patente='" & pPatente & "'"
   sAsignCmd
   cmdTemp.Execute
   xUpDirecciones = True
End Function

Public Function xUpEnvios(ByVal pNroCarta As String, ByVal pFecha As String, ByVal pCodEntrega As String, ByVal pPatente As String, ByVal pFechaAct As String, ByVal pTipo As String) As Boolean
   strSql1 = "UPDATE envios SET nrocarta='" & pNroCarta & "', fecha='" & Format(pFecha, "yyyy-mm-dd") & "',codentrega='" & pCodEntrega & "' WHERE patente='" & pPatente & "' and fecha='" & Format(pFechaAct, "yyyy-mm-dd") & "' And tipo = '" & pTipo & "'"
   sAsignCmd
   cmdTemp.Execute
   xUpEnvios = True
End Function

Public Function xInsertRegPagos(ByVal pPatente As String, ByVal pObs As String, ByVal pEstado As String, ByVal pTipo As String) As Boolean
   Dim strSql
   strSql = "insert into regpagos values ('" & pPatente & "',now(),'" & pEstado & "','" & pObs & "','" & pTipo & "')"
   cmdTemp.CommandText = strSql
   cmdTemp.CommandType = 1
   cmdTemp.CommandTimeout = 36000
   Set cmdTemp.ActiveConnection = DBConn
   cmdTemp.Execute
   xInsertRegPagos = True
End Function

Public Function xUpdateHisRegPagos(ByVal pPatente As String, ByVal pEstado As String, ByVal pTipo As String) As Boolean
   Dim strSql
   strSql = "update regpagos set estado='" & pEstado & pEstado & "' where patente='" & pPatente & "' and estado='" & pEstado & "' And tipo = '" & pTipo & "'"
   cmdTemp.CommandText = strSql
   cmdTemp.CommandType = 1
   cmdTemp.CommandTimeout = 36000
   Set cmdTemp.ActiveConnection = DBConn
   cmdTemp.Execute
   xUpdateHisRegPagos = True
End Function

Public Function xInsDirecciones(ByVal pPatente As String, ByVal pName As String, ByVal pDomicilio As String, _
                                 ByVal pCodPcia As String, ByVal pCP As String) As Boolean
  
   strSql1 = "select * from direcciones where patente='" & pPatente & "'"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   If Rs1.EOF Then
      strSql1 = "insert into direcciones (patente, nombre, domicilio, codpcia, codpostal) values ('" & pPatente & "','" & pName & "','" & pDomicilio & "','" & pCodPcia & "','" & pCP & "')"
      sAsignCmd
      cmdTemp.Execute
      xInsDirecciones = True
   Else
      xInsDirecciones = False
      MsgBox "Ya existe la patente a ingresar.", vbExclamation, sMessage
   End If
End Function

Public Function xInsColores(ByVal pCodigo As String, ByVal pDescr As String) As Boolean
   strSql1 = "select * from colores where codigo='" & pCodigo & "'"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   If Rs1.EOF Then
      strSql1 = "insert into colores  values ('" & pCodigo & "','" & pDescr & "',null)"
      sAsignCmd
      cmdTemp.Execute
      xInsColores = True
      Exit Function
   Else
      xInsColores = False
      MsgBox "Ya existe el código a ingresar.", vbExclamation, sMessage
   End If
End Function

Public Function xInsEnvios(ByVal pPatente As String, ByVal pNroCarta As String, ByVal pFecha As String, _
                     ByVal pCodEntrega As String, ByVal pObs As String, ByVal pTipo As String) As Boolean
strSql1 = "insert into envios (patente,nrocarta,fecha,codentrega,obs,tipo) values " _
        & " ('" & pPatente & "','" & pNroCarta & "','" & Format(pFecha, "yyyy-mm-dd") & "','" & pCodEntrega & "','" & pObs & "','" & pTipo & "')"
sAsignCmd
cmdTemp.Execute
xInsEnvios = True
End Function

Public Function xInsMarcas(ByVal pCodigo As String, ByVal pDescr As String) As Boolean
strSql1 = "select * from marcas where codigo='" & pCodigo & "'"
sAsignCmd
Set Rs1 = cmdTemp.Execute
If Rs1.EOF Then
   strSql1 = "insert into marcas  values ('" & pCodigo & "','" & pDescr & "',null)"
   sAsignCmd
   cmdTemp.Execute
   xInsMarcas = True
Else
   xInsMarcas = False
   MsgBox "Ya existe el código a ingresar.", vbExclamation, sMessage
End If
End Function

Public Function xInsModelos(ByVal pCodigo As String, ByVal pCodMarca As String, ByVal pDescr As String) As Boolean
strSql1 = "select * from modelos where codigo='" & pCodigo & "' and codmarca='" & pCodMarca & "'"
sAsignCmd
Set Rs1 = cmdTemp.Execute
If Rs1.EOF Then
   strSql1 = "insert into modelos  values ('" & pCodMarca & "','" & pCodigo & "','" & pDescr & "',null)"
   sAsignCmd
   cmdTemp.Execute
   xInsModelos = True
Else
   xInsModelos = False
   MsgBox "Ya existe el código a ingresar.", vbExclamation, sMessage
End If
End Function

Public Function xUpColores(ByVal pCodigo As String, ByVal pDescr As String, ByVal pNull As Boolean)
Dim mUpDato As String
mUpDato = ""
If pNull Then
   mUpDato = ", baja= now() "
End If
strSql1 = "update colores set descripcion='" & pDescr & "' " & mUpDato & " where codigo='" & pCodigo & "'"
sAsignCmd
cmdTemp.Execute
End Function

Public Function xUpMarcas(ByVal pCodigo As String, ByVal pDescr As String, ByVal pNull As String)
Dim mUpDato As String
mUpDato = ""
If pNull Then
   mUpDato = ", baja= now() "
End If
strSql1 = "update marcas set descripcion='" & pDescr & "' " & mUpDato & " where codigo='" & pCodigo & "'"
sAsignCmd
cmdTemp.Execute
End Function

Public Function xUpModelos(ByVal pCodigo As String, ByVal pCodMarca As String, ByVal pDescr As String, ByVal pNull As String)
   Dim mUpDato As String
   mUpDato = ""
   If pNull Then
      mUpDato = ", baja= now() "
   End If
   strSql1 = "update modelos set descripcion='" & pDescr & "' " & mUpDato & " where codigo='" & pCodigo & "' and codmarca='" & pCodMarca & "'"
   sAsignCmd
   cmdTemp.Execute
End Function

Public Function xUpDatosVehic(ByVal pPatente As String, ByVal pCodMarca As String, ByVal pCodModelo As String, ByVal pCodColor As String)
   strSql1 = "update Registros set codmarca='" & pCodMarca & "', modelo='" & pCodModelo & "', " _
      & " color='" & pCodColor & "' where patente='" & pPatente & "'"
   sAsignCmd
   On Error Resume Next
   cmdTemp.Execute
   If Err.Description <> "" Then
      MsgBox "error en el sistema: " & Chr(13) & Err.Description, vbCritical, sMessage
   Else
      MsgBox "actualización exitosa.", vbInformation, sMessage
   End If
End Function
