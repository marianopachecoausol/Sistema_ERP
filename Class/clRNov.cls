VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clRNov"
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
   Set DBConn = mObj.oAbreConexion("5", mIPServer, "regnov", "ssvv", "ssvv")
   'Set DBConn = mObj.oAbreConexion("4", "192.168.2.155", "regnov", "ssvv", "ssvv")
   'Set DBConn = mObj.oAbreConexion("4", "10.10.20.11", "regnov", "root", "123456")
End Sub

Private Sub class_terminate()
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

Public Function oListOperativos(ByVal pDate1 As Date, ByVal pDate2 As Date) As ADODB.Recordset
   Dim mRec As New ADODB.Recordset
   Dim mQuery As String
   mRec.CursorType = adOpenDynamic
   mQuery = "SELECT  RIGHT(descripcion,4) as Movil, Count(fecha) as Total FROM novedades WHERE fecha BETWEEN '" & Format(pDate1, "yyyy-mm-dd") & "' and '" & Format(pDate2, "yyyy-mm-dd") & "' AND CodNov='P' GROUP BY RIGHT(descripcion,4)"
   cmdTemp.CommandText = mQuery
   cmdTemp.CommandType = 1
   Set cmdTemp.ActiveConnection = DBConn
   cmdTemp.CommandTimeout = 36000
   Set mRec = cmdTemp.Execute
   Set oListOperativos = mRec
End Function

Public Function oMovilTipo(ByVal pTipo As String) As ADODB.Recordset  'todos los colores
   Dim Rs1, strSql1
   Set Rs1 = New ADODB.Recordset
   strSql1 = "select * from moviles where codtipomov='" & pTipo & "' order by 2"
   cmdTemp.CommandText = strSql1
   cmdTemp.CommandType = 1
   cmdTemp.CommandTimeout = 36000
   Set cmdTemp.ActiveConnection = DBConn
   Set Rs1 = cmdTemp.Execute
   Set oMovilTipo = Rs1
End Function

Public Function oTotalTareas(ByVal pFecha1 As String, ByVal pFecha2 As String) As ADODB.Recordset   'todos los colores
   Dim Rs1, strSql1
   Set Rs1 = New ADODB.Recordset
   strSql1 = "select O.codigo, O.descripcion, count(left(N.Descripcion,4)) as total from novedades N, otros O " _
         & " where O.Codigo = Left(N.Descripcion, 4) and N.fecha between '" & Format(pFecha1, "yyyy-mm-dd") & " 00:00:00' " _
         & " and '" & Format(pFecha2, "yyyy-mm-dd") & " 23:59:59' group by O.codigo"
   cmdTemp.CommandText = strSql1
   cmdTemp.CommandType = 1
   cmdTemp.CommandTimeout = 36000
   Set cmdTemp.ActiveConnection = DBConn
   Set Rs1 = cmdTemp.Execute
   Set oTotalTareas = Rs1
End Function

Public Function oMovilesGCO(ByVal pTipo As String) As ADODB.Recordset
   strSql1 = "select * from moviles where CodTipoMov in ('" & pTipo & "') order by Descripcion"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oMovilesGCO = Rs1
End Function

Public Function oMovilesPATGRU() As ADODB.Recordset
   strSql1 = "select * from moviles where CodTipoMov IN ('PAT','GRU') AND left(Codigo,2) in('G0','M0') order by Descripcion"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oMovilesPATGRU = Rs1
End Function



Public Function oServMoviles2(ByVal pMov As String, ByVal pFecha1 As String, ByVal pFecha2 As String) As ADODB.Recordset
    'strSql1 = "SELECT Left(Fecha,10) as mFecha,Fecha,Codigo,Km,Sent,CodNov,Demora, CodServ1, codramal FROM novedades N " _

    strSql1 = "SELECT Left(CAST(Fecha AS CHAR),10)  as mFecha,Fecha,Codigo,Km,Sent,CodNov,Demora, CodServ1, codramal FROM novedades N " _
    & " WHERE N.Fecha between '" & Format(pFecha1, "yyyy-mm-dd 00:00:00") & "' AND '" & Format(pFecha2, "yyyy-mm-dd 23:59:59") & "' " _
    & " AND N.Codigo <> ''AND N.Codigo IS NOT NULL " _
    & " AND (Mov1 = '" & pMov & "' OR Mov2 = '" & pMov & "' OR Mov3 = '" & pMov & "' ) " _
    & " AND NOT EXISTS (SELECT 1 FROM novedades " _
    & "                   WHERE  Fecha between '" & Format(pFecha1, "yyyy-mm-dd 00:00:00") & "' AND '" & Format(pFecha2, "yyyy-mm-dd 23:59:59") & "'" _
    & " AND ((CodNov = 'MM' AND Mov1 = N.Mov1 AND CodNov = N.CodNov and codigo = N.Codigo and Fecha < N.Fecha) " _
    & "or (CodNov = 'MM' AND Mov1 = N.Mov1  AND N.CodNov = 'A' and codigo = N.Codigo  and Fecha < N.Fecha) " _
    & "or (CodNov = 'A' AND Mov1 = N.Mov1  AND N.CodNov = 'MM' and codigo = N.Codigo  and Fecha < N.Fecha) " _
    & " or (CodNov = 'L' AND Mov1 = N.Mov1 AND CodNov = N.CodNov and codigo = N.Codigo and Fecha > N.Fecha)) " _
    & " OR  (CodNov in ('MM','A') AND Mov1 = N.Mov1 AND N.CodNov in ('N','QQ','RR')  and codigo = N.Codigo and Fecha > N.Fecha )) " _
    & " ORDER BY  codigo, Left(Fecha,10),fecha "
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oServMoviles2 = Rs1
End Function

Public Function oServMoviles4(ByVal pCodigo As String, pMov As String, pFecha1 As String, ByVal pFecha2 As String) As ADODB.Recordset
  strSql1 = "SELECT Left(Fecha,10) as mFecha,Fecha FROM novedades N " _
    & " WHERE N.Fecha between '" & Format(pFecha1, "yyyy-mm-dd 00:00:00") & "' AND '" & Format(pFecha2, "yyyy-mm-dd 23:59:59") & "' " _
    & " AND N.Codigo = '" & pCodigo & "' " _
    & " AND N.CodNov = 'E' " _
    & " AND (Mov1 = '" & pMov & "' OR Mov2 = '" & pMov & "' OR Mov3 = '" & pMov & "' ) " _
    & " ORDER BY  Left(Fecha,10),fecha, codigo "
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oServMoviles4 = Rs1
End Function




Public Function oNovedDistCampo(ByVal pFecha1 As String, ByVal pFecha2 As String, ByVal pCampo As String, ByVal pWhere As String) As ADODB.Recordset
   strSql1 = "select distinct " & pCampo & " from novedades where fecha between '" & Format(pFecha1, "yyyy-mm-dd") & " 00:00:00' and '" & Format(pFecha2, "yyyy-mm-dd") & " 23:59:59' " & pWhere
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oNovedDistCampo = Rs1
End Function

Public Function oNovedadesMovil(ByVal pFecha1 As String, ByVal pFecha2 As String, ByVal pMoviles As String) As ADODB.Recordset
   strSql1 = "select novedades.Fecha,novedades.Codigo, novedades.codramal,novedades.Km,novedades.Sent,novedades.Descripcion,CONCAT_WS('-',clima.Codigo,clima.Detalle) AS climax,novedades.Mov1,novedades.Mov2,novedades.Mov3 " _
      & " from novedades LEFT JOIN clima ON clima.Codigo = novedades.CodClima WHERE (Mov1 IN (" & pMoviles & ") OR Mov2 IN (" & pMoviles & ") OR Mov3 IN (" & pMoviles & ")) " _
      & " AND novedades.Fecha between '" & Format(pFecha1 & ":00", "yyyy-mm-dd hh:mm:ss") & "' and '" & Format(pFecha2 & ":59", "yyyy-mm-dd hh:mm:ss") & "' order by Fecha"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oNovedadesMovil = Rs1
End Function

Public Function oNovedadesFecha(ByVal pFecha1 As String, ByVal pFecha2 As String, ByVal pwCodNov As String, ByVal pwCodMovil As String, ByVal pOrder As String) As ADODB.Recordset
   strSql1 = "select * from novedades where Fecha between '" & Format(pFecha1, "yyyy/mm/dd") & " 00:00:00' and '" & Format(pFecha2, "yyyy/mm/dd") & " 23:59:59' " _
      & pwCodNov & " " & pwCodMovil & " " & pOrder
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oNovedadesFecha = Rs1
End Function

Public Function oNovedadesFechaHora(ByVal pFecha1 As String, ByVal pFecha2 As String, ByVal pwCodNov As String, ByVal pwCodMovil As String, ByVal pOrder As String) As ADODB.Recordset
   strSql1 = "select * from novedades where Fecha between '" & Format(pFecha1, "yyyy/mm/dd hh:mm:ss") & "' and '" & Format(pFecha2, "yyyy/mm/dd hh:mm:ss") & "' " _
      & pwCodNov & " " & pwCodMovil & " " & pOrder
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oNovedadesFechaHora = Rs1
End Function

Public Function oCodAlfaMinFecha(ByVal pFecha1 As String, ByVal pFecha2 As String) As ADODB.Recordset
   strSql1 = "select distinct(codigo) as codigo, min(fecha) as fecha from novedades " _
   & " where Fecha between '" & Format(pFecha1, "yyyy/mm/dd hh:mm:ss") & "' and '" & Format(pFecha2, "yyyy/mm/dd hh:mm:ss") & "' " _
   & " and codigo<>'' group by codigo order by fecha"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oCodAlfaMinFecha = Rs1
End Function

Public Function oNovedCodigoFecha(ByVal pwCodigo As String, ByVal pFecha1 As String, ByVal pFecha2 As String) As ADODB.Recordset
Dim mWhere As String
   mWhere = ""
   If Trim(pFecha1) <> "" And Trim(pFecha2) <> "" Then
      mWhere = " and Fecha between '" & Format(pFecha1, "yyyy/mm/dd") & " 00:00:00' and '" & Format(pFecha2, "yyyy/mm/dd") & " 23:59:59' "
   End If
   'strSql1 = "SELECT DISTINCT Codigo, left(Fecha,10) as mFecha FROM novedades WHERE " & pwCodigo & mWhere & " ORDER BY Left(Fecha,10), Codigo"
    strSql1 = "SELECT DISTINCT Codigo, Left(CAST(Fecha AS CHAR),10) as mFecha FROM novedades WHERE " & pwCodigo & mWhere & " ORDER BY Left(Fecha,10), Codigo"
   
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oNovedCodigoFecha = Rs1
End Function

Public Function oMovilKms(ByVal pFecha1 As String, ByVal pFecha2 As String, ByVal pCodMovil As String) As ADODB.Recordset
   strSql1 = "SELECT distinct movilturno.FechaInic,movilturno.CodMovil,movilturno.KmInicial,movilturno.KmFinal,movilturno.Chofer,CONCAT_WS('-',patrulleros.Codigo,patrulleros.Nombre) AS mPatrullero, turnos.Descripcion " _
      & " from ((movilturno LEFT JOIN patrulleros ON patrulleros.Codigo = movilturno.CodPatrullero) LEFT JOIN turnos ON turnos.Codigo = movilturno.CodTurno) " _
      & " where movilturno.FechaInic between '" & Format(pFecha1, "yyyy-mm-dd") & " 00:00:00' and '" & Format(pFecha2, "yyyy-mm-dd") & " 23:59:59' AND movilturno.KmFinal IS NOT NULL AND movilturno.KmFinal <> '' "
      If pCodMovil <> "" Then
         strSql1 = strSql1 & "and movilturno.CodMovil in ('" & pCodMovil & "') "
      End If
      strSql1 = strSql1 & " order by fechainic,codmovil ASC"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oMovilKms = Rs1
End Function

Public Function oObtOrigen(ByVal pWhere As String) As ADODB.Recordset
   strSql1 = "select CodOrigen FROM novedades WHERE " & pWhere & " ORDER BY Fecha ASC LIMIT 1"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oObtOrigen = Rs1
End Function

Public Function oMovilesCodigo(ByVal pCodigo As String, ByVal pFecha As String) As ADODB.Recordset
   'strSql1 = "select Mov1, Mov2, Mov3  FROM novedades2 WHERE codigo='" & pCodigo & "'"
   strSql1 = "select Mov1, Mov2, Mov3  FROM novedades2 WHERE codigo='" & pCodigo & "' and fecha='" & Format(pFecha, "yyyy-mm-dd hh:mm:ss") & "'"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oMovilesCodigo = Rs1
End Function


Public Function oCodigosAmbu(ByVal pFecha1 As String, ByVal pFecha2 As String) As ADODB.Recordset
   Dim Rs1, strSql1
   Set Rs1 = New ADODB.Recordset
   strSql1 = "SELECT Codigo, Fecha, CodNov, CodServ1 From novedades WHERE Fecha between '" & Format(pFecha1, "yyyy-mm-dd") & " 00:00:00" & "'" _
      & " and '" & Format(pFecha2, "yyyy-mm-dd") & " 23:59:59' and (Mov1 = 'AMBU' or Mov1 = 'AMB1') And CodNov in ('M','L') ORDER BY Codigo,Fecha"
   cmdTemp.CommandText = strSql1
   cmdTemp.CommandType = 1
   cmdTemp.CommandTimeout = 36000
   Set cmdTemp.ActiveConnection = DBConn
   Set Rs1 = cmdTemp.Execute
   Set oCodigosAmbu = Rs1
End Function

Public Function oMovilesXCodigo(ByVal pCodigo As String, ByVal pFecha As String) As ADODB.Recordset
   strSql1 = "select Mov1, Mov2, Mov3 from novedades where codigo='" & pCodigo & "' and fecha = '" & Format(pFecha, "yyyy-mm-dd hh:mm:ss") & "'"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oMovilesXCodigo = Rs1
End Function

Public Function oMovilesDistTool(ByVal pEstado As String, ByVal pWhere As String) As ADODB.Recordset
   strSql1 = "select distinct tooltip from moviles where tooltip<>'' and estado='" & pEstado & "' " & pWhere & " order by tooltip asc"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oMovilesDistTool = Rs1
End Function

Public Function oMovilTurnoChofer(ByVal pCodMovil As String) As ADODB.Recordset
   strSql1 = "select movilturno.CodMovil,movilturno.Chofer,movilturno.FechaInic,movilturno.KmInicial,movilturno.KmFinal," _
         & "turnos.Descripcion, movilturno.codpatrullero2 FROM movilturno LEFT JOIN turnos ON turnos.Codigo = movilturno.CodTurno " _
         & "where CodMovil = '" & pCodMovil & "' ORDER BY FechaInic DESC"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oMovilTurnoChofer = Rs1
End Function

Public Function oMovilTurnoFecha(ByVal pFecha As String, ByVal pCodMovil As String) As ADODB.Recordset
   strSql1 = "select A.*, B.Descripcion FROM movilturno A INNER JOIN turnos B ON A.CodTurno = B.Codigo " _
      & " WHERE A.CodMovil = '" & pCodMovil & "' AND A.FechaInic < '" & Format(pFecha, "yyyy/mm/dd hh:mm:ss") & "' ORDER BY A.FechaInic DESC"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oMovilTurnoFecha = Rs1
End Function

Public Function oMovilTurnoPatruFecha(ByVal pFecha As String, ByVal pCodMovil As String) As ADODB.Recordset
   strSql1 = "select A.*, B.Nombre, C.Descripcion FROM (movilturno A INNER JOIN patrulleros B ON B.Codigo = A.CodPatrullero) LEFT JOIN turnos C " _
   & " ON C.Codigo = A.CodTurno WHERE A.CodMovil = '" & pCodMovil & "' AND A.FechaInic < '" & Format(pFecha, "yyyy/mm/dd hh:mm:ss") & "' ORDER BY A.FechaInic DESC"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oMovilTurnoPatruFecha = Rs1
End Function

Public Function oMovilTurnoPatru(ByVal pCodMovil As String) As ADODB.Recordset
   strSql1 = "select movilturno.CodMovil,patrulleros.Nombre,movilturno.FechaInic,movilturno.KmInicial,movilturno.KmFinal,turnos.Descripcion, movilturno.Codpatrullero2, movilturno.Codpatrullero3  " _
            & " FROM ((movilturno LEFT JOIN patrulleros ON movilturno.CodPatrullero =patrulleros.Codigo) LEFT JOIN turnos ON turnos.Codigo = movilturno.CodTurno) " _
         & "where CodMovil = '" & pCodMovil & "' ORDER BY FechaInic DESC"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oMovilTurnoPatru = Rs1
End Function

Public Function oMovilTurnos(ByVal pCodMovil As String, ByVal pLimit As Integer) As ADODB.Recordset
   strSql1 = "select A.* ,B.Nombre,C.Descripcion FROM ((movilturno A LEFT JOIN patrulleros B ON A.CodPatrullero = B.Codigo)" _
      & "LEFT JOIN turnos C ON C.Codigo = A.CodTurno) where A.CodMovil = '" & pCodMovil & "' ORDER BY A.FechaInic DESC limit " & pLimit
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oMovilTurnos = Rs1
End Function

Public Function oNovedCountCodOrig(ByVal pFecha1 As String, ByVal pFecha2 As String, ByVal pCampo As String, ByVal pCodOrigen As String) As ADODB.Recordset
   strSql1 = "select " & pCampo & ", Count(*) AS Total from novedades where Fecha between '" & Format(pFecha1, "yyyy/mm/dd") & " 00:00:00' and '" & Format(pFecha2, "yyyy/mm/dd") & " 23:59:59' " & pCodOrigen & " GROUP BY " & pCampo & " ORDER BY " & pCampo
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oNovedCountCodOrig = Rs1
End Function

Public Function iCountCodOrig(ByVal pFecha1 As String, ByVal pFecha2 As String, ByVal pCodOrigen As String) As ADODB.Recordset
   strSql1 = "select A.CodOrigen, O.descripcion, Count(A.fecha) AS Total from novedades A, origen O " _
      & " where  A.Fecha between '" & Format(pFecha1, "yyyy/mm/dd") & " 00:00:00' and '" & Format(pFecha2, "yyyy/mm/dd") & " 23:59:59' and A.codorigen=O.codigo " _
      & " and A.codnov='A' " & pCodOrigen & " GROUP BY A.CodOrigen ORDER BY A.CodOrigen"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set iCountCodOrig = Rs1
End Function

Public Function oHoraFinAcc(ByVal pFecha As String, ByVal pCodAlfa As String) As String
   strSql1 = "select * from novedades where codigo='" & pCodAlfa & "' and left(Mov1,2)='M0' " _
      & " and date_format(fecha,'%m')='" & Format(Month(pFecha), "00") & "' and date_format(fecha,'%Y')='" & Format(Year(pFecha), "0000") & "' order by fecha desc"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   If Not Rs1.EOF Then
      oHoraFinAcc = Format(Rs1!Fecha, "hh:mm")
   Else
      oHoraFinAcc = ""
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

Public Function iCountNov2() As Integer
   iCountNov2 = 0
   strSql1 = "select count(fecha) as total from novedades2"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   If Not Rs1.EOF Then
      iCountNov2 = Rs1!Total
   End If
End Function
Public Function iCountNovPendientes() As Integer
   iCountNovPendientes = 0
   strSql1 = "select count(fecha) as total from novedades2 where finalizado = 0"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   If Not Rs1.EOF Then
      iCountNovPendientes = Rs1!Total
   End If
End Function



Public Function oCountServMes(ByVal pMes As String, ByVal pAnio As String) As ADODB.Recordset
   strSql1 = "select N.codserv1, count(distinct N.fecha) as total from novedades N, servicios S " _
      & " where month(N.fecha)=" & pMes & " and year(N.fecha)=" & pAnio & " and N.codserv1=S.codigo group by N.codserv1"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oCountServMes = Rs1
End Function

Public Function iCountTareas(ByVal pMes As String, ByVal pAnio As String, ByVal pCodNov As String, ByVal pTiposObj As String) As Integer
   iCountTareas = 0
   strSql1 = "SELECT COUNT(Fecha) AS Total From novedades WHERE CodNov IN ('" & pCodNov & "') " _
      & " and month(fecha)=" & pMes & " and year(fecha)=" & pAnio & " and (Left(Descripcion,3) IN (" & pTiposObj & "))"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   If Not Rs1.EOF Then
      iCountTareas = Rs1!Total
   End If
End Function

Public Function oNovedAnios() As ADODB.Recordset
   strSql1 = "SELECT distinct(year(fecha)) as anios FROM novedades order by 1"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oNovedAnios = Rs1
End Function

Public Function iCountServicios(ByVal pFecha1 As String, ByVal pFecha2 As String, ByVal pCodNov As String, _
                        ByVal pCodMovil As String, ByVal pLeftMovil As String) As Integer
Dim mWhere As String
   If Trim(pLeftMovil) <> "" Then
      mWhere = " and ((left(Mov1," & Len(pLeftMovil) & ")='" & pLeftMovil & "' or left(Mov2," & Len(pLeftMovil) & ")='" & pLeftMovil & "' or left(Mov3," & Len(pLeftMovil) & ")='" & pLeftMovil & "')) "
   End If
   If Trim(pCodMovil) <> "" Then
      If mWhere <> "" Then
         mWhere = " and (" & Mid(mWhere, 5, Len(mWhere) - 5) & " or (Mov1 = '" & pCodMovil & "' or Mov2 = '" & pCodMovil & "' or Mov3 = '" & pCodMovil & "'))"
      Else
         mWhere = " and (Mov1 = '" & pCodMovil & "' or Mov2 = '" & pCodMovil & "' or Mov3 = '" & pCodMovil & "')"
      End If
   End If
   strSql1 = "select count(fecha) as Total from novedades WHERE CodNov in (" & pCodNov & ") " _
      & "and fecha between '" & Format(pFecha1, "yyyy-mm-dd") & " 00:00:00' and '" & Format(pFecha2, "yyyy-mm-dd") & " 23:59:59' " _
      & mWhere
   sAsignCmd
   iCountServicios = 0
   Set Rs1 = cmdTemp.Execute
   If Not Rs1.EOF Then
     iCountServicios = Rs1!Total
   End If
End Function

Public Function iCountServMovil(ByVal pFecha1 As String, ByVal pFecha2 As String, ByVal pCodNov As String, _
                        ByVal pMovil As String, ByVal pLeftMovil As String) As Integer
Dim mWhere As String
   mWhere = ""
   If Trim(pLeftMovil) <> "" Then
      mWhere = " and " & pLeftMovil
   End If
   If Trim(pMovil) <> "" Then
      If mWhere <> "" Then
         mWhere = " and (" & Mid(mWhere, 5, Len(mWhere) - 4) & " or " & pMovil & ") "
      Else
         mWhere = " and " & pMovil
      End If
   End If
   strSql1 = "select count(fecha) as Total from novedades WHERE CodNov in (" & pCodNov & ") " _
      & "and fecha between '" & Format(pFecha1, "yyyy-mm-dd") & " 00:00:00' and '" & Format(pFecha2, "yyyy-mm-dd") & " 23:59:59' " _
      & mWhere
   sAsignCmd
   iCountServMovil = 0
   Set Rs1 = cmdTemp.Execute
   If Not Rs1.EOF Then
     iCountServMovil = Rs1!Total
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

Public Function xUpNovedadesSet(ByVal pSets As String, ByVal pWhere As String)
   strSql1 = "update novedades2 set " & pSets & " where " & pWhere
   sAsignCmd
   cmdTemp.Execute
End Function

Public Function xUpUltimos(ByVal pNroParte As Integer)
   strSql1 = "update ultimos set nroparte='" & pNroParte + 1 & "'"
   sAsignCmd
   cmdTemp.Execute
End Function

Public Function xUpNovedCodAlfa(ByVal pFecha As String, ByVal pCodAlfa As String)
   strSql1 = "update novedades set codigo='" & pCodAlfa & "' where Fecha='" & Format(pFecha, "yyyy-mm-dd hh:mm:ss") & "'"
   sAsignCmd
   cmdTemp.Execute
   strSql1 = "update novedades2 set codigo='" & pCodAlfa & "' where Fecha='" & Format(pFecha, "yyyy-mm-dd hh:mm:ss") & "'"
   sAsignCmd
   cmdTemp.Execute
End Function

Public Function xUpNovedCodNov(ByVal pFecha As String, ByVal pCodAlfa As String, ByVal pCodNov As String)
   strSql1 = "update novedades set codnov='" & pCodNov & "' where Fecha='" & Format(pFecha, "yyyy-mm-dd hh:mm:ss") & "' and codigo='" & pCodAlfa & "'"
   sAsignCmd
   cmdTemp.Execute
   strSql1 = "update novedades2 set codnov='" & pCodNov & "' where Fecha='" & Format(pFecha, "yyyy-mm-dd hh:mm:ss") & "' and codigo='" & pCodAlfa & "'"
   sAsignCmd
   cmdTemp.Execute
End Function

Public Function xUpActualizar(ByVal pPC As String, ByVal pFlag As Integer)
   strSql1 = "update actualizar set flag=" & pFlag & " where pc='" & pPC & "'"
   sAsignCmd
   cmdTemp.Execute
End Function

Public Function xUpActualizarNot(ByVal pPC As String, ByVal pFlag As Integer)
   strSql1 = "update actualizar set flag=" & pFlag & " where pc <> '" & pPC & "'"
   sAsignCmd
   cmdTemp.Execute
End Function

Public Function xUpClimaTag(ByVal pImagen As String, ByVal pIndice As Integer)
   strSql1 = "update climatag set tag=" & pIndice & " where IndImagen='" & pImagen & "'"
   sAsignCmd
   cmdTemp.Execute
End Function

Public Function xUpCodigos(ByVal pCodigo As String, ByVal pEstado As String)
   strSql1 = "update codigos set estado='" & pEstado & "' where codalfa='" & pCodigo & "'"
   sAsignCmd
   cmdTemp.Execute
End Function
       
Public Function xUpEstMoviles(ByVal pCodigo As String, ByVal pEstado As String, ByVal pToolTip As String, ByVal pCodNov As String)
   strSql1 = "update moviles set estado='" & pEstado & "', tooltip='" & pToolTip & "', codnov='" & pCodNov & "' where codigo='" & pCodigo & "'"
   sAsignCmd
   cmdTemp.Execute
End Function

Public Function xUpEstToolMov(ByVal pWhere As String, ByVal pEstado As String, ByVal pToolTip As String)
   strSql1 = "update moviles set estado='" & pEstado & "', tooltip='" & pToolTip & "' where " & pWhere
   sAsignCmd
   cmdTemp.Execute
End Function

Public Function xUpToolMov(ByVal pWhere As String, ByVal pToolTip As String)
   strSql1 = "update moviles set " & pToolTip & " where " & pWhere
   sAsignCmd
   cmdTemp.Execute
End Function

Public Function xUpMovilesCodNov(ByVal pCodigo As String, ByVal pCodNov As String)
   strSql1 = "update moviles set codnov='" & pCodNov & "' where codigo='" & pCodigo & "'"
   sAsignCmd
   cmdTemp.Execute
End Function

Public Function xInsNovedades(ByVal pFecha As String, ByVal pCodigo As String, ByVal pCodOper As String, _
               ByVal pKm As String, ByVal pSent As String, ByVal pCodOrigen As String, ByVal pDescr As String, _
               ByVal pCodNov As String, ByVal pCodClima As String, ByVal pCodMov1 As String, ByVal pCodServ1 As String, _
               ByVal pParte1 As Integer, ByVal pCodMovil2 As String, ByVal pCodServ2 As String, ByVal pParte2 As Integer, _
               ByVal pCodMovil3 As String, ByVal pCodServ3 As String, ByVal pParte3 As Integer, ByVal pDemora As String, ByVal pTipoNov As String, ByVal pCodRamal As String, ByVal pCodReferencia) As Boolean
Dim mFecha As String
   If Trim(pDemora) = "" Then
      pDemora = "0000-00-00 00:00:00"
   Else
      pDemora = Format(pDemora, "yyyy-mm-dd hh:mm:ss")
   End If
   If Trim(pFecha) = "" Then
      mFecha = "NOW()"
   Else
      mFecha = "'" & Format(pFecha, "yyyy-mm-dd hh:mm:ss") & "'"
   End If
   
   pKm = Replace(pKm, ",", ".") 'mp 20160309
   
   strSql1 = "insert into novedades2 (fecha,codigo,codoper,km,sent,codorigen,descripcion,codnov,codclima," _
   & "mov1,codserv1,parte1,mov2,codserv2,parte2,mov3,codserv3,parte3,demora,tiponov, codramal,codreferencia) values " _
   & " (" & mFecha & ",'" & pCodigo & "','" & pCodOper & "'," & pKm & "," & pSent & "," _
   & "'" & pCodOrigen & "','" & pDescr & "','" & pCodNov & "','" & pCodClima & "','" & pCodMov1 & "','" & pCodServ1 & "'," & pParte1 & ",'" & pCodMovil2 & "'," _
   & "'" & pCodServ2 & "'," & pParte2 & ",'" & pCodMovil3 & "','" & pCodServ3 & "'," & pParte3 & ",'" & pDemora & "','" & pTipoNov & "'," & pCodRamal & "," & pCodReferencia & ")"
   sAsignCmd
   cmdTemp.Execute
   xInsNovedades = True
End Function

Public Function xInsNov_old(ByVal pFecha As String, ByVal pCodigo As String, ByVal pCodOper As String, _
               ByVal pKm As String, ByVal pSent As String, ByVal pCodOrigen As String, ByVal pDescr As String, _
               ByVal pCodNov As String, ByVal pCodClima As String, ByVal pCodMov1 As String, ByVal pCodServ1 As String, _
               ByVal pParte1 As Integer, ByVal pCodMovil2 As String, ByVal pCodServ2 As String, ByVal pParte2 As Integer, _
               ByVal pCodMovil3 As String, ByVal pCodServ3 As String, ByVal pParte3 As Integer, ByVal pDemora As String, ByVal pTipoNov As String, ByVal pCodRamal As String, pCodReferencia As String, pFinalizado As Integer) As Boolean
Dim mFecha As String
   If Trim(pDemora) = "" Then
      pDemora = "0000-00-00 00:00:00"
   Else
      pDemora = Format(pDemora, "yyyy-mm-dd hh:mm:ss")
   End If
   If Trim(pFecha) = "" Then
      mFecha = "NOW()"
   Else
      mFecha = "'" & Format(pFecha, "yyyy-mm-dd hh:mm:ss") & "'"
   End If
   pKm = Replace(pKm, ",", ".")
   strSql1 = "insert into novedades (fecha,codigo,codoper,km,sent,codorigen,descripcion,codnov,codclima," _
   & "mov1,codserv1,parte1,mov2,codserv2,parte2,mov3,codserv3,parte3,demora,tiponov, codramal, codreferencia,finalizado) values " _
   & " (" & mFecha & ",'" & pCodigo & "','" & pCodOper & "'," & pKm & ",'" & pSent & "'," _
   & "'" & pCodOrigen & "','" & pDescr & "','" & pCodNov & "','" & pCodClima & "','" & pCodMov1 & "','" & pCodServ1 & "'," & pParte1 & ",'" & pCodMovil2 & "'," _
   & "'" & pCodServ2 & "'," & pParte2 & ",'" & pCodMovil3 & "','" & pCodServ3 & "'," & pParte3 & ",'" & pDemora & "','" & pTipoNov & "'," & pCodRamal & "," & pCodReferencia & "," & pFinalizado & ")"
   sAsignCmd
   On Error Resume Next
   cmdTemp.Execute
   If Err.Description <> "" Then
      MsgBox Err.Description
      mFecha = mFecha
   End If
   xInsNov_old = True
End Function

Public Function xUpMovilTurnosFin(ByVal pFechaInic As String, ByVal pCodMovil As String, ByVal pFechaFin As String, ByVal pKmFin As String)
   strSql1 = "update movilturno set fechafinal='" & Format(pFechaFin, "yyyy-mm-dd hh:mm:ss") & "', kmfinal='" & pKmFin & "' " _
           & "where fechainic='" & Format(pFechaInic, "yyyy-mm-dd hh:mm:ss") & "' and codmovil='" & pCodMovil & "'"
   sAsignCmd
   cmdTemp.Execute
End Function

Public Function xUpMovilTurnos(ByVal pFechaInic As String, ByVal pCodPatrullero As String, ByVal pChofer As String, _
                               ByVal pPolice As String, ByVal pKmIni As String, ByVal pCodTurno As String, ByVal pCodPat2 As String, ByVal pCodPat3 As String)
   strSql1 = "update movilturno set codpatrullero='" & pCodPatrullero & "', chofer='" & pChofer & "', police='" & pPolice & "', " _
           & "kminicial='" & pKmIni & "', codturno='" & pCodTurno & "', codpatrullero2='" & pCodPat2 & "', codpatrullero3='" & pCodPat3 & "' where fechainic='" & Format(pFechaInic, "yyyy-mm-dd hh:mm:ss") & "'"
   sAsignCmd
   cmdTemp.Execute
End Function

Public Function xUpTurnosKm(ByVal pCodMovil As String, ByVal pFechaInic As String, pKmI As String, ByVal pKmF As String)
   strSql1 = "update movilturno set kminicial='" & pKmI & "', kmfinal='" & pKmF & "' where codmovil='" & pCodMovil & "' and fechainic='" & Format(pFechaInic, "yyyy-mm-dd hh:mm:ss") & "'"
   sAsignCmd
   cmdTemp.Execute
End Function

Public Function xInsMovilTurno(ByVal pCodMovil As String, ByVal pPatrullero As String, ByVal pChofer As String, _
               ByVal pPolice As String, ByVal pFechaIni As String, ByVal pKmIni As String, ByVal pCodTurno As String, ByVal pCodPat2 As String, ByVal pCodPat3 As String)
   strSql1 = "insert into movilturno (codmovil,codpatrullero,chofer,police,fechainic,kminicial,codturno, codpatrullero2, codpatrullero3) values " _
      & "('" & pCodMovil & "','" & pPatrullero & "','" & pChofer & "','" & pPolice & "','" & Format(pFechaIni, "yyyy-mm-dd hh:mm:ss") & "'," _
      & "'" & pKmIni & "','" & pCodTurno & "','" & pCodPat2 & "','" & pCodPat3 & "')"
   sAsignCmd
   cmdTemp.Execute
End Function

Public Function xInsCombustible(ByVal pFecha As String, ByVal pCodMovil As String, ByVal pKilometraje As String, ByVal pLitros As String)
   strSql1 = "insert into combustible (fecha,codmovil,kilometros,litros) values " _
      & " ('" & Format(pFecha, "yyyy-mm-dd hh:mm:ss") & "','" & pCodMovil & "','" & pKilometraje & "'," & pLitros & ")"
      sAsignCmd
      cmdTemp.Execute
End Function

Public Function xInsCodigos(ByVal pCodAlfa As String)
   strSql1 = "insert into codigos (CodAlfa, Estado) VALUES ('" & pCodAlfa & "','1')"
   sAsignCmd
   cmdTemp.Execute
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

Public Function xInsertUserRegNov(ByVal pCodUser As String, ByVal pPCname As String)
   strSql1 = "select * from actualizar where pc='" & Mid(pPCname, 1, Len(pPCname) - 1) & "'"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   If Not Rs1.EOF Then
      strSql1 = "update actualizar set usuario='" & pCodUser & "' WHERE PC='" & Mid(pPCname, 1, Len(pPCname) - 1) & "'"
   Else
      strSql1 = "insert into actualizar VALUES ('" & Mid(pPCname, 1, Len(pPCname) - 1) & "','" & pCodUser & "','0')"
   End If
   sAsignCmd
   On Error Resume Next
   cmdTemp.Execute
End Function

Public Function xDelCodigos()
   strSql1 = "delete from codigos"
   sAsignCmd
   cmdTemp.Execute
End Function

Public Function xInEnlace(ByVal pCodigo As String, ByVal pNroFicha As String, ByVal pFecha As String, ByVal pCodOper As String)
   strSql1 = "insert into enlace (codigo, fecha, nroficha, codoperador, fechaproc) values " _
      & "('" & pCodigo & "','" & Format(pFecha, "yyyy-mm-dd") & "','" & pNroFicha & "', '" & pCodOper & "',current_timestamp)"
   sAsignCmd
   On Error Resume Next
   cmdTemp.Execute
   If Err.Description <> "" Then
      MsgBox "Existió un error al guardar el dato:" & Chr(13) & Err.Description, vbCritical, sMessage
   End If
End Function

Public Function xUpEnlace(ByVal pCodigo As String, ByVal pFecha As String, ByVal pNroFicha As String)
   strSql1 = "update enlace set nroficha='" & pNroFicha & "' where " _
      & "codigo='" & pCodigo & "' and fecha='" & Format(pFecha, "yyyy-mm-dd") & "'"
   sAsignCmd
   On Error Resume Next
   cmdTemp.Execute
End Function

Public Function xDepuraNov()
   strSql1 = "delete from novedades2 where " _
      & "fecha < DATE_ADD(current_timestamp,INTERVAL -10 DAY);"
   sAsignCmd
   On Error Resume Next
   cmdTemp.Execute
End Function

Public Function xInTareaDatos(ByVal pCodAlfa As String, ByVal pMov As String, ByVal pC1 As Integer, ByVal pC2 As Integer, _
   ByVal pC3 As Integer, ByVal pC4 As Integer, ByVal pC5 As Integer, ByVal pC6 As Integer, ByVal pC7 As Integer, _
   ByVal pC8 As Integer, ByVal pC9 As Integer, ByVal pBI As Integer, ByVal pBE As Integer, ByVal pCant As String)
   strSql1 = "insert into tareas_datos values (current_timestamp, '" & pCodAlfa & "', '" & pMov & "', " & pC1 & ", " & pC2 & ", " & pC3 _
   & ", " & pC4 & ", " & pC5 & ", " & pC6 & ", " & pC7 & ", " & pC8 & ", " & pC9 & ", " & pBI & ", " & pBE & ", '" & pCant & "')"
   sAsignCmd
   On Error Resume Next
   cmdTemp.Execute
End Function

Public Sub dKmIniFin(ByVal pCodSent As String, ByRef pIni As Double, ByRef pFin As Double)
   pIni = 0
   pFin = 0
   strSql1 = "select * from sentidos where codigo=" & pCodSent
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   If Not Rs1.EOF Then
      pIni = Rs1!km_i
      pFin = Rs1!km_F
   End If
End Sub

Public Function xInLogWaze(ByVal pCodAlfa As String, ByVal pCodTipo As String, ByVal pDescr As String)
   strSql1 = "insert into d_logwaze values (0, current_timestamp, '" & Trim(Right(MDI.mUser, 25)) & "', " & pCodTipo & ", '" & pCodAlfa & "', '" & pDescr & "')"
   sAsignCmd
   On Error Resume Next
   cmdTemp.Execute
   If Err.Description <> "" Then
      MsgBox Err.Description
   End If
   
End Function


Public Function getRamalWazeXml(ByVal pSentido As Integer) As String
'Dim mObj As New clRNov
Dim mRamal As Integer
Dim sRamal As String
Dim sSentido As String
   mRamal = sTablaDescr("sentidos", " codigo=" & pSentido, 2)
   sSentido = sTablaDescr("sentidos", " codigo=" & pSentido, 1)
   sRamal = sTablaDescr("ramales", " codigo=" & mRamal, 1)
   
   getRamalWazeXml = sRamal & " - " & Mid(sSentido, 4, Len(sSentido) - 3)
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
'=======               Llamadas a Store Procedures MySQL: Schema regnov             ========
'===========================================================================================

Public Function waze_getAccidentesLiberados() As ADODB.Recordset
   Set waze_getAccidentesLiberados = oStoreProcedure("waze_getAccidentesLiberados")
End Function

Public Function waze_ultima_novedad(ByVal CodAlfa As String) As ADODB.Recordset
   Dim paramCodAlfa As parametroSql
   Dim aParamSql(0) As parametroSql

   paramCodAlfa.Nombre = "vcodalfa"
   paramCodAlfa.Tipo = adVarChar
   paramCodAlfa.Direccion = adParamInput
   paramCodAlfa.Tamanio = 22
   paramCodAlfa.Valor = CodAlfa
   aParamSql(0) = paramCodAlfa
  
   Set waze_ultima_novedad = oStoreProcedureWithParam("waze_ultima_novedad", aParamSql)
End Function

Public Sub waze_cerrar_accidente(ByVal CodAlfa As String)
   Dim paramCodAlfa As parametroSql
   Dim aParamSql(0) As parametroSql

   paramCodAlfa.Nombre = "vcodalfa"
   paramCodAlfa.Tipo = adVarChar
   paramCodAlfa.Direccion = adParamInput
   paramCodAlfa.Tamanio = 22
   paramCodAlfa.Valor = CodAlfa
   
   aParamSql(0) = paramCodAlfa
   xStoreProcedureWithParam "waze_cerrar_accidente", aParamSql
   
End Sub

Public Sub waze_enviar_accidente(ByVal CodAlfa As String, ByVal descripcion As String, ByVal codsubtipo As Integer)
   Dim paramCodAlfa As parametroSql
   Dim paramDescripcion As parametroSql
   Dim paramCodSubtipo As parametroSql
   Dim aParamSql(2) As parametroSql
      
   With paramCodAlfa
      .Nombre = "vcodalfa"
      .Tipo = adVarChar
      .Direccion = adParamInput
      .Tamanio = 22
      .Valor = CodAlfa
   End With
   aParamSql(0) = paramCodAlfa

   With paramDescripcion
      .Nombre = "vdescripcion"
      .Tipo = adVarChar
      .Direccion = adParamInput
      .Tamanio = 255
      .Valor = descripcion
   End With
   aParamSql(1) = paramDescripcion

   With paramCodSubtipo
      .Nombre = "vcodsubtipo"
      .Tipo = adInteger
      .Direccion = adParamInput
      .Tamanio = 4
      .Valor = codsubtipo
   End With
   aParamSql(2) = paramCodSubtipo

   xStoreProcedureWithParam "waze_enviar_accidente", aParamSql
End Sub

Public Sub waze_asignar_movil(ByVal CodAlfa As String)
   Dim paramCodAlfa As parametroSql
   Dim aParamSql(0) As parametroSql
   
   With paramCodAlfa
      .Nombre = "vcodalfa"
      .Tipo = adVarChar
      .Direccion = adParamInput
      .Tamanio = 22
      .Valor = CodAlfa
   End With
   aParamSql(0) = paramCodAlfa
   
   xStoreProcedureWithParam "waze_asignar_movil", aParamSql
End Sub

'Public Sub waze_liberar_movil(ByVal codalfa As String)
'   Dim paramCodAlfa As parametroSql
'   Dim aParamSql(0) As parametroSql
'
'   With paramCodAlfa
'      .Nombre = "vcodalfa"
'      .Tipo = adVarChar
'      .Direccion = adParamInput
'      .Tamanio = 22
'      .Valor = codalfa
'   End With
'   aParamSql(0) = paramCodAlfa
'
'   xStoreProcedureWithParam "waze_liberar_movil", aParamSql
'End Sub


Public Function waze_liberar_movil(ByVal CodAlfa As String) As ADODB.Recordset

   Dim paramCodAlfa As parametroSql
   Dim aParamSql(0) As parametroSql

   With paramCodAlfa
      .Nombre = "vcodalfa"
      .Tipo = adVarChar
      .Direccion = adParamInput
      .Tamanio = 22
      .Valor = CodAlfa
   End With
   aParamSql(0) = paramCodAlfa
   Set waze_liberar_movil = oStoreProcedureWithParam("waze_liberar_movil", aParamSql)
End Function
