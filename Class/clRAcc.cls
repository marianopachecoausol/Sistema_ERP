VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clRAcc"
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
Dim mAccidentes As String

Private Sub Class_Initialize()
   Set mObj = New clConnSistemas
   DBConn.ConnectionTimeout = 500
   'Set DBConn = mObj.oAbreConexion("4", mIPServer, "FichaAccid", "ssvv", "ssvv")
   Set DBConn = mObj.oAbreConexion("4", mIPServer, "FichaAccid", "ssvv", "ssvv") 'mp20190527
   'Set DBConn = mObj.oAbreConexion("4", "10.10.20.11", "FichaAccid", "root", "123456")
   mAccidentes = " and codtipoficha='01' "
End Sub

Private Sub sAsignCmd()
   cmdTemp.CommandText = strSql1
   cmdTemp.CommandType = 1
   cmdTemp.CommandTimeout = 36000
   Set cmdTemp.ActiveConnection = DBConn
End Sub

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

Public Function oTabla(ByVal pNameTabla As String, ByVal pOrden As String) As ADODB.Recordset
   strSql1 = "SELECT * FROM " & pNameTabla & " " & pOrden
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oTabla = Rs1
End Function

Public Function oTablaNotNull(ByVal pNameTabla As String, ByVal pOrden As String) As ADODB.Recordset
   strSql1 = "SELECT * FROM " & pNameTabla & " WHERE fechabaja is null " & pOrden
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

Public Function sTablaDescr(ByVal pNameTabla As String, ByVal pCodigo As String, ByVal pCampo As Integer) As String
   strSql1 = "SELECT * FROM " & pNameTabla & " WHERE " & pCodigo
   sAsignCmd
   sTablaDescr = ""
   Set Rs1 = cmdTemp.Execute
   If Not Rs1.EOF Then
      sTablaDescr = Rs1.Fields(pCampo)
   End If
End Function

Public Function sUltNroOrden() As String
   sUltNroOrden = "00000"
   strSql1 = "select max(nroorden) as ult from Ficha"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   If Not Rs1.EOF And Rs1!ult <> "Nulo" Then
      sUltNroOrden = Format(Val(Rs1!ult) + 1, "00000")
   Else
      sUltNroOrden = "00001"
   End If
End Function

Public Function bExistFicha(ByVal pNroOrden As String) As Boolean
   strSql1 = "select fecha from Ficha where nroorden='" & pNroOrden & "'"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   bExistFicha = Not Rs1.EOF
End Function

Public Function bExistFichaCodAlfa(ByVal pFecha As String, ByVal pCodAlfa As String) As Boolean
   strSql1 = "select fecha from Ficha where fecha='" & Format(pFecha, "yyyy-mm-dd") & "' and codalfa='" & pCodAlfa & "'"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   bExistFichaCodAlfa = Not Rs1.EOF
End Function

Public Function oAccidTroncal(ByVal pDesde, ByVal pHasta) As ADODB.Recordset
   strSql1 = "SELECT COUNT(*) AS Total FROM Ficha WHERE ((Progresiva >= 25.75 AND Progresiva < 26.25) or (Progresiva >= 57.50 AND Progresiva < 58)) AND LugarAccid <> '09' AND (Fecha >= '" & Format(pDesde, "yyyy/mm/dd") & " 00:00:00' AND Fecha <= '" & Format(pHasta, "yyyy/mm/dd") & " 23:59:59')" & mAccidentes
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oAccidTroncal = Rs1
End Function

Public Function oBuscar(ByVal pTabla As String, ByVal pOrden As String) As ADODB.Recordset
   If Len(pTabla) > 6 Then
       strSql1 = "SELECT distinct A.nroorden, fecha FROM " & pTabla & " " & pOrden
   Else
      strSql1 = "SELECT distinct nroorden, fecha FROM " & pTabla & " " & pOrden
   End If
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oBuscar = Rs1
End Function

Public Function oAccidTraza(ByVal pDesde, ByVal pHasta, ByVal pBoolAccid As Boolean) As ADODB.Recordset
   strSql1 = "SELECT COUNT(*) AS Total FROM Ficha WHERE ((Progresiva < 25.75 OR Progresiva >= 26.25) AND (Progresiva < 57.50 OR Progresiva >= 58)) AND LugarAccid <> '09' AND (Fecha >= '" & Format(pDesde, "yyyy/mm/dd") & " 00:00:00' AND Fecha <= '" & Format(pHasta, "yyyy/mm/dd") & " 23:59:59') "
   If pBoolAccid Then
      strSql1 = strSql1 & " and codtipoficha='01' "
   Else
      strSql1 = strSql1 & " and codtipoficha='02' "
   End If
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oAccidTraza = Rs1
End Function

Public Function oAccidColect(ByVal pDesde, ByVal pHasta) As ADODB.Recordset
   strSql1 = "SELECT COUNT(*) AS Total FROM Ficha WHERE LugarAccid = '09' AND (Fecha >= '" & Format(pDesde, "yyyy/mm/dd") & " 00:00:00' AND Fecha <= '" & Format(pHasta, "yyyy/mm/dd") & " 23:59:59')" & mAccidentes
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oAccidColect = Rs1
End Function

Public Function oMarcasVehic(ByVal pCodTipo As String) As ADODB.Recordset
   strSql1 = "select * from Marca where codtipovehic='" & pCodTipo & "' " 'and fechabaja is null"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oMarcasVehic = Rs1
End Function

Public Function oMarcasVehicDescr(ByVal pCodTipo As String, ByVal pCodMarca As String) As ADODB.Recordset
   strSql1 = "select * from Marca where codtipovehic='" & pCodTipo & "' and codmarca='" & pCodMarca & "' " 'and fechabaja is null"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oMarcasVehicDescr = Rs1
End Function

Public Function oInterTerceros(ByVal pNroOrden As String) As ADODB.Recordset
   strSql1 = "SELECT B.descripcion, A.Movil, A.personal, A.dependencia  FROM interterceros A, movilext B  Where A.codtipo = B.Codigo and nroorden='" & pNroOrden & "' order by 1"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oInterTerceros = Rs1
End Function

Public Function oFichasViejas(ByVal pNroOrden As String) As Boolean
   strSql1 = "select * from Ficha where nroorden='" & pNroOrden & "' and codalfa is not null and codconfig is not null and codvisibilidad is not null and codinconv is not null and codpavim is not null"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   oFichasViejas = Rs1.EOF
End Function

Public Function oNroOrdenFechas(ByVal pFecha1 As String, ByVal pFecha2 As String) As ADODB.Recordset
   strSql1 = "select nroorden, fecha from Ficha where fecha between '" & Format(pFecha1, "yyyy-mm-dd") & "' and '" & Format(pFecha2, "yyyy-mm-dd") & "' order by 1,2"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oNroOrdenFechas = Rs1
End Function

Public Function oFichaAccidentados(ByVal pFecha1 As String, ByVal pFecha2 As String) As ADODB.Recordset
   strSql1 = "select distinctrow B.nroorden,letra, A.nrovictima,  B.fecha, B.Hora, B.Progresiva, A.herido, A.fallecio, A.codestado, B.lugaraccid " _
      & " FROM VictimasInvolucr A, Ficha B where  B.NroOrden = A.NroOrden AND B.Fecha between '" & Format(pFecha1, "yyyy-mm-dd") & "' AND '" & Format(pFecha2, "yyyy-mm-dd") & "' " & mAccidentes & " order by 1,2"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oFichaAccidentados = Rs1
End Function

Public Function oTablaNroOrden(ByVal pNameTabla As String, ByVal pNroOrden As String, ByVal pOrden As String) As ADODB.Recordset
   strSql1 = "SELECT * FROM " & pNameTabla & " where NroOrden='" & pNroOrden & "' " & pOrden
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oTablaNroOrden = Rs1
End Function

Public Function oFichasFechaProgr(ByVal pFecha1 As String, ByVal pFecha2 As String, ByVal pProg1 As String, ByVal pProg2 As String, ByVal pLugar As String) As ADODB.Recordset
   Dim mWhere As String
      
   If pProg1 <> "" Or pProg2 <> "" Then
      mWhere = " and progresiva between " & pProg1 & " and " & pProg2 & ""
   End If
   If pLugar <> "" Then
      mWhere = mWhere & " and " & pLugar
   End If
   strSql1 = "select * from Ficha where fecha between '" & Format(pFecha1, "yyyy-mm-dd") & "' and '" & Format(pFecha2, "yyyy-mm-dd") & "' " _
        & mWhere & mAccidentes & " order by 1"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oFichasFechaProgr = Rs1
End Function

Public Function oInfPersonal(ByVal pFecha1 As String, ByVal pFecha2 As String) As ADODB.Recordset
   If pFecha1 > "01/03/2008" Then
      strSql1 = "select distinct A.NroOrden, B.CodMovil, A.fecha, A.hora, A.progresiva, A.accidotro, A.Acciconotro,A.codcoliscontra1, clima1 " _
        & " from Ficha A, intergco B Where A.NroOrden = B.NroOrden  and A.fecha between '" & Format(pFecha1, "yyyy-mm-dd") & "' " _
        & " and '" & Format(pFecha2, "yyyy-mm-dd") & "' group by nroorden "
   Else
      strSql1 = "select distinct A.NroOrden, A.MovilNro, A.fecha, A.hora, A.progresiva, A.accidotro, A.Acciconotro,A.codcoliscontra1, clima1 " _
        & " from Ficha A Where A.fecha between '" & Format(pFecha1, "yyyy-mm-dd") & "' " _
        & " and '" & Format(pFecha2, "yyyy-mm-dd") & "' group by nroorden "
   End If
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oInfPersonal = Rs1
End Function

Public Function oFichasLegales(ByVal pFecha As String)
   strSql1 = "SELECT nroorden, fecha FROM Ficha where fecha >='" & pFecha & "' order by 1"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oFichasLegales = Rs1
End Function

Public Function iTotalHeridosNroOrden(ByVal pNroOrden As String) As Integer
   strSql1 = "select count(nrovictima) as Total From VictimasInvolucr where nroorden='" & pNroOrden & "' and (herido in ('1','2') or codestado in ('02','03'))"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   iTotalHeridosNroOrden = 0
   If Not Rs1.EOF Then
      iTotalHeridosNroOrden = Rs1!Total
   End If
End Function

Public Function iTotalHeridosNroOrdenMIGRACION(ByVal pNroOrden As String) As Integer
   strSql1 = "select count(nrovictima) as Total From VictimasInvolucrMIGRACION where nroorden='" & pNroOrden & "' and codestado in ('02','03')"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   iTotalHeridosNroOrdenMIGRACION = 0
   If Not Rs1.EOF Then
      iTotalHeridosNroOrdenMIGRACION = Rs1!Total
   End If
End Function


Public Function iTotalMuertosNroOrden(ByVal pNroOrden As String) As Integer
   strSql1 = "select count(nrovictima) as Total From VictimasInvolucr where nroorden='" & pNroOrden & "' and (fallecio in('1','2') or codestado in ('04','05'))"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   iTotalMuertosNroOrden = 0
   If Not Rs1.EOF Then
      iTotalMuertosNroOrden = Rs1!Total
   End If
End Function

Public Function iTotalMuertosNroOrdenMIGRACION(ByVal pNroOrden As String) As Integer
   strSql1 = "select count(nrovictima) as Total From VictimasInvolucrMIGRACION where nroorden='" & pNroOrden & "' and  codestado in ('04','05')"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   iTotalMuertosNroOrdenMIGRACION = 0
   If Not Rs1.EOF Then
      iTotalMuertosNroOrdenMIGRACION = Rs1!Total
   End If
End Function

Public Function iTotalVehiculosNroOrden(ByVal pNroOrden As String) As Integer
   strSql1 = "select count(letra) as Total From VehiculosInvolucr where nroorden='" & pNroOrden & "'"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   iTotalVehiculosNroOrden = 0
   If Not Rs1.EOF Then
      iTotalVehiculosNroOrden = Rs1!Total
   End If
End Function

Public Function iAccidFechaHora(ByVal pFecha1 As String, ByVal pFecha2 As String, _
                                ByVal pHora1 As String, ByVal pHora2 As String) As Integer
   strSql1 = "select count(nroorden) as total from Ficha where " _
      & " fecha between '" & Format(pFecha1, "yyyy-mm-dd") & "' AND '" & Format(pFecha2, "yyyy-mm-dd") & "' " _
      & " and hora between '" & pHora1 & "' and '" & pHora2 & "' " & mAccidentes
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   iAccidFechaHora = 0
   If Not Rs1.EOF Then
      iAccidFechaHora = Rs1!Total
   End If
End Function

Public Function iAccidentesProgr(ByVal pFecha1 As String, ByVal pFecha2 As String, _
                                 ByVal pProgr1 As Double, ByVal pProgr2 As Double, ByVal pTraza As Boolean) As Integer
   strSql1 = "select count(*) as total from Ficha where Progresiva >= " & pProgr1 & " and Progresiva < " & pProgr2 & " " _
      & " and Fecha between '" & Format(pFecha1, "yyyy-mm-dd") & "' AND '" & Format(pFecha2, "yyyy-mm-dd") & "' "
      
   If pTraza Then
      strSql1 = strSql1 & " and lugaraccid <> '09'" & mAccidentes
   Else
      strSql1 = strSql1 & " and lugaraccid = '09'" & mAccidentes
   End If
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   iAccidentesProgr = 0
   If Not Rs1.EOF Then
      iAccidentesProgr = Rs1!Total
   End If
End Function

Public Function iAccidVictProgr(ByVal pFecha1 As String, ByVal pFecha2 As String, _
                                 ByVal pProgr1 As Double, ByVal pProgr2 As Double, ByVal pTraza As Boolean) As Integer
         strSql1 = "select count(distinct A.nroorden) as total from VictimasInvolucr A, Ficha B where B.Progresiva >= " & pProgr1 & " and B.Progresiva < " & pProgr2 & " " _
      & " and B.Fecha between '" & Format(pFecha1, "yyyy-mm-dd") & "' AND '" & Format(pFecha2, "yyyy-mm-dd") & "' " _
      & " and A.nroorden = B.nroorden and (codestado in ('02','03','04','05') or herido in ('1','2') or fallecio in ('1','2','3')) "
   If pTraza Then
      strSql1 = strSql1 & " and lugaraccid <> '09'" & mAccidentes
   Else
      strSql1 = strSql1 & " and lugaraccid = '09'" & mAccidentes
   End If
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   iAccidVictProgr = 0
   If Not Rs1.EOF Then
      iAccidVictProgr = Rs1!Total
   End If
End Function

Public Function iMuertosHeridosProgr(ByVal pFecha1 As String, ByVal pFecha2 As String, _
                                 ByVal pProgr1 As Double, ByVal pProgr2 As Double, ByVal pTraza As Boolean, ByVal pHerido As Boolean) As Integer
   strSql1 = "select count(A.nrovictima) as Total From VictimasInvolucr A, Ficha B Where A.nroorden = B.nroorden " _
      & " and B.Progresiva >= " & pProgr1 & " AND Progresiva <" & pProgr2 & " " _
      & " and B.Fecha between '" & Format(pFecha1, "yyyy-mm-dd") & "' AND '" & Format(pFecha2, "yyyy-mm-dd") & "' "
   If pTraza Then
      strSql1 = strSql1 & " and lugaraccid <> '09'"
   Else
      strSql1 = strSql1 & " and lugaraccid = '09'"
   End If
   If pHerido Then
      strSql1 = strSql1 & " and (A.herido <> '' or A.codestado in ('02','03'))" & mAccidentes
   Else
      strSql1 = strSql1 & " and (A.fallecio <> '' or A.codestado in ('04','05'))" & mAccidentes
   End If
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   iMuertosHeridosProgr = 0
   If Not Rs1.EOF Then
      iMuertosHeridosProgr = Rs1!Total
   End If
End Function

Public Function oMuertosHeridosProgr(ByVal pFecha1 As String, ByVal pFecha2 As String, _
                                 ByVal pProgr1 As Double, ByVal pProgr2 As Double, ByVal pTraza As Boolean, ByVal pHerido As Boolean) As ADODB.Recordset
   strSql1 = "select count(A.nrovictima) as Total From VictimasInvolucr A, Ficha B Where A.nroorden = B.nroorden " _
      & " and B.Progresiva >= " & pProgr1 & " AND Progresiva <" & pProgr2 & " " _
      & " and B.Fecha between '" & Format(pFecha1, "yyyy-mm-dd") & "' AND '" & Format(pFecha2, "yyyy-mm-dd") & "' "
   If pTraza Then
      strSql1 = strSql1 & " and lugaraccid <> '09'"
   Else
      strSql1 = strSql1 & " and lugaraccid = '09'"
   End If
   If pHerido Then
      strSql1 = strSql1 & " and (A.herido <> '' or A.codestado in ('02','03'))" & mAccidentes
   Else
      strSql1 = strSql1 & " and (A.fallecio <> '' or A.codestado in ('04','05'))" & mAccidentes
   End If
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oMuertosHeridosProgr = Rs1
End Function

Public Function iCantVehicFechas(ByVal pFecha1 As String, ByVal pFecha2 As String, ByVal pCodTipoVehic As String) As Integer
   strSql1 = "select count(A.nroorden) as total from VehiculosInvolucr A, Ficha B where" _
      & " B.Fecha between '" & Format(pFecha1, "yyyy-mm-dd") & "' AND '" & Format(pFecha2, "yyyy-mm-dd") & "' " _
      & " and A.nroorden = B.nroorden "
   If pCodTipoVehic <> "" Then
      strSql1 = strSql1 & " and A.codtipovehic in (" & pCodTipoVehic & ")" & mAccidentes
   End If
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   iCantVehicFechas = 0
   If Not Rs1.EOF Then
      iCantVehicFechas = Rs1!Total
   End If
End Function


Public Function iCantVehicFechasConParamTraza(ByVal pFecha1 As String, ByVal pFecha2 As String, ByVal pCodTipoVehic As String, ByVal pTraza As Boolean) As Integer
   strSql1 = "select count(A.nroorden) as total from VehiculosInvolucr A, Ficha B where" _
      & " B.Fecha between '" & Format(pFecha1, "yyyy-mm-dd") & "' AND '" & Format(pFecha2, "yyyy-mm-dd") & "' " _
      & " and A.nroorden = B.nroorden "
   If pCodTipoVehic <> "" Then
      strSql1 = strSql1 & " and A.codtipovehic in (" & pCodTipoVehic & ")" & mAccidentes
   End If
   
   If pTraza Then
      strSql1 = strSql1 & " and lugaraccid <> '09'"
   Else
      strSql1 = strSql1 & " and lugaraccid = '09'"
   End If
      
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   iCantVehicFechasConParamTraza = 0
   If Not Rs1.EOF Then
      iCantVehicFechasConParamTraza = Rs1!Total
   End If
End Function


Public Function iCantVehicNroFicha(ByVal pNroOrden As String, ByVal pCodTipoVehic As String) As Integer
   strSql1 = "select count(A.nroorden) as total from VehiculosInvolucr A, Ficha B where " _
      & " B.Nroorden='" & pNroOrden & "' " _
      & " and A.nroorden = B.nroorden "
   If pCodTipoVehic <> "" Then
      strSql1 = strSql1 & " and A.codtipovehic in (" & pCodTipoVehic & ")" & mAccidentes
   End If
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   iCantVehicNroFicha = 0
   If Not Rs1.EOF Then
      iCantVehicNroFicha = Rs1!Total
   End If
End Function

Public Function iCantVehicNroFicha2(ByVal pNroOrden As String, ByVal pCodTipoVehic As String) As Integer
   strSql1 = "select count(A.nroorden) as total from VehiculosInvolucr A, Ficha B where " _
      & " B.Nroorden='" & pNroOrden & "' " _
      & " and A.nroorden = B.nroorden "
   If pCodTipoVehic <> "" Then
      strSql1 = strSql1 & " and A.codtipovehic in (" & pCodTipoVehic & ")"
   End If
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   iCantVehicNroFicha2 = 0
   If Not Rs1.EOF Then
      iCantVehicNroFicha2 = Rs1!Total
   End If
End Function


Public Function iCantCausaFechas(ByVal pFecha1 As String, ByVal pFecha2 As String, ByVal pCodWhere As String) As Integer
   strSql1 = "select count(nroorden) as total from Ficha  where Fecha between '" & Format(pFecha1, "yyyy-mm-dd") & "' AND '" & Format(pFecha2, "yyyy-mm-dd") & "' "
   If pCodWhere <> "" Then
      strSql1 = strSql1 & " and " & pCodWhere & mAccidentes
   End If
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   iCantCausaFechas = 0
   If Not Rs1.EOF Then
      iCantCausaFechas = Rs1!Total
   End If
End Function

Public Function iTotalMuertosCodigo(ByVal pFecha1 As String, ByVal pFecha2 As String, ByVal pCodWhere As String) As Integer
   strSql1 = "select count(A.nroorden) as Total From VictimasInvolucr A, Ficha B where A.nroorden=B.NroOrden " _
      & " and B.Fecha between '" & Format(pFecha1, "yyyy-mm-dd") & "' AND '" & Format(pFecha2, "yyyy-mm-dd") & "' " & mAccidentes
   If pCodWhere <> "" Then
      strSql1 = strSql1 & " and " & pCodWhere
   End If
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   iTotalMuertosCodigo = 0
   If Not Rs1.EOF Then
      iTotalMuertosCodigo = Rs1!Total
   End If
End Function

Public Function iTotalMuertos3(ByVal pFecha1 As String, ByVal pFecha2 As String, ByVal pWhere As String) As Integer
   strSql1 = "SELECT COUNT(A.nroorden) AS Fall FROM VictimasInvolucr A, VehiculosInvolucr B, Ficha C WHERE A.NroOrden = B.NroOrden AND C.NroOrden = B.NroOrden " _
      & " and C.Fecha between '" & Format(pFecha1, "yyyy-mm-dd") & "' AND '" & Format(pFecha2, "yyyy-mm-dd") & "' " & pWhere & mAccidentes
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   iTotalMuertos3 = 0
   If Not Rs1.EOF Then
      iTotalMuertos3 = Rs1!fall
   End If
End Function

Public Function oConsulta2004(ByVal pFecha1 As String, ByVal pFecha2 As String) As ADODB.Recordset
   strSql1 = "SELECT DISTINCTROW Ficha.Fecha,Ficha.Hora,Ficha.Progresiva,CONCAT_WS('-',ConOtroVehic.CodConOtro,ConOtroVehic.Detalle) AS mAccidconOtro," _
      & " CONCAT_WS('-',CausaConductor.CodCausaCond,CausaConductor.Descripcion) AS mCausaCond,CONCAT_WS('-',CausaVehic.CodCausaVehic,CausaVehic.Detalle) AS mCausaVehic ,CONCAT_WS('-',clima.CodClima,clima.Detalle) AS mClima, " _
      & " CONCAT_WS('-',SentidoTrans.CodSentidoTrans,SentidoTrans.Detalle) AS mSentTrans,CONCAT_WS('-',LugarAccid.CodLugarAccid,LugarAccid.Detalle) AS mLugarAccid,Space(5) AS mHora FROM ((((((Ficha LEFT JOIN ConOtroVehic ON ConOtroVehic.CodConOtro = Ficha.AcciConOtro) LEFT JOIN " _
      & " CausaConductor ON CausaConductor.CodCausaCond = Ficha.CodCausaCond1) LEFT JOIN CausaVehic ON CausaVehic.CodCausaVehic = Ficha.CausaVehic) LEFT JOIN clima ON clima.CodClima = Ficha.Clima1) LEFT JOIN  SentidoTrans ON SentidoTrans.CodSentidoTrans = Ficha.SentidoTrans) LEFT JOIN LugarAccid ON LugarAccid.CodLugarAccid = Ficha.LugarAccid) " _
      & " WHERE Ficha.Fecha between '" & Format(pFecha1, "yyyy-mm-dd") & "' AND '" & Format(pFecha2, "yyyy-mm-dd") & "' " & mAccidentes & " ORDER BY Ficha.Fecha"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oConsulta2004 = Rs1
End Function

Public Function oPeligroSector(ByVal pFecha1 As String, ByVal pFecha2 As String) As ADODB.Recordset
   strSql1 = "SELECT Ficha.NroOrden,CONCAT_WS('-',SentidoTrans.CodSentidoTrans,SentidoTrans.Detalle) AS mSentTrans,Ficha.Progresiva,Ficha.Fecha,Ficha.Hora,CONCAT_WS('-',LugarAccid.CodLugarAccid,LugarAccid.Detalle) AS mLugarAccid,CONCAT_WS('-',otros.CodOtros,otros.Detalle) As mAccidOtros,CONCAT_WS('-',ConOtroVehic.CodConOtro,ConOtroVehic.Detalle) AS mAccidConOtro,CONCAT_WS('-',ColisionContra.CodColision,ColisionContra.Descripcion) AS mColisContra,CONCAT_WS('-',CausaVehic.CodCausaVehic,CausaVehic.Detalle) AS mCausaVehic,CONCAT_WS('-',CausaConductor.CodCausaCond,CausaConductor.Descripcion) AS mCausaCond,CONCAT_WS('-',clima.CodClima,clima.Detalle) AS mClima,Space(5) AS mHora" _
      & " FROM ((((((((Ficha LEFT JOIN SentidoTrans ON SentidoTrans.CodSentidoTrans = Ficha.SentidoTrans) LEFT JOIN LugarAccid ON LugarAccid.CodLugarAccid = Ficha.LugarAccid) LEFT JOIN otros ON otros.CodOtros = Ficha.AccidOtro) LEFT JOIN ConOtroVehic ON ConOtroVehic.CodConOtro = Ficha.AcciConOtro) LEFT JOIN ColisionContra ON ColisionContra.CodColision = Ficha.CodColisContra1) LEFT JOIN CausaVehic ON CausaVehic.CodCausaVehic = Ficha.CausaVehic) LEFT JOIN CausaConductor ON CausaConductor.CodCausaCond = Ficha.CodCausaCond1) LEFT JOIN clima ON clima.CodClima = Ficha.Clima1) " _
      & " WHERE Ficha.Fecha between '" & Format(pFecha1, "yyyy-mm-dd") & "' AND '" & Format(pFecha2, "yyyy-mm-dd") & "' " & mAccidentes & " ORDER BY Ficha.Fecha"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oPeligroSector = Rs1
End Function

Public Function oEvolucion() As ADODB.Recordset
   strSql1 = "SELECT DISTINCTROW Patrullero.Nombre,Count(Ficha.NroOrden) AS TotalOrdens,Ficha.Fecha FROM Ficha,Patrullero WHERE Patrullero.CodPatrullero = Ficha.CodPatrullero GROUP BY Patrullero.Nombre, Ficha.Fecha"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oEvolucion = Rs1
End Function

Public Function oTotDiaSemana(ByVal pFecha1 As String, ByVal pFecha2 As String, ByVal mWhere As String) As ADODB.Recordset
   strSql1 = "SELECT Count(*) AS Fall, Weekday(Ficha.Fecha) AS Dias FROM VictimasInvolucr, Ficha Where Ficha.Fecha between '" & Format(pFecha1, "yyyy-mm-dd") & "' AND '" & Format(pFecha2, "yyyy-mm-dd") & "' " & mWhere & mAccidentes & " GROUP BY Weekday(Ficha.Fecha) "
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oTotDiaSemana = Rs1
End Function

Public Function oTotFichaDiaSemana(ByVal pFecha1 As String, ByVal pFecha2 As String, ByVal mWhere As String) As ADODB.Recordset
   strSql1 = "SELECT Count(*) AS Fall, Weekday(Ficha.Fecha) AS Dias FROM Ficha Where Ficha.Fecha between '" & Format(pFecha1, "yyyy-mm-dd") & "' AND '" & Format(pFecha2, "yyyy-mm-dd") & "' " & mWhere & mAccidentes & " GROUP BY Weekday(Ficha.Fecha) "
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oTotFichaDiaSemana = Rs1
End Function

Public Function iTotAccMesHora(ByVal pFecha1 As String, ByVal pFecha2 As String, _
                        ByVal pHora1 As String, ByVal pHora2 As String, ByVal mWhere As String) As Integer
   iTotAccMesHora = 0
   strSql1 = "SELECT COUNT(nroorden) AS Cont FROM Ficha WHERE Ficha.Fecha between '" & Format(pFecha1, "yyyy-mm-dd") & "' AND '" & Format(pFecha2, "yyyy-mm-dd") & "' " _
           & " AND Hora between '" & pHora1 & "' and '" & pHora2 & "' " & mWhere & mAccidentes
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   If Not Rs1.EOF Then
      iTotAccMesHora = Rs1!Cont
   End If
End Function
 
Public Function iTotAccFicVicMesHora(ByVal pFecha1 As String, ByVal pFecha2 As String, _
                        ByVal pHora1 As String, ByVal pHora2 As String, ByVal mWhere As String) As Integer
   iTotAccFicVicMesHora = 0
   strSql1 = "SELECT COUNT(*) AS Cont FROM Ficha, VictimasInvolucr WHERE Ficha.NroOrden = VictimasInvolucr.NroOrden " _
           & " AND Ficha.Fecha between '" & Format(pFecha1, "yyyy-mm-dd") & "' AND '" & Format(pFecha2, "yyyy-mm-dd") & "' " _
           & " AND Hora between '" & pHora1 & "' and '" & pHora2 & "' " & mWhere & mAccidentes
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   If Not Rs1.EOF Then
      iTotAccFicVicMesHora = Rs1!Cont
   End If
End Function

Public Function iTotAccDia(ByVal pFecha1 As String, ByVal pFecha2 As String, ByVal pDia As Integer, ByVal mWhere As String) As Integer
   iTotAccDia = 0
   strSql1 = "SELECT DISTINCTROW COUNT(*) AS Dia FROM Ficha WHERE (Ficha.Fecha >= '" & Format(pFecha1, "yyyy-mm-dd") & "' AND Ficha.Fecha <= '" & Format(pFecha2, "yyyy-mm-dd") & "') AND WEEKDAY(Ficha.Fecha) = " & pDia & mWhere & mAccidentes
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   If Not Rs1.EOF Then
      iTotAccDia = Rs1!Dia
   End If
End Function

Public Function iTotAccDiaVict(ByVal pFecha1 As String, ByVal pFecha2 As String, ByVal pDia As Integer, ByVal mWhere As String) As Integer
   iTotAccDiaVict = 0
   strSql1 = "SELECT DISTINCTROW COUNT(Ficha.nroorden) AS Dia FROM Ficha, VictimasInvolucr WHERE (Ficha.Fecha >= '" & Format(pFecha1, "yyyy-mm-dd") & "' AND Ficha.Fecha <= '" & Format(pFecha2, "yyyy-mm-dd") & "') AND WEEKDAY(Ficha.Fecha) = " & pDia & mWhere & mAccidentes
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   If Not Rs1.EOF Then
      iTotAccDiaVict = Rs1!Dia
   End If
End Function

Public Function oOtrosTrasl(ByVal pFecha1 As String, ByVal pFecha2 As String) As ADODB.Recordset
   strSql1 = "SELECT DISTINCTROW VictimasInvolucr.NroOrden,Ficha.Fecha,Ficha.Progresiva,Ficha.Hora,MedioTrasl.Descripcion,Space(5) AS mHora FROM " _
              & "(Ficha INNER JOIN VictimasInvolucr ON Ficha.NroOrden = VictimasInvolucr.NroOrden) LEFT JOIN MedioTrasl ON MedioTrasl.CodMedioTrasl = VictimasInvolucr.CodMedioTrasl WHERE ((Ficha.Fecha >= '" & Format(pFecha1, "yyyy-mm-dd") & "' AND Ficha.Fecha <= '" & Format(pFecha2, "yyyy-mm-dd") & "') AND (VictimasInvolucr.CodMedioTrasl = MedioTrasl.CodMedioTrasl) and ficha.codtipoficha='01' ) ORDER BY VictimasInvolucr.NroOrden"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oOtrosTrasl = Rs1
End Function

Public Function oPuntosNegros(ByVal pFecha1 As String, ByVal pFecha2 As String) As ADODB.Recordset
   strSql1 = "SELECT DISTINCTROW Ficha.NroOrden,CONCAT_WS('-',SentidoTrans.CodSentidoTrans,SentidoTrans.Detalle) AS mSentTrans,Ficha.Progresiva,Ficha.Fecha,Ficha.Hora,CONCAT_WS('-',LugarAccid.CodLugarAccid,LugarAccid.Detalle) AS mLugarAccid,CONCAT_WS('-',otros.CodOtros,otros.Detalle) As mAccidOtros,CONCAT_WS('-',ConOtroVehic.CodConOtro,ConOtroVehic.Detalle) AS mAccidconOtro,CONCAT_WS('-',ColisionContra.CodColision,ColisionContra.Descripcion) AS mColisContra,CONCAT_WS('-',CausaVehic.CodCausaVehic,CausaVehic.Detalle) AS mCausaVehic,CONCAT_WS('-',CausaConductor.CodCausaCond,CausaConductor.Descripcion) AS mCausaCond,CONCAT_WS('-',clima.CodClima,clima.Detalle) AS mClima, Space(5) AS mHora  " _
         & "FROM ((((((((Ficha LEFT JOIN otros ON otros.COdOtros = Ficha.AccidOtro) LEFT JOIN ConOtroVehic ON ConOtroVehic.CodCOnOtro = Ficha.AcciConOtro) LEFT JOIN CausaConductor ON CausaConductor.CodCausaCond = Ficha.CodCausaCond1) LEFT JOIN CausaVehic ON CausaVehic.CodCausaVehic = Ficha.CausaVehic) LEFT JOIN clima ON clima.CodClima = Ficha.Clima1) LEFT JOIN SentidoTrans ON SentidoTrans.CodSentidoTrans = Ficha.SentidoTrans) LEFT JOIN LugarAccid ON LugarAccid.CodLugarAccid = Ficha.LugarAccid) LEFT JOIN ColisionContra ON ColisionContra.CodColision = Ficha.CodColisContra1) WHERE Ficha.Fecha between '" & Format(pFecha1, "yyyy-mm-dd") & "' AND '" & Format(pFecha2, "yyyy-mm-dd") & "' " & mAccidentes & " ORDER BY Ficha.Fecha, Ficha.NroOrden"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oPuntosNegros = Rs1
End Function

Public Function oPeatonCiclis(ByVal pFecha1 As String, ByVal pFecha2 As String) As ADODB.Recordset
   strSql1 = "SELECT DISTINCTROW Ficha.NroOrden,Ficha.Fecha,Ficha.Hora,Ficha.Progresiva,CONCAT_WS('-',SentidoTrans.CodSentidoTrans,SentidoTrans.Detalle) AS mSentTrans,CONCAT_WS('-',LugarAccid.CodLugarAccid,LugarAccid.Detalle) AS mLugarAccid,CONCAT_WS('-',otros.CodOtros,otros.Detalle) As mAccidOtros,TotalHeridos.TotalHerido,TotalFallecidos.TotalFallec,Space(5) AS mHora " _
           & "FROM ((((((Ficha LEFT JOIN TotalFallecidos ON Ficha.NroOrden = TotalFallecidos.NroOrden) LEFT JOIN TotalHeridos ON Ficha.NroOrden = TotalHeridos.NroOrden) INNER JOIN TotalVehiculos ON Ficha.NroOrden = TotalVehiculos.NroOrden) LEFT JOIN otros ON Ficha.AccidOtro = otros.CodOtros) LEFT JOIN LugarAccid ON Ficha.LugarAccid = LugarAccid.CodLugarAccid) LEFT JOIN SentidoTrans ON Ficha.SentidoTrans = SentidoTrans.CodSentidoTrans)  WHERE (Ficha.Fecha between '" & Format(pFecha1, "yyyy-mm-dd") & "' AND '" & Format(pFecha2, "yyyy-mm-dd") & "') AND ((Ficha.AccidOtro = '02') OR (Ficha.AccidOtro = '03')) " & mAccidentes & "  ORDER BY Ficha.NroOrden,Ficha.Fecha"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oPeatonCiclis = Rs1
End Function

Public Function oTipoVehic(ByVal pFecha1 As String, ByVal pFecha2 As String, ByVal pCodTipoVehic As String) As ADODB.Recordset
   strSql1 = "SELECT DISTINCTROW Ficha.NroOrden,Ficha.Fecha, CONCAT_WS('-',TipoVehiculo.CodtipoVehic,TipoVehiculo.Descripcion) AS mTipoVehic " _
             & "FROM ((Ficha INNER JOIN VehiculosInvolucr ON Ficha.NroOrden = VehiculosInvolucr.NroOrden) LEFT JOIN TipoVehiculo ON TipoVehiculo.CodTipoVehic = VehiculosInvolucr.CodTipoVehic) WHERE (Ficha.Fecha between '" & Format(pFecha1, "yyyy-mm-dd") & "' AND '" & Format(pFecha2, "yyyy-mm-dd") & "') AND (VehiculosInvolucr.CodTipoVehic >= '" & pCodTipoVehic & "') AND (VehiculosInvolucr.CodTipoVehic = TipoVehiculo.CodtipoVehic) and Ficha.codtipoficha='01' ORDER BY Ficha.NroOrden"

   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oTipoVehic = Rs1
End Function

Public Function oParametrizada(ByVal pFecha1 As String, ByVal pFecha2 As String, ByVal pFrom As String, ByVal pWhere As String) As ADODB.Recordset
   strSql1 = "SELECT DISTINCTROW Ficha.NroOrden FROM " & pFrom & " WHERE Ficha.Fecha between '" & Format(pFecha1, "yyyy-mm-dd") & "' AND '" & Format(pFecha2, "yyyy-mm-dd") & "'" & pWhere
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oParametrizada = Rs1
End Function

Public Function iCountHeridos(ByVal pMes As String, ByVal pAnio As String, _
                        ByVal pLeves As Boolean, ByVal pGraves As Boolean, ByVal pTipo As String) As Integer
Dim mWhere As String

   mWhere = ""
   If pLeves And pGraves Then
      mWhere = " and (V.herido in ('1','2') or V.codestado in ('02','03'))"
   Else
      If pLeves Then
         mWhere = " and (V.herido='1' or V.codestado in ('02')) "
      End If
      If pGraves Then
         mWhere = " and (V.herido='2' or V.codestado in ('03')) "
      End If
   End If
   strSql1 = "SELECT count(V.nroorden) as total FROM  Ficha F, VictimasInvolucr V where " _
     & " month(F.fecha)=" & pMes & "   and year(F.fecha)=" & pAnio & " and F.nroorden=V.nroorden " & mWhere & mAccidentes
   sAsignCmd
   iCountHeridos = 0
   Set Rs1 = cmdTemp.Execute
   If Not Rs1.EOF Then
      iCountHeridos = Rs1!Total
   End If
End Function

Public Function iCountMedTrasl(ByVal pMes As String, ByVal pAnio As String, ByVal pCodTipo As String) As Integer
   strSql1 = "SELECT count(V.nroorden) as total FROM  Ficha F, VictimasInvolucr V where V.codmediotrasl='" & pCodTipo & "' " _
     & " and month(F.fecha)=" & pMes & "   and year(F.fecha)=" & pAnio & " and F.nroorden=V.nroorden  " & mAccidentes
    sAsignCmd
   iCountMedTrasl = 0
   Set Rs1 = cmdTemp.Execute
   If Not Rs1.EOF Then
      iCountMedTrasl = Rs1!Total
   End If
End Function


Public Function iCountFallHerVictVehic(ByVal pFecha1 As String, ByVal pFecha2 As String, _
                                  ByVal pTraza As Boolean, ByVal pHerido As Boolean, ByVal pWhere As String) As Integer
   Dim Rs1 As New ADODB.Recordset
   Dim strSql1 As String

   strSql1 = "select count(B.nrovictima) as Total From Ficha A, VictimasInvolucr B Where A.nroorden = B.nroorden " _
      & " and A.Fecha between '" & Format(pFecha1, "yyyy-mm-dd") & "' AND '" & Format(pFecha2, "yyyy-mm-dd") & "' "
   If pTraza Then
      strSql1 = strSql1 & " and A.lugaraccid <> '09'"
   Else
      strSql1 = strSql1 & " and A.lugaraccid = '09'"
   End If
   If pHerido Then
      strSql1 = strSql1 & " and (B.herido <> '' or B.codestado in ('02','03'))" & mAccidentes
   Else
      strSql1 = strSql1 & " and (B.fallecio <> '' or B.codestado in ('04','05'))" & mAccidentes
   End If
   If Trim(pWhere) <> "" Then
      strSql1 = strSql1 & " " & pWhere
   End If
   cmdTemp.CommandText = strSql1
   cmdTemp.CommandType = 1
   cmdTemp.CommandTimeout = 36000
   Set cmdTemp.ActiveConnection = DBConn
   Set Rs1 = cmdTemp.Execute
   iCountFallHerVictVehic = 0
   If Not Rs1.EOF Then
      iCountFallHerVictVehic = Rs1!Total
   End If
End Function

Public Function oCountFallHerTipo(ByVal pNroFicha As String, ByVal pHerido As Boolean, ByVal pWhere As String) As ADODB.Recordset
   strSql1 = "select codestado, count(nrovictima) as Total From VictimasInvolucr  Where nroorden='" & pNroFicha & "'"
   If pHerido Then
      strSql1 = strSql1 & " and (herido <> '' or codestado in ('02','03'))"
   Else
      strSql1 = strSql1 & " and (fallecio <> '' or codestado in ('04','05'))"
   End If
   If Trim(pWhere) <> "" Then
      strSql1 = strSql1 & " " & pWhere
   End If
    strSql1 = strSql1 & " group by codestado"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oCountFallHerTipo = Rs1
End Function

Public Function iCountColisContra(ByVal pFecha1 As String, ByVal pFecha2 As String, ByVal pKm1 As String, ByVal pKm2 As String, ByVal pCodColis As String) As Integer
   strSql1 = "SELECT count(codcoliscontra1) as total FROM Ficha " _
      & "where fecha between '" & Format(pFecha1, "yyyy-mm-dd") & "' and '" & Format(pFecha2, "yyyy-mm-dd") & "'" _
      & " and progresiva >=" & pKm1 & " and progresiva < " & pKm2 & " and codcoliscontra1='" & pCodColis & "'" & mAccidentes
   sAsignCmd
   iCountColisContra = 0
   Set Rs1 = cmdTemp.Execute
   If Not Rs1.EOF Then
      iCountColisContra = Rs1!Total
   End If
End Function

Public Function iCountIncidentes(ByVal pFecha1 As String, ByVal pFecha2 As String) As Integer
   strSql1 = "SELECT count(nroorden) as total FROM Ficha " _
      & "where fecha between '" & Format(pFecha1, "yyyy-mm-dd") & "' and '" & Format(pFecha2, "yyyy-mm-dd") & "' and codtipoficha='02'"
   sAsignCmd
   iCountIncidentes = 0
   Set Rs1 = cmdTemp.Execute
   If Not Rs1.EOF Then
      iCountIncidentes = Rs1!Total
   End If
End Function

Public Function iTotalHeridosTipoNroOrden(ByVal pNroOrden As String, ByVal pLeve As Boolean) As Integer
   strSql1 = "select count(nrovictima) as Total From VictimasInvolucr where nroorden='" & pNroOrden & "' "
   If pLeve Then
      strSql1 = strSql1 & " and (herido in ('1') or codestado in ('02'))"
   Else
      strSql1 = strSql1 & " and (herido in ('2') or codestado in ('03'))"
   End If
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   iTotalHeridosTipoNroOrden = 0
   If Not Rs1.EOF Then
      iTotalHeridosTipoNroOrden = Rs1!Total
   End If
End Function

Public Function sDanios(ByVal pNroFicha As String) As String
   strSql1 = "SELECT nroorden, sum(grail + grailp + njcab + njbarr + pcab + pbarr + cilum + svert + psos + otros)  as total FROM daniosgco where nroorden='" & pNroFicha & "' group by 1"
   sAsignCmd
   sDanios = "NO"
   On Error Resume Next
   Set Rs1 = cmdTemp.Execute
   If Err.Description = "" Then
      If Not Rs1.EOF Then
         If Rs1!Total > 0 Then
            sDanios = "SI"
         End If
      End If
   End If
End Function

Public Function xInsertFicha(ByVal pNroOrden As String, ByVal pFecha As String, ByVal pHoraAvi As String, ByVal pHoraArr As String, _
                             ByVal pKm As String, ByVal pCarril As String, ByVal pCodConOtro As String, ByVal pCodCContra As String, _
                             ByVal pCodOtros As String, _
                             ByVal pSent As String, _
                             ByVal pCodCSecund As String, _
                             ByVal pCodClima As String, _
                             ByVal pCodERoda As String, _
                             ByVal pSenial As String, _
                             ByVal pCodIlumina As String, _
                             ByVal pCodCausaCond As String, _
                             ByVal pCodCausaCond2 As String, _
                             ByVal pCodCausaCond3 As String, _
                             ByVal pCodCausaVeh As String, _
                             ByVal pAnexos As String, _
                             ByVal pCodAlfa As String, _
                             ByVal pCodConf As String, _
                             ByVal pCodVisibilidad As String, _
                             ByVal pCodPavim As String, _
                             ByVal pCodInconv As String, _
                             ByVal pObs As String, _
                             ByVal pCodTipoFicha As String, ByVal pCodRamal As String) As Boolean
                             
   Dim mFlag As Boolean
   
   xInsertFicha = False
   strSql1 = "select * from Ficha where nroorden='" & pNroOrden & "'"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   If Rs1.EOF Then
       pKm = Replace(pKm, ",", ".") 'mp 20160309
       '(NroOrden,CodPatrullero,Policia,MovilNro,Fecha,Hora,HoraLlegada,Progresiva,Tramo,Interseccion,Carril,AutoridGCO,AcciconOtro,CodColisContra1,CodColisContra2,CodColisContra3,AccidOtro,SentidoTrans, LugarAccid,Clima1,EstCalzada,EstBanquina,DemarcHoriz,DemarcVert,Iluminac,CodCausaCond1,CodCausaCond2,CodCausaCond3,CausaVehic,DanosGCO,BombName,BombDepto,AmbulName,AmbulDepend,PoliciaName,PoliciaDepend,GendarName,GendarDepend,GruaName,GruaEmpr,PoliCientName,PoliCientDepend,OtraAutoridad,Foto)
       strSql1 = "insert into Ficha values ('" & pNroOrden & "','','','','" & Format(pFecha, "yyyy-mm-dd") & "','" & pHoraAvi & "','" & pHoraArr & "'," _
         & pKm & ",'','','" & pCarril & "','','" & pCodConOtro & "','" & pCodCContra & "','','','" & pCodOtros & "'," & pSent & ",'" & pCodCSecund & "','" & pCodClima & "'," _
         & "'" & pCodERoda & "','','" & pSenial & "','','" & pCodIlumina & "','" & pCodCausaCond & "','" & pCodCausaCond2 & "','" & pCodCausaCond3 & "','" & pCodCausaVeh & "','','','','','','','','','','','','','',''," _
         & "'" & pAnexos & "','" & pCodAlfa & "','" & pCodConf & "','" & pCodVisibilidad & "','" & pCodInconv & "','" & pCodPavim & "','" & pObs & "','" & pCodTipoFicha & "', " & pCodRamal & ")"
      cmdTemp.CommandText = strSql1
      
      On Error Resume Next
      cmdTemp.Execute
      
      If Err.Description <> "" Then
         MsgBox Err.Description
         xInsertFicha = True
      End If
   Else
      xInsertFicha = True
   End If
End Function



Public Function xInsertFichaMIGRACION(ByVal pNroOrden As String, ByVal pFecha As String, ByVal pHoraAvi As String, ByVal pHoraArr As String, _
                             ByVal pKm As String, ByVal pCarril As String, ByVal pCodConOtro As String, ByVal pCodCContra As String, ByVal pCodCContra2 As String, _
                             ByVal pCodOtros As String, _
                             ByVal pSent As String, _
                             ByVal pCodCSecund As String, _
                             ByVal pCodClima As String, _
                             ByVal pCodERoda As String, _
                             ByVal pSenial As String, _
                             ByVal pCodIlumina As String, _
                             ByVal pCodCausaCond As String, _
                             ByVal pCodCausaCond2 As String, _
                             ByVal pCodCausaCond3 As String, _
                             ByVal pCodCausaVeh As String, _
                             ByVal pAnexos As String, _
                             ByVal pCodAlfa As String, _
                             ByVal pCodConf As String, _
                             ByVal pCodVisibilidad As String, _
                             ByVal pCodPavim As String, _
                             ByVal pCodInconv As String, _
                             ByVal pObs As String, _
                             ByVal pCodTipoFicha As String, ByVal pCodRamal As String) As Boolean
                             
   Dim mFlag As Boolean
   
   xInsertFichaMIGRACION = False
   strSql1 = "select * from Ficha where nroorden='" & pNroOrden & "'"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   If Rs1.EOF Then
       pKm = Replace(pKm, ",", ".") 'mp 20160309
       '(NroOrden,CodPatrullero,Policia,MovilNro,Fecha,Hora,HoraLlegada,Progresiva,Tramo,Interseccion,Carril,AutoridGCO,AcciconOtro,CodColisContra1,CodColisContra2,CodColisContra3,AccidOtro,SentidoTrans, LugarAccid,Clima1,EstCalzada,EstBanquina,DemarcHoriz,DemarcVert,Iluminac,CodCausaCond1,CodCausaCond2,CodCausaCond3,CausaVehic,DanosGCO,BombName,BombDepto,AmbulName,AmbulDepend,PoliciaName,PoliciaDepend,GendarName,GendarDepend,GruaName,GruaEmpr,PoliCientName,PoliCientDepend,OtraAutoridad,Foto)
       strSql1 = "insert into FichaMIGRACION values ('" & pNroOrden & "','','','','" & Format(pFecha, "yyyy-mm-dd") & "','" & pHoraAvi & "','" & pHoraArr & "'," _
         & pKm & ",'','','" & pCarril & "','','" & pCodConOtro & "','" & pCodCContra & "','" & pCodCContra2 & "','','" & pCodOtros & "'," & pSent & ",'" & pCodCSecund & "','" & pCodClima & "'," _
         & "'" & pCodERoda & "','','" & pSenial & "','','" & pCodIlumina & "','" & pCodCausaCond & "','" & pCodCausaCond2 & "','" & pCodCausaCond3 & "','" & pCodCausaVeh & "','','','','','','','','','','','','','',''," _
         & "'" & pAnexos & "','" & pCodAlfa & "','" & pCodConf & "','" & pCodVisibilidad & "','" & pCodInconv & "','" & pCodPavim & "','" & pObs & "','" & pCodTipoFicha & "', " & pCodRamal & ")"
      cmdTemp.CommandText = strSql1
      
      On Error Resume Next
      cmdTemp.Execute
      
      If Err.Description <> "" Then
         MsgBox Err.Description
         xInsertFichaMIGRACION = True
      End If
   Else
      xInsertFichaMIGRACION = True
   End If
End Function

















Public Function xUpFicha(ByVal pNroOrden As String, ByVal pFecha As String, ByVal pHoraAvi As String, ByVal pHoraArr As String, _
                           ByVal pKm As String, ByVal pCarril As String, ByVal pCodConOtro As String, ByVal pCodCContra As String, _
                           ByVal pCodOtros As String, _
                           ByVal pSent As String, _
                           ByVal pCodCSecund As String, _
                           ByVal pCodClima As String, _
                           ByVal pCodERoda As String, _
                           ByVal pSenial As String, _
                           ByVal pCodIlumina As String, _
                           ByVal pCodCausaCond As String, _
                           ByVal pCodCausaCond2 As String, _
                           ByVal pCodCausaCond3 As String, _
                           ByVal pCodCausaVeh As String, _
                           ByVal pAnexos As String, _
                           ByVal pCodAlfa As String, _
                           ByVal pCodConf As String, _
                           ByVal pCodVisibilidad As String, _
                           ByVal pCodPavim As String, _
                           ByVal pCodInconv As String, _
                           ByVal pObs As String, _
                           ByVal pCodTipoFicha As String, _
                           ByVal pCodRamal As String) As Boolean
   Dim mFlag As Boolean
   
   xUpFicha = True
   
   pKm = Replace(pKm, ",", ".") 'mp 20160309
   strSql1 = "update Ficha set fecha='" & Format(pFecha, "yyyy-mm-dd") & "', Hora='" & pHoraAvi & "', HoraLlegada='" & pHoraArr & "'," _
      & "progresiva=" & pKm & ", carril='" & pCarril & "', acciconotro='" & pCodConOtro & "', codcoliscontra1='" & pCodCContra & "', " _
      & "accidotro='" & pCodOtros & "', sentidotrans=" & pSent & ", lugaraccid='" & pCodCSecund & "', clima1='" & pCodClima & "', " _
      & "estcalzada='" & pCodERoda & "', demarchoriz='" & pSenial & "', iluminac='" & pCodIlumina & "', codcausacond1='" & pCodCausaCond & "', " _
      & "codcausacond2='" & pCodCausaCond2 & "', codcausacond3='" & pCodCausaCond3 & "',causavehic='" & pCodCausaVeh & "', " _
      & "foto='" & pAnexos & "', codalfa='" & pCodAlfa & "', codconfig='" & pCodConf & "', codvisibilidad='" & pCodVisibilidad & "', " _
      & "codinconv='" & pCodInconv & "', codpavim='" & pCodPavim & "', obs='" & pObs & "', codtipoficha='" & pCodTipoFicha & "', " _
      & "codramal=" & pCodRamal _
      & " WHERE nroorden='" & pNroOrden & "'"
         
      sAsignCmd
      On Error Resume Next
      cmdTemp.Execute
      If Err.Description <> "" Then
         MsgBox Err.Description
         xUpFicha = True
      End If
End Function

Public Function xInsFichaOlder(ByVal pNroOrden As String, ByVal pCodPatrullero As String, ByVal pPolicia As String, ByVal pCodMovil As String, _
                             ByVal pFecha As String, ByVal pHora As String, ByVal pHoraLleg As String, ByVal pKm As Double, _
                             ByVal pTramo As String, ByVal pInters As String, ByVal pCarril As String, ByVal pAutoridad As String, _
                             ByVal pAcciConOtro As String, ByVal pCodColis1 As String, ByVal pCodColis2 As String, ByVal pCodColis3 As String, _
                             ByVal pAccidOtro As String, ByVal pSent As String, ByVal pLugarAccid As String, ByVal pClima As String, _
                             ByVal pEstCalz As String, ByVal pEstBanq As String, ByVal pDemarcH As String, ByVal pDemarcV As String, _
                             ByVal pCodIlum As String, ByVal pCodCausa1 As String, ByVal pCodCausa2 As String, ByVal pCodCausa3 As String, _
                             ByVal pCodCausaVeh1 As String, ByVal pDanosGCo As String, ByVal pBombName As String, ByVal pBombDepto As String, _
                             ByVal pAMBUname As String, ByVal pAMBUdep As String, ByVal pPoliName As String, ByVal pPoliDepend As String, _
                             ByVal pGendName As String, ByVal pGendDepend As String, ByVal pGRUAname As String, ByVal pGRUAempr As String, _
                             ByVal pPoliCient As String, ByVal pPoliCientDep As String, ByVal pOtraAutor As String, ByVal pFoto As String) As Boolean
   strSql1 = "select * from Ficha where nroorden='" & pNroOrden & "'"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   If Rs1.EOF Then
      strSql1 = "insert into Ficha (NroOrden,CodPatrullero,Policia,MovilNro,Fecha,Hora,HoraLlegada,Progresiva,Tramo,Interseccion,Carril,AutoridGCO,AcciconOtro,CodColisContra1,CodColisContra2,CodColisContra3,AccidOtro,SentidoTrans, LugarAccid,Clima1,EstCalzada,EstBanquina,DemarcHoriz,DemarcVert,Iluminac,CodCausaCond1,CodCausaCond2,CodCausaCond3,CausaVehic,DanosGCO,BombName,BombDepto,AmbulName,AmbulDepend,PoliciaName,PoliciaDepend,GendarName,GendarDepend,GruaName,GruaEmpr,PoliCientName,PoliCientDepend,OtraAutoridad,Foto) values " _
         & "('" & pNroOrden & "','" & pCodPatrullero & "','" & pPolicia & "','" & pCodMovil & "','" & Format(pFecha, "yyyy/mm/dd") & "','" & pHora & "','" & pHoraLleg & "'," & pKm & "," _
         & "'" & pTramo & "','" & pInters & "','" & pCarril & "','" & pAutoridad & "','" & pAcciConOtro & "','" & pCodColis1 & "','" & pCodColis2 & "','" & pCodColis3 & "','" & pAccidOtro & "','" & pSent & "','" & pLugarAccid & "','" & pClima & "','" & pEstCalz & "'," _
         & "'" & pEstBanq & "','" & pDemarcH & "','" & pDemarcV & "','" & pCodIlum & "','" & pCodCausa1 & "','" & pCodCausa2 & "','" & pCodCausa3 & "'," _
         & "'" & pCodCausaVeh1 & "','" & pDanosGCo & "','" & pBombName & "','" & pBombDepto & "','" & pAMBUname & "'," _
         & "'" & pAMBUdep & "','" & pPoliName & "','" & pPoliDepend & "','" & pGendName & "','" & pGendDepend & "'," _
         & "'" & pGRUAname & "','" & pGRUAempr & "','" & pPoliCient & "','" & pPoliCientDep & "','" & pOtraAutor & "','" & pFoto & "')"
      cmdTemp.CommandText = strSql1
      cmdTemp.Execute
      xInsFichaOlder = True
   Else
      xInsFichaOlder = False
   End If
End Function

Public Function xUpFichaOlder(ByVal pNroOrden As String, ByVal pCodPatrullero As String, ByVal pPolicia As String, ByVal pCodMovil As String, _
                             ByVal pFecha As String, ByVal pHora As String, ByVal pHoraLleg As String, ByVal pKm As Double, _
                             ByVal pTramo As String, ByVal pInters As String, ByVal pCarril As String, ByVal pAutoridad As String, _
                             ByVal pAcciConOtro As String, ByVal pCodColis1 As String, ByVal pCodColis2 As String, ByVal pCodColis3 As String, _
                             ByVal pAccidOtro As String, ByVal pSent As String, ByVal pLugarAccid As String, ByVal pClima As String, _
                             ByVal pEstCalz As String, ByVal pEstBanq As String, ByVal pDemarcH As String, ByVal pDemarcV As String, _
                             ByVal pCodIlum As String, ByVal pCodCausa1 As String, ByVal pCodCausa2 As String, ByVal pCodCausa3 As String, _
                             ByVal pCodCausaVeh1 As String, ByVal pDanosGCo As String, ByVal pBombName As String, ByVal pBombDepto As String, _
                             ByVal pAMBUname As String, ByVal pAMBUdep As String, ByVal pPoliName As String, ByVal pPoliDepend As String, _
                             ByVal pGendName As String, ByVal pGendDepend As String, ByVal pGRUAname As String, ByVal pGRUAempr As String, _
                             ByVal pPoliCient As String, ByVal pPoliCientDep As String, ByVal pOtraAutor As String, ByVal pFoto As String) As Boolean
   strSql1 = "select * from Ficha where nroorden='" & pNroOrden & "'"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   If Rs1.EOF Then
      strSql1 = "update Ficha SET CodPatrullero = '" & pCodPatrullero & "' ,Policia = '" & pPolicia & "',MovilNro = '" & pCodMovil & "', " _
            & " Fecha = '" & Format(pFecha, "yyyy/mm/dd") & "',Hora = '" & pHora & "',HoraLlegada = '" & pHoraLleg & "',Progresiva = " & pKm & "," _
            & " Tramo = '" & pTramo & "',Interseccion = '" & pInters & "',Carril = '" & pCarril & "',AutoridGCO = '" & pAutoridad & "', " _
            & " AcciconOtro = '" & pAcciConOtro & "',CodColisContra1 = '" & pCodColis1 & "',CodColisContra2 = '" & pCodColis2 & "',CodColisContra3 = '" & pCodColis3 & "'," _
            & " AccidOtro = '" & pAccidOtro & "',SentidoTrans = '" & pSent & "', LugarAccid = '" & pLugarAccid & "',Clima1 = '" & pClima & "'," _
            & " EstCalzada = '" & pEstCalz & "',EstBanquina = '" & pEstBanq & "',DemarcHoriz = '" & pDemarcH & "',DemarcVert = '" & pDemarcV & "', " _
            & " Iluminac = '" & pCodIlum & "',CodCausaCond1 = '" & pCodCausa1 & "',CodCausaCond2 = '" & pCodCausa2 & "',CodCausaCond3 = '" & pCodCausa3 & "'," _
            & " CausaVehic = '" & pCodCausaVeh1 & "',DanosGCO = '" & pDanosGCo & "',BombName = '" & pBombName & "',BombDepto = '" & pBombDepto & "', " _
            & " AmbulName = '" & pAMBUname & "',AmbulDepend = '" & pAMBUdep & "',PoliciaName = '" & pPoliName & "',PoliciaDepend = '" & pPoliDepend & "'," _
            & " GendarName = '" & pGendName & "',GendarDepend = '" & pGendDepend & "',GruaName = '" & pGRUAname & "',GruaEmpr = '" & pGRUAempr & "'," _
            & " PoliCientName = '" & pPoliCient & "',PoliCientDepend = '" & pPoliCientDep & "',OtraAutoridad = '" & pOtraAutor & "',Foto = '" & pFoto & "' " _
            & " where NroOrden =  '" & pNroOrden & "'"
      cmdTemp.CommandText = strSql1
      cmdTemp.Execute
      xUpFichaOlder = True
   Else
      xUpFichaOlder = False
   End If
End Function

Public Function xInsertIntGCO(ByVal pNroOrden As String, _
                             ByVal pPatrullero1 As String, _
                             ByVal pPatrullero2 As String, _
                             ByVal pPolad As String, _
                             ByVal pCodMovil As String) As Boolean
   Dim mFlag As Boolean
   strSql1 = "select * from intergco where nroorden='" & pNroOrden & "' and codmovil='" & pCodMovil & "'"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   If Rs1.EOF Then
      strSql1 = "insert into intergco values ('" & pNroOrden & "','" & pCodMovil & "','" & pPatrullero1 & "','" & pPatrullero2 & "','" & pPolad & "')"
      cmdTemp.CommandText = strSql1
      cmdTemp.Execute
      xInsertIntGCO = True
   Else
      xInsertIntGCO = False
   End If
End Function

Public Function xInsertIntTerc(ByVal pNroOrden As String, _
                             ByVal pCodTipo As String, _
                             ByVal pMovil As String, _
                             ByVal pPersonal As String, _
                             ByVal pDependencia As String) As Boolean
   Dim mFlag As Boolean
   strSql1 = "select * from interterceros where nroorden='" & pNroOrden & "' and codtipo='" & pCodTipo & "' and movil='" & pMovil & "'"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   If Rs1.EOF Then
      strSql1 = "insert into interterceros values ('" & pNroOrden & "','" & pCodTipo & "','" & pMovil & "','" & pPersonal & "','" & pDependencia & "')"
      cmdTemp.CommandText = strSql1
      cmdTemp.Execute
      xInsertIntTerc = True
   Else
      xInsertIntTerc = False
   End If
End Function

Public Function xInsertDescr(ByVal pNroOrden As String, _
                             ByVal pArribo As String, _
                             ByVal pComentario As String, _
                             ByVal pMecanica As String, _
                             ByVal pObs As String) As Boolean
   Dim mFlag As Boolean
   strSql1 = "select * from fichadescr where nroorden='" & pNroOrden & "'"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   If Rs1.EOF Then
      strSql1 = "insert into fichadescr values ('" & pNroOrden & "','" & pArribo & "','" & pComentario & "','" & pMecanica & "','" & pObs & "')"
      cmdTemp.CommandText = strSql1
      cmdTemp.Execute
      xInsertDescr = True
   Else
      xInsertDescr = False
   End If
End Function

Public Function xInsertObserv(ByVal pNroOrden As String, _
                             ByVal pIndice As Integer, _
                             ByVal pDescr As String) As Boolean
   Dim mFlag As Boolean
   strSql1 = "select * from fichaobs where nroorden='" & pNroOrden & "' and indice=" & pIndice & ""
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   If Rs1.EOF Then
      strSql1 = "insert into fichaobs values ('" & pNroOrden & "'," & pIndice & ",'" & pDescr & "')"
      cmdTemp.CommandText = strSql1
      cmdTemp.Execute
      xInsertObserv = True
   Else
      xInsertObserv = False
   End If
End Function

Public Function xInsertDaniosGCO(ByVal pNroOrden As String, _
                             ByVal pD1 As Integer, _
                             ByVal pD2 As Integer, _
                             ByVal pD3 As Integer, _
                             ByVal pD4 As Integer, _
                             ByVal pD5 As Integer, _
                             ByVal pD6 As Integer, _
                             ByVal pD7 As Integer, _
                             ByVal pD8 As Integer, _
                             ByVal pD9 As Integer, _
                             ByVal pD10 As Integer) As Boolean
   Dim mFlag As Boolean
   strSql1 = "select * from daniosgco where nroorden='" & pNroOrden & "'"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   If Rs1.EOF Then
      strSql1 = "insert into daniosgco values ('" & pNroOrden & "'," & pD1 & "," & pD2 & "," & pD3 & "," & pD4 & "," & pD5 & "," & pD6 & "," & pD7 & "," & pD8 & "," & pD9 & "," & pD10 & ")"
      cmdTemp.CommandText = strSql1
      cmdTemp.Execute
      xInsertDaniosGCO = True
   Else
      xInsertDaniosGCO = False
   End If
End Function

Public Function xInsertVehiculos(ByVal pNroOrden As String, _
                             ByVal pLetra As String, _
                             ByVal pTitular As String, _
                             ByVal pCodTVehic As String, _
                             ByVal pDominio As String, _
                             ByVal pCodColor As String, _
                             ByVal pCodMarca As String, _
                             ByVal pCodMod As String, _
                             ByVal pCodSeguro As String, _
                             ByVal pNroPol As String, _
                             ByVal pEstGral As String, _
                             ByVal pEstNeum As String, _
                             ByVal pAnexoEst As String) As Boolean
   Dim mFlag As Boolean
   strSql1 = "select * from VehiculosInvolucr where nroorden='" & pNroOrden & "' and letra='" & pLetra & "'"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   If Rs1.EOF Then
      strSql1 = "insert into VehiculosInvolucr values ('" & pNroOrden & "','" & pLetra & "','" & pCodTVehic & "'," _
         & "'" & pCodMarca & "','" & pCodMod & "','" & pDominio & "','','','','','','" & pCodSeguro & "','" & pNroPol & "'," _
         & "'" & pTitular & "','" & pCodColor & "','" & pEstGral & "','" & pEstNeum & "','" & pAnexoEst & "')"
      xInsertVehiculos = False
      cmdTemp.CommandText = strSql1
      
      On Error Resume Next
      cmdTemp.Execute
      If Err.Description <> "" Then
         MsgBox Err.Description
         xInsertVehiculos = False
      Else
         xInsertVehiculos = True
      End If
   Else
      xInsertVehiculos = False
   End If
End Function


Public Function xInsertVehiculosMIGRACION(ByVal pNroOrden As String, ByVal pLetra As String, _
                                          ByVal pCodTVehic As String) As Boolean
   Dim mFlag As Boolean
   strSql1 = "select * from VehiculosInvolucrMIGRACION where nroorden='" & pNroOrden & "' and letra='" & pLetra & "'"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   If Rs1.EOF Then
      strSql1 = "insert into VehiculosInvolucrMIGRACION values ('" & pNroOrden & "','" & pLetra & "','" & pCodTVehic & "')"
      xInsertVehiculosMIGRACION = False
      cmdTemp.CommandText = strSql1
      
      On Error Resume Next
      cmdTemp.Execute
      If Err.Description <> "" Then
         MsgBox Err.Description
         xInsertVehiculosMIGRACION = False
      Else
         xInsertVehiculosMIGRACION = True
      End If
   Else
      xInsertVehiculosMIGRACION = False
   End If
End Function




Public Function xInsVehicOlder(ByVal pNroOrden As String, ByVal pLetra As String, _
                             ByVal pCodTipoVeh As String, ByVal pCodMarca As String, _
                             ByVal pCodModelo As String, ByVal pDominio As String, _
                             ByVal pConductor As String, ByVal pCodTipoDoc As String, _
                             ByVal pNroDoc As String, ByVal pDomicilio As String, _
                             ByVal pTel As String, ByVal pCodCiaSeg As String, _
                             ByVal pPoliza As String) As Boolean
   Dim mFlag As Boolean
   strSql1 = "select * from VehiculosInvolucr where nroorden='" & pNroOrden & "' and letra='" & pLetra & "'"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   If Rs1.EOF Then
      strSql1 = "insert into VehiculosInvolucr (NroOrden,Letra,CodTipoVehic,CodMarca,Modelo,Dominio,Conductor,CodTipoDoc,NroDocu,Domicilio,Telefono,CodCiaSeguro,NroPoliza) values " _
            & "('" & pNroOrden & "','" & pLetra & "','" & pCodTipoVeh & "','" & pCodMarca & "','" & pCodModelo & "','" & pDominio & "'," _
            & "'" & pConductor & "','" & pCodTipoDoc & "','" & pCodTipoDoc & "','" & pNroDoc & "','" & pDomicilio & "','" & pTel & "'," _
            & "'" & pCodCiaSeg & "','" & pPoliza & "')"
      cmdTemp.CommandText = strSql1
      
      On Error Resume Next
      cmdTemp.Execute
      If Err.Description <> "" Then
         MsgBox Err.Description
         xInsVehicOlder = True
      End If
   Else
      xInsVehicOlder = False
   End If
End Function

Public Function xInsertVictimas(ByVal pNroOrden As String, _
                             ByVal pNro As String, _
                             ByVal pLetra As String, _
                             ByVal pNombre As String, _
                             ByVal pDomic As String, _
                             ByVal pCodTDoc As String, _
                             ByVal pNroDoc As String, _
                             ByVal pEdad As String, _
                             ByVal pTel As String, _
                             ByVal pCodTrasl As String, _
                             ByVal pCodMedTrals As String, _
                             ByVal pCodEst As String, _
                             ByVal pCintu As String, _
                             ByVal pConductor As String) As Boolean
   Dim mFlag As Boolean
   strSql1 = "select * from VictimasInvolucr where nroorden='" & pNroOrden & "' and nrovictima='" & pNro & "'"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   If Rs1.EOF Then
      strSql1 = "insert into VictimasInvolucr values ('" & pNroOrden & "','" & pNro & "','" & pNombre & "','" & pDomic & "', " _
         & "'" & pCodTDoc & "','" & pNroDoc & "','','" & pCodTrasl & "','" & pCodMedTrals & "','','','" & pLetra & "','','" & pEdad & "','','" & pCintu & "','" & pTel & "', " _
         & "'" & pConductor & "','" & pCodEst & "')"
      cmdTemp.CommandText = strSql1
      
      On Error Resume Next
      cmdTemp.Execute
      
      If Err.Description <> "" Then
         MsgBox Err.Description
         xInsertVictimas = False
      Else
         xInsertVictimas = True
      End If
   Else
      xInsertVictimas = False
   End If
End Function

Public Function xInsertVictimasMIGRACION(ByVal pNroOrden As String, ByVal pNro As String, ByVal pCodEst As String) As Boolean
   Dim mFlag As Boolean
   strSql1 = "select * from VictimasInvolucrMIGRACION where NroOrden='" & pNroOrden & "' and NroVictima='" & pNro & "'"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   If Rs1.EOF Then
      strSql1 = "insert into VictimasInvolucrMIGRACION values ('" & pNroOrden & "','" & pNro & "','" & pCodEst & "')"
      cmdTemp.CommandText = strSql1
      
      On Error Resume Next
      cmdTemp.Execute
      
      If Err.Description <> "" Then
         MsgBox Err.Description
         xInsertVictimasMIGRACION = False
      Else
         xInsertVictimasMIGRACION = True
      End If
   Else
      xInsertVictimasMIGRACION = False
   End If
End Function



Public Function xInsVictimasOlder(ByVal pNroOrden As String, ByVal pNroVict As String, _
                             ByVal pNombre As String, ByVal pDomic As String, _
                             ByVal pCodTDoc As String, ByVal pNroDoc As String, _
                             ByVal pCodNacion As String, ByVal pCodTrasl As String, _
                             ByVal pCodMedTrals As String, ByVal pHerido As String, _
                             ByVal pFallecio As String, ByVal pLetra As String, _
                             ByVal pSexo As String, ByVal pEdad As String, _
                             ByVal pCodCivil As String, ByVal pCinturon As String) As Boolean
                             
   
   strSql1 = "select * from VictimasInvolucr where nroorden='" & pNroOrden & "' and nrovictima='" & pNroVict & "'"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   If Rs1.EOF Then
      strSql1 = "insert into VictimasInvolucr (NroOrden, NroVictima, Nombre, Domicilio,TipoDocu,NroDocu,CodNacion,CodLugarTrasl,CodMedioTrasl,Herido,Fallecio,Letra,Sexo,Edad,CodEstCivil,Cinturon) values " _
      & " ('" & pNroOrden & "','" & pNroVict & "','" & pNombre & "','" & pDomic & "', '" & pCodTDoc & "','" & pNroDoc & "'," _
      & " '" & pCodNacion & "','" & pCodTrasl & "', '" & pCodMedTrals & "','" & pHerido & "','" & pFallecio & "','" & pLetra & "'," _
      & " '" & pSexo & "','" & pEdad & "','" & pCodCivil & "','" & pCinturon & "')"
      cmdTemp.CommandText = strSql1
      cmdTemp.Execute
      xInsVictimasOlder = True
   Else
      xInsVictimasOlder = False
   End If
End Function

Public Function xInsMarcas(ByVal pCodTipoVeh As String, ByVal pCodMarca As String, ByVal pDescr As String)
   strSql1 = "insert into Marca (CodTipoVehic, CodMarca, Descripcion) values ('" & pCodTipoVeh & "','" & pCodMarca & "','" & pDescr & "')"
   sAsignCmd
   cmdTemp.Execute
End Function

Public Function xUpdMarcas(ByVal pCodTipoVeh As String, ByVal pCodMarca As String, ByVal pDescr As String)
   strSql1 = "update Marca set Descripcion='" & pDescr & "' where CodTipoVehic='" & pCodTipoVeh & "' and CodMarca='" & pCodMarca & "'"
   sAsignCmd
   cmdTemp.Execute
End Function

Public Function xUpdFichaDescr(ByVal pNroOrden As String, ByVal pDescrArribo As String, ByVal pComentarios As String, _
                              ByVal pMecanica As String, ByVal pObsPatru As String)
   strSql1 = "update fichadescr set descr_arribo = '" & pDescrArribo & "', comentarios='" & pComentarios & "'," _
         & " mecanica='" & pMecanica & "', obs_patrullero='" & pObsPatru & "' where NroOrden =  '" & pNroOrden & "'"
   sAsignCmd
   cmdTemp.Execute
End Function

Public Function xUpAuxiliar(ByVal pNroOrden As String)
   strSql1 = "update Auxiliar set ultnroorden='" & pNroOrden & "'"
   sAsignCmd
   cmdTemp.Execute
End Function

Public Function xDeleteTable(ByVal pTabla As String, ByVal pWhere As String)
   strSql1 = "delete from " & pTabla & " where " & pWhere
   sAsignCmd
   cmdTemp.Execute
End Function

Public Function xInsTabla(ByVal pTabla As String, ByVal pCodigo As String, ByVal pDescr As String)
   If pTabla = "TipoVehiculo" Or pTabla = "LugarTrasl" Or pTabla = "CiaSeguros" Then
      strSql1 = "insert into " & pTabla & " values ('" & pCodigo & "','" & pDescr & "',null)"
   Else
      strSql1 = "insert into " & pTabla & " values ('" & pCodigo & "','" & pDescr & "')"
   End If
   sAsignCmd
   cmdTemp.Execute
End Function

Public Function xUpdTabla(ByVal pTabla As String, ByVal pCodigo As String, ByVal pDescr As String)
   strSql1 = "update " & pTabla & " set " & pDescr & " where " & pCodigo
   sAsignCmd
   cmdTemp.Execute
End Function

Public Function xDelTabla(ByVal pTabla As String, ByVal pCodigo As String)
   strSql1 = "delete from " & pTabla & " where " & pCodigo
   sAsignCmd
   cmdTemp.Execute
End Function

Public Function xInLogDel(ByVal pNroFicha As String, ByVal pFecha As String, ByVal pCodUser As String, ByVal pIP As String)
   pFecha = Format(pFecha, "yyyy-mm-dd hh:mm:ss")
   strSql1 = "insert into logs_delficha values (current_timestamp,'" & pNroFicha & "','" & pFecha & "','" & Trim(Right(MDI.mUser, 20)) & "','" & Left(pIP, Len(pIP) - 1) & "')"
   sAsignCmd
   cmdTemp.Execute
End Function

Public Function oTotVehTipoNroOrden(ByVal pNroOrden As String) As ADODB.Recordset
   strSql1 = "SELECT codtipovehic, count(codtipovehic) as total From VehiculosInvolucr " _
   & " where nroorden='" & pNroOrden & "' group by codtipovehic order by 1"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oTotVehTipoNroOrden = Rs1
End Function

Public Function oTotVehTipoNroOrdenMIGRACION(ByVal pNroOrden As String) As ADODB.Recordset
   strSql1 = "SELECT codtipovehic, count(codtipovehic) as total From VehiculosInvolucrMIGRACION " _
   & " where nroorden='" & pNroOrden & "' group by codtipovehic order by 1"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oTotVehTipoNroOrdenMIGRACION = Rs1
End Function

Public Function iTotalHerLevesNroOrden(ByVal pNroOrden As String) As Integer
   strSql1 = "select count(nrovictima) as Total From VictimasInvolucr where nroorden='" & pNroOrden & "' and (herido in ('1') or codestado in ('02'))"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   iTotalHerLevesNroOrden = 0
   If Not Rs1.EOF Then
      iTotalHerLevesNroOrden = Rs1!Total
   End If
End Function

Public Function iTotalHerLevesNroOrdenMIGRACION(ByVal pNroOrden As String) As Integer
   strSql1 = "select count(nrovictima) as Total From VictimasInvolucrMIGRACION where nroorden='" & pNroOrden & "' and codestado in ('02')"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   iTotalHerLevesNroOrdenMIGRACION = 0
   If Not Rs1.EOF Then
      iTotalHerLevesNroOrdenMIGRACION = Rs1!Total
   End If
End Function

Public Function iCountSinLesionHerVictVehic(ByVal pFecha1 As String, ByVal pFecha2 As String, _
                                  ByVal pTraza As Boolean, ByVal pWhere As String) As Integer
   Dim Rs1 As New ADODB.Recordset
   Dim strSql1 As String

   strSql1 = "select count(B.nrovictima) as Total From Ficha A, VictimasInvolucr B Where A.nroorden = B.nroorden " _
      & " and A.Fecha between '" & Format(pFecha1, "yyyy-mm-dd") & "' AND '" & Format(pFecha2, "yyyy-mm-dd") & "' " _
      & " and B.codestado='01'"
   If pTraza Then
      strSql1 = strSql1 & " and A.lugaraccid <> '09'" & mAccidentes
   Else
      strSql1 = strSql1 & " and A.lugaraccid = '09'" & mAccidentes
   End If
   
   If Trim(pWhere) <> "" Then
      strSql1 = strSql1 & " " & pWhere
   End If
   cmdTemp.CommandText = strSql1
   cmdTemp.CommandType = 1
   cmdTemp.CommandTimeout = 36000
   Set cmdTemp.ActiveConnection = DBConn
   Set Rs1 = cmdTemp.Execute
   iCountSinLesionHerVictVehic = 0
   If Not Rs1.EOF Then
      iCountSinLesionHerVictVehic = Rs1!Total
   End If
End Function


