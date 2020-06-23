VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clMantEd"
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
DBConn.ConnectionTimeout = 500
DBConn.ConnectionTimeout = 500
'Set DBConn = mObj.oAbreConexion("4", mIPServer, "MantEd", "ssvv", "ssvv") 'oAbreConexion esta en el modulo1
Set DBConn = mObj.oAbreConexion("5", mIPServer, "MantEd", "ssvv", "ssvv") '20190527
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

Public Function oEjecutarSelect(ByVal pSentencia As String) As ADODB.Recordset
strSql1 = pSentencia
sAsignCmd
Set oEjecutarSelect = cmdTemp.Execute
End Function

Public Function sCampoDescrip(ByVal pTabla As String, ByVal pWhereCodigo As String, ByVal pCampoDescr As Integer) As String
strSql1 = "SELECT * FROM " & pTabla & " WHERE " & pWhereCodigo
sAsignCmd
Set Rs1 = cmdTemp.Execute
sCampoDescrip = ""
If Not Rs1.EOF Then
   sCampoDescrip = Rs1.Fields(pCampoDescr)
End If
End Function

Public Function ObtMaxParte() As Integer
strSql1 = "SELECT MAX(Parte) AS MaxParte FROM Registros"
sAsignCmd
Set Rs1 = cmdTemp.Execute
ObtMaxParte = NVL(Rs1!MaxParte, 0)
End Function

Public Sub InsRegistros(ByVal pParte As String, ByVal pFechaSolic As String, ByVal pCodEdificio As String, ByVal pDescripSolic As String, ByVal pPrioridad As String, ByVal pCodSuperv As String, ByVal pEstado As String, ByVal pOpGen As String, ByVal pOrigen As String, ByVal pTMant As String, ByVal pRubro As String, ByVal pEstim As Integer, ByVal pObserv As String, ByVal pSolicEspecial As Integer)
Dim TiempoAdmis As String

TiempoAdmis = pEstim * 1.5
TiempoAdmis = Replace(TiempoAdmis, ",", ".")
strSql1 = "INSERT INTO Registros (Parte, FechaSolic, CodEdificio, DescripSolic, Prioridad, Origen, CodSuperv, Estado, OpGen, TipoMant, CodRubro, TiempoEstim, TiempoAdmis, Observaciones,SolicEspecial) VALUES (" & pParte & ",'" & pFechaSolic & "','" & pCodEdificio & "','" & pDescripSolic & "','" & pPrioridad & "','" & pOrigen & "','" & pCodSuperv & "','" & pEstado & "','" & pOpGen & "','" & pTMant & "','" & pRubro & "'," & pEstim & "," & TiempoAdmis & ",'" & pObserv & "'," & pSolicEspecial & ")"
sAsignCmd
cmdTemp.Execute
End Sub

Public Sub UpdRegistros(ByVal pEdOP As String, ByVal pDscOP As String, ByVal pFechaAsist As String, ByVal pHoraIniAsist As String, ByVal pHoraFinAsist As String, ByVal pTipoMant As String, ByVal pCodRubro As String, ByVal pTpoEstim As String, ByVal pTpoReal As String, ByVal pTiempoAdmis As String, ByVal pManoObra As String, ByVal pHoras As String, ByVal pPesos As String, ByVal pMateriales As String, ByVal pObservaciones As String, ByVal pFecPro As String, ByVal pFecTer As String, ByVal pEstado As String, ByVal pParte As String)

pTpoEstim = Replace(pTpoEstim, ",", ".")
pTpoReal = Replace(pTpoReal, ",", ".")
pTiempoAdmis = Replace(pTiempoAdmis, ",", ".")
pHoras = Replace(pHoras, ",", ".")
pPesos = Replace(pPesos, ",", ".")
strSql1 = "UPDATE Registros SET CodEdificio='" & pEdOP & "',DescripSolic='" & pDscOP & "', " & IIf(pFechaAsist <> "", "FechaAsist='" & Format(pFechaAsist, "yyyy-mm-dd") & "',", "") & IIf(pHoraIniAsist <> "", "HoraIniAsist='" & pHoraIniAsist & "',", "") & IIf(pHoraFinAsist <> "", "HoraFinAsist='" & pHoraFinAsist & "',", "") & " TipoMant='" & pTipoMant & "', CodRubro='" & pCodRubro & "', TiempoEstim=" & pTpoEstim & "," & IIf(pTpoReal <> "", "TiempoReal=" & pTpoReal & ",", "") & " TiempoAdmis = " & pTiempoAdmis & "," & IIf(pManoObra <> "", "ManoObra='" & pManoObra & "',", "") & IIf(pHoras <> "", "Horas=" & pHoras & ",", "") & IIf(pPesos <> "", "Pesos=" & pPesos & ",", "") & IIf(pMateriales <> "", "Materiales='" & pMateriales & "',", "") & IIf(pObservaciones <> "", "Observaciones='" & pObservaciones & "',", "") & " FecPro = '" & Format(pFecPro, "yyyy-mm-dd hh:mm:ss") & "'," & IIf(pFecTer <> "", "FecTer='" & Format(pFecTer, "yyyy-mm-dd hh:mm:ss") & "',", "") & " Estado='" & pEstado & "' WHERE Parte=" & pParte
sAsignCmd
cmdTemp.Execute

End Sub

Public Function ObtCostoMO(pPeriodo) As Double
strSql1 = "SELECT Valor FROM Costos WHERE Periodo = (SELECT MAX(Periodo) FROM Costos WHERE Codigo = 'MAO' AND Periodo <= '" & pPeriodo & "')"
sAsignCmd
Set Rs1 = cmdTemp.Execute
ObtCostoMO = NVL(Rs1!Valor, 0)
End Function

Public Sub DelRegistros(ByVal pParte As Double)
strSql1 = "DELETE FROM Registros WHERE Parte = " & pParte
sAsignCmd
cmdTemp.Execute
End Sub

Public Sub AnularParte(ByVal pParte As Double, ByVal pOpAnu As String, pFecAnu As String, pMotivoAnu As String)
   strSql1 = "INSERT INTO AnulacionesParte() VALUES (" & pParte & ",'" & pOpAnu & "','" & Format(pFecAnu, "yyyy-mm-dd hh:mm:ss") & "','" & pMotivoAnu & "') "
   sAsignCmd
   cmdTemp.Execute
End Sub

'Public Sub InsAdicional(ByVal pParte As Double, ByVal pFechaSolic As String, ByVal pFechaAsist As String, ByVal pHoraIniAsist As String, ByVal pHoraFinAsist As String, ByVal pCodEdificio As String, ByVal pDescripSolic As String, ByVal pPrioridad As String, ByVal pOrigen As String, ByVal pCodSuperv As String, ByVal pTipoMant As String, ByVal pCodRubro As String, ByVal pTiempoEstim As Double, ByVal pTiempoReal As Double, ByVal pTiempoAdmis As Double, ByVal pManoObra As String, ByVal pHoras As Double, ByVal pPesos As Double, ByVal pMateriales As String, ByVal pObservaciones As String, ByVal pEstado As String, ByVal pOpGen As String, ByVal pFecPro As String, ByVal pFecTer As String)

Public Sub InsAdicional(ByVal pParte As Double, ByVal pFechaSolic As String, ByVal pFechaAsist As String, ByVal pHoraIniAsist As String, ByVal pHoraFinAsist As String, ByVal pCodEdificio As String, ByVal pDescripSolic As String, ByVal pPrioridad As String, ByVal pOrigen As String, ByVal pCodSuperv As String, ByVal pTipoMant As String, ByVal pCodRubro As String, ByVal pTiempoEstim As String, ByVal pTiempoReal As String, ByVal pTiempoAdmis As String, ByVal pManoObra As String, ByVal pHoras As String, ByVal pPesos As String, ByVal pMateriales As String, ByVal pObservaciones As String, ByVal pEstado As String, ByVal pOpGen As String, ByVal pFecPro As String, ByVal pFecTer As String)
pTiempoEstim = Replace(pTiempoEstim, ",", ".")
pTiempoReal = Replace(pTiempoReal, ",", ".")
pTiempoAdmis = Replace(pTiempoAdmis, ",", ".")
pHoras = Replace(pHoras, ",", ".")
pPesos = Replace(pPesos, ",", ".")
 
strSql1 = "INSERT INTO Registros VALUES (" & pParte & ",'" & pFechaSolic & "','" & pFechaAsist & "','" & pHoraIniAsist & "','" & pHoraFinAsist & "','" & pCodEdificio & "','" & pDescripSolic & "','" & pPrioridad & "','" & pOrigen & "','" & pCodSuperv & "','" & pTipoMant & "','" & pCodRubro & "'," & pTiempoEstim & "," & pTiempoReal & "," & pTiempoAdmis & ",'" & pManoObra & "'," & pHoras & "," & pPesos & ",'" & pMateriales & "','" & pObservaciones & "','" & pEstado & "','" & pOpGen & "','" & Format(pFecPro, "yyyy-mm-dd hh:mm:ss") & "','" & Format(pFecTer, "yyyy-mm-dd hh:mm:ss") & "',null,null,null,0)"
sAsignCmd
cmdTemp.Execute
End Sub

Public Sub UpdFechaSolic(ByVal pParte As Double, ByVal pFechaSolic As String)
strSql1 = "UPDATE Registros SET FechaSolic = '" & Format(pFechaSolic, "yyyy-mm-dd hh:mm:ss") & "' WHERE Parte = " & pParte
sAsignCmd
cmdTemp.Execute
End Sub

Public Sub UpdValReg(ByVal pEstado As String, pOpValid As String, ByVal pObsValid As String, ByVal pParte As String)
strSql1 = "UPDATE Registros SET Estado = '" & pEstado & "', ObservVal = '" & pObsValid & "', OpVal = '" & pOpValid & "', FecVal = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' WHERE Parte = " & pParte
sAsignCmd
cmdTemp.Execute
End Sub

Public Function ObtCodSuperv(ByVal pParte As String) As String
strSql1 = "SELECT CodSuperv FROM Registros WHERE Parte = '" & pParte & "'"
sAsignCmd
Set Rs1 = cmdTemp.Execute
ObtCodSuperv = NVL(Rs1!CodSuperv, "")
End Function


Public Function iCountSolpasExcelPartesPendientes() As Integer
   iCountSolpasExcelPartesPendientes = 0
   strSql1 = "SELECT COUNT(AUX.ZonaMantEdil) AS CantSolapas " & _
               "FROM " & _
               "( " & _
                  "SELECT DISTINCT E.ZonaMantEdil " & _
                  "FROM " & _
                     "Registros R " & _
                  "INNER JOIN " & _
                     "Edificios E ON R.CodEdificio = CONCAT(E.Tramo , ' - ' , E.Descripcion) " & _
                  "WHERE Estado IN ('P','G') " & _
                  "AND FecTer is null " & _
                  "AND SolicEspecial = 0 " & _
               "UNION " & _
                  "SELECT DISTINCT 'Solicitud Especial' AS ZonaMantEdil FROM Registros " & _
                  "WHERE Estado IN ('P','G') " & _
                  "AND FecTer is null " & _
                  "AND SolicEspecial = 1 " & _
               ") AUX ;"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   If Not Rs1.EOF Then
      iCountSolpasExcelPartesPendientes = Rs1!cantSolapas
   End If
End Function


Public Function oDetallePartesPendientes(ByVal pZonaMantEdil As String) As ADODB.Recordset

      strSql1 = "SELECT AUX.* FROM " & _
                  "( " & _
                     "SELECT Parte,E.Descripcion,FechaSolic,DescripSolic,Prioridad,TipoMant,CASE WHEN SolicEspecial = 0 THEN ZonaMantEdil ELSE 'Solicitud Especial' END AS ZonaMantEdil " & _
                     "FROM " & _
                        "Registros R  " & _
                      "Left Join  " & _
                      "AnulacionesParte A ON R.Parte = A.ParteAnu " & _
                     "INNER JOIN " & _
                        "Edificios E ON R.CodEdificio = CONCAT(E.ZonaMantEdil , ' - ' , E.Descripcion)  " & _
                     "WHERE Estado in ('P','G')  " & _
                     "AND FecTer is null  " & _
                     "AND A.ParteAnu IS NULL " & _
                     "ORDER BY ZonaMantEdil  " & _
                  ") AUX  " & _
                  "WHERE AUX.ZonaMantEdil = '" & pZonaMantEdil & "'  " & _
                  "ORDER BY AUX.Parte"
                  
                  
                  
                  
'         strSql1 = "SELECT AUX.* FROM " & _
'               "( " & _
'                  "SELECT Parte,E.Descripcion,FechaSolic,DescripSolic,Prioridad,TipoMant,CASE WHEN SolicEspecial = 0 THEN ZonaMantEdil ELSE 'Solicitud Especial' END AS ZonaMantEdil " & _
'                  "FROM " & _
'                     "Registros R  " & _
'                  "INNER JOIN " & _
'                     "Edificios E ON R.CodEdificio = CONCAT(E.Tramo , ' - ' , E.Descripcion)  " & _
'                  "WHERE Estado in ('P','G')  " & _
'                  "AND FecTer is null  " & _
'                  "ORDER BY ZonaMantEdil  " & _
'               ") AUX  " & _
'               "WHERE AUX.ZonaMantEdil = '" & pZonaMantEdil & "'  " & _
'               "ORDER BY AUX.Parte"
                                    
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oDetallePartesPendientes = Rs1
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
