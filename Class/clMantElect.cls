VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clMantElect"
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

Private Type parametroSql
    Nombre        As String
    Tipo          As DataTypeEnum
    Direccion     As ParameterDirectionEnum
    Tamanio       As ADO_LONGPTR
    Valor         As Variant
End Type


Dim miArray() As String

Private Sub Class_Initialize()
DBConn.ConnectionTimeout = 500
DBConn.ConnectionTimeout = 500
Set DBConn = mObj.oAbreConexion("5", mIPServer, "MantElect", "ssvv", "ssvv") 'oAbreConexion esta en el modulo1
Rs1.CursorType = adOpenDynamic
End Sub

Private Sub Class_Terminate()
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

Private Sub sAsignCmd_SP(ByVal cmd As ADODB.Command)
   cmd.CommandText = strSql1
   cmd.CommandType = adCmdStoredProc
   cmd.CommandTimeout = 36000
   Set cmd.ActiveConnection = DBConn
End Sub

'Private Sub sAsignCmdPrueba()
'cmdTemp.CommandText = strSql1
'cmdTemp.CommandType = adCmdText
'cmdTemp.CommandTimeout = 36000
'Set cmdTemp.ActiveConnection = DBConn
'End Sub

Public Function oEjecutarSelect(ByVal pSentencia As String) As ADODB.Recordset
strSql1 = pSentencia
sAsignCmd
Set oEjecutarSelect = cmdTemp.Execute
End Function


Public Function oEjecutarSelectPrueba(ByVal pSentencia As String) As ADODB.Recordset
strSql1 = pSentencia
'sAsignCmdPrueba
Set oEjecutarSelectPrueba = cmdTemp.Execute
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


Public Sub InsRegistros(ByVal pParte As String, ByVal pFechaSolic As String, ByVal pCodEdificio As String, ByVal pDescripSolic As String, ByVal pPrioridad As String, ByVal pCodSuperv As String, ByVal pEstado As String, ByVal pOpGen As String, ByVal pSectorAire As Integer)
strSql1 = "INSERT INTO Registros (Parte, FechaSolic, CodEdificio, Descripcion, Prioridad, CodSuperv, Estado, OpGen, Origen,SectorAire) VALUES (" & pParte & ",'" & pFechaSolic & "','" & pCodEdificio & "','" & pDescripSolic & "','" & pPrioridad & "','" & pCodSuperv & "','" & pEstado & "','" & pOpGen & "', 'O'," & pSectorAire & ")"
sAsignCmd
cmdTemp.Execute
End Sub

Public Sub UpdRegistros(ByVal pFechaIniAsist As String, ByVal pFechaFinAsist As String, ByVal pSegundaDesc As String, ByVal pRubro As String, ByVal pSubRubro As String, ByVal pCantidad As String, ByVal pHoras As String, ByVal pEstado As String, ByVal pOpPro As String, ByVal pFecPro As String, ByVal pOpTer As String, ByVal pFecTer As String, ByVal pParte As String)
'strSql1 = "UPDATE Registros SET FechaIniAsist = '" & Format(pFechaIniAsist, "yyyy-mm-dd hh:mm:ss") & "'" & IIf(pFechaFinAsist <> "", ", FechaFinAsist = '" & Format(pFechaFinAsist, "yyyy-mm-dd hh:mm:ss") & "'", "") & ", SegundaDesc = '" & pSegundaDesc & "', Rubro = '" & pRubro & "', SubRubro = '" & pSubRubro & "'" & IIf(pCantidad <> "", ", Cantidad = " & pCantidad, "") & IIf(pHoras <> "", ", Horas = " & pHoras, "") & ", Estado = '" & pEstado & "'" & IIf(pOpPro <> "", ", OpPro = '" & pOpPro & "'", "") & IIf(pFecPro <> "", ", FecPro = '" & Format(pFecPro, "yyyy-mm-dd hh:mm:ss") & "'", "") & IIf(pOpTer <> "", ", OpTer = '" & pOpTer & "'", "") & IIf(pFecTer <> "", ",  FecTer = '" & Format(pFecTer, "yyyy-mm-dd hh:mm:ss") & "'", "") & " WHERE Parte = " & pParte
pCantidad = Replace(pCantidad, ",", ".")
pHoras = Replace(pHoras, ",", ".")
strSql1 = "UPDATE Registros SET FechaIniAsist = '" & Format(pFechaIniAsist, "yyyy-mm-dd hh:mm:ss") & "'" & IIf(pFechaFinAsist <> "", ", FechaFinAsist = '" & Format(pFechaFinAsist, "yyyy-mm-dd hh:mm:ss") & "'", "") & ", SegundaDesc = '" & pSegundaDesc & "', Rubro = '" & pRubro & "', SubRubro = '" & pSubRubro & "', Interv = 1" & IIf(pCantidad <> "", ", Cantidad = " & pCantidad, "") & IIf(pHoras <> "", ", Horas = " & pHoras, "") & ", Estado = '" & pEstado & "'" & IIf(pOpPro <> "", ", OpPro = '" & pOpPro & "'", "") & IIf(pFecPro <> "", ", FecPro = '" & Format(pFecPro, "yyyy-mm-dd hh:mm:ss") & "'", "") & IIf(pOpTer <> "", ", OpTer = '" & pOpTer & "'", "") & IIf(pFecTer <> "", ",  FecTer = '" & Format(pFecTer, "yyyy-mm-dd hh:mm:ss") & "'", "") & " WHERE Parte = " & pParte
sAsignCmd
cmdTemp.Execute
End Sub

Public Sub UpdValReg(ByVal pEstado As String, pOpValid As String, ByVal pObsValid As String, ByVal pParte As String)
strSql1 = "UPDATE Registros SET Estado = '" & pEstado & "', ObserValid = '" & pObsValid & "', OpVal = '" & pOpValid & "', FecVal = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' WHERE Parte = " & pParte
sAsignCmd
cmdTemp.Execute
End Sub
Public Sub InsRelev(ByVal pParte As String, ByVal pFechaSolic As String, ByVal pCodEdificio As String, ByVal pDescripcion As String, ByVal pPrioridad As String, ByVal pCodSuperv As String, ByVal pFechaIniAsist As String, ByVal pFechaFinAsist As String, ByVal pSegundaDesc As String, ByVal pRubro As String, ByVal pSubRubro As String, ByVal pCantidad As String, ByVal pHoras As String, ByVal pEstado As String, ByVal pOrigen As String, ByVal pOpGen As String, ByVal pOpPro As String, ByVal pFecPro As String, ByVal pOpTer As String, ByVal pFecTer As String, ByVal pObserValid As String, ByVal pOpVal As String, ByVal pFecVal As String, ByVal pInterv As Integer)
'Public Sub InsRelev(ByVal pParte As String, ByVal pFechaSolic As String, ByVal pCodEdificio As String, ByVal pDescripcion As String, ByVal pPrioridad As String, ByVal pCodSuperv As String, ByVal pFechaIniAsist As String, ByVal pFechaFinAsist As String, ByVal pSegundaDesc As String, ByVal pRubro As String, ByVal pSubRubro As String, ByVal pCantidad As Double, ByVal pHoras As String, ByVal pEstado As String, ByVal pOrigen As String, ByVal pOpGen As String, ByVal pOpPro As String, ByVal pFecPro As String, ByVal pOpTer As String, ByVal pFecTer As String, ByVal pObserValid As String, ByVal pOpVal As String, ByVal pFecVal As String, ByVal pInterv As Integer)
strSql1 = "INSERT INTO Registros (Parte, FechaSolic, CodEdificio, Descripcion, Prioridad, CodSuperv, FechaIniAsist, FechaFinAsist, SegundaDesc, Rubro, SubRubro, Cantidad, Horas, Estado, Origen, OpGen, OpPro, FecPro, OpTer, FecTer, ObserValid, OpVal, FecVal, Interv) VALUES (" & pParte & ",'" & Format(pFechaSolic, "yyyy-mm-dd hh:mm:ss") & "','" & pCodEdificio & "','" & pDescripcion & "','" & pPrioridad & "','" & pCodSuperv & "','" & Format(pFechaIniAsist, "yyyy-mm-dd hh:mm:ss") & "','" & Format(pFechaFinAsist, "yyyy-mm-dd hh:mm:ss") & "','" & pSegundaDesc & "','" & pRubro & "','" & pSubRubro & "'," & pCantidad & "," & pHoras & ",'" & pEstado & "','" & pOrigen & "','" & pOpGen & "','" & pOpPro & "','" & Format(pFecPro, "yyyy-mm-dd hh:mm:ss") & "','" & pOpTer & "','" & Format(pFecTer, "yyyy-mm-dd hh:mm:ss") & "','" & pObserValid & "','" & pOpVal & "','" & Format(pFecVal, "yyyy-mm-dd hh:mm:ss") & "'," & pInterv & ")"
sAsignCmd
cmdTemp.Execute
End Sub

Public Function ObtCodSuperv(ByVal pParte As String) As String
strSql1 = "SELECT CodSuperv FROM Registros WHERE Parte = '" & pParte & "'"
sAsignCmd
Set Rs1 = cmdTemp.Execute
ObtCodSuperv = NVL(Rs1!CodSuperv, "")
End Function


Public Function esSupervisorElectrico(ByVal pCodSuperv As String) As Boolean
   Dim ret As Boolean
   ret = False

   strSql1 = "SELECT * FROM MailsxElectrico where codusuario = '" & pCodSuperv & "' and SectorAire = 0;"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute

   If Not Rs1.EOF Then
      ret = True
   End If
   esSupervisorElectrico = ret
End Function

Public Function esParteOperaciones(ByVal pParte As String) As Boolean
   Dim ret As Boolean
   ret = False

   strSql1 = " SELECT AUX.Parte FROM " & _
             "( " & _
             " SELECT Parte FROM COM_Comunicados_Det " & _
             " UNION " & _
             " SELECT Parte FROM REL_Relevamientos_Det " & _
             " ) AUX " & _
             " WHERE Parte = " & pParte & ";"
   
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   
   'Si no trae datos entonces es parte de Operaciones
   If Rs1.EOF Then
      ret = True
   End If
   esParteOperaciones = ret
End Function


Public Function xDeleteMoTecnicos(ByVal pParte As Double) As Boolean
   strSql1 = "Delete from Partes_MoTecnicos where Parte =" & pParte & ";"
   sAsignCmd
   cmdTemp.Execute
   xDeleteMoTecnicos = True
End Function


Public Sub xInsertMoTecnico(ByVal pParte As String, ByVal pCodMoTecnico As String)
   strSql1 = "Insert into Partes_MoTecnicos values (" & pParte & ",'" & pCodMoTecnico & "')"
   sAsignCmd
   On Error Resume Next
   cmdTemp.Execute
   If Err.Description <> "" Then
      MsgBox Err.Description
   End If
End Sub

Public Function xDeleteVehAsignados(ByVal pParte As Double) As Boolean
   strSql1 = "Delete from Partes_Vehiculos where Parte =" & pParte & ";"
   sAsignCmd
   cmdTemp.Execute
   xDeleteVehAsignados = True
End Function

Public Sub xInsertVehAsignados(ByVal pParte As String, ByVal pCodVehiculo As String, ByVal pKmInicio As String, ByVal pKmFin As String)
   pKmInicio = Replace(pKmInicio, ",", ".")
   pKmFin = Replace(pKmFin, ",", ".")
   
   strSql1 = "Insert into Partes_Vehiculos values (" & pParte & ",'" & pCodVehiculo & "'," & pKmInicio & "," & pKmFin & ")"
   sAsignCmd
   On Error Resume Next
   cmdTemp.Execute
   If Err.Description <> "" Then
      MsgBox Err.Description
   End If
End Sub


Public Function xinsOT(ByVal pCodUsuario As String, ByRef pvPartes_OT() As Double, ByRef pvVehiculos_OT() As String, ByRef pvVehiculosEsp_OT() As String, ByRef pvMO_Tecnicos_OT() As String, ByRef pvSubrubros_OT() As String, ByRef FechaOT As Date) As Integer
   On Error GoTo ErrorHandler
   Dim strSql1 As String
   Dim mi As Integer
   Dim rs As ADODB.Recordset
   Dim retOT As Integer
   
   cmdTemp.CommandTimeout = 36000
   Set cmdTemp.ActiveConnection = DBConn
   cmdTemp.ActiveConnection.BeginTrans

   cmdTemp.CommandType = adCmdStoredProc
  
   insOT_H pCodUsuario, cmdTemp
   Set rs = cmdTemp.Execute
   
   retOT = rs!vIdOt
   FechaOT = rs!vfecha

   rs.Close
     
   If pvPartes_OT(0) <> 0 Then
      For mi = LBound(pvPartes_OT) To UBound(pvPartes_OT)
         insOT_Partes retOT, pvPartes_OT(mi), FechaOT, pCodUsuario, cmdTemp
         'cierraOT_Partes retOT, pvPartes_OT(mi), FechaOT, pCodUsuario, cmdTemp
         
         cmdTemp.Execute
      Next
   End If
      
   If pvVehiculos_OT(0) <> "00" Then
      For mi = LBound(pvVehiculos_OT) To UBound(pvVehiculos_OT)
         insOT_Vehiculos retOT, pvVehiculos_OT(mi), 0, 0, pCodUsuario, cmdTemp
         cmdTemp.Execute
      Next
   End If

   If pvVehiculosEsp_OT(0) <> "00" Then
      For mi = LBound(pvVehiculosEsp_OT) To UBound(pvVehiculosEsp_OT)
         insOT_Vehiculos retOT, pvVehiculosEsp_OT(mi), 0, 0, pCodUsuario, cmdTemp
         cmdTemp.Execute
      Next
   End If
   
   If pvMO_Tecnicos_OT(0) <> "00" Then
      For mi = LBound(pvMO_Tecnicos_OT) To UBound(pvMO_Tecnicos_OT)
         insOT_MO_Tecnicos retOT, pvMO_Tecnicos_OT(mi), pCodUsuario, cmdTemp
         cmdTemp.Execute
      Next
   End If
 
   If pvSubrubros_OT(0) <> "000000" Then
      For mi = LBound(pvSubrubros_OT) To UBound(pvSubrubros_OT)
         insOT_Subrubros retOT, pvSubrubros_OT(mi), pCodUsuario, cmdTemp
         cmdTemp.Execute
      Next
   End If
 
   cmdTemp.ActiveConnection.CommitTrans
   cmdTemp.CommandType = 1
   xinsOT = retOT
   Exit Function
ErrorHandler:

   MsgBox "Se produjo un Error al generar la Orden de Trabajo. Comunicarse con el Administrador" & vbCrLf & "Detalle del error: " & Err.Number & " " & Err.Description, vbCritical, "Error"
   cmdTemp.ActiveConnection.RollbackTrans
   cmdTemp.CommandType = 1

End Function


'---------------------------------------
'Sp: insOT_H
'---------------------------------------
Public Sub insOT_H(CodUsuario As String, ByRef cmd As ADODB.Command)
   Dim parametroSql As parametroSql
   Dim aParamSql(0) As parametroSql
   parametroSql.Nombre = "vCodUsuario"
   parametroSql.Tipo = adVarChar
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 25
   parametroSql.Valor = CodUsuario
   aParamSql(0) = parametroSql
   xStoreProcedureWithParam_Transaccional "insOT_H", aParamSql, cmd
End Sub
   
Public Sub cerrarOT_H(ByVal IdOT As Integer, ByVal FecInicio As Date, ByVal FecFin As Date, ByVal CodUsuario_Ult_UPD As String, ByRef cmd As ADODB.Command)
   Dim parametroSql As parametroSql
   Dim aParamSql(3) As parametroSql
   
   parametroSql.Nombre = "vIdOT"
   parametroSql.Tipo = adInteger
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 4
   parametroSql.Valor = IdOT
   aParamSql(0) = parametroSql
      
   parametroSql.Nombre = "vFecInicio"
   parametroSql.Tipo = adDBTimeStamp
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 8
   parametroSql.Valor = FecInicio
   aParamSql(1) = parametroSql
      
   parametroSql.Nombre = "vFecFin"
   parametroSql.Tipo = adDBTimeStamp
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 8
   parametroSql.Valor = FecFin
   aParamSql(2) = parametroSql
    
   parametroSql.Nombre = "vCodUsuario_Ult_UPD"
   parametroSql.Tipo = adVarChar
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 25
   parametroSql.Valor = CodUsuario_Ult_UPD
   aParamSql(3) = parametroSql

   xStoreProcedureWithParam_Transaccional "cerrarOT_H", aParamSql, cmd
End Sub





'---------------------------------------
'insOT_Partes
'---------------------------------------

Public Sub insOT_Partes(IdOT As Integer, Parte As Double, Fecha As Date, CodUsuario As String, ByRef cmd As ADODB.Command)

   Dim parametroSql As parametroSql
   Dim aParamSql(3) As parametroSql
   
   parametroSql.Nombre = "vIdOT"
   parametroSql.Tipo = adInteger
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 4
   parametroSql.Valor = IdOT
   aParamSql(0) = parametroSql
   
   parametroSql.Nombre = "vParte"
   parametroSql.Tipo = adDouble
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 8
   parametroSql.Valor = Parte
   aParamSql(1) = parametroSql
   
   parametroSql.Nombre = "vFecha"
   parametroSql.Tipo = adDBTimeStamp
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 8
   parametroSql.Valor = Fecha
   aParamSql(2) = parametroSql
   
   
   parametroSql.Nombre = "vCodUsuario"
   parametroSql.Tipo = adVarChar
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 25
   parametroSql.Valor = CodUsuario
   aParamSql(3) = parametroSql
   
   xStoreProcedureWithParam_Transaccional "insOT_Partes", aParamSql, cmd

End Sub

'---------------------------------------
'cerrarOT_Partes
'---------------------------------------
'cancelarPartes_Det mIdCancel, pvPartes_OT(mi)
Public Sub cerrarOT_Partes(ByVal IdOT As Integer, Parte As Double, Estado As String, CodUsuario As String, ByRef cmd As ADODB.Command)
   Dim parametroSql As parametroSql
   Dim aParamSql(3) As parametroSql
   
   parametroSql.Nombre = "vIdOT"
   parametroSql.Tipo = adInteger
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 4
   parametroSql.Valor = IdOT
   aParamSql(0) = parametroSql
   
   parametroSql.Nombre = "vParte"
   parametroSql.Tipo = adDouble
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 8
   parametroSql.Valor = Parte
   aParamSql(1) = parametroSql
   
   parametroSql.Nombre = "vEstado"
   parametroSql.Tipo = adVarChar
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 2
   parametroSql.Valor = Estado
   aParamSql(2) = parametroSql
  
   parametroSql.Nombre = "vCodUsuario"
   parametroSql.Tipo = adVarChar
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 25
   parametroSql.Valor = CodUsuario
   aParamSql(3) = parametroSql

   xStoreProcedureWithParam_Transaccional "cerrarOT_Partes", aParamSql, cmd
End Sub

Public Sub cancelarPartesNoFinalizadosEnOT(ByVal IdOT As Integer, CodUsuario As String, ByRef cmd As ADODB.Command)
'Public Function cancelarPartesNoFinalizadosEnOT(ByVal IdOT As Integer, CodUsuario As String, ByRef cmd As ADODB.Command) As ADODB.Recordset

   Dim parametroSql As parametroSql
   Dim aParamSql(1) As parametroSql
   
   parametroSql.Nombre = "vIdOT"
   parametroSql.Tipo = adInteger
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 4
   parametroSql.Valor = IdOT
   aParamSql(0) = parametroSql
   
   parametroSql.Nombre = "vCodUsuario"
   parametroSql.Tipo = adVarChar
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 25
   parametroSql.Valor = CodUsuario
   aParamSql(1) = parametroSql

   xStoreProcedureWithParam_Transaccional "cancelarPartesNoFinalizadosEnOT", aParamSql, cmd
End Sub


'---------------------------------------
'insOT_Vehiculos
'---------------------------------------

Public Sub insOT_Vehiculos(ByVal IdOT As Integer, ByVal CodVehiculo As String, ByRef KmInicial As Integer, ByRef KmFinal As Integer, ByVal CodUsuario As String, ByRef cmd As ADODB.Command)

Dim parametroSql As parametroSql
Dim aParamSql(4) As parametroSql
   
With parametroSql
   .Nombre = "vIdOT"
   .Tipo = adInteger
   .Direccion = adParamInput
   .Tamanio = 4
   .Valor = IdOT
    aParamSql(0) = parametroSql
   
   .Nombre = "vCodVehiculo"
   .Tipo = adChar
   .Direccion = adParamInput
   .Tamanio = 2
   .Valor = CodVehiculo
    aParamSql(1) = parametroSql
   
   .Nombre = "vKmInicial"
   .Tipo = adInteger
   .Direccion = adParamInput
   .Tamanio = 4
   .Valor = KmInicial
    aParamSql(2) = parametroSql
   
   .Nombre = "vKmFinal"
   .Tipo = adInteger
   .Direccion = adParamInput
   .Tamanio = 4
   .Valor = KmFinal
    aParamSql(3) = parametroSql
     
   .Nombre = "vCodUsuario"
   .Tipo = adVarChar
   .Direccion = adParamInput
   .Tamanio = 25
   .Valor = CodUsuario
   aParamSql(4) = parametroSql
End With
   
   xStoreProcedureWithParam_Transaccional "insOT_Vehiculos", aParamSql, cmd

End Sub

'---------------------------------------
'delOT_Vehiculos
'---------------------------------------

Public Sub delOT_Vehiculos(ByVal IdOT As Integer, ByRef cmd As ADODB.Command)

   Dim parametroSql As parametroSql
   Dim aParamSql(0) As parametroSql
   
   parametroSql.Nombre = "vIdOT"
   parametroSql.Tipo = adInteger
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 4
   parametroSql.Valor = IdOT
   aParamSql(0) = parametroSql
   
   xStoreProcedureWithParam_Transaccional "delOT_Vehiculos", aParamSql, cmd

End Sub

















'---------------------------------------
'insOT_MO_Tecnicos
'---------------------------------------

Public Sub insOT_MO_Tecnicos(ByVal IdOT As Integer, CodMO_Tecnico As String, CodUsuario As String, ByRef cmd As ADODB.Command)

   Dim parametroSql As parametroSql
   Dim aParamSql(2) As parametroSql
   
   parametroSql.Nombre = "vIdOT"
   parametroSql.Tipo = adInteger
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 4
   parametroSql.Valor = IdOT
   aParamSql(0) = parametroSql
   
   parametroSql.Nombre = "vCodMO_Tecnico"
   parametroSql.Tipo = adChar
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 2
   parametroSql.Valor = CodMO_Tecnico
   aParamSql(1) = parametroSql
   
'   parametroSql.Nombre = "vFecha"
'   parametroSql.Tipo = adDBTimeStamp
'   parametroSql.Direccion = adParamInput
'   parametroSql.Tamanio = 8
'   parametroSql.Valor = Fecha
'   aParamSql(2) = parametroSql
'
   parametroSql.Nombre = "vCodUsuario"
   parametroSql.Tipo = adVarChar
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 25
   parametroSql.Valor = CodUsuario
   aParamSql(2) = parametroSql
   
   xStoreProcedureWithParam_Transaccional "insOT_MO_Tecnicos", aParamSql, cmd

End Sub


'---------------------------------------
'delOT_MO_Tecnicos
'---------------------------------------

Public Sub delOT_MO_Tecnicos(ByVal IdOT As Integer, ByRef cmd As ADODB.Command)

   Dim parametroSql As parametroSql
   Dim aParamSql(0) As parametroSql
   
   parametroSql.Nombre = "vIdOT"
   parametroSql.Tipo = adInteger
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 4
   parametroSql.Valor = IdOT
   aParamSql(0) = parametroSql
   
   xStoreProcedureWithParam_Transaccional "delOT_MO_Tecnicos", aParamSql, cmd

End Sub

















'---------------------------------------
'insOT_Subrubros
'---------------------------------------

Public Sub insOT_Subrubros(ByVal IdOT As Integer, CodSubrubro As String, CodUsuario As String, ByRef cmd As ADODB.Command)

   Dim parametroSql As parametroSql
   Dim aParamSql(2) As parametroSql
   
   parametroSql.Nombre = "vIdOT"
   parametroSql.Tipo = adInteger
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 4
   parametroSql.Valor = IdOT
   aParamSql(0) = parametroSql
   
   parametroSql.Nombre = "vCodSubrubro"
   parametroSql.Tipo = adVarChar
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 6
   parametroSql.Valor = CodSubrubro
   aParamSql(1) = parametroSql
   
'   parametroSql.Nombre = "vFecha"
'   parametroSql.Tipo = adDBTimeStamp
'   parametroSql.Direccion = adParamInput
'   parametroSql.Tamanio = 8
'   parametroSql.Valor = Fecha
'   aParamSql(2) = parametroSql
'
   parametroSql.Nombre = "vCodUsuario"
   parametroSql.Tipo = adVarChar
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 25
   parametroSql.Valor = CodUsuario
   aParamSql(2) = parametroSql
   
   xStoreProcedureWithParam_Transaccional "insOT_Subrubros", aParamSql, cmd

End Sub


'---------------------------------------
'delOT_Subrubros
'---------------------------------------

Public Sub delOT_Subrubros(ByVal IdOT As Integer, ByRef cmd As ADODB.Command)

   Dim parametroSql As parametroSql
   Dim aParamSql(0) As parametroSql
   
   parametroSql.Nombre = "vIdOT"
   parametroSql.Tipo = adInteger
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 4
   parametroSql.Valor = IdOT
   aParamSql(0) = parametroSql
   
   xStoreProcedureWithParam_Transaccional "delOT_Subrubros", aParamSql, cmd

End Sub

'---------------------------------------
'insOT_MovimientoMateriales
'---------------------------------------
Public Sub insOT_MovimientoMateriales(ByVal IdOT As Integer, CodProducto As String, CodUbicacion As String, Cantidad As Double, CantidadBD As Double, CodUsuario As String, ByRef cmd As ADODB.Command)
   Dim parametroSql As parametroSql
   Dim aParamSql(5) As parametroSql
   
   With parametroSql
      .Nombre = "vIdOT"
      .Tipo = adInteger
      .Direccion = adParamInput
      .Tamanio = 4
      .Valor = IdOT
      aParamSql(0) = parametroSql
      
      .Nombre = "vCodProducto"
      .Tipo = adVarChar
      .Direccion = adParamInput
      .Tamanio = 6
      .Valor = CodProducto
      aParamSql(1) = parametroSql
      
      .Nombre = "vCodUbicacion"
      .Tipo = adVarChar
      .Direccion = adParamInput
      .Tamanio = 4
      .Valor = CodUbicacion
      aParamSql(2) = parametroSql
   
      .Nombre = "vCantidad"
      .Tipo = adDouble
      .Direccion = adParamInput
      .Tamanio = 8
      .Valor = Cantidad
      aParamSql(3) = parametroSql
      
      .Nombre = "vCantidadBD"
      .Tipo = adDouble
      .Direccion = adParamInput
      .Tamanio = 8
      .Valor = CantidadBD
      aParamSql(4) = parametroSql
      
      .Nombre = "vCodUsuario"
      .Tipo = adVarChar
      .Direccion = adParamInput
      .Tamanio = 25
      .Valor = CodUsuario
      aParamSql(5) = parametroSql
   End With
   
   xStoreProcedureWithParam_Transaccional "insOT_MovimientoMateriales", aParamSql, cmd
End Sub




Public Sub xCerrarOT(ByVal pIdOT As String, ByVal pFecIni As Date, ByVal pFecFin As Date, ByVal pCodUsuario As String, _
                     ByRef pvPartes_OT() As Double, ByRef pvPartes_Estado() As String, ByRef pvPartesOriginal_OT() As Double, ByRef pvVehiculos_OT() As String, ByRef pvVehiculos_KmIni() As Integer, _
                     ByRef pvVehiculos_KmFin() As Integer, ByRef pvVehiculosEsp_OT() As String, ByRef pvVehiculosEsp_KmIni() As Integer, _
                     ByRef pvVehiculosEsp_KmFin() As Integer, ByRef pvMO_Tecnicos_OT() As String, ByRef pvSubrubros_OT() As String, _
                     ByRef pvMat_CodProd() As String, ByRef pvMat_CodUbic() As String, ByRef pvMat_Cantidad() As Double, ByRef pvMat_CantidadBD() As Double)
   On Error GoTo ErrorHandler
   'Dim strSql1 As String
   
   Dim mi As Integer
   Dim mErrMail As Integer
   Dim mListaDestinatarios As String
   Dim mListaDestinatariosOperac As String
   Dim mTextoMail As String
   Dim mTextoMailOperac As String
   Dim sectorAireFinalizado As Integer
   Dim sectorAireNOFinalizado As Integer
   
      
   cmdTemp.CommandTimeout = 36000
   Set cmdTemp.ActiveConnection = DBConn
   cmdTemp.ActiveConnection.BeginTrans
   cmdTemp.CommandType = adCmdStoredProc
  
   cerrarOT_H pIdOT, pFecIni, pFecFin, pCodUsuario, cmdTemp
   cmdTemp.Execute

   delOT_Vehiculos pIdOT, cmdTemp
   cmdTemp.Execute
   
   delOT_MO_Tecnicos pIdOT, cmdTemp
   cmdTemp.Execute
   
   delOT_Subrubros pIdOT, cmdTemp
   cmdTemp.Execute
   
   If pvPartes_OT(0) <> 0 Then
      For mi = LBound(pvPartes_OT) To UBound(pvPartes_OT)
           cerrarOT_Partes pIdOT, pvPartes_OT(mi), pvPartes_Estado(mi), pCodUsuario, cmdTemp
         cmdTemp.Execute
      Next
   End If
   
   cancelarPartesNoFinalizadosEnOT pIdOT, pCodUsuario, cmdTemp
   cmdTemp.Execute

   If pvVehiculos_OT(0) <> "00" Then
      For mi = LBound(pvVehiculos_OT) To UBound(pvVehiculos_OT)
         insOT_Vehiculos pIdOT, pvVehiculos_OT(mi), pvVehiculos_KmIni(mi), pvVehiculos_KmFin(mi), pCodUsuario, cmdTemp
         cmdTemp.Execute
      Next
   End If

   If pvVehiculosEsp_OT(0) <> "00" Then
      For mi = LBound(pvVehiculosEsp_OT) To UBound(pvVehiculosEsp_OT)
         insOT_Vehiculos pIdOT, pvVehiculosEsp_OT(mi), pvVehiculosEsp_KmIni(mi), pvVehiculosEsp_KmFin(mi), pCodUsuario, cmdTemp
         cmdTemp.Execute
      Next
   End If

   If pvMO_Tecnicos_OT(0) <> "00" Then
      For mi = LBound(pvMO_Tecnicos_OT) To UBound(pvMO_Tecnicos_OT)
         insOT_MO_Tecnicos pIdOT, pvMO_Tecnicos_OT(mi), pCodUsuario, cmdTemp
         cmdTemp.Execute
      Next
   End If

   If pvSubrubros_OT(0) <> "000000" Then
      For mi = LBound(pvSubrubros_OT) To UBound(pvSubrubros_OT)
         insOT_Subrubros pIdOT, pvSubrubros_OT(mi), pCodUsuario, cmdTemp
         cmdTemp.Execute
      Next
   End If

   If pvMat_CodProd(0) <> "000000" Then
      For mi = LBound(pvMat_CodProd) To UBound(pvMat_CodProd)
            If pvMat_Cantidad(mi) <> pvMat_CantidadBD(mi) Then
               insOT_MovimientoMateriales pIdOT, pvMat_CodProd(mi), pvMat_CodUbic(mi), pvMat_Cantidad(mi), pvMat_CantidadBD(mi), pCodUsuario, cmdTemp
               cmdTemp.Execute
            End If
      Next
   End If
 
   cmdTemp.ActiveConnection.CommitTrans
   cmdTemp.CommandType = 1
   
   Dim mRec1 As ADODB.Recordset
   Dim Prioridad As String
   Dim Lugar As String
   Dim Problema As String
   
  
   sectorAireFinalizado = 0
   sectorAireNOFinalizado = 0
   
   Dim qPartesFinalizados As Integer
   qPartesFinalizados = 0
     
     
   '==========================================INICIO: ARMO TEXTO MAIL PARTES SI FINALIZADOS========================================================
   If pvPartes_OT(0) <> 0 Then
      For mi = LBound(pvPartes_OT) To UBound(pvPartes_OT)
         'mp20200401
         'Set mRec1 = oEjecutarSelect("SELECT CodEdificio,Descripcion,Prioridad,SectorAire FROM Registros WHERE Parte = " & pvPartes_OT(mi))
         Set mRec1 = oEjecutarSelect("SELECT R.CodEdificio,R.Descripcion,R.Prioridad,R.SectorAire,O.Finalizado " & _
         " FROM Registros R" & _
         " Left Join OT_Partes O ON R.Parte = O.Parte" & _
         " WHERE IdOT = " & pIdOT & _
         " AND R.Parte = " & pvPartes_OT(mi))
         
         
         If mi = 0 Then
            sectorAireFinalizado = mRec1!SectorAire
         End If
         
         
         mTextoMailOperac = ""
         If Not mRec1.EOF Then
            mTextoMail = mTextoMail & vbTab & "Parte " & pvPartes_OT(mi) & " -" & IIf(mRec1!Finalizado = "SI", "FINALIZADO", "NO TERMINADO") & "- (Prioridad  " & mRec1!Prioridad & "): " & mRec1!CodEdificio & "  -  " & mRec1!descripcion & "" & vbCrLf
            
            '-------------------Inicio: Envio Mails Partes Supervisores Operaciones-------------------------------------------------------------
            'mp20200401
            'If esParteOperaciones(pvPartes_OT(mi)) Then
            If esParteOperaciones(pvPartes_OT(mi)) And mRec1!Finalizado = "SI" Then
               Dim mRec2 As ADODB.Recordset
               mTextoMailOperac = vbCrLf & "Se ha resuelto el Parte  " & pvPartes_OT(mi) & " de Mantenimiento Eléctrico: " & vbCrLf & vbCrLf & "     Descripción de la solicitud:  " & mRec1!descripcion & vbCrLf & vbCrLf & "Verifique el servicio realizado. Gracias"
               
               Set mRec2 = oEjecutarSelect("SELECT DISTINCT Email FROM MailsxSuperv WHERE CodSuperv = '" & ObtCodSuperv(pvPartes_OT(mi)) & "' And FechaBaja IS NULL ")
            
               If Not mRec2.EOF Then
                  mListaDestinatariosOperac = ""
                  Do While Not mRec2.EOF
                     mListaDestinatariosOperac = mListaDestinatariosOperac & mRec2!Email & ";"
                     mRec2.MoveNext
                  Loop
                  If Not fEnviar_Mail_CDO("", mListaDestinatariosOperac, "ausolmail@ausol.com.ar", "Repuesta a Solicitud de Servicios", mTextoMailOperac, "", "587", "ausolmail@ausol.com.ar", "sgedosmildiecisiete1$", True, False) Then
                     mErrMail = mErrMail + 1
                  End If
               End If
               mRec2.Close
            End If
            '-------------------Fin: Envio Mails Partes Supervisores Operaciones-------------------------------------------------------------
            mRec1.Close
      
         End If
         
         qPartesFinalizados = qPartesFinalizados + 1
      Next
   End If
   
   If qPartesFinalizados >= 1 Then
      mTextoMail = "Se han comenzado los siguientes partes correspondientes a esta OT: " & vbCrLf & vbCrLf & mTextoMail
   Else
      mTextoMail = "No se han comenzado partes para esta Orden de Trabajo " & vbCrLf
   End If
   
   mTextoMail = "Orden de Trabajo cerrada: " & pIdOT & vbCrLf & vbCrLf & mTextoMail & vbCrLf
   
   '==========================================FIN: ARMO TEXTO MAIL PARTES SI FINALIZADOS========================================================
   
   
   '==========================================INICIO: ARMO TEXTO MAIL PARTES NO FINALIZDOS (NO COMENZADOS)======================================
   Dim vPartesNoFinalizados_OT() As Double
   Dim cantPartesNoFinalizados As Integer
   Dim mtextoPartesNoFin As String

   cantPartesNoFinalizados = 0
      
   For mi = LBound(pvPartesOriginal_OT) To UBound(pvPartesOriginal_OT)
      If Not estaEnVectorPartesFinalizados(pvPartesOriginal_OT(mi), pvPartes_OT()) Then
         cantPartesNoFinalizados = cantPartesNoFinalizados + 1
         ReDim Preserve vPartesNoFinalizados_OT(0 To cantPartesNoFinalizados - 1)
         vPartesNoFinalizados_OT(cantPartesNoFinalizados - 1) = pvPartesOriginal_OT(mi)
     End If
   Next
      
   'Si existen partes no finalizados(no comenzdos) => armo texto para incluir en el mail
   If cantPartesNoFinalizados > 0 Then
      
      mtextoPartesNoFin = "Por otro lado se informa que los siguientes partes de esta Orden de trabajo no se han comenzado por lo que están nuevamente disponibles para nuevas OTs:" & vbCrLf & vbCrLf
      
      For mi = LBound(vPartesNoFinalizados_OT) To UBound(vPartesNoFinalizados_OT)
         
         Set mRec1 = oEjecutarSelect("SELECT CodEdificio,Descripcion,Prioridad,SectorAire FROM Registros WHERE Parte = " & vPartesNoFinalizados_OT(mi))
         If Not mRec1.EOF Then
            
            If mi = 0 Then
               sectorAireNOFinalizado = mRec1!SectorAire
            End If
            
            Prioridad = mRec1!Prioridad
            Lugar = mRec1!CodEdificio
            Problema = mRec1!descripcion
         End If
         mRec1.Close
         
         mtextoPartesNoFin = mtextoPartesNoFin & vbTab & "Parte " & vPartesNoFinalizados_OT(mi) & " (Prioridad  " & Prioridad & "): " & Lugar & "  -  " & Problema & "" & vbCrLf
      Next
   End If
   
   
   If cantPartesNoFinalizados > 0 Then
      mTextoMail = mTextoMail & mtextoPartesNoFin
   End If
   
   '==========================================FIN: ARMO TEXTO MAIL PARTES NO FINALIZDOS (NO COMENZADOS)==========================================
   mListaDestinatarios = ""
    
    
   '=======================================INICIO: ENVIO MAILS A ELECTRICOS (ELECTRICOS O ELECTRICOS AIRE)=======================================
    'Tomé como referencia el valor del primer elemento de cada vector para ver si era del sector aire o no.
   If sectorAireFinalizado = 1 Or sectorAireNOFinalizado = 1 Then
      Set Rs1 = oEjecutarSelect("SELECT DISTINCT Email FROM MailsxElectrico WHERE SectorAire = 1 and FechaBaja IS NULL ")
   Else
      Set Rs1 = oEjecutarSelect("SELECT DISTINCT Email FROM MailsxElectrico WHERE SectorAire = 0 and FechaBaja IS NULL ")
   End If
   
   
   If Not Rs1.EOF Then
      Do While Not Rs1.EOF
         mListaDestinatarios = mListaDestinatarios & Rs1!Email & ";"
         Rs1.MoveNext
      Loop
       
      If Not fEnviar_Mail_CDO("", mListaDestinatarios, "ausolmail@ausol.com.ar", " Cierre Orden de Trabajo: " & pIdOT, mTextoMail, "", "587", "ausolmail@ausol.com.ar", "sgedosmildiecisiete1$", True, False) Then
         mErrMail = mErrMail + 1
      End If
   End If
   If mErrMail = 0 Then
      MsgBox "Se ha cerrado la OT " & pIdOT & "  y enviado el mensaje correctamente a los responsables.", vbInformation, "Atención"
   Else
      MsgBox "Se ha cerrado la OT " & pIdOT & ", pero NO se ha enviado el correo correctamente", vbExclamation, "Atención"
   End If
   
   
   Set mRec1 = Nothing
   Exit Sub
   '=======================================FIN: ENVIO MAILS A ELECTRICOS (ELECTRICOS O ELECTRICOS AIRE)=======================================
ErrorHandler:

   MsgBox "Se produjo un Error al intentar cerrar la Orden de Trabajo. Comunicarse con el Administrador" & vbCrLf & "Detalle del error: " & Err.Number & " " & Err.Description, vbCritical, "Error"
   cmdTemp.ActiveConnection.RollbackTrans
   cmdTemp.CommandType = 1

End Sub

Public Sub xInsOT_Abastecimiento(ByRef pvEgresosCodProducto() As String, ByRef pvEgresosCodUbicacion() As String, ByRef pvEgresosCantidad() As Double, ByRef pvEgresosCantidadBD() As Double, ByVal pNroVale As String, pCodTipoVale As String, ByVal pIdOT As String, ByVal pCodUbicacion As String, ByVal pCodUsuario As String, ByRef bRet As Boolean)

On Error GoTo ErrorHandler
Dim strSql1 As String
Dim mi As Integer
'Dim prueba As Integer
   
   cmdTemp.CommandType = 1
   cmdTemp.CommandTimeout = 36000
   Set cmdTemp.ActiveConnection = DBConn
   cmdTemp.ActiveConnection.BeginTrans

   
   strSql1 = "INSERT INTO OT_Abastecimiento_H (IdOT,NroVale,CodTipoVale,CodUbicacion,Fecha,CodUsuario)  VALUES (" & pIdOT & ",'" & pNroVale & "','" & pCodTipoVale & "','" & pCodUbicacion & "', NOW(),'" & pCodUsuario & "');"

   cmdTemp.CommandText = strSql1
   cmdTemp.Execute

   cmdTemp.CommandType = adCmdStoredProc

   For mi = LBound(pvEgresosCodProducto) To UBound(pvEgresosCodProducto)
      
'     Genero un error(division por cero) ara probar si funcion el rollBackTrans
'      If mI = 1 Then
'         prueba = 1 / 0
'      End If

'insOT_Abastecimiento_Det`(IN vCodProducto CHAR(6),IN vCodUbicacion CHAR(4), IN vCantidad Double, IN vIdOT INTEGER, IN vCodUsuario VARCHAR(25))
      
      insOT_Abastecimiento_Det pvEgresosCodProducto(mi), pvEgresosCodUbicacion(mi), pvEgresosCantidad(mi), pvEgresosCantidadBD(mi), pIdOT, pCodUsuario, cmdTemp
      cmdTemp.Execute
   Next

   cmdTemp.ActiveConnection.CommitTrans
   cmdTemp.CommandType = 1
   Exit Sub
ErrorHandler:

   MsgBox "Se produjo un Error al realizar el ajuste. Comunicarse con el Administrador" & vbCrLf & "Detalle del error: " & Err.Number & " " & Err.Description, vbCritical, "Error"
   cmdTemp.ActiveConnection.RollbackTrans
   cmdTemp.CommandType = 1
   bRet = False

End Sub

'---------------------------------------
'Sp: insOT_Abastecimiento_Det
'---------------------------------------
Public Sub insOT_Abastecimiento_Det(ByVal CodProducto As String, ByVal CodUbicacion As String, ByVal Cantidad As Double, ByVal CantidadBD As Double, ByVal IdOT As Integer, ByVal CodUsuario As String, ByRef cmd As ADODB.Command)
   Dim parametroSql As parametroSql
   Dim aParamSql(5) As parametroSql
   With parametroSql
   
      .Nombre = "vCodProducto"
      .Tipo = adChar
      .Direccion = adParamInput
      .Tamanio = 6
      .Valor = CodProducto
      aParamSql(0) = parametroSql
      
      .Nombre = "vCodUbicacion"
      .Tipo = adChar
      .Direccion = adParamInput
      .Tamanio = 4
      .Valor = CodUbicacion
      aParamSql(1) = parametroSql
      
      .Nombre = "vCantidad"
      .Tipo = adDouble
      .Direccion = adParamInput
      .Tamanio = 8
      .Valor = Cantidad
      aParamSql(2) = parametroSql
      
      .Nombre = "vCantidadBD"
      .Tipo = adDouble
      .Direccion = adParamInput
      .Tamanio = 8
      .Valor = CantidadBD
      aParamSql(3) = parametroSql
      
      .Nombre = "vIdOT"
      .Tipo = adInteger
      .Direccion = adParamInput
      .Tamanio = 4
      .Valor = IdOT
      aParamSql(4) = parametroSql
      
      .Nombre = "vCodUsuario"
      .Tipo = adVarChar
      .Direccion = adParamInput
      .Tamanio = 25
      .Valor = CodUsuario
      aParamSql(5) = parametroSql
      
   End With
   xStoreProcedureWithParam_Transaccional "insOT_Abastecimiento_Det", aParamSql, cmd
 End Sub

Public Sub xInsMat_Ingresos(ByRef pvIngresosCodProducto() As String, ByRef pvIngresosCodUbicacion() As String, ByRef pvIngresosCantidad() As Double, ByRef pvIngresosCantidadBD() As Double, ByVal pNroVale As String, ByVal pCodTipoVale As String, ByVal pCodUbicacion As String, ByVal pCodUsuario As String, ByRef bRet As Boolean)

On Error GoTo ErrorHandler
Dim strSql1 As String
Dim mi As Integer
   
   cmdTemp.CommandType = 1
   cmdTemp.CommandTimeout = 36000
   Set cmdTemp.ActiveConnection = DBConn
   cmdTemp.ActiveConnection.BeginTrans

   
   strSql1 = "INSERT INTO Mat_Ingresos_H (NroVale,CodTipoVale,CodUbicacion, Fecha,CodUsuario)  VALUES ('" & pNroVale & "','" & pCodTipoVale & "','" & pCodUbicacion & "', NOW(),'" & pCodUsuario & "');"

   cmdTemp.CommandText = strSql1
   cmdTemp.Execute

   cmdTemp.CommandType = adCmdStoredProc

   For mi = LBound(pvIngresosCodProducto) To UBound(pvIngresosCodProducto)
      
      insMat_Ingresos_Det pvIngresosCodProducto(mi), pvIngresosCodUbicacion(mi), pvIngresosCantidad(mi), pvIngresosCantidadBD(mi), pNroVale, pCodTipoVale, pCodUsuario, cmdTemp
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
'Sp: insMat_Ingresos_Det
'---------------------------------------
Public Sub insMat_Ingresos_Det(ByVal CodProducto As String, ByVal CodUbicacion As String, ByVal Cantidad As Double, ByVal CantidadBD As Double, ByVal NroVale As String, ByVal CodTipoVale As String, ByVal CodUsuario As String, ByRef cmd As ADODB.Command)
   Dim parametroSql As parametroSql
   Dim aParamSql(6) As parametroSql
   
   With parametroSql
      
      .Nombre = "vCodProducto"
      .Tipo = adChar
      .Direccion = adParamInput
      .Tamanio = 6
      .Valor = CodProducto
      aParamSql(0) = parametroSql
      
      .Nombre = "vCodUbicacion"
      .Tipo = adChar
      .Direccion = adParamInput
      .Tamanio = 4
      .Valor = CodUbicacion
      aParamSql(1) = parametroSql
      
      .Nombre = "vCantidad"
      .Tipo = adDouble
      .Direccion = adParamInput
      .Tamanio = 8
      .Valor = Cantidad
      aParamSql(2) = parametroSql
      
      .Nombre = "vCantidadBD"
      .Tipo = adDouble
      .Direccion = adParamInput
      .Tamanio = 8
      .Valor = CantidadBD
      aParamSql(3) = parametroSql
      
      .Nombre = "vNroVale"
      .Tipo = adVarChar
      .Direccion = adParamInput
      .Tamanio = 9
      .Valor = NroVale
      aParamSql(4) = parametroSql
      
      .Nombre = "vCodTipoVale"
      .Tipo = adChar
      .Direccion = adParamInput
      .Tamanio = 1
      .Valor = CodTipoVale
      aParamSql(5) = parametroSql
      
      .Nombre = "vCodUsuario"
      .Tipo = adVarChar
      .Direccion = adParamInput
      .Tamanio = 25
      .Valor = CodUsuario
      aParamSql(6) = parametroSql
      
   End With
   xStoreProcedureWithParam_Transaccional "insMat_Ingresos_Det", aParamSql, cmd
 End Sub


Public Sub xAjustarMaterialesOT(ByVal pIdOT As String, ByRef pvMat_CodProd_Ajuste() As String, ByRef pvMat_CodUbic_Ajuste() As String, _
                                 ByRef pvMat_Cantidad_Ajuste() As Double, ByRef pvMat_CantidadBD_Ajuste() As Double, _
                                 ByRef pvMat_CodProd_Orig() As String, ByRef pvMat_CodUbic_Orig() As String, _
                                 ByRef pvMat_Cantidad_Orig() As Double, ByRef pvMat_CantidadBD_Orig() As Double, pCodUsuario As String)
   On Error GoTo ErrorHandler
   Dim mi As Integer
  
   cmdTemp.CommandTimeout = 36000
   Set cmdTemp.ActiveConnection = DBConn
   cmdTemp.ActiveConnection.BeginTrans
   cmdTemp.CommandType = adCmdStoredProc
 
   If pvMat_CodProd_Ajuste(0) <> "000000" Then
      For mi = LBound(pvMat_CodProd_Ajuste) To UBound(pvMat_CodProd_Ajuste)
         If pvMat_Cantidad_Ajuste(mi) <> pvMat_CantidadBD_Ajuste(mi) Then
            insOT_MovimientoMateriales pIdOT, pvMat_CodProd_Ajuste(mi), pvMat_CodUbic_Ajuste(mi), pvMat_Cantidad_Ajuste(mi), pvMat_CantidadBD_Ajuste(mi), pCodUsuario, cmdTemp
            cmdTemp.Execute
         End If
      Next
      For mi = LBound(pvMat_CodProd_Orig) To UBound(pvMat_CodProd_Orig)
         If Not estaEnVectorAjuste(pvMat_CodProd_Orig(mi), pvMat_CodProd_Ajuste()) Then
            insOT_MovimientoMateriales pIdOT, pvMat_CodProd_Orig(mi), pvMat_CodUbic_Orig(mi), 0, pvMat_CantidadBD_Orig(mi), pCodUsuario, cmdTemp
            cmdTemp.Execute
         End If
      Next
   End If
 
   cmdTemp.ActiveConnection.CommitTrans
   cmdTemp.CommandType = 1
   Exit Sub
ErrorHandler:
   MsgBox "Se produjo un Error al intentar ajustar los materiales de la Orden de Trabajo. Comunicarse con el Administrador" & vbCrLf & "Detalle del error: " & Err.Number & " " & Err.Description, vbCritical, "Error"
   cmdTemp.ActiveConnection.RollbackTrans
   cmdTemp.CommandType = 1
End Sub


Public Sub xAjustarAbasteciemientoOT(ByVal pIdOT As String, ByRef pvMat_CodProd_Ajuste() As String, ByRef pvMat_CodUbic_Ajuste() As String, _
                                 ByRef pvMat_Cantidad_Ajuste() As Double, ByRef pvMat_CantidadBD_Ajuste() As Double, _
                                 ByRef pvMat_CodProd_Orig() As String, ByRef pvMat_CodUbic_Orig() As String, _
                                 ByRef pvMat_Cantidad_Orig() As Double, ByRef pvMat_CantidadBD_Orig() As Double, pCodUsuario As String)
   On Error GoTo ErrorHandler
   Dim mi As Integer
  
   cmdTemp.CommandTimeout = 36000
   Set cmdTemp.ActiveConnection = DBConn
   cmdTemp.ActiveConnection.BeginTrans
   cmdTemp.CommandType = adCmdStoredProc
 
   If pvMat_CodProd_Ajuste(0) <> "000000" Then
      For mi = LBound(pvMat_CodProd_Ajuste) To UBound(pvMat_CodProd_Ajuste)
         If pvMat_Cantidad_Ajuste(mi) <> pvMat_CantidadBD_Ajuste(mi) Then
            insOT_Abastecimiento_Det pvMat_CodProd_Ajuste(mi), pvMat_CodUbic_Ajuste(mi), pvMat_Cantidad_Ajuste(mi), pvMat_CantidadBD_Ajuste(mi), pIdOT, pCodUsuario, cmdTemp
            cmdTemp.Execute
         End If
      Next
      For mi = LBound(pvMat_CodProd_Orig) To UBound(pvMat_CodProd_Orig)
         If Not estaEnVectorAjuste(pvMat_CodProd_Orig(mi), pvMat_CodProd_Ajuste()) Then
            insOT_Abastecimiento_Det pvMat_CodProd_Orig(mi), pvMat_CodUbic_Orig(mi), 0, pvMat_CantidadBD_Orig(mi), pIdOT, pCodUsuario, cmdTemp
            cmdTemp.Execute
         End If
      Next
   End If
 
   cmdTemp.ActiveConnection.CommitTrans
   cmdTemp.CommandType = 1
   Exit Sub
ErrorHandler:
   MsgBox "Se produjo un Error al intentar ajustar los materiales de la Orden de Trabajo. Comunicarse con el Administrador" & vbCrLf & "Detalle del error: " & Err.Number & " " & Err.Description, vbCritical, "Error"
   cmdTemp.ActiveConnection.RollbackTrans
   cmdTemp.CommandType = 1
End Sub

Public Sub xAjustarIngreso(ByVal pNroVale As String, ByVal pCodTipoVale As String, ByRef pvMat_CodProd_Ajuste() As String, ByRef pvMat_CodUbic_Ajuste() As String, _
                           ByRef pvMat_Cantidad_Ajuste() As Double, ByRef pvMat_CantidadBD_Ajuste() As Double, _
                           ByRef pvMat_CodProd_Orig() As String, ByRef pvMat_CodUbic_Orig() As String, _
                           ByRef pvMat_Cantidad_Orig() As Double, ByRef pvMat_CantidadBD_Orig() As Double, pCodUsuario As String)
   On Error GoTo ErrorHandler
   Dim mi As Integer
  
   cmdTemp.CommandTimeout = 36000
   Set cmdTemp.ActiveConnection = DBConn
   cmdTemp.ActiveConnection.BeginTrans
   cmdTemp.CommandType = adCmdStoredProc
 
   If pvMat_CodProd_Ajuste(0) <> "000000" Then
      For mi = LBound(pvMat_CodProd_Ajuste) To UBound(pvMat_CodProd_Ajuste)
         If pvMat_Cantidad_Ajuste(mi) <> pvMat_CantidadBD_Ajuste(mi) Then
            'insMat_Ingresos_Det 'mp20191010
            insMat_Ingresos_Det pvMat_CodProd_Ajuste(mi), pvMat_CodUbic_Ajuste(mi), pvMat_Cantidad_Ajuste(mi), pvMat_CantidadBD_Ajuste(mi), pNroVale, pCodTipoVale, pCodUsuario, cmdTemp
            cmdTemp.Execute
         End If
      Next
      For mi = LBound(pvMat_CodProd_Orig) To UBound(pvMat_CodProd_Orig)
         If Not estaEnVectorAjuste(pvMat_CodProd_Orig(mi), pvMat_CodProd_Ajuste()) Then
            insMat_Ingresos_Det pvMat_CodProd_Orig(mi), pvMat_CodUbic_Orig(mi), 0, pvMat_CantidadBD_Orig(mi), pNroVale, pCodTipoVale, pCodUsuario, cmdTemp
            cmdTemp.Execute
         End If
      Next
   End If
 
   cmdTemp.ActiveConnection.CommitTrans
   cmdTemp.CommandType = 1
   Exit Sub
ErrorHandler:
   MsgBox "Se produjo un Error al intentar ajustar el Ingreso de materiales. Comunicarse con el Administrador" & vbCrLf & "Detalle del error: " & Err.Number & " " & Err.Description, vbCritical, "Error"
   cmdTemp.ActiveConnection.RollbackTrans
   cmdTemp.CommandType = 1
End Sub

Public Function estaEnVectorAjuste(ByVal pCodProd As String, ByRef pvMat_CodProd_Ajuste() As String) As Boolean
   Dim ret As Boolean
   Dim mi As Integer
   
   ret = False
   
   For mi = LBound(pvMat_CodProd_Ajuste) To UBound(pvMat_CodProd_Ajuste)
         If pCodProd = pvMat_CodProd_Ajuste(mi) Then
            ret = True
            mi = 9999
            
         End If
   Next
   estaEnVectorAjuste = ret
End Function

Public Function estaEnVectorPartesFinalizados(ByVal pParte As Double, ByRef pPartes_OT() As Double) As Boolean
   Dim ret As Boolean
   Dim mi As Integer
   
   ret = False
   
   For mi = LBound(pPartes_OT) To UBound(pPartes_OT)
         If pParte = pPartes_OT(mi) Then
            ret = True
            mi = 9999
         End If
   Next
   estaEnVectorPartesFinalizados = ret
End Function

Public Sub xGenerarComunicado(ByVal pNroComunicado As String, ByVal pCodRamal As String, ByRef pvParte_NroParte() As Integer, ByRef pvParte_CodActivo() As String, ByRef pvParte_DescActivo() As String, _
                           ByRef pvParte_CodProblema() As String, ByRef pvParte_DescProblema() As String, _
                           ByRef pvPartePrioridad() As String, pCodUsuario As String)
   
   On Error GoTo ErrorHandler
   Dim strSql1 As String
   Dim mi As Integer
   Dim mNroParte As String
   Dim mDescRamal As String
   Dim mErrMail As Integer
   Dim mListaDestinatarios As String
   Dim mTextoMail As String
   
   
   cmdTemp.CommandType = 1
   cmdTemp.CommandTimeout = 36000
   Set cmdTemp.ActiveConnection = DBConn
   cmdTemp.ActiveConnection.BeginTrans

   strSql1 = "INSERT INTO COM_Comunicados_H (NroComunicado,CodRamal,Fecha,CodUsuario)  VALUES ('" & pNroComunicado & "','" & pCodRamal & "', NOW(),'" & pCodUsuario & "');"
   cmdTemp.CommandText = strSql1
   cmdTemp.Execute

   mNroParte = ObtMaxParte
   mDescRamal = sCampoDescrip("COM_Ramales", "Codigo = '" & pCodRamal & "'", 1)
   mErrMail = 0
   
   For mi = LBound(pvParte_CodActivo) To UBound(pvParte_CodActivo)

      Dim pCodSuperv As String
      pCodSuperv = "srelev"

     strSql1 = "INSERT INTO Registros (Parte, FechaSolic, CodEdificio, Descripcion, Prioridad, CodSuperv, Estado, OpGen, Origen,SectorAire) VALUES (" & mi + (mNroParte + 1) & ", NOW(),'" & pvParte_DescActivo(mi) & "','" & pvParte_DescProblema(mi) & "','" & pvPartePrioridad(mi) & "','" & pCodSuperv & "','G','" & pCodUsuario & "', 'O',0);"
     cmdTemp.CommandText = strSql1
     cmdTemp.Execute

     strSql1 = "INSERT INTO COM_Comunicados_Det (NroComunicado, Parte, CodActivo, Desc_OtroActivo, CodActivoProblema, Desc_OtroActivoProblema) VALUES ('" & pNroComunicado & "'," & mi + (mNroParte + 1) & ",'" & pvParte_CodActivo(mi) & "','" & pvParte_DescActivo(mi) & "','" & pvParte_CodProblema(mi) & "','" & pvParte_DescProblema(mi) & "');"
     cmdTemp.CommandText = strSql1
     cmdTemp.Execute
     
     pvParte_NroParte(mi) = mi + (mNroParte + 1)

   Next

   cmdTemp.ActiveConnection.CommitTrans
   
   '----------------------SECCION ENVIO MAIL----------------------------------------------------------------------------------------
   mTextoMail = "COMUNICADO NRO: " & pNroComunicado & vbCrLf
   mTextoMail = mTextoMail & "RAMAL: " & mDescRamal & vbCrLf & vbCrLf
   For mi = LBound(pvParte_DescProblema) To UBound(pvParte_DescProblema)
      mTextoMail = mTextoMail & vbTab & "Parte " & pvParte_NroParte(mi) & " (Prioridad  " & pvPartePrioridad(mi) & "): " & pvParte_DescActivo(mi) & "  -  " & pvParte_DescProblema(mi) & "" & vbCrLf
   Next
   
   If mTextoMail <> "" Then
      mListaDestinatarios = ""
      Set Rs1 = oEjecutarSelect("SELECT DISTINCT Email FROM MailsxElectrico WHERE SectorAire = 0 and FechaBaja IS NULL ")
      If Not Rs1.EOF Then
          Do While Not Rs1.EOF
              mListaDestinatarios = mListaDestinatarios & Rs1!Email & ";"
              Rs1.MoveNext
          Loop
          If Not fEnviar_Mail_CDO("", mListaDestinatarios, "ausolmail@ausol.com.ar", " Nuevo Comunicado: " & pNroComunicado & "  -  Ramal " & mDescRamal, "Se ha generado un nuevo Comunicado de Mant. Eléctrico, según detalle:" & vbCrLf & vbCrLf & mTextoMail, "", "587", "ausolmail@ausol.com.ar", "sgedosmildiecisiete1$", True, False) Then
              mErrMail = mErrMail + 1
          End If
      End If
      If mErrMail = 0 Then
         MsgBox "Se ha generado el Comunicado " & pNroComunicado & " y enviado el mensaje correctamente a los responsables.", vbInformation, "Atención"
      Else
         MsgBox "Se ha generado el Comunicado  " & pNroComunicado & ", pero NO se ha enviado el correo correctamente", vbExclamation, "Atención"
      End If
   End If
   '----------------------FIN SECCION ENVIO MAIL----------------------------------------------------------------------------------------
   Exit Sub
ErrorHandler:
   cmdTemp.ActiveConnection.RollbackTrans
   MsgBox "Se produjo un Error al cargar el Comunicado. Comunicarse con el Administrador" & vbCrLf & "Detalle del error: " & Err.Number & " " & Err.Description, vbCritical, "Error"
End Sub

Public Sub xGenerarRelevamiento(ByVal pCodRamal As String, ByRef pvParte_NroParte() As Integer, ByRef pvParte_DescProblema() As String, ByRef pvPartePrioridad() As String, pCodUsuario As String)

   On Error GoTo ErrorHandler
   Dim strSql1 As String
   Dim mi As Integer
   Dim mNroParte As String
   Dim rs As ADODB.Recordset
   Dim mIdRele As Integer
   Dim mDescRamal As String
   Dim pCodSuperv As String
   Dim pruebaDiv As Integer
   Dim mErrMail As Integer
   Dim mListaDestinatarios As String
   Dim mTextoMail As String
      
   pCodSuperv = "srelev"
   
   mDescRamal = sCampoDescrip("COM_Ramales", "Codigo = '" & pCodRamal & "'", 1)
   
   cmdTemp.CommandType = adCmdStoredProc
   cmdTemp.CommandTimeout = 36000
   Set cmdTemp.ActiveConnection = DBConn
   cmdTemp.ActiveConnection.BeginTrans
   
   insRelevamiento_H pCodRamal, pCodUsuario, cmdTemp
      
   Set rs = cmdTemp.Execute
   mIdRele = rs!vIdRele
   rs.Close
     
   mNroParte = ObtMaxParte
   mErrMail = 0
   cmdTemp.CommandType = 1
   
   
   For mi = LBound(pvParte_DescProblema) To UBound(pvParte_DescProblema)

     strSql1 = "INSERT INTO Registros (Parte, FechaSolic, CodEdificio, Descripcion, Prioridad, CodSuperv, Estado, OpGen, Origen,SectorAire) VALUES (" & mi + (mNroParte + 1) & ", NOW(),'" & mDescRamal & "','" & pvParte_DescProblema(mi) & "','" & pvPartePrioridad(mi) & "','" & pCodSuperv & "','G','" & pCodUsuario & "', 'O',0);"
     cmdTemp.CommandText = strSql1
     cmdTemp.Execute

     strSql1 = "INSERT INTO REL_Relevamientos_Det (idRele, Parte, Desc_Problema) VALUES (" & mIdRele & "," & mi + (mNroParte + 1) & ",'" & pvParte_DescProblema(mi) & "');"
     cmdTemp.CommandText = strSql1
     cmdTemp.Execute
     
     pvParte_NroParte(mi) = mi + (mNroParte + 1)
     
   Next

   cmdTemp.ActiveConnection.CommitTrans
   
   mTextoMail = "RAMAL: " & mDescRamal & vbCrLf & vbCrLf
   For mi = LBound(pvParte_DescProblema) To UBound(pvParte_DescProblema)
      mTextoMail = mTextoMail & vbTab & "Parte " & pvParte_NroParte(mi) & " (Prioridad  " & pvPartePrioridad(mi) & "): " & pvParte_DescProblema(mi) & "" & vbCrLf
   Next
   
   If mTextoMail <> "" Then
      mListaDestinatarios = ""
      Set Rs1 = oEjecutarSelect("SELECT DISTINCT Email FROM MailsxElectrico WHERE SectorAire = 0 and FechaBaja IS NULL ")
      If Not Rs1.EOF Then
          Do While Not Rs1.EOF
              mListaDestinatarios = mListaDestinatarios & Rs1!Email & ";"
              Rs1.MoveNext
          Loop
          If Not fEnviar_Mail_CDO("", mListaDestinatarios, "ausolmail@ausol.com.ar", " Nuevos Relevamientos - Ramal " & mDescRamal, "Se han generado nuevos relevamientos de Mant. Eléctrico, según detalle:" & vbCrLf & vbCrLf & mTextoMail, "", "587", "ausolmail@ausol.com.ar", "sgedosmildiecisiete1$", True, False) Then
              mErrMail = mErrMail + 1
          End If
      End If
      If mErrMail = 0 Then
         MsgBox "Se ha generado el Relevamiento y enviado el mensaje correctamente a los responsables.", vbInformation, "Atención"
      Else
         MsgBox "Se ha generado el Relevamiento, pero NO se ha enviado el correo correctamente", vbExclamation, "Atención"
      End If
   End If
   Exit Sub
ErrorHandler:
   cmdTemp.ActiveConnection.RollbackTrans
   MsgBox "Se produjo un Error al cargar el Relevamiento. Comunicarse con el Administrador" & vbCrLf & "Detalle del error: " & Err.Number & " " & Err.Description, vbCritical, "Error"
End Sub

Public Sub xGenerarRelevamiento_Columnas(ByVal pCodRamal As String, ByRef pvParte_NroParte() As Integer, ByRef pvParte_CodActivo() As String, ByRef pvParte_DescActivo() As String, _
                           ByRef pvParte_CodProblema() As String, ByRef pvParte_DescProblema() As String, _
                           ByRef pvPartePrioridad() As String, pCodUsuario As String)
   
   On Error GoTo ErrorHandler
   Dim strSql1 As String
   Dim mi As Integer
   Dim mNroParte As String
   Dim rs As ADODB.Recordset
   Dim mIdRele As Integer
   Dim mDescRamal As String
   Dim pCodSuperv As String
   Dim pruebaDiv As Integer
   Dim mErrMail As Integer
   Dim mListaDestinatarios As String
   Dim mTextoMail As String
      
   pCodSuperv = "srelev"
   
   mDescRamal = sCampoDescrip("COM_Ramales", "Codigo = '" & pCodRamal & "'", 1)
   
   cmdTemp.CommandType = adCmdStoredProc
   cmdTemp.CommandTimeout = 36000
   Set cmdTemp.ActiveConnection = DBConn
   cmdTemp.ActiveConnection.BeginTrans
   
   insRelevamiento_H pCodRamal, pCodUsuario, cmdTemp
      
   Set rs = cmdTemp.Execute
   mIdRele = rs!vIdRele
   rs.Close
     
   mNroParte = ObtMaxParte
   mErrMail = 0
   cmdTemp.CommandType = 1
   
   
   For mi = LBound(pvParte_DescProblema) To UBound(pvParte_DescProblema)

     strSql1 = "INSERT INTO Registros (Parte, FechaSolic, CodEdificio, Descripcion, Prioridad, CodSuperv, Estado, OpGen, Origen,SectorAire) VALUES (" & mi + (mNroParte + 1) & ", NOW(),'" & pvParte_DescActivo(mi) & "','" & pvParte_DescProblema(mi) & "','" & pvPartePrioridad(mi) & "','" & pCodSuperv & "','G','" & pCodUsuario & "', 'O',0);"
     cmdTemp.CommandText = strSql1
     cmdTemp.Execute
     
     strSql1 = "INSERT INTO REL_Relevamientos_Det_Columnas (IdRele,Parte, CodActivo, Desc_OtroActivo, CodActivoProblema, Desc_OtroActivoProblema) VALUES (" & mIdRele & "," & mi + (mNroParte + 1) & ",'" & pvParte_CodActivo(mi) & "','" & pvParte_DescActivo(mi) & "','" & pvParte_CodProblema(mi) & "','" & pvParte_DescProblema(mi) & "');"
     cmdTemp.CommandText = strSql1
     cmdTemp.Execute
     
     pvParte_NroParte(mi) = mi + (mNroParte + 1)
     
   Next

   cmdTemp.ActiveConnection.CommitTrans
   
   mTextoMail = "RAMAL: " & mDescRamal & vbCrLf & vbCrLf
   For mi = LBound(pvParte_DescProblema) To UBound(pvParte_DescProblema)
      mTextoMail = mTextoMail & vbTab & "Parte " & pvParte_NroParte(mi) & " (Prioridad  " & pvPartePrioridad(mi) & "): " & pvParte_DescActivo(mi) & "   -   " & pvParte_DescProblema(mi) & "" & vbCrLf
   Next
   
   If mTextoMail <> "" Then
      mListaDestinatarios = ""
      Set Rs1 = oEjecutarSelect("SELECT DISTINCT Email FROM MailsxElectrico WHERE SectorAire = 0 and FechaBaja IS NULL ")
      If Not Rs1.EOF Then
          Do While Not Rs1.EOF
              mListaDestinatarios = mListaDestinatarios & Rs1!Email & ";"
              Rs1.MoveNext
          Loop
          If Not fEnviar_Mail_CDO("", mListaDestinatarios, "ausolmail@ausol.com.ar", " Nuevos Relevamientos - Ramal " & mDescRamal, "Se han generado nuevos relevamientos de Mant. Eléctrico, según detalle:" & vbCrLf & vbCrLf & mTextoMail, "", "587", "ausolmail@ausol.com.ar", "sgedosmildiecisiete1$", True, False) Then
              mErrMail = mErrMail + 1
          End If
      End If
      If mErrMail = 0 Then
         MsgBox "Se ha generado el Relevamiento y enviado el mensaje correctamente a los responsables.", vbInformation, "Atención"
      Else
         MsgBox "Se ha generado el Relevamiento, pero NO se ha enviado el correo correctamente", vbExclamation, "Atención"
      End If
   End If
   Exit Sub
ErrorHandler:
   cmdTemp.ActiveConnection.RollbackTrans
   MsgBox "Se produjo un Error al cargar el Relevamiento. Comunicarse con el Administrador" & vbCrLf & "Detalle del error: " & Err.Number & " " & Err.Description, vbCritical, "Error"
End Sub

















'---------------------------------------
'Sp: getConsumoMatXidOTyUbicacion
'---------------------------------------
Public Function getConsumoMatXidOTyUbicacion(ByVal IdOT As String, ByVal CodUbicacion As String) As ADODB.Recordset
   Dim paramCodAlfa As parametroSql
   Dim aParamSql(1) As parametroSql

   paramCodAlfa.Nombre = "vIdOT"
   paramCodAlfa.Tipo = adInteger
   paramCodAlfa.Direccion = adParamInput
   paramCodAlfa.Tamanio = 4
   paramCodAlfa.Valor = IdOT
   aParamSql(0) = paramCodAlfa
  
   paramCodAlfa.Nombre = "vCodUbicacion"
   paramCodAlfa.Tipo = adChar
   paramCodAlfa.Direccion = adParamInput
   paramCodAlfa.Tamanio = 4
   paramCodAlfa.Valor = CodUbicacion
   aParamSql(1) = paramCodAlfa
  
   Set getConsumoMatXidOTyUbicacion = oStoreProcedureWithParam("getConsumoMatXidOTyUbicacion", aParamSql)
End Function

'---------------------------------------
'Sp: getConsumoAbastecimientoXidOTyUbicacion
'---------------------------------------
Public Function getConsumoAbastecimientoXidOTyUbicacion(ByVal IdOT As String, ByVal CodUbicacion As String) As ADODB.Recordset
   Dim paramCodAlfa As parametroSql
   Dim aParamSql(1) As parametroSql

   paramCodAlfa.Nombre = "vIdOT"
   paramCodAlfa.Tipo = adInteger
   paramCodAlfa.Direccion = adParamInput
   paramCodAlfa.Tamanio = 4
   paramCodAlfa.Valor = IdOT
   aParamSql(0) = paramCodAlfa
  
   paramCodAlfa.Nombre = "vCodUbicacion"
   paramCodAlfa.Tipo = adChar
   paramCodAlfa.Direccion = adParamInput
   paramCodAlfa.Tamanio = 4
   paramCodAlfa.Valor = CodUbicacion
   aParamSql(1) = paramCodAlfa
  
   Set getConsumoAbastecimientoXidOTyUbicacion = oStoreProcedureWithParam("getConsumoAbastecimientoXidOTyUbicacion", aParamSql)
End Function

'---------------------------------------
'Sp: getIngresoXnroVale_CodTipoVale_Ubicacion
'---------------------------------------
Public Function getIngresoXnroVale_CodTipoVale_Ubicacion(ByVal NroVale As String, ByVal CodTipoVale As String, ByVal CodUbicacion As String) As ADODB.Recordset
   Dim paramCodAlfa As parametroSql
   Dim aParamSql(2) As parametroSql

   With paramCodAlfa
       .Nombre = "vNroVale"
       .Tipo = adVarChar
       .Direccion = adParamInput
       .Tamanio = 9
       .Valor = NroVale
       aParamSql(0) = paramCodAlfa
      
       .Nombre = "vCodTipoVale"
       .Tipo = adChar
       .Direccion = adParamInput
       .Tamanio = 1
       .Valor = CodTipoVale
       aParamSql(1) = paramCodAlfa
      
       .Nombre = "vCodUbicacion"
       .Tipo = adChar
       .Direccion = adParamInput
       .Tamanio = 4
       .Valor = CodUbicacion
       aParamSql(2) = paramCodAlfa
End With
   Set getIngresoXnroVale_CodTipoVale_Ubicacion = oStoreProcedureWithParam("getIngresoXnroVale_CodTipoVale_Ubicacion", aParamSql)
End Function




Public Sub xinsRepo(ByVal pCodUsuario As String, ByVal pVeh_Origen As String, ByVal pVeh_Destino As String, ByRef pvCodProd() As String, ByRef pvStVeh_Inv() As Double, _
                         ByRef pvStVeh_Sist() As Double, ByRef pvCant_Repuesta() As Double, ByRef pvStDepo() As Double, ByRef pvRepo_RepoSugerida_xMR() As Double)
   
  'Los parametros pVeh_Origen y pVeh_Destino hacen referencia a Ubicaciones de Inventario las cuales estan relacionadas con Vehiculos.
   
   On Error GoTo ErrorHandler
   Dim strSql1 As String
   Dim mi As Integer
   Dim rs As ADODB.Recordset
   Dim mIdRepo As Integer
   
   If pvCodProd(0) <> "000000" Then
   
      cmdTemp.CommandTimeout = 36000
      Set cmdTemp.ActiveConnection = DBConn
      cmdTemp.ActiveConnection.BeginTrans
      cmdTemp.CommandType = adCmdStoredProc
  
      'insReposicion_H(ByVal Veh_Origen As String, ByVal Veh_Destino As String, CodUsuario As String, ByRef cmd As ADODB.Command)
      insReposicion_H pVeh_Origen, pVeh_Destino, pCodUsuario, cmdTemp
      Set rs = cmdTemp.Execute
      mIdRepo = rs!vIdRepo
      rs.Close
     
   
      For mi = LBound(pvCodProd) To UBound(pvCodProd)
         'insReposiciones_Det retOT, pvPartes_OT(mi), FechaOT, pCodUsuario, cmdTemp
         'PROCEDURE `insReposiciones_Det`(IN vIdRepo INTEGER, IN vCodProducto CHAR(6),IN vCodUbicacion_Origen CHAR(4), IN vCodUbicacion_Destino CHAR(4), IN vStVeh_Inv Double, IN vStVeh_Sist Double,IN vCant_Repuesta Double, IN vStDepo Double, IN vCodUsuario VARCHAR(25))
         insReposiciones_Det mIdRepo, pvCodProd(mi), pVeh_Origen, pVeh_Destino, pvStVeh_Inv(mi), pvStVeh_Sist(mi), pvCant_Repuesta(mi), pvStDepo(mi), pvRepo_RepoSugerida_xMR(mi), pCodUsuario, cmdTemp
         cmdTemp.Execute
      Next
   
      cmdTemp.ActiveConnection.CommitTrans
      cmdTemp.CommandType = 1
      Exit Sub
   End If
      

ErrorHandler:

   MsgBox "Se produjo un Error al realizar la Reposición. Comunicarse con el Administrador" & vbCrLf & "Detalle del error: " & Err.Number & " " & Err.Description, vbCritical, "Error"
   cmdTemp.ActiveConnection.RollbackTrans
   cmdTemp.CommandType = 1
End Sub


'---------------------------------------
'Sp: insRelevamiento_H
'---------------------------------------
Public Sub insRelevamiento_H(ByVal CodRamal As String, CodUsuario As String, ByRef cmd As ADODB.Command)
   Dim parametroSql As parametroSql
   Dim aParamSql(1) As parametroSql
   
   With parametroSql
   
      .Nombre = "vCodRamal"
      .Tipo = adVarChar
      .Direccion = adParamInput
      .Tamanio = 2
      .Valor = CodRamal
      aParamSql(0) = parametroSql
      
      .Nombre = "vCodUsuario"
      .Tipo = adVarChar
      .Direccion = adParamInput
      .Tamanio = 25
      .Valor = CodUsuario
      aParamSql(1) = parametroSql
      
   End With
   xStoreProcedureWithParam_Transaccional "insRelevamiento_H", aParamSql, cmd
End Sub








'---------------------------------------
'Sp: insReposicion_H
'---------------------------------------
Public Sub insReposicion_H(ByVal Veh_Origen As String, ByVal Veh_Destino As String, CodUsuario As String, ByRef cmd As ADODB.Command)
   Dim parametroSql As parametroSql
   Dim aParamSql(2) As parametroSql
   
   With parametroSql
   
      .Nombre = "vVeh_Origen"
      .Tipo = adChar
      .Direccion = adParamInput
      .Tamanio = 4
      .Valor = Veh_Origen
      aParamSql(0) = parametroSql
      
      .Nombre = "vVeh_Destino"
      .Tipo = adChar
      .Direccion = adParamInput
      .Tamanio = 4
      .Valor = Veh_Destino
      aParamSql(1) = parametroSql
   
      .Nombre = "vCodUsuario"
      .Tipo = adVarChar
      .Direccion = adParamInput
      .Tamanio = 25
      .Valor = CodUsuario
      aParamSql(2) = parametroSql
      
   End With
   xStoreProcedureWithParam_Transaccional "insReposicion_H", aParamSql, cmd
End Sub

Public Sub insReposiciones_Det(ByVal IdRepo As Integer, ByVal CodProducto As String, ByVal CodUbicacion_Origen As String, _
                                 ByVal CodUbicacion_Destino As String, ByVal StVeh_Inv As Double, ByVal StVeh_Sist As Double, _
                                 ByVal Cant_Repuesta As Double, ByVal StDepo As Double, ByVal RepoSugerida_xMR As Double, ByVal CodUsuario As String, ByRef cmd As ADODB.Command)
                                 
   Dim parametroSql As parametroSql
   Dim aParamSql(9) As parametroSql
   
   With parametroSql
      
      .Nombre = "vIdRepo"
      .Tipo = adInteger
      .Direccion = adParamInput
      .Tamanio = 4
      .Valor = IdRepo
      aParamSql(0) = parametroSql
      
      .Nombre = "vCodProducto"
      .Tipo = adChar
      .Direccion = adParamInput
      .Tamanio = 6
      .Valor = CodProducto
      aParamSql(1) = parametroSql
      
      .Nombre = "vCodUbicacion_Origen"
      .Tipo = adChar
      .Direccion = adParamInput
      .Tamanio = 4
      .Valor = CodUbicacion_Origen
      aParamSql(2) = parametroSql
      
      .Nombre = "vCodUbicacion_Destino"
      .Tipo = adChar
      .Direccion = adParamInput
      .Tamanio = 4
      .Valor = CodUbicacion_Destino
      aParamSql(3) = parametroSql
      
      .Nombre = "vStVeh_Inv"
      .Tipo = adDouble
      .Direccion = adParamInput
      .Tamanio = 8
      .Valor = StVeh_Inv
      aParamSql(4) = parametroSql
      
      .Nombre = "vStVeh_Sist"
      .Tipo = adDouble
      .Direccion = adParamInput
      .Tamanio = 8
      .Valor = StVeh_Sist
      aParamSql(5) = parametroSql
     
      .Nombre = "vCant_Repuesta"
      .Tipo = adDouble
      .Direccion = adParamInput
      .Tamanio = 8
      .Valor = Cant_Repuesta
      aParamSql(6) = parametroSql
      
      .Nombre = "vStDepo"
      .Tipo = adDouble
      .Direccion = adParamInput
      .Tamanio = 8
      .Valor = StDepo
      aParamSql(7) = parametroSql
      
      .Nombre = "vRepoSugerida_xMR"
      .Tipo = adDouble
      .Direccion = adParamInput
      .Tamanio = 8
      .Valor = RepoSugerida_xMR
      aParamSql(8) = parametroSql
            
      .Nombre = "vCodUsuario"
      .Tipo = adVarChar
      .Direccion = adParamInput
      .Tamanio = 25
      .Valor = CodUsuario
      aParamSql(9) = parametroSql
      
   End With
   xStoreProcedureWithParam_Transaccional "insReposiciones_Det", aParamSql, cmd
 End Sub

Public Sub xCancelarPartes(ByVal pMotivo As String, ByRef pvPartes_Cancel() As Double, ByRef vParte_CodEdificio() As String, ByRef pvParte_Descripcion() As String, ByRef pvParte_SecAire() As Integer, ByRef pvParte_OpGen() As String, ByRef pvParte_CodSuperv() As String, ByVal pCodUsuario As String)
   On Error GoTo ErrorHandler
   Dim rs As ADODB.Recordset
   Dim mi As Integer
   Dim mErrMail As Integer
   Dim mListaDestinatarios As String
   Dim mTextoMail As String
   Dim mTextoMailOperac As String
   Dim mIdCancel As Integer

   cmdTemp.CommandTimeout = 36000
   Set cmdTemp.ActiveConnection = DBConn
   cmdTemp.ActiveConnection.BeginTrans
   cmdTemp.CommandType = adCmdStoredProc
  
   insCancelaciones_H pMotivo, pCodUsuario, cmdTemp

   Set rs = cmdTemp.Execute
   mIdCancel = rs!vIdCancel
   rs.Close

   If pvPartes_Cancel(0) <> 0 Then
      For mi = LBound(pvPartes_Cancel) To UBound(pvPartes_Cancel)
           insCancelaciones_Det mIdCancel, pvPartes_Cancel(mi), cmdTemp
            cmdTemp.Execute
      Next
   End If

   cmdTemp.ActiveConnection.CommitTrans
   cmdTemp.CommandType = 1
   
   If pvPartes_Cancel(0) <> 0 Then
      For mi = LBound(pvPartes_Cancel) To UBound(pvPartes_Cancel)
         mTextoMail = ""
         mListaDestinatarios = ""
         'Si es parte de Operaciones => agrego a la lista de destinatarios a los Superv de Estacion que correspondan
         If esParteOperaciones(pvPartes_Cancel(mi)) Then
            Set Rs1 = oEjecutarSelect("SELECT DISTINCT Email FROM MailsxSuperv WHERE CodSuperv = '" & pvParte_CodSuperv(mi) & "' And FechaBaja IS NULL ")
            If Not Rs1.EOF Then
               Do While Not Rs1.EOF
                  mListaDestinatarios = mListaDestinatarios & Rs1!Email & ";"
                  Rs1.MoveNext
               Loop
            End If
         End If
         'Agrego a los supervisores electricos a la lista de destinatarios
         Set Rs1 = oEjecutarSelect("SELECT DISTINCT Email FROM MailsxElectrico WHERE SectorAire = " & pvParte_SecAire(mi) & " and FechaBaja IS NULL ")
         If Not Rs1.EOF Then
            Do While Not Rs1.EOF
               mListaDestinatarios = mListaDestinatarios & Rs1!Email & ";"
                  Rs1.MoveNext
               Loop
         End If
          
         mTextoMail = "A continuación se detalla el parte cancelado y el motivo de su cancelación: " & vbCrLf & vbCrLf
         mTextoMail = mTextoMail & vbTab & "Nro Parte: " & pvPartes_Cancel(mi) & vbCrLf
         mTextoMail = mTextoMail & vbTab & "Descripción: " & vParte_CodEdificio(mi) & " - " & pvParte_Descripcion(mi) & vbCrLf
         mTextoMail = mTextoMail & vbTab & "Parte generado por: " & pvParte_OpGen(mi) & vbCrLf
         mTextoMail = mTextoMail & vbTab & "Motivo de cancelación: " & pMotivo
          
         If Not fEnviar_Mail_CDO("", mListaDestinatarios, "ausolmail@ausol.com.ar", "Cancelación de Parte: " & pvPartes_Cancel(mi), mTextoMail, "", "587", "ausolmail@ausol.com.ar", "sgedosmildiecisiete1$", True, False) Then
            mErrMail = mErrMail + 1
         End If
      Next
   End If
   
   If mErrMail = 0 Then
      MsgBox "Se han cancelado exitosamente los partes seleccionados y enviado los mensajes correctamente a los responsables.", vbInformation, "Atención"
   Else
      MsgBox "Se han cancelado exitosamente los partes seleccionados, pero NO se ha enviado correctamente algún mensaje", vbExclamation, "Atención"
   End If
   Exit Sub
ErrorHandler:
   MsgBox "Se produjo un Error al intentar cancelar los partes. Comunicarse con el Administrador" & vbCrLf & "Detalle del error: " & Err.Number & " " & Err.Description, vbCritical, "Error"
   cmdTemp.ActiveConnection.RollbackTrans
   cmdTemp.CommandType = 1
End Sub

'---------------------------------------
'cancelarPartes_H
'---------------------------------------
Public Sub insCancelaciones_H(ByVal Motivo As String, ByVal CodUsuario As String, ByRef cmd As ADODB.Command)
   Dim parametroSql As parametroSql
   Dim aParamSql(1) As parametroSql
   
   parametroSql.Nombre = "vMotivo"
   parametroSql.Tipo = adVarChar
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 90
   parametroSql.Valor = Motivo
   aParamSql(0) = parametroSql
    
   parametroSql.Nombre = "vCodUsuario"
   parametroSql.Tipo = adVarChar
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 25
   parametroSql.Valor = CodUsuario
   aParamSql(1) = parametroSql

   xStoreProcedureWithParam_Transaccional "insCancelaciones_H", aParamSql, cmd
End Sub

'---------------------------------------
'insCancelaciones_Det
'---------------------------------------
Public Sub insCancelaciones_Det(IdCancel As Integer, Parte As Double, ByRef cmd As ADODB.Command)
   Dim parametroSql As parametroSql
   Dim aParamSql(1) As parametroSql
   
   parametroSql.Nombre = "vIdCancel"
   parametroSql.Tipo = adInteger
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 4
   parametroSql.Valor = IdCancel
   aParamSql(0) = parametroSql
   
   parametroSql.Nombre = "vParte"
   parametroSql.Tipo = adDouble
   parametroSql.Direccion = adParamInput
   parametroSql.Tamanio = 8
   parametroSql.Valor = Parte
   aParamSql(1) = parametroSql
   
   xStoreProcedureWithParam_Transaccional "insCancelaciones_Det", aParamSql, cmd
End Sub


'---------------------------------------
'Sp: getRPT_Partes
'---------------------------------------
Public Function getRPT_Partes(ByVal FecInicio As Date, ByVal FecFin As Date) As ADODB.Recordset
   Dim paramCodAlfa As parametroSql
   Dim aParamSql(1) As parametroSql

   paramCodAlfa.Nombre = "vFecInicio"
   paramCodAlfa.Tipo = adDBTimeStamp
   paramCodAlfa.Direccion = adParamInput
   paramCodAlfa.Tamanio = 8
   paramCodAlfa.Valor = FecInicio
   aParamSql(0) = paramCodAlfa
      
   paramCodAlfa.Nombre = "vFecFin"
   paramCodAlfa.Tipo = adDBTimeStamp
   paramCodAlfa.Direccion = adParamInput
   paramCodAlfa.Tamanio = 8
   paramCodAlfa.Valor = FecFin
   aParamSql(1) = paramCodAlfa
  
   Set getRPT_Partes = oStoreProcedureWithParam("getRPT_Partes", aParamSql)
End Function

'---------------------------------------
'Sp: getRPT_OTs
'---------------------------------------
Public Function getRPT_OTs(ByVal FecInicio As Date, ByVal FecFin As Date) As ADODB.Recordset
   Dim paramCodAlfa As parametroSql
   Dim aParamSql(1) As parametroSql

   paramCodAlfa.Nombre = "vFecInicio"
   paramCodAlfa.Tipo = adDBTimeStamp
   paramCodAlfa.Direccion = adParamInput
   paramCodAlfa.Tamanio = 8
   paramCodAlfa.Valor = FecInicio
   aParamSql(0) = paramCodAlfa
      
   paramCodAlfa.Nombre = "vFecFin"
   paramCodAlfa.Tipo = adDBTimeStamp
   paramCodAlfa.Direccion = adParamInput
   paramCodAlfa.Tamanio = 8
   paramCodAlfa.Valor = FecFin
   aParamSql(1) = paramCodAlfa
  
   Set getRPT_OTs = oStoreProcedureWithParam("getRPT_OTs", aParamSql)
End Function

'---------------------------------------
'Sp: getRPT_Comunicados
'---------------------------------------
Public Function getRPT_Comunicados(ByVal FecInicio As Date, ByVal FecFin As Date, Estado As String) As ADODB.Recordset
   Dim paramCodAlfa As parametroSql
   Dim aParamSql(2) As parametroSql

   paramCodAlfa.Nombre = "vFecInicio"
   paramCodAlfa.Tipo = adDBTimeStamp
   paramCodAlfa.Direccion = adParamInput
   paramCodAlfa.Tamanio = 8
   paramCodAlfa.Valor = FecInicio
   aParamSql(0) = paramCodAlfa
      
   paramCodAlfa.Nombre = "vFecFin"
   paramCodAlfa.Tipo = adDBTimeStamp
   paramCodAlfa.Direccion = adParamInput
   paramCodAlfa.Tamanio = 8
   paramCodAlfa.Valor = FecFin
   aParamSql(1) = paramCodAlfa
  
   paramCodAlfa.Nombre = "vEstado"
   paramCodAlfa.Tipo = adVarChar
   paramCodAlfa.Direccion = adParamInput
   paramCodAlfa.Tamanio = 7
   paramCodAlfa.Valor = Estado
   aParamSql(2) = paramCodAlfa
  
   Set getRPT_Comunicados = oStoreProcedureWithParam("getRPT_Comunicados", aParamSql)
End Function


'---------------------------------------
'Sp: getRPT_PartesLuminarias
'---------------------------------------
Public Function getRPT_PartesLuminarias(ByVal FecInicio As Date, ByVal FecFin As Date) As ADODB.Recordset
   Dim paramCodAlfa As parametroSql
   Dim aParamSql(1) As parametroSql

   paramCodAlfa.Nombre = "vFecInicio"
   paramCodAlfa.Tipo = adDBTimeStamp
   paramCodAlfa.Direccion = adParamInput
   paramCodAlfa.Tamanio = 8
   paramCodAlfa.Valor = FecInicio
   aParamSql(0) = paramCodAlfa
      
   paramCodAlfa.Nombre = "vFecFin"
   paramCodAlfa.Tipo = adDBTimeStamp
   paramCodAlfa.Direccion = adParamInput
   paramCodAlfa.Tamanio = 8
   paramCodAlfa.Valor = FecFin
   aParamSql(1) = paramCodAlfa
  
   Set getRPT_PartesLuminarias = oStoreProcedureWithParam("getRPT_PartesLuminarias", aParamSql)
End Function

'---------------------------------------
'Sp: getRPT_Hist_Intervencion_Columna
'---------------------------------------
Public Function getRPT_Hist_Intervencion_Columna(ByVal CodActivo As Integer) As ADODB.Recordset
   Dim paramSQL As parametroSql
   Dim aParamSql(0) As parametroSql

   paramSQL.Nombre = "vCodActivo"
   paramSQL.Tipo = adInteger
   paramSQL.Direccion = adParamInput
   paramSQL.Tamanio = 4
   paramSQL.Valor = CodActivo
   aParamSql(0) = paramSQL
      
   Set getRPT_Hist_Intervencion_Columna = oStoreProcedureWithParam("getRPT_Hist_Intervencion_Columna", aParamSql)
End Function


Friend Sub xStoreProcedureWithParam_Transaccional(ByVal pNameStore As String, ByRef pArrayParamSql() As parametroSql, ByRef cmd As ADODB.Command)
  'Dim cmd As New ADODB.Command
   Dim i As Integer
   strSql1 = pNameStore
   cmd.CommandText = strSql1
   'sAsignCmd_SP cmd
   
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

