VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clPeek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Private DBConn As New ADODB.Connection
Private cmdTemp As New ADODB.Command
Private mObj As New clConnSistemas
Private Rs1 As New ADODB.Recordset
Private mDOS As New FileSystemObject
Private VarTexto As String
Private mErrorMsg As Boolean
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
   Set DBConn = mObj.oAbreConexion("5", mIPServer, "Peek", "ssvv", "ssvv")
  'Set DBConn = mObj.oAbreConexion("4", "127.0.0.1", "peek", "root", "clave")
   mErrorMsg = False
End Sub

Private Sub Class_Terminate()
   DBConn.Close
   Set DBConn = Nothing
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

Public Sub xStoreProcedure(ByVal pNameStore As String)
   Dim cmd As New ADODB.Command
   strSql1 = pNameStore
   sAsignCmd_SP cmd
   On Error Resume Next
   cmd.Execute
   Set cmd = Nothing
End Sub

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
   On Error Resume Next
   cmd.Execute
   Set cmd = Nothing
End Sub


Public Function oPeekCarril_Sent(ByVal pCodCont As String, ByVal pSent As String) As ADODB.Recordset
   strSql1 = "SELECT * FROM Peek_Carril where codcont='" & pCodCont & "'"
   If pSent <> "" Then
      strSql1 = strSql1 & " and sentido in ('" & pSent & "')"
   End If
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oPeekCarril_Sent = Rs1
End Function

Public Function oSumSent(ByVal pCodCont As String, ByVal pYear As Integer, ByVal pMes As Integer, ByVal pCarrilesA As String, ByVal pCarrilesD As String, ByVal pOrder As String) As ADODB.Recordset
   strSql1 = "select " & pOrder & ", Sum(" & pCarrilesA & ") as A, sum(" & pCarrilesD & ") AS D From " & pCodCont & " " _
      & " where month(fecha)=" & pMes & " and year(Fecha)=" & pYear & " group by " & pOrder
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oSumSent = Rs1
End Function

Public Function oSumSentCol(ByVal pCodCont As String, ByVal pYear As Integer, ByVal pMes As Integer, _
                            ByVal pCarrilesA As String, ByVal pCarrilesD As String, _
                            ByVal pCarrilesColA As String, ByVal pCarrilesColD As String) As ADODB.Recordset
strSql1 = "select Fecha, Sum(" & pCarrilesA & ") as A, sum(" & pCarrilesD & ") AS D, sum(" & pCarrilesColA & ") AS ColA, " _
      & " sum(" & pCarrilesColD & ") AS ColD From " & pCodCont & " " _
      & " where month(fecha)=" & pMes & " and year(Fecha)=" & pYear & " group by fecha"
sAsignCmd
Set Rs1 = cmdTemp.Execute
Set oSumSentCol = Rs1
End Function

Public Function oSumCarriles(ByVal pCodCont As String, ByVal pYear As Integer, ByVal pMes As Integer, _
                            ByVal pCarriles As String, ByVal pOrder As String, Optional pClase As String) As ADODB.Recordset
Dim mClase As String
   mClase = ""
   If pClase <> "" Then mClase = pClase
   strSql1 = "select " & pOrder & ", " & pCarriles & " From " & pCodCont & " where month(fecha)=" & pMes & " and year(Fecha)=" & pYear & " " & mClase & " group by " & pOrder
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oSumCarriles = Rs1
End Function

Public Function oErrorPeek(ByVal pCodCont As String, ByVal pYear As Integer, ByVal pMes As Integer) As ADODB.Recordset
Dim mWhere As String
   mWhere = ""
   If Trim(pCodCont) <> "" Then
      mWhere = " AND Peek='" & pCodCont & "'"
   End If
   strSql1 = "SELECT DISTINCT Peek,Fecha,Descripcion,Archivo FROM Peek_Error INNER JOIN Errores ON Peek_Error.CodError=Errores.Codigo " _
         & " where month(fecha)=" & pMes & " and year(Fecha)=" & pYear & mWhere
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oErrorPeek = Rs1
End Function

Public Function iMaxCarril(ByVal pCodCont As String) As Integer
Dim mMax As Integer
   strSql1 = "SELECT carril  FROM Peek_Carril WHERE CodCont = '" & pCodCont & "'"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   mMax = 0
   Do While Not Rs1.EOF
      If Val(Rs1!carril) > mMax Then
         mMax = Rs1!carril
      End If
      Rs1.MoveNext
   Loop
   iMaxCarril = mMax
End Function

Public Function iCountFallas(ByVal pFecha1 As String, ByVal pFecha2 As String, ByVal pCont As String, ByVal pEstado As String) As Integer
Dim mMax As Integer
   iCountFallas = 0
   strSql1 = "select count(distinct(date_format(fecha, '%Y-%m-%d'))) as total From fallasred where " _
      & " fecha between '" & Format(pFecha1, "yyyy-mm-dd") & " 00:00:00' and '" & Format(pFecha2, "yyyy-mm-dd") & " 23:59:59' " _
      & " and codcontador='" & pCont & "' and estado='" & pEstado & "'"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   If Not Rs1.EOF Then
      iCountFallas = Rs1!Total
   End If
End Function

Public Function oFallas() As ADODB.Recordset
   strSql1 = "select * from fallasred order by 2 desc, 1"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oFallas = Rs1
End Function

Public Function oFallasMes(ByVal pCont As String, ByVal pFecha1 As String, ByVal pFecha2 As String) As ADODB.Recordset
   strSql1 = "select fecha, estado from fallasred where codcontador='" & pCont & "' and fecha between '" & Format(pFecha1, "yyyy-mm-dd") & " 00:00:00' and '" & Format(pFecha2, "yyyy-mm-dd") & " 23:59:59'"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oFallasMes = Rs1
End Function

Public Function bExistDato(ByVal pCont As String, ByVal pFecha As String) As Boolean
   strSql1 = "select fecha from " & pCont & " where fecha='" & pFecha & "'"
   sAsignCmd
   On Error Resume Next
   Set Rs1 = cmdTemp.Execute
   If Err.Description = "" Then
      bExistDato = Not Rs1.EOF
   Else
      bExistDato = True
   End If
   Rs1.Close
End Function

Public Function PasarDatos(mDir As String)
 Dim mBand As String
 Dim mFecha As String
 Dim mHora As String
 Dim mi As Integer
 Dim mj As Integer
 Dim mLane As Integer
 Dim mFile As String
 Dim mError As Boolean
 Dim mVeloc As Integer
 Dim mLong As Integer
 Dim mWhere As String
 Dim mValue As String
 Dim mConErrores As Boolean
 Dim mCtrlError As Boolean
 Dim mExistFecha As Boolean
 
 mError = True
 mCtrlError = False
 If MsgBox("¿Está Seguro de Pasar Datos a la Base", vbYesNo, "Peek - Atención!!") = vbYes Then
    sMsgEspere MDI, "Procesando datos... aguarde un momento.", True
    mFile = Dir(App.Path & "/PEEK/Datos/" & mDir & "/*.prn")
    While mFile <> ""
       Open App.Path & "/PEEK/Datos/" & mDir & "/" & Trim(mFile) For Input As #1
       mError = False
       Line Input #1, VarTexto
       mFecha = "20" & Mid(VarTexto, 39, 2) & "-" & Mid(VarTexto, 35, 2) & "-" & Mid(VarTexto, 37, 2) & ""
       If Not IsDate(mFecha) Then
          mFecha = "" & Year(Now) & "-" & Mid(mFile, 2, 2) & "-" & Mid(mFile, 4, 2) & ""
       End If
       mExistFecha = bExistDato(mDir, mFecha)
       Line Input #1, VarTexto
       Line Input #1, VarTexto
       Line Input #1, VarTexto
       If Left(VarTexto, 2) = "01" Then 'Configuración de Carriles
          mLane = Mid(VarTexto, 4, 2) 'Cantidad de Carriles
        Else
          mError = True
          DBConn.Execute ("INSERT INTO Peek_Error (Peek,Fecha,CodError,Archivo) VALUES ('" & mDir & "','" & mFecha & "','02','" & mFile & ".prn')")
       End If
       Rs1.Open "SELECT COUNT(Carril) AS mCarril FROM Peek_Carril WHERE CodCont='" & UCase(mDir) & "'", DBConn
       If Mid(VarTexto, 4, 2) <> "" & Format(Rs1!mCarril, "0#") & "" Then
          mError = True
          DBConn.Execute ("INSERT INTO Peek_Error (Peek,Fecha,CodError,Archivo) VALUES ('" & mDir & "','" & mFecha & "','03','" & mFile & ".prn')")
       End If
       Rs1.Close
       'Contro de Errores en Longitud
       Line Input #1, VarTexto
   If mCtrlError Then
       If Left(VarTexto, 2) <> "06" Then 'Si el Estudio es Longitud
          DBConn.Execute ("INSERT INTO Peek_Error (Peek,Fecha,CodError,Archivo) VALUES ('" & mDir & "','" & mFecha & "','02','" & mFile & ".prn')")
          mError = True
       End If
       If Mid(VarTexto, 4, 2) <> "08" Then
          DBConn.Execute ("INSERT INTO Peek_Error (Peek,Fecha,CodError,Archivo) VALUES ('" & mDir & "','" & mFecha & "','04','" & mFile & ".prn')")
          mError = True
       End If
   End If
       'Contro de Errores en Velocidades
       Line Input #1, VarTexto
   If mCtrlError Then
       If Left(VarTexto, 2) <> "03" Then 'Si el Estudio es Longitud
          DBConn.Execute ("INSERT INTO Peek_Error (Peek,Fecha,CodError,Archivo) VALUES ('" & mDir & "','" & mFecha & "','02','" & mFile & ".prn')")
          mError = True
       End If
       If Mid(VarTexto, 4, 2) <> "07" Then
          DBConn.Execute ("INSERT INTO Peek_Error (Peek,Fecha,CodError,Archivo) VALUES ('" & mDir & "','" & mFecha & "','06','" & mFile & ".prn')")
          mError = True
       Else
          If Not (Mid(VarTexto, 14, 4) = "0031" And Mid(VarTexto, 19, 4) = "0037" And Mid(VarTexto, 24, 4) = "0050" And Mid(VarTexto, 29, 4) = "0062" And Mid(VarTexto, 34, 4) = "0068" And Mid(VarTexto, 39, 4) = "0081" And Mid(VarTexto, 44, 4) = "0147") Then
             DBConn.Execute ("INSERT INTO Peek_Error (Peek,Fecha,CodError,Archivo) VALUES ('" & mDir & "','" & mFecha & "','07','" & mFile & ".prn')")
             mError = True
          End If
       End If
   End If
       '----  Comienzo de Datos ---------------------
       If Not mError Then
          mValue = ""
          For mi = 1 To mLane
             mValue = mValue & ",Carril" & mi
          Next
          Line Input #1, VarTexto
          mHora = Mid(VarTexto, 9, 2) & ":" & Mid(VarTexto, 11, 2)
          While Not EOF(1)
             mWhere = ""
             mHora = Mid(VarTexto, 9, 2) & ":" & Mid(VarTexto, 11, 2)
             mVeloc = Left(VarTexto, 2)
             mLong = Mid(VarTexto, 4, 2)
             mj = 14
             For mi = 0 To mLane - 1
                mWhere = mWhere & Mid(VarTexto, mj, 4) & ","
                mj = mj + 5
             Next
             mWhere = Mid(mWhere, 1, Len(mWhere) - 1)
             If Not mExistFecha Then
               DBConn.Execute "INSERT INTO " & mDir & " (Fecha,Hora,Veloc,Longi" & mValue & ") VALUES ('" & mFecha & "','" & mHora & "','" & mVeloc & "','" & mLong & "'," & mWhere & ")"
             End If
             Line Input #1, VarTexto
          Wend
          Close #1
          mDOS.MoveFile App.Path & "/PEEK/Datos/" & mDir & "/" & mFile, App.Path & "/PEEK/Datos/" & mDir & "/" & Mid(mFile, 1, Len(mFile) - 3) & "txt"
       Else
          mErrorMsg = mError
          Close #1
          mDOS.MoveFile App.Path & "/PEEK/Datos/" & mDir & "/" & mFile, App.Path & "/PEEK/Datos/" & mDir & "/" & Mid(mFile, 1, Len(mFile) - 3) & "BAD"
       End If
       VarTexto = ""
       mFile = Dir(App.Path & "/PEEK/Datos/" & mDir & "/*.prn")
    Wend
    DBConn.Execute "UPDATE " & mDir & " SET Hora='00:00' WHERE Hora='24:00'"
    If mErrorMsg Then
       MsgBox "Hubo Errores en la Introducción de Datos, Ver en Informes - Errores", vbCritical, sMessage
    End If
    sMsgEspere MDI, "", False
    MsgBox "Pasaje a Tabla Finalizado", vbInformation, sMessage
 End If
End Function

Public Function xInsFallasRed(ByVal pCodigo As String, _
                              ByVal pFecha As String, _
                              ByVal pEstado As String, _
                              ByVal pDescr As String) As Boolean
   
   strSql1 = "insert into fallasred values ('" & pCodigo & "','" & Format(pFecha, "yyyy-mm-dd hh:mm") & "', " _
         & " '" & pEstado & "', '" & pDescr & "')"
   cmdTemp.CommandText = strSql1
   cmdTemp.CommandType = 1
   cmdTemp.CommandTimeout = 36000
   Set cmdTemp.ActiveConnection = DBConn
   cmdTemp.Execute
   xInsFallasRed = True
End Function

Public Function oSumCarrilWT(ByVal pCodCont As String, ByVal pYear As Integer, ByVal pMes As Integer, ByVal pClases As String) As ADODB.Recordset
   strSql1 = "select fecha, carril, sum(c1+c2+c3+c4+c5+c6+c7+c8+c9) as total From " & pCodCont & " " _
      & " where month(fecha)=" & pMes & " and year(Fecha)=" & pYear & pClases & "  group by fecha, carril order by 1, 2"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oSumCarrilWT = Rs1
End Function

Public Function oSumCarrilWTHora(ByVal pCodCont As String, ByVal pYear As Integer, ByVal pMes As Integer, ByVal pClases As String) As ADODB.Recordset
   strSql1 = "select fecha, carril, hora, sum(c1+c2+c3+c4+c5+c6+c7+c8+c9) as total From " & pCodCont & " " _
      & " where month(fecha)=" & pMes & " and year(Fecha)=" & pYear & pClases & "  group by fecha, carril, hora order by 1, 2"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oSumCarrilWTHora = Rs1
End Function

Public Function oSumVelWT(ByVal pCodCont As String, ByVal pYear As Integer, ByVal pMes As Integer, ByVal pCarriles As String) As ADODB.Recordset
   strSql1 = "SELECT carril, sum(v1) as v1, sum(v2) as v2, sum(v3) as v3, sum(v4) as v4, sum(v5) as v5, sum(v6) as v6, sum(v7) as v7 From " & pCodCont & " " _
      & " where month(fecha)=" & pMes & " and year(Fecha)=" & pYear & pCarriles & "  group by carril order by 1"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oSumVelWT = Rs1
End Function

Public Function oSumVelWT2(ByVal pCodCont As String, ByVal pYear As Integer, ByVal pMes As Integer, ByVal pCarriles As String) As ADODB.Recordset
   strSql1 = "SELECT carril, sum(v1) as v1, sum(v2) as v2, sum(v3) as v3, sum(v4) as v4, sum(v5) as v5, sum(v6) as v6, sum(v7) as v7 From " & pCodCont & " " _
      & " where month(fecha)=" & pMes & " and year(Fecha)=" & pYear & pCarriles & "  group by carril order by 1"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oSumVelWT2 = Rs1
End Function



Public Function oSumLongWT(ByVal pCodCont As String, ByVal pYear As Integer, ByVal pMes As Integer, ByVal pCarriles As String) As ADODB.Recordset
   strSql1 = "SELECT carril, sum(c1) as c1, sum(c2) as c2, sum(c3) as c3, sum(c4) as c4, sum(c5) as c5, sum(c6) as c6, sum(c7) as c7, sum(c8) as c8 From " & pCodCont & " " _
      & " where month(fecha)=" & pMes & " and year(Fecha)=" & pYear & pCarriles & "  group by carril order by 1"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oSumLongWT = Rs1
End Function

Public Function oSumVelWTFecha(ByVal pCodCont As String, ByVal pYear As Integer, ByVal pMes As Integer, ByVal pCarriles As String) As ADODB.Recordset
   strSql1 = "SELECT fecha, carril, sum(v1) as v1, sum(v2) as v2, sum(v3) as v3, sum(v4) as v4, sum(v5) as v5, sum(v6) as v6, sum(v7) as v7 From " & pCodCont & " " _
      & " where month(fecha)=" & pMes & " and year(Fecha)=" & pYear & pCarriles & "  group by fecha, carril order by 1, 2"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oSumVelWTFecha = Rs1
End Function

Public Function oSumLongWTFecha(ByVal pCodCont As String, ByVal pYear As Integer, ByVal pMes As Integer, ByVal pCarriles As String) As ADODB.Recordset
   strSql1 = "SELECT fecha, carril, carril, sum(c1) as c1, sum(c2) as c2, sum(c3) as c3, sum(c4) as c4, sum(c5) as c5, sum(c6) as c6, sum(c7) as c7, sum(c8) as c8 From " & pCodCont & " " _
      & " where month(fecha)=" & pMes & " and year(Fecha)=" & pYear & pCarriles & "  group by fecha, carril order by 1, 2"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oSumLongWTFecha = Rs1
End Function

Public Function oSumVelWTFechaHora(ByVal pCodCont As String, ByVal pYear As Integer, ByVal pMes As Integer, ByVal pCarriles As String) As ADODB.Recordset
   strSql1 = "SELECT fecha, hora, carril, sum(v1) as v1, sum(v2) as v2, sum(v3) as v3, sum(v4) as v4, sum(v5) as v5, sum(v6) as v6, sum(v7) as v7 From " & pCodCont & " " _
      & " where month(fecha)=" & pMes & " and year(Fecha)=" & pYear & pCarriles & "  group by fecha, hora, carril order by 1, 2, 3"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oSumVelWTFechaHora = Rs1
End Function

Public Function oSumVelWTFechaHoraParaAtascos(ByVal pCodCont As String, ByVal pYear As Integer, ByVal pMes As Integer, ByVal pCarriles As String) As ADODB.Recordset
   strSql1 = "SELECT * From " & pCodCont & " " _
      & " where month(fecha)=" & pMes & " and year(Fecha)=" & pYear & "  ORDER BY 1, 2"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oSumVelWTFechaHoraParaAtascos = Rs1
   
   
End Function



Public Function oSumLongWTFechaHora(ByVal pCodCont As String, ByVal pYear As Integer, ByVal pMes As Integer, ByVal pCarriles As String) As ADODB.Recordset
   strSql1 = "SELECT fecha, hora, carril, carril, sum(c1) as c1, sum(c2) as c2, sum(c3) as c3, sum(c4) as c4, sum(c5) as c5, sum(c6) as c6, sum(c7) as c7, sum(c8) as c8 From " & pCodCont & " " _
      & " where month(fecha)=" & pMes & " and year(Fecha)=" & pYear & pCarriles & "  group by fecha, hora, carril order by 1, 2, 3"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oSumLongWTFechaHora = Rs1
End Function

Public Function fImportWT(ByVal mDir As String)
 Dim mFecha As String
 Dim mLineaTxt As String
 Dim mFile As String
 Dim mError As Boolean
 Dim mExistFecha As Boolean
   mError = True

'   Select Case Trim(mDir)
'      Case "km14"
'         mTablas(1) = "wt14v"
'         mTablas(2) = "wt14c"
'
'      Case "km23"
'         mTablas(1) = "wt23v"
'         mTablas(2) = "wt23c"
'
'      Case "km32"
'         mTablas(1) = "wt32v"
'         mTablas(2) = "wt32c"
'
'      Case "km36"
'         mTablas(1) = "wt36v"
'         mTablas(2) = "wt36c"
'
'      Case "rta5"
'         mTablas(1) = "wta5v"
'         mTablas(2) = "wta5c"
'      Case "Test"
'         mTablas(1) = "wttev"
'         mTablas(2) = "wttec"
'      Case "Balb"
'         mTablas(1) = "wtbav"
'         mTablas(2) = "wtbac"
'      Case "Belg"
'         mTablas(1) = "wtbev"
'         mTablas(2) = "wtbec"
'      Case "Boqu"
'         mTablas(1) = "wtbov"
'         mTablas(2) = "wtboc"
'      Case "Esco"
'         mTablas(1) = "wtesv"
'         mTablas(2) = "wtesc"
'      Case "Masc"
'         mTablas(1) = "wtmav"
'         mTablas(2) = "wtmac"
'      Case "Pi46"
'         mTablas(1) = "wtpiv"
'         mTablas(2) = "wtpic"
'      Case "Urug"
'         mTablas(1) = "wturv"
'         mTablas(2) = "wturc"
'
'   End Select
   
   If MsgBox("¿Está Seguro de Pasar Datos a la Base?", vbYesNo, sMessage) = vbYes Then
      sMsgEspere MDI, "Procesando datos... aguarde un momento.", True
      mFile = Dir(App.Path & "/PEEK/DatosWT/" & mDir & "/*.txt")
      Do While mFile <> ""
         mError = False
         Open App.Path & "/PEEK/Datoswt/" & mDir & "/" & Trim(mFile) For Input As #1
         mFecha = "" & Left(mFile, 4) & "-" & Mid(mFile, 5, 2) & "-" & Mid(mFile, 7, 2)
         Do
            Line Input #1, mLineaTxt
         Loop Until LCase(Left(Trim(Left(mLineaTxt, 5)), 1)) = "c"
         If Not IsDate(mFecha) Then
            mError = True
         End If
         Select Case UCase(mDir)
            Case "KM14"
               mExistFecha = bExistDato("wt14v", mFecha)
            Case "KM23"
               mExistFecha = bExistDato("wt23v", mFecha)
            Case "KM32"
               mExistFecha = bExistDato("wt32v", mFecha)
            Case "KM36"
               mExistFecha = bExistDato("wt36v", mFecha)
            Case "RTA5"
               mExistFecha = bExistDato("wta5v", mFecha)
            Case "TEST"
               mExistFecha = bExistDato("wttev", mFecha)
            Case "BALB"
               mExistFecha = bExistDato("wtbav", mFecha)
            Case "BELG"
               mExistFecha = bExistDato("wtbev", mFecha)
            Case "BOQU"
               mExistFecha = bExistDato("wtbov", mFecha)
            Case "ESCO"
               mExistFecha = bExistDato("wtesv", mFecha)
            Case "MASC"
               mExistFecha = bExistDato("wtmav", mFecha)
            Case "PI46"
               mExistFecha = bExistDato("wtpiv", mFecha)
            Case "URUG"
               mExistFecha = bExistDato("wturv", mFecha)
            Case "U2UG"
               mExistFecha = bExistDato("wtu2v", mFecha)
         End Select
         
         If Not mError Then 'And Not mExistFecha Then
         
         
            Do While Not EOF(1)
               If Trim(mLineaTxt) <> "" Then
                  ins_wtxxg mLineaTxt, mDir
               End If
               Line Input #1, mLineaTxt
            Loop
               If Trim(mLineaTxt) <> "" Then
                  ins_wtxxg mLineaTxt, mDir
               End If
            
            ins_wtxxc mDir
            ins_wtxxv mDir

            Close #1
            mDOS.MoveFile App.Path & "/PEEK/DatosWT/" & mDir & "/" & mFile, App.Path & "/PEEK/DatosWT/" & mDir & "/" & Mid(mFile, 1, Len(mFile) - 3) & "OK"
         Else
            mErrorMsg = mError
            Close #1
            mDOS.MoveFile App.Path & "/PEEK/DatosWT/" & mDir & "/" & mFile, App.Path & "/PEEK/DatosWT/" & mDir & "/" & Mid(mFile, 1, Len(mFile) - 3) & "BAD"
         End If
         mFile = Dir(App.Path & "/PEEK/DatosWT/" & mDir & "/*.txt")
      Loop
      If mErrorMsg Then
         MsgBox "Hubo Errores en la Introducción de Datos, Ver en Informes - Errores", vbCritical, sMessage
      End If
      sMsgEspere MDI, "", False
      MsgBox "Pasaje a Tabla Finalizado", vbInformation, sMessage
   End If
End Function



'Public Function fImportWT(ByVal mDir As String)
Public Function fImportWTPorVehiculo(ByVal mDir As String)
 Dim mBand As String
 
 Dim mFecha As String
 Dim mCarril As String
 Dim mLength As String
 Dim mKPH As String
 Dim mClass As String
 Dim mRange As String
 Dim mHora As String
 
 
 
 Dim mFechaA As String
 Dim mFechaC
 
 Dim mHoraA As String
 Dim mLineaTxt As String
 Dim mWhere As String
 Dim mValueV As String
 Dim mValueL As String
 'Dim mCarril As String
 Dim mTablas(2) As String
 
 Dim mDatosV(12, 7) As Integer
 Dim mDatosL(12, 8) As Integer
 Dim mDatosLAux(8) As Integer
 Dim mFile As String
 Dim mi As Integer
 Dim mj As Integer
 Dim mK As Integer
 Dim mLane As Integer
 Dim mPos As Integer
 Dim mCant As Integer
 Dim mError As Boolean
 Dim mFlag As Boolean
 Dim mCtrlError As Boolean
 Dim mExistFecha As Boolean
 Dim mPosAuxiliar As Integer
 Dim mTotalXLongitud As Integer
 Dim mTotalTransitos As Integer
 
 
   mError = True
   mCtrlError = False
   Select Case Trim(mDir)
      Case "km14"
         mTablas(1) = "wt14v"
      Case "km23"
         mTablas(1) = "wt23v"
      Case "km32"
         mTablas(1) = "wt32v"
      Case "km36"
         mTablas(1) = "wt36v"
      Case "rta5"
         mTablas(1) = "wta5v"
      Case "Test"
         mTablas(1) = "wtVehiculoTe"
      Case "Balb"
         mTablas(1) = "wtbav"
      Case "Belg"
         mTablas(1) = "wtbev"
      Case "Boqu"
         mTablas(1) = "wtbov"
      Case "Esco"
         mTablas(1) = "wtesv"
      Case "Masc"
         mTablas(1) = "wtmav"
      Case "Pi46"
         mTablas(1) = "wtpiv"
      Case "Urug"
         mTablas(1) = "wturv"
   End Select
   
   If MsgBox("¿Está Seguro de Pasar Datos a la Base?", vbYesNo, sMessage) = vbYes Then
      sMsgEspere MDI, "Procesando datos... aguarde un momento.", True
      mFile = Dir(App.Path & "/PEEK/DatosWT/PorVehiculo/" & mDir & "/*.txt")
      Do While mFile <> ""
         mError = False
         Open App.Path & "/PEEK/Datoswt/PorVehiculo/" & mDir & "/" & Trim(mFile) For Input As #1
         mFecha = "" & Left(mFile, 4) & "-" & Mid(mFile, 5, 2) & "-" & Mid(mFile, 7, 2)
         Do
            Line Input #1, mLineaTxt
         
         Loop Until LCase(Left(Trim(Left(mLineaTxt, 5)), 1)) = "c"
      '--   If Not IsDate(mFecha) Then
         '--   mError = True
         '--End If
         '--mPos = InStr(1, mLineaTxt, mFecha)
'         Select Case UCase(mDir)
'            Case "KM14"
'               mExistFecha = bExistDato("wt14v", mFecha)
'            Case "KM23"
'               mExistFecha = bExistDato("wt23v", mFecha)
'            Case "KM32"
'               mExistFecha = bExistDato("wt32v", mFecha)
'            Case "KM36"
'               mExistFecha = bExistDato("wt36v", mFecha)
'            Case "RTA5"
'               mExistFecha = bExistDato("wta5v", mFecha)
'            Case "TEST"
'               mExistFecha = bExistDato("wtTev", mFecha)
'            Case "BALB"
'               mExistFecha = bExistDato("wtbav", mFecha)
'            Case "BELG"
'               mExistFecha = bExistDato("wtbev", mFecha)
'            Case "BOQU"
'               mExistFecha = bExistDato("wtbov", mFecha)
'            Case "ESCO"
'               mExistFecha = bExistDato("wtesv", mFecha)
'            Case "MASC"
'               mExistFecha = bExistDato("wtmav", mFecha)
'            Case "PI46"
'               mExistFecha = bExistDato("wtpiv", mFecha)
'            Case "URUG"
'               mExistFecha = bExistDato("wturv", mFecha)
'         End Select
         '----  Comienzo de Datos ---------------------
         If Not mError Then 'And Not mExistFecha Then
                     
            mFlag = True
            Do While Not EOF(1)
               'If Not EOF(1) And Trim(mLineaTxt) <> "" Then
                  mCarril = Trim(Mid(mLineaTxt, 2, 4))
                  mLength = Trim(Mid(mLineaTxt, 12, 7))
                  mKPH = Trim(Mid(mLineaTxt, 22, 7))
                  If mKPH = "" Then
                     mKPH = "0,0"
                  End If
                  
                  mClass = Trim(Mid(mLineaTxt, 29, 7))
                  mRange = Trim(Mid(mLineaTxt, 37, 7))
                  mFecha = Trim(Mid(mLineaTxt, 45, 26))
                                                                     
                  xInsWTPorVehiculo mTablas(1), mCarril, mLength, mKPH, mClass, mRange, mFecha
                  'xInsWT mTablas(2), mFechaA, mHoraA, "C" & mI, mValueL & "0,0,0,0"
                  Line Input #1, mLineaTxt
               'End If
            Loop
            
                  mCarril = Trim(Mid(mLineaTxt, 2, 4))
                  mLength = Trim(Mid(mLineaTxt, 12, 7))
                  mKPH = Trim(Mid(mLineaTxt, 22, 7))
                  If mKPH = "" Then
                     mKPH = "0,0"
                  End If
                  
                  mClass = Trim(Mid(mLineaTxt, 29, 7))
                  mRange = Trim(Mid(mLineaTxt, 37, 7))
                  mFecha = Trim(Mid(mLineaTxt, 45, 26))
                                                                     
                  xInsWTPorVehiculo mTablas(1), mCarril, mLength, mKPH, mClass, mRange, mFecha
            
            Close #1
            mDOS.MoveFile App.Path & "/PEEK/DatosWT/PorVehiculo/" & mDir & "/" & mFile, App.Path & "/PEEK/DatosWT/PorVehiculo/" & mDir & "/" & Mid(mFile, 1, Len(mFile) - 3) & "OK"
         Else
            mErrorMsg = mError
            Close #1
            mDOS.MoveFile App.Path & "/PEEK/DatosWT/PorVehiculo/" & mDir & "/" & mFile, App.Path & "/PEEK/DatosWT/PorVehiculo/" & mDir & "/" & Mid(mFile, 1, Len(mFile) - 3) & "BAD"
         End If
         mFile = Dir(App.Path & "/PEEK/DatosWT/PorVehiculo/" & mDir & "/*.txt")
      Loop
      'DBConn.Execute "UPDATE " & mDir & " SET Hora='00:00' WHERE Hora='24:00'"
      If mErrorMsg Then
         MsgBox "Hubo Errores en la Introducción de Datos, Ver en Informes - Errores", vbCritical, sMessage
      End If
      sMsgEspere MDI, "", False
      MsgBox "Pasaje a Tabla Finalizado", vbInformation, sMessage
   End If
End Function

Public Function xInsWTPorVehiculo(ByVal pCont As String, ByVal pCarril As String, ByVal pLength As String, _
                     ByVal pKPH As String, ByVal pClass As String, ByVal pRange As String, ByVal pFecha As String)
   strSql1 = "insert into " & pCont & " values ('" & pCarril & "', " _
         & Replace(pLength, ",", ".") & ", " & Replace(pKPH, ",", ".") & "," & pClass & "," & pRange & ",'" & pFecha & "',0,0,0,0,0.00)"
   sAsignCmd
   cmdTemp.Execute
   
   

   
'                  mCarril = Trim(Mid(mLineaTxt, 2, 4))
'                  mLength = Trim(Mid(mLineaTxt, 12, 7))
'                  mKPH = Trim(Mid(mLineaTxt, 22, 7))
'                  mClass = Trim(Mid(mLineaTxt, 29, 7))
'                  mRange = Trim(Mid(mLineaTxt, 37, 7))
'                  mFecha = Trim(Mid(mLineaTxt, 45, 26))

   
   
   
End Function





















Public Function fImportWT_BACKUP(ByVal mDir As String)
 Dim mBand As String
 Dim mFecha As String
 Dim mFechaC
 Dim mHora As String
 Dim mHoraA As String
 Dim mLineaTxt As String
 Dim mWhere As String
 Dim mValueV As String
 Dim mValueL As String
 Dim mCarril As String
 Dim mTablas(2) As String
 
 Dim mDatosV(12, 7) As Integer
 Dim mDatosL(12, 8) As Integer
 Dim mFile As String
 Dim mi As Integer
 Dim mj As Integer
 Dim mLane As Integer
 Dim mPos As Integer
 Dim mCant As Integer
 Dim mError As Boolean
 Dim mFlag As Boolean
 Dim mCtrlError As Boolean
 Dim mExistFecha As Boolean
 
   mError = True
   mCtrlError = False
   Select Case Trim(mDir)
      Case "km14"
         mTablas(1) = "wt14v"
         mTablas(2) = "wt14c"
         
      Case "km23"
         mTablas(1) = "wt23v"
         mTablas(2) = "wt23c"
      
      Case "km32"
         mTablas(1) = "wt32v"
         mTablas(2) = "wt32c"
      
      Case "km36"
         mTablas(1) = "wt36v"
         mTablas(2) = "wt36c"
         
      Case "rta5"
         mTablas(1) = "wta5v"
         mTablas(2) = "wta5c"
      Case "Test"
         mTablas(1) = "wttev"
         mTablas(2) = "wttec"
      Case "Balb"
         mTablas(1) = "wtbav"
         mTablas(2) = "wtbac"
      Case "Belg"
         mTablas(1) = "wtbev"
         mTablas(2) = "wtbec"
      Case "Boqu"
         mTablas(1) = "wtbov"
         mTablas(2) = "wtboc"
      Case "Esco"
         mTablas(1) = "wtesv"
         mTablas(2) = "wtesc"
      Case "Masc"
         mTablas(1) = "wtmav"
         mTablas(2) = "wtmac"
      Case "Pi42"
         mTablas(1) = "wtpiv"
         mTablas(2) = "wtpic"
      Case "Urug"
         mTablas(1) = "wturv"
         mTablas(2) = "wturc"
         
   End Select
   
   If MsgBox("¿Está Seguro de Pasar Datos a la Base?", vbYesNo, sMessage) = vbYes Then
      sMsgEspere MDI, "Procesando datos... aguarde un momento.", True
      mFile = Dir(App.Path & "/PEEK/DatosWT/" & mDir & "/*.txt")
      Do While mFile <> ""
         mError = False
         Open App.Path & "/PEEK/Datoswt/" & mDir & "/" & Trim(mFile) For Input As #1
         mFecha = "" & Left(mFile, 4) & "-" & Mid(mFile, 5, 2) & "-" & Mid(mFile, 7, 2)
         Do
            Line Input #1, mLineaTxt
         
         Loop Until LCase(Left(Trim(Left(mLineaTxt, 5)), 1)) = "c"
         If Not IsDate(mFecha) Then
            mError = True
         End If
         mPos = InStr(1, mLineaTxt, mFecha)
         Select Case UCase(mDir)
            Case "KM14"
               mExistFecha = bExistDato("wt14v", mFecha)
            Case "KM23"
               mExistFecha = bExistDato("wt23v", mFecha)
            Case "KM32"
               mExistFecha = bExistDato("wt32v", mFecha)
            Case "KM36"
               mExistFecha = bExistDato("wt36v", mFecha)
            Case "RTA5"
               mExistFecha = bExistDato("wta5v", mFecha)
            Case "TEST"
               mExistFecha = bExistDato("wttev", mFecha)
            Case "BALB"
               mExistFecha = bExistDato("wtbav", mFecha)
            Case "BELG"
               mExistFecha = bExistDato("wtbev", mFecha)
            Case "BOQU"
               mExistFecha = bExistDato("wtbov", mFecha)
            Case "ESCO"
               mExistFecha = bExistDato("wtesv", mFecha)
            Case "MASC"
               mExistFecha = bExistDato("wtmav", mFecha)
            Case "PI42"
               mExistFecha = bExistDato("wtpiv", mFecha)
            Case "URUG"
               mExistFecha = bExistDato("wturv", mFecha)
         End Select
         '----  Comienzo de Datos ---------------------
         If Not mError And Not mExistFecha Then
            For mi = 1 To 12
               For mj = 1 To 7
                  mDatosV(mi, mj) = 0
                  mDatosL(mi, mj) = 0
               Next
            Next
            mDatosL(12, 8) = 0
            mFlag = True
            Do While Not EOF(1)
               If Trim(mLineaTxt) <> "" Then
                  mFechaC = DateAdd("n", -5, Trim(Mid(mLineaTxt, mPos, 20)))
               End If
               'If Format(Left(mFechaC, 10), "yyyy-mm-dd") = mFecha And Not EOF(1) And Trim(mLineaTxt) <> "" Then'20161003
               If Not EOF(1) And Trim(mLineaTxt) <> "" Then
                  mLane = CInt(Right(Trim(Left(mLineaTxt, 5)), 1))
                  If mFlag Then
                     'mHora = Mid(Trim(Mid(mLineaTxt, mPos, 20)), 12, 2)
                    ' mHora = mId(Format(DateAdd("n", -15, Trim(mId(mLineaTxt, mPos, 20))), "yyyy-mm-dd hh:mm:ss"), 12, 2)
                    mHora = Mid(Format(DateAdd("n", -5, Trim(Mid(mLineaTxt, mPos, 22))), "yyyy-mm-dd hh:mm:ss"), 12, 2)
                     mHoraA = mHora
                     mFlag = False
                  End If
                  If mHoraA = mHora Then
                     mHoraA = mHora
                     mj = 1
                     'vector de longitudes
                     For mi = 46 To 81 Step 5
                        If Trim(Mid(mLineaTxt, mi, 5)) = "-" Then
                           mCant = 0
                        Else
                           mCant = CInt(Trim(Mid(mLineaTxt, mi, 5)))
                        End If
                        mDatosL(mLane, mj) = mCant + mDatosL(mLane, mj)
                        mj = mj + 1
                     Next
                     mj = 1
                     'vector velocidades
                     'For mI = 126 To 162 Step 6
                     For mi = 135 To 171 Step 6
                        If Trim(Mid(mLineaTxt, mi, 6)) = "-" Then
                           mCant = 0
                        Else
                           mCant = CInt(Trim(Mid(mLineaTxt, mi, 6)))
                        End If
                        mDatosV(mLane, mj) = mCant + mDatosV(mLane, mj)
                        mj = mj + 1
                     Next
                     Line Input #1, mLineaTxt
                     If Not EOF(1) And Trim(mLineaTxt) <> "" Then
                       ' mHora = mId(Format(DateAdd("n", -15, Trim(mId(mLineaTxt, mPos, 20))), "yyyy-mm-dd hh:mm:ss"), 12, 2)
                        mHora = Mid(Format(DateAdd("n", -5, Trim(Mid(mLineaTxt, mPos, 22))), "yyyy-mm-dd hh:mm:ss"), 12, 2)
                     End If
                  Else
                     mi = 0
                     For mi = 1 To 12
                        mValueV = ""
                        mValueL = ""
                        For mj = 1 To 8
                           If mj < 8 Then
                              mValueV = mValueV & mDatosV(mi, mj) & ","
                              mDatosV(mi, mj) = 0
                           End If
                           mValueL = mValueL & mDatosL(mi, mj) & ","
                           mDatosL(mi, mj) = 0
                        Next
                        'xInsWT mTablas(1), mFecha, mHoraA, "C" & mI, mValueV & "0,0,0,0,0"'20161003
                        'xInsWT mTablas(2), mFecha, mHoraA, "C" & mI, mValueL & "0,0,0,0"'20161003
                        xInsWT mTablas(1), mFechaC, mHoraA, "C" & mi, mValueV & "0,0,0,0,0"
                        xInsWT mTablas(2), mFechaC, mHoraA, "C" & mi, mValueL & "0,0,0,0"
                        
                     Next
                     mHoraA = mHora
                  End If
               Else
                  Line Input #1, mLineaTxt
                  If Not EOF(1) And Trim(mLineaTxt) <> "" Then
                     mHora = Mid(Format(DateAdd("n", -5, Trim(Mid(mLineaTxt, mPos, 20))), "yyyy-mm-dd hh:mm:ss"), 12, 2)
                  End If
               End If
            Loop
            mi = 0
            For mi = 1 To 12
               mValueV = ""
               mValueL = ""
               For mj = 1 To 8
                  If mj < 8 Then
                     mValueV = mValueV & mDatosV(mi, mj) & ","
                     mDatosV(mi, mj) = 0
                  End If
                  mValueL = mValueL & mDatosL(mi, mj) & ","
                  mDatosL(mi, mj) = 0
               Next
               'xInsWT mTablas(1), mFecha, mHoraA, "C" & mI, mValueV & "0,0,0,0,0" '20161003
               'xInsWT mTablas(2), mFecha, mHoraA, "C" & mI, mValueL & "0,0,0,0" '20161003
               xInsWT mTablas(1), mFechaC, mHoraA, "C" & mi, mValueV & "0,0,0,0,0"
               xInsWT mTablas(2), mFechaC, mHoraA, "C" & mi, mValueL & "0,0,0,0"
            Next
            Close #1
            mDOS.MoveFile App.Path & "/PEEK/DatosWT/" & mDir & "/" & mFile, App.Path & "/PEEK/DatosWT/" & mDir & "/" & Mid(mFile, 1, Len(mFile) - 3) & "OK"
         Else
            mErrorMsg = mError
            Close #1
            mDOS.MoveFile App.Path & "/PEEK/DatosWT/" & mDir & "/" & mFile, App.Path & "/PEEK/DatosWT/" & mDir & "/" & Mid(mFile, 1, Len(mFile) - 3) & "BAD"
         End If
         mFile = Dir(App.Path & "/PEEK/DatosWT/" & mDir & "/*.txt")
      Loop
      'DBConn.Execute "UPDATE " & mDir & " SET Hora='00:00' WHERE Hora='24:00'"
      If mErrorMsg Then
         MsgBox "Hubo Errores en la Introducción de Datos, Ver en Informes - Errores", vbCritical, sMessage
      End If
      sMsgEspere MDI, "", False
      MsgBox "Pasaje a Tabla Finalizado", vbInformation, sMessage
   End If
End Function



Public Function xInsWT(ByVal pCont As String, ByVal pFecha As String, ByVal pHora As String, _
                     ByVal pCarril As String, ByVal pValores As String)
   strSql1 = "insert into " & pCont & " values ('" & Format(pFecha, "yyyy-mm-dd hh:mm") & "', " _
         & " '" & pHora & "', '" & pCarril & "'," & pValores & ")"
   sAsignCmd
   cmdTemp.Execute
End Function

Public Function sCarrilesSent(ByVal pCodCont As String, ByVal pSent As String) As String
   sCarrilesSent = ""
   Set Rs1 = oPeekCarril_Sent(pCodCont, pSent)
   If Not Rs1.EOF Then
      Do While Not Rs1.EOF
         sCarrilesSent = sCarrilesSent & "'C" & Rs1!carril & "', "
         Rs1.MoveNext
      Loop
      sCarrilesSent = Mid(sCarrilesSent, 1, Len(sCarrilesSent) - 2)
   End If
   Rs1.Close
End Function

Public Function oWTDias(ByVal pCodCont As String, ByVal pYear As Integer, ByVal pMes As Integer) As ADODB.Recordset
   strSql1 = "select distinct fecha from " & pCodCont & " where month(fecha)=" & pMes & " and year(fecha)=" & pYear & " order by fecha"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oWTDias = Rs1
End Function

'===========================================================================================
'=======               Llamadas a Store Procedures MySQL: Schema Peek             ========
'===========================================================================================

Public Sub ins_wtxxg(ByVal linea As String, ByVal Directorio As String)
   Dim paramLinea As parametroSql

   Dim aParamSql(0) As parametroSql
      
   With paramLinea
      .Nombre = "str"
      .Tipo = adVarChar
      .Direccion = adParamInput
      .Tamanio = 500
      .Valor = linea
   End With
   aParamSql(0) = paramLinea

   Select Case Directorio
            Case "Test"
               xStoreProcedureWithParam "ins_wtteg", aParamSql
            Case "Balb"
               xStoreProcedureWithParam "ins_wtbag", aParamSql
            Case "Belg"
               xStoreProcedureWithParam "ins_wtbeg", aParamSql
            Case "Boqu"
               xStoreProcedureWithParam "ins_wtbog", aParamSql
            Case "Esco"
               xStoreProcedureWithParam "ins_wtesg", aParamSql
            Case "Masc"
               xStoreProcedureWithParam "ins_wtmag", aParamSql
            Case "Pi46"
               xStoreProcedureWithParam "ins_wtpig", aParamSql
            Case "Urug"
               xStoreProcedureWithParam "ins_wturg", aParamSql
            Case "U2ug"
               xStoreProcedureWithParam "ins_wtu2g", aParamSql
         End Select
 
End Sub
Public Sub ins_wtxxc(ByVal Directorio As String)
   Select Case Directorio
            Case "Test"
               xStoreProcedure "ins_wttec"
            Case "Balb"
               xStoreProcedure "ins_wtbac"
            Case "Belg"
               xStoreProcedure "ins_wtbec"
            Case "Boqu"
               xStoreProcedure "ins_wtboc"
            Case "Esco"
               xStoreProcedure "ins_wtesc"
            Case "Masc"
               xStoreProcedure "ins_wtmac"
            Case "Pi46"
               xStoreProcedure "ins_wtpic"
            Case "Urug"
               xStoreProcedure "ins_wturc"
            Case "U2ug"
               xStoreProcedure "ins_wtu2c"
         End Select
End Sub
Public Sub ins_wtxxv(ByVal Directorio As String)
   Select Case Directorio
            Case "Test"
               xStoreProcedure "ins_wttev"
            Case "Balb"
               xStoreProcedure "ins_wtbav"
            Case "Belg"
               xStoreProcedure "ins_wtbev"
            Case "Boqu"
               xStoreProcedure "ins_wtbov"
            Case "Esco"
               xStoreProcedure "ins_wtesv"
            Case "Masc"
               xStoreProcedure "ins_wtmav"
            Case "Pi46"
               xStoreProcedure "ins_wtpiv"
            Case "Urug"
               xStoreProcedure "ins_wturv"
            Case "U2ug"
               xStoreProcedure "ins_wtu2v"
         End Select
End Sub







