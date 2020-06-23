VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clLogUser"
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
   'Set DBConn = mObj.oAbreConexion("4", mIPServer, "loguser", "ssvv", "ssvv") 'oAbreConexion esta en el modulo1
   Set DBConn = mObj.oAbreConexion("5", mIPServer, "loguser", "ssvv", "ssvv") '20190527
   'Set DBConn = mObj.oAbreConexion("4", mIPServer, "loguser", "root", "123456") 'oAbreConexion esta en el modulo1
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

Public Function oTabla(ByVal pTabla As String, ByVal pWhere As String) As ADODB.Recordset
   strSql1 = "SELECT * FROM " & LCase(pTabla) & " " & pWhere
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oTabla = Rs1
End Function

Public Function sCampoDescrip(ByVal pTabla As String, ByVal pWhereCodigo As String, ByVal pCampoDescr As Integer) As String
   strSql1 = "select * from " & LCase(pTabla) & " where " & pWhereCodigo
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   sCampoDescrip = ""
   If Not Rs1.EOF Then
      sCampoDescrip = Rs1.Fields(pCampoDescr)
   End If
End Function

Public Function oValidUsuario(ByVal pCodUsuario As String, ByVal pClave As String) As ADODB.Recordset
   'If mIPServer <> "10.10.30.208" Then
      'If mIPServer <> "10.128.1.209" Then
   If mIPServer <> "desa-SSVV" Then
      strSql1 = "select * from usuarios where codusuario='" & pCodUsuario & "' and clave=md5('" & pClave & "') and fechabaja is null"
   Else
     strSql1 = "select * from usuarios where codusuario='" & pCodUsuario & "' and clave=md5('" & pClave & "') and fechabaja is null"
   End If
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oValidUsuario = Rs1
End Function

Public Function oSistemasUsuario(ByVal pCodUsuario As String) As ADODB.Recordset
   strSql1 = "SELECT DISTINCT CODSISTEMA FROM permisos WHERE CODUSUARIO = '" & pCodUsuario & "'"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oSistemasUsuario = Rs1
End Function

Public Function oMenuSistema(ByVal pCodSistema As String) As ADODB.Recordset
   strSql1 = "select distinct * from menues where codsistema='" & pCodSistema & "'"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oMenuSistema = Rs1
End Function

Public Function oPermisoMenu(ByVal pCodUsuario As String, ByVal pCodSistema As String) As ADODB.Recordset
   strSql1 = "SELECT distinctrow  P.*, M.Descripcion FROM permisos P, menues M WHERE P.CodSistema = '" & pCodSistema & "' AND P.CodUsuario='" & pCodUsuario & "' and P.codmenu=M.codigomenu order by 4"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Set oPermisoMenu = Rs1
End Function

Public Function bRepetClave(ByVal pCodUser As String, ByVal pClave As String) As Boolean
   strSql1 = "select * from usuarios WHERE CodUsuario='" & pCodUser & "' AND md5('" & pClave & "') in (clave, clave1, clave2, clave3, clave4, clave5)"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   bRepetClave = (Not Rs1.EOF)
End Function

Public Function bUserBloqueado(ByVal pCodUsuario As String) As Boolean
   bUserBloqueado = True
   strSql1 = "select fechadelay, NOW() AS FechaActual from usuarios where CodUsuario='" & pCodUsuario & "'"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   If Not Rs1.EOF Then
     If DateDiff("n", NVL(Rs1!fechadelay, "20/06/2000 01:00:00"), Rs1!FechaActual) > 15 Then  '15 Minutos de inhabilitación de ingreso al sistema
        bUserBloqueado = False
     End If
   Else
      bUserBloqueado = False
      MsgBox "Usuario Inexistente", vbExclamation, sMessage
   End If
   Rs1.Close
End Function

Public Function bDurMinClave(ByVal pCodUser As String) As Boolean
   bDurMinClave = False
   strSql1 = "select FechaChgPwd, NOW() AS FechaActual from usuarios where CodUsuario='" & pCodUser & "'"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   If Not Rs1.EOF Then
      If DateDiff("d", Rs1!fechachgpwd, Rs1!FechaActual) <= 15 Then
         bDurMinClave = True
      End If
   End If
End Function

Public Function bControlClave(ByVal pCodUser As String, ByVal pClave As String) As Boolean
   bControlClave = False
   'strSql1 = "select fechachgpwd, NOW() AS FechaActual from usuarios where CodUsuario='" & pCodUser & "' and FechaBaja IS NULL and CodPerfil <> '1' "
   strSql1 = "select fechachgpwd, NOW() AS FechaActual from usuarios where CodUsuario='" & pCodUser & "' and FechaBaja IS NULL and CodPerfil not in ('0','1') "
   If pClave <> "" Then
      strSql1 = strSql1 & " and Clave=md5('" & pClave & "') "
   End If
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   If Not Rs1.EOF Then
      If DateDiff("d", NVL(Rs1!fechachgpwd, "20/06/2000 01:00:00"), Rs1!FechaActual) > 60 Then '60 días
         bControlClave = True
      End If
   End If
End Function

Public Function sActivarMenu(ByVal pCodUsuario As String, ByVal pCodSistema As Integer)
Dim mi As Integer
Dim mj As Integer
   strSql1 = "select codmenu from permisos where codusuario='" & pCodUsuario & "' and codsistema='" & pCodSistema & "'"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   Do While Not Rs1.EOF
      For mi = 0 To MDI.Count - 1
    '  MsgBox (MDI(mI).Name)
         If Right(Rs1!codmenu, 1) = ")" Then
            'MsgBox (MDI(mI).Name)
            If Mid(Rs1!codmenu, 1, Len(Rs1!codmenu) - 3) = MDI(mi).Name Then
               For mj = 0 To Left(Right(Rs1!codmenu, 2), 1) - 1
                  mi = mi + 1
               Next
               MDI(mi).Enabled = True
               mi = 999
            Else
               If Mid(Rs1!codmenu, 1, Len(Rs1!codmenu) - 4) = MDI(mi).Name Then
                  For mj = 0 To Left(Right(Rs1!codmenu, 3), 2) - 1
                     mi = mi + 1
                  Next
                  MDI(mi).Enabled = True
                  mi = 999
               End If
            End If
         Else
            If Rs1!codmenu = MDI(mi).Name Then
               MDI(mi).Enabled = True
               mi = 999
            End If
         End If
      Next
      Rs1.MoveNext
   Loop
End Function

Public Function sEmail(ByVal pCodigo As String) As String  'Todas las Vías
   strSql1 = "SELECT email FROM usuarios WHERE CodUsuario='" & pCodigo & "'"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   sEmail = ""
   If Not Rs1.EOF Then
      sEmail = Mid(Rs1!Email, 1, Len(Rs1!Email) - 11)
   End If
End Function

Public Function xDeletePermisos(ByVal pCodUser As String, ByVal pCodSistema As String) As Boolean
   strSql1 = "delete from permisos where codusuario='" & pCodUser & "' and codsistema='" & pCodSistema & "'"
   sAsignCmd
   cmdTemp.Execute
   xDeletePermisos = True
End Function

Public Function xInsertMD5(ByVal pCodUser As String, ByVal pClave As String)
   strSql1 = "select * from usuariosmd5 where codigo='" & pCodUser & "'"
   sAsignCmd
   Set Rs1 = cmdTemp.Execute
   strSql1 = "insert into usuariosmd5 values ('" & pCodUser & "',md5('" & pClave & "'),null,null,null,null,null)"
   If Not Rs1.EOF Then
      strSql1 = "update usuariosmd5 set clave6=clave5, clave5=clave4, clave4=clave3, clave3=clave2, clave2=clave1, clave1=md5('" & pClave & "') where codigo='" & pCodUser & "'"
   End If
   cmdTemp.CommandText = strSql1
   cmdTemp.Execute
End Function

Public Function xUpClavesUsuarios(ByVal pCodUser As String, ByVal pClave As String, ByVal pClave1 As String, ByVal pClave2 As String, ByVal pClave3 As String, ByVal pClave4 As String) As Boolean
   strSql1 = "UPDATE usuarios SET Clave5='" & pClave4 & "', Clave4='" & pClave3 & "',Clave3='" & pClave2 & "',Clave2='" & pClave1 & "', Clave1='" & pClave & "' WHERE CodUsuario= '" & pCodUser & "'"
   sAsignCmd
   cmdTemp.Execute
   xUpClavesUsuarios = True
End Function

Public Function xUpUnLockUser(ByVal pCodUser As String) As Boolean
   strSql1 = "UPDATE usuarios SET  FechaDelay=null where codusuario='" & pCodUser & "'"
   sAsignCmd
   cmdTemp.Execute
   xUpUnLockUser = True
End Function

Public Function xUpResetClaveUser(ByVal pCodUser As String) As Boolean
   strSql1 = "UPDATE usuarios SET clave='qaTyrufqm1Uyc', chgpass='0', FechaDelay=null, FechaChgPwd='2006-06-20 09:30:00' where codusuario='" & pCodUser & "'"
   sAsignCmd
   cmdTemp.Execute
   xUpResetClaveUser = True
End Function

Public Function xUpChgClaveUser(ByVal pCodUser As String, ByVal pClave As String) As Boolean
   strSql1 = "UPDATE usuarios SET Clave=md5('" & pClave & "'), ChgPass='1', FechaChgPwd=NOW() WHERE CodUsuario= '" & pCodUser & "'"
   sAsignCmd
   cmdTemp.Execute
   xUpChgClaveUser = True
End Function

Public Function xUpBajaUser(ByVal pCodUser As String)
   strSql1 = "UPDATE usuarios SET FechaBaja=CURRENT_DATE WHERE CodUsuario='" & pCodUser & "'"
   sAsignCmd
   cmdTemp.Execute
End Function

Public Function xInsNewUser(ByVal pCodUser As String, ByVal pApe As String, ByVal pName As String, ByVal pCodPerfil As String, ByVal pEmail As String)
   strSql1 = "INSERT INTO usuarios (codusuario, apellido, nombres, clave, codperfil, chgpass, email) VALUES " _
      & "('" & pCodUser & "','" & UCase(pApe) & "', '" & UCase(pName) & "', MD5('qazwsx'), " & pCodPerfil & ",'0','" & LCase(pEmail) & "')"
   sAsignCmd
   On Error Resume Next
   cmdTemp.Execute
   If Err.Description <> "" Then
      MsgBox "Existió un error: " & Chr(13) & "  " & Err.Description
   End If
End Function

Public Function xUpDatosUser(ByVal pCodUser As String, ByVal pCodPerfil As String, ByVal pApe As String, ByVal pName As String, ByVal pEmail As String)
   strSql1 = "UPDATE usuarios SET apellido='" & UCase(pApe) & "', nombres='" & UCase(pName) & "', CodPerfil='" & pCodPerfil & "', email='" & LCase(pEmail) & "' WHERE CodUsuario='" & pCodUser & "'"
   sAsignCmd
   On Error Resume Next
   cmdTemp.Execute
   If Err.Description <> "" Then
      MsgBox "Existió un error: " & Chr(13) & "  " & Err.Description
   End If
End Function

Public Function xSetDelayUser(ByVal pCodUser) As Boolean
   strSql1 = "update usuarios set FechaDelay=NOW() where codusuario='" & pCodUser & "'"
   sAsignCmd
   cmdTemp.Execute
   xSetDelayUser = True
End Function

Public Function xInsertPermisos(ByVal pCodUser As String, ByVal pCodMenu As String, ByVal pCodSistema As String)
   strSql1 = "insert into permisos values ('" & pCodUser & "','" & pCodMenu & "','" & pCodSistema & "')"
   sAsignCmd
   On Error Resume Next
   cmdTemp.Execute
   If Err.Description <> "" Then
      MsgBox Err.Description
   End If
End Function

Public Function UpdZonaxSup(ByVal pCodUser As String, pCodZona As String)
   strSql1 = "UPDATE usuarios SET Email = '" & pCodZona & "' WHERE CodUsuario = '" & pCodUser & "'"
   sAsignCmd
   cmdTemp.Execute
End Function
