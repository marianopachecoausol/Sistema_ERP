VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Dim mObjInven As New clInven
'Dim mRec As ADODB.Recordset
'Dim mRenglonProducto As Integer
'Dim mRenglonEgreso As Integer
'Dim mCodProducto As String
'Dim cboListIndex As Integer
'
'Private Sub Combo5_Click_OLD()
'   Dim mI As Integer
'
'   Combo4.Enabled = True
'   Combo2.Enabled = True
'
'   If cboListIndex <> Combo5.ListIndex Then
'      sLlenoUsuariosRet
'      sLlenoUsuariosAut
'      If (cboListIndex <> -1) Then
'
'         If MsgBox("Si selecciona otra Bodega se perderán los consumos cargados hasta el momento. ¿ Desea continuar ? ", vbYesNo, "Cambio de Bodega") = vbYes Then
'            Text9.Text = ""
'            Text10.Text = ""
'
'            'Elimino los registros de la grilla superior (productos)
'            For mI = FlexProd.Rows To 3 Step -1
'               FlexProd.RemoveItem mI
'            Next
'
'            'Elimino los registros de la grilla inferior (consumos)
'            For mI = FlexEgreso.Rows To 3 Step -1
'               FlexEgreso.RemoveItem mI
'            Next
'
'            mRenglonProducto = 0
'            mRenglonEgreso = 0
'         Else
'            Combo5.ListIndex = cboListIndex
'            sLlenoUsuariosRet
'            sLlenoUsuariosAut
'         End If
'
'         cboListIndex = Combo5.ListIndex
'
'      Else
'         cboListIndex = Combo5.ListIndex
'      End If
'
'   End If
'End Sub
'
'Private Sub Combo5_Click()
'   Dim mI As Integer
'
'   Combo4.Enabled = True
'   Combo2.Enabled = True
'
'   If cboListIndex <> Combo5.ListIndex Then
'      sLlenoUsuariosRet
'      sLlenoUsuariosAut
'      If (cboListIndex <> -1) Then
'         'Si tengo algun registro en la grilla inferior(Egresos)
'         If FlexEgreso.Rows > 2 Then
'            If MsgBox("Si selecciona otra Bodega se perderán los consumos cargados hasta el momento. ¿ Desea continuar ? ", vbYesNo, "Cambio de Bodega") = vbYes Then
'               Text9.Text = ""
'               Text10.Text = ""
'
'               'Elimino los registros de la grilla superior (productos)
'               For mI = FlexProd.Rows To 3 Step -1
'                  FlexProd.RemoveItem mI
'               Next
'
'               'Elimino los registros de la grilla inferior (consumos)
'               For mI = FlexEgreso.Rows To 3 Step -1
'                  FlexEgreso.RemoveItem mI
'               Next
'
'               mRenglonProducto = 0
'               mRenglonEgreso = 0
'            Else
'               Combo5.ListIndex = cboListIndex
'               sLlenoUsuariosRet
'               sLlenoUsuariosAut
'            End If
'         Else
'            Text9.Text = ""
'            Text10.Text = ""
'
'            'Elimino los registros de la grilla superior (productos)
'            For mI = FlexProd.Rows To 3 Step -1
'               FlexProd.RemoveItem mI
'            Next
'
'         End If
'
'         cboListIndex = Combo5.ListIndex
'
'      Else
'         cboListIndex = Combo5.ListIndex
'      End If
'
'   End If
'End Sub
'
'Private Sub sLlenoUsuariosRet()
'Dim mCodBodega As String
'Dim mObjInven2 As New clInven
'Dim mRec2 As New ADODB.Recordset
'
'   mCodBodega = Trim(Left(Combo5.Text, 4))
'   Combo4.Clear
'
'   Set mRec2 = mObjInven.oEjecutarSelect("SELECT CONCAT (P.Apellido,',', P.Nombres) AS Descripcion,P.CodUsuario AS CodUsuario FROM " & _
'   " UsuariosRet_Bodegas UB " & _
'   " Inner Join " & _
'   " Personal P ON UB.CodUsuario = P.CodUsuario " & _
'   " WHERE UB.CodBodega = '" & mCodBodega & "' AND P.Fecha_Baja IS NULL " & _
'   " ORDER BY P.Apellido;")
'
'
'   Do While Not mRec2.EOF
'      Combo4.AddItem mRec2!descripcion & Space(60) & mRec2!CodUsuario
'      mRec2.MoveNext
'   Loop
'   mRec2.Close
'   Set mObjInven2 = Nothing
'   Set mRec2 = Nothing
'End Sub
'
'
'Private Sub sLlenoUsuariosAut()
'Dim mCodBodega As String
'Dim mObjInven2 As New clInven
'Dim mRec2 As New ADODB.Recordset
'
'   mCodBodega = Trim(Left(Combo5.Text, 4))
'   Combo2.Clear
'
'   Set mRec2 = mObjInven.oEjecutarSelect("SELECT CONCAT (P.Apellido,',', P.Nombres) AS Descripcion,P.CodUsuario AS CodUsuario FROM " & _
'   " UsuariosAut_Bodegas UB " & _
'   " Inner Join " & _
'   " Personal P ON UB.CodUsuario = P.CodUsuario " & _
'   " WHERE UB.CodBodega = '" & mCodBodega & "' AND P.Fecha_Baja IS NULL " & _
'   " ORDER BY P.Apellido;")
'
'
'   Do While Not mRec2.EOF
'      Combo2.AddItem mRec2!descripcion & Space(60) & mRec2!CodUsuario
'      mRec2.MoveNext
'   Loop
'   mRec2.Close
'   Set mObjInven2 = Nothing
'   Set mRec2 = Nothing
'End Sub
'
'Private Sub Command1_Click()
'   Dim mI As Integer
'   Dim mJ As Integer
'
'   mRenglonProducto = 0
'
'   'Elimino los registros (de la consulta anterior) de la grilla superior
'   For mI = FlexProd.Rows To 3 Step -1
'      FlexProd.RemoveItem mI
'   Next
'
'   Set mRec = mObjInven.getStockXBodegaConFiltroProducto(Left(Combo5.Text, 4), Text9.Text)
'
'   'Cargo la Grilla Superior con lo devuelto por el sp "getStockXBodegaConFiltroProducto"
'   If Not mRec.EOF Then
'      mI = 1
'      Do While Not mRec.EOF
'         mI = mI + 1
'
'         FlexProd.AddItem ""
'         FlexProd.TextMatrix(mI, 1) = mRec!Producto
'         FlexProd.TextMatrix(mI, 2) = mRec!Ubicacion
'         FlexProd.TextMatrix(mI, 3) = mRec!Stock
'         FlexProd.TextMatrix(mI, 4) = mRec!UnidadMedida
'         FlexProd.TextMatrix(mI, 5) = mRec!CodigoSap
'         FlexProd.TextMatrix(mI, 6) = mRec!CodProducto
'         FlexProd.TextMatrix(mI, 7) = mRec!CodUbicacion
'
'         mRec.MoveNext
'      Loop
'   End If
'   mRec.Close
'
'   'Si presiono el boton Buscar y algun "producto/ubicacion" de la grilla de arriba, esta en la grilla inferior
'   'entonces que me actualice en la grilla superior el stock de ese o esos "producto/ubicacion" descontando el consumo de la grilla inferior
'   For mI = 2 To FlexProd.Rows - 1
'      For mJ = 2 To FlexEgreso.Rows - 1
'         If FlexProd.TextMatrix(mI, 6) = FlexEgreso.TextMatrix(mJ, 6) And FlexProd.TextMatrix(mI, 7) = FlexEgreso.TextMatrix(mJ, 7) Then
'            FlexProd.TextMatrix(mI, 3) = CDbl(Replace(Trim(FlexProd.TextMatrix(mI, 3)), ".", ",")) - CDbl(Replace(Trim(FlexEgreso.TextMatrix(mJ, 3)), ".", ","))
'            mJ = 999
'         End If
'      Next
'   Next
'End Sub
'
'Private Sub command3_Click(Index As Integer)
'   Dim iStock As Double
'   Dim mI As Integer
'   Dim mRec1 As New ADODB.Recordset
'
'   If Index = 0 Then
'      If fValidaEgreso() Then
'            FlexEgreso.AddItem vbTab & FlexProd.TextMatrix(mRenglonProducto, 1) & vbTab & FlexProd.TextMatrix(mRenglonProducto, 2) & vbTab & Text10.Text & vbTab & FlexProd.TextMatrix(mRenglonProducto, 4) & vbTab & FlexProd.TextMatrix(mRenglonProducto, 5) & vbTab & FlexProd.TextMatrix(mRenglonProducto, 6) & vbTab & FlexProd.TextMatrix(mRenglonProducto, 7)
'            FlexProd.TextMatrix(mRenglonProducto, 3) = CDbl(Replace(Trim(FlexProd.TextMatrix(mRenglonProducto, 3)), ".", ",")) - CDbl(Replace(Trim(Text10.Text), ".", ","))
'            Text10.Text = ""
'            Text10.SetFocus
'      End If
'   Else
'      For mI = 2 To FlexProd.Rows - 1
'
'         If FlexProd.TextMatrix(mI, 6) = FlexEgreso.TextMatrix(mRenglonEgreso, 6) And FlexProd.TextMatrix(mI, 7) = FlexEgreso.TextMatrix(mRenglonEgreso, 7) Then
'            Set mRec1 = mObjInven.oEjecutarSelect("SELECT Fecha,CodProducto,CodUbicacion, IFNULL(Stock,0) As Stock " & _
'                                       " FROM Movimientos2 M " & _
'                                       " WHERE CodProducto  = '" & FlexEgreso.TextMatrix(mRenglonEgreso, 6) & "' and CodUbicacion = '" & FlexEgreso.TextMatrix(mRenglonEgreso, 7) & "'" & _
'                                       " AND Fecha = (SELECT Max(Fecha) FROM Movimientos2 WHERE CodProducto = M.CodProducto AND CodUbicacion = M.CodUbicacion)")
'
'            If Not mRec1.EOF Then
'               iStock = mRec1!Stock
'            Else
'               iStock = 0
'            End If
'            mRec1.Close
'
'            FlexProd.TextMatrix(mI, 3) = iStock
'
'            mI = 9999
'         End If
'      Next
'
'      If FlexEgreso.Rows > 2 And mRenglonEgreso > 1 Then
'         FlexEgreso.RemoveItem (mRenglonEgreso)
'      End If
'
'      mRenglonEgreso = 0
'   End If
'End Sub
'
''Boton de confirmacion de Consumo de materiales
'Private Sub command3_Click(Index As Integer)
'   If Index = 0 Then
'      Dim vEgresosCodProducto() As String
'      Dim vEgresosCodUbicacion() As String
'      Dim vEgresosCantidad() As Double
'      Dim cantEgresos As Integer
'      Dim mJ As Integer
'      Dim mCodTipoVale As String
'      Dim mCodBodega As String
'      Dim mCodUsuarioRet As String
'      Dim mCodUsuarioAut As String
'      Dim mResultado As Boolean
'
'      If fValidaConfirmarConsumo() Then
'         cantEgresos = FlexEgreso.Rows - 2
'
'         ReDim vEgresosCodProducto(0 To cantEgresos - 1) As String
'         ReDim vEgresosCodUbicacion(0 To cantEgresos - 1) As String
'         ReDim vEgresosCantidad(0 To cantEgresos - 1) As Double
'
'
'         For mJ = 2 To FlexEgreso.Rows - 1
'            vEgresosCodProducto(mJ - 2) = FlexEgreso.TextMatrix(mJ, 6)
'            vEgresosCodUbicacion(mJ - 2) = FlexEgreso.TextMatrix(mJ, 7)
'            vEgresosCantidad(mJ - 2) = CDbl(Replace(FlexEgreso.TextMatrix(mJ, 3), ".", ","))
'         Next
'
'         If Option1.Value Then
'            mCodTipoVale = "C"
'         Else
'            mCodTipoVale = "M"
'         End If
'
'         mCodBodega = Left(Combo5.Text, 4)
'         mCodUsuarioRet = Trim(Right(Combo4.Text, 25))
'         mCodUsuarioAut = Trim(Right(Combo2.Text, 25))
'         mResultado = True
'         'OK 'Inserto en Consumo_H ->OK: FALTA TIPOVALE,CODBODETA,USUARIORETIRA,USURIOSIST
'         mObjInven.xInsEgreso vEgresosCodProducto(), vEgresosCodUbicacion(), vEgresosCantidad(), Trim(Text8.Text), mCodTipoVale, mCodBodega, mCodUsuarioRet, mCodUsuarioAut, Trim(Right(MDI.mUser, 15)), mResultado
'         'OK 'Inserto en Consumo_Det
'
'         If mResultado Then
'            MsgBox "El consumo se ha realizado exitosamente", vbInformation, "Consumo"
'            limpioFormulario
'            VerificaStockMin vEgresosCodProducto(), mCodBodega
'
'         End If
'
'
'         'Validado: 'Que se haya completado el campo Numero Vale.
'         'Validado: 'Que el Numero de Vale sea un valor entero.
'         'Validado: 'Que se haya completado el combo "Retirado por:"
'         'Validado: 'Que esten chequeados alguna de las dos radio button (Option1 u Option2)
'         'Validado: 'Que en la grilla inferior "Egresos" exista al menos un registro.
'         'Validado: 'Que no exista el Registro (Nro Vale, TipoVale ) en la tabla Consumo_H
'      End If
'   Else
'      Unload Me
'   End If
'
'End Sub
'
'
'Private Sub VerificaStockMin(ByRef pvEgresosCodProducto() As String, ByVal pCodBodega As String)
'   Dim mI As Integer
'
'   For mI = LBound(pvEgresosCodProducto) To UBound(pvEgresosCodProducto)
'      VerificaStockMinYnotifica pvEgresosCodProducto(mI), pCodBodega
'   Next
'End Sub
'
'Private Sub VerificaStockMinYnotifica(ByVal pCodProducto As String, pCodBodega As String)
'
'   Dim mRec1 As ADODB.Recordset
'   Dim mListaDestinatarios As String
'   Dim mTextoMail As String
'
'   mListaDestinatarios = ""
'   mTextoMail = ""
'
'
'   Set mRec = mObjInven.oEjecutarSelect(" SELECT  M.CodProducto, P.CodigoSap, P.Descripcion AS Producto, U.CodBodega,  B.Descripcion AS Bodega,  SUM(Stock) AS Stock, " & _
'      " IFNULL(SM.Stock_Min, 0) As Stock_Min,  SUM(Stock) - IFNULL(SM.Stock_Min, 0) AS StockMenosStockMin,  Med.Descripcion AS UnidadMedida, IFNULL(N.CodProducto,'XXXXXX') As ProductoNotificado " & _
'      " FROM  " & _
'      " Movimientos2 M " & _
'      "  INNER JOIN " & _
'      " Producto P ON M.CodProducto = P.Codigo " & _
'      "  INNER JOIN " & _
'      " Ubicaciones U ON  M.CodUbicacion = U.Codigo " & _
'      "  INNER JOIN " & _
'      " Bodegas B ON B.Codigo = U.CodBodega  " & _
'      "  INNER JOIN " & _
'      " UnidadMedida Med ON P.CodUnidadMedida = Med.Codigo " & _
'      "  LEFT JOIN " & _
'      " StocksMinimos SM ON SM.CodBodega = B.Codigo AND SM.CodProducto = M.CodProducto " & _
'      "  LEFT JOIN  " & _
'      "  StockMinimo_Notificaciones N ON N.CodProducto = M.CodProducto AND N.CodBodega = B.Codigo " & _
'      " WHERE Fecha = (SELECT MAX(Fecha) " & _
'      "                 From Movimientos2 " & _
'      "                 WHERE CodProducto = M.CodProducto AND CodUbicacion = M.CodUbicacion) " & _
'      " AND M.CodProducto = '" & pCodProducto & "' " & _
'      " AND U.CodBodega = '" & pCodBodega & "' " & _
'      " GROUP BY   M.CodProducto, P.Descripcion,U.CodBodega, B.Descripcion,Med.Descripcion, N.CodProducto;")
'
'      Set mRec1 = mObjInven.oEjecutarSelect(" SELECT DISTINCT P.Email FROM " & _
'                                       "    Usuario_Bodega_Notificacion U " & _
'                                       " INNER JOIN " & _
'                                       "  Personal P ON P.CodUsuario = U.CodUsuario " & _
'                                       " WHERE U.CodBodega = '& pCodBodega &'; ")
'
'   Do While Not mRec1.EOF
'      mListaDestinatarios = mListaDestinatarios & mRec!Email & ";"
'      mRec1.MoveNext
'   Loop
'   mRec1.Close
'
'   'Si estoy por debajo del stock minimo y no se notifico anteriormente entonces enivo correo y cargo flag de notificado.
'   If CDbl(Replace(mRec!StockMenosStockMin, ".", ",")) <= 0 And mRec!ProductoNotificado = "XXXXXX" Then
'      Set mRec1 = mObjInven.oEjecutarSelect(" SELECT DISTINCT P.Email FROM " & _
'                                    "    Usuario_Bodega_Notificacion U " & _
'                                    " INNER JOIN " & _
'                                    "  Personal P ON P.CodUsuario = U.CodUsuario " & _
'                                    " WHERE U.CodBodega = '" & pCodBodega & "'; ")
'
'      Do While Not mRec1.EOF
'         mListaDestinatarios = mListaDestinatarios & mRec1!Email & ";"
'         mRec1.MoveNext
'      Loop
'      mRec1.Close
'
'      mTextoMail = vbCrLf & _
'                  " A continuación se detallan los datos del producto que ha llegado a su Stock Mínimo: " & vbCrLf & _
'                   vbCrLf & _
'                   vbCrLf & _
'                  Space(5) & "Cód. Producto: " & mRec!CodProducto & vbCrLf & _
'                  Space(5) & "Código SAP: " & mRec!CodigoSap & vbCrLf & _
'                  Space(5) & "Producto: " & mRec!Producto & vbCrLf & _
'                  Space(5) & "Bodega: " & mRec!Bodega & vbCrLf & _
'                  Space(5) & "Stock Actual: " & Format(mRec!Stock, "#.00") & " " & mRec!UnidadMedida & vbCrLf & _
'                  Space(5) & "Stock Mínimo: " & Format(mRec!Stock_Min, "#.00") & " " & mRec!UnidadMedida & vbCrLf
'
'
'      If fEnviar_Mail_CDO("", mListaDestinatarios, "ausolmail@ausol.com.ar", " Sist. Global - Inventario: Stock mínimo alcanzó su límite", mTextoMail, "", "587", "ausolmail@ausol.com.ar", "sgedosmildiecisiete1$", True, False) Then
'         mObjInven.xInsStockMinimo_Notificaciones mRec!CodProducto, mRec!CodBodega
'      End If
'   End If
'   mRec.Close
'End Sub
'
'
'Private Sub limpioFormulario()
'   Dim mI As Integer
'
'   Text9.Text = ""
'   Text10.Text = ""
'   Text8.Text = ""
'
'   Option1.Value = False
'   Option2.Value = False
'
'   'Elimino los registros (de la consulta anterior) de la grilla superior
'   For mI = FlexProd.Rows To 3 Step -1
'      FlexProd.RemoveItem mI
'   Next
'
'   mRenglonProducto = 0
'
'   'Elimino los registros de la grilla inferior
'   For mI = FlexEgreso.Rows To 3 Step -1
'      FlexEgreso.RemoveItem mI
'   Next
'
'   mRenglonEgreso = 0
'
'   Combo5.Clear
'   Set mRec = mObjInven.oTabla("Bodegas", " where Fecha_Baja IS NULL and Codigo IN (SELECT CodBodega FROM Usuario_AccesoBodega WHERE CodUsuario = '" & Trim(Right(MDI.mUser, 15)) & "') order by Codigo")
'
'
'   Do While Not mRec.EOF
'      Combo5.AddItem "" & mRec!Codigo & " " & mRec!descripcion & ""
'      mRec.MoveNext
'   Loop
'   mRec.Close
'   cboListIndex = Combo5.ListIndex
'
'   Combo4.Clear
'   Combo2.Clear
'   Combo4.Enabled = False
'   Combo2.Enabled = False
'End Sub
'
'
'
'Private Sub FlexProd_Click()
'   Dim mI As Integer
'
'   If FlexProd.MouseRow > 0 Then
'
'      If mRenglonProducto <> 0 Then
'         FlexProd.Row = mRenglonProducto
'         For mI = 1 To FlexProd.Cols - 1
'            FlexProd.Col = mI
'            FlexProd.CellBackColor = vbWhite
'         Next
'      End If
'
'      mRenglonProducto = FlexProd.MouseRow
'
'      FlexProd.Row = mRenglonProducto
'      For mI = 1 To FlexProd.Cols - 1
'         FlexProd.Col = mI
'         FlexProd.CellBackColor = &H80000003
'      Next
'
'      If mRenglonProducto > 1 Then
'          mCodProducto = FlexProd.TextMatrix(mRenglonProducto, 4)
'      End If
'   Else
'      FlexProd.Row = mRenglonProducto
'      For mI = 1 To FlexProd.Cols - 1
'         FlexProd.Col = mI
'         FlexProd.CellBackColor = vbWhite
'      Next
'      mRenglonProducto = 0
'   End If
'End Sub
'
'Private Sub Form_Load()
'
'   Inven010_frm.Width = 21270
'   Inven010_frm.Height = 13950
'
'   sAlinearForm Me
'
'   Combo4.Enabled = False
'   Combo2.Enabled = False
'
'   'TODO(Realizado): Debe traer solo las bodegas que puede administrar el usuario. Tabla Futura Tabla: Usuarios-Bodegas (o sera mejor hacerlo por Almacen)
'   Set mRec = mObjInven.oTabla("Bodegas", " where Fecha_Baja IS NULL and Codigo IN (SELECT CodBodega FROM Usuario_AccesoBodega WHERE CodUsuario = '" & Trim(Right(MDI.mUser, 15)) & "') order by Codigo")
'
'
'   Do While Not mRec.EOF
'      Combo5.AddItem "" & mRec!Codigo & " " & mRec!descripcion & ""
'      mRec.MoveNext
'   Loop
'   mRec.Close
'
'   FlexProd.ColWidth(0) = 200
'   FlexProd.ColWidth(1) = 10700
'   FlexProd.ColWidth(2) = 4500
'   FlexProd.ColWidth(3) = 1500
'   FlexProd.ColWidth(4) = 1900
'   FlexProd.ColWidth(5) = 1250
'   FlexProd.ColWidth(6) = 0
'   FlexProd.ColWidth(7) = 0
'
'   FlexProd.TextMatrix(0, 1) = "Producto"
'   FlexProd.TextMatrix(0, 2) = "Ubicación"
'   FlexProd.TextMatrix(0, 3) = "Stock"
'   FlexProd.TextMatrix(0, 4) = "Unid.Medida"
'   FlexProd.TextMatrix(0, 5) = "Cód.Sap"
'   FlexProd.TextMatrix(0, 6) = "Cód. Producto"
'   FlexProd.TextMatrix(0, 7) = "Cód. Ubicacion"
'
'   FlexProd.RowHeight(1) = 0
'
'   FlexEgreso.ColWidth(0) = 200
'   FlexEgreso.ColWidth(1) = 10700
'   FlexEgreso.ColWidth(2) = 4500
'   FlexEgreso.ColWidth(3) = 1500
'   FlexEgreso.ColWidth(4) = 1900
'   FlexEgreso.ColWidth(5) = 1250
'   FlexEgreso.ColWidth(6) = 0
'   FlexEgreso.ColWidth(7) = 0
'
'
'   FlexEgreso.TextMatrix(0, 1) = "Producto"
'   FlexEgreso.TextMatrix(0, 2) = "Ubicación"
'   FlexEgreso.TextMatrix(0, 3) = "Cantidad"
'   FlexEgreso.TextMatrix(0, 4) = "Unid.Medida"
'   FlexEgreso.TextMatrix(0, 5) = "Cód.Sap"
'   FlexEgreso.TextMatrix(0, 6) = "Cód. Producto"
'   FlexEgreso.TextMatrix(0, 7) = "Cód. Ubicacion"
'
'   FlexEgreso.RowHeight(1) = 0
'
'   cboListIndex = Combo5.ListIndex
'End Sub
'
'Private Function fValidaEgreso() As Boolean
'   Dim mRet As Boolean
'   Dim mMensajeError As String
'   Dim mJ As Integer
'   Dim mCantidaStock As Double
'   Dim sStock As String
'   Dim iStock As Double
'   Dim mRec1 As New ADODB.Recordset
'   Dim posInstr As Integer
'   Dim qtyDecimales As Integer
'   Dim mCodTipoVale As String
'
'   mRet = True
'
'   If Trim(Text8.Text) = "" Then
'      mRet = False
'      mMensajeError = "Debe completar el Número de Vale"
'   End If
'
'   If mRet Then
'      If Not IsNumeric(Trim(Text8.Text)) Then
'         mRet = False
'         mMensajeError = "El Nro. Vale debe ser numérico !!"
'      End If
'   End If
'
'
'   If mRet Then
'      If Len(Trim(Text8.Text)) <> 9 Then
'         mRet = False
'         mMensajeError = "El Nro. Vale debe tener 9 caracteres numéricos !!"
'      End If
'   End If
'
'   If mRet Then
'      If ((Not Option1.Value) And (Not Option2.Value)) Then
'         mRet = False
'         mMensajeError = "Debe completar el Tipo de Vale"
'      End If
'   End If
'
'
'   If mRet Then
'      If Option1.Value Then
'         mCodTipoVale = "C"
'      Else
'         mCodTipoVale = "M"
'      End If
'
'      Set mRec1 = mObjInven.oEjecutarSelect("SELECT NroVale,CodTipoVale FROM Consumos_H WHERE NroVale = " & Trim(Text8.Text) & " and CodTipoVale = '" & mCodTipoVale & "'; ")
'      If Not mRec1.EOF Then
'         mRet = False
'         mMensajeError = "Ya se han realizado consumos para ese Número y Tipo de Vale !!!"
'      End If
'      mRec1.Close
'   End If
'
'   If mRet Then
'      If mRenglonProducto = 0 Then
'         mRet = False
'         mMensajeError = "Debe seleccionar un Producto de la Grilla superior"
'      End If
'   End If
'
'   If mRet Then
'      If mRenglonProducto <> 0 And FlexProd.TextMatrix(mRenglonProducto, 1) = "" Then
'         mRet = False
'         mMensajeError = "Debe seleccionar un Producto de la Grilla superior"
'      End If
'   End If
'
'   If mRet Then
'      If Trim(Text10.Text) = "" Then
'         mRet = False
'         mMensajeError = "Debe completar el campo: 'Cantidad'. "
'      End If
'   End If
'
'   If mRet Then
'      If Not IsNumeric(Replace(Text10.Text, ".", ",")) Then
'         mRet = False
'         mMensajeError = "La Cantidad ingresada no es un valor numérico"
'      End If
'   End If
'
'   If mRet Then
'      If CDbl(Replace(Trim(Text10.Text), ".", ",")) <= 0 Then
'         mRet = False
'         mMensajeError = "La Cantidad ingresada no puede ser menor o igual a cero."
'      End If
'   End If
'
'   'Valido que no supere los 2 digitos decimales
'   If mRet Then
'      posInstr = InStr(1, Replace(Trim(Text10.Text), ".", ","), ",")
'
'      If posInstr <> 0 Then
'         qtyDecimales = Len(Right(Trim(Text10.Text), Len(Trim(Text10.Text)) - posInstr))
'      End If
'
'      If qtyDecimales > 2 Then
'         mRet = False
'         mMensajeError = "El campo 'Cantidad' solo admite hasta dos dígitos decimales."
'      End If
'   End If
'
'   'Valido si el Producto y la Ubicacion seleccionados coinciden con un Egreso ya registrado
'   If mRet Then
'      For mJ = 2 To FlexEgreso.Rows - 1
'         If FlexEgreso.TextMatrix(mJ, 6) = FlexProd.TextMatrix(mRenglonProducto, 6) And FlexEgreso.TextMatrix(mJ, 7) = FlexProd.TextMatrix(mRenglonProducto, 7) Then
'            mRet = False
'            mMensajeError = "El Producto y la Ubicación elegidos ya han sido seleccionados"
'            mJ = 999
'         End If
'      Next
'   End If
'
'   'Valido si el saldo del stock es insuficiente para ese Producto/Ubicación
'   If mRet Then
'
'      Set mRec1 = mObjInven.oEjecutarSelect("SELECT Fecha,CodProducto,CodUbicacion, IFNULL(Stock,0) As Stock " & _
'                                                " FROM Movimientos2 M " & _
'                                                " WHERE CodProducto  = '" & FlexProd.TextMatrix(mRenglonProducto, 6) & "' and CodUbicacion = '" & FlexProd.TextMatrix(mRenglonProducto, 7) & "'" & _
'                                                " AND Fecha = (SELECT Max(Fecha) FROM Movimientos2 WHERE CodProducto = M.CodProducto AND CodUbicacion = M.CodUbicacion)")
'      If Not mRec1.EOF Then
'         iStock = mRec1!Stock
'      Else
'         iStock = 0
'      End If
'      mRec1.Close
'
'      If CDbl(Replace(Trim(Text10.Text), ".", ",")) > iStock Then
'         mRet = False
'         mMensajeError = "El stock es insuficiente para ese Producto en esa Ubicación"
'      End If
'   End If
'
'   If Not mRet Then
'         MsgBox mMensajeError, vbCritical, "Atención"
'   End If
'   fValidaEgreso = mRet
'End Function
'
'Private Function fValidaConfirmarConsumo() As Boolean
'
'  'Validado: 'Que se haya completado el campo Numero Vale.
'  'Validado: 'Que el Numero de Vale sea un valor entero.
'  'Validado: 'Que se haya completado el combo "Retirado por:"
'  'Validado:  'Que esten chequeados alguna de las dos radio button (Option1 u Option2)
'  'Validado: 'Que en la grilla inferior "Egresos" exista al menos un registro.
'  'Validado: 'Que no exista el Registro (Nro Vale, TipoVale ) en la tabla Consumo_H
'
'   Dim mRet As Boolean
'   Dim mMensajeError As String
'   Dim mCodTipoVale As String
'   Dim mRec1 As New ADODB.Recordset
'
'   mRet = True
'
'   If Trim(Text8.Text) = "" Then
'      mRet = False
'      mMensajeError = "Debe completar el Número de Vale"
'   End If
'
'   If mRet Then
'      If Not IsNumeric(Trim(Text8.Text)) Then
'         mRet = False
'         mMensajeError = "El Nro. Vale debe ser numérico !!"
'      End If
'   End If
'
'   If mRet Then
'      If Len(Trim(Text8.Text)) <> 9 Then
'         mRet = False
'         mMensajeError = "El Nro. Vale debe tener 9 caracteres numéricos !!"
'      End If
'   End If
'
'
'   If mRet Then
'      If ((Not Option1.Value) And (Not Option2.Value)) Then
'         mRet = False
'         mMensajeError = "Debe completar el Tipo de Vale"
'      End If
'   End If
'
'
'   If mRet Then
'      If Trim(Right(Combo4.Text, 25)) = "" Then
'         mRet = False
'         mMensajeError = "Debe completar el campo: 'Retirado por'"
'      End If
'   End If
'
'   If mRet Then
'      If Trim(Right(Combo2.Text, 25)) = "" Then
'         mRet = False
'         mMensajeError = "Debe completar el campo: 'Autorizado por'"
'      End If
'   End If
'
'
'   If mRet Then
'      If FlexEgreso.Rows <= 2 Then
'         mRet = False
'         mMensajeError = "Al menos debe existir un registro en la Grilla Egresos"
'      End If
'   End If
'
'   If mRet Then
'      If Option1.Value Then
'         mCodTipoVale = "C"
'      Else
'         mCodTipoVale = "M"
'      End If
'
'      Set mRec1 = mObjInven.oEjecutarSelect("SELECT NroVale,CodTipoVale FROM Consumos_H WHERE NroVale = " & Trim(Text8.Text) & " and CodTipoVale = '" & mCodTipoVale & "'; ")
'      If Not mRec1.EOF Then
'         mRet = False
'         mMensajeError = "Ya se han realizado consumos para ese Número y Tipo de Vale !!!"
'      End If
'      mRec1.Close
'   End If
'
'   If Not mRet Then
'         MsgBox mMensajeError, vbCritical, "Atención"
'   End If
'
'   fValidaConfirmarConsumo = mRet
'End Function
'
'Private Sub Form_Unload(Cancel As Integer)
'   Set mObjInven = Nothing
'   Set mRec = Nothing
'   ShowMenu 12, True, False
'End Sub
'
'Private Sub text10_KeyPress(KeyAscii As Integer)
'      If KeyAscii <> 46 Then
'         KeyAscii = fNumeroKeyPress(KeyAscii)
'      End If
'End Sub
'
'Private Sub text8_KeyPress(KeyAscii As Integer)
'         KeyAscii = fNumeroKeyPress(KeyAscii)
'End Sub
'
'Private Sub FlexEgreso_Click()
'   Dim mI As Integer
'
'   If FlexEgreso.MouseRow > 0 Then
'
'      If mRenglonEgreso <> 0 Then
'         If FlexEgreso.Rows > mRenglonEgreso Then
'            FlexEgreso.Row = mRenglonEgreso
'            For mI = 1 To FlexEgreso.Cols - 1
'               FlexEgreso.Col = mI
'               FlexEgreso.CellBackColor = vbWhite
'            Next
'         End If
'      End If
'
'      mRenglonEgreso = FlexEgreso.MouseRow
'
'      FlexEgreso.Row = mRenglonEgreso
'      For mI = 1 To FlexEgreso.Cols - 1
'         FlexEgreso.Col = mI
'         FlexEgreso.CellBackColor = &H80000003
'      Next
'
'      If mRenglonEgreso > 1 Then
'          mCodProducto = FlexEgreso.TextMatrix(mRenglonEgreso, 4)
'      End If
'   Else
'      FlexEgreso.Row = mRenglonEgreso
'      For mI = 1 To FlexProd.Cols - 1
'         FlexEgreso.Col = mI
'         FlexEgreso.CellBackColor = vbWhite
'      Next
'      mRenglonEgreso = 0
'   End If
'End Sub


