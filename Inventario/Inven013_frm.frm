VERSION 5.00
Begin VB.Form Inven013_frm 
   Caption         =   "Reporte de Stock"
   ClientHeight    =   2790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5310
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   2790
   ScaleWidth      =   5310
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   375
      Index           =   1
      Left            =   3240
      TabIndex        =   3
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   550
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Bodega:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "Inven013_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mObj As New clInven
Dim mRec As ADODB.Recordset
Dim XLS As EXCEL.Application
Dim mi As Integer
Dim mFechaEjec As Date

Private Sub Command1_Click(Index As Integer)
   If Index = 0 Then
      If Combo1.ListIndex <> -1 Then
   
         sMsgEspere Me, "Procesando datos...", True
         mFechaEjec = Now()
         
         Set XLS = CreateObject("Excel.Application")
         
         sPlanilla1
         sPlanilla2
         sPlanilla3
         
         XLS.Worksheets(1).Select
         
         sMsgEspere Me, "", False
         XLS.Application.Visible = True
         
      Else
         MsgBox "Debe seleccionar una Bodega.", vbExclamation, "Reporte de Stock."
      End If
   Else
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   Me.Width = 5430
   Me.Height = 3300
   sAlinearForm Me
   
   Set mRec = mObj.oTabla("Bodegas", " where Fecha_Baja IS NULL and Codigo IN (SELECT CodBodega FROM Usuario_AccesoBodega WHERE CodUsuario = '" & Trim(Right(MDI.mUser, 15)) & "') order by Codigo")
   
   Do While Not mRec.EOF
      Combo1.AddItem "" & mRec!Codigo & " " & mRec!descripcion & ""
      mRec.MoveNext
   Loop
   mRec.Close
End Sub

Private Sub sPlanilla1()
   mi = 10
   sCabecera1
   
   Set mRec = mObj.oEjecutarSelect(" SELECT  CodProducto,CodigoSAP, P.Descripcion AS Producto, CodBodega, B.Descripcion AS Bodega, SUM(Stock) AS Stock, Med.Descripcion AS UnidadMedida " & _
   "FROM  " & _
   " Movimientos2 M " & _
   "  INNER JOIN " & _
   " Producto P ON M.CodProducto = P.Codigo " & _
   "  INNER JOIN " & _
   " Ubicaciones U ON  M.CodUbicacion = U.Codigo AND U.CodBodega = '" & Left(Combo1.Text, 4) & "' " & _
   "  INNER JOIN " & _
   " Bodegas B ON B.Codigo = U.CodBodega  " & _
   "  INNER JOIN " & _
   " UnidadMedida Med ON P.CodUnidadMedida = Med.Codigo " & _
   " WHERE Fecha = (SELECT MAX(Fecha) " & _
   "                 From Movimientos2 " & _
   "                 WHERE CodProducto = M.CodProducto AND CodUbicacion = M.CodUbicacion) " & _
   " GROUP BY   CodProducto, P.Descripcion,CodBodega, B.Descripcion,Med.Descripcion " & _
   " ORDER BY   P.Descripcion ;")
   
   Do While Not mRec.EOF
      With XLS
      
      .Cells(mi, 1).Formula = NVL(mRec!CodProducto, "")
      .Cells(mi, 2).Formula = NVL(mRec!CodigoSap, "")
      .Cells(mi, 3).Formula = NVL(mRec!Producto, "")
      .Cells(mi, 4).Formula = NVL(mRec!CodBodega, "")
      .Cells(mi, 5).Formula = NVL(mRec!Bodega, "")
      .Cells(mi, 6).Formula = NVL(mRec!Stock, "")
      .Cells(mi, 7).Formula = NVL(mRec!UnidadMedida, "")
     
      End With
      mRec.MoveNext
      mi = mi + 1
   Loop
   mRec.Close
End Sub

Private Sub sPlanilla2()
   mi = 10
   sCabecera2
   
   Set mRec = mObj.oEjecutarSelect(" SELECT  M.CodProducto,P.CodigoSap,  P.Descripcion AS Producto, U.CodBodega,  B.Descripcion AS Bodega,  SUM(Stock) AS Stock, " & _
   " IFNULL(SM.Stock_Min, 0) As Stock_Min,  SUM(Stock) - IFNULL(SM.Stock_Min, 0) AS StockMenosStockMin,  Med.Descripcion AS UnidadMedida " & _
   "FROM  " & _
   " Movimientos2 M " & _
   "  INNER JOIN " & _
   " Producto P ON M.CodProducto = P.Codigo " & _
   "  INNER JOIN " & _
   " Ubicaciones U ON  M.CodUbicacion = U.Codigo AND U.CodBodega = '" & Left(Combo1.Text, 4) & "' " & _
   "  INNER JOIN " & _
   " Bodegas B ON B.Codigo = U.CodBodega  " & _
   "  INNER JOIN " & _
   " UnidadMedida Med ON P.CodUnidadMedida = Med.Codigo " & _
   "  LEFT JOIN " & _
   "  StocksMinimos SM ON SM.CodBodega = B.Codigo AND SM.CodProducto = M.CodProducto " & _
   " WHERE Fecha = (SELECT MAX(Fecha) " & _
   "                 From Movimientos2 " & _
   "                 WHERE CodProducto = M.CodProducto AND CodUbicacion = M.CodUbicacion) " & _
   " GROUP BY   M.CodProducto, P.Descripcion,U.CodBodega, B.Descripcion,Med.Descripcion " & _
   " ORDER BY   P.Descripcion;")
   
   Do While Not mRec.EOF
      With XLS
      
      .Cells(mi, 1).Formula = NVL(mRec!CodProducto, "")
      .Cells(mi, 2).Formula = NVL(mRec!CodigoSap, "")
      .Cells(mi, 3).Formula = NVL(mRec!Producto, "")
      .Cells(mi, 4).Formula = NVL(mRec!CodBodega, "")
      .Cells(mi, 5).Formula = NVL(mRec!Bodega, "")
      .Cells(mi, 6).Formula = NVL(mRec!Stock, "")
      .Cells(mi, 7).Formula = NVL(mRec!Stock_Min, "")
      .Cells(mi, 8).Formula = NVL(mRec!StockMenosStockMin, "")
      .Cells(mi, 9).Formula = NVL(mRec!UnidadMedida, "")
     
      End With
      mRec.MoveNext
      mi = mi + 1
   Loop
   mRec.Close
End Sub

Private Sub sPlanilla3()
   mi = 10
   sCabecera3
   
   Set mRec = mObj.oEjecutarSelect(" SELECT  CodProducto,P.CodigoSAP, P.Descripcion AS Producto, CodUbicacion, U.Descripcion AS Ubicacion, Stock, Med.Descripcion AS UnidadMedida " & _
   "FROM  " & _
   " Movimientos2 M " & _
   "  INNER JOIN " & _
   " Producto P ON M.CodProducto = P.Codigo " & _
   "  INNER JOIN " & _
   " Ubicaciones U ON  M.CodUbicacion = U.Codigo AND U.CodBodega = '" & Left(Combo1.Text, 4) & "' " & _
   "  INNER JOIN " & _
   " UnidadMedida Med ON P.CodUnidadMedida = Med.Codigo " & _
   " WHERE Fecha = (SELECT MAX(Fecha) " & _
   "                 From Movimientos2 " & _
   "                 WHERE CodProducto = M.CodProducto AND CodUbicacion = M.CodUbicacion) " & _
   " ORDER BY P.Descripcion, U.Descripcion ")
   
   Do While Not mRec.EOF
      With XLS
      
      .Cells(mi, 1).Formula = NVL(mRec!CodProducto, "")
      .Cells(mi, 2).Formula = NVL(mRec!CodigoSap, "")
      .Cells(mi, 3).Formula = NVL(mRec!Producto, "")
      .Cells(mi, 4).Formula = NVL(mRec!CodUbicacion, "")
      .Cells(mi, 5).Formula = NVL(mRec!Ubicacion, "")
      .Cells(mi, 6).Formula = NVL(mRec!Stock, "")
      .Cells(mi, 7).Formula = NVL(mRec!UnidadMedida, "")
     
      End With
      mRec.MoveNext
      mi = mi + 1
   Loop
   mRec.Close
End Sub

Private Sub sCabecera1()
   Dim sAlmacen As String

   With XLS
      .WorkBooks.Add
      .Worksheets(1).Select
      .Worksheets(1).Name = "Stock por Bodega"
      .Columns("A:A").ColumnWidth = 15 ' CodProducto
      .Columns("B:B").ColumnWidth = 15 ' CodigoSAP
      .Columns("C:C").ColumnWidth = 70 ' Producto
      .Columns("D:D").ColumnWidth = 15 ' CodBodega
      .Columns("E:E").ColumnWidth = 30 ' Bodega
      .Columns("F:F").ColumnWidth = 15 ' Stock
      .Columns("G:G").ColumnWidth = 30 ' Unidad Medida
         
      'Formateo el Cod.Producto: 00000X
      .Range("A10:A65536").Select
      .Selection.NumberFormat = "000000"
      
      'Formateo el Cod.Bodega: 000X
      .Range("D10:D65536").Select
      .Selection.NumberFormat = "0000"
      
      'Formateo: Encabezdo con Negrita.
      .Range("A9:G9").Select
      .Selection.Font.Bold = True
      .Selection.Interior.ColorIndex = 15
      
      With .Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With
      
      With .Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With
      
      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With
      
      With .Selection.Borders(xlRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With
      

      .Cells(1, 1).Formula = "AUTOPISTAS DEL SOL S.A."
      .Cells(3, 1).Formula = "REPORTE: STOCKS POR BODEGA"
      
      .Range("A1:A1").Select
      .Selection.Font.ColorIndex = 5
      .Selection.Font.Size = 14
      .Selection.Font.Bold = True
      
      .Range("A3:A3").Select
      .Selection.Font.Size = 12
      .Selection.Font.Bold = True
      
      .Range("A9:A9").Select
      
     
      Set mRec = mObj.oEjecutarSelect(" SELECT CONCAT(A.Codigo, ' - ', A.Descripcion) AS Almacen FROM Bodegas B " & _
                                  " Inner Join " & _
                                  " Almacenes A ON B.CodAlmacen = A.Codigo " & _
                                  " where B.Codigo = '" & Left(Combo1.Text, 4) & " ';")
                                  
      sAlmacen = mRec!Almacen
      mRec.Close

      .Cells(5, 1).Formula = "Almacén: " & sAlmacen
      .Cells(6, 1).Formula = "Bodega: " & Left(Combo1.Text, 4) & " - " & Right(Combo1.Text, Len(Combo1.Text) - 4)
      .Cells(7, 1).Formula = "Fecha ejecución del Reporte: " & mFechaEjec
      
      .Cells(9, 1).Formula = "Cód.Producto"
      .Cells(9, 2).Formula = "Código SAP"
      .Cells(9, 3).Formula = "Producto"
      .Cells(9, 4).Formula = "Cód.Bodega"
      .Cells(9, 5).Formula = "Bodega"
      .Cells(9, 6).Formula = "Stock"
      .Cells(9, 7).Formula = "Unidad Medida"
   End With
End Sub

Private Sub sCabecera2()
   Dim sAlmacen As String

   With XLS
      '.WorkBooks.Add
      .Worksheets(2).Select
      .Worksheets(2).Name = "Stock Vs. Stock Min."
      .Columns("A:A").ColumnWidth = 15 ' CodProducto
      .Columns("B:B").ColumnWidth = 15 ' CodigoSap
      .Columns("C:C").ColumnWidth = 70 ' Producto
      .Columns("D:D").ColumnWidth = 15 ' CodBodega
      .Columns("E:E").ColumnWidth = 30 ' Bodega
      .Columns("F:F").ColumnWidth = 15 ' Stock
      .Columns("G:G").ColumnWidth = 15 ' Stock Minimo
      .Columns("H:H").ColumnWidth = 18 ' Stock - Stock Min.
      .Columns("I:I").ColumnWidth = 30 ' Unidad Medida
         
      'Formateo el Cod.Producto: 00000X
      .Range("A10:A65536").Select
      .Selection.NumberFormat = "000000"
      
      'Formateo el Cod.Bodega: 000X
      .Range("D10:D65536").Select
      .Selection.NumberFormat = "0000"
      
      'Formateo: Encabezdo con Negrita.
      .Range("A9:I9").Select
      .Selection.Font.Bold = True
      .Selection.Interior.ColorIndex = 15
      
      With .Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With
      
      With .Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With
      
      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With
      
      With .Selection.Borders(xlRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With
      
      .Cells(1, 1).Formula = "AUTOPISTAS DEL SOL S.A."
      .Cells(3, 1).Formula = "REPORTE: STOCK vs. STOCK MÍNIMO"
      
      
      .Range("A1:A1").Select
      .Selection.Font.ColorIndex = 5
      .Selection.Font.Size = 14
      .Selection.Font.Bold = True
      
      .Range("A3:A3").Select
      .Selection.Font.Size = 12
      .Selection.Font.Bold = True
      
      .Range("A9:A9").Select
      
      
      Set mRec = mObj.oEjecutarSelect(" SELECT CONCAT(A.Codigo, ' - ', A.Descripcion) AS Almacen FROM Bodegas B " & _
                                  " Inner Join " & _
                                  " Almacenes A ON B.CodAlmacen = A.Codigo " & _
                                  " where B.Codigo = '" & Left(Combo1.Text, 4) & " ';")
                                  
      sAlmacen = mRec!Almacen
      mRec.Close

      .Cells(5, 1).Formula = "Almacén: " & sAlmacen
      .Cells(6, 1).Formula = "Bodega: " & Left(Combo1.Text, 4) & " - " & Right(Combo1.Text, Len(Combo1.Text) - 4)
      .Cells(7, 1).Formula = "Fecha ejecución del Reporte: " & mFechaEjec
      
      .Cells(9, 1).Formula = "Cód.Producto"
      .Cells(9, 2).Formula = "Código SAP"
      .Cells(9, 3).Formula = "Producto"
      .Cells(9, 4).Formula = "Cód.Bodega"
      .Cells(9, 5).Formula = "Bodega"
      .Cells(9, 6).Formula = "Stock"
      .Cells(9, 7).Formula = "Stock Mín."
      .Cells(9, 8).Formula = "Stock - Stock Mín."
      .Cells(9, 9).Formula = "Unidad Medida"
   End With
End Sub

Private Sub sCabecera3()
   Dim sAlmacen As String

   With XLS
      '.WorkBooks.Add
      .Worksheets(3).Select
      .Worksheets(3).Name = "Stock por Ubicación"
      .Columns("A:A").ColumnWidth = 15 ' CodProducto
      .Columns("B:B").ColumnWidth = 15 ' CodigoSAP
      .Columns("C:C").ColumnWidth = 70 ' Producto
      .Columns("D:D").ColumnWidth = 15 ' CodUbicacion
      .Columns("E:E").ColumnWidth = 30 ' Ubicacion
      .Columns("F:F").ColumnWidth = 15 ' Stock
      .Columns("G:G").ColumnWidth = 30 ' Unidad Medida
         
      'Formateo el Cod.Producto: 00000X
      .Range("A10:A65536").Select
      .Selection.NumberFormat = "000000"
      
      'Formateo el Cod.Ubicacion: 000X
      .Range("D10:D65536").Select
      .Selection.NumberFormat = "0000"
      
      'Formateo: Encabezdo con Negrita.
      .Range("A9:G9").Select
      .Selection.Font.Bold = True
      .Selection.Interior.ColorIndex = 15
      
      With .Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With
      
      With .Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With
      
      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With
      
      With .Selection.Borders(xlRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With
      
      .Cells(1, 1).Formula = "AUTOPISTAS DEL SOL S.A."
      .Cells(3, 1).Formula = "REPORTE: STOCK POR UBICACION"
      
      .Range("A1:A1").Select
      .Selection.Font.ColorIndex = 5
      .Selection.Font.Size = 14
      .Selection.Font.Bold = True
      
      .Range("A3:A3").Select
      .Selection.Font.Size = 12
      .Selection.Font.Bold = True
      
      .Range("A9:A9").Select
      
      
     Set mRec = mObj.oEjecutarSelect(" SELECT CONCAT(A.Codigo, ' - ', A.Descripcion) AS Almacen FROM Bodegas B " & _
                                  " Inner Join " & _
                                  " Almacenes A ON B.CodAlmacen = A.Codigo " & _
                                  " where B.Codigo = '" & Left(Combo1.Text, 4) & " ';")
                                  
      sAlmacen = mRec!Almacen
      mRec.Close

      .Cells(5, 1).Formula = "Almacén: " & sAlmacen
      .Cells(6, 1).Formula = "Bodega: " & Left(Combo1.Text, 4) & " - " & Right(Combo1.Text, Len(Combo1.Text) - 4)
      .Cells(7, 1).Formula = "Fecha ejecución del Reporte: " & mFechaEjec
      
      .Cells(9, 1).Formula = "Cód.Producto"
      .Cells(9, 2).Formula = "Código SAP"
      .Cells(9, 3).Formula = "Producto"
      .Cells(9, 4).Formula = "Cód.Ubicación"
      .Cells(9, 5).Formula = "Ubicación"
      .Cells(9, 6).Formula = "Stock"
      .Cells(9, 7).Formula = "Unidad Medida"
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set mObj = Nothing
   Set mRec = Nothing
   Set XLS = Nothing
   ShowMenu 12, True, False
End Sub

Public Sub StockDebajoDelMinimo()
         'sMsgEspere Me, "Procesando datos...", True
         mFechaEjec = Now()
         
         Set XLS = CreateObject("Excel.Application")
         
         sPlanillaDebajoStockMinimo

         XLS.Worksheets(1).Select
         
         'sMsgEspere Me, "", False
         XLS.Application.Visible = True

         Unload Me
         ShowMenu 12, True, False
End Sub

Private Sub sPlanillaDebajoStockMinimo()
   mi = 8
   sCabeceraDebajoStockMinimo
   
   Set mRec = mObj.oEjecutarSelect("  SELECT A.Codigo As CodAlmacen, A.Descripcion AS Almacen, " & _
   " U.CodBodega,  B.Descripcion AS Bodega, " & _
   " M.CodProducto,P.CodigoSap, P.Descripcion AS Producto, " & _
   "   SUM(Stock) AS Stock, " & _
   " IFNULL(SM.Stock_Min, 0) As Stock_Min, " & _
   " SUM(Stock) - IFNULL(SM.Stock_Min, 0) AS StockMenosStockMin, " & _
   " Med.Descripcion AS UnidadMedida " & _
   " FROM " & _
   "  Movimientos2 M " & _
   " INNER JOIN Producto P ON M.CodProducto = P.Codigo " & _
   " INNER JOIN Ubicaciones U ON  M.CodUbicacion = U.Codigo " & _
   " INNER JOIN Bodegas B ON B.Codigo = U.CodBodega " & _
   " INNER JOIN Almacenes A ON A.Codigo = B.CodAlmacen " & _
   " INNER JOIN UnidadMedida Med ON P.CodUnidadMedida = Med.Codigo " & _
   " LEFT JOIN StocksMinimos SM ON SM.CodBodega = B.Codigo AND SM.CodProducto = M.CodProducto " & _
   " WHERE Fecha = (SELECT MAX(Fecha) From Movimientos2 WHERE CodProducto = M.CodProducto AND CodUbicacion = M.CodUbicacion) " & _
   " AND U.CodBodega IN (SELECT CodBodega FROM Usuario_AccesoBodega WHERE CodUsuario = '" & Trim(Right(MDI.mUser, 15)) & "')" & _
   " GROUP BY  A.Codigo,  A.Descripcion,U.CodBodega, B.Descripcion, M.CodProducto, P.Descripcion,Med.Descripcion " & _
   " HAVING SUM(Stock) < IFNULL(Stock_Min, 0) " & _
   " ORDER BY   B.Descripcion, P.Descripcion;")

   Do While Not mRec.EOF
      With XLS
      
      .Cells(mi, 1).Formula = NVL(mRec!Almacen, "")
      .Cells(mi, 2).Formula = NVL(mRec!Bodega, "")
      .Cells(mi, 3).Formula = NVL(mRec!CodProducto, "")
      .Cells(mi, 4).Formula = NVL(mRec!CodigoSap, "")
      .Cells(mi, 5).Formula = NVL(mRec!Producto, "")
      .Cells(mi, 6).Formula = NVL(mRec!Stock, "")
      .Cells(mi, 7).Formula = NVL(mRec!Stock_Min, "")
      .Cells(mi, 8).Formula = NVL(mRec!StockMenosStockMin, "")
      .Cells(mi, 9).Formula = NVL(mRec!UnidadMedida, "")
     
      End With
      mRec.MoveNext
      mi = mi + 1
   Loop
   mRec.Close
End Sub

Private Sub sCabeceraDebajoStockMinimo()
   
   With XLS
      .WorkBooks.Add
      .Worksheets(1).Select
      .Worksheets(1).Name = "Stock menor Minimo."
      .Columns("A:A").ColumnWidth = 30 ' Almacen
      .Columns("B:B").ColumnWidth = 30 ' Bodega
      .Columns("C:C").ColumnWidth = 15 ' CodProducto
      .Columns("D:D").ColumnWidth = 15 ' CodigoSap
      .Columns("E:E").ColumnWidth = 70 ' Producto
      .Columns("F:F").ColumnWidth = 15 ' Stock
      .Columns("G:G").ColumnWidth = 15 ' Stock Minimo
      .Columns("H:H").ColumnWidth = 18 ' Stock - Stock Min.
      .Columns("I:I").ColumnWidth = 30 ' Unidad Medida
         
      'Formateo el Cod.Producto: 00000X
      .Range("C8:C65536").Select
      .Selection.NumberFormat = "000000"
      
      'Formateo: Encabezdo con Negrita.
      .Range("A7:I7").Select
      .Selection.Font.Bold = True
      .Selection.Interior.ColorIndex = 15
      
      With .Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With
      
      With .Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With
      
      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With
      
      With .Selection.Borders(xlRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      .Cells(1, 1).Formula = "AUTOPISTAS DEL SOL S.A."
      .Cells(3, 1).Formula = "REPORTE: STOCK POR DEBAJO DEL STOCK MÍNIMO"
      .Cells(5, 1).Formula = "Fecha ejecución del Reporte: " & mFechaEjec
      
      .Range("A1:A1").Select
      .Selection.Font.ColorIndex = 5
      .Selection.Font.Size = 14
      .Selection.Font.Bold = True
      
      .Range("A3:A3").Select
      .Selection.Font.Size = 12
      .Selection.Font.Bold = True
      
      .Range("A7:A7").Select
      
      .Cells(7, 1).Formula = "Almacén"
      .Cells(7, 2).Formula = "Bodega"
      .Cells(7, 3).Formula = "Cód.Producto"
      .Cells(7, 4).Formula = "Código SAP"
      .Cells(7, 5).Formula = "Producto"
      .Cells(7, 6).Formula = "Stock"
      .Cells(7, 7).Formula = "Stock Mín."
      .Cells(7, 8).Formula = "Stock - Stock Mín."
      .Cells(7, 9).Formula = "Unidad Medida"
   End With
End Sub
