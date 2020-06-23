VERSION 5.00
Begin VB.Form Inven012_frm 
   Caption         =   "Movimientos"
   ClientHeight    =   2790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4455
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   2790
   ScaleWidth      =   4455
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   2880
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1320
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   1
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Fechas:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
End
Attribute VB_Name = "Inven012_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mObj As New clInven
Dim mRec As New ADODB.Recordset
Dim XLS As EXCEL.Application
Dim mI As Integer
Dim mFechaEjec As Date
Public mReporte As String


Private Sub Command1_Click(Index As Integer)
   If Index = 0 Then
      If Fecha_ok(Text1(0).Text) And Fecha_ok(Text1(1).Text) Then
         If DateDiff("d", Text1(0).Text, Text1(1).Text) >= 0 Then
            sMsgEspere Me, "Procesando datos...", True
            mFechaEjec = Now()
            
            Set XLS = CreateObject("Excel.Application")
            
            If mReporte = "Movimientos" Then
               sPlanilla
            Else
               sPlanillaAjustes
            End If
            
            sMsgEspere Me, "", False
            XLS.Application.Visible = True
         Else
            MsgBox "Fecha Inicial mayor a la Final", vbCritical, sMessage
         End If
      End If
   Else
      Unload Me
   End If
End Sub

Private Sub sPlanilla()

   mI = 9
   sCabecera
   
   Set mRec = mObj.oEjecutarSelect("SELECT M.IdMov, M.Fecha, CASE WHEN M.CodTipoMovimiento = 'I' THEN 'Ingreso' ELSE 'Egreso' END AS CodTipoMovimiento , M.CodProducto,P.CodigoSap, P.Descripcion As DescripcionProducto, " & _
   " M.CodUbicacion, U.Descripcion As DescripiconUbicacion, U.CodBodega, B.Descripcion AS DescripcionBodega, " & _
   " B.CodAlmacen, A.Descripcion AS DescripcionAlmacen, M.Cantidad, M.CodUsuario " & _
   " FROM Movimientos2 M " & _
   " INNER JOIN " & _
   "   Producto P ON M.CodProducto = P.Codigo " & _
   " INNER JOIN " & _
   "  Ubicaciones U ON M.CodUbicacion = U.Codigo " & _
   " INNER JOIN " & _
   "  Bodegas B ON U.CodBodega = B.Codigo " & _
   " INNER JOIN " & _
   "  Almacenes A ON B.CodAlmacen = A.Codigo " & _
   "  Where Fecha between '" & Format(Text1(0).Text, "yyyy-mm-dd") & " 00:00:00' and '" & Format(Text1(1).Text, "yyyy-mm-dd") & " 23:59:59 '" & _
   "  AND B.Codigo IN (SELECT CodBodega FROM Usuario_AccesoBodega WHERE CodUsuario = '" & Trim(Right(MDI.mUser, 15)) & "') " & _
   "  ORDER BY Fecha, IdMov; ")
   
   Do While Not mRec.EOF
      With XLS
         '.Cells(mi, 1).Formula = NVL(mRec!IdMov, "")
         .Cells(mI, 2).Formula = NVL(mRec!Fecha, "")
         .Cells(mI, 3).Formula = NVL(mRec!CodTipoMovimiento, "")
         .Cells(mI, 4).Formula = NVL(mRec!CodProducto, "")
         .Cells(mI, 5).Formula = NVL(mRec!CodigoSap, "")
         .Cells(mI, 6).Formula = NVL(mRec!DescripcionProducto, "")
         .Cells(mI, 7).Formula = NVL(mRec!DescripcionAlmacen, "")
         .Cells(mI, 8).Formula = NVL(mRec!DescripcionBodega, "")
         .Cells(mI, 9).Formula = NVL(mRec!DescripiconUbicacion, "")
         .Cells(mI, 10).Formula = NVL(mRec!Cantidad, "")
         .Cells(mI, 11).Formula = NVL(mRec!CodUsuario, "")
      End With
      mRec.MoveNext
      mI = mI + 1
   Loop
   mRec.Close
End Sub

Private Sub sCabecera()
   With XLS
      .WorkBooks.Add
      .Worksheets(1).Select
      .Worksheets(1).Name = "Movimientos"
      '.Columns("A:A").ColumnWidth = 15 ' NumMovimiento
      .Columns("B:B").ColumnWidth = 16 ' Fecha
      .Columns("C:C").ColumnWidth = 15 ' Tipo Movimiento
      .Columns("D:D").ColumnWidth = 15 ' Cod. Producto
      .Columns("E:E").ColumnWidth = 15 ' Codigo Sap
      .Columns("F:F").ColumnWidth = 70 ' Producto
      .Columns("G:G").ColumnWidth = 25 ' Almacen
      .Columns("H:H").ColumnWidth = 25 ' Bodega
      .Columns("I:I").ColumnWidth = 25 ' Ubicacion
      .Columns("J:J").ColumnWidth = 15 ' Cantidad
      .Columns("K:K").ColumnWidth = 15 ' CodUsuario

      'Formateo el Cod.Producto: 000x
      .Range("D9:D65536").Select
      .Selection.NumberFormat = "000000"
      
      'Formateo: Encabezod con Negrita.
      .Range("B8:K8").Select
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
      .Cells(3, 1).Formula = "REPORTE: MOVIMIENTOS"
      
      .Range("A1:A1").Select
      .Selection.Font.ColorIndex = 5
      .Selection.Font.Size = 14
      .Selection.Font.Bold = True
      
      .Range("A3:A3").Select
      .Selection.Font.Size = 12
      .Selection.Font.Bold = True
      
      .Range("B8:B8").Select
      
      .Cells(5, 1).Formula = "Rango de Fechas: " & Text1(0).Text & " - " & Text1(1).Text
      .Cells(6, 1).Formula = "Fecha ejecución del Reporte: " & mFechaEjec
     
      .Cells(8, 2).Formula = "Fecha"
      .Cells(8, 3).Formula = "Tipo Movimiento"
      .Cells(8, 4).Formula = "Cód. Producto"
      .Cells(8, 5).Formula = "Código SAP"
      .Cells(8, 6).Formula = "Producto"
      .Cells(8, 7).Formula = "Almacén"
      .Cells(8, 8).Formula = "Bodega"
      .Cells(8, 9).Formula = "Ubicación"
      .Cells(8, 10).Formula = "Cantidad"
      .Cells(8, 11).Formula = "Cód. Usuario"
   End With
End Sub

Private Sub sPlanillaAjustes()

   mI = 9
   sCabeceraAjustes
   
   Set mRec = mObj.oEjecutarSelect("SELECT M.IdMov, M.Fecha, CASE WHEN M.CodTipoMovimiento = 'I' THEN 'Ingreso' ELSE 'Egreso' END AS CodTipoMovimiento , M.CodProducto,P.CodigoSap, P.Descripcion As DescripcionProducto, " & _
   " M.CodUbicacion, U.Descripcion As DescripiconUbicacion, U.CodBodega, B.Descripcion AS DescripcionBodega, " & _
   " B.CodAlmacen, A.Descripcion AS DescripcionAlmacen, M.Cantidad, M.CodUsuario, AJ.MotivoDesc " & _
   " FROM Movimientos2 M " & _
   " INNER JOIN " & _
   "   Producto P ON M.CodProducto = P.Codigo " & _
   " INNER JOIN " & _
   "  Ubicaciones U ON M.CodUbicacion = U.Codigo " & _
   " INNER JOIN " & _
   "  Bodegas B ON U.CodBodega = B.Codigo " & _
   " INNER JOIN " & _
   "  Almacenes A ON B.CodAlmacen = A.Codigo " & _
   " INNER JOIN   Ajustes AJ ON AJ.IdMov = M.IdMov " & _
   "  Where Fecha between '" & Format(Text1(0).Text, "yyyy-mm-dd") & " 00:00:00' and '" & Format(Text1(1).Text, "yyyy-mm-dd") & " 23:59:59 '" & _
   "  AND B.Codigo IN (SELECT CodBodega FROM Usuario_AccesoBodega WHERE CodUsuario = '" & Trim(Right(MDI.mUser, 15)) & "') " & _
   "  ORDER BY Fecha, IdMov; ")
   
   Do While Not mRec.EOF
      With XLS
         '.Cells(mi, 1).Formula = NVL(mRec!IdMov, "")
         .Cells(mI, 2).Formula = NVL(mRec!Fecha, "")
         .Cells(mI, 3).Formula = NVL(mRec!CodTipoMovimiento, "")
         .Cells(mI, 4).Formula = NVL(mRec!CodProducto, "")
         .Cells(mI, 5).Formula = NVL(mRec!CodigoSap, "")
         .Cells(mI, 6).Formula = NVL(mRec!DescripcionProducto, "")
         .Cells(mI, 7).Formula = NVL(mRec!DescripcionAlmacen, "")
         .Cells(mI, 8).Formula = NVL(mRec!DescripcionBodega, "")
         .Cells(mI, 9).Formula = NVL(mRec!DescripiconUbicacion, "")
         .Cells(mI, 10).Formula = NVL(mRec!Cantidad, "")
         .Cells(mI, 11).Formula = NVL(mRec!MotivoDesc, "")
         .Cells(mI, 12).Formula = NVL(mRec!CodUsuario, "")
      End With
      mRec.MoveNext
      mI = mI + 1
   Loop
   mRec.Close
End Sub

Private Sub sCabeceraAjustes()
   With XLS
      .WorkBooks.Add
      .Worksheets(1).Select
      .Worksheets(1).Name = "Ajustes"
      '.Columns("A:A").ColumnWidth = 15 ' NumMovimiento
      .Columns("B:B").ColumnWidth = 16 ' Fecha
      .Columns("C:C").ColumnWidth = 15 ' Tipo Movimiento
      .Columns("D:D").ColumnWidth = 15 ' Cod. Producto
      .Columns("E:E").ColumnWidth = 15 ' Codigo Sap
      .Columns("F:F").ColumnWidth = 70 ' Producto
      .Columns("G:G").ColumnWidth = 25 ' Almacen
      .Columns("H:H").ColumnWidth = 25 ' Bodega
      .Columns("I:I").ColumnWidth = 25 ' Ubicacion
      .Columns("J:J").ColumnWidth = 15 ' Cantidad
      .Columns("K:K").ColumnWidth = 100 ' MotivoDescripcion
      .Columns("L:L").ColumnWidth = 15 ' CodUsuario

      'Formateo el Cod.Producto: 000x
      .Range("D9:D65536").Select
      .Selection.NumberFormat = "000000"
      
      'Formateo: Encabezod con Negrita.
      .Range("B8:L8").Select
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
      .Cells(3, 1).Formula = "REPORTE: AJUSTES"
      
      .Range("A1:A1").Select
      .Selection.Font.ColorIndex = 5
      .Selection.Font.Size = 14
      .Selection.Font.Bold = True
      
      .Range("A3:A3").Select
      .Selection.Font.Size = 12
      .Selection.Font.Bold = True
      
      .Range("B8:B8").Select
      
      .Cells(5, 1).Formula = "Rango de Fechas: " & Text1(0).Text & " - " & Text1(1).Text
      .Cells(6, 1).Formula = "Fecha ejecución del Reporte: " & mFechaEjec
     
      .Cells(8, 2).Formula = "Fecha"
      .Cells(8, 3).Formula = "Tipo Movimiento"
      .Cells(8, 4).Formula = "Cód. Producto"
      .Cells(8, 5).Formula = "Código SAP"
      .Cells(8, 6).Formula = "Producto"
      .Cells(8, 7).Formula = "Almacén"
      .Cells(8, 8).Formula = "Bodega"
      .Cells(8, 9).Formula = "Ubicación"
      .Cells(8, 10).Formula = "Cantidad"
      .Cells(8, 11).Formula = "Motivo Ajuste"
      .Cells(8, 12).Formula = "Cód. Usuario"
      
   End With
End Sub

Private Sub Form_Load()
   Me.Width = 4575
   Me.Height = 3300
   sAlinearForm Me
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = fDateKeyPress(Text1(Index), KeyAscii)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set mObj = Nothing
   Set mRec = Nothing
   Set XLS = Nothing
   ShowMenu 12, True, False
End Sub
