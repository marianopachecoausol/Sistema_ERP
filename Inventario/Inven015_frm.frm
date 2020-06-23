VERSION 5.00
Begin VB.Form Inven015_frm 
   Caption         =   "Consumos"
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
Attribute VB_Name = "Inven015_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mObj As New clInven
Dim mRec As New ADODB.Recordset
Dim XLS As EXCEL.Application
Dim mi As Integer
Dim mFechaEjec As Date

Private Sub Command1_Click(Index As Integer)
   If Index = 0 Then
      If Fecha_ok(Text1(0).Text) And Fecha_ok(Text1(1).Text) Then
         If DateDiff("d", Text1(0).Text, Text1(1).Text) >= 0 Then
            sMsgEspere Me, "Procesando datos...", True
            mFechaEjec = Now()
            
            Set XLS = CreateObject("Excel.Application")
      
            sPlanilla
      
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

   mi = 9
   sCabecera
   
   Set mRec = mObj.oEjecutarSelect(" SELECT A.Descripcion as DescripcionAlmacen, B.Descripcion AS DescripcionBodega, " & _
   " M.CodProducto, P.CodigoSap, P.Descripcion AS DescripcionProducto, UM.Descripcion AS DescripcionUnidadMedida, " & _
   " SUM(M.Cantidad) AS Consumido " & _
   " FROM " & _
   " Movimientos2 M " & _
   "   INNER JOIN " & _
   " Producto P ON P.Codigo = M.CodProducto " & _
   "  INNER JOIN  " & _
   " Ubicaciones U ON M.CodUbicacion = U.Codigo " & _
   "   INNER JOIN " & _
   " Bodegas B ON U.CodBodega = B.Codigo " & _
   "   INNER JOIN " & _
   " Almacenes A ON B.CodAlmacen = A.Codigo " & _
   "   INNER JOIN  " & _
   " UnidadMedida UM ON UM.Codigo = P.CodUnidadMedida " & _
   " WHERE CodTipoMovimiento = 'E' " & _
   " AND Fecha between '" & Format(Text1(0).Text, "yyyy-mm-dd") & " 00:00:00' and '" & Format(Text1(1).Text, "yyyy-mm-dd") & " 23:59:59 '" & _
   " AND B.Codigo IN (SELECT CodBodega FROM Usuario_AccesoBodega WHERE CodUsuario = '" & Trim(Right(MDI.mUser, 15)) & "')" & _
   " GROUP BY  A.Descripcion, B.Descripcion, M.CodProducto, P.CodigoSap, P.Descripcion, UM.Descripcion " & _
   " ORDER BY A.Descripcion,  B.Descripcion, P.Descripcion; ")
   
   
   Do While Not mRec.EOF
      With XLS
      
      '.Cells(mi, 1).Formula = NVL(mRec!IdMov, "")
      .Cells(mi, 2).Formula = NVL(mRec!DescripcionAlmacen, "")
      .Cells(mi, 3).Formula = NVL(mRec!DescripcionBodega, "")
      .Cells(mi, 4).Formula = NVL(mRec!CodProducto, "")
      .Cells(mi, 5).Formula = NVL(mRec!CodigoSap, "")
      .Cells(mi, 6).Formula = NVL(mRec!DescripcionProducto, "")
      .Cells(mi, 7).Formula = NVL(mRec!Consumido, "")
      .Cells(mi, 8).Formula = NVL(mRec!DescripcionUnidadMedida, "")
      
      End With
      mRec.MoveNext
      mi = mi + 1
   Loop
   mRec.Close
End Sub

Private Sub sCabecera()
   With XLS
      .WorkBooks.Add
      .Worksheets(1).Select
      .Worksheets(1).Name = "Consumos"
     
      .Columns("B:B").ColumnWidth = 25 ' Almacen
      .Columns("C:C").ColumnWidth = 25 ' Bodega
      .Columns("D:D").ColumnWidth = 15 ' Cod. Producto
      .Columns("E:E").ColumnWidth = 15 ' Codigo Sap
      .Columns("F:F").ColumnWidth = 70 ' Producto
      .Columns("G:G").ColumnWidth = 15 ' Consumido
      .Columns("H:H").ColumnWidth = 25 ' UnidadMedida
         
      'Formateo el Cod.Producto: 000x
      .Range("D9:D65536").Select
      .Selection.NumberFormat = "000000"
      
      'Formateo: Encabezod con Negrita.
      .Range("B8:H8").Select
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
      .Cells(3, 1).Formula = "REPORTE: CONSUMOS"
      
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
      
      .Cells(8, 2).Formula = "Almacén"
      .Cells(8, 3).Formula = "Bodega"
      .Cells(8, 4).Formula = "Cód. Producto"
      .Cells(8, 5).Formula = "Código SAP"
      .Cells(8, 6).Formula = "Producto"
      .Cells(8, 7).Formula = "Consumido"
      .Cells(8, 8).Formula = "Unid. de Medida"


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
