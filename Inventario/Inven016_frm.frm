VERSION 5.00
Begin VB.Form Inven016_frm 
   Caption         =   "Egresos por personal"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5205
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3510
   ScaleWidth      =   5205
   Begin VB.ComboBox Combo2 
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
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1140
      Width           =   3495
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
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   300
      Width           =   3495
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
      Index           =   1
      Left            =   3360
      TabIndex        =   4
      Top             =   1920
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
      Left            =   1560
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   375
      Index           =   1
      Left            =   3000
      TabIndex        =   1
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Empleado:"
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
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label2 
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
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   855
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
      Top             =   1920
      Width           =   855
   End
End
Attribute VB_Name = "Inven016_frm"
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

Private Sub Combo1_Click()
   Combo2.Enabled = True
   sLlenoUsuariosRet
End Sub

Private Sub sLlenoUsuariosRet()
   Dim mCodBodega As String
   Dim mObj2 As New clInven
   Dim mRec2 As New ADODB.Recordset
   
   mCodBodega = Trim(Left(Combo1.Text, 4))
   Combo2.Clear
   
   Set mRec2 = mObj.oEjecutarSelect("SELECT CONCAT (P.Apellido,',', P.Nombres) AS Descripcion,P.CodUsuario AS CodUsuario FROM " & _
   " UsuariosRet_Bodegas UB " & _
   " Inner Join " & _
   " Personal P ON UB.CodUsuario = P.CodUsuario " & _
   " WHERE UB.CodBodega = '" & mCodBodega & "' AND P.Fecha_Baja IS NULL " & _
   " ORDER BY P.Apellido;")
   
   Do While Not mRec2.EOF
      Combo2.AddItem mRec2!descripcion & Space(60) & mRec2!CodUsuario
      mRec2.MoveNext
   Loop
   mRec2.Close
   Set mObj2 = Nothing
   Set mRec2 = Nothing
End Sub

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
   mi = 10
   sCabecera
   
   Set mRec = mObj.oEjecutarSelect("  " & _
   " SELECT A.Descripcion as DescripcionAlmacen, B.Descripcion AS DescripcionBodega, " & _
   " M.CodProducto, P.CodigoSap, P.Descripcion AS DescripcionProducto, M.Cantidad, UM.Descripcion AS DescripcionUnidadMedida, " & _
   " M.Fecha, CONCAT(PE.Apellido,',',PE.Nombres) as UsuarioAutoriza " & _
   " FROM  Movimientos2 M " & _
   " INNER JOIN  Producto P ON P.Codigo = M.CodProducto " & _
   " INNER JOIN   Ubicaciones U ON M.CodUbicacion = U.Codigo " & _
   " INNER JOIN  Bodegas B ON U.CodBodega = B.Codigo " & _
   " INNER JOIN  Almacenes A ON B.CodAlmacen = A.Codigo " & _
   " INNER JOIN   UnidadMedida UM ON UM.Codigo = P.CodUnidadMedida " & _
   " LEFT JOIN Consumos_Det CD ON CD.IdMov = M.IdMov " & _
   " LEFT JOIN Consumos_H CH ON CH.NroVale = CD.NroVale AND CH.CodTipoVale = CD.CodTipoVale " & _
   " LEFT JOIN Personal PE ON PE.CodUsuario = CH.CodUsuarioAutoriza " & _
   " WHERE CodTipoMovimiento = 'E' " & _
   " AND M.Fecha BETWEEN '" & Format(Text1(0).Text, "yyyy-mm-dd") & " 00:00:00' and '" & Format(Text1(1).Text, "yyyy-mm-dd") & " 23:59:59 ' " & _
   " AND CH.CodUsuarioRetira ='" & Trim(Right(Combo2.Text, 25)) & "'" & _
   " ORDER BY A.Descripcion,  B.Descripcion, P.Descripcion; ")
   
   Do While Not mRec.EOF
      With XLS
      
      .Cells(mi, 2).Formula = NVL(mRec!DescripcionAlmacen, "")
      .Cells(mi, 3).Formula = NVL(mRec!DescripcionBodega, "")
      .Cells(mi, 4).Formula = NVL(mRec!CodProducto, "")
      .Cells(mi, 5).Formula = NVL(mRec!CodigoSap, "")
      .Cells(mi, 6).Formula = NVL(mRec!DescripcionProducto, "")
      .Cells(mi, 7).Formula = NVL(mRec!Cantidad, "")
      .Cells(mi, 8).Formula = NVL(mRec!DescripcionUnidadMedida, "")
      .Cells(mi, 9).Formula = NVL(mRec!Fecha, "")
      .Cells(mi, 10).Formula = NVL(mRec!UsuarioAutoriza, "")
      
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
      .Worksheets(1).Name = "Egresos por Personal"
    
      .Columns("B:B").ColumnWidth = 25 ' Almacen
      .Columns("C:C").ColumnWidth = 25 ' Bodega
      .Columns("D:D").ColumnWidth = 15 ' Cod. Producto
      .Columns("E:E").ColumnWidth = 15 ' Codigo Sap
      .Columns("F:F").ColumnWidth = 70 ' Producto
      .Columns("G:G").ColumnWidth = 15 ' Cantidad
      .Columns("H:H").ColumnWidth = 25 ' UnidadMedida
      .Columns("I:I").ColumnWidth = 15 ' Fecha
      .Columns("J:J").ColumnWidth = 40 ' UsuarrioAutoriza
         
      'Formateo el Cod.Producto: 000x
      .Range("D6:D65536").Select
      .Selection.NumberFormat = "000000"
      
      
      'Formateo el campo Fecha
      .Range("I10:I65536").Select
      .Selection.NumberFormat = "dd-mm-yyyy"
      
      'Formateo: Encabezod con Negrita.
      .Range("B9:J9").Select
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
      .Cells(3, 1).Formula = "REPORTE: RETIRO DE PRODUCTOS POR PERSONAL"
      
      .Range("A1:A1").Select
      .Selection.Font.ColorIndex = 5
      .Selection.Font.Size = 14
      .Selection.Font.Bold = True
      
      .Range("A3:A3").Select
      .Selection.Font.Size = 12
      .Selection.Font.Bold = True
      
      .Range("B9:B9").Select
      
      .Cells(5, 1).Formula = "Retirados por: " & Trim(Left(Combo2.Text, 60))
      .Cells(6, 1).Formula = "Rango de Fechas: " & Text1(0).Text & " - " & Text1(1).Text
      .Cells(7, 1).Formula = "Fecha ejecución del Reporte: " & mFechaEjec

      .Cells(9, 2).Formula = "Almacén"
      .Cells(9, 3).Formula = "Bodega"
      .Cells(9, 4).Formula = "Cód. Producto"
      .Cells(9, 5).Formula = "Código SAP"
      .Cells(9, 6).Formula = "Producto"
      .Cells(9, 7).Formula = "Consumido"
      .Cells(9, 8).Formula = "Unid. de Medida"
      .Cells(9, 9).Formula = "Fecha"
      .Cells(9, 10).Formula = "Autorizador"
   End With
End Sub

Private Sub Form_Load()
   Me.Width = 5325
   Me.Height = 4020
   sAlinearForm Me
   
   Combo2.Enabled = False
   
   Set mRec = mObj.oTabla("Bodegas", " where Fecha_Baja IS NULL and Codigo IN (SELECT CodBodega FROM Usuario_AccesoBodega WHERE CodUsuario = '" & Trim(Right(MDI.mUser, 15)) & "') order by Codigo")
   Do While Not mRec.EOF
      Combo1.AddItem "" & mRec!Codigo & " " & mRec!descripcion & ""
      mRec.MoveNext
   Loop
   mRec.Close

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
