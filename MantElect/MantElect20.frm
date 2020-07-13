VERSION 5.00
Begin VB.Form MantElect20 
   Caption         =   "Ordenes de Trabajo"
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
Attribute VB_Name = "MantElect20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mObj As New clMantElect
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
      
            sPlanilla1
            sPlanilla2
            sPlanilla3
            sPlanilla4
            XLS.Worksheets(1).Select

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

Private Sub sPlanilla1()
   mi = 9
   sCabecera1
   
   Set mRec = mObj.getRPT_OTs(CDate(Text1(0).Text), CDate(Text1(1).Text))
   
   Do While Not mRec.EOF
      With XLS
         .Cells(mi, 2).Formula = NVL(mRec!IdOT, "")
         .Cells(mi, 3).Formula = NVL(mRec!FechaCarga, "")
         .Cells(mi, 4).Formula = NVL(mRec!FechaInicio, "")
         .Cells(mi, 5).Formula = NVL(mRec!FechaFin, "")
         .Cells(mi, 6).Formula = NVL(mRec!Elect_o_AA, "")
         .Cells(mi, 7).Formula = NVL(mRec!Usuario, "")
         
         .Cells(mi, 8).Formula = NVL(mRec!CodigoSap, "")
         .Cells(mi, 9).Formula = NVL(mRec!Producto, "")
         .Cells(mi, 10).Formula = NVL(mRec!Cantidad, "")
         .Cells(mi, 11).Formula = NVL(mRec!UnidadMedida, "")
         .Cells(mi, 12).Formula = NVL(mRec!Ubicacion, "")
         .Cells(mi, 13).Formula = NVL(mRec!NroVale, "")
      End With
      mRec.MoveNext
      mi = mi + 1
   Loop
   mRec.Close
End Sub

Private Sub sCabecera1()
   With XLS
      .WorkBooks.Add
      .Sheets.Add '4

      .Worksheets(1).Select
      .Worksheets(1).Name = "OTs_Consumo_Materiales "
     
      .Columns("A:A").ColumnWidth = 2
      .Columns("B:B").ColumnWidth = 12 ' Ord.Trabajo
      .Columns("C:C").ColumnWidth = 12 ' Fecha Carga
      .Columns("D:D").ColumnWidth = 12 ' Fecha Inicio
      .Columns("E:E").ColumnWidth = 12 ' Fecha Cierre
      .Columns("F:F").ColumnWidth = 17 ' Elect/AA
      .Columns("G:G").ColumnWidth = 27 ' O.T. cerrada por
      
      .Columns("H:H").ColumnWidth = 11 ' Cód. SAP
      .Columns("I:I").ColumnWidth = 63 ' Producto
      .Columns("J:J").ColumnWidth = 11 ' Cantidad
      .Columns("K:K").ColumnWidth = 12 ' Unidad Medida
      .Columns("L:L").ColumnWidth = 13 ' Ubicacion
      .Columns("M:M").ColumnWidth = 11 ' Nro. Vale

      .Columns("B:B").HorizontalAlignment = xlHAlignCenter
      .Columns("C:C").HorizontalAlignment = xlHAlignCenter
      .Columns("D:D").HorizontalAlignment = xlHAlignCenter
      .Columns("E:E").HorizontalAlignment = xlHAlignCenter
      .Columns("F:F").HorizontalAlignment = xlHAlignCenter
      .Columns("H:H").HorizontalAlignment = xlHAlignCenter
      .Columns("K:K").HorizontalAlignment = xlHAlignCenter
      .Columns("M:M").HorizontalAlignment = xlHAlignCenter
      
      'Formateo OrdTrabajo: 000x
      .Range("B9:B65536").Select
      .Selection.NumberFormat = "000000"

      'Formateo NroVale: 000x
      .Range("M9:M65536").Select
      .Selection.NumberFormat = "000000000"

      'Formateo: Encabezd con Negrita.
      .Range("B8:M8").Select
      .Selection.Font.Bold = True
      .Selection.Interior.ColorIndex = 15
      .Selection.HorizontalAlignment = xlHAlignCenter

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
      .Cells(3, 1).Formula = "REPORTE: Ordenes de Trabajo"

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

      .Cells(8, 2).Formula = "Ord.Trabajo"
      .Cells(8, 3).Formula = "Fecha Carga"
      .Cells(8, 4).Formula = "Fecha Inicio"
      .Cells(8, 5).Formula = "Fecha Cierre"
      .Cells(8, 6).Formula = "Eléctrico/A.A."
      .Cells(8, 7).Formula = "O.T. cerrada por"
      .Cells(8, 8).Formula = "Cód. SAP"
      .Cells(8, 9).Formula = "Producto"
      .Cells(8, 10).Formula = "Cantidad"
      .Cells(8, 11).Formula = "Unid.Medida"
      .Cells(8, 12).Formula = "Ubicación"
      .Cells(8, 13).Formula = "Nro. Vale"
   End With
End Sub

Private Sub sPlanilla2()
   mi = 9
   sCabecera2
   
   Set mRec = mObj.getRPT_OTs_Tecnicos(CDate(Text1(0).Text), CDate(Text1(1).Text))
   
   Do While Not mRec.EOF
      With XLS
         .Cells(mi, 2).Formula = NVL(mRec!IdOT, "")
         .Cells(mi, 3).Formula = NVL(mRec!FechaCarga, "")
         .Cells(mi, 4).Formula = NVL(mRec!FechaInicio, "")
         .Cells(mi, 5).Formula = NVL(mRec!FechaFin, "")
         .Cells(mi, 6).Formula = NVL(mRec!Elect_o_AA, "")
         .Cells(mi, 7).Formula = NVL(mRec!Usuario, "")
         
         .Cells(mi, 8).Formula = NVL(mRec!Tecnico, "")

      End With
      mRec.MoveNext
      mi = mi + 1
   Loop
   mRec.Close
End Sub

Private Sub sCabecera2()
   With XLS
      .Worksheets(2).Select
      .Worksheets(2).Name = "OTs_Tecnicos"
     
      .Columns("A:A").ColumnWidth = 2
      .Columns("B:B").ColumnWidth = 12 ' Ord.Trabajo
      .Columns("C:C").ColumnWidth = 12 ' Fecha Carga
      .Columns("D:D").ColumnWidth = 12 ' Fecha Inicio
      .Columns("E:E").ColumnWidth = 12 ' Fecha Cierre
      .Columns("F:F").ColumnWidth = 17 ' Elect/AA
      .Columns("G:G").ColumnWidth = 27 ' O.T. cerrada por
      .Columns("H:H").ColumnWidth = 31 ' Tecnicos

      .Columns("B:B").HorizontalAlignment = xlHAlignCenter
      .Columns("C:C").HorizontalAlignment = xlHAlignCenter
      .Columns("D:D").HorizontalAlignment = xlHAlignCenter
      .Columns("E:E").HorizontalAlignment = xlHAlignCenter
      .Columns("F:F").HorizontalAlignment = xlHAlignCenter
      
      'Formateo OrdTrabajo: 000x
      .Range("B9:B65536").Select
      .Selection.NumberFormat = "000000"

      'Formateo: Encabezd con Negrita.
      .Range("B8:H8").Select
      .Selection.Font.Bold = True
      .Selection.Interior.ColorIndex = 15
      .Selection.HorizontalAlignment = xlHAlignCenter

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
      .Cells(3, 1).Formula = "REPORTE: Ordenes de Trabajo"

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

      .Cells(8, 2).Formula = "Ord.Trabajo"
      .Cells(8, 3).Formula = "Fecha Carga"
      .Cells(8, 4).Formula = "Fecha Inicio"
      .Cells(8, 5).Formula = "Fecha Cierre"
      .Cells(8, 6).Formula = "Eléctrico/A.A."
      .Cells(8, 7).Formula = "O.T. cerrada por"
      .Cells(8, 8).Formula = "Técnico"

   End With
End Sub

Private Sub sPlanilla3()
   mi = 9
   sCabecera3
   
   Set mRec = mObj.getRPT_OTs_Subrubros(CDate(Text1(0).Text), CDate(Text1(1).Text))
   
   Do While Not mRec.EOF
      With XLS
         .Cells(mi, 2).Formula = NVL(mRec!IdOT, "")
         .Cells(mi, 3).Formula = NVL(mRec!FechaCarga, "")
         .Cells(mi, 4).Formula = NVL(mRec!FechaInicio, "")
         .Cells(mi, 5).Formula = NVL(mRec!FechaFin, "")
         .Cells(mi, 6).Formula = NVL(mRec!Elect_o_AA, "")
         .Cells(mi, 7).Formula = NVL(mRec!Usuario, "")
         
         .Cells(mi, 8).Formula = NVL(mRec!Rubro, "")
         .Cells(mi, 9).Formula = NVL(mRec!SubRubro, "")
      End With
      mRec.MoveNext
      mi = mi + 1
   Loop
   mRec.Close
End Sub

Private Sub sCabecera3()
   With XLS
      .Worksheets(3).Select
      .Worksheets(3).Name = "Ord.Trabajo_SubRubros"
     
      .Columns("A:A").ColumnWidth = 2
      .Columns("B:B").ColumnWidth = 12 ' Ord.Trabajo
      .Columns("C:C").ColumnWidth = 12 ' Fecha Carga
      .Columns("D:D").ColumnWidth = 12 ' Fecha Inicio
      .Columns("E:E").ColumnWidth = 12 ' Fecha Cierre
      .Columns("F:F").ColumnWidth = 17 ' Elect/AA
      .Columns("G:G").ColumnWidth = 27 ' O.T. cerrada por
      
      .Columns("H:H").ColumnWidth = 40 ' Rubro
      .Columns("I:I").ColumnWidth = 70 ' SubRubro

      .Columns("B:B").HorizontalAlignment = xlHAlignCenter
      .Columns("C:C").HorizontalAlignment = xlHAlignCenter
      .Columns("D:D").HorizontalAlignment = xlHAlignCenter
      .Columns("E:E").HorizontalAlignment = xlHAlignCenter
      .Columns("F:F").HorizontalAlignment = xlHAlignCenter

      'Formateo OrdTrabajo: 000x
      .Range("B9:B65536").Select
      .Selection.NumberFormat = "000000"

      'Formateo: Encabezd con Negrita.
      .Range("B8:I8").Select
      .Selection.Font.Bold = True
      .Selection.Interior.ColorIndex = 15
      .Selection.HorizontalAlignment = xlHAlignCenter

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
      .Cells(3, 1).Formula = "REPORTE: Ordenes de Trabajo"

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

      .Cells(8, 2).Formula = "Ord.Trabajo"
      .Cells(8, 3).Formula = "Fecha Carga"
      .Cells(8, 4).Formula = "Fecha Inicio"
      .Cells(8, 5).Formula = "Fecha Cierre"
      .Cells(8, 6).Formula = "Eléctrico/A.A."
      .Cells(8, 7).Formula = "O.T. cerrada por"
      .Cells(8, 8).Formula = "Rubro"
      .Cells(8, 9).Formula = "Sub Rubro"
   End With
End Sub

Private Sub sPlanilla4()
   mi = 9
   sCabecera4
   
   Set mRec = mObj.getRPT_OTs_Vehiculos(CDate(Text1(0).Text), CDate(Text1(1).Text))
   
   Do While Not mRec.EOF
      With XLS
         .Cells(mi, 2).Formula = NVL(mRec!IdOT, "")
         .Cells(mi, 3).Formula = NVL(mRec!FechaCarga, "")
         .Cells(mi, 4).Formula = NVL(mRec!FechaInicio, "")
         .Cells(mi, 5).Formula = NVL(mRec!FechaFin, "")
         .Cells(mi, 6).Formula = NVL(mRec!Elect_o_AA, "")
         .Cells(mi, 7).Formula = NVL(mRec!Usuario, "")
         
         .Cells(mi, 8).Formula = NVL(mRec!Vehiculo, "")
         .Cells(mi, 9).Formula = NVL(mRec!KmInicial, "")
         .Cells(mi, 10).Formula = NVL(mRec!KmFinal, "")

      End With
      mRec.MoveNext
      mi = mi + 1
   Loop
   mRec.Close
End Sub

Private Sub sCabecera4()
   With XLS
      .Worksheets(4).Select
      .Worksheets(4).Name = "OTs_Vehículos"
     
      .Columns("A:A").ColumnWidth = 2
      .Columns("B:B").ColumnWidth = 12 ' Ord.Trabajo
      .Columns("C:C").ColumnWidth = 12 ' Fecha Carga
      .Columns("D:D").ColumnWidth = 12 ' Fecha Inicio
      .Columns("E:E").ColumnWidth = 12 ' Fecha Cierre
      .Columns("F:F").ColumnWidth = 17 ' Elect/AA
      .Columns("G:G").ColumnWidth = 27 '  O.T. cerrada por
      
      .Columns("H:H").ColumnWidth = 27 ' Vehiculo
      .Columns("I:I").ColumnWidth = 17 ' KmInicnal
      .Columns("J:J").ColumnWidth = 17 ' KmFinal

      .Columns("B:B").HorizontalAlignment = xlHAlignCenter
      .Columns("C:C").HorizontalAlignment = xlHAlignCenter
      .Columns("D:D").HorizontalAlignment = xlHAlignCenter
      .Columns("E:E").HorizontalAlignment = xlHAlignCenter
      .Columns("F:F").HorizontalAlignment = xlHAlignCenter
      .Columns("I:I").HorizontalAlignment = xlHAlignRight
      .Columns("J:J").HorizontalAlignment = xlHAlignRight
      
      'Formateo OrdTrabajo: 000x
      .Range("B9:B65536").Select
      .Selection.NumberFormat = "000000"

      'Formateo: Encabezd con Negrita.
      .Range("B8:J8").Select
      .Selection.Font.Bold = True
      .Selection.Interior.ColorIndex = 15
      .Selection.HorizontalAlignment = xlHAlignCenter

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
      .Cells(3, 1).Formula = "REPORTE: Ordenes de Trabajo"

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

      .Cells(8, 2).Formula = "Ord.Trabajo"
      .Cells(8, 3).Formula = "Fecha Carga"
      .Cells(8, 4).Formula = "Fecha Inicio"
      .Cells(8, 5).Formula = "Fecha Cierre"
      .Cells(8, 6).Formula = "Eléctrico/A.A."
      .Cells(8, 7).Formula = "O.T. cerrada por"
      .Cells(8, 8).Formula = "Vehículo"
      .Cells(8, 9).Formula = "Km Inicial"
      .Cells(8, 10).Formula = "Km Final"
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
   ShowMenu 47, True, False
End Sub
