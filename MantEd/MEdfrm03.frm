VERSION 5.00
Begin VB.Form MEdfrm03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modulo de Reportes"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   2760
      TabIndex        =   1
      Top             =   1740
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   2760
      TabIndex        =   0
      Top             =   1020
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Volver"
      Height          =   375
      Index           =   1
      Left            =   2880
      TabIndex        =   3
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   2
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Hasta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   1320
      TabIndex        =   6
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Desde"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   1320
      TabIndex        =   5
      Top             =   1080
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Reportes Diarios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1680
      TabIndex        =   4
      Top             =   120
      Width           =   1785
   End
End
Attribute VB_Name = "MEdfrm03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mObj As New clMantEd
Dim mRec As New ADODB.Recordset
Dim WorkBooks As Object
Dim XLS As EXCEL.Application
Public mSector As String

Private Sub Command1_Click(Index As Integer)
If Index = 0 Then
   If fValida Then
      Select Case mSector
         Case 0, 1
            Generar_XLS_EstSol
         Case 2
            Generar_XLS_Diar
         Case 3
            Generar_XLS_Comp
         Case 4
            Generar_XLS_Verif
         Case 5
            Generar_XLS_Comp_X_Zonas
      End Select
   End If
Else
   Unload Me
End If
End Sub

Private Sub Form_Load()
sAlinearForm Me
If mSector = 2 Or mSector = 3 Then
   Text1(0).Text = "01/01/2015"
   Text1(1).Text = Right(100 + Day(Date), 2) & "/" & Right(100 + Month(Date), 2) & "/" & Year(Date)
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set mObj = Nothing
Set mRec = Nothing
ShowMenu 20, True, False
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = fDateKeyPress(Text1(Index), KeyAscii)
End Sub

Private Function fValida() As Boolean
Dim mRet As Boolean
Dim mI As Integer
mRet = True
For mI = 0 To Text1.UBound
   If mRet Then
      mRet = (Text1(mI).Text <> "")
   End If
Next
If mRet Then
   mRet = Fecha_ok(Text1(0).Text)
End If
If mRet Then
   mRet = Fecha_ok(Text1(1).Text)
End If
If mRet Then
   mRet = (DateDiff("n", Text1(0).Text, Text1(1).Text) >= 0)
End If
If Not mRet Then
   MsgBox "Verifique que los datos estén ingresados correctamente", vbCritical, "Atención"
End If
fValida = mRet
End Function


Public Sub GenXLS_PartesPendientes()

Dim mI As Integer
Dim mFila As Integer
'Dim cantSolapas As Integer
Dim mfechaEjecucionReport As Date
Dim mObj2 As New clMantEd
Dim mRec2 As New ADODB.Recordset

'Dim marrZonasEdilicias(4) As String
'Dim marrDescriZonasEdilicias(4) As String

Dim marrZonasEdilicias(3) As String
Dim marrDescriZonasEdilicias(3) As String

'marrZonasEdilicias(0) = "Solicitud Especial"
'marrZonasEdilicias(1) = "Z1"
'marrZonasEdilicias(2) = "Z2"
'marrZonasEdilicias(3) = "Z3"
'marrZonasEdilicias(4) = "Z4"
'
'marrDescriZonasEdilicias(0) = "Solicitud Especial"
'marrDescriZonasEdilicias(1) = "ZONA 1 TIGRE-DEBT-MRQZ-C.REAL"
'marrDescriZonasEdilicias(2) = "ZONA CENTRAL BUEN AYRE-202-BELGRANO-197 "
'marrDescriZonasEdilicias(3) = "ZONA 3 PILAR-DECALADO PILAR"
'marrDescriZonasEdilicias(4) = "ZONA 4 CAMPANA-DECALADO CAMPANA"




marrZonasEdilicias(0) = "Z1"
marrZonasEdilicias(1) = "Z2"
marrZonasEdilicias(2) = "Z3"
marrZonasEdilicias(3) = "Z4"


marrDescriZonasEdilicias(0) = "ZONA 1 TIGRE-DEBT-MRQZ-C.REAL"
marrDescriZonasEdilicias(1) = "ZONA CENTRAL BUEN AYRE-202-BELGRANO-197 "
marrDescriZonasEdilicias(2) = "ZONA 3 PILAR-DECALADO PILAR"
marrDescriZonasEdilicias(3) = "ZONA 4 CAMPANA-DECALADO CAMPANA"



Screen.MousePointer = vbHourglass
Set XLS = CreateObject("Excel.Application")
XLS.WorkBooks.Add

mfechaEjecucionReport = Now()
'cantSolapas = mObj.iCountSolpasExcelPartesPendientes()

'Set mRec = mObj.oEjecutarSelect("SELECT DISTINCT ZonaMantEdil " & _
'                                    "FROM " & _
'                                       "Registros R " & _
'                                    "INNER JOIN " & _
'                                       "Edificios E ON R.CodEdificio = CONCAT(E.Tramo , ' - ' , E.Descripcion) " & _
'                                    "WHERE Estado IN ('P','G') " & _
'                                    "AND FecTer IS NULL " & _
'                                    "AND SolicEspecial = 0 " & _
'                                 "UNION " & _
'                                 "SELECT DISTINCT 'Solicitud Especial' AS ZonaMantEdil " & _
'                                    "FROM " & _
'                                       "Registros R " & _
'                                    "INNER JOIN " & _
'                                       "Edificios E ON R.CodEdificio = CONCAT(E.Tramo , ' - ' , E.Descripcion) " & _
'                                    "WHERE Estado in ('P','G') " & _
'                                 "AND FecTer is null " & _
'                                 "AND SolicEspecial = 1 " & _
'                                 "ORDER BY 1 ")

'For mI = XLS.Sheets.Count To cantSolapas - 1
'   XLS.Sheets.Add
'Next
'
For mI = XLS.Sheets.Count To UBound(marrZonasEdilicias)
   XLS.Sheets.Add
Next
                                 
                                 
                                 
'For mI = 1 To cantSolapas
For mI = 1 To UBound(marrZonasEdilicias) + 1
      'XLS.Sheets(mI).Name = mRec!ZonaMantEdil
      XLS.Sheets(mI).Name = marrZonasEdilicias(mI - 1)
      XLS.Sheets(mI).Select
      XLS.Sheets(mI).PageSetup.Orientation = xlLandscape
   
      XLS.Cells(1, 2).Formula = "AUTOPISTAS DEL SOL S.A.   -  SERVICIO DE MANTENIMIENTO DE ESTACION  -  "

      XLS.Range("B1:B1").Font.Bold = True
      XLS.Range("B1:B1").Font.Size = 12
      XLS.Range("B1:H1").Merge
      XLS.Range("B1:H1").HorizontalAlignment = xlCenter
      
      XLS.Range("B3:D3").Merge
      XLS.Range("B3:D3").HorizontalAlignment = xlCenter
      XLS.Range("E3:H3").Merge
      XLS.Range("E3:H3").HorizontalAlignment = xlCenter
      XLS.Range("B3:E3").Font.Bold = True
      XLS.Range("B3:E3").Font.Size = 8
     
      XLS.Cells(3, 5).Formula = Format(mfechaEjecucionReport, "DD/MM/YYYY")
      XLS.Cells(3, 2).Formula = "PARTE DE TRABAJO " & marrDescriZonasEdilicias(mI - 1)
      
      XLS.Range("B4:H4").Interior.Color = RGB(196, 194, 194)
      XLS.Range("A4:H4").HorizontalAlignment = xlCenter
      XLS.Range("A4:H4").VerticalAlignment = xlCenter
      XLS.Range("A4:H4").Font.Bold = True
      XLS.Range("A4:H4").Font.Size = 8
        
      XLS.Cells(4, 2).Formula = "Tarea"
      'XLS.Cells(7, 3).Formula = "Fecha Solicitud"
      XLS.Cells(4, 3).Formula = "Lugar"
      XLS.Cells(4, 4).Formula = "Descripción del la solicitud"
      XLS.Cells(4, 5).Formula = "Material Utilizado"
      XLS.Cells(4, 6).Formula = "Hora Inicio"
      XLS.Cells(4, 7).Formula = "Hora Fin"
      XLS.Cells(4, 8).Formula = "Estado"
      
      'Set mRec2 = mObj2.oDetallePartesPendientes(mRec!ZonaMantEdil)
      Set mRec2 = mObj2.oDetallePartesPendientes(marrZonasEdilicias(mI - 1))
      If Not mRec2.EOF Then
         mFila = 5
         Do While Not mRec2.EOF
            XLS.Cells(mFila, 2).Formula = mRec2!Parte
          '  XLS.Cells(mFila, 3).Formula = "'" & Format(mRec2!FechaSolic, "DD/MM/YYYY")
            XLS.Cells(mFila, 3).Formula = mRec2!descripcion
            XLS.Cells(mFila, 3).WrapText = True
            XLS.Cells(mFila, 4).Formula = mRec2!DescripSolic
            XLS.Cells(mFila, 4).WrapText = True
            
            XLS.Range("B8:H" & mFila).Font.Size = 8
            
            mFila = mFila + 1
            mRec2.MoveNext
         Loop
      End If
      
      For mFila = 3 To 22
            XLS.Range("B3:H" & mFila).VerticalAlignment = xlCenter
            XLS.Range("B3:H" & mFila).Borders(xlEdgeTop).LineStyle = xlContinuous
            XLS.Range("B3:H" & mFila).Borders(xlEdgeBottom).LineStyle = xlContinuous
            XLS.Range("B3:H" & mFila).Borders(xlEdgeLeft).LineStyle = xlContinuous
            XLS.Range("B3:H" & mFila).Borders(xlEdgeRight).LineStyle = xlContinuous
            XLS.Range("B3:H" & mFila).Borders(xlInsideVertical).LineStyle = xlContinuous
            If mFila >= 5 Then
               XLS.Cells(mFila, 1).RowHeight = 22.5
            End If
      Next
   
      XLS.Columns("A:A").ColumnWidth = 0.33
      XLS.Columns("B:B").ColumnWidth = 5.14
      XLS.Columns("C:C").ColumnWidth = 13
      XLS.Columns("D:D").ColumnWidth = 70
      XLS.Columns("E:E").ColumnWidth = 14
      XLS.Columns("F:F").ColumnWidth = 7.86
      XLS.Columns("G:G").ColumnWidth = 7.86
      XLS.Columns("H:H").ColumnWidth = 5.86
      
'      mRec.MoveNext
Next

'mRec.Close
mRec2.Close
Set mRec2 = Nothing
Set mObj2 = Nothing

XLS.Sheets(1).Select
XLS.Visible = True

Set XLS = Nothing
Screen.MousePointer = vbArrow

End Sub



Private Sub Generar_XLS_EstSol()
Dim mFila As Double
Screen.MousePointer = vbHourglass
Set XLS = CreateObject("Excel.Application")
XLS.WorkBooks.Add



XLS.Cells(1, 1).Formula = "AUTOPISTAS DEL SOL S.A."
XLS.Cells(2, 1).Formula = "SOLICITUDES DE SERVICIO AL SECTOR DE MANTENIMIENTO EDILICIO"
XLS.Cells(4, 1).Formula = "Período: " & Text1(0).Text & " al " & Text1(1).Text
XLS.Range("A6:I7").WrapText = True
XLS.Range("A6:I7").HorizontalAlignment = xlCenter
XLS.Range("A6:I7").VerticalAlignment = xlCenter
XLS.Range("A6:A7").Merge
XLS.Cells(6, 1).Formula = "Parte"
XLS.Range("B6:B7").Merge
XLS.Cells(6, 2).Formula = "Fecha Solicitud"
XLS.Range("C6:C7").Merge
XLS.Cells(6, 3).Formula = "Edificio"
XLS.Range("D6:D7").Merge
XLS.Cells(6, 4).Formula = "Descripcion Solicitud"
XLS.Range("E6:E7").Merge
XLS.Cells(6, 5).Formula = "Prioridad"
XLS.Cells(6, 6).Formula = "Fecha Termacion"
XLS.Range("F6:G6").Merge
XLS.Cells(7, 6).Formula = "Estimada"
XLS.Range("H6:H7").Merge
XLS.Cells(7, 7).Formula = "Real"
XLS.Cells(6, 8).Formula = "Estado"
XLS.Range("I6:I7").Merge
XLS.Cells(6, 9).Formula = "Solicitante"
'Set mRec = mObj.oEjecutarSelect("SELECT * FROM Registros WHERE FechaSolic BETWEEN '" & Format(Text1(0).Text, "YYYY-MM-DD") & " 00:00:00" & "' And '" & Format(Text1(1).Text, "YYYY-MM-DD") & " 23:59:59" & "' " & IIf(mSector = 0, " And Origen = 'S'", "") & " ORDER BY Parte") ' trae los anulados, esta incorrecto

Set mRec = mObj.oEjecutarSelect("SELECT * FROM Registros R " & _
" Left Join  AnulacionesParte A ON R.Parte = A.ParteAnu " & _
" WHERE FechaSolic BETWEEN '" & Format(Text1(0).Text, "YYYY-MM-DD") & " 00:00:00" & "' And '" & Format(Text1(1).Text, "YYYY-MM-DD") & " 23:59:59" & "' " & IIf(mSector = 0, " And Origen = 'S'", "") & " AND A.ParteAnu IS NULL  ORDER BY Parte")



If Not mRec.EOF Then
   mFila = 8
   Do While Not mRec.EOF
      XLS.Cells(mFila, 1).Formula = mRec!Parte
      XLS.Cells(mFila, 2).Formula = "'" & Format(mRec!FechaSolic, "DD/MM/YYYY")
      XLS.Cells(mFila, 3).Formula = mRec!CodEdificio
      XLS.Cells(mFila, 4).Formula = mRec!DescripSolic
      XLS.Cells(mFila, 5).Formula = mRec!Prioridad
      XLS.Cells(mFila, 6).Formula = "'" & Format(mRec!FechaSolic + mRec!TiempoEstim, "DD/MM/YYYY")
      XLS.Cells(mFila, 7).Formula = "'" & Format(mRec!FechaAsist, "DD/MM/YYYY")
      XLS.Cells(mFila, 8).Formula = mRec!estado
      XLS.Cells(mFila, 9).Formula = mRec!OpGen
      mFila = mFila + 1
      mRec.MoveNext
   Loop
End If

XLS.Range("A1:I7").Font.Bold = True
XLS.Range("A1:A2").Font.Size = 14
XLS.Range("A1:I7").Font.Bold = True
XLS.Range("A4:A4").Font.Size = 12
XLS.Columns("A:G").ColumnWidth = 6
XLS.Columns("B:H").ColumnWidth = 12
XLS.Columns("C:C").ColumnWidth = 25
XLS.Columns("D:D").ColumnWidth = 50
XLS.Columns("E:E").ColumnWidth = 10
XLS.Columns("I:I").ColumnWidth = 13

XLS.Visible = True
Set XLS = Nothing
Screen.MousePointer = vbArrow
End Sub

Private Sub Generar_XLS_Diar()
Dim mFila As Double
Screen.MousePointer = vbHourglass
Set XLS = CreateObject("Excel.Application")
XLS.WorkBooks.Add

XLS.Cells(1, 1).Formula = "AUTOPISTAS DEL SOL S.A."
XLS.Cells(2, 1).Formula = "AVANCE DE TAREAS - MANTENIMIENTO EDILICIO"
XLS.Cells(4, 1).Formula = "Período: " & Text1(0).Text & " al " & Text1(1).Text
XLS.Range("A6:Q7").WrapText = True
XLS.Range("A6:Q7").HorizontalAlignment = xlCenter
XLS.Range("A6:Q7").VerticalAlignment = xlCenter
XLS.Range("A6:A7").Merge
XLS.Cells(6, 1).Formula = "No. TAREA"
XLS.Range("B6:B7").Merge
XLS.Cells(6, 2).Formula = "PEDIDO"
XLS.Range("C6:C7").Merge
XLS.Cells(6, 3).Formula = "ASISTIDO"
XLS.Range("D6:D7").Merge
XLS.Cells(6, 4).Formula = "HORA ING."
XLS.Range("E6:E7").Merge
XLS.Cells(6, 5).Formula = "HORA EGR."
XLS.Range("F6:F7").Merge
XLS.Cells(6, 6).Formula = "EDIFICIO"
XLS.Range("G6:G7").Merge
XLS.Cells(6, 7).Formula = "DESCRIPCION DEL PEDIDO"
XLS.Range("H6:H7").Merge
XLS.Cells(6, 8).Formula = "OBSERVACION"
XLS.Range("I6:P6").Merge
XLS.Cells(6, 9).Formula = "Mano de Obra"
XLS.Cells(7, 9).Formula = "AP"
XLS.Cells(7, 10).Formula = "DC"
XLS.Cells(7, 11).Formula = "DM"
XLS.Cells(7, 12).Formula = "HC"
XLS.Cells(7, 13).Formula = "HS"
XLS.Cells(7, 14).Formula = "LO"
XLS.Cells(7, 15).Formula = "MP"
XLS.Cells(7, 16).Formula = "NG"
XLS.Range("Q6:Q7").Merge
XLS.Cells(6, 17).Formula = "Materiales"
mFila = 8

'Set mRec = mObj.oEjecutarSelect("SELECT * FROM Registros WHERE FechaSolic BETWEEN '" & Format(Text1(0).Text, "YYYY-MM-DD") & " 00:00:00" & "' And '" & Format(Text1(1).Text, "YYYY-MM-DD") & " 23:59:59" & "' " & IIf(mSector = 0, " And Origen = 'S'", "") & " And Estado = 'G' ORDER BY Parte") ' trae los anulados
Set mRec = mObj.oEjecutarSelect("SELECT R.* FROM Registros R " & _
"  Left Join  AnulacionesParte A ON R.Parte = A.ParteAnu " & _
" WHERE FechaSolic BETWEEN '" & Format(Text1(0).Text, "YYYY-MM-DD") & " 00:00:00" & "' And '" & Format(Text1(1).Text, "YYYY-MM-DD") & " 23:59:59" & "' " & IIf(mSector = 0, " And Origen = 'S'", "") & " And Estado = 'G' AND A.ParteAnu IS NULL  ORDER BY Parte")

If Not mRec.EOF Then
   Do While Not mRec.EOF
      XLS.Cells(mFila, 1).Formula = mRec!Parte
      XLS.Cells(mFila, 2).Formula = "'" & Format(mRec!FechaSolic, "DD/MM/YYYY")
      XLS.Cells(mFila, 3).Formula = "'" & Format(mRec!FechaAsist, "DD/MM/YYYY")
      XLS.Cells(mFila, 4).Formula = mRec!HoraIniAsist
      XLS.Cells(mFila, 5).Formula = mRec!HoraFinAsist
      XLS.Cells(mFila, 6).Formula = mRec!CodEdificio
      XLS.Cells(mFila, 7).Formula = mRec!DescripSolic
      XLS.Cells(mFila, 8).Formula = mRec!Observaciones
      XLS.Cells(mFila, 9).Formula = ""
      XLS.Cells(mFila, 10).Formula = ""
      XLS.Cells(mFila, 11).Formula = ""
      XLS.Cells(mFila, 12).Formula = ""
      XLS.Cells(mFila, 13).Formula = ""
      XLS.Cells(mFila, 14).Formula = ""
      XLS.Cells(mFila, 15).Formula = ""
      XLS.Cells(mFila, 16).Formula = ""
      XLS.Cells(mFila, 17).Formula = mRec!Materiales
      mFila = mFila + 1
      mRec.MoveNext
   Loop
End If

XLS.Range("A1:Q7").Font.Bold = True
XLS.Range("A1:A2").Font.Size = 14
XLS.Range("A1:Q7").Font.Bold = True
XLS.Range("A4:A4").Font.Size = 12
XLS.Columns("A:A").ColumnWidth = 7
XLS.Columns("B:C").ColumnWidth = 11
XLS.Columns("D:E").ColumnWidth = 7
XLS.Columns("F:F").ColumnWidth = 21
XLS.Columns("G:H").ColumnWidth = 30
XLS.Columns("I:P").ColumnWidth = 3
XLS.Columns("Q:Q").ColumnWidth = 17

XLS.Range("A8:Q" & mFila).WrapText = True
XLS.Range("A8:A" & mFila).HorizontalAlignment = xlCenter
XLS.Range("A8:Q" & mFila).VerticalAlignment = xlCenter
XLS.Rows("8:" & mFila).EntireRow.AutoFit
XLS.Range("A6:Q" & mFila).Borders(xlEdgeTop).LineStyle = xlContinuous
XLS.Range("A6:Q" & mFila).Borders(xlEdgeBottom).LineStyle = xlContinuous
XLS.Range("A6:Q" & mFila).Borders(xlEdgeLeft).LineStyle = xlContinuous
XLS.Range("A6:Q" & mFila).Borders(xlEdgeRight).LineStyle = xlContinuous
XLS.Range("A6:Q" & mFila).Borders(xlInsideVertical).LineStyle = xlContinuous
XLS.Range("A6:Q" & mFila).Borders(xlInsideHorizontal).LineStyle = xlContinuous

XLS.Visible = True
Set XLS = Nothing
Screen.MousePointer = vbArrow
End Sub

Private Sub Generar_XLS_Comp()
Dim mFila As Double
Screen.MousePointer = vbHourglass
Set XLS = CreateObject("Excel.Application")
XLS.WorkBooks.Add

XLS.Cells(1, 1).Formula = "AUTOPISTAS DEL SOL S.A."
XLS.Cells(2, 1).Formula = "AVANCE DE TAREAS - MANTENIMIENTO EDILICIO"
XLS.Cells(4, 1).Formula = "Período: " & Text1(0).Text & " al " & Text1(1).Text
XLS.Range("A6:AA7").WrapText = True
XLS.Range("A6:AA7").HorizontalAlignment = xlCenter
XLS.Range("A6:AA7").VerticalAlignment = xlCenter
XLS.Range("A6:A7").Merge
XLS.Cells(6, 1).Formula = "No. TAREA"
XLS.Range("B6:B7").Merge
XLS.Cells(6, 2).Formula = "PEDIDO"
XLS.Range("C6:C7").Merge
XLS.Cells(6, 3).Formula = "ASISTIDO"
XLS.Range("D6:D7").Merge
XLS.Cells(6, 4).Formula = "HORA ING."
XLS.Range("E6:E7").Merge
XLS.Cells(6, 5).Formula = "HORA EGR."
XLS.Range("F6:F7").Merge
XLS.Cells(6, 6).Formula = "EDIFICIO"
XLS.Range("G6:G7").Merge
XLS.Cells(6, 7).Formula = "DESCRIPCION DEL PEDIDO"
XLS.Range("H6:H7").Merge
XLS.Cells(6, 8).Formula = "OBSERVACION"
XLS.Range("I6:K6").Merge
XLS.Cells(6, 9).Formula = "TIPO DE MANTENIMIENTO"
XLS.Cells(7, 9).Formula = "Preventivo"
XLS.Cells(7, 10).Formula = "Predictivo"
XLS.Cells(7, 11).Formula = "Correctivo"
XLS.Range("L6:L7").Merge
XLS.Cells(6, 12).Formula = "RUBRO"
XLS.Range("M6:O6").Merge
XLS.Cells(6, 13).Formula = "TIEMPO DE RESPUESTA"
XLS.Cells(7, 13).Formula = "Estimado"
XLS.Cells(7, 14).Formula = "Real"
XLS.Cells(7, 15).Formula = "Admisible"
XLS.Range("P6:P7").Merge
XLS.Cells(6, 16).Formula = "ALERTAS"
XLS.Range("Q6:Z6").Merge
XLS.Cells(6, 17).Formula = "MANO DE OBRA"
XLS.Cells(7, 17).Formula = "AP"
XLS.Cells(7, 18).Formula = "DC"
XLS.Cells(7, 19).Formula = "DM"
XLS.Cells(7, 20).Formula = "HC"
XLS.Cells(7, 21).Formula = "HS"
XLS.Cells(7, 22).Formula = "LO"
XLS.Cells(7, 23).Formula = "MP"
XLS.Cells(7, 24).Formula = "NG"
XLS.Cells(7, 25).Formula = "Hs/Per"
XLS.Cells(7, 26).Formula = "$/Tarea"
'XLS.Range("Y6:Y7").Merge
XLS.Range("AA6:AA7").Merge
XLS.Cells(6, 27).Formula = "Materiales"

'Set mRec = mObj.oEjecutarSelect("SELECT * FROM Registros WHERE FechaSolic BETWEEN '" & Format(Text1(0).Text, "YYYY-MM-DD") & " 00:00:00" & "' And '" & Format(Text1(1).Text, "YYYY-MM-DD") & " 23:59:59" & "' ORDER BY Parte")

Set mRec = mObj.oEjecutarSelect("SELECT R.* FROM Registros R " & _
" Left Join  AnulacionesParte A ON R.Parte = A.ParteAnu " & _
" WHERE FechaSolic BETWEEN '" & Format(Text1(0).Text, "YYYY-MM-DD") & " 00:00:00" & "' And '" & Format(Text1(1).Text, "YYYY-MM-DD") & " 23:59:59" & "' AND A.ParteAnu IS NULL  ORDER BY Parte")

If Not mRec.EOF Then
   mFila = 8
   Do While Not mRec.EOF
      XLS.Cells(mFila, 1).Formula = mRec!Parte
      If mRec!TiempoEstim <= 3 Then
         XLS.Range("A" & mFila & ":A" & mFila).Interior.ColorIndex = 3
      End If
      XLS.Cells(mFila, 2).Formula = Format(mRec!FechaSolic, "m/d/yyyy")
      If IsNull(mRec!FechaAsist) Then
         If DateDiff("d", mRec!FechaSolic + mRec!TiempoAdmis, Date) > 0 Then
            XLS.Range("B" & mFila & ":B" & mFila).Interior.ColorIndex = 44
         Else
            If DateDiff("d", mRec!FechaSolic + mRec!TiempoEstim, Date) > 0 Then
               XLS.Range("B" & mFila & ":B" & mFila).Interior.ColorIndex = 6
            End If
         End If
      End If
      XLS.Cells(mFila, 3).Formula = Format(mRec!FechaAsist, "m/d/yyyy")
      XLS.Cells(mFila, 4).Formula = mRec!HoraIniAsist
      XLS.Cells(mFila, 5).Formula = mRec!HoraFinAsist
      XLS.Cells(mFila, 6).Formula = mRec!CodEdificio
      XLS.Cells(mFila, 7).Formula = mRec!DescripSolic
      XLS.Cells(mFila, 8).Formula = mRec!Observaciones
      XLS.Cells(mFila, IIf(Left(mRec!TipoMant, 1) = "C", 11, IIf(Left(mRec!TipoMant, 4) = "Prev", 9, 10))).Formula = "x"
      XLS.Cells(mFila, 12).Formula = mObj.sCampoDescrip("Rubros", "Codigo = '" & mRec!CodRubro & "'", 1)
      XLS.Cells(mFila, 13).Formula = mRec!TiempoEstim
      XLS.Cells(mFila, 14).Formula = mRec!TiempoReal
      XLS.Cells(mFila, 15).Formula = mRec!TiempoAdmis
      If Not IsNull(mRec!FechaAsist) Then
         If mRec!TiempoReal <= mRec!TiempoEstim Then
            XLS.Range("P" & mFila & ":P" & mFila).Interior.ColorIndex = 4
         Else
            If mRec!TiempoReal > mRec!TiempoEstim And mRec!TiempoReal <= mRec!TiempoAdmis Then
               XLS.Range("P" & mFila & ":P" & mFila).Interior.ColorIndex = 6
            Else
               If mRec!TiempoReal > mRec!TiempoEstim And mRec!TiempoReal > mRec!TiempoAdmis Then
                  XLS.Range("P" & mFila & ":P" & mFila).Interior.ColorIndex = 3
               End If
            End If
         End If
      End If
      XLS.Cells(mFila, 17).Formula = IIf(InStr(1, mRec!ManoObra, "AP-") > 0, "x", "")
      XLS.Cells(mFila, 18).Formula = IIf(InStr(1, mRec!ManoObra, "DC-") > 0, "x", "")
      XLS.Cells(mFila, 19).Formula = IIf(InStr(1, mRec!ManoObra, "DM-") > 0, "x", "")
      XLS.Cells(mFila, 20).Formula = IIf(InStr(1, mRec!ManoObra, "HC-") > 0, "x", "")
      XLS.Cells(mFila, 21).Formula = IIf(InStr(1, mRec!ManoObra, "HS-") > 0, "x", "")
      XLS.Cells(mFila, 22).Formula = IIf(InStr(1, mRec!ManoObra, "LO-") > 0, "x", "")
      XLS.Cells(mFila, 23).Formula = IIf(InStr(1, mRec!ManoObra, "MP-") > 0, "x", "")
      XLS.Cells(mFila, 24).Formula = IIf(InStr(1, mRec!ManoObra, "NG-") > 0, "x", "")
      XLS.Cells(mFila, 25).Formula = mRec!Horas
      XLS.Cells(mFila, 26).Formula = mRec!Pesos
      XLS.Cells(mFila, 27).Formula = mRec!Materiales
      mFila = mFila + 1
      mRec.MoveNext
   Loop
End If

XLS.Range("A1:AA7").Font.Bold = True
XLS.Range("A1:A2").Font.Size = 14
XLS.Range("A1:AA7").Font.Bold = True
XLS.Range("A4:A4").Font.Size = 12
XLS.Columns("A:A").ColumnWidth = 7
XLS.Columns("B:C").ColumnWidth = 11
XLS.Columns("D:E").ColumnWidth = 7
XLS.Columns("F:F").ColumnWidth = 21
XLS.Columns("G:H").ColumnWidth = 30
XLS.Columns("I:K").ColumnWidth = 10
XLS.Columns("L:L").ColumnWidth = 18
XLS.Columns("M:O").ColumnWidth = 10
XLS.Columns("P:P").ColumnWidth = 9
XLS.Columns("Q:X").ColumnWidth = 3
XLS.Columns("Y:Z").ColumnWidth = 8
XLS.Columns("AA:AA").ColumnWidth = 18

XLS.Range("A8:AA" & mFila).WrapText = True
XLS.Range("A8:A" & mFila).HorizontalAlignment = xlCenter
XLS.Range("A8:AA" & mFila).VerticalAlignment = xlCenter
XLS.Range("I8:K" & mFila).HorizontalAlignment = xlCenter
XLS.Range("Q8:V" & mFila).HorizontalAlignment = xlCenter

XLS.Rows("8:" & mFila).EntireRow.AutoFit
XLS.Range("A6:AA" & mFila).Borders(xlEdgeTop).LineStyle = xlContinuous
XLS.Range("A6:AA" & mFila).Borders(xlEdgeBottom).LineStyle = xlContinuous
XLS.Range("A6:AA" & mFila).Borders(xlEdgeLeft).LineStyle = xlContinuous
XLS.Range("A6:AA" & mFila).Borders(xlEdgeRight).LineStyle = xlContinuous
XLS.Range("A6:AA" & mFila).Borders(xlInsideVertical).LineStyle = xlContinuous
XLS.Range("A6:AA" & mFila).Borders(xlInsideHorizontal).LineStyle = xlContinuous

XLS.Visible = True
Set XLS = Nothing
Screen.MousePointer = vbArrow
End Sub


Private Sub Generar_XLS_Comp_X_Zonas()
Dim mFila As Double
Screen.MousePointer = vbHourglass
Set XLS = CreateObject("Excel.Application")
XLS.WorkBooks.Add

Dim mI As Integer
Dim marrZonasEdilicias(3) As String
Dim marrDescriZonasEdilicias(3) As String

marrZonasEdilicias(0) = "Z1"
marrZonasEdilicias(1) = "Z2"
marrZonasEdilicias(2) = "Z3"
marrZonasEdilicias(3) = "Z4"

marrDescriZonasEdilicias(0) = "ZONA 1 TIGRE-DEBT-MRQZ-C.REAL"
marrDescriZonasEdilicias(1) = "ZONA CENTRAL BUEN AYRE-202-BELGRANO-197 "
marrDescriZonasEdilicias(2) = "ZONA 3 PILAR-DECALADO PILAR"
marrDescriZonasEdilicias(3) = "ZONA 4 CAMPANA-DECALADO CAMPANA"



For mI = XLS.Sheets.Count To UBound(marrZonasEdilicias)
   XLS.Sheets.Add
Next

For mI = 1 To UBound(marrZonasEdilicias) + 1


'----------------------------------------------------------------------------------------------------------------------------------------------


      XLS.Sheets(mI).Name = marrZonasEdilicias(mI - 1)
      XLS.Sheets(mI).Select
      XLS.Sheets(mI).PageSetup.Orientation = xlLandscape

      XLS.Cells(1, 1).Formula = "AUTOPISTAS DEL SOL S.A."
      XLS.Cells(2, 1).Formula = "AVANCE DE TAREAS - MANTENIMIENTO EDILICIO"
      XLS.Cells(4, 1).Formula = "Período: " & Text1(0).Text & " al " & Text1(1).Text & " - " & marrDescriZonasEdilicias(mI - 1)
      XLS.Range("A6:AA7").WrapText = True
      XLS.Range("A6:AA7").HorizontalAlignment = xlCenter
      XLS.Range("A6:AA7").VerticalAlignment = xlCenter
      XLS.Range("A6:A7").Merge
      XLS.Cells(6, 1).Formula = "No. TAREA"
      XLS.Range("B6:B7").Merge
      XLS.Cells(6, 2).Formula = "PEDIDO"
      XLS.Range("C6:C7").Merge
      XLS.Cells(6, 3).Formula = "ASISTIDO"
      XLS.Range("D6:D7").Merge
      XLS.Cells(6, 4).Formula = "HORA ING."
      XLS.Range("E6:E7").Merge
      XLS.Cells(6, 5).Formula = "HORA EGR."
      XLS.Range("F6:F7").Merge
      XLS.Cells(6, 6).Formula = "EDIFICIO"
      XLS.Range("G6:G7").Merge
      XLS.Cells(6, 7).Formula = "DESCRIPCION DEL PEDIDO"
      XLS.Range("H6:H7").Merge
      XLS.Cells(6, 8).Formula = "OBSERVACION"
      XLS.Range("I6:K6").Merge
      XLS.Cells(6, 9).Formula = "TIPO DE MANTENIMIENTO"
      XLS.Cells(7, 9).Formula = "Preventivo"
      XLS.Cells(7, 10).Formula = "Predictivo"
      XLS.Cells(7, 11).Formula = "Correctivo"
      XLS.Range("L6:L7").Merge
      XLS.Cells(6, 12).Formula = "RUBRO"
      XLS.Range("M6:O6").Merge
      XLS.Cells(6, 13).Formula = "TIEMPO DE RESPUESTA"
      XLS.Cells(7, 13).Formula = "Estimado"
      XLS.Cells(7, 14).Formula = "Real"
      XLS.Cells(7, 15).Formula = "Admisible"
      XLS.Range("P6:P7").Merge
      XLS.Cells(6, 16).Formula = "ALERTAS"
      XLS.Range("Q6:Z6").Merge
      XLS.Cells(6, 17).Formula = "MANO DE OBRA"
      XLS.Cells(7, 17).Formula = "AP"
      XLS.Cells(7, 18).Formula = "DC"
      XLS.Cells(7, 19).Formula = "DM"
      XLS.Cells(7, 20).Formula = "HC"
      XLS.Cells(7, 21).Formula = "HS"
      XLS.Cells(7, 22).Formula = "LO"
      XLS.Cells(7, 23).Formula = "MP"
      XLS.Cells(7, 24).Formula = "NG"
      XLS.Cells(7, 25).Formula = "Hs/Per"
      XLS.Cells(7, 26).Formula = "$/Tarea"
      'XLS.Range("Y6:Y7").Merge
      XLS.Range("AA6:AA7").Merge
      XLS.Cells(6, 27).Formula = "Materiales"
      
      'Set mRec = mObj.oEjecutarSelect("SELECT * FROM Registros WHERE FechaSolic BETWEEN '" & Format(Text1(0).Text, "YYYY-MM-DD") & " 00:00:00" & "' And '" & Format(Text1(1).Text, "YYYY-MM-DD") & " 23:59:59" & "' ORDER BY Parte")
      
      Set mRec = mObj.oEjecutarSelect("SELECT R.* FROM Registros R " & _
      " Left Join  AnulacionesParte A ON R.Parte = A.ParteAnu " & _
      " INNER JOIN Edificios E ON R.CodEdificio = CONCAT(E.ZonaMantEdil , ' - ' , E.Descripcion) " & _
      " WHERE FechaSolic BETWEEN '" & Format(Text1(0).Text, "YYYY-MM-DD") & " 00:00:00" & "' And '" & Format(Text1(1).Text, "YYYY-MM-DD") & " 23:59:59" & "' AND E.ZonaMantEdil ='" & marrZonasEdilicias(mI - 1) & "' AND A.ParteAnu IS NULL  ORDER BY Parte")
      
      If Not mRec.EOF Then
         mFila = 8
         Do While Not mRec.EOF
            XLS.Cells(mFila, 1).Formula = mRec!Parte
            If mRec!TiempoEstim <= 3 Then
               XLS.Range("A" & mFila & ":A" & mFila).Interior.ColorIndex = 3
            End If
            XLS.Cells(mFila, 2).Formula = Format(mRec!FechaSolic, "m/d/yyyy")
            If IsNull(mRec!FechaAsist) Then
               If DateDiff("d", mRec!FechaSolic + mRec!TiempoAdmis, Date) > 0 Then
                  XLS.Range("B" & mFila & ":B" & mFila).Interior.ColorIndex = 44
               Else
                  If DateDiff("d", mRec!FechaSolic + mRec!TiempoEstim, Date) > 0 Then
                     XLS.Range("B" & mFila & ":B" & mFila).Interior.ColorIndex = 6
                  End If
               End If
            End If
            XLS.Cells(mFila, 3).Formula = Format(mRec!FechaAsist, "m/d/yyyy")
            XLS.Cells(mFila, 4).Formula = mRec!HoraIniAsist
            XLS.Cells(mFila, 5).Formula = mRec!HoraFinAsist
            XLS.Cells(mFila, 6).Formula = mRec!CodEdificio
            XLS.Cells(mFila, 7).Formula = mRec!DescripSolic
            XLS.Cells(mFila, 8).Formula = mRec!Observaciones
            XLS.Cells(mFila, IIf(Left(mRec!TipoMant, 1) = "C", 11, IIf(Left(mRec!TipoMant, 4) = "Prev", 9, 10))).Formula = "x"
            XLS.Cells(mFila, 12).Formula = mObj.sCampoDescrip("Rubros", "Codigo = '" & mRec!CodRubro & "'", 1)
            XLS.Cells(mFila, 13).Formula = mRec!TiempoEstim
            XLS.Cells(mFila, 14).Formula = mRec!TiempoReal
            XLS.Cells(mFila, 15).Formula = mRec!TiempoAdmis
            If Not IsNull(mRec!FechaAsist) Then
               If mRec!TiempoReal <= mRec!TiempoEstim Then
                  XLS.Range("P" & mFila & ":P" & mFila).Interior.ColorIndex = 4
               Else
                  If mRec!TiempoReal > mRec!TiempoEstim And mRec!TiempoReal <= mRec!TiempoAdmis Then
                     XLS.Range("P" & mFila & ":P" & mFila).Interior.ColorIndex = 6
                  Else
                     If mRec!TiempoReal > mRec!TiempoEstim And mRec!TiempoReal > mRec!TiempoAdmis Then
                        XLS.Range("P" & mFila & ":P" & mFila).Interior.ColorIndex = 3
                     End If
                  End If
               End If
            End If
            XLS.Cells(mFila, 17).Formula = IIf(InStr(1, mRec!ManoObra, "AP-") > 0, "x", "")
            XLS.Cells(mFila, 18).Formula = IIf(InStr(1, mRec!ManoObra, "DC-") > 0, "x", "")
            XLS.Cells(mFila, 19).Formula = IIf(InStr(1, mRec!ManoObra, "DM-") > 0, "x", "")
            XLS.Cells(mFila, 20).Formula = IIf(InStr(1, mRec!ManoObra, "HC-") > 0, "x", "")
            XLS.Cells(mFila, 21).Formula = IIf(InStr(1, mRec!ManoObra, "HS-") > 0, "x", "")
            XLS.Cells(mFila, 22).Formula = IIf(InStr(1, mRec!ManoObra, "LO-") > 0, "x", "")
            XLS.Cells(mFila, 23).Formula = IIf(InStr(1, mRec!ManoObra, "MP-") > 0, "x", "")
            XLS.Cells(mFila, 24).Formula = IIf(InStr(1, mRec!ManoObra, "NG-") > 0, "x", "")
            XLS.Cells(mFila, 25).Formula = mRec!Horas
            XLS.Cells(mFila, 26).Formula = mRec!Pesos
            XLS.Cells(mFila, 27).Formula = mRec!Materiales
            mFila = mFila + 1
            mRec.MoveNext
         Loop
      End If
      
      XLS.Range("A1:AA7").Font.Bold = True
      XLS.Range("A1:A2").Font.Size = 14
      XLS.Range("A1:AA7").Font.Bold = True
      XLS.Range("A4:A4").Font.Size = 12
      XLS.Columns("A:A").ColumnWidth = 7
      XLS.Columns("B:C").ColumnWidth = 11
      XLS.Columns("D:E").ColumnWidth = 7
      XLS.Columns("F:F").ColumnWidth = 21
      XLS.Columns("G:H").ColumnWidth = 30
      XLS.Columns("I:K").ColumnWidth = 10
      XLS.Columns("L:L").ColumnWidth = 18
      XLS.Columns("M:O").ColumnWidth = 10
      XLS.Columns("P:P").ColumnWidth = 9
      XLS.Columns("Q:X").ColumnWidth = 3
      XLS.Columns("Y:Z").ColumnWidth = 8
      XLS.Columns("AA:AA").ColumnWidth = 18
      
      XLS.Range("A8:AA" & mFila).WrapText = True
      XLS.Range("A8:A" & mFila).HorizontalAlignment = xlCenter
      XLS.Range("A8:AA" & mFila).VerticalAlignment = xlCenter
      XLS.Range("I8:K" & mFila).HorizontalAlignment = xlCenter
      XLS.Range("Q8:V" & mFila).HorizontalAlignment = xlCenter
      
      XLS.Rows("8:" & mFila).EntireRow.AutoFit
      XLS.Range("A6:AA" & mFila).Borders(xlEdgeTop).LineStyle = xlContinuous
      XLS.Range("A6:AA" & mFila).Borders(xlEdgeBottom).LineStyle = xlContinuous
      XLS.Range("A6:AA" & mFila).Borders(xlEdgeLeft).LineStyle = xlContinuous
      XLS.Range("A6:AA" & mFila).Borders(xlEdgeRight).LineStyle = xlContinuous
      XLS.Range("A6:AA" & mFila).Borders(xlInsideVertical).LineStyle = xlContinuous
      XLS.Range("A6:AA" & mFila).Borders(xlInsideHorizontal).LineStyle = xlContinuous
      
'----------------------------------------------------------------------------------------------------------------------------------------------
Next


XLS.Sheets(1).Select
XLS.Visible = True

mRec.Close
Set mRec = Nothing

XLS.Visible = True
Set XLS = Nothing
Screen.MousePointer = vbArrow
End Sub


Private Sub Generar_XLS_Verif()
Dim mFila As Double
Screen.MousePointer = vbHourglass
Set XLS = CreateObject("Excel.Application")
XLS.WorkBooks.Add

XLS.Cells(1, 1).Formula = "AUTOPISTA DEL SOL S.A."
XLS.Cells(2, 1).Formula = "VERIFICACIÓN DE TRABAJOS REALIZADOS"
XLS.Cells(4, 1).Formula = "Período: " & Text1(0).Text & " al " & Text1(1).Text
XLS.Cells(6, 1).Formula = "Parte"
XLS.Cells(6, 2).Formula = "Fecha Solicitud"
XLS.Cells(6, 3).Formula = "Lugar"
XLS.Cells(6, 4).Formula = "Descripción de la Solicitud"
XLS.Cells(6, 5).Formula = "Fecha Fin Asistencia"
XLS.Cells(6, 6).Formula = "Comentarios Mant. Edilicio"
XLS.Cells(6, 7).Formula = "Observaciones Supervisión"
XLS.Cells(6, 8).Formula = "Estado"
mFila = 7

'Set mRec = mObj.oEjecutarSelect("SELECT * FROM Registros WHERE FechaSolic BETWEEN '" & Format(Text1(0).Text, "YYYY-MM-DD") & " 00:00:00" & "' And '" & Format(Text1(1).Text, "YYYY-MM-DD") & " 23:59:59" & "' And Estado IN ('R', 'A') And Origen = 'S' ORDER BY Parte") ' trae los anulados tambien

Set mRec = mObj.oEjecutarSelect(" SELECT R.* FROM Registros R " & _
" Left Join  AnulacionesParte A ON R.Parte = A.ParteAnu  " & _
" WHERE FechaSolic BETWEEN '" & Format(Text1(0).Text, "YYYY-MM-DD") & " 00:00:00" & "' And '" & Format(Text1(1).Text, "YYYY-MM-DD") & " 23:59:59" & "' And Estado IN ('R', 'A') And Origen = 'S' AND A.ParteAnu IS NULL ORDER BY Parte")



If Not mRec.EOF Then
   Do While Not mRec.EOF
      XLS.Cells(mFila, 1).Formula = mRec!Parte
      XLS.Cells(mFila, 2).Formula = Format(mRec!FechaSolic, "MM/DD/YYYY HH:MM:SS")
      XLS.Cells(mFila, 3).Formula = mRec!CodEdificio
      XLS.Cells(mFila, 4).Formula = mRec!DescripSolic
      XLS.Cells(mFila, 5).Formula = Format(mRec!FechaAsist, "MM/DD/YYYY") & " " & mRec!HoraFinAsist
      XLS.Cells(mFila, 6).Formula = mRec!Observaciones
      XLS.Cells(mFila, 7).Formula = mRec!ObservVal
      XLS.Cells(mFila, 8).Formula = IIf(mRec!estado = "R", "Rechazado", "Aceptado")
      mFila = mFila + 1
      mRec.MoveNext
   Loop
End If

XLS.Range("A1:H6").Font.Bold = True
XLS.Range("A1:A2").Font.Size = 14
XLS.Range("A4:A4").Font.Size = 12
XLS.Columns("A:A").ColumnWidth = 7
XLS.Columns("B:B").ColumnWidth = 20
XLS.Columns("C:C").ColumnWidth = 23
XLS.Columns("D:D").ColumnWidth = 40
XLS.Columns("E:E").ColumnWidth = 20
XLS.Columns("F:G").ColumnWidth = 40
XLS.Columns("H:H").ColumnWidth = 10

XLS.Range("A6:H" & mFila - 1).WrapText = True
XLS.Range("A6:A" & mFila - 1).HorizontalAlignment = xlCenter
XLS.Range("A6:H" & mFila - 1).VerticalAlignment = xlCenter
XLS.Rows("6:" & mFila - 1).EntireRow.AutoFit

XLS.Visible = True
Set XLS = Nothing
Screen.MousePointer = vbArrow
End Sub
