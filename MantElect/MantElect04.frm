VERSION 5.00
Begin VB.Form MantElect04 
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
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   2760
      TabIndex        =   0
      Top             =   1020
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   2760
      TabIndex        =   1
      Top             =   1740
      Width           =   975
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
      TabIndex        =   6
      Top             =   120
      Width           =   1785
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
      TabIndex        =   4
      Top             =   1800
      Width           =   1095
   End
End
Attribute VB_Name = "MantElect04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mObj As New clMantElect
Dim mRec As New ADODB.Recordset
Dim WorkBooks As Object
Dim XLS As EXCEL.Application
Public mReporte As Integer

Private Sub Command1_Click(Index As Integer)
If Index = 0 Then
   If fValida Then
      Select Case mReporte
         Case 0
            GenXLS_Detalle
            
         Case 1
            GenXLS_Diario
         Case 2
            GenXLS_Detalle_Nuevo_Formato
      End Select
   End If
Else
   Unload Me
End If
End Sub

Private Sub Form_Load()
sAlinearForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set mObj = Nothing
Set mRec = Nothing
ShowMenu 47, True, False
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

Private Sub GenXLS_Detalle()
Dim mFila As Double
Screen.MousePointer = vbHourglass
Set XLS = CreateObject("Excel.Application")
XLS.WorkBooks.Add

XLS.Cells(1, 1).Formula = "AUTOPISTAS DEL SOL S.A."
XLS.Cells(2, 1).Formula = "SOLICITUDES DE SERVICIO AL SECTOR DE MANTENIMIENTO ELECTRICO"
XLS.Cells(4, 1).Formula = "Período: " & Text1(0).Text & " al " & Text1(1).Text

XLS.Range("A6:W7").WrapText = True
XLS.Range("A6:W7").HorizontalAlignment = xlCenter
XLS.Range("A6:W7").VerticalAlignment = xlCenter
XLS.Range("A6:A7").Merge
XLS.Cells(6, 1).Formula = "Parte"
XLS.Range("B6:B7").Merge
XLS.Cells(6, 2).Formula = "Fecha Solicitud"
XLS.Range("C6:C7").Merge
XLS.Cells(6, 3).Formula = "Lugar"
XLS.Range("D6:D7").Merge
XLS.Cells(6, 4).Formula = "Descripcion Solicitud"
XLS.Range("E6:E7").Merge
XLS.Cells(6, 5).Formula = "Prioridad"
XLS.Range("F6:F7").Merge
XLS.Cells(6, 6).Formula = "Supervisión"
XLS.Cells(6, 7).Formula = "Fecha Asistencia"
XLS.Range("G6:H6").Merge
XLS.Cells(7, 7).Formula = "Inicial"
XLS.Cells(7, 8).Formula = "Final"
XLS.Range("I6:I7").Merge
XLS.Cells(6, 9).Formula = "Segunda Descripción"
XLS.Range("J6:J7").Merge
XLS.Cells(6, 10).Formula = "Rubro"
XLS.Range("K6:K7").Merge
XLS.Cells(6, 11).Formula = "Sub Rubro"
XLS.Range("L6:L7").Merge
XLS.Cells(6, 12).Formula = "Unidad"
XLS.Range("M6:M7").Merge
XLS.Cells(6, 13).Formula = "Interv"

XLS.Range("N6:N7").Merge
XLS.Cells(6, 14).Formula = "Cantidad"
XLS.Range("O6:O7").Merge
XLS.Cells(6, 15).Formula = "Horas"
XLS.Range("P6:P7").Merge
XLS.Cells(6, 16).Formula = "Estado"
XLS.Range("Q6:Q7").Merge
XLS.Cells(6, 17).Formula = "Op. Generación"
XLS.Range("R6:R7").Merge
XLS.Cells(6, 18).Formula = "Op. Proceso"
XLS.Range("S6:S7").Merge
XLS.Cells(6, 19).Formula = "Fecha de Proceso"
XLS.Range("T6:T7").Merge
XLS.Cells(6, 20).Formula = "Op. Terminacion"
XLS.Range("U6:U7").Merge
XLS.Cells(6, 21).Formula = "Fecha Terminación"
XLS.Range("V6:V7").Merge
XLS.Cells(6, 22).Formula = "Validación - Observaciones"
XLS.Range("W6:W7").Merge
XLS.Cells(6, 23).Formula = "Op. Validación"
XLS.Range("X6:X7").Merge
XLS.Cells(6, 24).Formula = "Fecha de Validación"

Set mRec = mObj.oEjecutarSelect("SELECT * FROM Registros WHERE FechaSolic BETWEEN '" & Format(Text1(0).Text, "YYYY-MM-DD") & " 00:00:00" & "' And '" & Format(Text1(1).Text, "YYYY-MM-DD") & " 23:59:59" & "' ORDER BY Parte")
If Not mRec.EOF Then
   mFila = 8
   Do While Not mRec.EOF
      XLS.Cells(mFila, 1).Formula = mRec!Parte
      XLS.Cells(mFila, 2).Formula = "'" & Format(mRec!FechaSolic, "DD/MM/YYYY HH:MM:SS")
      XLS.Cells(mFila, 3).Formula = mRec!CodEdificio
      XLS.Cells(mFila, 4).Formula = mRec!descripcion
      XLS.Cells(mFila, 5).Formula = mRec!Prioridad
      XLS.Cells(mFila, 6).Formula = mRec!CodSuperv
      XLS.Cells(mFila, 7).Formula = "'" & Format(mRec!FechaIniAsist, "DD/MM/YYYY HH:MM:SS")
      XLS.Cells(mFila, 8).Formula = "'" & Format(mRec!FechaFinAsist, "DD/MM/YYYY HH:MM:SS")
      XLS.Cells(mFila, 9).Formula = mRec!SegundaDesc
      XLS.Cells(mFila, 10).Formula = NVL(mRec!Rubro, "") & " - " & mObj.sCampoDescrip("Rubros", "Codigo = '" & mRec!Rubro & "'", 1)
      XLS.Cells(mFila, 11).Formula = NVL(mRec!SubRubro, "") & " - " & mObj.sCampoDescrip("SubRubros", "Codigo = '" & mRec!SubRubro & "'", 2)
      XLS.Cells(mFila, 12).Formula = mObj.sCampoDescrip("SubRubros", "Codigo = '" & mRec!SubRubro & "'", 3)
      XLS.Cells(mFila, 13).Formula = mRec!Interv
      XLS.Cells(mFila, 14).Formula = mRec!Cantidad
      XLS.Cells(mFila, 15).Formula = mRec!Horas
      XLS.Cells(mFila, 16).Formula = mRec!estado
      XLS.Cells(mFila, 17).Formula = mRec!OpGen
      XLS.Cells(mFila, 18).Formula = mRec!OpPro
      XLS.Cells(mFila, 19).Formula = "'" & Format(mRec!FecPro, "DD/MM/YYYY HH:MM:SS")
      XLS.Cells(mFila, 20).Formula = mRec!OpTer
      XLS.Cells(mFila, 21).Formula = "'" & Format(mRec!FecTer, "DD/MM/YYYY HH:MM:SS")
      XLS.Cells(mFila, 22).Formula = NVL(mRec!ObserValid, "")
      XLS.Cells(mFila, 23).Formula = NVL(mRec!OpVal, "")
      XLS.Cells(mFila, 24).Formula = "'" & Format(mRec!FecVal, "DD/MM/YYYY HH:MM:SS")
      mFila = mFila + 1
      mRec.MoveNext
   Loop
End If

XLS.Range("A1:W7").Font.Bold = True
XLS.Range("A1:A2").Font.Size = 14
XLS.Range("A1:W7").Font.Bold = True
XLS.Range("A4:A4").Font.Size = 12
XLS.Columns("A:A").ColumnWidth = 6
XLS.Columns("B:B").ColumnWidth = 18
XLS.Columns("C:C").ColumnWidth = 31
XLS.Columns("D:D").ColumnWidth = 50
XLS.Columns("E:E").ColumnWidth = 9
XLS.Columns("F:F").ColumnWidth = 15
XLS.Columns("G:H").ColumnWidth = 18
XLS.Columns("I:I").ColumnWidth = 50
XLS.Columns("J:J").ColumnWidth = 28
XLS.Columns("K:K").ColumnWidth = 50
XLS.Columns("L:L").ColumnWidth = 7
XLS.Columns("M:N").ColumnWidth = 9
XLS.Columns("O:O").ColumnWidth = 6
XLS.Columns("P:P").ColumnWidth = 7
XLS.Columns("Q:Q").ColumnWidth = 15
XLS.Columns("R:R").ColumnWidth = 15
XLS.Columns("S:S").ColumnWidth = 18
XLS.Columns("T:T").ColumnWidth = 15
XLS.Columns("U:U").ColumnWidth = 18
XLS.Columns("V:V").ColumnWidth = 50
XLS.Columns("W:W").ColumnWidth = 15
XLS.Columns("X:X").ColumnWidth = 18

XLS.Range("D8:D" & mFila - 1).WrapText = True
XLS.Range("I8:I" & mFila - 1).WrapText = True
XLS.Range("K8:K" & mFila - 1).WrapText = True
XLS.Range("V8:V" & mFila - 1).WrapText = True

XLS.Visible = True
Set XLS = Nothing
Screen.MousePointer = vbArrow
End Sub


Private Sub GenXLS_Detalle_Nuevo_Formato()
Dim mFila As Double
Screen.MousePointer = vbHourglass
Set XLS = CreateObject("Excel.Application")
XLS.WorkBooks.Add

XLS.Cells(1, 1).Formula = "AUTOPISTAS DEL SOL S.A."
XLS.Cells(2, 1).Formula = "SOLICITUDES DE SERVICIO AL SECTOR DE MANTENIMIENTO ELECTRICO"
XLS.Cells(4, 1).Formula = "Período: " & Text1(0).Text & " al " & Text1(1).Text

XLS.Range("A6:W6").WrapText = True
XLS.Range("A6:W6").HorizontalAlignment = xlCenter
XLS.Range("A6:W6").VerticalAlignment = xlCenter

XLS.Cells(6, 1).Formula = "Parte"
XLS.Cells(6, 2).Formula = "Fecha-Hora Solicitud"
XLS.Cells(6, 3).Formula = "Fecha Solicitud"
XLS.Cells(6, 4).Formula = "Descripcion Solicitud"
XLS.Cells(6, 5).Formula = "Prioridad"
XLS.Cells(6, 6).Formula = "Supervision"
XLS.Cells(6, 7).Formula = "Edificio"
XLS.Cells(6, 8).Formula = "Fecha-Hora Cierre"
XLS.Cells(6, 9).Formula = "Fecha Cierre"
XLS.Cells(6, 10).Formula = "Area"
XLS.Cells(6, 11).Formula = "Rubro"
XLS.Cells(6, 12).Formula = "SubRubro"
XLS.Cells(6, 13).Formula = "Unidad"
XLS.Cells(6, 14).Formula = "Interv"
XLS.Cells(6, 15).Formula = "Cantidad"
XLS.Cells(6, 16).Formula = "Horas"
XLS.Cells(6, 17).Formula = "Estados"
XLS.Cells(6, 18).Formula = "Op. Generación"
XLS.Cells(6, 19).Formula = "Op. Proceso"
XLS.Cells(6, 20).Formula = "Fecha Terminacion"
XLS.Cells(6, 21).Formula = "Validacion - Observaciones"
XLS.Cells(6, 22).Formula = "Op. Validación"
XLS.Cells(6, 23).Formula = "Fecha Validación"

Set mRec = mObj.oEjecutarSelect(" SELECT  Parte, " _
& "                                 FechaSolic AS FechaHoraSolic," _
& "                                 DATE_FORMAT(FechaSolic,'%d/%m/%Y') as FechaSolic," _
& "                                 R.Descripcion, Prioridad," _
& "                                 CodSuperv," _
& "                                 TRIM(SUBSTRING(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(CodEdificio,'- Vías',''),' - Edificio',''),'ASC',''),'DESC',''),'Decalado',''),6)) as Edificio," _
& "                                 FechaFinAsist AS FechaHoraFinAsist, " _
& "                                 DATE_FORMAT(FechaFinAsist,'%d/%m/%Y') as FechaFinAsist, " _
& "                                 CASE  WHEN SectorAire = 1 THEN 'Aire Acondicionado' ELSE 'Eléctrico' END AS Area, " _
& "                                 CONCAT(U.Codigo,' - ', U.Descripcion) as Rubro, " _
& "                                 CONCAT(S.Codigo,' - ', S.Descripcion) AS SubRubro, " _
& "                                 Unidad, Interv, Cantidad, Horas, Estado, OpGen, OpPro," _
& "                                 FecTer AS FechaHoraTer , DATE_FORMAT(FecTer,'%d/%m/%Y') as FecTer, " _
& "                                 ObserValid, OpVal, " _
& "                                 DATE_FORMAT(FecVal,'%d/%m/%Y') as FecVal  " _
& "                                FROM MantElect.Registros R " _
& "                                Left Join MantElect.Rubros U ON U.Codigo = R.Rubro " _
& "                                Left Join MantElect.SubRubros S ON S.Codigo = R.SubRubro " _
& "                                WHERE CodSuperv NOT IN ('soccovi','satclie','srelev','sbalbin') " _
& "                                AND  FechaSolic BETWEEN '" & Format(Text1(0).Text, "YYYY-MM-DD") & " 00:00:00" & "' And '" & Format(Text1(1).Text, "YYYY-MM-DD") & " 23:59:59" & "' ORDER BY Parte; ")


If Not mRec.EOF Then
   mFila = 7
   Do While Not mRec.EOF
      XLS.Cells(mFila, 1).Formula = mRec!Parte
      XLS.Cells(mFila, 2).Formula = "'" & Format(mRec!FechaHoraSolic, "DD/MM/YYYY HH:MM:SS")
      XLS.Cells(mFila, 3).Formula = mRec!FechaSolic
      XLS.Cells(mFila, 4).Formula = mRec!descripcion
      XLS.Cells(mFila, 5).Formula = mRec!Prioridad
      XLS.Cells(mFila, 6).Formula = mRec!CodSuperv
      XLS.Cells(mFila, 7).Formula = mRec!Edificio
      XLS.Cells(mFila, 8).Formula = "'" & Format(mRec!FechaHoraFinAsist, "DD/MM/YYYY HH:MM:SS")
      XLS.Cells(mFila, 9).Formula = mRec!FechaFinAsist
      XLS.Cells(mFila, 10).Formula = mRec!Area
      XLS.Cells(mFila, 11).Formula = mRec!Rubro
      XLS.Cells(mFila, 12).Formula = mRec!SubRubro
      XLS.Cells(mFila, 13).Formula = mRec!Unidad
      XLS.Cells(mFila, 14).Formula = mRec!Interv
      XLS.Cells(mFila, 15).Formula = mRec!Cantidad
      XLS.Cells(mFila, 16).Formula = mRec!Horas
      XLS.Cells(mFila, 17).Formula = mRec!estado
      XLS.Cells(mFila, 18).Formula = mRec!OpGen
      XLS.Cells(mFila, 19).Formula = mRec!OpPro
      XLS.Cells(mFila, 20).Formula = mRec!FecTer
      XLS.Cells(mFila, 21).Formula = NVL(mRec!ObserValid, "")
      XLS.Cells(mFila, 22).Formula = mRec!OpVal
      XLS.Cells(mFila, 23).Formula = mRec!FecVal
      
      mFila = mFila + 1
      mRec.MoveNext
   Loop
End If

XLS.Range("A1:W6").Font.Bold = True
XLS.Range("A1:A2").Font.Size = 14
XLS.Range("A1:W6").Font.Bold = True
XLS.Range("A4:A4").Font.Size = 12
XLS.Columns("A:A").ColumnWidth = 6
XLS.Columns("B:B").ColumnWidth = 20
XLS.Columns("C:C").ColumnWidth = 15
XLS.Columns("D:D").ColumnWidth = 90
XLS.Columns("E:E").ColumnWidth = 15
XLS.Columns("F:F").ColumnWidth = 15
XLS.Columns("G:G").ColumnWidth = 20
XLS.Columns("H:H").ColumnWidth = 20
XLS.Columns("I:I").ColumnWidth = 15
XLS.Columns("J:J").ColumnWidth = 25
XLS.Columns("K:K").ColumnWidth = 40
XLS.Columns("L:L").ColumnWidth = 50
XLS.Columns("M:M").ColumnWidth = 9
XLS.Columns("N:N").ColumnWidth = 9
XLS.Columns("O:O").ColumnWidth = 10
XLS.Columns("P:P").ColumnWidth = 15
XLS.Columns("Q:Q").ColumnWidth = 10
XLS.Columns("R:R").ColumnWidth = 15
XLS.Columns("S:S").ColumnWidth = 18
XLS.Columns("T:T").ColumnWidth = 15
XLS.Columns("U:U").ColumnWidth = 50
XLS.Columns("V:V").ColumnWidth = 18
XLS.Columns("W:W").ColumnWidth = 15
XLS.Columns("X:X").ColumnWidth = 15

XLS.Range("D8:D" & mFila - 1).WrapText = True
XLS.Range("I8:I" & mFila - 1).WrapText = True
XLS.Range("K8:K" & mFila - 1).WrapText = True
XLS.Range("V8:V" & mFila - 1).WrapText = True

XLS.Visible = True
Set XLS = Nothing
Screen.MousePointer = vbArrow
End Sub




















































Private Sub GenXLS_Diario()
Dim mFila As Double
Dim mI As Integer
Dim mLista As String
Dim mSelect As String
Dim mRubro As String
Dim mFilaSum As Integer
Dim mJ As Integer

Dim mdiego As String

Screen.MousePointer = vbHourglass
Set XLS = CreateObject("Excel.Application")
XLS.WorkBooks.Add

For mI = XLS.Sheets.Count To DateDiff("d", Format(Text1(0).Text, "DD/MM/YYYY"), Format(Text1(1).Text, "DD/MM/YYYY")) + 1
   XLS.Sheets.Add
Next
mLista = "{"
For mI = 0 To DateDiff("d", Format(Text1(0).Text, "DD/MM/YYYY"), Format(Text1(1).Text, "DD/MM/YYYY")) + 1
   If mI <> DateDiff("d", Format(Text1(0).Text, "DD/MM/YYYY"), Format(Text1(1).Text, "DD/MM/YYYY")) + 1 Then
      XLS.Sheets(mI + 1).Name = Val(Mid(CDate(Text1(0).Text) + mI, 1, 2))
      XLS.Sheets(mI + 1).Select
      XLS.Cells(4, 1).Formula = "Día: " & CDate(Text1(0).Text) + mI
      mLista = mLista & Val(Mid(CDate(Text1(0).Text) + mI, 1, 2)) & Chr(92)
   Else
      XLS.Sheets(mI + 1).Name = "Resumen Cant"
      XLS.Sheets(mI + 1).Select
      XLS.Cells(4, 1).Formula = "Hoja Resumen"
   End If

   XLS.Cells(1, 1).Formula = "AUTOPISTA DEL SOL S.A."
   XLS.Cells(2, 1).Formula = "PARTE DIARIO DE MANTENIMIENTO"

   XLS.Cells(7, 1).Formula = "Grupo"
   XLS.Cells(7, 2).Formula = "Descripcion de Grupo"
   XLS.Cells(7, 3).Formula = "Cod. tarea"
   XLS.Cells(7, 4).Formula = "Descripcion de tarea"
   XLS.Cells(7, 5).Formula = "2da. Descripcion de tarea"
   XLS.Cells(7, 6).Formula = "Un"
   XLS.Cells(6, 6).Formula = "CANT TAREAS"
   XLS.Cells(6, 8).Formula = "MAT"
   XLS.Cells(6, 9).Formula = "MANO DE OBRA (Hs)"

   Range("F6:I6").Select
   
   With Selection
      .HorizontalAlignment = xlCenterAcrossSelection
   End With
   
   XLS.Cells(7, 7).Formula = "Interv."
   XLS.Cells(7, 8).Formula = "Nº Registro Salida de Depósito"
   XLS.Cells(7, 9).Formula = "Cant"
   XLS.Cells(7, 10).Formula = "Apellido"
   XLS.Cells(7, 11).Formula = "Normal"
   XLS.Cells(7, 12).Formula = "Extra"
   XLS.Cells(7, 13).Formula = "Total Normal"
   XLS.Cells(7, 14).Formula = "Total Extra"

   Range("H6:H7").Select
   With Selection.Interior
      .ColorIndex = 7
      .Pattern = xlSolid
      .PatternColorIndex = xlAutomatic
   End With
   
   XLS.Range("I6:N6").Merge
   Range("I6:N7").Select
   With Selection.Interior
      .ColorIndex = 40
      .Pattern = xlSolid
      .PatternColorIndex = xlAutomatic
   End With

   XLS.Range("O6:W6").Merge
   XLS.Cells(6, 15).Formula = "FLETES (Hs)"
   XLS.Cells(7, 15).Formula = "Avalos"
   XLS.Cells(7, 16).Formula = "Bidone"
   XLS.Cells(7, 17).Formula = "Fidanza"
   XLS.Cells(7, 18).Formula = "Lucero"
   XLS.Cells(7, 19).Formula = "Menseguez Alberto"
   XLS.Cells(7, 20).Formula = "Menseguez Pablo"
   XLS.Cells(7, 21).Formula = "Pedernera"
   XLS.Cells(7, 22).Formula = "Schefer"
   XLS.Cells(7, 23).Formula = "Temez"
   
   Range("X6:W7").Select
   With Selection.Interior
      .ColorIndex = 35
      .Pattern = xlSolid
      .PatternColorIndex = xlAutomatic
   End With
   Range("O6:W7").Borders(xlEdgeLeft).LineStyle = xlContinuous

   XLS.Range("X6:AC6").Merge
   XLS.Cells(6, 24).Formula = "EQUIPOS (Hs)"
   XLS.Cells(7, 24).Formula = "Minicargadora"
   XLS.Cells(7, 25).Formula = "Retropala Terex 44"
   XLS.Cells(7, 26).Formula = "Retropala JCB"
   XLS.Cells(7, 27).Formula = "Minicargadora con martillo"
   XLS.Cells(7, 28).Formula = "Minicargadora con implemento barredora"
   XLS.Cells(7, 29).Formula = "Motoniveladora"
   
   Range("X6:AC7").Select
   With Selection.Interior
      .ColorIndex = 19
      .Pattern = xlSolid
      .PatternColorIndex = xlAutomatic
   End With
   Range("X6:AC7").Borders(xlEdgeLeft).LineStyle = xlContinuous
   
   XLS.Range("AD6:AJ6").Merge
   XLS.Cells(6, 30).Formula = "SUBCONTRATOS (Hs)"
   XLS.Cells(7, 30).Formula = "Bidone Almeja"
   XLS.Cells(7, 31).Formula = "Autotrol (semaforización)"
   XLS.Cells(7, 32).Formula = "Depacom (Limpieza desagues)"
   XLS.Cells(7, 33).Formula = "Aserradora"
   XLS.Cells(7, 34).Formula = "Martillo Neumático"
   XLS.Cells(7, 35).Formula = "Rodillo Compactador"
   XLS.Cells(7, 36).Formula = "Grúas"

   Range("AD6:AJ7").Select
   With Selection.Interior
      .ColorIndex = 45
      .Pattern = xlSolid
      .PatternColorIndex = xlAutomatic
   End With
   Range("AD6:AJ7").Borders(xlEdgeLeft).LineStyle = xlContinuous
   Range("AD6:AJ7").Borders(xlEdgeRight).LineStyle = xlContinuous
   mFila = 8
   If mI <> DateDiff("d", Format(Text1(0).Text, "DD/MM/YYYY"), Format(Text1(1).Text, "DD/MM/YYYY")) + 1 Then
      Set mRec = mObj.oEjecutarSelect("SELECT * FROM Registros WHERE FechaFinAsist BETWEEN '" & Format(CDate(Text1(0).Text) + mI, "YYYY-MM-DD") & " 00:00:00' And '" & Format(CDate(Text1(0).Text) + mI, "YYYY-MM-DD") & " 23:59:59' And Estado <> 'G' ORDER BY Parte")
      If Not mRec.EOF Then
         Do While Not mRec.EOF
            XLS.Cells(mFila, 1).Formula = mRec!Rubro
            XLS.Cells(mFila, 2).Formula = mObj.sCampoDescrip("Rubros", "Codigo = '" & mRec!Rubro & "'", 1)
            XLS.Cells(mFila, 3).Formula = mRec!SubRubro
            XLS.Cells(mFila, 4).Formula = mObj.sCampoDescrip("SubRubros", "Codigo = '" & mRec!SubRubro & "'", 2)
            XLS.Cells(mFila, 5).Formula = mRec!SegundaDesc
            XLS.Cells(mFila, 6).Formula = mObj.sCampoDescrip("SubRubros", "Codigo = '" & mRec!SubRubro & "'", 3)
            XLS.Cells(mFila, 7).Formula = mRec!Interv
            XLS.Cells(mFila, 8).Formula = ""
            XLS.Cells(mFila, 9).Formula = mRec!Cantidad
            XLS.Cells(mFila, 10).Formula = ""
            XLS.Cells(mFila, 11).Formula = 0
            XLS.Cells(mFila, 12).Formula = 0
            XLS.Cells(mFila, 13).Formula = mRec!Horas
            XLS.Cells(mFila, 14).Formula = 0
            mFila = mFila + 1
            mRec.MoveNext
         Loop
      End If
      mRec.Close
   Else
      mLista = Mid(mLista, 1, Len(mLista) - 1)
      mLista = mLista & "}"
      mSelect = "SELECT RUBRO, SUBRUBRO, SUM(INTERV) INTERV, SUM(CANT) CANT, SUM(HORAS) HORAS FROM (SELECT RUBRO AS RUBRO, SUBRUBRO AS SUBRUBRO, SUM(INTERV) INTERV, SUM(CANTIDAD) AS CANT, SUM(HORAS) AS HORAS FROM Registros WHERE FechaFinAsist BETWEEN '" & Format(CDate(Text1(0).Text), "YYYY-MM-DD") & " 00:00:00' And '" & Format(CDate(Text1(1).Text), "YYYY-MM-DD") & " 23:59:59' And ESTADO <> 'G' GROUP BY RUBRO, SUBRUBRO UNION ALL SELECT CODRUBRO AS RUBRO, CODIGO AS SUBRUBRO, 0 AS INTERV, 0 AS CANT, 0 AS HORAS FROM SubRubros ORDER BY RUBRO, SUBRUBRO, CANT) AS b GROUP BY b.RUBRO, b.SUBRUBRO ORDER BY b.RUBRO, b.SUBRUBRO"
      Set mRec = mObj.oEjecutarSelect(mSelect)
      mRubro = ""
      If Not mRec.EOF Then
         Do While Not mRec.EOF
            mRubro = mRec!Rubro
            XLS.Cells(mFila, 1).Formula = mRec!Rubro
            XLS.Cells(mFila, 2).Formula = mObj.sCampoDescrip("Rubros", "Codigo = '" & mRec!Rubro & "'", 1)
            XLS.Range("A" & mFila & ":AG" & mFila).Font.Bold = True
            mFilaSum = mFila
            mFila = mFila + 1
            Do While Not mRec.EOF And mRubro = mRec!Rubro
               XLS.Cells(mFila, 1).Formula = mRec!SubRubro
               XLS.Cells(mFila, 2).Formula = mObj.sCampoDescrip("SubRubros", "Codigo = '" & mRec!SubRubro & "'", 2)
               XLS.Cells(mFila, 6).Formula = mObj.sCampoDescrip("SubRubros", "Codigo = '" & mRec!SubRubro & "'", 3)
               XLS.Cells(mFila, 7).Formula = mRec!Interv
               XLS.Cells(mFila, 9).Formula = mRec!Cant
               XLS.Cells(mFila, 13).Formula = mRec!Horas
               XLS.Cells(mFila, 14).Formula = 0

'mdiego = "{=SUMA(SUMAR.SI(INDIRECTO(di & " & """!C8:""" & " & SUSTITUIR(DIRECCION(FILA();COLUMNA() + 3; 4); FILA();" & """""" & ") & 50);$A9;INDIRECTO(di & " & """!""" & " & SUSTITUIR(DIRECCION(FILA();COLUMNA() + 3;4);FILA();" & """""" & ") & " & """8:""" & " & SUSTITUIR(DIRECCION(FILA();COLUMNA() + 3;4);FILA();" & """""" & ") & 50)))}"

'               XLS.Range("L" & mFila).Formula = mdiego
'               XLS.Range("L" & mFila).Formula = "'" & mId(XLS.Range("L" & mFila).Formula, 2, Len(XLS.Range("L" & mFila).Formula) - 2)
'               XLS.Range("L" & mFila).Select
'               SendKeys "{F2}^+{ENTER}"
'               DoEvents
               
'               Range("A1").Formula = "=Sum(A5:B7)" 'introducimos la formula normalmente (no es matricial, no tiene ese límite de 255)
'Range("a1").Select 'seleccionamos la celda antes de enviar las teclas
'SendKeys "{F2}^+{ENTER}"  'F2 para editar y ctrl (^) Mays (+)  Enter para introducir como matricial
'DoEvents ' para enviar las teclas al sistema



               'XLS.Range("L" & mFila).Formula = mdiego
 
'               SendKeys "{F2}", True

 '              SendKeys "^+{ENTER}", True
               
'               XLS.Cells(mFila, 15).Formula = mdiego
               
               
'               XLS.Range("O" & mFila).Select
'               XLS.Selection.FormulaArray = Right(mdiego, Len(mdiego) - 1)
'               XLS.Cells(mFila, 16).Formula = "=SUM(SUMIF(INDIRECT(dias & " & """!C8:""" & " & SUBSTITUTE(ADDRESS(ROW();COLUMN() + 3; 4); ROW();" & """""" & ") & 50);$A9;INDIRECT(dias & " & """!""" & " & SUBSTITUTE(ADDRESS(ROW();COLUMN() + 3;4); ROW();" & """""" & ") & " & """8:""" & " & SUBSTITUTE(ADDRESS(ROW();COLUMN() + 3;4); ROW();" & """""" & ") & 50)))"
               mFila = mFila + 1
               mRec.MoveNext
               If mRec.EOF Then
                  Exit Do
               End If
            Loop
            XLS.Cells(mFilaSum, 7).FormulaR1C1 = "=SUM(R[1]C:R[" & mFila - mFilaSum - 1 & "]C)"
            XLS.Cells(mFilaSum, 9).FormulaR1C1 = "=SUM(R[1]C:R[" & mFila - mFilaSum - 1 & "]C)"
            For mJ = 13 To 36
               XLS.Cells(mFilaSum, mJ).FormulaR1C1 = "=SUM(R[1]C:R[" & mFila - mFilaSum - 1 & "]C)"
            Next
            If mRec.EOF Then
               Exit Do
            End If
         Loop
      End If
      mRec.Close
      'XLS.ActiveWorkbook.Names.Add Name:="dias", RefersToR1C1:="=" & mLista
       XLS.ActiveWorkbook.Names.Add Name:="di", RefersToR1C1:="=" & mLista
   End If
   
   XLS.Range("A1:AJ7").Font.Bold = True
   XLS.Range("A1:A2").Font.Size = 14
   XLS.Range("A1:AJ7").Font.Bold = True
   XLS.Range("A4:A4").Font.Size = 12
   XLS.Columns("A:A").ColumnWidth = 8
   XLS.Columns("B:B").ColumnWidth = 20
   XLS.Columns("C:C").ColumnWidth = 6
   XLS.Columns("D:E").ColumnWidth = 30
   XLS.Columns("F:F").ColumnWidth = 4
   XLS.Columns("G:I").ColumnWidth = 30
   XLS.Columns("G:I").ColumnWidth = 9

   XLS.Range("A7:AJ7").Borders(xlEdgeBottom).LineStyle = xlContinuous
   XLS.Range("H7:H7").WrapText = True
   XLS.Range("H7:H7").Orientation = 90
   XLS.Range("O6:AJ6").HorizontalAlignment = xlCenter
   XLS.Range("O7:AJ7").WrapText = True
   XLS.Range("O7:AJ7").Orientation = 90

   XLS.Range("A7:AJ" & mFila - 1).WrapText = True
   XLS.Range("A7:AJ7").HorizontalAlignment = xlCenter
   XLS.Range("A7:AJ" & mFila - 1).VerticalAlignment = xlCenter
   XLS.Range("A7:AJ" & mFila - 1).Borders(xlEdgeTop).LineStyle = xlContinuous
   XLS.Range("A7:AJ" & mFila - 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
   XLS.Range("A7:AJ" & mFila - 1).Borders(xlEdgeLeft).LineStyle = xlContinuous
   XLS.Range("A7:AJ" & mFila - 1).Borders(xlEdgeRight).LineStyle = xlContinuous
   XLS.Range("A7:AJ" & mFila - 1).Borders(xlInsideVertical).LineStyle = xlContinuous
   XLS.Range("O6:AJ6").Borders(xlEdgeTop).LineStyle = xlContinuous
   XLS.Range("O6:AJ6").Borders(xlEdgeLeft).LineStyle = xlContinuous
   XLS.Range("O6:AJ6").Borders(xlEdgeRight).LineStyle = xlContinuous
   XLS.Rows("7:7").RowHeight = 90
   XLS.Cells(1, 1).Select
Next

XLS.Columns("B:B").ColumnWidth = 70
XLS.Columns("C:E").Select
With Selection
   Selection.Delete Shift:=xlToLeft
End With

'dsdss
'XLS.Sheets(5).Select
'XLS.Range("L9").Select
'SendKeys "{F2}+{ENTER}", True
'
'SendKeys "{HOME}", True
'SendKeys "{DEL}", True
'SendKeys "{END}", True
'SendKeys "{BKSP}", True
'
'SendKeys "^+{ENTER}", True



'dasd

XLS.Range("A1").Select

XLS.Sheets(1).Select



XLS.Visible = True


Set XLS = Nothing
Screen.MousePointer = vbArrow
End Sub

