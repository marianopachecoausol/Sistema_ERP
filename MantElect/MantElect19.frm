VERSION 5.00
Begin VB.Form MantElect19 
   Caption         =   "Estado de Partes"
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
Attribute VB_Name = "MantElect19"
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
   
   Set mRec = mObj.getRPT_Partes(CDate(Text1(0).Text), CDate(Text1(1).Text))
   
   Do While Not mRec.EOF
      With XLS
      
      .Cells(mi, 2).Formula = NVL(mRec!Parte, "")
      .Cells(mi, 3).Formula = NVL(mRec!FechaSolic, "")
      .Cells(mi, 4).Formula = NVL(mRec!IdOT, "")
      .Cells(mi, 5).Formula = NVL(mRec!Categoria, "")
      .Cells(mi, 6).Formula = NVL(mRec!CodEdificio, "")
      .Cells(mi, 7).Formula = NVL(mRec!Lugar, "")
      .Cells(mi, 8).Formula = NVL(mRec!descripcion, "")
      .Cells(mi, 9).Formula = NVL(mRec!EstadoDesc, "")
      .Cells(mi, 10).Formula = NVL(mRec!FechaFin, "")
      .Cells(mi, 11).Formula = NVL(mRec!DiasAbierto, "")
      .Cells(mi, 12).Formula = NVL(mRec!Prioridad, "")
      
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
      .Worksheets(1).Name = "Partes"
     
      .Columns("A:A").ColumnWidth = 2
      .Columns("B:B").ColumnWidth = 9  ' Parte
      .Columns("C:C").ColumnWidth = 13 ' Fec.Carga
      .Columns("D:D").ColumnWidth = 12 ' Ord. Trabajo
      .Columns("E:E").ColumnWidth = 13 ' Origen
      .Columns("F:F").ColumnWidth = 24 ' Zona/Ramal/Comunicado
      .Columns("G:G").ColumnWidth = 30 ' Lugar
      .Columns("H:H").ColumnWidth = 100 'Problema
      .Columns("I:I").ColumnWidth = 10  'Estado
      .Columns("J:J").ColumnWidth = 13  'Fecha Cierre
      .Columns("K:K").ColumnWidth = 12  'Días Abiertos
      .Columns("L:L").ColumnWidth = 11  'Prioridad
            
      .Columns("B:B").HorizontalAlignment = xlHAlignCenter
      .Columns("C:C").HorizontalAlignment = xlHAlignCenter
      .Columns("D:D").HorizontalAlignment = xlHAlignCenter
      .Columns("E:E").HorizontalAlignment = xlHAlignCenter
      .Columns("F:F").HorizontalAlignment = xlHAlignCenter
      .Columns("I:I").HorizontalAlignment = xlHAlignCenter
      .Columns("J:J").HorizontalAlignment = xlHAlignCenter
      .Columns("K:K").HorizontalAlignment = xlHAlignCenter
      .Columns("L:L").HorizontalAlignment = xlHAlignCenter

      'Formateo la Ord.Trabajo: 000x
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
      .Cells(3, 1).Formula = "REPORTE: Estado de Partes"

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

      .Cells(8, 2).Formula = "Partes"
      .Cells(8, 3).Formula = "Fecha Carga"
      .Cells(8, 4).Formula = "Ord.Trabajo"
      .Cells(8, 5).Formula = "Origen"
      .Cells(8, 6).Formula = "Zona/Ramal/Comunicado"
      .Cells(8, 7).Formula = "Lugar"
      .Cells(8, 8).Formula = "Problema"
      .Cells(8, 9).Formula = "Estado"
      .Cells(8, 10).Formula = "Fecha Cierre"
      .Cells(8, 11).Formula = "Días abierto"
      .Cells(8, 12).Formula = "Prioridad"


 
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
