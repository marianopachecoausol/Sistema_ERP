VERSION 5.00
Begin VB.Form MantElect22 
   Caption         =   "Partes de luminarias"
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
Attribute VB_Name = "MantElect22"
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
   
   Set mRec = mObj.getRPT_PartesLuminarias(CDate(Text1(0).Text), CDate(Text1(1).Text))
   
   Do While Not mRec.EOF
      With XLS
         .Cells(mi, 2).Formula = NVL(mRec!Parte, "")
         .Cells(mi, 3).Formula = NVL(mRec!Fecha, "")
         .Cells(mi, 4).Formula = NVL(mRec!Columna, "")
         .Cells(mi, 5).Formula = NVL(mRec!CantLuminariasApagadas, "")
         .Cells(mi, 6).Formula = NVL(mRec!Problema, "")
         .Cells(mi, 7).Formula = NVL(mRec!DetectadoPor, "")
         .Cells(mi, 8).Formula = NVL(mRec!Estado, "")
         .Cells(mi, 9).Formula = NVL(mRec!Ramal, "")
         
         .Cells(mi, 10).Formula = NVL(mRec!Acceso, "")

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
      .Worksheets(1).Name = "Partes Luminarias"
     
      .Columns("A:A").ColumnWidth = 2
      .Columns("B:B").ColumnWidth = 12 ' Parte
      .Columns("C:C").ColumnWidth = 12 ' Fecha
      .Columns("D:D").ColumnWidth = 13 ' Columna
      .Columns("E:E").ColumnWidth = 22 ' CantLuminariasApagadas
      .Columns("F:F").ColumnWidth = 26 ' Problema
      
      .Columns("G:G").ColumnWidth = 17 ' DetectadoPor
      
      .Columns("H:H").ColumnWidth = 15 ' Estado
      .Columns("I:I").ColumnWidth = 19 ' Ramal
      .Columns("J:J").ColumnWidth = 17 ' Acceso

      .Columns("B:B").HorizontalAlignment = xlHAlignCenter
      .Columns("C:C").HorizontalAlignment = xlHAlignCenter
      .Columns("D:D").HorizontalAlignment = xlHAlignCenter
      .Columns("E:E").HorizontalAlignment = xlHAlignCenter
      .Columns("F:F").HorizontalAlignment = xlHAlignCenter
      .Columns("G:G").HorizontalAlignment = xlHAlignCenter
      .Columns("H:H").HorizontalAlignment = xlHAlignCenter
      .Columns("I:I").HorizontalAlignment = xlHAlignCenter
      .Columns("J:J").HorizontalAlignment = xlHAlignCenter
            
      'Formateo OrdTrabajo: 000x
'      .Range("B9:B65536").Select
'      .Selection.NumberFormat = "000000"

      'Formateo NroVale: 000x
'      .Range("M9:M65536").Select
'      .Selection.NumberFormat = "000000000"

      'Formateo: Encabezd con Negrita.
      .Range("B8:J8").Select
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
      .Cells(3, 1).Formula = "REPORTE: Partes de Luminarias"

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

      .Cells(8, 2).Formula = "Parte"
      .Cells(8, 3).Formula = "Fecha"
      .Cells(8, 4).Formula = "Columna"
      .Cells(8, 5).Formula = "Luminarias apagadas"
      .Cells(8, 6).Formula = "Problema"
      .Cells(8, 7).Formula = "Detectado por"
      .Cells(8, 8).Formula = "Estado"
      .Cells(8, 9).Formula = "Ramal"
      .Cells(8, 10).Formula = "Acceso"

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
