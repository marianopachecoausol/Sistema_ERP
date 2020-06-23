VERSION 5.00
Begin VB.Form RAcc12 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Waze - Cierre de Accidentes."
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5505
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   5505
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Volver"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   1
      Left            =   2925
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1650
      Width           =   1290
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Generar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   0
      Left            =   1350
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1650
      Width           =   1290
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
      Left            =   3075
      MaxLength       =   10
      TabIndex        =   3
      Top             =   900
      Width           =   1250
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
      Left            =   1575
      MaxLength       =   10
      TabIndex        =   1
      Top             =   900
      Width           =   1250
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nros de fichas Vs Códigos alfanuméricos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   1
      Left            =   150
      TabIndex        =   2
      Top             =   225
      Width           =   4920
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   0
      Left            =   750
      TabIndex        =   0
      Top             =   975
      Width           =   735
   End
End
Attribute VB_Name = "RAcc12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Si este formulario se usa para varios reportes se puede usar TAG de command1
'O procesos públicos para iniciar el formulario

Private Sub Form_Load()
   sAlinearForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ShowMenu 2, True, False
End Sub

Private Sub Command1_Click(Index As Integer)
   Select Case Index
      Case 0
         If sValidFechaDesdeHasta(Text1(0).Text, Text1(1).Text) Then
            sFichasCodigos Text1(0).Text, Text1(1).Text
         End If
      Case 1
         Unload Me
   End Select
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = fDateKeyPress(Text1(Index), KeyAscii)
End Sub

Private Sub sFichasCodigos(ByVal pFecha1 As String, ByVal pFecha2 As String)
Dim mObj As New clRNov
Dim mRec As New ADODB.Recordset
Dim XLS As EXCEL.Application
Dim mI As Integer
   
   Set mRec = mObj.oTabla("enlace", "where fecha between '" & Format(pFecha1, "yyyy-mm-dd") & "' and '" & Format(pFecha2, "yyyy-mm-dd") & "' order by fecha")
   If Not mRec.EOF Then
      sMsgEspere Me, "Buscando datos...", True
      Set XLS = CreateObject("Excel.Application")
      XLS.WorkBooks.Add
      sHeaderXLS XLS
      mI = 5
      Do While Not mRec.EOF
         XLS.Cells(mI, 1) = mRec!Fecha
         XLS.Cells(mI, 2) = NVL(mRec!nroficha, "")
         XLS.Cells(mI, 3) = NVL(mRec!Codigo, "")
         mI = mI + 1
         mRec.MoveNext
      Loop
      XLS.Range("A5:A" & mI).Select
      XLS.Selection.NumberFormat = "d-mmm-yyyy"
      sMsgEspere Me, "", False
      XLS.Application.Visible = True
   Else
      MsgBox "No existen datos para el rango de fechas.", vbInformation, sMessage
   End If
   mRec.Close
   Set mObj = Nothing
   Set mRec = Nothing
   
End Sub

Private Sub sHeaderXLS(ByRef pObj As Object)
   With pObj
      .Worksheets(1).Select
      .Worksheets(1).Name = "Datos"
      .Cells(1, 1).Formula = "Fichas de Accidentes Vs Códigos alfanuméricos."
      .Cells(2, 1).Formula = "Período de consulta: " & Text1(0).Text & " al " & Text1(1).Text
      .Cells(4, 1).Formula = "Fecha"
      .Cells(4, 2).Formula = "Nro Ficha"
      .Cells(4, 3).Formula = "Código Alfa"
      .Rows("1:2").Select
      With .Selection.Font
          .Name = "Arial"
          .FontStyle = "Normal"
          .Size = 12
          .Strikethrough = False
          .Superscript = False
          .Subscript = False
          .OutlineFont = False
          .Shadow = False
          .Underline = xlUnderlineStyleNone
          .ColorIndex = xlAutomatic
      End With
   End With
End Sub
