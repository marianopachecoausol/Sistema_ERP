VERSION 5.00
Begin VB.Form MantElect23 
   Caption         =   "Histórico Intervención de Columna."
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6765
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   2190
   ScaleWidth      =   6765
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   840
      Width           =   975
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   5160
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   375
      Index           =   1
      Left            =   4170
      TabIndex        =   1
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Index           =   0
      Left            =   1260
      TabIndex        =   0
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Km:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   2400
      TabIndex        =   9
      Top             =   900
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Acceso:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   900
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Progresiva:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   4080
      TabIndex        =   7
      Top             =   900
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Ramal:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   285
      Width           =   855
   End
End
Attribute VB_Name = "MantElect23"
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

Private Sub Combo1_Click()
   sLlenoComboAcceso
End Sub

Private Sub Combo3_Click()
   sLlenoComboKm
End Sub

Private Sub Combo4_Click()
   sLlenoComboProgresiva
End Sub

Private Sub Command1_Click(Index As Integer)
  ' Dim mCodActivo As String
  ' Dim mDescActivo As String
   
   If Index = 0 Then
      If Combo1.Text <> "" And Combo3.Text <> "" And Combo4.Text <> "" And Combo5.Text <> "" Then
      
         sMsgEspere Me, "Procesando datos...", True
         mFechaEjec = Now()

         Set XLS = CreateObject("Excel.Application")

         sPlanilla

         sMsgEspere Me, "", False
         XLS.Application.Visible = True
      Else
         MsgBox "Faltan ingresar datos !!!!", vbCritical, sMessage
      End If
   Else
      Unload Me
   End If
End Sub

Private Sub sPlanilla()
   mi = 9
   sCabecera
   
   Set mRec = mObj.getRPT_Hist_Intervencion_Columna(Trim(Right(Combo5, 20)))
   
  Do While Not mRec.EOF
      With XLS
         .Cells(mi, 2).Formula = NVL(mRec!Parte, "")
         .Cells(mi, 3).Formula = NVL(mRec!FechaSolic, "")
         .Cells(mi, 4).Formula = NVL(mRec!IdOT, "")
         .Cells(mi, 5).Formula = NVL(mRec!CodEdificio, "")
         .Cells(mi, 6).Formula = NVL(mRec!descripcion, "")
         .Cells(mi, 7).Formula = NVL(mRec!FechaFin, "")

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
      .Worksheets(1).Name = "Columna"
     
      .Columns("A:A").ColumnWidth = 2
      .Columns("B:B").ColumnWidth = 12 ' Parte
      .Columns("C:C").ColumnWidth = 17 ' Fecha Carga
      .Columns("D:D").ColumnWidth = 13 ' OT
      .Columns("E:E").ColumnWidth = 22 ' Columna
      .Columns("F:F").ColumnWidth = 40 ' Problema
      .Columns("G:G").ColumnWidth = 17 ' Fecha Cierre
   
      .Columns("B:B").HorizontalAlignment = xlHAlignCenter
      .Columns("C:C").HorizontalAlignment = xlHAlignCenter
      .Columns("D:D").HorizontalAlignment = xlHAlignCenter
      .Columns("E:E").HorizontalAlignment = xlHAlignCenter
      .Columns("F:F").HorizontalAlignment = xlHAlignCenter
      .Columns("G:G").HorizontalAlignment = xlHAlignCenter
            
      'Formateo OrdTrabajo: 000x
'      .Range("B9:B65536").Select
'      .Selection.NumberFormat = "000000"

      'Formateo NroVale: 000x
'      .Range("M9:M65536").Select
'      .Selection.NumberFormat = "000000000"

      'Formateo: Encabezd con Negrita.
      .Range("B8:G8").Select
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
      .Cells(3, 1).Formula = "REPORTE: Histórico de Intervención de Columna"

      .Range("A1:A1").Select
      .Selection.Font.ColorIndex = 5
      .Selection.Font.Size = 14
      .Selection.Font.Bold = True

      .Range("A3:A3").Select
      .Selection.Font.Size = 12
      .Selection.Font.Bold = True

      .Range("B8:B8").Select

      .Cells(5, 1).Formula = "Columna: " & mObj.sCampoDescrip("COM_Activos", "Codigo = " & Trim(Right(Combo5.Text, 20)), 1)
      .Cells(6, 1).Formula = "Fecha ejecución del Reporte: " & mFechaEjec

      .Cells(8, 2).Formula = "Parte"
      .Cells(8, 3).Formula = "Fecha Carga"
      .Cells(8, 4).Formula = "Ord. Trabajo"
      .Cells(8, 5).Formula = "Columna"
      .Cells(8, 6).Formula = "Problema"
      .Cells(8, 7).Formula = "Fecha Cierre"


   End With
End Sub

Private Sub Form_Load()
   Me.Width = 6885
   Me.Height = 2700
   sAlinearForm Me
   
   Combo3.Enabled = False
   Combo4.Enabled = False
   Combo5.Enabled = False
   
   sLlenoRamal
   'cboRamalListIndex = Combo1.ListIndex
   
End Sub

Private Sub sLlenoRamal()
   Dim mRec1 As New ADODB.Recordset
   
   Combo1.Clear
   Set mRec1 = mObj.oEjecutarSelect("SELECT Codigo, Descripcion From COM_Ramales ORDER BY Descripcion; ")

   Do While Not mRec1.EOF
      Combo1.AddItem mRec1!descripcion & Space(100) & mRec1!Codigo
      mRec1.MoveNext
   Loop
   
   mRec1.Close
   Set mRec1 = Nothing
End Sub

Private Sub sLlenoComboAcceso()
   Dim mCodRamal As String
   Dim mCodTipoActivo As String
   'Dim mObj As New clInven
   Dim mRec1 As New ADODB.Recordset
   
   
   mCodRamal = Right(Combo1.Text, 1)
   mCodTipoActivo = "01"
   Combo3.Clear
   Combo4.Clear
   Combo5.Clear

   Set mRec1 = mObj.oEjecutarSelect("SELECT DISTINCT Acceso FROM COM_Activos " & _
                                    " WHERE CodRamal = '" & mCodRamal & "'" & _
                                    " AND CodTipoActivo = '" & mCodTipoActivo & "' ORDER BY Acceso")
    
  If Not mRec1.EOF Then
     Combo3.Enabled = True
      
      Do While Not mRec1.EOF
         Combo3.AddItem "" & mRec1!Acceso
         mRec1.MoveNext
      Loop
  End If
   Combo4.Enabled = False
   Combo5.Enabled = False
   
   mRec1.Close
   Set mRec1 = Nothing
End Sub


Private Sub sLlenoComboKm()
   Dim mCodRamal As String
   Dim mCodTipoAcceso As String
   Dim mCodAcceso As String
   
   Dim mRec1 As New ADODB.Recordset
   
   mCodRamal = Right(Combo1.Text, 1)
   mCodTipoAcceso = "01"
   mCodAcceso = Trim(Combo3.Text)
   
   
   Combo4.Clear
   Combo4.Enabled = True
   Combo5.Clear
   Combo5.Enabled = False

   Set mRec1 = mObj.oEjecutarSelect("SELECT DISTINCT Km FROM COM_Activos " & _
                                    " WHERE CodRamal = '" & mCodRamal & "'" & _
                                    " AND CodTipoActivo = '" & mCodTipoAcceso & "'" & _
                                    " AND Acceso = '" & mCodAcceso & "'" & _
                                    " ORDER BY Km")
   
   Do While Not mRec1.EOF
     Combo4.AddItem "" & mRec1!km
     mRec1.MoveNext
   Loop
   mRec1.Close
   Set mRec1 = Nothing
End Sub


Private Sub sLlenoComboProgresiva()
   Dim mCodRamal As String
   Dim mCodTipoAcceso As String
   Dim mCodAcceso As String
   Dim mKm As String
   
   Dim mRec1 As New ADODB.Recordset
   
   mCodRamal = Right(Combo1.Text, 1)
   mCodTipoAcceso = "01"
   mCodAcceso = Trim(Combo3.Text)
   mKm = Trim(Combo4.Text)
   
   Combo5.Clear
   Combo5.Enabled = True
   
   Set mRec1 = mObj.oEjecutarSelect("SELECT DISTINCT Progresiva, Codigo FROM COM_Activos " & _
                                    " WHERE CodRamal = '" & mCodRamal & "'" & _
                                    " AND CodTipoActivo = '" & mCodTipoAcceso & "'" & _
                                    " AND Acceso = '" & mCodAcceso & "'" & _
                                    " AND Km = '" & mKm & "'" & _
                                    " ORDER BY Progresiva")
   
   Do While Not mRec1.EOF
     Combo5.AddItem "" & mRec1!Progresiva & Space(100) & mRec1!Codigo
     mRec1.MoveNext
   Loop
   mRec1.Close
   'Set mObj = Nothing
   Set mRec1 = Nothing
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set mObj = Nothing
   Set mRec = Nothing
   Set XLS = Nothing
   ShowMenu 47, True, False
End Sub
