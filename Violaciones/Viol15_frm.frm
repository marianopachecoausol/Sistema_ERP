VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Viol15_frm 
   BackColor       =   &H0081898F&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stand By Vehículos"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10140
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   10140
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   5640
      TabIndex        =   11
      Top             =   180
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H8000000D&
      Height          =   705
      Index           =   1
      Left            =   1140
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   1260
      Width           =   6255
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   0
      Left            =   1140
      TabIndex        =   7
      Top             =   900
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Salir"
      Height          =   315
      Left            =   9240
      TabIndex        =   3
      Top             =   180
      Width           =   795
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1140
      MaxLength       =   15
      TabIndex        =   2
      Top             =   240
      Width           =   1755
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1140
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   180
      Visible         =   0   'False
      Width           =   1755
   End
   Begin MSFlexGridLib.MSFlexGrid mFlex1 
      Height          =   3800
      Left            =   15
      TabIndex        =   0
      Top             =   2040
      Width           =   10155
      _ExtentX        =   17912
      _ExtentY        =   6694
      _Version        =   327680
      Cols            =   5
      FixedCols       =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   4320
      TabIndex        =   10
      Top             =   240
      Width           =   1275
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   2
      Left            =   9480
      MouseIcon       =   "Viol15_frm.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "Viol15_frm.frx":0152
      Stretch         =   -1  'True
      Tag             =   "0"
      ToolTipText     =   "Stand By"
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FAF4E7&
      BackStyle       =   0  'Transparent
      Caption         =   "Usar tecla % como comodín, Ej. AA% busca todas las patentes que comienzan con AA."
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   660
      Width           =   6195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Obs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   660
      TabIndex        =   6
      Top             =   1260
      Width           =   345
   End
   Begin VB.Image Image2 
      Enabled         =   0   'False
      Height          =   495
      Index           =   1
      Left            =   8400
      MouseIcon       =   "Viol15_frm.frx":0341
      MousePointer    =   99  'Custom
      Picture         =   "Viol15_frm.frx":0493
      Stretch         =   -1  'True
      Tag             =   "0"
      ToolTipText     =   "Stand By"
      Top             =   1200
      Width           =   495
   End
   Begin VB.Image Image2 
      Enabled         =   0   'False
      Height          =   495
      Index           =   0
      Left            =   7680
      MouseIcon       =   "Viol15_frm.frx":0A79
      MousePointer    =   99  'Custom
      Picture         =   "Viol15_frm.frx":0BCB
      Stretch         =   -1  'True
      Tag             =   "1"
      ToolTipText     =   "Activo"
      Top             =   1200
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   3060
      MouseIcon       =   "Viol15_frm.frx":1116
      MousePointer    =   99  'Custom
      Picture         =   "Viol15_frm.frx":1268
      Stretch         =   -1  'True
      Tag             =   "B"
      Top             =   180
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Est.Actual:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Patente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   180
      TabIndex        =   4
      Top             =   240
      Width           =   810
   End
End
Attribute VB_Name = "Viol15_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim XLS As EXCEL.Application
Dim mObj As New clViolaciones
Dim mRec As New ADODB.Recordset
Dim mI As Integer

Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Form_Load()
sAlinearForm Me
sInitForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set mObj = Nothing
Set mRec = Nothing
Set XLS = Nothing
ShowMenu 5, True, False
End Sub

Private Sub Combo1_Click()
   Set mRec = mObj.oPatentesStBy(Combo1.Text, Left(Combo2.Text, 1))
   If Not mRec.EOF Then
      If mRec!estado = "1" Then
         Image2(0).Enabled = True
         Image2(1).Enabled = False
         Text2(0).Text = "Stand By - " & mRec!Fecha
         Text2(0).BackColor = QBColor(12)
         Text2(0).ForeColor = &HFFFFFF
      Else
         Image2(0).Enabled = False
         Image2(1).Enabled = True
         Text2(0).Text = "Activo - " & mRec!Fecha
         Text2(0).BackColor = &HC0FFC0
         Text2(0).ForeColor = &H962D0A
      End If
      Text2(1).Text = NVL(mRec!OBS, "")
   Else
      Image2(0).Enabled = False
      Image2(1).Enabled = True
      Text2(0).Text = "Activo"
      Text2(0).BackColor = &HC0FFC0
      Text2(0).ForeColor = &H962D0A
      Text2(1).Text = ""
   End If
   mRec.Close
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Image1_Click()
If Combo2.Text <> "" Then
   If Image1.Tag = "B" Then
      If Text1.Text <> "" Then
         sMsgEspere Me, "Buscando datos...", True
         Set mRec = mObj.oPatentesReg(Trim(Text1.Text))
         If Not mRec.EOF Then
            Combo1.Clear
            Do While Not mRec.EOF
               Combo1.AddItem mRec.Fields(0)
               mRec.MoveNext
            Loop
            Combo1.Visible = True
            Text1.Visible = False
            Image1.Tag = "R"
            Image1.Picture = LoadPicture(App.Path & "\ERP\imagenes\redo.gif")
            Combo2.Enabled = False
         Else
            MsgBox "No existen datos", vbInformation, sMessage
         End If
         mRec.Close
         sMsgEspere Me, "", False
      End If
   Else
      Image1.Tag = "B"
      Image1.Picture = LoadPicture(App.Path & "\Violaciones\imagenes\buscar.gif")
      Image2(0).Enabled = False
      Image2(1).Enabled = False
      Combo1.Clear
      Combo1.Visible = False
      Text1.Visible = True
      Text1.Text = ""
      Text2(0).Text = ""
      Text2(0).BackColor = &HE0E0E0
      Text2(1).Text = ""
      Combo2.Enabled = True
      Combo2.Text = Combo2.List(0)
   End If
Else
   MsgBox "Debe seleccionar el tipo de incidencia", vbCritical, "Atención"
End If
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.BorderStyle = 1
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.BorderStyle = 0
End Sub

Private Sub Image2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2(Index).BorderStyle = 1
End Sub

Private Sub Image2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2(Index).BorderStyle = 0
End Sub

Private Sub Image2_Click(Index As Integer)
   Dim mFlag As Boolean
   Select Case Index
      Case 0
         If MsgBox("Pasar a estado ACTIVO el vehículo: " & Combo1.Text & "?", vbYesNo, sMessage) = vbYes Then
            mFlag = mObj.xUpdateHisRegPagos(Trim(Combo1.Text), "0", Left(Combo2.Text, 1))
            mFlag = mObj.xUpdateHisRegPagos(Trim(Combo1.Text), "1", Left(Combo2.Text, 1))
            mFlag = mObj.xInsertRegPagos(Trim(Combo1.Text), Trim(Text2(1).Text), "0", Left(Combo2.Text, 1))
            sLlenoFlex True
            Combo1.ListIndex = -1
         End If
      Case 1
         If MsgBox("Pasar a estado STAND BY el vehículo: " & Combo1.Text & "?", vbYesNo, sMessage) = vbYes Then
            mFlag = mObj.xUpdateHisRegPagos(Trim(Combo1.Text), "0", Left(Combo2.Text, 1))
            mFlag = mObj.xUpdateHisRegPagos(Trim(Combo1.Text), "1", Left(Combo2.Text, 1))
            mFlag = mObj.xInsertRegPagos(Trim(Combo1.Text), Trim(Text2(1).Text), "1", Left(Combo2.Text, 1))
            sLlenoFlex True
            Combo1.ListIndex = -1
         End If
     Case 2
         sMsgEspere Me, "Generando informe en Excel...", True
         sImprimirXLS
         sMsgEspere Me, "", False
   End Select
End Sub

Private Sub sInitForm()
With mFlex1
   .ColWidth(0) = 450   'Nro
   .ColWidth(1) = 1400  'Patente
   .ColWidth(2) = 1700  'Fecha
   .ColWidth(3) = 800   'Stand by
   .ColWidth(4) = 6400  'obs
   .TextMatrix(0, 0) = "Nro"
   .TextMatrix(0, 1) = "Patente"
   .TextMatrix(0, 2) = "Fecha"
   .TextMatrix(0, 3) = "StBy"
   .TextMatrix(0, 4) = "Observaciones"
End With
sLlenoFlex False
Combo2.AddItem ""
Combo2.AddItem "V - Violaciones"
Combo2.AddItem "D - Rec. Deuda"
End Sub

Private Sub sLlenoFlex(ByVal pBorrar As Boolean)
   Dim mStBy As String
   Dim mJ As Integer
   If pBorrar Then
      sBorraFlexDatos Viol15_frm.mFlex1
   End If
   Set mRec = mObj.oTabla("regpagos", "where estado in ('0', '1') order by patente, fecha")
   Do While Not mRec.EOF
      mStBy = "NO"
      If mRec!estado = "1" Then
         mStBy = "SI"
      End If
      mFlex1.AddItem mI & vbTab & mRec.Fields(0) & vbTab & mRec.Fields(1) & vbTab & mStBy & vbTab & mRec.Fields(3)
      mRec.MoveNext
   Loop
   mRec.Close
   If mFlex1.Rows > 2 Then
      mFlex1.RemoveItem 1
   End If
   sSetFlexNroFila Viol15_frm.mFlex1, 0
   sSetFlex2Colors Viol15_frm.mFlex1, &HFFFFFF, &HE0E0E0
   mFlex1.ColAlignment(2) = 2
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 37 Then
      KeyAscii = fAlfaNumKeyPress(KeyAscii)
      KeyAscii = fUcaseKeyPress(KeyAscii)
   End If
End Sub

Private Sub sImprimirXLS()
   Dim mJ As Integer
   Set XLS = CreateObject("Excel.Application")
   sCabecera
   With XLS
      .Application.DisplayAlerts = False
      .Sheets(3).Select
      .ActiveWindow.SelectedSheets.Delete
      .Sheets(2).Select
      .ActiveWindow.SelectedSheets.Delete
      .Application.DisplayAlerts = True
      sFormatCells "A1:E1", 15
      For mI = 1 To mFlex1.Rows - 1
         For mJ = 1 To 5
            XLS.Cells(mI + 1, mJ).Formula = mFlex1.TextMatrix(mI, mJ - 1)
         Next
      Next
      sFormatCells "A2:E" & (mI), 2
      .Visible = True
   End With
   Set XLS = Nothing
End Sub
   
Private Sub sCabecera()
   With XLS
      .WorkBooks.Add
      .Worksheets(1).Select
      .Worksheets(1).Name = "Datos"
      .Columns("A:A").ColumnWidth = 5    'nro
      .Columns("B:B").ColumnWidth = 15   'patente
      .Columns("C:C").ColumnWidth = 15   'fecha
      .Columns("D:D").ColumnWidth = 10   'stby
      .Columns("E:E").ColumnWidth = 250  'Obs
      .Range("A1:E1").Font.Bold = True
      .Range("A:E").Font.Name = "Arial"
      .Range("A:E").Font.Size = 10
      .Cells.Select
      .Selection.Interior.ColorIndex = 2
      .Selection.Interior.Pattern = xlSolid
      .Cells(1, 1).Formula = "Nro"
      .Cells(1, 2).Formula = "Patente"
      .Cells(1, 3).Formula = "Fecha"
      .Cells(1, 4).Formula = "StdBy"
      .Cells(1, 5).Formula = "Observaciones"
   End With
End Sub

Private Sub sFormatCells(pRango As String, pColor As Integer)
   With XLS
      .Range(pRango).Select
      .Selection.Interior.ColorIndex = pColor
      .Selection.Interior.Pattern = xlSolid
      .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
      .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
      .Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
      .Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
      .Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
      On Error Resume Next
      .Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
   End With
End Sub
