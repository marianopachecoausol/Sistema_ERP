VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form RNov10_frm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  Datos de pedidos de ambulancias"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10905
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   5775
   ScaleWidth      =   10905
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E9E9E9&
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   2
      Left            =   11055
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   675
      Width           =   1665
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Actualizar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   1
      Left            =   11055
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   225
      Width           =   1665
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E9E9E9&
      Caption         =   "Grabar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   0
      Left            =   9200
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   375
      Width           =   1290
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   3
      Left            =   6975
      MaxLength       =   8
      TabIndex        =   9
      Top             =   675
      Width           =   915
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   5550
      MaxLength       =   10
      TabIndex        =   8
      Top             =   675
      Width           =   1290
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   3450
      MaxLength       =   8
      TabIndex        =   7
      Top             =   675
      Width           =   915
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   2025
      MaxLength       =   10
      TabIndex        =   6
      Top             =   675
      Width           =   1290
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2025
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   225
      Width           =   3915
   End
   Begin MSFlexGridLib.MSFlexGrid Flex1 
      Height          =   4440
      Left            =   75
      TabIndex        =   0
      Top             =   1275
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   7832
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      ForeColor       =   4210752
      BackColorFixed  =   14602686
      ForeColorFixed  =   4210752
      BackColorSel    =   12648384
      ForeColorSel    =   4194304
      BackColorBkg    =   16382457
      GridColor       =   14737632
      GridColorFixed  =   12632256
      GridLinesFixed  =   1
      BorderStyle     =   0
      Appearance      =   0
      MousePointer    =   99
      MouseIcon       =   "RNov10_frm.frx":0000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00DED1BE&
      X1              =   8850
      X2              =   0
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00DED1BE&
      FillColor       =   &H00F9F9F9&
      FillStyle       =   0  'Solid
      Height          =   1065
      Left            =   8775
      Top             =   150
      Width           =   2040
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   240
      Index           =   0
      Left            =   6975
      TabIndex        =   10
      Top             =   270
      Width           =   915
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "arribo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   4
      Left            =   4725
      TabIndex        =   4
      Top             =   765
      Width           =   600
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "asignado"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   3
      Left            =   975
      TabIndex        =   3
      Top             =   765
      Width           =   885
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "móvil"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   2
      Left            =   6150
      TabIndex        =   2
      Top             =   300
      Width           =   540
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "pedido y cód. alfa"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   0
      Left            =   75
      TabIndex        =   1
      Top             =   300
      Width           =   1740
   End
End
Attribute VB_Name = "RNov10_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim mI As Integer

   With Flex1
      .TextMatrix(0, 0) = "NRO"
      .TextMatrix(0, 1) = "FECHA"
      .TextMatrix(0, 2) = "HORA"
      .TextMatrix(0, 3) = "COD.ALFA"
      .TextMatrix(0, 4) = "MOVIL"
      .TextMatrix(0, 5) = "ASIGNADO"
      .TextMatrix(0, 6) = "ARRIBO"
      .ColWidth(0) = 500
      .ColWidth(1) = 1300
      .ColWidth(2) = 900
      .ColWidth(3) = 1300
      .ColWidth(4) = 1100
      .ColWidth(5) = 2100
      .ColWidth(6) = 2100
      .ColWidth(7) = 400
      .ColWidth(8) = 400
      .ColWidth(9) = 0
      .Row = 0
      For mI = 0 To Flex1.Cols - 1
         .Col = mI
         .CellFontBold = True
      Next
   End With
   sInitForm
End Sub

Private Sub Combo1_Click()
   Label2(0).Caption = Trim(Right(Combo1.Text, 5))
End Sub

Private Sub Command1_Click(Index As Integer)
Dim mObj As New clRNov
Dim mI As Integer

   Select Case Index
      Case 0 'grabar
         If fValid() Then
            mObj.xInArriboAMBU Trim(mId(Combo1.Text, 24, 10)), Trim(Left(Combo1.Text, 20)), Label2(0).Caption, Text1(0).Text & " " & Text1(1).Text, Text1(2).Text & " " & Text1(3).Text
            sInitForm
         End If
         
      Case 1 'modificar
         If fValid() Then
            mObj.xUpArriboAMBU Flex1.Tag, Trim(mId(Combo1.Text, 24, 10)), Trim(Left(Combo1.Text, 20)), Label2(0).Caption, Text1(0).Text & " " & Text1(1).Text, Text1(2).Text & " " & Text1(3).Text
            sInitForm
            Command1_Click 2
         End If
         
      Case 2 'cancelar
         sInitForm
         For mI = Me.Height To 6150 Step 15
            Me.Height = mI
         Next
         Command1(0).Left = 9200
         Command1(1).Left = 11000
         Command1(2).Left = 11000
         Flex1.Tag = ""
         Flex1.Enabled = True
   End Select
   Set mObj = Nothing
End Sub

Private Sub Flex1_DblClick()
Dim mObj As New clRNov
Dim mI As Integer

   If Flex1.Row > 0 Then
      If Flex1.Col = 7 Then 'Actualizar
         Flex1.Enabled = False
         Combo1.Clear
         Combo1.AddItem Flex1.TextMatrix(Flex1.Row, 1) & " " & Flex1.TextMatrix(Flex1.Row, 2) & "  -  " & Flex1.TextMatrix(Flex1.Row, 3) & Space(20) & Flex1.TextMatrix(Flex1.Row, 4)
         Combo1.ListIndex = 0
         Text1(0).Text = Left(Flex1.TextMatrix(Flex1.Row, 5), 10)
         Text1(1).Text = Right(Flex1.TextMatrix(Flex1.Row, 5), 8)
         Text1(2).Text = Left(Flex1.TextMatrix(Flex1.Row, 6), 10)
         Text1(3).Text = Right(Flex1.TextMatrix(Flex1.Row, 6), 8)
         Flex1.Tag = Flex1.TextMatrix(Flex1.Row, 9)
         For mI = Me.Height To 1700 Step -15
            Me.Height = mI
         Next
         Command1(0).Left = 11000
         Command1(1).Left = 9000
         Command1(2).Left = 9000
         Me.Refresh
         
      End If
      If Flex1.Col = 8 Then
         If MsgBox("Seguro de borrar el registro con cód. alfanumérico=" & Flex1.TextMatrix(Flex1.Row, 3) & "?", vbOKCancel, sMessage) = vbOK Then
            mObj.xDelArriboAmbu Flex1.TextMatrix(Flex1.Row, 9), Flex1.TextMatrix(Flex1.Row, 1), Flex1.TextMatrix(Flex1.Row, 2)
            Unload Me
            RNov10_frm.Refresh
         End If
      End If
   End If
   Set mObj = Nothing
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 2
         KeyAscii = fDateKeyPress(Text1(Index), KeyAscii)
      Case 1, 3
         KeyAscii = fHoraKeyPress(Text1(Index), KeyAscii)
   End Select
End Sub

'---------------------------------------------------------------------------------------
' FUNCIONES
'---------------------------------------------------------------------------------------

Private Function fValid() As Boolean
Dim mText As String
   
   If Combo1.ListIndex < 0 Then mText = ". seleccionar un cód. alfanumérico" & Chr(13)
   fValid = Fecha_ok(Text1(0).Text)
   fValid = fValid And HoraLong_ok(Text1(1).Text)
   fValid = fValid And Fecha_ok(Text1(2).Text)
   fValid = fValid And HoraLong_ok(Text1(3).Text)
   If DateDiff("s", Text1(0).Text & " " & Text1(1).Text, Text1(2).Text & " " & Text1(3).Text) <= 0 Then
      mText = ". fecha asignado mayor a la de arribo" & Chr(13)
   End If
   If mText <> "" Then
      MsgBox mText, vbCritical, sMessage
      fValid = False
   End If
   
End Function

Private Sub sInitForm()
Dim mObj As New clRNov
Dim mRec As New ADODB.Recordset
Dim mI As Integer

   Set mRec = mObj.oAmbusPend()
   If Not mRec.EOF Then
      Combo1.Clear
      Do While Not mRec.EOF
         Combo1.AddItem mRec!Fecha & "  -  " & mRec!Codigo & Space(20) & mRec!Mov1
         mRec.MoveNext
      Loop
   End If
   mRec.Close
   sBorraFlexDatos Me.Flex1
   mI = 1
   Set mRec = mObj.oTabla("arribos_ambu", "order by fecha desc limit 100")
   If Not mRec.EOF Then
      Do While Not mRec.EOF
         Flex1.AddItem mI & vbTab & Left(mRec!Fecha, 10) & vbTab & Right(mRec!Fecha, 8) & vbTab & mRec!CodAlfa & vbTab & mRec!CodMovil & vbTab & mRec!asignado & vbTab & mRec!arribo & vbTab & "M" & vbTab & "X" & vbTab & mRec!Codigo
         mI = mI + 1
         mRec.MoveNext
      Loop
      Flex1.RemoveItem 1
   End If
   mRec.Close
   With Flex1
      For mI = 1 To .Rows - 1
         .Row = mI
         .Col = 3
         .CellAlignment = 4
         .Col = 7
         .CellAlignment = 4
         .CellFontBold = True
         .CellForeColor = vbBlue 'QBColor(10)
         .Col = 8
         .CellAlignment = 4
         .CellFontBold = True
         .CellForeColor = vbRed ' QBColor(9)
      Next
   End With
   For mI = 0 To Text1.UBound
      Text1(mI).Text = ""
   Next
   Set mObj = Nothing
   Set mRec = Nothing
End Sub
