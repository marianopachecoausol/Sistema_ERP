VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Viol1_frm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Módulo de Consulta de Violaciones"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10920
   ControlBox      =   0   'False
   Icon            =   "Viol1_frm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   10920
   Begin VB.Frame Frame1 
      Height          =   6975
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   10815
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   2
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2160
         Width           =   1935
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   360
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3975
         Left            =   75
         TabIndex        =   11
         Top             =   2880
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   7011
         _Version        =   327680
         Cols            =   10
         FixedCols       =   0
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Volver"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   9120
         TabIndex        =   10
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Consultar"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   9120
         TabIndex        =   9
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   4920
         MaxLength       =   15
         TabIndex        =   1
         Top             =   1440
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2160
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   3
         Left            =   3960
         MaxLength       =   5
         TabIndex        =   8
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   2
         Left            =   2760
         MaxLength       =   10
         TabIndex        =   7
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   1800
         MaxLength       =   5
         TabIndex        =   6
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   600
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   1500
         TabIndex        =   12
         Top             =   120
         Width           =   7335
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Consulta de Violaciones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   360
            Left            =   2040
            TabIndex        =   13
            Top             =   240
            Width           =   3045
         End
      End
      Begin VB.Image Image1 
         Height          =   495
         Left            =   7800
         MouseIcon       =   "Viol1_frm.frx":030A
         MousePointer    =   99  'Custom
         Picture         =   "Viol1_frm.frx":045C
         Stretch         =   -1  'True
         ToolTipText     =   "Imprimir resultado"
         Top             =   1200
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   6960
         TabIndex        =   19
         Top             =   2520
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Patente"
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
         Index           =   4
         Left            =   4920
         TabIndex        =   18
         Top             =   1200
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Estación Final"
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
         Index           =   3
         Left            =   2760
         TabIndex        =   17
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Estación Inicial"
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
         Index           =   2
         Left            =   600
         TabIndex        =   16
         Top             =   1920
         Width           =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha-Hora Final"
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
         Left            =   2760
         TabIndex        =   15
         Top             =   1200
         Width           =   1470
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha-Hora Inicial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   14
         Top             =   1200
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Viol1_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mObj As New clViolaciones
Dim mObjPea As New clPeaje
Dim mData As Database
Dim mRec As New ADODB.Recordset
Dim mI As Integer

Private Sub Form_Load()
Dim mObjPea As New clPeaje
   Me.Width = 11010
   Me.Height = 7410
   sAlinearForm Me
   Set mData = OpenDatabase(App.Path & "\Violaciones\Auxiliar.mdb")
   Set mRec = mObjPea.oEstaciones("")
   Do While Not mRec.EOF
      Combo1(0).AddItem mRec!CODIGO_ESTACION & "-" & mRec!Descripcion_Estacion
      Combo1(1).AddItem mRec!CODIGO_ESTACION & "-" & mRec!Descripcion_Estacion
      mRec.MoveNext
   Loop
   mRec.Close
   sCabeceraMS
   sSetFlexRowColor Viol1_frm.MSFlexGrid1, 0, &HE0E0E0
   Text1(0).Text = "01/01/2002"
   Text1(1).Text = "00:00"
   Text1(2).Text = Format(Now, "dd/mm/yyyy")
   Text1(3).Text = Format(Now, "hh:mm")
   Combo1(2).AddItem ""
   Combo1(2).AddItem "Violaciones"
   Combo1(2).AddItem "Rec. Deudas"
   Set mObjPea = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   mData.Close
   Set mData = Nothing
   Set mRec = Nothing
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 27 Then
      Combo1(Index).ListIndex = -1
   End If
End Sub

Private Sub Command1_Click(Index As Integer)
Dim mJ As Integer
Dim mIntervEst As String
Dim mCont As Integer
Dim mColor
Dim mMarca As String
Dim mModelo As String
Dim mDescrColor As String
Dim mEstacion As String
Dim mVia As String
   If Index = 0 Then
      If Command1(0).Caption = "Otra Consulta" Then
         Label3.Visible = False
         Image1.Visible = False
         Text1(4).Text = ""
         Combo1(0).ListIndex = -1
         Combo1(1).ListIndex = -1
         Combo1(2).ListIndex = -1
         Command1(0).Caption = "Consultar"
         sBorraFlexDatos Viol1_frm.MSFlexGrid1
         Text1(4).SetFocus
      Else
         mIntervEst = ""
         If Combo1(0).ListIndex <> -1 And Combo1(1).ListIndex <> -1 Then
            mI = Left(Combo1(0).Text, 2)
            mJ = Left(Combo1(1).Text, 2)
            If (mJ - mI) < 0 Then
               mJ = Left(Combo1(0).Text, 2)
               mI = Left(Combo1(1).Text, 2)
            End If
            For mCont = mI To mJ
               mIntervEst = mIntervEst & "'" & Format(mCont, "0#") & "',"
            Next
            mIntervEst = mId(mIntervEst, 1, Len(mIntervEst) - 1)
         End If
         If fValida Then
            sMsgEspere Me, "Buscando información... aguarde un momento.", True
            Set mRec = mObj.oViolFechasPatEst(Trim(Text1(4).Text), Text1(0).Text & " " & Text1(1).Text & ":00", Text1(2).Text & " " & Text1(3).Text & ":00", mIntervEst, IIf(Left(Combo1(2).Text, 1) = "V", "V", "D"))
            Do While Not mRec.EOF
               mMarca = mObj.sCampoDescrip("marcas", "codigo='" & NVL(mRec!CodMarca, "") & "'", 1)
               mModelo = mObj.sCampoDescrip("modelos", "codmarca='" & NVL(mRec!CodMarca, "") & "' and codigo='" & NVL(mRec!modelo, "") & "'", 2)
               mDescrColor = mObj.sCampoDescrip("colores", "codigo='" & NVL(mRec!Color, "") & "'", 1)
               mEstacion = mRec!Estacion & "-" & mObjPea.sCampoDescrip("PEA23T00", "CODIGO_ESTACION='" & NVL(mRec!Estacion, "") & "'", 1)
               MSFlexGrid1.AddItem mRec!Fecha & vbTab & Format(mRec!Hora, "hh:mm") & vbTab & mEstacion & vbTab & mRec!Via & vbTab & mRec!patente & vbTab & mMarca & vbTab & mModelo & vbTab & mDescrColor & vbTab & NVL(mRec!OBS, "") & vbTab & NVL(mRec!pago, "")
               mRec.MoveNext
            Loop
            mRec.Close
            sSetFlexColOrder Viol1_frm.MSFlexGrid1, 1
            sSetFlex2Colors Viol1_frm.MSFlexGrid1, &HC1DBD8, &HDCEBE9
            Label3.Caption = "Total = " & MSFlexGrid1.Rows - 2
            Label3.Visible = True
            If MSFlexGrid1.Rows > 2 Then
               Image1.Visible = True
               MSFlexGrid1.RemoveItem 1
            End If
            Command1(0).Caption = "Otra Consulta"
            sMsgEspere Me, "", False
         End If
      End If
   Else
      Unload Me
      ShowMenu 5, True, False
   End If
   Set mObj = Nothing
   Set mObjPea = Nothing
End Sub

Private Sub Image1_Click()
Dim mObjAccess As New clAccess
Dim mJ As Integer
Dim mVector(9) As String
Dim mAuxi
   
sMsgEspere Me, "Procesando...", True
mObjAccess.mBorrarAuxi "\Violaciones\Auxiliar", "Auxi"
mData.Execute ("CREATE TABLE Auxi (Orden INTEGER,Fecha DATE,Hora TEXT,Estacion TEXT,Via TEXT,Patente TEXT,Modelo TEXT,Color TEXT, Obs TEXT)")
For mI = 1 To MSFlexGrid1.Rows - 1
   For mJ = 0 To MSFlexGrid1.Cols - 1
      mVector(mJ) = MSFlexGrid1.TextMatrix(mI, mJ)
   Next
   mData.Execute ("INSERT INTO Auxi (Orden,Fecha,Hora,Estacion,Via,Patente,Modelo,Color, Obs) VALUES(" & mI & ",'" & mVector(0) & "','" & mVector(1) & "','" & mVector(2) & "','" & mVector(3) & "','" & mVector(4) & "','" & mVector(6) & "','" & mVector(7) & "','" & mVector(8) & "')")
Next
mAuxi = mData.OpenRecordset("select * from Auxi")
With CrystalReport1
   .WindowTitle = "Reporte Consulta de Violaciones"
   .DataFiles(0) = App.Path & "\Violaciones\Auxiliar.mdb"
   .Formulas(0) = "Listado = 'Período de Consulta Desde " & Text1(0).Text & " Hasta " & Text1(2).Text & ".'"
   .Formulas(1) = "Total = 'Total de Violaciones: " & Trim(mId(Label3.Caption, 8, 13)) & "'"
   .ReportFileName = App.Path & "\Violaciones\rep01.rpt"
   .WindowState = crptMaximized
   .Action = 1
End With
sMsgEspere Me, "", False
Set mObjAccess = Nothing
Set mAuxi = Nothing
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Image1.BorderStyle = 1
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Image1.BorderStyle = 0
End Sub

Private Sub MSFlexGrid1_DblClick()
If MSFlexGrid1.Col = 9 And MSFlexGrid1.Row >= 1 Then
   mObj.UpdPagos IIf(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 9) = "S", "", "S"), MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0), Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)), Text1(4).Text
   MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 9) = IIf(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 9) = "S", "", "S")
   mObj.InsLogPagos Trim(Right(MDI.mUser, 15)), Now, Text1(4).Text, MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 9), MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0), Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)), Trim(Left(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2), 3))
End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 2 'Fecha
         KeyAscii = fDateKeyPress(Text1(Index), KeyAscii)
      Case 1, 3 'Hora
         KeyAscii = fHoraKeyPress(Text1(Index), KeyAscii)
      Case 4 'Patente
         If KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 241 Then
            KeyAscii = KeyAscii - 32
         Else
            If KeyAscii <> 8 And KeyAscii <> 37 And Not (KeyAscii >= 48 And KeyAscii <= 57) And Not (KeyAscii >= 65 And KeyAscii <= 90) Then
               KeyAscii = 0
            End If
         End If
   End Select
End Sub

Private Function sCabeceraMS()
   With MSFlexGrid1
      .ColWidth(0) = 1100
      .ColWidth(1) = 650
      .ColWidth(2) = 1900
      .ColWidth(3) = 600
      .ColWidth(4) = 1400
      .ColWidth(5) = 1800
      .ColWidth(6) = 1700
      .ColWidth(7) = 1200
      .ColWidth(8) = 8500
      .ColWidth(9) = 500
      .Font = "Arial"
      sSetFlexColOrder Viol1_frm.MSFlexGrid1, 1
      .TextMatrix(0, 0) = "Fecha"
      .TextMatrix(0, 1) = "Hora"
      .TextMatrix(0, 2) = "Estación"
      .TextMatrix(0, 3) = "Vía"
      .TextMatrix(0, 4) = "Patente"
      .TextMatrix(0, 5) = "Marca"
      .TextMatrix(0, 6) = "Modelo"
      .TextMatrix(0, 7) = "Color"
      .TextMatrix(0, 8) = "Obs"
      .TextMatrix(0, 9) = "Pago"
      .Row = 0
      For mI = 0 To .Cols - 1
         .CellFontBold = True
      Next
   End With
End Function

Private Function fValida() As Boolean
fValida = Fecha_ok(Text1(0).Text)
fValida = fValida And Hora_ok(Text1(1).Text)
fValida = fValida And Fecha_ok(Text1(2).Text)
fValida = fValida And Hora_ok(Text1(3).Text)
If fValida Then
   fValida = (DateDiff("n", Text1(0).Text & " " & Text1(1).Text, Text1(2).Text & " " & Text1(3).Text) > 0)
   fValida = fValida And (Trim(Text1(4).Text) <> "")
Else
   MsgBox "Verificar las fechas ingresadas.", vbCritical, sMessage
End If
If fValida Then
   fValida = fValida And Combo1(2).Text <> ""
End If
If Not fValida Then
   MsgBox "Verificar los datos ingresados, o falta completar.", vbCritical, sMessage
End If
End Function
