VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Viol16_frm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4725
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   10545
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   10545
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00F0F0F0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004A4A4A&
      Height          =   330
      Index           =   2
      Left            =   1350
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2850
      Width           =   2115
   End
   Begin MSComctlLib.ProgressBar PBar1 
      Height          =   165
      Left            =   4575
      TabIndex        =   15
      Top             =   3825
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00DED1BE&
      Caption         =   "buscar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   2925
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1275
      Width           =   990
   End
   Begin MSFlexGridLib.MSFlexGrid Flex1 
      Height          =   2490
      Left            =   4575
      TabIndex        =   8
      Top             =   1275
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   4392
      _Version        =   327680
      Cols            =   5
      FixedCols       =   0
      BackColor       =   16777215
      ForeColor       =   4868682
      BackColorFixed  =   13619151
      ForeColorFixed  =   4868682
      BackColorBkg    =   15790320
      GridColorFixed  =   12632256
      Appearance      =   0
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00F0F0F0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004A4A4A&
      Height          =   330
      Index           =   1
      Left            =   1350
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2325
      Width           =   2115
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00F0F0F0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004A4A4A&
      Height          =   330
      Index           =   0
      Left            =   1350
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1800
      Width           =   2115
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004A4A4A&
      Height          =   315
      Index           =   0
      Left            =   1350
      MaxLength       =   6
      TabIndex        =   1
      Top             =   1275
      Width           =   1440
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00CFCFCF&
      Caption         =   "cancelar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   1
      Left            =   4575
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00CFCFCF&
      Caption         =   "actualizar registros"
      Enabled         =   0   'False
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
      Index           =   0
      Left            =   1350
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4200
      Width           =   2190
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0056534E&
      Height          =   210
      Index           =   6
      Left            =   9675
      TabIndex        =   17
      Top             =   1050
      Width           =   510
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "registros:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0056534E&
      Height          =   195
      Index           =   5
      Left            =   8775
      TabIndex        =   16
      Top             =   1050
      Width           =   825
   End
   Begin VB.Image Image2 
      Height          =   315
      Index           =   2
      Left            =   4050
      MousePointer    =   99  'Custom
      Picture         =   "Viol16_frm.frx":0000
      Stretch         =   -1  'True
      Top             =   2850
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image Image2 
      Height          =   315
      Index           =   1
      Left            =   4050
      MousePointer    =   99  'Custom
      Picture         =   "Viol16_frm.frx":03AE
      Stretch         =   -1  'True
      Top             =   2325
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image Image2 
      Height          =   315
      Index           =   0
      Left            =   4050
      MousePointer    =   99  'Custom
      Picture         =   "Viol16_frm.frx":075C
      Stretch         =   -1  'True
      Top             =   1800
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image Image1 
      Height          =   315
      Index           =   2
      Left            =   3600
      MousePointer    =   99  'Custom
      Picture         =   "Viol16_frm.frx":0B0A
      Stretch         =   -1  'True
      ToolTipText     =   "abm de colores"
      Top             =   2850
      Width           =   315
   End
   Begin VB.Image Image1 
      Height          =   315
      Index           =   1
      Left            =   3600
      MousePointer    =   99  'Custom
      Picture         =   "Viol16_frm.frx":10CA
      Stretch         =   -1  'True
      ToolTipText     =   "abm de modelos"
      Top             =   2325
      Width           =   315
   End
   Begin VB.Image Image1 
      Height          =   315
      Index           =   0
      Left            =   3600
      MousePointer    =   99  'Custom
      Picture         =   "Viol16_frm.frx":168A
      Stretch         =   -1  'True
      ToolTipText     =   "abm de marcas"
      Top             =   1800
      Width           =   315
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "datos existentes"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0056534E&
      Height          =   210
      Index           =   4
      Left            =   4575
      TabIndex        =   14
      Top             =   1050
      Width           =   1620
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "marca"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   3
      Left            =   450
      TabIndex        =   13
      Top             =   1875
      Width           =   765
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "color"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   2
      Left            =   450
      TabIndex        =   12
      Top             =   2925
      Width           =   765
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "modelo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   1
      Left            =   450
      TabIndex        =   11
      Top             =   2400
      Width           =   765
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "patente"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   0
      Left            =   450
      TabIndex        =   10
      Top             =   1350
      Width           =   765
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   10125
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   225
      Width           =   285
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00F0F0F0&
      FillStyle       =   0  'Solid
      Height          =   690
      Index           =   1
      Left            =   0
      Top             =   4050
      Width           =   10590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Modificar datos de vehículos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   270
      Left            =   225
      TabIndex        =   0
      Top             =   300
      Width           =   3690
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H004A4A4A&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00F0F0F0&
      FillStyle       =   0  'Solid
      Height          =   840
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   10590
   End
End
Attribute VB_Name = "Viol16_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
   sAlinearForm Me
   sInitForm
End Sub

Private Sub Combo1_Click(Index As Integer)
   If Index = 0 And Combo1(0).ListIndex > -1 Then
      sLlenoCombo "modelos", " and codmarca='" & Right(Combo1(0).Text, 2) & "' order by 3", Combo1(1), 2, 1
   End If
End Sub

Private Sub Command1_Click(Index As Integer)
   Select Case Index
      Case 0 'grabar- modificar
         If fValid() Then
            sGrabar
         Else
            If MsgBox("falta completar datos del vehículo, desea grabar?", vbYesNo, sMessage) = vbYes Then
               sGrabar
            End If
         End If
      Case 1 'cancelar
         sHabilitar False
         Combo1(0).ListIndex = -1
         Combo1(1).ListIndex = -1
         Combo1(2).ListIndex = -1
         
      Case 2 'buscar
         If fValPatente(Text1(0).Text) Then
            sMsgEspere Me, "buscando datos...", True
            sLlenaFlex Text1(0).Text
            sMsgEspere Me, "", False
         End If
   End Select
End Sub

Private Sub sInitForm()
   With Flex1
      .TextMatrix(0, 0) = "nro"
      .TextMatrix(0, 1) = "fecha"
      .TextMatrix(0, 2) = "marca"
      .TextMatrix(0, 3) = "modelo"
      .TextMatrix(0, 4) = "color"
      .ColWidth(0) = 400
      .ColWidth(1) = 1200
      .ColWidth(2) = 1300
      .ColWidth(3) = 1300
      .ColWidth(4) = 2000
   End With
   sLlenoCombo "marcas", "order by 2", Combo1(0), 1, 0
   sLlenoCombo "colores", "order by 2", Combo1(2), 1, 0
End Sub

Private Sub Image1_Click(Index As Integer)
Dim mVec(3) As Integer
   mVec(0) = 2
   mVec(1) = 1
   mVec(2) = 0
   Viol3_frm.sInitForm mVec(Index)
   Viol3_frm.pViol16_View = True
   Viol3_frm.Show
   Image1(Index).Visible = False
   Image2(Index).Visible = True
   Me.Enabled = False
End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Image1(Index).BorderStyle = 1
   Image1(Index).Picture = LoadPicture(App.Path & "\erp\imagenes\abm_g.gif")
End Sub

Private Sub Image1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Image1(Index).BorderStyle = 0
   Image1(Index).Picture = LoadPicture(App.Path & "\erp\imagenes\abm.gif")
End Sub

Private Sub Image2_Click(Index As Integer)
   Select Case Index
      Case 0
         sLlenoCombo "marcas", "order by 2", Combo1(0), 1, 0
      Case 1
         sLlenoCombo "modelos", " and codmarca='" & Right(Combo1(0).Text, 2) & "' order by 3", Combo1(1), 2, 1
      Case 2
         sLlenoCombo "colores", "order by 2", Combo1(2), 1, 0
   End Select
   Image1(Index).Visible = True
   Image2(Index).Visible = False
End Sub

Private Sub Label2_Click()
   Unload Me
   ShowMenu 5, True, False
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = fAlfaNumKeyPress(KeyAscii)
   KeyAscii = fUcaseKeyPress(KeyAscii)
End Sub

Private Sub sGrabar()
Dim mObj As New clViolaciones
   mObj.xUpDatosVehic Text1(0).Text, Trim(Right(Combo1(0).Text, 3)), Trim(Right(Combo1(1).Text, 3)), Trim(Right(Combo1(2).Text, 3))
   Set mObj = Nothing
End Sub

Private Function fValid() As Boolean
   fValid = Combo1(0).ListIndex > -1
   fValid = fValid And Combo1(1).ListIndex > -1
   fValid = fValid And Combo1(2).ListIndex > -1
End Function

Private Function fValPatente(ByVal pPatente As String)
   pPatente = Trim(pPatente)
   fValPatente = False
   If Len(pPatente) >= 6 Then
      fValPatente = (Asc(Left(pPatente, 1)) >= 65 And Asc(Left(pPatente, 1)) <= 90)
      fValPatente = fValPatente And (Asc(Mid(pPatente, 2, 1)) >= 65 And Asc(Mid(pPatente, 2, 1)) <= 90)
      fValPatente = fValPatente And (Asc(Mid(pPatente, 3, 1)) >= 65 And Asc(Mid(pPatente, 3, 1)) <= 90)
      fValPatente = fValPatente And (Asc(Mid(pPatente, 4, 1)) >= 48 And Asc(Mid(pPatente, 4, 1)) <= 57)
      fValPatente = fValPatente And (Asc(Mid(pPatente, 5, 1)) >= 48 And Asc(Mid(pPatente, 5, 1)) <= 57)
      fValPatente = fValPatente And (Asc(Mid(pPatente, 6, 1)) >= 48 And Asc(Mid(pPatente, 6, 1)) <= 57)
   End If
   If Not fValPatente Then MsgBox "patente incorrecta.", vbExclamation, sMessage
End Function

Private Sub sLlenaFlex(ByVal pPatente As String)
Dim mObj As New clViolaciones
Dim mRec As New ADODB.Recordset
Dim mMarca As String
Dim mModelo As String
Dim mColor As String
Dim mI As Integer
   
   sBorraFlexDatos Me.Flex1
   Set mRec = mObj.oViolFechasPatEst(pPatente, "01/01/2002", Date, "")
   If Not mRec.EOF Then
      sHabilitar True
      PBar1.mIn = 0
      PBar1.Max = mObj.iCountViolFechaPatente2(pPatente, "01/01/2002")
      PBar1.Visible = True
      mI = 1
      Do While Not mRec.EOF
         mMarca = Trim(NVL(mRec!CodMarca, ""))
         If Len(mMarca) = 2 Then
            mMarca = mObj.sCampoDescrip("marcas", "codigo='" & NVL(mRec!CodMarca, "") & "'", 1)
         End If
         mModelo = Trim(NVL(mRec!modelo, ""))
         If Len(mModelo) = 2 Then
            mModelo = mObj.sCampoDescrip("modelos", "codigo='" & NVL(mRec!modelo, "") & "' and codmarca='" & NVL(mRec!CodMarca, "") & "'", 2)
         End If
         mColor = Trim(NVL(mRec!Color, ""))
         If Len(mColor) = 2 Then
            mColor = mObj.sCampoDescrip("colores", "codigo='" & Trim(NVL(mRec!Color, "")) & "'", 1)
         End If
         Flex1.AddItem mI & vbTab & mRec!Fecha & vbTab & mMarca & vbTab & mModelo & vbTab & mColor
         mI = mI + 1
         mRec.MoveNext
         PBar1.Value = mI - 1
      Loop
      PBar1.Visible = False
      Label3(6).Caption = (mI - 1)
   Else
      MsgBox "No existen datos para esta patente.", vbExclamation, sMessage
   End If
   mRec.Close
   If Flex1.Rows > 2 Then
      Flex1.RemoveItem 1
      sSetFlex2Colors Me.Flex1, &HFFFFFF, &HE6E6E6
   End If
   Set mObj = Nothing
   Set mRec = Nothing
End Sub

Private Sub sHabilitar(ByVal pFlag As Boolean)
   Label3(6).Caption = 0
   Command1(0).Enabled = pFlag
   Command1(1).Enabled = pFlag
   Command1(2).Enabled = Not pFlag
   Combo1(0).Enabled = pFlag
   Combo1(1).Enabled = pFlag
   Combo1(2).Enabled = pFlag
   Text1(0).Enabled = Not pFlag
   If Not pFlag Then
      sBorraFlexDatos Me.Flex1
   End If
End Sub

Private Sub sLlenoCombo(ByVal pTabla As String, ByVal pOrder As String, ByRef pObj As Object, ByVal pCD As Integer, ByVal pCC)
   Dim mObj As New clViolaciones
   Dim mRec As New ADODB.Recordset
   pObj.Clear
   Set mRec = mObj.oTablaNotNull(pTabla, pOrder)
   Do While Not mRec.EOF
      pObj.AddItem mRec.Fields(pCD) & Space(80) & mRec.Fields(pCC)
      mRec.MoveNext
   Loop
   mRec.Close
   Set mObj = Nothing
   Set mRec = Nothing
End Sub
