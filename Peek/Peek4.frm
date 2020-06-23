VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Peek4_frm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estado de conexiones"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton Option1 
      BackColor       =   &H8000000E&
      Caption         =   "Km 29"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   4440
      TabIndex        =   33
      Tag             =   "Km29"
      Top             =   1500
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C1DBD8&
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   3900
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Index           =   1
      Left            =   2340
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   360
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H8000000E&
      Caption         =   "Rta 5"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   4440
      TabIndex        =   7
      Tag             =   "Rta5"
      Top             =   2955
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H8000000E&
      Caption         =   "Km 45"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   4440
      TabIndex        =   6
      Tag             =   "Km47"
      Top             =   2595
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H8000000E&
      Caption         =   "Km36"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   4440
      TabIndex        =   5
      Tag             =   "Km36"
      Top             =   2235
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H8000000E&
      Caption         =   "Km 32"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   4440
      TabIndex        =   4
      Tag             =   "Km32"
      Top             =   1875
      Width           =   915
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H8000000E&
      Caption         =   "Km 23"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   4440
      TabIndex        =   3
      Tag             =   "Km23"
      Top             =   1140
      Width           =   915
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H8000000E&
      Caption         =   "Km 14"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   4440
      TabIndex        =   2
      Tag             =   "Km14"
      Top             =   780
      Width           =   915
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Index           =   0
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   2235
   End
   Begin MSFlexGridLib.MSFlexGrid Flex1 
      Height          =   3555
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   6271
      _Version        =   327680
      Rows            =   7
      Cols            =   7
      FixedCols       =   0
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Index           =   6
      Left            =   5550
      TabIndex        =   36
      Top             =   1500
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Index           =   6
      Left            =   6675
      TabIndex        =   35
      Top             =   1500
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   6
      Left            =   7725
      TabIndex        =   34
      Top             =   1500
      Width           =   600
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   7800
      TabIndex        =   31
      Top             =   480
      Width           =   450
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desconect."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   1
      Left            =   6600
      TabIndex        =   30
      Top             =   480
      Width           =   990
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Conectado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   5460
      TabIndex        =   29
      Top             =   480
      Width           =   930
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   5
      Left            =   7740
      TabIndex        =   28
      Top             =   2880
      Width           =   600
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   4
      Left            =   7740
      TabIndex        =   27
      Top             =   2520
      Width           =   600
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   3
      Left            =   7740
      TabIndex        =   26
      Top             =   2160
      Width           =   600
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   2
      Left            =   7740
      TabIndex        =   25
      Top             =   1800
      Width           =   600
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   1
      Left            =   7740
      TabIndex        =   24
      Top             =   1140
      Width           =   600
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   7740
      TabIndex        =   23
      Top             =   780
      Width           =   600
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Index           =   5
      Left            =   6660
      TabIndex        =   22
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Index           =   4
      Left            =   6660
      TabIndex        =   21
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Index           =   3
      Left            =   6660
      TabIndex        =   20
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Index           =   2
      Left            =   6660
      TabIndex        =   19
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Index           =   1
      Left            =   6660
      TabIndex        =   18
      Top             =   1140
      Width           =   1100
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Index           =   0
      Left            =   6660
      TabIndex        =   17
      Top             =   780
      Width           =   1100
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Index           =   5
      Left            =   5520
      TabIndex        =   16
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Index           =   4
      Left            =   5520
      TabIndex        =   15
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Index           =   3
      Left            =   5520
      TabIndex        =   14
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Index           =   2
      Left            =   5520
      TabIndex        =   13
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Index           =   1
      Left            =   5520
      TabIndex        =   12
      Top             =   1140
      Width           =   1100
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Index           =   0
      Left            =   5520
      TabIndex        =   11
      Top             =   780
      Width           =   1100
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Año"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   1
      Left            =   2340
      TabIndex        =   10
      Top             =   120
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mes"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   120
      Width           =   330
   End
End
Attribute VB_Name = "Peek4_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mObj As New clPeek
Dim mRec As New ADODB.Recordset
Dim mFecha1 As String
Dim mFecha2 As String
Dim mDiaI As Integer

Private Sub Form_Load()
   sInitForm
   sAlinearForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set mObj = Nothing
   Set mRec = Nothing
   ShowMenu 13, True, False
End Sub

Private Sub Combo1_Click(Index As Integer)
Dim mDia As Integer
Dim mMaxDia As Integer
Dim mI As Integer
Dim mRow As Integer
   
   If Combo1(0).ListIndex > -1 And Combo1(1).ListIndex > -1 Then
      sDespintar
      For mRow = 1 To Flex1.Rows - 1
         For mI = 0 To 6
            Flex1.TextMatrix(mRow, mI) = ""
         Next
      Next
      mFecha1 = "01/" & Right(Combo1(0).Text, 2) & "/" & Combo1(1).Text
      mDiaI = WeekDay("01/" & Right(Combo1(0).Text, 2) & "/" & Combo1(1).Text) - 1
      mMaxDia = Day(DateAdd("d", -1, DateAdd("m", 1, "01/" & Right(Combo1(0).Text, 2) & "/" & Combo1(1).Text)))
      mFecha2 = mMaxDia & "/" & Right(Combo1(0).Text, 2) & "/" & Combo1(1).Text
      mRow = 1
      mDia = mDiaI
      For mI = 1 To mMaxDia
         If mDia > 6 Then
            mDia = 0
            mRow = mRow + 1
         End If
         Flex1.TextMatrix(mRow, mDia) = mI
         mDia = mDia + 1
      Next
      For mI = 0 To Option1.UBound
         If Option1(mI).Value = True Then
            Option1_Click mI
            Exit For
         End If
      Next
      sPintar
      sEstados
   End If
   
End Sub

Private Sub Command1_Click()
   Unload Me
End Sub



Private Sub Option1_Click(Index As Integer)
Dim mDia As Integer
Dim mRow As Integer
Dim mCol As Integer

  ' sMsgEspere Me, "Buscando Datos...", True
   sDespintar
   Set mRec = mObj.oFallasMes(Option1(Index).Tag, mFecha1, mFecha2)
   Do While Not mRec.EOF
      mDia = Int(Day(mRec!Fecha))
      mRow = Int((mDia + mDiaI) / 7) + 1
      mCol = Int(((mDia + mDiaI) Mod 7) - 1)
      Flex1.Row = mRow
      Flex1.Col = mCol
      If mRec!estado = "OK" Then
         Flex1.CellBackColor = &H80FF80
      Else
         Flex1.CellBackColor = &HFF&
      End If
      mRec.MoveNext
   Loop
   sPintar
   mRec.Close
End Sub

Private Sub sInitForm()
Dim mI As Integer
Dim mJ As Integer

   Me.Width = 8640
   Me.Height = 4785
   With Flex1
      .Row = 0
      For mI = 0 To 6
         .Col = mI
         .ColWidth(mI) = 600
         .ColAlignment(mI) = 4
         .CellFontBold = True
      Next
      For mI = 0 To Flex1.Rows - 1
         .RowHeight(mI) = 500
      Next
      .TextMatrix(0, 0) = "D"
      .TextMatrix(0, 1) = "L"
      .TextMatrix(0, 2) = "M"
      .TextMatrix(0, 3) = "M"
      .TextMatrix(0, 4) = "J"
      .TextMatrix(0, 5) = "V"
      .TextMatrix(0, 6) = "S"
   End With
   For mI = 1 To 12
      Combo1(0).AddItem MonthName(mI) & Space(40) & Format(mI, "00")
   Next
   For mI = Year(Date) To 2008 Step -1
      Combo1(1).AddItem mI
   Next
   'Combo1(1).AddItem "2009"
   
   Combo1(1).ListIndex = 0
End Sub

Private Sub sDespintar()
Dim mI As Integer
Dim mJ As Integer
   
   For mI = 1 To Flex1.Rows - 1
      For mJ = 0 To Flex1.Cols - 1
        Flex1.Row = mI
        Flex1.Col = mJ
        Flex1.CellBackColor = &HFFFFFF
      Next
   Next
End Sub

Private Sub sPintar()
Dim mI As Integer

   For mI = 1 To Flex1.Rows - 1
      If Trim(Flex1.TextMatrix(mI, 0)) <> "" Then
         Flex1.Row = mI
         Flex1.Col = 0
         Flex1.CellBackColor = &HE0E0E0
      End If
      If (Flex1.TextMatrix(mI, 6)) <> "" Then
         Flex1.Row = mI
         Flex1.Col = 6
         Flex1.CellBackColor = &HE0E0E0
      End If
   Next
End Sub

Private Sub sEstados()
Dim mI As Integer
Dim mOk As Integer
Dim mKO As Integer
   For mI = 0 To Option1.UBound
      mOk = mObj.iCountFallas(mFecha1, mFecha2, Option1(mI).Tag, "OK")
      mKO = mObj.iCountFallas(mFecha1, mFecha2, Option1(mI).Tag, "KO")
      If (mOk + mKO) > 0 Then
         Label2(mI).Caption = mOk & " - " & Format((mOk / (mOk + mKO) * 100), "00.0") & "%"
         Label3(mI).Caption = mKO & " - " & Format((mKO / (mOk + mKO) * 100), "00.0") & "%"
         Label4(mI).Caption = mKO + mOk
      Else
         Label2(mI).Caption = "S/D"
         Label3(mI).Caption = "S/D"
         Label4(mI).Caption = "S/D"
      End If
   Next
End Sub
