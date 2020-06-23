VERSION 5.00
Begin VB.Form MantElect05 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modulo de Relevamientos"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   16830
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   12
      Left            =   14520
      MaxLength       =   2
      TabIndex        =   14
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   2040
      MaxLength       =   5
      TabIndex        =   2
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Index           =   0
      Left            =   6720
      TabIndex        =   17
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Volver"
      Height          =   375
      Index           =   1
      Left            =   9000
      TabIndex        =   18
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   11
      Left            =   15960
      MaxLength       =   5
      TabIndex        =   16
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   10
      Left            =   15240
      MaxLength       =   5
      TabIndex        =   15
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   9120
      MaxLength       =   3
      TabIndex        =   12
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   6120
      MaxLength       =   150
      TabIndex        =   4
      Top             =   840
      Width           =   5775
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   960
      MaxLength       =   10
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   9
      Left            =   9720
      MaxLength       =   90
      TabIndex        =   13
      Top             =   1800
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   13200
      MaxLength       =   10
      TabIndex        =   6
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   14280
      MaxLength       =   5
      TabIndex        =   7
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   16080
      MaxLength       =   5
      TabIndex        =   9
      Top             =   1080
      Width           =   495
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   840
      Width           =   3015
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   1
      Left            =   12120
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   840
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   2
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1800
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   3
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1800
      Width           =   5655
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   15000
      MaxLength       =   10
      TabIndex        =   8
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Interv."
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
      Index           =   14
      Left            =   14500
      TabIndex        =   34
      Top             =   1560
      Width           =   570
   End
   Begin VB.Line Line2 
      X1              =   14400
      X2              =   14400
      Y1              =   1440
      Y2              =   2280
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Horas"
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
      Index           =   13
      Left            =   16005
      TabIndex        =   33
      Top             =   1560
      Width           =   510
   End
   Begin VB.Line Line1 
      Index           =   11
      X1              =   13080
      X2              =   16680
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line1 
      Index           =   10
      X1              =   120
      X2              =   16680
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      Index           =   8
      X1              =   3120
      X2              =   3120
      Y1              =   1440
      Y2              =   2280
   End
   Begin VB.Label Label1 
      Caption         =   "Registro de Relevamientos previos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6360
      TabIndex        =   32
      Top             =   120
      Width           =   4185
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Parte"
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
      Left            =   240
      TabIndex        =   31
      Top             =   600
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha / Hora Relev."
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
      Left            =   915
      TabIndex        =   30
      Top             =   600
      Width           =   1770
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Lugar"
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
      Left            =   3000
      TabIndex        =   29
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Descripcion del Lugar"
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
      Left            =   6120
      TabIndex        =   28
      Top             =   600
      Width           =   1875
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Prioridad"
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
      Left            =   12150
      TabIndex        =   27
      Top             =   600
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha / Hora Inicio"
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
      Index           =   6
      Left            =   13170
      TabIndex        =   26
      Top             =   840
      Width           =   1680
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Segunda Descripcion"
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
      Index           =   10
      Left            =   9840
      TabIndex        =   25
      Top             =   1560
      Width           =   1830
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Unid"
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
      Index           =   11
      Left            =   9120
      TabIndex        =   24
      Top             =   1560
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Cant."
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
      Index           =   12
      Left            =   15270
      TabIndex        =   23
      Top             =   1560
      Width           =   465
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   2760
      X2              =   2760
      Y1              =   480
      Y2              =   1440
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   120
      X2              =   16680
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   120
      X2              =   120
      Y1              =   480
      Y2              =   2280
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   840
      X2              =   840
      Y1              =   480
      Y2              =   1440
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   14880
      X2              =   14880
      Y1              =   720
      Y2              =   1440
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   6000
      X2              =   6000
      Y1              =   480
      Y2              =   1440
   End
   Begin VB.Line Line1 
      Index           =   6
      X1              =   12000
      X2              =   12000
      Y1              =   480
      Y2              =   1440
   End
   Begin VB.Line Line1 
      Index           =   12
      X1              =   120
      X2              =   16680
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      Index           =   13
      X1              =   9000
      X2              =   9000
      Y1              =   1440
      Y2              =   2280
   End
   Begin VB.Line Line1 
      Index           =   9
      X1              =   16680
      X2              =   16680
      Y1              =   480
      Y2              =   2280
   End
   Begin VB.Line Line1 
      Index           =   7
      X1              =   13080
      X2              =   13080
      Y1              =   480
      Y2              =   1440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Rubro"
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
      Index           =   8
      Left            =   315
      TabIndex        =   22
      Top             =   1560
      Width           =   525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Sub Rubro"
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
      Index           =   9
      Left            =   3360
      TabIndex        =   21
      Top             =   1560
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha / Hora Fin"
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
      Index           =   7
      Left            =   15030
      TabIndex        =   20
      Top             =   840
      Width           =   1470
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Asistencia"
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
      Index           =   5
      Left            =   14460
      TabIndex        =   19
      Top             =   525
      Width           =   885
   End
   Begin VB.Line Line1 
      Index           =   14
      X1              =   9600
      X2              =   9600
      Y1              =   1440
      Y2              =   2280
   End
   Begin VB.Line Line1 
      Index           =   15
      X1              =   15120
      X2              =   15120
      Y1              =   1440
      Y2              =   2280
   End
   Begin VB.Line Line1 
      Index           =   16
      X1              =   15840
      X2              =   15840
      Y1              =   1440
      Y2              =   2280
   End
End
Attribute VB_Name = "MantElect05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mObj As New clMantElect
Dim mObjPea As New clPeaje
Dim mRec As New ADODB.Recordset
Dim mObjLuser As New clLogUser

Private Sub Combo1_Click(Index As Integer)
Select Case Index
   Case 2
      Combo1(3).Clear
      Set mRec = mObj.oEjecutarSelect("SELECT * FROM SubRubros WHERE CodRubro = '" & Left(Combo1(2).Text, 8) & "' AND FechaBaja IS NULL ORDER BY Codigo")
      If Not mRec.EOF Then
         Do While Not mRec.EOF
            Combo1(3).AddItem mRec!Codigo & "-" & mRec!descripcion
            mRec.MoveNext
         Loop
      End If
      mRec.Close
   Case 3
      Set mRec = mObj.oEjecutarSelect("SELECT * FROM SubRubros WHERE CodRubro = '" & Left(Combo1(2).Text, 8) & "' AND FechaBaja IS NULL ORDER BY Codigo")
      Text1(8).Text = mRec!Unidad
      mRec.Close
End Select
End Sub

Private Sub Command1_Click(Index As Integer)
Dim mOperador As String
Dim mi As Integer
If Index = 0 Then
   If fValida Then
      If MsgBox("¿Está Seguro de Grabar los datos?", vbYesNo, sMessage) = vbYes Then
         'Grabo en Registros
         'mOperador = Left(mObjLuser.sCampoDescrip("USUARIOS", "CodUsuario = '" & Trim(Right(MDI.mUser, 20)) & "'", 6), InStr(1, mObjLuser.sCampoDescrip("USUARIOS", "CodUsuario = '" & Trim(Right(MDI.mUser, 20)) & "'", 6), "@") - 1)
         mOperador = Trim(Right(MDI.mUser, 20))
                 
         mObj.InsRelev mObj.ObtMaxParte + 1, Text1(1).Text & " " & Text1(2).Text & ":00", Combo1(0).Text, Text1(3).Text, Combo1(1).Text, "MantElect", Text1(4).Text & " " & Text1(5).Text & ":00", Text1(6).Text & " " & Text1(7).Text & ":00", Text1(9).Text, Left(Combo1(2).Text, 8), Left(Combo1(3).Text, 6), Text1(10).Text, Text1(11).Text, "T", "M", mOperador, mOperador, Now, mOperador, Now, "", "", "", Text1(12).Text
         For mi = 0 To Text1.UBound
            Text1(mi).Text = ""
         Next
         For mi = 0 To Combo1.UBound
            Combo1(mi).ListIndex = -1
         Next
         Text1(0).Text = mObj.ObtMaxParte + 1
      End If
   End If
Else
   Unload Me
End If
End Sub

Private Sub Form_Load()
Dim mi As Integer
MantElect05.Top = 100
MantElect05.Left = (MDI.Width - MantElect05.Width) / 2

'Set mRec = mObj.oEjecutarSelect("SELECT * FROM EdifRelev WHERE FechaBaja IS NULL")
Set mRec = mObj.oEjecutarSelect("SELECT * FROM Edificios WHERE FechaBaja IS NULL")
If Not mRec.EOF Then
   Do While Not mRec.EOF
      'Combo1(0).AddItem mRec!descripcion
      Combo1(0).AddItem mRec!Tramo & " - " & mRec!descripcion
      mRec.MoveNext
   Loop
End If
mRec.Close

Combo1(1).AddItem "Alta"
Combo1(1).AddItem "Media"
Combo1(1).AddItem "Baja"

Set mRec = mObj.oEjecutarSelect("SELECT * FROM Rubros WHERE FechaBaja IS NULL")
If Not mRec.EOF Then
   Do While Not mRec.EOF
      Combo1(2).AddItem mRec!Codigo & " - " & mRec!descripcion
      mRec.MoveNext
   Loop
End If
mRec.Close

Text1(0).Text = mObj.ObtMaxParte + 1
Text1(0).Enabled = False
Text1(8).Enabled = False
Text1(11).Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set mObj = Nothing
Set mRec = Nothing
ShowMenu 47, True, False
End Sub

Private Function fValida() As Boolean
Dim mRet As Boolean
Dim mi As Integer

mRet = True
For mi = 0 To Text1.UBound
   If mRet Then
      mRet = (Text1(mi).Text <> "")
   End If
Next
If mRet Then
   For mi = 0 To Combo1.UBound
      If mRet Then
         mRet = (Combo1(mi).Text <> "")
      End If
   Next
End If
If mRet Then
   mRet = Fecha_ok(Text1(1).Text)
End If
If mRet Then
   mRet = Hora_ok(Text1(2).Text)
End If
If mRet Then
   mRet = Fecha_ok(Text1(4).Text)
End If
If mRet Then
   mRet = Hora_ok(Text1(5).Text)
End If
If mRet Then
   mRet = Fecha_ok(Text1(6).Text)
End If
If mRet Then
   mRet = Hora_ok(Text1(7).Text)
End If
If mRet Then
   mRet = DateDiff("s", CDate(Text1(4).Text & " " & Text1(5).Text & ":00"), CDate(Text1(6).Text & " " & Text1(7).Text & ":00")) > 0
End If
If Not mRet Then
   MsgBox "Verifique que los datos estén ingresados correctamente", vbCritical, "Atención"
End If
fValida = mRet
End Function

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
   Case 1, 4, 6
      KeyAscii = fDateKeyPress(Text1(Index), KeyAscii)
   Case 2, 5, 7
      KeyAscii = fHoraKeyPress(Text1(Index), KeyAscii)
   Case 3, 9
      KeyAscii = fAlfaNumKeyPress(KeyAscii)
   Case 10, 11, 12
      KeyAscii = fNumDoubleKeyPress(KeyAscii)
End Select
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim mRet As Boolean
Select Case Index
   Case 4, 5, 6, 7
      mRet = (Text1(4).Text <> "" And Text1(5).Text <> "" And Text1(6).Text <> "" And Text1(7).Text <> "")
      If mRet Then
         If DateDiff("s", CDate(Text1(4).Text & " " & Text1(5).Text & ":00"), CDate(Text1(6).Text & " " & Text1(7).Text & ":00")) >= 0 Then
            Text1(11).Text = Replace(Redondeo(DateDiff("n", CDate(Text1(4).Text & " " & Text1(5).Text & ":00"), CDate(Text1(6).Text & " " & Text1(7).Text & ":00")) / 60, 2), ",", ".")
         Else
            MsgBox "Verifique las fechas de Asistencia", vbCritical, "Atención"
            Text1(Index).Text = ""
            Text1(Index).SetFocus
         End If
      End If
End Select
End Sub
