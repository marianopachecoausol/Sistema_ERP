VERSION 5.00
Begin VB.Form MEdfrm05 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento Edilicio"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   16980
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   13
      Left            =   2280
      MaxLength       =   19
      TabIndex        =   2
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   12
      Left            =   4560
      MaxLength       =   150
      TabIndex        =   18
      Top             =   3840
      Width           =   9135
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   11
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   17
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   10
      Left            =   1200
      MaxLength       =   5
      TabIndex        =   16
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   9
      Left            =   360
      MaxLength       =   5
      TabIndex        =   15
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   6480
      MaxLength       =   5
      TabIndex        =   13
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   5760
      MaxLength       =   5
      TabIndex        =   12
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   7680
      MaxLength       =   90
      TabIndex        =   7
      Top             =   840
      Width           =   6015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Index           =   0
      Left            =   5280
      TabIndex        =   19
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Volver"
      Height          =   375
      Index           =   1
      Left            =   7680
      TabIndex        =   20
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1200
      MaxLength       =   19
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   5040
      MaxLength       =   5
      TabIndex        =   11
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   3000
      MaxLength       =   10
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   4080
      MaxLength       =   5
      TabIndex        =   4
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   4680
      MaxLength       =   5
      TabIndex        =   5
      Top             =   840
      Width           =   495
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   840
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   1
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1800
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   2
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1800
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   3
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1800
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Columns         =   8
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "MEdfrm05.frx":0000
      Left            =   360
      List            =   "MEdfrm05.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   14
      Top             =   2880
      Width           =   16335
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   16800
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line1 
      Index           =   19
      X1              =   4440
      X2              =   4440
      Y1              =   3480
      Y2              =   4320
   End
   Begin VB.Line Line1 
      Index           =   16
      X1              =   240
      X2              =   16800
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Observaciones"
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
      Index           =   16
      Left            =   4680
      TabIndex        =   38
      Top             =   3600
      Width           =   1275
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Materiales"
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
      Index           =   15
      Left            =   2040
      TabIndex        =   37
      Top             =   3600
      Width           =   885
   End
   Begin VB.Line Line1 
      Index           =   18
      X1              =   1920
      X2              =   1920
      Y1              =   3480
      Y2              =   4320
   End
   Begin VB.Line Line1 
      Index           =   17
      X1              =   1080
      X2              =   1080
      Y1              =   3480
      Y2              =   4320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "$/Tarea"
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
      Left            =   1155
      TabIndex        =   36
      Top             =   3600
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Hs/Per"
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
      Left            =   360
      TabIndex        =   35
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Mano de Obra"
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
      Left            =   360
      TabIndex        =   34
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Line Line1 
      Index           =   14
      X1              =   6360
      X2              =   6360
      Y1              =   1680
      Y2              =   2400
   End
   Begin VB.Line Line1 
      Index           =   11
      X1              =   4920
      X2              =   7080
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line1 
      Index           =   10
      X1              =   240
      X2              =   16800
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line1 
      Index           =   8
      X1              =   2640
      X2              =   2640
      Y1              =   1320
      Y2              =   2400
   End
   Begin VB.Label Label1 
      Caption         =   "Registro de Partes Adicionales"
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
      Left            =   5160
      TabIndex        =   33
      Top             =   75
      Width           =   3720
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
      Left            =   440
      TabIndex        =   32
      Top             =   600
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha / Hora Solic."
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
      Left            =   1160
      TabIndex        =   31
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Edificio"
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
      Left            =   5520
      TabIndex        =   30
      Top             =   600
      Width           =   645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Descripcion de la Solicitud"
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
      Left            =   7800
      TabIndex        =   29
      Top             =   600
      Width           =   2295
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
      Index           =   5
      Left            =   390
      TabIndex        =   28
      Top             =   1560
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha / Hora Asistencia"
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
      Left            =   3050
      TabIndex        =   27
      Top             =   600
      Width           =   2085
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tiempos"
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
      Left            =   5640
      TabIndex        =   26
      Top             =   1440
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Estim."
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
      Left            =   5040
      TabIndex        =   25
      Top             =   1800
      Width           =   525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Real"
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
      Left            =   5820
      TabIndex        =   24
      Top             =   1800
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Admis."
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
      Left            =   6465
      TabIndex        =   23
      Top             =   1800
      Width           =   570
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   2880
      X2              =   2880
      Y1              =   480
      Y2              =   1320
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   240
      X2              =   16800
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   240
      X2              =   240
      Y1              =   480
      Y2              =   4320
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   1080
      X2              =   1080
      Y1              =   480
      Y2              =   1320
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   5280
      X2              =   5280
      Y1              =   480
      Y2              =   1320
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   7560
      X2              =   7560
      Y1              =   480
      Y2              =   1320
   End
   Begin VB.Line Line1 
      Index           =   6
      X1              =   4920
      X2              =   4920
      Y1              =   1320
      Y2              =   2400
   End
   Begin VB.Line Line1 
      Index           =   12
      X1              =   240
      X2              =   16800
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line1 
      Index           =   13
      X1              =   5640
      X2              =   5640
      Y1              =   1680
      Y2              =   2400
   End
   Begin VB.Line Line1 
      Index           =   15
      X1              =   7080
      X2              =   7080
      Y1              =   1320
      Y2              =   2400
   End
   Begin VB.Line Line1 
      Index           =   9
      X1              =   16800
      X2              =   16800
      Y1              =   480
      Y2              =   4320
   End
   Begin VB.Line Line1 
      Index           =   7
      X1              =   1320
      X2              =   1320
      Y1              =   1320
      Y2              =   2400
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Mant."
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
      Left            =   1515
      TabIndex        =   22
      Top             =   1560
      Width           =   930
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
      Index           =   7
      Left            =   2880
      TabIndex        =   21
      Top             =   1560
      Width           =   525
   End
End
Attribute VB_Name = "MEdfrm05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mObj As New clMantEd
Dim mObjLuser As New clLogUser
Dim mRec As New ADODB.Recordset

Private Sub Command1_Click(Index As Integer)
Dim mFecPro As String
Dim mFecTer As String
Dim mi As Integer
Dim mStrMO As String
Dim mNroParte As Double

If Index = 0 Then
   If fValida Then
      If MsgBox("¿Está Seguro de Grabar esta Orden?", vbYesNo, sMessage) = vbYes Then
         mFecPro = Now
         mFecTer = mFecPro

         'Veo el string para ManoObra
         mStrMO = ""
         For mi = 0 To List1.ListCount - 1
            If List1.Selected(mi) Then
               mStrMO = mStrMO & Left(List1.List(mi), InStr(1, List1.List(mi), "-"))
            End If
         Next

         'Grabo en Registros
         mNroParte = mObj.ObtMaxParte + 1
         'mObj.InsAdicional mNroParte, Format(CDate(Text1(1).Text & " " & Text1(13).Text), "yyyy-mm-dd hh:mm:ss"), Format(CDate(Text1(2).Text), "yyyy-mm-dd"), Text1(3).Text, Text1(4).Text, Combo1(0).Text, Text1(5).Text, Combo1(1).Text, "M", "MantEdil", Combo1(2).Text, Combo1(3).Text, Val(Text1(6).Text), Val(Text1(7).Text), Val(Text1(8).Text), mStrMO, Val(Text1(9).Text), Val(Text1(10).Text), Text1(11).Text, Text1(12).Text, "T", Trim(Right(MDI.mUser, 20)), mFecPro, mFecTer
         mObj.InsAdicional mNroParte, Format(CDate(Text1(1).Text & " " & Text1(13).Text), "yyyy-mm-dd hh:mm:ss"), Format(CDate(Text1(2).Text), "yyyy-mm-dd"), Text1(3).Text, Text1(4).Text, Combo1(0).Text, Text1(5).Text, Combo1(1).Text, "M", "MantEdil", Combo1(2).Text, Combo1(3).Text, Text1(6).Text, Text1(7).Text, Text1(8).Text, mStrMO, Text1(9).Text, Text1(10).Text, Text1(11).Text, Text1(12).Text, "T", Trim(Right(MDI.mUser, 20)), mFecPro, mFecTer
         'Blanqueo campos
         For mi = 0 To Text1.UBound
            Text1(mi).Text = ""
         Next
         For mi = 0 To Combo1.UBound
            Combo1(mi).ListIndex = -1
         Next
         
         'Blanqueo la Lista de mano de obra
         For mi = 0 To List1.ListCount - 1
            List1.Selected(mi) = False
         Next
         
         
         mNroParte = mObj.ObtMaxParte + 1
         Text1(0).Text = mNroParte
      End If
   End If
Else
   Unload Me
End If
End Sub

Private Sub Form_Load()
Dim mi As Integer
Dim mNroParte As Double

MEdfrm05.Top = 100
MEdfrm05.Left = (MDI.Width - MEdfrm05.Width) / 2

Set mRec = mObj.oEjecutarSelect("SELECT * FROM Edificios WHERE FechaBaja IS NULL ORDER BY ZonaMantEdil, Descripcion")
If Not mRec.EOF Then
   Do While Not mRec.EOF
      'Combo1(0).AddItem mRec!Tramo & " - " & mRec!descripcion
      Combo1(0).AddItem mRec!ZonaMantEdil & " - " & mRec!descripcion
      mRec.MoveNext
   Loop
End If
mRec.Close

Combo1(1).AddItem "Alta"
Combo1(1).AddItem "Media"
Combo1(1).AddItem "Baja"

Combo1(2).AddItem "Preventivo"
Combo1(2).AddItem "Predictivo"
Combo1(2).AddItem "Correctivo"

Set mRec = mObj.oEjecutarSelect("SELECT * FROM Rubros WHERE FechaBaja IS NULL")
If Not mRec.EOF Then
   Do While Not mRec.EOF
      Combo1(3).AddItem mRec!Codigo & " - " & mRec!descripcion
      mRec.MoveNext
   Loop
End If
mRec.Close

Set mRec = mObj.oEjecutarSelect("SELECT * FROM ManoObra WHERE FechaBaja IS NULL")
If Not mRec.EOF Then
   Do While Not mRec.EOF
      List1.AddItem mRec!Codigo & "-" & mRec!descripcion
      mRec.MoveNext
   Loop
End If
mRec.Close

mNroParte = mObj.ObtMaxParte + 1
Text1(0).Text = mNroParte
Text1(0).Enabled = False
Text1(9).Enabled = False
Text1(10).Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set mObj = Nothing
Set mRec = Nothing
ShowMenu 20, True, False
End Sub

Private Function fValida() As Boolean
Dim mRet As Boolean
Dim mi As Integer

mRet = True
For mi = 1 To Text1.UBound
   If mi <> 11 And mi <> 12 Then
      If mRet Then
         mRet = (Text1(mi).Text <> "")
      End If
   End If
Next
If mRet Then
   mRet = Fecha_ok(Text1(1).Text)
End If
If mRet Then
   mRet = Hora_ok(Text1(13).Text)
End If
If mRet Then
   mRet = Fecha_ok(Text1(2).Text)
End If
If mRet Then
   mRet = Hora_ok(Text1(3).Text)
End If
If mRet Then
   mRet = Hora_ok(Text1(4).Text)
End If
If mRet Then
   If DateDiff("n", CDate(Text1(1).Text & " " & Text1(13).Text & ":00"), CDate(Text1(2).Text & " " & Text1(3).Text & ":00")) <= 0 Then
      MsgBox "Verificar la fecha de Asistencia", vbCritical, "Atención"
      mRet = False
   End If
End If
If mRet Then
   For mi = 0 To Combo1.UBound
      If mRet Then
         mRet = (Combo1(mi).Text <> "")
      End If
   Next
End If
If Not mRet Then
   MsgBox "Verifique que los datos estén ingresados correctamente", vbCritical, "Atención"
End If
fValida = mRet
End Function

Private Sub List1_Click()
Dim mi As Integer
Dim mCant As Integer
If Text1(2).Text <> "" Then
   mCant = 0
   For mi = 0 To List1.ListCount - 1
      If List1.Selected(mi) Then
         mCant = mCant + 1
      End If
   Next
   Text1(10).Text = mCant * mObj.ObtCostoMO(Right(Text1(2).Text, 4) & Mid(Text1(2).Text, 4, 2))
End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
   Case 1, 2
      KeyAscii = fDateKeyPress(Text1(Index), KeyAscii)
   Case 3, 4, 13
      KeyAscii = fHoraKeyPress(Text1(Index), KeyAscii)
   Case 6, 7, 8, 9, 10
      KeyAscii = fNumDoubleKeyPress(KeyAscii)
End Select
End Sub
Private Sub Text1_LostFocus(Index As Integer)
If Index = 2 Then
   If Text1(1).Text <> "" And IsDate(Text1(1).Text) And Text1(2).Text <> "" And IsDate(Text1(2).Text) Then
      Text1(7).Text = DateDiff("d", CDate(Text1(1).Text), CDate(Text1(2).Text))
   Else
      If Not (IsDate(Text1(1).Text)) Then
         MsgBox "La Fecha de Solicitud es inválida", vbCritical, "Atención"
         Text1(1).SetFocus
         Exit Sub
      ElseIf Not (IsDate(Text1(2).Text)) Then
         MsgBox "La Fecha de Asistencia inválida", vbCritical, "Atención"
         Text1(2).SetFocus
         Exit Sub
      End If
   End If
End If
If Index = 4 Then
   If Text1(3).Text <> "" And IsDate(Text1(3).Text) And Text1(4).Text <> "" And IsDate(Text1(4).Text) Then
      Text1(9).Text = DateDiff("n", CDate(Text1(3).Text), CDate(Text1(4).Text)) / 60
   Else
      If Not (IsDate(Text1(3).Text)) Then
         MsgBox "La hora inicial de asistencia es inválida", vbCritical, "Atención"
         Text1(3).SetFocus
         Exit Sub
      ElseIf Not (IsDate(Text1(4).Text)) Then
         MsgBox "La hora final de asistencia es inválida", vbCritical, "Atención"
         Text1(4).SetFocus
         Exit Sub
      End If
   End If
End If
If Index = 6 Then
   Text1(8).Text = Val(Text1(6).Text) * 1.5
End If
End Sub
