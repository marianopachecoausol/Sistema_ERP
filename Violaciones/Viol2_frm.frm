VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Viol2_frm 
   BackColor       =   &H00B3C1CC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Módulo de Carga de Violaciones"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10890
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   10890
   Begin MSFlexGridLib.MSFlexGrid MSFlex 
      Height          =   5535
      Left            =   50
      TabIndex        =   14
      Top             =   2400
      Width           =   10820
      _ExtentX        =   19103
      _ExtentY        =   9763
      _Version        =   327680
      Cols            =   10
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   12704728
      GridColor       =   8388608
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00B3C1CC&
      Height          =   2355
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   10780
      Begin VB.CheckBox Check1 
         BackColor       =   &H00B3C1CC&
         Caption         =   "Es Deuda?"
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
         Left            =   5880
         TabIndex        =   15
         Top             =   1710
         Width           =   1300
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   5
         Left            =   4320
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00B3C1CC&
         Height          =   1575
         Left            =   7440
         TabIndex        =   28
         Top             =   720
         Width           =   3255
         Begin VB.CommandButton Command1 
            Caption         =   "Salir"
            Height          =   375
            Index           =   2
            Left            =   1560
            TabIndex        =   13
            Top             =   960
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00B3C1CC&
            Caption         =   "Modificar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   1
            Left            =   1440
            TabIndex        =   12
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H8000000B&
            Caption         =   "Grabar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   240
         TabIndex        =   10
         Top             =   2000
         Width           =   6975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   4
         Left            =   2280
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1320
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   3
         Left            =   240
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1320
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   2
         Left            =   5760
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   2
         Left            =   5880
         MaxLength       =   8
         TabIndex        =   9
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   1
         Left            =   1680
         MaxLength       =   5
         TabIndex        =   2
         Top             =   600
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   4560
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   2520
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   0
         Left            =   240
         MaxLength       =   10
         TabIndex        =   1
         Top             =   600
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00B3C1CC&
         Height          =   615
         Left            =   7440
         TabIndex        =   16
         Top             =   120
         Width           =   3255
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00B3C1CC&
            Caption         =   "Carga de Violaciones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   360
            Left            =   240
            TabIndex        =   17
            Top             =   165
            Width           =   2700
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00B3C1CC&
         Caption         =   "Nuevo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   5040
         MouseIcon       =   "Viol2_frm.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   31
         ToolTipText     =   "Agregar Color"
         Top             =   1060
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00B3C1CC&
         Caption         =   "Nuevo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   3360
         MouseIcon       =   "Viol2_frm.frx":0152
         MousePointer    =   99  'Custom
         TabIndex        =   30
         ToolTipText     =   "Agregar Modelo"
         Top             =   1060
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00B3C1CC&
         Caption         =   "Nuevo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   2
         Left            =   1320
         MouseIcon       =   "Viol2_frm.frx":02A4
         MousePointer    =   99  'Custom
         TabIndex        =   29
         ToolTipText     =   "Agregar Marca"
         Top             =   1060
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00B3C1CC&
         Caption         =   "Color"
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
         Left            =   4320
         TabIndex        =   27
         Top             =   1080
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00B3C1CC&
         Caption         =   "Sentido"
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
         Left            =   5760
         TabIndex        =   26
         Top             =   360
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00B3C1CC&
         Caption         =   "Hora"
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
         Left            =   1680
         TabIndex        =   25
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00B3C1CC&
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
         Index           =   6
         Left            =   240
         TabIndex        =   24
         Top             =   1720
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00B3C1CC&
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
         Index           =   5
         Left            =   5880
         TabIndex        =   23
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00B3C1CC&
         Caption         =   "Modelo"
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
         Left            =   2280
         TabIndex        =   22
         Top             =   1080
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00B3C1CC&
         Caption         =   "Marca"
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
         Left            =   240
         TabIndex        =   21
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00B3C1CC&
         Caption         =   "Vía"
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
         Left            =   4560
         TabIndex        =   20
         Top             =   360
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00B3C1CC&
         Caption         =   "Estación"
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
         Left            =   2520
         TabIndex        =   19
         Top             =   360
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00B3C1CC&
         Caption         =   "Fecha"
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
         TabIndex        =   18
         Top             =   360
         Width           =   540
      End
   End
End
Attribute VB_Name = "Viol2_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mObjViol As New clViolaciones
Dim mObjLUser As New clLogUser
Dim mObjPea As New clPeaje
Dim mRec As New ADODB.Recordset
Dim mVectEstaciones(18) As String
Dim mI As Integer

Private Sub Form_Load()
Me.MousePointer = 11
sAlinearForm Me
sMsgEspere Me, "Iniciando...", True
sInitForm
sMsgEspere Me, "", False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set mObjPea = Nothing
Set mObjViol = Nothing
Set mRec = Nothing
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
   Case 0  'GRABAR
      If fValidar Then
         Me.MousePointer = 11
         sMsgEspere Me, "Procesando... espere un momento", True
         If Command1(1).Visible Then
            mObjViol.xInsRegistros Text1(0).Text, Text1(1).Text, Left(Combo1(0).Text, 2), Combo1(1).Text, Right(Combo1(2).Text, 1), Trim(Text1(2).Text), Right(Combo1(4).Text, 2), Right(Combo1(5).Text, 2), Trim(Text1(3).Text), "", "", "", "", Right(Combo1(3).Text, 2), "", IIf(Check1.Value = 0, "V", "D")
            MSFlex.AddItem Text1(0).Text & vbTab & Text1(1).Text & vbTab & Combo1(0).Text & vbTab & Combo1(1).Text & vbTab & Left(Combo1(2).Text, 1) & vbTab & Combo1(3).Text & vbTab & Combo1(4).Text & vbTab & Trim(Text1(2).Text) & vbTab & Combo1(5).Text & vbTab & Trim(Text1(3).Text), 1
            sLimpiar
         Else 'Actualizar
            If mObjViol.xUpdRegistros(Text1(0).Text, Text1(1).Text, Left(Combo1(0).Text, 2), Combo1(1).Text, Right(Combo1(2).Text, 1), Trim(Text1(2).Text), Right(Combo1(4).Text, 2), Right(Combo1(5).Text, 2), Trim(Text1(3).Text), Right(Combo1(3).Text, 2), MSFlex.TextMatrix(MSFlex.Row, 0), MSFlex.TextMatrix(MSFlex.Row, 1), Left(MSFlex.TextMatrix(MSFlex.Row, 2), 2), MSFlex.TextMatrix(MSFlex.Row, 7)) Then
               sInputGrid
               Command1(1).Visible = True
               Command1(2).Caption = "Salir"
               sLimpiar
            End If
         End If
         Me.MousePointer = 0
         sMsgEspere Me, "", False
         Text1(0).SetFocus
      End If
   Case 1  'MODIFICAR
       If MSFlex.TextMatrix(MSFlex.Row, 0) = Date - 1 Then
          sMsgEspere Me, "Procesando... espere un momento", True
          sLlenar
          MSFlex.Col = 0
          MSFlex.ColSel = 9
          MSFlex.Enabled = False
          Command1(2).Caption = "Volver"
          Command1(1).Visible = False
          sMsgEspere Me, "", False
       Else
          MsgBox "Sólo se pueden modificar registros de ayer", vbCritical, "Atención"
       End If
   Case 2  'SALIR
       If Command1(2).Caption = "Salir" Then
          Unload Me
          ShowMenu 5, True, False
       Else
          Command1(2).Caption = "Salir"
          Command1(1).Visible = True
          sLimpiar
       End If
End Select
End Sub

Private Sub Combo1_Click(Index As Integer)
Select Case Index
   Case 0
      If Combo1(0).Text <> "" Then
         Combo1(1).Clear
         Combo1(2).ListIndex = -1
         Set mRec = mObjPea.oViasEstacion(Left(Combo1(0).Text, 2))
         Do While Not mRec.EOF
            Combo1(1).AddItem Format(mRec!NUM_VIA, "0#") & mRec!Sentido '& Space(20) & mRec!TIPO_COBRO
            mRec.MoveNext
         Loop
         mRec.Close
      End If
   Case 1, 2
      If Combo1(1).Text <> "" Then
         If Right(Combo1(1).Text, 1) = "A" Then
            Combo1(2).ListIndex = 1
         Else
            Combo1(2).ListIndex = 2
         End If
      End If
   Case 3
      Combo1(4).Clear
      Set mRec = mObjViol.oTabla("modelos", " where codmarca='" & Right(Combo1(Index).Text, 2) & "' order by 3")
      sLlenoCbo Viol2_frm.Combo1(4), mRec, 2, 1
   End Select
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 27 Then
   Combo1(Index).ListIndex = -1
End If
If Index = 0 Then
   Combo1(1).Clear
End If
End Sub

Private Sub Label3_Click(Index As Integer)
Viol3_frm.sInitForm Index
Viol3_frm.pViol2_View = True
Viol3_frm.Show
End Sub

Private Sub Label3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3(Index).BorderStyle = 1
End Sub

Private Sub Label3_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3(Index).BorderStyle = 0
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
   Case 0 'Fecha
      KeyAscii = fDateKeyPress(Text1(Index), KeyAscii)
   Case 1 'Hora
      KeyAscii = fHoraKeyPress(Text1(Index), KeyAscii)
   Case 2
      If KeyAscii = 32 Then
         KeyAscii = 0
      Else
         KeyAscii = fAlfaNumKeyPress(KeyAscii)
         KeyAscii = fUcaseKeyPress(KeyAscii)
      End If
   Case 3
End Select
End Sub

Private Function sLlenarFGrid()
Dim mMarca As String
Dim mModelo As String
Dim mColor As String
Dim mWhere As String
mWhere = ""
If Combo1(0).ListCount < 10 Then
   For mI = 0 To Combo1(0).ListCount - 1
      mWhere = "'" & Left(Combo1(0).List(mI), 2) & "',"
   Next
   mWhere = "where estacion in (" & mWhere & mId(mWhere, 1, Len(mWhere) - 1) & ") "
End If
Set mRec = mObjViol.oTabla("Registros", mWhere & " order by fecha desc limit 500")
Do While Not mRec.EOF
   mMarca = mObjViol.sCampoDescrip("marcas", "codigo='" & mRec!CodMarca & "'", 1) & Space(40) & NVL(mRec!CodMarca, "")
   mModelo = mObjViol.sCampoDescrip("modelos", "codigo='" & mRec!modelo & "' AND codmarca='" & mRec!CodMarca & "'", 2) & Space(40) & NVL(mRec!modelo, "")
   mColor = NVL(mRec!Color, "")
   If Len(mColor) < 3 Then
      mColor = mObjViol.sCampoDescrip("colores", "codigo='" & mRec!Color & "'", 1) & Space(40) & mRec!Color
   End If
   MSFlex.AddItem Format(mRec.Fields(0), "dd/mm/yyyy") & vbTab & mRec.Fields(1) & vbTab & mRec.Fields(2) & "-" & Trim(mVectEstaciones(mRec.Fields(2))) & vbTab & mRec.Fields(3) & vbTab & mRec.Fields(4) & vbTab & mMarca & vbTab & mModelo & vbTab & mRec!patente & vbTab & mColor & vbTab & mRec!OBS
   mRec.MoveNext
Loop
If MSFlex.Rows > 2 Then
   MSFlex.RemoveItem 1
End If
sSetFlex2Colors Viol2_frm.MSFlex, &HF0F9F9, &HD7FDFF
sSetFlexColOrder Viol2_frm.MSFlex, 1
End Function

Private Sub sInputGrid()
With MSFlex
   .TextMatrix(.Row, 0) = Text1(0).Text
   .TextMatrix(.Row, 1) = Text1(1).Text
   .TextMatrix(.Row, 2) = Combo1(0).Text
   .TextMatrix(.Row, 3) = Combo1(1).Text
   .TextMatrix(.Row, 4) = Left(Combo1(2).Text, 1)
   .TextMatrix(.Row, 5) = Combo1(3).Text
   .TextMatrix(.Row, 6) = Combo1(4).Text
   .TextMatrix(.Row, 7) = Trim(Text1(2).Text)
   .TextMatrix(.Row, 8) = Combo1(5).Text
   .TextMatrix(.Row, 9) = Trim(Text1(3).Text)
End With
End Sub

Private Function fValidar() As Boolean
fValidar = False
If Fecha_ok(Text1(0).Text) And Hora_ok(Text1(1).Text) Then
   If DateDiff("d", Date, Text1(0).Text) <= 0 Then
      If Text1(2).Text <> "" Then
         If Len(Text1(2).Text) >= 6 Then
            If Combo1(0).ListIndex > -1 Then
               fValidar = True
            Else
               MsgBox "Debe elegir al menos la Estación", vbExclamation, "Sistema de Violaciones"
            End If
         Else
            MsgBox "La Patente debe tener al menos 6 dígitos", vbExclamation, "Sistema de Violaciones"
         End If
      Else
         MsgBox "Es necesario ingresar la Patente", vbExclamation, "Sistema de Violaciones"
      End If
   Else
      MsgBox "Fecha de ingreso es mayor a la actual.", vbExclamation, "Sistema de Violaciones"
   End If
End If
End Function

Private Sub sLimpiar()
Text1(1).Text = ""
Text1(2).Text = ""
For mI = 1 To Combo1.UBound
   Combo1(mI).ListIndex = -1
Next
Check1.Value = 0
MSFlex.Enabled = True
End Sub

Private Sub sLlenar()
Dim mJ As Integer
With MSFlex
   Text1(0).Text = .TextMatrix(.Row, 0)
   Text1(1).Text = .TextMatrix(.Row, 1)
   For mJ = 0 To 4
      For mI = 0 To Combo1(mJ).ListCount - 1
         If Trim(Left(Combo1(mJ).List(mI), 2)) = Trim(Left(.TextMatrix(.Row, mJ + 2), 2)) Then
            Combo1(mJ).ListIndex = mI
         End If
      Next
   Next
   For mJ = 3 To 4
      For mI = 0 To Combo1(mJ).ListCount - 1
         If Trim(Right(Combo1(mJ).List(mI), 2)) = Trim(Right(.TextMatrix(.Row, mJ + 2), 2)) Then
            Combo1(mJ).ListIndex = mI
         End If
      Next
   Next
   For mI = 0 To Combo1(5).ListCount - 1
      If Trim(Left(Combo1(5).List(mI), 2)) = Trim(Left(.TextMatrix(.Row, 8), 2)) Then
         Combo1(5).ListIndex = mI
      End If
   Next
   Text1(2).Text = MSFlex.TextMatrix(.Row, 7)
   Text1(3).Text = MSFlex.TextMatrix(.Row, 9)
End With
End Sub

Private Sub sInitForm()
Dim mSector As String
mSector = ""
Select Case mObjLUser.sCampoDescrip("USUARIOS", "codusuario='" & Trim(Right(MDI.mUser, 16)) & "'", 6)
   Case "svergaraa@gco.com.ar", "svergarad@gco.com.ar"
      mSector = "01"
   Case "ssantarosa@gco.com.ar"
      mSector = "02"
   Case "situzaingo@gco.com.ar", "sdecalada@gco.com.ar"
      mSector = "03"
   Case "slujan@gco.com.ar"
      mSector = "04"
End Select
Set mRec = mObjPea.oEstaciones(mSector)
Do While Not mRec.EOF
   Combo1(0).AddItem Format(mRec!CODIGO_ESTACION, "0#") & " - " & mRec!Descripcion_Estacion
   mVectEstaciones(mRec!CODIGO_ESTACION) = mRec!Descripcion_Estacion
   mRec.MoveNext
Loop
mRec.Close
Set mObjLUser = Nothing
Combo1(2).AddItem " "
Combo1(2).AddItem "Ascendente                              A"
Combo1(2).AddItem "Descendente                             D"
Set mRec = mObjViol.oTabla("marcas", " order by 2")
sLlenoCbo Viol2_frm.Combo1(3), mRec, 1, 0
Set mRec = mObjViol.oTabla("colores", " order by 2")
sLlenoCbo Viol2_frm.Combo1(5), mRec, 1, 0
With MSFlex
   .ColWidth(0) = 1000
   .ColWidth(1) = 600
   .ColWidth(2) = 1200
   .ColWidth(3) = 500
   .ColWidth(4) = 400
   .ColWidth(5) = 1700
   .ColWidth(6) = 1500
   .ColWidth(7) = 1300
   .ColWidth(8) = 1000
   .ColWidth(9) = 5000
   .TextMatrix(0, 0) = "Fecha"
   .TextMatrix(0, 1) = "Hora"
   .TextMatrix(0, 2) = "Estación"
   .TextMatrix(0, 3) = "Vía"
   .TextMatrix(0, 4) = "Sen"
   .TextMatrix(0, 5) = "Marca"
   .TextMatrix(0, 6) = "Modelo"
   .TextMatrix(0, 7) = "Patente"
   .TextMatrix(0, 8) = "Color"
   .TextMatrix(0, 9) = "Observaciones"
   For mI = 0 To 9
      .Col = mI
      .CellFontBold = True
   Next
End With
MSFlex.Row = 0
For mI = 0 To MSFlex.Cols - 1
   MSFlex.Col = mI
   MSFlex.CellFontBold = True
Next
sLlenarFGrid
Me.MousePointer = 0
End Sub
