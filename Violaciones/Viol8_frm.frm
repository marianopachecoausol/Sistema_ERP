VERSION 5.00
Begin VB.Form Viol8_frm 
   BackColor       =   &H00CECECE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Módulo de Modificaciones de Direcciones"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   6255
   Begin VB.CommandButton Command2 
      Caption         =   "&Salir"
      Height          =   375
      Index           =   1
      Left            =   4560
      TabIndex        =   9
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Modificar"
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
      Left            =   2160
      TabIndex        =   8
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar"
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
      Left            =   2400
      TabIndex        =   7
      Top             =   480
      Width           =   855
   End
   Begin VB.ComboBox Combo2 
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   2880
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2520
      Width           =   3015
   End
   Begin VB.ComboBox Combo2 
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   360
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   360
      TabIndex        =   4
      Top             =   1920
      Width           =   4215
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   360
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CECECE&
      Caption         =   "Código Postal"
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
      Left            =   2880
      TabIndex        =   13
      Top             =   2280
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CECECE&
      Caption         =   "Provincia"
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
      Left            =   360
      TabIndex        =   12
      Top             =   2280
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CECECE&
      Caption         =   "Dirección"
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
      Left            =   360
      TabIndex        =   11
      Top             =   1680
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CECECE&
      Caption         =   "Nombre"
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
      Left            =   360
      TabIndex        =   10
      Top             =   1080
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CECECE&
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
      Height          =   240
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   810
   End
End
Attribute VB_Name = "Viol8_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public pPatente As String
Dim mObjViol As New clViolaciones
Dim mRec As New ADODB.Recordset
Dim mI As Integer
  
Private Sub Form_Load()
   Me.Height = 4380
   Me.Width = 6345
   sAlinearForm Me
   Set mRec = mObjViol.oTabla("provincias", "order by 2")
   sLlenoCbo Viol8_frm.Combo2(0), mRec, 1, 0
   If pPatente = "" Then
      Set mRec = mObjViol.oDistPatenteTabla("direcciones", "", "", "")
   Else
      Set mRec = mObjViol.oDistPatenteTabla("direcciones", pPatente, "", "")
      If mRec.EOF Then
         Combo1.AddItem pPatente
         Command2(0).Enabled = True
      End If
   End If
   Do While Not mRec.EOF
     Combo1.AddItem mRec!patente
     mRec.MoveNext
   Loop
   mRec.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set mObjViol = Nothing
   Set mRec = Nothing
   If pPatente = "" Then
      ShowMenu 5, True, False
   Else
      pPatente = ""
   End If
End Sub

Private Sub Combo1_Click()
   If Combo1.ListIndex > -1 Then
      sMsgEspere Me, "Buscando datos....", True
      sBlanquear
      Set mRec = mObjViol.oTabla("direcciones", " where patente='" & Trim(Combo1.Text) & "'")
      If Not mRec.EOF Then
         Command2(0).Enabled = True
         Text1(1).Text = NVL(mRec!nombre, "")
         Text1(2).Text = NVL(mRec!domicilio, "")
         For mI = 0 To Combo2(0).ListCount - 1
            If Right(Combo2(0).List(mI), 2) = mRec!codpcia Then
               Combo2(0).ListIndex = mI
            End If
         Next
         For mI = 0 To Combo2(1).ListCount - 1
            If Trim(Left(Combo2(1).List(mI), InStr(Combo2(1).List(mI), "-") - 2)) = mRec!codpostal Then
               Combo2(1).ListIndex = mI
            End If
         Next
      End If
      mRec.Close
      sMsgEspere Me, "", False
   End If
End Sub

Private Sub Combo2_Click(Index As Integer)
   Dim mRec1 As New ADODB.Recordset
   If Index = 0 Then
      Combo2(1).Clear
      Set mRec1 = mObjViol.oTabla("postal", " where codpcia='" & Right(Combo2(0).Text, 2) & "' order by 1")
      Do While Not mRec1.EOF
         Combo2(1).AddItem mRec1!Codigo & " - " & mRec1!Descripcion
         mRec1.MoveNext
      Loop
      mRec1.Close
   End If
   Set mRec1 = Nothing
End Sub

Private Sub Command1_Click()
 Dim mWhere As String
   If Command1.Caption = "Buscar" Then
      Command1.Caption = "OK"
      Command2(0).Enabled = False
      sBlanquear
      sCambiarObj True
      Command1.SetFocus
   Else
      Combo1.Clear
      Set mRec = mObjViol.oDistPatenteTabla("direcciones", Trim(Text1(0).Text), Trim(Text1(2).Text), Trim(Text1(1).Text))
      Do While Not mRec.EOF
         Combo1.AddItem mRec!patente
         mRec.MoveNext
      Loop
      mRec.Close
      Command1.Caption = "Buscar"
      sCambiarObj False
   End If
End Sub

Private Sub Command2_Click(Index As Integer)
   Dim mFlag As Boolean
   Select Case Index
      Case 0
         If Command2(0).Caption = "&Modificar" Then
            sVerObj True
            Command2(0).Caption = "&Grabar"
            Command2(1).Caption = "Volver"
         Else 'Grabo y listo
            If Text1(1).Text <> "" And Text1(2).Text <> "" And Combo2(0).Text <> "" And Combo2(1).Text <> "" Then
               Set mRec = mObjViol.oDatosPatente(Trim(Combo1.Text))
               If Not mRec.EOF Then
                  mFlag = mObjViol.xUpDirecciones(Trim(Text1(1).Text), Trim(Text1(2).Text), Right(Combo2(0).Text, 2), Trim(Left(Combo2(1).Text, InStr(Combo2(1).Text, "-") - 2)), Trim(Combo1.Text))
               Else
                  mFlag = mObjViol.xInsDirecciones(Trim(Combo1.Text), Trim(Text1(1).Text), Trim(Text1(2).Text), Right(Combo2(0).Text, 2), Trim(Left(Combo2(1).Text, InStr(Combo2(1).Text, "-") - 2)))
               End If
               mRec.Close
               sVerObj False
               Command2(0).Caption = "&Modificar"
            Else
               MsgBox "Falta completar algún datos.", vbCritical, sMessage & "Atención"
            End If
         End If
         
      Case 1
         If Command2(1).Caption = "Volver" Then
            sVerObj False
            Command2(0).Caption = "&Modificar"
            Command2(1).Caption = "&Salir"
         Else
            If pPatente <> "" Then
               Viol6_frm.Enabled = True
            End If
            Unload Viol8_frm
         End If
   End Select
End Sub

Private Sub sBlanquear()
   Text1(0).Text = ""
   Text1(1).Text = ""
   Text1(2).Text = ""
   Combo2(0).ListIndex = -1
   Combo2(1).ListIndex = -1
End Sub

Private Sub sCambiarObj(ByVal pFlag As Boolean)
   If pFlag Then
      Combo2(0).BackColor = &HCECECE
      Combo2(1).BackColor = &HCECECE
   Else
      Combo2(0).BackColor = &HFFFFFF
      Combo2(1).BackColor = &HFFFFFF
   End If
   Combo1.Visible = Not pFlag
   Text1(0).Visible = pFlag
   Text1(1).Enabled = pFlag
   Text1(2).Enabled = pFlag
End Sub

Private Sub sVerObj(ByVal pFlag As Boolean)
   Combo1.Enabled = Not pFlag
   Command1.Enabled = Not pFlag
   Text1(1).Enabled = pFlag
   Text1(2).Enabled = pFlag
   Combo2(0).Enabled = pFlag
   Combo2(1).Enabled = pFlag
End Sub
