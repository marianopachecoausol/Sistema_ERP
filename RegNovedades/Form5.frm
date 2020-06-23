VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form RNov5_frm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Asignación de Turnos"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8145
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   5010
      Left            =   50
      TabIndex        =   0
      Top             =   0
      Width           =   8010
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   4
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   3450
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   5160
         MaxLength       =   6
         TabIndex        =   6
         Top             =   2400
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   1440
         TabIndex        =   14
         Top             =   3840
         Width           =   3735
         Begin VB.CommandButton Command1 
            Caption         =   "&Volver"
            Height          =   495
            Index           =   1
            Left            =   2160
            TabIndex        =   8
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Aceptar"
            Default         =   -1  'True
            Height          =   495
            Index           =   0
            Left            =   360
            TabIndex        =   7
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   4800
         MaxLength       =   30
         TabIndex        =   5
         Top             =   3120
         Width           =   2775
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   2
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2400
         Width           =   2775
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   5160
         MaxLength       =   6
         TabIndex        =   3
         Top             =   1680
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Tag             =   "0"
         Top             =   1680
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   3
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   3075
         Width           =   2775
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Ver detalle"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   6450
         MousePointer    =   99  'Custom
         TabIndex        =   30
         Top             =   1500
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Patrullero auxiliar"
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
         Left            =   1320
         TabIndex        =   18
         Top             =   2820
         Width           =   1485
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   855
         Left            =   6240
         Stretch         =   -1  'True
         ToolTipText     =   "Click Para Detalle"
         Top             =   3840
         Width           =   1410
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1320
         TabIndex        =   17
         Top             =   1080
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Km Final"
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
         Left            =   5160
         TabIndex        =   16
         Top             =   2160
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Asignación de Turnos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   2520
         TabIndex        =   15
         Top             =   600
         Width           =   2610
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Turno"
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
         Left            =   3120
         TabIndex        =   13
         Top             =   1440
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Km Inicial"
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
         Left            =   5160
         TabIndex        =   12
         Top             =   1440
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Policía"
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
         Left            =   4800
         TabIndex        =   11
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Patrullero"
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
         Left            =   1320
         TabIndex        =   10
         Top             =   2160
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Móvil"
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
         Left            =   1320
         TabIndex        =   9
         Top             =   1440
         Width           =   465
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   5265
      Left            =   10000
      TabIndex        =   20
      Top             =   0
      Width           =   8250
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Volver"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   7200
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   50
         Width           =   990
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00EAFFEE&
         Caption         =   "Actualizar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   5700
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   50
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox Text2 
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
         ForeColor       =   &H00400000&
         Height          =   285
         Index           =   1
         Left            =   4650
         TabIndex        =   25
         Top             =   75
         Width           =   850
      End
      Begin VB.TextBox Text2 
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
         ForeColor       =   &H00400000&
         Height          =   285
         Index           =   0
         Left            =   3000
         TabIndex        =   24
         Top             =   75
         Width           =   850
      End
      Begin VB.ComboBox Combo2 
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
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   75
         Width           =   1365
      End
      Begin MSFlexGridLib.MSFlexGrid Flex1 
         Height          =   4700
         Left            =   105
         TabIndex        =   21
         ToolTipText     =   "Doble clic en la fila para actualizar kilometraje."
         Top             =   375
         Width           =   8145
         _ExtentX        =   14367
         _ExtentY        =   8281
         _Version        =   327680
         Cols            =   7
         FixedCols       =   0
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Km.F:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   4050
         TabIndex        =   27
         Top             =   150
         Width           =   510
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Km. I:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   2400
         TabIndex        =   26
         Top             =   150
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Móvil:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   22
         Top             =   150
         Width           =   510
      End
   End
End
Attribute VB_Name = "RNov5_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Fecha As String
Public mModif As Boolean

Private Sub Form_Load()
Dim mObj As New clRNov
Dim mRec As New ADODB.Recordset
   
   sAlinearForm Me
   mModif = False
   Set mRec = mObj.oTabla("moviles", "WHERE CodTipoMov in ('PAT','GRU') ORDER BY Descripcion")
   sLlenoCbo Me.Combo1(0), mRec, 1, 0
   Set mRec = mObj.oTabla("patrulleros", "WHERE Fecha_Baja IS NULL ORDER BY 2")
   Do While Not mRec.EOF
      Combo1(2).AddItem mRec!Codigo & " - " & mRec!nombre
      Combo1(3).AddItem mRec!Codigo & " - " & mRec!nombre
      Combo1(4).AddItem mRec!Codigo & " - " & mRec!nombre
      mRec.MoveNext
   Loop
   mRec.Close
   Image1.Picture = LoadPicture()
   Set mRec = Nothing
   Set mObj = Nothing
End Sub

Private Sub Combo2_Click()
Dim mObj As New clRNov
Dim mRec As New ADODB.Recordset
Dim mChofer As String
   
   sBorraFlexDatos Flex1
   If Combo2.ListIndex = 0 Then
      sInitFlex
   End If
   If Combo2.ListIndex > 0 Then
      Set mRec = mObj.oMovilTurnos(Combo2.Text, 15)
      Do While Not mRec.EOF
         mChofer = mRec!chofer
         If Left(Combo2.Text, 1) = "M" Then
            mChofer = NVL(mRec!nombre, "")
         End If
         Flex1.AddItem mRec!CodMovil & vbTab & mRec!fechainic & vbTab & mRec!KmInicial & vbTab & mRec!KmFinal & vbTab _
            & mChofer & vbTab & mRec!police & vbTab & mRec!Descripcion
         mRec.MoveNext
      Loop
      mRec.Close
   End If
   If Flex1.Rows > 2 Then
      Flex1.RemoveItem 1
   End If
   sSetFlex2Colors Flex1, &HFFFFFF, &HF5F5F5
   Set mRec = Nothing
   Set mObj = Nothing
End Sub

Private Sub Command1_Click(Index As Integer)
Dim mObj As New clRNov
Dim mError As Boolean
Dim mFlag As Boolean
Dim mFecha As Date
Dim mFechaF As Date
Dim mPatru As String
Dim mQuery As String
Dim Km1 As Long
Dim Km2 As Long
Dim mResp As String
Dim mI As Integer
 
   mError = False
   mFecha = Now
   mPatru = ""
   mFlag = False
   Select Case Index
      Case 0
         If Label2(5).Visible Then ' KM FINAL
            If Text1(2).Text <> "" Then
               If IsDate(Right(Label3.Caption, 19)) Then
                  mFecha = Right(Label3.Caption, 19)
               Else
                  mFecha = Right(Label3.Caption, 24)
               End If
               
               Km1 = Text1(0).Text
               Km2 = Trim(Text1(2).Text)
               If Km1 < Km2 Then
                  mResp = vbYes
                  If (Km2 - Km1) > 600 Then
                     mResp = MsgBox("Según los Kilometrajes el móvil Recorrió: " & (Km2 - Km1) & " Km. Es Correcto este Dato?", vbYesNo, sMessage)
                  End If
                  If mResp = vbYes Then
                     mFechaF = Now
                     mObj.xUpMovilTurnosFin mFecha, Right(Combo1(0).Text, 4), Now(), Trim(Text1(2).Text)
                     Label3.Visible = False
                     Label2(5).Visible = False
                     Text1(0).Text = Text1(2).Text
                     Text1(0).Enabled = True
                     Text1(2).Text = ""
                     Text1(2).Visible = False
                     If Right(Combo1(1).Text, 1) = "1" Or Right(Combo1(1).Text, 1) = "2" Or Right(Combo1(1).Text, 1) = "3" Then
                        Combo1(2).Enabled = True
                        Combo1(2).ListIndex = -1
                     End If
                     Text1(1).Enabled = True
                     Text1(1).Text = ""
                     For mI = 1 To 4
                        Combo1(mI).ListIndex = -1
                        Combo1(mI).Enabled = True
                     Next
                  End If
               Else
                  MsgBox "Km final es menor que km Inicial", vbCritical, sMessage
               End If
            Else
               MsgBox "Faltan Ingresar Datos.", vbCritical, sMessage
            End If
         
         Else 'KM INICIAL
            If Combo1(0).Text <> "" And Text1(0).Text <> "" And Combo1(1).Text <> "" And Text1(1).Text <> "" Then
               If Command1(0).Caption <> "&Aceptar" Then
                  mFecha = Right(Label3.Caption, 19)
               End If
               If Left(Combo1(0).Text, 4) = "MÓVI" Then
                  If Combo1(2).Text = "" Then
                     MsgBox "Faltan Ingresar Datos.", vbCritical, sMessage
                  Else
                     If fKilometros(Km2, Trim(Text1(0).Text)) Then
                        mPatru = Left(Combo1(2).Text, 3)
                        If Command1(0).Caption = "Actualizar" Then
                           mObj.xUpMovilTurnos mFecha, mPatru, "", Trim(Text1(1).Text), Trim(Text1(0).Text), Right(Combo1(1).Text, 1), Trim(Left(Combo1(3).Text, 3)), Trim(Left(Combo1(4).Text, 3))
                        Else
                            mObj.xInsMovilTurno Right(Combo1(0).Text, 4), mPatru, "", Trim(Text1(1).Text), Now(), Trim(Text1(0).Text), Right(Combo1(1).Text, 1), Trim(Left(Combo1(3).Text, 3)), Trim(Left(Combo1(4).Text, 3))
                        End If
                        sFormEnabled True
                        sClearAll
                     End If
                  End If
               Else
                  If fKilometros(Km2, Trim(Text1(0).Text)) Then
                     If Command1(0).Caption = "Actualizar" Then
                        mObj.xUpMovilTurnos mFecha, "", Trim(Text1(1).Text), "", Trim(Text1(0).Text), Right(Combo1(1).Text, 1), Trim(Left(Combo1(3).Text, 3)), Trim(Left(Combo1(4).Text, 3))
                     Else
                        mObj.xInsMovilTurno Right(Combo1(0).Text, 4), "", Trim(Text1(1).Text), "", Now(), Trim(Text1(0).Text), Right(Combo1(1).Text, 1), Trim(Left(Combo1(3).Text, 3)), Trim(Left(Combo1(4).Text, 3))
                     End If
                     sFormEnabled True
                     sClearAll
                  End If
               End If
            Else
               MsgBox "Faltan Ingresar Datos.", vbCritical, sMessage
            End If
         End If
         If mModif Then
            mModif = False
            Command1(0).Caption = "&Aceptar"
            Combo1(0).Enabled = True
            Image1.Enabled = True
         End If
    Case 1
      If mModif Then
         Command1(0).Caption = "&Aceptar"
         Command1(0).Enabled = False
         Command1(1).Enabled = False
         Combo1(0).Enabled = True
         Frame3.Visible = True
         mModif = False
         Image1.Enabled = True
      Else
         Unload RNov5_frm
         sFormEnabled True
      End If
   End Select
   Set mObj = Nothing
End Sub

Private Sub Command2_Click(Index As Integer)
Dim mObj As New clRNov
Dim mdif As Integer
Dim mI As Integer

   If Index = 0 Then 'actualizar
      'habria que comparar
      mdif = 100
      If Text2(1).Text <> "" Then
         mdif = Val(Text2(1).Text) - Val(Text2(0).Text)
      End If
      If mdif <= 1000 Then
         mObj.xUpTurnosKm Command2(0).Tag, Flex1.TextMatrix(Command2(1).Tag, 1), Text2(0).Text, Text2(1).Text
         Command2(0).Visible = False
         Command2(1).Caption = "Volver"
         Flex1.TextMatrix(Command2(1).Tag, 2) = Text2(0).Text
         Flex1.TextMatrix(Command2(1).Tag, 3) = Text2(1).Text
         Command2(0).Tag = ""
         Command2(1).Tag = ""
         Text2(0).Enabled = False
         Text2(1).Enabled = False
         Text2(0).Text = ""
         Text2(1).Text = ""
         Flex1.Enabled = True
      Else
         MsgBox "Existe una diferencia de más de 1000 km " & Chr(13) & "entre el km inicial y final. Verifique.", vbExclamation, sMessage
      End If
   Else
      If Command2(1).Caption = "Volver" Then
         Frame3.Left = 10000
         Frame1.Left = 50
      Else
         Command2(0).Visible = False
         Command2(1).Caption = "Volver"
         Command2(0).Tag = ""
         Command2(1).Tag = ""
         For mI = 0 To 1
            Text2(mI).Enabled = False
            Text2(mI).Text = ""
         Next
      End If
      Flex1.Enabled = True
   End If
End Sub

Private Sub Combo1_Click(Index As Integer)
Dim mObj As New clRNov
Dim mRec As New ADODB.Recordset
Dim nMov As String
Dim mI As Integer

   If Index = 0 Then
      If Not mModif Then
         sClearAll
      End If
      If Left(Combo1(0).Text, 4) = "MÓVI" Then
         Label2(7).Visible = True
         Combo1(3).Visible = True
         Combo1(4).Visible = True
         Image1.Picture = LoadPicture(App.Path & "\RegNovedades\Image\Patrulla_Vial.JPG")
         If Combo1(0).Tag <> "1" Then
            Combo1(1).Clear
            Set mRec = mObj.oTabla("turnos", "where CodTipoMovil = 'PAT'")
            sLlenoCbo Me.Combo1(1), mRec, 1, 0
            Combo1(0).Tag = "1"
            Combo1(2).Visible = True
            Text1(1).Top = 3180
            Text1(1).Left = 4800
            Label2(1).Caption = "Patrullero"
            Label2(2).Visible = True
         Else
            Combo1(1).ListIndex = -1
         End If
      Else
         If Combo1(0).Tag <> "2" Then
            Image1.Picture = LoadPicture(App.Path & "\RegNovedades\Image\Servicio_Vial.JPG")
            Combo1(1).Clear
            Set mRec = mObj.oTabla("turnos", "where CodTipoMovil = 'GRU'")
            sLlenoCbo Me.Combo1(1), mRec, 1, 0
            Combo1(0).Tag = "2"
            For mI = 2 To 4
               Combo1(mI).ListIndex = -1
               Combo1(mI).Visible = False
            Next
            Text1(1).Top = 2400
            Text1(1).Left = 1320
            Label2(1).Caption = "Chofer"
            Label2(2).Visible = False
            Label2(7).Visible = False
         Else
            Combo1(1).ListIndex = -1
         End If
      End If
      nMov = Right(Combo1(Index).Text, 4)
      Fecha = ""
      Set mRec = mObj.oTabla("movilturno", "where CodMovil = '" & nMov & "' AND (KmFinal IS NULL or KmFinal = '') ORDER BY FechaInic")
      If Not mRec.EOF Then
         sSelectCboRigth Combo1(1), mRec!codturno
         Text1(0).Text = mRec!KmInicial
         If mRec!CodPatrullero <> "" Then
             sSelectCboLeft Combo1(2), mRec!CodPatrullero
             sSelectCboLeft Combo1(3), NVL(mRec!codpatrullero2, "")
             sSelectCboLeft Combo1(4), NVL(mRec!codpatrullero3, "")
         End If
         If Left(mRec!CodMovil, 1) = "G" Or Left(mRec!CodMovil, 1) = "6" Then
            Text1(1).Text = mRec!chofer
            Text1(1).Left = 1320
         Else
            Text1(1).Text = mRec!police
            Text1(1).Left = 4800
         End If
         Label2(5).Visible = True
         Text1(2).Visible = True
         Fecha = mRec!fechainic
         For mI = 1 To 4
            Combo1(mI).Enabled = False
         Next
         Text1(0).Enabled = False
         Text1(1).Enabled = False
         Label3.Caption = "Fecha Inicio: " & mRec!fechainic
         Label3.Visible = True
      Else
         mRec.Close
         Label2(5).Visible = False
         Text1(2).Visible = False
         Set mRec = mObj.oTabla("movilturno", "where CodMovil = '" & nMov & "' AND (KmFinal IS NOT NULL OR KmFinal <> '') ORDER BY FechaInic DESC")
         If Not mRec.EOF Then
            Text1(0).Text = mRec!KmFinal
         End If
      End If
      mRec.Close
   End If
   Set mRec = Nothing
   Set mObj = Nothing
End Sub

Private Sub Flex1_DblClick()
   With Flex1
      If Flex1.Row > 0 And Trim(Flex1.TextMatrix(Flex1.Row, 0)) <> "" Then
         Command2(0).Tag = .TextMatrix(.Row, 0)    'codmovil
         Command2(1).Tag = .Row                    'fila
         Command2(0).Visible = True
         Command2(1).Caption = "Cancelar"
         Text2(0).Enabled = True
         Text2(0).Text = .TextMatrix(.Row, 2)
         If Trim(.TextMatrix(.Row, 3)) <> "" Then
            Text2(1).Enabled = True
            Text2(1).Text = .TextMatrix(.Row, 3)
         End If
         Flex1.Enabled = False
      End If
   End With
End Sub

Private Sub Label5_Click()
Dim mI As Integer
   sMsgEspere Me, "Buscando datos...", True
   sInitFlex
   Frame1.Left = -10000
   Frame3.Left = -50
   sMsgEspere Me, "", False
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 27 Then
      Combo1(Index).ListIndex = -1
   End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   If Index = 0 Or Index = 2 Then
      KeyAscii = fNumeroKeyPress(KeyAscii)
   Else
      KeyAscii = fAlfaNumSimbKeyPress(KeyAscii)
      KeyAscii = fUcaseKeyPress(KeyAscii)
   End If
End Sub

Sub sClearAll()
   Combo1(1).ListIndex = -1
   Combo1(2).ListIndex = -1
   Combo1(3).ListIndex = -1
   Combo1(4).ListIndex = -1
   Text1(0).Text = ""
   Text1(1).Text = ""
   Text1(2).Text = ""
   Text1(2).Visible = False
   Label2(5).Visible = False
   Combo1(1).Enabled = True
   Combo1(2).Enabled = True
   Combo1(3).Enabled = True
   Combo1(4).Enabled = True
   Text1(0).Enabled = True
   Text1(1).Enabled = True
   Label3.Visible = False
End Sub

Private Function fKilometros(ByRef pKm2 As Long, ByVal pKm1 As Long) As Boolean
Dim mObj As New clRNov
Dim mRec As New ADODB.Recordset
   pKm2 = pKm1
   fKilometros = True
   Set mRec = mObj.oTabla("movilturno", "where CodMovil = '" & Right(Combo1(0).Text, 4) & "' AND KmFinal <> '' order by Fechainic desc")
   If Not mRec.EOF Then
      If mRec!KmFinal > Text1(0).Text Then
         fKilometros = False
         MsgBox "Último Km Final para Móvil " & Left(mRec!CodMovil & mRec!KmFinal, 4) & " fue " & Trim(Mid(mRec!CodMovil & mRec!KmFinal, 5, 12)) & Chr(13) & Chr(10) & "Ingrese Km Inicial Mayor o Igual a " & Trim(Mid(mRec!CodMovil & mRec!KmFinal, 5, 12)), vbCritical, sMessage & " - Error en Kilometro Inicial!!"
      End If
      pKm2 = mRec!KmFinal
      If (pKm1 - pKm2) > 600 Then
         If MsgBox("Según los Kilometrajes el móvil Recorrió: " & (pKm1 - pKm2) & " Km. Es Correcto este Dato?", vbYesNo, sMessage) = vbNo Then
            fKilometros = False
         End If
      End If
   End If
   mRec.Close
   Set mRec = Nothing
   Set mObj = Nothing
End Function

Private Sub sInitFlex()
Dim mObj As New clRNov
Dim mRec As New ADODB.Recordset
Dim mRec2 As New ADODB.Recordset
Dim mChofer As String
Dim mI As Integer

   sBorraFlexDatos Flex1
   With Flex1
      .ColWidth(0) = 900
      .ColWidth(1) = 1900
      .ColWidth(2) = 1000
      .ColWidth(3) = 1000
      .ColWidth(4) = 3500
      .ColWidth(5) = 2500
      .ColWidth(6) = 2200
      .TextMatrix(0, 0) = "Móvil"
      .TextMatrix(0, 1) = "Fecha y hora"
      .TextMatrix(0, 2) = "Km.Inicial"
      .TextMatrix(0, 3) = "Km.Final"
      .TextMatrix(0, 4) = "Patrullero / chofer"
      .TextMatrix(0, 5) = "Policía"
      .TextMatrix(0, 6) = "Turno"
      .Row = 0
      For mI = 0 To .Cols - 1
         .Col = mI
         .CellFontBold = True
      Next
   End With
   Combo2.Clear
   Combo2.AddItem "Todos..."
   Set mRec = mObj.oMovilesPATGRU
   Do While Not mRec.EOF
      Combo2.AddItem mRec!Codigo
      Set mRec2 = mObj.oMovilTurnos(mRec!Codigo, 1)
      If Not mRec2.EOF Then
         mChofer = mRec2!chofer
         If mRec!codtipomov = "PAT" Then
            mChofer = NVL(mRec2!nombre, "")
         End If
         Flex1.AddItem mRec!Codigo & vbTab & mRec2!fechainic & vbTab & mRec2!KmInicial & vbTab & mRec2!KmFinal & vbTab _
            & mChofer & vbTab & mRec2!police & vbTab & mRec2!Descripcion
      Else
         Flex1.AddItem mRec!Codigo
      End If
      mRec2.Close
      mRec.MoveNext
   Loop
   mRec.Close
   If Flex1.Rows > 2 Then
      Flex1.RemoveItem 1
   End If
   sSetFlex2Colors Flex1, &HFFFFFF, &HF5F5F5
   Set mRec = Nothing
   Set mObj = Nothing
End Sub

Private Sub sFormEnabled(ByVal pFlag As Boolean)
   RNov1a_frm.Enabled = pFlag
   RNov1b_frm.Enabled = pFlag
   RNov1d_frm.Enabled = pFlag
End Sub
