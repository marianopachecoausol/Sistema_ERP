VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form MantElect02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Formulario de registro de Reparaciones"
   ClientHeight    =   9540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   21390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9540
   ScaleWidth      =   21390
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   17885
      MaxLength       =   10
      TabIndex        =   8
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Volver"
      Height          =   375
      Index           =   1
      Left            =   12725
      TabIndex        =   17
      Top             =   8880
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Index           =   0
      Left            =   10445
      TabIndex        =   16
      Top             =   8880
      Width           =   1095
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
      Index           =   1
      Left            =   15005
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   840
      Width           =   855
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
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   18965
      MaxLength       =   5
      TabIndex        =   9
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   17165
      MaxLength       =   5
      TabIndex        =   7
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   16085
      MaxLength       =   10
      TabIndex        =   6
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   9720
      MaxLength       =   90
      TabIndex        =   13
      Top             =   1800
      Width           =   5295
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   960
      MaxLength       =   19
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   6120
      MaxLength       =   150
      TabIndex        =   4
      Top             =   840
      Width           =   8660
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   9120
      MaxLength       =   3
      TabIndex        =   12
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   9
      Left            =   15240
      MaxLength       =   5
      TabIndex        =   14
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   10
      Left            =   15960
      MaxLength       =   5
      TabIndex        =   15
      Top             =   1800
      Width           =   615
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6180
      Left            =   120
      TabIndex        =   0
      Top             =   2475
      Width           =   21200
      _ExtentX        =   37386
      _ExtentY        =   10901
      _Version        =   327680
      Cols            =   16
   End
   Begin VB.Line Line1 
      Index           =   16
      X1              =   15840
      X2              =   15840
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
      Index           =   14
      X1              =   9600
      X2              =   9600
      Y1              =   1440
      Y2              =   2280
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
      Left            =   17405
      TabIndex        =   32
      Top             =   520
      Width           =   885
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
      Left            =   17915
      TabIndex        =   31
      Top             =   840
      Width           =   1470
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
      TabIndex        =   30
      Top             =   1560
      Width           =   915
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
      TabIndex        =   29
      Top             =   1560
      Width           =   525
   End
   Begin VB.Line Line1 
      Index           =   7
      X1              =   15965
      X2              =   15965
      Y1              =   480
      Y2              =   1440
   End
   Begin VB.Line Line1 
      Index           =   9
      X1              =   19565
      X2              =   19565
      Y1              =   480
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
      Index           =   12
      X1              =   120
      X2              =   19565
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      Index           =   6
      X1              =   14885
      X2              =   14885
      Y1              =   480
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
      Index           =   4
      X1              =   17765
      X2              =   17765
      Y1              =   720
      Y2              =   1440
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   840
      X2              =   840
      Y1              =   480
      Y2              =   1440
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   120
      X2              =   120
      Y1              =   480
      Y2              =   2280
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   120
      X2              =   19565
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   2760
      X2              =   2760
      Y1              =   480
      Y2              =   1440
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
      TabIndex        =   28
      Top             =   1560
      Width           =   465
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
      TabIndex        =   27
      Top             =   1560
      Width           =   405
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
      TabIndex        =   26
      Top             =   1560
      Width           =   1830
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
      Left            =   16055
      TabIndex        =   25
      Top             =   840
      Width           =   1680
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
      Left            =   15035
      TabIndex        =   24
      Top             =   600
      Width           =   765
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
      Index           =   3
      Left            =   6120
      TabIndex        =   23
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Lugarx"
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
      TabIndex        =   22
      Top             =   600
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha / Hora Solicit."
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
      TabIndex        =   21
      Top             =   600
      Width           =   1800
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
      TabIndex        =   20
      Top             =   600
      Width           =   465
   End
   Begin VB.Label Label1 
      Caption         =   "Registro de Reparaciones de Trabajo"
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
      Left            =   6960
      TabIndex        =   19
      Top             =   120
      Width           =   4515
   End
   Begin VB.Line Line1 
      Index           =   8
      X1              =   3120
      X2              =   3120
      Y1              =   1440
      Y2              =   2280
   End
   Begin VB.Line Line1 
      Index           =   10
      X1              =   120
      X2              =   19565
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      Index           =   11
      X1              =   15965
      X2              =   19565
      Y1              =   720
      Y2              =   720
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
      Left            =   16000
      TabIndex        =   18
      Top             =   1560
      Width           =   510
   End
End
Attribute VB_Name = "MantElect02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mObj As New clMantElect
Dim mObjPea As New clPeaje
Dim mRec As New ADODB.Recordset
Dim mRenglon As Integer
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
      Text1(7).Text = mRec!Unidad
      mRec.Close
End Select
End Sub

Private Sub Command1_Click(Index As Integer)
Dim mEstado As String
Dim mFecPro As String
Dim mFecTer As String
Dim mOkGrb As Boolean
Dim mEstadoAnt As String
Dim mTextoMail As String
Dim mErrMail As Integer
Dim mListaDestinatarios As String
Dim mSectorAire As String

If Index = 0 Then
   If fValida1 Then
      If MsgBox("¿Está Seguro de Grabar esta Orden?", vbYesNo, sMessage) = vbYes Then
         mEstadoAnt = MSFlexGrid1.TextMatrix(mRenglon, 14)
         mEstado = "P"
         mFecPro = Now
         mFecTer = ""
         mOkGrb = True
         If MsgBox("¿Está terminado el trabajo?", vbYesNo, sMessage) = vbYes Then
            mOkGrb = False
            mEstado = "T"
            mFecTer = IIf(MSFlexGrid1.TextMatrix(mRenglon, 14) = "G", mFecPro, Now)
            If fValida2 Then
               mOkGrb = True
            End If
            If mOkGrb Then
               If DateDiff("n", CDate(Text1(1).Text), CDate(Text1(3).Text & " " & Text1(4).Text & ":00")) <= 0 Then
                  mOkGrb = False
                  MsgBox "Verificar la fecha de Asistencia", vbCritical, "Atención"
               End If
            End If
         End If

         If mOkGrb Then
            'Completo el FlexGrid
            MSFlexGrid1.TextMatrix(mRenglon, 6) = Text1(3).Text & " " & Text1(4).Text & ":00"
            MSFlexGrid1.TextMatrix(mRenglon, 7) = IIf(Text1(5).Text <> "", Text1(5).Text & " " & Text1(6).Text & ":00", "")
            MSFlexGrid1.TextMatrix(mRenglon, 8) = Text1(8).Text
            MSFlexGrid1.TextMatrix(mRenglon, 9) = Combo1(2).Text
            MSFlexGrid1.TextMatrix(mRenglon, 10) = Combo1(3).Text
            MSFlexGrid1.TextMatrix(mRenglon, 11) = Text1(7).Text
            MSFlexGrid1.TextMatrix(mRenglon, 12) = Text1(9).Text
            MSFlexGrid1.TextMatrix(mRenglon, 13) = Text1(10).Text
            MSFlexGrid1.TextMatrix(mRenglon, 14) = mEstado

            mSectorAire = IIf(MSFlexGrid1.TextMatrix(mRenglon, 15) = "Si", "1", "0")
            
            'Actualizo en Registros
            mObj.UpdRegistros Text1(3).Text & " " & Text1(4).Text & ":00", IIf(Text1(5).Text <> "", Text1(5).Text & " " & Text1(6).Text & ":00", ""), Text1(8).Text, Left(Combo1(2).Text, 8), Left(Combo1(3).Text, 6), IIf(Text1(9).Text <> "", Text1(9).Text, ""), IIf(Text1(10).Text <> "", Text1(10).Text, ""), mEstado, IIf(mEstadoAnt = "G", Trim(Right(MDI.mUser, 20)), ""), IIf(mEstadoAnt = "G", mFecPro, ""), IIf(mEstado = "T", Trim(Right(MDI.mUser, 20)), ""), IIf(mFecTer <> "", mFecTer, ""), Text1(0).Text
            If mEstado = "T" Then
               mErrMail = 0
               
               mTextoMail = vbCrLf & "Se ha resuelto el Parte  " & Text1(0).Text & " de Mantenimiento Eléctrico: " & vbCrLf & vbCrLf & "     Descripción de la solicitud:  " & Text1(2).Text & vbCrLf & vbCrLf & "Verifique el servicio realizado. Gracias"
               Set mRec = mObj.oEjecutarSelect("SELECT DISTINCT Email  FROM MailsxSuperv WHERE CodSuperv = '" & mObj.ObtCodSuperv(Text1(0).Text) & "' And FechaBaja IS NULL UNION SELECT DISTINCT Email FROM MailsxElectrico WHERE SectorAire = " & mSectorAire & " AND FechaBaja IS NULL ")
               'Set mRec = mObj.oEjecutarSelect("SELECT * FROM MailsxSuperv WHERE CodSuperv = '" & mObj.ObtCodSuperv(Text1(0).Text) & "' And Email <> '" & mObjLuser.sCampoDescrip("USUARIOS", "CodUsuario = '" & Trim(Right(MDI.mUser, 20)) & "'", 6) & "' And FechaBaja IS NULL")
 
               If Not mRec.EOF Then
                  mListaDestinatarios = ""
                  Do While Not mRec.EOF
                     mListaDestinatarios = mListaDestinatarios & mRec!Email & ";"
                     mRec.MoveNext
                  Loop
                  If Not fEnviar_Mail_CDO("", mListaDestinatarios, "ausolmail@ausol.com.ar", "Repuesta a Solicitud de Servicios", mTextoMail, "", "587", "ausolmail@ausol.com.ar", "sgedosmildiecisiete1$", True, False) Then
                     mErrMail = mErrMail + 1
                  End If
               End If
               If mErrMail = 0 Then
                  MsgBox "Se ha grabado la solicitud y enviado el mensaje correctamente a los responsables.", vbInformation, "Atención"
               Else
                  MsgBox "Se ha grabado la solicitud correctamente, pero se NO ha enviado el correo correctamente", vbExclamation, "Atención"
               End If
            End If
         End If
      End If
   End If
Else
   Unload Me
End If
End Sub

Private Sub Form_Load()
Dim mi As Integer
MantElect02.Top = 100
MantElect02.Left = (MDI.Width - MantElect02.Width) / 2

Set mRec = mObj.oEjecutarSelect("SELECT * FROM Edificios WHERE FechaBaja IS NULL")
If Not mRec.EOF Then
   Do While Not mRec.EOF
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

MSFlexGrid1.ColWidth(0) = 200
MSFlexGrid1.ColWidth(1) = 500
MSFlexGrid1.ColWidth(2) = 1700
MSFlexGrid1.ColWidth(3) = 3000
MSFlexGrid1.ColWidth(4) = 4000
MSFlexGrid1.ColWidth(5) = 750
MSFlexGrid1.ColWidth(6) = 1700
MSFlexGrid1.ColWidth(7) = 1700
MSFlexGrid1.ColWidth(8) = 4000
MSFlexGrid1.ColWidth(9) = 2200
MSFlexGrid1.ColWidth(10) = 4000
MSFlexGrid1.ColWidth(11) = 500
MSFlexGrid1.ColWidth(12) = 500
MSFlexGrid1.ColWidth(13) = 600
MSFlexGrid1.ColWidth(14) = 400
MSFlexGrid1.ColWidth(15) = 0

For mi = 0 To MSFlexGrid1.Cols - 1
   MSFlexGrid1.ColAlignment(mi) = 2
Next

MSFlexGrid1.TextMatrix(0, 1) = "Parte"
MSFlexGrid1.TextMatrix(0, 2) = "Fecha Solicitud"
MSFlexGrid1.TextMatrix(0, 3) = "Lugar"
MSFlexGrid1.TextMatrix(0, 4) = "Descripcion de la Solicitud"
MSFlexGrid1.TextMatrix(0, 5) = "Prioridad"
MSFlexGrid1.TextMatrix(0, 6) = "Fecha Ini. Asist."
MSFlexGrid1.TextMatrix(0, 7) = "Fecha Fin Asist."
MSFlexGrid1.TextMatrix(0, 8) = "Segunda Descripcion"
MSFlexGrid1.TextMatrix(0, 9) = "Rubro"
MSFlexGrid1.TextMatrix(0, 10) = "Sub Rubro"
MSFlexGrid1.TextMatrix(0, 11) = "Unid."
MSFlexGrid1.TextMatrix(0, 12) = "Cant."
MSFlexGrid1.TextMatrix(0, 13) = "Horas"
MSFlexGrid1.TextMatrix(0, 14) = "Est."
MSFlexGrid1.TextMatrix(0, 15) = "Sector Aire"


'IIf(a > b, "a is Big","b is Big")






'Set mRec = mObj.oEjecutarSelect("SELECT * FROM Registros WHERE Estado NOT IN ('A', 'T') AND Origen = 'O' ORDER BY Parte")
Set mRec = mObj.oEjecutarSelect("SELECT R.* " & _
                                    "FROM Registros R " & _
                                        "Inner Join " & _
                                    "MailsxElectrico M ON R.SectorAire = M.SectorAire " & _
                                "WHERE Estado NOT IN ('A', 'T') AND Origen = 'O' AND M.codusuario = '" & Trim(Right(MDI.mUser, 20)) & "';")

If Not mRec.EOF Then
   mi = 1
   Do While Not mRec.EOF
      mi = mi + 1
      MSFlexGrid1.AddItem ""
      MSFlexGrid1.TextMatrix(mi, 1) = mRec!Parte
      MSFlexGrid1.TextMatrix(mi, 2) = NVL(mRec!FechaSolic, "")
      MSFlexGrid1.TextMatrix(mi, 3) = NVL(mRec!CodEdificio, "")
      MSFlexGrid1.TextMatrix(mi, 4) = NVL(mRec!descripcion, "")
      MSFlexGrid1.TextMatrix(mi, 5) = NVL(mRec!Prioridad, "")
      MSFlexGrid1.TextMatrix(mi, 6) = NVL(mRec!FechaIniAsist, "")
      MSFlexGrid1.TextMatrix(mi, 7) = NVL(mRec!FechaFinAsist, "")
      MSFlexGrid1.TextMatrix(mi, 8) = NVL(mRec!SegundaDesc, "")
      MSFlexGrid1.TextMatrix(mi, 9) = NVL(mRec!Rubro, "") & " - " & mObj.sCampoDescrip("Rubros", "Codigo = '" & mRec!Rubro & "'", 1)
      MSFlexGrid1.TextMatrix(mi, 10) = NVL(mRec!SubRubro, "") & "-" & mObj.sCampoDescrip("SubRubros", "Codigo = '" & mRec!SubRubro & "'", 2)
      MSFlexGrid1.TextMatrix(mi, 11) = mObj.sCampoDescrip("SubRubros", "Codigo = '" & mRec!SubRubro & "'", 3)
      MSFlexGrid1.TextMatrix(mi, 12) = NVL(mRec!Cantidad, "")
      MSFlexGrid1.TextMatrix(mi, 13) = NVL(mRec!Horas, "")
      MSFlexGrid1.TextMatrix(mi, 14) = NVL(mRec!estado, "")
      MSFlexGrid1.TextMatrix(mi, 15) = IIf(mRec!SectorAire = 1, "Si", "No")
      
      mRec.MoveNext
   Loop
   MSFlexGrid1.RemoveItem 1
End If
mRec.Close

Text1(0).Enabled = False
Text1(1).Enabled = False

Text1(2).Enabled = True
Combo1(0).Enabled = False
Combo1(1).Enabled = False
Text1(7).Enabled = False
Text1(10).Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set mObj = Nothing
Set mRec = Nothing
ShowMenu 47, True, False
End Sub

Private Function fValida1() As Boolean
Dim mRet As Boolean
Dim mi As Integer
mRet = mRenglon <> 0
If mRet Then
   mRet = Fecha_ok(Text1(3).Text)
   If mRet Then
      mRet = Hora_ok(Text1(4).Text)
   End If
   If mRet Then
      mRet = DateDiff("s", CDate(Text1(1).Text), CDate(Text1(3).Text & " " & Text1(4).Text & ":00")) > 0
   End If
   If mRet Then
      mRet = (Combo1(2).Text <> "")
   End If
   If mRet Then
      mRet = (Combo1(3).Text <> "")
   End If
   If mRet Then
      mRet = (Text1(8).Text <> "")
   End If
   If Not mRet Then
      MsgBox "Verifique que los datos estén ingresados correctamente", vbCritical, "Atención"
   End If
Else
   MsgBox "Seleccione un pedido de la grilla", vbCritical, "Atención"
End If
fValida1 = mRet
End Function

Private Sub MSFlexGrid1_Click()
Dim mi As Integer
Dim mj As Integer
Dim mFound As Boolean
Dim mHoraIniAsist As String
Dim mHoraFinAsist As String

If MSFlexGrid1.MouseCol = 0 And MSFlexGrid1.MouseRow > 0 Then
   mRenglon = MSFlexGrid1.MouseRow
   Text1(0).Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
   Text1(1).Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
   For mi = 0 To Combo1(0).ListCount - 1
      If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = Combo1(0).List(mi) Then
         Combo1(0).ListIndex = mi
      End If
   Next
   Text1(2).Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4)
   For mi = 0 To Combo1(1).ListCount - 1
      If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5) = Combo1(1).List(mi) Then
         Combo1(1).ListIndex = mi
      End If
   Next
   Text1(3).Text = NVL(Left(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6), 10), "")
   
   
   mHoraIniAsist = NVL(Right(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6), 24), "")
   If mHoraIniAsist <> "" Then
      mHoraIniAsist = Format(mHoraIniAsist, "hh:mm")
   End If
   Text1(4).Text = mHoraIniAsist
   'Text1(4).Text = NVL(Right(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6), 8), "")
   
   
   
   
   Text1(5).Text = NVL(Left(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7), 10), "")
   
   
   mHoraFinAsist = NVL(Right(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7), 24), "")
   If mHoraFinAsist <> "" Then
      mHoraFinAsist = Format(mHoraFinAsist, "hh:mm")
   End If
   Text1(6).Text = mHoraFinAsist
   'Text1(6).Text = NVL(Right(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7), 8), "")
   
   mFound = False
   For mi = 0 To Combo1(2).ListCount - 1
      If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 9) = Combo1(2).List(mi) Then
         mFound = True
         Combo1(2).ListIndex = mi
      End If
   Next
   If Not mFound Then
      Combo1(2).ListIndex = -1
   End If
   
   mFound = False
   For mi = 0 To Combo1(3).ListCount - 1
      If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 10) = Combo1(3).List(mi) Then
         mFound = True
         Combo1(3).ListIndex = mi
      End If
   Next
   If Not mFound Then
      Combo1(3).ListIndex = -1
   End If
   
   Text1(7).Text = NVL(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 11), "")
   Text1(8).Text = NVL(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 8), "")
   Text1(9).Text = NVL(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 12), "")
   Text1(10).Text = NVL(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 13), "")
Else
   mRenglon = 0
End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
   Case 3, 5
      KeyAscii = fDateKeyPress(Text1(Index), KeyAscii)
   Case 4, 6
      KeyAscii = fHoraKeyPress(Text1(Index), KeyAscii)
   Case 8
      KeyAscii = fAlfaNumKeyPress(KeyAscii)
   Case 9, 10
      KeyAscii = fNumDoubleKeyPress(KeyAscii)
End Select
End Sub

Private Function fValida2() As Boolean
Dim mRet As Boolean
Dim mi As Integer
mRet = mRenglon <> 0
If mRet Then
   mRet = Fecha_ok(Text1(5).Text)
   If mRet Then
      mRet = Hora_ok(Text1(6).Text)
   End If
   If mRet Then
      mRet = DateDiff("s", CDate(Text1(3).Text & " " & Text1(4).Text & ":00"), CDate(Text1(5).Text & " " & Text1(6).Text & ":00")) >= 0
   End If
   If mRet Then
      mRet = (Text1(9).Text <> "")
   End If
   If mRet Then
      mRet = (Text1(10).Text <> "")
   End If
   If Not mRet Then
      MsgBox "Verifique que los datos estén ingresados correctamente", vbCritical, "Atención"
   End If
Else
   MsgBox "Seleccione un pedido de la grilla", vbCritical, "Atención"
End If
fValida2 = mRet
End Function

Private Sub Text1_LostFocus(Index As Integer)
Dim mRet As Boolean
Select Case Index
   Case 3, 4, 5, 6
      mRet = (Text1(3).Text <> "" And Text1(4).Text <> "" And Text1(5).Text <> "" And Text1(6).Text <> "")
      If mRet Then
         If DateDiff("s", CDate(Text1(3).Text & " " & Text1(4).Text & ":00"), CDate(Text1(5).Text & " " & Text1(6).Text & ":00")) >= 0 Then
            'Text1(10).Text = Redondeo(DateDiff("n", CDate(Text1(3).Text & " " & Text1(4).Text & ":00"), CDate(Text1(5).Text & " " & Text1(6).Text & ":00")) / 60, 2)
            Text1(10).Text = Replace(Redondeo(DateDiff("n", CDate(Text1(3).Text & " " & Text1(4).Text & ":00"), CDate(Text1(5).Text & " " & Text1(6).Text & ":00")) / 60, 2), ",", ".")
         Else
            MsgBox "Verifique las fechas de Asistencia", vbCritical, "Atención"
            Text1(Index).Text = ""
            Text1(Index).SetFocus
         End If
      End If
End Select
End Sub
