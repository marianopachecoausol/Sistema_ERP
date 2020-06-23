VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form MantElect03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Formulario de Registro de Validaciones de trabajos realizados"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   18495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   18495
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   7080
      MaxLength       =   150
      TabIndex        =   4
      Top             =   960
      Width           =   5895
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   4560
      TabIndex        =   3
      Top             =   960
      Width           =   495
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Index           =   0
      Left            =   7560
      TabIndex        =   1
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Volver"
      Height          =   375
      Index           =   1
      Left            =   9840
      TabIndex        =   0
      Top             =   7440
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5340
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   18315
      _ExtentX        =   32306
      _ExtentY        =   9419
      _Version        =   327680
      Cols            =   10
   End
   Begin VB.Line Line1 
      Index           =   10
      X1              =   4320
      X2              =   13200
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label1 
      Caption         =   "Registro de Validaciones de trabajos realizados"
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
      TabIndex        =   9
      Top             =   120
      Width           =   5745
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
      Left            =   4575
      TabIndex        =   8
      Top             =   720
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Validación"
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
      Left            =   5625
      TabIndex        =   7
      Top             =   720
      Width           =   900
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
      Index           =   3
      Left            =   7200
      TabIndex        =   6
      Top             =   720
      Width           =   1275
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   4320
      X2              =   13200
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   4320
      X2              =   4320
      Y1              =   600
      Y2              =   1440
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   5280
      X2              =   5280
      Y1              =   600
      Y2              =   1440
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   6840
      X2              =   6840
      Y1              =   600
      Y2              =   1440
   End
   Begin VB.Line Line1 
      Index           =   9
      X1              =   13200
      X2              =   13200
      Y1              =   600
      Y2              =   1440
   End
End
Attribute VB_Name = "MantElect03"
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

Private Sub Command1_Click(Index As Integer)
Dim mEstado As String
Dim mFecVal As String
Dim mErrMail As Integer
Dim mListaDestinatarios As String
Dim mSectorAire As String


If Index = 0 Then
   If fValida Then
      If MsgBox("¿Está Seguro de Validar este trabajo?", vbYesNo, sMessage) = vbYes Then
         mEstado = Left(Combo1(0).Text, 1)
         mFecVal = Now
         mSectorAire = IIf(MSFlexGrid1.TextMatrix(mRenglon, 9) = "Si", "1", "0")
         
         'Completo el FlexGrid
         MSFlexGrid1.TextMatrix(mRenglon, 7) = Text1(1).Text
         MSFlexGrid1.TextMatrix(mRenglon, 8) = Left(Combo1(0).Text, 1)
         
         
         'Actualizo en Registros
         'mObj.UpdValReg Left(Combo1(0).Text, 1), Left(mObjLuser.sCampoDescrip("USUARIOS", "CodUsuario = '" & Trim(Right(MDI.mUser, 20)) & "'", 6), InStr(1, mObjLuser.sCampoDescrip("USUARIOS", "CodUsuario = '" & Trim(Right(MDI.mUser, 20)) & "'", 6), "@") - 1), Text1(1).Text, Text1(0).Text
         mObj.UpdValReg Left(Combo1(0).Text, 1), Trim(Right(MDI.mUser, 20)), Text1(1).Text, Text1(0).Text

         mErrMail = 0
         
         'Set mRec = mObj.oEjecutarSelect("SELECT * FROM MailsxSuperv WHERE CodSuperv = '" & Left(mObjLuser.sCampoDescrip("USUARIOS", "CodUsuario = '" & Trim(Right(MDI.mUser, 20)) & "'", 6), InStr(1, mObjLuser.sCampoDescrip("USUARIOS", "CodUsuario = '" & Trim(Right(MDI.mUser, 20)) & "'", 6), "@") - 1) & "' And FechaBaja IS NULL")
         Set mRec = mObj.oEjecutarSelect("SELECT DISTINCT Email FROM MailsxSuperv WHERE CodSuperv = '" & mObj.ObtCodSuperv(Text1(0).Text) & "' And FechaBaja IS NULL UNION SELECT DISTINCT Email FROM MailsxElectrico WHERE SectorAire = " & mSectorAire & " AND FechaBaja IS NULL ")
         If Not mRec.EOF Then
            Do While Not mRec.EOF
               mListaDestinatarios = mListaDestinatarios & mRec!Email & ";"
               mRec.MoveNext
            Loop
            If Not fEnviar_Mail_CDO("", mListaDestinatarios, "ausolmail@ausol.com.ar", "Validación de Tareas", "Se " & IIf(Left(Combo1(0).Text, 1) = "A", "ACEPTÓ", "RECHAZÓ") & " la tarea " & Text1(0).Text & "." & vbCrLf & "Observaciones: " & Text1(1).Text, "", "587", "ausolmail@ausol.com.ar", "sgedosmildiecisiete1$", True, False) Then
               mErrMail = mErrMail + 1
            End If
         End If
         If mErrMail = 0 Then
            MsgBox "Se ha validado la solicitud y enviado el mensaje correctamente a los responsables.", vbInformation, "Atención"
         Else
            MsgBox "Se ha validado la solicitud correctamente, pero se NO ha enviado el correo correctamente", vbExclamation, "Atención"
         End If
      End If
   End If
Else
   Unload Me
End If
End Sub

Private Sub Form_Load()
Dim mi As Integer
MantElect03.Top = 100
MantElect03.Left = (MDI.Width - MantElect03.Width) / 2

Combo1(0).AddItem "Aceptar"
Combo1(0).AddItem "Rechazar"

MSFlexGrid1.ColWidth(0) = 200
MSFlexGrid1.ColWidth(1) = 500
MSFlexGrid1.ColWidth(2) = 1700
MSFlexGrid1.ColWidth(3) = 3000
MSFlexGrid1.ColWidth(4) = 4000
MSFlexGrid1.ColWidth(5) = 1700
MSFlexGrid1.ColWidth(6) = 4000
MSFlexGrid1.ColWidth(7) = 4000
MSFlexGrid1.ColWidth(8) = 400
MSFlexGrid1.ColWidth(9) = 0

For mi = 0 To MSFlexGrid1.Cols - 1
   MSFlexGrid1.ColAlignment(mi) = 2
Next

MSFlexGrid1.TextMatrix(0, 1) = "Parte"
MSFlexGrid1.TextMatrix(0, 2) = "Fecha Solicitud"
MSFlexGrid1.TextMatrix(0, 3) = "Lugar"
MSFlexGrid1.TextMatrix(0, 4) = "Descripcion de la Solicitud"
MSFlexGrid1.TextMatrix(0, 5) = "Fecha Fin Asist."
MSFlexGrid1.TextMatrix(0, 6) = "Segunda Descripcion"
MSFlexGrid1.TextMatrix(0, 7) = "Observaciones"
MSFlexGrid1.TextMatrix(0, 8) = "Est"
MSFlexGrid1.TextMatrix(0, 9) = "Sector Aire"

'Set mRec = mObj.oEjecutarSelect("SELECT * FROM Registros WHERE Estado = 'T' And Origen = 'O' ORDER BY Parte")
Set mRec = mObj.oEjecutarSelect("SELECT * FROM Registros R INNER JOIN (SELECT * FROM Usuarios_Supervision U WHERE codusuario = '" & Trim(Right(MDI.mUser, 20)) & "') U ON R.CodSuperv = U.CodSuperv WHERE Estado = 'T' And Origen = 'O' ORDER BY Parte")
If Not mRec.EOF Then
   mi = 1
   Do While Not mRec.EOF
      mi = mi + 1
      MSFlexGrid1.AddItem ""
      MSFlexGrid1.TextMatrix(mi, 1) = mRec!Parte
      MSFlexGrid1.TextMatrix(mi, 2) = NVL(mRec!FechaSolic, "")
      MSFlexGrid1.TextMatrix(mi, 3) = NVL(mRec!CodEdificio, "")
      MSFlexGrid1.TextMatrix(mi, 4) = NVL(mRec!descripcion, "")
      MSFlexGrid1.TextMatrix(mi, 5) = NVL(mRec!FechaFinAsist, "")
      MSFlexGrid1.TextMatrix(mi, 6) = NVL(mRec!SegundaDesc, "")
      MSFlexGrid1.TextMatrix(mi, 7) = NVL(mRec!ObserValid, "")
      MSFlexGrid1.TextMatrix(mi, 8) = NVL(mRec!estado, "")
      MSFlexGrid1.TextMatrix(mi, 9) = IIf(mRec!SectorAire = 1, "Si", "No")
      mRec.MoveNext
   Loop
   MSFlexGrid1.RemoveItem 1
End If
mRec.Close

Text1(0).Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set mObj = Nothing
Set mRec = Nothing
ShowMenu 47, True, False
End Sub

Private Function fValida() As Boolean
Dim mRet As Boolean
mRet = mRenglon <> 0
If mRet Then
   mRet = (Combo1(0).Text <> "")
   If Not mRet Then
      MsgBox "Seleccione una validación a la tarea", vbCritical, "Atención"
   End If
Else
   MsgBox "Seleccione un pedido de la grilla", vbCritical, "Atención"
End If
fValida = mRet
End Function

Private Sub MSFlexGrid1_Click()
If MSFlexGrid1.MouseCol = 0 And MSFlexGrid1.MouseRow > 0 Then
   mRenglon = MSFlexGrid1.MouseRow
   Text1(0).Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
Else
   mRenglon = 0
End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = fAlfaNumKeyPress(KeyAscii)
End Sub
