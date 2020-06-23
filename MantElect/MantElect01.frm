VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form MantElect01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modulo de Solicitudes de Reparaciones Electricas"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   11835
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "La solicitud atiende a un pedido de asistencia:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   15
      Top             =   2020
      Width           =   4575
      Begin VB.OptionButton Option2 
         Caption         =   "Eléctrico."
         Height          =   255
         Left            =   900
         TabIndex        =   5
         Top             =   860
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "De aire acondicionado."
         Height          =   375
         Left            =   900
         TabIndex        =   4
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   2
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   840
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Grabar Solicitud y Enviar Mensaje"
      Enabled         =   0   'False
      Height          =   495
      Index           =   1
      Left            =   8040
      TabIndex        =   8
      Top             =   6200
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1520
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Volver"
      Height          =   375
      Index           =   2
      Left            =   11040
      TabIndex        =   7
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Agregar"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   4200
      TabIndex        =   6
      Top             =   6360
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   1
      Left            =   9240
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1520
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   3360
      MaxLength       =   90
      TabIndex        =   2
      Top             =   1520
      Width           =   5775
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2055
      Left            =   120
      TabIndex        =   9
      Top             =   4000
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   3625
      _Version        =   327680
      Cols            =   5
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Supervisión"
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
      Left            =   480
      TabIndex        =   14
      Top             =   600
      Width           =   1005
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   120
      Y1              =   480
      Y2              =   3640
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   11700
      Y1              =   480
      Y2              =   480
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
      Index           =   0
      Left            =   480
      TabIndex        =   13
      Top             =   1300
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Registro de solicitudes de servicios"
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
      Left            =   3000
      TabIndex        =   12
      Top             =   45
      Width           =   4245
   End
   Begin VB.Line Line8 
      X1              =   11700
      X2              =   11700
      Y1              =   480
      Y2              =   3640
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
      Index           =   2
      Left            =   9240
      TabIndex        =   11
      Top             =   1300
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Descripcion del problema"
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
      Left            =   3600
      TabIndex        =   10
      Top             =   1300
      Width           =   2160
   End
   Begin VB.Line Line9 
      X1              =   120
      X2              =   11700
      Y1              =   3640
      Y2              =   3640
   End
End
Attribute VB_Name = "MantElect01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mObj As New clMantElect
Dim mObjPea As New clPeaje
Dim mRec As New ADODB.Recordset
Dim mObjLuser As New clLogUser
Dim cboListIndex As Integer


Private Sub Combo1_Click(Index As Integer)
Dim mi As Integer
Select Case Index
      Case 2
         If cboListIndex <> Combo1(2).ListIndex Then
         sLlenoEdificios
            If cboListIndex <> -1 Then
               If (MSFlexGrid1.Rows >= 2 And MSFlexGrid1.TextMatrix(1, 1) <> "") Then
                  If MsgBox("Si selecciona otra Supervisión se eliminarán las solicitudes cargadas. ¿ Desea continuar ? ", vbYesNo, "Cambio de Supervisión") = vbYes Then
   
                     For mi = MSFlexGrid1.Rows - 2 To 1 Step -1
                        MSFlexGrid1.RemoveItem mi
                     Next
                     MSFlexGrid1.TextMatrix(1, 1) = ""
                     MSFlexGrid1.TextMatrix(1, 2) = ""
                     MSFlexGrid1.TextMatrix(1, 3) = ""
                     Command1(1).Enabled = False
                  
                  Else
                     Combo1(2).ListIndex = cboListIndex
                  End If
               End If
               cboListIndex = Combo1(2).ListIndex
            Else
               cboListIndex = Combo1(2).ListIndex
            End If
         End If
End Select
End Sub

Private Sub Command1_Click(Index As Integer)
Dim mi As Integer
Dim mErrMail As Integer
Dim mTextoMail As String
Dim mTextoMailAire As String
Dim mNroParte As String
Dim mListaDestinatarios As String
Dim mAireAcond As String
Dim mIntAireAcond As Integer

If Index = 0 Then
   If fValida Then
      
      'MsgBox (Option1.Value)
      
      'Si el OptionButton de Aire Acondicionado esta seleccionado entonces mAireAcond = "Si" caso contrario "No"
      
      If Option1.Value Then
        mAireAcond = "Si"
      Else
        mAireAcond = "No"
      End If
      
      
      MSFlexGrid1.AddItem vbTab & Combo1(0).Text & vbTab & Text1(0).Text & vbTab & Combo1(1).Text & vbTab & mAireAcond
      Command1(1).Enabled = True
      If MSFlexGrid1.TextMatrix(1, 1) = "" Then
         MSFlexGrid1.RemoveItem 1
      End If
   End If
   
   Option1.Value = False
   Option2.Value = False
  
Else
   If Index = 1 Then
      mErrMail = 0
      mNroParte = mObj.ObtMaxParte
      'mTextoMail = vbCrLf
      
      For mi = 1 To MSFlexGrid1.Rows - 1
         
         If MSFlexGrid1.TextMatrix(mi, 4) = "Si" Then
            mTextoMailAire = mTextoMailAire & "Parte " & (mNroParte + mi) & ": " & MSFlexGrid1.TextMatrix(mi, 1) & vbTab & " // " & MSFlexGrid1.TextMatrix(mi, 2) & vbTab & " // " & MSFlexGrid1.TextMatrix(mi, 3) & vbCrLf
            mIntAireAcond = 1
         Else
            mTextoMail = mTextoMail & "Parte " & (mNroParte + mi) & ": " & MSFlexGrid1.TextMatrix(mi, 1) & vbTab & " // " & MSFlexGrid1.TextMatrix(mi, 2) & vbTab & " // " & MSFlexGrid1.TextMatrix(mi, 3) & vbCrLf
            mIntAireAcond = 0
         End If
         
         'mTextoMail = mTextoMail & "Parte " & (mNroParte + mI) & ": " & MSFlexGrid1.TextMatrix(mI, 1) & vbTab & " // " & MSFlexGrid1.TextMatrix(mI, 2) & vbTab & " // " & MSFlexGrid1.TextMatrix(mI, 3) & vbCrLf
         'mTextoMail = mTextoMail & Format(mI, "00") & ") " & MSFlexGrid1.TextMatrix(mI, 1) & vbTab & " // " & MSFlexGrid1.TextMatrix(mI, 2) & vbTab & " // " & MSFlexGrid1.TextMatrix(mI, 3) & vbCrLf
         mObj.InsRegistros mi + mNroParte, Format(Now, "yyyy-mm-dd hh:mm:ss"), MSFlexGrid1.TextMatrix(mi, 1), MSFlexGrid1.TextMatrix(mi, 2), MSFlexGrid1.TextMatrix(mi, 3), Trim(Right(Combo1(2).Text, 10)), "G", Trim(Right(MDI.mUser, 20)), mIntAireAcond
        'mObj.InsRegistros mi + mNroParte, Format(Now, "yyyy-mm-dd hh:mm:ss"), MSFlexGrid1.TextMatrix(mi, 1), MSFlexGrid1.TextMatrix(mi, 2), MSFlexGrid1.TextMatrix(mi, 3), Left(mObjLuser.sCampoDescrip("USUARIOS", "CodUsuario = '" & Trim(Right(MDI.mUser, 20)) & "'", 2), InStr(1, mObjLuser.sCampoDescrip("USUARIOS", "CodUsuario = '" & Trim(Right(MDI.mUser, 20)) & "'", 2), "@") - 1), "G", Trim(Right(MDI.mUser, 20))
      Next
        
      If mTextoMailAire <> "" Then
        Set mRec = mObj.oEjecutarSelect("SELECT DISTINCT Email FROM MailsxSuperv WHERE CodSuperv = '" & Trim(Right(Combo1(2).Text, 10)) & "' And FechaBaja IS NULL UNION SELECT DISTINCT Email FROM MailsxElectrico WHERE SectorAire = 1 and FechaBaja IS NULL ")
        If Not mRec.EOF Then
            'mListaDestinatarios = ""
            Do While Not mRec.EOF
                mListaDestinatarios = mListaDestinatarios & mRec!Email & ";"
                mRec.MoveNext
            Loop
            'If Not fEnviar_Mail_CDO("", mListaDestinatarios, "ausolmail@ausol.com.ar", " Solicitud de Servicios", "Se ha generado el parte " & Text1(0).Text & ", solicitando los servicios de personal de Mant. Eléctrico, segun detalle:" & vbCrLf & mTextoMail, "", "587", "ausolmail@ausol.com.ar", "sgedosmildiecisiete1$", True, False) Then
            If Not fEnviar_Mail_CDO("", mListaDestinatarios, "ausolmail@ausol.com.ar", " Solicitud de Servicios", "Se ha realizado una solicitud de servicio de personal de Mant. Eléctrico, según detalle:" & vbCrLf & mTextoMailAire, "", "587", "ausolmail@ausol.com.ar", "sgedosmildiecisiete1$", True, False) Then
                mErrMail = mErrMail + 1
            End If
        End If
      End If
                  
      mListaDestinatarios = ""
                  
      If mTextoMail <> "" Then
        Set mRec = mObj.oEjecutarSelect("SELECT DISTINCT Email FROM MailsxSuperv WHERE CodSuperv = '" & Trim(Right(Combo1(2).Text, 10)) & "' And FechaBaja IS NULL UNION SELECT DISTINCT Email FROM MailsxElectrico WHERE SectorAire = 0 and FechaBaja IS NULL ")
        If Not mRec.EOF Then
          
            Do While Not mRec.EOF
                mListaDestinatarios = mListaDestinatarios & mRec!Email & ";"
                mRec.MoveNext
            Loop
            'If Not fEnviar_Mail_CDO("", mListaDestinatarios, "ausolmail@ausol.com.ar", " Solicitud de Servicios", "Se ha generado el parte " & Text1(0).Text & ", solicitando los servicios de personal de Mant. Eléctrico, segun detalle:" & vbCrLf & mTextoMail, "", "587", "ausolmail@ausol.com.ar", "sgedosmildiecisiete1$", True, False) Then
            If Not fEnviar_Mail_CDO("", mListaDestinatarios, "ausolmail@ausol.com.ar", " Solicitud de Servicios", "Se ha realizado una solicitud de servicio de personal de Mant. Eléctrico, según detalle:" & vbCrLf & mTextoMail, "", "587", "ausolmail@ausol.com.ar", "sgedosmildiecisiete1$", True, False) Then
                mErrMail = mErrMail + 1
            End If
        End If
      End If
      
      If mErrMail = 0 Then
        MsgBox "Se ha grabado la solicitud y enviado el mensaje correctamente a los responsables.", vbInformation, "Atención"
      Else
        MsgBox "Se ha grabado la solicitud correctamente, pero se NO ha enviado el correo correctamente", vbExclamation, "Atención"
      End If
   End If
   Unload Me
End If
End Sub

Private Sub Form_Load()
Dim mi As Integer
Dim mTramo As String

MantElect01.Top = 100
MantElect01.Left = (MDI.Width - MantElect01.Width) / 2



Set mRec = mObj.oEjecutarSelect(" SELECT S.CodSuperv, S.Descripcion " & _
                                      "FROM Usuarios_Supervision U " & _
                                         "Inner Join " & _
                                   "Supervisiones S ON S.CodSuperv = U.CodSuperv " & _
                                   "WHERE codusuario = '" & Trim(Right(MDI.mUser, 20)) & "'")
                                 
                                 
   Do While Not mRec.EOF
     Combo1(2).AddItem mRec!descripcion & Space(50) & mRec!CodSuperv
     mRec.MoveNext
   Loop
                                 
                                 
                                 
                                 
                                 
'   If Not mRec.EOF Then
'   Do While Not mRec.EOF
'      'Combo1(0).AddItem mRec!Tramo & " - " & mRec!descripcion
'      MsgBox (mRec!descripcion)
'      mRec.MoveNext
'   Loop
'End If
                                 


'Veo que tramos debe mostrar el combo de Edificios
'mTramo = ""
'Select Case Left(mObjLuser.sCampoDescrip("USUARIOS", "CodUsuario = '" & Trim(Right(MDI.mUser, 20)) & "'", 3), InStr(1, mObjLuser.sCampoDescrip("USUARIOS", "CodUsuario = '" & Trim(Right(MDI.mUser, 20)) & "'", 3), "@") - 1)
''   Case "svergaraa", "svergarad"
''      mTramo = "T1"
''   Case "ssantarosa"
''      mTramo = "T2"
''   Case "situzaingo", "sdecalada"
''      mTramo = "IT"
''   Case "slujan"
''      mTramo = "T3"
''End Select
'
'
'   Case "tesotigre", "suptigre", "marasc1", "realasc"
'      mTramo = "Z1"
'   Case "tescampana", "supcampana", "supcampanadecalado", "belgrano", "197desc1", "ruta197desc", "mpacheco"
'      mTramo = "Z2"
'   Case "tespilar", "suppilar", "suppilardec", "bayre", "buenayredesc", "r202a", "r202"
'      mTramo = "Z3"
'End Select




'Set mRec = mObj.oEjecutarSelect("SELECT * FROM Edificios WHERE FechaBaja IS NULL " & IIf(mTramo <> "", " And Tramo = '" & mTramo & "'", ""))
'If Not mRec.EOF Then
'   Do While Not mRec.EOF
'      Combo1(0).AddItem mRec!Tramo & " - " & mRec!descripcion
'      mRec.MoveNext
'   Loop
'End If
mRec.Close
Combo1(1).AddItem "Alta"
Combo1(1).AddItem "Media"
Combo1(1).AddItem "Baja"
MSFlexGrid1.ColWidth(0) = 200
MSFlexGrid1.ColWidth(1) = 1800
MSFlexGrid1.ColWidth(2) = 6900
MSFlexGrid1.ColWidth(3) = 700
MSFlexGrid1.ColWidth(4) = 1600
MSFlexGrid1.TextMatrix(0, 1) = "Lugar"
MSFlexGrid1.TextMatrix(0, 2) = "Descripcion del Problema"
MSFlexGrid1.TextMatrix(0, 3) = "Prioridad"
MSFlexGrid1.TextMatrix(0, 4) = "Aire acondicionado"
For mi = 0 To MSFlexGrid1.Cols - 1
   MSFlexGrid1.ColAlignment(mi) = 0
Next

cboListIndex = Combo1(2).ListIndex

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set mObj = Nothing
Set mRec = Nothing
Set mObjLuser = Nothing
ShowMenu 47, True, False
End Sub

Private Function fValida() As Boolean
Dim mRet As Boolean
Dim mi As Integer
Dim mRepe As Boolean
mRepe = False
mRet = (Text1(0).Text <> "")
If mRet Then
   For mi = 1 To Combo1.UBound
      If mRet Then
         mRet = (Combo1(mi).Text <> "")
      End If
   Next
End If

If mRet Then
    If Not (Option1.Value Or Option2.Value) Then
        mRet = False
    End If
End If


If mRet Then
   If MSFlexGrid1.Rows > 1 Then
      For mi = 1 To MSFlexGrid1.Rows - 1
         If MSFlexGrid1.TextMatrix(mi, 1) = Combo1(0).Text And MSFlexGrid1.TextMatrix(mi, 2) = Text1(0).Text And mRet Then
            mRet = False
            mRepe = True
         End If
      Next
   End If
End If




If Not mRet Then
   If mRepe Then
      MsgBox "Ya está ingresada dicha solicitud", vbCritical, "Atención"
   Else
      MsgBox "Verifique que todos los datos estén ingresados correctamente", vbCritical, "Atención"
   End If
End If
fValida = mRet
End Function

Private Sub MSFlexGrid1_DblClick()
Dim mi As Integer
Dim mj As Integer
If MSFlexGrid1.Row > 0 And MSFlexGrid1.TextMatrix(1, 1) <> "" And MSFlexGrid1.Col = 1 Then
   If MsgBox("¿Está Seguro de Eliminar este Registro?", vbYesNo, sMessage) = vbYes Then
      If MSFlexGrid1.Rows > 2 Then
         For mi = MSFlexGrid1.Row To MSFlexGrid1.Rows - 2
            For mj = 1 To MSFlexGrid1.Cols - 1
               MSFlexGrid1.TextMatrix(mi, mj) = MSFlexGrid1.TextMatrix(mi + 1, mj)
            Next
         Next
         MSFlexGrid1.RemoveItem (MSFlexGrid1.Rows - 1)
      Else
         MSFlexGrid1.AddItem ""
         MSFlexGrid1.RemoveItem 1
      End If
   End If
End If
End Sub



Private Sub sLlenoEdificios()
Dim mCodSupervision As String
Dim mObj As New clMantElect
Dim mRec As New ADODB.Recordset
   
   mCodSupervision = Trim(Right(Combo1(2).Text, 10))
   Combo1(0).Clear
   Set mRec = mObj.oEjecutarSelect("SELECT * FROM Edificios WHERE CodSuperv = '" & mCodSupervision & "' order by 2")
   Do While Not mRec.EOF
     'Combo1(0).AddItem mRec!descripcion & Space(60) & mRec!Codigo
      Combo1(0).AddItem mRec!Tramo & " - " & mRec!descripcion
     mRec.MoveNext
   Loop
   mRec.Close
   Set mObj = Nothing
   Set mRec = Nothing
End Sub
