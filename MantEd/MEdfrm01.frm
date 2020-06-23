VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form MEdfrm01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modulo de Solicitudes de Reparaciones Edilicias"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   10215
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   2
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   2400
      MaxLength       =   90
      TabIndex        =   2
      Top             =   1520
      Width           =   6615
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   1
      Left            =   9120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1520
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Agregar"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   3480
      TabIndex        =   4
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Volver"
      Height          =   375
      Index           =   2
      Left            =   5640
      TabIndex        =   5
      Top             =   4800
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1520
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Grabar Solicitud y Enviar Mensaje"
      Enabled         =   0   'False
      Height          =   495
      Index           =   1
      Left            =   8040
      TabIndex        =   6
      Top             =   4680
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1695
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   2990
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
      TabIndex        =   12
      Top             =   600
      Width           =   1005
   End
   Begin VB.Line Line9 
      X1              =   120
      X2              =   10080
      Y1              =   2280
      Y2              =   2280
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
      Left            =   2640
      TabIndex        =   11
      Top             =   1300
      Width           =   2160
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
      Left            =   9120
      TabIndex        =   10
      Top             =   1300
      Width           =   765
   End
   Begin VB.Line Line8 
      X1              =   10080
      X2              =   10080
      Y1              =   480
      Y2              =   2280
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
      TabIndex        =   9
      Top             =   45
      Width           =   4245
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
      TabIndex        =   8
      Top             =   1300
      Width           =   495
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   10080
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   120
      Y1              =   480
      Y2              =   2280
   End
End
Attribute VB_Name = "MEdfrm01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mObj As New clMantEd
'Dim mObjPea As New clPeaje
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

'Private Sub Command1_Click(Index As Integer)
'Dim mI As Integer
'Dim mErrMail As Integer
'Dim mTextoMail As String
'Dim mNroParte As String
'If Index = 0 Then
'   If fValida Then
'      MSFlexGrid1.AddItem vbTab & Combo1(0).Text & vbTab & Text1(0).Text & vbTab & Combo1(1).Text
'      Command1(1).Enabled = True
'      If MSFlexGrid1.TextMatrix(1, 1) = "" Then
'         MSFlexGrid1.RemoveItem 1
'      End If
'      'Limpio los textBoxs
'      For mI = 0 To Text1.Count - 1
'         Text1(mI).Text = ""
'      Next
'      'Limpio los comboBoxs
'      For mI = 0 To Combo1.Count - 1
'         Combo1(mI).ListIndex = -1
'      Next
'
'
'   End If
'Else
'   If Index = 1 Then
'      mNroParte = mObj.ObtMaxParte
'      mTextoMail = vbCrLf
'      For mI = 1 To MSFlexGrid1.Rows - 1
'         mTextoMail = mTextoMail & Format(mI, "00") & ") " & MSFlexGrid1.TextMatrix(mI, 1) & vbTab & " // " & MSFlexGrid1.TextMatrix(mI, 2) & vbTab & " // " & MSFlexGrid1.TextMatrix(mI, 3) & vbCrLf
'         mObj.InsRegistros mI + mNroParte, Format(Now, "yyyy-mm-dd hh:mm:ss"), MSFlexGrid1.TextMatrix(mI, 1), MSFlexGrid1.TextMatrix(mI, 2), MSFlexGrid1.TextMatrix(mI, 3), Left(mObjLuser.sCampoDescrip("USUARIOS", "CodUsuario = '" & Trim(Right(MDI.mUser, 20)) & "'", 3), InStr(1, mObjLuser.sCampoDescrip("USUARIOS", "CodUsuario = '" & Trim(Right(MDI.mUser, 20)) & "'", 3), "@") - 1), "G", Trim(Right(MDI.mUser, 20)), "S", "", "", 0, ""
'      Next
'      Set mRec = mObj.oEjecutarSelect("SELECT * FROM MailsxSuperv WHERE CodSuperv = '" & Left(mObjLuser.sCampoDescrip("USUARIOS", "CodUsuario = '" & Trim(Right(MDI.mUser, 20)) & "'", 3), InStr(1, mObjLuser.sCampoDescrip("USUARIOS", "CodUsuario = '" & Trim(Right(MDI.mUser, 20)) & "'", 3), "@") - 1) & "' And FechaBaja IS NULL")
'      mErrMail = 0
'      If Not mRec.EOF Then
'         Do While Not mRec.EOF
'             If Not fEnviar_Mail_CDO("", mRec!Email, "ausolmail@ausol.com.ar", "Solicitud de Servicios", "Se ha generado el parte " & Text1(0).Text & ", solicitando los servicios de personal de Mant. Edilicio, segun detalle:" & vbCrLf & mTextoMail, "", "587", "system\ausolmail", "sgedosmildiecisiete1$", True, False) Then
'               mErrMail = mErrMail + 1
'            End If
'            mRec.MoveNext
'         Loop
'      End If
'      If mErrMail = 0 Then
'         MsgBox "Se ha grabado la solicitud y enviado el mensaje correctamente a los responsables.", vbInformation, "Atención"
'      Else
'         MsgBox "Se ha grabado la solicitud correctamente, pero NO se ha enviado el correo correctamente", vbExclamation, "Atención"
'      End If
'   End If
'   Unload Me
'End If
'End Sub

Private Sub Command1_Click(Index As Integer)
Dim mi As Integer
Dim mErrMail As Integer
Dim mTextoMail As String

Dim mNroParte As String
Dim mListaDestinatarios As String
Dim mZonaMantEdil As String

mZonaMantEdil = ""


mZonaMantEdil = mObj.sTablaDescr("Edificios", "Descripcion = '" & Combo1(0).Text & "'", 4)


If Index = 0 Then
   If fValida Then
      MSFlexGrid1.AddItem vbTab & Combo1(0).Text & vbTab & Text1(0).Text & vbTab & Combo1(1).Text & vbTab & mZonaMantEdil & " - " & Combo1(0).Text
      Command1(1).Enabled = True
      
      If MSFlexGrid1.TextMatrix(1, 1) = "" Then
         MSFlexGrid1.RemoveItem 1
      End If
   End If
   

Else
   If Index = 1 Then
      mErrMail = 0
      mNroParte = mObj.ObtMaxParte
      'mTextoMail = vbCrLf
      
      For mi = 1 To MSFlexGrid1.Rows - 1
         mTextoMail = mTextoMail & "Parte " & (mNroParte + mi) & ": " & MSFlexGrid1.TextMatrix(mi, 1) & vbTab & " // " & MSFlexGrid1.TextMatrix(mi, 2) & vbTab & " // " & MSFlexGrid1.TextMatrix(mi, 3) & vbCrLf
         'mTextoMail = mTextoMail & "Parte " & (mNroParte + mi) & ": " & Right(MSFlexGrid1.TextMatrix(mi, 1), Len(MSFlexGrid1.TextMatrix(mi, 1)) - 5) & vbTab & " // " & MSFlexGrid1.TextMatrix(mi, 2) & vbTab & " // " & MSFlexGrid1.TextMatrix(mi, 3) & vbCrLf '--Saco codigo de zona
         'mObj.InsRegistros mi + mNroParte, Format(Now, "yyyy-mm-dd hh:mm:ss"), MSFlexGrid1.TextMatrix(mi, 1), MSFlexGrid1.TextMatrix(mi, 2), MSFlexGrid1.TextMatrix(mi, 3), Trim(Right(Combo1(2).Text, 10)), "G", Trim(Right(MDI.mUser, 20)), "S", "", "", 0, "", 0
         mObj.InsRegistros mi + mNroParte, Format(Now, "yyyy-mm-dd hh:mm:ss"), MSFlexGrid1.TextMatrix(mi, 4), MSFlexGrid1.TextMatrix(mi, 2), MSFlexGrid1.TextMatrix(mi, 3), Trim(Right(Combo1(2).Text, 10)), "G", Trim(Right(MDI.mUser, 20)), "S", "", "", 0, "", 0 '-- inserto el edificio del a fila oculta de la grilla
      Next
        
      mListaDestinatarios = ""
                  
      If mTextoMail <> "" Then
       
        Set mRec = mObj.oEjecutarSelect("SELECT DISTINCT Email FROM MailsxSuperv WHERE CodSuperv = '" & Trim(Right(Combo1(2).Text, 10)) & "' And FechaBaja IS NULL UNION SELECT DISTINCT Email FROM MailsxEdilicio WHERE FechaBaja IS NULL ")
        If Not mRec.EOF Then
          
            Do While Not mRec.EOF
               mListaDestinatarios = mListaDestinatarios & mRec!Email & ";"
               mRec.MoveNext
            Loop
            
            If Not fEnviar_Mail_CDO("", mListaDestinatarios, "ausolmail@ausol.com.ar", " MANT. EDILICIO - Solicitud de Servicios", "Se ha realizado una solicitud de servicio de personal de Mant. Edilicio, según detalle:" & vbCrLf & mTextoMail, "", "587", "ausolmail@ausol.com.ar", "sgedosmildiecisiete1$", True, False) Then
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

MEdfrm01.Top = 100
MEdfrm01.Left = (MDI.Width - MEdfrm01.Width) / 2

''Veo que tramos debe mostrar el combo de Edificios
'mTramo = ""
'Select Case Left(mObjLuser.sCampoDescrip("USUARIOS", "CodUsuario = '" & Trim(Right(MDI.mUser, 20)) & "'", 3), InStr(1, mObjLuser.sCampoDescrip("USUARIOS", "CodUsuario = '" & Trim(Right(MDI.mUser, 20)) & "'", 3), "@") - 1)
'   Case "marasc1", "realasc", "suptigre", "testigre"
'      mTramo = "Z1"
'   Case "supcampana", "supcampanadecalado", "tescampana"
'      mTramo = "Z2"
'   Case "suppilar", "suppilardec", "tespilar"
'      mTramo = "Z3"
'   Case "197desc1", "bayre", "belgrano", "buenayredesc", "r202", "ruta197desc", "sup202a"
'      mTramo = "Z4"
'   Case "aakel", "aghelfi", "agomez", "dcardone", "dmartinez", "epinto", "ezambelli", "jgrigorakis", "lbartone", "malbanesi", "mlaplace", "mnavarro", "rfanti", "ssanmartin", "usosa", "mpacheco"
'      mTramo = "Z1','Z2','Z3','Z4"
'End Select



   Set mRec = mObj.oEjecutarSelect(" SELECT S.CodSuperv, S.Descripcion " & _
                                      "FROM Usuarios_Supervision U " & _
                                         "Inner Join " & _
                                   "Supervisiones S ON S.CodSuperv = U.CodSuperv " & _
                                   "WHERE codusuario = '" & Trim(Right(MDI.mUser, 20)) & "'")
                                 
                                 
   Do While Not mRec.EOF
     Combo1(2).AddItem mRec!descripcion & Space(50) & mRec!CodSuperv
     mRec.MoveNext
   Loop
                                 
''Set mRec = mObj.oEjecutarSelect("SELECT * FROM Edificios WHERE FechaBaja IS NULL " & IIf(mTramo <> "", " And Tramo = '" & mTramo & "'", ""))
'Set mRec = mObj.oEjecutarSelect("SELECT * FROM Edificios WHERE FechaBaja IS NULL " & IIf(mTramo <> "", " And Tramo IN ('" & mTramo & "') order by 1,2", ""))
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
MSFlexGrid1.ColWidth(4) = 1800
MSFlexGrid1.TextMatrix(0, 1) = "Edificio"
MSFlexGrid1.TextMatrix(0, 2) = "Descripcion del Problema"
MSFlexGrid1.TextMatrix(0, 3) = "Prioridad"
For mi = 0 To MSFlexGrid1.Cols - 1
   MSFlexGrid1.ColAlignment(mi) = 0
Next

cboListIndex = Combo1(2).ListIndex

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set mObj = Nothing
Set mRec = Nothing
Set mObjLuser = Nothing
ShowMenu 20, True, False
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
Dim mObj As New clMantEd
Dim mRec As New ADODB.Recordset
   
   mCodSupervision = Trim(Right(Combo1(2).Text, 10))
   Combo1(0).Clear
   Set mRec = mObj.oEjecutarSelect("SELECT * FROM Edificios WHERE CodSuperv = '" & mCodSupervision & "' order by 2")
   Do While Not mRec.EOF
     'Combo1(0).AddItem mRec!descripcion & Space(60) & mRec!Codigo
      'Combo1(0).AddItem mRec!Tramo & " - " & mRec!descripcion
      Combo1(0).AddItem mRec!descripcion '-- Saco la zona
     mRec.MoveNext
   Loop
   mRec.Close
   Set mObj = Nothing
   Set mRec = Nothing
End Sub
