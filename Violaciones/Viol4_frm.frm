VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Viol4_frm 
   BackColor       =   &H00CECECE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Módulo Obtener Violaciones"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13515
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   13515
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   3720
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   900
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   540
      Width           =   2715
   End
   Begin VB.CommandButton Command2 
      Caption         =   "XLS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5280
      TabIndex        =   19
      Top             =   1080
      Width           =   735
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   6360
      TabIndex        =   18
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Agregar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   5160
      TabIndex        =   17
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   5040
      MaxLength       =   15
      TabIndex        =   16
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Listado &Direcciones"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   2
      Left            =   9360
      TabIndex        =   12
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Listado &Cartas Doc."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   1
      Left            =   9360
      TabIndex        =   11
      Top             =   120
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid MsFlex 
      Height          =   5895
      Left            =   60
      TabIndex        =   9
      Top             =   1560
      Width           =   13395
      _ExtentX        =   23627
      _ExtentY        =   10398
      _Version        =   327680
      Cols            =   12
      FixedCols       =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Obtener"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   7800
      TabIndex        =   8
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   2040
      MaxLength       =   5
      TabIndex        =   4
      Top             =   900
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   2640
      MaxLength       =   10
      TabIndex        =   2
      Top             =   180
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1020
      MaxLength       =   10
      TabIndex        =   1
      Top             =   180
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CECECE&
      Caption         =   "Tipo"
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
      Left            =   3240
      TabIndex        =   21
      Top             =   960
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CECECE&
      Caption         =   "Acciones"
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
      Left            =   180
      TabIndex        =   20
      Top             =   600
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CECECE&
      Caption         =   "Patentes"
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
      Left            =   5040
      TabIndex        =   15
      Top             =   120
      Width           =   765
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00CECECE&
      Caption         =   "Total de Pasadas"
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
      Left            =   11520
      TabIndex        =   14
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00CECECE&
      Caption         =   "Para Detalles hacer Doble Clic sobre una Patente"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   3510
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00CECECE&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   11760
      TabIndex        =   10
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CECECE&
      Caption         =   "Cant. Violaciones >="
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
      Left            =   180
      TabIndex        =   7
      Top             =   960
      Width           =   1770
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CECECE&
      Caption         =   "al"
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
      Left            =   2280
      TabIndex        =   6
      Top             =   240
      Width           =   165
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CECECE&
      Caption         =   "Fechas"
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
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   630
   End
End
Attribute VB_Name = "Viol4_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mObjViol As New clViolaciones
Dim mObjPea As New clPeaje
Dim mRec1 As New ADODB.Recordset
Dim mRec2 As New ADODB.Recordset
Dim mRs As New ADODB.Recordset
Dim XLS As EXCEL.Application
Dim mI As Integer
Public mTipo As String

Private Sub Form_Load()
Me.Width = 13605
Me.Height = 7905
sAlinearForm Me
sTituloFlex
sInitForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set mObjViol = Nothing
Set mObjPea = Nothing
Set mRec1 = Nothing
Set mRec2 = Nothing
Set mRs = Nothing
ShowMenu 5, True, False
End Sub

Private Sub Command1_Click(Index As Integer)
Dim mNombre As String
Dim mDirec As String
Dim mCP As String
Dim mLocalidad As String
Dim mPcia As String
Dim mEnvCarta As String
Dim mFecha As String
Dim mRango As String
Dim mEstaciones(18) As String
Dim mWhere As String
Dim mObs As String
Dim mUltEntrega As String
Dim mTuvoCD As String
Dim mJ As Integer
Dim mT As Integer
Dim mInd2 As Integer
Dim mCont As Integer
Dim mFlag As Boolean
Dim mProvincia As String

   Select Case Index
      Case 0
         sMsgEspere Me, "Procesando... espere un momento.", True
         mJ = 0
         If Label2.Caption <> "" Then
            sBorraFlexDatos Viol4_frm.MsFlex
         End If
         Command1(1).Enabled = False
         Command1(2).Enabled = False
         If fValida Then 'valida fechas y diferencia de fechas
            mWhere = ""
            If List1.ListCount > 0 Then
               mWhere = "AND patente IN ("
               For mI = 0 To List1.ListCount - 1
                  mWhere = mWhere & "'" & List1.List(mI) & "',"
               Next
               mWhere = mId(mWhere, 1, Len(mWhere) - 1) & ")"
            End If
            mTipo = IIf(Left(Combo2.Text, 1) = "V", "V", "D")
            Select Case Combo1.ListIndex
               Case 0, 4
                  Set mRec1 = mObjViol.oViolxPatenteDate(Text1(0).Text, Text1(1).Text, Trim(Text1(2).Text), mWhere, mTipo)
                  Do While Not mRec1.EOF
                     mUltEntrega = mObjViol.fUltEnvio(mRec1!patente, mTipo) 'si tuvo entregas
                     If mUltEntrega = "" Then
                        Set mRec2 = mObjViol.oDatosPatente(mRec1!patente)
                        If Not mRec2.EOF And Combo1.ListIndex = 0 Then
                           mLocalidad = ""
                           mProvincia = ""
                           If fVerStandBy(mRec1!patente, mTipo) Then
                              mLocalidad = mObjViol.sCampoDescrip("postal", "codigo='" & mRec2!codpostal & "' and codpcia='" & mRec2!codpcia & "'", 2)
                              mProvincia = mObjViol.sCampoDescrip("provincias", "codigo='" & mRec2!codpcia & "'", 1)
                              MsFlex.AddItem mRec1!patente & vbTab & mRec2!nombre & vbTab & mRec2!domicilio & vbTab & mLocalidad & vbTab & mProvincia & vbTab & mRec2!codpostal & vbTab & mRec1!Total & vbTab & "NO" & vbTab & "X" & vbTab & "" & vbTab & mUltEntrega & vbTab & ""
                              Command1(1).Enabled = True
                           End If
                        Else
                           If Combo1.ListIndex = 4 And mRec2.EOF Then
                              If fVerStandBy(mRec1!patente, mTipo) Then
                                 MsFlex.AddItem mRec1!patente & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & mRec1!Total & vbTab & mEnvCarta & vbTab & "X" & vbTab & "NO" & vbTab & mUltEntrega & vbTab & ""
                                 Command1(2).Enabled = True
                              End If
                           End If
                        End If
                        mRec2.Close
                     End If
                     mRec1.MoveNext
                  Loop
                  mRec1.Close
                  'Todos los vehículos que tengan pasos en el período ingresado, con una cantidad >= a la ingresada y
                  'además que tengan envios "sin recibir" a partir del 01/07/2007.
                  'NO IMPORTA CANT. DE ENVIOS.
               Case 1 'Acción 2
                  Set mRec1 = mObjViol.oViolxPatenteNoEntrega(Text1(0).Text, Text1(1).Text, Trim(Text1(2).Text), mWhere, mTipo)
                  Do While Not mRec1.EOF
                     'Si no esta en stand by pasa
                     If fVerStandBy(mRec1!patente, mTipo) Then
                        sInitVariables mNombre, mDirec, mLocalidad, mCP, mPcia, mEnvCarta, mObs, mTuvoCD
                        'Si no tuvo CD recibidas
                        If mObjViol.bCDRechazadas(mRec1!patente, "01/01/2007", Text1(1).Text, mTipo) Then
                           sDatosPatente mRec1!patente, mNombre, mDirec, mLocalidad, mCP, mPcia
                           mT = 0
                           Set mRec2 = mObjViol.oCDporPatente(mRec1!patente, mTipo)
                           Do While Not mRec2.EOF
                              If mObs = "" Then
                                 mObs = NVL(mRec2("OBS"), " ")
                                 If mObs = "" Then
                                    mObs = " "
                                 End If
                                 mTuvoCD = "SI"
                                 mUltEntrega = NVL(mRec2!NROCARTA, "")
                              End If
                              mEnvCarta = mEnvCarta & NVL(mRec2("NROCARTA"), "") & " - "
                              mRec2.MoveNext
                              mT = mT + 1
                           Loop
                           mRec2.Close
                           If mTuvoCD = "SI" And mT < 4 Then  'mt<4 que tenga menos de 4 cartas documentos enviadas
                              MsFlex.AddItem mRec1!patente & vbTab & mNombre & vbTab & mDirec & vbTab & mLocalidad & vbTab & mPcia & vbTab & mCP & vbTab & mRec1!Total & vbTab & mTuvoCD & vbTab & "X" & vbTab & mEnvCarta & vbTab & mUltEntrega & vbTab & mObs
                              Command1(1).Enabled = True
                              Command1(2).Enabled = True
                           End If
                        End If
                     End If
                     mRec1.MoveNext
                  Loop
                  mRec1.Close
               Case 2, 3 'Acción 3 (PATENTES CON UNA C.D. ENTREGADA Y SIN PAGAR....)
                  mWhere = ""
                  Set mRec1 = mObjPea.oPatentePagoViol
                  If Not mRec1.EOF Then
                     mWhere = " AND A.patente NOT IN ('"
                     Do While Not mRec1.EOF
                        mWhere = mWhere & mRec1!patente & "','"
                        mRec1.MoveNext
                     Loop
                     mWhere = mId(mWhere, 1, (Len(mWhere) - 2)) & ")"
                  End If
                  mRec1.Close
                  Set mRec1 = mObjViol.oCDEnvioRecib(Text1(0).Text, Text1(1).Text, Trim(Text1(2).Text), mWhere, mTipo)
                  If Not mRec1.EOF Then
                     Do While Not mRec1.EOF
                        If fVerStandBy(mRec1!patente, mTipo) Then
                           sInitVariables mNombre, mDirec, mLocalidad, mCP, mPcia, mEnvCarta, mObs, mTuvoCD
                           sDatosPatente mRec1!patente, mNombre, mDirec, mLocalidad, mCP, mPcia
                           'Detalle de cartas por patente ordenada por fecha DESC
                           mFlag = False
                           Set mRec2 = mObjViol.oCDporPatente(mRec1!patente, mTipo)
                           If Not mRec2.EOF Then
                              If Combo1.ListIndex = 2 Then
                                 mFlag = (mRec2!codentrega = "00")
                              Else
                                 mFlag = (mRec2!codentrega <> "00")
                              End If
                           End If
                           If mFlag Then
                              Do While Not mRec2.EOF
                                 If mObs = "" Then
                                    mCont = -1
                                    If DateDiff("d", Text1(0).Text, mRec2!Fecha) < 0 Then
                                       mCont = mObjViol.iCountViolFechaPatente2(mRec1!patente, mRec2!Fecha, mTipo)
                                    End If
                                    'Si tiene pasadas después de la última carta entregada
                                    If mCont >= Trim(Text1(2).Text) Then
                                       mObs = "Catidad de pasos: " & mCont
                                       mTuvoCD = "SI"
                                       mUltEntrega = NVL(mRec2!NROCARTA, "")
                                    Else
                                       mObs = " "
                                    End If
                                 End If
                                 mEnvCarta = mEnvCarta & NVL(mRec2("NROCARTA"), "") & " - "
                                 mRec2.MoveNext
                                 mT = mT + 1
                              Loop
                           End If
                           mRec2.Close
                           'datos de cartas y totales
                           MsFlex.AddItem mRec1!patente & vbTab & mNombre & vbTab & mDirec & vbTab & mLocalidad & vbTab & mPcia & vbTab & mCP & vbTab & mRec1!Total & vbTab & mTuvoCD & vbTab & "X" & vbTab & mEnvCarta & vbTab & mUltEntrega & vbTab & mObs
                        End If
                        mRec1.MoveNext
                     Loop
                     Command1(1).Enabled = True
                  End If
                  mRec1.Close
               Case 5 'Violadores Stand-By
                  Set mRec1 = mObjViol.oTabla("regpagos", "where estado='1' And tipo = '" & mTipo & "' order by 1")
                  If Not mRec1.EOF Then
                     Do While Not mRec1.EOF
                        mWhere = " and patente='" & mRec1!patente & "'"
                        'obtener ultima fecha de envio de CD para pasar como parametro de primer fecha...
                        mFecha = mObjViol.fUltFechaEnvio(mRec1!patente, mTipo)
                        If mFecha = "" Then
                           mFecha = Text1(0).Text
                        Else
                           If DateDiff("d", Text1(0).Text, mFecha) < 0 Then
                              mFecha = Text1(0).Text
                           End If
                        End If
                        Set mRec2 = mObjViol.oViolxPatenteDate(mFecha, Text1(1).Text, Trim(Text1(2).Text), mWhere, mTipo)
                        If Not mRec2.EOF Then
                           sDatosPatente mRec1!patente, mNombre, mDirec, mLocalidad, mCP, mPcia
                           sDatosCartasPat mRec1!patente, mObs, mTuvoCD, mUltEntrega, mEnvCarta
                           MsFlex.AddItem mRec1!patente & vbTab & mNombre & vbTab & mDirec & vbTab & mLocalidad & vbTab & mPcia & vbTab & mCP & vbTab & mRec2!Total & vbTab & mTuvoCD & vbTab & "X" & vbTab & mEnvCarta & vbTab & mUltEntrega & vbTab & ""
                        End If
                        mRec2.Close
                        mRec1.MoveNext
                     Loop
                  Else
                     MsgBox "Sin Datos para el período", vbInformation, sMessage
                  End If
                  mRec1.Close
               Case 6 'Violadores para legales
                  Set mRec1 = mObjViol.oCDRecibidas(Text1(0).Text, Text1(1).Text, mTipo)
                  If Not mRec1.EOF Then
                     Do While Not mRec1.EOF
                        If fVerStandBy(mRec1!patente, mTipo) Then
                           If mRec1!Total = 1 Then
                              mFecha = mRec1!Fecha
                           Else
                              mFecha = mObjViol.fUltFechaEnvio(mRec1!patente, mTipo)
                           End If
                           Set mRec2 = mObjViol.oViolxPatenteDate(mFecha, Now, "10", " and patente='" & mRec1!patente & "'", mTipo)
                           If Not mRec2.EOF Then
                              mCont = 0
                              sInitVariables mNombre, mDirec, mLocalidad, mCP, mPcia, mEnvCarta, mObs, mTuvoCD
                              mCont = mRec2!Total
                              mTuvoCD = "SI"
                              mRec2.Close
                              Set mRec2 = mObjViol.oDatosPatente(mRec1!patente)
                              If Not mRec2.EOF Then
                                 mNombre = mRec2!nombre
                                 mDirec = mRec2!domicilio
                                 mCP = mRec2!codpostal
                                 mLocalidad = mObjViol.sCampoDescrip("postal", "codigo='" & mRec2!codpostal & "' and codpcia='" & mRec2!codpcia & "'", 2)
                                 mPcia = mObjViol.sCampoDescrip("provincias", "codigo='" & mRec2!codpcia & "'", 1)
                              End If
                              mUltEntrega = mObjViol.fUltEnvio(mRec1!patente, mTipo)
                              If mNombre <> "" And mDirec <> "" Then
                                 MsFlex.AddItem mRec1!patente & vbTab & mNombre & vbTab & mDirec & vbTab & mLocalidad & vbTab & mPcia & vbTab & mCP & vbTab & mCont & vbTab & mTuvoCD & vbTab & "X" & vbTab & mEnvCarta & vbTab & mUltEntrega & vbTab & mObs
                                 Command1(1).Enabled = True
                              End If
                           End If
                           mRec2.Close
                        End If
                        mRec1.MoveNext
                     Loop
                  Else
                     MsgBox "Sin Datos para el período", vbInformation, sMessage
                  End If
                  mRec1.Close
               Case 7 'Pasos de usuarios en Stand By
                  Set mRec1 = mObjViol.oEjecutarSelect("SELECT patente, COUNT(*) As Total FROM Registros a WHERE a.tipo = '" & mTipo & "' And a.fecha BETWEEN '" & Format(Text1(0).Text, "yyyy-mm-dd") & "' AND '" & Format(Text1(1).Text, "yyyy-mm-dd") & "' And a.fecha > (SELECT MAX(b.fecha) FROM regpagos b WHERE b.tipo = '" & mTipo & "' And a.patente = b.patente) GROUP BY patente HAVING Total >= '" & Text1(2).Text & "' ORDER BY patente")
                  If Not mRec1.EOF Then
                     Do While Not mRec1.EOF
                        sDatosPatente mRec1!patente, mNombre, mDirec, mLocalidad, mCP, mPcia
                        MsFlex.AddItem mRec1!patente & vbTab & mNombre & vbTab & mDirec & vbTab & mLocalidad & vbTab & mPcia & vbTab & mCP & vbTab & mRec1!Total & vbTab & "" & vbTab & "X" & vbTab & "" & vbTab & "" & vbTab & ""
                        mRec1.MoveNext
                     Loop
                  Else
                     MsgBox "Sin Datos para el período", vbInformation, sMessage
                  End If
                  mRec1.Close
            End Select
            sSetFlex2Colors Viol4_frm.MsFlex, &HE0E0E0, &HFFFFFF
            If MsFlex.Rows > 2 Then
               MsFlex.RemoveItem 1
            End If
            Label2.Caption = (MsFlex.Rows - 1)
         End If
         sMsgEspere Me, "", False
      Case 1
         'Genera listado de Patentes y Pasadas para Correo
          Viol4_frm.MousePointer = 11
          sMsgEspere Me, "Procesando... espere un momento.", True
          Set mRec2 = mObjPea.oEstaciones("")
          While Not mRec2.EOF
             mEstaciones(mRec2!CODIGO_ESTACION) = Trim(mRec2!Descripcion_Estacion)
             mRec2.MoveNext
          Wend
          mRec2.Close
          Set XLS = CreateObject("Excel.Application")
          sCabecera Index
          mRango = "A1:J1"
          'nombre y dirección
          sFormatCells mRango, 15
          sCabecera2
          sFormatCells "A1:D1", 15
          mJ = 2
          mInd2 = 1
          For mI = 1 To Label2.Caption
             XLS.Sheets(1).Select
             mNombre = MsFlex.TextMatrix(mI, 1)
             mDirec = MsFlex.TextMatrix(mI, 2)
             mObs = MsFlex.TextMatrix(mI, 10)
             If Trim(mNombre) <> "" And Trim(mDirec) <> "" And MsFlex.TextMatrix(mI, 8) = "X" Then
                XLS.Cells(mJ, 2).Formula = mNombre
                XLS.Cells(mJ, 3).Formula = mDirec
                XLS.Cells(mJ, 9).Formula = mObs
                For mT = 3 To 6  'mT es la columna del XLS
                   XLS.Cells(mJ, mT + 1).Formula = MsFlex.TextMatrix(mI, mT)
                Next
                XLS.Cells(mJ, 1).Formula = Trim(MsFlex.TextMatrix(mI, 0))
                XLS.Cells(mJ, 8).Formula = mObjViol.fUltEnvio(Trim(MsFlex.TextMatrix(mI, 0)), mTipo)
                mJ = mJ + 1
                'Pasadas
                Set mRec2 = mObjViol.oViolFechasPatente(Text1(0).Text, Text1(1).Text, Trim(MsFlex.TextMatrix(mI, 0)), mTipo)
                XLS.Sheets(2).Select
                While Not mRec2.EOF
                   mInd2 = mInd2 + 1
                   XLS.Cells(mInd2, 1) = mEstaciones(Int(Val(mRec2!Estacion)))
                   XLS.Cells(mInd2, 2) = mRec2!Fecha
                   XLS.Cells(mInd2, 3) = mRec2!Hora
                   XLS.Cells(mInd2, 4) = MsFlex.TextMatrix(mI, 0)
                   mRec2.MoveNext
                Wend
                mRec2.Close
             End If
          Next
          XLS.Sheets(1).Select
          sFormatCells "A2:I" & (mJ - 1), 2
          XLS.Sheets(2).Select
          sFormatCells "A2:D" & (mInd2), 2
          Viol4_frm.MousePointer = 0
          XLS.Visible = True
          Set XLS = Nothing
          sMsgEspere Me, "", False
      Case 2
        'Genera listado de Patentes para averiguar direcciones
         sMsgEspere Me, "Procesando... espere un momento.", True
         Set XLS = CreateObject("Excel.Application")
         sCabecera Index
         XLS.Application.DisplayAlerts = False
         XLS.Sheets(3).Select
         XLS.ActiveWindow.SelectedSheets.Delete
         XLS.Sheets(2).Select
         XLS.ActiveWindow.SelectedSheets.Delete
         XLS.Application.DisplayAlerts = True
         mRango = "A1:E1"
         sFormatCells mRango, 15
         mJ = 2
         For mI = 1 To Label2.Caption
            mNombre = MsFlex.TextMatrix(mI, 1)
            mDirec = MsFlex.TextMatrix(mI, 2)
            If Trim(mNombre) = "" Or Trim(mDirec) = "" Then
               XLS.Cells(mJ, 1).Formula = MsFlex.TextMatrix(mI, 0)
               mJ = mJ + 1
            End If
         Next
         sFormatCells "A2:E" & (mJ - 1), 2
         XLS.Visible = True
         Set XLS = Nothing
         sMsgEspere Me, "", False
      Case 3 'agregar patentes al LIST
         If Text1(3).Text <> "" Then
            mFlag = True
            For mI = 0 To List1.ListCount - 1
               If List1.List(mI) = Trim(Text1(3).Text) Then
                  MsgBox "La patente ya Existe en el listado", vbCritical, sMessage
                  mFlag = False
               End If
            Next
            If mFlag Then
               List1.AddItem Trim(Text1(3).Text)
               Text1(3).Text = ""
            End If
         Else
            MsgBox "Agregar primero una patente", vbCritical, sMessage
         End If
   End Select
End Sub

Private Sub Command2_Click()
Set XLS = CreateObject("Excel.Application")
XLS.Application.WorkBooks.Open filename:="C:\patentes.xls"
XLS.Worksheets(1).Select
mI = 1
Do While XLS.Cells(mI, 1) <> ""
   List1.AddItem XLS.Cells(mI, 1)
   mI = mI + 1
Loop
MsgBox "Pasaje Finalizado", vbInformation, sMessage
Set XLS = Nothing
End Sub

Private Sub List1_DblClick()
List1.RemoveItem List1.ListIndex
End Sub

Private Sub MsFlex_DblClick()
If MsFlex.Row > 0 Then
   Select Case MsFlex.Col
      Case 0
         If Left(Combo1.Text, 1) <> "8" Then    'Pregunto si no son los que pasaron luego de pagar para la última opción del combo1.
            Viol5_frm.mpFechaD = Text1(0).Text
            Viol5_frm.mpFechaH = Text1(1).Text
         Else
            Set mRec1 = mObjViol.oEjecutarSelect("SELECT MAX(fecha) As MAXFECHA FROM regpagos WHERE tipo = '" & mTipo & "' And patente = '" & MsFlex.Text & "'")
            Viol5_frm.mpFechaD = mRec1!MAXFECHA
            Viol5_frm.mpFechaH = Format(Date, "yyyy-mm-dd")
            mRec1.Close
         End If
         Viol5_frm.mpPatente = MsFlex.Text
         Viol5_frm.Show
         Viol4_frm.Enabled = False
      Case 8
         If MsFlex.Text = "X" Then
            MsFlex.Text = ""
         Else
            MsFlex.Text = "X"
         End If
   End Select
End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
   Case 0, 1
      KeyAscii = fDateKeyPress(Text1(Index), KeyAscii)
   Case 2
      KeyAscii = fNumeroKeyPress(KeyAscii)
   Case 3
      KeyAscii = fAlfaNumKeyPress(KeyAscii)
      KeyAscii = fUcaseKeyPress(KeyAscii)
End Select
End Sub

Private Sub sInitForm()
Combo1.AddItem "1) Datos Sin Envíos previos"
Combo1.AddItem "2) C.D Enviadas s/recibir"
Combo1.AddItem "3) Primer C.D. recibida sin pagar"
Combo1.AddItem "4) C.D. un solo envío sin recibir"
Combo1.AddItem "5) Listado patentes sin direcciones"
Combo1.AddItem "6) Violadores Stand-By"
Combo1.AddItem "7) Violadores para legales"
Combo1.AddItem "8) Pasos de usuarios en Stand By"
Combo2.AddItem "Violacion"
Combo2.AddItem "Rec.Deuda"
End Sub

Private Sub sTituloFlex()
With MsFlex
   .ColWidth(0) = 1000
   .ColWidth(1) = 2700
   .ColWidth(2) = 3000
   .ColWidth(3) = 1800
   .ColWidth(4) = 1500
   .ColWidth(5) = 800
   .ColWidth(6) = 700
   .ColWidth(7) = 700
   .ColWidth(8) = 400
   .ColWidth(9) = 7000
   .ColWidth(10) = 2000
   .ColWidth(11) = 15000
   .TextMatrix(0, 0) = "Patente"
   .TextMatrix(0, 1) = "Nombre"
   .TextMatrix(0, 2) = "Direcc"
   .TextMatrix(0, 3) = "Localidad"
   .TextMatrix(0, 4) = "Pcia"
   .TextMatrix(0, 5) = "CP"
   .TextMatrix(0, 6) = "C.Viol"
   .TextMatrix(0, 7) = "E.Cart"
   .TextMatrix(0, 8) = "Sel"
   .TextMatrix(0, 9) = "Nro. Cartas"
   .TextMatrix(0, 10) = "Ult. Entrega"
   .TextMatrix(0, 11) = "Obs"
   .Row = 0
   For mI = 0 To 11
      .Col = mI
      .CellFontBold = True
   Next
End With
End Sub

Private Sub sCabecera(pItem As Integer)
Dim mRango As String
XLS.WorkBooks.Add
XLS.Worksheets(1).Select
If pItem = 2 Then
   mRango = "A1:E1"
   XLS.Columns("D:D").ColumnWidth = 15    'provincia
   XLS.Columns("E:E").ColumnWidth = 10   'cod. postal
   XLS.Worksheets(1).Name = "Datos"
Else
   mRango = "A1:J1"
   XLS.Columns("D:D").ColumnWidth = 24    'localidad
   XLS.Columns("E:E").ColumnWidth = 15    'provincia
   XLS.Columns("F:F").ColumnWidth = 10  'cod. postal
   XLS.Columns("G:G").ColumnWidth = 7.7   'pasadas
   XLS.Columns("H:H").NumberFormat = "@"
   XLS.Columns("H:H").ColumnWidth = 24   'Ult. Envio
   XLS.Columns("I:I").ColumnWidth = 250   'Ult. Entrega
   XLS.Columns("J:J").ColumnWidth = 250   'observaciones
   XLS.Worksheets(1).Name = "Nombr-Direcc."
   XLS.Worksheets(2).Name = "Estac-Día-Hora"
   XLS.Worksheets(3).Name = "Tarifario"
End If
XLS.Range(mRango).Font.Bold = True
XLS.Range(mRango).Font.Name = "Arial"
XLS.Range(mRango).Font.Size = 10
XLS.Cells.Select
XLS.Selection.Interior.ColorIndex = 2
XLS.Selection.Interior.Pattern = xlSolid
XLS.Columns("A:A").ColumnWidth = 10.7  'patente
XLS.Columns("B:B").ColumnWidth = 30.8  'nombre
XLS.Columns("C:C").ColumnWidth = 35.5  'direccion
XLS.Cells(1, 1).Formula = "Patente"
XLS.Cells(1, 2).Formula = "Nombre"
XLS.Cells(1, 3).Formula = "Dirección"
If pItem = 2 Then
   XLS.Cells(1, 4).Formula = "Provincia"
   XLS.Cells(1, 5).Formula = "Cod.Postal"
Else
   XLS.Cells(1, 4).Formula = "Localidad"
   XLS.Cells(1, 5).Formula = "Provincia"
   XLS.Cells(1, 6).Formula = "Cod.Postal"
   XLS.Cells(1, 7).Formula = "Pasadas"
   XLS.Cells(1, 8).Formula = "Ult. Envío"
   XLS.Cells(1, 9).Formula = "Ult. Entrega"
   XLS.Cells(1, 10).Formula = "Observaciones"
End If
End Sub

Private Sub sCabecera2()
Dim mRango As String
mRango = "A1:E1"
XLS.Worksheets(2).Select
XLS.Columns("A:A").ColumnWidth = 14    'estacion
XLS.Columns("B:B").ColumnWidth = 10    'fecha
XLS.Columns("B:B").NumberFormat = "d-mmm-yy"
XLS.Columns("B:B").HorizontalAlignment = xlCenter
XLS.Columns("C:C").ColumnWidth = 7     'hora
XLS.Columns("C:C").HorizontalAlignment = xlCenter
XLS.Columns("E:E").ColumnWidth = 10.7  'patente
XLS.Range(mRango).Font.Bold = True
XLS.Range(mRango).Font.Name = "Arial"
XLS.Range(mRango).Font.Size = 10
XLS.Cells.Select
XLS.Selection.Interior.ColorIndex = 2
XLS.Selection.Interior.Pattern = xlSolid
XLS.Cells(1, 1).Formula = "Estacion"
XLS.Cells(1, 2).Formula = "Fecha"
XLS.Cells(1, 3).Formula = "Hora"
XLS.Cells(1, 4).Formula = "Patente"
End Sub

Private Sub sFormatCells(pRango As String, pColor As Integer)
XLS.Range(pRango).Select
XLS.Selection.Interior.ColorIndex = pColor
XLS.Selection.Interior.Pattern = xlSolid
XLS.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
XLS.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
XLS.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
XLS.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
XLS.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
XLS.Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
XLS.Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
On Error Resume Next
XLS.Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
End Sub

Private Sub sInitVariables(ByRef pNombre As String, ByRef pDirec As String, ByRef pLocalidad As String, ByRef pCP As String, ByRef pProv As String, ByRef pNumCarta As String, ByRef pObs As String, ByRef pTuvoCD As String)
pNombre = ""
pDirec = ""
pLocalidad = ""
pCP = ""
pProv = ""
pNumCarta = ""
pObs = ""
pTuvoCD = "NO"
End Sub

Private Function fVerStandBy(ByVal pPatente As String, ByVal pTipo As String) As Boolean
fVerStandBy = True
If Trim(pPatente) <> "" Then
   Set mRs = mObjViol.oPatentesStBy(pPatente, pTipo)
   If Not mRs.EOF Then
      If mRs!estado = "1" Then
         fVerStandBy = False
      End If
   End If
   mRs.Close
End If
End Function

Private Sub sDatosPatente(ByVal pPatente As String, ByRef pNombre As String, ByRef pDirec As String, ByRef pLocal As String, ByRef pCP As String, ByRef pProv As String)
Set mRs = mObjViol.oDatosPatente(mRec1!patente)
pNombre = ""
pDirec = ""
pLocal = ""
pCP = ""
pProv = ""
If Not mRs.EOF Then
   pNombre = NVL(mRs!nombre, "")
   pDirec = NVL(mRs!domicilio, "")
   'pLocal = NVL(mRs!localidad, "")
   pCP = NVL(mRs!codpostal, "")
  ' pProv = NVL(mRs!prov, "")
   
   pLocal = mObjViol.sCampoDescrip("postal", "codigo='" & mRs!codpostal & "' and codpcia='" & mRs!codpcia & "'", 2)
   pProv = mObjViol.sCampoDescrip("provincias", "codigo='" & mRs!codpcia & "'", 1)
   
   
End If
mRs.Close
End Sub

Private Sub sDatosCartasPat(ByVal pPatente As String, ByRef pObs As String, ByRef pTuvoCD As String, ByRef pUltEntrega As String, ByRef pEnvCartas As String)
Set mRs = mObjViol.oCDporPatente(pPatente, mTipo)
pObs = ""
pTuvoCD = "NO"
pUltEntrega = ""
pEnvCartas = ""
Do While Not mRs.EOF
   If pObs = "" Then
      pObs = NVL(mRs("OBS"), " ")
      pTuvoCD = "SI"
      pUltEntrega = NVL(mRs!NROCARTA, "")
   End If
   pEnvCartas = pEnvCartas & NVL(mRs("NROCARTA"), "") & " - "
   mRs.MoveNext
Loop
mRs.Close
End Sub

Private Function fValida() As Boolean
fValida = Fecha_ok(Text1(0).Text)
fValida = fValida And Fecha_ok(Text1(1).Text)
fValida = fValida And (Text1(2).Text <> "")
fValida = fValida And (Combo1.Text <> "")
fValida = fValida And (Combo2.Text <> "")
If fValida Then
   fValida = (DateDiff("d", Text1(0).Text, Text1(1).Text) > 0)
End If
If Not fValida Then
   MsgBox "Verifique los datos de entrada", vbCritical, "Atención"
End If
End Function

