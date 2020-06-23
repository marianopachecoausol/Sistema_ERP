VERSION 5.00
Begin VB.Form RNov11 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WAZE"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9300
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   9300
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   4560
      Width           =   3855
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2040
      Width           =   5415
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3960
      MaxLength       =   5
      TabIndex        =   6
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2520
      MaxLength       =   10
      TabIndex        =   5
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3960
      MaxLength       =   5
      TabIndex        =   4
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2520
      MaxLength       =   10
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
      Height          =   615
      Index           =   1
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E8D9B9&
      Caption         =   "Enviar"
      Height          =   615
      Index           =   0
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7080
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   5160
      Width           =   5415
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   5
      Left            =   2520
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   5760
      Width           =   6015
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   2520
      MaxLength       =   5
      TabIndex        =   7
      Top             =   3840
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1560
      Width           =   3855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ramal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   1680
      TabIndex        =   22
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   7560
      TabIndex        =   21
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009F6713&
      Height          =   375
      Left            =   5880
      TabIndex        =   20
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SubTipo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   1440
      TabIndex        =   19
      Top             =   2120
      Width           =   795
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Hora Final"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   585
      TabIndex        =   18
      Top             =   3360
      Width           =   1635
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Hora Inicial"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   465
      TabIndex        =   17
      Top             =   2760
      Width           =   1740
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sentido"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   1440
      TabIndex        =   16
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Km"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   1890
      TabIndex        =   15
      Top             =   3960
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Novedad"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   1320
      TabIndex        =   14
      Top             =   5760
      Width           =   840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   9360
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ENVIAR INFORMACIÓN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009F6713&
      Height          =   585
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Width           =   4515
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   1800
      TabIndex        =   0
      Top             =   1680
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   5760
      Picture         =   "RNov11.frx":0000
      Stretch         =   -1  'True
      Top             =   75
      Width           =   3315
   End
End
Attribute VB_Name = "RNov11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   sInitForm
   sAlinearForm Me
End Sub

Private Sub Combo1_Click(Index As Integer)
   If Index = 0 Then
      sLlenoSubTipos Trim(Right(Combo1(0).Text, 1))
   End If
   
   If Index = 3 Then
      sLlenoSentido
   End If
   
   
End Sub

Private Sub Command1_Click(Index As Integer)
   Select Case Index
      Case 0 'enviar waze
         If fValidar() = True Then
            Dim mObj As New clRNov
            mObj.waze_enviar_accidente Label3.Caption, Text1(5).Text, Trim(Right(Combo1(1).Text, 2))
            sDoXML Label3.Caption
            Set mObj = Nothing
         End If
         
            Unload RNov1b_frm
            Unload RNov1d_frm
            Load RNov1b_frm
            Load RNov1d_frm

      Case 1
      
   End Select
End Sub


Private Sub Text1_Change(Index As Integer)
   If Index = 5 Then
      Label4.Caption = Len(Text1(Index).Text) & " / 250"
   End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 2 'fecha
         KeyAscii = fDateKeyPress(Text1(Index), KeyAscii)
      Case 1, 3 'hora
         KeyAscii = fHoraKeyPress(Text1(Index), KeyAscii)
      Case 4 'km
         KeyAscii = fKmsKeyPress(Text1(Index), KeyAscii)
      Case 5 'descr
         KeyAscii = fAlfaNumSimbKeyPress(KeyAscii)
         KeyAscii = fUcaseKeyPress(KeyAscii)
   End Select
End Sub


'----------------------------------------------------------------------------------------------------------
'PROCESOS Y FUNCIONES
'----------------------------------------------------------------------------------------------------------

Private Sub sInitForm()
Dim mObj As New clRNov
Dim mRec As New ADODB.Recordset

   Set mRec = mObj.oTabla("s_tiposincid", "order by 1")
   sLlenoCbo Combo1(0), mRec, 1, 0
   
   Set mRec = mObj.oTabla("ramales", "")
   Do While Not mRec.EOF
     Combo1(3).AddItem mRec!descripcion & Space(50) & mRec!Abrevia & Space(2) & mRec!Codigo
     mRec.MoveNext
   Loop
   mRec.Close
   

   
   
   Set mRec = mObj.oTabla("sentidos", "order by 1")
   If Not mRec.EOF Then
      Do While Not mRec.EOF
         'Combo1(2).AddItem mRec!codglobal & "-" & mRec!descripcion & Space(50) & mRec!Codigo
         Combo1(2).AddItem mRec!descripcion & Space(50) & mRec!Codigo
         mRec.MoveNext
      Loop
   End If
   mRec.Close
   Combo1(0).ListIndex = 0 'lo dejo en ACCIDENTE
   
   Combo1(0).Enabled = False
   Combo1(2).Enabled = False
   Combo1(3).Enabled = False
   Text1(4).Enabled = False
   
   Set mObj = Nothing
   Set mRec = Nothing
End Sub

Private Sub sLlenoSubTipos(ByVal pCodTipo As String)
Dim mObj As New clRNov
Dim mRec As New ADODB.Recordset
   
   Combo1(1).Clear
   Set mRec = mObj.oTabla("s_subtiposincid", " where codincid=" & pCodTipo & " order by 3")
   sLlenoCbo Combo1(1), mRec, 2, 0

   Set mObj = Nothing
   Set mRec = Nothing
End Sub

Public Sub sInitFormCod(ByVal pCodAlfa As String, ByVal pTipo As Integer)
Dim mObj As New clRNov
Dim mRec As New ADODB.Recordset
Dim mI As Integer
   
   Select Case pTipo
      Case 0
         Text1(0).Enabled = False
         Text1(1).Enabled = False
         Label1(5).Visible = False
         Text1(2).Visible = False
         Text1(3).Visible = False
      Case 1
      
      
   End Select
  
   'On Error Resume Next
   If Not mObj.bExistDatoTabla("d_waze", "codalfa='" & pCodAlfa & "'") Then

       Set mRec = mObj.waze_ultima_novedad(pCodAlfa)

       If Not mRec.EOF Then

                Label3.Caption = UCase(pCodAlfa)
               'no va nada xq la fecha se se obtiene en el procedimiento
               Text1(0).Text = Format(Now(), "dd/mm/yyyy")
               Text1(1).Text = Format(Now(), "HH:MM")

               Text1(4).Text = Format(mRec!km, "#0.00")

               For mI = 0 To Combo1(3).ListCount - 1
                  'If CInt(Trim(Right(Combo1(3).List(mI), 2))) = CInt(Trim(mRec!codramal)) Then
                  If CInt(Trim(Right(Combo1(3).List(mI), 2))) = mRec!codramal Then
                     Combo1(3).ListIndex = mI
                     mI = 999
                  End If
               Next
               For mI = 0 To Combo1(2).ListCount - 1
                  'If Trim(Right(Combo1(2).List(mI), 2)) = mRec!sent Then
                  If CInt(Trim(Right(Combo1(2).List(mI), 2))) = mRec!sent Then
                     Combo1(2).ListIndex = mI
                      mI = 999
                  End If
               Next



       End If
       mRec.Close

   Else
      If MsgBox("Ya envió este incidente, desea enviar una actualización a Waze?", vbYesNo, sMessage) = vbYes Then
         Set mRec = mObj.oTabla("d_waze", " where codigo='" & pCodAlfa & "' order by fecha")
         If Not mRec.EOF Then
            Label3.Caption = UCase(pCodAlfa)
            For mI = 0 To 3
               Text1(mI).Enabled = False
            Next
            Do While Not mRec.EOF
               If mRec!CodNov = "N" Then
                  Text1(0).Text = Format(mRec!Fecha, "dd/mm/yyyy")
                  Text1(1).Text = Format(mRec!Fecha, "HH:SS")
                  
                  Text1(4).Text = Format(mRec!km, "#0.00")
                  For mI = 0 To Combo1(2).ListCount - 1
                     If Trim(Left(Combo1(2).List(mI), 1)) = mRec!sent Then
                        Combo1(2).ListIndex = mI
                     End If
                  Next
               End If
               mRec.MoveNext
            Loop
         End If
         mRec.Close
      End If
   End If
   Set mObj = Nothing
   Set mRec = Nothing

End Sub

Private Function fValidar() As Boolean
Dim mText As String

   mText = ""
   fValidar = False
   If Fecha_ok(Text1(0).Text) = False Then mText = ". Fecha Inicial"
   If Hora_ok(Text1(1).Text) = False Then mText = mText & Chr(13) & ". Hora Inicial"
   'Fecha/Hora de cierre debería estar desde el detalle de RegNov
   If Combo1(0).ListIndex < 0 Then mText = mText & Chr(13) & ". Tipo de Incidente"
   If Combo1(2).ListIndex < 0 Then mText = mText & Chr(13) & ". Sentido de tránsito"
   If Trim(Text1(5).Text) = "" Then mText = mText & Chr(13) & ". Descripción del Incidente"
   If mText = "" Then
      fValidar = True
   Else
      MsgBox "Verificar lo siguiente: " & Chr(13) & mText, vbCritical, sMessage
   End If
End Function



Private Sub sDoXML_old20181112(ByVal pCodAlfa As String)
Dim mObj As New clRNov
Dim mRec As New ADODB.Recordset
Dim mNovedad As String
Dim mFechaCracion As String

   
   Set mRec = mObj.oTabla("d_waze", " where codalfa='" & pCodAlfa & "' order by fecha desc ")
   If Not mRec.EOF Then
   
      mFechaCracion = mObj.sFechaMySQL
   
      Open App.Path & "\RegNovedades\waze\ausolAR_" & pCodAlfa & ".xml" For Output As #1
      Print #1, "<?xml version='1.0' encoding='UTF-8'?>"
      Print #1, "<incidents xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xsi:noNamespaceSchemaLocation='http://www.gstatic.com/road-incidents/incidents_feed.xsd'>"
      Print #1, "<incident id='" & pCodAlfa & "'>"
      Print #1, " <creationtime>" & Format(mFechaCracion, "yyyy-mm-dd") & "T" & Format(mFechaCracion, "HH:mm:ss") & "-03:00" & "</creationtime>"
'      Print #1, "<creationtime>" & Format(Date, "yyyy-mm-dd") & "T" & Format(Time, "HH:mm:ss") & "-03:00" & "</creationtime>"
      Print #1, "<type>ACCIDENT</type>"
'      If Combo1(1).ListIndex > -1 Then
'         Print #1, "<subtype>" & mObj.sTablaDescr("s_subtiposincid", " codigo=" & Trim(Right(Combo1(1).Text, 2)), 3) & "</subtype>"
'      End If
      Print #1, "<subtype>" & mObj.sTablaDescr("s_subtiposincid", " codigo=" & mRec!codsubtipo, 3) & "</subtype>"
      Print #1, " <description>" & mRec!descripcion & "</description>"
      Print #1, "<location>"
      Print #1, "<polyline>" & mRec!lat & " " & mRec!lon & "</polyline>"
      Print #1, "</location>"
      'Print #1, " <starttime>" & Format(mRec!Fecha, "yyyy-mm-dd") & "T" & Format(mRec!Fecha, "HH:mm:ss") & "-03:00" & "</starttime>"
      Print #1, " <starttime>" & Format(mRec!fecha_i, "yyyy-mm-dd") & "T" & Format(mRec!fecha_i, "HH:mm:ss") & "-03:00" & "</starttime>"
      Print #1, " <name>Autopistas del Sol S.A.</name>"
      Print #1, "</incident>"
      Print #1, "</incidents>"
   
      Close #1
         
      'sEnviarFile App.Path & "\RegNovedades\waze\ausolAR_" & pCodAlfa & ".xml"
      
      'va waze_enviar_accidente MP
      'mObj.xInWaze pCodAlfa, Trim(Right(Combo1(0).Text, 2)), Trim(Right(Combo1(1).Text, 2)), Trim(Text1(4).Text), Trim(Right(Combo1(2).Text, 2)), Now(), Text1(2).Text & " " & Text1(3).Text, Trim(Text1(4).Text), mLat, mLong
        
   End If
   mRec.Close
   
   Unload Me
   Set mObj = Nothing
   Set mRec = Nothing
End Sub


'Private Function getRamalWazeXml(ByVal pSentido As Integer) As String
'Dim mObj As New clRNov
'Dim mRamal As Integer
'Dim sRamal As String
'Dim sSentido As String
'   mRamal = mObj.sTablaDescr("sentidos", " codigo=" & pSentido, 2)
'   sSentido = mObj.sTablaDescr("sentidos", " codigo=" & pSentido, 1)
'   sRamal = mObj.sTablaDescr("ramales", " codigo=" & mRamal, 1)
'
'   getRamalWazeXml = sRamal & " - " & Mid(sSentido, 4, Len(sSentido) - 3)
'End Function





Private Sub sDoXML(ByVal pCodAlfa As String)
Dim mObj As New clRNov
Dim mRec As New ADODB.Recordset
Dim mNovedad As String
Dim mFechaCracion As String




   Set mRec = mObj.oTabla("d_waze", " where estadowaze in (0, 1, 2) and codalfa='" & pCodAlfa & "' order by fecha desc ")
   If Not mRec.EOF Then
      Open App.Path & "\RegNovedades\waze\ausolAR.xml" For Output As #1
      'MsgBox App.Path & "\RegNovedades\waze\ausolAR.xml"
      Print #1, "<?xml version=""1.0"" encoding=""UTF-8""?>"
      Print #1, "<incidents xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:noNamespaceSchemaLocation=""http://www.gstatic.com/road-incidents/incidents_feed.xsd"">"
      Do While Not mRec.EOF
      
         Print #1, "<incident id=""" & mRec!CodAlfa & """ > "
         Print #1, "  <creationtime>" & Format(mRec!fecha_i, "yyyy-mm-dd") & "T" & Format(mRec!fecha_i, "HH:mm:ss") & "-03:00" & "</creationtime>"
         Print #1, "  <description>" & mRec!descripcion & "</description>"
         Print #1, "  <location>"




         Print #1, "<street>" & mObj.getRamalWazeXml(mRec!codsent) & "</street>"

         Print #1, "     <direction>ONE_DIRECTION</direction>"
         Print #1, "     <polyline>" & mRec!lat & " " & mRec!lon & "</polyline>"
         Print #1, "  </location>"
         Print #1, " <reference>AUSOL</reference>"
         Print #1, " <starttime>" & Format(mRec!fecha_i, "yyyy-mm-dd") & "T" & Format(mRec!fecha_i, "HH:mm:ss") & "-03:00" & "</starttime>"
         Print #1, " <endtime>" & Format(DateAdd("h", 4, mRec!fecha_upd), "yyyy-mm-dd") & "T" & Format(DateAdd("h", 4, mRec!fecha_upd), "HH:mm:ss") & "-03:00" & "</endtime>"
         Print #1, " <type>ACCIDENT</type>"
         Print #1, " <subtype>" & mObj.sTablaDescr("s_subtiposincid", " codigo=" & mRec!codsubtipo, 3) & "</subtype>"
         Print #1, "</incident>"
         
         mRec.MoveNext
      Loop
      mRec.Close
      
      
      Set mRec = mObj.oTabla("d_waze", " where estadowaze in (0, 1, 2) and codalfa<>'" & pCodAlfa & "' order by fecha desc ")
      Do While Not mRec.EOF
         Print #1, "<incident id=""" & mRec!CodAlfa & """ > "
         Print #1, "  <creationtime>" & Format(mRec!fecha_i, "yyyy-mm-dd") & "T" & Format(mRec!fecha_i, "HH:mm:ss") & "-03:00" & "</creationtime>"
         Print #1, "  <description>" & mRec!descripcion & "</description>"
         Print #1, "  <location>"

         Print #1, "<street>" & mObj.getRamalWazeXml(mRec!codsent) & "</street>"

         Print #1, "     <direction>ONE_DIRECTION</direction>"
         Print #1, "     <polyline>" & mRec!lat & " " & mRec!lon & "</polyline>"
         Print #1, "  </location>"
         Print #1, " <reference>AUSOL</reference>"
         Print #1, " <starttime>" & Format(mRec!fecha_i, "yyyy-mm-dd") & "T" & Format(mRec!fecha_i, "HH:mm:ss") & "-03:00" & "</starttime>"
         If DateDiff("n", DateAdd("h", 4, mRec!fecha_i), Now) > 0 Then
            mObj.xUpWazeDateUp mRec!CodAlfa
            Print #1, " <endtime>" & Format(DateAdd("h", 1, Now), "yyyy-mm-dd") & "T" & Format(DateAdd("h", 1, Now), "HH:mm:ss") & "-03:00" & "</endtime>"
         Else
            Print #1, " <endtime>" & Format(DateAdd("h", 4, mRec!fecha_i), "yyyy-mm-dd") & "T" & Format(DateAdd("h", 4, mRec!fecha_i), "HH:mm:ss") & "-03:00" & "</endtime>"
         End If
         Print #1, " <type>ACCIDENT</type>"
         Print #1, " <subtype>" & mObj.sTablaDescr("s_subtiposincid", " codigo=" & mRec!codsubtipo, 3) & "</subtype>"
         Print #1, "</incident>"
         mRec.MoveNext
      Loop
      mRec.Close
      
      'Verifico la existencia de Cierres menores a 10 minutos
      Set mRec = mObj.oTabla("d_waze", " where estadowaze=3 and fecha_f > date_add(current_timestamp, interval -10 MINUTE)   order by fecha desc ")
       Do While Not mRec.EOF
         Print #1, "<incident id=""" & mRec!CodAlfa & """ > "
         Print #1, "  <creationtime>" & Format(mRec!fecha_i, "yyyy-mm-dd") & "T" & Format(mRec!fecha_i, "HH:mm:ss") & "-03:00" & "</creationtime>"
         Print #1, "  <updatetime>" & Format(mRec!fecha_upd, "yyyy-mm-dd") & "T" & Format(mRec!fecha_upd, "HH:mm:ss") & "-03:00" & "</updatetime>"
         Print #1, "  <description>" & mRec!descripcion & "</description>"
         Print #1, "  <location>"

         Print #1, "<street>" & mObj.getRamalWazeXml(mRec!codsent) & "</street>"

         Print #1, "     <direction>ONE_DIRECTION</direction>"
         Print #1, "     <polyline>" & mRec!lat & " " & mRec!lon & "</polyline>"
         Print #1, "  </location>"
         Print #1, " <reference>AUSOL</reference>"
         Print #1, " <starttime>" & Format(mRec!fecha_i, "yyyy-mm-dd") & "T" & Format(mRec!fecha_i, "HH:mm:ss") & "-03:00" & "</starttime>"
         Print #1, " <endtime>" & Format(mRec!fecha_f, "yyyy-mm-dd") & "T" & Format(mRec!fecha_f, "HH:mm:ss") & "-03:00" & "</endtime>"
         Print #1, " <type>ACCIDENT</type>"
         Print #1, " <subtype>" & mObj.sTablaDescr("s_subtiposincid", " codigo=" & mRec!codsubtipo, 3) & "</subtype>"
         Print #1, "</incident>"
                  
         mRec.MoveNext
      Loop
      
      Print #1, "</incidents>"
   
      Close #1
      Sleep 2
      
       sEnviarFile "ausolAR.xml"  'Este está ok!
      
   End If
   mRec.Close
   
   Unload Me
   Set mObj = Nothing
   Set mRec = Nothing
End Sub

Private Sub sEnviarFile(ByVal pArchivo As String)
Dim Fso As New FileSystemObject
Dim Persistent As Boolean
Dim errResult

   Fso.CopyFile App.Path & "\RegNovedades\waze\" & pArchivo, "L:\" & pArchivo
   DoEvents
   
   Sleep 2
   
   Me.SetFocus
''   mObjNet.RemoveNetworkDrive "J:"
   MsgBox "Datos enviado a Waze.", vbInformation, sMessage

End Sub





Private Sub sEnviarFileOLD20181112(ByVal mArchivo As String)
Dim oShell As WshShell
Dim oExec As WshExec
Dim ret As String
Dim mComando As String
Dim ejecutar_Dos
    mComando = App.Path & "\pscp.exe -pw AdminGCO2010$ -sftp " & mArchivo & " ausolwaze@52.162.163.73:/home/ausolwaze"
    Set oShell = New WshShell
    DoEvents

    ' ejecutar el comando
    Set oExec = oShell.Exec("%comspec% /c " & mComando)
    ret = oExec.StdOut.ReadAll()

    ' retornar la salida y devolverla a la función
    'Text1.Text = ret  ' Replace(ret, Chr(10), vbNewLine)

    DoEvents
    Me.SetFocus
    
    MsgBox "Datos enviado a Waze.", vbInformation, sMessage
    
End Sub

Private Sub sLlenoSentido()
Dim mCodRamal As String
Dim mObj As New clRNov
Dim mRec As New ADODB.Recordset
   
   mCodRamal = Right(Combo1(3).Text, 1)
   Combo1(2).Clear
   Set mRec = mObj.oTabla("sentidos", "where codRamal = " & mCodRamal & " order by 2")
   Do While Not mRec.EOF
     Combo1(2).AddItem mRec!descripcion & Space(60) & mRec!Codigo
     mRec.MoveNext
   Loop
   mRec.Close
   Set mObj = Nothing
   Set mRec = Nothing
End Sub


'Private Sub xCoorde(ByVal pCodSent As String, ByVal pKm As String, ByRef pLat As Double, ByRef pLon As Double)
'Dim mObj As New clRNov
'Dim mRec As New ADODB.Recordset
'
'
'
'
'
'
'End Sub
