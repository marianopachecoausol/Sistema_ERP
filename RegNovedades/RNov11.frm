VERSION 5.00
Begin VB.Form RNov11 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WAZE"
   ClientHeight    =   6990
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
   ScaleHeight     =   6990
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
      Index           =   1
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   17
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
      TabIndex        =   15
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
      TabIndex        =   14
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
      TabIndex        =   13
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
      TabIndex        =   12
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
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E8D9B9&
      Caption         =   "Enviar"
      Height          =   615
      Index           =   0
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6240
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
      Top             =   4320
      Width           =   2415
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
      TabIndex        =   8
      Top             =   4920
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
      TabIndex        =   2
      Top             =   1560
      Width           =   3855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   7560
      TabIndex        =   20
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
      TabIndex        =   19
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
      TabIndex        =   18
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
      TabIndex        =   16
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
      TabIndex        =   6
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
      Left            =   1485
      TabIndex        =   5
      Top             =   4440
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
      TabIndex        =   4
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
      Left            =   1350
      TabIndex        =   3
      Top             =   4920
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
         Name            =   "Calibri"
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
      TabIndex        =   1
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
      Height          =   1335
      Left            =   5280
      Picture         =   "RNov11.frx":0000
      Top             =   80
      Width           =   3795
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
End Sub

Private Sub Command1_Click(Index As Integer)
   Select Case Index
      Case 0 'enviar waze
         If fValidar() = True Then
            sDoXML Label3.Caption
         End If
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
   Set mRec = mObj.oTabla("s_sentidos", "order by 1")
   If Not mRec.EOF Then
      Do While Not mRec.EOF
         Combo1(2).AddItem mRec!codglobal & "-" & mRec!descripcion & Space(50) & mRec!Codigo
         mRec.MoveNext
      Loop
   End If
   mRec.Close
   Combo1(0).ListIndex = 0 'lo dejo en ACCIDENTE
   Combo1(0).Enabled = False
      
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
   
   If mObj.bExistDatoTabla("d_waze", "codalfa='" & pCodAlfa & "'") = False Then
      Set mRec = mObj.oTabla("Novedades2", " where codigo='" & pCodAlfa & "' order by fecha")
      If Not mRec.EOF Then
         Label3.Caption = UCase(pCodAlfa)
         Do While Not mRec.EOF
            If mRec!CodNov = "N" Then
               Text1(0).Text = Format(mRec!Fecha, "dd/mm/yyyy")
               Text1(1).Text = Format(mRec!Fecha, "HH:SS")
               Text1(4).Text = Format(mRec!km, "#0.00")
               For mI = 0 To Combo1(2).ListCount - 1
                  If Trim(Left(Combo1(2).List(mI), 1)) = mRec!Sent Then
                     Combo1(2).ListIndex = mI
                  End If
               Next
               
            End If
            mRec.MoveNext
         Loop
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
                     If Trim(Left(Combo1(2).List(mI), 1)) = mRec!Sent Then
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



Private Sub sDoXML(ByVal pCodAlfa As String)
Dim mObj As New clRNov
Dim mRec As New ADODB.Recordset
Dim mFechaI As String
Dim mNovedad As String
Dim mLat As Double
Dim mLong As Double

    
   mLat = "-34.6357383"
   mLong = "-58.7861602"
   
   
   
   mObj.dCoordenadas Trim(Right(Combo1(2).Text, 2)), Text1(4).Text, mLat, mLong
   
   
'   <?xml version="1.0" encoding="UTF-8"?>
'  <incidents xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="http://www.gstatic.com/road-incidents/incidents_feed.xsd">
'    <incident id="AQdGFM3wDo">
'      <creationtime>2015-11-16T13:25:29-05:00</creationtime>
'      <updatetime>2015-12-02T14:53:20-05:00</updatetime>
'      <type>ROAD_CLOSED</type>
'      <description>Construction on I-95 NB between Exit 184: ME 222 and Exit 185: ME 15. All northbound lanes closed due to road works.</description>
'      <location>
'        <street>I-95 NB</street>
'        <polyline>44.808819 -68.793266 44.819252 -68.775346</polyline>
'      </location>
'      <starttime>2016-04-11T00:00:00-04:00</starttime>
'      <endtime>2016-04-15T24:00:00-04:00</endtime>
'    </incident>4
'  </incidents>
   
   
   Set mRec = mObj.oTabla("novedades2", " where codigo='" & pCodAlfa & "' order by fecha ")
   If Not mRec.EOF Then
      Open App.Path & "\RegNovedades\waze\auoesteAR_" & pCodAlfa & ".xml" For Output As #1
      Print #1, "<?xml version='1.0' encoding='UTF-8'?>"
      Print #1, "<incidents xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xsi:noNamespaceSchemaLocation='http://www.gstatic.com/road-incidents/incidents_feed.xsd'>"
      Print #1, "<incident id='" & pCodAlfa & "'>"
      Print #1, "<creationtime>" & Format(Date, "yyyy-mm-dd") & "T" & Format(Time, "HH:mm:ss") & "-03:00" & "</creationtime>"
      Print #1, "<type>ACCIDENT</type>"
      If Combo1(1).ListIndex > -1 Then
         Print #1, "<subtype>" & mObj.sTablaDescr("s_subtiposincid", " codigo=" & Trim(Right(Combo1(1).Text, 2)), 3) & "</subtype>"
      End If
      Print #1, " <description>" & Trim(Text1(5).Text) & "</description>"
      Print #1, "<location>"
      Print #1, "<polyline>" & mLat & " " & mLong & "</polyline>"
      Print #1, "</location>"
      Print #1, " <starttime>" & Format(mRec!Fecha, "yyyy-mm-dd") & "T" & Format(mRec!Fecha, "HH:mm:ss") & "-03:00" & "</starttime>"
      Print #1, "</incident>"
      Print #1, "</incidents>"
   
      Close #1
         
      sEnviarFile App.Path & "\RegNovedades\waze\auoesteAR_" & pCodAlfa & ".xml"
      
      
      mObj.xInWaze pCodAlfa, Trim(Right(Combo1(0).Text, 2)), Trim(Right(Combo1(1).Text, 2)), Trim(Text1(4).Text), Trim(Right(Combo1(2).Text, 2)), Now(), Text1(2).Text & " " & Text1(3).Text, Trim(Text1(4).Text), mLat, mLong
        
   End If
   mRec.Close
   
   Unload Me
   Set mObj = Nothing
   Set mRec = Nothing
End Sub

Private Sub sEnviarFile(ByVal mArchivo As String)
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
