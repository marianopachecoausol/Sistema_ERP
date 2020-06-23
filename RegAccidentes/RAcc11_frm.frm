VERSION 5.00
Begin VB.Form RAcc11_frm 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2790
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4350
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4350
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
      Height          =   375
      Index           =   1
      Left            =   3120
      MaskColor       =   &H00404040&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2280
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   435
      Index           =   0
      Left            =   1620
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2220
      Width           =   1155
   End
   Begin VB.ComboBox Combo1 
      ForeColor       =   &H00404040&
      Height          =   315
      Index           =   1
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1260
      Width           =   1515
   End
   Begin VB.ComboBox Combo1 
      ForeColor       =   &H00404040&
      Height          =   315
      Index           =   0
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   840
      Width           =   2475
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Año:"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   3
      Top             =   1380
      Width           =   330
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mes:"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   345
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Informe Índices"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      Left            =   960
      TabIndex        =   0
      Top             =   180
      Width           =   1935
   End
End
Attribute VB_Name = "RAcc11_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mObjRN As New clRNov
Dim mObjRAcc As New clRAcc
Dim mRec As New ADODB.Recordset
Dim XLS As EXCEL.Application
Dim mi As Integer
Dim mAccTotal(3) As Integer
Dim mFecha1 As String
Dim mFecha2 As String
Dim mMes As String
Dim mAnio As String

Private Sub Form_Load()
   sAlinearForm Me
   sMsgEspere Me, "Iniciando...", True
   For mi = 1 To 12
      Combo1(0).AddItem Format(mi, "00") & " - " & MonthName(mi)
   Next
   Set mRec = mObjRN.oNovedAnios
   Do While Not mRec.EOF
      Combo1(1).AddItem mRec.Fields(0)
      mRec.MoveNext
   Loop
   mRec.Close
   sMsgEspere Me, "", False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set mObjRN = Nothing
   Set mObjRAcc = Nothing
   Set mRec = Nothing
   Set XLS = Nothing
   ShowMenu 2, True, False
End Sub

Private Sub Command1_Click(Index As Integer)
   If Index = 0 Then
      'Ver si se hace por planillas o sale todo de una
      'para esto se necesitaria un select case
      If fValidar Then
         sMsgEspere Me, "Procesando datos...", True
         mMes = Left(Combo1(0).Text, 2)
         mAnio = Combo1(1).Text
         mFecha1 = "01/" & mMes & "/" & mAnio
         mFecha2 = DateAdd("d", -1, DateAdd("m", 1, "01/" & mMes & "/" & mAnio))
         Set XLS = CreateObject("Excel.Application")
         
         sPlanilla1
         sPlanilla2
         sPlanilla3
         sPlanilla4
         sPlanilla5
         sPlanilla6
         sPlanilla7
         sPlanilla8
         sPlanilla9_10
         
         sMsgEspere Me, "", False
         XLS.Application.Visible = True
      End If
   Else
      Unload Me
   End If
End Sub

Private Function fValidar() As Boolean
   fValidar = (Combo1(0).ListIndex > -1 And Combo1(1).ListIndex > -1)
   fValidar = fValidar And Combo1(0).ListIndex > -1
   If Combo1(1).Text = Year(Date) Then
       fValidar = fValidar And (Val(Left(Combo1(0).Text, 2)) <= Month(Date))
   End If
End Function

Private Sub sPlanilla1()
   sCabecera1
   sServGruas
   sDepejeVias
   sAccidentes
End Sub

Private Sub sPlanilla2()
Dim mCant As Integer
   sCabecera2
   'Incidentes
   XLS.Cells(14, 2).Formula = mObjRAcc.iCountIncidentes(mFecha1, mFecha2)
   'Accidentes
   XLS.Cells(15, 2).Formula = mAccTotal(1)
   XLS.Cells(16, 2).Formula = mAccTotal(2)
   XLS.Cells(17, 2).Formula = mAccTotal(3)
   XLS.Cells(18, 2).Formula = mAccTotal(1) + mAccTotal(2) + mAccTotal(3)
   'Intervenciones de patrullas
   XLS.Cells(19, 2).Formula = mObjRN.iCountServMovil(mFecha1, mFecha2, "'A','MM','E'", "", "left(Mov1,1)='M'")
   XLS.Cells(20, 2).Formula = mObjRAcc.iTotalMuertosCodigo(mFecha1, mFecha2, "(A.fallecio <> '' or A.codestado in ('04','05')) and B.AccidOtro in ('02','03')") ' and B.lugaraccid = '09'")
   XLS.Cells(21, 2).Formula = mObjRAcc.iTotalMuertosCodigo(mFecha1, mFecha2, "(A.fallecio <> '' or A.codestado in ('04','05')) and B.AccidOtro not in ('02','03')") ' and B.lugaraccid = '09'")
   'Serv. Grúas Livianas
   mCant = 0
   For mi = 1 To 3
      mCant = mCant + mObjRN.iCountServMovil(mFecha1, mFecha2, "'A','MM'", "Mov1='GMUL'", "left(Mov" & mi & ",2)='G0'")
   Next
   XLS.Cells(23, 2).Formula = mCant
   'Serv. Grúas Pesadas
   mCant = 0
   For mi = 1 To 3
      mCant = mCant + mObjRN.iCountServMovil(mFecha1, mFecha2, "'A','MM'", "", "left(Mov" & mi & ",2)='GP'")
   Next
   XLS.Cells(24, 2).Formula = mCant
   'Serv. Bomberos
   XLS.Cells(25, 2).Formula = mObjRN.iCountServicios(mFecha1, mFecha2, "'M'", "BOMB", "")
   'Serv. Vittal
   XLS.Cells(26, 2).Formula = mObjRN.iCountServicios(mFecha1, mFecha2, "'M'", "AMBU", "")
   'Operativos Gendarmería
   XLS.Cells(31, 2).Formula = mObjRN.iCountServicios(mFecha1, mFecha2, "'P'", "GEND", "")
   'Operativos Policía
   XLS.Cells(32, 2).Formula = mObjRN.iCountServicios(mFecha1, mFecha2, "'P'", "POLI", "")
   XLS.Cells(37, 2).Formula = "=Planilla1!B29"
   XLS.Cells(38, 2).Formula = "=Planilla1!B31"
End Sub

Private Sub sPlanilla3() 'Arribos de Ambulancias por minutos de demora y códigos de pedidos
Dim mCodigo As String
Dim mFecha As Date
Dim mVecDatos(3, 9) As Integer
Dim mj As Integer
   Set mRec = mObjRN.oCodigosAmbu(mFecha1, mFecha2)
   If Not mRec.EOF Then
      For mi = 1 To 3
         For mj = 1 To 9
            mVecDatos(mi, mj) = 0
         Next
      Next
      sCabecera3
      Do While Not mRec.EOF
         If mRec!CodNov = "M" Then
            mCodigo = mRec!Codigo
            mFecha = mRec!Fecha
            mi = Val(NVL(mRec!codserv1, "1"))
         Else
            If mCodigo = mRec!Codigo Then
               Select Case DateDiff("s", mFecha, mRec!Fecha)
                  Case 0 To 180: mj = 1
                  Case 181 To 360: mj = 2
                  Case 361 To 540: mj = 3
                  Case 541 To 720: mj = 4
                  Case 721 To 900: mj = 5
                  Case 901 To 1080: mj = 6
                  Case 1081 To 1260: mj = 7
                  Case 1261 To 1440: mj = 8
                  Case Is > 1440: mj = 9
               End Select
               mVecDatos(mi, mj) = mVecDatos(mi, mj) + 1
            End If
         End If
         mRec.MoveNext
      Loop
      With XLS
         For mj = 1 To 9
            .Cells(13, mj + 1).Formula = " " & mVecDatos(2, mj)
            .Cells(14, mj + 1).Formula = " " & mVecDatos(1, mj)
            .Cells(15, mj + 1).Formula = " " & mVecDatos(3, mj)
         Next
         .Cells(13, 11).Formula = "=SUM(B13:J13)"
         .Cells(14, 11).Formula = "=SUM(B14:J14)"
         .Cells(15, 11).Formula = "=SUM(B15:J15)"
         .Cells(16, 11).Formula = "=SUM(K13:K15)"
         
         For mi = 2 To 10
            .Cells(17, mi).Formula = "=" & Chr(mi + 64) & "13/K13"
            .Cells(19, mi).Formula = "=" & Chr(mi + 64) & "14/K14"
         Next
         .Cells(18, 6).Formula = "=(SUM(B13:F13))/K13"
         .Cells(18, 8).Formula = "=(SUM(B13:H13))/K13"
         .Cells(18, 9).Formula = "=(SUM(B13:I13))/K13"
         .Cells(18, 10).Formula = "=(SUM(B13:J13))/K13"
         .Cells(20, 9).Formula = "=(SUM(B14:I14))/K14"
         .Cells(20, 10).Formula = "=(SUM(B14:J14))/K14"
         .Cells(41, 2).Formula = "=SUM(B14:F14)"
         .Cells(42, 2).Formula = "=F13"
         .Cells(46, 2).Formula = "=SUM(B14:I14)"
         .Cells(47, 2).Formula = "=I14"
      End With
   End If
   mRec.Close
   
End Sub

Private Sub sPlanilla4()
Dim mFecha As String
Dim mVecDatos(2, 9) As Integer
Dim mj As Integer
   
   Set mRec = mObjRAcc.oTabla("Ficha", "where fecha between '" & Format(mFecha1, "yyyy-mm-dd") & "' and '" & Format(mFecha2, "yyyy-mm-dd") & "' order by fecha")
   If Not mRec.EOF Then
      For mj = 1 To 9
         mVecDatos(1, mj) = 0
         mVecDatos(2, mj) = 0
      Next
      sCabecera4
      Do While Not mRec.EOF
         mFecha = Format(mRec!Fecha & " " & mRec!HoraLlegada, "dd-mm-yyyy hh:mm")
         If Left(mRec!hora, 2) > Left(mRec!HoraLlegada, 2) Then
            mFecha = Format(DateAdd("d", 1, mRec!Fecha) & " " & mRec!HoraLlegada, "dd-mm-yyyy hh:mm")
         End If
         Select Case DateDiff("s", Format(mRec!Fecha & " " & mRec!hora, "dd-mm-yyyy hh:mm"), mFecha)
            Case 0 To 180: mj = 1
            Case 181 To 360: mj = 2
            Case 361 To 540: mj = 3
            Case 541 To 720: mj = 4
            Case 721 To 900: mj = 5
            Case 901 To 1080: mj = 6
            Case 1081 To 1260: mj = 7
            Case 1261 To 1440: mj = 8
            Case Is > 1440: mj = 9
         End Select
         If NVL(mRec!codtipoficha, "01") = "02" Then
            mVecDatos(2, mj) = mVecDatos(2, mj) + 1
         Else
            mVecDatos(1, mj) = mVecDatos(1, mj) + 1
         End If
         mRec.MoveNext
      Loop
      With XLS
         For mj = 1 To 9
            .Cells(12, mj + 1).Formula = " " & mVecDatos(1, mj)
            .Cells(13, mj + 1).Formula = " " & mVecDatos(2, mj) 'falta para incidencias
         Next
         .Cells(12, 11).Formula = "=SUM(B12:J12)"
         .Cells(13, 11).Formula = "=SUM(B13:J13)"
         For mi = 2 To 10
            .Cells(15, mi).Formula = "=" & Chr(mi + 64) & "12/K12"
            .Cells(17, mi).Formula = "=" & Chr(mi + 64) & "13/K13"
         Next
         .Cells(16, 5).Formula = "=(SUM(B12:E12))/K12"
         .Cells(16, 6).Formula = "=F12/K12"
         .Cells(16, 10).Formula = "=(SUM(G12:J12))/K12"
         'incidencias
         .Cells(17, 6).Formula = "=(SUM(B13:F13))/K13"
         .Cells(17, 7).Formula = "=G13/K13"
         .Cells(17, 8).Formula = "=H13/K13"
         .Cells(17, 9).Formula = "=I13/K13"
         .Cells(17, 10).Formula = "=J13/K13"
         'accidentes
         .Cells(37, 2).Formula = "=SUM(B12:F12)"
         .Cells(38, 2).Formula = "=B21/K12"
         'incidencias
         .Cells(42, 2).Formula = "=SUM(B13:I13)"
         .Cells(43, 2).Formula = "=B42/K13"
      End With
   End If
   mRec.Close
End Sub

Private Sub sPlanilla5OLD()
Dim mj As Integer
Dim mDatos(2, 6) As Integer
   For mi = 1 To 6
      mDatos(1, mj) = 0
      mDatos(2, mj) = 0
   Next
   sCabecera5
   With XLS
      .Cells(4, 2).Formula = mObjRAcc.iCantCausaFechas(mFecha1, mFecha2, " lugaraccid<>'09'")
      .Cells(5, 2).Formula = mObjRAcc.iAccidVictProgr(mFecha1, mFecha2, 13, 67, True)
      mDatos(1, 1) = mObjRAcc.iMuertosHeridosProgr(mFecha1, mFecha2, 13, 67, True, False)
      mDatos(1, 2) = mObjRAcc.iMuertosHeridosProgr(mFecha1, mFecha2, 13, 67, True, True)
      'peatones
      mDatos(1, 3) = mObjRAcc.iCountFallHerVictVehic(mFecha1, mFecha2, True, False, " and A.Accidotro='03'")
      mDatos(1, 4) = mObjRAcc.iCountFallHerVictVehic(mFecha1, mFecha2, True, True, " and A.Accidotro='03'")
      'ciclistas
      mDatos(1, 5) = mObjRAcc.iCountFallHerVictVehic(mFecha1, mFecha2, True, False, " and A.Accidotro='02'")
      mDatos(1, 6) = mObjRAcc.iCountFallHerVictVehic(mFecha1, mFecha2, True, True, " and A.Accidotro='02'")
      'COLECTORA
      .Cells(18, 2).Formula = mObjRAcc.iCantCausaFechas(mFecha1, mFecha2, " lugaraccid='09'")
      .Cells(19, 2).Formula = mObjRAcc.iAccidVictProgr(mFecha1, mFecha2, 13, 67, False)
      mDatos(2, 1) = mObjRAcc.iMuertosHeridosProgr(mFecha1, mFecha2, 13, 67, False, False)
      mDatos(2, 2) = mObjRAcc.iMuertosHeridosProgr(mFecha1, mFecha2, 13, 67, False, True)
      'peatones
      mDatos(2, 3) = mObjRAcc.iCountFallHerVictVehic(mFecha1, mFecha2, False, False, " and A.Accidotro='03'")
      mDatos(2, 4) = mObjRAcc.iCountFallHerVictVehic(mFecha1, mFecha2, False, True, " and A.Accidotro='03'")
      'ciclista
      mDatos(2, 5) = mObjRAcc.iCountFallHerVictVehic(mFecha1, mFecha2, False, False, " and A.Accidotro='02'")
      mDatos(2, 6) = mObjRAcc.iCountFallHerVictVehic(mFecha1, mFecha2, False, True, " and A.Accidotro='02'")
      For mi = 3 To 6
         .Cells(5 + mi, 2).Formula = mDatos(1, mi)
         .Cells(19 + mi, 2).Formula = mDatos(2, mi)
      Next
      .Cells(6, 2).Formula = mDatos(1, 1) - (mDatos(1, 3) + mDatos(1, 5))
      .Cells(7, 2).Formula = mDatos(1, 2) - (mDatos(1, 4) + mDatos(1, 6))
      .Cells(20, 2).Formula = mDatos(2, 1) - (mDatos(2, 3) + mDatos(2, 5))
      .Cells(21, 2).Formula = mDatos(2, 2) - (mDatos(2, 4) + mDatos(2, 6))
      
      'heridos sin lesiones
      .Cells(16, 2).Formula = mObjRAcc.iCountSinLesionHerVictVehic(mFecha1, mFecha2, True, "")  'MP20180629
      .Cells(30, 2).Formula = mObjRAcc.iCountSinLesionHerVictVehic(mFecha1, mFecha2, False, "") 'MP20180629

      
   End With
End Sub

Private Sub sPlanilla5()
Dim mj As Integer
Dim mDatos(2, 6) As Integer
   For mi = 1 To 6
      mDatos(1, mj) = 0
      mDatos(2, mj) = 0
   Next
   sCabecera5
   With XLS
      .Cells(4, 2).Formula = mObjRAcc.iCantCausaFechas(mFecha1, mFecha2, " lugaraccid<>'09'")
      .Cells(5, 2).Formula = mObjRAcc.iAccidVictProgr(mFecha1, mFecha2, 0, 73, True)
      mDatos(1, 1) = mObjRAcc.iMuertosHeridosProgr(mFecha1, mFecha2, 0, 73, True, False)
      mDatos(1, 2) = mObjRAcc.iMuertosHeridosProgr(mFecha1, mFecha2, 0, 73, True, True)
      'peatones
      mDatos(1, 3) = mObjRAcc.iCountFallHerVictVehic(mFecha1, mFecha2, True, False, " and A.Accidotro='03'")
      mDatos(1, 4) = mObjRAcc.iCountFallHerVictVehic(mFecha1, mFecha2, True, True, " and A.Accidotro='03'")
      'ciclistas
      mDatos(1, 5) = mObjRAcc.iCountFallHerVictVehic(mFecha1, mFecha2, True, False, " and A.Accidotro='02'")
      mDatos(1, 6) = mObjRAcc.iCountFallHerVictVehic(mFecha1, mFecha2, True, True, " and A.Accidotro='02'")
      'COLECTORA
      .Cells(18, 2).Formula = mObjRAcc.iCantCausaFechas(mFecha1, mFecha2, " lugaraccid='09'")
      .Cells(19, 2).Formula = mObjRAcc.iAccidVictProgr(mFecha1, mFecha2, 0, 73, False)
      mDatos(2, 1) = mObjRAcc.iMuertosHeridosProgr(mFecha1, mFecha2, 0, 73, False, False)
      mDatos(2, 2) = mObjRAcc.iMuertosHeridosProgr(mFecha1, mFecha2, 0, 73, False, True)
      'peatones
      mDatos(2, 3) = mObjRAcc.iCountFallHerVictVehic(mFecha1, mFecha2, False, False, " and A.Accidotro='03'")
      mDatos(2, 4) = mObjRAcc.iCountFallHerVictVehic(mFecha1, mFecha2, False, True, " and A.Accidotro='03'")
      'ciclista
      mDatos(2, 5) = mObjRAcc.iCountFallHerVictVehic(mFecha1, mFecha2, False, False, " and A.Accidotro='02'")
      mDatos(2, 6) = mObjRAcc.iCountFallHerVictVehic(mFecha1, mFecha2, False, True, " and A.Accidotro='02'")
      For mi = 3 To 6
         .Cells(5 + mi, 2).Formula = mDatos(1, mi)
         .Cells(19 + mi, 2).Formula = mDatos(2, mi)
      Next
      .Cells(6, 2).Formula = mDatos(1, 1) - (mDatos(1, 3) + mDatos(1, 5))
      .Cells(7, 2).Formula = mDatos(1, 2) - (mDatos(1, 4) + mDatos(1, 6))
      .Cells(20, 2).Formula = mDatos(2, 1) - (mDatos(2, 3) + mDatos(2, 5))
      .Cells(21, 2).Formula = mDatos(2, 2) - (mDatos(2, 4) + mDatos(2, 6))
      
      'heridos sin lesiones
      .Cells(16, 2).Formula = mObjRAcc.iCountSinLesionHerVictVehic(mFecha1, mFecha2, True, "")  'MP20180629
      .Cells(30, 2).Formula = mObjRAcc.iCountSinLesionHerVictVehic(mFecha1, mFecha2, False, "") 'MP20180629

      
   End With
End Sub


Private Sub sPlanilla6()
Dim mRec1 As New ADODB.Recordset
Dim mTexto As String
Dim mHoraFin As String
Dim mj As Integer

   mi = 6
   sCabecera6
   'Set mRec = mObjRAcc.oFichasFechaProgr(mFecha1, mFecha2, "", "", "")
   'Set mRec = mObjRAcc.oTabla("Ficha", " where fecha between '" & Format(mFecha1, "yyyy-mm-dd") & "' and '" & Format(mFecha2, "yyyy-mm-dd") & "' order by 1")
   
   Set mRec = mObjRAcc.oEjecutarSelect(" SELECT * FROM Ficha F " & _
                                       " INNER JOIN " & _
                                       " regnov.ramales R ON F.codramal = R.codigo " & _
                                       " where fecha between '" & Format(mFecha1, "yyyy-mm-dd") & "' and '" & Format(mFecha2, "yyyy-mm-dd") & "' order by 1")
   
'   SELECT * FROM Ficha F
'Inner Join
'regnov.ramales R ON F.codramal = R.codigo
'where Fecha between '2020-01-01' and '2020-01-31' order by 1;
   
   
   
   Do While Not mRec.EOF
      mTexto = ""
      With XLS
         .Cells(mi, 1).Formula = NVL(mRec!CodAlfa, "")
         .Cells(mi, 2).Formula = NVL(mRec!NroOrden, "")
         .Cells(mi, 3).Formula = mRec!Fecha
         .Cells(mi, 4).Formula = mRec!hora
         .Cells(mi, 5).Formula = mRec!Progresiva
         .Cells(mi, 36).Formula = mRec!descripcion
         If mRec!lugaraccid = "09" Then
            .Cells(mi, 6).Formula = "Colect"
         Else
            .Cells(mi, 6).Formula = mObjRAcc.sTablaDescr("SentidoTrans", "codsentidotrans='" & NVL(mRec!SentidoTrans, "") & "'", 1)
         End If
         .Cells(mi, 7).Formula = mRec!HoraLlegada
         Set mRec1 = mObjRAcc.oCountFallHerTipo(mRec!NroOrden, True, "")
         Do While Not mRec1.EOF
            mTexto = mTexto & mRec1!Total & "-" & mObjRAcc.sTablaDescr("estadoocupa", "codigo='" & mRec1!codestado & "'", 1) & " "
            mRec1.MoveNext
         Loop
         mRec1.Close
         .Cells(mi, 8).Formula = mTexto  'Cant. de heridos
         .Cells(mi, 9).Formula = mObjRAcc.iTotalMuertosNroOrden(mRec!NroOrden)
         .Cells(mi, 10).Formula = mObjRAcc.iTotalVehiculosNroOrden(mRec!NroOrden)
         If NVL(mRec!codtipoficha, "01") = "02" Then
            .Cells(mi, 11).Formula = "Incidente"
         Else
            .Cells(mi, 11).Formula = "Colisión"
            mTexto = mObjRAcc.sTablaDescr("ColisionContra", "codcolision='" & mRec!CodColisContra1 & "'", 1)
            If mTexto <> "" Then
               .Cells(mi, 11).Formula = mTexto
            End If
         End If
         .Cells(mi, 12).Formula = mObjRAcc.sTablaDescr("CausaConductor", "codcausacond='" & mRec!CodCausaCond1 & "'", 1)
         mHoraFin = mObjRN.oHoraFinAcc(mRec!Fecha, NVL(mRec!CodAlfa, ""))
         .Cells(mi, 13).Formula = mHoraFin
         .Cells(mi, 14).Formula = mObjRAcc.sTablaDescr("Clima", "codclima='" & mRec!Clima1 & "'", 1)
         .Cells(mi, 15).Formula = mObjRAcc.sTablaDescr("fichaobs", "nroorden='" & mRec!NroOrden & "'", 2)
         Set mRec1 = mObjRAcc.oTabla("interterceros", " where nroorden='" & mRec!NroOrden & "' and codtipo in ('01','03','04','05','08')")
         Do While Not mRec1.EOF
            Select Case mRec1!codtipo
               Case "01"
                  .Cells(mi, 16).Formula = NVL(mRec1!dependencia, "")
               Case "03"
                  .Cells(mi, 20).Formula = NVL(mRec1!dependencia, "")
               Case "04"
                  .Cells(mi, 19).Formula = NVL(mRec1!dependencia, "")
               Case "05"
                  .Cells(mi, 17).Formula = NVL(mRec1!dependencia, "")
               Case "08"
                  .Cells(mi, 18).Formula = NVL(mRec1!dependencia, "")
            End Select
            mRec1.MoveNext
         Loop
         mRec1.Close
         
         .Cells(mi, 21).Formula = mObjRAcc.iCantVehicNroFicha2(mRec!NroOrden, "'01'")  'bicicletas
         .Cells(mi, 22).Formula = mObjRAcc.iCantVehicNroFicha2(mRec!NroOrden, "'02','03'") 'motos
         .Cells(mi, 23).Formula = mObjRAcc.iCantVehicNroFicha2(mRec!NroOrden, "'04','05','15'") 'autos
         .Cells(mi, 24).Formula = mObjRAcc.iCantVehicNroFicha2(mRec!NroOrden, "'06','07','08','09','10','11','12','13','14','16'") 'camiones y varios
         
         
         For mj = 1 To 11
            If Mid(mRec!carril, mj, 1) = "1" Then
               .Cells(mi, 24 + mj) = "X"
            End If
         Next
      End With
      
      mRec.MoveNext
      mi = mi + 1
   Loop
   mRec.Close
End Sub

Private Sub sPlanilla7()
   sCabecera7
   With XLS
      .Cells(5, 2).Formula = mObjRAcc.iTotalMuertosCodigo(mFecha1, mFecha2, "edad <= 18 and (fallecio <> '' or codestado in ('04','05'))")
      .Cells(6, 2).Formula = mObjRAcc.iTotalMuertosCodigo(mFecha1, mFecha2, "edad between 19 and 25 and (fallecio <> '' or codestado in ('04','05'))")
      .Cells(7, 2).Formula = mObjRAcc.iTotalMuertosCodigo(mFecha1, mFecha2, "edad between 26 and 45 and (fallecio <> '' or codestado in ('04','05'))")
      .Cells(8, 2).Formula = mObjRAcc.iTotalMuertosCodigo(mFecha1, mFecha2, "edad > 45 and (fallecio <> '' or codestado in ('04','05'))")
      .Cells(9, 2).Formula = "=SUM(B5:B8)"
      If .Cells(9, 2).Formula <> "0" Then
         .Cells(10, 2).Formula = ""
         .Cells(11, 2).Formula = ""
         .Cells(12, 2).Formula = ""
         .Cells(13, 2).Formula = ""
      End If
   End With
End Sub

Private Sub sPlanilla8()
   With XLS
      .Worksheets(8).Select
      .Worksheets(8).Name = "Planilla8"
      .Columns("A:A").ColumnWidth = 35
      .Cells(2, 1).Formula = "Tipo de Colisiones contra..."
      .Cells(3, 1).Formula = "Mes:"
      .Cells(3, 2).Formula = " " & Mid(Combo1(0).Text, 6, 3) & "-" & Right(Combo1(1).Text, 2)
      .Cells(9, 1).Formula = "12.900 al 21.620"
      .Cells(29, 1).Formula = "21.620 al 35.570"
      .Cells(49, 1).Formula = "35.570 al 65.140"
      mi = 10
      Set mRec = mObjRAcc.oTablaNotNull("ColisionContra", "order by 1")
      Do While Not mRec.EOF
         .Cells(mi, 1).Formula = " " & mRec!CodColision & "-" & mRec!descripcion
         .Cells(mi + 20, 1).Formula = " " & mRec!CodColision & "-" & mRec!descripcion
         .Cells(mi + 40, 1).Formula = " " & mRec!CodColision & "-" & mRec!descripcion
         .Cells(mi, 2).Formula = mObjRAcc.iCountColisContra(mFecha1, mFecha2, "12.9", "21.62", mRec!CodColision)
         .Cells(mi + 20, 2).Formula = mObjRAcc.iCountColisContra(mFecha1, mFecha2, "21.62", "35.57", mRec!CodColision)
         .Cells(mi + 40, 2).Formula = mObjRAcc.iCountColisContra(mFecha1, mFecha2, "35.57", "65.15", mRec!CodColision)
         If mRec!CodColision = "12" Then
            .Cells(mi, 2).Formula = Val(.Cells(mi, 2).Formula) + mObjRAcc.iCountColisContra(mFecha1, mFecha2, "12.9", "21.62", "")
            .Cells(mi + 20, 2).Formula = Val(.Cells(mi + 20, 2).Formula) + mObjRAcc.iCountColisContra(mFecha1, mFecha2, "21.62", "35.57", "")
            .Cells(mi + 40, 2).Formula = Val(.Cells(mi + 40, 2).Formula) + mObjRAcc.iCountColisContra(mFecha1, mFecha2, "35.57", "65.15", "")
         End If
         mi = mi + 1
         mRec.MoveNext
      Loop
      mRec.Close
   End With
End Sub

Private Sub sPlanilla9_10()
Dim mDias(7) As String
Dim mj As Integer
mDias(1) = "Domingo"
mDias(2) = "Lunes"
mDias(3) = "Martes"
mDias(4) = "Miércoles"
mDias(5) = "Jueves"
mDias(6) = "Viernes"
mDias(7) = "Sábado"
   
   With XLS
      For mj = 9 To 10
         .Worksheets(mj).Select
         .Worksheets(mj).Name = "Planilla" & mj
         .Columns("B:H").ColumnWidth = 15
         .Range("B4:B200").Select
         .Selection.NumberFormat = "dd-mm-yyyy"
         .Cells(2, 1).Formula = "Mes:"
         .Cells(2, 2).Formula = " " & Mid(Combo1(0).Text, 6, 3) & "-" & Right(Combo1(1).Text, 2)
         .Cells(3, 1).Formula = "Nro Orden"
         .Cells(3, 2).Formula = "Fecha"
         .Cells(3, 3).Formula = "Día"
         .Cells(3, 4).Formula = "Hora"
         .Cells(3, 5).Formula = "Progresiva"
         .Cells(3, 6).Formula = "Vehículos"
         .Cells(3, 7).Formula = "Fallecidos"
         .Cells(3, 8).Formula = "Heridos Leves"
         .Cells(3, 9).Formula = "Heridos Graves"
         mi = 4
         If mj = 9 Then
            Set mRec = mObjRAcc.oFichasFechaProgr(mFecha1, mFecha2, "", "", "lugaraccid<>'09'")
         Else
            Set mRec = mObjRAcc.oFichasFechaProgr(mFecha1, mFecha2, "", "", "lugaraccid='09'")
         End If
         Do While Not mRec.EOF
            .Cells(mi, 1).Formula = mRec!NroOrden
            .Cells(mi, 2).Formula = mRec!Fecha
            .Cells(mi, 3).Formula = mDias(WeekDay(mRec!Fecha))
            .Cells(mi, 4).Formula = mRec!hora
            .Cells(mi, 5).Formula = mRec!Progresiva
            .Cells(mi, 6).Formula = mObjRAcc.iTotalVehiculosNroOrden(mRec!NroOrden)
            .Cells(mi, 7).Formula = mObjRAcc.iTotalMuertosNroOrden(mRec!NroOrden)
            .Cells(mi, 8).Formula = mObjRAcc.iTotalHeridosTipoNroOrden(mRec!NroOrden, True)
            .Cells(mi, 9).Formula = mObjRAcc.iTotalHeridosTipoNroOrden(mRec!NroOrden, False)
            mi = mi + 1
            mRec.MoveNext
         Loop
         mRec.Close
      Next
   End With
End Sub

Private Sub sServGruas()
Dim mDatosServ(8) As Integer
   For mi = 1 To 8
      mDatosServ(mi) = 0
   Next
   Set mRec = mObjRN.oCountServMes(Left(Combo1(0).Text, 2), Combo1(1).Text)
   Do While Not mRec.EOF
      Select Case mRec!codserv1
         Case "AN"
            mDatosServ(1) = NVL(mRec!Total, 0)
         Case "AM"
            mDatosServ(2) = NVL(mRec!Total, 0)
         Case "AC"
            mDatosServ(3) = NVL(mRec!Total, 0)
         Case "ARG", "COL", "SAC", "S1", "S2"
            mDatosServ(4) = mDatosServ(4) + NVL(mRec!Total, 0)
         Case "CN"
            mDatosServ(5) = NVL(mRec!Total, 0)
         Case "CM"
            mDatosServ(6) = NVL(mRec!Total, 0)
         Case "CC"
            mDatosServ(7) = NVL(mRec!Total, 0)
         Case "CRG"
            mDatosServ(8) = NVL(mRec!Total, 0)
      End Select
      mRec.MoveNext
   Loop
   mRec.Close
   XLS.Cells(13, 2).Formula = mDatosServ(1) + mDatosServ(2) + mDatosServ(3) + mDatosServ(4) + mDatosServ(5) + mDatosServ(6) + mDatosServ(7) + mDatosServ(8)
   For mi = 1 To 4
      XLS.Cells(mi + 15, 2).Formula = mDatosServ(mi)
      XLS.Cells(mi + 21, 2).Formula = mDatosServ(mi + 4)
   Next
End Sub

Private Sub sDepejeVias()
Dim mParam As String
Dim mCodTipo As String
   mCodTipo = "OBJ"
   For mi = 29 To 31 Step 2
      mParam = ""
      Set mRec = mObjRN.oTabla("otros", "where codtipotro='" & mCodTipo & "'")
      Do While Not mRec.EOF
         mParam = mParam & "'" & mRec!Codigo & "',"
         mRec.MoveNext
      Loop
      mRec.Close
      mParam = Mid(mParam, 1, Len(mParam) - 1)
      XLS.Cells(mi, 2).Formula = mObjRN.iCountTareas(Left(Combo1(0).Text, 2), Combo1(1).Text, "E','R", mParam)
      mCodTipo = "ANI"
   Next
End Sub

Private Sub sAccidentes()
   'Cant. de Accidentes
   Set mRec = mObjRAcc.oAccidTraza(mFecha1, mFecha2, True)
   mAccTotal(1) = mRec!Total 'Accidentes Calzada Principal
   mRec.Close
   Set mRec = mObjRAcc.oAccidColect(mFecha1, mFecha2)
   mAccTotal(2) = mRec!Total 'Accidentes Colectora
   mRec.Close
   Set mRec = mObjRAcc.oAccidTroncal(mFecha1, mFecha2)
   mAccTotal(3) = mRec!Total 'Accidentes Troncal
   mRec.Close
   XLS.Cells(33, 2).Formula = mAccTotal(1) + mAccTotal(2) + mAccTotal(3)
   
   'HERIDOS
   XLS.Cells(35, 2).Formula = mObjRAcc.iCountHeridos(mMes, mAnio, True, False, "")
   XLS.Cells(36, 2).Formula = mObjRAcc.iCountHeridos(mMes, mAnio, False, True, "")
   'MUERTOS
   XLS.Cells(37, 2).Formula = mObjRAcc.iTotalMuertosCodigo(mFecha1, mFecha2, "(A.fallecio <> '' or A.codestado in ('04','05'))")
   
   'Ambulancias
   XLS.Cells(38, 2).Formula = mObjRAcc.iCountMedTrasl(mMes, mAnio, "05")
   XLS.Cells(39, 2).Formula = mObjRAcc.iCountMedTrasl(mMes, mAnio, "06")
   
   '*** ACCIDENTES POR CAUSAS
   XLS.Cells(42, 2).Formula = mObjRAcc.iCantCausaFechas(mFecha1, mFecha2, "codtipoficha='01'")
   'peaton/ciclista
   XLS.Cells(43, 2).Formula = mObjRAcc.iCantCausaFechas(mFecha1, mFecha2, "accidotro in ('02','03')")
   'vehículos
   XLS.Cells(44, 2).Formula = mObjRAcc.iCantCausaFechas(mFecha1, mFecha2, "causavehic<>''")
   'Animales
   XLS.Cells(45, 2).Formula = mObjRAcc.iCantCausaFechas(mFecha1, mFecha2, "codcoliscontra1='07'")
   'Mat. s/calzada
    XLS.Cells(46, 2).Formula = mObjRAcc.iCantCausaFechas(mFecha1, mFecha2, "codcoliscontra1='09'")
   If Val(XLS.Cells(42, 2).Formula) <> 0 Then
      XLS.Cells(42, 2).Formula = Val(XLS.Cells(42, 2).Formula) - (Val(XLS.Cells(43, 2).Formula) + Val(XLS.Cells(44, 2).Formula) + Val(XLS.Cells(45, 2).Formula) + Val(XLS.Cells(46, 2).Formula))
   End If

   'CANTIDAD DE VEHICULOS INVOLUCRADOS TIPO DE VEHIC
   XLS.Cells(49, 2).Formula = mObjRAcc.iCantVehicFechas(mFecha1, mFecha2, "'06','07','08','09','10','11','12','13','14','16'") 'camiones y otros
   XLS.Cells(50, 2).Formula = mObjRAcc.iCantVehicFechas(mFecha1, mFecha2, "'04','05','15'") 'autos
   XLS.Cells(51, 2).Formula = mObjRAcc.iCantVehicFechas(mFecha1, mFecha2, "'02','03'") 'motos
   XLS.Cells(52, 2).Formula = mObjRAcc.iCantVehicFechas(mFecha1, mFecha2, "'01'") 'bicicletas
   
   'CANTIDAD DE VEHICULOS INVOLUCRADOS TIPO DE VEHIC EN CALZADA PRINCIPAL
   
   XLS.Cells(49, 4).Formula = mObjRAcc.iCantVehicFechasConParamTraza(mFecha1, mFecha2, "'06','07','08','09','10','11','12','13','14','16'", True) 'camiones y otros
   XLS.Cells(50, 4).Formula = mObjRAcc.iCantVehicFechasConParamTraza(mFecha1, mFecha2, "'04','05','15'", True) 'autos
   XLS.Cells(51, 4).Formula = mObjRAcc.iCantVehicFechasConParamTraza(mFecha1, mFecha2, "'02','03'", True) 'motos
   XLS.Cells(52, 4).Formula = mObjRAcc.iCantVehicFechasConParamTraza(mFecha1, mFecha2, "'01'", True) 'bicicletas
   
   'CANTIDAD DE VEHICULOS INVOLUCRADOS TIPO DE VEHIC EN COLECTORA
   
   XLS.Cells(49, 5).Formula = mObjRAcc.iCantVehicFechasConParamTraza(mFecha1, mFecha2, "'06','07','08','09','10','11','12','13','14','16'", False) 'camiones y otros
   XLS.Cells(50, 5).Formula = mObjRAcc.iCantVehicFechasConParamTraza(mFecha1, mFecha2, "'04','05','15'", False) 'autos
   XLS.Cells(51, 5).Formula = mObjRAcc.iCantVehicFechasConParamTraza(mFecha1, mFecha2, "'02','03'", False) 'motos
   XLS.Cells(52, 5).Formula = mObjRAcc.iCantVehicFechasConParamTraza(mFecha1, mFecha2, "'01'", False) 'bicicletas
   
   
   
End Sub

Private Sub sCabecera1()
   With XLS
      .WorkBooks.Add
      .Sheets.Add '4
      .Sheets.Add '5
      .Sheets.Add '6
      .Sheets.Add '7
      .Sheets.Add '8
      .Sheets.Add '9
      .Sheets.Add '10
      .Worksheets(1).Select
      .Worksheets(1).Name = "Planilla1"
      .Columns("A:A").ColumnWidth = 45
      .Cells(13, 1).Formula = "SERVICIOS"
      .Cells(15, 1).Formula = "Automóviles y camionetas"
      .Cells(16, 1).Formula = "* Neumáticos"
      .Cells(17, 1).Formula = "* Mecánica y otros"
      .Cells(18, 1).Formula = "* Combustible"
      .Cells(19, 1).Formula = "* Remolque por grúas contratadas"
      .Cells(21, 1).Formula = "CAMIONES"
      .Cells(22, 1).Formula = "* Neumáticos"
      .Cells(23, 1).Formula = "* Mecánica y otros"
      .Cells(24, 1).Formula = "* Combustible"
      .Cells(25, 1).Formula = "* Remolque"
      .Cells(29, 1).Formula = "Despeje de vía por material en calzada"
      .Cells(31, 1).Formula = "Despeje de vía por animales sueltos"
      .Cells(33, 1).Formula = "ACCIDENTES DE TRÁNSITO"
      .Cells(35, 1).Formula = "* Heridos leves"
      .Cells(36, 1).Formula = "* Heridos graves"
      .Cells(37, 1).Formula = "* Fallecidos"
      .Cells(38, 1).Formula = "* Trasl./atenc. por ambulancias del servicio"
      .Cells(39, 1).Formula = "* Otros traslados"
      .Cells(41, 1).Formula = "CAUSAS"
      .Cells(42, 1).Formula = "* Counductor"
      .Cells(43, 1).Formula = "* Peatón/ciclista"
      .Cells(44, 1).Formula = "* Vehículo"
      .Cells(45, 1).Formula = "* Animal"
      .Cells(46, 1).Formula = "* Material en calzada"
      .Cells(48, 1).Formula = "VEHICULOS COMPROMETIDOS"
      .Cells(48, 4).Formula = "C.Prinicipal"
      .Cells(48, 5).Formula = "Colectora"
      .Cells(49, 1).Formula = "* Camiones"
      .Cells(50, 1).Formula = "* Automóvile"
      .Cells(51, 1).Formula = "* Motos"
      .Cells(52, 1).Formula = "* Bicicletas"
      .Columns("B:B").ColumnWidth = 12
      .Cells(12, 2).Formula = " " & Mid(Combo1(0).Text, 6, 3) & "-" & Right(Combo1(1).Text, 2)
   End With
End Sub

Private Sub sCabecera2()
   With XLS
      .Worksheets(2).Select
      .Worksheets(2).Name = "Planilla2"
      .Columns("A:A").ColumnWidth = 45
      .Cells(14, 1).Formula = "Incidentes"
      .Cells(15, 1).Formula = "Accidentes en calzada principal"
      .Cells(16, 1).Formula = "Accidentes en colectora"
      .Cells(17, 1).Formula = "Accidentes en Peajes Troncales"
      .Cells(18, 1).Formula = "Accidentes Totales"
      .Cells(19, 1).Formula = "Intervenciones de patrullas"
      .Cells(20, 1).Formula = "Víctimas fatales peatones/ciclistas en calzada"
      .Cells(21, 1).Formula = "Víctimas fatales calzada principal"
      
      .Cells(23, 1).Formula = "Servicios GRUA LIVIANA"
      .Cells(24, 1).Formula = "Servicios GRUA PESADA"
      .Cells(25, 1).Formula = "Servicios Bomberos"
      .Cells(26, 1).Formula = "Servicios Emergencia Médica VITTAL"
      
      .Cells(28, 1).Formula = "Robo en Cabina de Peaje"
      .Cells(29, 1).Formula = "Robo a Usuario en tránsito"
      
      .Cells(31, 1).Formula = "Operativos Gendarmería Nacional"
      .Cells(32, 1).Formula = "Operativos Policía Pcia. Bs. As."
      .Cells(33, 1).Formula = "Op. Policía Seg. Vial Control Peaje"
      
      .Cells(37, 1).Formula = "Objetos retirados de calzada"
      .Cells(38, 1).Formula = "Incidentes en calzada con animales"
      .Cells(12, 2).Formula = " " & Mid(Combo1(0).Text, 6, 3) & "-" & Right(Combo1(1).Text, 2)
   End With
End Sub

Private Sub sCabecera3()
Dim mj As Integer
   With XLS
      '.WorkBooks.Add
      .Worksheets(3).Select
      .Worksheets(3).Name = "Planilla3"
      .Columns("A:A").ColumnWidth = 35
      .Columns("B:J").ColumnWidth = 7
      .Rows("6:6").RowHeight = 23
      .Cells(10, 1).Formula = "Discriminación por hora del día"
      .Cells(11, 1).Formula = "Intervenciones VITTAL - " & Mid(Combo1(0).Text, 6, 3) & "-" & Right(Combo1(1).Text, 2)
      .Cells(12, 1).Formula = "Horario en minutos"
      .Cells(13, 1).Formula = "Código Rojo"
      .Cells(14, 1).Formula = "Código Amarillo"
      .Cells(15, 1).Formula = "Código Verde"
      .Cells(17, 1).Formula = "% Código Rojo"
      .Cells(18, 1).Formula = "Acumulado Cod. Rojo"
      .Cells(19, 1).Formula = "% Código Amarillo"
      .Cells(20, 1).Formula = "Acumulado Cod. Amarillo"
      
      .Cells(39, 1).Formula = "Servicios de ambulancia VITTAL Acumulado"
      .Cells(40, 1).Formula = "Código Rojo"
      .Cells(41, 1).Formula = "Total Mes"
      .Cells(42, 1).Formula = "% arribos hasta 15 minutos"
      
      .Cells(45, 1).Formula = "Código Amarillo"
      .Cells(46, 1).Formula = "Total Mes"
      .Cells(47, 1).Formula = "% arribos hasta 24 minutos"
      
      .Cells(40, 2).Formula = " " & Mid(Combo1(0).Text, 6, 3) & "-" & Right(Combo1(1).Text, 2)
      .Cells(45, 2).Formula = " " & Mid(Combo1(0).Text, 6, 3) & "-" & Right(Combo1(1).Text, 2)
      mj = 0
      For mi = 2 To 9
         .Cells(13, mi).Formula = " " & mj & " a " & (mj + 3)
         mj = mj + 3
      Next
      .Cells(13, 10).Formula = " > a 24"
      .Cells(12, 11).Formula = " TOTAL"
      
      'porcentajes
      .Range("B17:J20").Select
      .Selection.Style = "Percent"
   End With
End Sub

Private Sub sCabecera4()
Dim mj As Integer
   With XLS
      '.WorkBooks.Add
      '.Sheets.Add
      .Worksheets(4).Select
      .Worksheets(4).Name = "Planilla4"
      .Columns("A:A").ColumnWidth = 35
      .Columns("B:J").ColumnWidth = 7
      .Rows("6:6").RowHeight = 23
      .Cells(10, 1).Formula = "Intervenciones Patrullas - " & Mid(Combo1(0).Text, 6, 3) & "-" & Right(Combo1(1).Text, 2)
      .Cells(11, 1).Formula = "Horario en minutos"
      .Cells(12, 1).Formula = "Accidentes"
      .Cells(13, 1).Formula = "Incidentes"
      .Cells(15, 1).Formula = "% Accidentes"
      .Cells(16, 1).Formula = "Acumulado Accidentes"
      .Cells(17, 1).Formula = "% Incidentes"
      .Cells(18, 1).Formula = "Acumulado Incidentes"
      
      .Cells(35, 1).Formula = "Arribos de PAtrullas Acumulados"
      .Cells(36, 1).Formula = "Accidentes"
      .Cells(37, 1).Formula = "Total Mes"
      .Cells(38, 1).Formula = "% arribos hasta 15 minutos"
      
      .Cells(41, 1).Formula = "Incidentes"
      .Cells(42, 1).Formula = "Total Mes"
      .Cells(43, 1).Formula = "% arribos hasta 24 minutos"
      
      .Cells(36, 2).Formula = " " & Mid(Combo1(0).Text, 6, 3) & "-" & Right(Combo1(1).Text, 2)
      .Cells(41, 2).Formula = " " & Mid(Combo1(0).Text, 6, 3) & "-" & Right(Combo1(1).Text, 2)
      mj = 0
      For mi = 2 To 9
         .Cells(11, mi).Formula = " " & mj & " a " & (mj + 3)
         mj = mj + 3
      Next
      .Cells(11, 10).Formula = " > a 24"
      .Cells(10, 11).Formula = " TOTAL"
      
      'porcentajes
      .Range("B15:J18").Select
      .Selection.Style = "Percent"
      .Range("B38").Select
      .Selection.Style = "Percent"
      .Range("B43").Select
      .Selection.Style = "Percent"
   End With
End Sub

Private Sub sCabecera5()
Dim mj As Integer
   With XLS
      .Worksheets(5).Select
      .Worksheets(5).Name = "Planilla5"
      .Columns("A:A").ColumnWidth = 35
      .Cells(3, 1).Formula = "CALZADA PPAL."
      .Cells(17, 1).Formula = "COLECTORA"
      For mi = 0 To 14 Step 14
         .Cells(3 + mi, 2).Formula = " " & Mid(Combo1(0).Text, 6, 3) & "-" & Right(Combo1(1).Text, 2)
         .Cells(4 + mi, 1).Formula = "Accidentes"
         .Cells(5 + mi, 1).Formula = "Accidentes con víctimas"
         .Cells(6 + mi, 1).Formula = "Víctimas mortales conduc/acomp."
         .Cells(7 + mi, 1).Formula = "Heridos conduc/acomp."
         .Cells(8 + mi, 1).Formula = "Víctimas Mortales Peatones"
         .Cells(9 + mi, 1).Formula = "Heridos Peatones"
         .Cells(10 + mi, 1).Formula = "Víctimas Ciclistas"
         .Cells(11 + mi, 1).Formula = "Heridos Ciclistas"
         .Cells(12 + mi, 1).Formula = "TOTAL Víctimas Mortales"
         .Cells(13 + mi, 1).Formula = "TOTAl Víctimas Heridos"
         .Cells(14 + mi, 1).Formula = "Víctimas mortales moto."
         .Cells(15 + mi, 1).Formula = "Heridos moto."
         .Cells(16 + mi, 1).Formula = "Víctimas Sin Lesiones."
      Next
      .Cells(12, 2).Formula = "=sum(B6+B8+B10)"
      .Cells(13, 2).Formula = "=sum(B7+B9+B11)"
      .Cells(26, 2).Formula = "=sum(B20+B22+B24)"
      .Cells(27, 2).Formula = "=sum(B21+B23+B25)"
   End With
End Sub

Private Sub sCabecera6()
   With XLS
      .Worksheets(6).Select
      .Worksheets(6).Name = "Planilla6"
      .Columns("A:A").ColumnWidth = 15
      .Columns("B:B").ColumnWidth = 15
      .Columns("C:C").ColumnWidth = 15
      .Columns("G:J").ColumnWidth = 10
      .Columns("K:K").ColumnWidth = 20
      .Columns("L:N").ColumnWidth = 10
      .Columns("O:O").ColumnWidth = 20
      .Columns("P:P").ColumnWidth = 22
      .Columns("Q:Q").ColumnWidth = 22
      .Columns("R:R").ColumnWidth = 22
      .Columns("S:S").ColumnWidth = 22
      .Columns("T:T").ColumnWidth = 22
      .Columns("U:U").ColumnWidth = 7
      .Columns("V:V").ColumnWidth = 7
      .Columns("W:W").ColumnWidth = 7
      .Columns("X:X").ColumnWidth = 7
      .Columns("Y:Y").ColumnWidth = 7
      .Columns("Z:Z").ColumnWidth = 7
      
      .Range("C6:C300").Select
      .Selection.NumberFormat = "dd-mm-yyyy"
      .Cells(3, 1).Formula = "Mes:"
      .Cells(3, 2).Formula = " " & Mid(Combo1(0).Text, 6, 3) & "-" & Right(Combo1(1).Text, 2)
      .Cells(5, 1).Formula = "Cód.Alfa"
      .Cells(5, 2).Formula = "Nro.Ficha"
      .Cells(5, 3).Formula = "Fecha"
      .Cells(5, 4).Formula = "Aviso"
      .Cells(5, 5).Formula = "Prog."
      .Cells(5, 6).Formula = "Calz."
      .Cells(5, 7).Formula = "Arribo"
      .Cells(5, 8).Formula = "Herido"
      .Cells(5, 9).Formula = "Fallec"
      .Cells(5, 10).Formula = "Cant. Vehic. Involucr."
      .Cells(5, 11).Formula = "Tipo"
      .Cells(5, 12).Formula = "Posible Causa"
      .Cells(5, 13).Formula = "Liberación de Calzada"
      .Cells(5, 14).Formula = "Condiciones climáticas"
      .Cells(5, 15).Formula = "Observaciones"
      .Cells(5, 16).Formula = "Bomberos"
      .Cells(5, 17).Formula = "Gendarería"
      .Cells(5, 18).Formula = "Policía Científica"
      .Cells(5, 19).Formula = "Policía Jurisdicción"
      .Cells(5, 20).Formula = "Policía Vial"
      .Cells(5, 21).Formula = "Bicicletas"
      .Cells(5, 22).Formula = "Motos"
      .Cells(5, 23).Formula = "Autos"
      .Cells(5, 24).Formula = "Camiones/Varios"
      .Cells(5, 25).Formula = "Carril 1"
      .Cells(5, 26).Formula = "Carril 2"
      .Cells(5, 27).Formula = "Carril 3"
      .Cells(5, 28).Formula = "Carril 4"
      .Cells(5, 29).Formula = "Carril 5"
      .Cells(5, 30).Formula = "Carril 6"
      .Cells(5, 31).Formula = "Carril 7"
      .Cells(5, 32).Formula = "Carril 8"
      .Cells(5, 33).Formula = "Carril 9"
      .Cells(5, 34).Formula = "Banq. Int"
      .Cells(5, 35).Formula = "Banq. Ext"
      .Cells(5, 36).Formula = "Ramal"
   End With
End Sub

Private Sub sCabecera7()
   With XLS
      .Worksheets(7).Select
      .Worksheets(7).Name = "Planilla7"
      .Columns("A:A").ColumnWidth = 15
      .Cells(2, 1).Formula = "Víctimas fatales según edad."
      .Cells(4, 1).Formula = "Mes:"
      .Cells(4, 2).Formula = " " & Mid(Combo1(0).Text, 6, 3) & "-" & Right(Combo1(1).Text, 2)
      .Cells(5, 1).Formula = "Menores de 18"
      .Cells(6, 1).Formula = "Entre 19 y 25"
      .Cells(7, 1).Formula = "Entre 26 y 45"
      .Cells(8, 1).Formula = "Mayor a 45"
      .Cells(9, 1).Formula = "TOTAL"
      .Cells(10, 1).Formula = "% Menores de 18"
      .Cells(11, 1).Formula = "% Entre 19 y 25"
      .Cells(12, 1).Formula = "% Entre 26 y 45"
      .Cells(13, 1).Formula = "% Mayor a 45"
   End With
End Sub
