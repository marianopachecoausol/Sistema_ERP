VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form RAcc10_frm 
   BackColor       =   &H00CECECE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Módulo de Reportes"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   5985
   Begin VB.CheckBox Check1 
      BackColor       =   &H00CECECE&
      Caption         =   "para Legales"
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   1680
      Visible         =   0   'False
      Width           =   2535
   End
   Begin Crystal.CrystalReport CReport1 
      Left            =   300
      Top             =   300
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00DFE7DE&
      Caption         =   "Todas"
      Height          =   255
      Index           =   2
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1260
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   8
      Top             =   2400
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C1DBD8&
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
      Height          =   435
      Index           =   0
      Left            =   3240
      MaskColor       =   &H8000000B&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2340
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   3180
      MaxLength       =   10
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1200
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Final"
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
      Left            =   3240
      TabIndex        =   6
      Top             =   960
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nro Ficha"
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
      Left            =   1680
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Inicial"
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
      Left            =   1680
      TabIndex        =   1
      Top             =   960
      Width           =   1110
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   5955
   End
End
Attribute VB_Name = "RAcc10_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mObjAcc As New clAccess

Private Sub Form_Load()
   sAlinearForm Me
   Check1.Value = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set mObjAcc = Nothing
   ShowMenu 2, True, False
End Sub

Private Sub Command1_Click(Index As Integer)
Dim mObj As New clRAcc
Dim mRec As New ADODB.Recordset

   Select Case Command1(Index).Caption
      Case "Buscar"
         If Fecha_ok(Text1(0).Text) And Fecha_ok(Text1(1).Text) Then
            If DateDiff("d", Text1(0).Text, Text1(1).Text) >= 0 Then
               sMsgEspere Me, "Buscando... espere un momento", True
               Set mRec = mObj.oNroOrdenFechas(Text1(0).Text, Text1(1).Text)
               If Not mRec.EOF Then
                  Combo1.Clear
                  sOcultarObj True
                  Do While Not mRec.EOF
                     Combo1.AddItem mRec!NroOrden & " - " & mRec!Fecha
                     mRec.MoveNext
                  Loop
               Else
                  MsgBox "Consulta sin Resultados", vbInformation, sMessage
               End If
               sMsgEspere Me, "", False
               mRec.Close
            Else
               MsgBox "Fecha Inicial mayor a la Final", vbCritical, sMessage
            End If
         End If
         
      Case "Volver"
         sOcultarObj False
         Combo1.Clear
         
      Case "Generar"
         If Combo1.ListIndex > -1 Then
            sMsgEspere Me, "Generando el informe...", True
            mObjAcc.mBorrarAuxi "\RegAccidentes\FichaAccid", "Auxi"
            mObjAcc.mBorrarAuxi "\RegAccidentes\FichaAccid", "Auxi2"
            mObjAcc.mBorrarAuxi "\RegAccidentes\FichaAccid", "Auxi3"
            mObjAcc.mBorrarAuxi "\RegAccidentes\FichaAccid", "Auxi4"
            mObjAcc.mBorrarAuxi "\RegAccidentes\FichaAccid", "Auxi5"
            If DateDiff("d", Right(Combo1.Text, 10), "01/03/2008") > 0 Then
               sVerFichaOlder  'para fichas viejas
            Else
               sVerReporte (Check1)
            End If
            sMsgEspere Me, "", False
         Else
            MsgBox "Elegir una ficha", vbExclamation, sMessage
         End If
         
      Case "Todas"
         sMsgEspere Me, "Buscando... espere un momento", True
         Set mRec = mObj.oTabla("Ficha", " order by NroOrden, Fecha")
         If Not mRec.EOF Then
            Combo1.Clear
            sOcultarObj True
            Do While Not mRec.EOF
               Combo1.AddItem mRec!NroOrden & " - " & mRec!Fecha
               mRec.MoveNext
            Loop
         Else
            MsgBox "Consulta sin Resultados", vbInformation, sMessage
         End If
         mRec.Close
         sMsgEspere Me, "", False
      
      Case "Salir"
         Unload Me
   End Select
   
   Set mObj = Nothing
   Set mRec = Nothing
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = fDateKeyPress(Text1(Index), KeyAscii)
End Sub

Private Sub sOcultarObj(ByVal pFlag As Boolean)
   Combo1.Visible = pFlag
   Text1(0).Visible = Not pFlag
   Text1(1).Visible = Not pFlag
   Label1(1).Visible = Not pFlag
   Label1(2).Visible = Not pFlag
   Label1(3).Visible = pFlag
   Command1(2).Visible = Not pFlag
   Check1.Visible = pFlag
   
   If pFlag Then
      Command1(0).Caption = "Generar"
      Command1(1).Caption = "Volver"
   Else
      Command1(0).Caption = "Buscar"
      Command1(1).Caption = "Salir"
      Text1(0).Text = ""
      Text1(1).Text = ""
   End If
End Sub

Private Sub sVerFichaOlder()
Dim mObj As New clRAcc
Dim mRec As New ADODB.Recordset
Dim mData As Database
Dim mDatos(8, 12) As String
Dim mCarril(7) As String
Dim mValues As String
Dim mI As Integer
Dim mJ As Integer
Dim mDescrMid As Integer
   mValues = ""
   Set mData = OpenDatabase(App.Path & "\RegAccidentes\FichaAccid.mdb")
   mData.Execute "create table Auxi (nroficha text, patru text, police text, mov text, fecha text, hora_av text, " _
      & "hora_ar text, km text, tramo text, inter text, carril_1 text, carril_2 text, carril_3 text, carril_4 text, " _
      & "carril_5 text, carril_6 text, carril_7 text, autoridad text, cov text, cc_1 text, cc_2 text, cc_3 text, " _
      & "otroacc text, sent text, lugaracc text, clima text, ecalz text, ebaq text, dhor text, dvert text, ilum text, " _
      & "ccond_1 text, ccond_2 text, ccond_3 text, cveh text, danos text, bombname text, bombdepto text, ambuname text, " _
      & "ambudep text, poliname text, polidep text, gendname text, genddep text, grua text, gruaempr text, policienname text, " _
      & "policiendep text, otraauto text, foto text)"
      
   Set mRec = mObj.oTablaNroOrden("Ficha", Left(Combo1.Text, 5), " order by fecha")
   If Not mRec.EOF Then
      mValues = "'" & mRec.Fields(0) & "','" & mObj.sTablaDescr("Patrullero", "codpatrullero='" & Trim(mRec.Fields(1)) & "'", 1) & "',"
      For mI = 2 To 9
         mValues = mValues & "'" & mRec.Fields(mI) & "', "
      Next
      'carriles
      For mJ = 1 To 7
         mCarril(mJ) = " "
      Next
      For mI = 1 To Len(mRec!carril) Step 2
         mJ = Mid(mRec!carril, mI, 2)
         If mJ <= 7 Then
            mCarril(mJ) = "X"
         End If
      Next
      For mI = 1 To 7
         mValues = mValues & "'" & mCarril(mI) & "', "
      Next
      For mI = 11 To 43
         mValues = mValues & "'" & mRec.Fields(mI) & "', "
      Next
      mValues = Mid(mValues, 1, Len(mValues) - 2)
   End If
   mRec.Close
   On Error Resume Next
   mData.Execute "insert into Auxi values (" & mValues & ")"
   mValues = ""
   If Err.Description <> "" Then
      MsgBox "Error en primer Auxi" & Chr(13) & "Descripción: " & Chr(13) & Err.Description
   End If
  
   'Descripción de una ficha
   mData.Execute "create table Auxi4 (A1 text, A2 text, A3 text, A4 text, A5 text, A6 text, A7 text, A8 text," _
      & " B1 text, B2 text, B3 text, B4 text, B5 text, B6 text, B7 text, B8 text)"
   mData.Execute "create table Auxi5 (C1 text, C2 text, C3 text, C4 text, C5 text, C6 text, C7 text, C8 text, " _
      & " E1 text, E2 text, E3 text, E4 text, E5 text, E6 text, E7 text, E8 text)"
         
   For mI = 1 To 4
      For mJ = 1 To 8
         mDatos(mI, mJ) = ""
      Next
   Next
   Set mRec = mObj.oTablaNroOrden("fichadescr", Left(Combo1.Text, 5), "")
   If Not mRec.EOF Then
      For mJ = 1 To 4
         mDescrMid = 1
         For mI = 1 To 8
            If Len(mRec.Fields(mJ)) >= (mI * 110) Then
               mDatos(mJ, mI) = Mid(mRec.Fields(mJ), mDescrMid, 110)
            Else
               If Mid(mRec.Fields(mJ), mDescrMid, 110) <> "" Then
                  mDatos(mJ, mI) = Mid(mRec.Fields(mJ), mDescrMid, 110)
               Else     'nuevo
                  mDatos(mJ, mI) = ""   'nuevo
               End If
            End If
            mDescrMid = mDescrMid + 110
         Next
      Next
   End If
   mRec.Close
   For mI = 1 To 2
      For mJ = 1 To 8
         fReplace (mDatos(mI, mJ))
         mDatos(mI, mJ) = Replace(mDatos(mI, mJ), Chr(13) & Chr(10), " ")
         mValues = mValues & "'" & mDatos(mI, mJ) & "', "
      Next
   Next
   mValues = Mid(mValues, 1, Len(mValues) - 2)
   mData.Execute "insert into Auxi4 values (" & mValues & ")"
   mValues = ""
   For mI = 3 To 4
      For mJ = 1 To 8
         fReplace (mDatos(mI, mJ))
         mDatos(mI, mJ) = Replace(mDatos(mI, mJ), Chr(13) & Chr(10), " ")
         mValues = mValues & "'" & mDatos(mI, mJ) & "', "
      Next
   Next
   mValues = Mid(mValues, 1, Len(mValues) - 2)
   mData.Execute "insert into Auxi5 values (" & mValues & ")"
   
   'Vehiculos Involucradas
   mData.Execute "create table Auxi2 (letra_1 text, tipoveh_1 text, marca_1 text, modelo_1 text, pat_1 text, cond_1 text, doc_1 text, tel_1 text, domic_1 text, seguro_1 text, " _
   & "letra_2 text, tipoveh_2 text, marca_2 text, modelo_2 text, pat_2 text, cond_2 text, doc_2 text, tel_2 text, domic_2 text, seguro_2 text," _
   & "letra_3 text, tipoveh_3 text, marca_3 text, modelo_3 text, pat_3 text, cond_3 text, doc_3 text, tel_3 text, domic_3 text, seguro_3 text," _
   & "letra_4 text, tipoveh_4 text, marca_4 text, modelo_4 text, pat_4 text, cond_4 text, doc_4 text, tel_4 text, domic_4 text, seguro_4 text," _
   & "letra_5 text, tipoveh_5 text, marca_5 text, modelo_5 text, pat_5 text, cond_5 text, doc_5 text, tel_5 text, domic_5 text, seguro_5 text," _
   & "letra_6 text, tipoveh_6 text, marca_6 text, modelo_6 text, pat_6 text, cond_6 text, doc_6 text, tel_6 text, domic_6 text, seguro_6 text," _
   & "letra_7 text, tipoveh_7 text, marca_7 text, modelo_7 text, pat_7 text, cond_7 text, doc_7 text, tel_7 text, domic_7 text, seguro_7 text," _
   & "letra_8 text, tipoveh_8 text, marca_8 text, modelo_8 text, pat_8 text, cond_8 text, doc_8 text, tel_8 text, domic_8 text, seguro_8 text)"
   
   For mI = 1 To 9
      For mJ = 1 To 11
         mDatos(mI, mJ) = ""
      Next
   Next
   Set mRec = mObj.oTablaNroOrden("VehiculosInvolucr", Left(Combo1.Text, 5), "order by letra")
   If Not mRec.EOF Then
      mI = 1
      Do While Not mRec.EOF
         If mI <= 8 Then
            mDatos(mI, 1) = mRec!L
            mDatos(mI, 2) = mRec!CodTipoVehic
            mDatos(mI, 3) = mObj.sTablaDescr("Marca", " codtipovehic='" & mRec!CodTipoVehic & "' and codmarca='" & mRec!CodMarca & "'", 2)
            mDatos(mI, 4) = mRec!modelo
            mDatos(mI, 5) = mRec!Dominio
            mDatos(mI, 6) = mRec!conductor 'falta conductor
            mDatos(mI, 7) = mObj.sTablaDescr("TipoDocu", " codtipodocu='" & mRec!codtipodoc & "'", 1) & " - " & mRec!nrodocu
            mDatos(mI, 8) = mRec!Telefono 'telefono
            mDatos(mI, 9) = mRec!domicilio 'domicilio
            mDatos(mI, 10) = mObj.sTablaDescr("CiaSeguros", " codciaseguro='" & mRec!CodCiaSeguro & "'", 1) & " - " & mRec!NroPoliza 'seguro y póliza CiaSeguros
         End If
         mI = mI + 1
         mRec.MoveNext
      Loop
   End If
   mRec.Close
   mValues = ""
   For mI = 1 To 8
      For mJ = 1 To 10
         mValues = mValues & "'" & mDatos(mI, mJ) & "',"
      Next
   Next
   mValues = Mid(mValues, 1, Len(mValues) - 1)
   mData.Execute "insert into Auxi2 values(" & mValues & ")"
   
   'Victimas Involucradas
   mData.Execute "create table Auxi3 (num_1 text, name_1 text, domic_1 text, doc_1 text, lugar_1 text, herido_1 text, fallecio_1 text, letra_1 text, sexo_1 text, edad_1 text, civil_1 text, cint_1 text," _
      & "num_2 text, name_2 text, domic_2 text, doc_2 text, lugar_2 text, herido_2 text, fallecio_2 text, letra_2 text, sexo_2 text, edad_2 text, civil_2 text, cint_2 text," _
      & "num_3 text, name_3 text, domic_3 text, doc_3 text, lugar_3 text, herido_3 text, fallecio_3 text, letra_3 text, sexo_3 text, edad_3 text, civil_3 text, cint_3 text," _
      & "num_4 text, name_4 text, domic_4 text, doc_4 text, lugar_4 text, herido_4 text, fallecio_4 text, letra_4 text, sexo_4 text, edad_4 text, civil_4 text, cint_4 text," _
      & "num_5 text, name_5 text, domic_5 text, doc_5 text, lugar_5 text, herido_5 text, fallecio_5 text, letra_5 text, sexo_5 text, edad_5 text, civil_5 text, cint_5 text," _
      & "num_6 text, name_6 text, domic_6 text, doc_6 text, lugar_6 text, herido_6 text, fallecio_6 text, letra_6 text, sexo_6 text, edad_6 text, civil_6 text, cint_6 text," _
      & "num_7 text, name_7 text, domic_7 text, doc_7 text, lugar_7 text, herido_7 text, fallecio_7 text, letra_7 text, sexo_7 text, edad_7 text, civil_7 text, cint_7 text," _
      & "num_8 text, name_8 text, domic_8 text, doc_8 text, lugar_8 text, herido_8 text, fallecio_8 text, letra_8 text, sexo_8 text, edad_8 text, civil_8 text, cint_8 text)"
      
   For mI = 1 To 8
      For mJ = 1 To 12
         mDatos(mI, mJ) = ""
      Next
   Next
   Set mRec = mObj.oTablaNroOrden("VictimasInvolucr", Left(Combo1.Text, 5), "order by nrovictima")
   If Not mRec.EOF Then
      mI = 1
      Do While Not mRec.EOF
         If mI <= 8 Then
            mDatos(mI, 1) = mRec!nrovictima
            mDatos(mI, 2) = NVL(mRec!nombre, "")
            mDatos(mI, 3) = NVL(mRec!domicilio, "")
            mDatos(mI, 4) = mObj.sTablaDescr("TipoDocu", " codtipodocu='" & NVL(mRec!codtipodoc, "") & "'", 1) & " - " & mRec!nrodocu
            mDatos(mI, 5) = mObj.sTablaDescr("LugarTrasl", " codlugartrasl='" & NVL(mRec!codlugartrasl, "") & "'", 1) & " - " & mObj.sTablaDescr("MedioTrasl", " codmediotrasl='" & NVL(mRec!codmediotrasl, "") & "'", 1)
            mDatos(mI, 6) = mRec!herido
            mDatos(mI, 7) = mRec!fallecio
            mDatos(mI, 8) = mRec!letra
            If NVL(mRec!Codsexo, "") = "1" Then
               mDatos(mI, 9) = "F"
            End If
            If NVL(mRec!Codsexo, "") = "2" Then
               mDatos(mI, 9) = "M"
            End If
            mDatos(mI, 10) = mRec!edad
            mDatos(mI, 11) = Left(mObj.sTablaDescr("EstadoCivil", " codestcivil='" & NVL(mRec!CodEstCivil, "") & "'", 1), 5)
            mDatos(mI, 12) = mRec!cinturon
         End If
         mI = mI + 1
         mRec.MoveNext
      Loop
   End If
   mRec.Close
   mValues = ""
   For mI = 1 To 8
      For mJ = 1 To 12
         mValues = mValues & "'" & mDatos(mI, mJ) & "',"
      Next
   Next
   mValues = Mid(mValues, 1, Len(mValues) - 1)
   mData.Execute "insert into Auxi3 values(" & mValues & ")"
   
   
   With CReport1
      .WindowTitle = "reporte"
      .DataFiles(0) = App.Path & "\RegAccidentes\FichaAccid.mdb"
      .ReportFileName = App.Path & "\RegAccidentes\" & "Rep23.rpt"
      .Action = 1
   End With
      With CReport1
      .WindowTitle = "reporte"
      .DataFiles(0) = App.Path & "\RegAccidentes\FichaAccid.mdb"
      .ReportFileName = App.Path & "\RegAccidentes\" & "Rep23i.rpt"
      .Action = 1
   End With
   
   Set mData = Nothing
   Set mRec = Nothing
   Set mObj = Nothing
End Sub

Private Sub sVerReporte(ByVal pLegales As Boolean)
Dim mObj As New clRAcc
Dim mObjNov As New clRNov
Dim mRec As New ADODB.Recordset
Dim mRec1 As New ADODB.Recordset
Dim mData As Database
Dim mVecDatos(125) As String
Dim mDatVehic(87) As String
Dim mWhere As String
Dim mI As Integer
Dim mJ As Integer
Dim mDescrMid As Integer
   mWhere = ""
   For mI = 1 To 100
      mVecDatos(mI) = ""
   Next
   Set mData = OpenDatabase(App.Path & "\RegAccidentes\FichaAccid.mdb")
   mData.Execute "create table Auxi (nroficha text, fecha text, codigo text, sent text, km text, av text, arr text, carril text, " _
         & "conf text, estrod text, visib text, calzSec text, paviment text, clima text, ilumina text, senial text, inconv text, " _
         & "cotro text, otros text, ccontra text, ccond1 text, ccond2 text, ccond3 text, cvehic text, " _
         & " obs_1 text, obs_2a text,obs_2b text, obs_2c text, obs_3 text, obs_4a text, obs_4b text, obs_5a text, obs_5b text, obs_5c text, obs_5d text, ramal text, foto text)"
   Set mRec = mObj.oTablaNroOrden("Ficha", Left(Combo1.Text, 5), " order by fecha")
   If Not mRec.EOF Then
      mVecDatos(1) = mRec!NroOrden
      mVecDatos(2) = mRec!Fecha
      mVecDatos(3) = mRec!CodAlfa
'      Select Case mRec!SentidoTrans
'         Case "01"
'            mVecDatos(4) = "nn"
'         Case "02"
'            mVecDatos(4) = "no"
'         Case "03"
'            mVecDatos(4) = "on"
'         Case "04", ""
'            mVecDatos(5) = "oo"
'      End Select
      mVecDatos(4) = mObjNov.sTablaDescr("sentidos", "codigo=" & mRec!SentidoTrans, 1) 'mp 20160314
      mVecDatos(4) = Right(mVecDatos(4), Len(mVecDatos(4)) - 3) 'mp 20160314
      mVecDatos(125) = mObjNov.sTablaDescr("ramales", "codigo=" & mRec!codramal, 1) 'mp 20160314
      mVecDatos(5) = mRec!Progresiva
      mVecDatos(6) = mRec!hora
      mVecDatos(7) = mRec!HoraLlegada
      mVecDatos(8) = sConvert(mRec!carril)
      mVecDatos(9) = mRec!CODCONFIG
      mVecDatos(10) = mRec!EstCalzada
      mVecDatos(11) = mRec!CODVISIBILIDAD
      mVecDatos(12) = mRec!lugaraccid        'CALZ SECUND
      mVecDatos(13) = mRec!CODPAVIM          'PAVIMENTO
      mVecDatos(14) = mRec!Clima1            'CLIMA
      mVecDatos(15) = mRec!Iluminac          'ILUMINACION
      mVecDatos(16) = sConvert(mRec!DemarcHoriz)
      mVecDatos(17) = mRec!CODINCONV
      mVecDatos(18) = mRec!AcciconOtro
      mVecDatos(19) = mRec!AccidOtro
      mVecDatos(20) = mRec!CodColisContra1
      mVecDatos(21) = mRec!CodCausaCond1
      mVecDatos(22) = mRec!CodCausaCond2
      mVecDatos(23) = mRec!CodCausaCond3
      mVecDatos(24) = mRec!causaVehic
      mVecDatos(73) = sConvert(mRec!Foto)
      mVecDatos(114) = mRec!OBS
   End If
   mRec.Close
   For mI = 1 To 24
      mWhere = mWhere & "'" & mVecDatos(mI) & "',"
   Next
   Set mRec = mObj.oTablaNroOrden("fichaobs", Left(Combo1.Text, 5), " order by 2")
   Do While Not mRec.EOF
      Select Case mRec!Indice
         Case 0
            mVecDatos(115) = Left(mRec!descripcion, 40)   'obs_2
            mVecDatos(116) = Mid(mRec!descripcion, 41, 40)
            mVecDatos(117) = Mid(mRec!descripcion, 81, 40)
         Case 1
            mVecDatos(118) = mRec!descripcion   'obs_3
         Case 2
            mVecDatos(119) = Left(mRec!descripcion, 20)  'obs_4
            mVecDatos(120) = Mid(mRec!descripcion, 21, 40)
         Case 17
            mVecDatos(121) = Left(mRec!descripcion, 40)
            mVecDatos(122) = Mid(mRec!descripcion, 41, 40)
            mVecDatos(123) = Mid(mRec!descripcion, 81, 40)
            mVecDatos(124) = Mid(mRec!descripcion, 121, 40)
     End Select
     mRec.MoveNext
   Loop
   mRec.Close
   'For mI = 90 To 100
   For mI = 114 To 125
      mWhere = mWhere & "'" & mVecDatos(mI) & "',"
   Next
   mWhere = mWhere & "'" & mVecDatos(73) & "'"
   mData.Execute "insert into Auxi values (" & mWhere & ")"
   mWhere = ""
   
   'Daños a GCO
   mData.Execute "create table Auxi2 (D1 text, D2 text, D3 text, D4 text, D5 text, D6 text, D7 text, D8 text, D9 text, D10 text, " _
      & " pat1a text, pat2a text, polad1a text, mov1a text, pat1b text, pat2b text, polad1b text, mov1b text, " _
      & " pat1c text, pat2c text, polad1c text, mov1c text, pat1d text, pat2d text, polad1d text, mov1d text, " _
      & " terc1a text, terc2a text, terc3a text, terc4a text, terc1b text, terc2b text, terc3b text, terc4b text, " _
      & " terc1c text, terc2c text, terc3c text, terc4c text, terc1d text, terc2d text, terc3d text, terc4d text, " _
      & " terc1e text, terc2e text, terc3e text, terc4e text, terc1f text, terc2f text, terc3f text, terc4f text, " _
      & " terc1g text, terc2g text, terc3g text, terc4g text, terc1h text, terc2h text, terc3h text, terc4h text, foto text, " _
      & " A1 text, A2 text, A3 text, A4 text, A5 text, A6 text, A7 text, A8 text, A9 text, A10 text)"
   
   Set mRec = mObj.oTablaNroOrden("daniosgco", Left(Combo1.Text, 5), "")
   If Not mRec.EOF Then
      For mI = 1 To 10
          mWhere = mWhere & "'" & mRec.Fields(mI) & "',"
      Next
   Else
      mWhere = "'0','0','0','0','0','0','0','0','0','0',"
   End If
   mRec.Close
   
   'Intervención GCO
   Set mRec = mObj.oTablaNroOrden("intergco", Left(Combo1.Text, 5), "")
   If Not mRec.EOF Then
      mI = 25
      Do While Not mRec.EOF
         mVecDatos(mI) = mRec.Fields(2)
         mVecDatos(mI + 1) = mRec.Fields(3)
         mVecDatos(mI + 2) = mRec.Fields(4)
         mVecDatos(mI + 3) = mRec.Fields(1)
         mI = mI + 4
         mRec.MoveNext
      Loop
   End If
   mRec.Close
   For mI = 25 To 40
      mWhere = mWhere & "'" & mVecDatos(mI) & "',"
   Next
   
   'Intervención Terceros
   Set mRec = mObj.oInterTerceros(Left(Combo1.Text, 5))
   If Not mRec.EOF Then
      mI = 41
      Do While Not mRec.EOF
         If mI <= 69 Then
            mVecDatos(mI) = mRec.Fields(0)
            mVecDatos(mI + 1) = mRec.Fields(1)
            mVecDatos(mI + 2) = mRec.Fields(2)
            mVecDatos(mI + 3) = mRec.Fields(3)
            mI = mI + 4
         End If
         mRec.MoveNext
      Loop
   End If
   mRec.Close
   For mI = 41 To 72
      mWhere = mWhere & "'" & mVecDatos(mI) & "',"
   Next
   mWhere = mWhere & "'" & mVecDatos(73) & "',"
   
   Set mRec = mObj.oTablaNroOrden("fichadescr", Left(Combo1.Text, 5), "")
   If Not mRec.EOF Then
      For mJ = 1 To 4
         mDescrMid = 1
         For mI = 1 To 10
            If Len(mRec.Fields(mJ)) >= (mI * 110) Then
               mVecDatos((10 * mJ) + 63 + mI) = Mid(mRec.Fields(mJ), mDescrMid, 110)
            Else
               If Mid(mRec.Fields(mJ), mDescrMid, 110) <> "" Then
                  mVecDatos((10 * mJ) + 63 + mI) = Mid(mRec.Fields(mJ), mDescrMid, 110)
               Else     'nuevo
                  mVecDatos((10 * mJ) + 63 + mI) = ""  'nuevo
               End If
            End If
            mDescrMid = mDescrMid + 110
         Next
      Next
   End If
   mRec.Close

   For mI = 74 To 83
      fReplace (mVecDatos(mI))
      mVecDatos(mI) = Replace(mVecDatos(mI), Chr(13) & Chr(10), " ")
      mWhere = mWhere & "'" & UCase(mVecDatos(mI)) & "',"
   Next
   
   
   mWhere = Mid(mWhere, 1, Len(mWhere) - 1)
   mData.Execute "insert into Auxi2 values (" & mWhere & ")"
   
   mWhere = ""
   mData.Execute "create table Auxi4 (B1 text, B2 text, B3 text, B4 text, B5 text, B6 text, B7 text, B8 text, B9 text, B10 text, " _
      & " C1 text, C2 text, C3 text, C4 text, C5 text, C6 text, C7 text, C8 text, C9 text, C10 text, " _
      & " E1 text, E2 text, E3 text, E4 text, E5 text, E6 text, E7 text, E8 text, E9 text, E10 text)"
   
   For mI = 84 To 113
      fReplace (mVecDatos(mI))
      mVecDatos(mI) = Replace(mVecDatos(mI), Chr(13) & Chr(10), " ")
      mWhere = mWhere & "'" & UCase(mVecDatos(mI)) & "',"
   Next
   mWhere = Mid(mWhere, 1, Len(mWhere) - 1)
   mData.Execute "insert into Auxi4 values (" & mWhere & ")"
   mWhere = ""

   mData.Execute "create table Auxi3 (codtipo text, dominio text, color text, marca text, modelo text, seguro text, poliza text, titular text, " _
      & "nomb text, domic text, tel text, doc text, " _
      & "nro1 text, nomb1 text, domic1 text, doc1 text, edad1 text, tel1 text, trasl1 text, est1 text, " _
      & "nro2 text, nomb2 text, domic2 text, doc2 text, edad2 text, tel2 text, trasl2 text, est2 text, " _
      & "nro3 text, nomb3 text, domic3 text, doc3 text, edad3 text, tel3 text, trasl3 text, est3 text, " _
      & "nro4 text, nomb4 text, domic4 text, doc4 text, edad4 text, tel4 text, trasl4 text, est4 text, " _
      & "nro5 text, nomb5 text, domic5 text, doc5 text, edad5 text, tel5 text, trasl5 text, est5 text, " _
      & "nro6 text, nomb6 text, domic6 text, doc6 text, edad6 text, tel6 text, trasl6 text, est6 text, " _
      & "nro7 text, nomb7 text, domic7 text, doc7 text, edad7 text, tel7 text, trasl7 text, est7 text, " _
      & "nro8 text, nomb8 text, domic8 text, doc8 text, edad8 text, tel8 text, trasl8 text, est8 text, " _
      & "nro9 text, nomb9 text, domic9 text, doc9 text, edad9 text, tel9 text, trasl9 text, est9 text, " _
      & "estgral text, estneum text, anexo text ) "
   Set mRec = mObj.oTablaNroOrden("VehiculosInvolucr", Left(Combo1.Text, 5), "order by 2")
   Do While Not mRec.EOF
      For mI = 1 To 87
         mDatVehic(mI) = ""
      Next
      mWhere = ""
      mDatVehic(1) = mRec!CodTipoVehic
      mDatVehic(2) = mRec!Dominio
      mDatVehic(3) = ""
      Set mRec1 = mObj.oTablaCodigo("colores", "codigo='" & mRec!codcolor & "'")
      If Not mRec1.EOF Then
         mDatVehic(3) = mRec1!descripcion
      End If
      mRec1.Close
      mDatVehic(4) = ""
      Set mRec1 = mObj.oMarcasVehicDescr(mRec!CodTipoVehic, mRec!CodMarca)
      If Not mRec1.EOF Then
         mDatVehic(4) = mRec1!descripcion
      End If
      mRec1.Close
      mDatVehic(5) = mRec!modelo
      mDatVehic(6) = ""
      Set mRec1 = mObj.oTablaCodigo("CiaSeguros", "codciaseguro = '" & mRec!CodCiaSeguro & "' ")
      If Not mRec1.EOF Then
         mDatVehic(6) = mRec1!descripcion
      End If
      mRec1.Close
      mDatVehic(7) = mRec!NroPoliza
      mDatVehic(8) = mRec!TITULAR
      mDatVehic(85) = mRec!ESTGRAL
      mDatVehic(86) = mRec!ESTNEUM
      mDatVehic(87) = mRec!ANEXOEST
      Set mRec1 = mObj.oTablaNroOrden("VictimasInvolucr", Left(Combo1.Text, 5), " and letra='" & mRec!letra & "' order by 2")
      mJ = 13
      Do While Not mRec1.EOF
         If mRec1!conductor = "1" Then
            mDatVehic(9) = mRec1!nombre
            mDatVehic(10) = mRec1!domicilio
            mDatVehic(11) = mRec1!TEL
            mDatVehic(12) = mRec1!nrodocu
         End If
            If mJ >= 85 Then
               For mI = 1 To 87
                  mWhere = mWhere & "'" & mDatVehic(mI) & "',"
               Next
               mWhere = Mid(mWhere, 1, Len(mWhere) - 1)
               mData.Execute "insert into Auxi3 values (" & mWhere & ")"
               mJ = 13
               For mI = 13 To 87
                  mDatVehic(mI) = ""
               Next
               mWhere = ""
            End If
            mDatVehic(mJ) = mRec1!nrovictima
            mDatVehic(mJ + 1) = mRec1!nombre
            mDatVehic(mJ + 2) = mRec1!domicilio
            mDatVehic(mJ + 3) = mRec1!nrodocu
            mDatVehic(mJ + 4) = mRec1!edad
            mDatVehic(mJ + 5) = mRec1!TEL
            mDatVehic(mJ + 6) = mObj.sTablaDescr("LugarTrasl", " codlugartrasl='" & mRec1!codlugartrasl & "'", 1)
            mDatVehic(mJ + 7) = mObj.sTablaDescr("estadoocupa", " codigo='" & mRec1!codestado & "'", 1)
            mJ = mJ + 8
         mRec1.MoveNext
      Loop
      mRec1.Close
            
      mRec.MoveNext
      For mI = 1 To 87
         mWhere = mWhere & "'" & mDatVehic(mI) & "',"
      Next
      mWhere = Mid(mWhere, 1, Len(mWhere) - 1)
      mData.Execute "insert into Auxi3 values (" & mWhere & ")"
   Loop
   mRec.Close
   'Ingresa peatones si hay
   
      Set mRec = mObj.oTablaNroOrden("VictimasInvolucr", Left(Combo1.Text, 5), " and letra='' order by 2")
   If Not mRec.EOF Then
      mJ = 13
      For mI = 1 To 87
         mDatVehic(mI) = ""
      Next
      Do While Not mRec.EOF
         If mJ >= 85 Then
            For mI = 1 To 87
               mWhere = mWhere & "'" & mDatVehic(mI) & "',"
            Next
            mWhere = Mid(mWhere, 1, Len(mWhere) - 1)
            mData.Execute "insert into Auxi3 values (" & mWhere & ")"
            mJ = 13
            For mI = 13 To 87
               mDatVehic(mI) = ""
            Next
            mWhere = ""
         End If
         mDatVehic(mJ) = mRec!nrovictima
         mDatVehic(mJ + 1) = mRec!nombre
         mDatVehic(mJ + 2) = mRec!domicilio
         mDatVehic(mJ + 3) = mRec!nrodocu
         mDatVehic(mJ + 4) = mRec!edad
         mDatVehic(mJ + 5) = mRec!TEL
         mDatVehic(mJ + 6) = mObj.sTablaDescr("LugarTrasl", " codlugartrasl='" & mRec!codlugartrasl & "'", 1)
         mDatVehic(mJ + 7) = mObj.sTablaDescr("estadoocupa", " codigo='" & mRec!codestado & "'", 1)
         mJ = mJ + 8
         mRec.MoveNext
      Loop
      mWhere = ""
      For mI = 1 To 87
         mWhere = mWhere & "'" & mDatVehic(mI) & "',"
      Next
      mWhere = Mid(mWhere, 1, Len(mWhere) - 1)
      mData.Execute "insert into Auxi3 values (" & mWhere & ")"
   End If
   mRec.Close
   'fin peatones
  
   Dim mAuxi
   Dim mAuxi2
   Dim mAuxi3
   Dim mAuxi4
   
   Set mAuxi = mData.OpenRecordset("SELECT count(*) as Total FROM Auxi")
   Set mAuxi2 = mData.OpenRecordset("SELECT count(*) as Total FROM Auxi2")
   Set mAuxi3 = mData.OpenRecordset("SELECT count(*) as Total FROM Auxi3")
   Set mAuxi4 = mData.OpenRecordset("SELECT count(*) as Total FROM Auxi4")
   mAuxi.Close
   mAuxi2.Close
   mAuxi3.Close
   mAuxi4.Close

   With CReport1
      .WindowTitle = "reporte"
      .DataFiles(0) = App.Path & "\RegAccidentes\FichaAccid.mdb"
      If pLegales Then
         .ReportFileName = App.Path & "\RegAccidentes\" & "Rep22l.rpt"
      Else
         .ReportFileName = App.Path & "\RegAccidentes\" & "Rep22.rpt"
      End If
      .Action = 1
   End With
    With CReport1
      .WindowTitle = "reporte"
      .DataFiles(0) = App.Path & "\RegAccidentes\FichaAccid.mdb"
      .ReportFileName = App.Path & "\RegAccidentes\" & "Rep22i.rpt"
      .Action = 1
   End With
   Set mData = Nothing
   Set mRec = Nothing
   Set mRec1 = Nothing
   Set mObj = Nothing
   Set mObjNov = Nothing
End Sub

Private Function sConvert(ByVal pDato As String) As String
Dim mI As Integer
   sConvert = ""
   For mI = 1 To Len(pDato)
      If Mid(pDato, mI, 1) = "0" Then
         sConvert = sConvert & "o"
      Else
         sConvert = sConvert & "n"
      End If
   Next
End Function

Public Sub sVerFicha(ByVal pNroFicha As String)
Dim mObj As New clRAcc
   Combo1.Clear
   Combo1.AddItem pNroFicha
   Combo1.ListIndex = 0
   sMsgEspere Me, "Generando informe...", True
   mObjAcc.mBorrarAuxi "\RegAccidentes\FichaAccid", "Auxi"
   mObjAcc.mBorrarAuxi "\RegAccidentes\FichaAccid", "Auxi2"
   mObjAcc.mBorrarAuxi "\RegAccidentes\FichaAccid", "Auxi3"
   mObjAcc.mBorrarAuxi "\RegAccidentes\FichaAccid", "Auxi4"
   mObjAcc.mBorrarAuxi "\RegAccidentes\FichaAccid", "Auxi5"
   If DateDiff("d", Right(pNroFicha, 10), "01/03/2008") > 0 Then
      sVerFichaOlder  'para fichas viejas
   Else
      sVerReporte (Check1)
   End If
   sMsgEspere Me, "", False
   Set mObj = Nothing
   RAcc10_frm.Visible = False

End Sub

