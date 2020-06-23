VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form RAcc6_frm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form6"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   ControlBox      =   0   'False
   LinkTopic       =   "Form9"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   6825
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   600
         Top             =   480
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   5040
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   5160
         MaxLength       =   5
         TabIndex        =   9
         Top             =   2040
         Width           =   855
      End
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   1920
         TabIndex        =   8
         Top             =   2760
         Width           =   2775
         Begin VB.CommandButton Command1 
            Caption         =   "&Generar"
            Default         =   -1  'True
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   0
            Left            =   240
            TabIndex        =   3
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Cancelar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   1
            Left            =   1440
            TabIndex        =   4
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   3720
         MaxLength       =   10
         TabIndex        =   2
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   1
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hora"
         Height          =   195
         Left            =   5160
         TabIndex        =   10
         Top             =   1800
         Width           =   345
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
         Height          =   195
         Left            =   3720
         TabIndex        =   7
         Top             =   1800
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial"
         Height          =   195
         Left            =   1800
         TabIndex        =   6
         Top             =   1800
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         TabIndex        =   5
         Top             =   720
         Width           =   75
      End
   End
End
Attribute VB_Name = "RAcc6_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mData As Database
Dim mObjRAcc As New clRAcc
Dim mObjAcc As New clAccess
Dim mRec As New ADODB.Recordset
Public mDesde As String
Public mHasta As String
Public mImprimir As Boolean
Public mReporte As String
Dim Auxi

Private Sub Form_Load()
   sInitForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
   mData.Close
   Set mData = Nothing
   Set mObjRAcc = Nothing
   Set mObjAcc = Nothing
   Set mRec = Nothing
   ShowMenu 2, True, False
End Sub

Private Sub Command1_Click(Index As Integer)
Dim mRec1 As New ADODB.Recordset
Dim mBorrar As Boolean
Dim mFlag As Boolean
Dim mi As Integer
Dim mj As Integer
Dim mAuxi As Variant

Dim DatosCol(7, 4) As Integer
Dim DatosTraz(7, 4) As Integer
Dim DatosTronc(2, 4) As Integer
Dim DatosVehic(4) As Integer
Dim DatosCausas(5) As Integer
Dim DatosClima(8) As Integer
Dim DatosDead(10) As Integer
Dim DatosOtrosAmb(13) As Integer
Dim DatosCond(17) As Integer
Dim mProgr(12) As Double
Dim MesHora(24, 3) As Integer
Dim Dias(3, 7) As Integer
Dim Horas(25) As String
Dim xBorrar As Boolean

Dim mDatosStr(10) As String
Dim Value1, Value2 As String
Dim mFecha As String
Dim mNroOrden As String
Dim mNroOrdenOld As String
Dim mHora As String
Dim mKm As String
Dim mLugar As String
Dim mValorX As String
Dim mLetra As String

   mFlag = False
   If Index = 0 Then
      If mReporte <> "Evaluación" Then
         If sValidFechaDesdeHasta(Text1(0).Text, Text1(1).Text) Then
            mObjAcc.mBorrarAuxi "\RegAccidentes\FichaAccid", "Auxi"
            mDesde = Text1(0).Text
            mHasta = Text1(1).Text
            mFlag = True
         End If
      Else
         mFlag = True
      End If
      If mFlag Then
         Value1 = ""
         Value2 = ""
         sMsgEspere Me, "Generando informe... aguarde un momento.", True
         mImprimir = False
         Select Case mReporte
            Case "Principal"  'Viejo y Nuevo
               Set mRec = mObjRAcc.oFichaAccidentados(Text1(0).Text, Text1(1).Text)
               If Not mRec.EOF Then
                  mData.Execute ("CREATE TABLE Auxi (NroOrden TEXT,Fecha TEXT,Hora TEXT,Progresiva TEXT,TotalFallec INTEGER,TotalHerido INTEGER, TotalVehic INTEGER,mLugarAccid TEXT, mHora TEXT)")
                  mNroOrden = mRec!NroOrden
                  mNroOrdenOld = mRec!NroOrden
                  For mi = 1 To 4
                     DatosVehic(mi) = 0
                  Next
                  mLetra = ""
                  Do While Not mRec.EOF
                     mNroOrden = mRec!NroOrden
                     If mNroOrdenOld = mNroOrden Then
                        mFecha = mRec!Fecha
                        mHora = mRec!hora
                        mKm = mRec!Progresiva
                        mLugar = mObjRAcc.sTablaDescr("LugarAccid", "codlugaraccid='" & mRec!lugaraccid & "'", 1)
                        If NVL(mRec!letra, "") <> "" Then
                           If mLetra <> mRec!letra Then
                              DatosVehic(1) = DatosVehic(1) + 1
                              mLetra = mRec!letra
                           End If
                        End If
                        If mRec!herido <> "" Or mRec!codestado = "02" Or mRec!codestado = "03" Then
                           DatosVehic(2) = DatosVehic(2) + 1
                        End If
                        If mRec!fallecio <> "" Or mRec!codestado = "04" Or mRec!codestado = "05" Then
                           DatosVehic(3) = DatosVehic(3) + 1
                        End If
                        mRec.MoveNext
                     Else
                        mData.Execute ("INSERT INTO Auxi (NroOrden,Fecha,Hora,Progresiva,TotalFallec,TotalHerido,TotalVehic,mLugarAccid, mHora) VALUES " _
                         & " ('" & mNroOrdenOld & "','" & mFecha & "','" & mHora & "'," & mKm & "," & DatosVehic(3) & "," & DatosVehic(2) & "," & DatosVehic(1) & ",'" & mLugar & "','" & mHora & "')")
                        For mi = 1 To 4
                           DatosVehic(mi) = 0
                        Next
                        mNroOrdenOld = mRec!NroOrden
                        mLetra = ""
                     End If
                  Loop
                  mData.Execute ("INSERT INTO Auxi (NroOrden,Fecha,Hora,Progresiva,TotalFallec,TotalHerido,TotalVehic,mLugarAccid, mHora) VALUES " _
                         & " ('" & mNroOrdenOld & "','" & mFecha & "','" & mHora & "'," & mKm & "," & DatosVehic(3) & "," & DatosVehic(2) & "," & DatosVehic(1) & ",'" & mLugar & "','" & mHora & "')")
                  mRec.Close
                  Set mAuxi = mData.OpenRecordset("SELECT * FROM Auxi")
                  If PrintRep(mAuxi) Then
                     CrystalReport1.ReportFileName = App.Path & "\RegAccidentes\rep06.rpt"
                     CrystalReport1.Action = 1
                  End If
               End If
         
            Case "Consulta_2004" 'MYSQL
               mData.Execute ("CREATE TABLE Auxi (Fecha DATETIME,Hora TEXT,Progresiva DOUBLE,mAccidconOtro TEXT,mCausaCond TEXT,mCausaVehic TEXT,mClima TEXT,mSentTrans TEXT,mLugarAccid TEXT,mHora TEXT)")
               Set mRec = mObjRAcc.oConsulta2004(Text1(0).Text, Text1(1).Text)
               Do While Not mRec.EOF
                  mData.Execute ("INSERT INTO Auxi (Fecha,Hora,Progresiva,mAccidconOtro,mCausaCond,mCausaVehic,mClima,mSentTrans,mLugarAccid,mHora) VALUES ('" & Format(mRec.Fields(0), "dd/mm/yyyy") & "','" & Format(mRec.Fields(1), "HH:MM") & "'," & NVL(mRec.Fields(2), 0) & ",'" & NVL(mRec.Fields(3), "") & "','" & NVL(mRec.Fields(4), "") & "','" & NVL(mRec.Fields(5), "") & "','" & NVL(mRec.Fields(6), "") & "','" & mRec.Fields(7) & "','" & mRec.Fields(8) & "','" & mRec.Fields(9) & "')")
                  mRec.MoveNext
               Loop
               mRec.Close
               mData.Execute "update Auxi set mHora = hora"
               Set mAuxi = mData.OpenRecordset("SELECT * FROM Auxi")
               If PrintRep(mAuxi) Then
                  CrystalReport1.ReportFileName = App.Path & "\Regaccidentes\rep07.rpt"
                  CrystalReport1.Action = 1
               End If
         
            Case "Peligrosidad por Sector" 'Viejo y Nuevo
               mData.Execute ("CREATE TABLE Auxi (NroOrden TEXT,mSentTrans TEXT,Progresiva DOUBLE,Fecha DATETIME,Hora TEXT,mLugarAccid TEXT,mAccidOtros TEXT,mAccidconOtro TEXT,mColisContra TEXT,mCausaVehic TEXT,mCausaCond TEXT,mClima TEXT,mHora TEXT)")
               Set mRec = mObjRAcc.oPeligroSector(Text1(0).Text, Text1(1).Text)
               Do While Not mRec.EOF
                  mData.Execute ("INSERT INTO Auxi (NroOrden,mSentTrans,Progresiva,Fecha,Hora,mLugarAccid,mAccidOtros,mAccidconOtro,mColisContra,mCausaVehic,mCausaCond,mClima,mHora) VALUES ('" & mRec.Fields(0) & "','" & mRec.Fields(1) & "'," & mRec.Fields(2) & ",'" & Format(mRec.Fields(3), "dd/mm/yyyy") & "','" & Format(mRec.Fields(4), "HH:MM") & "','" & NVL(mRec.Fields(5), "") & "','" & NVL(mRec.Fields(6), "") & "','" & NVL(mRec.Fields(7), "") & "','" & NVL(mRec.Fields(8), "") & "','" & NVL(mRec.Fields(9), "") & "','" & mRec.Fields(10) & "','" & mRec.Fields(11) & "','" & mRec.Fields(12) & "')")
                  mRec.MoveNext
               Loop
               mRec.Close
               mData.Execute "update Auxi set mHora = hora"
               Set mAuxi = mData.OpenRecordset("SELECT * FROM Auxi")
               If PrintRep(mAuxi) Then
                  CrystalReport1.ReportFileName = App.Path & "\RegAccidentes\rep08.rpt"
                  CrystalReport1.Action = 1
               End If
     
            Case "Evaluación" 'MySQL
               mObjAcc.mBorrarAuxi "\RegAccidentes\FichaAccid", "Auxi"
               mData.Execute ("CREATE TABLE Auxi (Nombre TEXT,TotalOrdens INTEGER,Fecha DATETIME)")
               Set mRec = mObjRAcc.oEvolucion()
               Do While Not mRec.EOF
                  mData.Execute ("INSERT INTO Auxi (Nombre,TotalOrdens,Fecha) VALUES ('" & mRec.Fields(0) & "'," & mRec.Fields(1) & ",'" & Format(mRec.Fields(2), "dd/mm/yyyy") & "')")
                  mRec.MoveNext
               Loop
               mRec.Close
               Set mAuxi = mData.OpenRecordset("SELECT * FROM Auxi")
               Label4.Visible = False
               Combo1.Visible = False
               If PrintRep(mAuxi) Then
                  CrystalReport1.ReportFileName = App.Path & "\RegAccidentes\rep09.rpt"
                  CrystalReport1.Action = 1
               End If
        
            Case "Día de Semana" 'Viejo y Nuevo
               mData.Execute ("CREATE TABLE Auxi (Flag INTEGER,xDias TEXT,xHerid INTEGER,xFallec INTEGER,xAccid INTEGER)")
               Set mRec = mObjRAcc.oTabla("Dias", "")
               Do While Not mRec.EOF
                  mData.Execute ("INSERT INTO Auxi (Flag,xDias,xHerid,xFallec,xAccid) VALUES (1,'" & mRec!Dia & "',0,0,0)")
                  mRec.MoveNext
               Loop
               mRec.Close
               Set mRec = mObjRAcc.oTotDiaSemana(Text1(0).Text, Text1(1).Text, " and VictimasInvolucr.NroOrden = Ficha.NroOrden AND (VictimasInvolucr.Herido in ('1','2') or VictimasInvolucr.codestado in ('02','03'))")
               Do While Not mRec.EOF
                  mData.Execute ("UPDATE Auxi SET xHerid = " & mRec!fall & " WHERE xDias = '" & mObjRAcc.sTablaDescr("Dias", " coddia='" & mRec!Dias & "'", 1) & "'")
                  mRec.MoveNext
               Loop
               mRec.Close
               Set mRec = mObjRAcc.oTotDiaSemana(Text1(0).Text, Text1(1).Text, " and VictimasInvolucr.NroOrden = Ficha.NroOrden AND (VictimasInvolucr.Fallecio in ('1','2') or VictimasInvolucr.codestado in ('04','05'))")
               Do While Not mRec.EOF
                  mData.Execute ("UPDATE Auxi SET xFallec = " & mRec!fall & " WHERE xDias = '" & mObjRAcc.sTablaDescr("Dias", " coddia='" & mRec!Dias & "'", 1) & "'")
                  mRec.MoveNext
               Loop
               mRec.Close
               Set mRec = mObjRAcc.oTotFichaDiaSemana(Text1(0).Text, Text1(1).Text, "")
               Do While Not mRec.EOF
                  mData.Execute ("UPDATE Auxi SET xAccid = " & mRec!fall & " WHERE xDias = '" & mObjRAcc.sTablaDescr("Dias", " coddia='" & mRec!Dias & "'", 1) & "'")
                  mRec.MoveNext
               Loop
               mRec.Close
               Set mAuxi = mData.OpenRecordset("SELECT * FROM Auxi")
               If PrintRep(mAuxi) Then
                  CrystalReport1.ReportFileName = App.Path & "\RegAccidentes\rep10.rpt"
                  CrystalReport1.Action = 1
               End If
       
            Case "Hora Día" 'Viejo y Nuevo
               For mi = 0 To 23
                  Horas(mi) = Format(mi, "00")
               Next
               Horas(24) = "24"
               mData.Execute ("CREATE TABLE Auxi (Flag INTEGER,xHoras TEXT,xAccid INTEGER,xHerid INTEGER,xFallec INTEGER)")
               For mi = 0 To 23
                  mj = 0
                  mValorX = " " & mObjRAcc.iTotAccMesHora(mDesde, mHasta, Horas(mi) & ":00", Horas(mi) & ":59", "")
                  mValorX = mValorX & "," & mObjRAcc.iTotAccFicVicMesHora(mDesde, mHasta, Horas(mi) & ":00", Horas(mi) & ":59", "AND (VictimasInvolucr.Herido in ('1','2') or VictimasInvolucr.codestado in ('02','03'))")
                  mValorX = mValorX & "," & mObjRAcc.iTotAccFicVicMesHora(mDesde, mHasta, Horas(mi) & ":00", Horas(mi) & ":59", "AND (VictimasInvolucr.Fallecio in ('1','2') or VictimasInvolucr.codestado in ('04','05'))")
                  mData.Execute ("INSERT INTO Auxi (Flag,xHoras,xAccid,xHerid,xFallec) VALUES (1,'" & Horas(mi) & " a " & Horas(mi + 1) & "'," & mValorX & ")")
               Next
               Set mAuxi = mData.OpenRecordset("SELECT * FROM Auxi")
               If PrintRep(mAuxi) Then
                  CrystalReport1.ReportFileName = App.Path & "\RegAccidentes\rep11.rpt"
                  CrystalReport1.Action = 1
               End If
              
            Case "Otros Traslados" 'MySQL
               mData.Execute ("CREATE TABLE Auxi (NroOrden TEXT,Fecha DATETIME,Progresiva DOUBLE,Hora TEXT,Descripcion TEXT,mHora TEXT)")
               Set mRec = mObjRAcc.oOtrosTrasl(Text1(0).Text, Text1(1).Text)
               Do While Not mRec.EOF
                  mData.Execute ("INSERT INTO Auxi (NroOrden,Fecha,Progresiva,Hora,Descripcion,mHora) VALUES ('" & mRec.Fields(0) & "','" & Format(mRec.Fields(1), "dd/mm/yyyy") & "'," & mRec.Fields(2) & ",'" & mRec.Fields(3) & "','" & mRec.Fields(4) & "','" & mRec.Fields(5) & "')")
                  mRec.MoveNext
               Loop
               mRec.Close
               mData.Execute "update Auxi set mHora = hora"
               Set mAuxi = mData.OpenRecordset("SELECT * FROM Auxi")
               If PrintRep(mAuxi) Then
                  CrystalReport1.ReportFileName = App.Path & "\RegAccidentes\rep12.rpt"
                  CrystalReport1.Action = 1
               End If
             
            Case "Principal Colectora" 'Módulo viejo y nuevo
               mData.Execute ("CREATE TABLE Auxi (NroOrden TEXT,Fecha DATE,Hora TEXT,Progresiva DOUBLE,TotalVehic INTEGER,TotalFallec INTEGER, TotalHerido INTEGER,mLugarAccid TEXT,mSentTrans TEXT,mHora TEXT)")
               Set mRec = mObjRAcc.oTablaCodigo("LugarAccid", "codlugaraccid='09'")
               If Not mRec.EOF Then
                  mLugar = mRec!DETALLE
               End If
               mRec.Close
               Set mRec = mObjRAcc.oTabla("SentidoTrans", " order by 1")
               Do While Not mRec.EOF
                  mDatosStr(Val(mRec.Fields(0))) = mRec.Fields(1)
                  mRec.MoveNext
               Loop
               mRec.Close
               Set mRec = mObjRAcc.oFichasFechaProgr(Text1(0).Text, Text1(1).Text, "", "", " lugaraccid='09'")
               If Not mRec.EOF Then
                  Do While Not mRec.EOF
                     For mi = 1 To 3
                        DatosDead(mi) = 0
                     Next
                     mValorX = mObjRAcc.sTablaDescr("SentidoTrans", "codsentidotrans='" & mRec!SentidoTrans & "'", 0)
                     DatosDead(1) = mObjRAcc.iTotalHeridosNroOrden(mRec!NroOrden)
                     DatosDead(2) = mObjRAcc.iTotalMuertosNroOrden(mRec!NroOrden)
                     DatosDead(3) = mObjRAcc.iTotalVehiculosNroOrden(mRec!NroOrden)
                     mData.Execute ("INSERT INTO Auxi VALUES ('" & mRec!NroOrden & "','" & Format(mRec!Fecha, "dd/mm/yyyy") & "','" & mRec!hora & "'," & mRec!Progresiva & "," & DatosDead(3) & "," & DatosDead(2) & "," & DatosDead(1) & ",'" & mLugar & "','" & mDatosStr(Val(mRec!SentidoTrans)) & "','" & mRec!hora & "')")
                     mRec.MoveNext
                  Loop
               End If
               mRec.Close
               
               Set mAuxi = mData.OpenRecordset("SELECT * FROM Auxi")
               If PrintRep(mAuxi) Then
                  CrystalReport1.ReportFileName = App.Path & "\RegAccidentes\rep13.rpt"
                  CrystalReport1.Action = 1
               End If
           
            Case "Puntos Negros" 'MySQL
               mData.Execute ("CREATE TABLE Auxi (NroOrden TEXT,mSentTrans TEXT,Progresiva DOUBLE,Fecha DATETIME,Hora TEXT,mLugarAccid TEXT,mAccidOtros TEXT,mAccidconOtro TEXT,mColisContra TEXT,mCausaVehic TEXT,mCausaCond TEXT,mClima TEXT,mHora TEXT)")
               Set mRec = mObjRAcc.oPuntosNegros(mDesde, mHasta)
               Do While Not mRec.EOF
                  mData.Execute ("INSERT INTO Auxi (NroOrden,mSentTrans,Progresiva,Fecha,Hora,mLugarAccid,mAccidOtros,mAccidconOtro,mColisContra,mCausaVehic,mCausaCond,mClima,mHora) VALUES ('" & mRec.Fields(0) & "','" & mRec.Fields(1) & "'," & mRec.Fields(2) & ",'" & Format(mRec.Fields(3), "dd/mm/yyyy") & "','" & Format(mRec.Fields(4), "HH:MM") & "','" & NVL(mRec.Fields(5), "") & "','" & NVL(mRec.Fields(6), "") & "','" & NVL(mRec.Fields(7), "") & "','" & NVL(mRec.Fields(8), "") & "','" & NVL(mRec.Fields(9), "") & "','" & mRec.Fields(10) & "','" & mRec.Fields(11) & "','" & mRec.Fields(12) & "')")
                  mRec.MoveNext
               Loop
               mRec.Close
               mData.Execute "update Auxi set mHora = hora"
               Set mAuxi = mData.OpenRecordset("SELECT * FROM Auxi")
               If PrintRep(mAuxi) Then
                  CrystalReport1.ReportFileName = App.Path & "\RegAccidentes\rep14.rpt"
                  CrystalReport1.Action = 1
               End If
      
            Case "Informe Personal" 'Viejo y NuevoW
               mData.Execute ("CREATE TABLE Auxi (MovilNro TEXT,Fecha DATETIME,Hora TEXT,Progresiva DOUBLE,TotalFallec INTEGER,TotalHerido INTEGER,mAccidOtros TEXT,mAccidconOtro TEXT,mColisContra TEXT,mClima TEXT,mHora TEXT)")
               Set mRec = mObjRAcc.oInfPersonal(Text1(0).Text, Text1(1).Text)
               Do While Not mRec.EOF
                  For mi = 1 To 4
                     mDatosStr(mi) = ""
                  Next
                  mDatosStr(1) = mObjRAcc.sTablaDescr("Otros", "codotros='" & mRec!AccidOtro & "'", 1)
                  mDatosStr(2) = mObjRAcc.sTablaDescr("ConOtroVehic", "codconotro='" & mRec!AcciconOtro & "'", 1)
                  mDatosStr(3) = mObjRAcc.sTablaDescr("ColisionContra", "codcolision='" & mRec!CodColisContra1 & "'", 1)
                  mDatosStr(4) = mObjRAcc.sTablaDescr("Clima", "codclima='" & mRec!Clima1 & "'", 1)
                  DatosDead(1) = mObjRAcc.iTotalHeridosNroOrden(mRec!NroOrden)
                  DatosDead(2) = mObjRAcc.iTotalMuertosNroOrden(mRec!NroOrden)
                  mData.Execute ("INSERT INTO Auxi (MovilNro,Fecha,Hora,Progresiva,TotalFallec,TotalHerido,mAccidOtros,mAccidconOtro,mColisContra,mClima,mHora) VALUES " _
                  & "('" & mRec.Fields(1) & "','" & Format(mRec.Fields(2), "dd/mm/yyyy") & "','" & mRec.Fields(3) & "'," & mRec.Fields(4) & "," & DatosDead(2) & "," & DatosDead(1) & ",'" & mDatosStr(1) & "','" & mDatosStr(2) & "','" & mDatosStr(3) & "','" & mDatosStr(4) & "','" & mRec.Fields(3) & "')")
                  mRec.MoveNext
               Loop
               mRec.Close
               mData.Execute "update Auxi set mHora = hora"
               Set mAuxi = mData.OpenRecordset("SELECT * FROM Auxi")
               If PrintRep(mAuxi) Then
                  CrystalReport1.ReportFileName = App.Path & "\RegAccidentes\rep15.rpt"
                  CrystalReport1.Action = 1
               End If
        
            Case "Peatón Ciclista" 'MySQL
               mData.Execute ("CREATE TABLE Auxi (NroOrden TEXT,Fecha DATETIME,Hora TEXT,Progresiva DOUBLE,mSentTrans TEXT,mLugarAccid TEXT,mAccidOtros TEXT,TotalFallec INTEGER,TotalHerido INTEGER,mHora TEXT)")
               Set mRec = mObjRAcc.oPeatonCiclis(mDesde, mHasta)
               Do While Not mRec.EOF
                  mData.Execute ("INSERT INTO Auxi (NroOrden,Fecha,Hora,Progresiva,mSentTrans,mLugarAccid,mAccidOtros,TotalFallec,TotalHerido,mHora) VALUES ('" & mRec.Fields(0) & "','" & Format(mRec.Fields(1), "dd/mm/yyyy") & "','" & Format(mRec.Fields(2), "HH:MM") & "'," & NVL(mRec.Fields(3), 0) & ",'" & NVL(mRec.Fields(4), "") & "','" & NVL(mRec.Fields(5), "") & "','" & NVL(mRec.Fields(6), "") & "'," & NVL(mRec.Fields(7), 0) & "," & NVL(mRec.Fields(8), 0) & ",'" & mRec.Fields(9) & "')")
                  mRec.MoveNext
               Loop
               mRec.Close
               mData.Execute "update Auxi set mHora = hora"
               Set mAuxi = mData.OpenRecordset("SELECT * FROM Auxi")
               If PrintRep(mAuxi) Then
                  CrystalReport1.ReportFileName = App.Path & "\RegAccidentes\rep16.rpt"
                  CrystalReport1.Action = 1
               End If
        
            Case "Tipo Vehículo" 'MySQL
               mData.Execute ("CREATE TABLE Auxi (NroOrden TEXT,Fecha DATETIME,mTipoVehic TEXT)")
               Set mRec = mObjRAcc.oTipoVehic(mDesde, mHasta, Left(Combo1.Text, 2))
               Do While Not mRec.EOF
                  mData.Execute ("INSERT INTO Auxi (NroOrden,Fecha,mTipoVehic) VALUES ('" & mRec.Fields(0) & "','" & Format(mRec.Fields(1), "dd/mm/yyyy") & "','" & NVL(mRec.Fields(2), "") & "')")
                  mRec.MoveNext
               Loop
               mRec.Close
               Set mAuxi = mData.OpenRecordset("SELECT * FROM Auxi")
               If PrintRep(mAuxi) Then
                  CrystalReport1.ReportFileName = App.Path & "\RegAccidentes\rep17.rpt"
                  CrystalReport1.Action = 1
               End If
             
            Case "Total Discriminado" 'Viejo y Nuevo
               For mi = 0 To 23
                  MesHora(mi, 0) = mObjRAcc.iAccidFechaHora(Text1(0).Text, Text1(1).Text, Format(mi, "00") & ":00", Format(mi, "00") & ":59")
                  MesHora(mi, 1) = mObjRAcc.iTotalMuertosCodigo(Text1(0).Text, Text1(1).Text, " hora between '" & Format(mi, "00") & ":00" & "' and '" & Format(mi, "00") & ":59" & "' and (herido in ('1','2') or codestado in ('02','03'))")
                  MesHora(mi, 2) = mObjRAcc.iTotalMuertosCodigo(Text1(0).Text, Text1(1).Text, " hora between '" & Format(mi, "00") & ":00" & "' and '" & Format(mi, "00") & ":59" & "' and (fallecio in ('1','2') or codestado in ('04','05'))")
               Next
               'DIAS DE LAS SEMANA
               For mi = 1 To 7
                  Dias(0, (mi - 1)) = mObjRAcc.iTotAccDia(mDesde, mHasta, (mi - 1), "")
                  Dias(1, (mi - 1)) = mObjRAcc.iTotAccDiaVict(mDesde, mHasta, (mi - 1), " AND Ficha.NroOrden = VictimasInvolucr.NroOrden AND (VictimasInvolucr.Herido in ('1','2') or codestado in ('02','03'))")
                  Dias(2, (mi - 1)) = mObjRAcc.iTotAccDiaVict(mDesde, mHasta, (mi - 1), " AND Ficha.NroOrden = VictimasInvolucr.NroOrden AND (VictimasInvolucr.Fallecio in ('1','2') or codestado in ('04','05'))")
               Next
               mData.Execute ("CREATE TABLE Auxi (CA1 Integer,CA2 Integer,CA3 Integer,CA4 Integer,CA5 Integer,CA6 Integer,CA7 Integer,CA8 Integer,CA9 Integer,CA10 Integer,CA11 Integer,CA12 Integer,CA13 Integer,CA14 Integer,CA15 Integer,CA16 Integer,CA17 Integer,CA18 Integer,CA19 Integer,CA20 Integer,CA21 Integer,CA22 Integer,CA23 Integer,CA24 Integer,CH1 Integer,CH2 Integer,CH3 Integer,CH4 Integer,CH5 Integer,CH6 Integer,CH7 Integer,CH8 Integer,CH9 Integer,CH10 Integer,CH11 Integer," _
                           & "CH12 Integer,CH13 Integer,CH14 Integer,CH15 Integer,CH16 Integer,CH17 Integer,CH18 Integer,CH19 Integer,CH20 Integer,CH21 Integer,CH22 Integer,CH23 Integer,CH24 Integer,CF1 Integer,CF2 Integer,CF3 Integer,CF4 Integer,CF5 Integer,CF6 Integer,CF7 Integer,CF8 Integer,CF9 Integer,CF10 Integer,CF11 Integer,CF12 Integer,CF13 Integer,CF14 Integer,CF15 Integer,CF16 Integer,CF17 Integer,CF18 Integer,CF19 Integer,CF20 Integer,CF21 Integer,CF22 Integer,CF23 Integer,CF24 Integer," _
                           & "DA1 Integer, DA2 Integer,DA3 Integer, DA4 Integer,DA5 Integer, DA6 Integer,DA7 Integer,DH1 Integer, DH2 Integer,DH3 Integer, DH4 Integer,DH5 Integer, DH6 Integer,DH7 Integer,DF1 Integer, DF2 Integer,DF3 Integer, DF4 Integer,DF5 Integer, DF6 Integer,DF7 Integer)")
                        
               For mj = 0 To 2
                  For mi = 0 To 23
                     Value1 = Value1 & MesHora(mi, mj) & ","
                  Next
               Next
               For mj = 0 To 2
                  For mi = 0 To 6
                     Value2 = Value2 & Dias(mj, mi) & ","
                  Next
               Next
               Value2 = Left(Value2, Len(Value2) - 1)
               mData.Execute ("INSERT INTO Auxi (CA1,CA2,CA3,CA4,CA5,CA6,CA7,CA8,CA9,CA10,CA11,CA12,CA13,CA14,CA15,CA16,CA17,CA18,CA19,CA20,CA21,CA22,CA23,CA24,CH1,CH2,CH3,CH4,CH5,CH6,CH7,CH8,CH9,CH10,CH11,CH12,CH13,CH14,CH15,CH16,CH17,CH18,CH19,CH20,CH21,CH22,CH23,CH24,CF1,CF2,CF3,CF4,CF5,CF6,CF7,CF8,CF9,CF10,CF11,CF12,CF13,CF14,CF15,CF16,CF17,CF18,CF19,CF20,CF21,CF22,CF23,CF24,DA1,DA2,DA3,DA4,DA5,DA6,DA7,DH1,DH2,DH3,DH4,DH5,DH6,DH7,DF1,DF2,DF3,DF4,DF5,DF6,DF7) VALUES (" & Value1 & Value2 & ")")
               Set mAuxi = mData.OpenRecordset("select * from auxi")
               mAuxi.Close
               For mi = 0 To mData.TableDefs.Count - 1
                  If mData.TableDefs(mi).Name = "Auxi" Then
                     CrystalReport1.WindowTitle = "Reporte " & mReporte
                     CrystalReport1.DataFiles(0) = App.Path & "\RegAccidentes\FichaAccid.mdb"
                     CrystalReport1.Formulas(0) = "Listado = '" & Text1(0).Text & " al " & Text1(1).Text & ".'"
                     CrystalReport1.ReportFileName = App.Path & "\RegAccidentes\rep20.rpt"
                     CrystalReport1.WindowState = crptMaximized
                     CrystalReport1.Action = 1
                     mi = 300
                     mImprimir = True
                  End If
               Next
             
            Case "Total Mensuales" 'MySQL   ----Modificación de intervalos de progresivas - 04/10/2004
               mProgr(0) = 12.9   '12.95
               mProgr(1) = 21.62  '20.8
               mProgr(2) = 27.18  '27.31
               mProgr(3) = 35.84  '36.7
               mProgr(4) = 38.57  '41.5
               mProgr(5) = 47.66  '47.66
               mProgr(6) = 63.3   '51.81
               mProgr(7) = 65.14  '65.14
               mProgr(8) = 25.75  'TRONCAL ITUZ
               mProgr(9) = 26.25  'TRONCAL ITUZ
               mProgr(10) = 57.5  'TRONCAL LUJAN
               mProgr(11) = 58    'TRONCAL LUJAN
               'Traza y Colectora
               For mi = 0 To 6
                  DatosTraz(mi, 0) = mObjRAcc.iAccidentesProgr(Text1(0).Text, Text1(1).Text, mProgr(mi), mProgr(mi + 1), True)
                  DatosTraz(mi, 1) = mObjRAcc.iAccidVictProgr(Text1(0).Text, Text1(1).Text, mProgr(mi), mProgr(mi + 1), True)
                  DatosTraz(mi, 2) = mObjRAcc.iMuertosHeridosProgr(Text1(0).Text, Text1(1).Text, mProgr(mi), mProgr(mi + 1), True, False)
                  DatosTraz(mi, 3) = mObjRAcc.iMuertosHeridosProgr(Text1(0).Text, Text1(1).Text, mProgr(mi), mProgr(mi + 1), True, True)
                  DatosCol(mi, 0) = mObjRAcc.iAccidentesProgr(Text1(0).Text, Text1(1).Text, mProgr(mi), mProgr(mi + 1), False)
                  DatosCol(mi, 1) = mObjRAcc.iAccidVictProgr(Text1(0).Text, Text1(1).Text, mProgr(mi), mProgr(mi + 1), False)
                  DatosCol(mi, 2) = mObjRAcc.iMuertosHeridosProgr(Text1(0).Text, Text1(1).Text, mProgr(mi), mProgr(mi + 1), False, False)
                  DatosCol(mi, 3) = mObjRAcc.iMuertosHeridosProgr(Text1(0).Text, Text1(1).Text, mProgr(mi), mProgr(mi + 1), False, True)
               Next
               'Troncales
               mj = 8
               For mi = 0 To 1
                  DatosTronc(mi, 0) = mObjRAcc.iAccidentesProgr(Text1(0).Text, Text1(1).Text, mProgr(mj), mProgr(mj + 1), True)
                  DatosTronc(mi, 1) = mObjRAcc.iAccidVictProgr(Text1(0).Text, Text1(1).Text, mProgr(mj), mProgr(mj + 1), True)
                  DatosTronc(mi, 2) = mObjRAcc.iMuertosHeridosProgr(Text1(0).Text, Text1(1).Text, mProgr(mj), mProgr(mj + 1), True, False)
                  DatosTronc(mi, 3) = mObjRAcc.iMuertosHeridosProgr(Text1(0).Text, Text1(1).Text, mProgr(mj), mProgr(mj + 1), True, True)
                  mj = 10
               Next
               For mi = 0 To 3
                  If DatosTraz(1, mi) <> 0 Then
                     DatosTraz(1, mi) = DatosTraz(1, mi) - DatosTronc(0, mi)
                  End If
                  If DatosTraz(5, mi) <> 0 Then
                     DatosTraz(5, mi) = DatosTraz(5, mi) - DatosTronc(1, mi)
                  End If
               Next
               '*** CALCULA CANTIDAD DE VEHICULOS INVOLUCRADOS TIPO DE VEHIC
               DatosVehic(0) = mObjRAcc.iCantVehicFechas(Text1(0).Text, Text1(1).Text, "'01'") 'bicicletas
               DatosVehic(1) = mObjRAcc.iCantVehicFechas(Text1(0).Text, Text1(1).Text, "'02','03'") 'motos
               DatosVehic(2) = mObjRAcc.iCantVehicFechas(Text1(0).Text, Text1(1).Text, "'04','05','15'") 'autos
               DatosVehic(3) = mObjRAcc.iCantVehicFechas(Text1(0).Text, Text1(1).Text, "'06','07','08','09','10','11','12','13','14','16'") 'camiones y otros
             
               '*** ACCIDENTES POR CAUSAS
               DatosCausas(0) = mObjRAcc.iCantCausaFechas(Text1(0).Text, Text1(1).Text, "codtipoficha='01'")
               DatosCausas(1) = mObjRAcc.iCantCausaFechas(Text1(0).Text, Text1(1).Text, "codcoliscontra1='07' and codtipoficha='01'")
               DatosCausas(2) = mObjRAcc.iCantCausaFechas(Text1(0).Text, Text1(1).Text, "causavehic<>''  and codtipoficha='01'")
               'peaton/ciclista
               DatosCausas(3) = mObjRAcc.iCantCausaFechas(Text1(0).Text, Text1(1).Text, "accidotro in ('02','03')")
               'Mat. s/calzada
               DatosCausas(4) = mObjRAcc.iCantCausaFechas(Text1(0).Text, Text1(1).Text, "codcoliscontra1='09'")
               If DatosCausas(0) <> 0 Then
                  DatosCausas(0) = DatosCausas(0) - (DatosCausas(1) + DatosCausas(2) + DatosCausas(3) + DatosCausas(4))
               End If
             
               '***CUENTA ACCIDENTES POR TIPO DE CLIMA
               DatosClima(0) = mObjRAcc.iCantCausaFechas(Text1(0).Text, Text1(1).Text, "clima1='01'")
               DatosClima(1) = mObjRAcc.iCantCausaFechas(Text1(0).Text, Text1(1).Text, "clima1='09'")
               DatosClima(2) = mObjRAcc.iCantCausaFechas(Text1(0).Text, Text1(1).Text, "clima1='03'")
               DatosClima(3) = mObjRAcc.iCantCausaFechas(Text1(0).Text, Text1(1).Text, "clima1='05'")
               DatosClima(4) = mObjRAcc.iCantCausaFechas(Text1(0).Text, Text1(1).Text, "clima1='06'")
               DatosClima(5) = mObjRAcc.iCantCausaFechas(Text1(0).Text, Text1(1).Text, "clima1='10'")
               DatosClima(6) = mObjRAcc.iCantCausaFechas(Text1(0).Text, Text1(1).Text, "clima1='02'")
               DatosClima(7) = mObjRAcc.iCantCausaFechas(Text1(0).Text, Text1(1).Text, "clima1='07'")
               DatosClima(8) = mObjRAcc.iCantCausaFechas(Text1(0).Text, Text1(1).Text, "clima1='11'")

             
               '***CUENTA HERIDOS Y MUERTOS EN ACCIDENTES
               'MUERTOS
               DatosDead(0) = mObjRAcc.iTotalMuertosCodigo(Text1(0).Text, Text1(1).Text, "(A.fallecio <> '' or A.codestado in ('04','05')) and A.letra='' and B.AccidOtro='03'")
               DatosDead(1) = mObjRAcc.iTotalMuertos3(Text1(0).Text, Text1(1).Text, " and (A.Fallecio <> '' OR A.codestado in ('04','05')) AND B.CodTipoVehic = '01'")
               DatosDead(2) = mObjRAcc.iTotalMuertosCodigo(Text1(0).Text, Text1(1).Text, " (A.fallecio <> '' or A.codestado in ('04','05'))")
               DatosDead(2) = DatosDead(2) - DatosDead(1) - DatosDead(0)
               
               DatosDead(3) = mObjRAcc.iTotalMuertosCodigo(Text1(0).Text, Text1(1).Text, " B.AccidOtro = '03' AND (A.Herido <> '' OR A.codestado in ('02','03')) AND A.Letra = ''")
               DatosDead(4) = mObjRAcc.iTotalMuertos3(Text1(0).Text, Text1(1).Text, "AND (A.Herido <> '' OR A.codestado in ('02','03')) AND B.CodTipoVehic = '01' AND A.Letra = B.Letra AND C.AccidOtro = '02'")
               DatosDead(5) = mObjRAcc.iTotalMuertosCodigo(Text1(0).Text, Text1(1).Text, " (A.Herido <> '' OR A.codestado in ('02','03'))")
               DatosDead(5) = DatosDead(5) - DatosDead(4) - DatosDead(3)
               DatosDead(6) = mObjRAcc.iTotalMuertosCodigo(Text1(0).Text, Text1(1).Text, " (fallecio <> '' or codestado in ('04','05'))")
               DatosDead(8) = mObjRAcc.iTotalMuertosCodigo(Text1(0).Text, Text1(1).Text, " (Herido = '1' OR Herido='02' or codestado='02')")
               DatosDead(9) = mObjRAcc.iTotalMuertosCodigo(Text1(0).Text, Text1(1).Text, " (Herido = '2' OR Herido='03' or codestado='03')")
               DatosDead(7) = DatosDead(8) + DatosDead(9)
               
               '***CUENTA ACCIDENTES CON TIPO DE ACCIDENT "OTROS"
               DatosOtrosAmb(0) = mObjRAcc.iCantCausaFechas(Text1(0).Text, Text1(1).Text, "AccidOtro = '03'")
               DatosOtrosAmb(1) = mObjRAcc.iCantCausaFechas(Text1(0).Text, Text1(1).Text, "AccidOtro = '01'")
               DatosOtrosAmb(2) = mObjRAcc.iCantCausaFechas(Text1(0).Text, Text1(1).Text, "AccidOtro = '04'")
               DatosOtrosAmb(3) = mObjRAcc.iCantCausaFechas(Text1(0).Text, Text1(1).Text, "AccidOtro = '02'")
               DatosOtrosAmb(5) = mObjRAcc.iCantCausaFechas(Text1(0).Text, Text1(1).Text, "AcciconOtro <> ''")
               DatosOtrosAmb(4) = mObjRAcc.iCantCausaFechas(Text1(0).Text, Text1(1).Text, "")
               DatosOtrosAmb(4) = DatosOtrosAmb(4) - (DatosOtrosAmb(0) + DatosOtrosAmb(1) + DatosOtrosAmb(2) + DatosOtrosAmb(3) + DatosOtrosAmb(5))
   
               'NO SE TOMA MAS EN CUENTA
               DatosOtrosAmb(6) = mObjRAcc.iTotalMuertosCodigo(Text1(0).Text, Text1(1).Text, "CodMedioTrasl = '01'")
               DatosOtrosAmb(7) = mObjRAcc.iTotalMuertosCodigo(Text1(0).Text, Text1(1).Text, "CodMedioTrasl = '02'")
               DatosOtrosAmb(8) = mObjRAcc.iTotalMuertosCodigo(Text1(0).Text, Text1(1).Text, "CodMedioTrasl = '03'")
               DatosOtrosAmb(9) = mObjRAcc.iTotalMuertosCodigo(Text1(0).Text, Text1(1).Text, "CodMedioTrasl = '04'")
               DatosOtrosAmb(10) = mObjRAcc.iTotalMuertosCodigo(Text1(0).Text, Text1(1).Text, "CodMedioTrasl = '05'")
               DatosOtrosAmb(11) = mObjRAcc.iTotalMuertosCodigo(Text1(0).Text, Text1(1).Text, "CodMedioTrasl = '06'")
               DatosOtrosAmb(12) = mObjRAcc.iTotalMuertosCodigo(Text1(0).Text, Text1(1).Text, "CodMedioTrasl = '07'")
               
               '*** CUENTA ACCIDENTES PRODUCIDOS POR CAUSAS DEL CONDUCTOR Y VEHICULO
               DatosCond(0) = mObjRAcc.iCantCausaFechas(Text1(0).Text, Text1(1).Text, "CodCausaCond1 = '01'")
               DatosCond(1) = mObjRAcc.iCantCausaFechas(Text1(0).Text, Text1(1).Text, "CodCausaCond1 = '02'")
               DatosCond(2) = mObjRAcc.iCantCausaFechas(Text1(0).Text, Text1(1).Text, "CodCausaCond1 = '03'")
               DatosCond(3) = mObjRAcc.iCantCausaFechas(Text1(0).Text, Text1(1).Text, "CodCausaCond1 = '04'")
               DatosCond(4) = mObjRAcc.iCantCausaFechas(Text1(0).Text, Text1(1).Text, "CodCausaCond1 = '05'")
               DatosCond(5) = mObjRAcc.iCantCausaFechas(Text1(0).Text, Text1(1).Text, "CodCausaCond1 = '06'")
               DatosCond(6) = mObjRAcc.iCantCausaFechas(Text1(0).Text, Text1(1).Text, "CodCausaCond1 = '07'")
               DatosCond(7) = mObjRAcc.iCantCausaFechas(Text1(0).Text, Text1(1).Text, "CodCausaCond1 = '08'")
               DatosCond(8) = mObjRAcc.iCantCausaFechas(Text1(0).Text, Text1(1).Text, "CodCausaCond1 = '09'")
               DatosCond(9) = mObjRAcc.iCantCausaFechas(Text1(0).Text, Text1(1).Text, "CodCausaCond1 = '10'")
               DatosCond(10) = mObjRAcc.iCantCausaFechas(Text1(0).Text, Text1(1).Text, "CodCausaCond1 = '11'")
               DatosCond(11) = mObjRAcc.iCantCausaFechas(Text1(0).Text, Text1(1).Text, "CodCausaCond1 = '12'")
               DatosCond(12) = mObjRAcc.iCantCausaFechas(Text1(0).Text, Text1(1).Text, "CausaVehic = '01'")
               DatosCond(13) = mObjRAcc.iCantCausaFechas(Text1(0).Text, Text1(1).Text, "CausaVehic = '02'")
               DatosCond(14) = mObjRAcc.iCantCausaFechas(Text1(0).Text, Text1(1).Text, "CausaVehic = '03'")
               DatosCond(15) = mObjRAcc.iCantCausaFechas(Text1(0).Text, Text1(1).Text, "CausaVehic = '04'")
               DatosCond(16) = mObjRAcc.iCantCausaFechas(Text1(0).Text, Text1(1).Text, "CausaVehic = '05'")
                         
               mData.Execute ("CREATE TABLE Auxi (CT11 Integer, CT12 Integer, CT13 Integer, CT14 Integer, CT21 Integer, CT22 Integer, CT23 Integer, CT24 Integer, CC11 Integer, CC12 Integer, CC13 Integer, CC14 Integer, CC21 Integer, CC22 Integer, CC23 Integer, CC24 Integer , CC31 Integer, CC32 Integer, CC33 Integer, CC34 Integer, CC41 Integer, CC42 Integer, CC43 Integer, CC44 Integer, CC51 Integer, CC52 Integer, CC53 Integer, CC54 Integer, CC61 Integer, CC62 Integer, CC63 Integer, CC64 Integer, CC71 Integer, CC72 Integer, CC73 Integer, CC74 Integer, " _
                             & "CP11 Integer, CP12 Integer, CP13 Integer, CP14 Integer, CP21 Integer, CP22 Integer, CP23 Integer, CP24 Integer , CP31 Integer, CP32 Integer, CP33 Integer, CP34 Integer, CP41 Integer, CP42 Integer, CP43 Integer, CP44 Integer, CP51 Integer, CP52 Integer, CP53 Integer, CP54 Integer, CP61 Integer, CP62 Integer, CP63 Integer, CP64 Integer, CP71 Integer, CP72 Integer, CP73 Integer, CP74 Integer, DV1 Integer, DV2 Integer, DV3 Integer, DV4 Integer, DC1 Integer, DC2 Integer, DC3 Integer, DC4 Integer, DC5 Integer, DT1 Integer, DT2 Integer, DT3 Integer, DT4 Integer, DT5 Integer, DT6 Integer, DT7 Integer, DT8 Integer, DT9 Integer,  " _
                             & "DM1 Integer, DM2 Integer, DM3 Integer, DM4 Integer, DM5 Integer, DM6 Integer, DM7 Integer, DM8 Integer, DM9 Integer, DM10 Integer, DOA1 Integer, DOA2 Integer, DOA3 Integer, DOA4 Integer, DOA5 Integer, DOA6 Integer, DOA7 Integer, DOA8 Integer, DOA9 Integer, DOA10 Integer, DOA11 Integer, DOA12 Integer, DOA13 Integer, DCD1 Integer, DCD2 Integer, DCD3 Integer, DCD4 Integer, DCD5 Integer, DCD6 Integer, DCD7 Integer, DCD8 Integer, DCD9 Integer, DCD10 Integer, DCD11 Integer, DCD12 Integer, DCD13 Integer, DCD14 Integer, DCD15 Integer, DCD16 Integer, DCD17 Integer)")
                
               Value1 = "" & DatosTronc(0, 0) & "," & DatosTronc(0, 1) & "," & DatosTronc(0, 2) & "," & DatosTronc(0, 3) & "," & DatosTronc(1, 0) & "," & DatosTronc(1, 1) & "," & DatosTronc(1, 2) & "," & DatosTronc(1, 3) & "," & DatosCol(0, 0) & "," & DatosCol(0, 1) & "," & DatosCol(0, 2) & "," & DatosCol(0, 3) & "," & DatosCol(1, 0) & "," & DatosCol(1, 1) & "," _
                      & "" & DatosCol(1, 2) & "," & DatosCol(1, 3) & "," & DatosCol(2, 0) & "," & DatosCol(2, 1) & "," & DatosCol(2, 2) & "," & DatosCol(2, 3) & "," & DatosCol(3, 0) & "," & DatosCol(3, 1) & "," & DatosCol(3, 2) & "," & DatosCol(3, 3) & "," & DatosCol(4, 0) & "," & DatosCol(4, 1) & "," & DatosCol(4, 2) & "," & DatosCol(4, 3) & "," & DatosCol(5, 0) & "," & DatosCol(5, 1) & "," & DatosCol(5, 2) & "," & DatosCol(5, 3) & "," & DatosCol(6, 0) & "," & DatosCol(6, 1) & "," & DatosCol(6, 2) & "," & DatosCol(6, 3) & "," & DatosTraz(0, 0) & "," & DatosTraz(0, 1) & "," & DatosTraz(0, 2) & "," & DatosTraz(0, 3) & "," & DatosTraz(1, 0) & "," & DatosTraz(1, 1) & "," & DatosTraz(1, 2) & "," & DatosTraz(1, 3) & "," & DatosTraz(2, 0) & "," & DatosTraz(2, 1) & "," & DatosTraz(2, 2) & "," & DatosTraz(2, 3) & "," & DatosTraz(3, 0) & "," & DatosTraz(3, 1) & "," & DatosTraz(3, 2) & "," & DatosTraz(3, 3) & "," & DatosTraz(4, 0) & "," & DatosTraz(4, 1) & ""
               Value2 = "," & DatosTraz(4, 2) & "," & DatosTraz(4, 3) & "," & DatosTraz(5, 0) & "," & DatosTraz(5, 1) & "," & DatosTraz(5, 2) & "," & DatosTraz(5, 3) & "," & DatosTraz(6, 0) & "," & DatosTraz(6, 1) & "," & DatosTraz(6, 2) & "," & DatosTraz(6, 3) & "," & DatosVehic(0) & "," & DatosVehic(1) & "," & DatosVehic(2) & "," & DatosVehic(3) & "," & DatosCausas(0) & "," & DatosCausas(1) & "," & DatosCausas(2) & "," & DatosCausas(3) & "," & DatosCausas(4) & "," & DatosClima(0) & "," & DatosClima(1) & "," & DatosClima(2) & "," & DatosClima(3) & "," & DatosClima(4) & "," & DatosClima(5) & "," & DatosClima(6) & "," & DatosClima(7) & "," & DatosClima(8) & ", " & DatosDead(0) & "," & DatosDead(1) & "," & DatosDead(2) & "," & DatosDead(3) & "," & DatosDead(4) & "," & DatosDead(5) & "," & DatosDead(6) & "," & DatosDead(7) & "," & DatosDead(8) & "," & DatosDead(9) & "," _
                      & "" & DatosOtrosAmb(0) & "," & DatosOtrosAmb(1) & "," & DatosOtrosAmb(2) & "," & DatosOtrosAmb(3) & "," & DatosOtrosAmb(4) & "," & DatosOtrosAmb(5) & "," & DatosOtrosAmb(6) & "," & DatosOtrosAmb(7) & "," & DatosOtrosAmb(8) & "," & DatosOtrosAmb(9) & "," & DatosOtrosAmb(10) & "," & DatosOtrosAmb(11) & "," & DatosOtrosAmb(12) & "," & DatosCond(0) & "," & DatosCond(1) & "," & DatosCond(2) & "," & DatosCond(3) & "," & DatosCond(4) & "," & DatosCond(5) & "," & DatosCond(6) & "," & DatosCond(7) & "," & DatosCond(8) & "," & DatosCond(9) & "," & DatosCond(10) & "," & DatosCond(11) & "," & DatosCond(12) & "," & DatosCond(13) & "," & DatosCond(14) & "," & DatosCond(15) & "," & DatosCond(16) & ""
                
               mData.Execute ("INSERT INTO Auxi (CT11,CT12,CT13,CT14,CT21,CT22,CT23,CT24,CC11,CC12,CC13,CC14,CC21,CC22,CC23,CC24,CC31,CC32,CC33,CC34,CC41,CC42,CC43,CC44,CC51,CC52,CC53,CC54,CC61,CC62,CC63,CC64,CC71,CC72,CC73,CC74,CP11,CP12,CP13,CP14,CP21,CP22,CP23,CP24,CP31,CP32,CP33,CP34,CP41,CP42,CP43,CP44,CP51,CP52,CP53,CP54,CP61,CP62,CP63,CP64,CP71,CP72,CP73,CP74,DV1,DV2,DV3,DV4,DC1,DC2,DC3,DC4,DC5,DT1,DT2,DT3,DT4,DT5,DT6,DT7,DT8,DT9,DM1,DM2,DM3,DM4,DM5,DM6,DM7,DM8,DM9,DM10,DOA1,DOA2,DOA3,DOA4,DOA5,DOA6,DOA7,DOA8,DOA9,DOA10,DOA11,DOA12,DOA13,DCD1,DCD2,DCD3,DCD4,DCD5,DCD6,DCD7,DCD8,DCD9,DCD10,DCD11,DCD12,DCD13,DCD14,DCD15,DCD16,DCD17) VALUES (" & Value1 & Value2 & ")")
               Set mAuxi = mData.OpenRecordset("select * from auxi")
               mAuxi.Close
               For mi = 0 To mData.TableDefs.Count - 1
                  If mData.TableDefs(mi).Name = "Auxi" Then
                     CrystalReport1.WindowTitle = "Reporte " & mReporte
                     CrystalReport1.DataFiles(0) = App.Path & "\RegAccidentes\FichaAccid.mdb"
                     CrystalReport1.Formulas(0) = "Listado = 'Total de Accidentes desde " & Text1(0).Text & " hasta " & Text1(1).Text & ".'"
                     CrystalReport1.ReportFileName = App.Path & "\RegAccidentes\rep18.rpt"
                     CrystalReport1.WindowState = crptMaximized
                     CrystalReport1.Action = 1
                     mi = 300
                     mImprimir = True
                  End If
               Next
            
            Case "Demoras de Patrullas" 'MySQL
               mj = 0
               For mi = 0 To 12
                 Horas(mi) = mj
                 mj = mj + 2
               Next
               Horas(13) = 10000
               'mData.Execute ("CREATE TABLE Auxi (Flagg INTEGER,xFicha TEXT,xFecha TEXT,xHora1 TEXT,xHora2 TEXT,xDemora INTEGER)")
               mData.Execute ("CREATE TABLE Auxi (Flagg INTEGER,xFicha TEXT,xFecha TEXT,xHora1 TEXT,xHora2 TEXT,xDemora INTEGER,xCodAlfa TEXT)")
               Set mRec = mObjRAcc.oTabla("Ficha", "where fecha between '" & Format(Text1(0).Text, "yyyy-mm-dd") & "' and '" & Format(Text1(1).Text, "yyyy-mm-dd") & "' order by 1")
               Do While Not mRec.EOF
                 'mData.Execute "INSERT INTO Auxi (Flagg,xFicha,xFecha,xHora1,xHora2,xDemora) VALUES (1,'" & mRec!NroOrden & "','" & mRec!Fecha & "','" & Format(mRec!hora, "hh:mm") & "','" & mRec!HoraLlegada & "',0)"
                 mData.Execute "INSERT INTO Auxi (Flagg,xFicha,xFecha,xHora1,xHora2,xDemora, xCodAlfa) VALUES (1,'" & mRec!NroOrden & "','" & mRec!Fecha & "','" & Format(mRec!hora, "hh:mm") & "','" & mRec!HoraLlegada & "',0,'" & mRec!CodAlfa & "')"
                  mRec.MoveNext
               Loop
               mRec.Close
               Set mAuxi = mData.OpenRecordset("SELECT xFicha,DATEDIFF('n',xHora1,xHora2) As Demora FROM Auxi")
               Do While Not mAuxi.EOF
                  mData.Execute ("UPDATE Auxi SET xDemora = " & mAuxi!demora & " WHERE xFicha='" & mAuxi!xFicha & "'")
                  mAuxi.MoveNext
               Loop
               mAuxi.Close
               For mi = 0 To 12
                  Set Auxi = mData.OpenRecordset("SELECT COUNT(*) AS Total  FROM Auxi WHERE xDemora >= " & Horas(mi) & " and xDemora < " & Horas(mi + 1) & " AND xHora1 <> ''")
                  Value1 = "De " & Horas(mi) & " a " & Horas(mi + 1)
                  mData.Execute "INSERT INTO Auxi (Flagg,xFicha,xFecha,xHora1,xHora2,xDemora) VALUES (2,'','','','" & Value1 & "'," & Auxi!Total & ")"
                  Auxi.Close
               Next
               mData.Execute ("UPDATE Auxi SET xHora2 = 'Mayor a 24' WHERE xHora2 = '" & Value1 & "'")
               Set mAuxi = mData.OpenRecordset("SELECT * FROM Auxi")
               If PrintRep(mAuxi) Then
                  CrystalReport1.ReportFileName = App.Path & "\RegAccidentes\rep21.rpt"
                  CrystalReport1.Action = 1
               End If
         End Select
         If mImprimir Then
            Text1(0).Text = ""
            Text1(1).Text = ""
            Combo1.ListIndex = -1
         End If
         sMsgEspere Me, "", False
      End If
   Else
      Unload Me
   End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   If Index = 2 Then
      KeyAscii = fNumeroKeyPress(KeyAscii)
   Else
      KeyAscii = fDateKeyPress(Text1(Index), KeyAscii)
   End If
End Sub

Private Function PrintRep(mRecord As Variant) As Boolean
   If mRecord.EOF Then
      MsgBox "Consulta sin Resultado", vbInformation, "Atención!!!"
      mRecord.Close
      PrintRep = False
   Else
      mRecord.MoveFirst
      mRecord.MoveLast
      mRecord.Close
      CrystalReport1.WindowTitle = "Reporte " & mReporte
      CrystalReport1.DataFiles(0) = "FichaAccid.mdb"
      CrystalReport1.Formulas(0) = "Listado = 'Listado de Reporte " & mReporte & " del " & Text1(0).Text & " al " & Text1(1).Text & ".'"
      CrystalReport1.WindowState = crptMaximized
      PrintRep = True
   End If
   mImprimir = PrintRep
End Function

Private Sub sInitForm()
   Set mData = OpenDatabase(App.Path & "\RegAccidentes\FichaAccid.mdb")
   Label1.Caption = "Reporte " & mReporte
   Label1.Left = (Frame1.Width - Label1.Width) / 2
   Me.Height = 4455
   Me.Width = 6915
   sAlinearForm Me
   If mReporte = "Evaluación" Then
      Label2.Visible = False
      Label3.Visible = False
      Label4.Visible = False
      Text1(0).Visible = False
      Text1(1).Visible = False
      Text1(2).Visible = False
      Combo1.Visible = False
   Else
      Label2.Visible = True
      Label3.Visible = True
      Label4.Visible = False
      Text1(0).Visible = True
      Text1(1).Visible = True
      Text1(2).Visible = False
      Combo1.Visible = False
   End If
   If mReporte = "Hora Día" Then
   Else
      If mReporte = "Tipo Vehículo" Then
         Label4.Visible = True
         Label4.Caption = "Tipo Vehic >="
         Combo1.Visible = True
         Set mRec = mObjRAcc.oTabla("TipoVehiculo", "")
         Do While Not mRec.EOF
            Combo1.AddItem mRec!CodTipoVehic & " " & Left(mRec!descripcion, 6)
            mRec.MoveNext
         Loop
         mRec.Close
      End If
   End If
End Sub
