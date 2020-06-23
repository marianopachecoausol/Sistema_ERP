VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form RNov7_frm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Módulo de Reportes."
   ClientHeight    =   4320
   ClientLeft      =   540
   ClientTop       =   1335
   ClientWidth     =   7710
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   7710
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   50
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2160
         TabIndex        =   24
         Top             =   2280
         Visible         =   0   'False
         Width           =   4035
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Agrupar por Códigos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   3540
         TabIndex        =   14
         Top             =   2700
         Visible         =   0   'False
         Width           =   2175
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   480
         Top             =   480
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Sin Detalle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   1920
         TabIndex        =   13
         Top             =   2640
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ListBox List1 
         Height          =   1035
         Left            =   6000
         TabIndex        =   17
         Top             =   2880
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3120
         MaxLength       =   7
         TabIndex        =   15
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   3
         Left            =   4080
         MaxLength       =   5
         TabIndex        =   4
         Text            =   "23:59"
         Top             =   1680
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   2
         Left            =   2760
         MaxLength       =   5
         TabIndex        =   2
         Text            =   "00:00"
         Top             =   1680
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   1
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   0
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   2040
         TabIndex        =   6
         Top             =   3120
         Width           =   3375
         Begin VB.CommandButton Command2 
            Caption         =   "Cancelar"
            Height          =   495
            Index           =   1
            Left            =   1920
            TabIndex        =   18
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Ok"
            Default         =   -1  'True
            Height          =   495
            Index           =   0
            Left            =   480
            TabIndex        =   16
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Cancelar"
            Height          =   495
            Index           =   1
            Left            =   2040
            TabIndex        =   23
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Aceptar"
            Height          =   495
            Index           =   0
            Left            =   360
            TabIndex        =   22
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1440
         TabIndex        =   21
         Top             =   2880
         Width           =   4095
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Selecc."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   6480
         TabIndex        =   20
         Top             =   1755
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Móviles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   6240
         TabIndex        =   19
         Top             =   2640
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar por Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   1
         Left            =   3960
         TabIndex        =   12
         Top             =   2280
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar por Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   0
         Left            =   1920
         TabIndex        =   11
         Top             =   2280
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Móvil"
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
         Left            =   4680
         TabIndex        =   10
         Top             =   1440
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Index           =   1
         Left            =   3120
         TabIndex        =   9
         Top             =   1440
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicio"
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
         Left            =   1560
         TabIndex        =   8
         Top             =   1440
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   960
         TabIndex        =   7
         Top             =   480
         Width           =   5655
      End
   End
End
Attribute VB_Name = "RNov7_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mData As Database
Dim mObj As New clRNov
Dim mRec As New ADODB.Recordset
Public mTabla As String

Private Sub Form_Load()
   sAlinearForm Me
   Set mData = OpenDatabase(App.Path & "\RegNovedades\RegNovPlus.mdb")
   sInitForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
   mData.Close
   Set mData = Nothing
   Set mObj = Nothing
   Set mRec = Nothing
End Sub

Private Sub Combo1_Click()
   Select Case mTabla
      Case "9"
         If Left(Combo1.Text, 3) <> "TOD" Then
            Label4.Visible = True
            Label5.Visible = True
            List1.Visible = True
         Else
            Label4.Visible = False
            Label5.Visible = False
            List1.Visible = False
         End If
      Case "10"
         If Combo1.ListIndex > -1 Then
            Text1(0).Text = Left(Combo1.Text, 10)
            Text1(1).Text = Left(Combo1.Text, 10)
         End If
   End Select
End Sub

Private Sub Command1_Click(Index As Integer)
Dim mObjRAcc As New clRAcc
Dim mObjAcc As New clAccess
Dim mRec1 As New ADODB.Recordset
Dim mRec2 As New ADODB.Recordset
Dim mBorrar As Boolean
Dim mFlag As Boolean
Dim xFlag As Boolean
Dim mArribo As Boolean
Dim mAuxi
 
Dim mMov, mTitulo, mSubTitulo As String
Dim mMovAux As String
Dim mRut As String
Dim mVaria As String
Dim xCodServ As String
Dim xServGRU(9) As String
Dim xServGRP(5) As String
Dim xFin As String
Dim mHora As String
Dim mCodRamal As String
Dim mSent As String
Dim mKm As String


Dim xDemora1 As Date
Dim xOcupa1 As Date
Dim xAsig As Date
Dim xFecha1 As Date
Dim xFecha2 As Date

Dim mI, mJ As Integer
Dim xMax As Integer
Dim xMaxTodo As Integer
Dim xAmbu(14, 3) As Integer
Dim TotMoviles As Integer
Dim TotMovilesP As Integer
Dim mCuadro(19) As Integer
Dim xAtend As Integer
Dim mKmF As Integer

Dim xDemora As Double
Dim xOcupa As Double
Dim mSumDem As Double
Dim mSumOcup As Double
Dim mSumCanc As Double
Dim mSumServ As Double

Dim xDemoraStr As String


   If Index = 0 Then
      mObjAcc.mBorrarAuxi "\RegNovedades\RegNovPlus", "Auxi"
      If sValida Then
         mFlag = True
         mVaria = ""
         mHora = ""
         sMsgEspere Me, "Generando informe... aguarde un momento.", True
         Select Case mTabla
            Case "0"
               If Hora_ok(Text1(2).Text) And Hora_ok(Text1(3).Text) Then
                  mTitulo = "REPORTE Novedades de Móviles."
                  mData.Execute ("CREATE TABLE Auxi (mFecha TEXT, mHora TEXT, Codigo TEXT, Ramal TEXT, Km DOUBLE, Sent TEXT, Descripcion TEXT, Climax TEXT, Mov1 TEXT, Mov2 TEXT, Mov3 TEXT)")
                  If Combo1.Text = "TODOS" Then
                     mSubTitulo = "Todos los Móviles.  (" & Text1(0).Text & " al " & Text1(1).Text & ")"
                     Set mRec = mObj.oMovilesGCO("PAT','GRU")
                     Do While Not mRec.EOF
                        mMov = mMov & "'" & mRec!Codigo & "',"
                        mRec.MoveNext
                     Loop
                     mRec.Close
                     mMov = Mid(mMov, 1, Len(mMov) - 1)
                     CrystalReport1.ReportFileName = App.Path & "\RegNovedades\" & "Rep04.rpt"
                  Else
                     mMov = "'" & Right(Combo1.Text, 4) & "'"
                     mSubTitulo = "Móvil: " & Right(Combo1.Text, 4) & ".  (" & Text1(0).Text & " al " & Text1(1).Text & ")"
                     CrystalReport1.ReportFileName = App.Path & "\RegNovedades\" & "Rep03.rpt"
                  End If
                  Set mRec = mObj.oNovedadesMovil(Text1(0).Text & " " & Text1(2).Text, Text1(1).Text & " " & Text1(3).Text, mMov)
                  
                  
                  Do While Not mRec.EOF
                     mCodRamal = Mid(mObj.sTablaDescr("ramales", " codigo=" & mRec!codramal, 2), 2, 2)
                     mSent = Left(mObj.sTablaDescr("sentidos", "codigo=" & mRec!sent, 1), 2)
                     mKm = mRec!km
                     mKm = Replace(mKm, ",", ".")
                     mData.Execute ("INSERT INTO Auxi (mFecha,mHora,Codigo,Ramal, Km,Sent,Descripcion,Climax,Mov1,Mov2,Mov3) VALUES ('" & Left(mRec!Fecha, 10) & "','" & Mid(mRec!Fecha, 12, 14) & "','" & mRec!Codigo & "','" & mCodRamal & "', " & mKm & ",'" & mSent & "','" & mRec!descripcion & "','" & mRec!climax & "','" & mRec!Mov1 & "','" & mRec!Mov2 & "','" & mRec!Mov3 & "')")
                     mRec.MoveNext
                  Loop
                  mRec.Close
               Else
                  mFlag = False
               End If
                 
            Case "1" 'MySQL
               mTitulo = "REPORTE Kilometros Recorridos"
               mData.Execute ("CREATE TABLE Auxi (mFecha TEXT, mHora TEXT, CodMovil TEXT, KmInicial TEXT, KmFinal TEXT, Chofer TEXT,mPatrullero TEXT,Descripcion TEXT)")
               CrystalReport1.ReportFileName = App.Path & "\RegNovedades\" & "Rep05.rpt"
               If Combo1.Text = "TODOS" Then
                  mMov = ""
                  mSubTitulo = "Todos los Móviles. (" & Text1(0).Text & " al " & Text1(1).Text & ")"
               Else
                  mMov = Right(Combo1.Text, 4)
                  mSubTitulo = "Móvil: " & mMov & ". (" & Text1(0).Text & " al " & Text1(1).Text & ")"
               End If
               Set mRec = mObj.oMovilKms(Text1(0).Text, Text1(1).Text, mMov)
               Do While Not mRec.EOF
                  xFecha1 = Left(mRec.Fields(0), 10)
                  mVaria = Mid(mRec.Fields(0), 12, 5)
                  mData.Execute ("INSERT INTO Auxi (mFecha,mHora,CodMovil,KmInicial,KmFinal,Chofer,mPatrullero,Descripcion) VALUES ('" & xFecha1 & "','" & mVaria & "','" & mRec.Fields(1) & "','" & mRec.Fields(2) & "','" & mRec.Fields(3) & "','" & mRec.Fields(4) & "','" & mRec.Fields(5) & "','" & mRec.Fields(6) & "')")
                  mRec.MoveNext
               Loop
               mRec.Close
                 
            Case "2" 'Arribos VITTAL
               mMov = Left(Combo1.Text, 4)
               mSubTitulo = "Móvil: " & mMov & ".  (" & Text1(0).Text & " al " & Text1(1).Text & ")"
               If Left(Combo1.Text, 8) <> "AMBU EXT" Then ' lo nuevo
                  mTitulo = "REPORTE Arribos de Móvil " & Left(Combo1.Text, 4)
                  Set mRec = mObj.oNovedadesFecha(Text1(0).Text, Text1(1).Text, "and codnov in('M','L','T') ", "and Mov1 = '" & mMov & "'", "order by fecha")
                  If Not mRec.EOF Then
                     mData.Execute ("CREATE TABLE Auxi (xCodigo TEXT,Fecha DATE, Ramal TEXT, Km TEXT,Sent TEXT,Servicio TEXT,Asign TEXT,Arribo TEXT,Demora TEXT,Codigo TEXT,Destino TEXT, Total INTEGER, Rojo INTEGER)")
                     mVaria = mRec!Codigo
                     Do While Not mRec.EOF
                        mCodRamal = Mid(mObj.sTablaDescr("ramales", " codigo=" & mRec!codramal, 2), 2, 2)
                        mSent = Left(mObj.sTablaDescr("sentidos", "codigo=" & mRec!sent, 1), 2)
  
                        Select Case mRec!CodNov
                           Case "M"
                              xCodServ = mObj.sTablaDescr("varios", "codigo='" & mRec!codserv1 & "'", 1)
                              xFin = mObj.sTablaDescr("varios", "codigo='" & mRec!TipoNov & "'", 1)
                              mKm = mRec!km
                              mKm = Replace(mKm, ",", ".")
                              mData.Execute "INSERT INTO Auxi (xCodigo,Fecha,Ramal, Km,Sent,Servicio,Asign,Arribo,Demora,Codigo,Destino,Total, Rojo) VALUES ('" & mRec!Codigo & "','" & mRec!Fecha & "','" & mCodRamal & "', '" & mKm & "','" & mSent & "','" & mRec!descripcion & "','" & Format(mRec!Fecha, "hh:mm") & "','','','" & xCodServ & "','" & xFin & "',0,0)"
                              
                           Case "L"
                              Set mAuxi = mData.OpenRecordset("SELECT TOP 1 * FROM Auxi WHERE xCodigo = '" & mRec!Codigo & "' AND Arribo = ''")
                              If Not mAuxi.EOF Then
                                 xDemora = DateDiff("n", mAuxi!Fecha, mRec!Fecha)
                                 mData.Execute ("UPDATE Auxi SET Arribo='" & Format(mRec!Fecha, "hh:mm") & "',Demora='" & xDemora & "' WHERE xCodigo = '" & mRec!Codigo & "' AND Fecha = #" & Format(mAuxi!Fecha, "mm/dd/yyyy hh:mm:ss") & "# AND Asign = '" & mAuxi!Asign & "'")
                              End If
                              mAuxi.Close
                           
                           Case "T"
                              Set mAuxi = mData.OpenRecordset("SELECT TOP 1 * FROM Auxi WHERE xCodigo = '" & mRec!Codigo & "' AND Arribo = ''")
                              If Not mAuxi.EOF Then
                                 mData.Execute ("UPDATE Auxi SET Servicio = '" & mRec!descripcion & "' WHERE xCodigo = '" & mRec!Codigo & "' AND Fecha = #" & Format(mAuxi!Fecha, "mm/dd/yyyy hh:mm:ss") & "#")
                              End If
                              mAuxi.Close
                        End Select
                        mRec.MoveNext
                     Loop
                     Set mAuxi = mData.OpenRecordset("SELECT COUNT(*) AS Total FROM Auxi WHERE Codigo = 'ROJO'")
                     mI = mAuxi!Total
                     mData.Execute ("UPDATE Auxi SET Rojo = '" & mI & "'")
                     mAuxi.Close
                     Set mAuxi = mData.OpenRecordset("SELECT COUNT(*) AS Total FROM Auxi")
                     mData.Execute ("UPDATE Auxi SET Total = '" & mAuxi!Total & "'")
                     CrystalReport1.ReportFileName = App.Path & "\RegNovedades\" & "Rep21.rpt"
                  Else
                     mFlag = False
                  End If
                  mRec.Close
               Else
                  mTitulo = "REPORTE Arribos de Móvil AMBU EXTERNOS"
                  Set mRec = mObj.oTabla("arribos_ambu", " where fecha between '" & Format(Text1(0).Text, "yyyy-mm-dd") & "' and '" & Format(Text1(1).Text, "yyyy-mm-dd") & "'")
                  If Not mRec.EOF Then
                     mData.Execute ("CREATE TABLE Auxi (xCodigo TEXT,Fecha DATE, Ramal TEXT, Km TEXT, Sent TEXT,Servicio TEXT,Asign TEXT,Arribo TEXT,Demora TEXT,Codigo TEXT,Destino TEXT, Total INTEGER, Rojo INTEGER)")
                     Do While Not mRec.EOF
                        mCodRamal = Mid(mObj.sTablaDescr("ramales", " codigo=" & mRec!codramal, 2), 2, 2)
                        mSent = Left(mObj.sTablaDescr("sentidos", "codigo=" & mRec!sent, 1), 2)
                        xAtend = DateDiff("n", mRec!asignado, mRec!arribo)
                        Set mRec1 = mObj.oTabla("novedades2", "where codigo='" & mRec!CodAlfa & "' and fecha='" & Format(mRec!Fecha, "yyyy-mm-dd hh:mm:ss") & "' and codnov='M'")
                        If Not mRec1.EOF Then
                           xCodServ = mObj.sTablaDescr("varios", "codigo='" & mRec1!codserv1 & "'", 1)
                           mKm = mRec1!km
                           mKm = Replace(mKm, ",", ".")
                           mData.Execute "insert into Auxi values ('" & mRec!CodAlfa & "','" & mRec!Fecha & "','" & mCodRamal & "','" & mKm & "','" & mSent & "','" & mRec1!descripcion & "','" & Format(mRec!asignado, "hh:mm") & "','" & Format(mRec!arribo, "hh:mm") & "','" & xAtend & "','" & xCodServ & "','',0,0)"
                        End If
                        mRec1.Close
                        mRec.MoveNext
                     Loop
                     Set mAuxi = mData.OpenRecordset("SELECT COUNT(*) AS Total FROM Auxi WHERE Codigo = 'ROJO'")
                     mI = mAuxi!Total
                     mData.Execute ("UPDATE Auxi SET Rojo = '" & mI & "'")
                     mAuxi.Close
                     Set mAuxi = mData.OpenRecordset("SELECT COUNT(*) AS Total FROM Auxi")
                     mData.Execute ("UPDATE Auxi SET Total = '" & mAuxi!Total & "'")
                     CrystalReport1.ReportFileName = App.Path & "\RegNovedades\" & "Rep21.rpt"
                  End If
                  mRec.Close
               End If
                 
            Case "3" 'Arribos Ambulancias
               mTitulo = "REPORTE Cuadro de Arribos de Ambulancias "
               mSubTitulo = "Período: " & Text1(0).Text & " al " & Text1(1).Text & ""
               Set mRec = mObj.oCodigosAmbu(Text1(0).Text, Text1(1).Text)
               For mI = 0 To 13
                  For mJ = 0 To 2
                     xAmbu(mI, mJ) = 0
                  Next
               Next
               If Not mRec.EOF Then
                  mData.Execute ("CREATE TABLE Auxi (DA1 INTEGER,DA2 INTEGER,DA3 INTEGER,DA4 INTEGER,DA5 INTEGER,DA6 INTEGER,DA7 INTEGER,DA8 INTEGER,DA9 INTEGER,DA10 INTEGER,DA11 INTEGER,DA12 INTEGER,DA13 INTEGER,DA14 INTEGER, " _
                              & "DR1 INTEGER,DR2 INTEGER,DR3 INTEGER,DR4 INTEGER,DR5 INTEGER,DR6 INTEGER,DR7 INTEGER,DR8 INTEGER,DR9 INTEGER,DR10 INTEGER,DR11 INTEGER,DR12 INTEGER,DR13 INTEGER,DR14 INTEGER, " _
                              & "DO1 INTEGER,DO2 INTEGER,DO3 INTEGER,DO4 INTEGER,DO5 INTEGER,DO6 INTEGER,DO7 INTEGER,DO8 INTEGER,DO9 INTEGER,DO10 INTEGER,DO11 INTEGER,DO12 INTEGER,DO13 INTEGER,DO14 INTEGER)")
                  mVaria = mRec!Codigo
                  xAsig = mRec!Fecha
                  xCodServ = mRec!codserv1
                  Do While Not mRec.EOF
                     If mRec!CodNov = "M" Then
                        mVaria = mRec!Codigo          'guarda el código alfanumérico
                        xAsig = mRec!Fecha            'Fecha de asignación
                        xCodServ = mRec!codserv1      'Cod de servicio, 1, 2 o 3
                     Else
                        If mRec!Codigo = mVaria And mRec!CodNov = "L" Then
                           xDemora = DateDiff("n", xAsig, mRec!Fecha)
                           For mI = 1 To 13 '14
                              If xDemora > ((mI * 2) - 2) And xDemora <= (mI * 2) Then
                                 xAmbu((mI - 1), (Val(xCodServ) - 1)) = xAmbu((mI - 1), (Val(xCodServ) - 1)) + 1
                                 mI = 99
                              End If
                           Next
                           If mI = 14 And xDemora > 26 Then
                              xAmbu((mI - 1), (Val(xCodServ) - 1)) = xAmbu((mI - 1), (Val(xCodServ) - 1)) + 1
                           End If
                        End If
                     End If
                     mRec.MoveNext
                  Loop
                  mVaria = ""
                  For mJ = 0 To 2
                     For mI = 0 To 13
                        mVaria = mVaria & xAmbu(mI, mJ) & ","
                     Next
                  Next
                  mVaria = Mid(mVaria, 1, Len(mVaria) - 1)
                  mData.Execute ("insert into AUXI values (" & mVaria & ")")
                  CrystalReport1.ReportFileName = App.Path & "\RegNovedades\" & "Rep22i.rpt"
               Else
                  MsgBox "No Existen Datos Para el Período", vbInformation, sMessage
               End If
               mRec.Close
   
            Case "4" 'MySQL
               mTitulo = "REPORTE de Emisoras"
               mData.Execute ("CREATE TABLE Auxi (mFecha TEXT,mHora TEXT,mEmisora TEXT,mDetalle TEXT)")
               If Combo1.Text = "TODAS" Then
                  mSubTitulo = "Todas las Emisoras. (" & Text1(0).Text & " al " & Text1(1).Text & ")"
                  mMov = ""
                  CrystalReport1.ReportFileName = App.Path & "\RegNovedades\" & "Rep06.rpt"
               Else
                  mMov = " AND LEFT(Descripcion,3) = '" & Right(Combo1.Text, 3) & "'"
                  mSubTitulo = "Emisora: " & mMov & " - " & Trim(Left(Combo1.Text, 25)) & ". (" & Text1(0).Text & " al " & Text1(1).Text & ")"
                   mSubTitulo = "Emisora: " & Trim(Left(Combo1.Text, 25)) & ". (" & Text1(0).Text & " al " & Text1(1).Text & ")"
                  CrystalReport1.ReportFileName = App.Path & "\RegNovedades\" & "Rep07.rpt"
               End If
               Set mRec = mObj.oNovedadesFecha(Text1(0).Text, Text1(1).Text, " and CodNov = 'F'", mMov, "order by fecha asc")
               Do While Not mRec.EOF
                  xFecha1 = Left(mRec!Fecha, 10)
                  mVaria = Mid(mRec!Fecha, 12, 5)
                  mData.Execute ("INSERT INTO Auxi (mFecha,mHora,mEmisora,mDetalle) VALUES ('" & xFecha1 & "','" & mVaria & "','" & Mid(mRec!descripcion, 7, 25) & "','" & Mid(mRec!descripcion, 36, 74) & "')")
                  mRec.MoveNext
               Loop
               mRec.Close
                
            Case "5"
               mTitulo = "REPORTE Origenes de Novedades"
               mData.Execute ("CREATE TABLE Auxi (CodOrigen TEXT, detalle TEXT, mTotal INTEGER)")
               If Combo1.Text = "TODOS" Then
                  mMov = ""
                  mSubTitulo = "Todos los Origenes . (" & Text1(0).Text & " al " & Text1(1).Text & ")"
                  CrystalReport1.ReportFileName = App.Path & "\RegNovedades\" & "Rep08.rpt"
               Else
                  mMov = " and codorigen='" & Left(Combo1.Text, 3) & "' "
                  mSubTitulo = "Origen:" & Trim(Combo1.Text) & ". (" & Text1(0).Text & " al " & Text1(1).Text & ")"
                  CrystalReport1.ReportFileName = App.Path & "\RegNovedades\" & "Rep08.rpt"
               End If
               Set mRec = mObj.iCountCodOrig(Text1(0).Text, Text1(1).Text, mMov)
               Do While Not mRec.EOF
                  mData.Execute ("INSERT INTO Auxi (CodOrigen,detalle,mTotal) VALUES ('" & mRec.Fields(0) & "','" & mRec.Fields(1) & "'," & mRec.Fields(2) & ")")
                  mRec.MoveNext
               Loop
               mRec.Close
                
            Case "6"
               mTitulo = "REPORTE Resumen de Rutinas"
               mSubTitulo = "Rutinas. (" & Text1(0).Text & " al " & Text1(1).Text & ")"
               mData.Execute ("CREATE TABLE Auxi (R1A integer,R1D Integer,R2A integer,R2D Integer,R3A integer,R3D Integer,R4A integer,R4D Integer,R5A integer,R5D Integer,R6A integer,R6D Integer,R7A integer,R7D Integer,R8A integer,R8D Integer,R9A integer,R9D Integer)")
               mData.Execute ("INSERT INTO Auxi (R1A,R1D,R2A,R2D,R3A,R3D,R4A,R4D,R5A,R5D,R6A,R6D,R7A,R7D,R8A,R8D,R9A,R9D) VALUES (0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)")
               For mI = 1 To 9
                  mRut = "RU" & mI
                  mJ = mObj.iCountTabla("novedades", Text1(0).Text, Text1(1).Text, "*", " and Left(Descripcion,3) = '" & mRut & "' and codnov in('C','MN') and sent in('A','K')")
                  mData.Execute ("UPDATE Auxi SET R" & mI & "A = " & mJ & "")
                  mJ = mObj.iCountTabla("novedades", Text1(0).Text, Text1(1).Text, "*", "and Left(Descripcion,3) = '" & mRut & "' and codnov in('C','MN') and sent in('D','B','T')")
                  mData.Execute ("UPDATE Auxi SET R" & mI & "D = " & mJ & " ")
               Next
               CrystalReport1.ReportFileName = App.Path & "\RegNovedades\" & "Rep09.rpt"
              
            Case "7" 'MySQL
               mTitulo = "REPORTE Rutinas por Fechas"
               mSubTitulo = "Fecha: (" & Text1(0).Text & " al " & Text1(1).Text & ")"
               mData.Execute ("CREATE TABLE Auxi (Codigo TEXT, CodRut TEXT,Fecha DATE, Km DOUBLE, Sent TEXT, KmF DOUBLE, SentF TEXT)")
               Set mRec = mObj.oNovedadesFecha(Text1(0).Text, Text1(1).Text, "and (codnov in ('C','MN') or Descripcion like '%Fin de Rutina.%')", "", "order by codigo,fecha")

               
               Do While Not mRec.EOF
                  mKm = mRec!km
                  mKm = Replace(mKm, ",", ".")
                  If mRec!CodNov = "C" Or mRec!CodNov = "MN" Then
                     mData.Execute ("INSERT INTO Auxi (Codigo,CodRut,Fecha,Km,Sent,KmF,SentF) VALUES " _
                        & " ('" & mRec!Codigo & "','" & Left(mRec!descripcion, 3) & "','" & Left(mRec!Fecha, 10) & "'," & mKm & ",'" & mRec!sent & "',99,'X')")
                  Else
                     mData.Execute "update Auxi set kmf=" & mKm & ", sentf='" & mRec!sent & "' where codigo='" & mRec!Codigo & "'"
                  End If
                  mRec.MoveNext
               Loop
               mRec.Close
               CrystalReport1.ReportFileName = App.Path & "\RegNovedades\" & "Rep10.rpt"
                              
            Case "8" 'MySQL
               mTitulo = "REPORTE Rutinas Detalle"
               CrystalReport1.ReportFileName = App.Path & "\RegNovedades\" & "Rep11.rpt"
               mData.Execute ("CREATE TABLE Auxi (mFecha TEXT,mHora TEXT,Km DOUBLE,Sent TEXT,Codigo TEXT,Descripcion TEXT,Mov1 TEXT, Ramal TEXT)")
               If Combo1.Text = "TODOS" Then
                  mSubTitulo = "Todos los Móviles. (" & Text1(0).Text & " al " & Text1(1).Text & ")"
                  mMov = ""
               Else
                  mMov = "and mov1='" & Right(Combo1.Text, 4) & "'"
                  mSubTitulo = "Móvil: " & Trim(Left(Combo1.Text, 25)) & ". (" & Text1(0).Text & " al " & Text1(1).Text & ")"
               End If
               Set mRec = mObj.oNovedadesFecha(Text1(0).Text, Text1(1).Text, " and codnov in ('C','MN')", mMov, "order by fecha")
               Do While Not mRec.EOF
                  mVaria = mVaria & "'" & mRec!Codigo & "',"
                  mRec.MoveNext
               Loop
               mRec.Close
               If Len(mVaria) > 0 Then
                  mVaria = Mid(mVaria, 1, Len(mVaria) - 1)
                  Set mRec = mObj.oNovedadesFecha(Text1(0).Text, Text1(1).Text, " and codigo in (" & mVaria & ")", "", "order by mov1, fecha")
                  mCodRamal = Mid(mObj.sTablaDescr("ramales", " codigo=" & mRec!codramal, 2), 2, 2)
                  mSent = Left(mObj.sTablaDescr("sentidos", "codigo=" & mRec!sent, 1), 2)
                  Do While Not mRec.EOF
                     mKm = mRec!km
                     mKm = Replace(mKm, ",", ".")
                      
                     mData.Execute ("INSERT INTO Auxi (mFecha,mHora,Km,Sent,Codigo,Descripcion,Mov1,Ramal) VALUES ('" & Format(mRec.Fields(0), "dd/mm/yyyy") & "','" & Format(mRec.Fields(0), "hh:mm") & "', " _
                     & " " & mKm & ",'" & mSent & "','" & mRec!Codigo & "','" & mRec!descripcion & "','" & mRec!Mov1 & "','" & mCodRamal & "')")
                     mRec.MoveNext
                  Loop
                  mRec.Close
               End If
              
            Case "9" 'MySQL
               If Hora_ok(Text1(2).Text) Then
                  mTitulo = "REPORTE Seguimiento de Móviles"
                  mData.Execute ("CREATE TABLE Auxi (mFecha DATE,Movil TEXT,Codigo TEXT,Km DOUBLE,Sent TEXT,Pedido TEXT,Asign TEXT,Arribo TEXT,Free TEXT,Demora TEXT,Ocupado TEXT,TotDia INTEGER,TotMov INTEGER,TotDiaP INTEGER,TotMovP INTEGER,Rango TEXT,GralMov INTEGER, GralMovP INTEGER, Ramal TEXT,DemoraMinutos INTEGER, OcupadoMinutos INTEGER)")
                  TotMoviles = 0
                  TotMovilesP = 0
                  If Left(Combo1.Text, 3) = "TOD" Then
                     Select Case Combo1.ListIndex
                        Case 0
                           mSubTitulo = "Todos los Móviles. (" & Text1(0).Text & " al " & Text1(1).Text & ")"
                           For mI = 4 To Combo1.ListCount - 1
                              Combo1.ListIndex = mI
                              mMov = Right(Combo1.Text, 4)
                              sInsertDatos mMov, Text1(2).Text
                           Next
                        Case 1
                           mSubTitulo = "Todos los Móviles Patrulla. (" & Text1(0).Text & " al " & Text1(1).Text & ")"
                           For mI = 4 To Combo1.ListCount - 1
                              Combo1.ListIndex = mI
                              If Left(Combo1.Text, 3) = "MÓV" Then
                                 mMov = Right(Combo1.Text, 4)
                                 sInsertDatos mMov, Text1(2).Text
                              End If
                           Next
                        Case 2
                           mSubTitulo = "Todas las Grúas Livianas. (" & Text1(0).Text & " al " & Text1(1).Text & ")"
                           For mI = 4 To Combo1.ListCount - 1
                              If Left(Right(Combo1.List(mI), 4), 2) = "G0" Or Left(Right(Combo1.List(mI), 4), 2) = "GM" Or Left(Right(Combo1.List(mI), 4), 2) = "69" Then
                                 mMov = Right(Combo1.List(mI), 4)
                                 sInsertDatos mMov, Text1(2).Text
                              End If
                           Next
                        Case 3
                            mSubTitulo = "Todas las Grúas Pesadas. (" & Text1(0).Text & " al " & Text1(1).Text & ")"
                            Set mRec = mObj.oMovilTipo("GRP")
                            Do While Not mRec.EOF
                               sInsertDatos mRec!Codigo, Text1(2).Text
                               mRec.MoveNext
                            Loop
                            mRec.Close
                     End Select
                  Else
                     If List1.Visible And List1.ListCount > 0 Then
                        mSubTitulo = "Móvil: "
                        For mI = 0 To List1.ListCount - 1
                          List1.ListIndex = mI
                          mMov = Trim(List1.Text)
                          mSubTitulo = mSubTitulo & mMov & ", "
                          sInsertDatos mMov, Text1(2).Text
                        Next
                        mSubTitulo = Mid(mSubTitulo, 1, Len(mSubTitulo) - 1) & ". (" & Text1(0).Text & " al " & Text1(1).Text & ")"
                        List1.Clear
                     Else
                        mMov = Right(Combo1.Text, 4)
                        mSubTitulo = "Móvil: " & mMov & ". (" & Text1(0).Text & " al " & Text1(1).Text & ")"
                        sInsertDatos mMov, Text1(2).Text
                     End If
                  End If
                  'Set mAuxi = mData.OpenRecordset("SELECT count(*) as rtotMovilesP FROM Auxi WHERE Demora > #" & Format(Text1(2).Text & ":59", "hh:mm") & "#")
                  Set mAuxi = mData.OpenRecordset("SELECT count(*) as rtotMovilesP FROM Auxi WHERE Demora > '" & Format(CDate(Text1(2).Text), "hh:mm:ss") & "'")
                  TotMovilesP = mAuxi!rTotMovilesP
                  mAuxi.Close
                  Set mAuxi = mData.OpenRecordset("SELECT COUNT(*) as rTotMoviles FROM Auxi")
                  TotMoviles = mAuxi!rTotMoviles
                  mAuxi.Close
                  mData.Execute ("UPDATE Auxi SET GralMov = " & TotMoviles & ", GralMovP=" & TotMovilesP & "")
                  If Option1.Value Then
                     CrystalReport1.ReportFileName = App.Path & "\RegNovedades\" & "Rep25.rpt"
                  Else
                     CrystalReport1.ReportFileName = App.Path & "\RegNovedades\" & "Rep12.rpt"
                  End If
                  
               End If
                 
            Case "10" 'MySQL
               mTitulo = "REPORTE Seguimiento Por Evento"
               mSubTitulo = "Código: " & Right(Combo1.Text, 7)
               mData.Execute ("CREATE TABLE Auxi (mFecha TEXT,mHora TEXT,Km DOUBLE,Sent TEXT,Codigo TEXT,Descripcion TEXT,Mov1 TEXT,Mov2 TEXT,Mov3 TEXT, Ramal TEXT)")
               Set mRec = mObj.oNovedadesFecha(DateAdd("d", -2, Text1(0).Text), DateAdd("d", 2, Text1(1).Text), "", " and codigo='" & Right(Combo1.Text, 7) & "'", "order by fecha")
               Do While Not mRec.EOF
                  xFecha1 = Format(Left(mRec.Fields(0), 10), "dd/mm/yyyy")
                  mHora = Mid(mRec.Fields(0), 12, 12)
                  mCodRamal = Mid(mObj.sTablaDescr("ramales", " codigo=" & mRec!codramal, 2), 2, 2)
                  mSent = Left(mObj.sTablaDescr("sentidos", "codigo=" & mRec!sent, 1), 2)
                  mKm = mRec!km
                  mKm = Replace(mKm, ",", ".")
                  mData.Execute ("INSERT INTO Auxi (mFecha,mHora,Km,Sent,Codigo,Descripcion,Mov1,Mov2,Mov3,Ramal) VALUES " _
                     & "('" & xFecha1 & "','" & mHora & "'," & mKm & ",'" & mSent & "','" & mRec!Codigo & "','" & mRec!descripcion & "','" & mRec!Mov1 & "','" & mRec!Mov2 & "','" & mRec!Mov3 & "','" & mCodRamal & "')")
                  mRec.MoveNext
               Loop
               mRec.Close
               CrystalReport1.ReportFileName = App.Path & "\RegNovedades\" & "Rep13.rpt"
              
            Case "11"
               mTitulo = "REPORTE Servicios Realizados por Tipo o Móvil"
               mData.Execute ("CREATE TABLE Auxi (Codigo TEXT,Detalle TEXT,Total INTEGER,Movil TEXT)")
               mSubTitulo = "Todos los Móviles. (" & Text1(0).Text & " al " & Text1(1).Text & ")"
               CrystalReport1.ReportFileName = App.Path & "\RegNovedades\" & "Rep14.rpt"
               If Combo1.Text = "TODOS" Then
                  Set mRec = mObj.oNovedCountCodOrig(Text1(0).Text, Text1(1).Text, "CodServ1", " and codserv1 not in('','1','2','3')")
               Else
                  If Combo1.Text = "TOTAL DETALLADO" Then
                     Set mRec = mObj.oNovedCountCodOrig(Text1(0).Text, Text1(1).Text, "CodServ1,mov1", " and codserv1 not in('','1','2')")
                     CrystalReport1.ReportFileName = App.Path & "\RegNovedades\" & "Rep15.rpt"
                  Else
                     mMov = Right(Combo1.Text, 4)
                     mSubTitulo = "Móvil: " & mMov & " - " & Trim(Left(Combo1.Text, 25)) & ". (" & Text1(0).Text & " al " & Text1(1).Text & ")"
                     Set mRec = mObj.oNovedCountCodOrig(Text1(0).Text, Text1(1).Text, "CodServ1", " and codserv1 not in('','1','2') and mov1='" & mMov & "'")
                  End If
               End If
               If Not mRec.EOF Then
                  Do While Not mRec.EOF
                     mVaria = mObj.sTablaDescr("servicios", "Codigo = '" & mRec!codserv1 & "'", 1)
                     If mVaria <> "" Then
                        If Combo1.Text = "TOTAL DETALLADO" Then
                           mData.Execute ("INSERT INTO Auxi (Codigo,Detalle,Total,Movil) VALUES ('" & mRec!codserv1 & "','" & mVaria & "'," & mRec!Total & ",'" & mRec!Mov1 & "' )")
                        Else
                           mData.Execute ("INSERT INTO Auxi (Codigo,Detalle,Total) VALUES ('" & mRec!codserv1 & "','" & mVaria & "'," & mRec!Total & " )")
                        End If
                     Else
                        mData.Execute ("INSERT INTO Auxi (Codigo,Detalle,Total) VALUES ('" & mRec!codserv1 & "','Inexistente'," & mRec!Total & " )")
                     End If
                     mRec.MoveNext
                  Loop
               Else
                  MsgBox "No Existen Datos Para Esta Consulta", vbCritical, "RegNov 3.1 - Atención!!"
                  mFlag = False
               End If
               mRec.Close
                 
            Case "12" 'MySQL
               mTitulo = "REPORTE Servicios Realizados por Progresiva"
               mSubTitulo = "Fecha: " & Text1(0).Text & " al " & Text1(1).Text
               CrystalReport1.ReportFileName = App.Path & "\RegNovedades\" & "Rep16.rpt"
               mData.Execute ("CREATE TABLE Auxi (Km INTEGER, Ascend INTEGER,Descend INTEGER)")
               For mI = 12 To 65
                  mJ = mObj.iCountTabla("novedades", Text1(0).Text, Text1(1).Text, "fecha", "AND Km between " & mI & " AND " & (mI + 1) & " AND Sent in ('A','K','S') AND CodServ1 <> ''")
                  xMax = mObj.iCountTabla("novedades", Text1(0).Text, Text1(1).Text, "fecha", "AND Km between " & mI & " AND " & (mI + 1) & " AND Sent in ('D','B','T') AND CodServ1 <> ''")
                  mData.Execute ("INSERT INTO Auxi (Km,Ascend,Descend) VALUES (" & mI & "," & mJ & "," & xMax & " )")
               Next
              
            Case "13" 'MySQL
               mTitulo = "REPORTE Servicios Realizados por Hora"
               mSubTitulo = "Fecha: " & Text1(0).Text & " al " & Text1(1).Text
               mData.Execute ("CREATE TABLE Auxi (Flag INTEGER,Hora1 INTEGER, Ascend INTEGER,Descend INTEGER)")
               For mI = 0 To 23
                  mJ = mObj.iCountTabla("novedades", Text1(0).Text, Text1(1).Text, "fecha", "AND RIGHT(Fecha,8) BETWEEN '" & Format(mI, "00") & ":00:00' AND '" & Format(mI, "00") & ":59:59' AND Sent in ('A','K','S') AND CodServ1 <> ''")
                  xMax = mObj.iCountTabla("novedades", Text1(0).Text, Text1(1).Text, "fecha", "AND RIGHT(Fecha,8) BETWEEN '" & Format(mI, "00") & ":00:00' AND '" & Format(mI, "00") & ":59:59' AND Sent in ('D','B','T') AND CodServ1 <> ''")
                  mData.Execute ("INSERT INTO Auxi (Flag,Hora1,Ascend,Descend) VALUES (1," & mI & "," & mJ & "," & xMax & " )")
               Next
               CrystalReport1.ReportFileName = App.Path & "\RegNovedades\" & "Rep17.rpt"
              
            Case "14" 'MySQL
               mTitulo = "REPORTE Servicios Realizados por Sentido"
               mSubTitulo = "Fecha: " & Text1(0).Text & " al " & Text1(1).Text
               mData.Execute ("CREATE TABLE Auxi (Codigo TEXT, Detalle TEXT,Total INTEGER)")
               mData.Execute ("INSERT INTO Auxi (Codigo,Detalle,Total) VALUES ('A','Ascendente',0)")
               mData.Execute ("INSERT INTO Auxi (Codigo,Detalle,Total) VALUES ('B','Colectora Descendente',0)")
               mData.Execute ("INSERT INTO Auxi (Codigo,Detalle,Total) VALUES ('D','Descendente',0)")
               mData.Execute ("INSERT INTO Auxi (Codigo,Detalle,Total) VALUES ('K','Colectora Ascendente',0)")
               mData.Execute ("INSERT INTO Auxi (Codigo,Detalle,Total) VALUES ('S','Troncal Ascendente',0)")
               mData.Execute ("INSERT INTO Auxi (Codigo,Detalle,Total) VALUES ('T','Troncal Descendente',0)")
               Set mRec = mObj.oNovedCountCodOrig(Text1(0).Text, Text1(1).Text, "sent", " AND Sent <> '' AND CodServ1 <> '' ")
               Do While Not mRec.EOF
                  mData.Execute ("UPDATE Auxi SET Total = " & mRec!Total & " WHERE Codigo = '" & mRec!sent & "'")
                  mRec.MoveNext
               Loop
               mRec.Close
               CrystalReport1.ReportFileName = App.Path & "\RegNovedades\" & "Rep18.rpt"
                 
            Case "15"
               xFecha1 = Now
               mTitulo = "REPORTE Promedio Móvil de Servicios Realizados"
               xServGRU(0) = "AC"
               xServGRU(1) = "AM"
               xServGRU(2) = "AN"
               xServGRU(3) = "ARG"
               xServGRU(4) = "CC"
               xServGRU(5) = "CM"
               xServGRU(6) = "CN"
               xServGRU(7) = "COL"
               xServGRU(8) = "SAC"
               xServGRP(0) = "CM"
               xServGRP(1) = "CN"
               xServGRP(2) = "COL"
               xServGRP(3) = "CRG"
               xServGRP(4) = "SAC"
               xMaxTodo = 1
               mData.Execute ("CREATE TABLE Auxi (Movil TEXT,CodServ TEXT,Deriv INTEGER,Suspendidos INTEGER,Atendidos INTEGER,Ocupa INTEGER,Demora INTEGER)")
               If Combo1.Text = "TODOS" Then
                  Combo1.ListIndex = 1
                  xMaxTodo = Combo1.ListCount - 1
               End If
               For mI = 1 To xMaxTodo
                  If xMaxTodo = 1 Then
                     mMov = Right(Combo1.Text, 4)
                  Else
                     Combo1.ListIndex = mI
                     mMov = Right(Combo1.Text, 4)
                  End If
                  If Combo1.ListIndex < 8 Then
                     xMax = 8
                  Else
                     xMax = 4
                  End If
                  mSumCanc = 0
                  For mJ = 0 To xMax
                     mVaria = ""
                     If xMax = 8 Then
                        Set mRec = mObj.oNovedDistCampo(Text1(0).Text, Text1(1).Text, "codigo", " and codserv1='" & xServGRU(mJ) & "' AND Mov1 = '" & mMov & "'")
                     Else
                        Set mRec = mObj.oNovedDistCampo(Text1(0).Text, Text1(1).Text, "codigo", " and codserv1='" & xServGRP(mJ) & "' AND Mov1 = '" & mMov & "'")
                     End If
                     Do While Not mRec.EOF
                        mVaria = mVaria & "'" & mRec!Codigo & "',"
                        mRec.MoveNext
                     Loop
                     mRec.Close
                     If Len(mVaria) > 0 Then
                        mVaria = Mid(mVaria, 1, Len(mVaria) - 1)
                        Set mRec = mObj.oNovedadesFecha(Text1(0).Text, Text1(1).Text, "and codnov ='A' AND Codigo IN (" & mVaria & ")", " and (Mov1 = '" & mMov & "' OR Mov2 = '" & mMov & "' OR Mov3 = '" & mMov & "')", "order by fecha")
                        mSumServ = 0
                        mSumDem = 0
                        mSumOcup = 0
                        Do While Not mRec.EOF
                           xFlag = False
                           If mRec!Demora <> "" Then 'Demora
                              xFlag = True
                           End If
                           xDemora = 0
                           Set mRec1 = mObj.oNovedadesFecha(Text1(0).Text, Text1(1).Text, " and codnov='N'", " and mov1='" & mMov & "' AND Codigo='" & mRec!Codigo & "'", "order by fecha")
                           mArribo = False
                           If Not mRec1.EOF Then 'Demora por Arribo
                              xOcupa1 = mRec1!Fecha
                              If xFlag Then
                                 xDemora = DateDiff("n", mRec!Demora, mRec1!Fecha)
                              Else
                                 xDemora = DateDiff("n", mRec!Fecha, mRec1!Fecha)
                              End If
                              mArribo = True
                           End If
                           mRec1.Close
                           mJ = mObj.iCountTabla("novedades", Text1(0).Text, Text1(1).Text, "fecha", " and codserv1 in ('SS','SVR','USA') and codigo='" & mRec!Codigo & "'")
                           If mJ = 0 Then 'No fue cancelado
                              Set mRec1 = mObj.oNovedadesFecha(Text1(0).Text, Text1(1).Text, " and codnov='I'", "and codigo='" & mRec!Codigo & "' AND Mov1='" & mMov & "'", "order by fecha")
                              If Not mRec1.EOF Then
                                 If mArribo Then
                                    xOcupa = DateDiff("n", xOcupa1, mRec1!Fecha)
                                    mSumDem = mSumDem + xDemora
                                    mSumOcup = mSumOcup + xOcupa
                                    mSumServ = mSumServ + 1
                                 Else
                                    xOcupa = DateDiff("n", mRec!Fecha, mRec1!Fecha)
                                 End If
                              Else
                                 mSumCanc = mSumCanc + 1
                              End If
                              mRec1.Close
                           Else
                              mSumCanc = mSumCanc + 1
                           End If
                           mRec.MoveNext
                        Loop
                        mRec.Close
                        If mSumServ <> 0 Then
                           If xMax = 8 Then
                              mData.Execute ("INSERT INTO Auxi(Movil,CodServ,Deriv,Suspendidos,Atendidos,Ocupa,Demora) VALUES ('" & mMov & "','" & xServGRU(mJ) & "',0," & mSumCanc & "," & mSumServ & ",'" & mSumOcup & "','" & mSumDem & "')")
                           Else
                              mData.Execute ("INSERT INTO Auxi(Movil,CodServ,Deriv,Suspendidos,Atendidos,Ocupa,Demora) VALUES ('" & mMov & "','" & xServGRP(mJ) & "',0," & mSumCanc & "," & mSumServ & ",'" & mSumOcup & "','" & mSumDem & "')")
                           End If
                        End If
                     End If
                  Next
               Next
               Set mRec = mObj.oTabla("servicios", "")
               Do While Not mRec.EOF
                  mVaria = mRec!Codigo & " -" & mRec!descripcion
                  mData.Execute ("UPDATE Auxi SET CodServ = '" & mVaria & "' WHERE CodServ = '" & mRec!Codigo & "'")
                  mRec.MoveNext
               Loop
               mRec.Close
               If xMaxTodo = 1 Then
                  mSubTitulo = "Móvil: " & mMov & " - " & Trim(Left(Combo1.Text, 25)) & ". (" & Text1(0).Text & " al " & Text1(1).Text & ")"
               Else
                  mSubTitulo = "Todos los Móviles. (" & Text1(0).Text & " al " & Text1(1).Text & ")"
               End If
               CrystalReport1.ReportFileName = App.Path & "\RegNovedades\" & "Rep19.rpt"
               xFecha2 = Now
               MsgBox "Minutos de Demora: " & DateDiff("n", xFecha1, xFecha2)
                 
            Case "16" 'MySQL
               xFin = Now
               If Hora_ok(Text1(2).Text) And Hora_ok(Text1(3).Text) Then
                  xFecha1 = Text1(0).Text & " " & Text1(2).Text & ":00"
                  xFecha2 = Text1(1).Text & " " & Text1(3).Text & ":59"
                  If DateDiff("n", xFecha1, xFecha2) > 0 Then
                     If Trim(Text3.Text) <> "" Then
                        mVaria = " and descripcion like '%" & Trim(Text3.Text) & "%' "
                     End If
                     mTitulo = "REPORTE Novedades"
                     mData.Execute ("CREATE TABLE Auxi (mFecha DATE,mHora TEXT,Km DOUBLE,Sent TEXT,Codigo TEXT,Descripcion TEXT,Mov1 TEXT,Mov2 TEXT,Mov3 TEXT, Ramal TEXT)")
                     mSubTitulo = "Fecha: " & Text1(0).Text & " " & Text1(2).Text & " al " & Text1(1).Text & " " & Text1(3).Text & ""
                     If Check1.Value = 1 Then
                        Set mRec = mObj.oCodAlfaMinFecha(xFecha1, xFecha2)
                        Set mRec1 = mObj.oNovedadesFechaHora(xFecha1, xFecha2, "AND codigo='' or codigo is null", "", "order by fecha")
                        Do While Not mRec.EOF And Not mRec1.EOF
                           If mRec!Fecha < mRec1!Fecha Then
                              Set mRec2 = mObj.oNovedadesFecha(DateAdd("d", -1, xFecha1), DateAdd("d", 1, xFecha2), " and codigo='" & mRec!Codigo & "'", "", " order by fecha")
                              Do While Not mRec2.EOF
                                 xFecha1 = Format(Left(mRec2!Fecha, 10), "dd/mm/yyyy")
                                 mVaria = Mid(mRec2!Fecha, 12, 14)
                                 mCodRamal = Mid(mObj.sTablaDescr("ramales", " codigo=" & mRec2!codramal, 2), 2, 2)
                                 mSent = Left(mObj.sTablaDescr("sentidos", "codigo=" & mRec2!sent, 1), 2)
                                 mKm = mRec2!km
                                 mKm = Replace(mKm, ",", ".")
                                 mData.Execute ("INSERT INTO Auxi (mFecha,mHora,Km,Sent,Codigo,Descripcion,Mov1,Mov2,Mov3,Ramal) VALUES ('" & xFecha1 & "','" & mVaria & "'," & mKm & ",'" & mSent & "','" & mRec2!Codigo & "','" & mRec2!descripcion & "','" & mRec2!Mov1 & "','" & mRec2!Mov2 & "','" & mRec2!Mov3 & "','" & mCodRamal & "')")
                                 mRec2.MoveNext
                              Loop
                              mRec2.Close
                              mRec.MoveNext
                           Else
                              xFecha1 = Format(Left(mRec1!Fecha, 10), "dd/mm/yyyy")
                              mVaria = Mid(mRec1!Fecha, 12, 14)
                              mCodRamal = Mid(mObj.sTablaDescr("ramales", " codigo=" & mRec1!codramal, 2), 2, 2)
                              mSent = Left(mObj.sTablaDescr("sentidos", "codigo=" & mRec1!sent, 1), 2)
                              mKm = mRec1!km
                              mKm = Replace(mKm, ",", ".")
                              mData.Execute ("INSERT INTO Auxi (mFecha,mHora,Km,Sent,Codigo,Descripcion,Mov1,Mov2,Mov3,Ramal) VALUES ('" & xFecha1 & "','" & mVaria & "'," & mKm & ",'" & mSent & "','" & mRec1!Codigo & "','" & mRec1!descripcion & "','" & mRec1!Mov1 & "','" & mRec1!Mov2 & "','" & mRec1!Mov3 & "','" & mCodRamal & "')")
                              mRec1.MoveNext
                           End If
                        Loop
                        Do While Not mRec.EOF
                           Set mRec2 = mObj.oNovedadesFecha(DateAdd("d", -1, xFecha1), DateAdd("d", 1, xFecha2), " and codigo='" & mRec!Codigo & "'", "", " order by fecha")
                           Do While Not mRec2.EOF
                              xFecha1 = Format(Left(mRec2!Fecha, 10), "dd/mm/yyyy")
                              mVaria = Mid(mRec2!Fecha, 12, 14)
                              mCodRamal = Mid(mObj.sTablaDescr("ramales", " codigo=" & mRec2!codramal, 2), 2, 2)
                              mSent = Left(mObj.sTablaDescr("sentidos", "codigo=" & mRec2!sent, 1), 2)
                              mKm = mRec2!km
                              mKm = Replace(mKm, ",", ".")
                              
                              mData.Execute ("INSERT INTO Auxi (mFecha,mHora,Km,Sent,Codigo,Descripcion,Mov1,Mov2,Mov3,Ramal) VALUES ('" & xFecha1 & "','" & mVaria & "'," & mKm & ",'" & mSent & "','" & mRec2!Codigo & "','" & mRec2!descripcion & "','" & mRec2!Mov1 & "','" & mRec2!Mov2 & "','" & mRec2!Mov3 & "','" & mCodRamal & "')")
                              mRec2.MoveNext
                           Loop
                           mRec2.Close
                           mRec.MoveNext
                        Loop
                        mRec.Close
                        Do While Not mRec1.EOF
                           xFecha1 = Format(Left(mRec1.Fields(0), 10), "dd/mm/yyyy")
                           mVaria = Mid(mRec1.Fields(0), 12, 14)
                           mCodRamal = Mid(mObj.sTablaDescr("ramales", " codigo=" & mRec1!codramal, 2), 2, 2)
                           mSent = Left(mObj.sTablaDescr("sentidos", "codigo=" & mRec1!sent, 1), 2)
                           mKm = mRec1!km
                           mKm = Replace(mKm, ",", ".")
                           mData.Execute ("INSERT INTO Auxi (mFecha,mHora,Km,Sent,Codigo,Descripcion,Mov1,Mov2,Mov3,Ramal) VALUES ('" & xFecha1 & "','" & mVaria & "'," & mKm & ",'" & mSent & "','" & mRec1!Codigo & "','" & mRec1!descripcion & "','" & mRec1!Mov1 & "','" & mRec1!Mov2 & "','" & mRec1!Mov3 & "','" & mCodRamal & "')")
                           mRec1.MoveNext
                        Loop
                        mRec1.Close
                        CrystalReport1.ReportFileName = App.Path & "\RegNovedades\" & "Rep23S.rpt"
                     Else
                        Set mRec = mObj.oNovedadesFechaHora(xFecha1, xFecha2, "", mVaria, "order by fecha asc")
                        Do While Not mRec.EOF
                           xFecha1 = Format(Left(mRec!Fecha, 10), "dd/mm/yyyy")
                           mVaria = Mid(mRec!Fecha, 12, 14)
                           mCodRamal = Mid(mObj.sTablaDescr("ramales", " codigo=" & mRec!codramal, 2), 2, 2)
                           mSent = Left(mObj.sTablaDescr("sentidos", "codigo=" & mRec!sent, 1), 2)
                           mKm = mRec!km
                           mKm = Replace(mKm, ",", ".")
                           
                           
                           mData.Execute ("INSERT INTO Auxi (mFecha,mHora,Km,Sent,Codigo,Descripcion,Mov1,Mov2,Mov3,Ramal) VALUES ('" & xFecha1 & "','" & mVaria & "'," & mKm & ",'" & mSent & "','" & mRec!Codigo & "','" & mRec!descripcion & "','" & mRec!Mov1 & "','" & mRec!Mov2 & "','" & mRec!Mov3 & "','" & mCodRamal & "')")
                           mRec.MoveNext
                        Loop
                        mRec.Close
                        CrystalReport1.ReportFileName = App.Path & "\RegNovedades\" & "Rep23.rpt"
                     End If
                  Else
                     mFlag = False
                  End If
               End If
             
            Case "17" 'MySQL
               mVaria = ""
               mTitulo = "REPORTE Estadísticas Retiro de Objetos"
               Set mRec = mObj.oNovedadesFecha(Text1(0).Text, Text1(1).Text, "and codnov='E'", "and codigo<>''", "order by codigo")
               Do While Not mRec.EOF
                  mVaria = mVaria & "'" & mRec!Codigo & "',"
                  mRec.MoveNext
               Loop
               mRec.Close
               If Len(mVaria) > 0 Then
                  mVaria = Mid(mVaria, 1, Len(mVaria) - 1)
                  If Combo1.Text = "TODOS" Then
                     mSubTitulo = "Todos los Móviles.  (" & Text1(0).Text & " al " & Text1(1).Text & ")"
                     Set mRec = mObj.oNovedadesFecha(Text1(0).Text, Text1(1).Text, "", " and Codigo IN (" & mVaria & ")", "order by codigo, fecha")
                  Else
                     mMov = Right(Combo1.Text, 4)
                     mSubTitulo = "Móvil: " & mMov & ".  (" & Text1(0).Text & " al " & Text1(1).Text & ")"
                     Set mRec = mObj.oNovedadesFecha(Text1(0).Text, Text1(1).Text, "and Codigo IN (" & mVaria & ")", "and mov1='" & mMov & "'", "order by codigo, fecha")
                  End If
                  mData.Execute ("CREATE TABLE Auxi (mFecha DATE,Movil TEXT,Codigo TEXT,Km DOUBLE,Sent TEXT,Descrip TEXT,Aviso TEXT,Arribo TEXT,Liberado TEXT,Demora TEXT)")
                  Do While Not mRec.EOF
                     If mRec!CodNov = "A" Then
                        xCodServ = mRec!Codigo
                        xAsig = mRec!Fecha
                     End If
                     If mRec!CodNov = "R" Then
                        If mRec!Codigo = xCodServ Then
                           mHora = DateDiff("n", xAsig, mRec!Fecha)
                           mMovAux = ""
                           Set mRec1 = mObj.oTabla("tareas_datos", "where codalfa='" & mRec!Codigo & "' and codmovil='" & mRec!Mov1 & "' and fecha between '" & Format(DateAdd("d", -1, mRec!Fecha), "yyyy-mm-dd hh:mm:ss") & "' and '" & Format(DateAdd("d", 1, mRec!Fecha), "yyyy-mm-dd hh:mm:ss") & "'")
                           If Not mRec1.EOF Then
                              If mRec1!c1 = 1 Then mMovAux = "C1 "
                              If mRec1!c2 = 1 Then mMovAux = mMovAux & " C2 "
                              If mRec1!c3 = 1 Then mMovAux = mMovAux & " C3 "
                              If mRec1!c4 = 1 Then mMovAux = mMovAux & " C4 "
                              If mRec1!c5 = 1 Then mMovAux = mMovAux & " C5 "
                              If mRec1!c6 = 1 Then mMovAux = mMovAux & " C6 "
                              If mRec1!c7 = 1 Then mMovAux = mMovAux & " C7 "
                              If mRec1!C8 = 1 Then mMovAux = mMovAux & " C8 "
                              If mRec1!c9 = 1 Then mMovAux = mMovAux & " C9 "
                              If mRec1!bi = 1 Then mMovAux = mMovAux & " BI "
                              If mRec1!be = 1 Then mMovAux = mMovAux & " BE "
                              mMovAux = mMovAux & "- Cant: " & mRec1!Cant
                           End If
                           mRec1.Close
                           mKm = mRec!km
                           mKm = Replace(mKm, ",", ".")
                           mData.Execute ("INSERT INTO Auxi (mFecha,Movil,Codigo,Km,Sent,Descrip,Aviso,Arribo,Liberado,Demora) VALUES ('" & Format(mRec!Fecha, "dd/mm/yyyy") & "','" & mRec!Mov1 & "','" & mRec!Codigo & "'," & mKm & ",'" & mRec!sent & "','" & mRec!descripcion & " _ " & mMovAux & "','" & Format(xAsig, "hh:mm") & "','" & Format(mRec!Fecha, "hh:mm") & "','','" & mHora & "')")
                        End If
                     End If
                     If mRec!CodNov = "S" Then
                        If mRec!Codigo = xCodServ Then
                           mData.Execute ("UPDATE Auxi SET Liberado='" & Format(mRec!Fecha, "hh:mm") & "' WHERE Codigo='" & mRec!Codigo & "'")
                        End If
                     End If
                     mRec.MoveNext
                  Loop
                  mRec.Close
                  CrystalReport1.ReportFileName = App.Path & "\RegNovedades\" & "Rep29.rpt"
               End If
                  
            Case "18" 'MySQL
               mTitulo = "REPORTE Cuadro de las Principales Variables"
               mSubTitulo = "Período de Consulta: (" & Text1(0).Text & " al " & Text1(1).Text & ")"
               mData.Execute ("CREATE TABLE Auxi (C0 INTEGER,C1 INTEGER,C2 INTEGER,C3 INTEGER,C4 INTEGER,C5 INTEGER,C6 INTEGER,C7 INTEGER,C8 INTEGER,C9 INTEGER,C10 INTEGER,C11 INTEGER,C12 INTEGER,C13 INTEGER,C14 INTEGER,C15 INTEGER,C16 INTEGER,C17 INTEGER,C18 INTEGER)")
               mCuadro(0) = mObjRAcc.iCountIncidentes(Text1(0).Text, Text1(1).Text)
               Set mRec = mObjRAcc.oAccidTraza(Text1(0).Text, Text1(1).Text, True)
               mCuadro(1) = mRec!Total 'Accidentes Calzada Principal
               mRec.Close
               Set mRec = mObjRAcc.oAccidColect(Text1(0).Text, Text1(1).Text)
               mCuadro(2) = mRec!Total 'Accidentes Colectora
               mRec.Close
               Set mRec = mObjRAcc.oAccidTroncal(Text1(0).Text, Text1(1).Text)
               mCuadro(3) = mRec!Total 'Accidentes Troncal
               mCuadro(4) = mCuadro(1) + mCuadro(2) + mCuadro(3) 'Accidentes Total
               mRec.Close
               
               Set mObjRAcc = Nothing
               mCuadro(5) = mObj.iCountTabla("novedades", Text1(0).Text, Text1(1).Text, "fecha", " and codnov in ('A','MM','E') AND left(Mov1,1)='M' ")
               mCuadro(6) = mObj.iCountTabla("novedades", Text1(0).Text, Text1(1).Text, "fecha", " and codnov in ('A','MM') AND (left(Mov1,2)='G0' OR mov1='GMUL') ")
               mCuadro(6) = mCuadro(6) + mObj.iCountTabla("novedades", Text1(0).Text, Text1(1).Text, "fecha", " and codnov in ('A','MM') AND (left(Mov2,2)='G0' OR mov1='GMUL') ")
               mCuadro(6) = mCuadro(6) + mObj.iCountTabla("novedades", Text1(0).Text, Text1(1).Text, "fecha", " and codnov in ('A','MM') AND (left(Mov3,2)='G0' OR mov1='GMUL') ")
               mCuadro(7) = mObj.iCountTabla("novedades", Text1(0).Text, Text1(1).Text, "fecha", " and codnov in ('A','MM') AND (left(Mov1,2)='GP' OR left(Mov2,2)='GP' or left(Mov3,2)='GP') ")
               mCuadro(8) = mObj.iCountTabla("novedades", Text1(0).Text, Text1(1).Text, "fecha", " and codnov in ('M') AND (Mov1='BOMB' OR Mov2='BOMB' or Mov3='BOMB') ")
               mCuadro(9) = mObj.iCountTabla("novedades", Text1(0).Text, Text1(1).Text, "fecha", " and codnov in ('M') AND (Mov1='AMBU' OR Mov2='AMBU' or Mov3='AMBU') AND TipoNov IN ('O','IO','XO')")
               mCuadro(10) = mObj.iCountTabla("novedades", Text1(0).Text, Text1(1).Text, "fecha", " and codnov in ('M') AND Mov1='AMBU' AND TipoNov IN ('A','IA','XA')")
               mCuadro(11) = mObj.iCountTabla("novedades", Text1(0).Text, Text1(1).Text, "fecha", " and codnov in ('M') AND Mov1='AMBU' AND TipoNov IN ('P','IP','XP')")
               mCuadro(12) = mCuadro(9) + mCuadro(10) + mCuadro(11) 'Serv. Ambu Total
               If mCuadro(12) = 0 Then
                  mCuadro(12) = mObj.iCountTabla("novedades", Text1(0).Text, Text1(1).Text, "fecha", " and codnov in ('M') AND Mov1='AMBU'")
               End If
               mCuadro(13) = mObj.iCountTabla("novedades", Text1(0).Text, Text1(1).Text, "fecha", " and codnov in ('H')")
               mCuadro(14) = mObj.iCountTabla("novedades", Text1(0).Text, Text1(1).Text, "fecha", " and codnov in ('G')")
               mCuadro(15) = mObj.iCountTabla("novedades", Text1(0).Text, Text1(1).Text, "fecha", " and ((codnov in ('P') AND Mov1='GEND') OR descripcion='Operativo GEND')")
               mCuadro(16) = mObj.iCountTabla("novedades", Text1(0).Text, Text1(1).Text, "fecha", " and ((codnov in ('P') AND Mov1='POLI') OR descripcion='Operativo POLI')")
               Set mRec = mObj.oTabla("otros", "where CodTipOtro = 'OBJ'")
               mVaria = ""
               Do While Not mRec.EOF
                  mVaria = mVaria & "'" & mRec!Codigo & "',"
                  mRec.MoveNext
               Loop
               mRec.Close
               mVaria = Mid(mVaria, 1, Len(mVaria) - 1)
               mCuadro(17) = mObj.iCountTabla("novedades", Text1(0).Text, Text1(1).Text, "fecha", " and codnov in ('E','R') and (Left(Descripcion,3) IN (" & mVaria & "))")
               mVaria = ""
               Set mRec = mObj.oTabla("otros", "where CodTipOtro = 'ANI'")
               Do While Not mRec.EOF
                  mVaria = mVaria & "'" & mRec!Codigo & "',"
                  mRec.MoveNext
               Loop
               mRec.Close
               mVaria = Mid(mVaria, 1, Len(mVaria) - 1)
               mCuadro(18) = mObj.iCountTabla("novedades", Text1(0).Text, Text1(1).Text, "fecha", " and codnov in ('E','R') and Left(Descripcion,3) IN (" & mVaria & ")")
               mData.Execute ("INSERT INTO Auxi (C0,C1,C2,C3,C4,C5,C6,C7,C8,C9,C10,C11,C12,C13,C14,C15,C16,C17,C18) VALUES (" & mCuadro(0) & "," & mCuadro(1) & "," & mCuadro(2) & "," & mCuadro(3) & "," & mCuadro(4) & "," & mCuadro(5) & "," & mCuadro(6) & "," & mCuadro(7) & "," & mCuadro(8) & "," & mCuadro(9) & "," & mCuadro(10) & "," & mCuadro(11) & "," & mCuadro(12) & "," & mCuadro(13) & "," & mCuadro(14) & "," & mCuadro(15) & "," & mCuadro(16) & "," & mCuadro(17) & "," & mCuadro(18) & ")")
               CrystalReport1.ReportFileName = App.Path & "\RegNovedades\" & "Rep24.rpt"
                  
            Case "19" 'MySQL  -  Búsqueda de turnos por Código de Evento
               mTitulo = "REPORTE Turnos de Choferes por Código de Evento"
               mSubTitulo = "Evento: " & Trim(Right(Combo1.Text, 7)) & ""
               mData.Execute ("CREATE TABLE Auxi (Codigo TEXT,Fecha TEXT,Hora TEXT,Movil TEXT,Chofer TEXT,Turno TEXT,KmInicial INTEGER,KmFinal INTEGER)")
               mVaria = Trim(Right(Combo1.Text, 7))
               xFlag = True
               mMovAux = ""
               Set mRec = mObj.oNovedDistCampo(Left(Combo1.Text, 10), Left(Combo1.Text, 10), "mov1,fecha", "and Codigo='" & Trim(Right(Combo1.Text, 7)) & "' AND Left(Mov1,1) in ('G','M') ORDER BY Mov1")
               Do While Not mRec.EOF
                  mMov = mRec!Mov1
                  If mMov <> mMovAux Then
                     mMovAux = mMov
                     If Left(mRec!Mov1, 1) = "G" Then
                        Set mRec1 = mObj.oMovilTurnoFecha(mRec!Fecha, mRec!Mov1)
                        If Not mRec1.EOF Then
                           xFecha1 = Left(mRec1!fechainic, 10)
                           mHora = Right(mRec1!fechainic, 8)
                           If NVL(mRec1!KmFinal, 0) = "" Then
                              mKmF = 0
                           End If
                           mData.Execute ("INSERT INTO Auxi (Codigo,Fecha,Hora,Movil,Chofer,Turno,KmInicial,KmFinal) VALUES ('" & mVaria & "','" & xFecha1 & "','" & mHora & "','" & mRec!Mov1 & "','" & mRec1!chofer & "','" & mRec1!descripcion & "'," & mRec1!KmInicial & "," & mKmF & ")")
                        End If
                        mRec1.Close
                     Else
                        Set mRec1 = mObj.oMovilTurnoPatruFecha(mRec!Fecha, mRec!Mov1)
                        If Not mRec1.EOF Then
                           xFecha1 = Left(mRec1!fechainic, 10)
                           mHora = Right(mRec1!fechainic, 8)
                           If NVL(mRec1!KmFinal, 0) = "" Then
                              mKmF = 0
                           End If
                           mRut = mRec1!Nombre & " - " & mObj.sTablaDescr("patrulleros", "codigo='" & NVL(mRec1!codpatrullero2, "") & "'", 1) & " - " & mObj.sTablaDescr("patrulleros", "codigo='" & NVL(mRec1!codpatrullero3, "") & "'", 1)
                           mData.Execute ("INSERT INTO Auxi (Codigo,Fecha,Hora,Movil,Chofer,Turno,KmInicial,KmFinal) VALUES ('" & mVaria & "','" & xFecha1 & "','" & mHora & "','" & mRec!Mov1 & "','" & mRut & "','" & mRec1!descripcion & "'," & mRec1!KmInicial & "," & mKmF & ")")
                        End If
                        mRec1.Close
                     End If
                  End If
                  mRec.MoveNext
               Loop
               mRec.Close
               CrystalReport1.ReportFileName = App.Path & "\RegNovedades\" & "Rep27.rpt"
               
              Case "20" 'MySQL   - Búsqueda de Turnos por Fecha y Móvil
                  mTitulo = "REPORTE Turnos de Choferes por Código de Evento"
                  mData.Execute ("CREATE TABLE Auxi (Fecha TEXT,Hora TEXT,Movil TEXT,Chofer TEXT,Turno TEXT,KmInicial INTEGER,KmFinal INTEGER)")
                  If Combo1.Text = "TODOS" Then
                     mSubTitulo = "Todos los Móviles. (" & Text1(0).Text & " al " & Text1(1).Text & ")"
                     For mI = 1 To Combo1.ListCount - 1
                        mMov = Right(Combo1.List(mI), 4)
                        If Left(mMov, 1) = "M" Then
                           Set mRec = mObj.oMovilTurnoPatru(mMov & "' and fechainic between '" & Format(Text1(0).Text, "yyyy-mm-dd") & " 00:00:00' and '" & Format(Text1(1).Text, "yyyy-mm-dd") & " 23:59:59")
                           Do While Not mRec.EOF
                              xFecha1 = Left(mRec!fechainic, 10)
                              mHora = Right(mRec!fechainic, 8)
                              If NVL(mRec!KmFinal, 0) = "" Then
                                 xFin = "0"
                              Else
                                 xFin = NVL(mRec!KmFinal, 0)
                              End If
                              mRut = mRec!Nombre & " - " & mObj.sTablaDescr("patrulleros", "codigo='" & NVL(mRec!codpatrullero2, "") & "'", 1) & " - " & mObj.sTablaDescr("patrulleros", "codigo='" & NVL(mRec!codpatrullero3, "") & "'", 1)
                              mData.Execute ("INSERT INTO Auxi (Fecha,Hora,Movil,Chofer,Turno,KmInicial,KmFinal) VALUES ('" & xFecha1 & "','" & mHora & "','" & mMov & "','" & mRut & "','" & mRec!descripcion & "'," & mRec!KmInicial & "," & xFin & ")")
                              mRec.MoveNext
                           Loop
                        Else
                           Set mRec = mObj.oMovilTurnoChofer(mMov & "' and fechainic between '" & Format(Text1(0).Text, "yyyy-mm-dd") & " 00:00:00' and '" & Format(Text1(1).Text, "yyyy-mm-dd") & " 23:59:59")
                           Do While Not mRec.EOF
                              xFecha1 = Left(mRec!fechainic, 10)
                              mHora = Right(mRec!fechainic, 8)
                              If NVL(mRec!KmFinal, 0) = "" Then
                                 xFin = "0"
                              Else
                                 xFin = NVL(mRec!KmFinal, "0")
                              End If
                              mData.Execute ("INSERT INTO Auxi (Fecha,Hora,Movil,Chofer,Turno,KmInicial,KmFinal) VALUES ('" & xFecha1 & "','" & mHora & "','" & mMov & "','" & mRec!chofer & "','" & mRec!descripcion & "'," & mRec!KmInicial & "," & xFin & ")")
                              mRec.MoveNext
                           Loop
                        End If
                        mRec.Close
                     Next
                  Else
                     mMov = Right(Combo1.Text, 4)
                     mSubTitulo = "Móvil " & mMov & ". (" & Text1(0).Text & " al " & Text1(1).Text & ")"
                     Set mRec = mObj.oMovilTurnoPatru(mMov & "' and fechainic between '" & Format(Text1(0).Text, "yyyy-mm-dd") & " 00:00:00' and '" & Format(Text1(1).Text, "yyyy-mm-dd") & " 23:59:59")
                     Do While Not mRec.EOF
                        xFecha1 = Left(mRec!fechainic, 10)
                        mHora = Right(mRec!fechainic, 8)
                        mData.Execute ("INSERT INTO Auxi (Fecha,Hora,Movil,Chofer,Turno,KmInicial,KmFinal) VALUES ('" & xFecha1 & "','" & mHora & "','" & mMov & "','" & mRec!Nombre & "','" & mRec!descripcion & "'," & mRec!KmInicial & "," & NVL(mRec!KmFinal, 0) & ")")
                        mRec.MoveNext
                     Loop
                     mRec.Close
                  End If
                  CrystalReport1.ReportFileName = App.Path & "\RegNovedades\" & "Rep28.rpt"
             
              Case "21" '21-05-2007
                  mTitulo = "REPORTE de Operativos"
                  mSubTitulo = "Fecha del " & Text1(0).Text & " al " & Text1(1).Text
                  Set mRec = mObj.oListOperativos(Text1(0).Text, Text1(1).Text)
                  mData.Execute ("CREATE TABLE Auxi (Movil TEXT,Cantidad INTEGER)")
                  Do While Not mRec.EOF
                     mData.Execute ("INSERT INTO Auxi VALUES ('" & mRec!Movil & "'," & mRec!Total & ")")
                     mRec.MoveNext
                  Loop
                  mRec.Close
                  CrystalReport1.ReportFileName = App.Path & "\RegNovedades\" & "Rep30.rpt"
                  
              Case "22"
                  mTitulo = "REPORTE Total de Tareas"
                  mSubTitulo = "Fecha del " & Text1(0).Text & " al " & Text1(1).Text
                  Set mRec = mObj.oTotalTareas(Text1(0).Text, Text1(1).Text)
                  mData.Execute ("CREATE TABLE Auxi (codigo TEXT, descr TEXT, cantidad INTEGER)")
                  Do While Not mRec.EOF
                     mData.Execute ("INSERT INTO Auxi VALUES ('" & mRec!Codigo & "','" & mRec!descripcion & "'," & mRec!Total & ")")
                     mRec.MoveNext
                  Loop
                  mRec.Close
                  CrystalReport1.ReportFileName = App.Path & "\RegNovedades\" & "Rep31.rpt"
                  
                  
                  
                  
                  
                  
                  
                  
                  
                  
                  
                  
            Case "23" 'MySQL
               If Hora_ok(Text1(2).Text) Then
                  mTitulo = "REPORTE Seguimiento de Móviles"
                  mData.Execute ("CREATE TABLE Auxi (mFecha DATE,Movil TEXT,Codigo TEXT,Km DOUBLE,Sent TEXT,Pedido TEXT,Asign TEXT,Arribo TEXT,Free TEXT,Demora TEXT,Ocupado TEXT,TotDia INTEGER,TotMov INTEGER,TotDiaP INTEGER,TotMovP INTEGER,Rango TEXT,GralMov INTEGER, GralMovP INTEGER, Ramal TEXT, DemoraMinutos INTEGER, OcupadoMinutos INTEGER )")
                  TotMoviles = 0
                  TotMovilesP = 0
                  If Left(Combo1.Text, 3) = "TOD" Then
                     Select Case Combo1.ListIndex
                        Case 0
                           mSubTitulo = "Todos los Móviles. (" & Text1(0).Text & " al " & Text1(1).Text & ")"
                           For mI = 4 To Combo1.ListCount - 1
                              Combo1.ListIndex = mI
                              mMov = Right(Combo1.Text, 4)
                              sInsertDatos mMov, Text1(2).Text
                           Next
                        Case 1
                           mSubTitulo = "Todos los Móviles Patrulla. (" & Text1(0).Text & " al " & Text1(1).Text & ")"
                           For mI = 4 To Combo1.ListCount - 1
                              Combo1.ListIndex = mI
                              If Left(Combo1.Text, 3) = "MÓV" Then
                                 mMov = Right(Combo1.Text, 4)
                                 sInsertDatos mMov, Text1(2).Text
                              End If
                           Next
                        Case 2
                           mSubTitulo = "Todas las Grúas Livianas. (" & Text1(0).Text & " al " & Text1(1).Text & ")"
                           For mI = 4 To Combo1.ListCount - 1
                              If Left(Right(Combo1.List(mI), 4), 2) = "G0" Or Left(Right(Combo1.List(mI), 4), 2) = "GM" Or Left(Right(Combo1.List(mI), 4), 2) = "69" Then
                                 mMov = Right(Combo1.List(mI), 4)
                                 sInsertDatos mMov, Text1(2).Text
                              End If
                           Next
                        Case 3
                            mSubTitulo = "Todas las Grúas Pesadas. (" & Text1(0).Text & " al " & Text1(1).Text & ")"
                            Set mRec = mObj.oMovilTipo("GRP")
                            Do While Not mRec.EOF
                               sInsertDatos mRec!Codigo, Text1(2).Text
                               mRec.MoveNext
                            Loop
                            mRec.Close
                     End Select
                  Else
                     If List1.Visible And List1.ListCount > 0 Then
                        mSubTitulo = "Móvil: "
                        For mI = 0 To List1.ListCount - 1
                          List1.ListIndex = mI
                          mMov = Trim(List1.Text)
                          mSubTitulo = mSubTitulo & mMov & ", "
                          sInsertDatos mMov, Text1(2).Text
                        Next
                        mSubTitulo = Mid(mSubTitulo, 1, Len(mSubTitulo) - 1) & ". (" & Text1(0).Text & " al " & Text1(1).Text & ")"
                        List1.Clear
                     Else
                        mMov = Right(Combo1.Text, 4)
                        mSubTitulo = "Móvil: " & mMov & ". (" & Text1(0).Text & " al " & Text1(1).Text & ")"
                        sInsertDatos mMov, Text1(2).Text
                     End If
                  End If
                  'Set mAuxi = mData.OpenRecordset("SELECT count(*) as rtotMovilesP FROM Auxi WHERE Demora > #" & Format(Text1(2).Text & ":59", "hh:mm") & "#")
                  Set mAuxi = mData.OpenRecordset("SELECT count(*) as rtotMovilesP FROM Auxi WHERE Demora > '" & Format(CDate(Text1(2).Text), "hh:mm:ss") & "'")
                  TotMovilesP = mAuxi!rTotMovilesP
                  mAuxi.Close
                  Set mAuxi = mData.OpenRecordset("SELECT COUNT(*) as rTotMoviles FROM Auxi")
                  TotMoviles = mAuxi!rTotMoviles
                  mAuxi.Close
                  mData.Execute ("UPDATE Auxi SET GralMov = " & TotMoviles & ", GralMovP=" & TotMovilesP & "")
                  mData.Execute ("UPDATE Auxi SET GralMov = " & TotMoviles & ", GralMovP=" & TotMovilesP & "")
                  If Option1.Value Then
                     CrystalReport1.ReportFileName = App.Path & "\RegNovedades\" & "Rep25.rpt"
                  Else
                     CrystalReport1.ReportFileName = App.Path & "\RegNovedades\" & "Rep32.rpt"
                  End If
                  
               End If
                  
                  
                  
                  
                  
                  
                  
                  
                  
                  
                  
                  
                  
                  





             
         End Select
         sMsgEspere Me, "", False
      Else
         mFlag = False
      End If
      If mFlag Then
         On Error Resume Next
         Set mAuxi = mData.OpenRecordset("SELECT count(*) as Total FROM Auxi")
         'If Err.Description <> "" Then
            If mAuxi!Total > 0 Then
               CrystalReport1.WindowTitle = mTitulo
               CrystalReport1.DataFiles(0) = App.Path & "\RegNovedades\RegNovPlus.mdb"
               CrystalReport1.Formulas(0) = "Titulo= '" & mTitulo & "'"
               CrystalReport1.Formulas(1) = "Movil= '" & mSubTitulo & "'"
               CrystalReport1.Action = 1
               CrystalReport1.Formulas(1) = ""
               CrystalReport1.Formulas(0) = ""
            Else
               MsgBox "Consulta Sin Resultados", vbInformation, "RegNov 3.1 - Atención!!"
            End If
            mAuxi.Close
         'Else
         '  MsgBox "Consulta Sin Resultados", vbInformation, "RegNov 3.1 - Atención!!"
         'End If
         
         Combo1.ListIndex = -1
          Select Case mTabla
            Case "3", "6", "7", "12", "13", "14", "16", "18", "21", "22"
                 Combo1.ListIndex = 0
          End Select
          Label4.Visible = False
          Label5.Visible = False
          List1.Visible = False
          If mTabla <> 10 Then
             Text1(0).Text = ""
             Text1(1).Text = ""
             If mTabla <> 23 Then
               Text1(2).Text = ""
             End If
             Text1(3).Text = ""
             Text3.Text = ""
          End If
          If mTabla = "19" Then
             Combo1.ListIndex = 0
             Combo1.Visible = False
             Text2.Visible = True
             Text2.Text = ""
             Text1(0).Text = "14/11/2002 09:30:00"
             Text1(1).Text = "14/11/2003 09:31:00"
             Label2(2).Left = 3120
             Command1(0).Visible = False
             Command1(1).Visible = False
             Command2(0).Visible = True
             Command2(1).Visible = True
          End If
      End If
   Else
      Unload RNov7_frm
      ShowMenu 1, True, False
   End If
   Set mObjAcc = Nothing
   Set mObjRAcc = Nothing
   Set mRec1 = Nothing
   Set mRec2 = Nothing
End Sub

Private Sub Command2_Click(Index As Integer)
Dim mFlags As Boolean
Dim mI As Long
Dim mMes As Integer
Dim xMes As Integer
   
   mFlags = True
   If Index = 0 Then
      If Text2.Visible Then
         If Text2.Text <> "" And Len(Text2.Text) = 7 Then 'BÚSQUEDA POR CÓDIGO
            Set mRec = mObj.oNovedCodigoFecha("codigo='" & Text2.Text & "'", "", "")
            If Not mRec.EOF Then
               Combo1.Clear
               mMes = Format(mRec!mFecha, "m")
               xMes = 1000
               Do While Not mRec.EOF
                  If mMes <> xMes Then
                     Combo1.AddItem Format(mRec!mFecha, "dd/mm/yyyy") & " - " & mRec!Codigo
                  End If
                  mRec.MoveNext
                  If Not mRec.EOF Then
                     xMes = Format(mRec!mFecha, "m")
                  End If
               Loop
            Else
               MsgBox "No Existe el Código en la Base de Datos", vbCritical, "RegNov 3.1 - Atención!!"
               mFlags = False
            End If
            mRec.Close
         Else
            MsgBox "Faltan Agregar Datos", vbCritical, "RegNov 3.1 - Atención!!"
            mFlags = False
         End If
      Else
         If sValidFechaDesdeHasta(Text1(0).Text, Text1(1).Text) Then
            Set mRec = mObj.oNovedCodigoFecha("codigo<>''", Text1(0).Text, Text1(1).Text)
            If Not mRec.EOF Then
               Combo1.Clear
               Do While Not mRec.EOF
                  Combo1.AddItem Format(mRec!mFecha, "dd/mm/yyyy") & " - " & mRec!Codigo
                  mRec.MoveNext
               Loop
            Else
              MsgBox "No Existen Códigos en la Base de Datos para ese Rango de Fechas", vbCritical, "RegNov 3.1 - Atención!!"
            End If
            mRec.Close
         Else
            mFlags = False
         End If
      End If
   End If
   If mFlags Then
      Command1(0).Visible = True
      Command1(1).Visible = True
      Command2(0).Visible = False
      Command2(1).Visible = False
      Label2(2).Left = 2600
      Label2(2).Visible = True
      Combo1.Left = 2600
      Combo1.ListIndex = -1
      Combo1.Visible = True
      If Text1(0).Text = "" Then
         Text1(0).Text = "01/01/2002"
         Text1(1).Text = "01/01/2002"
      End If
      Text2.Text = ""
      Text2.Visible = False
      If mTabla = "10" Then
         Label2(0).Visible = False
         Label2(1).Visible = False
         Text1(0).Visible = False
         Text1(1).Visible = False
         Label3(0).Visible = True
         Label3(1).Visible = True
      Else
         If Index = 1 Then
            ShowMenu 1, True, False
            Unload RNov7_frm
         End If
      End If
   End If
End Sub

Private Sub Label3_Click(Index As Integer)
   If mTabla = "10" Then
      Command1(0).Visible = False
      Command1(1).Visible = False
      Command2(0).Visible = True
      Command2(1).Visible = True
      Combo1.Visible = False
      Label3(0).Visible = False
      Label3(1).Visible = False
      If Index = 0 Then
         Label2(0).Left = 2460
         Text1(0).Left = 2460
         Label2(1).Left = 3920
         Text1(1).Left = 3920
         Label2(0).Visible = True
         Label2(1).Visible = True
         Label2(2).Visible = False
         Text1(0).Visible = True
         Text1(1).Visible = True
         Text1(0).Text = ""
         Text1(1).Text = ""
         Text1(0).SetFocus
      Else
         Text2.Visible = True
         Text2.SetFocus
         Label2(2).Left = 3120
      End If
   End If
End Sub

Private Sub Label3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Label3(Index).BorderStyle = 1
End Sub

Private Sub Label3_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Label3(Index).BorderStyle = 0
End Sub

Private Sub Label5_Click()
Dim mCont As Integer
   For mCont = 0 To List1.ListCount - 1
      List1.ListIndex = mCont
      If List1.Text = Right(Combo1.Text, 4) Then
         mCont = 999
      End If
   Next
   mCont = mCont - 1
   If mCont <> 999 Then
      List1.AddItem Right(Combo1.Text, 4)
   End If
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Label5.BorderStyle = 1
End Sub

Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Label5.BorderStyle = 0
End Sub

Private Sub List1_DblClick()
   List1.RemoveItem List1.ListIndex
End Sub

Private Sub Option1_DblClick()
   Option1.Value = False
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   If Index < 2 Then
      KeyAscii = fDateKeyPress(Text1(Index), KeyAscii)
   Else
      KeyAscii = fHoraKeyPress(Text1(Index), KeyAscii)
   End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
  KeyAscii = fAlfaNumKeyPress(KeyAscii)
  KeyAscii = fUcaseKeyPress(KeyAscii)
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 37 Then
      KeyAscii = fAlfaNumSimbKeyPress(KeyAscii)
      KeyAscii = fUcaseKeyPress(KeyAscii)
   End If
End Sub

Private Sub sInsertDatos(ByVal mMov As String, RangDemo As String)
 Dim xRec As New ADODB.Recordset
 Dim xCod, xSent As String
 Dim Code As String
 Dim xFlag(4) As Boolean   '(0)Asignado - (1) Arribo - (2) Liberado - (3) Demora
 Dim Primero As Boolean
 Dim OkDia As Boolean
 Dim mOkGrabar As Boolean
 Dim mRecAux
 Dim CodNov(4) As Date '(0)demora, (1)Asignado, (2)Arribo, (3)Liberado
 Dim xFecha As Date
 Dim OldFecha As Date
 Dim xDemora As Date
 Dim xOcupado As Date
 Dim D24 As Date
 Dim DemoP As Date
 Dim mI As Integer
 Dim TotDia As Integer
 Dim TotMov As Integer
 Dim TotDiaP As Integer
 Dim TotMovP As Integer
 Dim xKm As Double
 Dim xRamal As String
 Dim xCodRamal As String
 Dim mKm As String
 Dim xDemoraMinutos As Integer
 Dim xOcupadoMinutos As Integer
   D24 = "12/12/2002 23:59:59"
   Set xRec = mObj.oServMoviles2(mMov, Text1(0).Text, Text1(1).Text)
   Primero = True
   OkDia = False
   mOkGrabar = False
   TotDia = 0
   TotMov = 0
   TotDiaP = 0
   TotMovP = 0
   DemoP = "20/06/2005 " & RangDemo & ":59"
   Do While Not xRec.EOF
      For mI = 0 To 3
         xFlag(mI) = False
      Next
      xCod = xRec!Codigo
      Code = xRec!Codigo
      mOkGrabar = False
      Do While Not xRec.EOF And xCod = Code
         Select Case xRec!CodNov
            Case "A", "MM"
               If Primero Then
                  xFecha = xRec!mFecha
                  Primero = False
               End If
               xFlag(0) = True
               CodNov(1) = xRec!Fecha 'Asignado
               'xSent = xRec!Sent
               xKm = xRec!km
               xCodRamal = Mid(mObj.sTablaDescr("ramales", " codigo=" & xRec!codramal, 2), 2, 2)
               xSent = Left(mObj.sTablaDescr("sentidos", "codigo=" & xRec!sent, 1), 2)
               If xFecha <> xRec!mFecha Then
                  OldFecha = xFecha
                  xFecha = xRec!mFecha
                  OkDia = True
               End If
            Case "Q"
               If Not xFlag(3) Then
                  CodNov(0) = xRec!Fecha 'Demora
                  xFlag(3) = True
               End If
            Case "N", "QQ", "RR"
               If Not xFlag(1) Then
                  CodNov(2) = xRec!Fecha ' Arribo
                  xFlag(1) = True
               End If
            Case "L", "TT"
               If Not xFlag(2) Then
                  CodNov(3) = xRec!Fecha
                  xFlag(2) = True
               End If
         End Select
         xRec.MoveNext
         If Not xRec.EOF Then
            Code = xRec!Codigo
         Else
            Code = "Urzagasti"
         End If
      Loop
      If Not xFlag(3) Then
        CodNov(0) = CodNov(1)
      End If
      If xFlag(0) And xFlag(1) And xFlag(2) Then
         If Format(CodNov(0), "dd/mm/yyyy") > Format(CodNov(2), "dd/mm/yyyy") Then
            xDemora = D24 - TimeValue(Format(CodNov(0), "hh:mm:ss"))
            xDemora = xDemora + TimeValue(Format(CodNov(2), "hh:mm:ss"))
         Else
            'xDemora = TimeValue(Format(CodNov(0), "hh:mm:ss")) - TimeValue(Format(CodNov(2), "hh:mm:ss"))'20161028
            xDemora = D24 - TimeValue(Format(CodNov(0), "hh:mm:ss"))
            xDemora = xDemora + TimeValue(Format(CodNov(2), "hh:mm:ss"))
         End If
         xDemora = Format(xDemora, "hh:mm:ss")
         xDemoraMinutos = segToMin(DateDiff("s", CodNov(0), CodNov(2)))
         If Format(CodNov(1), "dd/mm/yyyy") > Format(CodNov(3), "dd/mm/yyyy") Then
            xOcupado = D24 - TimeValue(Format(CodNov(1), "hh:mm:ss"))
            xOcupado = xDemora + TimeValue(Format(CodNov(3), "hh:mm:ss"))
         Else
 '           xOcupado = TimeValue(Format(CodNov(3), "hh:mm:ss")) - TimeValue(Format(CodNov(1), "hh:mm:ss"))'20161028
             xOcupado = D24 - TimeValue(Format(CodNov(1), "hh:mm:ss"))
             xOcupado = xOcupado + TimeValue(Format(CodNov(3), "hh:mm:ss"))
         End If
         xOcupadoMinutos = segToMin(DateDiff("s", CodNov(1), CodNov(3)))
         mKm = xKm
         mKm = Replace(mKm, ",", ".")
         'mData.Execute ("INSERT INTO Auxi (mFecha,Movil,Codigo,Km,Sent,Pedido,Asign,Arribo,Free,Demora,Ocupado,TotDia,TotMov,TotDiaP,TotMovP,Rango,GralMov,GralMovP,Ramal) VALUES ('" & xFecha & "','" & mMov & "','" & xCod & "'," & mKm & ",'" & xSent & "','" & Format(CodNov(0), "hh:mm:ss") & "','" & Format(CodNov(1), "hh:mm:ss") & "','" & Format(CodNov(2), "hh:mm:ss") & "','" & Format(CodNov(3), "hh:mm:ss") & "','" & Format(xDemora, "hh:mm:ss") & "','" & Format(xOcupado, "hh:mm:ss") & "',0,0,0,0,'" & RangDemo & "',0,0,'" & xCodRamal & "')")
         mData.Execute ("INSERT INTO Auxi (mFecha,Movil,Codigo,Km,Sent,Pedido,Asign,Arribo,Free,Demora,Ocupado,TotDia,TotMov,TotDiaP,TotMovP,Rango,GralMov,GralMovP,Ramal,DemoraMinutos,OcupadoMinutos) VALUES ('" & xFecha & "','" & mMov & "','" & xCod & "'," & mKm & ",'" & xSent & "','" & Format(CodNov(0), "hh:mm:ss") & "','" & Format(CodNov(1), "hh:mm:ss") & "','" & Format(CodNov(2), "hh:mm:ss") & "','" & Format(CodNov(3), "hh:mm:ss") & "','" & Format(xDemora, "hh:mm:ss") & "','" & Format(xOcupado, "hh:mm:ss") & "',0,0,0,0,'" & RangDemo & "',0,0,'" & xCodRamal & "'," & xDemoraMinutos & "," & xOcupadoMinutos & ")")
      Else
         '--------------Asignado, No Arribo, y liberado
         If xFlag(0) And Not xFlag(1) And xFlag(2) Then
            CodNov(2) = CodNov(3)
            If Format(CodNov(0), "dd/mm/yyyy") > Format(CodNov(2), "dd/mm/yyyy") Then
               xDemora = D24 - TimeValue(Format(CodNov(0), "hh:mm:ss"))
               xDemora = xDemora + TimeValue(Format(CodNov(2), "hh:mm:ss"))
            Else
               'xDemora = TimeValue(Format(CodNov(2), "hh:mm:ss")) - TimeValue(Format(CodNov(0), "hh:mm:ss")) '20161028
               xDemora = D24 - TimeValue(Format(CodNov(0), "hh:mm:ss"))
               xDemora = xDemora + TimeValue(Format(CodNov(2), "hh:mm:ss"))
            End If
            xDemoraMinutos = segToMin(DateDiff("s", CodNov(0), CodNov(2)))
            xOcupado = "00:00:00"
            xOcupadoMinutos = 0
            mKm = xKm
            mKm = Replace(mKm, ",", ".")
            'mData.Execute ("INSERT INTO Auxi (mFecha,Movil,Codigo,Km,Sent,Pedido,Asign,Arribo,Free,Demora,Ocupado,TotDia,TotMov,TotDiaP,TotMovP,Rango,GralMov,GralMovP,Ramal) VALUES ('" & xFecha & "','" & mMov & "','" & xCod & "'," & mKm & ",'" & xSent & "','" & Format(CodNov(0), "hh:mm:ss") & "','" & Format(CodNov(1), "hh:mm:ss") & "','" & Format(CodNov(2), "hh:mm:ss") & "','" & Format(CodNov(3), "hh:mm:ss") & "','" & Format(xDemora, "hh:mm:ss") & "','" & Format(xOcupado, "hh:mm:ss") & "',0,0,0,0,'" & RangDemo & "',0,0,'" & xCodRamal & "')")
            mData.Execute ("INSERT INTO Auxi (mFecha,Movil,Codigo,Km,Sent,Pedido,Asign,Arribo,Free,Demora,Ocupado,TotDia,TotMov,TotDiaP,TotMovP,Rango,GralMov,GralMovP,Ramal,DemoraMinutos,OcupadoMinutos) VALUES ('" & xFecha & "','" & mMov & "','" & xCod & "'," & mKm & ",'" & xSent & "','" & Format(CodNov(0), "hh:mm:ss") & "','" & Format(CodNov(1), "hh:mm:ss") & "','" & Format(CodNov(2), "hh:mm:ss") & "','" & Format(CodNov(3), "hh:mm:ss") & "','" & Format(xDemora, "hh:mm:ss") & "','" & Format(xOcupado, "hh:mm:ss") & "',0,0,0,0,'" & RangDemo & "',0,0,'" & xCodRamal & "'," & xDemoraMinutos & "," & xOcupadoMinutos & ")")
         Else
            If Not xFlag(0) And xFlag(2) Then 'no asignado y liberado
               If Not xFlag(1) Then
                  CodNov(2) = CodNov(3) 'si no tiene arribo le pongo el liberado
               End If
                  Dim xRec2 As New ADODB.Recordset
               Set xRec2 = mObj.oServMoviles4(xCod, mMov, Text1(0).Text, Text1(1).Text)
               If Not xRec2.EOF Then
                  CodNov(0) = xRec2!Fecha
                  CodNov(1) = xRec2!Fecha
                  xRec2.Close
                  Set xRec2 = Nothing
                  If Format(CodNov(0), "dd/mm/yyyy") > Format(CodNov(2), "dd/mm/yyyy") Then
                     xDemora = D24 - TimeValue(Format(CodNov(0), "hh:mm:ss"))
                     xDemora = xDemora + TimeValue(Format(CodNov(2), "hh:mm:ss"))
                  Else
                     xDemora = TimeValue(Format(CodNov(2), "hh:mm:ss")) - TimeValue(Format(CodNov(0), "hh:mm:ss"))
                  End If
                  xDemora = Format(xDemora, "hh:mm:ss")
                  xDemoraMinutos = segToMin(DateDiff("s", CodNov(0), CodNov(2)))
                  
                  If Format(CodNov(1), "dd/mm/yyyy") > Format(CodNov(3), "dd/mm/yyyy") Then
                     xOcupado = D24 - TimeValue(Format(CodNov(1), "hh:mm:ss"))
                     xOcupado = xDemora + TimeValue(Format(CodNov(3), "hh:mm:ss"))
                  Else
                     'xOcupado = TimeValue(Format(CodNov(3), "hh:mm:ss")) - TimeValue(Format(CodNov(1), "hh:mm:ss"))
                     xOcupado = D24 - TimeValue(Format(CodNov(1), "hh:mm:ss"))
                     xOcupado = xOcupado + TimeValue(Format(CodNov(3), "hh:mm:ss"))
                  End If
                  xOcupadoMinutos = segToMin(DateDiff("s", CodNov(1), CodNov(3)))
               Else
                  xDemora = "00:00:00"
                  xOcupado = "00:00:00"
                  xDemoraMinutos = 0
                  xOcupadoMinutos = 0
               End If
               mKm = xKm
               mKm = Replace(mKm, ",", ".")
               'mData.Execute ("INSERT INTO Auxi (mFecha,Movil,Codigo,Km,Sent,Pedido,Asign,Arribo,Free,Demora,Ocupado,TotDia,TotMov,TotDiaP,TotMovP,Rango,GralMov,GralMovP,Ramal) VALUES ('" & xFecha & "','" & mMov & "','" & xCod & "'," & mKm & ",'" & xSent & "','" & Format(CodNov(0), "hh:mm:ss") & "','" & Format(CodNov(1), "hh:mm:ss") & "','" & Format(CodNov(2), "hh:mm:ss") & "','" & Format(CodNov(3), "hh:mm:ss") & "','" & Format(xDemora, "hh:mm:ss") & "','" & Format(xOcupado, "hh:mm:ss") & "',0,0,0,0,'" & RangDemo & "',0,0,'" & xCodRamal & "')")
               mData.Execute ("INSERT INTO Auxi (mFecha,Movil,Codigo,Km,Sent,Pedido,Asign,Arribo,Free,Demora,Ocupado,TotDia,TotMov,TotDiaP,TotMovP,Rango,GralMov,GralMovP,Ramal,DemoraMinutos,OcupadoMinutos) VALUES ('" & xFecha & "','" & mMov & "','" & xCod & "'," & mKm & ",'" & xSent & "','" & Format(CodNov(0), "hh:mm:ss") & "','" & Format(CodNov(1), "hh:mm:ss") & "','" & Format(CodNov(2), "hh:mm:ss") & "','" & Format(CodNov(3), "hh:mm:ss") & "','" & Format(xDemora, "hh:mm:ss") & "','" & Format(xOcupado, "hh:mm:ss") & "',0,0,0,0,'" & RangDemo & "',0,0,'" & xCodRamal & "'," & xDemoraMinutos & "," & xOcupadoMinutos & " )")
            End If
         End If
      End If
      For mI = 0 To 3
         CodNov(mI) = "00:00:00"
      Next
      If OkDia Then
         'Set mRecAux = mData.OpenRecordset("SELECT count(*) as rTotDiaP FROM Auxi WHERE Demora > #" & Format(DemoP, "hh:mm") & "# AND Movil = '" & mMov & "' AND mFecha = #" & Format(OldFecha, "mm/dd/yyyy") & "#")
         Set mRecAux = mData.OpenRecordset("SELECT count(*) as rTotDiaP FROM Auxi WHERE Demora >'" & Format(CDate(DemoP), "hh:mm:ss") & "' AND Movil = '" & mMov & "' AND mFecha = #" & Format(OldFecha, "mm/dd/yyyy") & "#")
         TotDiaP = mRecAux!rTotDiaP
         mRecAux.Close
         Set mRecAux = mData.OpenRecordset("SELECT count(*) as rTotDia FROM Auxi WHERE Movil = '" & mMov & "' AND mFecha = #" & Format(OldFecha, "mm/dd/yyyy") & "#")
         TotDia = mRecAux!rTotDia
         mRecAux.Close
         mData.Execute ("UPDATE Auxi SET TotDia=" & TotDia & ",TotDiaP=" & TotDiaP & " WHERE mFecha = #" & Format(OldFecha, "mm/dd/yyyy") & "# AND Movil = '" & mMov & "'")
         OkDia = False
      End If
   Loop
   xRec.Close
   OldFecha = xFecha
   'Set mRecAux = mData.OpenRecordset("SELECT count(*) as rTotDiaP FROM Auxi WHERE Demora > #" & Format(DemoP, "hh:mm") & "# AND Movil = '" & mMov & "' AND mFecha = #" & Format(OldFecha, "mm/dd/yyyy") & "#")
   Set mRecAux = mData.OpenRecordset("SELECT count(*) as rTotDiaP FROM Auxi WHERE Demora >'" & Format(CDate(DemoP), "hh:mm:ss") & "' AND Movil = '" & mMov & "' AND mFecha = #" & Format(OldFecha, "mm/dd/yyyy") & "#")
   TotDiaP = mRecAux!rTotDiaP
   mRecAux.Close
   Set mRecAux = mData.OpenRecordset("SELECT count(*) as rTotDia FROM Auxi WHERE Movil = '" & mMov & "' AND mFecha = #" & Format(OldFecha, "mm/dd/yyyy") & "#")
   TotDia = mRecAux!rTotDia
   mRecAux.Close
   mData.Execute ("UPDATE Auxi SET TotDia=" & TotDia & ",TotDiaP=" & TotDiaP & " WHERE mFecha = #" & Format(OldFecha, "mm/dd/yyyy") & "# AND Movil = '" & mMov & "'")
   OkDia = False
   'Set mRecAux = mData.OpenRecordset("SELECT count(*) as rTotMovP FROM Auxi WHERE Demora > #" & Format(DemoP, "hh:mm") & "# AND Movil = '" & mMov & "'")
   Set mRecAux = mData.OpenRecordset("SELECT count(*) as rTotMovP FROM Auxi WHERE Demora > '" & Format(DemoP, "hh:mm:ss") & "' AND Movil = '" & mMov & "'")
   TotMovP = mRecAux!rTotMovP
   mRecAux.Close
   Set mRecAux = mData.OpenRecordset("SELECT count(*) as rTotMov FROM Auxi WHERE Movil='" & mMov & "'")
   TotMov = mRecAux!rTotMov
   mRecAux.Close
   mData.Execute ("UPDATE Auxi SET TotMov=" & TotMov & ",TotMovP=" & TotMovP & " WHERE Movil = '" & mMov & "'")
   Set xRec = Nothing
End Sub
Function segToMin(ByVal min As Long) As Long
   Dim cociente, parteDecimal As Long
   cociente = min / 60
   parteDecimal = cociente - Fix(cociente)
   cociente = cociente - parteDecimal
   If parteDecimal > (1 / 2) Then
      cociente = cociente + 1
   End If
   segToMin = cociente
End Function

'Private Sub sInsertDatos(ByVal mMov As String, RangDemo As String)
' Dim xRec As New ADODB.Recordset
' Dim xCod, xSent As String
' Dim Code As String
' Dim xFlag(4) As Boolean   '(0)Asignado - (1) Arribo - (2) Liberado - (3) Demora
' Dim Primero As Boolean
' Dim OkDia As Boolean
' Dim mOkGrabar As Boolean
' Dim mRecAux
' Dim CodNov(4) As Date '(0)demora, (1)Asignado, (2)Arribo, (3)Liberado
' Dim xFecha As Date
' Dim OldFecha As Date
' Dim xDemora As Date
' Dim xOcupado As Date
' Dim D24 As Date
' Dim DemoP As Date
' Dim mI As Integer
' Dim TotDia As Integer
' Dim TotMov As Integer
' Dim TotDiaP As Integer
' Dim TotMovP As Integer
' Dim xKm As Double
' Dim xRamal As String
' Dim xCodRamal As String
' Dim mKm As String
' Dim xDemoraMinutos As Integer
' Dim xOcupadoMinutos As Integer
' Dim xDemoraStr As String
'
'
'
'
'   D24 = "12/12/2002 23:59:59"
'   Set xRec = mObj.oServMoviles2(mMov, Text1(0).Text, Text1(1).Text)
'   Primero = True
'   OkDia = False
'   mOkGrabar = False
'   TotDia = 0
'   TotMov = 0
'   TotDiaP = 0
'   TotMovP = 0
'   DemoP = "20/06/2005 " & RangDemo & ":59"
'   Do While Not xRec.EOF
'   xDemoraMinutos = 0
'   xOcupadoMinutos = 0
'      For mI = 0 To 3
'         xFlag(mI) = False
'      Next
'      xCod = xRec!Codigo
'      Code = xRec!Codigo
'      mOkGrabar = False
'      Do While Not xRec.EOF And xCod = Code
'         Select Case xRec!CodNov
'            Case "A", "MM" ', "E"
'               If Primero Then
'                  xFecha = xRec!mFecha
'                  Primero = False
'               End If
'               xFlag(0) = True
'               CodNov(1) = xRec!Fecha 'Asignado
'               'xSent = xRec!Sent
'               xKm = xRec!Km
'               xCodRamal = Mid(mObj.sTablaDescr("ramales", " codigo=" & xRec!codramal, 2), 2, 2)
'               xSent = Left(mObj.sTablaDescr("sentidos", "codigo=" & xRec!Sent, 1), 2)
'
'               If xFecha <> xRec!mFecha Then
'                  OldFecha = xFecha
'                  xFecha = xRec!mFecha
'                  OkDia = True
'               End If
'
'            Case "Q"
'               If Not xFlag(3) Then
'                  CodNov(0) = xRec!Fecha 'Demora
'                  xFlag(3) = True
'               End If
'
'            Case "N", "QQ", "RR"
'               If Not xFlag(1) Then
'                  CodNov(2) = xRec!Fecha ' Arribo
'                  xFlag(1) = True
'               End If
'
'            Case "L", "TT" ', "S"
'               If Not xFlag(2) Then
'                  CodNov(3) = xRec!Fecha
'                  xFlag(2) = True
'               End If
'         End Select
'         xRec.MoveNext
'
'         If Not xRec.EOF Then
'            Code = xRec!Codigo
'         Else
'            Code = "Urzagasti"
'         End If
'      Loop
'      If Not xFlag(3) Then
'        CodNov(0) = CodNov(1)
'      End If
'      If xFlag(0) And xFlag(1) And xFlag(2) Then
'         If Format(CodNov(0), "dd/mm/yyyy") > Format(CodNov(2), "dd/mm/yyyy") Then
'            xDemora = D24 - TimeValue(Format(CodNov(0), "hh:mm:ss"))
'            xDemora = xDemora + TimeValue(Format(CodNov(2), "hh:mm:ss"))
'         Else
'            xDemora = TimeValue(Format(CodNov(2), "hh:mm:ss")) - TimeValue(Format(CodNov(0), "hh:mm:ss"))
'         End If
'         xDemora = Format(xDemora, "hh:mm:ss")
'
'         'xDemoraMinutos = cantidadMinutos(Hour(xDemora), Minute(xDemora), Second(xDemora))
'
'
'         xDemoraMinutos = DateDiff("n", CodNov(0), CodNov(2))
'
'
'         If Format(CodNov(1), "dd/mm/yyyy") > Format(CodNov(3), "dd/mm/yyyy") Then
'            xOcupado = D24 - TimeValue(Format(CodNov(1), "hh:mm:ss"))
'            xOcupado = xDemora + TimeValue(Format(CodNov(3), "hh:mm:ss"))
'         Else
'            xOcupado = TimeValue(Format(CodNov(3), "hh:mm:ss")) - TimeValue(Format(CodNov(1), "hh:mm:ss"))
'         End If
'         'xOcupadoMinutos = cantidadMinutos(Hour(xOcupado), Minute(xOcupado), Second(xOcupado))
'         xOcupadoMinutos = DateDiff("n", CodNov(1), CodNov(3))
'         mKm = xKm
'         mKm = Replace(mKm, ",", ".")
'         mData.Execute ("INSERT INTO Auxi (mFecha,Movil,Codigo,Km,Sent,Pedido,Asign,Arribo,Free,Demora,Ocupado,TotDia,TotMov,TotDiaP,TotMovP,Rango,GralMov,GralMovP,Ramal,DemoraMinutos,OcupadoMinutos) VALUES ('" & xFecha & "','" & mMov & "','" & xCod & "'," & mKm & ",'" & xSent & "','" & Format(CodNov(0), "hh:mm:ss") & "','" & Format(CodNov(1), "hh:mm:ss") & "','" & Format(CodNov(2), "hh:mm:ss") & "','" & Format(CodNov(3), "hh:mm:ss") & "','" & Format(xDemora, "hh:mm:ss") & "','" & Format(xOcupado, "hh:mm:ss") & "',0,0,0,0,'" & RangDemo & "',0,0,'" & xCodRamal & "'," & xDemoraMinutos & "," & 0 & ")")
'         'mData.Execute ("INSERT INTO Auxi (mFecha,Movil,Codigo,Km,Sent,Pedido,Asign,Arribo,Free,Demora,Ocupado,TotDia,TotMov,TotDiaP,TotMovP,Rango,GralMov,GralMovP,Ramal) VALUES ('" & xFecha & "','" & mMov & "','" & xCod & "'," & mKm & ",'" & xSent & "','" & Format(CodNov(0), "hh:mm:ss") & "','" & Format(CodNov(1), "hh:mm:ss") & "','" & Format(CodNov(2), "hh:mm:ss") & "','" & Format(CodNov(3), "hh:mm:ss") & "','" & xDemoraStr & "','" & Format(xOcupado, "hh:mm:ss") & "',0,0,0,0,'" & RangDemo & "',0,0,'" & xCodRamal & "')")
'      Else
'         If xFlag(0) And Not xFlag(1) And xFlag(2) Then
'            CodNov(2) = CodNov(3)
'            If Format(CodNov(0), "dd/mm/yyyy") > Format(CodNov(2), "dd/mm/yyyy") Then
'               xDemora = D24 - TimeValue(Format(CodNov(0), "hh:mm:ss"))
'               xDemora = xDemora + TimeValue(Format(CodNov(2), "hh:mm:ss"))
'            Else
'               xDemora = TimeValue(Format(CodNov(2), "hh:mm:ss")) - TimeValue(Format(CodNov(0), "hh:mm:ss"))
'            End If
'            'xDemoraStr = DateDiff("n", CodNov(0), CodNov(2))
'            'xDemoraMinutos = cantidadMinutos(Hour(xDemora), Minute(xDemora), Second(xDemora))
'            xDemoraMinutos = DateDiff("n", CodNov(0), CodNov(2))
'            xOcupado = "00:00:00"
'
'            mKm = xKm
'            mKm = Replace(mKm, ",", ".")
'            mData.Execute ("INSERT INTO Auxi (mFecha,Movil,Codigo,Km,Sent,Pedido,Asign,Arribo,Free,Demora,Ocupado,TotDia,TotMov,TotDiaP,TotMovP,Rango,GralMov,GralMovP,Ramal,DemoraMinutos,OcupadoMinutos) VALUES ('" & xFecha & "','" & mMov & "','" & xCod & "'," & mKm & ",'" & xSent & "','" & Format(CodNov(0), "hh:mm:ss") & "','" & Format(CodNov(1), "hh:mm:ss") & "','" & Format(CodNov(2), "hh:mm:ss") & "','" & Format(CodNov(3), "hh:mm:ss") & "','" & Format(xDemora, "hh:mm:ss") & "','" & Format(xOcupado, "hh:mm:ss") & "',0,0,0,0,'" & RangDemo & "',0,0,'" & xCodRamal & "'," & xDemoraMinutos & "," & 1 & ")")
'            'mData.Execute ("INSERT INTO Auxi (mFecha,Movil,Codigo,Km,Sent,Pedido,Asign,Arribo,Free,Demora,Ocupado,TotDia,TotMov,TotDiaP,TotMovP,Rango,GralMov,GralMovP,Ramal) VALUES ('" & xFecha & "','" & mMov & "','" & xCod & "'," & mKm & ",'" & xSent & "','" & Format(CodNov(0), "hh:mm:ss") & "','" & Format(CodNov(1), "hh:mm:ss") & "','" & Format(CodNov(2), "hh:mm:ss") & "','" & Format(CodNov(3), "hh:mm:ss") & "','" & Format(xDemora, "hh:mm:ss") & "','" & Format(xOcupado, "hh:mm:ss") & "',0,0,0,0,'" & RangDemo & "',0,0,'" & xCodRamal & "')")
'            'mData.Execute ("INSERT INTO Auxi (mFecha,Movil,Codigo,Km,Sent,Pedido,Asign,Arribo,Free,Demora,Ocupado,TotDia,TotMov,TotDiaP,TotMovP,Rango,GralMov,GralMovP,Ramal) VALUES ('" & xFecha & "','" & mMov & "','" & xCod & "'," & mKm & ",'" & xSent & "','" & Format(CodNov(0), "hh:mm:ss") & "','" & Format(CodNov(1), "hh:mm:ss") & "','" & Format(CodNov(2), "hh:mm:ss") & "','" & Format(CodNov(3), "hh:mm:ss") & "','" & xDemoraStr & "','" & Format(xOcupado, "hh:mm:ss") & "',0,0,0,0,'" & RangDemo & "',0,0,'" & xCodRamal & "')")
'         Else
'            If Not xFlag(0) And xFlag(2) Then
'               Dim xRec2 As New ADODB.Recordset
'               Set xRec2 = mObj.oServMoviles4(xCod, mMov, Text1(0).Text, Text1(1).Text)
'               CodNov(1) = xRec!Fecha
'               If Not xFlag(1) Then
'                  CodNov(2) = CodNov(1)
'               Else
'                  CodNov(2) = xRec!Fecha
'               End If
'               xRec2.Close
'               Set xRec2 = Nothing
'               If Format(CodNov(0), "dd/mm/yyyy") > Format(CodNov(2), "dd/mm/yyyy") Then
'                  xDemora = D24 - TimeValue(Format(CodNov(0), "hh:mm:ss"))
'                  xDemora = xDemora + TimeValue(Format(CodNov(2), "hh:mm:ss"))
'               Else
'                  xDemora = TimeValue(Format(CodNov(2), "hh:mm:ss")) - TimeValue(Format(CodNov(0), "hh:mm:ss"))
'               End If
'
'               mKm = xKm
'               mKm = Replace(mKm, ",", ".")
'
'               mData.Execute ("INSERT INTO Auxi (mFecha,Movil,Codigo,Km,Sent,Pedido,Asign,Arribo,Free,Demora,Ocupado,TotDia,TotMov,TotDiaP,TotMovP,Rango,GralMov,GralMovP,Ramal,DemoraMinutos,OcupadoMinutos) VALUES ('" & xFecha & "','" & mMov & "','" & xCod & "'," & mKm & ",'" & xSent & "','" & Format(CodNov(0), "hh:mm:ss") & "','" & Format(CodNov(1), "hh:mm:ss") & "','" & Format(CodNov(2), "hh:mm:ss") & "','" & Format(CodNov(3), "hh:mm:ss") & "','" & Format(xDemora, "hh:mm:ss") & "','" & Format(xOcupado, "hh:mm:ss") & "',0,0,0,0,'" & RangDemo & "',0,0,'" & xCodRamal & "'," & xDemoraMinutos & "," & 3 & ")")
'            End If
'
'
'
'            mKm = xKm
'            mKm = Replace(mKm, ",", ".")
'            mData.Execute ("INSERT INTO Auxi (mFecha,Movil,Codigo,Km,Sent,Pedido,Asign,Arribo,Free,Demora,Ocupado,TotDia,TotMov,TotDiaP,TotMovP,Rango,GralMov,GralMovP,Ramal,DemoraMinutos,OcupadoMinutos) VALUES ('" & xFecha & "','" & mMov & "','" & xCod & "'," & mKm & ",'" & xSent & "','" & Format(CodNov(0), "hh:mm:ss") & "','" & Format(CodNov(1), "hh:mm:ss") & "','" & Format(CodNov(2), "hh:mm:ss") & "','" & Format(CodNov(3), "hh:mm:ss") & "','" & Format(xDemora, "hh:mm:ss") & "','" & Format(xOcupado, "hh:mm:ss") & "',0,0,0,0,'" & RangDemo & "',0,0,'" & xCodRamal & "'," & xDemoraMinutos & "," & 2 & ")")
'            'mData.Execute ("INSERT INTO Auxi (mFecha,Movil,Codigo,Km,Sent,Pedido,Asign,Arribo,Free,Demora,Ocupado,TotDia,TotMov,TotDiaP,TotMovP,Rango,GralMov,GralMovP,Ramal) VALUES ('" & xFecha & "','" & mMov & "','" & xCod & "'," & mKm & ",'" & xSent & "','" & Format(CodNov(0), "hh:mm:ss") & "','" & Format(CodNov(1), "hh:mm:ss") & "','" & Format(CodNov(2), "hh:mm:ss") & "','" & Format(CodNov(3), "hh:mm:ss") & "','" & Format(xDemora, "hh:mm:ss") & "','" & Format(xOcupado, "hh:mm:ss") & "',0,0,0,0,'" & RangDemo & "',0,0,'" & xCodRamal & "')")
'            'mData.Execute ("INSERT INTO Auxi (mFecha,Movil,Codigo,Km,Sent,Pedido,Asign,Arribo,Free,Demora,Ocupado,TotDia,TotMov,TotDiaP,TotMovP,Rango,GralMov,GralMovP,Ramal) VALUES ('" & xFecha & "','" & mMov & "','" & xCod & "'," & mKm & ",'" & xSent & "','" & Format(CodNov(0), "hh:mm:ss") & "','" & Format(CodNov(1), "hh:mm:ss") & "','" & Format(CodNov(2), "hh:mm:ss") & "','" & Format(CodNov(3), "hh:mm:ss") & "','" & xDemoraStr & "','" & Format(xOcupado, "hh:mm:ss") & "',0,0,0,0,'" & RangDemo & "',0,0,'" & xCodRamal & "')")
'         End If
'      End If
'      For mI = 0 To 3
'         CodNov(mI) = "00:00:00"
'      Next
'      If OkDia Then
'         'Set mRecAux = mData.OpenRecordset("SELECT count(*) as rTotDiaP FROM Auxi WHERE Demora > #" & Format(DemoP, "hh:mm") & "# AND Movil = '" & mMov & "' AND mFecha = #" & Format(OldFecha, "mm/dd/yyyy") & "#")
'         Set mRecAux = mData.OpenRecordset("SELECT count(*) as rTotDiaP FROM Auxi WHERE Demora > '" & Format(CDate(DemoP), "hh:mm:ss") & "' AND Movil = '" & mMov & "' AND mFecha = #" & Format(OldFecha, "mm/dd/yyyy") & "#")
'         TotDiaP = mRecAux!rTotDiaP
'         mRecAux.Close
'         Set mRecAux = mData.OpenRecordset("SELECT count(*) as rTotDia FROM Auxi WHERE Movil = '" & mMov & "' AND mFecha = #" & Format(OldFecha, "mm/dd/yyyy") & "#")
'         TotDia = mRecAux!rTotDia
'         mRecAux.Close
'         mData.Execute ("UPDATE Auxi SET TotDia=" & TotDia & ",TotDiaP=" & TotDiaP & " WHERE mFecha = #" & Format(OldFecha, "mm/dd/yyyy") & "# AND Movil = '" & mMov & "'")
'         OkDia = False
'      End If
'   Loop
'   xRec.Close
'   OldFecha = xFecha
'   'Set mRecAux = mData.OpenRecordset("SELECT count(*) as rTotDiaP FROM Auxi WHERE Demora > #" & Format(DemoP, "hh:mm") & "# AND Movil = '" & mMov & "' AND mFecha = #" & Format(OldFecha, "mm/dd/yyyy") & "#")
'   Set mRecAux = mData.OpenRecordset("SELECT count(*) as rTotDiaP FROM Auxi WHERE Demora >'" & Format(CDate(DemoP), "hh:mm:ss") & "' AND Movil = '" & mMov & "' AND mFecha = #" & Format(OldFecha, "mm/dd/yyyy") & "#")
'   TotDiaP = mRecAux!rTotDiaP
'   mRecAux.Close
'   Set mRecAux = mData.OpenRecordset("SELECT count(*) as rTotDia FROM Auxi WHERE Movil = '" & mMov & "' AND mFecha = #" & Format(OldFecha, "mm/dd/yyyy") & "#")
'   TotDia = mRecAux!rTotDia
'   mRecAux.Close
'   mData.Execute ("UPDATE Auxi SET TotDia=" & TotDia & ",TotDiaP=" & TotDiaP & " WHERE mFecha = #" & Format(OldFecha, "mm/dd/yyyy") & "# AND Movil = '" & mMov & "'")
'   OkDia = False
'   'Set mRecAux = mData.OpenRecordset("SELECT count(*) as rTotMovP FROM Auxi WHERE Demora > #" & Format(DemoP, "hh:mm:ss") & "# AND Movil = '" & mMov & "'")
'   Set mRecAux = mData.OpenRecordset("SELECT count(*) as rTotMovP FROM Auxi WHERE Demora > '" & Format(DemoP, "hh:mm:ss") & "' AND Movil = '" & mMov & "'")
'   TotMovP = mRecAux!rTotMovP
'   mRecAux.Close
'   Set mRecAux = mData.OpenRecordset("SELECT count(*) as rTotMov FROM Auxi WHERE Movil='" & mMov & "'")
'   TotMov = mRecAux!rTotMov
'   mRecAux.Close
'   mData.Execute ("UPDATE Auxi SET TotMov=" & TotMov & ",TotMovP=" & TotMovP & " WHERE Movil = '" & mMov & "'")
'   'mData.Execute ("DELETE FROM Auxi WHERE Pedido = '00:00:00'")
'   Set xRec = Nothing
'End Sub
Private Sub sInitForm()
Dim mI As Integer
   Select Case mTabla
      Case "0"
         Set mRec = mObj.oMovilesGCO("PAT','GRU','GRP")
         Combo1.AddItem "TODOS"
         sLlenoCbo Me.Combo1, mRec, 1, 0
         Combo1.Left = 3100
         Combo1.Top = 2200
         Label2(0).Caption = "Fecha - Hora Inicial"
         Text1(2).Left = 2760
         Label2(1).Left = 4080
         Label2(1).Caption = "Fecha - Hora Final"
         Label2(2).Left = 2500
         Label2(2).Top = 2290
         Text1(1).Left = 4080
         Text1(3).Left = 5280
         Text1(2).Visible = True
         Text1(3).Visible = True
         
      Case "1", "9", "17", "23"
         Combo1.AddItem "TODOS"
         If mTabla = "9" Or mTabla = "23" Then
            Combo1.AddItem "TODAS PATR."
            Combo1.AddItem "TODAS G. LIV."
            Combo1.AddItem "TODAS G. PES."
            If mTabla = "9" Then
               Option1.Visible = True
               Label3(0).Caption = "Demora Mayor a:"
               Label3(0).Visible = True
               Label3(0).Left = 1560
               Text1(2).Visible = True
               Text1(2).Top = 2210
               Text1(2).Left = 3090
               Text1(2).TabIndex = 5
            End If



         End If
         Set mRec = mObj.oMovilesGCO("PAT','GRU','GRP")
         sLlenoCbo Me.Combo1, mRec, 1, 0
         
      Case "2"
         Combo1.AddItem "AMBU-AMBULANCIAS"
         Combo1.AddItem "BOMB-BOMBEROS "
         Combo1.AddItem "GEND-GENDARMERIA"
         Combo1.AddItem "POLI-POLICIA"
         Combo1.AddItem "AMBU EXT-AMBUL.EXT."
         Combo1.Width = 1600
         Combo1.ListIndex = 0
         
      Case "4"
         Combo1.AddItem "TODAS"
         Set mRec = mObj.oTabla("emisoras", "WHERE Fecha_Baja IS NULL ORDER BY Descripcion")
         sLlenoCbo Me.Combo1, mRec, 1, 0
         Label2(2).Caption = "Emisoras"
         Combo1.Width = 2100
     
      Case "5"
         Set mRec = mObj.oTabla("origen", "WHERE Fecha_Baja IS NULL ORDER BY Codigo")
         Combo1.AddItem "TODOS"
         Do While Not mRec.EOF
            Combo1.AddItem mRec!Codigo & "- " & mRec!descripcion
            mRec.MoveNext
         Loop
         mRec.Close
         Label2(2).Caption = "Origen"
         Combo1.Width = 2200
      
      Case "3", "6", "7", "12", "13", "14", "16", "18", "21", "22"
         Combo1.AddItem "h"
         Combo1.ListIndex = 0
         Combo1.Visible = False
         Label2(2).Visible = False
         If mTabla = "16" Then
            Label2(0).Caption = "Fecha - Hora Inicial"
            Text1(2).Left = 2760
            Label2(1).Left = 4080
            Label2(1).Caption = "Fecha - Hora Final"
            Text1(1).Left = 4080
            Text1(3).Left = 5280
            Text1(2).Visible = True
            Text1(3).Visible = True
            Text3.Visible = True
            Label3(0).Left = 1500
            Label3(0).Caption = "Filtro"
            Label3(0).Visible = True
            Check1.Visible = True
         Else
            Label2(0).Left = 2460
            Text1(0).Left = 2460
            Label2(1).Left = 3920
            Text1(1).Left = 3920
         End If
         
      Case "8"
         Combo1.AddItem "TODOS"
         Set mRec = mObj.oMovilesGCO("PAT")
         sLlenoCbo Me.Combo1, mRec, 1, 0
         
      Case "10"
         Label2(2).Caption = "Código"
         Combo1.Clear
         Combo1.Left = 2600
         Combo1.Width = 2300
         Label2(2).Left = 2600
         Label2(0).Visible = False
         Label2(1).Visible = False
         Text1(0).Text = "01/01/2002"
         Text1(0).Visible = False
         Text1(1).Text = "02/01/2002"
         Text1(1).Visible = False
         Label3(0).Visible = True
         Label3(1).Visible = True
         
      Case "11", "15"
         Combo1.AddItem "TODOS"
         Set mRec = mObj.oMovilesGCO("GRP','GRU")
         If mTabla = "11" Then
            Combo1.AddItem "TOTAL DETALLADO"
         End If
         sLlenoCbo Me.Combo1, mRec, 1, 0
         Combo1.Width = 1900
         
      Case "19"
        Label2(2).Caption = "Código"
        Combo1.Clear
        Command1(0).Visible = False
        Command1(1).Visible = False
        Command2(0).Visible = True
        Command2(1).Visible = True
        Text1(0).Visible = False
        Text1(1).Visible = False
        Label2(0).Visible = False
        Label2(1).Visible = False
        Combo1.Visible = False
        Text2.Visible = True
        Label2(2).Left = 3120
     
     Case "20"
         Combo1.AddItem "TODOS"
         If mTabla = "11" Then
            Combo1.AddItem "TOTAL DETALLADO"
         End If
         Set mRec = mObj.oMovilesPATGRU
         sLlenoCbo Me.Combo1, mRec, 1, 0
         Combo1.Width = 1900
   End Select
End Sub

Private Function sValida() As Boolean
   sValida = (Text1(0).Text <> "" And Text1(1).Text <> "" And Combo1.Text <> "")
   If Not sValida Then
      MsgBox "Faltan ingresar datos.", vbInformation, sMessage
   End If
   sValida = sValida And sValidFechaDesdeHasta(Text1(0).Text, Text1(1).Text)
End Function
