VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form Pek1_frm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contadores Peek Traffic"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8490
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   8490
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   240
         Top             =   3600
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   2
         Left            =   6280
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1800
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   4720
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1800
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   680
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1800
         Width           =   3815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Volver"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   4320
         TabIndex        =   2
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   2640
         TabIndex        =   1
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Consulta de Volúmen de Tránsito"
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
         Height          =   360
         Left            =   480
         TabIndex        =   9
         Top             =   480
         Width           =   7215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año"
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
         Left            =   6280
         TabIndex        =   7
         Top             =   1560
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes"
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
         Left            =   4720
         TabIndex        =   5
         Top             =   1560
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Contador"
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
         Left            =   680
         TabIndex        =   4
         Top             =   1560
         Width           =   780
      End
   End
End
Attribute VB_Name = "Pek1_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mData As Database
Dim mObj As New clPeek
Dim mRec As New ADODB.Recordset
Dim mi As Integer

Private Sub Form_Load()
   Set mData = OpenDatabase(App.Path & "\Peek\Peek.mdb")
   mRec.CursorType = adOpenDynamic
   Me.Width = 8610
   Me.Height = 4815
   sAlinearForm Me
   For mi = 1 To 12
      Combo1(1).AddItem MonthName(mi) & Space(20) & Format(mi, "00")
   Next
   For mi = 2016 To Year(Now)
      Combo1(2).AddItem mi
   Next
   Set mRec = mObj.oTabla("Contadores", "")
   sLlenoCbo Pek1_frm.Combo1(0), mRec, 1, 0
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   mData.Close
   Set mData = Nothing
   Set mObj = Nothing
   Set mRec = Nothing
   ShowMenu 13, True, False
End Sub

Private Sub Command1_Click(Index As Integer)
Dim mObjAcc As New clAccess
Dim mFecha As Date
Dim mHora As String
Dim mTitulo As String
Dim mCarriles As Integer
Dim mClases As Integer
Dim mReporte As Boolean
Dim mFlag As Boolean
Dim mWhere As String
Dim mCarril_A As String
Dim mCarril_D As String
Dim mValues As String
Dim mClase As String
Dim mContador As String
Dim mValores(8) As Integer
Dim mRs

   If Index = 0 Then
      If Combo1(0).ListIndex <> -1 And Combo1(1).ListIndex <> -1 And Combo1(2).ListIndex <> -1 Then
         mObjAcc.mBorrarAuxi "\Peek\Peek", "Auxi"
         mWhere = ""
         mReporte = True
         Select Case Command1(1).Tag
            Case "0"
               sMsgEspere Me, "Procesando datos...", True
               Select Case Command1(0).Tag
                  Case 0 'Volumétrico mensual
                     If Trim(Right(Combo1(0).Text, 4)) <> "Km23" Then
                        CrystalReport1.ReportFileName = App.Path & "\Peek\" & "Volumen.rpt"
                        mData.Execute ("CREATE TABLE Auxi (Flag TEXT,Fecha TEXT,TotalA INTEGER,TotalD INTEGER,Hora TEXT,Carril INTEGER)")
                        mCarril_A = fCarrilesSent("A")
                        mCarril_D = fCarrilesSent("D")
                        Set mRec = mObj.oSumSent(Trim(Right(Combo1(0).Text, 4)), Combo1(2).Text, Val(Right(Combo1(1).Text, 2)), mCarril_A, mCarril_D, "Fecha")
                        If Not mRec.EOF Then
                           mTitulo = "Volumétrico Mensual"
                           Do While Not mRec.EOF
                               mData.Execute ("INSERT INTO Auxi (Flag,Fecha,TotalA,TotalD) VALUES ('1','" & Format(mRec!Fecha, "dd/mm/yyyy") & "'," & mRec!A & "," & mRec!D & ")")
                               mRec.MoveNext
                           Loop
                        Else
                          
                           'completar Auxi con todos los días en 0.
                           If LCase(Trim(Right(Combo1(0).Text, 4))) = "km32" Or LCase(Trim(Right(Combo1(0).Text, 4))) = "km36" Then
                                fLlenarAuxi "wt" & Right(Combo1(0).Text, 2) & "c", Combo1(2).Text, Val(Right(Combo1(1).Text, 2)) '
                           End If
                        End If
                        mRec.Close
                        'Wavetronix
                        fDatosWT LCase(Trim(Right(Combo1(0).Text, 4))), Command1(0).Tag
                     Else 'Con colectora
                        CrystalReport1.ReportFileName = App.Path & "\Peek\" & "Volumen_23.rpt"
                        mData.Execute ("CREATE TABLE Auxi (Flag TEXT,Fecha TEXT,TotalA INTEGER,TotalD INTEGER, ColA INTEGER, ColD INTEGER)")
                        mCarril_A = "Carril1 + Carril2 + Carril3 + Carril4"
                        mCarril_D = "Carril5 + Carril6 + Carril7 + Carril8"
                        Set mRec = mObj.oSumSentCol(Trim(Right(Combo1(0).Text, 4)), Combo1(2).Text, Val(Right(Combo1(1).Text, 2)), mCarril_A, mCarril_D, "Carril9 + Carril10", "Carril11 + Carril12")
                        If Not mRec.EOF Then
                           mTitulo = "Volumétrico Mensual"
                           Do While Not mRec.EOF
                              mData.Execute ("INSERT INTO Auxi (Flag,Fecha,TotalA,TotalD,ColA,ColD) VALUES ('1','" & Format(mRec!Fecha, "dd/mm/yyyy") & "'," & mRec!A & "," & mRec!D & "," & mRec!ColA & "," & mRec!ColD & ")")
                              mRec.MoveNext
                           Loop
                        Else
                           MsgBox "No Existen Datos en la Base", vbInformation, sMessage
                           fLlenarAuxi "wt" & Right(Combo1(0).Text, 2) & "c", Combo1(2).Text, Val(Right(Combo1(1).Text, 2)) '
                           'mReporte = False
                        End If
                        mRec.Close
                        'Wavetronix
                        fDatosWT LCase(Trim(Right(Combo1(0).Text, 4))), Command1(0).Tag
                     End If
               
                  Case 1 'Mensual por días
                     mTitulo = "Volumétrico Mensual por Días"
                     mData.Execute ("CREATE TABLE Auxi (Flag TEXT,Fecha TEXT,C1 INTEGER,C2 INTEGER,C3 INTEGER,C4 INTEGER,C5 INTEGER,C6 INTEGER,C7 INTEGER,C8 INTEGER,C9 INTEGER,C10 INTEGER,C11 INTEGER,C12 INTEGER)")
                     mWhere = fCarriles()
                     CrystalReport1.ReportFileName = App.Path & "\Peek\" & Trim(Right(Combo1(0).Text, 4)) & "_TD.rpt"
                     Set mRec = mObj.oSumCarriles(Trim(Right(Combo1(0).Text, 4)), Combo1(2).Text, Val(Right(Combo1(1).Text, 2)), mWhere, "Fecha")
                     If Not mRec.EOF Then
                        Do While Not mRec.EOF
                           mWhere = ""
                           mValues = ""
                           mCarriles = mObj.iMaxCarril(Trim(Right(Combo1(0).Text, 4)))
                           For mi = 1 To mCarriles
                              mWhere = mWhere & "," & mRec.Fields(mi).Name
                              mValues = mValues & "," & mRec.Fields(mi)
                           Next
                           mData.Execute ("INSERT INTO Auxi (Flag,Fecha" & mWhere & ") VALUES ('1','" & Format(mRec!Fecha, "dd/mm/yyyy") & "'" & mValues & ")")
                           mRec.MoveNext
                        Loop
                     End If
                     mRec.Close
                     'wavetronix
                     fDatosWT LCase(Trim(Right(Combo1(0).Text, 4))), Command1(0).Tag
                     
                  Case 2 'Volumétrico Detalle de Días //falta WT
                     mTitulo = "Volumétrico Detalle de Días"
                     CrystalReport1.ReportFileName = App.Path & "\Peek\" & Trim(Right(Combo1(0).Text, 4)) & "_Dia.rpt"
                     mData.Execute ("CREATE TABLE Auxi (Fecha DATE,Hora TEXT,C1 INTEGER,C2 INTEGER,C3 INTEGER,C4 INTEGER,C5 INTEGER,C6 INTEGER,C7 INTEGER,C8 INTEGER,C9 INTEGER,C10 INTEGER,C11 INTEGER,C12 INTEGER)")
                     mWhere = fCarriles()
                     Set mRec = mObj.oSumCarriles(Trim(Right(Combo1(0).Text, 4)), Combo1(2).Text, Val(Right(Combo1(1).Text, 2)), mWhere, "Fecha, hora")
                     If Not mRec.EOF Then
                        mCarriles = mObj.iMaxCarril(Trim(Right(Combo1(0).Text, 4)))
                        Do While Not mRec.EOF
                           mWhere = ""
                           mValues = ""
                           For mi = 1 To mCarriles
                              mWhere = mWhere & ",C" & mi
                              mValues = mValues & "," & mRec.Fields(mi + 1)
                           Next
                           mData.Execute ("INSERT INTO Auxi (Fecha,Hora" & mWhere & ") VALUES ('" & Format(mRec!Fecha, "dd/mm/yyyy") & "','" & mRec!hora & "'" & mValues & ")")
                           mRec.MoveNext

                        Loop
                     End If
                     mRec.Close
                     'wavetronix
                     fDatosWT LCase(Trim(Right(Combo1(0).Text, 4))), Command1(0).Tag
                End Select
                sMsgEspere Me, "", False
                
            Case "1", "2" 'CLASES por LONGITUDES y VELOCIDADES
            sMsgEspere Me, "Procesando datos...", True
               mCarriles = 0
               mWhere = fCarriles()
               Select Case Command1(0).Tag
                  Case 0 'MENSUAL
                     If Command1(1).Tag = "1" Then 'Longitudes
                        mTitulo = "Volumétrico Mensual por LONGITUDES de  Vehículos "
                        mData.Execute ("CREATE TABLE Auxi (Flag TEXT,Clase TEXT,C1 INTEGER,C2 INTEGER,C3 INTEGER,C4 INTEGER,C5 INTEGER,C6 INTEGER,C7 INTEGER,C8 INTEGER,C9 INTEGER,C10 INTEGER,C11 INTEGER,C12 INTEGER)")
                        Set mRec = mObj.oSumCarriles(Trim(Right(Combo1(0).Text, 4)), Combo1(2).Text, Val(Right(Combo1(1).Text, 2)), mWhere, "longi")
                     Else
                        mTitulo = "Volumétrico Mensual por VELOCIDADES de Vehículos "
                        mData.Execute ("CREATE TABLE Auxi (Flag TEXT,Veloc TEXT,C1 INTEGER,C2 INTEGER,C3 INTEGER,C4 INTEGER,C5 INTEGER,C6 INTEGER,C7 INTEGER,C8 INTEGER,C9 INTEGER,C10 INTEGER,C11 INTEGER,C12 INTEGER)")
                        Set mRec = mObj.oSumCarriles(Trim(Right(Combo1(0).Text, 4)), Combo1(2).Text, Val(Right(Combo1(1).Text, 2)), mWhere, "veloc")
                     End If
                     CrystalReport1.ReportFileName = App.Path & "\Peek\" & Trim(Right(Combo1(0).Text, 4)) & "_Class.rpt"
                     If Not mRec.EOF Then
                        mCarriles = mObj.iMaxCarril(Trim(Right(Combo1(0).Text, 4)))
                        Do While Not mRec.EOF
                           mWhere = ""
                           mValues = ""
                           For mi = 1 To mCarriles
                              mWhere = mWhere & ",C" & mi
                              mValues = mValues & "," & mRec.Fields(mi)
                           Next
                           If Command1(1).Tag = "1" Then
                              mData.Execute ("INSERT INTO Auxi (Flag,Clase" & mWhere & ") VALUES ('1','" & mRec.Fields(0) & "'" & mValues & ")")
                           Else
                              mData.Execute ("INSERT INTO Auxi (Flag,Veloc" & mWhere & ") VALUES ('2','" & mRec.Fields(0) & "'" & mValues & ")")
                           End If
                           mRec.MoveNext
                        Loop
                     End If
                     mRec.Close
                     
                     'Wavetronix
                     sDatosWTVelLon LCase(Trim(Right(Combo1(0).Text, 4))), Command1(0).Tag, Command1(1).Tag
                     
                     
                  Case 1 'DIARIO
                     If Command1(1).Tag = "1" Then
                        mTitulo = "Volumétrico Diario por LONGITUDES de Vehículos "
                        Set mRec = mObj.oSumCarriles(Trim(Right(Combo1(0).Text, 4)), Combo1(2).Text, Val(Right(Combo1(1).Text, 2)), mWhere, "fecha, longi")
                        mWhere = "Clase"
                     Else
                        mTitulo = "Volumétrico Diario por VELOCIDADES de Vehículos "
                        Set mRec = mObj.oSumCarriles(Trim(Right(Combo1(0).Text, 4)), Combo1(2).Text, Val(Right(Combo1(1).Text, 2)), mWhere, "fecha,veloc")
                        mWhere = "Veloc"
                     End If
                     CrystalReport1.ReportFileName = App.Path & "\Peek\" & Trim(Right(Combo1(0).Text, 4)) & "_Cl_Dia.rpt"
                     mData.Execute ("CREATE TABLE Auxi (Flag TEXT,Fecha TEXT," & mWhere & " TEXT,C1 INTEGER,C2 INTEGER,C3 INTEGER,C4 INTEGER,C5 INTEGER,C6 INTEGER,C7 INTEGER,C8 INTEGER,C9 INTEGER,C10 INTEGER,C11 INTEGER,C12 INTEGER)")
                     If Not mRec.EOF Then
                        mCarriles = mObj.iMaxCarril(Trim(Right(Combo1(0).Text, 4)))
                        Do While Not mRec.EOF
                           mWhere = ""
                           mValues = ""
                           For mi = 1 To mCarriles
                              mWhere = mWhere & ",C" & mi
                              mValues = mValues & "," & mRec.Fields(mi + 1)
                           Next
                           If Command1(1).Tag = "1" Then
                              mData.Execute ("INSERT INTO Auxi (Flag,Fecha,Clase" & mWhere & ") VALUES ('1','" & mRec.Fields(0) & "','" & mRec.Fields(1) & "'" & mValues & ")")
                           Else
                              mData.Execute ("INSERT INTO Auxi (Flag,Fecha,Veloc" & mWhere & ") VALUES ('2','" & mRec.Fields(0) & "','" & mRec.Fields(1) & "'" & mValues & ")")
                           End If
                           mRec.MoveNext
                        Loop
                    End If
                    mRec.Close
                     
                    'Wavetronix
                    sDatosWTVelLon LCase(Trim(Right(Combo1(0).Text, 4))), Command1(0).Tag, Command1(1).Tag
                  Case 2 'DIARIO X HORA
                        mTitulo = "Volumétrico Diario por hora por VELOCIDADES de Vehículos "
                        'Set mRec = mObj.oSumCarriles(Trim(Right(Combo1(0).Text, 4)), Combo1(2).Text, Val(Right(Combo1(1).Text, 2)), mWhere, "fecha,hora,veloc")
                        mWhere = "Veloc"


                        CrystalReport1.ReportFileName = App.Path & "\Peek\" & Trim(Right(Combo1(0).Text, 4)) & "_Cl_Dia_Hora.rpt"
                        mData.Execute ("CREATE TABLE Auxi (Flag TEXT,Fecha TEXT,Hora TEXT," & mWhere & " TEXT,C1menor50 INTEGER,C1mayor50 INTEGER,C2menor50 INTEGER,C2mayor50 INTEGER,C3menor50 INTEGER,C3mayor50 INTEGER,C4menor50 INTEGER,C4mayor50 INTEGER,C5menor50 INTEGER,C5mayor50 INTEGER,C6menor50 INTEGER,C6mayor50 INTEGER,C7menor50 INTEGER,C7mayor50 INTEGER,C8menor50 INTEGER,C8mayor50 INTEGER,C9menor50 INTEGER,C9mayor50 INTEGER,C10menor50 INTEGER,C10mayor50 INTEGER,C11menor50 INTEGER,C11mayor50 INTEGER,C12menor50 INTEGER,C12mayor50 INTEGER)")

'                     If Not mRec.EOF Then
'                        mCarriles = mObj.iMaxCarril(Trim(Right(Combo1(0).Text, 4)))
'                        Do While Not mRec.EOF
'                           mWhere = ""
'                           mValues = ""
'                           For mI = 1 To mCarriles
'                              mWhere = mWhere & ",C" & mI
'                              mValues = mValues & "," & mRec.Fields(mI + 1)
'                           Next
'                           If Command1(1).Tag = "1" Then
'                              mData.Execute ("INSERT INTO Auxi (Flag,Fecha,Hora,Clase" & mWhere & ") VALUES ('1','" & mRec.Fields(0) & "','" & mRec.Fields(1) & "','" & mRec.Fields(2) & "'" & mValues & ")")
'                           Else
'                              mData.Execute ("INSERT INTO Auxi (Flag,Fecha,Hora,Veloc" & mWhere & ") VALUES ('2','" & mRec.Fields(0) & "','" & mRec.Fields(1) & "','" & mRec.Fields(2) & "'" & mValues & ")")
'                           End If
'                           mRec.MoveNext
'                        Loop
'                    End If
'                    mRec.Close
                    sDatosWTVelLon LCase(Trim(Right(Combo1(0).Text, 4))), Command1(0).Tag, Command1(1).Tag
               sMsgEspere Me, "", False
               Exit Sub
               
               End Select
               sMsgEspere Me, "", False
               'Exit Sub
               
            Case "3"
               mReporte = False
               If Combo1(0).Text = "TODOS" Then
                  Set mRec = mObj.oErrorPeek("", Combo1(2).Text, Val(Right(Combo1(1).Text, 2)))
                  mTitulo = "Contador " & Trim(Right(Combo1(0).Text, 4)) & ": "
               Else
                  Set mRec = mObj.oErrorPeek(UCase(Trim(Right(Combo1(0).Text, 4))), Combo1(2).Text, Val(Right(Combo1(1).Text, 2)))
                  mTitulo = ""
               End If
               If Not mRec.EOF Then
                  mData.Execute ("CREATE TABLE Auxi (Peek TEXT,Fecha TEXT,Descripcion TEXT,Archivo TEXT)")
                  Do While Not mRec.EOF
                     mData.Execute ("INSERT INTO Auxi (Peek,Fecha,Descripcion,Archivo) VALUES ('" & mRec!Peek & "','" & Format(mRec!Fecha, "dd/mm/yyyy") & "','" & mRec!descripcion & "','" & mRec!Archivo & "')")
                     mRec.MoveNext
                  Loop
                  CrystalReport1.ReportFileName = App.Path & "\Peek\Errores.rpt"
                  CrystalReport1.WindowTitle = "Reporte del Contador " & Trim(Right(Combo1(0).Text, 4))
                  CrystalReport1.DataFiles(0) = "Peek.mdb"
                  CrystalReport1.Formulas(0) = "Titulo = '" & mTitulo & "Reporte de Errores  - Mes: " & Trim(Left(Combo1(1).Text, 12)) & ".'"
                  CrystalReport1.WindowState = crptMaximized
                  CrystalReport1.Action = 1
                  Combo1(0).ListIndex = -1
                  Combo1(1).ListIndex = -1
                  Combo1(2).ListIndex = -1
               Else
                  MsgBox "No Existen Errores para el Mes Solicitado", vbInformation, sMessage
               End If
               mRec.Close
               
            Case "4"
               mData.Execute ("CREATE TABLE Auxi (Fecha TEXT,Clase TEXT, C1A INTEGER,C2A INTEGER,C3A INTEGER,C4A INTEGER,C5A INTEGER,C6A INTEGER,C7A INTEGER, C1D INTEGER,C2D INTEGER,C3D INTEGER,C4D INTEGER,C5D INTEGER,C6D INTEGER,C7D INTEGER)")
               mCarril_A = fCarrilesSent("A")
               mCarril_D = fCarrilesSent("D")
               Select Case Command1(0).Tag
                  Case 0
                     mTitulo = "Volumétrico Mensual de VELOCIDADES por LONGITUDES de Vehículos "
                     Set mRec = mObj.oSumSent(Trim(Right(Combo1(0).Text, 4)), Combo1(2).Text, Val(Right(Combo1(1).Text, 2)), mCarril_A, mCarril_D, "Longi,Veloc")
                  Case 1
                     mTitulo = "Volumétrico DIARIO de VELOCIDADES por LONGITUDES de Vehículos"
                     Set mRec = mObj.oSumSent(Trim(Right(Combo1(0).Text, 4)), Combo1(2).Text, Val(Right(Combo1(1).Text, 2)), mCarril_A, mCarril_D, "Fecha,Longi,Veloc")
               End Select
               If Trim(Right(Combo1(0).Text, 4)) = "Km47" Then
                  CrystalReport1.ReportFileName = App.Path & "\Peek\VelxLong2.rpt"
               Else
                  CrystalReport1.ReportFileName = App.Path & "\Peek\VelxLong.rpt"
               End If
               If Not mRec.EOF Then
                   mClase = mRec!Longi
                   mFecha = "20/06/1974"
                   If Command1(0).Tag <> "0" Then 'Diario de LONGITUDES Por VELOCIDADES
                      mFecha = mRec!Fecha
                   End If
                   Do While Not mRec.EOF
                      If mClase <> mRec!Longi Then
                         mData.Execute ("INSERT INTO Auxi (Fecha,Clase," & Mid(mWhere, 1, Len(mWhere) - 1) & ") VALUES ('" & mFecha & "','" & mClase & "'," & Mid(mValues, 1, Len(mValues) - 1) & ")")
                         mClase = mRec!Longi
                         mWhere = "C" & mRec!veloc & "A," & "C" & mRec!veloc & "D,"
                         mValues = mRec!A & "," & mRec!D & ","
                         If Command1(0).Tag <> "0" Then
                            mFecha = mRec!Fecha
                         End If
                      Else
                         mWhere = mWhere & "C" & mRec!veloc & "A," & "C" & mRec!veloc & "D,"
                         mValues = mValues & mRec!A & "," & mRec!D & ","
                      End If
                      mRec.MoveNext
                   Loop
                   mData.Execute ("INSERT INTO Auxi (Fecha,Clase," & Mid(mWhere, 1, Len(mWhere) - 1) & ") VALUES ('" & mFecha & "','" & mClase & "'," & Mid(mValues, 1, Len(mValues) - 1) & ")")
               Else
                  MsgBox "No Existen Datos en la Base"
                  mReporte = False
               End If
               mRec.Close '
         End Select
         If mReporte Then
            Set mRs = mData.OpenRecordset("select * from Auxi ")
            mRs.Close
            CrystalReport1.WindowTitle = "Reporte del Contador " & Trim(Right(Combo1(0).Text, 4))
            CrystalReport1.DataFiles(0) = "Peek.mdb"
            'CrystalReport1.Formulas(0) = "Titulo = 'Contador " & Trim(Right(Combo1(0).Text, 4)) & ": " & mTitulo & "  - Mes: " & Trim(Left(Combo1(1).Text, 12)) & ".'"
            
            If Trim(Right(Combo1(0).Text, 4)) = "U2ug" Then
               Dim nombreContador As String
               nombreContador = Trim(Left(Combo1(0).Text, Len(Combo1(0).Text) - 4))
               nombreContador = Trim(Left(nombreContador, Len(nombreContador) - 6))
               CrystalReport1.Formulas(0) = "Titulo = '" & nombreContador & ": " & mTitulo & "  - Mes: " & Trim(Left(Combo1(1).Text, 12)) & ".'"
            Else
               CrystalReport1.Formulas(0) = "Titulo = '" & Trim(Left(Combo1(0).Text, Len(Combo1(0).Text) - 4)) & ": " & mTitulo & "  - Mes: " & Trim(Left(Combo1(1).Text, 12)) & ".'"
            End If
            
            
            
            CrystalReport1.WindowState = crptMaximized
            'CrystalReport1.SelectionFormula
            'CrystalReport1.SortFields (2)
            CrystalReport1.Action = 1
            Combo1(0).ListIndex = -1
            Combo1(1).ListIndex = -1
            Combo1(2).ListIndex = -1
         End If
      Else
         MsgBox "Faltan Completar Datos"
      End If
   Else
      Unload Pek1_frm
   End If
   Set mObjAcc = Nothing
End Sub

Private Function fDatosWT(ByVal pCont As String, ByVal pOpt As String)
Dim mRec2 As New ADODB.Recordset
Dim mRs
Dim mFechaIni As String
Dim mClase As String
Dim mi As Integer
Dim mj As Integer
Dim mFlag As Boolean

   mFechaIni = DateAdd("d", 1, Now)
   mFlag = True
   Select Case pCont
      Case "km14"
         mFechaIni = "01/10/2011"
      Case "km23"
         mFechaIni = "01/05/2010"
      Case "km29"
         mFechaIni = "01/12/2013"
      Case "km32"
         mFechaIni = "01/09/2010"
      Case "km36"
         mFechaIni = "01/02/2011"
      Case "rta5"
         mFechaIni = "01/11/2011"
      Case "urug"
         mFechaIni = "01/08/2016"
      Case "esco" ' agregar el idcontador en minuscula
         mFechaIni = "01/09/2016"
      Case "test"
         mFechaIni = "01/09/2016"
      Case "pi46"
         mFechaIni = "01/09/2016"
      Case "belg"
         mFechaIni = "01/01/2017"
      Case "masc"
         mFechaIni = "01/01/2017"
      Case "balb"
         mFechaIni = "01/01/2017"
      Case "boqu" ' agregar el idcontador en minuscula
         mFechaIni = "01/01/2017"
      Case "u2ug"
         mFechaIni = "01/01/2020"
      Case Else
         mFlag = False
   End Select
   
   If DateDiff("d", mFechaIni, "01/" & Right(Combo1(1).Text, 2) & "/" & Combo1(2).Text) >= 0 And mFlag Then
      Select Case pOpt
         Case "0" 'mensual
            mClase = " and carril in (" & mObj.sCarrilesSent(pCont, "A") & ")" 'detecta clases de WT
            For mi = 1 To 2
               'Set mRec2 = mObj.oSumCarriles("wt" & Right(pCont, 2) & "c", Combo1(2).Text, Val(Right(Combo1(1).Text, 2)), "sum(c1+c2+c3+c4+c5+c6+c7+c8+c9) as total", "fecha", mClase)
               Set mRec2 = mObj.oSumCarriles("wt" & Left(pCont, 2) & "c", Combo1(2).Text, Val(Right(Combo1(1).Text, 2)), "sum(c1+c2+c3+c4+c5+c6+c7+c8+c9) as total", "fecha", mClase)
               Do While Not mRec2.EOF
                  Set mRs = mData.OpenRecordset("select * from Auxi where fecha=#" & Mid(mRec2!Fecha, 4, 2) & "/" & Left(mRec2!Fecha, 2) & "/" & Right(mRec2!Fecha, 4) & "#")
                  If mRs.EOF Then
                     mData.Execute ("INSERT INTO Auxi (Flag,Fecha,TotalA,TotalD) VALUES ('1','" & Format(mRec2!Fecha, "dd/mm/yyyy") & "',0,0)")
                  End If
                  mRs.Close
                  If mi = 1 Then
                     mData.Execute "update Auxi set TotalA=" & mRec2!Total & " where fecha=#" & Mid(mRec2!Fecha, 4, 2) & "/" & Left(mRec2!Fecha, 2) & "/" & Right(mRec2!Fecha, 4) & "#"
                  Else
                     mData.Execute "update Auxi set TotalD=" & mRec2!Total & " where fecha=#" & Mid(mRec2!Fecha, 4, 2) & "/" & Left(mRec2!Fecha, 2) & "/" & Right(mRec2!Fecha, 4) & "#"
                  End If
                  mRec2.MoveNext
               Loop
               mRec2.Close
               mClase = " and carril in (" & mObj.sCarrilesSent(pCont, "D") & ")"  'detecta clases de WT
            Next
               
         Case "1" 'diario
            mClase = " and carril in (" & mObj.sCarrilesSent(pCont, "A','D") & ")" 'detecta clases de WT
            'Set mRec2 = mObj.oSumCarrilWT("wt" & Right(pCont, 2) & "c", Combo1(2).Text, Val(Right(Combo1(1).Text, 2)), mClase)
            Set mRec2 = mObj.oSumCarrilWT("wt" & Left(pCont, 2) & "c", Combo1(2).Text, Val(Right(Combo1(1).Text, 2)), mClase)
            Do While Not mRec2.EOF
               Set mRs = mData.OpenRecordset("select * from Auxi where fecha=#" & Mid(mRec2!Fecha, 4, 2) & "/" & Left(mRec2!Fecha, 2) & "/" & Right(mRec2!Fecha, 4) & "#")
               If mRs.EOF Then
                  mData.Execute ("INSERT INTO Auxi  VALUES ('1','" & Format(mRec2!Fecha, "dd/mm/yyyy") & "',0,0,0,0,0,0,0,0,0,0,0,0)")
               End If
               mRs.Close
               If Not (mRec2!carril = "C11" Or mRec2!carril = "C12") Then
                  mData.Execute "update Auxi set " & mRec2!carril & "=" & mRec2!Total & " where fecha=#" & Mid(mRec2!Fecha, 4, 2) & "/" & Left(mRec2!Fecha, 2) & "/" & Right(mRec2!Fecha, 4) & "# "
               End If
               mRec2.MoveNext
            Loop
            mRec2.Close
         
         Case "2" 'diario por hora
            mClase = " and carril in (" & mObj.sCarrilesSent(pCont, "A") & ")" 'detecta clases de WT
            For mi = 1 To 2
               'Set mRec2 = mObj.oSumCarrilWTHora("wt" & Right(pCont, 2) & "c", Combo1(2).Text, Val(Right(Combo1(1).Text, 2)), mClase)
               Set mRec2 = mObj.oSumCarrilWTHora("wt" & Left(pCont, 2) & "c", Combo1(2).Text, Val(Right(Combo1(1).Text, 2)), mClase)
               Do While Not mRec2.EOF
                  Set mRs = mData.OpenRecordset("select * from Auxi where fecha=#" & Mid(mRec2!Fecha, 4, 2) & "/" & Left(mRec2!Fecha, 2) & "/" & Right(mRec2!Fecha, 4) & "#")
                  If mRs.EOF Then
                     For mj = 0 To 23
                        mData.Execute ("INSERT INTO Auxi  VALUES ('" & Format(mRec2!Fecha, "dd/mm/yyyy") & "','" & Format(mj, "00") & ":00',0,0,0,0,0,0,0,0,0,0,0,0)")
                     Next
                  End If
                  mRs.Close
                  'If Not (mRec2!carril = "C9" Or mRec2!carril = "C10" Or mRec2!carril = "C11" Or mRec2!carril = "C12") Then
                  If Not (mRec2!carril = "C11" Or mRec2!carril = "C12") Then
                     mData.Execute "update Auxi set " & mRec2!carril & "=" & mRec2!Total & " where fecha=#" & Mid(mRec2!Fecha, 4, 2) & "/" & Left(mRec2!Fecha, 2) & "/" & Right(mRec2!Fecha, 4) & "# and hora='" & mRec2!hora & ":00' "
                  End If
                  mRec2.MoveNext
               Loop
               mRec2.Close
               mClase = " and carril in (" & mObj.sCarrilesSent(pCont, "D") & ")" 'detecta clases de WT
            Next
         
      End Select
   End If
   Set mRec2 = Nothing
End Function

Private Sub sDatosWTVelLon(ByVal pCont As String, ByVal pOpc As String, ByVal pLong As String)
Dim mRec2 As New ADODB.Recordset
Dim mRs
Dim mFechaIni As String
Dim mClase As String
Dim mi As Integer
Dim mj As Integer
Dim mFlag As Boolean

   mFechaIni = DateAdd("d", 1, Now)
   mFlag = True
   Select Case pCont
      Case "km14"
         mFechaIni = "01/10/2011"
      Case "balb"
         mFechaIni = "01/03/2017"
      Case "boqu"
         mFechaIni = "01/03/2017"
      Case "belg"
         mFechaIni = "01/03/207"
      Case "urug"
         mFechaIni = "01/03/2017"
      Case "masc"
         mFechaIni = "01/03/2017"
      Case "esco"
         mFechaIni = "01/03/2017"
      Case "pi46"
         mFechaIni = "01/03/2017"
      Case Else
         mFlag = False
   End Select
   
   If DateDiff("d", mFechaIni, "01/" & Right(Combo1(1).Text, 2) & "/" & Combo1(2).Text) >= 0 And mFlag Then
      mClase = " and carril in (" & mObj.sCarrilesSent(pCont, "A','D") & ")" 'detecta clases de WT
      Select Case pOpc
         Case "0" 'mensual
            If pLong = "1" Then 'Longitudes
               Set mRec2 = mObj.oSumLongWT("wt" & Left(pCont, 2) & "c", Combo1(2).Text, Val(Right(Combo1(1).Text, 2)), mClase)
            Else
               Set mRec2 = mObj.oSumVelWT("wt" & Left(pCont, 2) & "v", Combo1(2).Text, Val(Right(Combo1(1).Text, 2)), mClase)
            End If
            If Not mRec2.EOF Then
               Set mRs = mData.OpenRecordset("select * from Auxi")
               If mRs.EOF Then
                  For mi = 1 To 7
                     mData.Execute ("INSERT INTO Auxi  VALUES ('" & pLong & "','" & mi & "',0,0,0,0,0,0,0,0,0,0,0,0)")
                  Next
                  If pLong = "1" Then mData.Execute ("INSERT INTO Auxi  VALUES ('1','8',0,0,0,0,0,0,0,0,0,0,0,0)")
               End If
               mRs.Close
               Do While Not mRec2.EOF
                  If Command1(1).Tag = "1" Then 'Longitudes
                     For mi = 1 To 8
                        mData.Execute "update Auxi set " & mRec2!carril & "=" & mRec2.Fields(mi) & " where clase='" & mi & "'"
                     Next
                  Else
                     For mi = 1 To 7
                        mData.Execute "update Auxi set " & mRec2!carril & "=" & mRec2.Fields(mi) & " where veloc='" & mi & "'"
                     Next
                  End If
                  mRec2.MoveNext
               Loop
            End If
            mRec2.Close
            
         Case "1" 'diario
            If pLong = "1" Then 'Longitudes
               Set mRec2 = mObj.oSumLongWTFecha("wt" & Left(pCont, 2) & "c", Combo1(2).Text, Val(Right(Combo1(1).Text, 2)), mClase)
            Else
               Set mRec2 = mObj.oSumVelWTFecha("wt" & Left(pCont, 2) & "v", Combo1(2).Text, Val(Right(Combo1(1).Text, 2)), mClase)
            End If '--mp20170404
            Do While Not mRec2.EOF
               Set mRs = mData.OpenRecordset("select * from Auxi where fecha=#" & Mid(mRec2!Fecha, 4, 2) & "/" & Left(mRec2!Fecha, 2) & "/" & Right(mRec2!Fecha, 4) & "#")
               If mRs.EOF Then
                  For mi = 1 To 7
                     mData.Execute ("INSERT INTO Auxi  VALUES ('" & pLong & "','" & Format(mRec2!Fecha, "dd/mm/yyyy") & "','" & mi & "',0,0,0,0,0,0,0,0,0,0,0,0)")
                  Next
                  If pLong = "1" Then mData.Execute ("INSERT INTO Auxi  VALUES ('1','" & Format(mRec2!Fecha, "dd/mm/yyyy") & "','8',0,0,0,0,0,0,0,0,0,0,0,0)")
               End If
               mRs.Close
               If Command1(1).Tag = "1" Then 'Longitudes
                  For mi = 1 To 8
                     mData.Execute "update Auxi set " & mRec2!carril & "=" & mRec2.Fields(mi + 1) & " where clase='" & mi & "' and fecha=#" & Mid(mRec2!Fecha, 4, 2) & "/" & Left(mRec2!Fecha, 2) & "/" & Right(mRec2!Fecha, 4) & "# "
                  Next
               Else
                  For mi = 1 To 7
                     mData.Execute "update Auxi set " & mRec2!carril & "=" & mRec2.Fields(mi + 1) & " where veloc='" & mi & "' and fecha=#" & Mid(mRec2!Fecha, 4, 2) & "/" & Left(mRec2!Fecha, 2) & "/" & Right(mRec2!Fecha, 4) & "# "
                  Next
               End If
               mRec2.MoveNext
            Loop
            mRec2.Close
         Case "2" 'diario x hora
            If pLong = "1" Then 'Longitudes
               'Set mRec2 = mObj.oSumLongWTFechaHora("wt" & Left(pCont, 2) & "c", Combo1(2).Text, Val(Right(Combo1(1).Text, 2)), mClase)
            Else
              Set mRec2 = mObj.oSumVelWTFechaHoraParaAtascos("wt" & Left(pCont, 2) & "v_Atascos_carrilVeloc", Combo1(2).Text, Val(Right(Combo1(1).Text, 2)), mClase)
            End If
            
            Dim qtyCarriles As Integer
            Dim i As Integer
            Dim rango As String
            Dim mFila As Double
            Dim XLS As EXCEL.Application
            
            qtyCarriles = mObj.iMaxCarril(pCont)
            i = 0
            rango = ""
            Screen.MousePointer = vbHourglass
            
            Set XLS = CreateObject("Excel.Application")
            XLS.WorkBooks.Add
            
            XLS.Cells(1, 1).Formula = "AUTOPISTAS DEL SOL S.A."
            XLS.Cells(2, 1).Formula = "Indicador de atascos OEA"
            XLS.Cells(4, 1).Formula = UCase(Left(Combo1(0).Text, 40))
            XLS.Cells(5, 1).Formula = "PERIODO: " & Right(Combo1(1).Text, 2) & "/" & Combo1(2).Text
            XLS.Range("A1:A5").Font.Bold = True
            XLS.Range("A1:A1").Font.Size = 14
            XLS.Range("A2:A2").Font.Size = 12
                        
            XLS.Cells(7, 1).Formula = "FECHA"
            XLS.Cells(7, 2).Formula = "HORA"
            
            XLS.Cells(7, 3).Formula = "ASCENDENTE"
            XLS.Cells(7, qtyCarriles + 3).Formula = "DESCENDENTE"
            
            'Titulos celdas de carriles
            For i = 1 To qtyCarriles
               XLS.Cells(8, 1 + (i * 2)).Formula = "CARRIL " & i
            Next i
                        
            '--MERGE celdas de CARRILES
            For i = 1 To qtyCarriles
               rango = Chr(65 + (2 * i)) & "8:" & Chr(65 + (2 * i) + 1) & "8"
               XLS.Range(rango).Merge
            Next i
            
           'Titulos: Menro 50 km/h. , Mayor 50 km/h.
            For i = 1 To qtyCarriles
               XLS.Cells(9, 1 + (2 * i)).Formula = "Menor 50 km/h."
               XLS.Cells(9, 1 + (2 * i) + 1).Formula = "Mayor 50 km/h."
            Next i
            
            'ANCHO COLUMNAS: Mayor 50 km/h. Menor 50 km/h.
            rango = "C9:" & Chr(65 + (2 * i) + 1) & "9"
            XLS.Range(rango).ColumnWidth = 14
            
            'Merge Fecha, Hora
            XLS.Range("A7: A9").Merge
            XLS.Range("B7: B9").Merge
            
            'Merge celdas: Ascendente , Descendente
            XLS.Range("C7:" & Chr(65 + qtyCarriles + 1) & "7").Merge
            XLS.Range(Chr(65 + qtyCarriles + 2) & "7:" & Chr(65 + (2 * qtyCarriles) + 1) & "7").Merge
            
            'Formateo de header de grilla
            XLS.Range("A7:" & Chr(65 + (2 * qtyCarriles) + 1) & "9").HorizontalAlignment = xlCenter
            XLS.Range("A7:" & Chr(65 + (2 * qtyCarriles) + 1) & "9").VerticalAlignment = xlCenter
            XLS.Range("A7:" & Chr(65 + (2 * qtyCarriles) + 1) & "9").Font.Bold = True
            XLS.Range("A7:" & Chr(65 + (2 * qtyCarriles) + 1) & "9").Interior.Color = RGB(222, 222, 222)
            XLS.Range("A7:" & Chr(65 + (2 * qtyCarriles) + 1) & "9").Borders.Color = RGB(0, 0, 0)
            XLS.Range("A7:" & Chr(65 + (2 * qtyCarriles) + 1) & "9").Borders.Weight = 3
            
            If Not mRec2.EOF Then
            mFila = 10
               Do While Not mRec2.EOF
                  'Completo registros de grilla
                  XLS.Cells(mFila, 1).Formula = "'" & Format(mRec2!Fecha, "DD/MM/YYYY")
                  XLS.Cells(mFila, 2).Formula = mRec2!hora & ":00"
                  'Completo datos de las columnas: Mayor 50 km/h. Menor 50 km/h.
                  For i = 1 To qtyCarriles
                     XLS.Cells(mFila, (2 * i) + 1).Formula = mRec2.Fields(2 * i)
                     XLS.Cells(mFila, (2 * i) + 2).Formula = mRec2.Fields((2 * i) + 1)
                  Next i
                  mFila = mFila + 1
                  mRec2.MoveNext
               Loop
            End If
            mRec2.Close
            
            XLS.Visible = True
            Set XLS = Nothing
            Screen.MousePointer = vbArrow
''--------------------------------------------------------------------------------------------------------------------
''''            'MP20170404
''''            Do While Not mRec2.EOF
''''               Set mRs = mData.OpenRecordset("select * from Auxi where fecha=#" & Mid(mRec2!Fecha, 4, 2) & "/" & Left(mRec2!Fecha, 2) & "/" & Right(mRec2!Fecha, 4) & "# " & " and hora = '" & mRec2!Hora & ":00'")
''''               If mRs.EOF Then
''''                  mData.Execute ("INSERT INTO Auxi  VALUES ('" & pLong & "','" & Format(mRec2!Fecha, "dd/mm/yyyy") & "','" & mRec2!Hora & ":00','" & mI & "',0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)")
''''                  If pLong = "1" Then mData.Execute ("INSERT INTO Auxi  VALUES ('1','" & Format(mRec2!Fecha, "dd/mm/yyyy") & "','" & mRec2!Hora & "','8',0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)")
''''               End If
''''               mRs.Close
''''               If Command1(1).Tag = "1" Then 'Longitudes
'''''                  For mI = 1 To 8
'''''                    mData.Execute "update Auxi set " & mRec2!carril & "=" & mRec2.Fields(mI + 2) & " where clase='" & mI & "' and fecha=#" & Mid(mRec2!Fecha, 4, 2) & "/" & Left(mRec2!Fecha, 2) & "/" & Right(mRec2!Fecha, 4) & "# " & " and hora = '" & mRec2!Hora & "'"
'''''                  Next
''''               Else
''''                      mData.Execute "update Auxi set " & mRec2!carril & "menor50 = " & mRec2!vmenor50 & ", " & mRec2!carril & "mayor50 = " & mRec2!vmayor50 & " where fecha=#" & Mid(mRec2!Fecha, 4, 2) & "/" & Left(mRec2!Fecha, 2) & "/" & Right(mRec2!Fecha, 4) & "# " & " and hora = '" & mRec2!Hora & ":00'"
''''               End If
''''                   mRec2.MoveNext
''''            Loop
''''            mRec2.Close
      End Select
   End If
End Sub

Private Function fCarrilesSent(ByVal pSent As String) As String
   fCarrilesSent = ""
   Set mRec = mObj.oPeekCarril_Sent(Trim(Right(Combo1(0).Text, 4)), pSent)
   Do While Not mRec.EOF
     fCarrilesSent = fCarrilesSent & "Carril" & mRec!carril & "+"
     mRec.MoveNext
   Loop
   mRec.Close
   fCarrilesSent = Mid(fCarrilesSent, 1, Len(fCarrilesSent) - 1)
End Function

Private Function fCarriles() As String
   fCarriles = ""
   Set mRec = mObj.oPeekCarril_Sent(Trim(Right(Combo1(0).Text, 4)), "")
   Do While Not mRec.EOF
      fCarriles = fCarriles & "sum(Carril" & mRec!carril & ") as C" & mRec!carril & ","
      mRec.MoveNext
   Loop
   mRec.Close
   fCarriles = Mid(fCarriles, 1, Len(fCarriles) - 1)
End Function

Private Function fLlenarAuxi(ByVal pCont As String, ByVal pAnio As String, ByVal pMes As String)
Dim mRec1 As New ADODB.Recordset    'acá llena la tabla Auxi si el mes y año corresponden con datos de WT
   
   Set mRec1 = mObj.oWTDias(pCont, pAnio, pMes)
   Do While Not mRec1.EOF
      mData.Execute ("INSERT INTO Auxi (Flag,Fecha,TotalA,TotalD) VALUES ('1','" & Format(mRec1!Fecha, "dd/mm/yyyy") & "',0,0)")
      mRec1.MoveNext
   Loop
End Function

Private Sub sMensual(ByVal pCont As String, ByVal pFecha As String)
Dim mObj As New clPeek
Dim mRec As New ADODB.Recordset
Dim mCarril_A As String
Dim mCarril_D As String
   
   mCarril_A = fCarrilesSent("A")
   mCarril_D = fCarrilesSent("D")
      
   
   
   
   
End Sub
