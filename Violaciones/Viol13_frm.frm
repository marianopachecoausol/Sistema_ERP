VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form Viol13_frm 
   BackColor       =   &H00CECECE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes - "
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "Viol13_frm.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   6030
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1140
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   2
      Left            =   3120
      MaxLength       =   10
      TabIndex        =   1
      Top             =   1380
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   0
      Top             =   1380
      Visible         =   0   'False
      Width           =   975
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   240
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Salir"
      Height          =   375
      Index           =   1
      Left            =   3540
      TabIndex        =   4
      Top             =   2940
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
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
      Height          =   495
      Index           =   0
      Left            =   1800
      TabIndex        =   3
      Top             =   2820
      Width           =   1275
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   3120
      MaxLength       =   2
      TabIndex        =   2
      Top             =   1860
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CECECE&
      Caption         =   "Costo Administrativo $"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   1080
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CECECE&
      Caption         =   "Cant. MIN de Cartas "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   1800
      TabIndex        =   5
      Top             =   840
      Width           =   1800
   End
End
Attribute VB_Name = "Viol13_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mObjViol As New clViolaciones
Dim mObjPea As New clPeaje
Dim mRec As New ADODB.Recordset
Dim mRec2 As New ADODB.Recordset
Dim mData As Database
Public mReporte As Integer

Private Sub Form_Load()
sAlinearForm Me
sInitForm
Set mData = OpenDatabase(App.Path & "\Violaciones\Auxiliar.mdb")
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set mObjViol = Nothing
Set mObjPea = Nothing
Set mRec = Nothing
Set mRec2 = Nothing
mData.Close
ShowMenu 5, True, False
End Sub

Private Sub Command1_Click(Index As Integer)
 Dim mObjAcc As New clAccess
 Dim mRs As New ADODB.Recordset
 Dim mVector(20) As String
 Dim Entrega As String
 Dim mTotales(7) As Integer
 Dim mI As Integer
 Dim mFecha1, mFecha2 As String
 Dim mTotal As Integer
 Dim mParcial As Integer
 Dim mAuxi
 
 
 If Index = 0 Then
   mObjAcc.mBorrarAuxi "\Violaciones\Auxiliar", "Auxi"
   Select Case mReporte
      Case 0
         Set mRec = mObjViol.oCountEnviosCD(Trim(Text1(0).Text))
         If Not mRec.EOF Then
            Set mRs = mObjViol.oCodEntrega
            Do While Not mRs.EOF
               If Int(mRs!Codigo) <> 99 Then
                  mVector(Int(mRs!Codigo)) = mRs!Descripcion
               Else
                  mVector(20) = mRs!Descripcion
               End If
               mRs.MoveNext
            Loop
            mRs.Close
            mData.Execute ("CREATE TABLE Auxi (mPatente TEXT, mFecha TEXT, NroCarta TEXT, entrega TEXT, obs TEXT, nombre TEXT, total INTEGER)")
            Set mObjAcc = Nothing
            Do While Not mRec.EOF
                Set mRec2 = mObjViol.oDetalleEnviosCD(mRec!patente)
                Do While Not mRec2.EOF
                    Entrega = ""
                    If mRec2!codentrega <> "" Then
                        If Int(mRec2!codentrega) <> 99 Then
                           Entrega = mVector(Int(mRec2!codentrega))
                        Else
                           Entrega = mVector(16)
                        End If
                    End If
                    mData.Execute "INSERT INTO Auxi VALUES ('" & mRec2!patente & "','" & mRec2!Fecha & "','" & mRec2!NROCARTA & "','" & Entrega & "','" & mRec2!OBS & "','" & mRec2!nombre & "'," & mRec!Total & ")"
                    mRec2.MoveNext
                Loop
                mRec2.Close
                mRec.MoveNext
            Loop
            Set mAuxi = mData.OpenRecordset("SELECT count(*) as Total FROM Auxi")
            mAuxi.Close
            CrystalReport1.WindowTitle = "Reporte Envíos de Cartas Documentos"
            CrystalReport1.Formulas(0) = "Listado = 'Cantidad de envíos de Cartas Documentos >= " & Trim(Text1(0).Text) & "'"
            sImprimir "\Violaciones\rep02.rpt"
         Else
            MsgBox "No se encontraron datos con este parámetro.", vbInformation, sMessage
         End If
         mRec.Close
         
      Case 1
         If Fecha_ok(Text1(1).Text) And Fecha_ok(Text1(2).Text) Then
            If DateDiff("d", Text1(1).Text, Text1(2).Text) >= 0 Then
               mData.Execute ("CREATE TABLE Auxi (recibido INTEGER, rechazo INTEGER, error INTEGER, otros INTEGER, cartas INTEGER, monto DOUBLE, montoMes DOUBLE)")
               Set mRec = mObjViol.oTotalEnviosEntrega(Text1(1).Text, Text1(2).Text)
               For mI = 1 To 7
                  mTotales(mI) = 0
               Next
               Do While Not mRec.EOF
                  Select Case mRec!codentrega
                     Case "00"
                        mTotales(1) = mRec!Total
                     Case "04"
                        mTotales(2) = mRec!Total
                     Case "02", "03", "08", "14"
                        mTotales(3) = mTotales(3) + mRec!Total
                     Case Else
                        mTotales(4) = mTotales(4) + mRec!Total
                   End Select
                   mRec.MoveNext
               Loop
               mRec.Close
               Set mRec = mObjPea.oCartaDocCobradas(Text1(1).Text, Text1(2).Text)
               If Not mRec.EOF Then
                  mTotales(5) = NVL(mRec!Total, 0)
                  mTotales(6) = NVL(mRec!monto, 0)
               End If
               mRec.Close
               mData.Execute "INSERT INTO Auxi VALUES (" & mTotales(1) & "," & mTotales(2) & "," & mTotales(3) & "," & mTotales(4) & "," & mTotales(5) & "," & mTotales(6) & "," & mTotales(7) & ")"
               Set mAuxi = mData.OpenRecordset("SELECT count(*) as Total FROM Auxi")
               mAuxi.Close
               CrystalReport1.WindowTitle = "Reporte Envíos de Cartas Documentos"
               CrystalReport1.Formulas(0) = "Listado = 'Período evaluado: " & Trim(Text1(1).Text) & " al " & Trim(Text1(2).Text) & "'"
               sImprimir "\Violaciones\rep03.rpt"
            Else
               MsgBox "La fecha inicial es mayor a la final", vbCritical, sMessage
            End If
         End If
         
      Case 2
         If Combo1.ListIndex > -1 Then
               Me.MousePointer = 11
               mFecha1 = "01/" & Format(Trim(Right(Combo1.Text, 2)), "00") & "/" & Left(Combo1.Text, 4)
               mFecha2 = DateAdd("m", 1, mFecha1)
               mFecha2 = DateAdd("d", -1, mFecha2)
               mData.Execute ("CREATE TABLE Auxi (patente TEXT, nrocarta TEXT, fecha TEXT, pasadas INTEGER, Total INTEGER, ParciaL INTEGER)")
               Set mRec = mObjViol.oCartasEnPeriodo(mFecha1, mFecha2)
               Set mRec2 = mObjViol.oCountViolDspFecha(mFecha1, mFecha2)
               mTotal = 0
               mParcial = 0
               Do While Not mRec.EOF
                  mI = 0
                  mTotal = mTotal + 1
                  If Not mRec2.EOF Then
                     If mRec!patente = mRec2!patente Then
                        mI = mRec2!Total
                        mRec2.MoveNext
                        mParcial = mParcial + 1
                     End If
                  End If
                  mData.Execute ("INSERT INTO Auxi VALUES ('" & mRec!patente & "','" & mRec!NROCARTA & "','" & mRec!Fecha & "'," & mI & ",0,0)")
                  mRec.MoveNext
               Loop
               mRec.Close
               mRec2.Close
               mData.Execute ("UPDATE Auxi SET Total=" & mTotal & ", Parcial=" & mParcial & "")
               Set mAuxi = mData.OpenRecordset("SELECT count(*) as Total FROM Auxi")
               mAuxi.Close
               Me.MousePointer = 0
               CrystalReport1.WindowTitle = "Reporte Seguimiento de Patentes"
               CrystalReport1.Formulas(0) = "Listado = 'Período evaluado: " & mFecha1 & " al " & mFecha2 & "'"
               sImprimir "\Violaciones\rep04.rpt"
         Else
            MsgBox "Seleccionar un período", vbCritical, sMessage
         End If
   End Select
 Else
    Unload Viol13_frm
 End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
   Case 0
      KeyAscii = fNumeroKeyPress(KeyAscii)
   Case 1, 2
      KeyAscii = fDateKeyPress(Viol13_frm.Text1(Index), KeyAscii)
End Select
End Sub

Private Sub sInitForm()
Select Case mReporte
   Case 0
      Label1(0).Left = 1500
      Label1(0).Top = 1140
      Text1(0).Visible = True
      Text1(0).Left = 3300
      Text1(0).Top = 1080
      Me.Caption = Me.Caption & " Envíos de Cartas Documentos."
      Label1(0).Caption = "Cant. MIN de Cartas "
      
   Case 1
      Label1(0).Left = 1800
      Label1(0).Top = 840
      Text1(1).Visible = True
      Text1(1).Left = 1800
      Text1(1).Top = 1080
      Text1(2).Visible = True
      Text1(2).Left = 3240
      Text1(2).Top = 1080
      Me.Caption = Me.Caption & " Seguimientos de envíos de cartas documentos."
      Label1(0).Caption = "Fecha desde  -  hasta"
   
   Case 2
      Label1(0).Left = 1800
      Label1(0).Top = 840
      Label1(0).Caption = "Año y Mes"
      Combo1.Visible = True
      Me.Caption = Me.Caption & " Seguimientos de Patentes."
      Set mRec = mObjViol.oMesAnioEnvios
      Do While Not mRec.EOF
         Combo1.AddItem mRec!anio & " - " & MonthName(mRec!mes) & Space(20) & mRec!mes
         mRec.MoveNext
      Loop
      mRec.Close
End Select
End Sub

Private Sub sImprimir(ByVal pReport As String)
   CrystalReport1.DataFiles(0) = App.Path & "\Violaciones\Auxiliar.mdb"
   CrystalReport1.ReportFileName = App.Path & pReport
   CrystalReport1.WindowState = crptMaximized
   CrystalReport1.Action = 1
End Sub
