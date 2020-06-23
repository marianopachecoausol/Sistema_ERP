VERSION 5.00
Begin VB.Form RAcc14 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3375
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5475
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
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
   ScaleHeight     =   3375
   ScaleWidth      =   5475
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Datos migrados sist. anterior"
      Height          =   495
      Left            =   720
      TabIndex        =   6
      Top             =   1920
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
      Height          =   495
      Index           =   1
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Generar"
      Height          =   615
      Index           =   0
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   3120
      MaxLength       =   10
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fechas"
      Height          =   210
      Left            =   720
      TabIndex        =   1
      Top             =   1260
      Width           =   660
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Formulario para OCCOVI"
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
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   3195
   End
End
Attribute VB_Name = "RAcc14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   sAlinearForm Me
   sInitForm
End Sub

Private Sub Command1_Click(Index As Integer)
   Dim esMigracion As Boolean
   esMigracion = False
   
   If Index = 0 Then
      If sValidFechaDesdeHasta(Text1(0).Text, Text1(1).Text) = True Then
         If Check1 Then
            esMigracion = True
         End If
         sFormOCCOVI (esMigracion)
      Else
         MsgBox "Verificar fechas.", vbCritical, sMessage
      End If
   Else
      Unload Me
   End If
End Sub

'PROCESOS Y FUNCIONES
Private Sub sInitForm()
   'nada
End Sub

Private Sub sImportarExcel()
   Dim mDOS As New FileSystemObject
   Dim XLS As EXCEL.Application
   Dim mI As Integer
   Dim mJ As Integer
   Dim mFila As Integer
   Dim mArchivo As String
   
   Dim mObj As New clRNov
   Dim mObj2 As New clRAcc
      
   Dim Fecha As String
   Dim hora As String
   Dim xCodRamal As String
   Dim km As String
   Dim asc_desc As String
   Dim descTraza
   Dim codSentido As String
   Dim codClima As String
   Dim AccidOtro As String
   Dim AcciconOtro As String
   Dim CodColisContra1 As String
   Dim CodColisContra2 As String
   Dim cantHeridosLeves As Integer
   Dim cantHeridosGraves As Integer
   Dim cantMuertos As Integer
   Dim cantTotalVictimas As Integer
   Dim flagInsert As Boolean
   Dim NroOrden As Integer
   Dim strNroOrden As String

   Dim cantAutos As Integer
   Dim cantCamioneta As Integer
   Dim cantCamion As Integer
   Dim cantOmnibus As Integer
   Dim cantBicletas As Integer
   Dim cantMotos As Integer
   Dim cantOtros As Integer
   Dim cantTotalVehiculos As Integer


   mArchivo = "2017.xls"
   Set XLS = CreateObject("Excel.Application")
   With XLS
      .Application.WorkBooks.Open filename:=App.Path & "\RegAccidentes\Importacion para occovi\posta\" & mArchivo
      .Worksheets(1).Select
      .Worksheets(1).Name = "Hoja1"
      
      NroOrden = -4982
      
      
      mFila = 1
      Do While Trim(.Cells(mFila, 2)) <> ""
         strNroOrden = NroOrden
      
         Fecha = ""
         hora = ""
         xCodRamal = ""
         km = ""
         asc_desc = ""
         descTraza = ""
         codSentido = ""
         codClima = ""
         AccidOtro = ""
         AcciconOtro = ""
         CodColisContra1 = ""
         CodColisContra2 = ""
         cantHeridosLeves = 0
         cantHeridosGraves = 0
         cantMuertos = 0
         cantTotalVictimas = 0
         cantAutos = 0
         cantCamioneta = 0
         cantCamion = 0
         cantOmnibus = 0
         cantBicletas = 0
         cantMotos = 0
         cantOtros = 0
         cantTotalVehiculos = 0
      
         Fecha = Mid(.Cells(mFila, 2), 7, 4) & "-" & Mid(.Cells(mFila, 2), 4, 2) & "-" & Mid(.Cells(mFila, 2), 1, 2)
         hora = Mid(.Cells(mFila, 3), 1, 5) 'ver que pasa con hora con otra logitud de cadena
         xCodRamal = mObj.sTablaDescr("ramales", "abrevia='[" & .Cells(mFila, 4) & "]'", 0)
         km = Replace(.Cells(mFila, 5), ",", ".")
         codClima = mObj2.sTablaDescr("Clima", "Detalle='" & Trim(.Cells(mFila, 7)) & "' and fechabaja is null", 0)
         If codClima = "" Then
            codClima = "11" 'otro
         End If
      
         If Trim(.Cells(mFila, 8)) = "" Or Trim(.Cells(mFila, 8)) = "0" Then
            asc_desc = "Desc"
         Else
            asc_desc = "Asc"
         End If

       
         Select Case Trim(.Cells(mFila, 6))
            Case "P"
               descTraza = "CALZADA PRINCIPAL"
            Case "CP"
               descTraza = "COLECTORA PRINCIPAL"
            Case "CF"
               descTraza = "COLECTORA FRENTISTA"
         End Select

         codSentido = mObj.sTablaDescr("sentidos", "codramal=" & xCodRamal & _
            " and descripcion like '%" & descTraza & "%'" & _
            " and descripcion like '%" & asc_desc & "%'", 0)
       
        
         If Trim(.Cells(mFila, 10)) <> "" And Trim(.Cells(mFila, 10)) <> "0" Then
            AccidOtro = "01" 'Vuelco
         End If
                 
         If Trim(.Cells(mFila, 12)) <> "" And Trim(.Cells(mFila, 12)) <> "0" Then
            AcciconOtro = "02" ' Posterior
         End If

         If (Trim(.Cells(mFila, 11)) <> "" And Trim(.Cells(mFila, 11)) <> "0") And (Trim(.Cells(mFila, 12)) = "" Or Trim(.Cells(mFila, 12)) = "0") Then
            AcciconOtro = "01" 'Frontal
         End If
      
'         If Trim(.Cells(mFila, 11)) <> "" And Trim(.Cells(mFila, 11)) <> "0" Then
'            AcciconOtro = "01" ' Frontal
'         End If
'
'         If (Trim(.Cells(mFila, 12)) <> "" And Trim(.Cells(mFila, 12)) <> "0") And (Trim(.Cells(mFila, 11)) = "" Or Trim(.Cells(mFila, 11)) = "0") Then
'            AcciconOtro = "02" 'Posterior
'         End If
      
         If (Trim(.Cells(mFila, 13)) <> "" And Trim(.Cells(mFila, 13)) <> "0") And (Trim(.Cells(mFila, 11)) = "" Or Trim(.Cells(mFila, 11)) = "0") And (Trim(.Cells(mFila, 12)) = "" Or Trim(.Cells(mFila, 12)) = "0") Then
              AcciconOtro = "03" 'Diagonal
         End If
      
         If Trim(.Cells(mFila, 14)) <> "" And Trim(.Cells(mFila, 14)) <> "0" Then
            CodColisContra1 = "07" 'Animal
         End If
         
         If Trim(.Cells(mFila, 15)) <> "" And Trim(.Cells(mFila, 15)) <> "0" Then
            If Trim(.Cells(mFila, 14)) <> "" And Trim(.Cells(mFila, 14)) <> "0" Then
               CodColisContra2 = "12" 'Otros
            Else
               CodColisContra1 = "12" 'Otros
            End If
         End If
         'falta codifcar clima y ver que onda con codColisContra2
         flagInsert = mObj2.xInsertFichaMIGRACION(strNroOrden, Fecha, hora, "", km, "", AcciconOtro, CodColisContra1, CodColisContra2, AccidOtro, codSentido, "", codClima, "", "", "", _
         "", "", "", "", "", "", "", "", "", "", "", "", xCodRamal)
         
         
         If IsNumeric(Trim(.Cells(mFila, 16))) And Trim(.Cells(mFila, 16)) <> "" Then
            cantHeridosLeves = Trim(.Cells(mFila, 16))
         End If
      
         If IsNumeric(.Cells(mFila, 17)) And Trim(.Cells(mFila, 17)) <> "" Then
            cantHeridosGraves = Trim(.Cells(mFila, 17))
         End If
      
         If IsNumeric(Trim(.Cells(mFila, 18))) And Trim(.Cells(mFila, 18)) <> "" Then
            cantMuertos = Trim(.Cells(mFila, 18))
         End If
      
         cantTotalVictimas = cantHeridosLeves + cantHeridosGraves + cantMuertos
            
         'insercion en tabla VictimasInvolucr
         If cantTotalVictimas > 0 Then
            mI = 1
            For mJ = 1 To cantHeridosLeves
               flagInsert = mObj2.xInsertVictimasMIGRACION(strNroOrden, mI, "02")
               mI = mI + 1
            Next
            For mJ = 1 To cantHeridosGraves
               flagInsert = mObj2.xInsertVictimasMIGRACION(strNroOrden, mI, "03")
               mI = mI + 1
            Next
            For mJ = 1 To cantMuertos
               flagInsert = mObj2.xInsertVictimasMIGRACION(strNroOrden, mI, "04")
               mI = mI + 1
            Next
         End If
      
      
         If IsNumeric(Trim(.Cells(mFila, 19))) And Trim(.Cells(mFila, 19)) <> "" Then
            cantAutos = Trim(.Cells(mFila, 19))
         End If
     
         If IsNumeric(Trim(.Cells(mFila, 20))) And Trim(.Cells(mFila, 20)) <> "" Then
            cantCamioneta = Trim(.Cells(mFila, 20))
         End If
      
         If IsNumeric(Trim(.Cells(mFila, 21))) And Trim(.Cells(mFila, 21)) <> "" Then
            cantCamion = Trim(.Cells(mFila, 21))
         End If
      
         If IsNumeric(Trim(.Cells(mFila, 22))) And Trim(.Cells(mFila, 22)) <> "" Then
            cantOmnibus = Trim(.Cells(mFila, 22))
         End If
      
         If IsNumeric(Trim(.Cells(mFila, 23))) And Trim(.Cells(mFila, 23)) <> "" Then
            cantBicletas = Trim(.Cells(mFila, 23))
         End If
      
         If IsNumeric(Trim(.Cells(mFila, 24))) And Trim(.Cells(mFila, 24)) <> "" Then
            cantMotos = Trim(.Cells(mFila, 24))
         End If
      
         If IsNumeric(Trim(.Cells(mFila, 25))) And Trim(.Cells(mFila, 25)) <> "" Then
            cantOtros = Trim(.Cells(mFila, 25))
         End If
         
         cantTotalVehiculos = cantAutos + cantCamioneta + cantCamion + _
                            cantOmnibus + cantBicletas + cantMotos + cantOtros
      
         If cantTotalVehiculos > 0 Then
            mI = 1
            For mJ = 1 To cantAutos
               flagInsert = mObj2.xInsertVehiculosMIGRACION(strNroOrden, mI, "04")
               mI = mI + 1
            Next
            For mJ = 1 To cantCamioneta
               flagInsert = mObj2.xInsertVehiculosMIGRACION(strNroOrden, mI, "05")
               mI = mI + 1
            Next
            For mJ = 1 To cantCamion
               flagInsert = mObj2.xInsertVehiculosMIGRACION(strNroOrden, mI, "08")
               mI = mI + 1
            Next
             For mJ = 1 To cantOmnibus
               flagInsert = mObj2.xInsertVehiculosMIGRACION(strNroOrden, mI, "06")
               mI = mI + 1
            Next
            For mJ = 1 To cantBicletas
               flagInsert = mObj2.xInsertVehiculosMIGRACION(strNroOrden, mI, "01")
               mI = mI + 1
            Next
            For mJ = 1 To cantMotos
               flagInsert = mObj2.xInsertVehiculosMIGRACION(strNroOrden, mI, "03")
               mI = mI + 1
            Next
            For mJ = 1 To cantOtros
               flagInsert = mObj2.xInsertVehiculosMIGRACION(strNroOrden, mI, "16")
               mI = mI + 1
            Next
         End If
       
         NroOrden = NroOrden - 1
         mFila = mFila + 1
      Loop
   End With

   Set mObj = Nothing
   Set mObj2 = Nothing
      
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = fDateKeyPress(Text1(Index), KeyAscii)
End Sub

Private Sub sFormOCCOVI(ByVal pEsMigracion As Boolean)
'Private Sub sFormOCCOVIoriginal()
Dim mObj As New clRAcc
Dim mObjRN As New clRNov

Dim mRec As New ADODB.Recordset
Dim mRec2 As New ADODB.Recordset
Dim mDOS As New FileSystemObject
Dim XLS As EXCEL.Application
Dim mArchivo As String
Dim mI As Integer
Dim mT As Integer
Dim mTotalVic As Integer
Dim Fecha As String

   If pEsMigracion Then
      Set mRec = mObj.oTabla("FichaMIGRACION", " where fecha between '" & Format(Text1(0).Text, "yyyy-mm-dd") & " ' and '" & Format(Text1(1).Text, "yyyy-mm-dd") & "' order by fecha")
   Else
      Set mRec = mObj.oTabla("Ficha", " where fecha between '" & Format(Text1(0).Text, "yyyy-mm-dd") & " ' and '" & Format(Text1(1).Text, "yyyy-mm-dd") & "' order by fecha")
   End If
   
   If Not mRec.EOF Then
   
      sMsgEspere Me, "Procesando...", True
      mArchivo = "OCCOVIACC_" & Format(Now(), "yyyymmddhhmmss") & ".xls"
      mDOS.CopyFile App.Path & "\RegAccidentes\OccoviAcc.xls", App.Path & "\RegAccidentes\tmp\" & mArchivo, True
      Set XLS = CreateObject("Excel.Application")
      With XLS
      .Application.WorkBooks.Open filename:=App.Path & "\RegAccidentes\tmp\" & mArchivo
      .Worksheets(1).Select
      .Worksheets(1).Name = "Detalle"
      mI = 2
      Do While Not mRec.EOF
      
         '.Cells(mI, 1).Formula = "'" & mRec!nroorden
         'fijos
         .Cells(mI, 5).Formula = "'" & mObjRN.sTablaDescr("ramales", "codigo='" & mRec!codramal & "'", 3)
         
         .Cells(mI, 9).Formula = "Autopista"
         .Cells(mI, 11).Formula = "Peaje"
         .Cells(mI, 12).Formula = "OCRABA"
         .Cells(mI, 4).Formula = mObjRN.sTablaDescr("ramal_zonas", "codramal='" & mRec!codramal & "' and " & Replace(mRec!Progresiva, ",", ".") & " between km_i and km_f", 3)

         .Cells(mI, 6).Formula = mRec!Progresiva
         .Cells(mI, 7).Formula = mRec!Interseccion
         .Cells(mI, 8).Formula = Trim(Right(mObjRN.sTablaDescr("sentidos", "codigo='" & mRec!SentidoTrans & "'", 1), 4)) + "ENDENTE" 'SENTIDO TRANSITO
         
         .Cells(mI, 10).Formula = mObj.sTablaDescr("LugarAccid", "codlugaraccid='" & mRec!lugaraccid & "'", 1) 'LUGAR ACCIDENTE
         If .Cells(mI, 10).Formula = "" Then .Cells(mI, 10).Formula = "Calzada Principal"
         .Cells(mI, 13).Formula = "'" & mRec!Fecha
         .Cells(mI, 14).Formula = mRec!hora
         .Cells(mI, 15).Formula = "Habil"
         Dim mObjP As New clPeaje
         
         Fecha = Mid(mRec!Fecha, 7, 4) & "/" & Mid(mRec!Fecha, 4, 2) & "/" & Mid(mRec!Fecha, 1, 2)
         
         If mObjP.sEsFeriado(Fecha) = True Then
            .Cells(mI, 15).Formula = "Feriado"
         End If
         Set mObjP = Nothing
         If mRec!CodColisContra1 <> "" Then
            .Cells(mI, 16).Formula = "Colisión contra: " & mObj.sTablaDescr("ColisionContra", "CodColision='" & mRec!CodColisContra1 & "'", 1)
         End If
         If mRec!CodColisContra2 <> "" Then
            .Cells(mI, 16).Formula = .Cells(mI, 16).Formula & ", " & mObj.sTablaDescr("ColisionContra", "CodColision='" & mRec!CodColisContra2 & "'", 1)
         End If
         If mRec!CodColisContra3 <> "" Then
            .Cells(mI, 16).Formula = .Cells(mI, 16).Formula & ", " & mObj.sTablaDescr("ColisionContra", "CodColision='" & mRec!CodColisContra3 & "'", 1)
         End If

         If mRec!AcciconOtro <> "" Then
            .Cells(mI, 16).Formula = .Cells(mI, 16).Formula & " /Con otro vehículo: " & mObj.sTablaDescr("ConOtroVehic", "CodconOtro='" & mRec!AcciconOtro & "'", 1)
         End If
         
         If mRec!AccidOtro <> "" Then
            .Cells(mI, 16).Formula = .Cells(mI, 16).Formula & " /Otros: " & mObj.sTablaDescr("Otros", "CodOtros='" & mRec!AccidOtro & "'", 1)
         End If
         
         .Cells(mI, 37).Formula = ""
         If mRec!CodCausaCond1 <> "" Then
            .Cells(mI, 37).Formula = mObj.sTablaDescr("CausaConductor", "codcausacond='" & mRec!CodCausaCond1 & "'", 1)
         End If
         .Cells(mI, 38).Formula = "Se Ignora"
         If mRec!lugaraccid <> "" Then
            .Cells(mI, 38).Formula = mObj.sTablaDescr("LugarAccid", "CodLugarAccid='" & mRec!lugaraccid & "'", 1)
         End If
         .Cells(mI, 39).Formula = mObj.sTablaDescr("Clima", "Codclima='" & mRec!Clima1 & "'", 1)
         .Cells(mI, 40).Formula = mObj.sTablaDescr("EstCalzada", "codestcalzada='" & mRec!EstCalzada & "'", 1)
         .Cells(mI, 41).Formula = mObj.sTablaDescr("Iluminacion", "codilumin='" & mRec!Iluminac & "'", 1)
         
         'VICTIMAS
         mTotalVic = 0
         
         If pEsMigracion Then
            mTotalVic = mObj.iTotalHeridosNroOrdenMIGRACION(mRec!NroOrden)
            mT = mObj.iTotalHerLevesNroOrdenMIGRACION(mRec!NroOrden)
         Else
            mTotalVic = mObj.iTotalHeridosNroOrden(mRec!NroOrden)
            mT = mObj.iTotalHerLevesNroOrden(mRec!NroOrden)
         End If
         .Cells(mI, 19).Formula = mTotalVic - mT
         .Cells(mI, 20).Formula = mT
         
         If pEsMigracion Then
            mT = mObj.iTotalMuertosNroOrdenMIGRACION(mRec!NroOrden)
         Else
            mT = mObj.iTotalMuertosNroOrden(mRec!NroOrden)
         End If
         .Cells(mI, 18).Formula = mT
         mTotalVic = mT + mTotalVic
         .Cells(mI, 17).Formula = mTotalVic
         
         'VEHICULOS
         mTotalVic = 0
         If pEsMigracion Then
            Set mRec2 = mObj.oTotVehTipoNroOrdenMIGRACION(mRec!NroOrden)
         Else
            Set mRec2 = mObj.oTotVehTipoNroOrden(mRec!NroOrden)
         End If
         If Not mRec2.EOF Then
            Do While Not mRec2.EOF
               mTotalVic = mTotalVic + mRec2!Total
               Select Case mRec2!CodTipoVehic
                  Case "01"
                     .Cells(mI, 22).Formula = mRec2!Total
                  Case "03"
                     .Cells(mI, 24).Formula = mRec2!Total
                  Case "04", "15"
                     If .Cells(mI, 25).Formula <> "" Then
                        .Cells(mI, 25).Formula = CInt(.Cells(mI, 25).Formula) + mRec2!Total
                     Else
                        .Cells(mI, 25).Formula = mRec2!Total
                     End If
                  Case "05"
                     .Cells(mI, 26).Formula = mRec2!Total
                  Case "06"
                     .Cells(mI, 27).Formula = mRec2!Total
                  Case "07"
                     .Cells(mI, 28).Formula = mRec2!Total
                  Case "08"
                     .Cells(mI, 29).Formula = mRec2!Total
                  Case "09"
                     .Cells(mI, 30).Formula = mRec2!Total
                  Case "10"
                     .Cells(mI, 31).Formula = mRec2!Total
                  Case "11"
                     .Cells(mI, 32).Formula = mRec2!Total
                  Case "16"""
                     .Cells(mI, 35).Formula = mRec2!Total
               End Select
               mRec2.MoveNext
            Loop
         End If
         mRec2.Close
         .Cells(mI, 21).Formula = mTotalVic

         mI = mI + 1
         mRec.MoveNext
      Loop
            
      .Visible = True
      End With
      sMsgEspere Me, "", False
   End If
   mRec.Close
   
   Set mObjRN = Nothing
   Set mObj = Nothing
   Set mRec = Nothing
   Set mRec2 = Nothing
   
End Sub


'Private Sub sFormOCCOVI()
Private Sub sFormOCCOVInuevo()
Dim mObj As New clRAcc
Dim mObjRN As New clRNov

Dim mRec As New ADODB.Recordset
Dim mRec2 As New ADODB.Recordset
Dim mDOS As New FileSystemObject
Dim XLS As EXCEL.Application
Dim mArchivo As String
Dim mI As Integer
Dim mT As Integer
Dim mTotalVic As Integer
Dim Fecha As String

   Set mRec = mObj.oTabla("FichaMIGRACION", " where fecha between '" & Format(Text1(0).Text, "yyyy-mm-dd") & " ' and '" & Format(Text1(1).Text, "yyyy-mm-dd") & "' order by fecha")
   If Not mRec.EOF Then
   
      sMsgEspere Me, "Procesando...", True
      mArchivo = "OCCOVIACC_" & Format(Now(), "yyyymmddhhmmss") & ".xls"
      mDOS.CopyFile App.Path & "\RegAccidentes\OccoviAcc.xls", App.Path & "\RegAccidentes\tmp\" & mArchivo, True
      Set XLS = CreateObject("Excel.Application")
      With XLS
      .Application.WorkBooks.Open filename:=App.Path & "\RegAccidentes\tmp\" & mArchivo
      .Worksheets(1).Select
      .Worksheets(1).Name = "Detalle"
      mI = 2
      Do While Not mRec.EOF
      
         '.Cells(mI, 1).Formula = "'" & mRec!nroorden
         'fijos
         .Cells(mI, 5).Formula = "'" & mObjRN.sTablaDescr("ramales", "codigo='" & mRec!codramal & "'", 3)
         
         .Cells(mI, 9).Formula = "Autopista"
         .Cells(mI, 11).Formula = "Peaje"
         .Cells(mI, 12).Formula = "OCRABA"
         .Cells(mI, 4).Formula = mObjRN.sTablaDescr("ramal_zonas", "codramal='" & mRec!codramal & "' and " & Replace(mRec!Progresiva, ",", ".") & " between km_i and km_f", 3)

         .Cells(mI, 6).Formula = mRec!Progresiva
         .Cells(mI, 7).Formula = mRec!Interseccion
         .Cells(mI, 8).Formula = Trim(Right(mObjRN.sTablaDescr("sentidos", "codigo='" & mRec!SentidoTrans & "'", 1), 4)) + "ENDENTE" 'SENTIDO TRANSITO
         .Cells(mI, 10).Formula = mObj.sTablaDescr("LugarAccid", "codlugaraccid='" & mRec!lugaraccid & "'", 1) 'LUGAR ACCIDENTE
         If .Cells(mI, 10).Formula = "" Then .Cells(mI, 10).Formula = "Calzada Principal"
         .Cells(mI, 13).Formula = "'" & mRec!Fecha
         .Cells(mI, 14).Formula = mRec!hora
         .Cells(mI, 15).Formula = "Habil"
         Dim mObjP As New clPeaje
         
         Fecha = Mid(mRec!Fecha, 7, 4) & "/" & Mid(mRec!Fecha, 4, 2) & "/" & Mid(mRec!Fecha, 1, 2)
         
         If mObjP.sEsFeriado(Fecha) = True Then
            .Cells(mI, 15).Formula = "Feriado"
         End If
         Set mObjP = Nothing
         If mRec!CodColisContra1 <> "" Then
            .Cells(mI, 16).Formula = "Colisión contra: " & mObj.sTablaDescr("ColisionContra", "CodColision='" & mRec!CodColisContra1 & "'", 1)
         End If
         If mRec!CodColisContra2 <> "" Then
            .Cells(mI, 16).Formula = .Cells(mI, 16).Formula & ", " & mObj.sTablaDescr("ColisionContra", "CodColision='" & mRec!CodColisContra2 & "'", 1)
         End If
         If mRec!CodColisContra3 <> "" Then
            .Cells(mI, 16).Formula = .Cells(mI, 16).Formula & ", " & mObj.sTablaDescr("ColisionContra", "CodColision='" & mRec!CodColisContra3 & "'", 1)
         End If

         If mRec!AcciconOtro <> "" Then
            .Cells(mI, 16).Formula = .Cells(mI, 16).Formula & " /Con otro vehículo: " & mObj.sTablaDescr("ConOtroVehic", "CodconOtro='" & mRec!AcciconOtro & "'", 1)
         End If
         
         If mRec!AccidOtro <> "" Then
            .Cells(mI, 16).Formula = .Cells(mI, 16).Formula & " /Otros: " & mObj.sTablaDescr("Otros", "CodOtros='" & mRec!AccidOtro & "'", 1)
         End If
         
         .Cells(mI, 37).Formula = ""
         If mRec!CodCausaCond1 <> "" Then
            .Cells(mI, 37).Formula = mObj.sTablaDescr("CausaConductor", "codcausacond='" & mRec!CodCausaCond1 & "'", 1)
         End If
         .Cells(mI, 38).Formula = "Se Ignora"
         If mRec!lugaraccid <> "" Then
            .Cells(mI, 38).Formula = mObj.sTablaDescr("LugarAccid", "CodLugarAccid='" & mRec!lugaraccid & "'", 1)
         End If
         .Cells(mI, 39).Formula = mObj.sTablaDescr("Clima", "Codclima='" & mRec!Clima1 & "'", 1)
         .Cells(mI, 40).Formula = mObj.sTablaDescr("EstCalzada", "codestcalzada='" & mRec!EstCalzada & "'", 1)
         .Cells(mI, 41).Formula = mObj.sTablaDescr("Iluminacion", "codilumin='" & mRec!Iluminac & "'", 1)
         
         'VICTIMAS
         mTotalVic = 0
         
         mTotalVic = mObj.iTotalHeridosNroOrdenMIGRACION(mRec!NroOrden)
         mT = mObj.iTotalHerLevesNroOrdenMIGRACION(mRec!NroOrden)
         .Cells(mI, 19).Formula = mTotalVic - mT
         .Cells(mI, 20).Formula = mT
         
         mT = mObj.iTotalMuertosNroOrdenMIGRACION(mRec!NroOrden)
         .Cells(mI, 18).Formula = mT
         mTotalVic = mT + mTotalVic
         .Cells(mI, 17).Formula = mTotalVic
         
         'VEHICULOS
         mTotalVic = 0
         Set mRec2 = mObj.oTotVehTipoNroOrdenMIGRACION(mRec!NroOrden)
         If Not mRec2.EOF Then
            Do While Not mRec2.EOF
               mTotalVic = mTotalVic + mRec2!Total
               Select Case mRec2!CodTipoVehic
                  Case "01"
                     .Cells(mI, 22).Formula = mRec2!Total
                  Case "03"
                     .Cells(mI, 24).Formula = mRec2!Total
                  Case "04", "15"
                     If .Cells(mI, 25).Formula <> "" Then
                        .Cells(mI, 25).Formula = CInt(.Cells(mI, 25).Formula) + mRec2!Total
                     Else
                        .Cells(mI, 25).Formula = mRec2!Total
                     End If
                  Case "05"
                     .Cells(mI, 26).Formula = mRec2!Total
                  Case "06"
                     .Cells(mI, 27).Formula = mRec2!Total
                  Case "07"
                     .Cells(mI, 28).Formula = mRec2!Total
                  Case "08"
                     .Cells(mI, 29).Formula = mRec2!Total
                  Case "09"
                     .Cells(mI, 30).Formula = mRec2!Total
                  Case "10"
                     .Cells(mI, 31).Formula = mRec2!Total
                  Case "11"
                     .Cells(mI, 32).Formula = mRec2!Total
                  Case "16"""
                     .Cells(mI, 35).Formula = mRec2!Total
               End Select
               mRec2.MoveNext
            Loop
         End If
         mRec2.Close
         .Cells(mI, 21).Formula = mTotalVic
         
'         .Cells(mI, 5).Formula = mRec!progresiva
'         .Cells(mI, 5).Formula = mRec!progresiva
'         .Cells(mI, 5).Formula = mRec!progresiva
'         .Cells(mI, 5).Formula = mRec!progresiva
'         .Cells(mI, 5).Formula = mRec!progresiva
'
         mI = mI + 1
         mRec.MoveNext
      Loop
            
      .Visible = True
      End With
      sMsgEspere Me, "", False
   End If
   mRec.Close
   
   Set mObjRN = Nothing
   Set mObj = Nothing
   Set mRec = Nothing
   
End Sub

