VERSION 5.00
Begin VB.Form RNov12 
   Caption         =   "Waze - Cerrar accidentes."
   ClientHeight    =   3900
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10815
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   10815
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   495
      Index           =   1
      Left            =   9240
      TabIndex        =   2
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar accidentes waze"
      Height          =   495
      Index           =   0
      Left            =   6360
      TabIndex        =   1
      Top             =   3240
      Width           =   2655
   End
   Begin VB.ListBox List1 
      Height          =   2310
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   120
      Width           =   10575
   End
End
Attribute VB_Name = "RNov12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Dim bol As Boolean
Dim mCodigos As String
   mCodigos = ""
   If Index = 0 Then 'Si cierro accidente
      Dim i As Integer
      For i = 0 To List1.ListCount - 1
         List1.ListIndex = i
         If List1.Selected(i) = True Then
            Dim mObj As New clRNov
            mObj.waze_cerrar_accidente Trim((Left(List1.Text, 7)))
            mCodigos = mCodigos & "'" & Trim((Left(List1.Text, 7))) & "',"
          
            Set mObj = Nothing
            '1) HACER EL UPDATE SET ESTADOWAZE = 3 AL ULTIMO WAZE DE ESE ALFANUMERICO CON ESTADO 2
            '2) CREAR ARCHIVO
            '3) SE VA A CREAR UN PROCESO EN EL SERVIDOR QUE MUEVA LOS ARCHIVO MAS VIEJOS DE 20 MINUTOS
         End If
         
      Next
      
      If mCodigos <> "" Then
         sDoXMLCierre Mid(mCodigos, 1, Len(mCodigos) - 1)
      End If
      
   End If
   RNov1a_frm.Enabled = True
   RNov1b_frm.Enabled = True
   RNov1d_frm.Enabled = True
   Unload Me
End Sub




Private Sub Form_Load()
   Dim mObj As New clRNov
   Dim mRec As New ADODB.Recordset
   Dim i As Integer
'   mPc = Mid(MDI.mPCname, 1, Len(MDI.mPCname) - 1)
   Me.Height = 4410
   Me.Width = 10935
   
   Me.Top = RNov1a_frm.Top + RNov1a_frm.Height + 30
   Me.Left = (MDI.Width - Me.Width) / 2
   
   Set mRec = mObj.waze_getAccidentesLiberados
   Do While Not mRec.EOF
      List1.AddItem mRec!descripcion
      mRec.MoveNext
   Loop
   For i = 0 To List1.ListCount - 1
      List1.Selected(i) = True
   Next
   mRec.Close
   Set mObj = Nothing
   Set mRec = Nothing
End Sub


Private Sub sDoXMLCierreOLD20181112(ByVal pCodAlfa As String)
   Dim mObj As New clRNov
   Dim mRec As New ADODB.Recordset
   Dim mFechaCracion As String

   Set mRec = mObj.oTabla("d_waze", " where codalfa='" & pCodAlfa & "' order by fecha desc ")
   If Not mRec.EOF Then
   
      mFechaCracion = mObj.sFechaMySQL
      Open App.Path & "\RegNovedades\waze\ausolAR_" & pCodAlfa & ".xml" For Output As #1
      Print #1, "<?xml version='1.0' encoding='UTF-8'?>"
      Print #1, "<incidents xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xsi:noNamespaceSchemaLocation='http://www.gstatic.com/road-incidents/incidents_feed.xsd'>"
      Print #1, "<incident id='" & pCodAlfa & "'>"
      Print #1, "<creationtime>" & Format(mFechaCracion, "yyyy-mm-dd") & "T" & Format(mFechaCracion, "HH:mm:ss") & "-03:00" & "</creationtime>"
      Print #1, "<type>ACCIDENT</type>"
      Print #1, "<subtype>" & mObj.sTablaDescr("s_subtiposincid", " codigo=" & mRec!codsubtipo, 3) & "</subtype>"
      Print #1, " <description>" & mRec!descripcion & "</description>"
      Print #1, "<location>"
      Print #1, "<polyline>" & mRec!lat & " " & mRec!lon & "</polyline>"
      Print #1, "</location>"
      Print #1, " <starttime>" & Format(mRec!fecha_i, "yyyy-mm-dd") & "T" & Format(mRec!fecha_i, "HH:mm:ss") & "-03:00" & "</starttime>"
      Print #1, " <endtime>" & Format(mRec!fecha_f, "yyyy-mm-dd") & "T" & Format(mRec!fecha_f, "HH:mm:ss") & "-03:00" & "</endtime>"
      Print #1, " <name>Autopistas del Sol S.A.</name>"
      Print #1, "</incident>"
      Print #1, "</incidents>"
   
      Close #1
         
      sEnviarFile App.Path & "\RegNovedades\waze\ausolAR_" & pCodAlfa & ".xml"
              
   End If
   mRec.Close
   
   'Unload Me
   Set mObj = Nothing
   Set mRec = Nothing
End Sub



Private Sub sDoXMLCierre(ByVal pCodAlfa As String)
   Dim mObj As New clRNov
   Dim mRec As New ADODB.Recordset

   Set mRec = mObj.oTabla("d_waze", " where codalfa in (" & pCodAlfa & ") order by fecha desc ")
   If Not mRec.EOF Then
      Open App.Path & "\RegNovedades\waze\ausolAR.xml" For Output As #1
      Print #1, "<?xml version=""1.0"" encoding=""UTF-8""?>"
      Print #1, "<incidents xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:noNamespaceSchemaLocation=""http://www.gstatic.com/road-incidents/incidents_feed.xsd"">"
      Do While Not mRec.EOF
         Print #1, "<incident id=""" & mRec!CodAlfa & """ > "
         Print #1, "  <creationtime>" & Format(mRec!fecha_i, "yyyy-mm-dd") & "T" & Format(mRec!fecha_i, "HH:mm:ss") & "-03:00" & "</creationtime>"
         Print #1, "  <updatetime>" & Format(Now, "yyyy-mm-dd") & "T" & Format(Now, "HH:mm:ss") & "-03:00" & "</updatetime>"
         Print #1, "  <description>" & mRec!descripcion & "</description>"
         Print #1, "  <location>"
         
         Print #1, "<street>" & mObj.getRamalWazeXml(mRec!codsent) & "</street>"
         
         Print #1, "     <direction>ONE_DIRECTION</direction>"
         Print #1, "     <polyline>" & mRec!lat & " " & mRec!lon & "</polyline>"
         Print #1, "  </location>"
         Print #1, " <reference>AUSOL</reference>"
         Print #1, " <starttime>" & Format(mRec!fecha_i, "yyyy-mm-dd") & "T" & Format(mRec!fecha_i, "HH:mm:ss") & "-03:00" & "</starttime>"
         Print #1, " <endtime>" & Format(Now, "yyyy-mm-dd") & "T" & Format(Now, "HH:mm:ss") & "-03:00" & "</endtime>"
         Print #1, " <type>ACCIDENT</type>"
         Print #1, " <subtype>" & mObj.sTablaDescr("s_subtiposincid", " codigo=" & mRec!codsubtipo, 3) & "</subtype>"
         Print #1, "</incident>"
                  
         mRec.MoveNext
      Loop
      mRec.Close
      
      Set mRec = mObj.oTabla("d_waze", " where codalfa not in (" & pCodAlfa & ") and estadowaze in (0, 1, 2) order by fecha desc ")
       Do While Not mRec.EOF
         Print #1, "<incident id=""" & mRec!CodAlfa & """ > "
         Print #1, "  <creationtime>" & Format(mRec!fecha_i, "yyyy-mm-dd") & "T" & Format(mRec!fecha_i, "HH:mm:ss") & "-03:00" & "</creationtime>"
'         Print #1, "  <updatetime>" & Format(Now, "yyyy-mm-dd") & "T" & Format(Now, "HH:mm:ss") & "-03:00" & "</updatetime>"
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
      Set mRec = mObj.oTabla("d_waze", " where estadowaze=3 and fecha_f > date_add(current_timestamp, interval -10 MINUTE)  and codalfa not in (" & pCodAlfa & ") order by fecha desc ")
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
   
   'Unload Me
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
   ''MsgBox "Datos enviado a Waze.", vbInformation, sMessage

End Sub






Private Sub sEnviarFileOLD(ByVal mArchivo As String)
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
    
    'MsgBox "Datos enviado a Waze.", vbInformation, sMessage
    
End Sub
