Attribute VB_Name = "Module1"
Public Const sMessage = "Sistema Global - Versión 202006231620"
Public Const mIPServer = "desa-ssvv.ausol.corp"
'Public Const mIPServer = "10.128.1.209"
'Public Const mIPServer = "127.0.0.1"
'Public Const mIPServer = "192.168.2.155"  'IP Urzagasti para pruebas con MYSQL Local en otras PCs
'Public Const mIPServer = "10.10.20.11"  'IP Urzagasti para pruebas con MYSQL Local en otras PCs
Public Const mVidFServer = "\\gcopublic\Publico\Sistema Global\VIDEO-AUDITORIA\"

'ESTO ES PARA EL SONIDO DE AVISO EN EL REGISTRO DE NOVEDADES
Declare Function sndplaysound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Const SND_ASYNC = (1)
Const SND_NODEFAULT = (2)
Public mArchivoETR As String

Public O_Registro As WshShell

Public Const LocalMachine As String = "HKEY_LOCAL_MACHINE"
Public Const CurrentUser As String = "HKEY_CURRENT_USER"
Public Function Leer_Dato(Principal As String, Valor As String) As String
    Set O_Registro = New WshShell
    Leer_Dato = O_Registro.RegRead(Principal & "\Control Panel\International\" & Valor)
    Set O_Registro = Nothing
End Function


Public Sub PlaySound(ByVal mFileWav)
Dim rc
   rc = sndplaysound(App.Path & "\RegNovedades\Sonido\" & mFileWav, SND_NODEFAULT + SND_ASYNC)
End Sub

Public Function ShowMenu(mMenu As Integer, mView As Boolean, mCheckMin As Boolean)
Dim mCont As Integer
   If mView Then
      MDI.ERP_VenWind(mMenu).Visible = True
      MDI.ERP_VenWind(mMenu).Checked = True
      MDI.ERP_Vent.Visible = True
      MDI.ERP_Vacio.Visible = True
      MDI.ERP_VenMini.Visible = True
      MDI.ERP_VenMini.Enabled = True
      MDI.mMenuActivo = mMenu
      MDI.mFormActive mMenu
   End If
   If mCheckMin Then
      For mCont = 0 To MDI.ERP_VenWind.UBound
         MDI.ERP_VenWind(mCont).Checked = False
      Next
   End If
   '***************************************************************************
   ' Cuando se agregue un nuevo sistema que contenga un menu en MDI, se deberán
   ' agregar en el case las cabeceras del MENÚ
   '***************************************************************************
   Select Case mMenu
      Case 1 'Registro de Novedades
         MDI.RNov_Arch.Visible = mView
         MDI.RNov_Nove.Visible = mView
         MDI.RNov_Repo.Visible = mView
         MDI.RNov_Exit.Visible = mView
         MDI.Caption = Left(MDI.Caption, 40) & Space(10) & " -   Sistema de Registro de Novedades"
      Case 2 'Registro de Accidentología
         MDI.RAcc_Arch.Visible = mView
         MDI.RAcc_Fich.Visible = mView
         MDI.RAcc_Info.Visible = mView
         MDI.RAcc_Exit.Visible = mView
         MDI.Caption = Left(MDI.Caption, 40) & Space(10) & " -   Sistema de Accidentología"
      Case 9 'Cambio de Clave
         ERP4_frm.Visible = mView
         MDI.Caption = Left(MDI.Caption, 40) & Space(10) & " -   Cambio de Clave"
      Case 10 'Asignación de Permisos
         ERP2_frm.Visible = mView
         MDI.Caption = Left(MDI.Caption, 40) & Space(10) & " -   Asignación de Permisos"
      Case 11 'Alta de Usuarios
         ERP5_frm.Visible = mView
         MDI.Caption = Left(MDI.Caption, 40) & Space(10) & " -   Alta de Usuarios"
     
      Case 12 'Sistema Inventario
      
         MDI.Inven_Arch.Visible = mView
         MDI.Inven_Movi.Visible = mView
         MDI.Inven_Repo.Visible = mView
         MDI.Inven_Exit.Visible = mView
         MDI.Caption = Left(MDI.Caption, 40) & Space(10) & " -   Sistema de Gestión de Inventarios"
     
     Case 13 'Contador PEEK
         MDI.CPEEK_IDat.Visible = mView
         MDI.CPEEK_Info.Visible = mView
         MDI.CPEEK_Exit.Visible = mView
         MDI.Caption = Left(MDI.Caption, 40) & Space(10) & " -   Sistema Consultas de Contadores PEEKs"
      Case 20 'Sistema Mant. Edilicio
         MDI.MEdReg.Visible = mView
         MDI.MEdRep.Visible = mView
         MDI.MEdSal.Visible = mView
         MDI.Caption = Left(MDI.Caption, 40) & Space(10) & " -   Sistema de Gestión de Mantenimiento Edilicio"
      Case 47 'Sistema Mant. Eléctrico
         MDI.MElecReg.Visible = mView
         MDI.MElecRep.Visible = mView
         MDI.MElecSal.Visible = mView
         MDI.Caption = Left(MDI.Caption, 40) & Space(10) & " -   Sistema de Gestión de Mantenimiento Electrico"
   End Select
End Function

Public Function NVL(vntVariable As Variant, vntValor) As Variant
   NVL = IIf(IsNull(vntVariable), vntValor, vntVariable)
End Function

Public Sub sLlenoCbo(ByRef pObjCbo As Object, ByRef pRec As ADODB.Recordset, ByVal pCampoDescr As Integer, ByVal pCampoCod As Integer)
   Do While Not pRec.EOF
      pObjCbo.AddItem pRec.Fields(pCampoDescr) & Space(60) & pRec.Fields(pCampoCod)
      pRec.MoveNext
   Loop
   pRec.Close
End Sub
   
Public Sub sSelectCboRigth(ByRef pObjCbo As Object, ByVal pCompara As String)
Dim mi As Integer
   For mi = 0 To pObjCbo.ListCount - 1
      If Trim(Right(pObjCbo.List(mi), 3)) = pCompara Then
         pObjCbo.ListIndex = mi
      End If
   Next
End Sub
   
Public Sub sSelectCboLeft(ByRef pObjCbo As Object, ByVal pCompara As String)
Dim mi As Integer
   For mi = 0 To pObjCbo.ListCount - 1
      If Trim(Left(pObjCbo.List(mi), 3)) = pCompara Then
         pObjCbo.ListIndex = mi
      End If
   Next
End Sub

Public Function Progresiva_Ok(mParam As String, ByVal pCodSent As String) As Boolean
Dim mObj As New clRNov
Dim mRet As Boolean
Dim mi As Integer
Dim mj As Double
Dim mIni As Double
Dim mFin As Double
Dim mPunto As String

   mPunto = Leer_Dato(CurrentUser, "sDecimal")
     
  ' mParam = Replace(mParam, ".", ",") 'mp 20160309
   If mPunto = "," Then mParam = Replace(mParam, ".", ",")      'mp 20160309
   
   mRet = True
   If mParam <> "" And pCodSent <> "" Then
      mObj.dKmIniFin pCodSent, mIni, mFin
      mj = mParam
      If mj < mIni Or mj > mFin Then
         MsgBox "Progresiva Incorrecta. Fuera de Límites", vbCritical, sMessage
         mRet = False
      End If
   Else
      MsgBox "Verificar Progresiva o Sentido.", vbCritical, sMessage
      mRet = False
   End If
   Progresiva_Ok = mRet
End Function

Public Function Fecha_ok(pParam As String) As Boolean
Dim mRet As Boolean
Dim Dia As String
Dim mes As String
Dim Ano As String
   Dia = Mid(pParam, 1, 2) 'Asigna valor Dia
   mes = Mid(pParam, 4, 2) 'Asigna valor Mes
   Ano = Mid(pParam, 7, 4)  'Asigna valor Año
   mRet = True
   If pParam <> "" Then 'Valida que el textbox no este vacio
      If Len(pParam) = 10 Then 'Valida la longitud de la fecha
         If Mid(pParam, 3, 1) = "/" And (Mid(pParam, 6, 1) = "/") Then 'Valida formato de la fecha
            If mes > 12 Or mes < 1 Then
               MsgBox "Mes Inválido", vbCritical, "Atención!!! - " & sMessage
               mRet = False
            Else
               If Ano < 1920 Or Ano > 2030 Then
                  MsgBox "Año Inválido", vbCritical, "Atención!!! - " & sMessage
                  mRet = False
               Else
                  If mes = 4 Or mes = 6 Or mes = 9 Or mes = 11 Then
                     If Dia > 30 Then
                        MsgBox "Día Inválido", vbCritical, "Atención!!! - " & sMessage
                        mRet = False
                     End If
                  Else
                     If mes = 2 Then
                        If Ano Mod 4 = 0 Then
                           If Dia > 29 Then
                              MsgBox "Día Inválido", vbCritical, "Atención!!! - " & sMessage
                              mRet = False
                           End If
                        Else
                           If Dia > 28 Then
                              MsgBox "Día Inválido", vbCritical, "Atención!!! - " & sMessage
                              mRet = False
                           End If
                        End If
                     Else
                        If Dia > 31 Then
                            MsgBox "Día Inválido", vbCritical, "Atención!!! - " & sMessage
                            mRet = False
                        End If
                     End If
                  End If
               End If
            End If
         Else
            MsgBox "Formato Incorrecto. Ingrese DD/MM/AAAA", vbCritical, "Atención!!! - " & sMessage
            mRet = False
         End If
      Else
         MsgBox "Formato Incorrecto. Ingrese DD/MM/AAAA", vbCritical, "Atención!!! - " & sMessage
         mRet = False
      End If
   Else
      MsgBox "Debe Ingresar Fecha. Formato DD/MM/AAAA", vbCritical, "Atención!!! - " & sMessage
      mRet = False
   End If
   Fecha_ok = mRet
End Function

Public Function Hora_ok(pParam As String) As Boolean
Dim mRet As Boolean
Dim Hor As String
Dim min As String
   Hor = Mid(pParam, 1, 2) 'Asigna valor Hora
   min = Mid(pParam, 4, 2) 'Asigna valor Minutos
   mRet = True
   If pParam <> "" Then 'Valida que el textbox no este vacio
      If Len(pParam) = 5 Then  'Valida la longitud de la hora
         If Mid(pParam, 3, 1) = ":" Then  'Valida formato de la hora
            If Hor < 0 Or Hor > 23 Then
               MsgBox "Hora Inválida", vbCritical, "Atención!!! - " & sMessage
               mRet = False
            Else
               If min < 0 Or min > 59 Then
                  MsgBox "Minuto Inválido", vbCritical, "Atención!!! - " & sMessage
                  mRet = False
               End If
            End If
         Else
            MsgBox "Formato Incorrecto. Ingrese HH:MM", vbCritical, "Atención!!! - " & sMessage
            mRet = False
         End If
      Else
         MsgBox "Formato Incorrecto. Ingrese HH:MM", vbCritical, "Atención!!! - " & sMessage
         mRet = False
      End If
   Else
      MsgBox "Debe Ingresar Hora. Formato HH:MM", vbCritical, "Atención!!! - " & sMessage
      mRet = False
   End If
   Hora_ok = mRet
End Function

Public Function Hora_ok2(mObj1 As Object, mObj2 As Object) As Boolean
Dim Flag As Boolean
Dim Hora1, Hora2 As Date
   If mObj1.Text <> "" And mObj2.Text <> "" Then
      If Hora_ok(mObj1.Text) And Hora_ok(mObj2.Text) Then
         Hora1 = mObj1.Text
         Hora2 = mObj2.Text
         If Hora1 <= Hora2 Then
            Flag = True
         Else
            Flag = False
            MsgBox "Hora Inicial Mayor a la Final", vbCritical, "Atención!!! - " & sMessage
         End If
      Else
         Flag = False
      End If
   Else
      If mObj1.Text <> "" Or mObj2.Text <> "" Then
         MsgBox "Falta Completar una Hora", vbCritical, "Atención!!! - " & sMessage
         Flag = False
      Else
         Flag = True
      End If
   End If
   Hora_ok2 = Flag
End Function

Public Function mExitSist(mCodSist As Integer)
Dim mi As Integer
   MDI.Caption = Left(MDI.Caption, 35)
   MDI.ERP_VenWind(mCodSist).Checked = False
   MDI.ERP_VenWind(mCodSist).Visible = False
   MDI.mMenuActivo = 999
   For mi = 0 To MDI.ERP_VenWind.UBound
      If MDI.ERP_VenWind(mi).Visible Then
         mi = 998
      End If
   Next
   If mi = 999 Then
      MDI.ERP_VenMini.Visible = True
      MDI.ERP_Vent.Visible = True
      MDI.ERP_Vacio.Visible = True
   Else
      MDI.ERP_VenMini.Visible = False
      MDI.ERP_Vent.Visible = False
      MDI.ERP_Vacio.Visible = False
   End If
   MDI.ERP_VenMini.Enabled = False
End Function

Public Function ValReal(pParam As String) As Double
Dim mRet As Double
Dim mStr As String
Dim mi As Integer
Dim mRead As String
   mStr = ""
   For mi = 1 To Len(pParam)
      mRead = Mid(pParam, mi, 1)
      If Asc(mRead) = 45 Or Asc(mRead) = 46 Or (Asc(mRead) >= 48 And Asc(mRead) <= 57) Then
         mStr = mStr & mRead
      End If
   Next
   mRet = Val(mStr)
   ValReal = mRet
   End Function
   
   Public Function MonthName(ByVal pMes As Integer) As String
   Select Case pMes
      Case 1
         MonthName = "Enero"
      Case 2
         MonthName = "Febrero"
      Case 3
         MonthName = "Marzo"
      Case 4
         MonthName = "Abril"
      Case 5
         MonthName = "Mayo"
      Case 6
         MonthName = "Junio"
      Case 7
         MonthName = "Julio"
      Case 8
         MonthName = "Agosto"
      Case 9
         MonthName = "Septiembre"
      Case 10
         MonthName = "Octubre"
      Case 11
         MonthName = "Noviembre"
      Case 12
         MonthName = "Diciembre"
      Case Else
         MonthName = ""
   End Select
End Function

Public Function Sleep(ByVal lngSegundos)
   If lngSegundos = "" Then lngSegundos = 0
   If IsNull(lngSegundos) Then lngSegundos = 0
   If Not IsNumeric(lngSegundos) Then lngSegundos = 0
   If lngSegundos < 0 Then lngSegundos = 0
   Ahora = Now()
   Do While Now() < (Ahora + (lngSegundos / 86400))
   Loop
End Function
 
Public Function fNewCodAlfa() As String
Dim mObj As New clRNov
Dim mRec As New ADODB.Recordset
Dim mReturn As String
Dim mi As Integer
Dim mj As Integer
   Set mRec = mObj.oTabla("codigos", "where estado='1' limit 15")
   mi = mRec.RecordCount
   Randomize Len(MDI.mUser) + Second(Now)
   mj = Int((mi * Rnd) + 1)
   For mi = 1 To mj
      mRec.MoveNext
   Next
   mReturn = mRec.Fields(0)
   mObj.xUpCodigos mReturn, "0"
   fNewCodAlfa = mReturn
   Set mObj = Nothing
End Function

Public Function fInitRNov1a_frm()
   RNov1a_frm.Enabled = True
   RNov1a_frm.Command2(0).Visible = True
   RNov1a_frm.List1.Clear
   RNov1a_frm.List1.Enabled = True
   RNov1a_frm.List1.Visible = False
   RNov1a_frm.Label3.Visible = False
End Function

Public Function ClimaOK(ByVal mKmts As Double) As String
   Select Case mKmts
      Case Is <= 21.62
         ClimaOK = "0"
      Case Is <= 35.84
         ClimaOK = "1"
      Case Is <= 38.57
         ClimaOK = "2"
      Case Is <= 47.66
         ClimaOK = "3"
      Case Is <= 63.3
         ClimaOK = "4"
      Case Is <= 65.14
         ClimaOK = "5"
   End Select
End Function

Public Function fPasar_Km_Sent(pKm As String, pSent As String, ByVal pCodRamal As String, ByVal pCodReferencia As String)
Dim mi As Integer
   
   For mi = 0 To RNov1c_frm.Combo1(4).ListCount - 1
      If Left(Right(RNov1c_frm.Combo1(4).List(mi), 6), 2) = pCodRamal Then
         RNov1c_frm.Combo1(4).ListIndex = mi
         mi = 999
      End If
   Next
   
   For mi = 0 To RNov1c_frm.Combo1(5).ListCount - 1
      If Trim(Right(RNov1c_frm.Combo1(5).List(mi), 3)) = pCodReferencia Then
         RNov1c_frm.Combo1(5).ListIndex = mi
         mi = 999
      End If
   Next
   
   For mi = 0 To RNov1c_frm.Combo1(0).ListCount - 1
      If Left(RNov1c_frm.Combo1(0).List(mi), 2) = pSent Then
         RNov1c_frm.Combo1(0).ListIndex = mi
         mi = 999
      End If
   Next
   
   RNov1c_frm.Text1(1).Text = Format(pKm, "00.00")
End Function

Public Sub sObtOrigen(ByVal pMovil, ByVal pCodAlfa As String, pObjCmb As Object)
Dim mObj As New clRNov
Dim mRec As New ADODB.Recordset
Dim pOrigen As String
Dim mi As Integer
   pOrigen = "SSV"
   If pCodAlfa = "" Then
      If Left(pMovil, 1) <> "M" Then
         pOrigen = "ACA"
      End If
   Else
      If pMovil = "" Then
         Set mRec = mObj.oObtOrigen("Codigo='" & pCodAlfa & "'")
      Else
         Set mRec = mObj.oObtOrigen("(Mov1='" & pMovil & "' OR Mov2='" & pMovil & "' OR Mov3='" & pMovil & "') AND Codigo='" & pCodAlfa & "'")
         If Left(pMovil, 1) <> "M" Then
            pOrigen = "ACA"
         End If
      End If
      If Not mRec.EOF Then
         pOrigen = mRec!CodOrigen
      End If
      mRec.Close
   End If
   For mi = 0 To pObjCmb.ListCount - 1
      If Left(pObjCmb.List(mi), 3) = pOrigen Then
         pObjCmb.ListIndex = mi
         mi = 999
      End If
   Next
   Set mRec = Nothing
   Set mObj = Nothing
End Sub

Public Sub sAlinearForm(pObj As Object)
   pObj.Left = (MDI.Width - pObj.Width) / 2
   pObj.Top = (MDI.Height - pObj.Height) / 3
End Sub

Public Function sValidFechaDesdeHasta(ByVal pFdesde As String, ByVal pFhasta As String) As Boolean
   sValidFechaDesdeHasta = False
   If Fecha_ok(pFdesde) And Fecha_ok(pFhasta) Then
      If CDate(pFdesde) > CDate(pFhasta) Then
         MsgBox "El rango de fechas es incorrecto.", vbExclamation, sMessage
      Else
         sValidFechaDesdeHasta = True
      End If
   End If
End Function

Function Replace(ByVal Expresion As String, ByVal Encontrar As String, ByVal ReemplazarCon As String) As String
Dim mi As Long
Dim mj As Long
   mj = 1
   Do
      mi = InStr(mj, Expresion, Encontrar)
      If mi Then
         Expresion = Left$(Expresion, mi - 1) & ReemplazarCon & Mid$(Expresion, mi + Len(Encontrar))
         mj = mi + 1
      End If
   Loop While mi
   Replace = Expresion
End Function

Function fDateKeyPress(ByVal pObjText As Object, pTecla As Integer) As Integer
   If Not (pTecla >= 47 And pTecla <= 57) And pTecla <> 8 Then
      fDateKeyPress = 0
   Else
      If (Len(pObjText.Text) = 2 Or Len(pObjText.Text) = 5) And pTecla <> 8 Then
         If pTecla <> 47 Then
            pObjText.Text = pObjText.Text + "/"
         End If
         pObjText.SelStart = Len(pObjText.Text) + 1
      End If
      fDateKeyPress = pTecla
   End If
End Function

Function fHoraKeyPress(ByVal pObjText As Object, pTecla As Integer) As Integer
   If Not (pTecla >= 48 And pTecla <= 58) And pTecla <> 8 Then
      fHoraKeyPress = 0
   Else
      If (Len(pObjText.Text) = 2) And pTecla <> 8 Then
         pObjText.Text = pObjText.Text + ":"
         pObjText.SelStart = Len(pObjText.Text) + 1
      End If
      fHoraKeyPress = pTecla
   End If
End Function

Function fNumeroKeyPress(pTecla As Integer) As Integer
   If Not (pTecla >= 48 And pTecla <= 57) And pTecla <> 8 And pTecla <> 13 Then
      fNumeroKeyPress = 0
   Else
      fNumeroKeyPress = pTecla
   End If
End Function

Function fNumDoubleKeyPress(pTecla As Integer) As Integer
   If Not ((pTecla >= 48 And pTecla <= 57) Or pTecla = 46) And pTecla <> 8 And pTecla <> 13 Then
      fNumDoubleKeyPress = 0
   Else
      fNumDoubleKeyPress = pTecla
   End If
End Function

Function fKmsKeyPress(ByVal pObjText As Object, pTecla As Integer) As Integer
   If Not (pTecla >= 47 And pTecla <= 57) And pTecla <> 8 And pTecla <> 13 Then
      fKmsKeyPress = 0
   Else
      If (Len(pObjText.Text) = 2) And pTecla <> 8 Then
         pObjText.Text = pObjText.Text + "."
         pObjText.SelStart = Len(pObjText.Text) + 1
      End If
      fKmsKeyPress = pTecla
   End If
End Function

Function fAlfaNumSimbKeyPress(pTecla As Integer) As Integer
   If Not ((pTecla >= 44 And pTecla <= 57) Or (pTecla >= 97 And pTecla <= 122) Or (pTecla >= 65 And pTecla <= 90) Or pTecla = 32 Or pTecla = 8 Or pTecla = 209 Or pTecla = 241) Then
      fAlfaNumSimbKeyPress = 0
   Else
      fAlfaNumSimbKeyPress = pTecla
   End If
End Function
       
Function fAlfaNumKeyPress(pTecla As Integer) As Integer
   If Not ((pTecla >= 40 And pTecla <= 57) Or (pTecla >= 97 And pTecla <= 122) Or (pTecla >= 65 And pTecla <= 90) Or pTecla = 32 Or pTecla = 8 Or pTecla = 209 Or pTecla = 241) Then
      fAlfaNumKeyPress = 0
   Else
      fAlfaNumKeyPress = pTecla
   End If
End Function

Function fAlfaKeyPress(pTecla As Integer) As Integer
   If Not ((pTecla >= 97 And pTecla <= 122) Or (pTecla >= 65 And pTecla <= 90) Or pTecla = 32 Or pTecla = 8 Or pTecla = 209 Or pTecla = 241) Then
      fAlfaKeyPress = 0
   Else
      fAlfaKeyPress = pTecla
   End If
End Function

Function fUcaseKeyPress(ByVal pTecla As Integer) As Integer
   If (pTecla >= 97 And pTecla <= 122) Or pTecla = 241 Then
      pTecla = pTecla - 32
   End If
   fUcaseKeyPress = pTecla
End Function

Public Sub sMsgEspere(ByVal pObj As Object, ByVal pMensaje As String, ByVal pVisible As Boolean)
   If pVisible Then
      pObj.Enabled = False
      ERP6_frm.pTexto = pMensaje
      ERP6_frm.Show
      ERP6_frm.Refresh
   Else
      Unload ERP6_frm
      pObj.Enabled = True
   End If
End Sub

Public Sub sPintaFlex(ByVal pObj As Object, ByVal pRow As Integer, ByVal pCols As Integer, ByVal pColor As Variant)
Dim mi As Integer
   pObj.Row = pRow
   For mi = 0 To pCols - 1
      pObj.Col = mi
      pObj.CellBackColor = pColor
   Next
End Sub

Public Sub sSetFlexRowColor(ByRef pObj As Object, ByVal pRow As Integer, ByVal pColor As Double)
Dim mi As Integer
   pObj.Row = pRow
   For mi = 0 To pObj.Cols - 1
      pObj.Col = mi
      pObj.CellBackColor = pColor
   Next
End Sub

Public Sub sSetFlex2Colors(ByRef pObj As Object, ByVal pColor1 As Double, ByVal pColor2 As Double)
Dim mi As Integer
Dim mj As Integer
   For mi = 1 To pObj.Rows - 1
      pObj.Row = mi
      For mj = 0 To pObj.Cols - 1
         pObj.Col = mj
         If (mi Mod 2) = 0 Then
            pObj.CellBackColor = pColor1
         Else
            pObj.CellBackColor = pColor2
         End If
      Next
   Next
End Sub

Public Sub sSetFlexColOrder(ByRef pObj As Object, ByVal pAlign As Integer)
Dim mi As Integer
   For mi = 0 To pObj.Cols - 1
      pObj.ColAlignment(mi) = pAlign
   Next
End Sub

Public Sub sSetFlexNroFila(ByRef pObj As Object, ByVal pCol As Integer)
Dim mi As Integer
   For mi = 1 To pObj.Rows - 1
      pObj.TextMatrix(mi, pCol) = mi
   Next
End Sub

Public Sub sBorraFlexDatos(ByRef pObj As Object)
Dim mi As Integer
   For mi = pObj.Rows - 1 To 2 Step -1
      pObj.RemoveItem mi
   Next
   For mi = 0 To pObj.Cols - 1
      On Error Resume Next
      pObj.TextMatrix(1, mi) = ""
   Next
End Sub

Public Sub fReplace(ByRef pTexto As String)
   pTexto = Replace(pTexto, "|", "")
   pTexto = Replace(pTexto, "$", "")
   pTexto = Replace(pTexto, "&", "")
   pTexto = Replace(pTexto, "'", "")
End Sub

Public Sub sCleanPic(ByRef pFlex As Object, ByVal pCol As Integer)
Dim mi As Integer
   pFlex.Row = 0
   For mi = 0 To pFlex.Cols - 1
      pFlex.Col = mi
      Set pFlex.CellPicture = Nothing
   Next
   pFlex.Col = pCol
End Sub

Public Sub sOrderFlexBub(ByRef pFlex As Object, ByVal pCol As Integer, ByVal pTipoIf As String)
Dim mTmp(15) As String
Dim mi As Integer
Dim mj As Integer
Dim mT As Integer
Dim morder As String
Dim mFlag As Boolean
   morder = "ASC"
   sCleanPic pFlex, pCol
   If pFlex.Tag <> "" Then
      If CInt(Left(pFlex.Tag, 1)) = pCol Then
         morder = "DESC"
         If Mid(pFlex.Tag, 2, 1) = "D" Then
            morder = "ASC"
         End If
      End If
   End If
   pFlex.Tag = pCol & morder
   pFlex.Row = 0
   pFlex.Col = pCol
   Set pFlex.CellPicture = LoadPicture(App.Path & "\ERP\Imagenes\" & Left(morder, 1) & ".gif")
   pFlex.CellPictureAlignment = 7
   With pFlex
      For mi = 1 To pFlex.Rows - 1
         For mj = 1 To mi - 1
            Select Case pTipoIf
               Case "T" 'texto
                  If morder = "ASC" Then
                     mFlag = (.TextMatrix(mi, pCol) < .TextMatrix(mj, pCol))
                  Else
                     mFlag = (.TextMatrix(mi, pCol) > .TextMatrix(mj, pCol))
                  End If
               Case "F" 'fecha
                  If morder = "ASC" Then
                     mFlag = (DateDiff("d", .TextMatrix(mi, pCol), .TextMatrix(mj, pCol)) > 0)
                  Else
                     mFlag = (DateDiff("d", .TextMatrix(mi, pCol), .TextMatrix(mj, pCol)) < 0)
                  End If
               Case "N"
                  If morder = "ASC" Then
                     mFlag = (.TextMatrix(mi, pCol) < .TextMatrix(mj, pCol))
                  Else
                     mFlag = (.TextMatrix(mi, pCol) > .TextMatrix(mj, pCol))
                  End If
            End Select
            If mFlag = True Then
               For mT = 0 To .Cols - 1
                  mTmp(mT) = .TextMatrix(mj, mT)
               Next
               For mT = 0 To .Cols - 1
                  .TextMatrix(mj, mT) = .TextMatrix(mi, mT)
               Next
               For mT = 0 To .Cols - 1
                  .TextMatrix(mi, mT) = mTmp(mT)
               Next
            End If
         Next
      Next
   End With
End Sub

Public Function Pat_ok(ByVal pPat As String) As Boolean
Dim mPat As String
Dim mRet As Boolean
   mPat = Trim(UCase(pPat))
   mRet = False
   If Len(mPat) = 6 Then
      If Asc(Mid(mPat, 1, 1)) >= 65 And Asc(Mid(mPat, 1, 1)) <= 90 And _
         Asc(Mid(mPat, 2, 1)) >= 65 And Asc(Mid(mPat, 1, 1)) <= 90 And _
         Asc(Mid(mPat, 3, 1)) >= 65 And Asc(Mid(mPat, 1, 1)) <= 90 And _
         IsNumeric(Mid(mPat, 4, 3)) Then
         mRet = True
      End If
   End If
   Pat_ok = mRet
End Function
Public Function fGetCodigoReferencia(pCodAlfa As String) As String
    Dim mObj As New clRNov
    fGetCodigoReferencia = Trim(mObj.sTablaDescr("novedades2", "Codigo='" & pCodAlfa & "' and Fecha = (SELECT MAX(Fecha) from regnov.novedades2 where Codigo ='" & pCodAlfa & "')", 21)) 'obtengo código de tabla
    Set mObj = Nothing
End Function
Public Function fGetKm(pCodAlfa As String) As String
    Dim mObj As New clRNov
    fGetKm = Trim(mObj.sTablaDescr("novedades2", "Codigo='" & pCodAlfa & "' and Fecha = (SELECT MAX(Fecha) from regnov.novedades2 where Codigo ='" & pCodAlfa & "')", 3)) 'obtengo código de tabla
    Set mObj = Nothing
End Function



Public Function fEnviar_Mail_CDO(SerVidor_SMTP As String, Para As String, De As String, Asunto As String, Mensaje As String, Optional Path_Adjunto As String, Optional Puerto As String = "25", Optional Usuario As String, Optional PassWord As String, Optional Usar_Autentificacion As Boolean = True, Optional Usar_SSL As Boolean = True) As Boolean
   MousePointer = vbHourglass
   ' Variable de objeto Cdo.Message
   Dim Obj_Email As CDO.Message
   ' Crea un Nuevo objeto CDO.Message
   Set Obj_Email = New CDO.Message
   ' Indica el servidor Smtp para poder enviar el Mail ( puede ser el nombre _
     del servidor o su dirección IP )
   SerVidor_SMTP = "10.10.10.243"
   Obj_Email.Configuration.Fields(cdoSMTPServer) = SerVidor_SMTP
   Obj_Email.Configuration.Fields(cdoSendUsingMethod) = 2
   ' Puerto. Por defecto se usa el puerto 25, en el caso de Gmail se usan los puertos _
     465 o  el puerto 587 ( este último me dio error )
   Obj_Email.Configuration.Fields.Item _
       ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = CLng(Puerto)
   ' Indica el tipo de autentificación con el servidor de correo _
    El valor 0 no requiere autentificarse, el valor 1 es con autentificación
   Obj_Email.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/" & _
               "configuration/smtpauthenticate") = Abs(Usar_Autentificacion)
       ' Tiempo máximo de espera en segundos para la conexión
   Obj_Email.Configuration.Fields.Item _
       ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30
   ' Configura las opciones para el login en el SMTP
   If Usar_Autentificacion Then
      ' Id de usuario del servidor Smtp ( en el caso de gmail, debe ser la dirección de correro _
       mas el @gmail.com )
      Obj_Email.Configuration.Fields.Item _
          ("http://schemas.microsoft.com/cdo/configuration/sendusername") = Usuario
      ' Password de la cuenta
      Obj_Email.Configuration.Fields.Item _
          ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = PassWord
      ' Indica si se usa SSL para el envío. En el caso de Gmail requiere que esté en True
      Obj_Email.Configuration.Fields.Item _
          ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = Usar_SSL
      Obj_Email.Configuration.Fields.Item _
         ("http://schemas.microsoft.com/cdo/configuration/enablessl") = 1
   End If
   ' *********************************************************************************
   ' Estructura del mail
   '**********************************************************************************
   ' Dirección del Destinatario
   Obj_Email.To = Para
   ' Dirección del remitente
   Obj_Email.From = De
   ' Asunto del mensaje
   Obj_Email.Subject = Asunto
   ' Cuerpo del mensaje
   Obj_Email.TextBody = Mensaje
   'Ruta del archivo adjunto
   If Path_Adjunto <> vbNullString Then
      Obj_Email.AddAttachment (Path_Adjunto)
   End If
   ' Actualiza los datos antes de enviar
   Obj_Email.Configuration.Fields.Update
   On Error Resume Next
   ' Envía el email
   Obj_Email.Send

   If Err.Number = 0 Then
      fEnviar_Mail_CDO = True
   Else
      MsgBox Err.Description, vbCritical, " Error al enviar el amil "
   End If

   ' Descarga la referencia
   If Not Obj_Email Is Nothing Then
      Set Obj_Email = Nothing
   End If

   On Error GoTo 0
   MousePointer = vbNormal
 End Function
 Public Function ContarChar(ByVal pTexto As String, pChar As String) As Integer
Dim mi As Integer
Dim mCant As Integer
mCant = 0
For mi = 1 To Len(pTexto)
   If Mid(pTexto, mi, 1) = pChar Then
      mCant = mCant + 1
   End If
Next
ContarChar = mCant
End Function

Public Function Redondeo(pNum As Double, pDec As Integer)
Dim mRet As Double
mRet = Fix(pNum * 10 ^ pDec + Sgn(pNum) * 0.5) / 10 ^ pDec
Redondeo = mRet
End Function

